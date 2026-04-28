#!/usr/bin/env bash
# provision.sh — End-to-end Agent 365 setup for NinjaClaw-Nano.
#
# Walks through every step needed to get NinjaClaw-Nano showing in
# Microsoft 365 Admin Center → Agent 365 → Available agents:
#
#   1. Create the Entra Agent ID app + client secret + Graph permissions.
#   2. Set up the messaging endpoint (Azure Bot Framework relay OR Tailscale).
#   3. Publish the blueprint to the Agent 365 admin plane.
#
# Prints what's about to happen at each step and pauses for confirmation.

set -euo pipefail

HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
# shellcheck source=lib/common.sh
source "$HERE/lib/common.sh"

MODE=""
RESOURCE_GROUP=""
LOCATION="eastus"
ENDPOINT_URL=""
TENANT_ID=""
SUBSCRIPTION_ID=""
DISPLAY_NAME="NinjaClaw-Nano"
APPLY_PUBLISH=0
ASSUME_YES=0

usage() {
  cat <<EOF
Usage: $0 --mode bot-framework|tailscale [options]

Modes
  --mode bot-framework    Use Azure Bot Framework as the relay (recommended).
                          Required: --resource-group, --location, --endpoint-url
  --mode tailscale        Front the local listener with Caddy on the tailnet.
                          No Azure resources required.

Common options
  --tenant-id GUID        Entra tenant ID (default: az account show)
  --subscription-id GUID  Azure subscription (bot-framework mode only)
  --display-name NAME     App registration display name (default: NinjaClaw-Nano)
  --apply                 Actually POST the blueprint to the admin plane.
                          Without this, step 3 stays in dry-run.
  --yes                   Skip confirmation prompts.

Bot Framework options
  --resource-group RG     Azure resource group for the bot resource.
  --location REGION       Azure region (default: eastus).
  --endpoint-url URL      Public HTTPS URL of /api/messages on this host.

Examples
  # Full bot-framework deployment behind an App Service
  $0 --mode bot-framework \\
     --resource-group ninjaclaw-rg \\
     --location eastus \\
     --endpoint-url https://my-host.azurewebsites.net/api/messages \\
     --apply

  # Tailscale-only (testing or LAN-restricted)
  $0 --mode tailscale --apply
EOF
}

while [ $# -gt 0 ]; do
  case "$1" in
    --mode)            MODE="$2"; shift 2 ;;
    --resource-group)  RESOURCE_GROUP="$2"; shift 2 ;;
    --location)        LOCATION="$2"; shift 2 ;;
    --endpoint-url)    ENDPOINT_URL="$2"; shift 2 ;;
    --tenant-id)       TENANT_ID="$2"; shift 2 ;;
    --subscription-id) SUBSCRIPTION_ID="$2"; shift 2 ;;
    --display-name)    DISPLAY_NAME="$2"; shift 2 ;;
    --apply)           APPLY_PUBLISH=1; shift ;;
    --yes|-y)          ASSUME_YES=1; shift ;;
    -h|--help)         usage; exit 0 ;;
    *)                 a365_die "unknown arg: $1" ;;
  esac
done

case "$MODE" in
  bot-framework)
    [ -n "$RESOURCE_GROUP" ] || a365_die "--resource-group is required for bot-framework mode"
    [ -n "$ENDPOINT_URL" ]   || a365_die "--endpoint-url is required for bot-framework mode"
    ;;
  tailscale)
    : # no extra requirements
    ;;
  *) a365_die "must specify --mode bot-framework or --mode tailscale" ;;
esac

confirm() {
  [ "$ASSUME_YES" = "1" ] && return 0
  printf '\n%s [y/N] ' "$1"
  read -r reply
  case "$reply" in
    y|Y|yes|YES) return 0 ;;
    *)           a365_die "aborted" ;;
  esac
}

cat <<EOF

╔══════════════════════════════════════════════════════════════════════╗
║                NinjaClaw-Nano · Agent 365 Setup                      ║
╠══════════════════════════════════════════════════════════════════════╣
║  Mode:           $MODE
║  Display name:   $DISPLAY_NAME
║  Tenant:         ${TENANT_ID:-<from az login>}
║  Subscription:   ${SUBSCRIPTION_ID:-<current az subscription>}
EOF
[ "$MODE" = "bot-framework" ] && cat <<EOF
║  Resource group: $RESOURCE_GROUP
║  Location:       $LOCATION
║  Endpoint:       $ENDPOINT_URL
EOF
cat <<EOF
║  Publish:        $([ "$APPLY_PUBLISH" = "1" ] && echo "LIVE — blueprint will be POSTed" || echo "dry-run")
╚══════════════════════════════════════════════════════════════════════╝
EOF

confirm "proceed?"

# ---------------------------------------------------------------------------
# Step 1: Entra Agent ID
# ---------------------------------------------------------------------------
a365_log "step 1/3 — Entra Agent ID"
"$HERE/01-create-entra-agent.sh" \
  --display-name "$DISPLAY_NAME" \
  ${TENANT_ID:+--tenant-id "$TENANT_ID"} \
  ${SUBSCRIPTION_ID:+--subscription-id "$SUBSCRIPTION_ID"}

confirm "step 1 complete. continue to step 2 ($MODE messaging endpoint)?"

# ---------------------------------------------------------------------------
# Step 2: Messaging endpoint
# ---------------------------------------------------------------------------
a365_log "step 2/3 — messaging endpoint ($MODE)"
case "$MODE" in
  bot-framework)
    "$HERE/02-create-bot-framework.sh" \
      --resource-group "$RESOURCE_GROUP" \
      --location "$LOCATION" \
      --endpoint-url "$ENDPOINT_URL" \
      ${SUBSCRIPTION_ID:+--subscription-id "$SUBSCRIPTION_ID"}
    ;;
  tailscale)
    "$HERE/02-setup-tailscale.sh"
    ;;
esac

confirm "step 2 complete. continue to step 3 (publish blueprint)?"

# ---------------------------------------------------------------------------
# Step 3: Publish blueprint
# ---------------------------------------------------------------------------
a365_log "step 3/3 — publish blueprint"
if [ "$APPLY_PUBLISH" = "1" ]; then
  "$HERE/03-publish-blueprint.sh" --apply
else
  "$HERE/03-publish-blueprint.sh"
fi

cat <<EOF

──────────────────────────────────────────────────────────────────────
  All steps complete.

  Next steps:
    1. Restart NinjaClaw-Nano so the Agent 365 channel picks up the new
       AGENT365_* env vars: pkill -f 'NinjaClaw' && npm start
    2. Open Microsoft 365 Admin Center → Agent 365 → Available agents
       and verify "$DISPLAY_NAME" appears.
    3. Test with the Microsoft 365 Agents Playground:
       npx -y @microsoft/m365agentsplayground
──────────────────────────────────────────────────────────────────────
EOF
