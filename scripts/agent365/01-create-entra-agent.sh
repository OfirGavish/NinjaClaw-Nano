#!/usr/bin/env bash
# 01-create-entra-agent.sh
#
# Provisions the Entra Agent ID app registration that the NinjaClaw-Nano host
# uses to authenticate with Microsoft Agent 365 and to perform OBO token
# exchanges for governed MCP tools.
#
# Outputs:
#   - AGENT365_CLIENT_ID
#   - AGENT365_TENANT_ID
#   - AGENT365_CLIENT_SECRET
#   - AGENT365_OBJECT_ID  (used by the Bot Framework + blueprint publish steps)
#
# Idempotent: if an app with the same display name already exists in the
# tenant, the script reuses it rather than creating a duplicate.

set -euo pipefail

HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
# shellcheck source=lib/common.sh
source "$HERE/lib/common.sh"

PROJECT_ROOT="$(a365_project_root "$HERE")"
ENV_FILE="${ENV_FILE:-$PROJECT_ROOT/.env}"
DISPLAY_NAME="${AGENT365_DISPLAY_NAME:-NinjaClaw-Nano}"
TENANT_ID="${AGENT365_TENANT_ID:-}"
SUBSCRIPTION_ID="${AGENT365_SUBSCRIPTION_ID:-}"

usage() {
  cat <<EOF
Usage: $0 [--display-name NAME] [--tenant-id GUID] [--subscription-id GUID]

Options can also be supplied via env vars:
  AGENT365_DISPLAY_NAME, AGENT365_TENANT_ID, AGENT365_SUBSCRIPTION_ID

Writes credentials into .env at the repo root (or \$ENV_FILE).
EOF
}

while [ $# -gt 0 ]; do
  case "$1" in
    --display-name)    DISPLAY_NAME="$2"; shift 2 ;;
    --tenant-id)       TENANT_ID="$2"; shift 2 ;;
    --subscription-id) SUBSCRIPTION_ID="$2"; shift 2 ;;
    -h|--help)         usage; exit 0 ;;
    *)                 a365_die "unknown arg: $1" ;;
  esac
done

a365_require_cmd az jq
a365_az_login_check "$SUBSCRIPTION_ID"

if [ -z "$TENANT_ID" ]; then
  TENANT_ID="$(az account show --query tenantId -o tsv)"
  a365_log "resolved tenant: $TENANT_ID"
fi

a365_log "looking up existing app registration: $DISPLAY_NAME"
APP_JSON="$(az ad app list --display-name "$DISPLAY_NAME" --query '[0]' -o json)"
if [ "$APP_JSON" = "null" ] || [ -z "$APP_JSON" ]; then
  a365_log "creating Entra app: $DISPLAY_NAME"
  APP_JSON="$(az ad app create \
    --display-name "$DISPLAY_NAME" \
    --sign-in-audience "AzureADMyOrg" \
    --enable-id-token-issuance true \
    -o json)"
else
  a365_log "reusing existing app registration"
fi

CLIENT_ID="$(jq -r .appId <<<"$APP_JSON")"
OBJECT_ID="$(jq -r .id <<<"$APP_JSON")"

# Ensure a service principal exists for the app (required for token issuance
# and for assigning Graph permissions).
if ! az ad sp show --id "$CLIENT_ID" >/dev/null 2>&1; then
  a365_log "creating service principal"
  az ad sp create --id "$CLIENT_ID" >/dev/null
else
  a365_log "service principal already exists"
fi

# Mint (or rotate) a client secret. We always create a new one so the .env
# always contains a credential the host can actually use; old secrets are
# left in place and can be removed manually if needed.
a365_log "creating client secret (valid 2 years)"
SECRET_JSON="$(az ad app credential reset \
  --id "$CLIENT_ID" \
  --display-name "ninjaclaw-nano-$(date -u +%Y%m%d)" \
  --years 2 \
  --append \
  -o json)"
CLIENT_SECRET="$(jq -r .password <<<"$SECRET_JSON")"

# --- Microsoft Graph delegated permissions --------------------------------
# Resource App ID for Microsoft Graph is a fixed well-known GUID.
GRAPH_APP_ID="00000003-0000-0000-c000-000000000000"

# Map Graph delegated scope → permission ID. These are the well-known IDs
# from the Microsoft Graph permissions reference; they're stable.
declare -A GRAPH_SCOPES=(
  ["Mail.ReadWrite"]="024d486e-b451-40bb-833d-3e66d98c5c73"
  ["Mail.Send"]="e383f46e-2787-4529-855e-0e479a3ffac0"
  ["Calendars.ReadWrite"]="1ec239c2-d7c9-4623-a91a-a9775856bb36"
  ["Files.ReadWrite"]="5c28f0bf-8a70-41f1-8ab2-9032436ddb65"
  ["Sites.Read.All"]="205e70e5-aba6-4c52-a976-6d2d46c48043"
  ["Chat.ReadWrite"]="9ff7295e-131b-4d94-90e1-69fde507ac11"
  ["ChannelMessage.Read.All"]="767156cb-16ae-4d10-8f8b-41b657c8c8c8"
  ["User.Read"]="e1fe6dd8-ba31-4d61-89e7-88639da4683d"
)

a365_log "requesting Microsoft Graph delegated permissions"
for scope in "${!GRAPH_SCOPES[@]}"; do
  az ad app permission add \
    --id "$CLIENT_ID" \
    --api "$GRAPH_APP_ID" \
    --api-permissions "${GRAPH_SCOPES[$scope]}=Scope" \
    >/dev/null 2>&1 || a365_warn "could not add scope $scope (may already be present)"
done

a365_warn "permissions requested but NOT yet granted admin consent."
a365_warn "to consent: az ad app permission admin-consent --id $CLIENT_ID"
a365_warn "(requires Global Administrator or Privileged Role Administrator)"

# --- Persist to .env ------------------------------------------------------
a365_log "writing credentials to $ENV_FILE"
a365_env_set "$ENV_FILE" "AGENT365_CLIENT_ID"     "$CLIENT_ID"
a365_env_set "$ENV_FILE" "AGENT365_TENANT_ID"     "$TENANT_ID"
a365_env_set "$ENV_FILE" "AGENT365_CLIENT_SECRET" "$CLIENT_SECRET"
a365_env_set "$ENV_FILE" "AGENT365_OBJECT_ID"     "$OBJECT_ID"

a365_log "done. AGENT365_CLIENT_ID=$CLIENT_ID"
