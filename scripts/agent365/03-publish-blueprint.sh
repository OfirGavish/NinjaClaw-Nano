#!/usr/bin/env bash
# 03-publish-blueprint.sh
#
# Submits the agent blueprint to the Microsoft Agent 365 admin plane so that
# the agent appears in the Microsoft 365 Admin Center under
# Agent 365 → Agents → Available agents.
#
# Inputs:
#   - agent365/blueprint.json
#   - .env values written by previous scripts:
#       AGENT365_CLIENT_ID
#       AGENT365_TENANT_ID
#       AGENT365_OBJECT_ID
#       AGENT365_MESSAGING_ENDPOINT
#       AGENT365_BOT_NAME            (bot-framework mode)
#       AGENT365_TAILSCALE_FQDN      (tailscale mode)
#
# IMPORTANT — INTERNAL MICROSOFT ENDPOINT
# ---------------------------------------------------------------------------
# As of this writing, the Agent 365 blueprint publish API is gated to the
# Frontier preview and not on public docs. The default endpoint below is a
# best-guess based on the runtime SDK's `agent365.svc.cloud.microsoft` host;
# verify the exact path with internal Microsoft channels and override via
#
#     AGENT365_PUBLISH_API
#     AGENT365_PUBLISH_AUDIENCE   (token audience for OBO exchange)
#
# Until you've confirmed those, the script runs in --dry-run mode by default
# and just prints the request payload + the curl command it would execute.

set -euo pipefail

HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
# shellcheck source=lib/common.sh
source "$HERE/lib/common.sh"

PROJECT_ROOT="$(a365_project_root "$HERE")"
ENV_FILE="${ENV_FILE:-$PROJECT_ROOT/.env}"
BLUEPRINT_FILE="${AGENT365_BLUEPRINT:-$PROJECT_ROOT/agent365/blueprint.json}"

# TODO(internal): confirm with Microsoft Frontier program contact.
PUBLISH_API="${AGENT365_PUBLISH_API:-https://agent365.svc.cloud.microsoft/admin/v1/blueprints}"
PUBLISH_AUDIENCE="${AGENT365_PUBLISH_AUDIENCE:-https://agent365.svc.cloud.microsoft/.default}"

DRY_RUN=1
APPLY=0

usage() {
  cat <<EOF
Usage: $0 [--apply]

  --apply   actually POST to the Agent 365 admin API. Default is dry-run:
            print the payload + curl command without sending.

Required env (set automatically by previous scripts):
  AGENT365_CLIENT_ID
  AGENT365_TENANT_ID
  AGENT365_CLIENT_SECRET
  AGENT365_OBJECT_ID
  AGENT365_MESSAGING_ENDPOINT

Optional overrides (for internal endpoints):
  AGENT365_PUBLISH_API       (POST URL; default: $PUBLISH_API)
  AGENT365_PUBLISH_AUDIENCE  (token audience; default: $PUBLISH_AUDIENCE)
  AGENT365_BLUEPRINT         (path to blueprint.json; default: $BLUEPRINT_FILE)
EOF
}

while [ $# -gt 0 ]; do
  case "$1" in
    --apply)   APPLY=1; DRY_RUN=0; shift ;;
    -h|--help) usage; exit 0 ;;
    *)         a365_die "unknown arg: $1" ;;
  esac
done

a365_require_cmd jq curl
[ -f "$BLUEPRINT_FILE" ] || a365_die "blueprint not found: $BLUEPRINT_FILE"
[ -f "$ENV_FILE" ]       || a365_die "env file not found: $ENV_FILE"

# shellcheck disable=SC1090
set -a; . "$ENV_FILE"; set +a

[ -n "${AGENT365_CLIENT_ID:-}" ]            || a365_die "AGENT365_CLIENT_ID missing"
[ -n "${AGENT365_TENANT_ID:-}" ]            || a365_die "AGENT365_TENANT_ID missing"
[ -n "${AGENT365_CLIENT_SECRET:-}" ]        || a365_die "AGENT365_CLIENT_SECRET missing"
[ -n "${AGENT365_MESSAGING_ENDPOINT:-}" ]   || a365_die "AGENT365_MESSAGING_ENDPOINT missing — run 02-create-bot-framework.sh or 02-setup-tailscale.sh first"

a365_log "building publish payload from $BLUEPRINT_FILE"
PAYLOAD="$(jq -c \
  --arg appId    "$AGENT365_CLIENT_ID" \
  --arg tenantId "$AGENT365_TENANT_ID" \
  --arg objectId "${AGENT365_OBJECT_ID:-}" \
  --arg endpoint "$AGENT365_MESSAGING_ENDPOINT" \
  --arg botName  "${AGENT365_BOT_NAME:-}" \
  --arg mode     "${AGENT365_HOSTING_MODE:-bot-framework}" \
  '. + {
    deployment: {
      appId: $appId,
      tenantId: $tenantId,
      objectId: $objectId,
      messagingEndpoint: $endpoint,
      botName: ($botName | select(length > 0)),
      hostingMode: $mode
    }
  }' \
  "$BLUEPRINT_FILE")"

a365_log "acquiring access token (audience: $PUBLISH_AUDIENCE)"
TOKEN_RESPONSE="$(curl -fsSL -X POST \
  "https://login.microsoftonline.com/$AGENT365_TENANT_ID/oauth2/v2.0/token" \
  -H "Content-Type: application/x-www-form-urlencoded" \
  --data-urlencode "client_id=$AGENT365_CLIENT_ID" \
  --data-urlencode "client_secret=$AGENT365_CLIENT_SECRET" \
  --data-urlencode "grant_type=client_credentials" \
  --data-urlencode "scope=$PUBLISH_AUDIENCE" 2>&1)"

ACCESS_TOKEN="$(printf '%s' "$TOKEN_RESPONSE" | jq -r .access_token 2>/dev/null || echo "")"
if [ -z "$ACCESS_TOKEN" ] || [ "$ACCESS_TOKEN" = "null" ]; then
  a365_warn "could not obtain access token. response was:"
  printf '%s\n' "$TOKEN_RESPONSE" | sed 's/^/    /' >&2
  a365_warn "this usually means PUBLISH_AUDIENCE is wrong or admin consent has not been granted yet."
  a365_warn "consent: az ad app permission admin-consent --id $AGENT365_CLIENT_ID"
  if [ "$DRY_RUN" = "0" ]; then
    a365_die "no token, cannot --apply"
  fi
  ACCESS_TOKEN="<DRY-RUN-NO-TOKEN>"
fi

if [ "$DRY_RUN" = "1" ]; then
  cat <<EOF

──────────────────────────────────────────────────────────────────────
  DRY RUN — would POST the following:
──────────────────────────────────────────────────────────────────────
  curl -X POST "$PUBLISH_API" \\
    -H "Authorization: Bearer \$ACCESS_TOKEN" \\
    -H "Content-Type: application/json" \\
    --data-binary @-

  payload (\$(wc -c <<<"\$PAYLOAD") bytes):
$(printf '%s' "$PAYLOAD" | jq .)

  to actually publish, run:  $0 --apply
──────────────────────────────────────────────────────────────────────
EOF
  exit 0
fi

a365_log "POSTing blueprint to $PUBLISH_API"
HTTP_RESPONSE_FILE="$(mktemp)"
HTTP_STATUS="$(curl -sS -o "$HTTP_RESPONSE_FILE" -w '%{http_code}' \
  -X POST "$PUBLISH_API" \
  -H "Authorization: Bearer $ACCESS_TOKEN" \
  -H "Content-Type: application/json" \
  --data-binary "$PAYLOAD")"

a365_log "HTTP $HTTP_STATUS"
cat "$HTTP_RESPONSE_FILE"
echo
rm -f "$HTTP_RESPONSE_FILE"

case "$HTTP_STATUS" in
  2*) a365_log "blueprint published. check Microsoft 365 Admin Center → Agent 365 → Available agents" ;;
  401|403) a365_die "auth rejected. verify PUBLISH_AUDIENCE and admin consent." ;;
  404) a365_die "publish endpoint not found. confirm AGENT365_PUBLISH_API with internal contact." ;;
  *) a365_die "publish failed (HTTP $HTTP_STATUS)" ;;
esac
