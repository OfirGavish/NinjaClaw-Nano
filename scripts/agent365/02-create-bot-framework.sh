#!/usr/bin/env bash
# 02-create-bot-framework.sh
#
# Creates an Azure Bot resource that relays Microsoft 365 / Agent 365
# activities to the local NinjaClaw-Nano host. This is the recommended
# hosting path because it works behind NAT — the Bot Framework handles
# inbound traffic and posts activities to the agent's messaging endpoint
# over HTTPS using the agent's Entra credentials.
#
# Inputs (from .env, written by 01-create-entra-agent.sh):
#   AGENT365_CLIENT_ID
#   AGENT365_TENANT_ID
#
# Required arguments:
#   --resource-group  Azure resource group to host the bot.
#   --location        Azure region (e.g. eastus, westeurope).
#   --endpoint-url    Public HTTPS URL of /api/messages on the host.
#                     For Bot Framework relay, this must be reachable from
#                     Azure's outbound IPs. If your host is behind NAT,
#                     front it with Azure Front Door, Application Gateway,
#                     or an azurewebsites.net App Service.
#
# Outputs:
#   AGENT365_BOT_NAME
#   AGENT365_MESSAGING_ENDPOINT

set -euo pipefail

HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
# shellcheck source=lib/common.sh
source "$HERE/lib/common.sh"

PROJECT_ROOT="$(a365_project_root "$HERE")"
ENV_FILE="${ENV_FILE:-$PROJECT_ROOT/.env}"

RESOURCE_GROUP=""
LOCATION="${AGENT365_LOCATION:-global}"
ENDPOINT_URL=""
BOT_NAME="${AGENT365_BOT_NAME:-}"
SUBSCRIPTION_ID="${AGENT365_SUBSCRIPTION_ID:-}"

usage() {
  cat <<EOF
Usage: $0 --resource-group RG --location REGION --endpoint-url HTTPS_URL [--bot-name NAME]

Reads AGENT365_CLIENT_ID and AGENT365_TENANT_ID from \$ENV_FILE.
EOF
}

while [ $# -gt 0 ]; do
  case "$1" in
    --resource-group)  RESOURCE_GROUP="$2"; shift 2 ;;
    --location)        LOCATION="$2"; shift 2 ;;
    --endpoint-url)    ENDPOINT_URL="$2"; shift 2 ;;
    --bot-name)        BOT_NAME="$2"; shift 2 ;;
    --subscription-id) SUBSCRIPTION_ID="$2"; shift 2 ;;
    -h|--help)         usage; exit 0 ;;
    *)                 a365_die "unknown arg: $1" ;;
  esac
done

[ -n "$RESOURCE_GROUP" ] || a365_die "--resource-group is required"
[ -n "$ENDPOINT_URL" ]   || a365_die "--endpoint-url is required"
[ -f "$ENV_FILE" ]       || a365_die "env file not found: $ENV_FILE (run 01-create-entra-agent.sh first)"

# shellcheck disable=SC1090
set -a; . "$ENV_FILE"; set +a

[ -n "${AGENT365_CLIENT_ID:-}" ] || a365_die "AGENT365_CLIENT_ID missing in $ENV_FILE"
[ -n "${AGENT365_TENANT_ID:-}" ] || a365_die "AGENT365_TENANT_ID missing in $ENV_FILE"

BOT_NAME="${BOT_NAME:-ninjaclaw-nano-$(echo -n "$AGENT365_CLIENT_ID" | cut -c1-8)}"

a365_require_cmd az
a365_az_login_check "$SUBSCRIPTION_ID"

# Ensure the Bot Service extension is available; the standalone az command
# group `az bot` was deprecated in newer versions and replaced by
# `az resource create` against `Microsoft.BotService/botServices`. We use
# the resource path directly so this works on any reasonably modern az.
a365_log "creating Azure Bot resource: $BOT_NAME (rg=$RESOURCE_GROUP, location=$LOCATION)"

# Check if the bot already exists
EXISTING="$(az resource show \
  --resource-group "$RESOURCE_GROUP" \
  --resource-type "Microsoft.BotService/botServices" \
  --name "$BOT_NAME" \
  -o json 2>/dev/null || true)"

if [ -n "$EXISTING" ] && [ "$EXISTING" != "null" ]; then
  a365_log "bot $BOT_NAME already exists, updating endpoint"
  az resource update \
    --resource-group "$RESOURCE_GROUP" \
    --resource-type "Microsoft.BotService/botServices" \
    --name "$BOT_NAME" \
    --set "properties.endpoint=$ENDPOINT_URL" \
    >/dev/null
else
  # Properties shape per Microsoft.BotService/botServices REST contract.
  PROPERTIES="$(jq -nc \
    --arg name "$BOT_NAME" \
    --arg endpoint "$ENDPOINT_URL" \
    --arg appId "$AGENT365_CLIENT_ID" \
    --arg tenantId "$AGENT365_TENANT_ID" \
    '{
      displayName: $name,
      endpoint: $endpoint,
      msaAppId: $appId,
      msaAppType: "SingleTenant",
      msaAppTenantId: $tenantId,
      isCmekEnabled: false,
      publicNetworkAccess: "Enabled"
    }')"

  az resource create \
    --resource-group "$RESOURCE_GROUP" \
    --resource-type "Microsoft.BotService/botServices" \
    --name "$BOT_NAME" \
    --location "$LOCATION" \
    --is-full-object false \
    --properties "$PROPERTIES" \
    --api-version "2022-09-15" \
    >/dev/null
fi

# Enable the Microsoft Teams channel and the agentic Microsoft 365 channel.
# Channels are subresources of the bot.
for CHANNEL_NAME in "MsTeamsChannel" "M365Channel"; do
  a365_log "enabling channel: $CHANNEL_NAME"
  CHANNEL_PROPS="$(jq -nc --arg name "$CHANNEL_NAME" '{ channelName: $name, properties: { isEnabled: true } }')"
  az resource create \
    --resource-group "$RESOURCE_GROUP" \
    --resource-type "Microsoft.BotService/botServices/channels" \
    --name "$BOT_NAME/$CHANNEL_NAME" \
    --is-full-object false \
    --properties "$CHANNEL_PROPS" \
    --api-version "2022-09-15" \
    >/dev/null 2>&1 || a365_warn "could not enable $CHANNEL_NAME (may not be available in your tenant)"
done

a365_log "writing bot info to $ENV_FILE"
a365_env_set "$ENV_FILE" "AGENT365_BOT_NAME"            "$BOT_NAME"
a365_env_set "$ENV_FILE" "AGENT365_MESSAGING_ENDPOINT"  "$ENDPOINT_URL"
a365_env_set "$ENV_FILE" "AGENT365_HOSTING_MODE"        "bot-framework"

a365_log "done. bot=$BOT_NAME  endpoint=$ENDPOINT_URL"
a365_warn "next: run 03-publish-blueprint.sh to register with the Agent 365 admin center"
