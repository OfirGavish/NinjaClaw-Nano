#!/usr/bin/env bash
# 02-setup-tailscale.sh
#
# Alternative hosting path for users who don't want to put their NinjaClaw-Nano
# host endpoint on the public internet via Bot Framework relay. This script
# fronts the local agent-365 listener with Caddy on the tailnet, terminating
# TLS using Tailscale's built-in HTTPS certificate provisioning.
#
# Result: a stable wss:// URL of the form
#   https://<machine-name>.<tailnet-name>.ts.net/api/messages
#
# That URL is the value you'll register as the messaging endpoint in the
# Agent 365 admin blueprint.
#
# Caveats vs Bot Framework relay:
#   - Microsoft 365 cloud cannot reach a tailnet directly. This path is for
#     scenarios where you've added the M365 / Agent 365 service identity to
#     your tailnet via a tailscale ACL, OR where you're testing locally
#     against the Microsoft 365 Agents Playground.
#   - For real M365 production traffic, use 02-create-bot-framework.sh.

set -euo pipefail

HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
# shellcheck source=lib/common.sh
source "$HERE/lib/common.sh"

PROJECT_ROOT="$(a365_project_root "$HERE")"
ENV_FILE="${ENV_FILE:-$PROJECT_ROOT/.env}"
AGENT_PORT="${AGENT365_PORT:-3979}"

usage() {
  cat <<EOF
Usage: $0 [--port PORT]

Requires:
  - tailscale CLI installed and logged in (tailscale up)
  - HTTPS enabled on the tailnet (tailscale set --advertise-features=https)
  - root or sudo access (Caddy binds :443 and reads tailscale state)

Writes the resulting public messaging endpoint URL into \$ENV_FILE as
AGENT365_MESSAGING_ENDPOINT.
EOF
}

while [ $# -gt 0 ]; do
  case "$1" in
    --port)    AGENT_PORT="$2"; shift 2 ;;
    -h|--help) usage; exit 0 ;;
    *)         a365_die "unknown arg: $1" ;;
  esac
done

a365_require_cmd tailscale curl

# Resolve the tailscale FQDN of this machine.
TS_STATUS_JSON="$(tailscale status --json 2>/dev/null || true)"
[ -n "$TS_STATUS_JSON" ] || a365_die "tailscale not running. run: tailscale up"

TS_FQDN="$(printf '%s' "$TS_STATUS_JSON" | jq -r '.Self.DNSName' | sed 's/\.$//')"
[ -n "$TS_FQDN" ] && [ "$TS_FQDN" != "null" ] || a365_die "could not resolve tailscale FQDN"

a365_log "tailscale FQDN: $TS_FQDN"

# Verify HTTPS is enabled on this tailnet.
if ! printf '%s' "$TS_STATUS_JSON" | jq -e '.CurrentTailnet.MagicDNSSuffix' >/dev/null; then
  a365_warn "MagicDNS may not be enabled. enable it in the Tailscale admin console"
fi

# Provision the TLS cert (one-shot; tailscale renews automatically).
a365_log "provisioning Tailscale HTTPS certificate (may take ~30s on first run)"
CERT_DIR="${AGENT365_CERT_DIR:-/var/lib/ninjaclaw-nano/agent365}"
sudo mkdir -p "$CERT_DIR"
sudo chown "$(id -u):$(id -g)" "$CERT_DIR"
tailscale cert --cert-file "$CERT_DIR/cert.pem" --key-file "$CERT_DIR/key.pem" "$TS_FQDN"

# Install / refresh Caddy as the TLS-terminating reverse proxy. Using Caddy
# (rather than tailscale serve) so we get reload-friendly config and explicit
# logs. Caddy is installed only if missing; we don't manage upgrades.
if ! command -v caddy >/dev/null 2>&1; then
  a365_log "installing Caddy"
  if command -v apt-get >/dev/null 2>&1; then
    sudo apt-get update -y
    sudo apt-get install -y debian-keyring debian-archive-keyring apt-transport-https curl gnupg
    curl -fsSL "https://dl.cloudsmith.io/public/caddy/stable/gpg.key" | sudo gpg --dearmor -o /usr/share/keyrings/caddy-stable-archive-keyring.gpg
    curl -fsSL "https://dl.cloudsmith.io/public/caddy/stable/debian.deb.txt" | sudo tee /etc/apt/sources.list.d/caddy-stable.list >/dev/null
    sudo apt-get update -y
    sudo apt-get install -y caddy
  elif command -v brew >/dev/null 2>&1; then
    brew install caddy
  else
    a365_die "no supported package manager found; install Caddy manually from https://caddyserver.com/docs/install"
  fi
fi

CADDY_CONF="/etc/caddy/ninjaclaw-agent365.caddy"
sudo tee "$CADDY_CONF" >/dev/null <<EOF
# Managed by NinjaClaw-Nano scripts/agent365/02-setup-tailscale.sh
# Re-run that script to regenerate. Manual edits will be overwritten.
$TS_FQDN {
    tls $CERT_DIR/cert.pem $CERT_DIR/key.pem

    handle /api/messages {
        reverse_proxy 127.0.0.1:$AGENT_PORT
    }

    handle /api/health {
        respond "ok" 200
    }

    handle {
        respond "NinjaClaw-Nano Agent 365 endpoint" 200
    }

    log {
        output file /var/log/caddy/ninjaclaw-agent365.log
        format json
    }
}
EOF

# Wire the snippet into the main Caddyfile if it isn't already.
MAIN_CADDYFILE="/etc/caddy/Caddyfile"
sudo touch "$MAIN_CADDYFILE"
if ! sudo grep -qF "import $CADDY_CONF" "$MAIN_CADDYFILE"; then
  echo "import $CADDY_CONF" | sudo tee -a "$MAIN_CADDYFILE" >/dev/null
fi

a365_log "reloading Caddy"
if command -v systemctl >/dev/null 2>&1; then
  sudo systemctl reload caddy || sudo systemctl restart caddy
else
  sudo caddy reload --config "$MAIN_CADDYFILE" || a365_warn "manual Caddy reload required"
fi

ENDPOINT_URL="https://$TS_FQDN/api/messages"
a365_log "writing endpoint to $ENV_FILE"
a365_env_set "$ENV_FILE" "AGENT365_MESSAGING_ENDPOINT" "$ENDPOINT_URL"
a365_env_set "$ENV_FILE" "AGENT365_HOSTING_MODE"       "tailscale"
a365_env_set "$ENV_FILE" "AGENT365_TAILSCALE_FQDN"     "$TS_FQDN"

a365_log "done. messaging endpoint: $ENDPOINT_URL"
a365_warn "verify reachability from outside the tailnet before publishing the blueprint"
