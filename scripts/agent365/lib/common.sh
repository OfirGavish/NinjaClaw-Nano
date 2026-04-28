#!/usr/bin/env bash
# Common helpers for the Agent 365 provisioning scripts.
# Sourced by the other scripts in this directory.

set -euo pipefail

A365_LOG_PREFIX="${A365_LOG_PREFIX:-agent365}"

a365_log()  { printf '\033[1;36m[%s]\033[0m %s\n' "$A365_LOG_PREFIX" "$*"; }
a365_warn() { printf '\033[1;33m[%s]\033[0m %s\n' "$A365_LOG_PREFIX" "$*" >&2; }
a365_err()  { printf '\033[1;31m[%s]\033[0m %s\n' "$A365_LOG_PREFIX" "$*" >&2; }
a365_die()  { a365_err "$*"; exit 1; }

a365_require_cmd() {
  for cmd in "$@"; do
    command -v "$cmd" >/dev/null 2>&1 || a365_die "missing required command: $cmd"
  done
}

# Append (or replace) a KEY=VALUE pair in a .env file.
# Usage: a365_env_set <env-file> <KEY> <VALUE>
a365_env_set() {
  local file="$1" key="$2" value="$3"
  touch "$file"
  if grep -qE "^${key}=" "$file"; then
    # macOS sed and GNU sed differ on -i; use a portable temp-file dance.
    local tmp
    tmp="$(mktemp)"
    grep -vE "^${key}=" "$file" > "$tmp"
    printf '%s=%s\n' "$key" "$value" >> "$tmp"
    mv "$tmp" "$file"
  else
    printf '%s=%s\n' "$key" "$value" >> "$file"
  fi
}

# Verify the user is logged in to az and on the requested subscription.
a365_az_login_check() {
  local subscription_id="${1:-}"
  a365_require_cmd az
  if ! az account show >/dev/null 2>&1; then
    a365_die "az cli not logged in. run: az login"
  fi
  if [ -n "$subscription_id" ]; then
    a365_log "switching to subscription $subscription_id"
    az account set --subscription "$subscription_id" >/dev/null
  fi
}

# Resolve the project root (directory containing package.json) starting from
# the script's location.
a365_project_root() {
  local here="$1"
  local dir
  dir="$(cd "$(dirname "$here")" && pwd)"
  while [ "$dir" != "/" ]; do
    if [ -f "$dir/package.json" ]; then
      printf '%s\n' "$dir"
      return 0
    fi
    dir="$(dirname "$dir")"
  done
  a365_die "could not locate project root from $here"
}
