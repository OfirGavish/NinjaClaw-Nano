---
name: use-native-credential-proxy
description: Replace OneCLI gateway with the built-in credential proxy. For users who want simple .env-based credential management without installing OneCLI. Reads API key or OAuth token from .env and injects into container API requests.
---

# Use Native Credential Proxy

This skill replaces the OneCLI gateway with NinjaClaw's built-in credential proxy. Containers get credentials injected via a local HTTP proxy that reads from `.env` — no external services needed.

## Phase 1: Pre-flight

### Check if already applied

Check if `src/credential-proxy.ts` is imported in `src/index.ts`:

```bash
grep "credential-proxy" src/index.ts
```

If it shows an import for `startCredentialProxy`, the native proxy is already active. Skip to Phase 3 (Setup).

### Check if OneCLI is active

```bash
grep "@onecli-sh/sdk" package.json
```

If `@onecli-sh/sdk` appears, OneCLI is the active credential provider. Proceed with Phase 2 to replace it.

If neither check matches, you may be on an older version. Run `/update-NinjaClaw` first, then retry.

## Phase 2: Apply Code Changes

### Ensure upstream remote

```bash
git remote -v
```

If `upstream` is missing, add it:

```bash
git remote add upstream https://github.com/OfirGavish/NinjaClaw-Nano.git
```

### Merge the skill branch

```bash
git fetch upstream skill/native-credential-proxy
git merge upstream/skill/native-credential-proxy || {
  git checkout --theirs package-lock.json
  git add package-lock.json
  git merge --continue
}
```

This merges in:
- `src/credential-proxy.ts` and `src/credential-proxy.test.ts` (the proxy implementation)
- Restored credential proxy usage in `src/index.ts`, `src/container-runner.ts`, `src/container-runtime.ts`, `src/config.ts`
- Removed `@onecli-sh/sdk` dependency
- Restored `CREDENTIAL_PROXY_PORT` config (default 3001)
- Restored platform-aware proxy bind address detection
- Reverted setup skill to `.env`-based credential instructions

If the merge reports conflicts beyond `package-lock.json`, resolve them by reading the conflicted files and understanding the intent of both sides.

### Update main group COPILOT.md

Replace the OneCLI auth reference with the native proxy:

In `groups/main/COPILOT.md`, replace:
> OneCLI manages credentials (including github-copilot auth) — run `onecli --help`.

with:
> The native credential proxy manages credentials (including github-copilot auth) via `.env` — see `src/credential-proxy.ts`.

### Validate code changes

```bash
npm install
npm run build
npx vitest run src/credential-proxy.test.ts src/container-runner.test.ts
```

All tests must pass and build must be clean before proceeding.

## Phase 3: Setup Credentials

AskUserQuestion: Do you want to use your **Copilot subscription** (Pro/Max) or an **github-copilot API key**?

1. **Copilot subscription (Pro/Max)** — description: "Uses your existing Copilot Pro or Max subscription. You'll run `copilot setup-token` in another terminal to get your token."
2. **github-copilot API key** — description: "Pay-per-use API key from console.github-copilot.com."

### Subscription path

Tell the user to run `copilot setup-token` in another terminal and copy the token it outputs. Do NOT collect the token in chat.

Once they have the token, add it to `.env`:

```bash
# Add to .env (create file if needed)
echo 'GITHUB_TOKEN=<token>' >> .env
```

Note: `GITHUB_TOKEN` is also supported as a fallback.

### API key path

Tell the user to get an API key from https://console.github-copilot.com/settings/keys if they don't have one.

Add it to `.env`:

```bash
echo 'GITHUB_TOKEN=<key>' >> .env
```

### After either path

**If the user's response happens to contain a token or key** (starts with `sk-ant-` or looks like a token): write it to `.env` on their behalf using the appropriate variable name.

**Optional:** If the user needs a custom API endpoint, they can add `github-copilot_BASE_URL=<url>` to `.env` (defaults to `https://api.github.com`).

## Phase 4: Verify

1. Rebuild and restart:

```bash
npm run build
```

Then restart the service:
- macOS: `launchctl kickstart -k gui/$(id -u)/com.NinjaClaw`
- Linux: `systemctl --user restart NinjaClaw`
- WSL/manual: stop and re-run `bash start-NinjaClaw.sh`

2. Check logs for successful proxy startup:

```bash
tail -20 logs/NinjaClaw.log | grep "Credential proxy"
```

Expected: `Credential proxy started` with port and auth mode.

3. Send a test message in the registered chat to verify the agent responds.

4. Note: after applying this skill, the OneCLI credential steps in `/setup` no longer apply. `.env` is now the credential source.

## Troubleshooting

**"Credential proxy upstream error" in logs:** Check that `.env` has a valid `GITHUB_TOKEN` or `GITHUB_TOKEN`. Verify the API is reachable: `curl -s https://api.github.com/v1/messages -H "x-api-key: test" | head`.

**Port 3001 already in use:** Set `CREDENTIAL_PROXY_PORT=<other port>` in `.env` or as an environment variable.

**Container can't reach proxy (Linux):** The proxy binds to the `docker0` bridge IP by default. If that interface doesn't exist (e.g. rootless Docker), set `CREDENTIAL_PROXY_HOST=0.0.0.0` as an environment variable.

**OAuth token expired (401 errors):** Re-run `copilot setup-token` in a terminal and update the token in `.env`.

## Removal

To revert to OneCLI gateway:

1. Find the merge commit: `git log --oneline --merges -5`
2. Revert it: `git revert <merge-commit> -m 1` (undoes the skill branch merge, keeps your other changes)
3. `npm install` (re-adds `@onecli-sh/sdk`)
4. `npm run build`
5. Follow `/setup` step 4 to configure OneCLI credentials
6. Remove `GITHUB_TOKEN` / `GITHUB_TOKEN` from `.env`
