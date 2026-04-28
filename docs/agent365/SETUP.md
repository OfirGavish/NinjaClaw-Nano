# Agent 365 Setup — NinjaClaw-Nano

End-to-end guide for getting NinjaClaw-Nano registered with Microsoft Agent 365
so it appears in the **Microsoft 365 Admin Center → Agent 365 → Available agents**
list and accepts traffic from Teams, Outlook, Word, Excel, and PowerPoint
(plus the existing Telegram and web channels, which are unaffected).

> **Status:** Preview. Microsoft Agent 365 is part of the Frontier program;
> the blueprint publish API is internal-preview at the time of writing.
> Phases 1–2 (identity, observability, governed MCP tooling) are fully
> implemented and verifiable today; phase 3 (the publish API call) ships in
> dry-run by default until you confirm the internal endpoint with your
> Microsoft Frontier contact and run with `--apply`.

## Prerequisites

- A Microsoft 365 tenant enrolled in the Frontier preview.
- Tenant admin (or Application Administrator + Cloud Application
  Administrator) role for the Entra app registration step.
- `az` CLI v2.55+, `jq`, `curl`, and either `bash` 4+ on Linux/macOS or
  WSL on Windows.
- For **bot-framework** mode: an Azure subscription and a publicly reachable
  HTTPS endpoint for `/api/messages` (App Service, Front Door, or
  Application Gateway in front of your VM).
- For **tailscale** mode: `tailscale` installed and logged in, MagicDNS +
  HTTPS enabled on the tailnet.

## One-shot path

```bash
# Bot Framework relay (recommended — works behind NAT, full M365 reach)
./scripts/agent365/provision.sh \
    --mode bot-framework \
    --resource-group ninjaclaw-rg \
    --location eastus \
    --endpoint-url https://my-host.azurewebsites.net/api/messages \
    --apply
```

```bash
# Tailscale (LAN-restricted, useful for Frontier preview testing)
./scripts/agent365/provision.sh --mode tailscale --apply
```

`provision.sh` walks through three steps and pauses for confirmation
between each.

## What each step does

### 1. `01-create-entra-agent.sh`
- Creates (or reuses) an Entra app registration with display name
  `NinjaClaw-Nano`.
- Mints a 2-year client secret.
- Requests Microsoft Graph delegated scopes for Mail, Calendar, Files,
  Sites, Chat, ChannelMessage, and User.Read.
- Writes `AGENT365_CLIENT_ID`, `AGENT365_TENANT_ID`,
  `AGENT365_CLIENT_SECRET`, `AGENT365_OBJECT_ID` to `.env`.

> **Manual follow-up:** the script requests permissions but cannot grant
> admin consent on your behalf. Run this once after the script completes:
>
> ```bash
> az ad app permission admin-consent --id "$AGENT365_CLIENT_ID"
> ```

### 2a. `02-create-bot-framework.sh` (bot-framework mode)
- Creates an Azure Bot resource (`Microsoft.BotService/botServices`)
  pointing at your `--endpoint-url`.
- Enables the Microsoft Teams channel and the agentic M365 channel on
  the bot.
- Writes `AGENT365_BOT_NAME` and `AGENT365_MESSAGING_ENDPOINT` to `.env`.

### 2b. `02-setup-tailscale.sh` (tailscale mode)
- Provisions a Tailscale HTTPS certificate for the local machine.
- Installs and configures Caddy as a reverse proxy that terminates TLS
  on the tailnet FQDN and forwards `/api/messages` to the local
  NinjaClaw-Nano host on `AGENT365_PORT` (default `3979`).
- Writes `AGENT365_MESSAGING_ENDPOINT` (e.g.
  `https://machine.tailnet.ts.net/api/messages`) to `.env`.

### 3. `03-publish-blueprint.sh`
- Reads `agent365/blueprint.json` and merges in the deployment-specific
  fields (app ID, tenant, endpoint).
- Acquires a client-credentials token for `AGENT365_PUBLISH_AUDIENCE`.
- POSTs the merged payload to `AGENT365_PUBLISH_API`.
- Defaults to dry-run; pass `--apply` to actually submit.

> **You will need to confirm two values from internal Microsoft sources
> before the publish step succeeds:**
>
> - `AGENT365_PUBLISH_API` — the admin-plane URL that accepts blueprint
>   POSTs. The default in the script
>   (`https://agent365.svc.cloud.microsoft/admin/v1/blueprints`) is a
>   best-guess based on the runtime SDK's host.
> - `AGENT365_PUBLISH_AUDIENCE` — the token audience the admin plane
>   accepts. Default `https://agent365.svc.cloud.microsoft/.default`.
>
> Override either via env or by editing `.env` before running step 3.

## Verifying the result

1. Restart NinjaClaw-Nano so it picks up the new env vars:
   ```bash
   pkill -f 'NinjaClaw' && npm start
   ```
2. Tail the host logs and look for `Agent 365 channel listening`.
3. Open the **Microsoft 365 Admin Center → Agent 365 → Available agents**
   page in the tenant. The agent should appear with the icon, name, and
   description from `agent365/blueprint.json`.
4. End-to-end smoke test using the Microsoft 365 Agents Playground:
   ```bash
   npx -y @microsoft/m365agentsplayground
   ```
5. Check the Agent 365 observability blade — turns should appear as OTLP
   traces tagged with `agentic.message`, `agentic.email`, etc.

## Repeating / updating

All three scripts are idempotent:

- Re-running `01-create-entra-agent.sh` reuses the existing app and
  appends a new secret (old secrets stay valid until manually removed).
- Re-running `02-create-bot-framework.sh` updates the messaging endpoint.
- Re-running `03-publish-blueprint.sh --apply` republishes the blueprint
  (treat as an upsert; the admin plane keeps the same blueprint ID).

To change the displayed name, icon, or requested permissions, edit
`agent365/blueprint.json` and re-run step 3.

## Troubleshooting

| Symptom | Likely cause | Fix |
|---|---|---|
| `az ad app permission add` fails | scope GUID wrong / API not consented | check Microsoft Graph permissions reference for the current ID |
| Step 3 returns HTTP 401/403 | admin consent missing OR wrong audience | run `az ad app permission admin-consent` and verify `AGENT365_PUBLISH_AUDIENCE` |
| Step 3 returns HTTP 404 | publish URL is wrong | confirm `AGENT365_PUBLISH_API` with internal contact |
| Bot Framework channel posts return 401 | endpoint can't validate JWT | verify `AGENT365_CLIENT_ID` / `AGENT365_TENANT_ID` match the bot's `msaAppId` / `msaAppTenantId` |
| Tailscale endpoint not reachable from M365 | M365 cloud isn't on your tailnet | use bot-framework mode for production |
| Agent doesn't appear in admin center | blueprint ID collision OR pending tenant approval | check the response body of the publish call; some tenants require manual admin approval before listing |

## What this does NOT change

- Existing Telegram, web UI, and legacy Teams channels are untouched.
- `GITHUB_TOKEN` and the per-conversation Docker container model auth path
  are untouched. Agent 365 is for *the agent's identity in M365*, not for
  model inference.
- Per-VM, per-user isolation is preserved. Each user's instance has its
  own Entra app, its own bot resource, and its own blueprint deployment.
