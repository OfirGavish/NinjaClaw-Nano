# Agent 365 Integration — Design Doc

> Status: **Draft / Phase planning**
> Owner: @OfirGavish
> Target: NinjaClaw-Nano v.next

## Goal

Make NinjaClaw-Nano a first-class **Microsoft Agent 365** agent so that:

1. The agent is registered as an **Entra Agent ID** with an IT-approved blueprint (governance, lifecycle, audit).
2. Every turn emits **OpenTelemetry** spans that flow into the **Agent 365 Observability** ingest.
3. (Phase 2) Users can talk to the agent natively from **Teams, Outlook, Word comments**.
4. (Phase 2) The agent can read/write the user's **Outlook · Teams · OneDrive · SharePoint** data through governed **MCP** servers, on behalf of the user.

**Explicitly out of scope:** changing how the agent talks to its language model. The existing `GITHUB_TOKEN` → GitHub Copilot SDK flow stays exactly as it is. Agent 365 identity is for *the agent's identity in the M365 ecosystem*, not for model inference.

---

## Background — what Agent 365 actually is

Agent 365 is **not** an agent framework. It is an enterprise control plane that wraps an existing agent (built in any framework) and adds five things:

| Capability | What it is | Phase |
|---|---|---|
| **Entra Agent Identity** | Agent gets its own Entra-backed identity (and optionally a mailbox / user resources). Agent JWT validated by the SDK at the receive endpoint. | 1 |
| **Observability** | OTel spans/metrics/logs for every turn, tool call, inference event. Exported via OTLP (auth via agentic token) to `agent365.svc.cloud.microsoft`. | 1 |
| **Notifications** | Receive @mentions / emails from Teams, Outlook, Word as `Activity` objects. Reply with `sendActivity()` or `createEmailResponseActivity()`. | 2 |
| **Governed MCP tooling** | Admin-curated MCP servers (SharePoint search, Mail, Calendar, Teams). Permissions enforced at the gateway. | 2 |
| **Blueprints** | IT-approved templates published to M365 Admin Center. Users request agent instances; each instance inherits DLP, audit, external-access policies. | 3 |

Architecture per Microsoft's own layering:

```
┌─────────────────────────────────────────────┐
│ Enterprise Capabilities  ← Agent 365 SDK    │  added in this work
├─────────────────────────────────────────────┤
│ Agent Logic              ← NinjaClaw code   │  unchanged
├─────────────────────────────────────────────┤
│ LLM Orchestrator         ← GH Copilot SDK   │  unchanged
└─────────────────────────────────────────────┘
```

---

## Architecture

```mermaid
flowchart TB
    subgraph M365["Microsoft 365 cloud"]
        Teams["Teams chats / channels"]
        Outlook["Outlook (email)"]
        Word["Word comments"]
        Entra["Entra Agent ID<br/>blueprint + identity"]
        A365Obs["Agent 365<br/>Observability ingest"]
        Graph["Microsoft Graph<br/>(Mail / Calendar / OneDrive / SharePoint / Teams)"]
        MCP["Governed MCP servers<br/>(SharePoint search, Mail, etc.)"]
    end

    subgraph VM["Azure VM (per user)"]
        subgraph Host["NinjaClaw-Nano host process"]
            Gateway["Existing channels<br/>Telegram · Web UI · Teams (legacy)"]
            A365Channel["NEW: Agent 365 channel<br/>Express adapter @ /api/messages<br/>AgentApplication { authorization: agentic }"]
            Brain["NinjaBrain<br/>(SQLite knowledge)"]
            Otel["OTel exporter<br/>→ agent365.svc.cloud.microsoft"]
        end

        subgraph Container["Docker container (per conversation)"]
            CopilotSDK["GitHub Copilot SDK loop<br/>GITHUB_TOKEN unchanged"]
            Tools["bash · read · write · brain-cli<br/>+ Graph tools (Phase 2)"]
        end
    end

    Teams -->|user @mentions agent| A365Channel
    Outlook -->|email to agent| A365Channel
    Word -->|@mention in comment| A365Channel

    A365Channel -.->|validates JWT against| Entra
    A365Channel -->|spawns / dispatches turn| Container
    Gateway -->|spawns / dispatches turn| Container

    CopilotSDK -->|reads/writes| Brain
    CopilotSDK -->|tool calls| Tools

    Container -->|emits OTel spans| Otel
    A365Channel -->|emits OTel spans| Otel
    Otel -->|OTLP authenticated<br/>via agentic token| A365Obs

    Tools -.->|Phase 2: on-behalf-of user via<br/>aadObjectId from activity.from| MCP
    MCP -->|admin-governed access| Graph

    classDef new fill:#d4ffd4,stroke:#2a7a2a,stroke-width:2px
    classDef phase2 fill:#fff5d4,stroke:#a07700,stroke-width:1.5px,stroke-dasharray:4 3
    class A365Channel,Otel,Entra,A365Obs new
    class MCP,Graph,Tools phase2
```

Green = Phase 1. Yellow dashed = Phase 2.

### Key design decision: Agent 365 lives in the **host**, not the container

The Agent 365 SDK is fundamentally an HTTP receiver pattern (Express endpoint at `/api/messages`, Entra JWT validation, `AgentApplication` event routing). It must be reachable from the M365 cloud and must hold an Entra credential. Our containers are ephemeral, isolated, and intentionally credential-free.

**Therefore:** the Agent 365 endpoint runs in the **NinjaClaw-Nano host process** as a new channel adapter, alongside the existing Telegram / Web UI / Teams channels. When a turn arrives, the host dispatches it to a per-conversation Docker container the same way every other channel does today. This preserves our "no real secrets in containers" rule.

---

## Phase 1 — Identity + Observability

### Scope

- Register NinjaClaw-Nano as an Entra Agent ID via the Agent 365 CLI.
- Add an `agent365` channel in the host (Express + `@microsoft/agents-hosting`).
- Wire OTel export to Agent 365 Observability.
- Verify telemetry shows up in tenant.

### Packages to add

```json
{
  "dependencies": {
    "@microsoft/agents-hosting": "^1.2.2",
    "@microsoft/agents-activity": "^1.2.2",
    "@microsoft/agents-a365-runtime": "^0.1.0-preview",
    "@microsoft/agents-a365-observability": "^0.1.0-preview",
    "@microsoft/agents-a365-observability-hosting": "^0.1.0-preview",
    "@microsoft/agents-a365-notifications": "^0.1.0-preview"
  }
}
```

> Versions track `^0.1.0-preview.x` until GA. Pin the *exact* preview revision in `package.json` and refresh deliberately — this is preview SDK.

### New files

```
src/channels/agent365/
  index.ts            // Express receiver, JWT validation, dispatch into container
  agent.ts            // class Agent extends AgentApplication<TurnState>
  observability.ts    // token preload, OTel exporter wiring
```

### Receiver skeleton (`src/channels/agent365/index.ts`)

```ts
import {
  AuthConfiguration,
  authorizeJWT,
  CloudAdapter,
  loadAuthConfigFromEnv,
} from '@microsoft/agents-hosting';
import express from 'express';
import { agentApplication } from './agent';

const authConfig: AuthConfiguration = loadAuthConfigFromEnv();
const app = express().use(express.json());

// Health check must be BEFORE auth middleware
app.get('/api/health', (_, res) => res.status(200).json({ status: 'healthy' }));

app.use(authorizeJWT(authConfig));

app.post('/api/messages', (req, res) => {
  const adapter = agentApplication.adapter as CloudAdapter;
  adapter.process(req, res, ctx => agentApplication.run(ctx));
});

const port = Number(process.env.AGENT365_PORT) || 3978;
app.listen(port, '0.0.0.0', () => console.log(`Agent 365 channel on :${port}`));
```

### Agent class skeleton (`src/channels/agent365/agent.ts`)

```ts
import {
  AgentApplication, MemoryStorage, TurnContext, TurnState,
} from '@microsoft/agents-hosting';
import { ActivityTypes } from '@microsoft/agents-activity';
import { BaggageBuilder } from '@microsoft/agents-a365-observability';
import {
  AgenticTokenCacheInstance,
  BaggageBuilderUtils,
} from '@microsoft/agents-a365-observability-hosting';
import { getObservabilityAuthenticationScope } from '@microsoft/agents-a365-runtime';
import '@microsoft/agents-a365-notifications';

import { dispatchTurnToContainer } from '../../runtime/dispatch';

export class NinjaClawA365Agent extends AgentApplication<TurnState> {
  static authHandlerName = 'agentic';

  constructor() {
    super({
      storage: new MemoryStorage(),
      authorization: { agentic: { type: 'agentic' } },
    });

    this.onActivity(ActivityTypes.Message, (ctx, state) => this.handleMessage(ctx, state),
      [NinjaClawA365Agent.authHandlerName]);

    this.onActivity(ActivityTypes.InstallationUpdate, ctx =>
      ctx.activity.action === 'add'
        ? ctx.sendActivity('NinjaClaw-Nano hired. ⚔️')
        : ctx.sendActivity('Goodbye.'));
  }

  private async handleMessage(ctx: TurnContext, state: TurnState) {
    await this.preloadObservabilityToken(ctx);

    const baggage = BaggageBuilderUtils
      .fromTurnContext(new BaggageBuilder(), ctx)
      .sessionDescription('NinjaClaw-Nano turn')
      .build();

    await ctx.sendActivity('Got it — working on it…');
    await ctx.sendActivity({ type: 'typing' } as any);

    try {
      await baggage.run(async () => {
        const reply = await dispatchTurnToContainer({
          userText: ctx.activity.text ?? '',
          userDisplayName: ctx.activity.from?.name,
          userAadObjectId: ctx.activity.from?.aadObjectId, // <-- key for Phase 2
          channel: 'agent365',
        });
        await ctx.sendActivity(reply);
      });
    } finally {
      baggage.dispose();
    }
  }

  private async preloadObservabilityToken(ctx: TurnContext) {
    const agentId = ctx.activity?.recipient?.agenticAppId ?? '';
    const tenantId = ctx.activity?.recipient?.tenantId ?? '';
    await AgenticTokenCacheInstance.RefreshObservabilityToken(
      agentId, tenantId, ctx, this.authorization,
      getObservabilityAuthenticationScope(),
    );
  }
}

export const agentApplication = new NinjaClawA365Agent();
```

### Environment variables

| Var | Purpose |
|---|---|
| `clientId` | Entra Agent App (client) ID — populated by Agent 365 CLI provisioning |
| `tenantId` | Entra tenant ID |
| `clientSecret` *or* federated identity | Auth for the agent endpoint |
| `AGENT365_PORT` | Host port for the receiver (default `3978`) |
| `AGENT365_OBSERVABILITY_ENDPOINT` | Override only if not using default `agent365.svc.cloud.microsoft` |

`loadAuthConfigFromEnv()` from `@microsoft/agents-hosting` consumes these.

### Provisioning

Use the Agent 365 CLI (separate install) — *do not* hand-craft Entra app registrations. Order:

1. `agent365 init` — creates the blueprint scaffolding.
2. `agent365 deploy --target azure` — creates Entra Agent ID + Azure resources.
3. Save the emitted `clientId` / `tenantId` to the host `.env`.
4. `agent365 publish` — pushes the blueprint to admin center for tenant approval.

(Confirm exact CLI commands once we install it; the names above match the docs structure.)

### Acceptance criteria for Phase 1

- [ ] Agent endpoint reachable from M365 cloud (Bot Framework relay or direct, per Agent 365 docs).
- [ ] Test with **Microsoft 365 Agents Playground** (`@microsoft/m365agentsplayground`) — round-trip message works.
- [ ] OTel spans visible in the Agent 365 Observability blade for the test agent.
- [ ] Existing channels (Telegram, Web UI, Teams legacy) **unchanged and still working**. The Agent 365 channel is purely additive.
- [ ] `GITHUB_TOKEN` model auth path **unchanged**.

---

## Phase 2 — M365 Notifications + Graph data access

### Scope

- Handle `AgentNotificationActivity` (email notifications, Word comments, Teams @mentions arrive as activities).
- Add Graph tools to the container's tool set so the LLM can request reads/writes against Outlook / Teams / OneDrive / SharePoint.
- All Graph access is **on behalf of the calling user**, identified via `activity.from.aadObjectId`. Never the agent's own identity for user data.

### Auth pattern: On-Behalf-Of (OBO)

When a turn arrives, we have:

- `activity.from.aadObjectId` → the user's Entra Object ID
- The agent has an Entra Agent ID with admin-consented Graph permissions

Two viable paths:

1. **Governed MCP servers (preferred)** — admin pre-approves MCP servers (SharePoint, Mail, Calendar). The Agent 365 tooling SDK (`@microsoft/agents-a365-tooling`) handles token exchange and policy enforcement. Our container just calls the MCP server.
2. **Direct Graph with OBO token** — fallback when an MCP server doesn't exist for what we need. Agent acquires an OBO token via `this.authorization.exchangeToken(ctx, 'agentic', { scopes: [...] })` and injects it into the container as a short-lived header for that single turn.

Default to option 1; option 2 only with explicit per-tool justification because it leaks a Graph token (even short-lived) to the container.

### New tools in the container

```
brain-cli.sh        existing
graph-mail          NEW — read mail, send mail (via MCP)
graph-calendar      NEW — read events, create events
graph-files         NEW — search/read OneDrive + SharePoint
graph-teams         NEW — search Teams chats/channels
```

Each is a thin shell wrapper that POSTs to a host-side proxy, which talks to the governed MCP server. The container never holds a Graph token.

### Notification handling

```ts
this.onAgentNotification('agents:*', async (ctx, state, notif) => {
  switch (notif.notificationType) {
    case NotificationType.EmailNotification:
      await this.handleEmail(ctx, notif);
      break;
    // future: TeamsMentionNotification, WordCommentNotification, ...
    default:
      console.warn('Unhandled A365 notification:', notif.notificationType);
  }
}, 1, [NinjaClawA365Agent.authHandlerName]);
```

### Acceptance criteria for Phase 2

- [ ] Sending the agent an email triggers a turn, agent reads the email via Graph, replies via `createEmailResponseActivity`.
- [ ] @mention in Teams triggers a turn with full thread context.
- [ ] Container can answer "what did I email Bob about last week" by tool-calling the SharePoint/Mail MCP server.
- [ ] All Graph access logged with the calling user's `aadObjectId` for audit.

---

## Phase 3 — Blueprint and distribution

- Author `agent365.blueprint.json` with required permissions, MCP entitlements, DLP policy refs.
- `agent365 publish` to push to M365 Admin Center.
- Ship a one-command installer that, given a tenant + Azure subscription, provisions VM + agent identity + blueprint.

(Out of scope for this doc — separate design once Phase 2 lands.)

---

## Risks and open questions

| Risk | Mitigation |
|---|---|
| Agent 365 SDK is **preview** (`0.1.0-preview.x`). Breaking changes likely. | Pin exact preview version. Treat the channel as experimental until GA. Keep upstream legacy Teams channel running in parallel. |
| Endpoint must be reachable from M365 cloud. | Per-user VMs sit behind NAT. Need either Azure Bot Framework relay (preferred), Azure App Service, or a tunneling solution. Decide before Phase 1 starts. |
| Multiple users on one VM vs. one VM per user — affects identity model. | Today: one VM per user. Keep that. Each VM has its own agent identity. |
| Observability data leaves the VM and goes to MS-managed ingest. | Document this in the README. Provide an opt-out env var that disables the OTel exporter (still keeps notifications + identity). |
| Container OBO token in Phase 2 is a credential entering the container. | Default to MCP-server path; require explicit allowlist + short TTL for direct OBO. |
| GitHub Copilot SDK + OTel — possible double-instrumentation or noise. | Verify with a small spike before wiring globally. |

---

## Useful links

- Agent 365 developer hub: https://learn.microsoft.com/en-us/microsoft-agent-365/developer/
- Node.js SDK source: https://github.com/microsoft/Agent365-nodejs
- Reference sample (OpenAI, Node.js): https://github.com/microsoft/Agent365-Samples/tree/main/nodejs/openai/sample-agent
- Frontier preview enrollment: https://adoption.microsoft.com/copilot/frontier-program/
- Entra Agent blueprint: https://learn.microsoft.com/en-us/entra/agent-id/identity-platform/agent-blueprint
- npm packages: https://www.npmjs.com/search?q=%40microsoft%2Fagents-a365
