<p align="center">
  <img src="assets/NinjaClaw.png" alt="NinjaClaw-Nano" width="200">
</p>

<p align="center">
  <strong>NinjaClaw-Nano</strong> — A security-focused AI agent with NinjaBrain knowledge engine.<br>
  Lightweight, container-isolated, powered by GitHub Copilot SDK.
</p>

<p align="center">
  Based on <a href="https://github.com/qwibitai/nanoclaw">NanoClaw</a> · MIT License
</p>

---

## What is NinjaClaw-Nano?

A personal AI agent that runs in Docker containers for security, communicates via Telegram and Web UI, and learns from every conversation through its built-in **NinjaBrain** knowledge engine.

### Key Differences from NanoClaw

| Feature | NanoClaw | NinjaClaw-Nano |
|---------|----------|----------------|
| **AI Provider** | Claude Agent SDK (Anthropic) | GitHub Copilot SDK (any model) |
| **Default Model** | Claude Sonnet | Claude Sonnet 4.6 via Copilot |
| **Knowledge Engine** | ❌ None | ✅ NinjaBrain (FTS5 search, compiled truth, timeline) |
| **Channels** | WhatsApp (via skills) | Telegram, Web UI, Teams (built-in) |
| **Prompt Optimization** | ❌ None | ✅ Agent Lightning APO |
| **Brain Tools** | ❌ None | ✅ 6 MCP tools (search, get, put, link, list, stats) |

## Features

- **Container Isolation** — Agents run in Docker containers with filesystem isolation. Bash access is safe because commands run inside the container, not on your host.
- **NinjaBrain** — Structured knowledge engine with compiled truth + timeline model. Stores everything the agent learns about people, projects, companies, and concepts. Auto-injects relevant context before every response.
- **GitHub Copilot SDK** — Single-model agent loop. No split-brain architecture. The same model that plans also executes tools, reads files, writes code, and verifies its output.
- **Multi-Channel** — Telegram bot, Web UI (WebSocket + REST), and Microsoft Teams.
- **Agent Lightning APO** — Automatic Prompt Optimization runs daily to improve the system prompt based on conversation quality metrics.
- **Security-First** — Read-only project mount in containers, credential injection without exposing secrets, blocked dangerous commands.

## Quick Start

```bash
git clone https://github.com/OfirGavish/NinjaClaw-Nano.git
cd NinjaClaw-Nano
npm install
npm run build
```

### Configure

Create a `.env` file:

```bash
ASSISTANT_NAME=NinjaClaw-Nano
GITHUB_TOKEN=your_github_copilot_token
COPILOT_MODEL=claude-sonnet-4.6
TELEGRAM_BOT_TOKEN=your_telegram_bot_token
WEB_PORT=8077
```

### Build the Agent Container

```bash
cd container
docker build -t ninjaclaw-agent:latest .
```

### Run

```bash
npm start
```

## NinjaBrain

NinjaBrain is a structured knowledge engine that gives the agent persistent memory across conversations.

### How It Works

- **Entity Types**: person, company, concept, project, tool
- **Slug Format**: `type/name` (e.g., `people/ofir-gavish`, `projects/maester`)
- **Compiled Truth**: Current best understanding of an entity (full rewrite on update)
- **Timeline**: Append-only log of events with dates
- **FTS5 Search**: Full-text search across all knowledge pages
- **Cross-Links**: Relationships between entities (`works_at`, `created`, `uses`)

### The Compounding Loop

On every conversation, the agent:
1. **READ** — Searches brain before answering about known entities
2. **RESPOND** — Uses brain context for informed answers
3. **WRITE** — After learning something new → `brain_put`
4. **LINK** — If two entities are related → `brain_link`

### MCP Tools (available inside the container)

| Tool | Description |
|------|-------------|
| `brain_search` | Full-text search across knowledge pages |
| `brain_get` | Get a page by slug with cross-links |
| `brain_put` | Create or update a knowledge page |
| `brain_link` | Cross-reference two pages |
| `brain_list` | List pages by entity type |
| `brain_stats` | Knowledge base statistics |

## Agent Lightning

Automatic Prompt Optimization powered by [Microsoft Agent Lightning](https://github.com/microsoft/agent-lightning).

### Setup

```bash
cd agent-lightning
bash setup-agent-lightning.sh nano
```

### How It Works

- Runs daily at 3 AM via cron
- Reads conversation traces from SQLite
- Scores conversations (positive: helpful responses; negative: errors, failures)
- Runs APO beam search with textual gradients
- Outputs an optimized system prompt for review

### Manual Run

```bash
source ~/agl-env/bin/activate
python3 optimize_prompt.py \
    --db store/messages.db \
    --prompt groups/main/COPILOT.md \
    --agent-type nano
```

## Architecture

```
Telegram/Web → SQLite → Polling Loop → Container (GitHub Copilot SDK) → Response
                                              ↕
                                        NinjaBrain DB
```

- Single Node.js process orchestrator
- Agents execute in isolated Docker containers
- NinjaBrain auto-injects context before container spawn
- IPC via filesystem between host and container
- Per-group message queue with concurrency control

## Credits

- Based on [NanoClaw](https://github.com/qwibitai/nanoclaw) by Gavriel Cohen
- Powered by [GitHub Copilot SDK](https://github.com/github/copilot-sdk)
- Prompt optimization by [Agent Lightning](https://github.com/microsoft/agent-lightning)

## License

MIT — See [LICENSE](LICENSE) for details.
