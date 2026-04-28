/**
 * Settings module — read/write configuration values to the project .env file
 * from the web UI.
 *
 * Design constraints:
 *   - Source of truth is `.env` at the project root. We do NOT mutate
 *     process.env at runtime; channels reload values when restarted.
 *     (Restart prompt is shown in the UI.)
 *   - Secrets are returned masked (`••••••last4`) so the UI can show "set"
 *     vs "empty" without leaking the value back.
 *   - Writes are atomic: write to `.env.tmp` then rename. Existing comments
 *     and unrelated keys are preserved.
 *   - Only keys in the SETTINGS_SCHEMA whitelist can be read or written.
 *     This prevents the UI from being used to inject arbitrary env vars.
 */

import fs from 'fs';
import path from 'path';
import { logger } from './logger.js';

export type SettingsCategory =
  | 'github'
  | 'web'
  | 'telegram'
  | 'teams'
  | 'agent365'
  | 'onecli'
  | 'general';

export type SettingType = 'string' | 'secret' | 'number' | 'boolean' | 'url';

export interface SettingsField {
  key: string;
  label: string;
  description: string;
  type: SettingType;
  category: SettingsCategory;
  /** Hint shown to the user; not validated at write time. */
  placeholder?: string;
  /** When true, value is masked on read and never echoed back. */
  secret?: boolean;
  /** Optional grouping inside a category (e.g. "Bot Framework relay"). */
  group?: string;
}

/**
 * Whitelist of every setting the web UI is allowed to touch.
 * Anything not in this list is rejected by the API.
 */
export const SETTINGS_SCHEMA: SettingsField[] = [
  // GitHub Models / Copilot
  {
    key: 'GITHUB_TOKEN',
    label: 'GitHub Token',
    description:
      'Personal access token used by the agent runner to call GitHub Models. Must include `models:read` scope.',
    type: 'secret',
    category: 'github',
    secret: true,
    placeholder: 'github_pat_...',
  },

  // General
  {
    key: 'ASSISTANT_NAME',
    label: 'Assistant Name',
    description: 'Trigger word and display name (e.g. "Andy" → @Andy).',
    type: 'string',
    category: 'general',
    placeholder: 'Andy',
  },
  {
    key: 'TZ',
    label: 'Timezone',
    description: 'IANA timezone for scheduled tasks (e.g. America/New_York).',
    type: 'string',
    category: 'general',
    placeholder: 'UTC',
  },
  {
    key: 'MAX_CONCURRENT_CONTAINERS',
    label: 'Max Concurrent Containers',
    description: 'Upper bound on simultaneous agent containers.',
    type: 'number',
    category: 'general',
    placeholder: '5',
  },

  // Web UI
  {
    key: 'NINJACLAW_WEB_PORT',
    label: 'Web UI Port',
    description: 'TCP port the web UI listens on. Restart required.',
    type: 'number',
    category: 'web',
    placeholder: '8484',
  },
  {
    key: 'NINJACLAW_WEB_TOKEN',
    label: 'Web UI Access Token',
    description:
      'Bearer token required to access this UI and the WebSocket. Leave empty for dev mode (open access on localhost).',
    type: 'secret',
    category: 'web',
    secret: true,
  },
  {
    key: 'NINJACLAW_WEB_USER_NAME',
    label: 'Web UI Display Name',
    description: 'Name shown for messages sent from this UI.',
    type: 'string',
    category: 'web',
    placeholder: 'Ofir Gavish',
  },

  // Telegram
  {
    key: 'TELEGRAM_BOT_TOKEN',
    label: 'Telegram Bot Token',
    description: 'Token from @BotFather. Leave empty to disable Telegram.',
    type: 'secret',
    category: 'telegram',
    secret: true,
    placeholder: '0000000000:AA...',
  },
  {
    key: 'TELEGRAM_ALLOWED_CHAT_IDS',
    label: 'Allowed Chat IDs',
    description: 'Comma-separated Telegram chat IDs allowed to message the bot.',
    type: 'string',
    category: 'telegram',
    placeholder: '123456789,987654321',
  },

  // Teams (legacy Bot Framework)
  {
    key: 'TEAMS_APP_ID',
    label: 'Teams App ID',
    description: 'Microsoft App ID for the legacy Teams bot.',
    type: 'string',
    category: 'teams',
    placeholder: '00000000-0000-0000-0000-000000000000',
  },
  {
    key: 'TEAMS_APP_PASSWORD',
    label: 'Teams App Password',
    description: 'Bot Framework client secret.',
    type: 'secret',
    category: 'teams',
    secret: true,
  },

  // OneCLI credential gateway
  {
    key: 'ONECLI_URL',
    label: 'OneCLI URL',
    description: 'Base URL of the OneCLI credential gateway.',
    type: 'url',
    category: 'onecli',
    placeholder: 'https://onecli.example.com',
  },
  {
    key: 'ONECLI_API_KEY',
    label: 'OneCLI API Key',
    description: 'API key for the OneCLI gateway.',
    type: 'secret',
    category: 'onecli',
    secret: true,
  },

  // Agent 365 — identity
  {
    key: 'AGENT365_CLIENT_ID',
    label: 'Entra Client ID',
    description: 'App registration client ID. Leave empty to disable Agent 365.',
    type: 'string',
    category: 'agent365',
    group: 'Identity',
    placeholder: '00000000-0000-0000-0000-000000000000',
  },
  {
    key: 'AGENT365_TENANT_ID',
    label: 'Entra Tenant ID',
    description: 'Microsoft 365 tenant ID.',
    type: 'string',
    category: 'agent365',
    group: 'Identity',
  },
  {
    key: 'AGENT365_CLIENT_SECRET',
    label: 'Entra Client Secret',
    description: 'Secret for the app registration.',
    type: 'secret',
    category: 'agent365',
    group: 'Identity',
    secret: true,
  },
  {
    key: 'AGENT365_OBJECT_ID',
    label: 'App Object ID',
    description: 'Object ID of the service principal (set by 01-create-entra-agent.sh).',
    type: 'string',
    category: 'agent365',
    group: 'Identity',
  },
  {
    key: 'AGENT365_PORT',
    label: 'Listener Port',
    description: 'Local TCP port for the /api/messages endpoint.',
    type: 'number',
    category: 'agent365',
    group: 'Identity',
    placeholder: '3979',
  },

  // Agent 365 — hosting
  {
    key: 'AGENT365_HOSTING_MODE',
    label: 'Hosting Mode',
    description: '`bot-framework` (recommended) or `tailscale`.',
    type: 'string',
    category: 'agent365',
    group: 'Hosting',
  },
  {
    key: 'AGENT365_MESSAGING_ENDPOINT',
    label: 'Messaging Endpoint',
    description: 'Public HTTPS URL of /api/messages.',
    type: 'url',
    category: 'agent365',
    group: 'Hosting',
    placeholder: 'https://my-host.example.com/api/messages',
  },
  {
    key: 'AGENT365_BOT_NAME',
    label: 'Azure Bot Name',
    description: 'Name of the Microsoft.BotService resource (bot-framework mode).',
    type: 'string',
    category: 'agent365',
    group: 'Hosting',
  },
  {
    key: 'AGENT365_TAILSCALE_FQDN',
    label: 'Tailscale FQDN',
    description: 'Tailnet hostname (tailscale mode).',
    type: 'string',
    category: 'agent365',
    group: 'Hosting',
  },

  // Agent 365 — admin publish
  {
    key: 'AGENT365_PUBLISH_API',
    label: 'Publish API URL',
    description: 'Admin-plane endpoint for blueprint publish (preview, internal).',
    type: 'url',
    category: 'agent365',
    group: 'Admin Publish',
  },
  {
    key: 'AGENT365_PUBLISH_AUDIENCE',
    label: 'Publish Token Audience',
    description: 'OAuth scope for the publish API.',
    type: 'string',
    category: 'agent365',
    group: 'Admin Publish',
  },
  {
    key: 'AGENT365_OBSERVABILITY_ENDPOINT',
    label: 'Observability Endpoint',
    description: 'Optional override for the OTLP ingestion URL.',
    type: 'url',
    category: 'agent365',
    group: 'Admin Publish',
  },
];

const SCHEMA_BY_KEY: Map<string, SettingsField> = new Map(
  SETTINGS_SCHEMA.map((f) => [f.key, f]),
);

export interface SettingValueOut {
  key: string;
  /** Present (and full value) for non-secret keys. */
  value?: string;
  /** Present for secret keys when set: "••••••" + last 4 chars. */
  masked?: string;
  /** True when the key has a non-empty value in .env. */
  isSet: boolean;
}

export interface SettingsResponse {
  schema: SettingsField[];
  values: SettingValueOut[];
  envPath: string;
}

function envPath(): string {
  return path.join(process.cwd(), '.env');
}

function maskSecret(value: string): string {
  if (!value) return '';
  if (value.length <= 4) return '••••';
  return `••••••${value.slice(-4)}`;
}

/**
 * Read all whitelisted settings from .env. Secrets are masked.
 */
export function readSettings(): SettingsResponse {
  const file = envPath();
  let content = '';
  try {
    content = fs.readFileSync(file, 'utf-8');
  } catch {
    // File missing — return empty values; UI will let user create it.
  }

  const raw: Record<string, string> = {};
  for (const line of content.split('\n')) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith('#')) continue;
    const eqIdx = trimmed.indexOf('=');
    if (eqIdx === -1) continue;
    const key = trimmed.slice(0, eqIdx).trim();
    if (!SCHEMA_BY_KEY.has(key)) continue;
    let value = trimmed.slice(eqIdx + 1).trim();
    if (
      value.length >= 2 &&
      ((value.startsWith('"') && value.endsWith('"')) ||
        (value.startsWith("'") && value.endsWith("'")))
    ) {
      value = value.slice(1, -1);
    }
    raw[key] = value;
  }

  const values: SettingValueOut[] = SETTINGS_SCHEMA.map((field) => {
    const v = raw[field.key];
    const isSet = !!v;
    if (field.secret) {
      return { key: field.key, masked: isSet ? maskSecret(v) : '', isSet };
    }
    return { key: field.key, value: v ?? '', isSet };
  });

  return { schema: SETTINGS_SCHEMA, values, envPath: file };
}

export interface WriteResult {
  written: string[];
  skipped: string[];
  unknown: string[];
}

/**
 * Atomically merge updates into .env.
 *
 * - Only whitelisted keys are accepted; everything else goes into `unknown`.
 * - Empty string for a non-secret value clears the key (writes empty value).
 * - Empty string for a secret means "no change" (so we don't accidentally
 *   wipe the secret when the UI sends back a masked placeholder).
 * - Existing comments and unrelated keys are preserved verbatim.
 */
export function writeSettings(
  updates: Record<string, string>,
): WriteResult {
  const file = envPath();
  const result: WriteResult = { written: [], skipped: [], unknown: [] };

  let existing = '';
  try {
    existing = fs.readFileSync(file, 'utf-8');
  } catch {
    // File missing — start with empty content. We'll create it.
  }

  // Normalize updates: filter to whitelist + secret rules.
  const accepted: Map<string, string> = new Map();
  for (const [key, value] of Object.entries(updates)) {
    const field = SCHEMA_BY_KEY.get(key);
    if (!field) {
      result.unknown.push(key);
      continue;
    }
    // Reject anything that smells like the masked placeholder being echoed back.
    if (field.secret && value.includes('••')) {
      result.skipped.push(key);
      continue;
    }
    if (field.secret && value === '') {
      // Don't clear secrets on empty input — UI sends "" when the user didn't
      // touch the field. To explicitly clear, the user types "<clear>".
      result.skipped.push(key);
      continue;
    }
    let normalized = value;
    if (field.secret && value === '<clear>') normalized = '';
    accepted.set(key, normalized);
  }

  // Walk the existing file line by line, replacing matching keys in place.
  const lines = existing.split('\n');
  const seen = new Set<string>();
  const updated: string[] = [];
  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith('#')) {
      updated.push(line);
      continue;
    }
    const eqIdx = trimmed.indexOf('=');
    if (eqIdx === -1) {
      updated.push(line);
      continue;
    }
    const key = trimmed.slice(0, eqIdx).trim();
    if (accepted.has(key)) {
      updated.push(formatLine(key, accepted.get(key)!));
      seen.add(key);
      result.written.push(key);
    } else {
      updated.push(line);
    }
  }

  // Append any keys that weren't already in the file.
  const additions: string[] = [];
  for (const [key, value] of accepted.entries()) {
    if (seen.has(key)) continue;
    additions.push(formatLine(key, value));
    result.written.push(key);
  }
  let finalContent = updated.join('\n');
  if (additions.length) {
    if (finalContent && !finalContent.endsWith('\n')) finalContent += '\n';
    if (!finalContent.endsWith('\n\n')) finalContent += '\n';
    finalContent += '# Added by web UI\n';
    finalContent += additions.join('\n') + '\n';
  }

  // Atomic write via temp + rename.
  const tmp = `${file}.tmp-${process.pid}-${Date.now()}`;
  fs.writeFileSync(tmp, finalContent, { mode: 0o600 });
  // Match the original file's mode if it exists.
  try {
    const st = fs.statSync(file);
    fs.chmodSync(tmp, st.mode);
  } catch {
    /* new file, keep 0600 */
  }
  fs.renameSync(tmp, file);

  logger.info(
    { written: result.written, skipped: result.skipped.length, unknown: result.unknown.length },
    'settings updated via web UI',
  );

  return result;
}

function formatLine(key: string, value: string): string {
  // Quote values that contain whitespace or shell-special chars.
  if (/[\s"'#$`\\]/.test(value)) {
    const escaped = value.replace(/\\/g, '\\\\').replace(/"/g, '\\"');
    return `${key}="${escaped}"`;
  }
  return `${key}=${value}`;
}
