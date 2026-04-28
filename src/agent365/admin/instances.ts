/**
 * Agent 365 multi-instance management.
 *
 * Why a separate store instead of just .env:
 *   - .env holds at most one set of AGENT365_* variables. We want to manage
 *     N agent instances (one per "agent persona" or per "deployment slot")
 *     under a single published blueprint.
 *   - Per-instance secrets shouldn't all live in .env where channels load
 *     them at boot. The active instance is mirrored to .env so the existing
 *     channel runtime keeps working unchanged; the rest live in this store.
 *
 * Storage:
 *   - JSON file at ~/.config/NinjaClaw/agent365-instances.json (mode 0600).
 *   - Schema is intentionally tiny — instances reference a single shared
 *     blueprint (AGENT365_BLUEPRINT_ID).
 *
 * Provisioning:
 *   - Uses the admin's MSAL token from admin/oauth.ts. Calls Microsoft
 *     Graph (Application.ReadWrite.All) to create an Entra app + client
 *     secret, then optionally creates an Azure Bot Framework resource via
 *     Azure Resource Manager.
 *   - Mirrors the same shape as scripts/agent365/01-create-entra-agent.sh
 *     and 02-create-bot-framework.sh, so behaviour stays in sync.
 */

import fs from 'fs';
import os from 'os';
import path from 'path';
import { logger } from '../../logger.js';
import { acquireAdminToken } from './oauth.js';
import { writeSettings } from '../../settings.js';

const STORE_DIR = path.join(
  process.env.HOME || os.homedir(),
  '.config',
  'NinjaClaw',
);
const STORE_FILE = path.join(STORE_DIR, 'agent365-instances.json');

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';
const GRAPH_SCOPE = 'https://graph.microsoft.com/.default';
const ARM_BASE = 'https://management.azure.com';
const ARM_SCOPE = 'https://management.azure.com/.default';

/**
 * Microsoft Graph delegated permission ids requested when creating a new
 * agent instance app. End users granting consent via device-code flow get
 * this exact set; admins can extend the list later in Entra portal.
 *
 * GUIDs are stable Microsoft-published identifiers; verify with:
 *   GET https://graph.microsoft.com/v1.0/servicePrincipals(appId='00000003-0000-0000-c000-000000000000')
 *     ?$select=oauth2PermissionScopes
 */
const GRAPH_DELEGATED_PERMISSIONS = [
  // User.Read
  { id: 'e1fe6dd8-ba31-4d61-89e7-88639da4683d', type: 'Scope' },
  // offline_access
  { id: '7427e0e9-2fba-42fe-b0c0-848c9e6a8182', type: 'Scope' },
  // Mail.Read
  { id: '570282fd-fa5c-430d-a7fd-fc8dc98a9dca', type: 'Scope' },
  // Mail.Send
  { id: 'e383f46e-2787-4529-855e-0e479a3ffac0', type: 'Scope' },
  // Calendars.ReadWrite
  { id: '1ec239c2-d7c9-4623-a91a-a9775856bb36', type: 'Scope' },
  // Files.Read.All
  { id: 'df85f4d6-205c-4ac5-a5ea-6bf408dba283', type: 'Scope' },
  // Sites.Read.All
  { id: '205e70e5-aba6-4c52-a976-6d2d46c48043', type: 'Scope' },
  // Chat.Read
  { id: 'f501c180-9344-439a-bca0-6cbf209fd270', type: 'Scope' },
  // Chat.ReadWrite
  { id: '9ff7295e-131b-4d94-90e1-69fde507ac11', type: 'Scope' },
  // ChatMessage.Send
  { id: '116b7235-7cc6-461e-b163-8e55691d839e', type: 'Scope' },
  // Team.ReadBasic.All
  { id: '485be79e-c497-4b35-9400-0e3fa7f2a5d4', type: 'Scope' },
  // Channel.ReadBasic.All
  { id: '9d8982ae-4365-4f57-95e9-d6032a4c0b87', type: 'Scope' },
  // ChannelMessage.Read.All
  { id: '767156cb-16ae-4d10-8f8b-41b657c8c8c8', type: 'Scope' },
  // ChannelMessage.Send
  { id: 'ebf0f66e-9fb1-49e4-a278-222f76911cf4', type: 'Scope' },
] as const;

export type HostingMode = 'bot-framework' | 'tailscale';

export interface AgentInstance {
  id: string;
  displayName: string;
  clientId: string;
  tenantId: string;
  objectId: string;
  /** Stored locally only; never sent to the blueprint API. */
  clientSecret: string;
  hostingMode: HostingMode;
  messagingEndpoint?: string;
  botName?: string;
  /** Azure resource id of the Bot Framework resource (if created). */
  botResourceId?: string;
  /** Blueprint id at time of creation; informational only. */
  blueprintId?: string;
  createdAt: string;
  notes?: string;
}

interface StoreFile {
  version: 1;
  activeInstanceId?: string;
  instances: AgentInstance[];
}

function readStore(): StoreFile {
  try {
    const raw = fs.readFileSync(STORE_FILE, 'utf-8');
    const parsed = JSON.parse(raw) as StoreFile;
    if (parsed.version !== 1 || !Array.isArray(parsed.instances)) {
      throw new Error('unexpected store schema');
    }
    return parsed;
  } catch (err: any) {
    if (err.code === 'ENOENT') {
      return { version: 1, instances: [] };
    }
    throw err;
  }
}

function writeStore(store: StoreFile): void {
  fs.mkdirSync(STORE_DIR, { recursive: true });
  const tmp = `${STORE_FILE}.tmp`;
  fs.writeFileSync(tmp, JSON.stringify(store, null, 2), { mode: 0o600 });
  fs.renameSync(tmp, STORE_FILE);
}

/** Public-safe view of an instance (no secret). */
export interface InstanceSummary {
  id: string;
  displayName: string;
  clientId: string;
  tenantId: string;
  objectId: string;
  hostingMode: HostingMode;
  messagingEndpoint?: string;
  botName?: string;
  botResourceId?: string;
  blueprintId?: string;
  createdAt: string;
  notes?: string;
  isActive: boolean;
  hasSecret: boolean;
}

function summarize(instance: AgentInstance, activeId?: string): InstanceSummary {
  return {
    id: instance.id,
    displayName: instance.displayName,
    clientId: instance.clientId,
    tenantId: instance.tenantId,
    objectId: instance.objectId,
    hostingMode: instance.hostingMode,
    messagingEndpoint: instance.messagingEndpoint,
    botName: instance.botName,
    botResourceId: instance.botResourceId,
    blueprintId: instance.blueprintId,
    createdAt: instance.createdAt,
    notes: instance.notes,
    isActive: instance.id === activeId,
    hasSecret: !!instance.clientSecret,
  };
}

export function listInstances(): {
  activeInstanceId?: string;
  instances: InstanceSummary[];
} {
  const store = readStore();
  return {
    activeInstanceId: store.activeInstanceId,
    instances: store.instances.map((i) => summarize(i, store.activeInstanceId)),
  };
}

export function getInstance(id: string): InstanceSummary | null {
  const store = readStore();
  const found = store.instances.find((i) => i.id === id);
  return found ? summarize(found, store.activeInstanceId) : null;
}

/**
 * Internal: full record (including clientSecret) for the user OAuth layer.
 * Never expose this over the HTTP API.
 */
export function getInstanceWithSecret(id: string): AgentInstance | null {
  const store = readStore();
  return store.instances.find((i) => i.id === id) || null;
}

export function getActiveInstanceId(): string | undefined {
  return readStore().activeInstanceId;
}

/**
 * Patch an existing instance's Entra app to enable user-delegated Graph
 * sign-in (device-code) and add the standard delegated permission set.
 *
 * Used to upgrade instances created before user-delegated support shipped.
 * Idempotent: re-running just rewrites the same fields.
 */
export async function enableUserDelegated(
  id: string,
): Promise<{ ok: true; clientId: string; tenantId: string }> {
  const store = readStore();
  const instance = store.instances.find((i) => i.id === id);
  if (!instance) throw new Error(`instance not found: ${id}`);
  await graphFetch('PATCH', `/applications/${instance.objectId}`, {
    isFallbackPublicClient: true,
    requiredResourceAccess: [
      {
        resourceAppId: '00000003-0000-0000-c000-000000000000',
        resourceAccess: GRAPH_DELEGATED_PERMISSIONS,
      },
    ],
  });
  return { ok: true, clientId: instance.clientId, tenantId: instance.tenantId };
}

/**
 * Mark an instance as active and mirror its credentials into .env so the
 * existing channel runtime (which reads AGENT365_* from process.env) picks
 * it up on the next restart.
 */
export function setActiveInstance(id: string): InstanceSummary {
  const store = readStore();
  const instance = store.instances.find((i) => i.id === id);
  if (!instance) throw new Error(`instance not found: ${id}`);

  store.activeInstanceId = id;
  writeStore(store);

  // Mirror creds into .env. Empty/clear semantics: writeSettings only
  // accepts <clear> for secrets, and skips empty strings entirely. So we
  // pass <clear> for the secret and just omit empty optional keys.
  const updates: Record<string, string> = {
    AGENT365_CLIENT_ID: instance.clientId,
    AGENT365_TENANT_ID: instance.tenantId,
    AGENT365_OBJECT_ID: instance.objectId,
    AGENT365_CLIENT_SECRET: instance.clientSecret || '<clear>',
    AGENT365_HOSTING_MODE: instance.hostingMode,
  };
  if (instance.messagingEndpoint) {
    updates.AGENT365_MESSAGING_ENDPOINT = instance.messagingEndpoint;
  }
  if (instance.botName) {
    updates.AGENT365_BOT_NAME = instance.botName;
  }
  writeSettings(updates);

  return summarize(instance, id);
}

// ---------------------------------------------------------------------------
// Provisioning helpers — admin Graph + ARM calls
// ---------------------------------------------------------------------------

async function graphFetch(
  method: string,
  url: string,
  body?: unknown,
): Promise<any> {
  const { token } = await acquireAdminToken([GRAPH_SCOPE]);
  const res = await fetch(`${GRAPH_BASE}${url}`, {
    method,
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
    },
    body: body !== undefined ? JSON.stringify(body) : undefined,
  });
  const text = await res.text();
  let json: any = undefined;
  if (text) {
    try { json = JSON.parse(text); } catch { json = text; }
  }
  if (!res.ok) {
    const detail = json?.error?.message || text || `HTTP ${res.status}`;
    throw new Error(`Graph ${method} ${url} failed: ${detail}`);
  }
  return json;
}

async function armFetch(
  method: string,
  url: string,
  body?: unknown,
): Promise<any> {
  const { token } = await acquireAdminToken([ARM_SCOPE]);
  const res = await fetch(`${ARM_BASE}${url}`, {
    method,
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
    },
    body: body !== undefined ? JSON.stringify(body) : undefined,
  });
  const text = await res.text();
  let json: any = undefined;
  if (text) {
    try { json = JSON.parse(text); } catch { json = text; }
  }
  if (!res.ok && res.status !== 404) {
    const detail = json?.error?.message || text || `HTTP ${res.status}`;
    throw new Error(`ARM ${method} ${url} failed: ${detail}`);
  }
  return { status: res.status, body: json };
}

export interface CreateInstanceOptions {
  displayName: string;
  /** When omitted, defaults to the admin's home tenant. */
  tenantId?: string;
  hostingMode: HostingMode;
  messagingEndpoint?: string;
  notes?: string;
  /** Bot Framework resource provisioning (only if hostingMode = bot-framework). */
  botFramework?: {
    subscriptionId: string;
    resourceGroup: string;
    location?: string;
    botName?: string;
  };
  /** Mark this instance active immediately (mirrors creds into .env). */
  setActive?: boolean;
}

export interface CreateInstanceResult {
  instance: InstanceSummary;
  steps: Array<{ step: string; status: 'ok' | 'skipped' | 'error'; detail?: string }>;
}

/**
 * Create a new agent instance: Entra app + secret + (optionally) Bot Framework
 * resource. Persists to local store and optionally mirrors to .env.
 */
export async function createInstance(
  opts: CreateInstanceOptions,
): Promise<CreateInstanceResult> {
  const steps: CreateInstanceResult['steps'] = [];

  // 1. Create Entra app registration via Graph.
  let app: any;
  try {
    app = await graphFetch('POST', '/applications', {
      displayName: opts.displayName,
      signInAudience: 'AzureADMyOrg',
      // Allow device-code / public-client flows so end users can sign in
      // and grant delegated Graph consent (mail/calendar/files/chat).
      isFallbackPublicClient: true,
      requiredResourceAccess: [
        {
          // Microsoft Graph
          resourceAppId: '00000003-0000-0000-c000-000000000000',
          resourceAccess: GRAPH_DELEGATED_PERMISSIONS,
        },
      ],
    });
    steps.push({ step: 'create-app', status: 'ok', detail: app.appId });
  } catch (err: any) {
    steps.push({ step: 'create-app', status: 'error', detail: err.message });
    throw err;
  }

  const clientId: string = app.appId;
  const objectId: string = app.id;

  // 2. Create the matching service principal (required for token issuance).
  try {
    await graphFetch('POST', '/servicePrincipals', { appId: clientId });
    steps.push({ step: 'create-sp', status: 'ok' });
  } catch (err: any) {
    // Tenants with auto-provisioning may already have created it.
    if (/already exists/i.test(err.message)) {
      steps.push({ step: 'create-sp', status: 'skipped', detail: 'already exists' });
    } else {
      steps.push({ step: 'create-sp', status: 'error', detail: err.message });
      // SP creation failure is non-fatal — Graph often eventually-creates one.
    }
  }

  // 3. Mint a client secret (2 years).
  let clientSecret = '';
  try {
    const exp = new Date();
    exp.setFullYear(exp.getFullYear() + 2);
    const passwordResp = await graphFetch(
      'POST',
      `/applications/${objectId}/addPassword`,
      {
        passwordCredential: {
          displayName: `ninjaclaw-${new Date().toISOString().slice(0, 10)}`,
          endDateTime: exp.toISOString(),
        },
      },
    );
    clientSecret = passwordResp.secretText;
    steps.push({ step: 'mint-secret', status: 'ok' });
  } catch (err: any) {
    steps.push({ step: 'mint-secret', status: 'error', detail: err.message });
    throw err;
  }

  const resolvedTenant = opts.tenantId || (await resolveAdminTenantId());

  let botResourceId: string | undefined;
  let botName: string | undefined;
  let messagingEndpoint = opts.messagingEndpoint;

  // 4. Create Bot Framework resource (optional).
  if (opts.hostingMode === 'bot-framework' && opts.botFramework) {
    const bf = opts.botFramework;
    botName = bf.botName || `ninjaclaw-${clientId.slice(0, 8)}`;
    const location = bf.location || 'global';
    const armUrl =
      `/subscriptions/${bf.subscriptionId}` +
      `/resourceGroups/${bf.resourceGroup}` +
      `/providers/Microsoft.BotService/botServices/${botName}` +
      `?api-version=2022-09-15`;
    if (!messagingEndpoint) {
      steps.push({
        step: 'create-bot',
        status: 'error',
        detail: 'messagingEndpoint required for bot-framework hosting',
      });
      throw new Error('messagingEndpoint required for bot-framework hosting');
    }
    try {
      const result = await armFetch('PUT', armUrl, {
        location,
        sku: { name: 'F0' },
        kind: 'azurebot',
        properties: {
          displayName: botName,
          endpoint: messagingEndpoint,
          msaAppId: clientId,
          msaAppType: 'SingleTenant',
          msaAppTenantId: resolvedTenant,
          publicNetworkAccess: 'Enabled',
        },
      });
      botResourceId = result.body?.id;
      steps.push({ step: 'create-bot', status: 'ok', detail: botResourceId });
    } catch (err: any) {
      steps.push({ step: 'create-bot', status: 'error', detail: err.message });
      throw err;
    }
  } else if (opts.hostingMode === 'bot-framework') {
    steps.push({
      step: 'create-bot',
      status: 'skipped',
      detail: 'no botFramework block supplied; create resource manually',
    });
  }

  const blueprintId = process.env.AGENT365_BLUEPRINT_ID || undefined;

  const instance: AgentInstance = {
    id: `inst_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`,
    displayName: opts.displayName,
    clientId,
    tenantId: resolvedTenant,
    objectId,
    clientSecret,
    hostingMode: opts.hostingMode,
    messagingEndpoint,
    botName,
    botResourceId,
    blueprintId,
    createdAt: new Date().toISOString(),
    notes: opts.notes,
  };

  const store = readStore();
  store.instances.push(instance);
  if (opts.setActive || !store.activeInstanceId) {
    store.activeInstanceId = instance.id;
  }
  writeStore(store);
  steps.push({ step: 'persist-store', status: 'ok' });

  if (opts.setActive || store.activeInstanceId === instance.id) {
    try {
      setActiveInstance(instance.id);
      steps.push({ step: 'mirror-env', status: 'ok' });
    } catch (err: any) {
      steps.push({ step: 'mirror-env', status: 'error', detail: err.message });
    }
  }

  return {
    instance: summarize(instance, store.activeInstanceId),
    steps,
  };
}

export interface DeleteInstanceOptions {
  /** When true, also deletes the Entra app + Bot Framework resource. */
  cascade?: boolean;
}export interface DeleteInstanceResult {
  removed: boolean;
  steps: Array<{ step: string; status: 'ok' | 'skipped' | 'error'; detail?: string }>;
}

export async function deleteInstance(
  id: string,
  opts: DeleteInstanceOptions = {},
): Promise<DeleteInstanceResult> {
  const steps: DeleteInstanceResult['steps'] = [];
  const store = readStore();
  const idx = store.instances.findIndex((i) => i.id === id);
  if (idx === -1) {
    return { removed: false, steps: [{ step: 'lookup', status: 'error', detail: 'not found' }] };
  }
  const instance = store.instances[idx];

  if (opts.cascade) {
    if (instance.botResourceId) {
      try {
        await armFetch(
          'DELETE',
          `${instance.botResourceId}?api-version=2022-09-15`,
        );
        steps.push({ step: 'delete-bot', status: 'ok' });
      } catch (err: any) {
        steps.push({ step: 'delete-bot', status: 'error', detail: err.message });
      }
    } else {
      steps.push({ step: 'delete-bot', status: 'skipped', detail: 'no bot resource id' });
    }
    try {
      await graphFetch('DELETE', `/applications/${instance.objectId}`);
      steps.push({ step: 'delete-app', status: 'ok' });
    } catch (err: any) {
      steps.push({ step: 'delete-app', status: 'error', detail: err.message });
    }
  }

  store.instances.splice(idx, 1);
  if (store.activeInstanceId === id) {
    store.activeInstanceId = store.instances[0]?.id;
  }
  writeStore(store);
  steps.push({ step: 'persist-store', status: 'ok' });

  if (store.activeInstanceId) {
    try {
      setActiveInstance(store.activeInstanceId);
      steps.push({ step: 'remirror-env', status: 'ok' });
    } catch (err: any) {
      steps.push({ step: 'remirror-env', status: 'error', detail: err.message });
    }
  }

  return { removed: true, steps };
}

async function resolveAdminTenantId(): Promise<string> {
  // Use Graph /me to discover the admin's home tenant.
  try {
    const me = await graphFetch('GET', '/organization');
    const value = me?.value?.[0];
    if (value?.id) return value.id;
  } catch (err) {
    logger.warn({ err }, 'failed to resolve admin tenant via /organization');
  }
  if (process.env.AGENT365_TENANT_ID) return process.env.AGENT365_TENANT_ID;
  throw new Error('could not determine tenantId; pass it explicitly');
}
