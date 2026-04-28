/**
 * Agent 365 admin OAuth — MSAL device-code flow.
 *
 * Why a separate identity from the per-instance agent app:
 *   - Different scopes (admin scopes vs agent scopes).
 *   - Different lifetime (90-day refresh vs 2-year client secret).
 *   - Different store (per-user, not per-deployment).
 *
 * The user runs `Sign in with Microsoft` from the Web UI. We hand back a
 * device code + verification URL; the user pastes the code on their phone
 * or another browser. We poll MSAL until the token arrives, then cache the
 * account so subsequent admin actions silently re-acquire tokens.
 *
 * Tokens land in `~/.config/NinjaClaw/agent365-admin.json` with mode 0600.
 * Never written to .env (different lifetime, different identity).
 */

import fs from 'fs';
import os from 'os';
import path from 'path';
import {
  PublicClientApplication,
  type AccountInfo,
  type AuthenticationResult,
  type DeviceCodeRequest,
} from '@azure/msal-node';
import { logger } from '../../logger.js';

/**
 * Default bootstrap app: the public Azure CLI app id. We piggyback on it
 * for the admin flow because (a) it already has user_impersonation against
 * the Microsoft Graph + management APIs, and (b) it removes the one-time
 * "register an app" prerequisite from the admin's setup story.
 *
 * Override with AGENT365_ADMIN_CLIENT_ID once you've registered a dedicated
 * "NinjaClaw Admin Console" app for finer-grained scope control.
 */
const DEFAULT_ADMIN_CLIENT_ID = '04b07795-8ddb-461a-bbee-02f9e1bf7b46'; // Azure CLI

const CACHE_DIR = path.join(
  process.env.HOME || os.homedir(),
  '.config',
  'NinjaClaw',
);
const TOKEN_CACHE_FILE = path.join(CACHE_DIR, 'agent365-admin.json');

/**
 * Default scopes asked for at sign-in. We request the union of what
 * blueprint publish + instance creation need; MSAL will surface the
 * consent prompt if any are missing.
 *
 * - `https://management.azure.com/user_impersonation` — Azure resource
 *   manager, used later for Bot Framework resource creation.
 * - `https://graph.microsoft.com/.default` — Graph (uses whichever delegated
 *   permissions Azure CLI is pre-consented for, including Application.ReadWrite
 *   when the admin account has it). Switching to a dedicated app id later
 *   lets us narrow this with explicit scopes.
 *
 * NOTE: the Frontier admin-plane scope (`https://agent365.svc.cloud.microsoft/.default`)
 * is not requested by default because Azure CLI's client id has no consent
 * mapping for it, which causes device-code flow to fail with `invalid_grant`.
 * Override via AGENT365_ADMIN_SCOPES once a registered admin app exists.
 *
 * Device-code flow can only request scopes for a SINGLE resource per call.
 * We bootstrap with Graph; the management.azure.com token is acquired later
 * via silent refresh (`acquireTokenSilent`) using the cached refresh token.
 */
const DEFAULT_ADMIN_SCOPES = [
  'https://graph.microsoft.com/.default',
];

export interface PendingDeviceFlow {
  flowId: string;
  userCode: string;
  verificationUri: string;
  expiresAt: number;
  message: string;
}

interface PendingFlowState extends PendingDeviceFlow {
  promise: Promise<AuthenticationResult | null>;
  status: 'pending' | 'success' | 'error';
  error?: string;
  account?: AccountInfo;
}

const pendingFlows = new Map<string, PendingFlowState>();

let pca: PublicClientApplication | null = null;

function adminClientId(): string {
  return (
    process.env.AGENT365_ADMIN_CLIENT_ID ||
    DEFAULT_ADMIN_CLIENT_ID
  );
}

function adminTenantId(): string {
  // 'organizations' = any work/school account; switch to the tenant id when
  // you want to lock the admin console to a single tenant.
  return process.env.AGENT365_ADMIN_TENANT || 'organizations';
}

function adminScopes(): string[] {
  const raw = process.env.AGENT365_ADMIN_SCOPES;
  if (!raw) return DEFAULT_ADMIN_SCOPES;
  return raw.split(',').map((s) => s.trim()).filter(Boolean);
}

function getPca(): PublicClientApplication {
  if (pca) return pca;
  fs.mkdirSync(CACHE_DIR, { recursive: true });
  pca = new PublicClientApplication({
    auth: {
      clientId: adminClientId(),
      authority: `https://login.microsoftonline.com/${adminTenantId()}`,
    },
    cache: { cachePlugin: makeFileCachePlugin() },
  });
  return pca;
}

/**
 * Disk cache plugin that persists MSAL's serialized cache (which contains
 * encrypted refresh tokens) into TOKEN_CACHE_FILE with restrictive perms.
 */
function makeFileCachePlugin() {
  return {
    async beforeCacheAccess(ctx: any) {
      try {
        const data = fs.readFileSync(TOKEN_CACHE_FILE, 'utf-8');
        ctx.tokenCache.deserialize(data);
      } catch {
        /* missing on first run */
      }
    },
    async afterCacheAccess(ctx: any) {
      if (!ctx.cacheHasChanged) return;
      const data = ctx.tokenCache.serialize();
      fs.writeFileSync(TOKEN_CACHE_FILE, data, { mode: 0o600 });
    },
  };
}

/**
 * Kick off a device-code flow. Returns immediately with the user-visible
 * code + verification URL; the caller polls getDeviceFlowStatus(flowId)
 * until status === 'success' or 'error'.
 */
export async function startDeviceFlow(): Promise<PendingDeviceFlow> {
  const flowId = `flow_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;

  // We need to capture the device-code response (URL + user code) BEFORE
  // MSAL starts polling. MSAL hands them to us via a callback inside
  // acquireTokenByDeviceCode. Use a short-lived deferred promise to
  // surface them to the API caller.
  let resolveCode: (v: PendingDeviceFlow) => void;
  let rejectCode: (e: Error) => void;
  const codePromise = new Promise<PendingDeviceFlow>((resolve, reject) => {
    resolveCode = resolve;
    rejectCode = reject;
  });

  const request: DeviceCodeRequest = {
    scopes: adminScopes(),
    deviceCodeCallback: (response) => {
      resolveCode({
        flowId,
        userCode: response.userCode,
        verificationUri: response.verificationUri,
        expiresAt: Date.now() + response.expiresIn * 1000,
        message: response.message,
      });
    },
  };

  // Kick off the long-running poll. We deliberately don't await it here —
  // the caller polls status separately.
  const tokenPromise = getPca()
    .acquireTokenByDeviceCode(request)
    .catch((err) => {
      // If we never even got a device code, surface the error to startDeviceFlow.
      rejectCode(err instanceof Error ? err : new Error(String(err)));
      throw err;
    });

  const flow = await codePromise;

  const state: PendingFlowState = {
    ...flow,
    promise: tokenPromise,
    status: 'pending',
  };
  pendingFlows.set(flowId, state);

  // Wire status updates without blocking.
  tokenPromise
    .then((result) => {
      if (result?.account) {
        state.status = 'success';
        state.account = result.account;
        logger.info(
          { username: result.account.username },
          'Agent 365 admin sign-in succeeded',
        );
      } else {
        state.status = 'error';
        state.error = 'no token returned';
      }
    })
    .catch((err) => {
      state.status = 'error';
      state.error = err?.message || String(err);
      logger.warn({ err }, 'Agent 365 admin sign-in failed');
    });

  return flow;
}

export interface DeviceFlowStatus {
  status: 'pending' | 'success' | 'error';
  username?: string;
  error?: string;
}

export function getDeviceFlowStatus(flowId: string): DeviceFlowStatus | null {
  const state = pendingFlows.get(flowId);
  if (!state) return null;
  return {
    status: state.status,
    username: state.account?.username,
    error: state.error,
  };
}

/**
 * Get the cached admin account, if any. Returns the first signed-in account
 * — we don't support multi-admin in the same install today.
 */
export async function getCachedAdminAccount(): Promise<AccountInfo | null> {
  try {
    const accounts = await getPca().getTokenCache().getAllAccounts();
    return accounts[0] || null;
  } catch (err) {
    logger.debug({ err }, 'Agent 365 admin: token cache read failed');
    return null;
  }
}

/**
 * Acquire a token for the given scopes. Tries silent first; if the cached
 * refresh token can't satisfy the request, throws — the caller is expected
 * to surface that as "please sign in again" in the UI.
 */
export async function acquireAdminToken(
  scopes: string[],
): Promise<{ token: string; expiresOn: number; username: string }> {
  const account = await getCachedAdminAccount();
  if (!account) {
    throw new Error('not signed in: call POST /api/agent365/signin first');
  }
  const result = await getPca().acquireTokenSilent({ account, scopes });
  if (!result?.accessToken) {
    throw new Error('silent token acquisition returned empty token');
  }
  return {
    token: result.accessToken,
    expiresOn: result.expiresOn ? result.expiresOn.getTime() : Date.now() + 3600_000,
    username: account.username,
  };
}

export async function signOut(): Promise<void> {
  const account = await getCachedAdminAccount();
  if (!account) return;
  await getPca().getTokenCache().removeAccount(account);
  pendingFlows.clear();
  try {
    fs.unlinkSync(TOKEN_CACHE_FILE);
  } catch {
    /* ignore */
  }
}
