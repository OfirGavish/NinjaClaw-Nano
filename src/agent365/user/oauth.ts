/**
 * Agent 365 — per-instance user (delegated) OAuth.
 *
 * Sibling of admin/oauth.ts but scoped to the **active agent instance's**
 * Entra app, not the admin's bootstrap app. The user signs in once with
 * device-code; we cache their refresh token and mint short-lived Graph
 * access tokens on demand.
 *
 * Why per-instance:
 *   - Each agent instance has its own clientId/tenantId (created via
 *     admin/instances.createInstance). User consent + tokens belong to that
 *     specific app — switching active instance switches identity.
 *   - Refresh tokens are kept in `~/.config/NinjaClaw/agent365-user-{id}.json`
 *     mode 0600. Never written to .env (different lifetime, different surface).
 *
 * Scopes asked at sign-in are delegated Graph permissions only. Admin
 * consent for the underlying app permissions is handled by the admin
 * during instance creation; user consent here covers the per-user data
 * surface (their mail / calendar / OneDrive / SharePoint / chats).
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
import { getInstanceWithSecret } from '../admin/instances.js';

const CACHE_DIR = path.join(
  process.env.HOME || os.homedir(),
  '.config',
  'NinjaClaw',
);

function tokenCacheFile(instanceId: string): string {
  // Sanitize: instance ids are `inst_<ts>_<rand>` but be defensive.
  const safe = instanceId.replace(/[^A-Za-z0-9_.-]/g, '_');
  return path.join(CACHE_DIR, `agent365-user-${safe}.json`);
}

/**
 * Default delegated scopes asked at sign-in. Covers the common "personal
 * assistant" surface: read+send mail, manage calendar, read OneDrive
 * and SharePoint files/sites, read/send Teams chats and channels, plus
 * profile + offline_access for refresh tokens.
 *
 * Override per deployment with AGENT365_USER_SCOPES (comma-separated).
 */
const DEFAULT_USER_SCOPES = [
  'https://graph.microsoft.com/User.Read',
  'https://graph.microsoft.com/Mail.Read',
  'https://graph.microsoft.com/Mail.Send',
  'https://graph.microsoft.com/Calendars.ReadWrite',
  'https://graph.microsoft.com/Files.Read.All',
  'https://graph.microsoft.com/Sites.Read.All',
  'https://graph.microsoft.com/Chat.Read',
  'https://graph.microsoft.com/Chat.ReadWrite',
  'https://graph.microsoft.com/ChatMessage.Send',
  'https://graph.microsoft.com/Team.ReadBasic.All',
  'https://graph.microsoft.com/Channel.ReadBasic.All',
  'https://graph.microsoft.com/ChannelMessage.Read.All',
  'https://graph.microsoft.com/ChannelMessage.Send',
  'offline_access',
];

function userScopes(): string[] {
  const raw = process.env.AGENT365_USER_SCOPES;
  if (!raw) return DEFAULT_USER_SCOPES;
  return raw.split(',').map((s) => s.trim()).filter(Boolean);
}

interface PendingFlowState {
  flowId: string;
  instanceId: string;
  userCode: string;
  verificationUri: string;
  expiresAt: number;
  message: string;
  promise: Promise<AuthenticationResult | null>;
  status: 'pending' | 'success' | 'error';
  error?: string;
  account?: AccountInfo;
}

const pendingFlows = new Map<string, PendingFlowState>();
const pcaCache = new Map<string, PublicClientApplication>();

function getPca(instanceId: string, clientId: string, tenantId: string): PublicClientApplication {
  const key = `${instanceId}|${clientId}|${tenantId}`;
  const existing = pcaCache.get(key);
  if (existing) return existing;
  fs.mkdirSync(CACHE_DIR, { recursive: true });
  const file = tokenCacheFile(instanceId);
  const pca = new PublicClientApplication({
    auth: {
      clientId,
      authority: `https://login.microsoftonline.com/${tenantId}`,
    },
    cache: {
      cachePlugin: {
        async beforeCacheAccess(ctx: any) {
          try {
            const data = fs.readFileSync(file, 'utf-8');
            ctx.tokenCache.deserialize(data);
          } catch {
            /* missing on first run */
          }
        },
        async afterCacheAccess(ctx: any) {
          if (!ctx.cacheHasChanged) return;
          fs.writeFileSync(file, ctx.tokenCache.serialize(), { mode: 0o600 });
        },
      },
    },
  });
  pcaCache.set(key, pca);
  return pca;
}

function loadInstance(instanceId: string): { clientId: string; tenantId: string } {
  const inst = getInstanceWithSecret(instanceId);
  if (!inst) throw new Error(`unknown instance id: ${instanceId}`);
  if (!inst.clientId || !inst.tenantId) {
    throw new Error(`instance ${instanceId} missing clientId or tenantId`);
  }
  return { clientId: inst.clientId, tenantId: inst.tenantId };
}

export interface UserDeviceFlow {
  flowId: string;
  userCode: string;
  verificationUri: string;
  expiresAt: number;
  message: string;
}

export async function startUserDeviceFlow(instanceId: string): Promise<UserDeviceFlow> {
  const { clientId, tenantId } = loadInstance(instanceId);
  const flowId = `uflow_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;

  let resolveCode: (v: UserDeviceFlow) => void;
  let rejectCode: (e: Error) => void;
  const codePromise = new Promise<UserDeviceFlow>((resolve, reject) => {
    resolveCode = resolve;
    rejectCode = reject;
  });

  const request: DeviceCodeRequest = {
    scopes: userScopes(),
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

  const pca = getPca(instanceId, clientId, tenantId);
  const tokenPromise = pca
    .acquireTokenByDeviceCode(request)
    .catch((err) => {
      rejectCode(err instanceof Error ? err : new Error(String(err)));
      throw err;
    });

  const flow = await codePromise;

  const state: PendingFlowState = {
    ...flow,
    instanceId,
    promise: tokenPromise,
    status: 'pending',
  };
  pendingFlows.set(flowId, state);

  tokenPromise
    .then((result) => {
      if (result?.account) {
        state.status = 'success';
        state.account = result.account;
        logger.info(
          { instanceId, username: result.account.username },
          'Agent 365 user sign-in succeeded',
        );
      } else {
        state.status = 'error';
        state.error = 'no token returned';
      }
    })
    .catch((err) => {
      state.status = 'error';
      state.error = err?.message || String(err);
      logger.warn({ err, instanceId }, 'Agent 365 user sign-in failed');
    });

  return flow;
}

export interface UserDeviceFlowStatus {
  status: 'pending' | 'success' | 'error';
  username?: string;
  tenantId?: string;
  error?: string;
}

export function getUserDeviceFlowStatus(flowId: string): UserDeviceFlowStatus | null {
  const state = pendingFlows.get(flowId);
  if (!state) return null;
  return {
    status: state.status,
    username: state.account?.username,
    tenantId: state.account?.tenantId,
    error: state.error,
  };
}

export async function getCachedUserAccount(instanceId: string): Promise<AccountInfo | null> {
  let clientId: string;
  let tenantId: string;
  try {
    ({ clientId, tenantId } = loadInstance(instanceId));
  } catch {
    return null;
  }
  try {
    const pca = getPca(instanceId, clientId, tenantId);
    const accounts = await pca.getTokenCache().getAllAccounts();
    return accounts[0] || null;
  } catch (err) {
    logger.debug({ err, instanceId }, 'Agent 365 user: token cache read failed');
    return null;
  }
}

export async function acquireUserToken(
  instanceId: string,
  scopes: string[] = userScopes(),
): Promise<{ token: string; expiresOn: number; username: string }> {
  const { clientId, tenantId } = loadInstance(instanceId);
  const pca = getPca(instanceId, clientId, tenantId);
  const accounts = await pca.getTokenCache().getAllAccounts();
  const account = accounts[0];
  if (!account) {
    throw new Error(
      `not signed in for instance ${instanceId}: call POST /api/agent365/user/signin first`,
    );
  }
  const result = await pca.acquireTokenSilent({ account, scopes });
  if (!result?.accessToken) {
    throw new Error('silent token acquisition returned empty token');
  }
  return {
    token: result.accessToken,
    expiresOn: result.expiresOn ? result.expiresOn.getTime() : Date.now() + 3600_000,
    username: account.username,
  };
}

export async function userSignOut(instanceId: string): Promise<void> {
  let clientId: string;
  let tenantId: string;
  try {
    ({ clientId, tenantId } = loadInstance(instanceId));
  } catch {
    // Instance gone — just remove the cache file.
    try { fs.unlinkSync(tokenCacheFile(instanceId)); } catch { /* ignore */ }
    return;
  }
  const pca = getPca(instanceId, clientId, tenantId);
  const accounts = await pca.getTokenCache().getAllAccounts();
  for (const a of accounts) {
    await pca.getTokenCache().removeAccount(a);
  }
  for (const [flowId, state] of pendingFlows) {
    if (state.instanceId === instanceId) pendingFlows.delete(flowId);
  }
  try { fs.unlinkSync(tokenCacheFile(instanceId)); } catch { /* ignore */ }
}
