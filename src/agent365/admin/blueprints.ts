/**
 * Agent 365 admin — blueprint publish via the admin plane.
 *
 * Wraps the same POST that scripts/agent365/03-publish-blueprint.sh does,
 * but driven from the Web UI using the admin's MSAL token instead of a
 * client-credentials secret in .env.
 *
 * Flow:
 *   1. Read agent365/blueprint.json from the project root.
 *   2. Merge in deployment-specific fields (current AGENT365_* env vars).
 *   3. Acquire an admin access token for AGENT365_PUBLISH_AUDIENCE.
 *   4. POST to AGENT365_PUBLISH_API.
 *
 * Persists the resulting blueprint id to .env (AGENT365_BLUEPRINT_ID).
 */

import fs from 'fs';
import path from 'path';
import { logger } from '../../logger.js';
import { acquireAdminToken } from './oauth.js';
import { writeSettings } from '../../settings.js';

const DEFAULT_PUBLISH_API =
  'https://agent365.svc.cloud.microsoft/admin/v1/blueprints';
const DEFAULT_PUBLISH_AUDIENCE =
  'https://agent365.svc.cloud.microsoft/.default';

export interface BlueprintPublishOptions {
  /** When true, returns the prepared payload + auth context without POSTing. */
  dryRun?: boolean;
  /** Optional override of the on-disk blueprint path. */
  blueprintPath?: string;
}

export interface BlueprintPublishResult {
  status: 'dry-run' | 'success' | 'error';
  payload: Record<string, unknown>;
  publishUrl: string;
  audience: string;
  username?: string;
  blueprintId?: string;
  httpStatus?: number;
  responseBody?: unknown;
  error?: string;
}

export async function publishBlueprint(
  opts: BlueprintPublishOptions = {},
): Promise<BlueprintPublishResult> {
  const blueprintPath =
    opts.blueprintPath ||
    path.join(process.cwd(), 'agent365', 'blueprint.json');
  const publishUrl =
    process.env.AGENT365_PUBLISH_API || DEFAULT_PUBLISH_API;
  const audience =
    process.env.AGENT365_PUBLISH_AUDIENCE || DEFAULT_PUBLISH_AUDIENCE;

  let blueprint: Record<string, unknown>;
  try {
    blueprint = JSON.parse(fs.readFileSync(blueprintPath, 'utf-8'));
  } catch (err: any) {
    return {
      status: 'error',
      payload: {},
      publishUrl,
      audience,
      error: `failed to read blueprint at ${blueprintPath}: ${err?.message}`,
    };
  }

  const payload = {
    ...blueprint,
    deployment: {
      appId: process.env.AGENT365_CLIENT_ID || '',
      tenantId: process.env.AGENT365_TENANT_ID || '',
      objectId: process.env.AGENT365_OBJECT_ID || '',
      messagingEndpoint: process.env.AGENT365_MESSAGING_ENDPOINT || '',
      botName: process.env.AGENT365_BOT_NAME || undefined,
      hostingMode: process.env.AGENT365_HOSTING_MODE || 'bot-framework',
    },
  };

  if (opts.dryRun) {
    return { status: 'dry-run', payload, publishUrl, audience };
  }

  let token: { token: string; username: string };
  try {
    token = await acquireAdminToken([audience]);
  } catch (err: any) {
    return {
      status: 'error',
      payload,
      publishUrl,
      audience,
      error: err?.message || String(err),
    };
  }

  try {
    const res = await fetch(publishUrl, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token.token}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(payload),
    });
    const responseBody = await safeReadJson(res);
    const blueprintId =
      (responseBody && typeof responseBody === 'object'
        ? (responseBody as any).id || (responseBody as any).blueprintId
        : undefined) || undefined;

    if (!res.ok) {
      return {
        status: 'error',
        payload,
        publishUrl,
        audience,
        username: token.username,
        httpStatus: res.status,
        responseBody,
        error: `publish API returned HTTP ${res.status}`,
      };
    }

    if (blueprintId) {
      try {
        writeSettings({ AGENT365_BLUEPRINT_ID: blueprintId });
      } catch (err) {
        logger.warn({ err }, 'failed to persist AGENT365_BLUEPRINT_ID');
      }
    }

    return {
      status: 'success',
      payload,
      publishUrl,
      audience,
      username: token.username,
      httpStatus: res.status,
      responseBody,
      blueprintId,
    };
  } catch (err: any) {
    return {
      status: 'error',
      payload,
      publishUrl,
      audience,
      username: token.username,
      error: err?.message || String(err),
    };
  }
}

async function safeReadJson(res: Response): Promise<unknown> {
  const text = await res.text().catch(() => '');
  if (!text) return undefined;
  try {
    return JSON.parse(text);
  } catch {
    return text.slice(0, 2000);
  }
}
