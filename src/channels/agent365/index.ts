/**
 * Agent 365 channel — entry point. Reads env, registers the channel, and
 * delegates the heavy lifting to sibling modules:
 *
 *   - agent.ts           — Agent365Channel class (orchestrator).
 *   - observability.ts   — A365 observability middleware + token preload.
 *   - mcp.ts             — MCP server discovery + OBO Graph fallback.
 *   - wpx-comments.ts    — Word/Excel/PowerPoint comment-reply via Graph.
 *
 * The channel is opt-in: it stays disabled unless AGENT365_CLIENT_ID is
 * set, so existing deployments are unaffected. The model auth flow
 * (GITHUB_TOKEN, etc.) inside the per-group containers is untouched.
 */

import { registerChannel } from '../registry.js';
import { readEnvFile } from '../../env.js';
import { Agent365Channel } from './agent.js';

const ENV_KEYS = [
  'AGENT365_CLIENT_ID',
  'AGENT365_TENANT_ID',
  'AGENT365_CLIENT_SECRET',
  'AGENT365_PORT',
  'AGENT365_OBSERVABILITY_ENDPOINT',
];

const envConfig = readEnvFile(ENV_KEYS);
const AGENT365_CLIENT_ID =
  process.env.AGENT365_CLIENT_ID || envConfig.AGENT365_CLIENT_ID || '';
const AGENT365_TENANT_ID =
  process.env.AGENT365_TENANT_ID || envConfig.AGENT365_TENANT_ID || '';
const AGENT365_CLIENT_SECRET =
  process.env.AGENT365_CLIENT_SECRET || envConfig.AGENT365_CLIENT_SECRET || '';
const AGENT365_PORT = parseInt(
  process.env.AGENT365_PORT || envConfig.AGENT365_PORT || '3979',
  10,
);
const AGENT365_OBSERVABILITY_ENDPOINT =
  process.env.AGENT365_OBSERVABILITY_ENDPOINT ||
  envConfig.AGENT365_OBSERVABILITY_ENDPOINT ||
  '';

registerChannel('agent365', (opts) => {
  if (!AGENT365_CLIENT_ID) return null;
  return new Agent365Channel(
    opts.onMessage,
    opts.onChatMetadata,
    opts.registeredGroups,
    {
      clientId: AGENT365_CLIENT_ID,
      tenantId: AGENT365_TENANT_ID,
      clientSecret: AGENT365_CLIENT_SECRET,
      port: AGENT365_PORT,
      observabilityEndpoint: AGENT365_OBSERVABILITY_ENDPOINT,
    },
  );
});
