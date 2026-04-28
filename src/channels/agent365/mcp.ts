/**
 * Agent 365 channel — MCP discovery + OBO Graph fallback.
 *
 * Two responsibilities:
 *
 *   1. discoverMcpServers(context, auth, mcpService) — calls the
 *      McpToolServerConfigurationService to list admin-curated MCP servers
 *      this user is authorized to call, and returns them in the SDK-shaped
 *      record consumed by the per-turn cache.
 *
 *   2. resolveGraphFallbackToken(context, auth) — when MCP discovery
 *      returns nothing (or admin policy doesn't curate a server for some
 *      workload), do a direct OBO exchange for `graph.microsoft.com/.default`
 *      using the agentic auth handler so containers can call Graph directly.
 *      Opt-in via AGENT365_GRAPH_FALLBACK=true.
 */

import {
  TurnContext,
  type Authorization,
} from '@microsoft/agents-hosting';
import { McpToolServerConfigurationService } from '@microsoft/agents-a365-tooling';
import { logger } from '../../logger.js';
import type { CachedMcpServer } from '../../agent365/mcp-context.js';

const GRAPH_SCOPE = 'https://graph.microsoft.com/.default';

export const AGENTIC_AUTH_HANDLER = 'agentic';

export async function discoverMcpServers(
  context: TurnContext,
  auth: Authorization | undefined,
  mcpService: McpToolServerConfigurationService,
): Promise<Record<string, CachedMcpServer>> {
  if (!auth) return {};
  const servers = await mcpService.listToolServers(
    context,
    auth,
    AGENTIC_AUTH_HANDLER,
  );
  const out: Record<string, CachedMcpServer> = {};
  for (const srv of servers) {
    if (!srv.url) continue;
    out[srv.mcpServerName] = {
      type: 'http',
      url: srv.url,
      headers: srv.headers,
      // The tooling gateway already filters tools per admin policy; expose
      // them all to the model and let the gateway enforce permissions.
      tools: ['*'],
    };
  }
  return out;
}

/**
 * Mint a Microsoft Graph access token via the SDK's agentic OBO exchange.
 * Used when admin policy hasn't curated an MCP server covering the
 * workload the user is asking about and the deployment opts in to direct
 * Graph fallback. Returns undefined on any failure; callers degrade
 * gracefully (no Graph access that turn).
 */
export async function resolveGraphFallbackToken(
  context: TurnContext,
  auth: Authorization | undefined,
): Promise<string | undefined> {
  if (!auth) return undefined;
  try {
    // Newer SDK signature: (context, authHandlerId, options).
    // We pass scopes explicitly so this works whether or not the agentic
    // handler has a default scope configured.
    const tokenResponse = await auth.exchangeToken(
      context,
      AGENTIC_AUTH_HANDLER,
      { scopes: [GRAPH_SCOPE] },
    );
    if (tokenResponse?.token) return tokenResponse.token;
    logger.debug('Agent 365: OBO Graph exchange returned empty token');
    return undefined;
  } catch (err) {
    logger.debug({ err }, 'Agent 365: OBO Graph fallback failed');
    return undefined;
  }
}

/**
 * Returns true when the Graph-fallback path should be tried. Today this
 * is purely opt-in via env to keep the default conservative.
 */
export function isGraphFallbackEnabled(): boolean {
  return (
    (process.env.AGENT365_GRAPH_FALLBACK || '').toLowerCase() === 'true'
  );
}
