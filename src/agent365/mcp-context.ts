/**
 * Per-JID MCP turn context cache for the Agent 365 channel.
 *
 * When an Agent 365 turn arrives, the channel pre-discovers MCP server
 * configurations (with bearer tokens already attached by the SDK's OBO
 * exchange) and stashes them here keyed by the channel JID. NinjaClaw's
 * queue picks the message up later via processGroupMessages → runAgent →
 * runContainerAgent and pulls the cached MCP servers to inject into the
 * Copilot SDK session.
 *
 * This is intentionally a small in-process cache — Agent 365 is an opt-in
 * preview channel, and the bearer tokens it caches are short-lived (the
 * issuing tenant typically returns ~60-minute TTLs). Entries expire on
 * read; consumers degrade gracefully (no Graph tools that turn) when the
 * cache miss happens.
 */

import { logger } from '../logger.js';

/**
 * Mirror of the Copilot SDK's MCPRemoteServerConfig shape so we don't have
 * to import the SDK from the host. The container-side runner uses the real
 * SDK type.
 */
export interface CachedMcpServer {
  type: 'http' | 'sse';
  url: string;
  headers?: Record<string, string>;
  tools: string[];
}

export interface McpTurnContext {
  /** Map of MCP server name → SDK-shaped config. */
  mcpServers: Record<string, CachedMcpServer>;
  /** Wall-clock ms when this entry should be discarded. */
  expiresAt: number;
  /** Calling user's Entra Object ID, for audit logs. */
  aadObjectId?: string;
  /**
   * True when the originating activity arrived as an email notification —
   * the channel uses this to format outbound replies as
   * `createEmailResponseActivity` instead of a plain message.
   */
  fromEmail: boolean;
  /**
   * On-behalf-of access token for Microsoft Graph (`graph.microsoft.com/.default`),
   * minted via the SDK's agentic auth handler when no admin-curated MCP
   * server covers the workload. Containers can use this for direct Graph
   * calls when MCP discovery returned no usable server.
   *
   * Single-turn credential. Same TTL as the surrounding cache entry.
   */
  graphToken?: string;
  /**
   * WPX (Word/Excel/PowerPoint) comment context, when the originating
   * activity was a comment notification. Lets the channel attempt a
   * comment-thread reply via Graph instead of a chat-style fallback.
   */
  wpxComment?: WpxCommentContext;
}

export interface WpxCommentContext {
  /** App that produced the comment ("word" | "excel" | "powerpoint"). */
  workload: 'word' | 'excel' | 'powerpoint';
  /** OneDrive driveItem ID of the document. */
  documentId?: string;
  /** ID of the comment that triggered the notification. */
  initiatingCommentId?: string;
  /** Active thread ID — what we reply to. */
  subjectCommentId?: string;
}

const cache = new Map<string, McpTurnContext>();

/** Conservative default — well below typical Entra access-token TTL. */
export const DEFAULT_MCP_TTL_MS = 50 * 60 * 1000;

export function setMcpTurnContext(jid: string, ctx: McpTurnContext): void {
  cache.set(jid, ctx);
  logger.debug(
    { jid, serverCount: Object.keys(ctx.mcpServers).length, fromEmail: ctx.fromEmail },
    'Agent 365: cached MCP turn context',
  );
}

/**
 * Pop the cached MCP context for a JID. Returns undefined if missing or
 * expired. Always removes the entry — MCP server bearer tokens are
 * single-turn credentials.
 */
export function consumeMcpTurnContext(jid: string): McpTurnContext | undefined {
  const entry = cache.get(jid);
  if (!entry) return undefined;
  cache.delete(jid);
  if (Date.now() >= entry.expiresAt) {
    logger.debug({ jid }, 'Agent 365: MCP turn context expired before use');
    return undefined;
  }
  return entry;
}

/** Test-only helper. */
export function _clearMcpCache(): void {
  cache.clear();
}
