/**
 * Agent 365 channel — orchestrator class.
 *
 * Responsibilities (kept thin; the heavy lifting lives in sibling modules):
 *
 *   - Construct CloudAdapter + AgentApplication using the agentic auth
 *     handler (Phase 1).
 *   - Wire observability middleware + per-turn token preload (Phase 1).
 *   - Register message + notification handlers, route everything through
 *     `routeInbound()` into NinjaClaw's existing dispatch pipeline (Phase 1+2).
 *   - At the start of every turn, discover MCP servers (Phase 2) and
 *     optionally mint an OBO Graph token as a fallback (follow-up).
 *   - On outbound `sendMessage`, prefer comment-thread replies for WPX
 *     notifications, then the email response activity for email
 *     notifications, then plain chat as the last resort (Phase 2 + follow-up).
 */

import express from 'express';
import {
  CloudAdapter,
  AgentApplication,
  MemoryStorage,
  TurnContext,
  TurnState,
  loadAuthConfigFromEnv,
  AuthConfiguration,
  authorizeJWT,
} from '@microsoft/agents-hosting';
import { ConversationReference } from '@microsoft/agents-activity';
import { McpToolServerConfigurationService } from '@microsoft/agents-a365-tooling';
// Side-effect import: module-augments AgentApplication with the
// onAgentic*Notification() methods.
import '@microsoft/agents-a365-notifications';
import {
  createEmailResponseActivity,
  NotificationType,
} from '@microsoft/agents-a365-notifications';

import { Channel, NewMessage, OnInboundMessage, OnChatMetadata, RegisteredGroup } from '../../types.js';
import { logger } from '../../logger.js';
import {
  setMcpTurnContext,
  DEFAULT_MCP_TTL_MS,
  type CachedMcpServer,
  type WpxCommentContext,
} from '../../agent365/mcp-context.js';
import { configureObservability, preloadObservabilityToken } from './observability.js';
import {
  AGENTIC_AUTH_HANDLER,
  discoverMcpServers,
  resolveGraphFallbackToken,
  isGraphFallbackEnabled,
} from './mcp.js';
import {
  extractWpxComment,
  replyAsComment,
} from './wpx-comments.js';

interface ChannelEnv {
  clientId: string;
  tenantId: string;
  clientSecret: string;
  port: number;
  observabilityEndpoint: string;
}

export class Agent365Channel implements Channel {
  name = 'agent365';
  private adapter: CloudAdapter | null = null;
  private app: AgentApplication<TurnState> | null = null;
  private mcpService: McpToolServerConfigurationService | null = null;
  private connected = false;
  private onMessage: OnInboundMessage;
  private onChatMetadata: OnChatMetadata;
  private registeredGroups: () => Record<string, RegisteredGroup>;

  // Conversation references for proactive sends (mirrors Teams channel).
  private convRefs = new Map<string, Partial<ConversationReference>>();
  // JIDs whose originating activity was an email notification.
  private emailJids = new Set<string>();
  // JIDs that arrived as a WPX comment notification, with the comment's IDs.
  private wpxJids = new Map<string, WpxCommentContext>();
  // Per-JID Graph access token from the OBO fallback path.
  private graphTokens = new Map<string, string>();

  constructor(
    onMessage: OnInboundMessage,
    onChatMetadata: OnChatMetadata,
    registeredGroups: () => Record<string, RegisteredGroup>,
    private env: ChannelEnv,
  ) {
    this.onMessage = onMessage;
    this.onChatMetadata = onChatMetadata;
    this.registeredGroups = registeredGroups;
  }

  async connect(): Promise<void> {
    if (!this.env.clientId) {
      logger.info('Agent 365: no AGENT365_CLIENT_ID, skipping');
      return;
    }

    // The hosting SDK's loadAuthConfigFromEnv() reads lower-cased env vars
    // (clientId, tenantId, clientSecret). Mirror our prefixed values into
    // process.env just-in-time so we don't collide with other Entra apps.
    if (!process.env.clientId) process.env.clientId = this.env.clientId;
    if (!process.env.tenantId && this.env.tenantId) {
      process.env.tenantId = this.env.tenantId;
    }
    if (!process.env.clientSecret && this.env.clientSecret) {
      process.env.clientSecret = this.env.clientSecret;
    }
    if (this.env.observabilityEndpoint && !process.env.A365_OBSERVABILITY_ENDPOINT) {
      process.env.A365_OBSERVABILITY_ENDPOINT = this.env.observabilityEndpoint;
    }

    let authConfig: AuthConfiguration;
    try {
      authConfig = loadAuthConfigFromEnv();
    } catch (err) {
      logger.error({ err }, 'Agent 365: failed to load auth config; channel disabled');
      return;
    }

    this.adapter = new CloudAdapter(authConfig);
    configureObservability(this.adapter);

    this.app = new AgentApplication<TurnState>({
      storage: new MemoryStorage(),
      adapter: this.adapter,
      authorization: {
        [AGENTIC_AUTH_HANDLER]: {
          type: 'agentic',
          scopes: [
            'Mail.ReadWrite',
            'Mail.Send',
            'Calendars.ReadWrite',
            'Files.ReadWrite',
            'Sites.Read.All',
            'Chat.ReadWrite',
            'ChannelMessage.Read.All',
            'User.Read',
          ],
        },
      },
    });

    this.mcpService = new McpToolServerConfigurationService();

    this.app.onTurn('beforeTurn', async (context: TurnContext, _state: TurnState) => {
      await preloadObservabilityToken(
        context,
        this.app?.authorization,
        this.env.clientId,
        this.env.tenantId,
      );
      return true;
    });

    this.app.onActivity('message', async (context: TurnContext, _state: TurnState) => {
      const activity = context.activity;
      let text = activity.text || '';
      if (!text.trim()) return;

      // Strip @mention text the same way the Teams channel does.
      for (const entity of activity.entities || []) {
        if (entity.type === 'mention') {
          const mentioned = (entity as any).mentioned;
          if (mentioned?.id === activity.recipient?.id) {
            text = text.replace((entity as any).text || '', '').trim();
          }
        }
      }

      await this.routeInbound(context, text, /* fromEmail */ false);
    });

    // Notifications. Each one routes the notification into the same
    // NinjaClaw dispatch pipeline as plain messages — the agent simply
    // sees a message in the chat. WPX notifications carry a comment
    // context for the comment-reply path.
    this.app.onAgenticEmailNotification(async (context, _state, notif) => {
      const text = this.extractNotificationText(notif) || '(empty email body)';
      await this.routeInbound(context, text, /* fromEmail */ true);
    });

    this.app.onAgenticWordNotification(async (context, _state, notif) => {
      const text = this.extractNotificationText(notif) || '(empty Word comment)';
      const comment = extractWpxComment(notif, 'word');
      await this.routeInbound(context, `[Word comment] ${text}`, false, comment);
    });

    this.app.onAgenticExcelNotification(async (context, _state, notif) => {
      const text = this.extractNotificationText(notif) || '(empty Excel comment)';
      const comment = extractWpxComment(notif, 'excel');
      await this.routeInbound(context, `[Excel comment] ${text}`, false, comment);
    });

    this.app.onAgenticPowerPointNotification(async (context, _state, notif) => {
      const text = this.extractNotificationText(notif) || '(empty PowerPoint comment)';
      const comment = extractWpxComment(notif, 'powerpoint');
      await this.routeInbound(context, `[PowerPoint comment] ${text}`, false, comment);
    });

    this.app.onLifecycleNotification(async (_context, _state, notif) => {
      logger.info(
        { notificationType: NotificationType[notif.notificationType] },
        'Agent 365: lifecycle notification received',
      );
    });

    const expressApp = express();
    expressApp.use(express.json());

    // Keep health unauthenticated so tunnels/load balancers can verify routing
    // without needing a Microsoft 365 activity token.
    expressApp.get('/api/health', (_req, res) => {
      res.status(200).json({
        status: 'healthy',
        channel: 'agent365',
        clientId: this.env.clientId,
      });
    });

    // CloudAdapter expects authorizeJWT() to populate req.user with the
    // caller identity. Without it, Agent 365/Bot Framework requests are not
    // authenticated as real platform activities.
    expressApp.use(authorizeJWT(authConfig));

    expressApp.post('/api/messages', async (req, res) => {
      if (!req.body || typeof req.body.type !== 'string') {
        res.status(400).json({ error: 'Invalid Bot Framework activity' });
        return;
      }

      try {
        await this.adapter!.process(req as any, res as any, async (context) => {
          await this.app!.run(context);
        });
      } catch (err) {
        logger.error({ err }, 'Agent 365: failed to process inbound activity');
        if (!res.headersSent) {
          res.status(500).json({ error: 'Failed to process activity' });
        }
      }
    });

    expressApp.listen(this.env.port, () => {
      logger.info(
        { port: this.env.port, clientId: this.env.clientId },
        'Agent 365 channel listening',
      );
    });

    this.connected = true;
  }

  async sendMessage(jid: string, text: string): Promise<void> {
    if (!this.adapter) return;
    const ref = this.convRefs.get(jid);
    if (!ref) {
      logger.warn({ jid }, 'No conversation reference for Agent 365 JID');
      return;
    }

    // Try comment-thread reply first if this JID arrived as a WPX comment.
    const wpxComment = this.wpxJids.get(jid);
    if (wpxComment) {
      const graphToken = this.graphTokens.get(jid);
      let replied = false;
      if (graphToken) {
        replied = await replyAsComment({ comment: wpxComment, graphToken, htmlBody: text });
      } else {
        logger.debug({ jid }, 'Agent 365: no Graph token for WPX reply, falling back');
      }
      this.wpxJids.delete(jid);
      this.graphTokens.delete(jid);
      if (replied) return;
      // Fall through to chat reply below.
    }

    // Chunk long messages, same heuristic as Teams.
    const chunks: string[] = [];
    for (let i = 0; i < text.length; i += 10000) {
      chunks.push(text.slice(i, i + 10000));
    }

    for (const chunk of chunks) {
      await this.adapter.continueConversation(
        this.env.clientId,
        ref as ConversationReference,
        async (ctx: TurnContext) => {
          if (this.emailJids.has(jid)) {
            const emailActivity = createEmailResponseActivity(chunk);
            await ctx.sendActivity(emailActivity);
          } else {
            await ctx.sendActivity(chunk);
          }
        },
      );
    }

    // Email replies are one-shot.
    if (this.emailJids.has(jid)) this.emailJids.delete(jid);
    this.graphTokens.delete(jid);
  }

  isConnected(): boolean {
    return this.connected;
  }

  ownsJid(jid: string): boolean {
    return jid.startsWith('agent365:');
  }

  async disconnect(): Promise<void> {
    this.connected = false;
  }

  // -----------------------------------------------------------------------
  // Internals
  // -----------------------------------------------------------------------

  private async routeInbound(
    context: TurnContext,
    text: string,
    fromEmail: boolean,
    wpxComment?: WpxCommentContext,
  ): Promise<void> {
    const activity = context.activity;
    const convId = activity.conversation?.id || 'unknown';
    const jid = `agent365:${convId}`;
    const senderName = activity.from?.name || 'Agent 365 User';
    const timestamp = activity.timestamp
      ? new Date(activity.timestamp).toISOString()
      : new Date().toISOString();

    logger.info({
      jid,
      channelId: activity.channelId,
      fromId: activity.from?.id,
      fromName: activity.from?.name,
      fromAadObjectId: activity.from?.aadObjectId,
      conversationTenantId: (activity.conversation as any)?.tenantId,
      activityTenantId: (activity as any).channelData?.tenant?.id,
      recipientId: activity.recipient?.id,
      serviceUrl: activity.serviceUrl,
    }, 'Agent 365: inbound activity identity');

    try {
      const ref = activity.getConversationReference();
      this.convRefs.set(jid, ref);
    } catch (err) {
      logger.debug({ err, jid }, 'Agent 365: could not capture conversation reference');
    }

    if (fromEmail) this.emailJids.add(jid);
    if (wpxComment) this.wpxJids.set(jid, wpxComment);

    const mcpServers = await this.safeDiscoverMcp(context, jid);

    // Mint OBO Graph token for two cases: (1) WPX comment reply path needs
    // it to call Graph; (2) opt-in fallback so containers without an MCP
    // server can still reach Graph directly.
    const needGraphToken =
      !!wpxComment ||
      (isGraphFallbackEnabled() && Object.keys(mcpServers).length === 0);
    const graphToken = needGraphToken
      ? await resolveGraphFallbackToken(context, this.app?.authorization)
      : undefined;
    if (graphToken) this.graphTokens.set(jid, graphToken);

    setMcpTurnContext(jid, {
      mcpServers,
      expiresAt: Date.now() + DEFAULT_MCP_TTL_MS,
      aadObjectId: activity.from?.aadObjectId,
      fromEmail,
      graphToken,
      wpxComment,
    });

    this.onChatMetadata(jid, timestamp, senderName, 'agent365');

    const msg: NewMessage = {
      id: activity.id || Date.now().toString(),
      chat_jid: jid,
      sender: `agent365:${activity.from?.id || 'unknown'}`,
      sender_name: senderName,
      content: text,
      timestamp,
      is_from_me: false,
    };

    this.onMessage(jid, msg);
  }

  private async safeDiscoverMcp(
    context: TurnContext,
    jid: string,
  ): Promise<Record<string, CachedMcpServer>> {
    if (!this.mcpService) return {};
    try {
      return await discoverMcpServers(context, this.app?.authorization, this.mcpService);
    } catch (err) {
      logger.warn({ err, jid }, 'Agent 365: MCP discovery failed');
      return {};
    }
  }

  private extractNotificationText(notif: any): string {
    if (notif.text && notif.text.trim()) return String(notif.text).trim();
    const html = notif.emailNotification?.htmlBody;
    if (html) {
      // Crude HTML → text. Good enough for routing into the LLM, which can
      // handle imperfect text. Avoids pulling in a full HTML parser.
      return String(html)
        .replace(/<style[\s\S]*?<\/style>/gi, ' ')
        .replace(/<script[\s\S]*?<\/script>/gi, ' ')
        .replace(/<[^>]+>/g, ' ')
        .replace(/&nbsp;/g, ' ')
        .replace(/\s+/g, ' ')
        .trim();
    }
    return '';
  }
}
