/**
 * NinjaClaw — Microsoft Teams channel for NinjaClaw.
 *
 * Implements the Channel interface using Bot Framework SDK.
 * Uses proactive messaging: returns HTTP 200 immediately, processes in background.
 * Same pattern as the Python NinjaClaw teams_bot.py.
 */

import express from 'express';
import {
  BotFrameworkAdapter,
  type BotFrameworkAdapterSettings,
  TurnContext,
  ActivityHandler,
  ConversationReference,
} from 'botbuilder';
import { registerChannel } from './registry.js';
import { Channel, NewMessage, OnInboundMessage, OnChatMetadata, RegisteredGroup } from '../types.js';
import { logger } from '../logger.js';
import { readEnvFile } from '../env.js';

const envConfig = readEnvFile(['TEAMS_BOT_APP_ID', 'TEAMS_BOT_APP_PASSWORD', 'TEAMS_BOT_TENANT_ID', 'TEAMS_BOT_PORT']);
const TEAMS_BOT_APP_ID = process.env.TEAMS_BOT_APP_ID || envConfig.TEAMS_BOT_APP_ID || '';
const TEAMS_BOT_APP_PASSWORD = process.env.TEAMS_BOT_APP_PASSWORD || envConfig.TEAMS_BOT_APP_PASSWORD || '';
const TEAMS_BOT_TENANT_ID = process.env.TEAMS_BOT_TENANT_ID || envConfig.TEAMS_BOT_TENANT_ID || '';
const TEAMS_BOT_PORT = parseInt(process.env.TEAMS_BOT_PORT || envConfig.TEAMS_BOT_PORT || '3978', 10);

class TeamsChannel implements Channel {
  name = 'teams';
  private adapter: BotFrameworkAdapter | null = null;
  private connected = false;
  private onMessage: OnInboundMessage;
  private onChatMetadata: OnChatMetadata;
  private registeredGroups: () => Record<string, RegisteredGroup>;
  // Store conversation references for proactive messaging
  private convRefs = new Map<string, Partial<ConversationReference>>();

  constructor(
    onMessage: OnInboundMessage,
    onChatMetadata: OnChatMetadata,
    registeredGroups: () => Record<string, RegisteredGroup>,
  ) {
    this.onMessage = onMessage;
    this.onChatMetadata = onChatMetadata;
    this.registeredGroups = registeredGroups;
  }

  async connect(): Promise<void> {
    if (!TEAMS_BOT_APP_ID || !TEAMS_BOT_APP_PASSWORD) {
      logger.info('Teams: no TEAMS_BOT_APP_ID/PASSWORD, skipping');
      return;
    }

    const settings: BotFrameworkAdapterSettings = {
      appId: TEAMS_BOT_APP_ID,
      appPassword: TEAMS_BOT_APP_PASSWORD,
      channelAuthTenant: TEAMS_BOT_TENANT_ID || undefined,
    };
    this.adapter = new BotFrameworkAdapter(settings);

    const app = express();
    app.use(express.json());

    app.post('/api/messages', (req, res) => {
      this.adapter!.process(req as any, res as any, async (context) => {
        if (context.activity.type === 'message') {
          let text = context.activity.text || '';
          if (!text.trim()) return;

          // Remove @mention
          for (const entity of context.activity.entities || []) {
            if (entity.type === 'mention') {
              const mentioned = (entity as any).mentioned;
              if (mentioned?.id === context.activity.recipient.id) {
                text = text.replace((entity as any).text || '', '').trim();
              }
            }
          }

          const convId = context.activity.conversation.id;
          const jid = `teams:${convId}`;
          const senderName = context.activity.from.name || 'Teams User';
          const timestamp = context.activity.timestamp?.toISOString() || new Date().toISOString();

          // Save conversation reference for proactive messaging
          const ref = TurnContext.getConversationReference(context.activity);
          this.convRefs.set(jid, ref);

          this.onChatMetadata(jid, timestamp, senderName, 'teams');

          const msg: NewMessage = {
            id: context.activity.id || Date.now().toString(),
            chat_jid: jid,
            sender: `teams:${context.activity.from.id}`,
            sender_name: senderName,
            content: text,
            timestamp,
            is_from_me: false,
          };

          // Send typing indicator
          await context.sendActivity({ type: 'typing' });

          this.onMessage(jid, msg);
        }
      });
    });

    app.listen(TEAMS_BOT_PORT, () => {
      logger.info({ port: TEAMS_BOT_PORT }, 'Teams channel listening');
    });

    this.connected = true;
  }

  async sendMessage(jid: string, text: string): Promise<void> {
    if (!this.adapter) return;
    const ref = this.convRefs.get(jid);
    if (!ref) {
      logger.warn({ jid }, 'No conversation reference for Teams JID');
      return;
    }

    // Split long messages (Teams ~28K limit, but keep chunks reasonable)
    const chunks = [];
    for (let i = 0; i < text.length; i += 10000) {
      chunks.push(text.slice(i, i + 10000));
    }

    for (const chunk of chunks) {
      await this.adapter.continueConversation(ref as ConversationReference, async (ctx) => {
        await ctx.sendActivity(chunk);
      });
    }
  }

  isConnected(): boolean {
    return this.connected;
  }

  ownsJid(jid: string): boolean {
    return jid.startsWith('teams:');
  }

  async setTyping(jid: string): Promise<void> {
    if (!this.adapter) return;
    const ref = this.convRefs.get(jid);
    if (!ref) return;
    await this.adapter.continueConversation(ref as ConversationReference, async (ctx) => {
      await ctx.sendActivity({ type: 'typing' });
    });
  }

  async disconnect(): Promise<void> {
    this.connected = false;
  }
}

registerChannel('teams', (opts) => {
  if (!TEAMS_BOT_APP_ID) return null;
  return new TeamsChannel(opts.onMessage, opts.onChatMetadata, opts.registeredGroups);
});
