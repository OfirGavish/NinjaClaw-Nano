/**
 * NinjaClaw — Telegram channel for NinjaClaw.
 *
 * Implements the Channel interface using grammY.
 * Messages flow: Telegram → this channel → NinjaClaw orchestrator → container → response → Telegram
 */

import { Bot } from 'grammy';
import { registerChannel } from './registry.js';
import { Channel, NewMessage, OnInboundMessage, OnChatMetadata, RegisteredGroup } from '../types.js';
import { logger } from '../logger.js';
import { readEnvFile } from '../env.js';

const envConfig = readEnvFile(['TELEGRAM_BOT_TOKEN', 'ALLOWED_TELEGRAM_IDS']);
const TELEGRAM_BOT_TOKEN = process.env.TELEGRAM_BOT_TOKEN || envConfig.TELEGRAM_BOT_TOKEN || '';
const ALLOWED_TELEGRAM_IDS = new Set(
  (process.env.ALLOWED_TELEGRAM_IDS || envConfig.ALLOWED_TELEGRAM_IDS || '')
    .split(',')
    .filter(Boolean)
    .map(Number),
);

function isAllowed(userId: number): boolean {
  if (ALLOWED_TELEGRAM_IDS.size === 0) return true;
  return ALLOWED_TELEGRAM_IDS.has(userId);
}

class TelegramChannel implements Channel {
  name = 'telegram';
  private bot: Bot | null = null;
  private connected = false;
  private onMessage: OnInboundMessage;
  private onChatMetadata: OnChatMetadata;
  private registeredGroups: () => Record<string, RegisteredGroup>;

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
    if (!TELEGRAM_BOT_TOKEN) {
      logger.info('Telegram: no TELEGRAM_BOT_TOKEN, skipping');
      return;
    }

    this.bot = new Bot(TELEGRAM_BOT_TOKEN);

    this.bot.on('message:text', (ctx) => {
      const userId = ctx.from?.id ?? 0;
      if (!isAllowed(userId)) return;

      const chatId = ctx.chat.id.toString();
      const jid = `telegram:${chatId}`;
      const senderName = [ctx.from?.first_name, ctx.from?.last_name].filter(Boolean).join(' ') || 'User';
      const text = ctx.message.text;
      const timestamp = new Date(ctx.message.date * 1000).toISOString();

      // Report chat metadata
      const chatName = ctx.chat.type === 'private'
        ? senderName
        : (ctx.chat as any).title || `Chat ${chatId}`;
      this.onChatMetadata(jid, timestamp, chatName, 'telegram', ctx.chat.type !== 'private');

      // Build NinjaClaw message
      const msg: NewMessage = {
        id: ctx.message.message_id.toString(),
        chat_jid: jid,
        sender: `telegram:${userId}`,
        sender_name: senderName,
        content: text,
        timestamp,
        is_from_me: false,
      };

      this.onMessage(jid, msg);
    });

    this.bot.start();
    this.connected = true;
    logger.info('Telegram channel connected');
  }

  async sendMessage(jid: string, text: string): Promise<void> {
    if (!this.bot) return;
    const chatId = jid.replace('telegram:', '');

    // Split long messages (Telegram 4096 char limit)
    const chunks = [];
    for (let i = 0; i < text.length; i += 4000) {
      chunks.push(text.slice(i, i + 4000));
    }

    for (const chunk of chunks) {
      try {
        await this.bot.api.sendMessage(chatId, chunk, { parse_mode: 'Markdown' });
      } catch {
        // Fallback without Markdown
        await this.bot.api.sendMessage(chatId, chunk).catch((err) => {
          logger.error({ err, chatId }, 'Failed to send Telegram message');
        });
      }
    }
  }

  isConnected(): boolean {
    return this.connected;
  }

  ownsJid(jid: string): boolean {
    return jid.startsWith('telegram:');
  }

  async setTyping(jid: string): Promise<void> {
    if (!this.bot) return;
    const chatId = jid.replace('telegram:', '');
    await this.bot.api.sendChatAction(chatId, 'typing').catch(() => {});
  }

  async disconnect(): Promise<void> {
    this.bot?.stop();
    this.connected = false;
  }
}

// Self-register with NinjaClaw's channel registry
registerChannel('telegram', (opts) => {
  if (!TELEGRAM_BOT_TOKEN) return null;
  return new TelegramChannel(opts.onMessage, opts.onChatMetadata, opts.registeredGroups);
});
