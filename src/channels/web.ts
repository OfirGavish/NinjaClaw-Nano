/**
 * NinjaClaw — Web Chat channel for NinjaClaw.
 *
 * Implements the Channel interface using Express + WebSocket.
 * Provides the same web UI as the Python NinjaClaw web_api.py.
 */

import express from 'express';
import { WebSocketServer } from 'ws';
import type WebSocket from 'ws';
import { createServer } from 'http';
import path from 'path';
import { registerChannel } from './registry.js';
import { Channel, NewMessage, OnInboundMessage, OnChatMetadata, RegisteredGroup } from '../types.js';
import { logger } from '../logger.js';
import { readEnvFile } from '../env.js';
import crypto from 'crypto';

const envConfig = readEnvFile(['NINJACLAW_WEB_PORT', 'NINJACLAW_WEB_TOKEN', 'NINJACLAW_WEB_USER_NAME']);
const WEB_PORT = parseInt(process.env.NINJACLAW_WEB_PORT || envConfig.NINJACLAW_WEB_PORT || '8484', 10);
const WEB_TOKEN = process.env.NINJACLAW_WEB_TOKEN || envConfig.NINJACLAW_WEB_TOKEN || '';
const WEB_USER_NAME = process.env.NINJACLAW_WEB_USER_NAME || envConfig.NINJACLAW_WEB_USER_NAME || 'Ofir Gavish';
const WEB_JID = 'web:main';

function checkToken(token: string): boolean {
  if (!WEB_TOKEN) return true; // Dev mode — no token required
  return crypto.timingSafeEqual(
    Buffer.from(token.padEnd(64, '\0')),
    Buffer.from(WEB_TOKEN.padEnd(64, '\0')),
  );
}

class WebChannel implements Channel {
  name = 'web';
  private connected = false;
  private clients = new Set<WebSocket>();
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
    const app = express();
    const server = createServer(app);
    const wss = new WebSocketServer({ server, path: '/ws/chat' });

    app.use(express.json());

    // Serve static web UI files
    const staticDir = path.join(process.cwd(), 'web_static');
    app.use('/static', express.static(staticDir));
    app.get('/', (_req: any, res: any) => {
      res.sendFile(path.join(staticDir, 'index.html'));
    });

    // REST chat endpoint
    app.post('/api/chat', async (req: any, res: any) => {
      const { message, token } = req.body;
      if (!checkToken(token || '')) {
        res.status(401).json({ error: 'Unauthorized' });
        return;
      }
      if (!message) {
        res.status(400).json({ error: 'message is required' });
        return;
      }

      const timestamp = new Date().toISOString();
      this.onChatMetadata(WEB_JID, timestamp, WEB_USER_NAME, 'web');

      const msg: NewMessage = {
        id: `web-${Date.now()}`,
        chat_jid: WEB_JID,
        sender: `web:${WEB_USER_NAME}`,
        sender_name: WEB_USER_NAME,
        content: message,
        timestamp,
        is_from_me: false,
      };
      this.onMessage(WEB_JID, msg);

      // The response will come through sendMessage → broadcast to WS clients
      res.json({ status: 'queued' });
    });

    // WebSocket chat
    wss.on('connection', (ws: any, req: any) => {
      const url = new URL(req.url || '/', `http://localhost:${WEB_PORT}`);
      const token = url.searchParams.get('token') || '';
      if (!checkToken(token)) {
        ws.close(4001, 'Unauthorized');
        return;
      }

      this.clients.add(ws);
      logger.info('Web chat client connected');

      ws.on('message', (raw: any) => {
        let data: { text?: string; type?: string };
        try {
          data = JSON.parse(raw.toString());
        } catch {
          data = { type: 'message', text: raw.toString() };
        }

        const text = data.text?.trim();
        if (!text) return;

        const timestamp = new Date().toISOString();
        this.onChatMetadata(WEB_JID, timestamp, WEB_USER_NAME, 'web');

        const msg: NewMessage = {
          id: `web-${Date.now()}`,
          chat_jid: WEB_JID,
          sender: `web:${WEB_USER_NAME}`,
          sender_name: WEB_USER_NAME,
          content: text,
          timestamp,
          is_from_me: false,
        };
        this.onMessage(WEB_JID, msg);
      });

      ws.on('close', () => {
        this.clients.delete(ws);
        logger.info('Web chat client disconnected');
      });
    });

    server.listen(WEB_PORT, '0.0.0.0', () => {
      logger.info({ port: WEB_PORT }, 'Web chat channel listening');
    });

    this.connected = true;
  }

  async sendMessage(jid: string, text: string): Promise<void> {
    if (!this.ownsJid(jid)) return;
    const payload = JSON.stringify({ type: 'message', text });
    for (const ws of this.clients) {
      if (ws.readyState === WebSocket.OPEN) {
        ws.send(payload);
      }
    }
  }

  isConnected(): boolean {
    return this.connected;
  }

  ownsJid(jid: string): boolean {
    return jid.startsWith('web:');
  }

  async setTyping(jid: string): Promise<void> {
    if (!this.ownsJid(jid)) return;
    const payload = JSON.stringify({ type: 'typing' });
    for (const ws of this.clients) {
      if (ws.readyState === WebSocket.OPEN) {
        ws.send(payload);
      }
    }
  }

  async disconnect(): Promise<void> {
    for (const ws of this.clients) ws.close();
    this.clients.clear();
    this.connected = false;
  }
}

registerChannel('web', (opts) => {
  return new WebChannel(opts.onMessage, opts.onChatMetadata, opts.registeredGroups);
});
