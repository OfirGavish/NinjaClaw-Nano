/**
 * NinjaClaw — Web Chat channel for NinjaClaw.
 *
 * Implements the Channel interface using Express + WebSocket.
 * Provides the same web UI as the Python NinjaClaw web_api.py.
 */

import express from 'express';
import { WebSocketServer, WebSocket } from 'ws';
import { createServer } from 'http';
import path from 'path';
import { registerChannel } from './registry.js';
import { Channel, NewMessage, OnInboundMessage, OnChatMetadata, RegisteredGroup } from '../types.js';
import { logger } from '../logger.js';
import { readEnvFile } from '../env.js';
import { readSettings, writeSettings } from '../settings.js';
import {
  startDeviceFlow,
  getDeviceFlowStatus,
  getCachedAdminAccount,
  signOut,
} from '../agent365/admin/oauth.js';
import { publishBlueprint } from '../agent365/admin/blueprints.js';
import {
  listInstances,
  getInstance,
  createInstance,
  deleteInstance,
  setActiveInstance,
  getActiveInstanceId,
  enableUserDelegated,
} from '../agent365/admin/instances.js';
import {
  startUserDeviceFlow,
  getUserDeviceFlowStatus,
  getCachedUserAccount,
  userSignOut,
} from '../agent365/user/oauth.js';
import {
  getMe,
  listMessages,
  getMessage,
  sendMail,
  listUpcomingEvents,
} from '../agent365/user/graph.js';
import { m365McpHandler } from '../agent365/user/mcp-server.js';
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

    // Settings page (served from web_static/settings.html)
    app.get('/settings', (_req: any, res: any) => {
      res.sendFile(path.join(staticDir, 'settings.html'));
    });

    // Settings API — read/write whitelisted .env values.
    // Auth: same token as chat. Token is required when WEB_TOKEN is set;
    // otherwise dev-mode (open access) like the rest of the channel.
    function settingsAuthOk(req: any): boolean {
      const header = req.header('authorization') || '';
      const bearer = header.startsWith('Bearer ') ? header.slice(7) : '';
      const queryToken = (req.query?.token as string) || '';
      const bodyToken = (req.body?.token as string) || '';
      const supplied = bearer || queryToken || bodyToken;
      return checkToken(supplied);
    }

    app.get('/api/settings', (req: any, res: any) => {
      if (!settingsAuthOk(req)) {
        res.status(401).json({ error: 'Unauthorized' });
        return;
      }
      try {
        res.json(readSettings());
      } catch (err) {
        logger.error({ err }, 'failed to read settings');
        res.status(500).json({ error: 'failed to read settings' });
      }
    });

    app.post('/api/settings', (req: any, res: any) => {
      if (!settingsAuthOk(req)) {
        res.status(401).json({ error: 'Unauthorized' });
        return;
      }
      const body = req.body || {};
      const updates = body.updates;
      if (!updates || typeof updates !== 'object' || Array.isArray(updates)) {
        res.status(400).json({ error: '`updates` object is required' });
        return;
      }
      try {
        const result = writeSettings(updates as Record<string, string>);
        res.json({
          ...result,
          notice:
            'Restart NinjaClaw for changes to take effect (channel processes only re-read .env at startup).',
        });
      } catch (err: any) {
        logger.error({ err }, 'failed to write settings');
        res
          .status(500)
          .json({ error: 'failed to write settings', detail: err?.message });
      }
    });

    // ---------------------------------------------------------------------
    // Agent 365 admin console — device-code sign-in + blueprint publish.
    // Uses the same web-token auth gate as the rest of /api/settings.
    // ---------------------------------------------------------------------

    app.get('/agent365', (_req: any, res: any) => {
      res.sendFile(path.join(staticDir, 'agent365.html'));
    });

    app.get('/api/agent365/account', async (req: any, res: any) => {
      if (!settingsAuthOk(req)) {
        res.status(401).json({ error: 'Unauthorized' });
        return;
      }
      const account = await getCachedAdminAccount();
      res.json({
        signedIn: !!account,
        username: account?.username,
        tenantId: account?.tenantId,
      });
    });

    app.post('/api/agent365/signin', async (req: any, res: any) => {
      if (!settingsAuthOk(req)) {
        res.status(401).json({ error: 'Unauthorized' });
        return;
      }
      try {
        const flow = await startDeviceFlow();
        res.json(flow);
      } catch (err: any) {
        logger.error({ err }, 'agent365 device-code start failed');
        res
          .status(500)
          .json({ error: 'sign-in start failed', detail: err?.message });
      }
    });

    app.get('/api/agent365/signin/:flowId', (req: any, res: any) => {
      if (!settingsAuthOk(req)) {
        res.status(401).json({ error: 'Unauthorized' });
        return;
      }
      const status = getDeviceFlowStatus(req.params.flowId);
      if (!status) {
        res.status(404).json({ error: 'unknown flow id' });
        return;
      }
      res.json(status);
    });

    app.post('/api/agent365/signout', async (req: any, res: any) => {
      if (!settingsAuthOk(req)) {
        res.status(401).json({ error: 'Unauthorized' });
        return;
      }
      await signOut();
      res.json({ ok: true });
    });

    app.post('/api/agent365/blueprint/publish', async (req: any, res: any) => {
      if (!settingsAuthOk(req)) {
        res.status(401).json({ error: 'Unauthorized' });
        return;
      }
      const dryRun = !!(req.body || {}).dryRun;
      try {
        const result = await publishBlueprint({ dryRun });
        // Use 422 (not 502) for publish errors — Cloudflare intercepts 502
        // from the origin and replaces the JSON body with its own HTML page.
        const httpStatus = result.status === 'error' ? 422 : 200;
        res.status(httpStatus).json(result);
      } catch (err: any) {
        logger.error({ err }, 'agent365 blueprint publish crashed');
        res.status(500).json({ error: err?.message || String(err) });
      }
    });

    // ---------------------------------------------------------------------
    // Agent 365 instances — multi-instance management under one blueprint.
    // ---------------------------------------------------------------------

    app.get('/api/agent365/instances', (req: any, res: any) => {
      if (!settingsAuthOk(req)) {
        res.status(401).json({ error: 'Unauthorized' });
        return;
      }
      try {
        res.json(listInstances());
      } catch (err: any) {
        logger.error({ err }, 'agent365 listInstances failed');
        res.status(500).json({ error: err?.message || String(err) });
      }
    });

    app.get('/api/agent365/instances/:id', (req: any, res: any) => {
      if (!settingsAuthOk(req)) {
        res.status(401).json({ error: 'Unauthorized' });
        return;
      }
      const instance = getInstance(req.params.id);
      if (!instance) {
        res.status(404).json({ error: 'instance not found' });
        return;
      }
      res.json(instance);
    });

    app.post('/api/agent365/instances', async (req: any, res: any) => {
      if (!settingsAuthOk(req)) {
        res.status(401).json({ error: 'Unauthorized' });
        return;
      }
      const body = req.body || {};
      if (!body.displayName || typeof body.displayName !== 'string') {
        res.status(400).json({ error: 'displayName is required' });
        return;
      }
      if (
        body.hostingMode !== 'bot-framework' &&
        body.hostingMode !== 'tailscale'
      ) {
        res
          .status(400)
          .json({ error: 'hostingMode must be "bot-framework" or "tailscale"' });
        return;
      }
      try {
        const result = await createInstance(body);
        res.status(201).json(result);
      } catch (err: any) {
        logger.error({ err }, 'agent365 createInstance failed');
        res
          .status(502)
          .json({ error: err?.message || String(err) });
      }
    });

    app.post('/api/agent365/instances/:id/activate', (req: any, res: any) => {
      if (!settingsAuthOk(req)) {
        res.status(401).json({ error: 'Unauthorized' });
        return;
      }
      try {
        const instance = setActiveInstance(req.params.id);
        res.json({
          instance,
          notice:
            'Restart NinjaClaw for the new instance credentials to take effect.',
        });
      } catch (err: any) {
        const status = /not found/i.test(err?.message) ? 404 : 500;
        res.status(status).json({ error: err?.message || String(err) });
      }
    });

    app.delete('/api/agent365/instances/:id', async (req: any, res: any) => {
      if (!settingsAuthOk(req)) {
        res.status(401).json({ error: 'Unauthorized' });
        return;
      }
      const cascade = req.query?.cascade === 'true';
      try {
        const result = await deleteInstance(req.params.id, { cascade });
        const status = result.removed ? 200 : 404;
        res.status(status).json(result);
      } catch (err: any) {
        logger.error({ err }, 'agent365 deleteInstance failed');
        res.status(500).json({ error: err?.message || String(err) });
      }
    });

    // ---------------------------------------------------------------------
    // Agent 365 — per-instance USER (delegated) Graph access.
    // The user signs in once; the agent then reads mail/calendar/files on
    // their behalf using the active instance's Entra app + delegated
    // permissions. Resolves the instance id from the request (?instanceId=)
    // or falls back to the active instance.
    // ---------------------------------------------------------------------

    function resolveInstanceId(req: any): string | null {
      const fromQuery = (req.query?.instanceId as string) || '';
      const fromBody = (req.body?.instanceId as string) || '';
      const id = fromQuery || fromBody || getActiveInstanceId();
      return id || null;
    }

    app.get('/api/agent365/user/account', async (req: any, res: any) => {
      if (!settingsAuthOk(req)) { res.status(401).json({ error: 'Unauthorized' }); return; }
      const instanceId = resolveInstanceId(req);
      if (!instanceId) { res.json({ signedIn: false, error: 'no active instance' }); return; }
      const account = await getCachedUserAccount(instanceId);
      if (!account) { res.json({ signedIn: false, instanceId }); return; }
      // Try to fetch /me for richer profile info; fall back gracefully.
      try {
        const me = await getMe(instanceId);
        res.json({ signedIn: true, instanceId, account: { username: account.username, tenantId: account.tenantId }, me });
      } catch (err: any) {
        res.json({ signedIn: true, instanceId, account: { username: account.username, tenantId: account.tenantId }, meError: err?.message });
      }
    });

    app.post('/api/agent365/user/signin', async (req: any, res: any) => {
      if (!settingsAuthOk(req)) { res.status(401).json({ error: 'Unauthorized' }); return; }
      const instanceId = resolveInstanceId(req);
      if (!instanceId) { res.status(400).json({ error: 'no active instance — create or activate one first' }); return; }
      try {
        const flow = await startUserDeviceFlow(instanceId);
        res.json({ ...flow, instanceId });
      } catch (err: any) {
        logger.error({ err, instanceId }, 'agent365 user device-code start failed');
        res.status(500).json({ error: 'sign-in start failed', detail: err?.message });
      }
    });

    app.get('/api/agent365/user/signin/:flowId', (req: any, res: any) => {
      if (!settingsAuthOk(req)) { res.status(401).json({ error: 'Unauthorized' }); return; }
      const status = getUserDeviceFlowStatus(req.params.flowId);
      if (!status) { res.status(404).json({ error: 'unknown flow id' }); return; }
      res.json(status);
    });

    app.post('/api/agent365/user/signout', async (req: any, res: any) => {
      if (!settingsAuthOk(req)) { res.status(401).json({ error: 'Unauthorized' }); return; }
      const instanceId = resolveInstanceId(req);
      if (!instanceId) { res.status(400).json({ error: 'no active instance' }); return; }
      await userSignOut(instanceId);
      res.json({ ok: true, instanceId });
    });

    app.post('/api/agent365/instances/:id/enable-user-delegated', async (req: any, res: any) => {
      if (!settingsAuthOk(req)) { res.status(401).json({ error: 'Unauthorized' }); return; }
      try {
        const result = await enableUserDelegated(req.params.id);
        res.json(result);
      } catch (err: any) {
        logger.error({ err, id: req.params.id }, 'agent365 enableUserDelegated failed');
        res.status(500).json({ error: err?.message || String(err) });
      }
    });

    // Graph proxies — keep surface minimal; gated by web token + per-user delegated scopes.
    app.get('/api/agent365/user/mail', async (req: any, res: any) => {
      if (!settingsAuthOk(req)) { res.status(401).json({ error: 'Unauthorized' }); return; }
      const instanceId = resolveInstanceId(req);
      if (!instanceId) { res.status(400).json({ error: 'no active instance' }); return; }
      try {
        const top = req.query?.top ? parseInt(String(req.query.top), 10) : undefined;
        const unreadOnly = req.query?.unreadOnly === 'true';
        const search = req.query?.search ? String(req.query.search) : undefined;
        const items = await listMessages(instanceId, { top, unreadOnly, search });
        res.json({ items });
      } catch (err: any) {
        res.status(500).json({ error: err?.message || String(err) });
      }
    });

    app.get('/api/agent365/user/mail/:id', async (req: any, res: any) => {
      if (!settingsAuthOk(req)) { res.status(401).json({ error: 'Unauthorized' }); return; }
      const instanceId = resolveInstanceId(req);
      if (!instanceId) { res.status(400).json({ error: 'no active instance' }); return; }
      try {
        const message = await getMessage(instanceId, req.params.id);
        res.json(message);
      } catch (err: any) {
        res.status(500).json({ error: err?.message || String(err) });
      }
    });

    app.post('/api/agent365/user/mail/send', async (req: any, res: any) => {
      if (!settingsAuthOk(req)) { res.status(401).json({ error: 'Unauthorized' }); return; }
      const instanceId = resolveInstanceId(req);
      if (!instanceId) { res.status(400).json({ error: 'no active instance' }); return; }
      const body = req.body || {};
      try {
        const result = await sendMail(instanceId, {
          to: Array.isArray(body.to) ? body.to : (body.to ? [String(body.to)] : []),
          cc: Array.isArray(body.cc) ? body.cc : undefined,
          bcc: Array.isArray(body.bcc) ? body.bcc : undefined,
          subject: String(body.subject || ''),
          bodyHtml: body.bodyHtml ? String(body.bodyHtml) : undefined,
          bodyText: body.bodyText ? String(body.bodyText) : undefined,
        });
        res.json(result);
      } catch (err: any) {
        res.status(500).json({ error: err?.message || String(err) });
      }
    });

    app.get('/api/agent365/user/events', async (req: any, res: any) => {
      if (!settingsAuthOk(req)) { res.status(401).json({ error: 'Unauthorized' }); return; }
      const instanceId = resolveInstanceId(req);
      if (!instanceId) { res.status(400).json({ error: 'no active instance' }); return; }
      try {
        const top = req.query?.top ? parseInt(String(req.query.top), 10) : undefined;
        const daysAhead = req.query?.daysAhead ? parseInt(String(req.query.daysAhead), 10) : undefined;
        const items = await listUpcomingEvents(instanceId, { top, daysAhead });
        res.json({ items });
      } catch (err: any) {
        res.status(500).json({ error: err?.message || String(err) });
      }
    });

    // M365 MCP server endpoint — exposed to the containerized agent so it can
    // call Graph (mail/calendar) as native MCP tools instead of curl shims.
    // Bearer token gated; stateless transport per request (see mcp-server.ts).
    app.all('/mcp/m365', async (req: any, res: any) => {
      if (!settingsAuthOk(req)) { res.status(401).json({ error: 'Unauthorized' }); return; }
      await m365McpHandler(req, res);
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
