/**
 * NinjaClaw — Microsoft 365 MCP server (host-side).
 *
 * Exposes the per-instance delegated Graph helpers in `./graph.ts` as
 * Model Context Protocol tools so the containerized agent can call them
 * natively (mail/calendar/files/sites) instead of going through ad-hoc curl shims.
 *
 * Mounted on the existing Web channel Express app at `/mcp/m365`.
 * Stateless transport: a fresh `McpServer` + `StreamableHTTPServerTransport`
 * pair is created per HTTP request — simplest correct shape for short-lived
 * tool calls and avoids any per-session bookkeeping on the host.
 *
 * Auth: the channel layer must gate this route with the same Bearer
 * `NINJACLAW_WEB_TOKEN` used elsewhere; this module only resolves the
 * `instanceId` (?instanceId=... query, fallback to active instance) and
 * delegates to graph.ts which calls MSAL's silent token cache.
 */

import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StreamableHTTPServerTransport } from '@modelcontextprotocol/sdk/server/streamableHttp.js';
import { z } from 'zod';
import { logger } from '../../logger.js';
import { getActiveInstanceId } from '../admin/instances.js';
import {
  listMessages,
  getMessage,
  sendMail,
  listUpcomingEvents,
  searchFiles,
  listRecentDriveItems,
  listDriveChildren,
  getDriveItem,
  readDriveItemText,
  searchSites,
  listSiteDrives,
  listTeamsChats,
  listTeamsChatMessages,
  sendTeamsChatMessage,
  listJoinedTeams,
  listTeamChannels,
  listTeamsChannelMessages,
  listTeamsChannelMessageReplies,
  sendTeamsChannelMessage,
  sendTeamsChannelReply,
} from './graph.js';

export const M365_MCP_TOOL_NAMES = [
  'm365_mail_list',
  'm365_mail_get',
  'm365_mail_send',
  'm365_calendar_upcoming',
  'm365_file_search',
  'm365_drive_recent',
  'm365_drive_list',
  'm365_file_get_metadata',
  'm365_file_read_text',
  'm365_site_search',
  'm365_site_drives',
  'm365_teams_chats_list',
  'm365_teams_chat_messages',
  'm365_teams_send_chat_message',
  'm365_teams_joined_teams',
  'm365_teams_channels',
  'm365_teams_channel_messages',
  'm365_teams_channel_message_replies',
  'm365_teams_send_channel_message',
  'm365_teams_send_channel_reply',
] as const;

function resolveInstanceId(req: any): string | undefined {
  const fromQuery = (req.query?.instanceId as string) || '';
  if (fromQuery) return fromQuery;
  return getActiveInstanceId();
}

function buildServer(instanceId: string): McpServer {
  const server = new McpServer(
    { name: 'ninjaclaw-m365', version: '1.0.0' },
    { capabilities: { tools: {} } },
  );

  server.registerTool(
    'm365_mail_list',
    {
      description:
        'List recent messages from the signed-in user inbox via Microsoft Graph (delegated). Returns subject/from/preview only.',
      inputSchema: {
        top: z.number().int().min(1).max(50).optional()
          .describe('Number of messages to return (default 10).'),
        unreadOnly: z.boolean().optional()
          .describe('If true, only unread messages.'),
        search: z.string().optional()
          .describe('Optional Graph $search query string.'),
      },
    },
    async (args) => {
      const items = await listMessages(instanceId, args ?? {});
      return {
        content: [{ type: 'text', text: JSON.stringify(items, null, 2) }],
      };
    },
  );

  server.registerTool(
    'm365_mail_get',
    {
      description:
        'Fetch a single message by id from the signed-in user mailbox.',
      inputSchema: {
        id: z.string().min(1).describe('The Graph message id.'),
      },
    },
    async ({ id }) => {
      const message = await getMessage(instanceId, id);
      return {
        content: [{ type: 'text', text: JSON.stringify(message, null, 2) }],
      };
    },
  );

  server.registerTool(
    'm365_mail_send',
    {
      description:
        'Send an email from the signed-in user via Microsoft Graph (delegated). Provide bodyHtml or bodyText.',
      inputSchema: {
        to: z.array(z.string().email()).min(1)
          .describe('One or more recipient email addresses.'),
        subject: z.string().min(1),
        bodyHtml: z.string().optional(),
        bodyText: z.string().optional(),
        cc: z.array(z.string().email()).optional(),
        bcc: z.array(z.string().email()).optional(),
      },
    },
    async (args) => {
      const result = await sendMail(instanceId, args);
      return {
        content: [{ type: 'text', text: JSON.stringify(result, null, 2) }],
      };
    },
  );

  server.registerTool(
    'm365_calendar_upcoming',
    {
      description:
        'List upcoming calendar events for the signed-in user. Returns start/end/subject/location.',
      inputSchema: {
        top: z.number().int().min(1).max(50).optional()
          .describe('Maximum events to return (default 10).'),
        daysAhead: z.number().int().min(1).max(60).optional()
          .describe('How many days forward to scan (default 7).'),
      },
    },
    async (args) => {
      const events = await listUpcomingEvents(instanceId, args ?? {});
      return {
        content: [{ type: 'text', text: JSON.stringify(events, null, 2) }],
      };
    },
  );

  server.registerTool(
    'm365_file_search',
    {
      description:
        'Search OneDrive and SharePoint files the signed-in user can access. Returns file metadata, web links, drive IDs, and item IDs for follow-up reads/lists.',
      inputSchema: {
        query: z.string().min(1).describe('Search query, for example a filename, project name, or phrase.'),
        top: z.number().int().min(1).max(50).optional()
          .describe('Maximum results to return (default 10).'),
      },
    },
    async (args) => {
      const results = await searchFiles(instanceId, args);
      return {
        content: [{ type: 'text', text: JSON.stringify(results, null, 2) }],
      };
    },
  );

  server.registerTool(
    'm365_drive_recent',
    {
      description:
        'List recent OneDrive/SharePoint files visible to the signed-in user.',
      inputSchema: {
        top: z.number().int().min(1).max(50).optional()
          .describe('Maximum recent files to return (default 10).'),
      },
    },
    async (args) => {
      const items = await listRecentDriveItems(instanceId, args ?? {});
      return {
        content: [{ type: 'text', text: JSON.stringify(items, null, 2) }],
      };
    },
  );

  server.registerTool(
    'm365_drive_list',
    {
      description:
        'List children in the signed-in user OneDrive root, a OneDrive folder, or a SharePoint document library/folder. Use driveId+itemId from search/site results for SharePoint folders.',
      inputSchema: {
        driveId: z.string().optional()
          .describe('Optional drive/document-library id. If omitted, uses the user OneDrive.'),
        itemId: z.string().optional()
          .describe('Optional folder item id. If omitted, lists the drive root.'),
        top: z.number().int().min(1).max(200).optional()
          .describe('Maximum children to return (default 50).'),
      },
    },
    async (args) => {
      const items = await listDriveChildren(instanceId, args ?? {});
      return {
        content: [{ type: 'text', text: JSON.stringify(items, null, 2) }],
      };
    },
  );

  server.registerTool(
    'm365_file_get_metadata',
    {
      description:
        'Get metadata for a OneDrive or SharePoint file/folder by itemId plus optional driveId, or by webUrl.',
      inputSchema: {
        driveId: z.string().optional()
          .describe('Optional drive/document-library id.'),
        itemId: z.string().optional()
          .describe('File or folder item id. Required unless webUrl is provided.'),
        webUrl: z.string().url().optional()
          .describe('A OneDrive or SharePoint browser URL. Required unless itemId is provided.'),
      },
    },
    async (args) => {
      const item = await getDriveItem(instanceId, args);
      return {
        content: [{ type: 'text', text: JSON.stringify(item, null, 2) }],
      };
    },
  );

  server.registerTool(
    'm365_file_read_text',
    {
      description:
        'Read text-like file content from OneDrive or SharePoint by itemId plus optional driveId, or by webUrl. Returns text for txt/md/csv/json/html/code files; Office/binary files return metadata and a note.',
      inputSchema: {
        driveId: z.string().optional()
          .describe('Optional drive/document-library id.'),
        itemId: z.string().optional()
          .describe('File item id. Required unless webUrl is provided.'),
        webUrl: z.string().url().optional()
          .describe('A OneDrive or SharePoint browser URL. Required unless itemId is provided.'),
        maxBytes: z.number().int().min(1000).max(1000000).optional()
          .describe('Maximum bytes of text content to return (default 200000).'),
      },
    },
    async (args) => {
      const item = await readDriveItemText(instanceId, args);
      return {
        content: [{ type: 'text', text: JSON.stringify(item, null, 2) }],
      };
    },
  );

  server.registerTool(
    'm365_site_search',
    {
      description:
        'Search SharePoint sites the signed-in user can access. Use returned siteId with m365_site_drives.',
      inputSchema: {
        query: z.string().min(1).describe('SharePoint site search query.'),
        top: z.number().int().min(1).max(50).optional()
          .describe('Maximum sites to return (default 10).'),
      },
    },
    async (args) => {
      const sites = await searchSites(instanceId, args);
      return {
        content: [{ type: 'text', text: JSON.stringify(sites, null, 2) }],
      };
    },
  );

  server.registerTool(
    'm365_site_drives',
    {
      description:
        'List document libraries/drives for a SharePoint site the signed-in user can access.',
      inputSchema: {
        siteId: z.string().min(1).describe('SharePoint site id returned by m365_site_search or file metadata.'),
      },
    },
    async ({ siteId }) => {
      const drives = await listSiteDrives(instanceId, siteId);
      return {
        content: [{ type: 'text', text: JSON.stringify(drives, null, 2) }],
      };
    },
  );

  server.registerTool(
    'm365_teams_chats_list',
    {
      description:
        'List Microsoft Teams one-on-one and group chats visible to the signed-in user. Use chatId with m365_teams_chat_messages or m365_teams_send_chat_message.',
      inputSchema: {
        top: z.number().int().min(1).max(50).optional()
          .describe('Maximum chats to return (default 20).'),
      },
    },
    async (args) => {
      const chats = await listTeamsChats(instanceId, args ?? {});
      return {
        content: [{ type: 'text', text: JSON.stringify(chats, null, 2) }],
      };
    },
  );

  server.registerTool(
    'm365_teams_chat_messages',
    {
      description:
        'Read recent messages from a Microsoft Teams one-on-one or group chat the signed-in user can access.',
      inputSchema: {
        chatId: z.string().min(1).describe('Teams chat id from m365_teams_chats_list.'),
        top: z.number().int().min(1).max(50).optional()
          .describe('Maximum messages to return (default 25).'),
      },
    },
    async (args) => {
      const messages = await listTeamsChatMessages(instanceId, args);
      return {
        content: [{ type: 'text', text: JSON.stringify(messages, null, 2) }],
      };
    },
  );

  server.registerTool(
    'm365_teams_send_chat_message',
    {
      description:
        'Send a Microsoft Teams chat message as the signed-in user. Requires confirm=true to prevent accidental sends. Provide bodyText or bodyHtml.',
      inputSchema: {
        chatId: z.string().min(1).describe('Teams chat id from m365_teams_chats_list.'),
        bodyText: z.string().optional().describe('Plain text body. It will be HTML-escaped for Teams.'),
        bodyHtml: z.string().optional().describe('Optional Teams-compatible HTML body.'),
        confirm: z.literal(true).describe('Must be true to send the message.'),
      },
    },
    async (args) => {
      const message = await sendTeamsChatMessage(instanceId, args);
      return {
        content: [{ type: 'text', text: JSON.stringify(message, null, 2) }],
      };
    },
  );

  server.registerTool(
    'm365_teams_joined_teams',
    {
      description:
        'List Microsoft Teams teams joined by the signed-in user. Use teamId with m365_teams_channels.',
      inputSchema: {
        top: z.number().int().min(1).max(200).optional()
          .describe('Maximum teams to return (default 50).'),
      },
    },
    async (args) => {
      const teams = await listJoinedTeams(instanceId, args ?? {});
      return {
        content: [{ type: 'text', text: JSON.stringify(teams, null, 2) }],
      };
    },
  );

  server.registerTool(
    'm365_teams_channels',
    {
      description:
        'List channels in a Microsoft Teams team visible to the signed-in user. Use channelId with channel message tools.',
      inputSchema: {
        teamId: z.string().min(1).describe('Team id from m365_teams_joined_teams.'),
        top: z.number().int().min(1).max(200).optional()
          .describe('Maximum channels to return (default 50).'),
      },
    },
    async (args) => {
      const channels = await listTeamChannels(instanceId, args);
      return {
        content: [{ type: 'text', text: JSON.stringify(channels, null, 2) }],
      };
    },
  );

  server.registerTool(
    'm365_teams_channel_messages',
    {
      description:
        'Read recent top-level Microsoft Teams channel messages. Use m365_teams_channel_message_replies to read a thread.',
      inputSchema: {
        teamId: z.string().min(1).describe('Team id from m365_teams_joined_teams.'),
        channelId: z.string().min(1).describe('Channel id from m365_teams_channels.'),
        top: z.number().int().min(1).max(50).optional()
          .describe('Maximum messages to return (default 25).'),
      },
    },
    async (args) => {
      const messages = await listTeamsChannelMessages(instanceId, args);
      return {
        content: [{ type: 'text', text: JSON.stringify(messages, null, 2) }],
      };
    },
  );

  server.registerTool(
    'm365_teams_channel_message_replies',
    {
      description:
        'Read replies in a Microsoft Teams channel message thread.',
      inputSchema: {
        teamId: z.string().min(1).describe('Team id from m365_teams_joined_teams.'),
        channelId: z.string().min(1).describe('Channel id from m365_teams_channels.'),
        messageId: z.string().min(1).describe('Top-level channel message id.'),
        top: z.number().int().min(1).max(50).optional()
          .describe('Maximum replies to return (default 25).'),
      },
    },
    async (args) => {
      const replies = await listTeamsChannelMessageReplies(instanceId, args);
      return {
        content: [{ type: 'text', text: JSON.stringify(replies, null, 2) }],
      };
    },
  );

  server.registerTool(
    'm365_teams_send_channel_message',
    {
      description:
        'Send a new top-level Microsoft Teams channel message as the signed-in user. Requires confirm=true to prevent accidental sends. Provide bodyText or bodyHtml.',
      inputSchema: {
        teamId: z.string().min(1).describe('Team id from m365_teams_joined_teams.'),
        channelId: z.string().min(1).describe('Channel id from m365_teams_channels.'),
        bodyText: z.string().optional().describe('Plain text body. It will be HTML-escaped for Teams.'),
        bodyHtml: z.string().optional().describe('Optional Teams-compatible HTML body.'),
        confirm: z.literal(true).describe('Must be true to send the message.'),
      },
    },
    async (args) => {
      const message = await sendTeamsChannelMessage(instanceId, args);
      return {
        content: [{ type: 'text', text: JSON.stringify(message, null, 2) }],
      };
    },
  );

  server.registerTool(
    'm365_teams_send_channel_reply',
    {
      description:
        'Reply to a Microsoft Teams channel message thread as the signed-in user. Requires confirm=true to prevent accidental sends. Provide bodyText or bodyHtml.',
      inputSchema: {
        teamId: z.string().min(1).describe('Team id from m365_teams_joined_teams.'),
        channelId: z.string().min(1).describe('Channel id from m365_teams_channels.'),
        messageId: z.string().min(1).describe('Top-level channel message id to reply to.'),
        bodyText: z.string().optional().describe('Plain text body. It will be HTML-escaped for Teams.'),
        bodyHtml: z.string().optional().describe('Optional Teams-compatible HTML body.'),
        confirm: z.literal(true).describe('Must be true to send the reply.'),
      },
    },
    async (args) => {
      const reply = await sendTeamsChannelReply(instanceId, args);
      return {
        content: [{ type: 'text', text: JSON.stringify(reply, null, 2) }],
      };
    },
  );

  return server;
}

/**
 * Express handler for the MCP Streamable HTTP endpoint.
 * Caller must apply auth (Bearer token) before invoking.
 *
 * Usage:
 *   app.post('/mcp/m365', requireAuth, m365McpHandler);
 *   app.get('/mcp/m365', requireAuth, m365McpHandler);
 *   app.delete('/mcp/m365', requireAuth, m365McpHandler);
 */
export async function m365McpHandler(req: any, res: any): Promise<void> {
  const instanceId = resolveInstanceId(req);
  if (!instanceId) {
    res.status(400).json({
      jsonrpc: '2.0',
      error: { code: -32000, message: 'No active Agent 365 instance — pass ?instanceId=...' },
      id: null,
    });
    return;
  }

  // Stateless: new transport + server per request.
  const transport = new StreamableHTTPServerTransport({
    sessionIdGenerator: undefined,
  });

  res.on('close', () => {
    transport.close().catch(() => {});
  });

  try {
    const server = buildServer(instanceId);
    await server.connect(transport);
    await transport.handleRequest(req, res, req.body);
  } catch (err: any) {
    logger.error({ err, instanceId }, 'M365 MCP request failed');
    if (!res.headersSent) {
      res.status(500).json({
        jsonrpc: '2.0',
        error: { code: -32603, message: err?.message || 'Internal error' },
        id: null,
      });
    }
  }
}
