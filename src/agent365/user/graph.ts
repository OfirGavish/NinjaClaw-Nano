/**
 * Agent 365 — Microsoft Graph helpers built on the per-instance user token.
 *
 * Thin wrappers over the v1.0 Graph endpoints we expose to the agent and
 * to the admin Web UI. Everything funnels through `graphGet` /
 * `graphPost` / `graphGetContent` so token refresh, error normalization,
 * and logging happen in one place.
 */

import { logger } from '../../logger.js';
import { acquireUserToken } from './oauth.js';

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';

async function graphGet(
  instanceId: string,
  pathAndQuery: string,
): Promise<any> {
  const { token } = await acquireUserToken(instanceId);
  const res = await fetch(`${GRAPH_BASE}${pathAndQuery}`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  const text = await res.text();
  let json: any;
  if (text) {
    try { json = JSON.parse(text); } catch { json = text; }
  }
  if (!res.ok) {
    const detail = json?.error?.message || text || `HTTP ${res.status}`;
    logger.warn({ status: res.status, pathAndQuery, detail }, 'Graph GET failed');
    throw new Error(`Graph GET ${pathAndQuery} failed: ${detail}`);
  }
  return json;
}

async function graphPost(
  instanceId: string,
  pathAndQuery: string,
  body: unknown,
): Promise<any> {
  const { token } = await acquireUserToken(instanceId);
  const res = await fetch(`${GRAPH_BASE}${pathAndQuery}`, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify(body),
  });
  if (!res.ok) {
    const text = await res.text();
    let detail = text;
    try { detail = JSON.parse(text)?.error?.message || text; } catch { /* keep raw */ }
    logger.warn({ status: res.status, pathAndQuery, detail }, 'Graph POST failed');
    throw new Error(`Graph POST ${pathAndQuery} failed: ${detail}`);
  }
  if (res.status === 202 || res.status === 204) return { ok: true };
  const text = await res.text();
  if (!text) return { ok: true };
  try { return JSON.parse(text); } catch { return text; }
}

async function graphGetContent(
  instanceId: string,
  pathAndQuery: string,
  maxBytes: number,
): Promise<{ contentType: string; buffer: Buffer; byteLength: number; truncated: boolean }> {
  const { token } = await acquireUserToken(instanceId);
  const res = await fetch(`${GRAPH_BASE}${pathAndQuery}`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!res.ok) {
    const text = await res.text();
    let detail = text;
    try { detail = JSON.parse(text)?.error?.message || text; } catch { /* keep raw */ }
    logger.warn({ status: res.status, pathAndQuery, detail }, 'Graph content GET failed');
    throw new Error(`Graph GET ${pathAndQuery} failed: ${detail}`);
  }
  const buffer = Buffer.from(await res.arrayBuffer());
  const limited = buffer.subarray(0, maxBytes);
  return {
    contentType: res.headers.get('content-type') || 'application/octet-stream',
    buffer: limited,
    byteLength: buffer.byteLength,
    truncated: buffer.byteLength > limited.byteLength,
  };
}

function clampInt(value: number | undefined, fallback: number, min: number, max: number): number {
  if (!Number.isFinite(value)) return fallback;
  return Math.min(Math.max(Math.trunc(value ?? fallback), min), max);
}

function truncateString(value: unknown, maxChars = 10_000): string | undefined {
  if (typeof value !== 'string') return undefined;
  if (value.length <= maxChars) return value;
  return `${value.slice(0, maxChars)}…`;
}

function htmlFromText(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;')
    .replace(/\r?\n/g, '<br>');
}

function teamsMessageBody(opts: { bodyHtml?: string; bodyText?: string }): { contentType: 'html'; content: string } {
  const content = opts.bodyHtml?.trim()
    ? opts.bodyHtml
    : htmlFromText(opts.bodyText || '');
  if (!content.trim()) throw new Error('Teams message body is required');
  // Graph Teams send APIs accept HTML bodies; escape plain text before posting.
  return { contentType: 'html', content };
}

function requireConfirmed(confirm: boolean | undefined, operation: string): void {
  if (confirm !== true) {
    throw new Error(`${operation}: confirm must be true to send a Teams message`);
  }
}

export interface GraphMe {
  id: string;
  displayName?: string;
  mail?: string;
  userPrincipalName?: string;
  jobTitle?: string;
}

export async function getMe(instanceId: string): Promise<GraphMe> {
  return graphGet(instanceId, '/me?$select=id,displayName,mail,userPrincipalName,jobTitle');
}

export interface MailMessageSummary {
  id: string;
  subject?: string;
  from?: string;
  receivedDateTime?: string;
  bodyPreview?: string;
  isRead?: boolean;
  webLink?: string;
}

function summarizeMessage(m: any): MailMessageSummary {
  return {
    id: m.id,
    subject: m.subject,
    from: m.from?.emailAddress?.address,
    receivedDateTime: m.receivedDateTime,
    bodyPreview: m.bodyPreview,
    isRead: m.isRead,
    webLink: m.webLink,
  };
}

export interface ListMailOptions {
  top?: number;
  unreadOnly?: boolean;
  search?: string;
}

export async function listMessages(
  instanceId: string,
  opts: ListMailOptions = {},
): Promise<MailMessageSummary[]> {
  const top = Math.min(Math.max(opts.top ?? 25, 1), 50);
  const params = new URLSearchParams();
  params.set('$top', String(top));
  params.set('$select', 'id,subject,from,receivedDateTime,bodyPreview,isRead,webLink');
  if (opts.search) {
    params.set('$search', `"${opts.search.replace(/"/g, '')}"`);
  } else {
    params.set('$orderby', 'receivedDateTime desc');
    if (opts.unreadOnly) params.set('$filter', 'isRead eq false');
  }
  const data = await graphGet(instanceId, `/me/messages?${params.toString()}`);
  return (data.value || []).map(summarizeMessage);
}

export interface MailMessageFull extends MailMessageSummary {
  toRecipients?: string[];
  ccRecipients?: string[];
  bodyContentType?: 'html' | 'text';
  bodyContent?: string;
}

export async function getMessage(
  instanceId: string,
  messageId: string,
): Promise<MailMessageFull> {
  const m = await graphGet(
    instanceId,
    `/me/messages/${encodeURIComponent(messageId)}?$select=id,subject,from,toRecipients,ccRecipients,receivedDateTime,bodyPreview,isRead,webLink,body`,
  );
  return {
    ...summarizeMessage(m),
    toRecipients: (m.toRecipients || []).map((r: any) => r.emailAddress?.address).filter(Boolean),
    ccRecipients: (m.ccRecipients || []).map((r: any) => r.emailAddress?.address).filter(Boolean),
    bodyContentType: m.body?.contentType,
    bodyContent: m.body?.content,
  };
}

export interface SendMailOptions {
  to: string[];
  subject: string;
  bodyHtml?: string;
  bodyText?: string;
  cc?: string[];
  bcc?: string[];
}

export async function sendMail(
  instanceId: string,
  opts: SendMailOptions,
): Promise<{ ok: true }> {
  if (!opts.to?.length) throw new Error('sendMail: at least one recipient required');
  if (!opts.subject) throw new Error('sendMail: subject required');
  const body = opts.bodyHtml
    ? { contentType: 'HTML', content: opts.bodyHtml }
    : { contentType: 'Text', content: opts.bodyText || '' };
  const message = {
    subject: opts.subject,
    body,
    toRecipients: opts.to.map((a) => ({ emailAddress: { address: a } })),
    ccRecipients: (opts.cc || []).map((a) => ({ emailAddress: { address: a } })),
    bccRecipients: (opts.bcc || []).map((a) => ({ emailAddress: { address: a } })),
  };
  await graphPost(instanceId, '/me/sendMail', { message, saveToSentItems: true });
  return { ok: true };
}

export interface CalendarEventSummary {
  id: string;
  subject?: string;
  start?: string;
  end?: string;
  location?: string;
  organizer?: string;
  webLink?: string;
}

export async function listUpcomingEvents(
  instanceId: string,
  opts: { top?: number; daysAhead?: number } = {},
): Promise<CalendarEventSummary[]> {
  const top = Math.min(Math.max(opts.top ?? 25, 1), 50);
  const days = Math.min(Math.max(opts.daysAhead ?? 7, 1), 60);
  const start = new Date();
  const end = new Date(start.getTime() + days * 24 * 60 * 60 * 1000);
  const params = new URLSearchParams({
    startDateTime: start.toISOString(),
    endDateTime: end.toISOString(),
    $top: String(top),
    $orderby: 'start/dateTime',
    $select: 'id,subject,start,end,location,organizer,webLink',
  });
  const data = await graphGet(instanceId, `/me/calendarView?${params.toString()}`);
  return (data.value || []).map((e: any) => ({
    id: e.id,
    subject: e.subject,
    start: e.start?.dateTime,
    end: e.end?.dateTime,
    location: e.location?.displayName,
    organizer: e.organizer?.emailAddress?.address,
    webLink: e.webLink,
  }));
}

export interface DriveItemSummary {
  id: string;
  name?: string;
  webUrl?: string;
  size?: number;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  driveId?: string;
  siteId?: string;
  parentPath?: string;
  folder?: boolean;
  packageType?: string;
  mimeType?: string;
}

export interface FileSearchResult extends DriveItemSummary {
  hitId?: string;
  rank?: number;
  summary?: string;
}

export interface SiteSummary {
  id: string;
  name?: string;
  displayName?: string;
  webUrl?: string;
  description?: string;
}

export interface DriveSummary {
  id: string;
  name?: string;
  webUrl?: string;
  driveType?: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
}

function summarizeDriveItem(item: any): DriveItemSummary {
  const parent = item.parentReference || {};
  return {
    id: item.id,
    name: item.name,
    webUrl: item.webUrl,
    size: item.size,
    createdDateTime: item.createdDateTime,
    lastModifiedDateTime: item.lastModifiedDateTime,
    driveId: parent.driveId,
    siteId: parent.siteId,
    parentPath: parent.path,
    folder: !!item.folder,
    packageType: item.package?.type,
    mimeType: item.file?.mimeType,
  };
}

function driveItemSelect(): string {
  return 'id,name,webUrl,size,createdDateTime,lastModifiedDateTime,parentReference,file,folder,package';
}

function shareIdFromUrl(webUrl: string): string {
  const encoded = Buffer.from(webUrl, 'utf-8')
    .toString('base64')
    .replace(/\+/g, '-')
    .replace(/\//g, '_')
    .replace(/=+$/g, '');
  return `u!${encoded}`;
}

function driveItemPath(opts: { driveId?: string; itemId?: string }): string {
  if (opts.driveId) {
    const driveId = encodeURIComponent(opts.driveId);
    return opts.itemId
      ? `/drives/${driveId}/items/${encodeURIComponent(opts.itemId)}`
      : `/drives/${driveId}/root`;
  }
  return opts.itemId
    ? `/me/drive/items/${encodeURIComponent(opts.itemId)}`
    : '/me/drive/root';
}

export async function searchFiles(
  instanceId: string,
  opts: { query: string; top?: number },
): Promise<FileSearchResult[]> {
  const query = opts.query?.trim();
  if (!query) throw new Error('searchFiles: query is required');
  const size = clampInt(opts.top, 10, 1, 50);
  const data = await graphPost(instanceId, '/search/query', {
    requests: [
      {
        entityTypes: ['driveItem'],
        query: { queryString: query },
        from: 0,
        size,
        fields: [
          'id',
          'name',
          'webUrl',
          'size',
          'createdDateTime',
          'lastModifiedDateTime',
          'parentReference',
          'file',
          'folder',
          'package',
        ],
      },
    ],
  });
  const containers = data?.value?.flatMap((v: any) => v.hitsContainers || []) || [];
  const hits = containers.flatMap((c: any) => c.hits || []);
  return hits.map((hit: any) => ({
    ...summarizeDriveItem(hit.resource || {}),
    hitId: hit.hitId,
    rank: hit.rank,
    summary: hit.summary,
  }));
}

export async function listRecentDriveItems(
  instanceId: string,
  opts: { top?: number } = {},
): Promise<DriveItemSummary[]> {
  const top = clampInt(opts.top, 10, 1, 50);
  const params = new URLSearchParams();
  params.set('$top', String(top));
  params.set('$select', driveItemSelect());
  const data = await graphGet(instanceId, `/me/drive/recent?${params.toString()}`);
  return (data.value || []).map(summarizeDriveItem);
}

export async function listDriveChildren(
  instanceId: string,
  opts: { driveId?: string; itemId?: string; top?: number } = {},
): Promise<DriveItemSummary[]> {
  const top = clampInt(opts.top, 50, 1, 200);
  const params = new URLSearchParams();
  params.set('$top', String(top));
  params.set('$select', driveItemSelect());
  const data = await graphGet(
    instanceId,
    `${driveItemPath(opts)}/children?${params.toString()}`,
  );
  return (data.value || []).map(summarizeDriveItem);
}

export async function getDriveItem(
  instanceId: string,
  opts: { driveId?: string; itemId?: string; webUrl?: string },
): Promise<DriveItemSummary> {
  const params = new URLSearchParams();
  params.set('$select', driveItemSelect());
  const path = opts.webUrl
    ? `/shares/${shareIdFromUrl(opts.webUrl)}/driveItem`
    : driveItemPath(opts);
  const data = await graphGet(instanceId, `${path}?${params.toString()}`);
  return summarizeDriveItem(data);
}

function isLikelyText(contentType: string, name?: string): boolean {
  const lowerType = contentType.toLowerCase();
  if (lowerType.startsWith('text/')) return true;
  if (/json|xml|yaml|csv|javascript|typescript|markdown|x-www-form-urlencoded/.test(lowerType)) {
    return true;
  }
  const lowerName = (name || '').toLowerCase();
  return /\.(txt|md|markdown|json|csv|tsv|xml|yml|yaml|html|css|js|ts|tsx|jsx|py|ps1|sh|log)$/i
    .test(lowerName);
}

export async function readDriveItemText(
  instanceId: string,
  opts: { driveId?: string; itemId?: string; webUrl?: string; maxBytes?: number },
): Promise<{ item?: DriveItemSummary; contentType: string; byteLength: number; truncated: boolean; text?: string; note?: string }> {
  if (!opts.webUrl && !opts.itemId) {
    throw new Error('readDriveItemText: provide itemId or webUrl');
  }
  const maxBytes = clampInt(opts.maxBytes, 200_000, 1_000, 1_000_000);
  const item = await getDriveItem(instanceId, opts);
  const contentPath = opts.webUrl
    ? `/shares/${shareIdFromUrl(opts.webUrl)}/driveItem/content`
    : `${driveItemPath({ driveId: opts.driveId || item.driveId, itemId: opts.itemId || item.id })}/content`;
  const content = await graphGetContent(instanceId, contentPath, maxBytes);
  if (!isLikelyText(content.contentType, item.name)) {
    return {
      item,
      contentType: content.contentType,
      byteLength: content.byteLength,
      truncated: content.truncated,
      note: 'File is not text-like; content was not returned. Use webUrl to open it or ask for a text/CSV/JSON/Markdown file.',
    };
  }
  return {
    item,
    contentType: content.contentType,
    byteLength: content.byteLength,
    truncated: content.truncated,
    text: content.buffer.toString('utf-8'),
  };
}

export async function searchSites(
  instanceId: string,
  opts: { query: string; top?: number },
): Promise<SiteSummary[]> {
  const query = opts.query?.trim();
  if (!query) throw new Error('searchSites: query is required');
  const top = clampInt(opts.top, 10, 1, 50);
  const params = new URLSearchParams();
  params.set('search', query);
  params.set('$select', 'id,name,displayName,webUrl,description');
  const data = await graphGet(instanceId, `/sites?${params.toString()}`);
  return (data.value || []).slice(0, top).map((site: any) => ({
    id: site.id,
    name: site.name,
    displayName: site.displayName,
    webUrl: site.webUrl,
    description: site.description,
  }));
}

export async function listSiteDrives(
  instanceId: string,
  siteId: string,
): Promise<DriveSummary[]> {
  if (!siteId) throw new Error('listSiteDrives: siteId is required');
  const params = new URLSearchParams();
  params.set('$select', 'id,name,webUrl,driveType,createdDateTime,lastModifiedDateTime');
  const data = await graphGet(
    instanceId,
    `/sites/${encodeURIComponent(siteId)}/drives?${params.toString()}`,
  );
  return (data.value || []).map((drive: any) => ({
    id: drive.id,
    name: drive.name,
    webUrl: drive.webUrl,
    driveType: drive.driveType,
    createdDateTime: drive.createdDateTime,
    lastModifiedDateTime: drive.lastModifiedDateTime,
  }));
}

export interface TeamsChatSummary {
  id: string;
  topic?: string;
  chatType?: string;
  createdDateTime?: string;
  lastUpdatedDateTime?: string;
  webUrl?: string;
}

export interface TeamsMessageIdentity {
  type: 'user' | 'application' | 'device' | 'conversation' | 'unknown';
  id?: string;
  displayName?: string;
  userIdentityType?: string;
}

export interface TeamsMessageSummary {
  id: string;
  replyToId?: string;
  messageType?: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  lastEditedDateTime?: string;
  deletedDateTime?: string;
  subject?: string;
  summary?: string;
  importance?: string;
  locale?: string;
  webUrl?: string;
  chatId?: string;
  channelIdentity?: { teamId?: string; channelId?: string };
  from?: TeamsMessageIdentity;
  bodyContentType?: string;
  bodyContent?: string;
  attachments?: Array<{ id?: string; name?: string; contentType?: string; contentUrl?: string }>;
}

export interface TeamSummary {
  id: string;
  displayName?: string;
  description?: string;
  webUrl?: string;
  isArchived?: boolean;
}

export interface TeamsChannelSummary {
  id: string;
  displayName?: string;
  description?: string;
  webUrl?: string;
  membershipType?: string;
  email?: string;
}

function summarizeChat(chat: any): TeamsChatSummary {
  return {
    id: chat.id,
    topic: chat.topic,
    chatType: chat.chatType,
    createdDateTime: chat.createdDateTime,
    lastUpdatedDateTime: chat.lastUpdatedDateTime,
    webUrl: chat.webUrl,
  };
}

function summarizeTeamsIdentity(from: any): TeamsMessageIdentity | undefined {
  if (!from) return undefined;
  for (const type of ['user', 'application', 'device', 'conversation'] as const) {
    const identity = from[type];
    if (identity) {
      return {
        type,
        id: identity.id,
        displayName: identity.displayName,
        userIdentityType: identity.userIdentityType,
      };
    }
  }
  return { type: 'unknown' };
}

function summarizeTeamsMessage(message: any): TeamsMessageSummary {
  return {
    id: message.id,
    replyToId: message.replyToId,
    messageType: message.messageType,
    createdDateTime: message.createdDateTime,
    lastModifiedDateTime: message.lastModifiedDateTime,
    lastEditedDateTime: message.lastEditedDateTime,
    deletedDateTime: message.deletedDateTime,
    subject: message.subject,
    summary: message.summary,
    importance: message.importance,
    locale: message.locale,
    webUrl: message.webUrl,
    chatId: message.chatId,
    channelIdentity: message.channelIdentity
      ? {
        teamId: message.channelIdentity.teamId,
        channelId: message.channelIdentity.channelId,
      }
      : undefined,
    from: summarizeTeamsIdentity(message.from),
    bodyContentType: message.body?.contentType,
    bodyContent: truncateString(message.body?.content),
    attachments: (message.attachments || []).map((attachment: any) => ({
      id: attachment.id,
      name: attachment.name,
      contentType: attachment.contentType,
      contentUrl: attachment.contentUrl,
    })),
  };
}

export async function listTeamsChats(
  instanceId: string,
  opts: { top?: number } = {},
): Promise<TeamsChatSummary[]> {
  const top = clampInt(opts.top, 20, 1, 50);
  const params = new URLSearchParams();
  params.set('$top', String(top));
  params.set('$select', 'id,topic,chatType,createdDateTime,lastUpdatedDateTime,webUrl');
  const data = await graphGet(instanceId, `/me/chats?${params.toString()}`);
  return (data.value || []).map(summarizeChat);
}

export async function listTeamsChatMessages(
  instanceId: string,
  opts: { chatId: string; top?: number },
): Promise<TeamsMessageSummary[]> {
  if (!opts.chatId) throw new Error('listTeamsChatMessages: chatId is required');
  const top = clampInt(opts.top, 25, 1, 50);
  const params = new URLSearchParams();
  params.set('$top', String(top));
  const data = await graphGet(
    instanceId,
    `/chats/${encodeURIComponent(opts.chatId)}/messages?${params.toString()}`,
  );
  return (data.value || []).map(summarizeTeamsMessage);
}

export async function sendTeamsChatMessage(
  instanceId: string,
  opts: { chatId: string; bodyHtml?: string; bodyText?: string; confirm?: boolean },
): Promise<TeamsMessageSummary> {
  if (!opts.chatId) throw new Error('sendTeamsChatMessage: chatId is required');
  requireConfirmed(opts.confirm, 'sendTeamsChatMessage');
  const data = await graphPost(
    instanceId,
    `/chats/${encodeURIComponent(opts.chatId)}/messages`,
    { body: teamsMessageBody(opts) },
  );
  return summarizeTeamsMessage(data);
}

export async function listJoinedTeams(
  instanceId: string,
  opts: { top?: number } = {},
): Promise<TeamSummary[]> {
  const top = clampInt(opts.top, 50, 1, 200);
  const params = new URLSearchParams();
  params.set('$select', 'id,displayName,description,webUrl,isArchived');
  const data = await graphGet(instanceId, `/me/joinedTeams?${params.toString()}`);
  return (data.value || []).slice(0, top).map((team: any) => ({
    id: team.id,
    displayName: team.displayName,
    description: team.description,
    webUrl: team.webUrl,
    isArchived: team.isArchived,
  }));
}

export async function listTeamChannels(
  instanceId: string,
  opts: { teamId: string; top?: number },
): Promise<TeamsChannelSummary[]> {
  if (!opts.teamId) throw new Error('listTeamChannels: teamId is required');
  const top = clampInt(opts.top, 50, 1, 200);
  const params = new URLSearchParams();
  params.set('$top', String(top));
  params.set('$select', 'id,displayName,description,webUrl,membershipType,email');
  const data = await graphGet(
    instanceId,
    `/teams/${encodeURIComponent(opts.teamId)}/channels?${params.toString()}`,
  );
  return (data.value || []).map((channel: any) => ({
    id: channel.id,
    displayName: channel.displayName,
    description: channel.description,
    webUrl: channel.webUrl,
    membershipType: channel.membershipType,
    email: channel.email,
  }));
}

export async function listTeamsChannelMessages(
  instanceId: string,
  opts: { teamId: string; channelId: string; top?: number },
): Promise<TeamsMessageSummary[]> {
  if (!opts.teamId) throw new Error('listTeamsChannelMessages: teamId is required');
  if (!opts.channelId) throw new Error('listTeamsChannelMessages: channelId is required');
  const top = clampInt(opts.top, 25, 1, 50);
  const params = new URLSearchParams();
  params.set('$top', String(top));
  const data = await graphGet(
    instanceId,
    `/teams/${encodeURIComponent(opts.teamId)}/channels/${encodeURIComponent(opts.channelId)}/messages?${params.toString()}`,
  );
  return (data.value || []).map(summarizeTeamsMessage);
}

export async function listTeamsChannelMessageReplies(
  instanceId: string,
  opts: { teamId: string; channelId: string; messageId: string; top?: number },
): Promise<TeamsMessageSummary[]> {
  if (!opts.teamId) throw new Error('listTeamsChannelMessageReplies: teamId is required');
  if (!opts.channelId) throw new Error('listTeamsChannelMessageReplies: channelId is required');
  if (!opts.messageId) throw new Error('listTeamsChannelMessageReplies: messageId is required');
  const top = clampInt(opts.top, 25, 1, 50);
  const params = new URLSearchParams();
  params.set('$top', String(top));
  const data = await graphGet(
    instanceId,
    `/teams/${encodeURIComponent(opts.teamId)}/channels/${encodeURIComponent(opts.channelId)}/messages/${encodeURIComponent(opts.messageId)}/replies?${params.toString()}`,
  );
  return (data.value || []).map(summarizeTeamsMessage);
}

export async function sendTeamsChannelMessage(
  instanceId: string,
  opts: { teamId: string; channelId: string; bodyHtml?: string; bodyText?: string; confirm?: boolean },
): Promise<TeamsMessageSummary> {
  if (!opts.teamId) throw new Error('sendTeamsChannelMessage: teamId is required');
  if (!opts.channelId) throw new Error('sendTeamsChannelMessage: channelId is required');
  requireConfirmed(opts.confirm, 'sendTeamsChannelMessage');
  const data = await graphPost(
    instanceId,
    `/teams/${encodeURIComponent(opts.teamId)}/channels/${encodeURIComponent(opts.channelId)}/messages`,
    { body: teamsMessageBody(opts) },
  );
  return summarizeTeamsMessage(data);
}

export async function sendTeamsChannelReply(
  instanceId: string,
  opts: { teamId: string; channelId: string; messageId: string; bodyHtml?: string; bodyText?: string; confirm?: boolean },
): Promise<TeamsMessageSummary> {
  if (!opts.teamId) throw new Error('sendTeamsChannelReply: teamId is required');
  if (!opts.channelId) throw new Error('sendTeamsChannelReply: channelId is required');
  if (!opts.messageId) throw new Error('sendTeamsChannelReply: messageId is required');
  requireConfirmed(opts.confirm, 'sendTeamsChannelReply');
  const data = await graphPost(
    instanceId,
    `/teams/${encodeURIComponent(opts.teamId)}/channels/${encodeURIComponent(opts.channelId)}/messages/${encodeURIComponent(opts.messageId)}/replies`,
    { body: teamsMessageBody(opts) },
  );
  return summarizeTeamsMessage(data);
}
