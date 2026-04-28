/**
 * Agent 365 channel — WPX (Word/Excel/PowerPoint) comment helpers.
 *
 * Microsoft 365 surfaces in-document comments through the same activity
 * pipeline as messages, decorated with a `WpxComment` entity. The hosting
 * SDK does not (yet) expose a `createWordCommentResponseActivity` analogue
 * to `createEmailResponseActivity`, so this module:
 *
 *   - Extracts a normalized WpxCommentContext from the notification activity.
 *   - Attempts to post the agent's reply as a real comment-thread reply via
 *     Microsoft Graph (using an OBO Graph access token).
 *   - Returns false on any failure so the caller can fall back to a plain
 *     chat reply — never block the user-visible response.
 *
 * Graph endpoints used (per workload):
 *   - Excel:      POST /me/drive/items/{itemId}/workbook/comments/{cid}/replies
 *   - Word/PPT:   not yet stable in v1.0/beta; we attempt a workload-shaped
 *                 best-effort POST and accept failure.
 *
 * Anything we can't post as a comment falls back to the email/chat path.
 */

import { logger } from '../../logger.js';
import type {
  WpxCommentContext,
} from '../../agent365/mcp-context.js';
import type { AgentNotificationActivity } from '@microsoft/agents-a365-notifications';

const WORD_HINT       = /word/i;
const EXCEL_HINT      = /excel|xls/i;
const POWERPOINT_HINT = /powerpoint|ppt/i;

/**
 * Pull a workload-tagged WpxCommentContext out of a notification activity,
 * or return undefined when the activity doesn't carry a comment.
 *
 * Workload is inferred from the channelData / valueType when present,
 * because the SDK shape doesn't carry it on WpxComment itself.
 */
export function extractWpxComment(
  notif: AgentNotificationActivity,
  workloadHint: 'word' | 'excel' | 'powerpoint',
): WpxCommentContext | undefined {
  const comment = notif.wpxCommentNotification;
  if (!comment) return undefined;
  return {
    workload: detectWorkload(notif, workloadHint),
    documentId: comment.documentId,
    initiatingCommentId: comment.initiatingCommentId,
    subjectCommentId: comment.subjectCommentId,
  };
}

function detectWorkload(
  notif: AgentNotificationActivity,
  hint: 'word' | 'excel' | 'powerpoint',
): 'word' | 'excel' | 'powerpoint' {
  const probe = `${notif.valueType || ''} ${JSON.stringify(notif.channelData || '')}`;
  if (POWERPOINT_HINT.test(probe)) return 'powerpoint';
  if (EXCEL_HINT.test(probe)) return 'excel';
  if (WORD_HINT.test(probe)) return 'word';
  return hint;
}

export interface ReplyAsCommentArgs {
  comment: WpxCommentContext;
  graphToken: string;
  htmlBody: string;
}

/**
 * Try to post `htmlBody` as a reply on the comment thread. Returns true on
 * a 2xx response, false otherwise. Never throws.
 */
export async function replyAsComment(args: ReplyAsCommentArgs): Promise<boolean> {
  const { comment, graphToken, htmlBody } = args;
  if (!comment.documentId || !comment.subjectCommentId) {
    logger.debug({ comment }, 'Agent 365: WPX reply missing IDs, falling back');
    return false;
  }

  const url = buildGraphReplyUrl(comment);
  if (!url) return false;

  // Graph comment-reply payload shape varies slightly per workload, but
  // {content, contentType} is accepted on the workbook endpoint and a
  // close fit for the others. We send HTML when available so formatting
  // round-trips when the workload supports it.
  const body = JSON.stringify({
    content: htmlBody,
    contentType: 'html',
  });

  try {
    const res = await fetch(url, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${graphToken}`,
        'Content-Type': 'application/json',
      },
      body,
    });
    if (res.ok) {
      logger.info(
        { workload: comment.workload, status: res.status },
        'Agent 365: posted WPX comment reply',
      );
      return true;
    }
    const detail = await safeReadText(res);
    logger.warn(
      { workload: comment.workload, status: res.status, detail },
      'Agent 365: WPX comment reply rejected, falling back',
    );
    return false;
  } catch (err) {
    logger.warn({ err, workload: comment.workload }, 'Agent 365: WPX comment reply failed');
    return false;
  }
}

function buildGraphReplyUrl(comment: WpxCommentContext): string | undefined {
  const item = encodeURIComponent(comment.documentId!);
  const cid = encodeURIComponent(comment.subjectCommentId!);
  switch (comment.workload) {
    case 'excel':
      return `https://graph.microsoft.com/v1.0/me/drive/items/${item}/workbook/comments/${cid}/replies`;
    case 'word':
    case 'powerpoint':
      // Word/PowerPoint comment replies are a beta surface; attempt the
      // beta endpoint and accept that it may return 404 in tenants where
      // it isn't enabled. Caller falls back to chat reply on failure.
      return `https://graph.microsoft.com/beta/me/drive/items/${item}/comments/${cid}/replies`;
    default:
      return undefined;
  }
}

async function safeReadText(res: Response): Promise<string> {
  try {
    return (await res.text()).slice(0, 500);
  } catch {
    return '';
  }
}
