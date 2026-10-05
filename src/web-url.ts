/**
 * Fastmail web-app deep links.
 *
 * Fastmail opens a specific email at
 *   https://app.fastmail.com/mail/<mailbox path>/<threadId>.<emailId>?u=<u>
 * where <mailbox path> is the email's *current* mailbox, named by walking the
 * mailbox tree up via parentId and joining names with "." (e.g. "Inbox.Feed").
 * A link built with the wrong mailbox does not open the email, so the path
 * must come from the email's own mailboxIds.
 */

export interface MailboxNode {
  id: string;
  name: string;
  parentId?: string | null;
  role?: string | null;
}

const WEB_BASE = 'https://app.fastmail.com/mail';

/** Dotted path for a mailbox (e.g. "Inbox.Feed"), or null if it can't be resolved. */
export function buildMailboxPath(mailboxId: string, mailboxes: MailboxNode[]): string | null {
  const byId = new Map(mailboxes.map(m => [m.id, m]));
  const names: string[] = [];
  const seen = new Set<string>();
  let current = byId.get(mailboxId);
  if (!current) return null;

  while (current) {
    if (seen.has(current.id)) return null; // parentId cycle
    seen.add(current.id);
    names.unshift(current.name);
    if (!current.parentId) break;
    current = byId.get(current.parentId);
    if (!current) return null; // dangling parent
  }
  return names.join('.');
}

/**
 * Choose which of an email's mailboxes to link through.
 * Rule: the first mailboxId (in the order the server returned mailboxIds)
 * whose mailbox exists in the tree. Simple and deterministic; emails in Feed
 * and Fleet are only ever in one mailbox in practice.
 */
export function pickMailboxPath(
  mailboxIds: Record<string, boolean> | undefined,
  mailboxes: MailboxNode[],
): string | null {
  for (const id of Object.keys(mailboxIds ?? {})) {
    const path = buildMailboxPath(id, mailboxes);
    if (path) return path;
  }
  return null;
}

export function buildWebUrl(args: {
  emailId: string;
  threadId?: string;
  mailboxPath: string | null;
  webUserId: string | null;
}): string {
  const { emailId, threadId, mailboxPath, webUserId } = args;
  if (!mailboxPath || !threadId) {
    // Legacy format: opens Fastmail but not the specific email.
    return `${WEB_BASE}/email/${emailId}`;
  }
  const path = mailboxPath.split('.').map(encodeURIComponent).join('.');
  const url = `${WEB_BASE}/${path}/${threadId}.${emailId}`;
  return webUserId ? `${url}?u=${encodeURIComponent(webUserId)}` : url;
}

/**
 * The web app's `u` query parameter. FASTMAIL_WEB_USER_ID wins if set;
 * otherwise it's the JMAP accountId without its leading "u" (verified
 * against the live account). Returns null (omit ?u) if neither applies.
 */
export function resolveWebUserId(
  accountId: string,
  env: Record<string, string | undefined> = process.env,
): string | null {
  const configured = env.FASTMAIL_WEB_USER_ID?.trim();
  if (configured) return configured;
  const match = /^u([0-9a-f]+)$/i.exec(accountId ?? '');
  return match ? match[1] : null;
}
