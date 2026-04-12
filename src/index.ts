#!/usr/bin/env node
import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  CallToolRequestSchema,
  ErrorCode,
  ListToolsRequestSchema,
  McpError,
} from '@modelcontextprotocol/sdk/types.js';
import { FastmailAuth, FastmailConfig } from './auth.js';
import { JmapClient } from './jmap-client.js';
import { ContactsCalendarClient } from './contacts-calendar.js';
import { CalDAVCalendarClient } from './caldav-client.js';

const server = new Server(
  {
    name: 'fastmail-mcp',
    version: '1.8.2',
  },
  {
    capabilities: {
      tools: {},
    },
  }
);

let jmapClient: JmapClient | null = null;
let contactsCalendarClient: ContactsCalendarClient | null = null;
let caldavClient: CalDAVCalendarClient | null = null;

function findEnvValue(keys: string[]): { value?: string; key?: string; wasPlaceholder: boolean } {
  const isPlaceholder = (val: string) => /\$\{[^}]+\}/.test(val.trim());
  for (const key of keys) {
    const raw = process.env[key];
    if (typeof raw === 'string' && raw.trim().length > 0) {
      if (isPlaceholder(raw)) {
        return { value: undefined, key, wasPlaceholder: true };
      }
      return { value: raw.trim(), key, wasPlaceholder: false };
    }
  }
  return { value: undefined, key: undefined, wasPlaceholder: false };
}

function maskSecret(value: string): string {
  if (value.length <= 6) return '***';
  return `${value.slice(0, 4)}…${value.slice(-2)} (len ${value.length})`;
}

function getAuthConfig(): FastmailConfig {
  const tokenInfo = findEnvValue([
    'FASTMAIL_API_TOKEN',
    'USER_CONFIG_FASTMAIL_API_TOKEN',
    'USER_CONFIG_fastmail_api_token',
    'fastmail_api_token',
  ]);
  const apiToken = tokenInfo.value;
  if (!apiToken) {
    throw new McpError(
      ErrorCode.InvalidRequest,
      'FASTMAIL_API_TOKEN environment variable is required'
    );
  }

  const baseInfo = findEnvValue([
    'FASTMAIL_BASE_URL',
    'USER_CONFIG_FASTMAIL_BASE_URL',
    'USER_CONFIG_fastmail_base_url',
    'fastmail_base_url',
  ]);

  return { apiToken, baseUrl: baseInfo.value };
}

function initializeClient(): JmapClient {
  if (jmapClient) {
    return jmapClient;
  }

  const auth = new FastmailAuth(getAuthConfig());
  jmapClient = new JmapClient(auth);
  return jmapClient;
}

function initializeContactsCalendarClient(): ContactsCalendarClient {
  if (contactsCalendarClient) {
    return contactsCalendarClient;
  }

  const auth = new FastmailAuth(getAuthConfig());
  contactsCalendarClient = new ContactsCalendarClient(auth);
  return contactsCalendarClient;
}

function initializeCalDAVClient(): CalDAVCalendarClient | null {
  if (caldavClient) return caldavClient;

  const username = findEnvValue([
    'FASTMAIL_CALDAV_USERNAME',
    'USER_CONFIG_FASTMAIL_CALDAV_USERNAME',
  ]).value;
  const password = findEnvValue([
    'FASTMAIL_CALDAV_PASSWORD',
    'USER_CONFIG_FASTMAIL_CALDAV_PASSWORD',
  ]).value;

  if (!username || !password) return null;

  caldavClient = new CalDAVCalendarClient({ username, password });
  return caldavClient;
}

server.setRequestHandler(ListToolsRequestSchema, async () => {
  return {
    tools: [
      {
        name: 'list_mailboxes',
        description: 'List all mailboxes in the Fastmail account',
        inputSchema: {
          type: 'object',
          properties: {},
        },
      },
      {
        name: 'list_emails',
        description: 'List emails from a mailbox',
        inputSchema: {
          type: 'object',
          properties: {
            mailboxId: {
              type: 'string',
              description: 'ID of the mailbox to list emails from (optional, defaults to all)',
            },
            limit: {
              type: 'number',
              description: 'Maximum number of emails to return (default: 20)',
              default: 20,
            },
          },
        },
      },
      {
        name: 'get_email',
        description: 'Get a specific email by ID',
        inputSchema: {
          type: 'object',
          properties: {
            emailId: {
              type: 'string',
              description: 'ID of the email to retrieve',
            },
          },
          required: ['emailId'],
        },
      },
      {
        name: 'get_emails',
        description: 'Batch-fetch multiple emails by ID in a single call. Prefer this over looping get_email when you need more than one.',
        inputSchema: {
          type: 'object',
          properties: {
            emailIds: {
              type: 'array',
              items: { type: 'string' },
              description: 'IDs of the emails to retrieve. Missing IDs are silently omitted from the response.',
            },
          },
          required: ['emailIds'],
        },
      },
      {
        name: 'send_email',
        description: 'Send an email',
        inputSchema: {
          type: 'object',
          properties: {
            to: {
              type: 'array',
              items: { type: 'string' },
              description: 'Recipient email addresses',
            },
            cc: {
              type: 'array',
              items: { type: 'string' },
              description: 'CC email addresses (optional)',
            },
            bcc: {
              type: 'array',
              items: { type: 'string' },
              description: 'BCC email addresses (optional)',
            },
            from: {
              type: 'string',
              description: 'Sender email address (optional, defaults to account primary email)',
            },
            mailboxId: {
              type: 'string',
              description: 'Mailbox ID to save the email to (optional, defaults to Drafts folder)',
            },
            subject: {
              type: 'string',
              description: 'Email subject',
            },
            textBody: {
              type: 'string',
              description: 'Plain text body (optional)',
            },
            htmlBody: {
              type: 'string',
              description: 'HTML body (optional)',
            },
            inReplyTo: {
              type: 'array',
              items: { type: 'string' },
              description: 'Message-ID(s) of the email being replied to (optional, for threading)',
            },
            references: {
              type: 'array',
              items: { type: 'string' },
              description: 'Full reference chain of Message-IDs (optional, for threading)',
            },
          },
          required: ['to', 'subject'],
        },
      },
      {
        name: 'reply_email',
        description: 'Reply to an existing email with proper threading headers (In-Reply-To, References). Automatically fetches the original email to build the reply chain. By default sends immediately; set send=false to save as a draft instead.',
        inputSchema: {
          type: 'object',
          properties: {
            originalEmailId: {
              type: 'string',
              description: 'ID of the email to reply to',
            },
            to: {
              type: 'array',
              items: { type: 'string' },
              description: 'Recipient email addresses (optional, defaults to the original sender)',
            },
            cc: {
              type: 'array',
              items: { type: 'string' },
              description: 'CC email addresses (optional)',
            },
            bcc: {
              type: 'array',
              items: { type: 'string' },
              description: 'BCC email addresses (optional)',
            },
            from: {
              type: 'string',
              description: 'Sender email address (optional, defaults to account primary email)',
            },
            textBody: {
              type: 'string',
              description: 'Plain text body (optional)',
            },
            htmlBody: {
              type: 'string',
              description: 'HTML body (optional)',
            },
            send: {
              type: 'boolean',
              description: 'Whether to send the reply immediately (default: true). Set to false to save as draft instead.',
            },
          },
          required: ['originalEmailId'],
        },
      },
      {
        name: 'create_draft',
        description: 'Create an email draft without sending it. Supports threading headers for replies. IMPORTANT: each call creates a new draft — do not call twice for the same message.',
        inputSchema: {
          type: 'object',
          properties: {
            to: {
              type: 'array',
              items: { type: 'string' },
              description: 'Recipient email addresses (optional)',
            },
            cc: {
              type: 'array',
              items: { type: 'string' },
              description: 'CC email addresses (optional)',
            },
            bcc: {
              type: 'array',
              items: { type: 'string' },
              description: 'BCC email addresses (optional)',
            },
            from: {
              type: 'string',
              description: 'Sender email address (optional, defaults to account primary email)',
            },
            mailboxId: {
              type: 'string',
              description: 'Mailbox ID to save the draft to (optional, defaults to Drafts folder)',
            },
            subject: {
              type: 'string',
              description: 'Email subject (optional)',
            },
            textBody: {
              type: 'string',
              description: 'Plain text body (optional)',
            },
            htmlBody: {
              type: 'string',
              description: 'HTML body (optional)',
            },
            inReplyTo: {
              type: 'array',
              items: { type: 'string' },
              description: 'Message-IDs to reply to (optional, for threading)',
            },
            references: {
              type: 'array',
              items: { type: 'string' },
              description: 'Message-IDs for References header (optional, for threading)',
            },
          },
        },
      },
      {
        name: 'edit_draft',
        description: 'Edit an existing draft email. Since JMAP emails are immutable, this atomically destroys the old draft and creates a new one with the updated fields. Only fields you provide will be changed; others are preserved from the original draft.',
        inputSchema: {
          type: 'object',
          properties: {
            emailId: {
              type: 'string',
              description: 'The ID of the draft email to edit',
            },
            to: {
              type: 'array',
              items: { type: 'string' },
              description: 'Updated recipient email addresses (optional, keeps existing if omitted)',
            },
            cc: {
              type: 'array',
              items: { type: 'string' },
              description: 'Updated CC email addresses (optional)',
            },
            bcc: {
              type: 'array',
              items: { type: 'string' },
              description: 'Updated BCC email addresses (optional)',
            },
            from: {
              type: 'string',
              description: 'Updated sender email address (optional)',
            },
            subject: {
              type: 'string',
              description: 'Updated email subject (optional)',
            },
            textBody: {
              type: 'string',
              description: 'Updated plain text body (optional)',
            },
            htmlBody: {
              type: 'string',
              description: 'Updated HTML body (optional)',
            },
          },
          required: ['emailId'],
        },
      },
      {
        name: 'send_draft',
        description: 'Send an existing draft email. The draft must have recipients (to/cc/bcc) and a from address. After sending, the email is moved to the Sent folder and the draft keyword is removed.',
        inputSchema: {
          type: 'object',
          properties: {
            emailId: {
              type: 'string',
              description: 'The ID of the draft email to send',
            },
          },
          required: ['emailId'],
        },
      },
      {
        name: 'search_emails',
        description: 'Search emails by subject or content',
        inputSchema: {
          type: 'object',
          properties: {
            query: {
              type: 'string',
              description: 'Search query string',
            },
            limit: {
              type: 'number',
              description: 'Maximum number of results (default: 20)',
              default: 20,
            },
          },
          required: ['query'],
        },
      },
      {
        name: 'list_contacts',
        description: 'List contacts from the address book',
        inputSchema: {
          type: 'object',
          properties: {
            limit: {
              type: 'number',
              description: 'Maximum number of contacts to return (default: 50)',
              default: 50,
            },
          },
        },
      },
      {
        name: 'get_contact',
        description: 'Get a specific contact by ID',
        inputSchema: {
          type: 'object',
          properties: {
            contactId: {
              type: 'string',
              description: 'ID of the contact to retrieve',
            },
          },
          required: ['contactId'],
        },
      },
      {
        name: 'search_contacts',
        description: 'Search contacts by name or email',
        inputSchema: {
          type: 'object',
          properties: {
            query: {
              type: 'string',
              description: 'Search query string',
            },
            limit: {
              type: 'number',
              description: 'Maximum number of results (default: 20)',
              default: 20,
            },
          },
          required: ['query'],
        },
      },
      {
        name: 'list_calendars',
        description: 'List all calendars',
        inputSchema: {
          type: 'object',
          properties: {},
        },
      },
      {
        name: 'list_calendar_events',
        description: 'List events from a calendar, optionally filtered by date range',
        inputSchema: {
          type: 'object',
          properties: {
            calendarId: {
              type: 'string',
              description: 'ID of the calendar (optional, defaults to all calendars)',
            },
            limit: {
              type: 'number',
              description: 'Maximum number of events to return (default: 50)',
              default: 50,
            },
            timeMin: {
              type: 'string',
              description: 'Lower bound for events (ISO 8601 UTC, e.g. 2026-03-23T00:00:00Z). Only events ending after this time are returned.',
            },
            timeMax: {
              type: 'string',
              description: 'Upper bound for events (ISO 8601 UTC, e.g. 2026-03-23T23:59:59Z). Only events starting before this time are returned.',
            },
          },
        },
      },
      {
        name: 'get_calendar_event',
        description: 'Get a specific calendar event by ID',
        inputSchema: {
          type: 'object',
          properties: {
            eventId: {
              type: 'string',
              description: 'ID of the event to retrieve',
            },
          },
          required: ['eventId'],
        },
      },
      {
        name: 'create_calendar_event',
        description: 'Create a new calendar event',
        inputSchema: {
          type: 'object',
          properties: {
            calendarId: {
              type: 'string',
              description: 'Calendar to create the event in. Accepts a JMAP calendar ID, a calendar name (case-insensitive, e.g. "Paul"), or a CalDAV URL. Prefer the name.',
            },
            title: {
              type: 'string',
              description: 'Event title',
            },
            description: {
              type: 'string',
              description: 'Event description (optional)',
            },
            start: {
              type: 'string',
              description: 'Start time. Accepts bare local (2026-03-25T14:00), ISO with Z (2026-03-25T03:00:00Z), ISO with offset (2026-03-25T14:00:00+11:00), or all-day date (2026-03-25). Bare local is interpreted in FASTMAIL_TIMEZONE (default: system timezone).',
            },
            end: {
              type: 'string',
              description: 'End time. Same formats as start.',
            },
            location: {
              type: 'string',
              description: 'Event location (optional)',
            },
            participants: {
              type: 'array',
              items: {
                type: 'object',
                properties: {
                  email: { type: 'string' },
                  name: { type: 'string' }
                }
              },
              description: 'Event participants (optional)',
            },
            timezone: {
              type: 'string',
              description: 'IANA timezone name used to interpret bare-local start/end (e.g. "Australia/Sydney"). Optional — defaults to FASTMAIL_TIMEZONE env var, then system timezone. Ignored for inputs that already specify UTC (Z) or an offset.',
            },
          },
          required: ['calendarId', 'title', 'start', 'end'],
        },
      },
      {
        name: 'find_duplicate_event',
        description: 'Check whether a calendar event with a similar title already exists within the given date window. Use before create_calendar_event to avoid duplicates from email re-processing or overlapping triage flows.',
        inputSchema: {
          type: 'object',
          properties: {
            title: {
              type: 'string',
              description: 'Event title to match (case-insensitive, substring match).',
            },
            start: {
              type: 'string',
              description: 'Start time of the proposed event. Same formats as create_calendar_event.',
            },
            end: {
              type: 'string',
              description: 'End time (optional). Defaults to start.',
            },
            calendarNames: {
              type: 'array',
              items: { type: 'string' },
              description: 'Calendar names or IDs to search. Omit to search all calendars.',
            },
            timezone: {
              type: 'string',
              description: 'IANA timezone for bare-local start/end (defaults to FASTMAIL_TIMEZONE then system).',
            },
          },
          required: ['title', 'start'],
        },
      },
      {
        name: 'delete_calendar_event',
        description: 'Delete a calendar event. `scope` and `instanceStart` are both optional: if omitted, the MCP fetches the event and picks sensible defaults — `all` for non-recurring events, `this` for recurring events (with `instanceStart` inferred from the event\'s start). Override either explicitly when you need the other behaviour (e.g. scope="all" on a recurring event to delete the whole series).',
        inputSchema: {
          type: 'object',
          properties: {
            eventId: {
              type: 'string',
              description: 'ID of the event to delete',
            },
            scope: {
              type: 'string',
              enum: ['this', 'all'],
              description: 'Optional. "this" to delete only the specified occurrence; "all" to delete the whole event/series. Auto-detected when omitted.',
            },
            instanceStart: {
              type: 'string',
              description: 'Optional. ISO 8601 start time of the occurrence to delete. Inferred from the event when scope is "this" and this is omitted.',
            },
          },
          required: ['eventId'],
        },
      },
      {
        name: 'move_calendar_event',
        description: 'Move a calendar event to a different calendar. `scope` and `instanceStart` are both optional: if omitted, the MCP fetches the event and picks sensible defaults — `all` for non-recurring events, `this` for recurring events (with `instanceStart` inferred from the event\'s start). Override either explicitly when you need the other behaviour.',
        inputSchema: {
          type: 'object',
          properties: {
            eventId: {
              type: 'string',
              description: 'ID of the event to move',
            },
            targetCalendarId: {
              type: 'string',
              description: 'ID of the target calendar (accepts calendar name or JMAP ID)',
            },
            scope: {
              type: 'string',
              enum: ['this', 'all'],
              description: 'Optional. "this" to move only the specified occurrence; "all" to move the whole event/series. Auto-detected when omitted.',
            },
            instanceStart: {
              type: 'string',
              description: 'Optional. ISO 8601 start time of the occurrence to move. Inferred from the event when scope is "this" and this is omitted.',
            },
          },
          required: ['eventId', 'targetCalendarId'],
        },
      },
      {
        name: 'get_recent_emails',
        description: 'Get the most recent emails from inbox (like top-ten)',
        inputSchema: {
          type: 'object',
          properties: {
            limit: {
              type: 'number',
              description: 'Number of recent emails to retrieve (default: 10, max: 50)',
              default: 10,
            },
            mailboxName: {
              type: 'string',
              description: 'Mailbox to search (default: inbox)',
              default: 'inbox',
            },
          },
        },
      },
      {
        name: 'mark_email_read',
        description: 'Mark an email as read or unread',
        inputSchema: {
          type: 'object',
          properties: {
            emailId: {
              type: 'string',
              description: 'ID of the email to mark',
            },
            read: {
              type: 'boolean',
              description: 'true to mark as read, false to mark as unread',
              default: true,
            },
          },
          required: ['emailId'],
        },
      },
      {
        name: 'pin_email',
        description: 'Pin or unpin an email',
        inputSchema: {
          type: 'object',
          properties: {
            emailId: {
              type: 'string',
              description: 'ID of the email to pin/unpin',
            },
            pinned: {
              type: 'boolean',
              description: 'true to pin, false to unpin',
              default: true,
            },
          },
          required: ['emailId'],
        },
      },
      {
        name: 'delete_email',
        description: 'Delete an email (move to trash)',
        inputSchema: {
          type: 'object',
          properties: {
            emailId: {
              type: 'string',
              description: 'ID of the email to delete',
            },
          },
          required: ['emailId'],
        },
      },
      {
        name: 'move_email',
        description: 'Move an email to a different mailbox',
        inputSchema: {
          type: 'object',
          properties: {
            emailId: {
              type: 'string',
              description: 'ID of the email to move',
            },
            targetMailboxId: {
              type: 'string',
              description: 'ID of the target mailbox',
            },
          },
          required: ['emailId', 'targetMailboxId'],
        },
      },
      {
        name: 'add_labels',
        description: 'Add labels (mailboxes) to an email without removing existing ones',
        inputSchema: {
          type: 'object',
          properties: {
            emailId: {
              type: 'string',
              description: 'ID of the email to add labels to',
            },
            mailboxIds: {
              type: 'array',
              items: { type: 'string' },
              description: 'Array of mailbox IDs to add as labels',
            },
          },
          required: ['emailId', 'mailboxIds'],
        },
      },
      {
        name: 'remove_labels',
        description: 'Remove specific labels (mailboxes) from an email',
        inputSchema: {
          type: 'object',
          properties: {
            emailId: {
              type: 'string',
              description: 'ID of the email to remove labels from',
            },
            mailboxIds: {
              type: 'array',
              items: { type: 'string' },
              description: 'Array of mailbox IDs to remove as labels',
            },
          },
          required: ['emailId', 'mailboxIds'],
        },
      },
      {
        name: 'advanced_search',
        description: 'Advanced email search with multiple criteria',
        inputSchema: {
          type: 'object',
          properties: {
            query: {
              type: 'string',
              description: 'Text to search for in subject/body',
            },
            from: {
              type: 'string',
              description: 'Filter by sender email',
            },
            to: {
              type: 'string',
              description: 'Filter by recipient email',
            },
            subject: {
              type: 'string',
              description: 'Filter by subject',
            },
            hasAttachment: {
              type: 'boolean',
              description: 'Filter emails with attachments',
            },
            isUnread: {
              type: 'boolean',
              description: 'Filter unread emails',
            },
            isPinned: {
              type: 'boolean',
              description: 'Filter pinned emails',
            },
            mailboxId: {
              type: 'string',
              description: 'Search within a specific mailbox. Accepts a JMAP mailbox ID, a name (case-insensitive, e.g. "Fleet"), or a role ("inbox", "archive"). Prefer the name.',
            },
            after: {
              type: 'string',
              description: 'Emails after this date (ISO 8601). Raw — does not compensate for Fastmail search index lag. Prefer `since` unless you specifically need exact boundary semantics.',
            },
            since: {
              type: 'string',
              description: 'Emails since this timestamp (ISO 8601). Internally subtracts 2 hours to work around Fastmail search index lag, so recent emails are not missed when paging forward by a `last_run` marker. This is the right choice for incremental syncs.',
            },
            before: {
              type: 'string',
              description: 'Emails before this date (ISO 8601)',
            },
            limit: {
              type: 'number',
              description: 'Maximum results (default: 50)',
              default: 50,
            },
          },
        },
      },
      {
        name: 'get_thread',
        description: 'Get all emails in a conversation thread',
        inputSchema: {
          type: 'object',
          properties: {
            threadId: {
              type: 'string',
              description: 'ID of the thread/conversation',
            },
          },
          required: ['threadId'],
        },
      },
      {
        name: 'get_mailbox_stats',
        description: 'Get statistics for a mailbox (unread count, total emails, etc.)',
        inputSchema: {
          type: 'object',
          properties: {
            mailboxId: {
              type: 'string',
              description: 'ID of the mailbox (optional, defaults to all mailboxes)',
            },
          },
        },
      },
      {
        name: 'get_account_summary',
        description: 'Get overall account summary with statistics',
        inputSchema: {
          type: 'object',
          properties: {},
        },
      },
      {
        name: 'bulk_mark_read',
        description: 'Mark multiple emails as read/unread',
        inputSchema: {
          type: 'object',
          properties: {
            emailIds: {
              type: 'array',
              items: { type: 'string' },
              description: 'Array of email IDs to mark',
            },
            read: {
              type: 'boolean',
              description: 'true to mark as read, false as unread',
              default: true,
            },
          },
          required: ['emailIds'],
        },
      },
      {
        name: 'bulk_pin',
        description: 'Pin or unpin multiple emails',
        inputSchema: {
          type: 'object',
          properties: {
            emailIds: {
              type: 'array',
              items: { type: 'string' },
              description: 'Array of email IDs to pin/unpin',
            },
            pinned: {
              type: 'boolean',
              description: 'true to pin, false to unpin',
              default: true,
            },
          },
          required: ['emailIds'],
        },
      },
      {
        name: 'bulk_move',
        description: 'Move multiple emails to a mailbox',
        inputSchema: {
          type: 'object',
          properties: {
            emailIds: {
              type: 'array',
              items: { type: 'string' },
              description: 'Array of email IDs to move',
            },
            targetMailboxId: {
              type: 'string',
              description: 'ID of target mailbox',
            },
          },
          required: ['emailIds', 'targetMailboxId'],
        },
      },
      {
        name: 'bulk_delete',
        description: 'Delete multiple emails (move to trash)',
        inputSchema: {
          type: 'object',
          properties: {
            emailIds: {
              type: 'array',
              items: { type: 'string' },
              description: 'Array of email IDs to delete',
            },
          },
          required: ['emailIds'],
        },
      },
      {
        name: 'bulk_add_labels',
        description: 'Add labels to multiple emails simultaneously',
        inputSchema: {
          type: 'object',
          properties: {
            emailIds: {
              type: 'array',
              items: { type: 'string' },
              description: 'Array of email IDs to add labels to',
            },
            mailboxIds: {
              type: 'array',
              items: { type: 'string' },
              description: 'Array of mailbox IDs to add as labels',
            },
          },
          required: ['emailIds', 'mailboxIds'],
        },
      },
      {
        name: 'bulk_remove_labels',
        description: 'Remove labels from multiple emails simultaneously',
        inputSchema: {
          type: 'object',
          properties: {
            emailIds: {
              type: 'array',
              items: { type: 'string' },
              description: 'Array of email IDs to remove labels from',
            },
            mailboxIds: {
              type: 'array',
              items: { type: 'string' },
              description: 'Array of mailbox IDs to remove as labels',
            },
          },
          required: ['emailIds', 'mailboxIds'],
        },
      },
    ],
  };
});

server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args } = request.params;

  try {

    const client = initializeClient();

    switch (name) {
      case 'list_mailboxes': {
        const mailboxes = await client.getMailboxes();
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(mailboxes, null, 2),
            },
          ],
        };
      }

      case 'list_emails': {
        const { mailboxId, limit } = args as any;
        const validLimit = Math.min(Math.max(Number(limit) || 20, 1), 50);
        const emails = await client.getEmails(mailboxId, validLimit);
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(emails, null, 2),
            },
          ],
        };
      }

      case 'get_email': {
        const { emailId } = args as any;
        if (!emailId) {
          throw new McpError(ErrorCode.InvalidParams, 'emailId is required');
        }
        const email = await client.getEmailById(emailId);
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(email, null, 2),
            },
          ],
        };
      }

      case 'get_emails': {
        const { emailIds } = args as any;
        if (!emailIds || !Array.isArray(emailIds) || emailIds.length === 0) {
          throw new McpError(ErrorCode.InvalidParams, 'emailIds must be a non-empty array');
        }
        const emails = await client.getEmailsByIds(emailIds);
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(emails, null, 2),
            },
          ],
        };
      }

      case 'send_email': {
        const { to, cc, bcc, from, mailboxId, subject, textBody, htmlBody, inReplyTo, references } = args as any;
        if (!to || !Array.isArray(to) || to.length === 0) {
          throw new McpError(ErrorCode.InvalidParams, 'to field is required and must be a non-empty array');
        }
        if (!subject) {
          throw new McpError(ErrorCode.InvalidParams, 'subject is required');
        }
        if (!textBody && !htmlBody) {
          throw new McpError(ErrorCode.InvalidParams, 'Either textBody or htmlBody is required');
        }

        const submissionId = await client.sendEmail({
          to,
          cc,
          bcc,
          from,
          mailboxId,
          subject,
          textBody,
          htmlBody,
          inReplyTo,
          references,
        });

        return {
          content: [
            {
              type: 'text',
              text: `Email sent successfully. Submission ID: ${submissionId}`,
            },
          ],
        };
      }

      case 'reply_email': {
        const { originalEmailId, to, cc, bcc, from, textBody, htmlBody, send: shouldSend = true } = args as any;
        if (!originalEmailId) {
          throw new McpError(ErrorCode.InvalidParams, 'originalEmailId is required');
        }
        if (shouldSend && !textBody && !htmlBody) {
          throw new McpError(ErrorCode.InvalidParams, 'Either textBody or htmlBody is required');
        }

        // Fetch the original email to get threading headers
        const originalEmail = await client.getEmailById(originalEmailId);

        // Build threading headers
        const originalMessageId = originalEmail.messageId?.[0];
        if (!originalMessageId) {
          throw new McpError(ErrorCode.InternalError, 'Original email does not have a Message-ID; cannot thread reply');
        }

        const inReplyToHeader = [originalMessageId];
        const referencesHeader = [
          ...(originalEmail.references || []),
          originalMessageId,
        ];

        // Build subject with Re: prefix
        let replySubject = originalEmail.subject || '';
        if (!/^Re:/i.test(replySubject)) {
          replySubject = `Re: ${replySubject}`;
        }

        // Default recipients to the original sender
        const replyTo = (to && Array.isArray(to) && to.length > 0)
          ? to
          : (Array.isArray(originalEmail.from) ? originalEmail.from.map((addr: any) => addr.email).filter(Boolean) : []);

        if (replyTo.length === 0) {
          throw new McpError(ErrorCode.InvalidParams, 'Could not determine reply recipient. Please provide "to" explicitly.');
        }

        const replyParams = {
          to: replyTo,
          cc,
          bcc,
          from,
          subject: replySubject,
          textBody,
          htmlBody,
          inReplyTo: inReplyToHeader,
          references: referencesHeader,
        };

        if (!shouldSend) {
          const emailId = await client.createDraft(replyParams);
          return {
            content: [
              {
                type: 'text',
                text: `Reply draft saved successfully (Email ID: ${emailId}). Subject: ${replySubject}`,
              },
            ],
          };
        }

        const submissionId = await client.sendEmail(replyParams);

        return {
          content: [
            {
              type: 'text',
              text: `Reply sent successfully. Submission ID: ${submissionId}`,
            },
          ],
        };
      }

      case 'create_draft': {
        const { to, cc, bcc, from, mailboxId, subject, textBody, htmlBody, inReplyTo, references } = args as any;

        if (!to?.length && !subject && !textBody && !htmlBody) {
          throw new McpError(ErrorCode.InvalidParams, 'At least one of to, subject, textBody, or htmlBody must be provided');
        }

        const emailId = await client.createDraft({
          to,
          cc,
          bcc,
          from,
          mailboxId,
          subject,
          textBody,
          htmlBody,
          inReplyTo,
          references,
        });

        const summary = [
          `Draft created successfully (Email ID: ${emailId}).`,
          subject ? `Subject: ${subject}` : null,
          to?.length ? `To: ${to.join(', ')}` : null,
          cc?.length ? `CC: ${cc.join(', ')}` : null,
        ].filter(Boolean).join(' ');

        return {
          content: [
            {
              type: 'text',
              text: summary,
            },
          ],
        };
      }

      case 'edit_draft': {
        const { emailId, to, cc, bcc, from, subject, textBody, htmlBody } = args as any;
        if (!emailId) {
          throw new McpError(ErrorCode.InvalidParams, 'emailId is required');
        }

        const newEmailId = await client.updateDraft(emailId, {
          to,
          cc,
          bcc,
          from,
          subject,
          textBody,
          htmlBody,
        });

        return {
          content: [
            {
              type: 'text',
              text: `Draft updated successfully. New Email ID: ${newEmailId} (old draft ${emailId} was replaced)`,
            },
          ],
        };
      }

      case 'send_draft': {
        const { emailId } = args as any;
        if (!emailId) {
          throw new McpError(ErrorCode.InvalidParams, 'emailId is required');
        }

        const submissionId = await client.sendDraft(emailId);

        return {
          content: [
            {
              type: 'text',
              text: `Draft sent successfully. Submission ID: ${submissionId}`,
            },
          ],
        };
      }

      case 'search_emails': {
        const { query, limit = 20 } = args as any;
        if (!query) {
          throw new McpError(ErrorCode.InvalidParams, 'query is required');
        }
        const emails = await client.searchEmails(query, limit);
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(emails, null, 2),
            },
          ],
        };
      }

      case 'list_contacts': {
        const { limit = 50 } = args as any;
        const contactsClient = initializeContactsCalendarClient();
        const contacts = await contactsClient.getContacts(limit);
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(contacts, null, 2),
            },
          ],
        };
      }

      case 'get_contact': {
        const { contactId } = args as any;
        if (!contactId) {
          throw new McpError(ErrorCode.InvalidParams, 'contactId is required');
        }
        const contactsClient = initializeContactsCalendarClient();
        const contact = await contactsClient.getContactById(contactId);
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(contact, null, 2),
            },
          ],
        };
      }

      case 'search_contacts': {
        const { query, limit = 20 } = args as any;
        if (!query) {
          throw new McpError(ErrorCode.InvalidParams, 'query is required');
        }
        const contactsClient = initializeContactsCalendarClient();
        const contacts = await contactsClient.searchContacts(query, limit);
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(contacts, null, 2),
            },
          ],
        };
      }

      case 'list_calendars': {
        try {
          const contactsClient = initializeContactsCalendarClient();
          const calendars = await contactsClient.getCalendars();
          return { content: [{ type: 'text', text: JSON.stringify(calendars, null, 2) }] };
        } catch {
          // JMAP calendars not available, try CalDAV
          const davClient = initializeCalDAVClient();
          if (!davClient) {
            throw new McpError(ErrorCode.InvalidRequest, 'JMAP calendars not available and CalDAV not configured. Set FASTMAIL_CALDAV_USERNAME and FASTMAIL_CALDAV_PASSWORD to use CalDAV.');
          }
          const calendars = await davClient.getCalendars();
          return { content: [{ type: 'text', text: JSON.stringify(calendars, null, 2) }] };
        }
      }

      case 'list_calendar_events': {
        const { calendarId, limit = 50, timeMin, timeMax } = args as any;
        try {
          const contactsClient = initializeContactsCalendarClient();
          const events = await contactsClient.getCalendarEvents(calendarId, limit);
          return { content: [{ type: 'text', text: JSON.stringify(events, null, 2) }] };
        } catch {
          const davClient = initializeCalDAVClient();
          if (!davClient) {
            throw new McpError(ErrorCode.InvalidRequest, 'JMAP calendars not available and CalDAV not configured. Set FASTMAIL_CALDAV_USERNAME and FASTMAIL_CALDAV_PASSWORD to use CalDAV.');
          }
          const events = await davClient.getCalendarEvents(calendarId, limit, timeMin, timeMax);
          return { content: [{ type: 'text', text: JSON.stringify(events, null, 2) }] };
        }
      }

      case 'get_calendar_event': {
        const { eventId } = args as any;
        if (!eventId) {
          throw new McpError(ErrorCode.InvalidParams, 'eventId is required');
        }
        try {
          const contactsClient = initializeContactsCalendarClient();
          const event = await contactsClient.getCalendarEventById(eventId);
          return { content: [{ type: 'text', text: JSON.stringify(event, null, 2) }] };
        } catch {
          const davClient = initializeCalDAVClient();
          if (!davClient) {
            throw new McpError(ErrorCode.InvalidRequest, 'JMAP calendars not available and CalDAV not configured. Set FASTMAIL_CALDAV_USERNAME and FASTMAIL_CALDAV_PASSWORD to use CalDAV.');
          }
          const event = await davClient.getCalendarEventById(eventId);
          return { content: [{ type: 'text', text: JSON.stringify(event, null, 2) }] };
        }
      }

      case 'find_duplicate_event': {
        const { title, start, end, calendarNames, timezone } = args as any;
        if (!title || !start) {
          throw new McpError(ErrorCode.InvalidParams, 'title and start are required');
        }
        try {
          const contactsClient = initializeContactsCalendarClient();
          const result = await contactsClient.findDuplicateEvent({ title, start, end, calendarNames, timezone });
          return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
        } catch {
          const davClient = initializeCalDAVClient();
          if (!davClient) {
            throw new McpError(ErrorCode.InvalidRequest, 'JMAP calendars not available and CalDAV not configured. Set FASTMAIL_CALDAV_USERNAME and FASTMAIL_CALDAV_PASSWORD to use CalDAV.');
          }
          const titleLower = String(title).trim().toLowerCase();
          if (!titleLower) throw new McpError(ErrorCode.InvalidParams, 'title is required for duplicate detection');

          const startDay = String(start).slice(0, 10);
          const endDay = end ? String(end).slice(0, 10) : startDay;
          const addDays = (d: string, delta: number) => {
            const dt = new Date(d + 'T00:00:00Z');
            dt.setUTCDate(dt.getUTCDate() + delta);
            return dt.toISOString().slice(0, 10);
          };
          const timeMin = addDays(startDay, -1) + 'T00:00:00Z';
          const timeMax = addDays(endDay, 1) + 'T23:59:59Z';

          const allCalendars = await davClient.getCalendars();
          let targetCalendars = allCalendars;
          if (calendarNames && calendarNames.length > 0) {
            const wanted = new Set(calendarNames.map((n: string) => n.toLowerCase()));
            targetCalendars = allCalendars.filter((c: any) =>
              wanted.has((c.displayName || '').toLowerCase()) ||
              wanted.has((c.url || '').toLowerCase()) ||
              wanted.has((c.id || '').toLowerCase())
            );
          }

          const searched: string[] = [];
          for (const cal of targetCalendars) {
            const calKey = cal.url || cal.id || cal.displayName;
            searched.push(calKey);
            const events = await davClient.getCalendarEvents(calKey, 200, timeMin, timeMax);
            for (const ev of events) {
              const evTitle = String((ev as any).title || '').trim().toLowerCase();
              if (!evTitle) continue;
              const evStart = String((ev as any).start || '').slice(0, 10);
              if (evStart < startDay || evStart > endDay) {
                // Allow events within the ±1 window too
                if (evStart < addDays(startDay, -1) || evStart > addDays(endDay, 1)) continue;
              }
              if (evTitle === titleLower || evTitle.includes(titleLower) || titleLower.includes(evTitle)) {
                return { content: [{ type: 'text', text: JSON.stringify({ found: true, event: ev, searched }, null, 2) }] };
              }
            }
          }
          return { content: [{ type: 'text', text: JSON.stringify({ found: false, event: null, searched }, null, 2) }] };
        }
      }

      case 'create_calendar_event': {
        const { calendarId, title, description, start, end, location, participants, timezone } = args as any;
        if (!calendarId || !title || !start || !end) {
          throw new McpError(ErrorCode.InvalidParams, 'calendarId, title, start, and end are required');
        }
        try {
          const contactsClient = initializeContactsCalendarClient();
          const eventId = await contactsClient.createCalendarEvent({
            calendarId, title, description, start, end, location, participants, timezone,
          });
          return { content: [{ type: 'text', text: `Calendar event created successfully. Event ID: ${eventId}` }] };
        } catch {
          const davClient = initializeCalDAVClient();
          if (!davClient) {
            throw new McpError(ErrorCode.InvalidRequest, 'JMAP calendars not available and CalDAV not configured. Set FASTMAIL_CALDAV_USERNAME and FASTMAIL_CALDAV_PASSWORD to use CalDAV.');
          }
          const eventId = await davClient.createCalendarEvent({
            calendarId, title, description, start, end, location,
          });
          return { content: [{ type: 'text', text: `Calendar event created via CalDAV. Event ID: ${eventId}` }] };
        }
      }

      case 'delete_calendar_event': {
        let { eventId, scope, instanceStart } = args as any;
        if (!eventId) {
          throw new McpError(ErrorCode.InvalidParams, 'eventId is required');
        }
        const needsInference = !scope || (scope === 'this' && !instanceStart);
        try {
          const contactsClient = initializeContactsCalendarClient();
          if (needsInference) {
            const ev = await contactsClient.getCalendarEventById(eventId);
            if (!ev) throw new McpError(ErrorCode.InvalidParams, `Event not found: ${eventId}`);
            const isRecurring = !!(ev.recurrenceRules || ev.recurrenceOverrides);
            if (!scope) scope = isRecurring ? 'this' : 'all';
            if (scope === 'this' && !instanceStart) instanceStart = ev.start;
          }
          await contactsClient.deleteCalendarEvent(eventId, scope, instanceStart);
          return { content: [{ type: 'text', text: `Calendar event deleted successfully (scope: ${scope})` }] };
        } catch (err) {
          if (err instanceof McpError) throw err;
          const davClient = initializeCalDAVClient();
          if (!davClient) {
            throw new McpError(ErrorCode.InvalidRequest, 'JMAP calendars not available and CalDAV not configured.');
          }
          if (needsInference) {
            const ev = await davClient.getCalendarEventById(eventId);
            if (!ev) throw new McpError(ErrorCode.InvalidParams, `Event not found: ${eventId}`);
            if (!scope) scope = ev.isRecurring ? 'this' : 'all';
            if (scope === 'this' && !instanceStart) instanceStart = ev.start;
          }
          await davClient.deleteCalendarEvent(eventId, scope, instanceStart);
          return { content: [{ type: 'text', text: `Calendar event deleted successfully (scope: ${scope})` }] };
        }
      }

      case 'move_calendar_event': {
        let { eventId, targetCalendarId, scope, instanceStart } = args as any;
        if (!eventId || !targetCalendarId) {
          throw new McpError(ErrorCode.InvalidParams, 'eventId and targetCalendarId are required');
        }
        const needsInference = !scope || (scope === 'this' && !instanceStart);
        try {
          const contactsClient = initializeContactsCalendarClient();
          if (needsInference) {
            const ev = await contactsClient.getCalendarEventById(eventId);
            if (!ev) throw new McpError(ErrorCode.InvalidParams, `Event not found: ${eventId}`);
            const isRecurring = !!(ev.recurrenceRules || ev.recurrenceOverrides);
            if (!scope) scope = isRecurring ? 'this' : 'all';
            if (scope === 'this' && !instanceStart) instanceStart = ev.start;
          }
          const resolvedTarget = await contactsClient.resolveCalendarId(targetCalendarId);
          await contactsClient.moveCalendarEvent(eventId, resolvedTarget, scope, instanceStart);
          return { content: [{ type: 'text', text: `Calendar event moved successfully (scope: ${scope})` }] };
        } catch (err) {
          if (err instanceof McpError) throw err;
          const davClient = initializeCalDAVClient();
          if (!davClient) {
            throw new McpError(ErrorCode.InvalidRequest, 'JMAP calendars not available and CalDAV not configured.');
          }
          if (needsInference) {
            const ev = await davClient.getCalendarEventById(eventId);
            if (!ev) throw new McpError(ErrorCode.InvalidParams, `Event not found: ${eventId}`);
            if (!scope) scope = ev.isRecurring ? 'this' : 'all';
            if (scope === 'this' && !instanceStart) instanceStart = ev.start;
          }
          await davClient.moveCalendarEvent(eventId, targetCalendarId, scope, instanceStart);
          return { content: [{ type: 'text', text: `Calendar event moved successfully (scope: ${scope})` }] };
        }
      }

      case 'get_recent_emails': {
        const { limit = 10, mailboxName = 'inbox' } = args as any;
        const client = initializeClient();
        const emails = await client.getRecentEmails(limit, mailboxName);
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(emails, null, 2),
            },
          ],
        };
      }

      case 'mark_email_read': {
        const { emailId, read = true } = args as any;
        if (!emailId) {
          throw new McpError(ErrorCode.InvalidParams, 'emailId is required');
        }
        const client = initializeClient();
        await client.markEmailRead(emailId, read);
        return {
          content: [
            {
              type: 'text',
              text: `Email ${read ? 'marked as read' : 'marked as unread'} successfully`,
            },
          ],
        };
      }

      case 'pin_email': {
        const { emailId, pinned = true } = args as any;
        if (!emailId) {
          throw new McpError(ErrorCode.InvalidParams, 'emailId is required');
        }
        const client = initializeClient();
        await client.pinEmail(emailId, pinned);
        return {
          content: [
            {
              type: 'text',
              text: `Email ${pinned ? 'pinned' : 'unpinned'} successfully`,
            },
          ],
        };
      }

      case 'delete_email': {
        const { emailId } = args as any;
        if (!emailId) {
          throw new McpError(ErrorCode.InvalidParams, 'emailId is required');
        }
        const client = initializeClient();
        await client.deleteEmail(emailId);
        return {
          content: [
            {
              type: 'text',
              text: 'Email deleted successfully (moved to trash)',
            },
          ],
        };
      }

      case 'move_email': {
        const { emailId, targetMailboxId } = args as any;
        if (!emailId || !targetMailboxId) {
          throw new McpError(ErrorCode.InvalidParams, 'emailId and targetMailboxId are required');
        }
        const client = initializeClient();
        await client.moveEmail(emailId, targetMailboxId);
        return {
          content: [
            {
              type: 'text',
              text: 'Email moved successfully',
            },
          ],
        };
      }

      case 'add_labels': {
        const { emailId, mailboxIds } = args as any;
        if (!emailId) {
          throw new McpError(ErrorCode.InvalidParams, 'emailId is required');
        }
        if (!mailboxIds || !Array.isArray(mailboxIds) || mailboxIds.length === 0) {
          throw new McpError(ErrorCode.InvalidParams, 'mailboxIds array is required and must not be empty');
        }
        const client = initializeClient();
        await client.addLabels(emailId, mailboxIds);
        return {
          content: [
            {
              type: 'text',
              text: `Labels added successfully to email`,
            },
          ],
        };
      }

      case 'remove_labels': {
        const { emailId, mailboxIds } = args as any;
        if (!emailId) {
          throw new McpError(ErrorCode.InvalidParams, 'emailId is required');
        }
        if (!mailboxIds || !Array.isArray(mailboxIds) || mailboxIds.length === 0) {
          throw new McpError(ErrorCode.InvalidParams, 'mailboxIds array is required and must not be empty');
        }
        const client = initializeClient();
        await client.removeLabels(emailId, mailboxIds);
        return {
          content: [
            {
              type: 'text',
              text: `Labels removed successfully from email`,
            },
          ],
        };
      }


      case 'advanced_search': {
        const { query, from, to, subject, hasAttachment, isUnread, isPinned, mailboxId, after, since, before, limit } = args as any;
        const client = initializeClient();
        const emails = await client.advancedSearch({
          query, from, to, subject, hasAttachment, isUnread, isPinned, mailboxId, after, since, before, limit
        });
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(emails, null, 2),
            },
          ],
        };
      }

      case 'get_thread': {
        const { threadId } = args as any;
        if (!threadId) {
          throw new McpError(ErrorCode.InvalidParams, 'threadId is required');
        }
        const client = initializeClient();
        try {
          const thread = await client.getThread(threadId);
          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify(thread, null, 2),
              },
            ],
          };
        } catch (error) {
          // Provide helpful error information
          throw new McpError(ErrorCode.InternalError, `Thread access failed: ${error instanceof Error ? error.message : String(error)}`);
        }
      }

      case 'get_mailbox_stats': {
        const { mailboxId } = args as any;
        const client = initializeClient();
        const stats = await client.getMailboxStats(mailboxId);
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(stats, null, 2),
            },
          ],
        };
      }

      case 'get_account_summary': {
        const client = initializeClient();
        const summary = await client.getAccountSummary();
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(summary, null, 2),
            },
          ],
        };
      }

      case 'bulk_mark_read': {
        const { emailIds, read = true } = args as any;
        if (!emailIds || !Array.isArray(emailIds) || emailIds.length === 0) {
          throw new McpError(ErrorCode.InvalidParams, 'emailIds array is required and must not be empty');
        }
        const client = initializeClient();
        await client.bulkMarkRead(emailIds, read);
        return {
          content: [
            {
              type: 'text',
              text: `${emailIds.length} emails ${read ? 'marked as read' : 'marked as unread'} successfully`,
            },
          ],
        };
      }

      case 'bulk_pin': {
        const { emailIds, pinned = true } = args as any;
        if (!emailIds || !Array.isArray(emailIds) || emailIds.length === 0) {
          throw new McpError(ErrorCode.InvalidParams, 'emailIds array is required and must not be empty');
        }
        const client = initializeClient();
        await client.bulkPinEmails(emailIds, pinned);
        return {
          content: [
            {
              type: 'text',
              text: `${emailIds.length} emails ${pinned ? 'pinned' : 'unpinned'} successfully`,
            },
          ],
        };
      }

      case 'bulk_move': {
        const { emailIds, targetMailboxId } = args as any;
        if (!emailIds || !Array.isArray(emailIds) || emailIds.length === 0) {
          throw new McpError(ErrorCode.InvalidParams, 'emailIds array is required and must not be empty');
        }
        if (!targetMailboxId) {
          throw new McpError(ErrorCode.InvalidParams, 'targetMailboxId is required');
        }
        const client = initializeClient();
        await client.bulkMove(emailIds, targetMailboxId);
        return {
          content: [
            {
              type: 'text',
              text: `${emailIds.length} emails moved successfully`,
            },
          ],
        };
      }

      case 'bulk_delete': {
        const { emailIds } = args as any;
        if (!emailIds || !Array.isArray(emailIds) || emailIds.length === 0) {
          throw new McpError(ErrorCode.InvalidParams, 'emailIds array is required and must not be empty');
        }
        const client = initializeClient();
        await client.bulkDelete(emailIds);
        return {
          content: [
            {
              type: 'text',
              text: `${emailIds.length} emails deleted successfully (moved to trash)`,
            },
          ],
        };
      }

      case 'bulk_add_labels': {
        const { emailIds, mailboxIds } = args as any;
        if (!emailIds || !Array.isArray(emailIds) || emailIds.length === 0) {
          throw new McpError(ErrorCode.InvalidParams, 'emailIds array is required and must not be empty');
        }
        if (!mailboxIds || !Array.isArray(mailboxIds) || mailboxIds.length === 0) {
          throw new McpError(ErrorCode.InvalidParams, 'mailboxIds array is required and must not be empty');
        }
        const client = initializeClient();
        await client.bulkAddLabels(emailIds, mailboxIds);
        return {
          content: [
            {
              type: 'text',
              text: `Labels added successfully to ${emailIds.length} emails`,
            },
          ],
        };
      }

      case 'bulk_remove_labels': {
        const { emailIds, mailboxIds } = args as any;
        if (!emailIds || !Array.isArray(emailIds) || emailIds.length === 0) {
          throw new McpError(ErrorCode.InvalidParams, 'emailIds array is required and must not be empty');
        }
        if (!mailboxIds || !Array.isArray(mailboxIds) || mailboxIds.length === 0) {
          throw new McpError(ErrorCode.InvalidParams, 'mailboxIds array is required and must not be empty');
        }
        const client = initializeClient();
        await client.bulkRemoveLabels(emailIds, mailboxIds);
        return {
          content: [
            {
              type: 'text',
              text: `Labels removed successfully from ${emailIds.length} emails`,
            },
          ],
        };
      }

      default:
        throw new McpError(ErrorCode.MethodNotFound, `Unknown tool: ${name}`);
    }
  } catch (error) {
    if (error instanceof McpError) {
      throw error;
    }
    throw new McpError(
      ErrorCode.InternalError,
      `Tool execution failed: ${error instanceof Error ? error.message : String(error)}`
    );
  }
});

async function runServer() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error('Fastmail MCP server running on stdio');
}

runServer().catch(() => {
  // Avoid logging raw error objects to prevent accidental PII leakage
  console.error('Fastmail MCP server failed to start');
  process.exit(1);
});