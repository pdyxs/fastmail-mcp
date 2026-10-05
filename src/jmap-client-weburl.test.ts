import { describe, it, beforeEach, afterEach, mock } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'fs';
import { JmapClient } from './jmap-client.js';
import { FastmailAuth } from './auth.js';

const MAILBOXES = JSON.parse(readFileSync(new URL('./fixtures/mailboxes.json', import.meta.url), 'utf8'));

function makeClient(accountId = 'u1234abcd') {
  const client = new JmapClient(new FastmailAuth({ apiToken: 'fake-token' }));
  mock.method(client, 'getSession', async () => ({
    apiUrl: 'https://api.example.com/jmap/api/',
    accountId,
    capabilities: {},
  }));
  const getMailboxes = mock.method(client, 'getMailboxes', async () => MAILBOXES);
  return { client, getMailboxes };
}

function emailResponse(list: any[]) {
  return { methodResponses: [['Email/get', { list }, 'email']] };
}

const FEED_EMAIL = { id: 'StmP5P6Dmx8Z', threadId: 'AvcOCME5z7sN', mailboxIds: { P2Ubw: true }, subject: 'Vol 247' };
const ARCHIVE_EMAIL = { id: 'e2', threadId: 't2', mailboxIds: { 'P6-': true }, subject: 'Old' };

describe('webUrl on fetched emails', () => {
  let savedEnv: string | undefined;
  beforeEach(() => {
    savedEnv = process.env.FASTMAIL_WEB_USER_ID;
    delete process.env.FASTMAIL_WEB_USER_ID;
  });
  afterEach(() => {
    if (savedEnv === undefined) delete process.env.FASTMAIL_WEB_USER_ID;
    else process.env.FASTMAIL_WEB_USER_ID = savedEnv;
  });

  it('getEmailById requests mailboxIds and threadId', async () => {
    const { client } = makeClient();
    const makeReq = mock.method(client, 'makeRequest', async () => emailResponse([FEED_EMAIL]));
    await client.getEmailById('StmP5P6Dmx8Z');
    const props = makeReq.mock.calls[0].arguments[0].methodCalls[0][1].properties;
    assert.ok(props.includes('mailboxIds'));
    assert.ok(props.includes('threadId'));
  });

  it('getEmailById links through the current mailbox path with derived u', async () => {
    const { client } = makeClient();
    mock.method(client, 'makeRequest', async () => emailResponse([FEED_EMAIL]));
    const email = await client.getEmailById('StmP5P6Dmx8Z');
    assert.equal(email.webUrl, 'https://app.fastmail.com/mail/Inbox.Feed/AvcOCME5z7sN.StmP5P6Dmx8Z?u=1234abcd');
  });

  it('uses FASTMAIL_WEB_USER_ID when the accountId cannot be derived', async () => {
    process.env.FASTMAIL_WEB_USER_ID = 'cfg42';
    const { client } = makeClient('acct-123');
    mock.method(client, 'makeRequest', async () => emailResponse([ARCHIVE_EMAIL]));
    const email = await client.getEmailById('e2');
    assert.equal(email.webUrl, 'https://app.fastmail.com/mail/Archive/t2.e2?u=cfg42');
  });

  it('getEmailsByIds builds a per-email mailbox path and requests mailboxIds', async () => {
    const { client } = makeClient();
    const makeReq = mock.method(client, 'makeRequest', async () => emailResponse([FEED_EMAIL, ARCHIVE_EMAIL]));
    const emails = await client.getEmailsByIds(['StmP5P6Dmx8Z', 'e2']);
    assert.ok(makeReq.mock.calls[0].arguments[0].methodCalls[0][1].properties.includes('mailboxIds'));
    assert.equal(emails[0].webUrl, 'https://app.fastmail.com/mail/Inbox.Feed/AvcOCME5z7sN.StmP5P6Dmx8Z?u=1234abcd');
    assert.equal(emails[1].webUrl, 'https://app.fastmail.com/mail/Archive/t2.e2?u=1234abcd');
  });

  it('caches the mailbox tree per client instance', async () => {
    const { client, getMailboxes } = makeClient();
    mock.method(client, 'makeRequest', async () => emailResponse([FEED_EMAIL]));
    await client.getEmailById('StmP5P6Dmx8Z');
    await client.getEmailById('StmP5P6Dmx8Z');
    await client.getEmailsByIds(['StmP5P6Dmx8Z']);
    assert.equal(getMailboxes.mock.calls.length, 1);
  });
});
