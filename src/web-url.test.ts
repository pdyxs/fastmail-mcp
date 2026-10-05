import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'fs';
import { buildMailboxPath, pickMailboxPath, buildWebUrl, resolveWebUserId } from './web-url.js';

const MAILBOXES = JSON.parse(readFileSync(new URL('./fixtures/mailboxes.json', import.meta.url), 'utf8'));

describe('buildMailboxPath', () => {
  it('joins a nested mailbox with its parents using "."', () => {
    // Feed's parent is Inbox
    assert.equal(buildMailboxPath('P2Ubw', MAILBOXES), 'Inbox.Feed');
  });

  it('returns a top-level mailbox name on its own', () => {
    assert.equal(buildMailboxPath('P6-', MAILBOXES), 'Archive');
  });

  it('handles three levels deep', () => {
    const boxes = [
      { id: 'a', name: 'A', parentId: null },
      { id: 'b', name: 'B', parentId: 'a' },
      { id: 'c', name: 'C', parentId: 'b' },
    ];
    assert.equal(buildMailboxPath('c', boxes), 'A.B.C');
  });

  it('returns null for an unknown mailbox id', () => {
    assert.equal(buildMailboxPath('nope', MAILBOXES), null);
  });

  it('does not loop forever on a parentId cycle', () => {
    const boxes = [
      { id: 'a', name: 'A', parentId: 'b' },
      { id: 'b', name: 'B', parentId: 'a' },
    ];
    assert.equal(buildMailboxPath('a', boxes), null);
  });
});

describe('pickMailboxPath', () => {
  it('uses the first mailboxId (in mailboxIds order) that exists in the tree', () => {
    assert.equal(pickMailboxPath({ missing: true, P2Ubw: true, 'P6-': true }, MAILBOXES), 'Inbox.Feed');
  });

  it('returns null when no mailboxIds resolve', () => {
    assert.equal(pickMailboxPath({ missing: true }, MAILBOXES), null);
    assert.equal(pickMailboxPath(undefined, MAILBOXES), null);
  });
});

describe('buildWebUrl', () => {
  it('builds mail/<path>/<threadId>.<emailId>?u=<u>', () => {
    assert.equal(
      buildWebUrl({ emailId: 'StmP5P6Dmx8Z', threadId: 'AvcOCME5z7sN', mailboxPath: 'Inbox.Feed', webUserId: 'abc123' }),
      'https://app.fastmail.com/mail/Inbox.Feed/AvcOCME5z7sN.StmP5P6Dmx8Z?u=abc123',
    );
  });

  it('omits ?u when there is no web user id', () => {
    assert.equal(
      buildWebUrl({ emailId: 'e1', threadId: 't1', mailboxPath: 'Archive', webUserId: null }),
      'https://app.fastmail.com/mail/Archive/t1.e1',
    );
  });

  it('URI-encodes path segments with spaces', () => {
    assert.equal(
      buildWebUrl({ emailId: 'e1', threadId: 't1', mailboxPath: 'Paper Trail', webUserId: null }),
      'https://app.fastmail.com/mail/Paper%20Trail/t1.e1',
    );
  });

  it('falls back to the legacy email URL when mailbox path or thread is unknown', () => {
    assert.equal(
      buildWebUrl({ emailId: 'e1', threadId: undefined, mailboxPath: null, webUserId: 'x' }),
      'https://app.fastmail.com/mail/email/e1',
    );
  });
});

describe('resolveWebUserId', () => {
  it('derives u from the JMAP accountId by dropping the leading "u"', () => {
    assert.equal(resolveWebUserId('u1234abcd', {}), '1234abcd');
  });

  it('prefers FASTMAIL_WEB_USER_ID when set', () => {
    assert.equal(resolveWebUserId('u1234abcd', { FASTMAIL_WEB_USER_ID: 'override' }), 'override');
  });

  it('returns null when the accountId is not u<hex> and no env is set', () => {
    assert.equal(resolveWebUserId('acct-123', {}), null);
  });
});
