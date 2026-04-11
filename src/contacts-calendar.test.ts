import { describe, it, beforeEach, mock } from 'node:test';
import assert from 'node:assert/strict';
import { ContactsCalendarClient } from './contacts-calendar.js';
import { FastmailAuth } from './auth.js';

const ACCOUNT_ID = 'acct-123';
const EVENT_ID = 'event-abc';
const TARGET_CALENDAR_ID = 'cal-456';

function makeClient(): ContactsCalendarClient {
  const auth = new FastmailAuth({ apiToken: 'fake-token' });
  const client = new ContactsCalendarClient(auth);

  mock.method(client, 'getSession', async () => ({
    apiUrl: 'https://api.example.com/jmap/api/',
    accountId: ACCOUNT_ID,
    capabilities: {
      'urn:ietf:params:jmap:calendars': {},
    },
  }));

  return client;
}

function stubMakeRequest(client: ContactsCalendarClient, response: any) {
  mock.method(client, 'makeRequest', async () => response);
}

function captureRequest(client: ContactsCalendarClient): { lastRequest: any } {
  const capture = { lastRequest: null as any };
  mock.method(client, 'makeRequest', async (req: any) => {
    capture.lastRequest = req;
    return {
      methodResponses: [
        ['CalendarEvent/set', { destroyed: [EVENT_ID], updated: { [EVENT_ID]: {} }, created: {} }, '0'],
      ],
    };
  });
  return capture;
}

// ---------- deleteCalendarEvent ----------

describe('deleteCalendarEvent – scope "all"', () => {
  let client: ContactsCalendarClient;

  beforeEach(() => { client = makeClient(); });

  it('sends CalendarEvent/set with destroy', async () => {
    const capture = captureRequest(client);
    await client.deleteCalendarEvent(EVENT_ID, 'all');

    const call = capture.lastRequest.methodCalls[0];
    assert.equal(call[0], 'CalendarEvent/set');
    assert.deepEqual(call[1].destroy, [EVENT_ID]);
    assert.equal(call[1].accountId, ACCOUNT_ID);
  });

  it('throws when server returns notDestroyed', async () => {
    stubMakeRequest(client, {
      methodResponses: [
        ['CalendarEvent/set', { notDestroyed: { [EVENT_ID]: { type: 'notFound' } } }, '0'],
      ],
    });
    await assert.rejects(
      () => client.deleteCalendarEvent(EVENT_ID, 'all'),
      /Failed to delete event/
    );
  });
});

describe('deleteCalendarEvent – scope "this"', () => {
  let client: ContactsCalendarClient;

  beforeEach(() => { client = makeClient(); });

  it('sends CalendarEvent/set with recurrenceOverrides excluded', async () => {
    const capture = captureRequest(client);
    await client.deleteCalendarEvent(EVENT_ID, 'this', '2026-04-10T14:00:00Z');

    const call = capture.lastRequest.methodCalls[0];
    assert.equal(call[0], 'CalendarEvent/set');
    const update = call[1].update[EVENT_ID];
    // recurrenceId strips Z and truncates to 19 chars
    assert.deepEqual(update['recurrenceOverrides/2026-04-10T14:00:00'], { excluded: true });
  });

  it('uses instanceStart without Z directly as recurrenceId', async () => {
    const capture = captureRequest(client);
    await client.deleteCalendarEvent(EVENT_ID, 'this', '2026-04-10T09:30:00');

    const call = capture.lastRequest.methodCalls[0];
    const update = call[1].update[EVENT_ID];
    assert.deepEqual(update['recurrenceOverrides/2026-04-10T09:30:00'], { excluded: true });
  });

  it('throws when instanceStart is missing', async () => {
    await assert.rejects(
      () => client.deleteCalendarEvent(EVENT_ID, 'this'),
      /instanceStart is required/
    );
  });

  it('throws when server returns notUpdated', async () => {
    stubMakeRequest(client, {
      methodResponses: [
        ['CalendarEvent/set', { notUpdated: { [EVENT_ID]: { type: 'notFound' } } }, '0'],
      ],
    });
    await assert.rejects(
      () => client.deleteCalendarEvent(EVENT_ID, 'this', '2026-04-10T14:00:00Z'),
      /Failed to exclude occurrence/
    );
  });
});

// ---------- moveCalendarEvent ----------

describe('moveCalendarEvent – scope "all"', () => {
  let client: ContactsCalendarClient;

  beforeEach(() => { client = makeClient(); });

  it('sends CalendarEvent/set updating calendarIds', async () => {
    const capture = captureRequest(client);
    await client.moveCalendarEvent(EVENT_ID, TARGET_CALENDAR_ID, 'all');

    const call = capture.lastRequest.methodCalls[0];
    assert.equal(call[0], 'CalendarEvent/set');
    const update = call[1].update[EVENT_ID];
    assert.deepEqual(update.calendarIds, { [TARGET_CALENDAR_ID]: true });
  });

  it('throws when server returns notUpdated', async () => {
    stubMakeRequest(client, {
      methodResponses: [
        ['CalendarEvent/set', { notUpdated: { [EVENT_ID]: { type: 'forbidden' } } }, '0'],
      ],
    });
    await assert.rejects(
      () => client.moveCalendarEvent(EVENT_ID, TARGET_CALENDAR_ID, 'all'),
      /Failed to move event/
    );
  });
});

describe('moveCalendarEvent – scope "this"', () => {
  let client: ContactsCalendarClient;

  beforeEach(() => { client = makeClient(); });

  it('creates detached occurrence in target and excludes from source', async () => {
    // getCalendarEventById is called first to get duration
    mock.method(client, 'getCalendarEventById', async () => ({
      id: EVENT_ID,
      title: 'Weekly Standup',
      description: 'Daily sync',
      location: 'Zoom',
      duration: 'PT30M',
    }));

    const capture = captureRequest(client);
    await client.moveCalendarEvent(EVENT_ID, TARGET_CALENDAR_ID, 'this', '2026-04-10T14:00:00Z');

    const call = capture.lastRequest.methodCalls[0];
    assert.equal(call[0], 'CalendarEvent/set');

    // New occurrence created in target calendar
    const created = call[1].create.newOccurrence;
    assert.deepEqual(created.calendarIds, { [TARGET_CALENDAR_ID]: true });
    assert.equal(created.title, 'Weekly Standup');
    assert.equal(created.start, '2026-04-10T14:00:00');
    // 14:00 + 30min = 14:30
    assert.equal(created.end, '2026-04-10T14:30:00');

    // Source occurrence excluded
    const update = call[1].update[EVENT_ID];
    assert.deepEqual(update['recurrenceOverrides/2026-04-10T14:00:00'], { excluded: true });
  });

  it('handles duration in days correctly', async () => {
    mock.method(client, 'getCalendarEventById', async () => ({
      id: EVENT_ID,
      title: 'All Day',
      duration: 'P1D',
    }));

    const capture = captureRequest(client);
    await client.moveCalendarEvent(EVENT_ID, TARGET_CALENDAR_ID, 'this', '2026-04-10T00:00:00Z');

    const call = capture.lastRequest.methodCalls[0];
    const created = call[1].create.newOccurrence;
    assert.equal(created.end, '2026-04-11T00:00:00');
  });

  it('throws when instanceStart is missing', async () => {
    mock.method(client, 'getCalendarEventById', async () => ({ id: EVENT_ID, title: 'X', duration: 'PT1H' }));
    await assert.rejects(
      () => client.moveCalendarEvent(EVENT_ID, TARGET_CALENDAR_ID, 'this'),
      /instanceStart is required/
    );
  });
});

// ---------- getCalendarEvents – calendarName annotation (M12) ----------

describe('getCalendarEvents – annotates calendarName', () => {
  let client: ContactsCalendarClient;

  beforeEach(() => { client = makeClient(); });

  it('populates calendarName using the primary calendarIds entry', async () => {
    // Seed calendar cache via getCalendars mock (used by annotateCalendarNames).
    mock.method(client, 'getCalendars', async () => [
      { id: 'cal-paul', name: 'Paul' },
      { id: 'cal-julie', name: 'Julie' },
    ]);
    stubMakeRequest(client, {
      methodResponses: [
        ['CalendarEvent/query', { ids: ['e1', 'e2'] }, 'query'],
        ['CalendarEvent/get', {
          list: [
            { id: 'e1', title: 'Paul Event', calendarIds: { 'cal-paul': true } },
            { id: 'e2', title: 'Julie Event', calendarIds: { 'cal-julie': true } },
          ],
        }, 'events'],
      ],
    });

    const events = await client.getCalendarEvents();
    assert.equal(events.length, 2);
    assert.equal(events[0].calendarName, 'Paul');
    assert.equal(events[1].calendarName, 'Julie');
  });

  it('leaves calendarName undefined when calendarIds empty / unknown', async () => {
    mock.method(client, 'getCalendars', async () => [{ id: 'cal-paul', name: 'Paul' }]);
    stubMakeRequest(client, {
      methodResponses: [
        ['CalendarEvent/query', { ids: ['e1'] }, 'query'],
        ['CalendarEvent/get', {
          list: [{ id: 'e1', title: 'Orphan', calendarIds: { 'cal-missing': true } }],
        }, 'events'],
      ],
    });

    const events = await client.getCalendarEvents();
    assert.equal(events[0].calendarName, undefined);
  });
});
