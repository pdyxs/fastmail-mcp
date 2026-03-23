import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import {
  extractVEvent,
  parseICalValue,
  formatICalDate,
  parseCalendarObject,
  expandRecurringEvent,
} from './caldav-client.js';

describe('extractVEvent', () => {
  it('extracts VEVENT block from iCalendar data', () => {
    const ical = [
      'BEGIN:VCALENDAR',
      'BEGIN:VTIMEZONE',
      'TZID:Europe/Rome',
      'DTSTART:19700101T000000',
      'END:VTIMEZONE',
      'BEGIN:VEVENT',
      'SUMMARY:Test Event',
      'DTSTART;TZID=Europe/Rome:20260320T083000',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\n');

    const vevent = extractVEvent(ical);
    assert.ok(vevent.includes('SUMMARY:Test Event'));
    assert.ok(vevent.includes('DTSTART;TZID=Europe/Rome:20260320T083000'));
    assert.ok(!vevent.includes('VTIMEZONE'));
    assert.ok(!vevent.includes('TZID:Europe/Rome'));
  });

  it('returns original data when no VEVENT block found', () => {
    const data = 'no vevent here';
    assert.equal(extractVEvent(data), data);
  });

  it('ignores VTIMEZONE DTSTART when extracting VEVENT', () => {
    const ical = [
      'BEGIN:VCALENDAR',
      'BEGIN:VTIMEZONE',
      'TZID:Europe/Rome',
      'BEGIN:STANDARD',
      'DTSTART:19701025T030000',
      'RRULE:FREQ=YEARLY;BYMONTH=10;BYDAY=-1SU',
      'END:STANDARD',
      'END:VTIMEZONE',
      'BEGIN:VEVENT',
      'DTSTART;TZID=Europe/Rome:20260320T083000',
      'SUMMARY:Meeting',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\n');

    const vevent = extractVEvent(ical);
    // Should only have the VEVENT DTSTART, not the VTIMEZONE one
    const dtstartMatches = vevent.match(/DTSTART/g);
    assert.equal(dtstartMatches?.length, 1);
    assert.ok(vevent.includes('20260320T083000'));
  });
});

describe('parseICalValue', () => {
  it('handles simple KEY:value format', () => {
    const vevent = 'SUMMARY:Test Event\nDTSTART:20260320T083000Z';
    assert.equal(parseICalValue(vevent, 'SUMMARY'), 'Test Event');
    assert.equal(parseICalValue(vevent, 'DTSTART'), '20260320T083000Z');
  });

  it('handles parameterized KEY;TZID=...:value format', () => {
    const vevent = 'DTSTART;TZID=Europe/Rome:20260320T083000\nSUMMARY:Test';
    assert.equal(parseICalValue(vevent, 'DTSTART'), '20260320T083000');
  });

  it('handles VALUE=DATE format', () => {
    const vevent = 'DTSTART;VALUE=DATE:20260324\nDTEND;VALUE=DATE:20260325';
    assert.equal(parseICalValue(vevent, 'DTSTART'), '20260324');
    assert.equal(parseICalValue(vevent, 'DTEND'), '20260325');
  });

  it('returns undefined for missing keys', () => {
    const vevent = 'SUMMARY:Test';
    assert.equal(parseICalValue(vevent, 'LOCATION'), undefined);
  });

  it('handles line folding (continuation lines)', () => {
    const vevent = 'DESCRIPTION:This is a long\n description that wraps\nSUMMARY:Test';
    assert.equal(parseICalValue(vevent, 'DESCRIPTION'), 'This is a longdescription that wraps');
  });
});

describe('formatICalDate', () => {
  it('formats datetime without timezone', () => {
    assert.equal(formatICalDate('20260320T083000'), '2026-03-20T08:30:00');
  });

  it('formats datetime with Z suffix', () => {
    assert.equal(formatICalDate('20260320T083000Z'), '2026-03-20T08:30:00Z');
  });

  it('formats all-day date', () => {
    assert.equal(formatICalDate('20260324'), '2026-03-24');
  });

  it('returns undefined for undefined input', () => {
    assert.equal(formatICalDate(undefined), undefined);
  });

  it('returns cleaned string for unrecognized formats', () => {
    assert.equal(formatICalDate('something-else'), 'something-else');
  });

  it('strips carriage returns', () => {
    assert.equal(formatICalDate('20260320T083000\r'), '2026-03-20T08:30:00');
  });
});

describe('parseCalendarObject', () => {
  it('parses a full calendar object with VTIMEZONE + VEVENT', () => {
    const data = [
      'BEGIN:VCALENDAR',
      'VERSION:2.0',
      'BEGIN:VTIMEZONE',
      'TZID:Europe/Rome',
      'BEGIN:STANDARD',
      'DTSTART:19701025T030000',
      'RRULE:FREQ=YEARLY;BYMONTH=10;BYDAY=-1SU',
      'TZOFFSETFROM:+0200',
      'TZOFFSETTO:+0100',
      'END:STANDARD',
      'BEGIN:DAYLIGHT',
      'DTSTART:19700329T020000',
      'RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=-1SU',
      'TZOFFSETFROM:+0100',
      'TZOFFSETTO:+0200',
      'END:DAYLIGHT',
      'END:VTIMEZONE',
      'BEGIN:VEVENT',
      'UID:abc123@fastmail',
      'DTSTART;TZID=Europe/Rome:20260320T083000',
      'DTEND;TZID=Europe/Rome:20260320T093000',
      'SUMMARY:Morning Meeting',
      'DESCRIPTION:Discuss project\\nSecond line',
      'LOCATION:Room A\\, Building 1',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');

    const event = parseCalendarObject({ data, url: 'https://caldav.example.com/cal/abc.ics' });

    assert.equal(event.id, 'abc123@fastmail');
    assert.equal(event.url, 'https://caldav.example.com/cal/abc.ics');
    assert.equal(event.title, 'Morning Meeting');
    assert.equal(event.description, 'Discuss project\nSecond line');
    assert.equal(event.location, 'Room A, Building 1');
    // Should get the VEVENT DTSTART, not the VTIMEZONE one
    assert.equal(event.start, '2026-03-20T08:30:00');
    assert.equal(event.end, '2026-03-20T09:30:00');
  });

  it('parses an all-day event', () => {
    const data = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:allday1@fastmail',
      'DTSTART;VALUE=DATE:20260324',
      'DTEND;VALUE=DATE:20260325',
      'SUMMARY:All Day Event',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');

    const event = parseCalendarObject({ data, url: '' });
    assert.equal(event.start, '2026-03-24');
    assert.equal(event.end, '2026-03-25');
    assert.equal(event.title, 'All Day Event');
  });

  it('parses a UTC event', () => {
    const data = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:utc1@fastmail',
      'DTSTART:20260320T083000Z',
      'DTEND:20260320T093000Z',
      'SUMMARY:UTC Event',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');

    const event = parseCalendarObject({ data, url: '' });
    assert.equal(event.start, '2026-03-20T08:30:00Z');
    assert.equal(event.end, '2026-03-20T09:30:00Z');
  });

  it('defaults title to Untitled when SUMMARY is missing', () => {
    const data = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:notitle@fastmail',
      'DTSTART:20260320T083000Z',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');

    const event = parseCalendarObject({ data, url: '' });
    assert.equal(event.title, 'Untitled');
  });

  it('handles missing optional fields', () => {
    const data = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:minimal@fastmail',
      'DTSTART:20260320T083000Z',
      'SUMMARY:Minimal',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');

    const event = parseCalendarObject({ data, url: '' });
    assert.equal(event.description, undefined);
    assert.equal(event.location, undefined);
    assert.equal(event.end, undefined);
  });
});

describe('expandRecurringEvent', () => {
  function makeObj(lines: string[]): { data: string; url: string } {
    return { data: lines.join('\r\n'), url: 'https://caldav.example.com/cal/event.ics' };
  }

  it('returns base event unchanged when no RRULE', () => {
    const obj = makeObj([
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:nonrecurring@test',
      'DTSTART:20260325T100000Z',
      'DTEND:20260325T110000Z',
      'SUMMARY:One-off',
      'END:VEVENT',
      'END:VCALENDAR',
    ]);
    const results = expandRecurringEvent(obj, '2026-03-25T00:00:00Z', '2026-03-31T23:59:59Z');
    assert.equal(results.length, 1);
    assert.equal(results[0].start, '2026-03-25T10:00:00Z');
  });

  it('expands weekly RRULE and returns occurrences in range', () => {
    // Weekly on Tuesdays, started 2024-06-18 (a Tuesday)
    const obj = makeObj([
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:climb@test',
      'DTSTART:20240618T180000Z',
      'DTEND:20240618T200000Z',
      'SUMMARY:Climb',
      'RRULE:FREQ=WEEKLY;BYDAY=TU',
      'END:VEVENT',
      'END:VCALENDAR',
    ]);
    // Week of 2026-03-25 to 2026-03-31 contains Tuesday 2026-03-31
    const results = expandRecurringEvent(obj, '2026-03-25T00:00:00Z', '2026-03-31T23:59:59Z');
    assert.equal(results.length, 1);
    assert.equal(results[0].title, 'Climb');
    assert.equal(results[0].start, '2026-03-31T18:00:00Z');
    assert.equal(results[0].end, '2026-03-31T20:00:00Z');
  });

  it('returns empty array when RRULE has no occurrences in range', () => {
    // Weekly on Sundays, querying a range with no Sunday
    const obj = makeObj([
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:sunday@test',
      'DTSTART:20240616T100000Z',
      'DTEND:20240616T110000Z',
      'SUMMARY:Sunday Event',
      'RRULE:FREQ=WEEKLY;BYDAY=SU',
      'END:VEVENT',
      'END:VCALENDAR',
    ]);
    // 2026-03-26 is a Thursday — query Mon-Sat (no Sunday)
    const results = expandRecurringEvent(obj, '2026-03-23T00:00:00Z', '2026-03-28T23:59:59Z');
    assert.equal(results.length, 0);
  });

  it('expands monthly RRULE', () => {
    const obj = makeObj([
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:monthly@test',
      'DTSTART:20240131T170000Z',
      'DTEND:20240131T210000Z',
      'SUMMARY:Tabletop Tuesdays',
      'RRULE:FREQ=MONTHLY;BYDAY=-1TU',
      'END:VEVENT',
      'END:VCALENDAR',
    ]);
    // Last Tuesday of March 2026 is 2026-03-31
    const results = expandRecurringEvent(obj, '2026-03-25T00:00:00Z', '2026-03-31T23:59:59Z');
    assert.equal(results.length, 1);
    assert.equal(results[0].start, '2026-03-31T17:00:00Z');
  });

  it('respects EXDATE exclusions', () => {
    const obj = makeObj([
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:climb-exdate@test',
      'DTSTART:20240618T180000Z',
      'DTEND:20240618T200000Z',
      'SUMMARY:Climb',
      'RRULE:FREQ=WEEKLY;BYDAY=TU',
      'EXDATE:20260331T180000Z',
      'END:VEVENT',
      'END:VCALENDAR',
    ]);
    const results = expandRecurringEvent(obj, '2026-03-25T00:00:00Z', '2026-03-31T23:59:59Z');
    assert.equal(results.length, 0);
  });

  it('handles all-day recurring events', () => {
    const obj = makeObj([
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:allday-weekly@test',
      'DTSTART;VALUE=DATE:20240101',
      'DTEND;VALUE=DATE:20240102',
      'SUMMARY:Weekly All Day',
      'RRULE:FREQ=WEEKLY;BYDAY=WE',
      'END:VEVENT',
      'END:VCALENDAR',
    ]);
    // 2026-03-25 is a Wednesday
    const results = expandRecurringEvent(obj, '2026-03-25T00:00:00Z', '2026-03-31T23:59:59Z');
    assert.equal(results.length, 1);
    assert.equal(results[0].start, '2026-03-25');
    assert.equal(results[0].end, '2026-03-26');
  });
});
