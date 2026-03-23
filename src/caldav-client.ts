import { DAVClient, DAVCalendar, DAVCalendarObject } from 'tsdav';
import { createRequire } from 'module';
const _require = createRequire(import.meta.url);
// rrule ships without "type":"module" so its .js files are CJS — load via createRequire
const { RRule, RRuleSet } = _require('rrule') as {
  RRule: typeof import('rrule').RRule;
  RRuleSet: typeof import('rrule').RRuleSet;
};

export interface CalDAVConfig {
  username: string;
  password: string;
  serverUrl?: string;
}

export interface CalendarInfo {
  id: string;
  displayName: string;
  url: string;
  description?: string;
  color?: string;
}

export interface CalendarEvent {
  id: string;
  url: string;
  title: string;
  description?: string;
  start?: string;
  end?: string;
  location?: string;
}

/**
 * Extract the VEVENT block from iCalendar data.
 * This avoids matching properties from VTIMEZONE or other components.
 */
export function extractVEvent(data: string): string {
  const match = data.match(/BEGIN:VEVENT[\s\S]*?END:VEVENT/);
  return match ? match[0] : data;
}

/**
 * Parse an iCalendar property value from within a VEVENT block.
 * Handles simple (KEY:value), parameterized (KEY;TZID=...:value),
 * and VALUE=DATE (KEY;VALUE=DATE:20260319) forms.
 * Also handles line folding (continuation lines starting with space/tab).
 */
export function parseICalValue(vevent: string, key: string): string | undefined {
  // Match KEY followed by either ; (params) or : (value), capturing the rest
  const regex = new RegExp(`^(${key}[;:].*)$`, 'm');
  const match = vevent.match(regex);
  if (!match) return undefined;

  // Handle line folding: continuation lines start with space or tab
  let fullLine = match[1];
  const lines = vevent.split(/\r?\n/);
  const matchIdx = lines.findIndex(l => l === fullLine || l.startsWith(fullLine));
  if (matchIdx >= 0) {
    for (let i = matchIdx + 1; i < lines.length; i++) {
      if (lines[i].startsWith(' ') || lines[i].startsWith('\t')) {
        fullLine += lines[i].substring(1);
      } else {
        break;
      }
    }
  }

  // Extract the value after the last colon in the property line
  // For DTSTART;TZID=Europe/Rome:20260320T083000 → 20260320T083000
  // For DTSTART:20220210T154500Z → 20220210T154500Z
  // For DTSTART;VALUE=DATE:20260324 → 20260324
  const colonIdx = fullLine.indexOf(':');
  if (colonIdx === -1) return undefined;
  return fullLine.substring(colonIdx + 1).trim();
}

/**
 * Format an iCalendar date/datetime string to ISO 8601.
 * Input formats: 20260320T083000, 20260320T083000Z, 20260324
 * Output: 2026-03-20T08:30:00, 2026-03-20T08:30:00Z, 2026-03-24
 */
export function formatICalDate(raw: string | undefined): string | undefined {
  if (!raw) return undefined;
  const cleaned = raw.replace(/\r/g, '');

  // All-day date: 20260324 (8 digits)
  if (/^\d{8}$/.test(cleaned)) {
    return `${cleaned.slice(0, 4)}-${cleaned.slice(4, 6)}-${cleaned.slice(6, 8)}`;
  }

  // DateTime: 20260320T083000 or 20260320T083000Z
  const dtMatch = cleaned.match(/^(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})(\d{2})(Z?)$/);
  if (dtMatch) {
    const [, y, m, d, hh, mm, ss, z] = dtMatch;
    return `${y}-${m}-${d}T${hh}:${mm}:${ss}${z}`;
  }

  return cleaned;
}

export function parseCalendarObject(obj: DAVCalendarObject): CalendarEvent {
  const vevent = extractVEvent(obj.data || '');
  const title = parseICalValue(vevent, 'SUMMARY') || 'Untitled';
  const description = parseICalValue(vevent, 'DESCRIPTION');
  const rawStart = parseICalValue(vevent, 'DTSTART');
  const rawEnd = parseICalValue(vevent, 'DTEND');
  const location = parseICalValue(vevent, 'LOCATION');
  const uid = parseICalValue(vevent, 'UID') || obj.url || '';

  return {
    id: uid,
    url: obj.url || '',
    title,
    description: description?.replace(/\\n/g, '\n').replace(/\\,/g, ',') || undefined,
    start: formatICalDate(rawStart),
    end: formatICalDate(rawEnd),
    location: location?.replace(/\\,/g, ',') || undefined,
  };
}

/**
 * Parse a compact iCal date/datetime value (from parseICalValue) into a Date.
 * Treats the time as UTC for rrule computation purposes.
 */
function parseICalDateToDate(raw: string): Date | undefined {
  const cleaned = raw.replace(/\r/g, '');
  // All-day: YYYYMMDD
  if (/^\d{8}$/.test(cleaned)) {
    return new Date(`${cleaned.slice(0, 4)}-${cleaned.slice(4, 6)}-${cleaned.slice(6, 8)}T00:00:00Z`);
  }
  // DateTime: YYYYMMDDTHHMMSS or YYYYMMDDTHHMMSSZ
  const m = cleaned.match(/^(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})(\d{2})(Z?)$/);
  if (m) {
    return new Date(`${m[1]}-${m[2]}-${m[3]}T${m[4]}:${m[5]}:${m[6]}Z`);
  }
  return undefined;
}

/**
 * Format a Date back to ISO 8601 in the same style as the original DTSTART.
 */
function formatOccurrenceDate(date: Date, isAllDay: boolean, hasZ: boolean): string {
  if (isAllDay) {
    return date.toISOString().slice(0, 10);
  }
  const iso = date.toISOString(); // e.g. 2026-03-31T18:00:00.000Z
  const base = `${iso.slice(0, 10)}T${iso.slice(11, 19)}`;
  return hasZ ? `${base}Z` : base;
}

/**
 * Expand a calendar object's recurring event into individual occurrences
 * within [timeMin, timeMax]. Returns the base event if no RRULE is present.
 * Returns an empty array if RRULE is present but no occurrences fall in range.
 */
export function expandRecurringEvent(
  obj: DAVCalendarObject,
  timeMin: string,
  timeMax: string,
): CalendarEvent[] {
  const vevent = extractVEvent(obj.data || '');
  const rruleStr = parseICalValue(vevent, 'RRULE');
  const baseEvent = parseCalendarObject(obj);

  if (!rruleStr) {
    return [baseEvent];
  }

  const rawStart = parseICalValue(vevent, 'DTSTART');
  if (!rawStart) return [baseEvent];

  const cleanedStart = rawStart.replace(/\r/g, '');
  const isAllDay = /^\d{8}$/.test(cleanedStart);
  const hasZ = cleanedStart.endsWith('Z');

  const dtstart = parseICalDateToDate(cleanedStart);
  if (!dtstart) return [baseEvent];

  const rawEnd = parseICalValue(vevent, 'DTEND');
  const dtend = rawEnd ? parseICalDateToDate(rawEnd.replace(/\r/g, '')) : undefined;
  const durationMs = dtend ? dtend.getTime() - dtstart.getTime() : 0;

  try {
    const rrule = new RRule({ ...RRule.parseString(rruleStr), dtstart });

    // Collect EXDATEs if present
    const exdateMatches = (obj.data || '').match(/^EXDATE[;:].+$/gm) || [];
    let rule: InstanceType<typeof RRule> | InstanceType<typeof RRuleSet> = rrule;
    if (exdateMatches.length > 0) {
      const set = new RRuleSet();
      set.rrule(rrule);
      for (const line of exdateMatches) {
        const val = line.substring(line.indexOf(':') + 1).trim();
        const exdate = parseICalDateToDate(val);
        if (exdate) set.exdate(exdate);
      }
      rule = set;
    }

    const after = new Date(timeMin);
    const before = new Date(timeMax);
    const occurrences = rule.between(after, before, true);

    return occurrences.map((occ: Date) => {
      const occEnd = durationMs > 0 ? new Date(occ.getTime() + durationMs) : undefined;
      return {
        ...baseEvent,
        start: formatOccurrenceDate(occ, isAllDay, hasZ),
        end: occEnd ? formatOccurrenceDate(occEnd, isAllDay, hasZ) : baseEvent.end,
      };
    });
  } catch {
    return [baseEvent];
  }
}

/**
 * Convert an ISO 8601 datetime string to iCalendar datetime format.
 * Handles timezone offsets by converting to UTC (YYYYMMDDTHHMMSSZ).
 * Handles all-day dates (YYYY-MM-DD → YYYYMMDD).
 */
function toICalDateTime(dateTimeStr: string): string {
  // All-day date: YYYY-MM-DD
  if (/^\d{4}-\d{2}-\d{2}$/.test(dateTimeStr)) {
    return dateTimeStr.replace(/-/g, '');
  }
  // Parse and convert to UTC
  const d = new Date(dateTimeStr);
  if (!isNaN(d.getTime())) {
    return d.toISOString().replace(/[-:]/g, '').replace(/\.\d{3}/, '');
  }
  // Fallback for already-compact formats
  return dateTimeStr.replace(/[-:]/g, '');
}

export class CalDAVCalendarClient {
  private config: CalDAVConfig;
  private client: DAVClient | null = null;
  private calendars: DAVCalendar[] | null = null;

  constructor(config: CalDAVConfig) {
    this.config = config;
  }

  private async getClient(): Promise<DAVClient> {
    if (this.client) return this.client;

    this.client = new DAVClient({
      serverUrl: this.config.serverUrl || 'https://caldav.fastmail.com',
      credentials: {
        username: this.config.username,
        password: this.config.password,
      },
      authMethod: 'Basic',
      defaultAccountType: 'caldav',
    });

    await this.client.login();
    return this.client;
  }

  async getCalendars(): Promise<CalendarInfo[]> {
    const client = await this.getClient();
    const calendars = await client.fetchCalendars();
    this.calendars = calendars;

    return calendars
      .filter(c => c.displayName !== 'DEFAULT_TASK_CALENDAR_NAME')
      .map(c => ({
        id: c.url || '',
        displayName: String(c.displayName || 'Unnamed'),
        url: c.url || '',
        description: c.description || undefined,
        color: (c as any).calendarColor || undefined,
      }));
  }

  async getCalendarEvents(calendarId?: string, limit: number = 50, timeMin?: string, timeMax?: string): Promise<CalendarEvent[]> {
    const client = await this.getClient();

    if (!this.calendars) {
      this.calendars = await client.fetchCalendars();
    }

    let targetCalendars = this.calendars.filter(
      c => c.displayName !== 'DEFAULT_TASK_CALENDAR_NAME'
    );
    if (calendarId) {
      targetCalendars = targetCalendars.filter(
        c => c.url === calendarId || c.displayName === calendarId
      );
    }

    const fetchOptions: Parameters<typeof client.fetchCalendarObjects>[0] = { calendar: targetCalendars[0] };
    if (timeMin || timeMax) {
      fetchOptions.timeRange = {
        start: timeMin || '1970-01-01T00:00:00Z',
        end: timeMax || '2099-12-31T23:59:59Z',
      };
    }

    const allEvents: CalendarEvent[] = [];
    for (const cal of targetCalendars) {
      const objects = await client.fetchCalendarObjects({ ...fetchOptions, calendar: cal });
      for (const obj of objects) {
        if (timeMin || timeMax) {
          const expanded = expandRecurringEvent(
            obj,
            timeMin || '1970-01-01T00:00:00Z',
            timeMax || '2099-12-31T23:59:59Z',
          );
          allEvents.push(...expanded);
        } else {
          allEvents.push(parseCalendarObject(obj));
        }
      }
      if (allEvents.length >= limit) break;
    }

    return allEvents.slice(0, limit);
  }

  async getCalendarEventById(eventId: string): Promise<CalendarEvent | null> {
    const client = await this.getClient();

    if (!this.calendars) {
      this.calendars = await client.fetchCalendars();
    }

    for (const cal of this.calendars) {
      const objects = await client.fetchCalendarObjects({ calendar: cal });
      for (const obj of objects) {
        const vevent = extractVEvent(obj.data || '');
        const uid = parseICalValue(vevent, 'UID');
        if (uid === eventId || obj.url === eventId) {
          return parseCalendarObject(obj);
        }
      }
    }

    return null;
  }

  async createCalendarEvent(event: {
    calendarId: string;
    title: string;
    description?: string;
    start: string;
    end: string;
    location?: string;
  }): Promise<string> {
    const client = await this.getClient();

    if (!this.calendars) {
      this.calendars = await client.fetchCalendars();
    }

    const targetCal = this.calendars.find(
      c => c.url === event.calendarId || c.displayName === event.calendarId
    );
    if (!targetCal) {
      throw new Error(`Calendar not found: ${event.calendarId}`);
    }

    const uid = `${Date.now()}-${Math.random().toString(36).slice(2)}@fastmail-mcp`;
    const now = new Date().toISOString().replace(/[-:]/g, '').replace(/\.\d{3}/, '');
    const ical = [
      'BEGIN:VCALENDAR',
      'VERSION:2.0',
      'PRODID:-//fastmail-mcp//CalDAV//EN',
      'BEGIN:VEVENT',
      `UID:${uid}`,
      `DTSTAMP:${now}`,
      `DTSTART:${toICalDateTime(event.start)}`,
      `DTEND:${toICalDateTime(event.end)}`,
      `SUMMARY:${event.title}`,
      event.description ? `DESCRIPTION:${event.description}` : '',
      event.location ? `LOCATION:${event.location}` : '',
      'END:VEVENT',
      'END:VCALENDAR',
    ].filter(Boolean).join('\r\n');

    await client.createCalendarObject({
      calendar: targetCal,
      filename: `${uid}.ics`,
      iCalString: ical,
    });

    return uid;
  }
}
