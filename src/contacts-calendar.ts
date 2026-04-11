import { JmapClient, JmapRequest } from './jmap-client.js';
import { defaultTz, normalizeEventDateTime } from './dates.js';

export class ContactsCalendarClient extends JmapClient {
  private calendarCache: Array<{ id: string; name: string }> | null = null;

  /**
   * Resolve a calendar name or ID to a JMAP calendar ID.
   *
   * Accepts: a JMAP calendar ID (passed through), a calendar name
   * (case-insensitive lookup), or undefined (returns undefined — caller
   * decides whether that's an error).
   *
   * CalDAV URLs are NOT resolved here — the caller should handle that by
   * falling through to the CalDAV client.
   */
  async resolveCalendarId(nameOrId: string): Promise<string> {
    if (!nameOrId) throw new Error('calendarId or calendarName is required');

    // CalDAV URLs shouldn't be resolved by this path — the caller handles them.
    if (nameOrId.startsWith('http://') || nameOrId.startsWith('https://')) {
      return nameOrId;
    }

    if (!this.calendarCache) {
      const cals = await this.getCalendars();
      this.calendarCache = cals.map((c: any) => ({ id: c.id, name: c.name }));
    }

    // Exact ID match
    const byId = this.calendarCache.find(c => c.id === nameOrId);
    if (byId) return byId.id;

    // Case-insensitive name match
    const lower = nameOrId.toLowerCase();
    const byName = this.calendarCache.find(c => c.name.toLowerCase() === lower);
    if (byName) return byName.id;

    // Partial name match as a last resort
    const byPartial = this.calendarCache.filter(c => c.name.toLowerCase().includes(lower));
    if (byPartial.length === 1) return byPartial[0].id;
    if (byPartial.length > 1) {
      throw new Error(
        `Ambiguous calendar name "${nameOrId}" — matches: ${byPartial.map(c => c.name).join(', ')}`
      );
    }

    const available = this.calendarCache.map(c => c.name).join(', ');
    throw new Error(`Calendar "${nameOrId}" not found. Available: ${available}`);
  }

  private async checkContactsPermission(): Promise<boolean> {
    const session = await this.getSession();
    return !!session.capabilities['urn:ietf:params:jmap:contacts'];
  }
  
  private async checkCalendarsPermission(): Promise<boolean> {
    const session = await this.getSession();
    return !!session.capabilities['urn:ietf:params:jmap:calendars'];
  }
  
  async getContacts(limit: number = 50): Promise<any[]> {
    // Check permissions first
    const hasPermission = await this.checkContactsPermission();
    if (!hasPermission) {
      throw new Error('Contacts access not available. This account may not have JMAP contacts permissions enabled. Please check your Fastmail account settings or contact support to enable contacts API access.');
    }

    const session = await this.getSession();
    
    // Try CardDAV namespace first, then Fastmail specific
    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:contacts'],
      methodCalls: [
        ['Contact/query', {
          accountId: session.accountId,
          limit
        }, 'query'],
        ['Contact/get', {
          accountId: session.accountId,
          '#ids': { resultOf: 'query', name: 'Contact/query', path: '/ids' },
          properties: ['id', 'name', 'emails', 'phones', 'addresses', 'notes']
        }, 'contacts']
      ]
    };

    try {
      const response = await this.makeRequest(request);
      return this.getListResult(response, 1);
    } catch (error) {
      // Fallback: try to get contacts using AddressBook methods
      const fallbackRequest: JmapRequest = {
        using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:contacts'],
        methodCalls: [
          ['AddressBook/get', {
            accountId: session.accountId
          }, 'addressbooks']
        ]
      };
      
      try {
        const fallbackResponse = await this.makeRequest(fallbackRequest);
        return this.getListResult(fallbackResponse, 0);
      } catch (fallbackError) {
        throw new Error(`Contacts not supported or accessible: ${error instanceof Error ? error.message : String(error)}. Try checking account permissions or enabling contacts API access in Fastmail settings.`);
      }
    }
  }

  async getContactById(id: string): Promise<any> {
    // Check permissions first
    const hasPermission = await this.checkContactsPermission();
    if (!hasPermission) {
      throw new Error('Contacts access not available. This account may not have JMAP contacts permissions enabled. Please check your Fastmail account settings or contact support to enable contacts API access.');
    }

    const session = await this.getSession();
    
    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:contacts'],
      methodCalls: [
        ['Contact/get', {
          accountId: session.accountId,
          ids: [id]
        }, 'contact']
      ]
    };

    try {
      const response = await this.makeRequest(request);
      return this.getListResult(response, 0)[0];
    } catch (error) {
      throw new Error(`Contact access not supported: ${error instanceof Error ? error.message : String(error)}. Try checking account permissions or enabling contacts API access in Fastmail settings.`);
    }
  }

  async searchContacts(query: string, limit: number = 20): Promise<any[]> {
    // Check permissions first
    const hasPermission = await this.checkContactsPermission();
    if (!hasPermission) {
      throw new Error('Contacts access not available. This account may not have JMAP contacts permissions enabled. Please check your Fastmail account settings or contact support to enable contacts API access.');
    }

    const session = await this.getSession();
    
    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:contacts'],
      methodCalls: [
        ['Contact/query', {
          accountId: session.accountId,
          filter: { text: query },
          limit
        }, 'query'],
        ['Contact/get', {
          accountId: session.accountId,
          '#ids': { resultOf: 'query', name: 'Contact/query', path: '/ids' },
          properties: ['id', 'name', 'emails', 'phones', 'addresses', 'notes']
        }, 'contacts']
      ]
    };

    try {
      const response = await this.makeRequest(request);
      return this.getListResult(response, 1);
    } catch (error) {
      throw new Error(`Contact search not supported: ${error instanceof Error ? error.message : String(error)}. Try checking account permissions or enabling contacts API access in Fastmail settings.`);
    }
  }

  async getCalendars(): Promise<any[]> {
    // Check permissions first
    const hasPermission = await this.checkCalendarsPermission();
    if (!hasPermission) {
      throw new Error('Calendar access not available. This account may not have JMAP calendar permissions enabled. Please check your Fastmail account settings or contact support to enable calendar API access.');
    }

    const session = await this.getSession();
    
    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:calendars'],
      methodCalls: [
        ['Calendar/get', {
          accountId: session.accountId
        }, 'calendars']
      ]
    };

    try {
      const response = await this.makeRequest(request);
      return this.getListResult(response, 0);
    } catch (error) {
      // Calendar access might require special permissions
      throw new Error(`Calendar access not supported or requires additional permissions. This may be due to account settings or JMAP scope limitations: ${error instanceof Error ? error.message : String(error)}. Try checking account permissions or enabling calendar API access in Fastmail settings.`);
    }
  }

  async getCalendarEvents(calendarId?: string, limit: number = 50): Promise<any[]> {
    // Check permissions first
    const hasPermission = await this.checkCalendarsPermission();
    if (!hasPermission) {
      throw new Error('Calendar access not available. This account may not have JMAP calendar permissions enabled. Please check your Fastmail account settings or contact support to enable calendar API access.');
    }

    const session = await this.getSession();

    const resolvedId = calendarId ? await this.resolveCalendarId(calendarId) : undefined;
    const filter = resolvedId ? { inCalendar: resolvedId } : {};

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:calendars'],
      methodCalls: [
        ['CalendarEvent/query', {
          accountId: session.accountId,
          filter,
          sort: [{ property: 'start', isAscending: true }],
          limit
        }, 'query'],
        ['CalendarEvent/get', {
          accountId: session.accountId,
          '#ids': { resultOf: 'query', name: 'CalendarEvent/query', path: '/ids' },
          properties: ['id', 'title', 'description', 'start', 'end', 'location', 'participants', 'calendarIds', 'recurrenceRules', 'recurrenceOverrides']
        }, 'events']
      ]
    };

    try {
      const response = await this.makeRequest(request);
      const events = this.getListResult(response, 1);
      return this.annotateCalendarNames(events);
    } catch (error) {
      throw new Error(`Calendar events access not supported: ${error instanceof Error ? error.message : String(error)}. Try checking account permissions or enabling calendar API access in Fastmail settings.`);
    }
  }

  /**
   * Populate each event with `calendarName` using its `calendarIds` map.
   * Callers shouldn't have to maintain their own ID-to-name table. If multiple
   * calendars are set, picks the first one (matches how Fastmail displays it).
   */
  private async annotateCalendarNames(events: any[]): Promise<any[]> {
    if (!events || events.length === 0) return events;
    if (!this.calendarCache) {
      const cals = await this.getCalendars();
      this.calendarCache = cals.map((c: any) => ({ id: c.id, name: c.name }));
    }
    const nameById = new Map(this.calendarCache.map(c => [c.id, c.name]));
    return events.map((ev: any) => {
      if (ev && ev.calendarIds && typeof ev.calendarIds === 'object') {
        const ids = Object.keys(ev.calendarIds).filter(k => ev.calendarIds[k]);
        const primary = ids[0];
        if (primary && nameById.has(primary)) {
          return { ...ev, calendarName: nameById.get(primary) };
        }
      }
      return ev;
    });
  }

  async getCalendarEventById(id: string): Promise<any> {
    // Check permissions first
    const hasPermission = await this.checkCalendarsPermission();
    if (!hasPermission) {
      throw new Error('Calendar access not available. This account may not have JMAP calendar permissions enabled. Please check your Fastmail account settings or contact support to enable calendar API access.');
    }

    const session = await this.getSession();
    
    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:calendars'],
      methodCalls: [
        ['CalendarEvent/get', {
          accountId: session.accountId,
          ids: [id],
          properties: ['id', 'title', 'description', 'start', 'end', 'duration', 'location', 'participants', 'calendarIds', 'recurrenceRules', 'recurrenceOverrides']
        }, 'event']
      ]
    };

    try {
      const response = await this.makeRequest(request);
      const events = this.getListResult(response, 0);
      const annotated = await this.annotateCalendarNames(events);
      return annotated[0];
    } catch (error) {
      throw new Error(`Calendar event access not supported: ${error instanceof Error ? error.message : String(error)}. Try checking account permissions or enabling calendar API access in Fastmail settings.`);
    }
  }

  async deleteCalendarEvent(eventId: string, scope: 'this' | 'all', instanceStart?: string): Promise<void> {
    const hasPermission = await this.checkCalendarsPermission();
    if (!hasPermission) {
      throw new Error('Calendar access not available. This account may not have JMAP calendar permissions enabled.');
    }
    const session = await this.getSession();

    if (scope === 'all') {
      const request: JmapRequest = {
        using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:calendars'],
        methodCalls: [
          ['CalendarEvent/set', {
            accountId: session.accountId,
            destroy: [eventId],
          }, 'deleteEvent'],
        ],
      };
      const response = await this.makeRequest(request);
      const result = this.getMethodResult(response, 0);
      if (result.notDestroyed?.[eventId]) {
        throw new Error(`Failed to delete event: ${JSON.stringify(result.notDestroyed[eventId])}`);
      }
    } else {
      if (!instanceStart) {
        throw new Error('instanceStart is required when scope is "this"');
      }
      // JMAP recurrenceId is LocalDateTime (YYYY-MM-DDTHH:MM:SS, no Z)
      const recurrenceId = instanceStart.replace(/Z$/, '').substring(0, 19);
      const request: JmapRequest = {
        using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:calendars'],
        methodCalls: [
          ['CalendarEvent/set', {
            accountId: session.accountId,
            update: {
              [eventId]: {
                [`recurrenceOverrides/${recurrenceId}`]: { excluded: true },
              },
            },
          }, 'updateEvent'],
        ],
      };
      const response = await this.makeRequest(request);
      const result = this.getMethodResult(response, 0);
      if (result.notUpdated?.[eventId]) {
        throw new Error(`Failed to exclude occurrence: ${JSON.stringify(result.notUpdated[eventId])}`);
      }
    }
  }

  async moveCalendarEvent(eventId: string, targetCalendarId: string, scope: 'this' | 'all', instanceStart?: string): Promise<void> {
    const hasPermission = await this.checkCalendarsPermission();
    if (!hasPermission) {
      throw new Error('Calendar access not available. This account may not have JMAP calendar permissions enabled.');
    }
    const session = await this.getSession();

    if (scope === 'all') {
      const request: JmapRequest = {
        using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:calendars'],
        methodCalls: [
          ['CalendarEvent/set', {
            accountId: session.accountId,
            update: {
              [eventId]: {
                calendarIds: { [targetCalendarId]: true },
              },
            },
          }, 'moveEvent'],
        ],
      };
      const response = await this.makeRequest(request);
      const result = this.getMethodResult(response, 0);
      if (result.notUpdated?.[eventId]) {
        throw new Error(`Failed to move event: ${JSON.stringify(result.notUpdated[eventId])}`);
      }
    } else {
      if (!instanceStart) {
        throw new Error('instanceStart is required when scope is "this"');
      }

      // Fetch source event to get duration
      const sourceEvent = await this.getCalendarEventById(eventId);
      if (!sourceEvent) {
        throw new Error(`Calendar event not found: ${eventId}`);
      }

      // Compute occurrence end from RFC 8984 duration or start/end
      let durationMs = 0;
      if (sourceEvent.duration) {
        const m = String(sourceEvent.duration).match(
          /^P(?:(\d+)Y)?(?:(\d+)M)?(?:(\d+)D)?(?:T(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?)?$/
        );
        if (m) {
          durationMs =
            (Number(m[1]) || 0) * 365 * 24 * 3600000 +
            (Number(m[2]) || 0) * 30 * 24 * 3600000 +
            (Number(m[3]) || 0) * 24 * 3600000 +
            (Number(m[4]) || 0) * 3600000 +
            (Number(m[5]) || 0) * 60000 +
            (Number(m[6]) || 0) * 1000;
        }
      } else if (sourceEvent.start && sourceEvent.end) {
        durationMs = new Date(sourceEvent.end).getTime() - new Date(sourceEvent.start).getTime();
      }

      const occStartLocal = instanceStart.replace(/Z$/, '').substring(0, 19);
      const occEndDate = new Date(new Date(instanceStart).getTime() + durationMs);
      const occEndLocal = occEndDate.toISOString().replace(/Z$/, '').substring(0, 19);

      const recurrenceId = occStartLocal;

      // Create detached occurrence in target calendar and exclude from source in one batch
      const request: JmapRequest = {
        using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:calendars'],
        methodCalls: [
          ['CalendarEvent/set', {
            accountId: session.accountId,
            create: {
              newOccurrence: {
                calendarIds: { [targetCalendarId]: true },
                title: sourceEvent.title,
                description: sourceEvent.description,
                start: occStartLocal,
                end: occEndLocal,
                location: sourceEvent.location,
              },
            },
            update: {
              [eventId]: {
                [`recurrenceOverrides/${recurrenceId}`]: { excluded: true },
              },
            },
          }, 'moveOccurrence'],
        ],
      };
      const response = await this.makeRequest(request);
      const result = this.getMethodResult(response, 0);
      if (result.notUpdated?.[eventId]) {
        throw new Error(`Failed to exclude occurrence from source: ${JSON.stringify(result.notUpdated[eventId])}`);
      }
    }
  }

  /**
   * Find an existing calendar event that looks like a duplicate of the one
   * about to be created. Use before createCalendarEvent to avoid creating
   * duplicates from email re-processing or overlapping triage flows.
   *
   * Matching is on title (case-insensitive, with substring fallback) within
   * the event's date window (inclusive). Searches across all provided
   * calendar names (or all calendars if omitted).
   */
  async findDuplicateEvent(params: {
    title: string;
    start: string;
    end?: string;
    calendarNames?: string[];
    timezone?: string;
  }): Promise<{ found: boolean; event: any | null; searched: string[] }> {
    const tzIfBare = params.timezone || defaultTz();
    const startNorm = normalizeEventDateTime(params.start, tzIfBare);
    const endNorm = params.end ? normalizeEventDateTime(params.end, tzIfBare) : startNorm;

    // Date window: allow ±1 day around the start so we catch events with
    // slightly different times but the same day. Callers doing tighter
    // matching can pass a narrower end.
    const windowStart = startNorm.start.slice(0, 10);
    const windowEnd = endNorm.start.slice(0, 10);

    const titleLower = params.title.trim().toLowerCase();
    if (!titleLower) throw new Error('title is required for duplicate detection');

    // Resolve which calendars to search
    let calendarIds: string[];
    if (params.calendarNames && params.calendarNames.length > 0) {
      calendarIds = await Promise.all(params.calendarNames.map(n => this.resolveCalendarId(n)));
    } else {
      const all = await this.getCalendars();
      calendarIds = all.map((c: any) => c.id);
    }

    const searched: string[] = [];
    for (const calId of calendarIds) {
      searched.push(calId);
      const events = await this.getCalendarEvents(calId, 200);
      for (const ev of events) {
        const evTitle = (ev.title || '').trim().toLowerCase();
        if (!evTitle) continue;

        const evStart = (ev.start || '').slice(0, 10);
        if (evStart < windowStart || evStart > windowEnd) continue;

        if (evTitle === titleLower || evTitle.includes(titleLower) || titleLower.includes(evTitle)) {
          return { found: true, event: ev, searched };
        }
      }
    }

    return { found: false, event: null, searched };
  }

  async createCalendarEvent(event: {
    calendarId: string;
    title: string;
    description?: string;
    start: string; // Bare local, ISO with Z, ISO with offset, or YYYY-MM-DD
    end: string;
    location?: string;
    participants?: Array<{ email: string; name?: string }>;
    timezone?: string; // Override the default for bare-local inputs
  }): Promise<string> {
    // Check permissions first
    const hasPermission = await this.checkCalendarsPermission();
    if (!hasPermission) {
      throw new Error('Calendar access not available. This account may not have JMAP calendar permissions enabled. Please check your Fastmail account settings or contact support to enable calendar API access.');
    }

    const session = await this.getSession();

    // Resolve calendar name → JMAP calendar ID (pass-through if already an ID).
    const resolvedCalendarId = await this.resolveCalendarId(event.calendarId);

    // Normalise to RFC 8984 LocalDateTime + timeZone. Fixes the bug where
    // offset-bearing inputs were passed through raw and silently rejected.
    const tzIfBare = event.timezone || defaultTz();
    const startNorm = normalizeEventDateTime(event.start, tzIfBare);
    const endNorm = normalizeEventDateTime(event.end, tzIfBare);

    const eventObject: Record<string, unknown> = {
      calendarId: resolvedCalendarId,
      title: event.title,
      description: event.description || '',
      start: startNorm.start,
      end: endNorm.start,
      location: event.location || '',
      participants: event.participants || []
    };
    if (startNorm.timeZone) {
      eventObject.timeZone = startNorm.timeZone;
    }

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:calendars'],
      methodCalls: [
        ['CalendarEvent/set', {
          accountId: session.accountId,
          create: { newEvent: eventObject }
        }, 'createEvent']
      ]
    };

    try {
      const response = await this.makeRequest(request);
      const result = this.getMethodResult(response, 0);
      const eventId = result.created?.newEvent?.id;
      if (!eventId) {
        throw new Error('Calendar event creation returned no event ID');
      }
      return eventId;
    } catch (error) {
      throw new Error(`Calendar event creation not supported: ${error instanceof Error ? error.message : String(error)}. Try checking account permissions or enabling calendar API access in Fastmail settings.`);
    }
  }
}