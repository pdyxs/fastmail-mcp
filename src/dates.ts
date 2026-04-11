import { readlinkSync } from 'fs';

/**
 * Read the system timezone from /etc/localtime, as ticktick-mcp does.
 * Falls back to UTC if unreadable.
 */
function systemTz(): string {
  try {
    const link = readlinkSync('/etc/localtime');
    const idx = link.indexOf('/zoneinfo/');
    return idx >= 0 ? link.slice(idx + '/zoneinfo/'.length) : 'UTC';
  } catch {
    return 'UTC';
  }
}

/**
 * The default timezone used when the caller passes a bare datetime with no
 * offset or Z suffix. Set FASTMAIL_TIMEZONE in the MCP launch wrapper to
 * override (e.g. Australia/Sydney).
 */
export function defaultTz(): string {
  return process.env.FASTMAIL_TIMEZONE || systemTz();
}

export interface NormalizedEventDateTime {
  /** JMAP LocalDateTime: YYYY-MM-DDTHH:MM:SS, no Z, no offset */
  start: string;
  /** IANA timezone name (e.g. "Australia/Sydney", "Etc/UTC"), or undefined for all-day */
  timeZone: string | undefined;
  isAllDay: boolean;
}

const BARE_LOCAL_RE = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}(:\d{2})?(\.\d+)?$/;
const ALL_DAY_RE = /^\d{4}-\d{2}-\d{2}$/;
const Z_SUFFIX_RE = /Z$/;
const OFFSET_SUFFIX_RE = /[+-]\d{2}:?\d{2}$/;

/**
 * Normalise a caller-provided datetime string into the JMAP CalendarEvent
 * shape (RFC 8984): a LocalDateTime without timezone information, plus a
 * separate timeZone field.
 *
 * Accepted inputs:
 *   - YYYY-MM-DD                         → all-day, no timeZone
 *   - YYYY-MM-DDTHH:MM:SS                → interpreted in `tzIfBare`
 *   - YYYY-MM-DDTHH:MM:SSZ               → UTC
 *   - YYYY-MM-DDTHH:MM:SS+11:00          → converted to UTC
 *
 * Fixes the bug where offset-bearing inputs were previously passed through
 * to JMAP as-is and silently rejected.
 */
export function normalizeEventDateTime(
  input: string,
  tzIfBare: string,
): NormalizedEventDateTime {
  const trimmed = input.trim();

  if (ALL_DAY_RE.test(trimmed)) {
    return { start: trimmed, timeZone: undefined, isAllDay: true };
  }

  if (BARE_LOCAL_RE.test(trimmed)) {
    const withSeconds = trimmed.length === 16 ? `${trimmed}:00` : trimmed;
    return { start: withSeconds.slice(0, 19), timeZone: tzIfBare, isAllDay: false };
  }

  if (Z_SUFFIX_RE.test(trimmed)) {
    return { start: trimmed.replace(Z_SUFFIX_RE, '').slice(0, 19), timeZone: 'Etc/UTC', isAllDay: false };
  }

  if (OFFSET_SUFFIX_RE.test(trimmed)) {
    // Convert to UTC wall time via Date, then strip.
    const d = new Date(trimmed);
    if (isNaN(d.getTime())) {
      throw new Error(`Invalid datetime: ${input}`);
    }
    return { start: d.toISOString().replace(Z_SUFFIX_RE, '').slice(0, 19), timeZone: 'Etc/UTC', isAllDay: false };
  }

  throw new Error(`Invalid datetime: ${input}. Expected YYYY-MM-DD, YYYY-MM-DDTHH:MM[:SS], or ISO 8601 with Z / offset.`);
}
