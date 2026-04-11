import { describe, it } from 'node:test';
import assert from 'node:assert';
import { normalizeEventDateTime } from './dates.js';

describe('normalizeEventDateTime', () => {
  it('handles bare local datetime with defaultTz', () => {
    const r = normalizeEventDateTime('2026-03-25T14:00:00', 'Australia/Sydney');
    assert.equal(r.start, '2026-03-25T14:00:00');
    assert.equal(r.timeZone, 'Australia/Sydney');
    assert.equal(r.isAllDay, false);
  });

  it('pads bare local with no seconds', () => {
    const r = normalizeEventDateTime('2026-03-25T14:00', 'Australia/Sydney');
    assert.equal(r.start, '2026-03-25T14:00:00');
    assert.equal(r.timeZone, 'Australia/Sydney');
  });

  it('strips Z suffix and marks UTC', () => {
    const r = normalizeEventDateTime('2026-03-25T03:00:00Z', 'Australia/Sydney');
    assert.equal(r.start, '2026-03-25T03:00:00');
    assert.equal(r.timeZone, 'Etc/UTC');
  });

  it('converts +11:00 offset to UTC wall time (previously rejected silently)', () => {
    const r = normalizeEventDateTime('2026-03-25T14:00:00+11:00', 'Australia/Sydney');
    assert.equal(r.start, '2026-03-25T03:00:00');
    assert.equal(r.timeZone, 'Etc/UTC');
  });

  it('converts -05:00 offset to UTC wall time', () => {
    const r = normalizeEventDateTime('2026-03-25T09:00:00-05:00', 'UTC');
    assert.equal(r.start, '2026-03-25T14:00:00');
    assert.equal(r.timeZone, 'Etc/UTC');
  });

  it('handles all-day YYYY-MM-DD', () => {
    const r = normalizeEventDateTime('2026-03-25', 'Australia/Sydney');
    assert.equal(r.start, '2026-03-25');
    assert.equal(r.timeZone, undefined);
    assert.equal(r.isAllDay, true);
  });

  it('rejects malformed input', () => {
    assert.throws(() => normalizeEventDateTime('not a date', 'UTC'), /Invalid datetime/);
    assert.throws(() => normalizeEventDateTime('2026/03/25', 'UTC'), /Invalid datetime/);
  });
});
