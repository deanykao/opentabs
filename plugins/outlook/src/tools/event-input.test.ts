import { describe, expect, test } from 'vitest';
import { pickOnlineMeetingProvider } from '../calendar-settings.js';
import { extractJoinUrl, mapRecurrence } from './calendar-schemas.js';
import { buildEventBody, buildRecurrence } from './event-input.js';

const TZ = 'Central Standard Time';

describe('buildEventBody', () => {
  test('applies the resolved time zone to start and end', () => {
    const body = buildEventBody({ start: '2026-06-02T13:00:00', end: '2026-06-02T14:00:00' }, { timeZone: TZ });
    expect(body.start).toEqual({ dateTime: '2026-06-02T13:00:00', timeZone: TZ });
    expect(body.end).toEqual({ dateTime: '2026-06-02T14:00:00', timeZone: TZ });
  });

  test('requests the online meeting provider only when turning a meeting online', () => {
    const ctx = { timeZone: TZ, onlineMeetingProvider: 'teamsForBusiness' };
    expect(buildEventBody({ is_online_meeting: true }, ctx)).toEqual({
      isOnlineMeeting: true,
      onlineMeetingProvider: 'teamsForBusiness',
    });
    expect(buildEventBody({ is_online_meeting: false }, ctx)).toEqual({ isOnlineMeeting: false });
  });

  test('includes only provided fields so updates do not clobber others', () => {
    expect(buildEventBody({ subject: 'x' }, { timeZone: TZ })).toEqual({ subject: 'x' });
  });

  test('room location carries the room mailbox', () => {
    expect(buildEventBody({ location_email: 'room@x.com' }, { timeZone: TZ }).location).toEqual({
      displayName: 'room@x.com',
      locationEmailAddress: 'room@x.com',
    });
  });

  test('reminder lead time turns the reminder on unless explicitly off', () => {
    expect(buildEventBody({ reminder_minutes_before_start: 5 }, { timeZone: TZ })).toEqual({
      reminderMinutesBeforeStart: 5,
      isReminderOn: true,
    });
    expect(buildEventBody({ is_reminder_on: false }, { timeZone: TZ })).toEqual({ isReminderOn: false });
  });
});

describe('buildRecurrence', () => {
  test('weekly series ending after N occurrences', () => {
    expect(
      buildRecurrence(
        {
          pattern: { type: 'weekly', days_of_week: ['tuesday'] },
          range: { type: 'numbered', start_date: '2026-06-02', number_of_occurrences: 2 },
        },
        TZ,
      ),
    ).toEqual({
      pattern: { type: 'weekly', interval: 1, daysOfWeek: ['tuesday'], firstDayOfWeek: 'sunday' },
      range: { type: 'numbered', startDate: '2026-06-02', numberOfOccurrences: 2, recurrenceTimeZone: TZ },
    });
  });

  test('relative monthly carries days and week index', () => {
    const r = buildRecurrence(
      {
        pattern: { type: 'relativeMonthly', interval: 1, days_of_week: ['tuesday'], index: 'second' },
        range: { type: 'endDate', start_date: '2026-06-09', end_date: '2026-12-31' },
      },
      TZ,
    );
    expect(r.pattern).toEqual({ type: 'relativeMonthly', interval: 1, daysOfWeek: ['tuesday'], index: 'second' });
    expect(r.range).toEqual({
      type: 'endDate',
      startDate: '2026-06-09',
      endDate: '2026-12-31',
      recurrenceTimeZone: TZ,
    });
  });

  test.each([
    [{ type: 'weekly' as const }, { type: 'noEnd' as const, start_date: '2026-06-02' }, /days_of_week/],
    [{ type: 'absoluteMonthly' as const }, { type: 'noEnd' as const, start_date: '2026-06-02' }, /day_of_month/],
    [
      { type: 'absoluteYearly' as const, day_of_month: 2 },
      { type: 'noEnd' as const, start_date: '2026-06-02' },
      /month/,
    ],
    [{ type: 'daily' as const }, { type: 'endDate' as const, start_date: '2026-06-02' }, /end_date/],
    [{ type: 'daily' as const }, { type: 'numbered' as const, start_date: '2026-06-02' }, /number_of_occurrences/],
  ])('rejects %o / %o missing its required part', (pattern, range, message) => {
    expect(() => buildRecurrence({ pattern, range }, TZ)).toThrow(message);
  });
});

describe('pickOnlineMeetingProvider', () => {
  test('falls back to Teams when the default provider is unknown (REST casing)', () => {
    expect(
      pickOnlineMeetingProvider({
        allowedOnlineMeetingProviders: ['TeamsForBusiness'],
        defaultOnlineMeetingProvider: 'Unknown',
      }),
    ).toBe('teamsForBusiness');
  });

  test('uses the default provider when it is allowed', () => {
    expect(
      pickOnlineMeetingProvider({
        allowedOnlineMeetingProviders: ['skypeForBusiness', 'teamsForBusiness'],
        defaultOnlineMeetingProvider: 'skypeForBusiness',
      }),
    ).toBe('skypeForBusiness');
  });

  test('returns undefined when no provider is allowed', () => {
    expect(pickOnlineMeetingProvider({ allowedOnlineMeetingProviders: ['unknown'] })).toBeUndefined();
  });
});

describe('extractJoinUrl', () => {
  test('prefers structured online meeting data', () => {
    expect(extractJoinUrl({ onlineMeeting: { joinUrl: 'https://a' }, onlineMeetingUrl: 'https://b' })).toBe(
      'https://a',
    );
  });

  test('finds the Teams link in an add-in meeting body and decodes HTML entities', () => {
    const content =
      '<a href="https://teams.microsoft.com/l/meetup-join/19%3ameeting_x%40thread.v2/0?context=%7b%7d&amp;tenant=1" class="me-email-headline">Join</a>';
    expect(extractJoinUrl({ onlineMeeting: null, body: { content } })).toBe(
      'https://teams.microsoft.com/l/meetup-join/19%3ameeting_x%40thread.v2/0?context=%7b%7d&tenant=1',
    );
  });

  test('finds a Zoom link in a plain-text body, skipping unrelated links', () => {
    const content =
      'Agenda: https://intranet.example.com/doc\r\nJoin Zoom Meeting\r\nhttps://bnsf.zoom.us/j/123456789?pwd=abc\r\n';
    expect(extractJoinUrl({ body: { content } })).toBe('https://bnsf.zoom.us/j/123456789?pwd=abc');
  });

  test('falls back to the body preview when the body was not selected', () => {
    expect(extractJoinUrl({ bodyPreview: 'Join Zoom Meeting\r\nhttps://bnsf.zoom.us/j/1?pwd=x\r\n' })).toBe(
      'https://bnsf.zoom.us/j/1?pwd=x',
    );
  });

  test('returns empty when there is no link', () => {
    expect(extractJoinUrl({ body: { content: 'no link' } })).toBe('');
  });
});

describe('mapRecurrence', () => {
  test('normalizes REST casing and blanks the no-end sentinel date', () => {
    expect(
      mapRecurrence({
        pattern: { type: 'Weekly', interval: 1, daysOfWeek: ['Tuesday'], firstDayOfWeek: 'Sunday', index: 'First' },
        range: { type: 'NoEnd', startDate: '2026-06-02', endDate: '0001-01-01', recurrenceTimeZone: TZ },
      }),
    ).toEqual({
      pattern: {
        type: 'weekly',
        interval: 1,
        days_of_week: ['tuesday'],
        day_of_month: 0,
        month: 0,
        index: 'first',
        first_day_of_week: 'sunday',
      },
      range: {
        type: 'noEnd',
        start_date: '2026-06-02',
        end_date: '',
        number_of_occurrences: 0,
        recurrence_time_zone: TZ,
      },
    });
  });
});
