import { api } from './outlook-api.js';
import { toCamelEnum } from './tools/calendar-schemas.js';

let mailboxTimeZone: Promise<string | undefined> | undefined;

/**
 * The mailbox's configured time zone (a Windows name such as "Pacific Standard Time"),
 * fetched once per page load. Resolves `undefined` when the setting is unreadable;
 * a failed lookup is not cached so the next call retries.
 */
const getMailboxTimeZone = (): Promise<string | undefined> => {
  mailboxTimeZone ??= api<{ timeZone?: string }>(
    { graph: '/me/mailboxSettings', rest: '/me/MailboxSettings' },
    {},
    'calendar',
  ).then(
    settings => settings.timeZone || undefined,
    () => {
      mailboxTimeZone = undefined;
      return undefined;
    },
  );
  return mailboxTimeZone;
};

/**
 * Time zone for interpreting and returning calendar times: the caller's explicit zone,
 * else the mailbox time zone (what Outlook itself shows), else UTC.
 */
export const resolveTimeZone = async (explicit: string | undefined): Promise<string> =>
  explicit ?? (await getMailboxTimeZone()) ?? 'UTC';

export interface RawOnlineMeetingSettings {
  allowedOnlineMeetingProviders?: string[];
  defaultOnlineMeetingProvider?: string;
}

/**
 * Pick the provider to request for a new online meeting. Setting `isOnlineMeeting`
 * alone yields no join link when the calendar's default provider is `unknown` (as in
 * tenants where Teams links normally come from the Teams add-in), so an explicit
 * allowed provider is required: the default when allowed, else Teams, else any.
 */
export const pickOnlineMeetingProvider = (settings: RawOnlineMeetingSettings): string | undefined => {
  const allowed = (settings.allowedOnlineMeetingProviders ?? []).map(toCamelEnum).filter(p => p !== 'unknown');
  const preferred = toCamelEnum(settings.defaultOnlineMeetingProvider ?? 'unknown');
  if (allowed.includes(preferred)) return preferred;
  return allowed.includes('teamsForBusiness') ? 'teamsForBusiness' : allowed[0];
};

const providerByCalendar = new Map<string, Promise<string | undefined>>();

/** Online-meeting provider for a calendar (default calendar when no id), cached per page load. */
export const getOnlineMeetingProvider = (calendarId: string | undefined): Promise<string | undefined> => {
  const key = calendarId ?? '';
  let provider = providerByCalendar.get(key);
  if (!provider) {
    provider = api<RawOnlineMeetingSettings>(
      calendarId ? `/me/calendars/${encodeURIComponent(calendarId)}` : '/me/calendar',
      { query: { $select: 'allowedOnlineMeetingProviders,defaultOnlineMeetingProvider' } },
      'calendar',
    ).then(pickOnlineMeetingProvider, () => {
      providerByCalendar.delete(key);
      return undefined;
    });
    providerByCalendar.set(key, provider);
  }
  return provider;
};
