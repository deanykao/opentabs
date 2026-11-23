import { z } from 'zod';
import { type RawEmailAddress, emailAddressSchema, mapEmailAddress } from './schemas.js';

// ── Input date / time strings ─────────────────────────────────────────────
//
// Plain regex patterns, never `z.iso.*`: those emit a JSON Schema `format` keyword
// that the MCP server's strict Ajv validator rejects ("unknown format"), which makes
// the whole tool uncallable. `date-time` would also demand an offset, which local
// times paired with a named time zone deliberately omit.

const DATE = '\\d{4}-(?:0[1-9]|1[0-2])-(?:0[1-9]|[12]\\d|3[01])';
const TIME = '(?:[01]\\d|2[0-3]):[0-5]\\d(?::[0-5]\\d(?:\\.\\d{1,7})?)?';

/** Local wall-clock date-time with no offset, e.g. `2026-06-02T13:00:00`. */
export const localDateTimeSchema = z
  .string()
  .regex(
    new RegExp(`^${DATE}T${TIME}$`),
    'Expected a local ISO 8601 date-time without offset, e.g. 2026-06-02T13:00:00',
  );

/** Date-time with an optional `Z` or `±hh:mm` offset, e.g. `2026-06-02T00:00:00-05:00`. */
export const offsetDateTimeSchema = z
  .string()
  .regex(
    new RegExp(`^${DATE}T${TIME}(?:Z|[+-](?:[01]\\d|2[0-3]):[0-5]\\d)?$`),
    'Expected an ISO 8601 date-time, e.g. 2026-06-02T00:00:00-05:00',
  );

/** Calendar date, e.g. `2026-06-02`. */
export const dateOnlySchema = z.string().regex(new RegExp(`^${DATE}$`), 'Expected a date as YYYY-MM-DD');

/**
 * Outlook REST returns PascalCase enum values (`Busy`, `TeamsForBusiness`) where
 * Graph returns camelCase (`busy`, `teamsForBusiness`). Normalize to Graph's form.
 */
export const toCamelEnum = (value: string): string => value.charAt(0).toLowerCase() + value.slice(1);

// ── Date / time ───────────────────────────────────────────────────────────

export const dateTimeZoneSchema = z.object({
  date_time: z.string().describe('Local date-time in ISO 8601 without offset (e.g. "2026-06-02T13:00:00")'),
  time_zone: z.string().describe('Time zone name (e.g. "UTC", "Eastern Standard Time")'),
});

export interface RawDateTimeZone {
  dateTime?: string;
  timeZone?: string;
}

export const mapDateTimeZone = (d: RawDateTimeZone | undefined | null) => ({
  date_time: d?.dateTime ?? '',
  time_zone: d?.timeZone ?? '',
});

// ── Attendee ───────────────────────────────────────────────────────────────

export const attendeeSchema = z.object({
  name: z.string().describe('Attendee display name'),
  address: z.string().describe('Attendee email address'),
  type: z.string().describe('Attendance type (required, optional, resource)'),
  response: z
    .string()
    .describe('Response status (none, accepted, declined, tentativelyAccepted, notResponded, organizer)'),
});

export interface RawAttendee extends RawEmailAddress {
  type?: string;
  status?: { response?: string; time?: string };
}

export const mapAttendee = (a: RawAttendee) => {
  const { name, address } = mapEmailAddress(a);
  return {
    name,
    address,
    type: toCamelEnum(a.type ?? 'required'),
    response: toCamelEnum(a.status?.response ?? 'none'),
  };
};

// ── Calendar ───────────────────────────────────────────────────────────────

export const calendarSchema = z.object({
  id: z.string().describe('Calendar ID'),
  name: z.string().describe('Calendar display name'),
  color: z.string().describe('Color preset name (auto, lightBlue, lightGreen, ...)'),
  hex_color: z.string().describe('Color as a hex string when available'),
  is_default: z.boolean().describe('Whether this is the default calendar'),
  can_edit: z.boolean().describe('Whether the user can edit events in this calendar'),
  can_share: z.boolean().describe('Whether the user can share this calendar'),
  can_view_private_items: z.boolean().describe('Whether the user can view private events'),
  owner: emailAddressSchema.describe('Calendar owner'),
});

export interface RawCalendar {
  id?: string;
  name?: string;
  color?: string;
  hexColor?: string;
  isDefaultCalendar?: boolean;
  canEdit?: boolean;
  canShare?: boolean;
  canViewPrivateItems?: boolean;
  owner?: { name?: string; address?: string };
}

export const mapCalendar = (c: RawCalendar) => ({
  id: c.id ?? '',
  name: c.name ?? '',
  color: toCamelEnum(c.color ?? 'auto'),
  hex_color: c.hexColor ?? '',
  is_default: c.isDefaultCalendar ?? false,
  can_edit: c.canEdit ?? false,
  can_share: c.canShare ?? false,
  can_view_private_items: c.canViewPrivateItems ?? false,
  owner: {
    name: c.owner?.name ?? '',
    address: c.owner?.address ?? '',
  },
});

// ── Recurrence ─────────────────────────────────────────────────────────────

export const recurrenceSchema = z.object({
  pattern: z.object({
    type: z.string().describe('daily, weekly, absoluteMonthly, relativeMonthly, absoluteYearly, relativeYearly'),
    interval: z.number().describe('Units between occurrences'),
    days_of_week: z.array(z.string()).describe('Days the event recurs on (weekly / relative patterns)'),
    day_of_month: z.number().describe('Day of the month (absolute patterns; 0 when unused)'),
    month: z.number().describe('Month of the year (yearly patterns; 0 when unused)'),
    index: z.string().describe('Week of the month for relative patterns (first..last)'),
    first_day_of_week: z.string().describe('First day of the week'),
  }),
  range: z.object({
    type: z.string().describe('endDate, noEnd, or numbered'),
    start_date: z.string().describe('First date of the series (YYYY-MM-DD)'),
    end_date: z.string().describe('Last date of the series for endDate ranges'),
    number_of_occurrences: z.number().describe('Occurrence count for numbered ranges (0 when unused)'),
    recurrence_time_zone: z.string().describe('Time zone the recurrence is evaluated in'),
  }),
});

export interface RawRecurrence {
  pattern?: {
    type?: string;
    interval?: number;
    daysOfWeek?: string[];
    dayOfMonth?: number;
    month?: number;
    index?: string;
    firstDayOfWeek?: string;
  };
  range?: {
    type?: string;
    startDate?: string;
    endDate?: string;
    numberOfOccurrences?: number;
    recurrenceTimeZone?: string;
  };
}

export const mapRecurrence = (r: RawRecurrence | null | undefined) =>
  r?.pattern
    ? {
        pattern: {
          type: toCamelEnum(r.pattern.type ?? ''),
          interval: r.pattern.interval ?? 1,
          days_of_week: (r.pattern.daysOfWeek ?? []).map(toCamelEnum),
          day_of_month: r.pattern.dayOfMonth ?? 0,
          month: r.pattern.month ?? 0,
          index: toCamelEnum(r.pattern.index ?? 'first'),
          first_day_of_week: toCamelEnum(r.pattern.firstDayOfWeek ?? 'sunday'),
        },
        range: {
          type: toCamelEnum(r.range?.type ?? ''),
          start_date: r.range?.startDate ?? '',
          // Graph and REST report "0001-01-01" for ranges without an end date.
          end_date: r.range?.endDate && !r.range.endDate.startsWith('0001-') ? r.range.endDate : '',
          number_of_occurrences: r.range?.numberOfOccurrences ?? 0,
          recurrence_time_zone: r.range?.recurrenceTimeZone ?? '',
        },
      }
    : null;

// ── Event ──────────────────────────────────────────────────────────────────

export const eventSummarySchema = z.object({
  id: z.string().describe('Event ID'),
  subject: z.string().describe('Event subject/title'),
  start: dateTimeZoneSchema.describe('Start date-time'),
  end: dateTimeZoneSchema.describe('End date-time'),
  is_all_day: z.boolean().describe('Whether the event lasts all day'),
  location: z.string().describe('Primary location display name'),
  organizer: emailAddressSchema.describe('Event organizer'),
  is_cancelled: z.boolean().describe('Whether the event has been cancelled'),
  show_as: z.string().describe('Free/busy status (free, tentative, busy, oof, workingElsewhere, unknown)'),
  is_online_meeting: z.boolean().describe('Whether the event is an online meeting'),
  online_meeting_url: z
    .string()
    .describe('Join URL for the online meeting, if any (also detected from a Teams/Zoom/Webex/Meet link in the body)'),
  type: z.string().describe('Event type (singleInstance, occurrence, exception, seriesMaster)'),
  response_status: z
    .string()
    .describe("The current user's response (none, accepted, declined, tentativelyAccepted, organizer, notResponded)"),
  web_link: z.string().describe('Link to open the event in Outlook on the web'),
  preview: z.string().describe('Body preview text'),
});

export const eventDetailSchema = eventSummarySchema.extend({
  body_type: z.string().describe('Body content type (text or html)'),
  body: z.string().describe('Full event body'),
  attendees: z.array(attendeeSchema).describe('Event attendees and their responses'),
  importance: z.string().describe('Importance level (low, normal, high)'),
  sensitivity: z.string().describe('Sensitivity (normal, personal, private, confidential)'),
  categories: z.array(z.string()).describe('Categories/labels'),
  is_reminder_on: z.boolean().describe('Whether a reminder is set'),
  reminder_minutes_before_start: z.number().describe('Minutes before start the reminder fires'),
  has_attachments: z.boolean().describe('Whether the event has attachments'),
  response_requested: z.boolean().describe('Whether the organizer requested a response'),
  allow_new_time_proposals: z.boolean().describe('Whether invitees may propose a new time'),
  series_master_id: z.string().describe('ID of the series master, for occurrences/exceptions'),
  is_recurring: z.boolean().describe('Whether the event is part of a recurring series'),
  recurrence: recurrenceSchema.nullable().describe('Recurrence pattern and range, for series masters'),
  created_at: z.string().describe('Creation datetime (ISO 8601)'),
  last_modified_at: z.string().describe('Last modified datetime (ISO 8601)'),
});

export interface RawEvent {
  id?: string;
  subject?: string;
  bodyPreview?: string;
  body?: { contentType?: string; content?: string };
  start?: RawDateTimeZone;
  end?: RawDateTimeZone;
  isAllDay?: boolean;
  location?: { displayName?: string };
  organizer?: RawEmailAddress;
  attendees?: RawAttendee[];
  isCancelled?: boolean;
  isOnlineMeeting?: boolean;
  onlineMeeting?: { joinUrl?: string } | null;
  onlineMeetingUrl?: string | null;
  showAs?: string;
  type?: string;
  responseStatus?: { response?: string; time?: string };
  importance?: string;
  sensitivity?: string;
  categories?: string[];
  isReminderOn?: boolean;
  reminderMinutesBeforeStart?: number;
  hasAttachments?: boolean;
  responseRequested?: boolean;
  allowNewTimeProposals?: boolean;
  seriesMasterId?: string;
  recurrence?: RawRecurrence | null;
  webLink?: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
}

/** Join links of the common meeting services, as they appear in invitation bodies. */
const MEETING_JOIN_URL = new RegExp(
  [
    'https://teams\\.(?:microsoft|live)\\.com/(?:l/meetup-join|meet)/',
    'https://[\\w.-]*zoom\\.us/(?:j|w|my)/',
    'https://[\\w.-]+\\.webex\\.com/(?:meet|join|[\\w.-]+/j\\.php)',
    'https://meet\\.google\\.com/[a-z]',
  ]
    .map(prefix => `${prefix}[^\\s"'<>]*`)
    .join('|'),
  'i',
);

/**
 * The event's online-meeting join URL. Meetings added by a meeting-service add-in
 * (Teams, Zoom, Webex) carry no `onlineMeeting` data — the link only exists in the
 * body, so fall back to the first join link in the body, then the body preview.
 */
export const extractJoinUrl = (e: RawEvent): string => {
  const structured = e.onlineMeeting?.joinUrl || e.onlineMeetingUrl;
  if (structured) return structured;
  const match = e.body?.content?.match(MEETING_JOIN_URL) ?? e.bodyPreview?.match(MEETING_JOIN_URL);
  return match ? match[0].replace(/&amp;/g, '&') : '';
};

export const mapEventSummary = (e: RawEvent) => {
  const joinUrl = extractJoinUrl(e);
  return {
    id: e.id ?? '',
    subject: e.subject ?? '(no subject)',
    start: mapDateTimeZone(e.start),
    end: mapDateTimeZone(e.end),
    is_all_day: e.isAllDay ?? false,
    location: e.location?.displayName ?? '',
    organizer: mapEmailAddress(e.organizer ?? {}),
    is_cancelled: e.isCancelled ?? false,
    show_as: toCamelEnum(e.showAs ?? 'unknown'),
    is_online_meeting: (e.isOnlineMeeting ?? false) || joinUrl !== '',
    online_meeting_url: joinUrl,
    type: toCamelEnum(e.type ?? 'singleInstance'),
    response_status: toCamelEnum(e.responseStatus?.response ?? 'none'),
    web_link: e.webLink ?? '',
    preview: e.bodyPreview ?? '',
  };
};

export const mapEventDetail = (e: RawEvent) => {
  const summary = mapEventSummary(e);
  return {
    ...summary,
    body_type: toCamelEnum(e.body?.contentType ?? 'text'),
    body: e.body?.content ?? '',
    attendees: (e.attendees ?? []).map(mapAttendee),
    importance: toCamelEnum(e.importance ?? 'normal'),
    sensitivity: toCamelEnum(e.sensitivity ?? 'normal'),
    categories: e.categories ?? [],
    is_reminder_on: e.isReminderOn ?? false,
    reminder_minutes_before_start: e.reminderMinutesBeforeStart ?? 0,
    has_attachments: e.hasAttachments ?? false,
    response_requested: e.responseRequested ?? false,
    allow_new_time_proposals: e.allowNewTimeProposals ?? true,
    series_master_id: e.seriesMasterId ?? '',
    is_recurring: Boolean(e.recurrence) || ['seriesMaster', 'occurrence', 'exception'].includes(summary.type),
    recurrence: mapRecurrence(e.recurrence),
    created_at: e.createdDateTime ?? '',
    last_modified_at: e.lastModifiedDateTime ?? '',
  };
};

// ── Request-body builders ───────────────────────────────────────────────────

/** Build a Graph/REST dateTimeTimeZone object. */
export const buildDateTime = (dateTime: string, timeZone: string) => ({ dateTime, timeZone });

/** Attendee shape accepted by create_event / update_event / find_meeting_times input. */
export const attendeeInputSchema = z.object({
  address: z.string().describe('Attendee email address'),
  name: z.string().optional().describe('Attendee display name'),
  type: z
    .enum(['required', 'optional', 'resource'])
    .optional()
    .describe('Attendance type (default: required). Use resource for rooms/equipment.'),
});

export interface AttendeeInput {
  address: string;
  name?: string;
  type?: 'required' | 'optional' | 'resource';
}

/** Build the attendees array for an event request body. */
export const buildAttendees = (attendees: AttendeeInput[]) =>
  attendees.map(a => ({
    emailAddress: { address: a.address, name: a.name },
    type: a.type ?? 'required',
  }));

// ── Schedule (free/busy) ─────────────────────────────────────────────────

export const scheduleItemSchema = z.object({
  status: z.string().describe('Free/busy status (free, tentative, busy, oof, workingElsewhere, unknown)'),
  start: dateTimeZoneSchema.describe('Slot start'),
  end: dateTimeZoneSchema.describe('Slot end'),
  subject: z.string().describe('Subject, when visible to the requester'),
  location: z.string().describe('Location, when visible'),
  is_private: z.boolean().describe('Whether the item is marked private'),
  is_meeting: z.boolean().describe('Whether the item is a meeting'),
  is_recurring: z.boolean().describe('Whether the item recurs'),
});

export const workingHoursSchema = z.object({
  days_of_week: z.array(z.string()).describe('Working days'),
  start_time: z.string().describe('Work day start time'),
  end_time: z.string().describe('Work day end time'),
  time_zone: z.string().describe('Working-hours time zone name'),
});

export const scheduleSchema = z.object({
  schedule_id: z.string().describe('The email address this schedule belongs to'),
  availability_view: z
    .string()
    .describe('Per-interval availability string: 0=free, 1=tentative, 2=busy, 3=oof, 4=workingElsewhere'),
  items: z.array(scheduleItemSchema).describe('Busy/tentative/oof time blocks in the requested window'),
  working_hours: workingHoursSchema.nullable().describe('Working hours, when available'),
  error: z.string().describe('Error message for this schedule, if the lookup failed'),
});

export interface RawScheduleItem {
  status?: string;
  start?: RawDateTimeZone;
  end?: RawDateTimeZone;
  subject?: string;
  location?: string;
  isPrivate?: boolean;
  isMeeting?: boolean;
  isRecurring?: boolean;
}

export interface RawSchedule {
  scheduleId?: string;
  availabilityView?: string;
  scheduleItems?: RawScheduleItem[];
  workingHours?: {
    daysOfWeek?: string[];
    startTime?: string;
    endTime?: string;
    timeZone?: { name?: string };
  };
  error?: { message?: string; responseCode?: string };
}

export const mapSchedule = (s: RawSchedule) => ({
  schedule_id: s.scheduleId ?? '',
  availability_view: s.availabilityView ?? '',
  items: (s.scheduleItems ?? []).map(item => ({
    status: toCamelEnum(item.status ?? 'unknown'),
    start: mapDateTimeZone(item.start),
    end: mapDateTimeZone(item.end),
    subject: item.subject ?? '',
    location: item.location ?? '',
    is_private: item.isPrivate ?? false,
    is_meeting: item.isMeeting ?? false,
    is_recurring: item.isRecurring ?? false,
  })),
  working_hours: s.workingHours
    ? {
        days_of_week: (s.workingHours.daysOfWeek ?? []).map(toCamelEnum),
        start_time: s.workingHours.startTime ?? '',
        end_time: s.workingHours.endTime ?? '',
        time_zone: s.workingHours.timeZone?.name ?? '',
      }
    : null,
  error: s.error?.message ?? '',
});

// ── Shared field lists for $select ──────────────────────────────────────────

export const EVENT_SUMMARY_FIELDS =
  'id,subject,start,end,isAllDay,location,organizer,isCancelled,showAs,isOnlineMeeting,onlineMeeting,type,responseStatus,webLink,bodyPreview';

export const EVENT_DETAIL_FIELDS = `${EVENT_SUMMARY_FIELDS},body,attendees,importance,sensitivity,categories,isReminderOn,reminderMinutesBeforeStart,hasAttachments,responseRequested,allowNewTimeProposals,seriesMasterId,recurrence,createdDateTime,lastModifiedDateTime`;
