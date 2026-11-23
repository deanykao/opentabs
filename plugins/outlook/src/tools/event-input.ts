import { ToolError } from '@opentabs-dev/plugin-sdk';
import { z } from 'zod';
import {
  type AttendeeInput,
  attendeeInputSchema,
  buildAttendees,
  buildDateTime,
  dateOnlySchema,
} from './calendar-schemas.js';

const DAYS = ['sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday'] as const;

export const recurrenceInputSchema = z.object({
  pattern: z.object({
    type: z
      .enum(['daily', 'weekly', 'absoluteMonthly', 'relativeMonthly', 'absoluteYearly', 'relativeYearly'])
      .describe(
        'daily; weekly (needs days_of_week); absoluteMonthly (day_of_month); relativeMonthly (days_of_week + index, e.g. 2nd Tuesday); absoluteYearly (month + day_of_month); relativeYearly (month + days_of_week + index)',
      ),
    interval: z
      .number()
      .int()
      .min(1)
      .optional()
      .describe('Units between occurrences, e.g. 2 = every other week (default 1)'),
    days_of_week: z.array(z.enum(DAYS)).optional().describe('Days for weekly / relative patterns'),
    day_of_month: z.number().int().min(1).max(31).optional().describe('Day of month for absolute patterns'),
    month: z.number().int().min(1).max(12).optional().describe('Month for yearly patterns'),
    index: z
      .enum(['first', 'second', 'third', 'fourth', 'last'])
      .optional()
      .describe('Week of the month for relative patterns (default first)'),
    first_day_of_week: z.enum(DAYS).optional().describe('First day of the week for weekly patterns (default sunday)'),
  }),
  range: z.object({
    type: z.enum(['endDate', 'noEnd', 'numbered']).describe('How the series ends'),
    start_date: dateOnlySchema.describe('First date of the series (YYYY-MM-DD); should be the event start date'),
    end_date: dateOnlySchema.optional().describe('Last date of the series (required for endDate)'),
    number_of_occurrences: z.number().int().min(1).optional().describe('Number of occurrences (required for numbered)'),
  }),
});

export type RecurrenceInput = z.infer<typeof recurrenceInputSchema>;

/** Optional fields shared by create_event and update_event. */
export const eventFieldsInput = {
  time_zone: z
    .string()
    .optional()
    .describe(
      'Time zone for start/end and recurrence (e.g. "Central Standard Time"). Defaults to your mailbox time zone.',
    ),
  body: z.string().optional().describe('Event body/description'),
  body_type: z.enum(['text', 'html']).optional().describe('Body content type (default: text)'),
  location: z.string().optional().describe('Location display name'),
  location_email: z
    .string()
    .optional()
    .describe('Room mailbox address for the location (from list_rooms). Also add the room as a resource attendee.'),
  attendees: z.array(attendeeInputSchema).optional().describe('Attendees to invite (replaces the list on update)'),
  is_all_day: z.boolean().optional().describe('All-day event. When true, start/end must be at midnight.'),
  is_online_meeting: z
    .boolean()
    .optional()
    .describe("Attach an online meeting (Teams) join link, using the calendar's allowed provider"),
  importance: z.enum(['low', 'normal', 'high']).optional().describe('Importance level'),
  show_as: z
    .enum(['free', 'tentative', 'busy', 'oof', 'workingElsewhere'])
    .optional()
    .describe('Free/busy status to show (default: busy)'),
  sensitivity: z.enum(['normal', 'personal', 'private', 'confidential']).optional().describe('Sensitivity'),
  categories: z.array(z.string()).optional().describe('Categories/labels'),
  reminder_minutes_before_start: z.number().int().min(0).optional().describe('Reminder lead time in minutes'),
  is_reminder_on: z.boolean().optional().describe('Turn the reminder on or off'),
  response_requested: z.boolean().optional().describe('Ask invitees to respond (default true)'),
  allow_new_time_proposals: z.boolean().optional().describe('Let invitees propose a new time (default true)'),
  recurrence: recurrenceInputSchema.optional().describe('Make the event a recurring series'),
};

export interface EventFields {
  subject?: string;
  start?: string;
  end?: string;
  body?: string;
  body_type?: 'text' | 'html';
  location?: string;
  location_email?: string;
  attendees?: AttendeeInput[];
  is_all_day?: boolean;
  is_online_meeting?: boolean;
  importance?: 'low' | 'normal' | 'high';
  show_as?: 'free' | 'tentative' | 'busy' | 'oof' | 'workingElsewhere';
  sensitivity?: 'normal' | 'personal' | 'private' | 'confidential';
  categories?: string[];
  reminder_minutes_before_start?: number;
  is_reminder_on?: boolean;
  response_requested?: boolean;
  allow_new_time_proposals?: boolean;
  recurrence?: RecurrenceInput;
}

/** Build a Graph patternedRecurrence, rejecting patterns missing their required parts. */
export const buildRecurrence = (r: RecurrenceInput, timeZone: string) => {
  const { pattern, range } = r;
  const needsDays = ['weekly', 'relativeMonthly', 'relativeYearly'].includes(pattern.type);
  const needsDayOfMonth = pattern.type === 'absoluteMonthly' || pattern.type === 'absoluteYearly';
  const needsMonth = pattern.type === 'absoluteYearly' || pattern.type === 'relativeYearly';
  if (needsDays && !pattern.days_of_week?.length) {
    throw ToolError.validation(`recurrence.pattern.days_of_week is required for ${pattern.type}.`);
  }
  if (needsDayOfMonth && pattern.day_of_month === undefined) {
    throw ToolError.validation(`recurrence.pattern.day_of_month is required for ${pattern.type}.`);
  }
  if (needsMonth && pattern.month === undefined) {
    throw ToolError.validation(`recurrence.pattern.month is required for ${pattern.type}.`);
  }
  if (range.type === 'endDate' && !range.end_date) {
    throw ToolError.validation('recurrence.range.end_date is required when range.type is endDate.');
  }
  if (range.type === 'numbered' && range.number_of_occurrences === undefined) {
    throw ToolError.validation('recurrence.range.number_of_occurrences is required when range.type is numbered.');
  }

  return {
    pattern: {
      type: pattern.type,
      interval: pattern.interval ?? 1,
      ...(needsDays ? { daysOfWeek: pattern.days_of_week } : {}),
      ...(needsDayOfMonth ? { dayOfMonth: pattern.day_of_month } : {}),
      ...(needsMonth ? { month: pattern.month } : {}),
      ...(pattern.type.startsWith('relative') ? { index: pattern.index ?? 'first' } : {}),
      ...(pattern.type === 'weekly' ? { firstDayOfWeek: pattern.first_day_of_week ?? 'sunday' } : {}),
    },
    range: {
      type: range.type,
      startDate: range.start_date,
      ...(range.type === 'endDate' ? { endDate: range.end_date } : {}),
      ...(range.type === 'numbered' ? { numberOfOccurrences: range.number_of_occurrences } : {}),
      recurrenceTimeZone: timeZone,
    },
  };
};

/**
 * Build the Graph-shaped (camelCase) event body for create/update from tool input.
 * Only provided fields are included, so the same builder serves PATCH semantics.
 * `onlineMeetingProvider` is the provider to request when `is_online_meeting` is true.
 */
export const buildEventBody = (
  params: EventFields,
  ctx: { timeZone: string; onlineMeetingProvider?: string },
): Record<string, unknown> => {
  const body: Record<string, unknown> = {};
  if (params.subject !== undefined) body.subject = params.subject;
  if (params.start !== undefined) body.start = buildDateTime(params.start, ctx.timeZone);
  if (params.end !== undefined) body.end = buildDateTime(params.end, ctx.timeZone);
  if (params.body !== undefined) {
    body.body = { contentType: params.body_type === 'html' ? 'HTML' : 'Text', content: params.body };
  }
  if (params.location !== undefined || params.location_email !== undefined) {
    body.location = {
      displayName: params.location ?? params.location_email,
      ...(params.location_email ? { locationEmailAddress: params.location_email } : {}),
    };
  }
  if (params.attendees !== undefined) body.attendees = buildAttendees(params.attendees);
  if (params.is_all_day !== undefined) body.isAllDay = params.is_all_day;
  if (params.is_online_meeting !== undefined) {
    body.isOnlineMeeting = params.is_online_meeting;
    if (params.is_online_meeting && ctx.onlineMeetingProvider) {
      body.onlineMeetingProvider = ctx.onlineMeetingProvider;
    }
  }
  if (params.importance !== undefined) body.importance = params.importance;
  if (params.show_as !== undefined) body.showAs = params.show_as;
  if (params.sensitivity !== undefined) body.sensitivity = params.sensitivity;
  if (params.categories !== undefined) body.categories = params.categories;
  if (params.reminder_minutes_before_start !== undefined) {
    body.reminderMinutesBeforeStart = params.reminder_minutes_before_start;
    body.isReminderOn = params.is_reminder_on ?? true;
  } else if (params.is_reminder_on !== undefined) {
    body.isReminderOn = params.is_reminder_on;
  }
  if (params.response_requested !== undefined) body.responseRequested = params.response_requested;
  if (params.allow_new_time_proposals !== undefined) body.allowNewTimeProposals = params.allow_new_time_proposals;
  if (params.recurrence !== undefined) body.recurrence = buildRecurrence(params.recurrence, ctx.timeZone);
  return body;
};
