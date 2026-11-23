import { defineTool } from '@opentabs-dev/plugin-sdk';
import { z } from 'zod';
import { getOnlineMeetingProvider, resolveTimeZone } from '../calendar-settings.js';
import { api } from '../outlook-api.js';
import {
  EVENT_DETAIL_FIELDS,
  type RawEvent,
  eventDetailSchema,
  localDateTimeSchema,
  mapEventDetail,
} from './calendar-schemas.js';
import { buildEventBody, eventFieldsInput } from './event-input.js';

export const createEvent = defineTool({
  name: 'create_event',
  displayName: 'Create Event',
  description:
    'Create a calendar event or meeting. Adding attendees sends them invitations. Set is_online_meeting to attach a Teams join link, recurrence to create a series, and location_email plus a resource attendee to book a room. Times are local wall-clock times in time_zone (default: your mailbox time zone).',
  summary: 'Create a calendar event',
  icon: 'calendar-plus',
  group: 'Calendar',
  input: z.object({
    subject: z.string().describe('Event subject/title'),
    start: localDateTimeSchema.describe('Start as local ISO 8601 without an offset (e.g. "2026-06-02T13:00:00")'),
    end: localDateTimeSchema.describe('End as local ISO 8601 without an offset (e.g. "2026-06-02T14:00:00")'),
    ...eventFieldsInput,
    calendar_id: z.string().optional().describe('Calendar to create the event in. Defaults to the primary calendar.'),
  }),
  output: z.object({
    event: eventDetailSchema.describe('The created event'),
  }),
  handle: async params => {
    const timeZone = await resolveTimeZone(params.time_zone);
    const onlineMeetingProvider = params.is_online_meeting
      ? await getOnlineMeetingProvider(params.calendar_id)
      : undefined;
    const body = buildEventBody(params, { timeZone, onlineMeetingProvider });

    const endpoint = params.calendar_id
      ? `/me/calendars/${encodeURIComponent(params.calendar_id)}/events`
      : '/me/events';
    const created = await api<RawEvent>(
      endpoint,
      {
        method: 'POST',
        body,
        query: { $select: EVENT_DETAIL_FIELDS },
        headers: { Prefer: `outlook.timezone="${timeZone}"` },
      },
      'calendar-write',
    );
    return { event: mapEventDetail(created) };
  },
});
