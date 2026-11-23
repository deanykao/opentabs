import { ToolError, defineTool } from '@opentabs-dev/plugin-sdk';
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

export const updateEvent = defineTool({
  name: 'update_event',
  displayName: 'Update Event',
  description:
    'Update properties of an existing calendar event. Only the fields you provide are changed. Updating a meeting you organize sends updated invitations to its attendees. To change start/end times, provide both start and end (local times in time_zone, default: your mailbox time zone). Pass an occurrence ID to change one instance, or the series master ID to change the whole series.',
  summary: 'Update a calendar event',
  icon: 'pencil',
  group: 'Calendar',
  input: z.object({
    event_id: z.string().describe('The event ID to update'),
    subject: z.string().optional().describe('New subject/title'),
    start: localDateTimeSchema.optional().describe('New start as local ISO 8601 without an offset'),
    end: localDateTimeSchema.optional().describe('New end as local ISO 8601 without an offset'),
    ...eventFieldsInput,
  }),
  output: z.object({
    event: eventDetailSchema.describe('The updated event'),
  }),
  handle: async params => {
    const { event_id, time_zone, ...fields } = params;
    if (Object.values(fields).every(v => v === undefined)) {
      throw ToolError.validation('Provide at least one field to update besides event_id.');
    }

    const timeZone = await resolveTimeZone(time_zone);
    const onlineMeetingProvider = params.is_online_meeting ? await getOnlineMeetingProvider(undefined) : undefined;
    const body = buildEventBody(fields, { timeZone, onlineMeetingProvider });

    const updated = await api<RawEvent>(
      `/me/events/${encodeURIComponent(event_id)}`,
      {
        method: 'PATCH',
        body,
        query: { $select: EVENT_DETAIL_FIELDS },
        headers: { Prefer: `outlook.timezone="${timeZone}"` },
      },
      'calendar-write',
    );
    return { event: mapEventDetail(updated) };
  },
});
