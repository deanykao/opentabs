import { defineTool } from '@opentabs-dev/plugin-sdk';
import { z } from 'zod';
import { resolveTimeZone } from '../calendar-settings.js';
import { apiPaged } from '../outlook-api.js';
import {
  EVENT_SUMMARY_FIELDS,
  type RawEvent,
  eventSummarySchema,
  mapEventSummary,
  offsetDateTimeSchema,
} from './calendar-schemas.js';

export const listEventInstances = defineTool({
  name: 'list_event_instances',
  displayName: 'List Event Instances',
  description:
    'List the occurrences (and exceptions) of a recurring series within a date range. Pass the series master ID (series_master_id from get_event). Use an occurrence ID with update_event or delete_event to change a single instance.',
  summary: 'List occurrences of a series',
  icon: 'calendar-range',
  group: 'Calendar',
  input: z.object({
    event_id: z.string().describe('Series master event ID'),
    start: offsetDateTimeSchema.describe(
      'Range start as ISO 8601; include a UTC offset (e.g. "2026-06-01T00:00:00-05:00"), otherwise UTC',
    ),
    end: offsetDateTimeSchema.describe('Range end as ISO 8601; include a UTC offset, otherwise UTC'),
    limit: z.number().int().min(1).max(500).optional().describe('Max occurrences to return (default 100)'),
    time_zone: z
      .string()
      .optional()
      .describe('Time zone for returned start/end times. Defaults to your mailbox time zone.'),
  }),
  output: z.object({
    time_zone: z.string().describe('Time zone of the returned event times'),
    events: z.array(eventSummarySchema).describe('Occurrences ordered by start time'),
    has_more: z.boolean().describe('True when more occurrences exist beyond limit'),
  }),
  handle: async params => {
    const tz = await resolveTimeZone(params.time_zone);
    const limit = params.limit ?? 100;
    const { items, hasMore } = await apiPaged<RawEvent>(
      `/me/events/${encodeURIComponent(params.event_id)}/instances`,
      {
        query: {
          startDateTime: params.start,
          endDateTime: params.end,
          $select: EVENT_SUMMARY_FIELDS,
          $top: Math.min(limit, 100),
        },
        headers: { Prefer: `outlook.timezone="${tz}"` },
      },
      'calendar',
      limit,
    );
    return { time_zone: tz, events: items.map(mapEventSummary), has_more: hasMore };
  },
});
