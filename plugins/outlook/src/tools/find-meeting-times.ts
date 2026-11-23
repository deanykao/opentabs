import { defineTool } from '@opentabs-dev/plugin-sdk';
import { z } from 'zod';
import { resolveTimeZone } from '../calendar-settings.js';
import { api } from '../outlook-api.js';
import {
  type RawDateTimeZone,
  attendeeInputSchema,
  buildAttendees,
  dateTimeZoneSchema,
  localDateTimeSchema,
  mapDateTimeZone,
  toCamelEnum,
} from './calendar-schemas.js';
import { type RawEmailAddress, emailAddressSchema, mapEmailAddress } from './schemas.js';

interface RawSuggestion {
  confidence?: number;
  organizerAvailability?: string;
  suggestionReason?: string;
  meetingTimeSlot?: { start?: RawDateTimeZone; end?: RawDateTimeZone };
  attendeeAvailability?: { attendee?: RawEmailAddress; availability?: string }[];
  locations?: { displayName?: string; locationEmailAddress?: string }[];
}

interface RawSuggestionsResult {
  emptySuggestionsReason?: string;
  meetingTimeSuggestions?: RawSuggestion[];
}

const suggestionSchema = z.object({
  start: dateTimeZoneSchema.describe('Suggested start'),
  end: dateTimeZoneSchema.describe('Suggested end'),
  confidence: z.number().describe('Likelihood (0-100) that all attendees can attend'),
  organizer_availability: z.string().describe('Your availability in the slot (free, tentative, busy, ...)'),
  reason: z.string().describe('Why the slot was suggested, when returned'),
  attendees: z
    .array(z.object({ address: z.string(), name: z.string(), availability: z.string() }))
    .describe('Each attendee’s availability in the slot'),
  locations: z.array(emailAddressSchema).describe('Suggested rooms/locations (name and room mailbox)'),
});

export const findMeetingTimes = defineTool({
  name: 'find_meeting_times',
  displayName: 'Find Meeting Times',
  description:
    'Suggest meeting times when the given attendees (and optionally rooms) are available, like the Outlook Scheduling Assistant. Searches within the window(s) you give, in work hours by default. Returns ranked slots with each attendee’s availability; if none are found, empty_reason explains why. Does not create anything — pass a chosen slot to create_event.',
  summary: 'Suggest meeting times',
  icon: 'calendar-search',
  group: 'Calendar',
  input: z.object({
    attendees: z
      .array(attendeeInputSchema)
      .describe('People to meet with (you are included as organizer automatically). Use type resource for rooms.'),
    duration_minutes: z.number().int().min(5).max(1440).describe('Meeting length in minutes'),
    start: localDateTimeSchema.describe('Search window start (local ISO 8601 without offset)'),
    end: localDateTimeSchema.describe('Search window end (local ISO 8601 without offset)'),
    time_zone: z
      .string()
      .optional()
      .describe('Time zone for the window and returned slots. Defaults to your mailbox time zone.'),
    activity_domain: z
      .enum(['work', 'personal', 'unrestricted'])
      .optional()
      .describe('Hours to consider: work (attendees’ working hours, default), personal, or unrestricted'),
    max_candidates: z.number().int().min(1).max(50).optional().describe('Max suggestions to return (default 10)'),
    minimum_attendee_percentage: z
      .number()
      .min(0)
      .max(100)
      .optional()
      .describe('Minimum confidence for a slot to be returned (default 50)'),
    is_organizer_optional: z.boolean().optional().describe('Allow slots where you are busy (default false)'),
    rooms: z
      .array(z.string())
      .optional()
      .describe('Room mailbox addresses to consider (from list_rooms); slots require one of them free'),
    suggest_rooms: z.boolean().optional().describe('Let Outlook suggest rooms when no rooms are given (default false)'),
  }),
  output: z.object({
    time_zone: z.string().describe('Time zone of the returned slots'),
    suggestions: z.array(suggestionSchema).describe('Suggested slots, best first'),
    empty_reason: z.string().describe('Why no suggestions were returned (empty when there are suggestions)'),
  }),
  handle: async params => {
    const tz = await resolveTimeZone(params.time_zone);
    const rooms = params.rooms ?? [];
    const body: Record<string, unknown> = {
      attendees: buildAttendees(params.attendees),
      timeConstraint: {
        activityDomain: params.activity_domain ?? 'work',
        timeSlots: [{ start: { dateTime: params.start, timeZone: tz }, end: { dateTime: params.end, timeZone: tz } }],
      },
      meetingDuration: `PT${params.duration_minutes}M`,
      maxCandidates: params.max_candidates ?? 10,
      isOrganizerOptional: params.is_organizer_optional ?? false,
      returnSuggestionReasons: true,
      minimumAttendeePercentage: params.minimum_attendee_percentage ?? 50,
    };
    if (rooms.length > 0 || params.suggest_rooms) {
      body.locationConstraint = {
        isRequired: rooms.length > 0,
        suggestLocation: rooms.length === 0,
        locations: rooms.map(address => ({
          resolveAvailability: true,
          displayName: address,
          locationEmailAddress: address,
        })),
      };
    }

    const data = await api<RawSuggestionsResult>(
      '/me/findMeetingTimes',
      { method: 'POST', body, headers: { Prefer: `outlook.timezone="${tz}"` } },
      'calendar',
    );

    return {
      time_zone: tz,
      suggestions: (data.meetingTimeSuggestions ?? []).map(s => ({
        start: mapDateTimeZone(s.meetingTimeSlot?.start),
        end: mapDateTimeZone(s.meetingTimeSlot?.end),
        confidence: s.confidence ?? 0,
        organizer_availability: toCamelEnum(s.organizerAvailability ?? 'unknown'),
        reason: s.suggestionReason ?? '',
        attendees: (s.attendeeAvailability ?? []).map(a => ({
          ...mapEmailAddress(a.attendee),
          availability: toCamelEnum(a.availability ?? 'unknown'),
        })),
        locations: (s.locations ?? []).map(l => ({ name: l.displayName ?? '', address: l.locationEmailAddress ?? '' })),
      })),
      empty_reason: data.emptySuggestionsReason ?? '',
    };
  },
});
