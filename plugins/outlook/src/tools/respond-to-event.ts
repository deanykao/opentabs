import { ToolError, defineTool } from '@opentabs-dev/plugin-sdk';
import { z } from 'zod';
import { resolveTimeZone } from '../calendar-settings.js';
import { api } from '../outlook-api.js';
import { buildDateTime, localDateTimeSchema } from './calendar-schemas.js';

const RESPONSE_ACTION: Record<string, string> = {
  accept: 'accept',
  decline: 'decline',
  tentative: 'tentativelyAccept',
};

export const respondToEvent = defineTool({
  name: 'respond_to_event',
  displayName: 'Respond to Event',
  description:
    'Respond to a meeting invitation by accepting, declining, or tentatively accepting. Optionally include a comment, choose whether to send a response back to the organizer, and (for decline or tentative) propose a new time.',
  summary: 'Accept, decline, or tentatively accept',
  icon: 'calendar-check',
  group: 'Calendar',
  input: z.object({
    event_id: z.string().describe('The event ID of the meeting invitation'),
    response: z.enum(['accept', 'decline', 'tentative']).describe('Your response to the invitation'),
    comment: z.string().optional().describe('Optional comment to include with the response'),
    send_response: z.boolean().optional().describe('Whether to send the response to the organizer (default: true)'),
    proposed_start: localDateTimeSchema
      .optional()
      .describe('Propose a new start (decline/tentative only; requires proposed_end and send_response)'),
    proposed_end: localDateTimeSchema.optional().describe('Proposed new end, paired with proposed_start'),
    time_zone: z.string().optional().describe('Time zone of the proposed times. Defaults to your mailbox time zone.'),
  }),
  output: z.object({
    success: z.boolean().describe('Whether the response was recorded'),
  }),
  handle: async params => {
    const proposing = params.proposed_start !== undefined || params.proposed_end !== undefined;
    if (proposing) {
      if (params.response === 'accept') {
        throw ToolError.validation('A new time can only be proposed when declining or tentatively accepting.');
      }
      if (!params.proposed_start || !params.proposed_end) {
        throw ToolError.validation('Provide both proposed_start and proposed_end.');
      }
      if (params.send_response === false) {
        throw ToolError.validation('Proposing a new time requires send_response to be true.');
      }
    }

    const body: Record<string, unknown> = {
      comment: params.comment,
      sendResponse: params.send_response ?? true,
    };
    if (params.proposed_start && params.proposed_end) {
      const tz = await resolveTimeZone(params.time_zone);
      body.proposedNewTime = {
        start: buildDateTime(params.proposed_start, tz),
        end: buildDateTime(params.proposed_end, tz),
      };
    }

    await api(
      `/me/events/${encodeURIComponent(params.event_id)}/${RESPONSE_ACTION[params.response]}`,
      { method: 'POST', body },
      'calendar-write',
    );
    return { success: true };
  },
});
