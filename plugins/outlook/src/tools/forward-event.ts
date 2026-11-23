import { defineTool } from '@opentabs-dev/plugin-sdk';
import { z } from 'zod';
import { api } from '../outlook-api.js';

export const forwardEvent = defineTool({
  name: 'forward_event',
  displayName: 'Forward Event',
  description:
    'Forward a meeting invitation to other people so they can join. As organizer, recipients are added to the meeting; as an attendee, the organizer is notified that you forwarded it.',
  summary: 'Forward a meeting invitation',
  icon: 'forward',
  group: 'Calendar',
  input: z.object({
    event_id: z.string().describe('The event ID to forward'),
    to: z
      .array(z.object({ address: z.string().describe('Email address'), name: z.string().optional() }))
      .min(1)
      .describe('Recipients'),
    comment: z.string().optional().describe('Optional note to include'),
  }),
  output: z.object({
    success: z.boolean().describe('Whether the invitation was forwarded'),
  }),
  handle: async params => {
    await api(
      `/me/events/${encodeURIComponent(params.event_id)}/forward`,
      {
        method: 'POST',
        body: {
          toRecipients: params.to.map(r => ({ emailAddress: { address: r.address, name: r.name } })),
          comment: params.comment,
        },
      },
      'calendar-write',
    );
    return { success: true };
  },
});
