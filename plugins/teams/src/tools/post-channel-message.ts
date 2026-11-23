import { defineTool } from '@opentabs-dev/plugin-sdk';
import { z } from 'zod';
import { textToHtml } from '../html.js';
import { graphApi } from '../teams-api.js';
import { formatField } from './message-input.js';

export const postChannelMessage = defineTool({
  name: 'post_channel_message',
  displayName: 'Post Channel Message',
  description:
    'Post a new top-level message to a team channel (via Microsoft Graph). To answer an existing message use reply_to_channel_message. Text is plain by default (escaped, newlines become line breaks); set format to "html" for raw HTML.',
  summary: 'Post a message to a channel',
  icon: 'send',
  group: 'Teams',
  input: z.object({
    team_id: z.string().min(1).describe('Team ID from list_teams'),
    channel_id: z.string().min(1).describe('Channel ID from list_channels'),
    text: z.string().min(1).describe('Message text'),
    subject: z.string().optional().describe('Optional subject line shown above the message'),
    format: formatField,
  }),
  output: z.object({
    id: z.string().describe('ID of the created message'),
    web_url: z.string().describe('Link to the message in Teams'),
  }),
  handle: async params => {
    const data = await graphApi<{ id?: string; webUrl?: string | null }>(
      `/teams/${encodeURIComponent(params.team_id)}/channels/${encodeURIComponent(params.channel_id)}/messages`,
      {
        method: 'POST',
        scope: 'ChannelMessage.Send',
        body: {
          ...(params.subject ? { subject: params.subject } : {}),
          body: { contentType: 'html', content: params.format === 'html' ? params.text : textToHtml(params.text) },
        },
      },
    );
    return { id: data.id ?? '', web_url: data.webUrl ?? '' };
  },
});
