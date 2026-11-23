import { defineTool } from '@opentabs-dev/plugin-sdk';
import { z } from 'zod';
import { textToHtml } from '../html.js';
import { graphApi } from '../teams-api.js';
import { formatField } from './message-input.js';

export const replyToChannelMessage = defineTool({
  name: 'reply_to_channel_message',
  displayName: 'Reply to Channel Message',
  description:
    'Reply to a message in a team channel thread (via Microsoft Graph). Use the message ID from read_channel_messages. Text is plain by default (escaped, newlines become line breaks); set format to "html" for raw HTML.',
  summary: 'Reply to a channel message',
  icon: 'reply',
  group: 'Teams',
  input: z.object({
    team_id: z.string().min(1).describe('Team ID from list_teams'),
    channel_id: z.string().min(1).describe('Channel ID from list_channels'),
    message_id: z.string().min(1).describe('ID of the top-level message to reply to'),
    text: z.string().min(1).describe('Reply text'),
    format: formatField,
  }),
  output: z.object({
    id: z.string().describe('ID of the created reply'),
    web_url: z.string().describe('Link to the reply in Teams'),
  }),
  handle: async params => {
    const data = await graphApi<{ id?: string; webUrl?: string | null }>(
      `/teams/${encodeURIComponent(params.team_id)}/channels/${encodeURIComponent(params.channel_id)}/messages/${encodeURIComponent(params.message_id)}/replies`,
      {
        method: 'POST',
        scope: 'ChannelMessage.Send',
        body: {
          body: { contentType: 'html', content: params.format === 'html' ? params.text : textToHtml(params.text) },
        },
      },
    );
    return { id: data.id ?? '', web_url: data.webUrl ?? '' };
  },
});
