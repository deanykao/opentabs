import { defineTool } from '@opentabs-dev/plugin-sdk';
import { z } from 'zod';
import { graphApi, graphApiFromLink } from '../teams-api.js';
import { channelMessageSchema, type GraphChannelMessage, type GraphPage, mapChannelMessage } from './graph-schemas.js';

export const readChannelReplies = defineTool({
  name: 'read_channel_replies',
  displayName: 'Read Channel Replies',
  description:
    'Read the replies to a channel message (via Microsoft Graph). Use the message ID from read_channel_messages. Pass next_cursor from the previous response as cursor for more.',
  summary: 'Read replies to a channel message',
  icon: 'messages-square',
  group: 'Teams',
  input: z.object({
    team_id: z.string().min(1).describe('Team ID from list_teams'),
    channel_id: z.string().min(1).describe('Channel ID from list_channels'),
    message_id: z.string().min(1).describe('ID of the parent message from read_channel_messages'),
    page_size: z.number().int().min(1).max(50).optional().describe('Replies per page (default 20, max 50)'),
    cursor: z.string().optional().describe('Opaque next_cursor from a previous read_channel_replies response'),
  }),
  output: z.object({
    replies: z.array(channelMessageSchema).describe('Replies to the message'),
    next_cursor: z.string().optional().describe('Pass as cursor to get the next page; absent when no more'),
  }),
  handle: async params => {
    const options = { scope: 'ChannelMessage.Read.All' };
    const data = params.cursor
      ? await graphApiFromLink<GraphPage<GraphChannelMessage>>(params.cursor, options)
      : await graphApi<GraphPage<GraphChannelMessage>>(
          `/teams/${encodeURIComponent(params.team_id)}/channels/${encodeURIComponent(params.channel_id)}/messages/${encodeURIComponent(params.message_id)}/replies`,
          { ...options, query: { $top: params.page_size ?? 20 } },
        );
    const next = data['@odata.nextLink'];
    return {
      replies: (data.value ?? []).map(mapChannelMessage),
      ...(next ? { next_cursor: next } : {}),
    };
  },
});
