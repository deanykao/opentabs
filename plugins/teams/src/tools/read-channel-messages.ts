import { defineTool } from '@opentabs-dev/plugin-sdk';
import { z } from 'zod';
import { graphApi, graphApiFromLink } from '../teams-api.js';
import { channelMessageSchema, type GraphChannelMessage, type GraphPage, mapChannelMessage } from './graph-schemas.js';

export const readChannelMessages = defineTool({
  name: 'read_channel_messages',
  displayName: 'Read Channel Messages',
  description:
    'Read top-level messages of a team channel, newest first (via Microsoft Graph). Replies are not included — use read_channel_replies with a message ID. Pass next_cursor from the previous response as cursor for older messages.',
  summary: 'Read messages in a channel',
  icon: 'message-square',
  group: 'Teams',
  input: z.object({
    team_id: z.string().min(1).describe('Team ID from list_teams'),
    channel_id: z.string().min(1).describe('Channel ID from list_channels'),
    page_size: z.number().int().min(1).max(50).optional().describe('Messages per page (default 20, max 50)'),
    cursor: z.string().optional().describe('Opaque next_cursor from a previous read_channel_messages response'),
  }),
  output: z.object({
    messages: z.array(channelMessageSchema).describe('Channel messages, newest first'),
    next_cursor: z.string().optional().describe('Pass as cursor to get the next page; absent when no more'),
  }),
  handle: async params => {
    const options = { scope: 'ChannelMessage.Read.All' };
    const data = params.cursor
      ? await graphApiFromLink<GraphPage<GraphChannelMessage>>(params.cursor, options)
      : await graphApi<GraphPage<GraphChannelMessage>>(
          `/teams/${encodeURIComponent(params.team_id)}/channels/${encodeURIComponent(params.channel_id)}/messages`,
          { ...options, query: { $top: params.page_size ?? 20 } },
        );
    const next = data['@odata.nextLink'];
    return {
      messages: (data.value ?? []).map(mapChannelMessage),
      ...(next ? { next_cursor: next } : {}),
    };
  },
});
