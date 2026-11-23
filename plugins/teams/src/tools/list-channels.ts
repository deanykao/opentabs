import { defineTool } from '@opentabs-dev/plugin-sdk';
import { z } from 'zod';
import { graphApi } from '../teams-api.js';

interface GraphChannel {
  id?: string;
  displayName?: string;
  description?: string | null;
  membershipType?: string;
  webUrl?: string | null;
  isArchived?: boolean;
}

export const listChannels = defineTool({
  name: 'list_channels',
  displayName: 'List Channels',
  description:
    'List the channels of a team (via Microsoft Graph). Use the team ID from list_teams. Channel IDs work with read_channel_messages and post_channel_message.',
  summary: 'List channels in a team',
  icon: 'hash',
  group: 'Teams',
  input: z.object({
    team_id: z.string().min(1).describe('Team ID from list_teams'),
  }),
  output: z.object({
    channels: z.array(
      z.object({
        id: z.string().describe('Channel ID (e.g., "19:...@thread.tacv2")'),
        name: z.string().describe('Channel display name'),
        description: z.string().describe('Channel description'),
        membership_type: z.string().describe('"standard", "private" or "shared"'),
        is_archived: z.boolean().describe('Whether the channel is archived'),
        web_url: z.string().describe('Link to the channel in Teams'),
      }),
    ),
  }),
  handle: async params => {
    const data = await graphApi<{ value?: GraphChannel[] }>(`/teams/${encodeURIComponent(params.team_id)}/channels`, {
      scope: 'Channel.ReadBasic.All',
    });
    return {
      channels: (data.value ?? []).map(c => ({
        id: c.id ?? '',
        name: c.displayName ?? '',
        description: c.description ?? '',
        membership_type: c.membershipType ?? '',
        is_archived: c.isArchived ?? false,
        web_url: c.webUrl ?? '',
      })),
    };
  },
});
