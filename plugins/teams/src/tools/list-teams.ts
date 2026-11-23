import { defineTool } from '@opentabs-dev/plugin-sdk';
import { z } from 'zod';
import { graphApi } from '../teams-api.js';

interface GraphTeam {
  id?: string;
  displayName?: string;
  description?: string | null;
  isArchived?: boolean;
  webUrl?: string | null;
}

export const listTeams = defineTool({
  name: 'list_teams',
  displayName: 'List Teams',
  description:
    'List the Microsoft Teams teams the current user has joined (via Microsoft Graph). Use the team ID with list_channels.',
  summary: 'List joined teams',
  icon: 'users',
  group: 'Teams',
  input: z.object({}),
  output: z.object({
    teams: z.array(
      z.object({
        id: z.string().describe('Team ID (group GUID)'),
        name: z.string().describe('Team display name'),
        description: z.string().describe('Team description'),
        is_archived: z.boolean().describe('Whether the team is archived'),
        web_url: z.string().describe('Link to the team in Teams'),
      }),
    ),
  }),
  handle: async () => {
    const data = await graphApi<{ value?: GraphTeam[] }>('/me/joinedTeams', { scope: 'Team.ReadBasic.All' });
    return {
      teams: (data.value ?? []).map(t => ({
        id: t.id ?? '',
        name: t.displayName ?? '',
        description: t.description ?? '',
        is_archived: t.isArchived ?? false,
        web_url: t.webUrl ?? '',
      })),
    };
  },
});
