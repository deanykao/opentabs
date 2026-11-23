import { defineTool } from '@opentabs-dev/plugin-sdk';
import { z } from 'zod';
import { substrateApi } from '../teams-api.js';

interface SubstrateSource {
  ClientConversationId?: string;
  ClientThreadId?: string;
  ChannelName?: string;
  TeamName?: string;
  DateTimeSent?: string;
  Preview?: string;
  From?: { EmailAddress?: { Address?: string; Name?: string } };
}

interface SubstrateResponse {
  EntitySets?: Array<{
    ResultSets?: Array<{
      Total?: number;
      MoreResultsAvailable?: boolean;
      Results?: Array<{ Source?: SubstrateSource; HitHighlightedSummary?: string }>;
    }>;
  }>;
}

export const searchMessages = defineTool({
  name: 'search_messages',
  displayName: 'Search Messages',
  description:
    'Search Teams chat and channel messages the current user can see, using the same Substrate search Teams uses. Returns previews with conversation_id and message_id; use read_messages with the conversation_id for context. Supports Teams search syntax (e.g., "from:", quoted phrases).',
  summary: 'Search Teams messages',
  icon: 'search',
  group: 'Messages',
  input: z.object({
    query: z.string().min(1).describe('Search query'),
    limit: z.number().int().min(1).max(50).optional().describe('Results to return (default 10, max 50)'),
    offset: z.number().int().min(0).optional().describe('Number of results to skip, for paging (default 0)'),
  }),
  output: z.object({
    results: z.array(
      z.object({
        conversation_id: z.string().describe('Conversation/thread ID containing the message'),
        message_id: z.string().describe('Message ID within the conversation (empty if not reported)'),
        from_name: z.string().describe('Sender display name'),
        from_email: z.string().describe('Sender email address'),
        sent_at: z.string().describe('Time the message was sent (ISO 8601)'),
        team_name: z.string().describe('Team name for channel messages, else empty'),
        channel_name: z.string().describe('Channel name for channel messages, else empty'),
        preview: z.string().describe('Message preview text'),
      }),
    ),
    total: z.number().describe('Total number of matches'),
    next_offset: z.number().optional().describe('Pass as offset to get the next page; absent when no more'),
  }),
  handle: async params => {
    const size = params.limit ?? 10;
    const from = params.offset ?? 0;
    const data = await substrateApi<SubstrateResponse>('/search/api/v2/query', {
      method: 'POST',
      body: {
        entityRequests: [
          {
            entityType: 'Message',
            contentSources: ['Teams'],
            query: { queryString: params.query },
            from,
            size,
          },
        ],
        cvid: crypto.randomUUID(),
        logicalId: crypto.randomUUID(),
        scenario: { name: 'opentabs_teams_search' },
      },
    });

    const resultSet = data.EntitySets?.[0]?.ResultSets?.[0];
    const results = (resultSet?.Results ?? []).map(r => {
      const src = r.Source ?? {};
      const messageId = /;messageid=(\d+)/.exec(src.ClientConversationId ?? '')?.[1] ?? '';
      return {
        conversation_id: src.ClientThreadId ?? '',
        message_id: messageId,
        from_name: src.From?.EmailAddress?.Name ?? '',
        from_email: src.From?.EmailAddress?.Address ?? '',
        sent_at: src.DateTimeSent ?? '',
        team_name: src.TeamName ?? '',
        channel_name: src.ChannelName ?? '',
        preview: src.Preview ?? r.HitHighlightedSummary ?? '',
      };
    });
    const hasMore = resultSet?.MoreResultsAvailable === true && results.length > 0;
    return {
      results,
      total: resultSet?.Total ?? results.length,
      ...(hasMore ? { next_offset: from + results.length } : {}),
    };
  },
});
