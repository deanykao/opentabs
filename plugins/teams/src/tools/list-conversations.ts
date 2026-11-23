import { defineTool } from '@opentabs-dev/plugin-sdk';
import { z } from 'zod';
import { cursorFromBackwardLink, endpointFromCursor } from '../pagination.js';
import { chatApi } from '../teams-api.js';
import { conversationSchema, mapConversation } from './schemas.js';

export const listConversations = defineTool({
  name: 'list_conversations',
  displayName: 'List Conversations',
  description:
    'List chat conversations (1:1 chats, group chats, meeting chats, and channel threads). Use the conversation ID from results with read_messages. To get older conversations, pass the next_cursor from the previous response as cursor; next_cursor is absent when there are no more.',
  summary: 'List chats (paginated)',
  icon: 'message-square',
  group: 'Chats',
  input: z.object({
    page_size: z
      .number()
      .int()
      .min(1)
      .max(50)
      .optional()
      .describe('Number of conversations to return (default 20, max 50)'),
    cursor: z
      .string()
      .optional()
      .describe('Opaque next_cursor from a previous list_conversations response, to fetch the next page'),
  }),
  output: z.object({
    conversations: z.array(conversationSchema).describe('List of conversations'),
    next_cursor: z.string().optional().describe('Pass as cursor to get the next page; absent when no more pages'),
  }),
  handle: async params => {
    const pageSize = params.page_size ?? 20;
    const data = params.cursor
      ? await chatApi<ConversationsResponse>(endpointFromCursor(params.cursor, '/conversations?', pageSize))
      : await chatApi<ConversationsResponse>('/v1/users/ME/conversations', {
          query: { view: 'superchat', pageSize, startTime: 0, targetType: 'Thread|Passport' },
        });

    const conversations = (data.conversations ?? []).map(c => mapConversation(c));
    const nextCursor = conversations.length > 0 ? cursorFromBackwardLink(data._metadata?.backwardLink) : undefined;
    return { conversations, ...(nextCursor ? { next_cursor: nextCursor } : {}) };
  },
});

interface ConversationsResponse {
  conversations?: Array<Record<string, unknown>>;
  _metadata?: { backwardLink?: string };
}
