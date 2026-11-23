import { defineTool } from '@opentabs-dev/plugin-sdk';
import { z } from 'zod';
import { cursorFromBackwardLink, endpointFromCursor } from '../pagination.js';
import { chatApi } from '../teams-api.js';
import { mapMessage, messageSchema } from './schemas.js';

export const readMessages = defineTool({
  name: 'read_messages',
  displayName: 'Read Messages',
  description:
    'Read messages from a chat conversation, newest first. Includes every message type with its message_type (system events such as "ThreadActivity/AddMember" and cards such as "RichText/Media_Card" included) unless skip_system is true. To get older messages, pass next_cursor from the previous response as cursor; next_cursor is absent when there are no more.',
  summary: 'Read messages from a chat (paginated)',
  icon: 'message-square',
  group: 'Messages',
  input: z.object({
    conversation_id: z.string().min(1).describe('Conversation/thread ID to read messages from'),
    page_size: z
      .number()
      .int()
      .min(1)
      .max(200)
      .optional()
      .describe('Number of messages to return (default 20, max 200)'),
    cursor: z
      .string()
      .optional()
      .describe('Opaque next_cursor from a previous read_messages response for the same conversation'),
    skip_system: z
      .boolean()
      .optional()
      .describe(
        'Omit non-conversational messages (message_type starting with "ThreadActivity/" or "Event/"); default false',
      ),
  }),
  output: z.object({
    messages: z.array(messageSchema).describe('List of messages (newest first)'),
    skipped_count: z.number().describe('Messages omitted because skip_system was true (0 otherwise)'),
    next_cursor: z.string().optional().describe('Pass as cursor to get older messages; absent when no more'),
  }),
  handle: async params => {
    const pageSize = params.page_size ?? 20;
    const base = `/v1/users/ME/conversations/${encodeURIComponent(params.conversation_id)}/messages`;
    const data = params.cursor
      ? await chatApi<MessagesResponse>(
          endpointFromCursor(params.cursor, `/conversations/${params.conversation_id}/messages`, pageSize),
        )
      : await chatApi<MessagesResponse>(base, { query: { pageSize } });

    const raw = data.messages ?? [];
    const isSystem = (m: Record<string, unknown>): boolean => {
      const type = typeof m.messagetype === 'string' ? m.messagetype : '';
      return type.startsWith('ThreadActivity/') || type.startsWith('Event/');
    };
    const kept = params.skip_system ? raw.filter(m => !isSystem(m)) : raw;
    const nextCursor = raw.length > 0 ? cursorFromBackwardLink(data._metadata?.backwardLink) : undefined;
    return {
      messages: kept.map(m => mapMessage(m)),
      skipped_count: raw.length - kept.length,
      ...(nextCursor ? { next_cursor: nextCursor } : {}),
    };
  },
});

interface MessagesResponse {
  messages?: Array<Record<string, unknown>>;
  _metadata?: { backwardLink?: string };
}
