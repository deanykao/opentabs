import { ToolError, defineTool } from '@opentabs-dev/plugin-sdk';
import { z } from 'zod';
import { randomClientMessageId } from '../html.js';
import { chatApi } from '../teams-api.js';

export const markChatRead = defineTool({
  name: 'mark_chat_read',
  displayName: 'Mark Chat Read',
  description:
    'Mark a chat as read up to a message by moving its consumption horizon. Defaults to the newest message in the chat. This clears the unread indicator for the current user.',
  summary: 'Mark a chat as read',
  icon: 'check-check',
  group: 'Chats',
  input: z.object({
    conversation_id: z.string().min(1).describe('Conversation/thread ID'),
    message_id: z
      .string()
      .regex(/^\d+$/, 'message_id must be the numeric ID from read_messages')
      .optional()
      .describe('Mark read up to this message ID (default: newest message)'),
  }),
  output: z.object({
    success: z.boolean().describe('Whether the read marker was updated'),
    read_up_to: z.string().describe('Message ID the chat is now marked read through'),
  }),
  handle: async params => {
    const base = `/v1/users/ME/conversations/${encodeURIComponent(params.conversation_id)}`;
    let messageId = params.message_id;
    if (!messageId) {
      const data = await chatApi<{ messages?: Array<{ id?: string }> }>(`${base}/messages`, {
        query: { pageSize: 20 },
      });
      // Message order in a page is not strictly chronological; IDs are arrival times, so take the max.
      const ids = (data.messages ?? []).map(m => m.id ?? '').filter(id => /^\d+$/.test(id));
      messageId = ids.reduce<string | undefined>(
        (best, id) => (best === undefined || BigInt(id) > BigInt(best) ? id : best),
        undefined,
      );
      if (!messageId) throw ToolError.validation('Conversation has no messages to mark as read');
    }

    // Consumption horizon format observed on conversations: "<messageId>;<nowMs>;<clientMessageId>".
    await chatApi<Record<string, unknown>>(`${base}/properties`, {
      method: 'PUT',
      query: { name: 'consumptionhorizon' },
      body: { consumptionhorizon: `${messageId};${String(Date.now())};${randomClientMessageId()}` },
    });
    return { success: true, read_up_to: messageId };
  },
});
