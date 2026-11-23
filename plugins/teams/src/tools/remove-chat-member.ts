import { defineTool } from '@opentabs-dev/plugin-sdk';
import { z } from 'zod';
import { threadApi } from '../teams-api.js';

export const removeChatMember = defineTool({
  name: 'remove_chat_member',
  displayName: 'Remove Chat Member',
  description:
    'Remove a user from a Teams group chat by MRI. Requires admin privileges in the chat. Applies to group chats, not team channels.',
  summary: 'Remove a user from a group chat',
  icon: 'user-minus',
  group: 'Members',
  input: z.object({
    conversation_id: z.string().min(1).describe('Conversation/thread ID of the group chat'),
    user: z.string().min(1).describe('User MRI to remove (e.g., "8:orgid:<guid>")'),
  }),
  output: z.object({
    success: z.boolean().describe('Whether the member was removed'),
  }),
  handle: async params => {
    await threadApi<Record<string, unknown>>(params.conversation_id, `/members/${encodeURIComponent(params.user)}`, {
      method: 'DELETE',
    });
    return { success: true };
  },
});
