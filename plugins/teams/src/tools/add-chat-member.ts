import { defineTool } from '@opentabs-dev/plugin-sdk';
import { z } from 'zod';
import { threadApi } from '../teams-api.js';

export const addChatMember = defineTool({
  name: 'add_chat_member',
  displayName: 'Add Chat Member',
  description:
    'Add a user to a Teams group chat by their MRI (e.g., "8:orgid:<guid>"). Applies to group chats, not team channels.',
  summary: 'Add a user to a group chat',
  icon: 'user-plus',
  group: 'Members',
  input: z.object({
    conversation_id: z.string().min(1).describe('Conversation/thread ID of the group chat'),
    user: z.string().min(1).describe('User MRI to add (e.g., "8:orgid:<guid>")'),
    role: z.enum(['User', 'Admin']).optional().describe('Role to grant in the chat (default "User")'),
  }),
  output: z.object({
    success: z.boolean().describe('Whether the member was added'),
  }),
  handle: async params => {
    await threadApi<Record<string, unknown>>(params.conversation_id, `/members/${encodeURIComponent(params.user)}`, {
      method: 'PUT',
      body: { role: params.role ?? 'User' },
    });
    return { success: true };
  },
});
