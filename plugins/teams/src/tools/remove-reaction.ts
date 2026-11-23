import { defineTool } from '@opentabs-dev/plugin-sdk';
import { z } from 'zod';
import { chatApi } from '../teams-api.js';

export const removeReaction = defineTool({
  name: 'remove_reaction',
  displayName: 'Remove Reaction',
  description:
    'Remove the current user\'s emoji reaction from a chat message. Use the same reaction key that was set (e.g., "like", "heart"). Message IDs come from read_messages.',
  summary: 'Remove a reaction from a chat message',
  icon: 'smile',
  group: 'Messages',
  input: z.object({
    conversation_id: z.string().min(1).describe('Conversation/thread ID'),
    message_id: z.string().min(1).describe('Message ID to remove the reaction from'),
    reaction: z.string().min(1).describe('Reaction key to remove (e.g., "like", "heart")'),
  }),
  output: z.object({
    success: z.boolean().describe('Whether the reaction was removed'),
  }),
  handle: async params => {
    await chatApi<Record<string, unknown>>(
      `/v1/users/ME/conversations/${encodeURIComponent(params.conversation_id)}/messages/${encodeURIComponent(params.message_id)}/properties`,
      {
        method: 'DELETE',
        query: { name: 'emotions' },
        body: { emotions: { key: params.reaction } },
      },
    );
    return { success: true };
  },
});
