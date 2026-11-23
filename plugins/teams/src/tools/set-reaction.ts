import { defineTool } from '@opentabs-dev/plugin-sdk';
import { z } from 'zod';
import { chatApi } from '../teams-api.js';

export const setReaction = defineTool({
  name: 'set_reaction',
  displayName: 'Add Reaction',
  description:
    'Add an emoji reaction to a chat message as the current user. Use a Teams reaction key such as "like", "heart", "laugh", "surprised", "sad", "angry" (the keys read_messages reports in reactions). Message IDs come from read_messages.',
  summary: 'React to a chat message',
  icon: 'smile-plus',
  group: 'Messages',
  input: z.object({
    conversation_id: z.string().min(1).describe('Conversation/thread ID'),
    message_id: z.string().min(1).describe('Message ID to react to'),
    reaction: z.string().min(1).describe('Reaction key (e.g., "like", "heart", "laugh")'),
  }),
  output: z.object({
    success: z.boolean().describe('Whether the reaction was added'),
  }),
  handle: async params => {
    await chatApi<Record<string, unknown>>(
      `/v1/users/ME/conversations/${encodeURIComponent(params.conversation_id)}/messages/${encodeURIComponent(params.message_id)}/properties`,
      {
        method: 'PUT',
        query: { name: 'emotions' },
        body: { emotions: { key: params.reaction, value: Date.now() } },
      },
    );
    return { success: true };
  },
});
