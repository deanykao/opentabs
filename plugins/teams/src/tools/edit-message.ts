import { defineTool } from '@opentabs-dev/plugin-sdk';
import { z } from 'zod';
import { buildMessageBody } from '../html.js';
import { chatApi } from '../teams-api.js';
import { formatField, mentionsField } from './message-input.js';

export const editMessage = defineTool({
  name: 'edit_message',
  displayName: 'Edit Message',
  description:
    'Edit a message in a Teams chat conversation. Only messages sent by the current user can be edited. Message IDs come from read_messages or send_message. Text is plain by default (escaped); set format to "html" for raw HTML.',
  summary: 'Edit a chat message',
  icon: 'pencil',
  group: 'Messages',
  input: z.object({
    conversation_id: z.string().min(1).describe('Conversation/thread ID'),
    message_id: z.string().min(1).describe('Message ID to edit'),
    text: z.string().min(1).describe('New message content'),
    format: formatField,
    mentions: mentionsField,
  }),
  output: z.object({
    success: z.boolean().describe('Whether the edit was successful'),
  }),
  handle: async params => {
    const { content, mentionsJson } = buildMessageBody(params.text, params.format ?? 'text', params.mentions);
    await chatApi<Record<string, unknown>>(
      `/v1/users/ME/conversations/${encodeURIComponent(params.conversation_id)}/messages/${encodeURIComponent(params.message_id)}`,
      {
        method: 'PUT',
        body: {
          content,
          messagetype: 'RichText/Html',
          contenttype: 'text',
          ...(mentionsJson ? { properties: { mentions: mentionsJson } } : {}),
        },
      },
    );
    return { success: true };
  },
});
