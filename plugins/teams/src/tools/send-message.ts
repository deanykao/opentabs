import { defineTool } from '@opentabs-dev/plugin-sdk';
import { z } from 'zod';
import { buildMessageBody, randomClientMessageId } from '../html.js';
import { chatApiRaw } from '../teams-api.js';
import { formatField, mentionsField } from './message-input.js';

export const sendMessage = defineTool({
  name: 'send_message',
  displayName: 'Send Message',
  description:
    'Send a message to a Teams chat conversation (use the conversation ID from list_conversations or create_chat). By default text is plain: &, <, > are escaped and newlines become line breaks. Set format to "html" to send raw HTML. To @mention people, pass mentions and write "@Display Name" in the text (or {{0}}, {{1}} placeholders in html format).',
  summary: 'Send a message to a chat',
  icon: 'send',
  group: 'Messages',
  input: z.object({
    conversation_id: z.string().min(1).describe('Conversation/thread ID to send the message to'),
    text: z.string().min(1).describe('Message text to send'),
    format: formatField,
    mentions: mentionsField,
  }),
  output: z.object({
    message_id: z
      .string()
      .describe('Server-assigned message ID; the same ID read_messages returns and edit_message/delete_message accept'),
    client_message_id: z.string().describe('Client-assigned message ID'),
  }),
  handle: async params => {
    const { content, mentionsJson } = buildMessageBody(params.text, params.format ?? 'text', params.mentions);
    const clientMsgId = randomClientMessageId();
    const { data, headers } = await chatApiRaw<{ OriginalArrivalTime?: number }>(
      `/v1/users/ME/conversations/${encodeURIComponent(params.conversation_id)}/messages`,
      {
        method: 'POST',
        body: {
          content,
          messagetype: 'RichText/Html',
          contenttype: 'text',
          clientmessageid: clientMsgId,
          ...(mentionsJson ? { properties: { mentions: mentionsJson } } : {}),
        },
      },
    );

    // The message ID equals the arrival time in epoch ms (the same value read_messages reports as `id`).
    const location = headers.get('Location') ?? '';
    const fromLocation = /\/messages\/([^/?]+)$/.exec(location)?.[1];
    return {
      message_id: fromLocation ?? String(data.OriginalArrivalTime ?? ''),
      client_message_id: clientMsgId,
    };
  },
});
