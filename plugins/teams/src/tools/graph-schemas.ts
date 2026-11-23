import { z } from 'zod';

// ---------------------------------------------------------------------------
// Microsoft Graph channel message schema (shared by channel read tools)
// ---------------------------------------------------------------------------

export const channelMessageSchema = z.object({
  id: z.string().describe('Message ID (use as message_id for read_channel_replies / reply_to_channel_message)'),
  reply_to_id: z.string().describe('ID of the parent message if this is a reply, else empty'),
  message_type: z.string().describe('Graph message type (e.g., "message", "systemEventMessage")'),
  created_at: z.string().describe('Creation time (ISO 8601)'),
  last_modified_at: z.string().describe('Last modification time (ISO 8601), empty if never modified'),
  deleted: z.boolean().describe('Whether the message was deleted'),
  subject: z.string().describe('Message subject, empty if none'),
  from_name: z.string().describe('Sender display name (empty for system messages)'),
  from_user_id: z.string().describe('Sender Entra ID object ID (empty for apps/system)'),
  content_type: z.string().describe('"html" or "text"'),
  content: z.string().describe('Message body'),
  mentions: z.array(z.object({ id: z.number(), name: z.string() })).describe('Mentions in the message'),
  attachments: z
    .array(z.object({ name: z.string(), content_type: z.string(), url: z.string() }))
    .describe('Attachments on the message'),
  reactions: z
    .array(z.object({ type: z.string(), count: z.number() }))
    .describe('Reaction types with counts (e.g., like, heart)'),
  web_url: z.string().describe('Link to the message in Teams'),
});

interface GraphReaction {
  reactionType?: string;
}

export interface GraphChannelMessage {
  id?: string;
  replyToId?: string | null;
  messageType?: string;
  createdDateTime?: string | null;
  lastModifiedDateTime?: string | null;
  deletedDateTime?: string | null;
  subject?: string | null;
  from?: { user?: { id?: string; displayName?: string } | null } | null;
  body?: { contentType?: string; content?: string | null };
  mentions?: Array<{ id?: number; mentionText?: string | null }>;
  attachments?: Array<{ name?: string | null; contentType?: string | null; contentUrl?: string | null }>;
  reactions?: GraphReaction[];
  webUrl?: string | null;
}

export const mapChannelMessage = (m: GraphChannelMessage): z.infer<typeof channelMessageSchema> => {
  const counts = new Map<string, number>();
  for (const r of m.reactions ?? []) {
    if (r.reactionType) counts.set(r.reactionType, (counts.get(r.reactionType) ?? 0) + 1);
  }
  return {
    id: m.id ?? '',
    reply_to_id: m.replyToId ?? '',
    message_type: m.messageType ?? '',
    created_at: m.createdDateTime ?? '',
    last_modified_at: m.lastModifiedDateTime ?? '',
    deleted: !!m.deletedDateTime,
    subject: m.subject ?? '',
    from_name: m.from?.user?.displayName ?? '',
    from_user_id: m.from?.user?.id ?? '',
    content_type: m.body?.contentType ?? '',
    content: m.body?.content ?? '',
    mentions: (m.mentions ?? []).map(x => ({ id: x.id ?? 0, name: x.mentionText ?? '' })),
    attachments: (m.attachments ?? []).map(a => ({
      name: a.name ?? '',
      content_type: a.contentType ?? '',
      url: a.contentUrl ?? '',
    })),
    reactions: [...counts].map(([type, count]) => ({ type, count })),
    web_url: m.webUrl ?? '',
  };
};

export interface GraphPage<T> {
  value?: T[];
  '@odata.nextLink'?: string;
}
