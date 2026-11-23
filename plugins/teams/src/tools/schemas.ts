import { z } from 'zod';

// ---------------------------------------------------------------------------
// Conversation (chat) schema
// ---------------------------------------------------------------------------

export const conversationSchema = z.object({
  id: z.string().describe('Conversation/thread ID'),
  topic: z.string().describe('Chat name/topic (empty for unnamed 1:1 chats)'),
  type: z.string().describe('Conversation type (e.g., "Conversation")'),
  thread_type: z.string().describe('Thread type (e.g., "chat", "meeting", "space")'),
  member_count: z.string().describe('Number of members in the conversation'),
  created_at: z.string().describe('When the conversation was created (ISO 8601)'),
  last_message_content: z.string().describe('Content of the last message'),
  last_message_type: z.string().describe('Message type of the last message'),
  last_message_time: z.string().describe('Timestamp of the last message'),
  last_message_from: z.string().describe('Display name of the last message sender'),
  version: z.number().describe('Conversation version'),
});

export type Conversation = z.infer<typeof conversationSchema>;

interface RawConversation {
  id?: string;
  type?: string;
  threadProperties?: Record<string, unknown>;
  lastMessage?: Record<string, unknown>;
  version?: number;
}

export const mapConversation = (c: RawConversation): Conversation => ({
  id: c.id ?? '',
  topic: String(c.threadProperties?.topic ?? ''),
  type: c.type ?? '',
  thread_type: String(c.threadProperties?.threadType ?? ''),
  member_count: String(c.threadProperties?.memberCount ?? ''),
  created_at: String(c.threadProperties?.createdat ?? ''),
  last_message_content: String(c.lastMessage?.content ?? ''),
  last_message_type: String(c.lastMessage?.messagetype ?? ''),
  last_message_time: String(c.lastMessage?.composetime ?? ''),
  last_message_from: String(c.lastMessage?.imdisplayname ?? ''),
  version: typeof c.version === 'number' ? c.version : 0,
});

// ---------------------------------------------------------------------------
// Mention schema
// ---------------------------------------------------------------------------

const mentionSchema = z.object({
  mri: z.string().describe('Mentioned user MRI (e.g., "8:orgid:...")'),
  display_name: z.string().describe('Display name of the mentioned user'),
});

// ---------------------------------------------------------------------------
// File attachment schema
// ---------------------------------------------------------------------------

const fileSchema = z.object({
  title: z.string().describe('File name'),
  type: z.string().describe('File type/extension (e.g., "png", "pdf")'),
  url: z.string().describe('File object URL (SharePoint/OneDrive; opening it requires the signed-in browser session)'),
  share_url: z.string().describe('Sharing URL if the file has one, else empty'),
  item_id: z.string().describe('Teams/SharePoint file item ID'),
});

// ---------------------------------------------------------------------------
// Message schema
// ---------------------------------------------------------------------------

export const messageSchema = z.object({
  id: z
    .string()
    .describe(
      'Message ID (arrival time in epoch ms); pass to edit_message, delete_message, set_reaction, remove_reaction',
    ),
  client_message_id: z.string().describe('Client-assigned message ID'),
  content: z.string().describe('Message content (may contain HTML)'),
  message_type: z
    .string()
    .describe('Message type (e.g., "RichText/Html", "Text", "ThreadActivity/AddMember", "RichText/Media_Card")'),
  from: z.string().describe('Sender MRI (e.g., "8:orgid:username")'),
  display_name: z.string().describe('Sender display name'),
  compose_time: z.string().describe('When the message was composed (ISO 8601)'),
  conversation_id: z.string().describe('ID of the conversation this message belongs to'),
  mentions: z.array(mentionSchema).describe('Users mentioned in this message'),
  files: z.array(fileSchema).describe('Files attached to this message'),
  reactions: z.array(z.string()).describe('Reaction emoji keys on this message (e.g., ["like", "heart"])'),
  is_deleted: z.boolean().describe('Whether the message was deleted (its content is then empty)'),
});

export type Message = z.infer<typeof messageSchema>;

interface RawMention {
  mri?: string;
  displayName?: string;
}

interface RawFile {
  title?: string;
  fileName?: string;
  type?: string;
  fileType?: string;
  objectUrl?: string;
  itemid?: string;
  id?: string;
  fileInfo?: { fileUrl?: string | null; shareUrl?: string | null };
}

interface RawMessage {
  id?: string;
  clientmessageid?: string;
  content?: string;
  messagetype?: string;
  from?: string;
  imdisplayname?: string;
  composetime?: string;
  conversationid?: string;
  properties?: {
    mentions?: string;
    files?: string;
    emotions?: Array<{ key?: string; users?: unknown[] }>;
    deletetime?: number | string;
  };
  annotationsSummary?: {
    emotions?: Array<{ key?: string; users?: unknown[] }>;
  };
}

/** Extract the MRI (e.g., "8:live:username") from a full contact URL. */
const extractMri = (from: string): string => {
  const match = /\/contacts\/(.+)$/.exec(from);
  return match?.[1] ?? from;
};

/** Parse mentions from the serialized JSON string in properties.mentions. */
const parseMentions = (raw?: string): Array<{ mri: string; display_name: string }> => {
  if (!raw) return [];
  try {
    const arr = JSON.parse(raw) as RawMention[];
    return arr.map(m => ({ mri: m.mri ?? '', display_name: m.displayName ?? '' }));
  } catch {
    return [];
  }
};

/** Parse files from the serialized JSON string in properties.files. */
const parseFiles = (raw?: string): Message['files'] => {
  if (!raw) return [];
  try {
    const arr = JSON.parse(raw) as RawFile[];
    return arr.map(f => ({
      title: f.title || f.fileName || '',
      type: f.fileType || f.type || '',
      url: f.objectUrl || f.fileInfo?.fileUrl || '',
      share_url: f.fileInfo?.shareUrl ?? '',
      item_id: f.itemid || f.id || '',
    }));
  } catch {
    return [];
  }
};

/**
 * Extract unique reaction keys from emotions/annotationsSummary. Removing the last
 * reaction of a kind leaves its entry behind with an empty `users` list, so only
 * entries with at least one reacting user (or no user list at all) count.
 */
const parseReactions = (msg: RawMessage): string[] => {
  const emotions = msg.properties?.emotions ?? msg.annotationsSummary?.emotions ?? [];
  const keys = new Set<string>();
  for (const e of emotions) {
    if (e.key && (!Array.isArray(e.users) || e.users.length > 0)) keys.add(e.key);
  }
  return [...keys];
};

export const mapMessage = (m: RawMessage): Message => ({
  id: m.id ?? '',
  client_message_id: m.clientmessageid ?? '',
  content: m.content ?? '',
  message_type: m.messagetype ?? '',
  from: extractMri(m.from ?? ''),
  display_name: m.imdisplayname ?? '',
  compose_time: m.composetime ?? '',
  conversation_id: m.conversationid ?? '',
  mentions: parseMentions(m.properties?.mentions),
  files: parseFiles(m.properties?.files),
  reactions: parseReactions(m),
  is_deleted: Boolean(m.properties?.deletetime),
});
