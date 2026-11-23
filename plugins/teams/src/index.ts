import { OpenTabsPlugin } from '@opentabs-dev/plugin-sdk';
import type { ToolDefinition } from '@opentabs-dev/plugin-sdk';
import { clearCaches, isTeamsAuthenticated, waitForTeamsAuth, watchForAuthAndNotify } from './teams-api.js';
import { addChatMember } from './tools/add-chat-member.js';
import { createChat } from './tools/create-chat.js';
import { deleteMessage } from './tools/delete-message.js';
import { editMessage } from './tools/edit-message.js';
import { getConversationDetails } from './tools/get-conversation-details.js';
import { getCurrentUser } from './tools/get-current-user.js';
import { getPresence } from './tools/get-presence.js';
import { listChannels } from './tools/list-channels.js';
import { listConversations } from './tools/list-conversations.js';
import { listTeams } from './tools/list-teams.js';
import { markChatRead } from './tools/mark-chat-read.js';
import { postChannelMessage } from './tools/post-channel-message.js';
import { readChannelMessages } from './tools/read-channel-messages.js';
import { readChannelReplies } from './tools/read-channel-replies.js';
import { readMessages } from './tools/read-messages.js';
import { removeChatMember } from './tools/remove-chat-member.js';
import { removeReaction } from './tools/remove-reaction.js';
import { replyToChannelMessage } from './tools/reply-to-channel-message.js';
import { searchMessages } from './tools/search-messages.js';
import { sendMessage } from './tools/send-message.js';
import { setChatTopic } from './tools/set-chat-topic.js';
import { setReaction } from './tools/set-reaction.js';

class TeamsPlugin extends OpenTabsPlugin {
  readonly name = 'teams';
  readonly description =
    'Microsoft Teams — read, search and send chat messages, read and post team channel messages, reactions, and presence';
  override readonly displayName = 'Microsoft Teams';
  readonly urlPatterns = ['*://teams.live.com/*', '*://teams.microsoft.com/*', '*://teams.cloud.microsoft/*'];
  override readonly homepage = 'https://teams.live.com/v2/';
  readonly tools: ToolDefinition[] = [
    // Chats
    listConversations,
    getConversationDetails,
    createChat,
    setChatTopic,
    markChatRead,
    // Messages
    sendMessage,
    readMessages,
    editMessage,
    deleteMessage,
    setReaction,
    removeReaction,
    searchMessages,
    // Members
    addChatMember,
    removeChatMember,
    // Teams & channels
    listTeams,
    listChannels,
    readChannelMessages,
    readChannelReplies,
    postChannelMessage,
    replyToChannelMessage,
    // People
    getCurrentUser,
    getPresence,
  ];

  override teardown(): void {
    clearCaches();
  }

  /**
   * Ready only when a Skype JWT can actually be obtained (the same path every
   * tool uses). The wait stays under the extension's 5 s `isReady()` budget; if
   * auth is not there yet, keep watching in the background and ask the extension
   * to re-probe as soon as it appears.
   */
  async isReady(): Promise<boolean> {
    if (isTeamsAuthenticated()) return true;
    if (await waitForTeamsAuth()) return true;
    watchForAuthAndNotify();
    return false;
  }
}

export default new TeamsPlugin();
