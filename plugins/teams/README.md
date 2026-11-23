# Microsoft Teams

Microsoft Teams — read, search and send chat messages, read and post team channel messages, reactions, and presence — gives AI agents access to Microsoft Teams through your authenticated browser session.

## Install

```bash
opentabs plugin install teams
```

Or install globally via npm:

```bash
npm install -g @opentabs-dev/opentabs-plugin-teams
```

## Setup

1. Open [teams.live.com](https://teams.live.com/v2/) in Chrome and log in
2. Open the OpenTabs side panel — the Microsoft Teams plugin should appear as **ready**

## Tools (22)

### Chats (5)

| Tool | Description | Type |
|---|---|---|
| `list_conversations` | List chats (paginated) | Read |
| `get_conversation_details` | Get chat details and members | Read |
| `create_chat` | Create a new chat conversation | Write |
| `set_chat_topic` | Set a chat topic | Write |
| `mark_chat_read` | Mark a chat as read | Write |

### Messages (7)

| Tool | Description | Type |
|---|---|---|
| `send_message` | Send a message to a chat | Write |
| `read_messages` | Read messages from a chat (paginated) | Read |
| `edit_message` | Edit a chat message | Write |
| `delete_message` | Delete a chat message | Write |
| `set_reaction` | React to a chat message | Write |
| `remove_reaction` | Remove a reaction from a chat message | Write |
| `search_messages` | Search Teams messages | Read |

### Members (2)

| Tool | Description | Type |
|---|---|---|
| `add_chat_member` | Add a user to a group chat | Write |
| `remove_chat_member` | Remove a user from a group chat | Write |

### Teams (6)

| Tool | Description | Type |
|---|---|---|
| `list_teams` | List joined teams | Read |
| `list_channels` | List channels in a team | Read |
| `read_channel_messages` | Read messages in a channel | Read |
| `read_channel_replies` | Read replies to a channel message | Read |
| `post_channel_message` | Post a message to a channel | Write |
| `reply_to_channel_message` | Reply to a channel message | Write |

### People (2)

| Tool | Description | Type |
|---|---|---|
| `get_current_user` | Get current user info | Read |
| `get_presence` | Get user presence | Read |

## How It Works

This plugin runs inside your Microsoft Teams tab through the [OpenTabs](https://opentabs.dev) Chrome extension. It uses your existing browser session — no API tokens or OAuth apps required. All operations happen as you, with your permissions.

## License

MIT
