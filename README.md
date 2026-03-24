# m365-actions-mcp-server

MCP server that adds **write operations** and **shared mailbox search** to Microsoft 365 — the missing piece that turns Anthropic's read-only ms365 connector into a full bidirectional integration.

## Why this exists

Anthropic's built-in Microsoft 365 connector gives Claude powerful read access: searching emails, browsing calendars, finding Teams messages, querying SharePoint. But it **cannot send, create, update, or delete anything**, and it has **limited support for shared/delegated mailboxes**. This server fills both gaps.

With both connectors running side by side, Claude can complete full workflows end-to-end: read an email and reply to it, check calendar availability and book a meeting, find a Teams conversation and respond — all without leaving the chat.

## Available Tools

### Email

| Tool | Description | Works with |
|------|-------------|------------|
| `m365_send_mail` | Send a new email | — |
| `m365_reply_mail` | Reply to an existing email (reply / reply-all) | `outlook_email_search` (Anthropic) to get message ID |
| `m365_get_attachments` | Download attachments from an email | `outlook_email_search` (Anthropic) to get message ID |
| `m365_search_shared_mail` | Search emails in shared/delegated mailboxes | Standalone, uses Graph API /users/{mailbox}/messages with KQL search |

### Calendar

| Tool | Description | Works with |
|------|-------------|------------|
| `m365_create_event` | Create a calendar event (with optional Teams link) | `find_meeting_availability` (Anthropic) to check slots |
| `m365_update_event` | Update an existing event | `outlook_calendar_search` (Anthropic) to get event ID |
| `m365_delete_event` | Delete an event | `outlook_calendar_search` (Anthropic) to get event ID |

### Teams

| Tool | Description | Works with |
|------|-------------|------------|
| `m365_list_teams_and_channels` | List joined teams and their channels | — |
| `m365_send_channel_message` | Send a message to a Teams channel | `m365_list_teams_and_channels` to get IDs |
| `m365_send_chat_message` | Send a message in a Teams chat (1:1 or group) | `chat_message_search` (Anthropic) to get chat ID |
| `m365_reply_to_message` | Reply to an existing Teams message | `chat_message_search` (Anthropic) to get message ID |

## How it works with the Anthropic connector

```
Anthropic ms365 connector (read)     m365-actions-mcp-server (write)
─────────────────────────────────    ──────────────────────────────────
outlook_email_search          ──→    m365_send_mail / m365_reply_mail
outlook_calendar_search       ──→    m365_create_event / m365_update_event
find_meeting_availability     ──→    m365_create_event
chat_message_search           ──→    m365_send_chat_message / m365_reply_to_message
sharepoint_search             ──→    (read-only, no write needed)
```

Claude automatically chains these tools together. For example: "Find the email from Marco about the invoice and reply saying we'll pay by Friday" triggers `outlook_email_search` (Anthropic) followed by `m365_reply_mail` (this server).

## Shared Mailbox Search

The m365_search_shared_mail tool fills a gap in the Anthropic connector: reliable search across shared/delegated mailboxes. It accesses /users/{mailbox}/messages directly via Microsoft Graph API, supporting:

- **KQL search** (from:, subject:, free text) via the search parameter
- **Date range filtering** via OData filter
- **Pagination fallback** with client-side filtering when search is unavailable
- **Folder targeting** (Inbox, Sent Items, etc.)

This is needed because the Anthropic connector outlook_email_search has unreliable indexing for shared mailboxes, and may miss emails in high-volume mailboxes.

The Mail.Read.Shared delegated permission must be granted in Azure AD for this tool to work.

## Prerequisites

- Node.js >= 18
- An app registered on Azure AD / Entra ID

### Azure App Configuration

1. Go to [Entra ID](https://portal.azure.com) -> App registrations -> New registration
2. Name: `M365 Actions MCP Server`
3. Supported account types: **Single tenant**
4. Redirect URI: platform **Web**, URI `http://localhost:3939/auth/callback`
5. Create a **Client Secret** (Certificates & secrets -> New client secret)
6. Add **Delegated** permissions (API permissions -> Microsoft Graph):
   - `Mail.Send`
   - `Mail.Read`
   - `Mail.Read.Shared`
   - `Calendars.ReadWrite`
   - `Chat.ReadWrite`
   - `ChannelMessage.Send`
   - `Team.ReadBasic.All`
   - `Channel.ReadBasic.All`
   - `User.Read`
   - `offline_access`
7. Click **Grant admin consent**

## Installation

```bash
git clone https://github.com/dfsolutions/m365-actions-mcp-server.git
cd m365-actions-mcp-server
npm install
```

## Configuration

Copy `.env.example` to `.env` and fill in your Azure app credentials:

```bash
cp .env.example .env
```

```env
M365_CLIENT_ID=your-client-id
M365_CLIENT_SECRET=your-client-secret
M365_TENANT_ID=your-tenant-id
M365_REDIRECT_URI=http://localhost:3939/auth/callback
M365_USER_EMAIL=your-email@domain.com
```

## Build and Run

```bash
npx tsc
node dist/index.js
```

On the first run, the browser will open for Microsoft login. The token is cached in `~/.m365-actions-tokens.json` and reused automatically.

## Integration with Claude Desktop / Cowork

Add to your MCP configuration file:

```json
{
  "mcpServers": {
    "m365-actions": {
      "command": "node",
      "args": ["/full/path/to/m365-actions-mcp-server/dist/index.js"],
      "env": {
        "M365_CLIENT_ID": "your-client-id",
        "M365_CLIENT_SECRET": "your-client-secret",
        "M365_TENANT_ID": "your-tenant-id",
        "M365_REDIRECT_URI": "http://localhost:3939/auth/callback",
        "M365_USER_EMAIL": "your-email@domain.com"
      }
    }
  }
}
```

Alternatively, if the variables are in the `.env` file inside the project folder:

```json
{
  "mcpServers": {
    "m365-actions": {
      "command": "node",
      "args": ["/full/path/to/m365-actions-mcp-server/dist/index.js"]
    }
  }
}
```

## Authentication

The server uses MSAL with the **Authorization Code (delegated)** flow:
- On first run, it opens the browser for login
- Saves access + refresh token in `~/.m365-actions-tokens.json`
- Automatically renews the token on expiry
- The refresh token lasts ~90 days, then a new login is required

## License

MIT
