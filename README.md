# m365-actions-mcp-server

MCP server for write operations on Microsoft 365 via Microsoft Graph API.

Designed to work alongside Anthropic's read-only ms365 connector to provide full bidirectional access.

## Available Tools

| Tool | Description |
|------|-------------|
| `m365_send_mail` | Send a new email |
| `m365_reply_mail` | Reply to an existing email (reply / reply-all) |
| `m365_get_attachments` | Download attachments from an email |
| `m365_create_event` | Create a calendar event |
| `m365_update_event` | Update an existing event |
| `m365_delete_event` | Delete an event |

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
   - `Calendars.ReadWrite`
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