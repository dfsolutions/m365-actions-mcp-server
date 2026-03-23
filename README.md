# m365-actions-mcp-server

MCP server per operazioni di scrittura su Microsoft 365 tramite Microsoft Graph API.

Da usare accanto al connector ms365 read-only di Anthropic per avere sia lettura che scrittura.

## Tool disponibili

| Tool | Descrizione |
|------|-------------|
| `m365_send_mail` | Invia una nuova email |
| `m365_reply_mail` | Rispondi a una mail esistente (reply / reply-all) |
| `m365_get_attachments` | Scarica gli allegati di una mail |
| `m365_create_event` | Crea un evento nel calendario |
| `m365_update_event` | Modifica un evento esistente |
| `m365_delete_event` | Elimina un evento |

## Prerequisiti

- Node.js >= 18
- Un'app registrata su Azure AD / Entra ID

### Configurazione app Azure

1. Vai su [Entra ID](https://portal.azure.com) → App registrations → New registration
2. Nome: `M365 Actions MCP Server`
3. Supported account types: **Single tenant**
4. Redirect URI: piattaforma **Web**, URI `http://localhost:3939/auth/callback`
5. Crea un **Client Secret** (Certificates & secrets → New client secret)
6. Aggiungi i permessi **Delegated** (API permissions → Microsoft Graph):
   - `Mail.Send`
   - `Mail.Read`
   - `Calendars.ReadWrite`
   - `User.Read`
   - `offline_access`
7. Premi **Grant admin consent**

## Installazione

```bash
git clone https://github.com/TUOUSER/m365-actions-mcp-server.git
cd m365-actions-mcp-server
npm install
```

## Configurazione

Copia `.env.example` in `.env` e compila con i dati della tua app Azure:

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

## Build e avvio

```bash
npx tsc
node dist/index.js
```

Al primo avvio si aprirà il browser per il login Microsoft. Il token viene cachato in `~/.m365-actions-tokens.json` e riutilizzato automaticamente.

## Integrazione con Claude Desktop / Cowork

Aggiungi al file di configurazione MCP:

```json
{
  "mcpServers": {
    "m365-actions": {
      "command": "node",
      "args": ["/percorso/completo/m365-actions-mcp-server/dist/index.js"],
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

In alternativa, se le variabili sono nel `.env` nella cartella del progetto, basta:

```json
{
  "mcpServers": {
    "m365-actions": {
      "command": "node",
      "args": ["/percorso/completo/m365-actions-mcp-server/dist/index.js"]
    }
  }
}
```

## Autenticazione

Il server usa MSAL con flusso **Authorization Code (delegated)**:
- Al primo avvio apre il browser per il login
- Salva access + refresh token in `~/.m365-actions-tokens.json`
- Rinnova automaticamente il token alla scadenza
- Il refresh token dura ~90 giorni, poi serve un nuovo login

## Licenza

MIT
