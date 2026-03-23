# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Comandi principali

```bash
npm run build      # Compila TypeScript → dist/
npm run dev        # Avvia in watch mode (tsx, no build richiesto)
npm run start      # Esegue il server compilato (dist/index.js)
npm run clean      # Rimuove dist/
```

Non ci sono test runner né linter configurati.

## Configurazione ambiente

Copia `.env.example` in `.env` e compila le variabili Azure AD:
- `M365_CLIENT_ID` — App registration ID
- `M365_CLIENT_SECRET` — App secret
- `M365_TENANT_ID` — Tenant ID
- `M365_REDIRECT_URI` — Default: `http://localhost:3939/auth/callback`

## Architettura

**Tipo di progetto:** MCP server che espone operazioni di scrittura su Microsoft 365 (Mail e Calendar) tramite Microsoft Graph API. Complementa il connector read-only di Anthropic.

**Flusso principale:**
1. Claude o altro client MCP chiama uno dei 6 tool registrati in `src/index.ts`
2. Il tool handler (in `src/tools/`) valida i parametri con Zod e chiama il Graph client
3. `src/graphClient.ts` ottiene un token OAuth valido tramite `src/auth.ts` (MSAL)
4. La risposta Graph viene mappata in contenuto MCP (testo + JSON strutturato)

**Autenticazione (`src/auth.ts`):**
- OAuth 2.0 con PKCE tramite MSAL Node
- Prima tenta silenziosamente dal cache (`~/.m365-actions-tokens.json`)
- Se il cache è vuoto, avvia un server locale su porta 3939 e apre il browser per il login interattivo
- I token vengono persistiti su disco e auto-rinnovati

**Tool MCP registrati (`src/index.ts`):**
| Tool | Descrizione |
|------|-------------|
| `m365_send_mail` | Invia nuova email |
| `m365_reply_mail` | Risponde a email esistente (per message ID) |
| `m365_create_event` | Crea evento calendario |
| `m365_update_event` | Modifica evento esistente (aggiornamento parziale) |
| `m365_delete_event` | Elimina evento (marcato come destructive) |

**Costanti e defaults (`src/constants.ts`):**
- Timezone default: `Europe/Rome`
- Token cache path: `~/.m365-actions-tokens.json`
- Safety limit corpo messaggi: 25.000 caratteri

**Gestione errori (`src/utils/errors.ts`):**
Mapper centralizzato per codici HTTP Graph API (400, 401, 403, 404, 429). Da consultare prima di aggiungere nuova logica di errore.

## Note operative

- Il progetto usa ES modules (`"type": "module"` in package.json) con `moduleResolution: Node16` — le importazioni devono includere l'estensione `.js` anche per file `.ts`
- TypeScript strict mode attivo
- Le risposte dei tool includono sia testo human-readable che `structuredContent` JSON
