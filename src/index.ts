#!/usr/bin/env node
/**
 * m365-actions-mcp-server
 *
 * MCP server per operazioni di scrittura su Microsoft 365:
 * - Invio mail (send, reply, reply-all)
 * - Gestione calendario (create, update, delete)
 *
 * Da usare accanto al connector ms365 read-only di Anthropic.
 */

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";

import {
  SendMailInputSchema,
  ReplyMailInputSchema,
  GetAttachmentsInputSchema,
  handleSendMail,
  handleReplyMail,
  handleGetAttachments,
} from "./tools/mail.js";

import {
  SearchMailInputSchema,
  handleSearchMail,
} from "./tools/mail-search.js";

import {
  CreateEventInputSchema,
  UpdateEventInputSchema,
  DeleteEventInputSchema,
  handleCreateEvent,
  handleUpdateEvent,
  handleDeleteEvent,
} from "./tools/calendar.js";

import {
  ListTeamsAndChannelsInputSchema,
  SendChannelMessageInputSchema,
  SendChatMessageInputSchema,
  ReplyToMessageInputSchema,
  handleListTeamsAndChannels,
  handleSendChannelMessage,
  handleSendChatMessage,
  handleReplyToMessage,
} from "./tools/teams.js";

// ── Server ───────────────────────────────────────────────

const server = new McpServer({
  name: "m365-actions-mcp-server",
  version: "1.0.0",
});

// ── Mail Tools ───────────────────────────────────────────

server.registerTool(
  "m365_send_mail",
  {
    title: "Invia Mail (Microsoft 365)",
    description: `Invia una nuova email tramite Microsoft 365 / Outlook.

Parametri:
  - to (string | string[]): Uno o più indirizzi email destinatari
  - subject (string): Oggetto della mail
  - body (string): Corpo della mail (HTML di default, o testo semplice)
  - cc (string | string[], opzionale): Destinatari in copia
  - content_type ("HTML" | "Text", default "HTML"): Formato del corpo

Restituisce conferma di invio con destinatari e oggetto.

Usa questo tool per inviare nuove email. Per rispondere a una mail esistente, usa m365_reply_mail.`,
    inputSchema: SendMailInputSchema,
    annotations: {
      readOnlyHint: false,
      destructiveHint: false,
      idempotentHint: false,
      openWorldHint: true,
    },
  },
  async (params) => handleSendMail(params)
);

server.registerTool(
  "m365_reply_mail",
  {
    title: "Rispondi a Mail (Microsoft 365)",
    description: `Rispondi a una mail esistente in Microsoft 365 / Outlook.

Parametri:
  - message_id (string): ID del messaggio a cui rispondere (ottenibile dal tool outlook_email_search del connector ms365)
  - body (string): Contenuto della risposta
  - reply_all (boolean, default false): Se true, risponde a tutti i destinatari

Restituisce conferma di invio.

Usa outlook_email_search per ottenere il message_id prima di chiamare questo tool.`,
    inputSchema: ReplyMailInputSchema,
    annotations: {
      readOnlyHint: false,
      destructiveHint: false,
      idempotentHint: false,
      openWorldHint: true,
    },
  },
  async (params) => handleReplyMail(params)
);

server.registerTool(
  "m365_get_attachments",
  {
    title: "Scarica Allegati Mail (Microsoft 365)",
    description: `Scarica gli allegati di una mail da Microsoft 365 / Outlook e li salva su disco.

Parametri:
  - message_id (string): ID del messaggio (ottenibile da outlook_email_search)
  - save_to (string, opzionale): Percorso cartella dove salvare i file

Restituisce la lista dei file scaricati con nome, percorso e dimensione.

Usa outlook_email_search per trovare la mail, poi questo tool per scaricare gli allegati.`,
    inputSchema: GetAttachmentsInputSchema,
    annotations: {
      readOnlyHint: true,
      destructiveHint: false,
      idempotentHint: true,
      openWorldHint: true,
    },
  },
  async (params) => handleGetAttachments(params)
);


server.registerTool(
  "m365_search_shared_mail",
  {
    title: "Cerca Mail in Casella Delegata (Microsoft 365)",
    description: `Cerca email in una casella delegata/condivisa di Microsoft 365 (es. df@dfsolutions.it, amministrazione@dfsolutions.it).

IMPORTANTE: Questo tool serve SOLO per cercare nelle caselle delegate/condivise.
Per la casella personale dell'utente, usa outlook_email_search del connector ms365 Anthropic.

Parametri:
  - mailbox (string, OBBLIGATORIO): Email della casella delegata (es. df@dfsolutions.it)
  - query (string, opzionale): Ricerca full-text in oggetto e corpo
  - sender (string, opzionale): Filtra per mittente
  - subject (string, opzionale): Filtra per oggetto
  - folder (string, opzionale): Cartella specifica (es. Inbox, Sent Items)
  - after/before (string, opzionale): Range date (ISO 8601)
  - has_attachments (boolean, opzionale): Solo mail con allegati
  - is_read (boolean, opzionale): Filtra per stato lettura
  - limit (number, default 10, max 50): Numero risultati

Restituisce lista di mail con ID, mittente, oggetto, data, anteprima e link web.`,
    inputSchema: SearchMailInputSchema,
    annotations: {
      readOnlyHint: true,
      destructiveHint: false,
      idempotentHint: true,
      openWorldHint: true,
    },
  },
  async (params) => handleSearchMail(params)
);
// ── Calendar Tools ───────────────────────────────────────

server.registerTool(
  "m365_create_event",
  {
    title: "Crea Evento Calendario (Microsoft 365)",
    description: `Crea un nuovo evento nel calendario Outlook / Microsoft 365.

Parametri:
  - subject (string): Titolo dell'evento
  - start (string): Data/ora di inizio ISO 8601 (es. "2026-03-24T10:00:00")
  - end (string): Data/ora di fine ISO 8601 (es. "2026-03-24T11:00:00")
  - attendees (string[], default []): Lista email dei partecipanti
  - body (string, opzionale): Descrizione o note dell'evento
  - location (string, opzionale): Luogo dell'evento
  - is_online_meeting (boolean, default false): Se true, genera un link Teams
  - timezone (string, default "Europe/Rome"): Fuso orario

Restituisce ID dell'evento, link web e link Teams (se online).

Usa find_meeting_availability del connector ms365 per verificare la disponibilità prima di creare l'evento.`,
    inputSchema: CreateEventInputSchema,
    annotations: {
      readOnlyHint: false,
      destructiveHint: false,
      idempotentHint: false,
      openWorldHint: true,
    },
  },
  async (params) => handleCreateEvent(params)
);

server.registerTool(
  "m365_update_event",
  {
    title: "Modifica Evento Calendario (Microsoft 365)",
    description: `Modifica un evento esistente nel calendario Outlook / Microsoft 365.

Parametri:
  - event_id (string): ID dell'evento da modificare (ottenibile da outlook_calendar_search)
  - subject (string, opzionale): Nuovo titolo
  - start (string, opzionale): Nuova data/ora inizio ISO 8601
  - end (string, opzionale): Nuova data/ora fine ISO 8601
  - attendees (string[], opzionale): Nuova lista partecipanti (sovrascrive i precedenti)
  - body (string, opzionale): Nuova descrizione
  - location (string, opzionale): Nuovo luogo
  - is_online_meeting (boolean, opzionale): Attiva/disattiva meeting Teams
  - timezone (string, default "Europe/Rome"): Fuso orario

Fornisci solo i campi da modificare. I campi omessi restano invariati.

Usa outlook_calendar_search del connector ms365 per ottenere l'event_id.`,
    inputSchema: UpdateEventInputSchema,
    annotations: {
      readOnlyHint: false,
      destructiveHint: false,
      idempotentHint: true,
      openWorldHint: true,
    },
  },
  async (params) => handleUpdateEvent(params)
);

server.registerTool(
  "m365_delete_event",
  {
    title: "Elimina Evento Calendario (Microsoft 365)",
    description: `Elimina un evento dal calendario Outlook / Microsoft 365.

Parametri:
  - event_id (string): ID dell'evento da eliminare (ottenibile da outlook_calendar_search)

ATTENZIONE: questa azione è irreversibile. L'evento verrà rimosso definitivamente.

Usa outlook_calendar_search del connector ms365 per ottenere l'event_id prima di chiamare questo tool.`,
    inputSchema: DeleteEventInputSchema,
    annotations: {
      readOnlyHint: false,
      destructiveHint: true,
      idempotentHint: true,
      openWorldHint: true,
    },
  },
  async (params) => handleDeleteEvent(params)
);

// ── Teams Tools ─────────────────────────────────────────

server.registerTool(
  "m365_list_teams_and_channels",
  {
    title: "Elenca Team e Canali (Microsoft Teams)",
    description: `Elenca i team Microsoft Teams dell'utente e i relativi canali.

Parametri:
  - team_id (string, opzionale): ID di un team specifico per elencarne i canali. Se omesso, elenca tutti i team.

Restituisce la lista di team (con ID, nome, descrizione) oppure la lista di canali di un team specifico.

Usa questo tool come primo passo per ottenere gli ID necessari a m365_send_channel_message e m365_reply_to_message.`,
    inputSchema: ListTeamsAndChannelsInputSchema,
    annotations: {
      readOnlyHint: true,
      destructiveHint: false,
      idempotentHint: true,
      openWorldHint: true,
    },
  },
  async (params) => handleListTeamsAndChannels(params)
);

server.registerTool(
  "m365_send_channel_message",
  {
    title: "Invia Messaggio in Canale Teams",
    description: `Invia un messaggio in un canale di Microsoft Teams.

Parametri:
  - team_id (string): ID del team (ottenibile da m365_list_teams_and_channels)
  - channel_id (string): ID del canale (ottenibile da m365_list_teams_and_channels)
  - body (string): Contenuto del messaggio (HTML o testo)
  - subject (string, opzionale): Oggetto/titolo del messaggio
  - content_type ("html" | "text", default "html"): Formato del contenuto

Restituisce l'ID del messaggio e il link diretto.

Complementa chat_message_search del connector ms365 Anthropic (che è read-only) aggiungendo la capacità di scrivere nei canali.`,
    inputSchema: SendChannelMessageInputSchema,
    annotations: {
      readOnlyHint: false,
      destructiveHint: false,
      idempotentHint: false,
      openWorldHint: true,
    },
  },
  async (params) => handleSendChannelMessage(params)
);

server.registerTool(
  "m365_send_chat_message",
  {
    title: "Invia Messaggio Chat Teams (1:1 / Gruppo)",
    description: `Invia un messaggio in una chat Teams esistente (1:1 o di gruppo).

Parametri:
  - chat_id (string): ID della chat (ottenibile da chat_message_search del connector ms365 Anthropic)
  - body (string): Contenuto del messaggio (HTML o testo)
  - content_type ("html" | "text", default "html"): Formato del contenuto

Restituisce l'ID del messaggio e il link diretto.

Usa chat_message_search del connector ms365 Anthropic per trovare la chat e ottenere il chat_id, poi questo tool per inviare il messaggio.`,
    inputSchema: SendChatMessageInputSchema,
    annotations: {
      readOnlyHint: false,
      destructiveHint: false,
      idempotentHint: false,
      openWorldHint: true,
    },
  },
  async (params) => handleSendChatMessage(params)
);

server.registerTool(
  "m365_reply_to_message",
  {
    title: "Rispondi a Messaggio Teams",
    description: `Rispondi a un messaggio esistente in Microsoft Teams (canale o chat).

Parametri:
  - context ("channel" | "chat"): Tipo di contesto del messaggio
  - team_id (string, se context="channel"): ID del team
  - channel_id (string, se context="channel"): ID del canale
  - chat_id (string, se context="chat"): ID della chat
  - message_id (string): ID del messaggio a cui rispondere
  - body (string): Contenuto della risposta
  - content_type ("html" | "text", default "html"): Formato del contenuto

Restituisce l'ID della risposta e il link diretto.

Per i canali: usa m365_list_teams_and_channels per team_id/channel_id.
Per le chat: usa chat_message_search del connector ms365 Anthropic per chat_id e message_id.`,
    inputSchema: ReplyToMessageInputSchema,
    annotations: {
      readOnlyHint: false,
      destructiveHint: false,
      idempotentHint: false,
      openWorldHint: true,
    },
  },
  async (params) => handleReplyToMessage(params)
);

// ── Start ────────────────────────────────────────────────

async function main(): Promise<void> {
  console.error("m365-actions-mcp-server v1.0.0 — avvio in modalità stdio...");

  const transport = new StdioServerTransport();
  await server.connect(transport);

  console.error("Server MCP connesso e in ascolto.");
}

main().catch((error) => {
  console.error("Errore fatale:", error);
  process.exit(1);
});

