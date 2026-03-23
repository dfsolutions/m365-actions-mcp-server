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
  CreateEventInputSchema,
  UpdateEventInputSchema,
  DeleteEventInputSchema,
  handleCreateEvent,
  handleUpdateEvent,
  handleDeleteEvent,
} from "./tools/calendar.js";

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
