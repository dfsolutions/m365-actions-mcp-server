import { z } from "zod";
import { getGraphClient } from "../graphClient.js";
import { handleGraphError } from "../utils/errors.js";
import { DEFAULT_TIMEZONE } from "../constants.js";
import type {
  CalendarEventPayload,
  CalendarEventUpdatePayload,
} from "../types.js";

// ── Schemas ──────────────────────────────────────────────

export const CreateEventInputSchema = z
  .object({
    subject: z
      .string()
      .min(1, "L'oggetto dell'evento è obbligatorio")
      .describe("Titolo dell'evento"),
    start: z
      .string()
      .min(1)
      .describe("Data/ora di inizio in formato ISO 8601 (es. 2026-03-24T10:00:00)"),
    end: z
      .string()
      .min(1)
      .describe("Data/ora di fine in formato ISO 8601 (es. 2026-03-24T11:00:00)"),
    attendees: z
      .array(z.string().email())
      .default([])
      .describe("Lista di email dei partecipanti (opzionale)"),
    body: z
      .string()
      .optional()
      .describe("Descrizione/note dell'evento (HTML o testo, opzionale)"),
    location: z
      .string()
      .optional()
      .describe("Luogo dell'evento (opzionale)"),
    is_online_meeting: z
      .boolean()
      .default(false)
      .describe("Se true, crea un meeting Teams associato (default: false)"),
    sensitivity: z
      .enum(["normal", "personal", "private", "confidential"])
      .default("normal")
      .describe("Visibilità: normal (default), personal, private (nasconde dettagli agli altri), confidential"),
    show_as: z
      .enum(["free", "tentative", "busy", "oof", "workingElsewhere", "unknown"])
      .default("busy")
      .describe("Stato disponibilità: busy (default), free, tentative, oof (fuori ufficio), workingElsewhere"),
    timezone: z
      .string()
      .default(DEFAULT_TIMEZONE)
      .describe(`Fuso orario (default: ${DEFAULT_TIMEZONE})`),
  })
  .strict();

export type CreateEventInput = z.infer<typeof CreateEventInputSchema>;

export const UpdateEventInputSchema = z
  .object({
    event_id: z
      .string()
      .min(1, "L'ID dell'evento è obbligatorio")
      .describe("ID dell'evento da modificare (ottenibile da outlook_calendar_search)"),
    subject: z.string().optional().describe("Nuovo titolo (opzionale)"),
    start: z
      .string()
      .optional()
      .describe("Nuova data/ora di inizio ISO 8601 (opzionale)"),
    end: z
      .string()
      .optional()
      .describe("Nuova data/ora di fine ISO 8601 (opzionale)"),
    attendees: z
      .array(z.string().email())
      .optional()
      .describe("Nuova lista partecipanti (opzionale, sovrascrive i precedenti)"),
    body: z.string().optional().describe("Nuova descrizione (opzionale)"),
    location: z.string().optional().describe("Nuovo luogo (opzionale)"),
    is_online_meeting: z
      .boolean()
      .optional()
      .describe("Attiva/disattiva meeting Teams (opzionale)"),
    sensitivity: z
      .enum(["normal", "personal", "private", "confidential"])
      .optional()
      .describe("Visibilità: normal, personal, private (nasconde dettagli agli altri), confidential"),
    show_as: z
      .enum(["free", "tentative", "busy", "oof", "workingElsewhere", "unknown"])
      .optional()
      .describe("Stato disponibilità: busy, free, tentative, oof (fuori ufficio), workingElsewhere"),
    timezone: z
      .string()
      .default(DEFAULT_TIMEZONE)
      .describe(`Fuso orario (default: ${DEFAULT_TIMEZONE})`),
  })
  .strict();

export type UpdateEventInput = z.infer<typeof UpdateEventInputSchema>;

export const DeleteEventInputSchema = z
  .object({
    event_id: z
      .string()
      .min(1, "L'ID dell'evento è obbligatorio")
      .describe("ID dell'evento da eliminare (ottenibile da outlook_calendar_search)"),
  })
  .strict();

export type DeleteEventInput = z.infer<typeof DeleteEventInputSchema>;

// ── Handlers ─────────────────────────────────────────────

export async function handleCreateEvent(params: CreateEventInput): Promise<{
  content: Array<{ type: "text"; text: string }>;
  structuredContent?: Record<string, unknown>;
}> {
  try {
    const client = await getGraphClient();

    const payload: CalendarEventPayload = {
      subject: params.subject,
      start: { dateTime: params.start, timeZone: params.timezone },
      end: { dateTime: params.end, timeZone: params.timezone },
      ...(params.attendees.length > 0
        ? {
            attendees: params.attendees.map((email) => ({
              emailAddress: { address: email },
              type: "required" as const,
            })),
          }
        : {}),
      ...(params.body
        ? { body: { contentType: "HTML" as const, content: params.body } }
        : {}),
      ...(params.location
        ? { location: { displayName: params.location } }
        : {}),
      ...(params.is_online_meeting
        ? { isOnlineMeeting: true, onlineMeetingProvider: "teamsForBusiness" as const }
        : {}),
      ...(params.sensitivity !== "normal"
        ? { sensitivity: params.sensitivity }
        : {}),
      ...(params.show_as !== "busy"
        ? { showAs: params.show_as }
        : {}),
    };

    const event = await client.api("/me/events").post(payload);

    const output = {
      status: "created",
      event_id: event.id,
      subject: event.subject,
      start: event.start,
      end: event.end,
      web_link: event.webLink,
      ...(event.onlineMeeting?.joinUrl
        ? { teams_link: event.onlineMeeting.joinUrl }
        : {}),
    };

    let text = `Evento "${params.subject}" creato con successo (${params.start} → ${params.end})`;
    if (event.onlineMeeting?.joinUrl) {
      text += `\nLink Teams: ${event.onlineMeeting.joinUrl}`;
    }

    return {
      content: [{ type: "text" as const, text }],
      structuredContent: output,
    };
  } catch (error) {
    return {
      content: [{ type: "text" as const, text: handleGraphError(error) }],
    };
  }
}

export async function handleUpdateEvent(params: UpdateEventInput): Promise<{
  content: Array<{ type: "text"; text: string }>;
  structuredContent?: Record<string, unknown>;
}> {
  try {
    const client = await getGraphClient();

    const payload: CalendarEventUpdatePayload = {};

    if (params.subject) payload.subject = params.subject;
    if (params.start)
      payload.start = { dateTime: params.start, timeZone: params.timezone };
    if (params.end)
      payload.end = { dateTime: params.end, timeZone: params.timezone };
    if (params.body)
      payload.body = { contentType: "HTML", content: params.body };
    if (params.location)
      payload.location = { displayName: params.location };
    if (params.attendees) {
      payload.attendees = params.attendees.map((email) => ({
        emailAddress: { address: email },
        type: "required" as const,
      }));
    }
    if (params.is_online_meeting !== undefined) {
      payload.isOnlineMeeting = params.is_online_meeting;
      if (params.is_online_meeting) {
        payload.onlineMeetingProvider = "teamsForBusiness";
      }
    }
    if (params.sensitivity) {
      payload.sensitivity = params.sensitivity;
    }
    if (params.show_as) {
      payload.showAs = params.show_as;
    }

    const event = await client
      .api(`/me/events/${params.event_id}`)
      .patch(payload);

    const output = {
      status: "updated",
      event_id: event.id,
      subject: event.subject,
      start: event.start,
      end: event.end,
    };

    return {
      content: [
        {
          type: "text" as const,
          text: `Evento "${event.subject}" aggiornato con successo.`,
        },
      ],
      structuredContent: output,
    };
  } catch (error) {
    return {
      content: [{ type: "text" as const, text: handleGraphError(error) }],
    };
  }
}

export async function handleDeleteEvent(params: DeleteEventInput): Promise<{
  content: Array<{ type: "text"; text: string }>;
  structuredContent?: Record<string, unknown>;
}> {
  try {
    const client = await getGraphClient();

    await client.api(`/me/events/${params.event_id}`).delete();

    const output = {
      status: "deleted",
      event_id: params.event_id,
    };

    return {
      content: [
        {
          type: "text" as const,
          text: `Evento ${params.event_id} eliminato con successo.`,
        },
      ],
      structuredContent: output,
    };
  } catch (error) {
    return {
      content: [{ type: "text" as const, text: handleGraphError(error) }],
    };
  }
}
