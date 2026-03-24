import { z } from "zod";
import { getGraphClient } from "../graphClient.js";
import { handleGraphError } from "../utils/errors.js";
import type {
  TeamsChatMessagePayload,
  TeamsChannelMessagePayload,
} from "../types.js";

// ── Schemas ──────────────────────────────────────────────

export const ListTeamsAndChannelsInputSchema = z
  .object({
    team_id: z
      .string()
      .optional()
      .describe(
        "ID del team di cui elencare i canali (opzionale). Se omesso, elenca tutti i team dell'utente."
      ),
  })
  .strict();

export type ListTeamsAndChannelsInput = z.infer<typeof ListTeamsAndChannelsInputSchema>;

export const SendChannelMessageInputSchema = z
  .object({
    team_id: z
      .string()
      .min(1, "L'ID del team è obbligatorio")
      .describe("ID del team (ottenibile da m365_list_teams_and_channels)"),
    channel_id: z
      .string()
      .min(1, "L'ID del canale è obbligatorio")
      .describe("ID del canale (ottenibile da m365_list_teams_and_channels)"),
    body: z
      .string()
      .min(1, "Il corpo del messaggio non può essere vuoto")
      .describe("Contenuto del messaggio (HTML o testo semplice)"),
    subject: z
      .string()
      .optional()
      .describe("Oggetto del messaggio (opzionale, visibile come titolo nel canale)"),
    content_type: z
      .enum(["html", "text"])
      .default("html")
      .describe("Tipo di contenuto: html (default) o text"),
  })
  .strict();

export type SendChannelMessageInput = z.infer<typeof SendChannelMessageInputSchema>;

export const SendChatMessageInputSchema = z
  .object({
    chat_id: z
      .string()
      .min(1, "L'ID della chat è obbligatorio")
      .describe(
        "ID della chat Teams (1:1 o di gruppo). Ottenibile da chat_message_search del connector ms365."
      ),
    body: z
      .string()
      .min(1, "Il corpo del messaggio non può essere vuoto")
      .describe("Contenuto del messaggio (HTML o testo semplice)"),
    content_type: z
      .enum(["html", "text"])
      .default("html")
      .describe("Tipo di contenuto: html (default) o text"),
  })
  .strict();

export type SendChatMessageInput = z.infer<typeof SendChatMessageInputSchema>;

export const ReplyToMessageInputSchema = z
  .object({
    context: z
      .enum(["channel", "chat"])
      .describe("Contesto del messaggio: 'channel' per canali team, 'chat' per chat dirette/gruppo"),
    team_id: z
      .string()
      .optional()
      .describe("ID del team (obbligatorio se context = 'channel')"),
    channel_id: z
      .string()
      .optional()
      .describe("ID del canale (obbligatorio se context = 'channel')"),
    chat_id: z
      .string()
      .optional()
      .describe("ID della chat (obbligatorio se context = 'chat')"),
    message_id: z
      .string()
      .min(1, "L'ID del messaggio è obbligatorio")
      .describe("ID del messaggio a cui rispondere"),
    body: z
      .string()
      .min(1, "Il corpo della risposta non può essere vuoto")
      .describe("Contenuto della risposta (HTML o testo semplice)"),
    content_type: z
      .enum(["html", "text"])
      .default("html")
      .describe("Tipo di contenuto: html (default) o text"),
  })
  .strict()
  .refine(
    (data) => {
      if (data.context === "channel") return !!data.team_id && !!data.channel_id;
      if (data.context === "chat") return !!data.chat_id;
      return false;
    },
    {
      message:
        "Per context='channel' servono team_id e channel_id. Per context='chat' serve chat_id.",
    }
  );

export type ReplyToMessageInput = z.infer<typeof ReplyToMessageInputSchema>;

// ── Handlers ─────────────────────────────────────────────

export async function handleListTeamsAndChannels(
  params: ListTeamsAndChannelsInput
): Promise<{
  content: Array<{ type: "text"; text: string }>;
  structuredContent?: Record<string, unknown>;
}> {
  try {
    const client = await getGraphClient();

    if (params.team_id) {
      // List channels for a specific team
      const response = await client
        .api(`/me/joinedTeams/${params.team_id}/channels`)
        .get();

      const channels = (response.value ?? []) as Array<{
        id: string;
        displayName: string;
        description: string | null;
        membershipType: string;
      }>;

      const output = {
        team_id: params.team_id,
        channels: channels.map((ch) => ({
          id: ch.id,
          name: ch.displayName,
          description: ch.description,
          type: ch.membershipType,
        })),
      };

      const lines = channels
        .map((ch) => `- ${ch.displayName} (${ch.membershipType}) → ${ch.id}`)
        .join("\n");

      return {
        content: [
          {
            type: "text" as const,
            text: `${channels.length} canale/i trovato/i:\n${lines}`,
          },
        ],
        structuredContent: output,
      };
    }

    // List all joined teams
    const response = await client.api("/me/joinedTeams").get();

    const teams = (response.value ?? []) as Array<{
      id: string;
      displayName: string;
      description: string | null;
    }>;

    const output = {
      teams: teams.map((t) => ({
        id: t.id,
        name: t.displayName,
        description: t.description,
      })),
    };

    const lines = teams
      .map((t) => `- ${t.displayName} → ${t.id}`)
      .join("\n");

    return {
      content: [
        {
          type: "text" as const,
          text: `${teams.length} team trovato/i:\n${lines}`,
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

export async function handleSendChannelMessage(
  params: SendChannelMessageInput
): Promise<{
  content: Array<{ type: "text"; text: string }>;
  structuredContent?: Record<string, unknown>;
}> {
  try {
    const client = await getGraphClient();

    const payload: TeamsChannelMessagePayload = {
      body: {
        contentType: params.content_type,
        content: params.body,
      },
      ...(params.subject ? { subject: params.subject } : {}),
    };

    const message = await client
      .api(
        `/teams/${params.team_id}/channels/${params.channel_id}/messages`
      )
      .post(payload);

    const output = {
      status: "sent",
      message_id: message.id,
      team_id: params.team_id,
      channel_id: params.channel_id,
      web_url: message.webUrl ?? null,
    };

    let text = `Messaggio inviato con successo nel canale (ID: ${message.id})`;
    if (message.webUrl) {
      text += `\nLink: ${message.webUrl}`;
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

export async function handleSendChatMessage(
  params: SendChatMessageInput
): Promise<{
  content: Array<{ type: "text"; text: string }>;
  structuredContent?: Record<string, unknown>;
}> {
  try {
    const client = await getGraphClient();

    const payload: TeamsChatMessagePayload = {
      body: {
        contentType: params.content_type,
        content: params.body,
      },
    };

    const message = await client
      .api(`/chats/${params.chat_id}/messages`)
      .post(payload);

    const output = {
      status: "sent",
      message_id: message.id,
      chat_id: params.chat_id,
      web_url: message.webUrl ?? null,
    };

    let text = `Messaggio inviato con successo nella chat (ID: ${message.id})`;
    if (message.webUrl) {
      text += `\nLink: ${message.webUrl}`;
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

export async function handleReplyToMessage(
  params: ReplyToMessageInput
): Promise<{
  content: Array<{ type: "text"; text: string }>;
  structuredContent?: Record<string, unknown>;
}> {
  try {
    const client = await getGraphClient();

    const payload: TeamsChatMessagePayload = {
      body: {
        contentType: params.content_type,
        content: params.body,
      },
    };

    let endpoint: string;
    if (params.context === "channel") {
      endpoint = `/teams/${params.team_id}/channels/${params.channel_id}/messages/${params.message_id}/replies`;
    } else {
      endpoint = `/chats/${params.chat_id}/messages/${params.message_id}/replies`;
    }

    const reply = await client.api(endpoint).post(payload);

    const output = {
      status: "replied",
      reply_id: reply.id,
      message_id: params.message_id,
      context: params.context,
      web_url: reply.webUrl ?? null,
    };

    let text = `Risposta inviata con successo (ID: ${reply.id})`;
    if (reply.webUrl) {
      text += `\nLink: ${reply.webUrl}`;
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
