import { z } from "zod";
import fs from "fs";
import path from "path";
import { getGraphClient } from "../graphClient.js";
import { handleGraphError } from "../utils/errors.js";
import type { SendMailPayload, MailRecipient } from "../types.js";

// ── Schemas ──────────────────────────────────────────────

export const SendMailInputSchema = z
  .object({
    to: z
      .union([z.string().email(), z.array(z.string().email()).min(1)])
      .describe("Destinatario/i: un indirizzo email o un array di indirizzi"),
    subject: z
      .string()
      .min(1, "L'oggetto non può essere vuoto")
      .max(998, "L'oggetto non può superare 998 caratteri")
      .describe("Oggetto della mail"),
    body: z
      .string()
      .min(1, "Il corpo della mail non può essere vuoto")
      .describe("Contenuto della mail (HTML o testo semplice)"),
    cc: z
      .union([z.string().email(), z.array(z.string().email())])
      .optional()
      .describe("Destinatario/i in copia (opzionale)"),
    content_type: z
      .enum(["HTML", "Text"])
      .default("HTML")
      .describe("Tipo di contenuto: HTML (default) o Text"),
  })
  .strict();

export type SendMailInput = z.infer<typeof SendMailInputSchema>;

export const ReplyMailInputSchema = z
  .object({
    message_id: z
      .string()
      .min(1, "L'ID del messaggio è obbligatorio")
      .describe("ID del messaggio a cui rispondere (ottenibile da outlook_email_search)"),
    body: z
      .string()
      .min(1, "Il corpo della risposta non può essere vuoto")
      .describe("Contenuto della risposta (HTML o testo semplice)"),
    reply_all: z
      .boolean()
      .default(false)
      .describe("Se true, risponde a tutti i destinatari (default: false)"),
  })
  .strict();

export type ReplyMailInput = z.infer<typeof ReplyMailInputSchema>;

export const GetAttachmentsInputSchema = z
  .object({
    message_id: z
      .string()
      .min(1, "L'ID del messaggio è obbligatorio")
      .describe("ID del messaggio di cui scaricare gli allegati (ottenibile da outlook_email_search)"),
    save_to: z
      .string()
      .optional()
      .describe("Percorso cartella dove salvare gli allegati (opzionale, default: cartella corrente)"),
  })
  .strict();

export type GetAttachmentsInput = z.infer<typeof GetAttachmentsInputSchema>;

// ── Helpers ──────────────────────────────────────────────

function toRecipients(emails: string | string[]): MailRecipient[] {
  const list = Array.isArray(emails) ? emails : [emails];
  return list.map((addr) => ({ emailAddress: { address: addr } }));
}

// ── Handlers ─────────────────────────────────────────────

export async function handleSendMail(params: SendMailInput): Promise<{
  content: Array<{ type: "text"; text: string }>;
  structuredContent?: Record<string, unknown>;
}> {
  try {
    const client = await getGraphClient();

    const payload: SendMailPayload = {
      message: {
        subject: params.subject,
        body: {
          contentType: params.content_type,
          content: params.body,
        },
        toRecipients: toRecipients(params.to),
        ...(params.cc ? { ccRecipients: toRecipients(params.cc) } : {}),
      },
    };

    await client.api("/me/sendMail").post(payload);

    const recipients = Array.isArray(params.to) ? params.to.join(", ") : params.to;
    const output = {
      status: "sent",
      to: recipients,
      subject: params.subject,
    };

    return {
      content: [
        {
          type: "text" as const,
          text: `Mail inviata con successo a ${recipients} — oggetto: "${params.subject}"`,
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

export async function handleReplyMail(params: ReplyMailInput): Promise<{
  content: Array<{ type: "text"; text: string }>;
  structuredContent?: Record<string, unknown>;
}> {
  try {
    const client = await getGraphClient();

    const endpoint = params.reply_all
      ? `/me/messages/${params.message_id}/replyAll`
      : `/me/messages/${params.message_id}/reply`;

    await client.api(endpoint).post({
      comment: params.body,
    });

    const action = params.reply_all ? "Risposta a tutti" : "Risposta";
    const output = {
      status: "replied",
      message_id: params.message_id,
      reply_all: params.reply_all,
    };

    return {
      content: [
        {
          type: "text" as const,
          text: `${action} inviata con successo al messaggio ${params.message_id}`,
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

export async function handleGetAttachments(params: GetAttachmentsInput): Promise<{
  content: Array<{ type: "text"; text: string }>;
  structuredContent?: Record<string, unknown>;
}> {
  try {
    const client = await getGraphClient();

    const response = await client
      .api(`/me/messages/${params.message_id}/attachments`)
      .get();

    const attachments = response.value as Array<{
      id: string;
      name: string;
      contentType: string;
      size: number;
      contentBytes?: string;
      "@odata.type": string;
    }>;

    if (!attachments || attachments.length === 0) {
      return {
        content: [
          {
            type: "text" as const,
            text: "Nessun allegato trovato per questo messaggio.",
          },
        ],
      };
    }

    const saveDir = params.save_to ?? ".";
    if (!fs.existsSync(saveDir)) {
      fs.mkdirSync(saveDir, { recursive: true });
    }

    const savedFiles: Array<{ name: string; path: string; size: number; contentType: string }> = [];

    for (const att of attachments) {
      if (att["@odata.type"] === "#microsoft.graph.fileAttachment" && att.contentBytes) {
        const filePath = path.join(saveDir, att.name);
        const buffer = Buffer.from(att.contentBytes, "base64");
        fs.writeFileSync(filePath, buffer);
        savedFiles.push({
          name: att.name,
          path: filePath,
          size: att.size,
          contentType: att.contentType,
        });
      }
    }

    if (savedFiles.length === 0) {
      return {
        content: [
          {
            type: "text" as const,
            text: `Trovati ${attachments.length} allegati ma nessuno è un file scaricabile (potrebbero essere riferimenti o elementi inline).`,
          },
        ],
      };
    }

    const output = {
      status: "downloaded",
      message_id: params.message_id,
      attachments: savedFiles,
    };

    const fileList = savedFiles
      .map((f) => `- ${f.name} (${(f.size / 1024).toFixed(1)} KB) → ${f.path}`)
      .join("\n");

    return {
      content: [
        {
          type: "text" as const,
          text: `${savedFiles.length} allegato/i scaricato/i:\n${fileList}`,
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
