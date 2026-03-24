import { z } from "zod";
import { getGraphClient } from "../graphClient.js";
import { handleGraphError } from "../utils/errors.js";

export const SearchMailInputSchema = z
  .object({
    mailbox: z
      .string()
      .email()
      .describe("Indirizzo email della casella delegata/condivisa (es. df@dfsolutions.it)"),
    query: z
      .string()
      .optional()
      .describe("Testo da cercare nel soggetto e corpo della mail"),
    sender: z
      .string()
      .optional()
      .describe("Filtra per mittente (email completa o parziale)"),
    subject: z
      .string()
      .optional()
      .describe("Filtra per oggetto della mail"),
    folder: z
      .string()
      .optional()
      .describe("Nome cartella (es. 'Inbox', 'Sent Items'). Default: tutte."),
    after: z
      .string()
      .optional()
      .describe("Mail ricevute dopo questa data (ISO 8601, es. '2026-03-01')"),
    before: z
      .string()
      .optional()
      .describe("Mail ricevute prima di questa data"),
    has_attachments: z.boolean().optional().describe("Filtra solo mail con allegati"),
    is_read: z.boolean().optional().describe("Filtra per stato lettura"),
    limit: z.number().min(1).max(50).default(10).describe("Max risultati (default 10, max 50)"),
  })
  .strict();

export type SearchMailInput = z.infer<typeof SearchMailInputSchema>;

export async function handleSearchMail(params: SearchMailInput): Promise<{
  content: Array<{ type: "text"; text: string }>;
  structuredContent?: Record<string, unknown>;
}> {
  try {
    const client = await getGraphClient();
    const basePath = `/users/${params.mailbox}`;

    let messagesEndpoint = `${basePath}/messages`;
    if (params.folder) {
      const foldersResp = await client
        .api(`${basePath}/mailFolders`)
        .select("id,displayName")
        .top(50)
        .get();
      const folders = foldersResp.value as Array<{ id: string; displayName: string }>;
      const match = folders.find(
        (f) => f.displayName.toLowerCase() === params.folder!.toLowerCase()
      );
      if (match) {
        messagesEndpoint = `${basePath}/mailFolders/${match.id}/messages`;
      } else {
        return {
          content: [{
            type: "text" as const,
            text: `Cartella "${params.folder}" non trovata in ${params.mailbox}. Disponibili: ${folders.map((f) => f.displayName).join(", ")}`,
          }],
        };
      }
    }

    // Build OData filters only for supported properties
    const filters: string[] = [];
    if (params.after) {
      filters.push(`receivedDateTime ge ${new Date(params.after).toISOString()}`);
    }
    if (params.before) {
      filters.push(`receivedDateTime lt ${new Date(params.before).toISOString()}`);
    }
    if (params.has_attachments !== undefined) {
      filters.push(`hasAttachments eq ${params.has_attachments}`);
    }
    if (params.is_read !== undefined) {
      filters.push(`isRead eq ${params.is_read}`);
    }

    // Determine if we need text-based searching
    const hasTextSearch = !!(params.sender || params.subject || params.query);

    // Build KQL search query for text-based filtering
    let kqlParts: string[] = [];
    if (params.sender) {
      kqlParts.push(`from:${params.sender}`);
    }
    if (params.subject) {
      kqlParts.push(`subject:${params.subject}`);
    }
    if (params.query) {
      kqlParts.push(params.query);
    }

    const fetchLimit = hasTextSearch ? 50 : params.limit;
    let messages: Array<any> = [];

    if (hasTextSearch && kqlParts.length > 0) {
      // Try $search with KQL syntax first (requires ConsistencyLevel: eventual)
      try {
        const kqlQuery = kqlParts.join(" AND ");
        let searchRequest = client
          .api(messagesEndpoint)
          .header("ConsistencyLevel", "eventual")
          .search(`"${kqlQuery}"`)
          .select("id,subject,from,toRecipients,receivedDateTime,bodyPreview,hasAttachments,isRead,importance,webLink")
          .top(fetchLimit);

        if (filters.length > 0) {
          searchRequest = searchRequest.filter(filters.join(" and "));
        }

        const searchResponse = await searchRequest.get();
        messages = searchResponse.value ?? [];
      } catch (_searchError) {
        // $search failed, fall back to pagination with client-side filtering
        const maxPages = 10; // fetch up to 500 results
        const pageSize = 50;
        let nextLink: string | undefined = undefined;

        for (let page = 0; page < maxPages; page++) {
          try {
            let pageResponse: any;
            if (page === 0) {
              let pageRequest = client
                .api(messagesEndpoint)
                .select("id,subject,from,toRecipients,receivedDateTime,bodyPreview,hasAttachments,isRead,importance,webLink")
                .orderby("receivedDateTime desc")
                .top(pageSize);
              if (filters.length > 0) {
                pageRequest = pageRequest.filter(filters.join(" and "));
              }
              pageResponse = await pageRequest.get();
            } else if (nextLink) {
              pageResponse = await client.api(nextLink).get();
            } else {
              break;
            }
            const pageMessages = pageResponse.value ?? [];
            messages.push(...pageMessages);
            nextLink = pageResponse["@odata.nextLink"];
            if (!nextLink || pageMessages.length < pageSize) break;
          } catch (_pageError) {
            break;
          }
        }

        // Apply client-side filtering
        if (params.sender) {
          const senderLower = params.sender.toLowerCase();
          messages = messages.filter((m: any) => {
            const addr = m.from?.emailAddress?.address?.toLowerCase() ?? "";
            const name = m.from?.emailAddress?.name?.toLowerCase() ?? "";
            return addr.includes(senderLower) || name.includes(senderLower);
          });
        }
        if (params.subject) {
          const subjectLower = params.subject.toLowerCase();
          messages = messages.filter((m: any) =>
            (m.subject ?? "").toLowerCase().includes(subjectLower)
          );
        }
        if (params.query) {
          const queryLower = params.query.toLowerCase();
          messages = messages.filter((m: any) => {
            const subj = (m.subject ?? "").toLowerCase();
            const preview = (m.bodyPreview ?? "").toLowerCase();
            const addr = m.from?.emailAddress?.address?.toLowerCase() ?? "";
            const name = m.from?.emailAddress?.name?.toLowerCase() ?? "";
            return subj.includes(queryLower) || preview.includes(queryLower) || addr.includes(queryLower) || name.includes(queryLower);
          });
        }
      }
    } else {
      // No text search needed, just use OData filters
      let request = client
        .api(messagesEndpoint)
        .select("id,subject,from,toRecipients,receivedDateTime,bodyPreview,hasAttachments,isRead,importance,webLink")
        .orderby("receivedDateTime desc")
        .top(params.limit);

      if (filters.length > 0) {
        request = request.filter(filters.join(" and "));
      }

      const response = await request.get();
      messages = response.value ?? [];
    }
    // Trim to requested limit
    messages = messages.slice(0, params.limit);



    if (!messages || messages.length === 0) {
      return {
        content: [{
          type: "text" as const,
          text: `Nessuna mail trovata nella casella ${params.mailbox} con i criteri specificati.`,
        }],
        structuredContent: { results: [], count: 0, mailbox: params.mailbox },
      };
    }

    const results = messages.map((m) => ({
      id: m.id,
      subject: m.subject ?? "(nessun oggetto)",
      sender: m.from?.emailAddress?.address ?? "sconosciuto",
      sender_name: m.from?.emailAddress?.name ?? "",
      to: m.toRecipients?.map((r: any) => r.emailAddress.address).join(", ") ?? "",
      date: m.receivedDateTime,
      preview: m.bodyPreview?.substring(0, 200) ?? "",
      has_attachments: m.hasAttachments,
      is_read: m.isRead,
      importance: m.importance,
      web_link: m.webLink,
    }));

    const summary = results
      .map(
        (r, i) =>
          `${i + 1}. [${r.date.substring(0, 16)}] Da: ${r.sender_name || r.sender}\n   Oggetto: ${r.subject}\n   ${r.preview.substring(0, 120)}...`
      )
      .join("\n\n");

    return {
      content: [{
        type: "text" as const,
        text: `Trovate ${results.length} mail nella casella ${params.mailbox}:\n\n${summary}`,
      }],
      structuredContent: { results, count: results.length, mailbox: params.mailbox },
    };
  } catch (error) {
    return {
      content: [{ type: "text" as const, text: handleGraphError(error) }],
    };
  }
}
