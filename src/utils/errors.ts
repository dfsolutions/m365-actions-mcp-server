import type { GraphErrorResponse } from "../types.js";

/**
 * Centralised error handler for Microsoft Graph API errors.
 * Returns a user-friendly, actionable error message.
 */
export function handleGraphError(error: unknown): string {
  if (error instanceof Error) {
    // Graph client errors include statusCode on the error object
    const graphErr = error as Error & {
      statusCode?: number;
      body?: string;
    };

    if (graphErr.statusCode) {
      let detail = "";
      if (graphErr.body) {
        try {
          const parsed: GraphErrorResponse = JSON.parse(graphErr.body);
          detail = parsed.error?.message ?? "";
        } catch {
          detail = graphErr.body;
        }
      }

      switch (graphErr.statusCode) {
        case 400:
          return `Errore 400 – Richiesta non valida. ${detail || "Controlla i parametri inviati."}`;
        case 401:
          return "Errore 401 – Token scaduto o non valido. Riavvia il server per ripetere il login.";
        case 403:
          return `Errore 403 – Permessi insufficienti. ${detail || "Verifica che l'app Azure abbia i permessi Mail.Send e Calendars.ReadWrite."}`;
        case 404:
          return `Errore 404 – Risorsa non trovata. ${detail || "Controlla l'ID fornito."}`;
        case 429:
          return "Errore 429 – Troppe richieste. Attendi qualche secondo e riprova.";
        default:
          return `Errore ${graphErr.statusCode} – ${detail || graphErr.message}`;
      }
    }

    return `Errore: ${graphErr.message}`;
  }

  return `Errore imprevisto: ${String(error)}`;
}
