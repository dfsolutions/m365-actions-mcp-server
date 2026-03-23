import {
  ConfidentialClientApplication,
  AuthorizationCodeRequest,
  AuthorizationUrlRequest,
  Configuration,
  CryptoProvider,
} from "@azure/msal-node";
import http from "http";
import open from "open";
import fs from "fs";
import path from "path";
import {
  M365_CLIENT_ID,
  M365_CLIENT_SECRET,
  M365_TENANT_ID,
  M365_REDIRECT_URI,
  GRAPH_SCOPES,
  TOKEN_CACHE_PATH,
} from "./constants.js";

const msalConfig: Configuration = {
  auth: {
    clientId: M365_CLIENT_ID,
    authority: `https://login.microsoftonline.com/${M365_TENANT_ID}`,
    clientSecret: M365_CLIENT_SECRET,
  },
  cache: {
    // We will handle persistence manually
  },
};

let msalClient: ConfidentialClientApplication;

function getMsalClient(): ConfidentialClientApplication {
  if (!msalClient) {
    msalClient = new ConfidentialClientApplication(msalConfig);

    // Restore cache from disk if available
    const cachePath = getTokenCachePath();
    if (fs.existsSync(cachePath)) {
      try {
        const cacheData = fs.readFileSync(cachePath, "utf-8");
        msalClient.getTokenCache().deserialize(cacheData);
        console.error("[auth] Cache token ripristinata da disco.");
      } catch {
        console.error("[auth] Cache token corrotta, verrà rigenerata.");
      }
    }
  }
  return msalClient;
}

function getTokenCachePath(): string {
  // Store in user home directory for persistence across runs
  const homeDir = process.env.HOME ?? process.env.USERPROFILE ?? ".";
  return path.join(homeDir, TOKEN_CACHE_PATH);
}

function persistCache(): void {
  try {
    const client = getMsalClient();
    const cacheData = client.getTokenCache().serialize();
    fs.writeFileSync(getTokenCachePath(), cacheData, "utf-8");
    console.error("[auth] Cache token salvata su disco.");
  } catch (err) {
    console.error("[auth] Impossibile salvare cache:", err);
  }
}

/**
 * Try to get a token silently from the cache.
 * Returns null if no cached account is found.
 */
async function acquireTokenSilent(): Promise<string | null> {
  const client = getMsalClient();
  const cache = client.getTokenCache();
  const accounts = await cache.getAllAccounts();

  if (accounts.length === 0) {
    return null;
  }

  try {
    const result = await client.acquireTokenSilent({
      account: accounts[0],
      scopes: GRAPH_SCOPES,
    });

    if (result?.accessToken) {
      persistCache();
      return result.accessToken;
    }
  } catch {
    console.error("[auth] Silent token acquisition fallita, serve login interattivo.");
  }

  return null;
}

/**
 * Interactive login via browser.
 * Opens the default browser, starts a local HTTP server to catch the redirect,
 * exchanges the authorization code for tokens, and persists the cache.
 */
async function acquireTokenInteractive(): Promise<string> {
  const client = getMsalClient();
  const cryptoProvider = new CryptoProvider();

  const { verifier, challenge } = await cryptoProvider.generatePkceCodes();

  const authCodeUrlParams: AuthorizationUrlRequest = {
    scopes: GRAPH_SCOPES,
    redirectUri: M365_REDIRECT_URI,
    codeChallenge: challenge,
    codeChallengeMethod: "S256",
  };

  const authUrl = await client.getAuthCodeUrl(authCodeUrlParams);

  // Extract port from redirect URI
  const redirectUrl = new URL(M365_REDIRECT_URI);
  const port = parseInt(redirectUrl.port || "3939", 10);

  return new Promise<string>((resolve, reject) => {
    const server = http.createServer(async (req, res) => {
      if (!req.url?.startsWith(redirectUrl.pathname)) {
        res.writeHead(404);
        res.end();
        return;
      }

      const url = new URL(req.url, `http://localhost:${port}`);
      const code = url.searchParams.get("code");

      if (!code) {
        const errorMsg = url.searchParams.get("error_description") ?? "Nessun codice ricevuto";
        res.writeHead(400, { "Content-Type": "text/html; charset=utf-8" });
        res.end(`<h2>Errore di autenticazione</h2><p>${errorMsg}</p>`);
        server.close();
        reject(new Error(errorMsg));
        return;
      }

      try {
        const tokenRequest: AuthorizationCodeRequest = {
          code,
          scopes: GRAPH_SCOPES,
          redirectUri: M365_REDIRECT_URI,
          codeVerifier: verifier,
        };

        const result = await client.acquireTokenByCode(tokenRequest);

        if (!result?.accessToken) {
          throw new Error("Nessun access token nella risposta");
        }

        persistCache();

        res.writeHead(200, { "Content-Type": "text/html; charset=utf-8" });
        res.end(
          "<h2>Login riuscito!</h2><p>Puoi chiudere questa finestra e tornare al terminale.</p>"
        );
        server.close();
        resolve(result.accessToken);
      } catch (err) {
        res.writeHead(500, { "Content-Type": "text/html; charset=utf-8" });
        res.end(`<h2>Errore</h2><p>${String(err)}</p>`);
        server.close();
        reject(err);
      }
    });

    server.listen(port, () => {
      console.error(`[auth] Server di callback in ascolto su porta ${port}`);
      console.error("[auth] Apro il browser per il login Microsoft...");
      open(authUrl).catch(() => {
        console.error(`[auth] Impossibile aprire il browser. Vai manualmente a:\n${authUrl}`);
      });
    });

    // Timeout after 5 minutes
    setTimeout(() => {
      server.close();
      reject(new Error("Timeout di autenticazione (5 minuti). Riavvia il server per riprovare."));
    }, 5 * 60 * 1000);
  });
}

/**
 * Main entry point: get a valid access token.
 * Tries silent first, falls back to interactive login.
 */
export async function getAccessToken(): Promise<string> {
  // Validate env vars
  if (!M365_CLIENT_ID || !M365_TENANT_ID || !M365_CLIENT_SECRET) {
    throw new Error(
      "Variabili d'ambiente mancanti: M365_CLIENT_ID, M365_TENANT_ID e M365_CLIENT_SECRET sono obbligatorie. " +
      "Copia .env.example in .env e compilalo con i dati della tua app Azure."
    );
  }

  // Try silent first
  const cachedToken = await acquireTokenSilent();
  if (cachedToken) {
    console.error("[auth] Token ottenuto dalla cache.");
    return cachedToken;
  }

  // Fall back to interactive
  console.error("[auth] Nessun token in cache, avvio login interattivo...");
  return acquireTokenInteractive();
}
