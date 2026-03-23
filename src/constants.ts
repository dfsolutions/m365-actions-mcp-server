import dotenv from "dotenv";
dotenv.config();

export const M365_CLIENT_ID = process.env.M365_CLIENT_ID ?? "";
export const M365_CLIENT_SECRET = process.env.M365_CLIENT_SECRET ?? "";
export const M365_TENANT_ID = process.env.M365_TENANT_ID ?? "";
export const M365_REDIRECT_URI =
  process.env.M365_REDIRECT_URI ?? "http://localhost:3939/auth/callback";
export const M365_USER_EMAIL = process.env.M365_USER_EMAIL ?? "";

export const GRAPH_SCOPES = [
  "Mail.Send",
  "Mail.Read",
  "Calendars.ReadWrite",
  "User.Read",
  "offline_access",
];

export const GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0";

export const TOKEN_CACHE_PATH = ".m365-actions-tokens.json";

export const CHARACTER_LIMIT = 25000;

export const DEFAULT_TIMEZONE = "Europe/Rome";
