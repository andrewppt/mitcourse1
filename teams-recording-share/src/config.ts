/**
 * Central configuration, loaded from environment variables (a `.env` file is loaded by the
 * entry points via `--env-file` or `dotenv`-style loading in `loadDotEnv`).
 *
 * Two auth modes:
 *  - "app": client-credentials (daemon/webhook service). Requires CLIENT_SECRET and, for
 *    meeting/recording/attendance calls, a Teams *application access policy* granted to the organizer(s).
 *  - "delegated": device-code sign-in as the organizer (CLI / MCP on your laptop). Token cache on disk.
 */
import fs from "node:fs";
import path from "node:path";

export type AuthMode = "app" | "delegated";
export type ExternalFallback = "users-link" | "anonymous-link" | "skip";

export interface AppConfig {
  tenantId: string;
  clientId: string;
  clientSecret?: string;
  authMode: AuthMode;
  /** Where the delegated (device-code) MSAL token cache is persisted. */
  tokenCachePath: string;
  /** Delegated scopes requested at sign-in. */
  delegatedScopes: string[];
  /** Organizer user object IDs (or UPNs) whose meetings the webhook service watches. Empty = tenant-wide subscription. */
  organizerUserIds: string[];
  /** Public HTTPS base URL of the webhook service (Graph must reach it). e.g. https://xyz.ngrok.app */
  publicBaseUrl?: string;
  port: number;
  /** Secret echoed back by Graph in every notification; reject notifications that do not carry it. */
  clientState: string;
  /** Role granted on the recording file. */
  shareRole: "read" | "write";
  /** Ask OneDrive to email each recipient a sharing invitation. */
  sendInvitation: boolean;
  invitationMessage: string;
  /** What to do for attendees the `invite` action cannot add directly (typically brand-new external guests under app-only auth). */
  externalFallback: ExternalFallback;
  /** Hours until a fallback sharing link expires (0 = never). */
  fallbackLinkExpiryHours: number;
  /** Also grant access to people who were invited on the calendar event but did not join. */
  includeCalendarInvitees: boolean;
  /** Extra addresses that always get access (e.g. a shared mailbox). */
  alwaysInclude: string[];
  /** Email domains treated as internal (defaults to the organizer's domain, derived at runtime). */
  internalDomains: string[];
  /** OneDrive folder recordings land in. Teams uses "Recordings". */
  recordingsFolder: string;
  /** How long to keep retrying to locate the OneDrive file after the recording notification (minutes). */
  driveLookupTimeoutMinutes: number;
  /** Directory for the service's on-disk state (processed recordings, subscription ids). */
  stateDir: string;
  /** Graph base URL, overridable for tests. */
  graphBaseUrl: string;
}

function bool(v: string | undefined, dflt: boolean): boolean {
  if (v === undefined || v === "") return dflt;
  return ["1", "true", "yes", "on"].includes(v.toLowerCase());
}
function list(v: string | undefined): string[] {
  return (v ?? "").split(/[,;\s]+/).map((s) => s.trim()).filter(Boolean);
}
function int(v: string | undefined, dflt: number): number {
  const n = Number(v);
  return Number.isFinite(n) && v !== undefined && v !== "" ? n : dflt;
}

/** Load KEY=VALUE lines from a .env file into process.env without overriding existing values. */
export function loadDotEnv(file = path.resolve(process.cwd(), ".env")): void {
  if (!fs.existsSync(file)) return;
  for (const raw of fs.readFileSync(file, "utf8").split(/\r?\n/)) {
    const line = raw.trim();
    if (!line || line.startsWith("#")) continue;
    const eq = line.indexOf("=");
    if (eq < 0) continue;
    const key = line.slice(0, eq).trim();
    let val = line.slice(eq + 1).trim();
    if ((val.startsWith('"') && val.endsWith('"')) || (val.startsWith("'") && val.endsWith("'"))) val = val.slice(1, -1);
    if (process.env[key] === undefined) process.env[key] = val;
  }
}

export function loadConfig(env: NodeJS.ProcessEnv = process.env): AppConfig {
  const authMode = ((env.AUTH_MODE ?? "delegated").toLowerCase() as AuthMode);
  if (authMode !== "app" && authMode !== "delegated") throw new Error(`AUTH_MODE must be "app" or "delegated", got "${env.AUTH_MODE}"`);
  const required = (k: string) => {
    const v = env[k];
    if (!v) throw new Error(`Missing required environment variable ${k}`);
    return v;
  };
  const cfg: AppConfig = {
    tenantId: required("TENANT_ID"),
    clientId: required("CLIENT_ID"),
    clientSecret: env.CLIENT_SECRET || undefined,
    authMode,
    tokenCachePath: env.TOKEN_CACHE_PATH || path.resolve(env.STATE_DIR || "./data", "msal-cache.json"),
    delegatedScopes: list(env.DELEGATED_SCOPES).length
      ? list(env.DELEGATED_SCOPES)
      : ["OnlineMeetings.Read", "OnlineMeetingRecording.Read.All", "OnlineMeetingArtifact.Read.All", "Files.ReadWrite", "Calendars.Read", "Mail.Send", "User.Read", "offline_access"],
    organizerUserIds: list(env.ORGANIZER_USER_IDS),
    publicBaseUrl: env.PUBLIC_BASE_URL?.replace(/\/+$/, "") || undefined,
    port: int(env.PORT, 3978),
    clientState: env.CLIENT_STATE || "",
    shareRole: (env.SHARE_ROLE as "read" | "write") || "read",
    sendInvitation: bool(env.SEND_INVITATION, true),
    invitationMessage: env.INVITATION_MESSAGE || "Here is the recording of our meeting.",
    externalFallback: ((env.EXTERNAL_FALLBACK ?? "users-link") as ExternalFallback),
    fallbackLinkExpiryHours: int(env.FALLBACK_LINK_EXPIRY_HOURS, 0),
    includeCalendarInvitees: bool(env.INCLUDE_CALENDAR_INVITEES, false),
    alwaysInclude: list(env.ALWAYS_INCLUDE).map((e) => e.toLowerCase()),
    internalDomains: list(env.INTERNAL_DOMAINS).map((d) => d.toLowerCase()),
    recordingsFolder: env.RECORDINGS_FOLDER || "Recordings",
    driveLookupTimeoutMinutes: int(env.DRIVE_LOOKUP_TIMEOUT_MINUTES, 20),
    stateDir: env.STATE_DIR || "./data",
    graphBaseUrl: env.GRAPH_BASE_URL || "https://graph.microsoft.com",
  };
  if (cfg.authMode === "app" && !cfg.clientSecret) throw new Error("AUTH_MODE=app requires CLIENT_SECRET");
  if (!["users-link", "anonymous-link", "skip"].includes(cfg.externalFallback)) throw new Error(`EXTERNAL_FALLBACK must be users-link | anonymous-link | skip`);
  return cfg;
}
