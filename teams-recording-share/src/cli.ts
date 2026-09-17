#!/usr/bin/env node
/**
 * Small dependency-free CLI around the same pipeline the MCP server exposes.
 *
 * Commands: login | whoami | share | attendees | recordings | help
 * Unlike the MCP server, stdout here is for humans; `--json` prints the raw result instead.
 */
import { loadConfig, loadDotEnv, type AppConfig } from "./config.js";
import { createLogger, type Logger } from "./log.js";
import { login, tokenProviderFromConfig, type TokenProvider } from "./auth.js";
import { GraphClient } from "./graph/client.js";
import {
  findOnlineMeetingByJoinUrl,
  getOnlineMeeting,
  isExternal,
  listAttendees,
  listCalendarInvitees,
  mergeAttendees,
} from "./graph/meetings.js";
import { listRecordingsFolder } from "./graph/drive.js";
import { shareRecordingWithAttendees } from "./core/share.js";
import type { Attendee, DriveItem, OnlineMeeting } from "./graph/types.js";
import type { ShareRequest, ShareResult } from "./core/types.js";

const log: Logger = createLogger("cli");

function out(line = ""): void {
  process.stdout.write(`${line}\n`);
}

/* ------------------------------------------------------------------- argv */

interface Args {
  flags: Record<string, string | boolean>;
  positionals: string[];
}

const BOOLEAN_FLAGS = new Set(["dry-run", "json", "calendar", "help", "verbose"]);

function parseArgs(argv: string[]): Args {
  const flags: Record<string, string | boolean> = {};
  const positionals: string[] = [];
  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i]!;
    if (arg === "--") {
      positionals.push(...argv.slice(i + 1));
      break;
    }
    if (arg.startsWith("--")) {
      const body = arg.slice(2);
      const eq = body.indexOf("=");
      if (eq >= 0) {
        flags[body.slice(0, eq)] = body.slice(eq + 1);
        continue;
      }
      const next = argv[i + 1];
      if (BOOLEAN_FLAGS.has(body) || next === undefined || next.startsWith("--")) {
        flags[body] = true;
      } else {
        flags[body] = next;
        i++;
      }
      continue;
    }
    if (arg === "-h") {
      flags.help = true;
      continue;
    }
    positionals.push(arg);
  }
  return { flags, positionals };
}

function str(args: Args, name: string): string | undefined {
  const v = args.flags[name];
  if (v === undefined) return undefined;
  if (typeof v === "boolean") throw new Error(`--${name} requires a value`);
  return v.trim() || undefined;
}

function bool(args: Args, name: string): boolean {
  const v = args.flags[name];
  if (v === undefined) return false;
  return typeof v === "boolean" ? v : !["0", "false", "no", "off"].includes(v.toLowerCase());
}

function num(args: Args, name: string): number | undefined {
  const v = str(args, name);
  if (v === undefined) return undefined;
  const n = Number(v);
  if (!Number.isFinite(n)) throw new Error(`--${name} must be a number, got "${v}"`);
  return n;
}

function list(args: Args, name: string): string[] | undefined {
  const v = str(args, name);
  if (v === undefined) return undefined;
  return v.split(/[,;\s]+/).map((s) => s.trim()).filter(Boolean);
}

/* ------------------------------------------------------------------ context */

interface Ctx {
  cfg: AppConfig;
  tokens: TokenProvider;
  graph: GraphClient;
}

function createCtx(): Ctx {
  const cfg = loadConfig();
  const tokens = tokenProviderFromConfig(cfg, {
    allowInteractive: false,
    onMessage: (message: string) => out(message),
  });
  return { cfg, tokens, graph: new GraphClient(tokens, { baseUrl: cfg.graphBaseUrl, logger: log }) };
}

interface MeUser {
  id: string;
  displayName?: string;
  userPrincipalName?: string;
}

let meCache: MeUser | null = null;

async function getMe(ctx: Ctx): Promise<MeUser> {
  if (meCache) return meCache;
  if (ctx.cfg.authMode !== "delegated") throw new Error("`/me` is only available with AUTH_MODE=delegated.");
  meCache = await ctx.graph.get<MeUser>("/me", { query: { $select: "id,displayName,userPrincipalName" } });
  return meCache;
}

async function resolveOrganizer(ctx: Ctx, given?: string): Promise<string> {
  const value = given?.trim();
  if (value && value !== "me") return value;
  if (ctx.cfg.authMode === "delegated") {
    const me = await getMe(ctx);
    if (!me.id) throw new Error("Could not resolve the signed-in user from `/me`.");
    return me.id;
  }
  if (ctx.cfg.organizerUserIds.length === 1) return ctx.cfg.organizerUserIds[0]!;
  throw new Error("--organizer is required with AUTH_MODE=app (no signed-in user to default to).");
}

async function resolveMeeting(ctx: Ctx, userId: string, meetingId?: string, joinWebUrl?: string): Promise<OnlineMeeting> {
  if (meetingId) return getOnlineMeeting(ctx.graph, userId, meetingId);
  if (joinWebUrl) {
    const meeting = await findOnlineMeetingByJoinUrl(ctx.graph, userId, joinWebUrl);
    if (!meeting) throw new Error(`No online meeting found for join URL ${joinWebUrl}`);
    return meeting;
  }
  throw new Error("Pass --meeting <onlineMeetingId> or --join-url <url>.");
}

function internalDomainsFor(ctx: Ctx, meeting?: OnlineMeeting): string[] {
  if (ctx.cfg.internalDomains.length) return ctx.cfg.internalDomains;
  const upn = meeting?.participants?.organizer?.upn ?? meCache?.userPrincipalName ?? undefined;
  const domain = upn?.split("@")[1]?.toLowerCase();
  return domain ? [domain] : [];
}

/* ----------------------------------------------------------------- printing */

function printJson(value: unknown): void {
  out(JSON.stringify(value, null, 2));
}

function formatSize(bytes?: number): string {
  if (!bytes || bytes <= 0) return "";
  const mb = bytes / (1024 * 1024);
  return mb >= 1 ? `${mb.toFixed(1)} MB` : `${Math.round(bytes / 1024)} KB`;
}

function printAttendees(attendees: Attendee[]): void {
  for (const a of attendees) {
    const bits: string[] = [a.source];
    if (a.role) bits.push(a.role);
    if (a.isExternal) bits.push("external");
    if (typeof a.totalAttendanceInSeconds === "number") bits.push(`${Math.round(a.totalAttendanceInSeconds / 60)}m`);
    out(`  ${a.email}${a.displayName ? `  (${a.displayName})` : ""}  [${bits.join(", ")}]`);
  }
}

function printShareResult(result: ShareResult): void {
  out(`${result.dryRun ? "DRY RUN — " : ""}Meeting: ${result.meeting.subject ?? "(no subject)"}`);
  out(`  id: ${result.meeting.id}`);
  out(`  organizer: ${result.meeting.organizerUserId}`);
  out(`Attendees resolved: ${result.attendees.length} (${result.attendees.filter((a) => a.isExternal).length} external)`);
  for (const rec of result.recordings) {
    out();
    out(`Recording ${rec.recordingId}${rec.createdDateTime ? ` (${rec.createdDateTime})` : ""}`);
    out(`  file: ${rec.driveItem ? `${rec.driveItem.name} — ${rec.driveItem.webUrl ?? rec.driveItem.id}` : "(not found)"}`);
    out(`  granted (${rec.granted.length}): ${rec.granted.map((p) => p.email).join(", ") || "-"}`);
    out(`  already had access (${rec.alreadyHadAccess.length}): ${rec.alreadyHadAccess.map((p) => p.email).join(", ") || "-"}`);
    if (rec.fallbackLink) {
      out(`  fallback link (${rec.fallbackLink.scope}): ${rec.fallbackLink.webUrl}`);
      if (rec.fallbackLink.sentTo.length) out(`    sent to: ${rec.fallbackLink.sentTo.join(", ")}`);
      if (rec.fallbackLink.expirationDateTime) out(`    expires: ${rec.fallbackLink.expirationDateTime}`);
    }
    for (const s of rec.skipped) out(`  skipped ${s.email}: ${s.reason}`);
    for (const e of rec.errors) out(`  error${e.email ? ` for ${e.email}` : ""}: ${e.message}`);
  }
  for (const w of result.warnings) out(`warning: ${w}`);
}

function printRecordings(items: DriveItem[], folder: string): void {
  out(`${items.length} item(s) in /${folder}`);
  for (const item of items) {
    out(`  ${item.createdDateTime ?? "?"}  ${item.name}${formatSize(item.size) ? `  (${formatSize(item.size)})` : ""}`);
    out(`    id: ${item.id}${item.webUrl ? `  ${item.webUrl}` : ""}`);
  }
}

/* ----------------------------------------------------------------- commands */

const HELP = `teams-recording-share — share Teams meeting recordings with everyone who attended

Usage:
  tsx src/cli.ts <command> [options]

Commands:
  login                         Sign in (browser window by default) and persist the token cache
  whoami                        Show the auth mode and the signed-in account
  share                         Grant attendees read access to a meeting's recording
  attendees                     List the people who attended a meeting
  recordings                    List recent files in the organizer's OneDrive Recordings folder
  help                          Show this help

share options:
  --meeting <onlineMeetingId>   Graph onlineMeeting id (MSp...)      (or --join-url)
  --join-url <url>              Teams join URL of the meeting
  --organizer <id|upn>          Organizer (default: the signed-in user in delegated mode)
  --recording <id>              Only this callRecording (default: all recordings of the meeting)
  --drive-item <id>             Explicit OneDrive item id (skips the file search)
  --extra a@b.com,c@d.com       Extra recipients beyond the attendees
  --dry-run                     Compute the plan without granting anything
  --json                        Print the raw ShareResult as JSON

attendees options:
  --meeting <id> | --join-url <url>   Meeting to inspect
  --organizer <id|upn>                Organizer (default: signed-in user)
  --calendar                          Also include calendar invitees who did not join
  --json                              Print JSON

recordings options:
  --organizer <id|upn>          Organizer (default: signed-in user)
  --top <n>                     How many items to list (default 25)
  --json                        Print JSON

Environment: TENANT_ID, CLIENT_ID, AUTH_MODE=delegated|app, CLIENT_SECRET (app mode), ... (see .env.example)`;

async function cmdLogin(_args: Args): Promise<void> {
  const cfg = loadConfig();
  if (cfg.authMode !== "delegated") {
    out("AUTH_MODE=app uses client credentials — no interactive sign-in is needed.");
    return;
  }
  const result = await login(cfg, log);
  out(`Signed in as ${result.username ?? "(unknown account)"}`);
  out(`Token cache: ${cfg.tokenCachePath}`);
}

async function cmdWhoami(args: Args): Promise<void> {
  const ctx = createCtx();
  const account = ctx.tokens.account ? await ctx.tokens.account() : undefined;
  const user = ctx.cfg.authMode === "delegated" ? await getMe(ctx) : null;
  const payload = {
    authMode: ctx.cfg.authMode,
    tenantId: ctx.cfg.tenantId,
    clientId: ctx.cfg.clientId,
    account: account ?? null,
    user,
    organizerUserIds: ctx.cfg.organizerUserIds,
    tokenCachePath: ctx.cfg.tokenCachePath,
  };
  if (bool(args, "json")) return printJson(payload);
  out(`Auth mode: ${payload.authMode}`);
  out(`Tenant:    ${payload.tenantId}`);
  out(`Client:    ${payload.clientId}`);
  if (account) out(`Account:   ${account}`);
  if (user) out(`User:      ${user.displayName ?? "?"} <${user.userPrincipalName ?? "?"}> (${user.id})`);
  if (!user && ctx.cfg.organizerUserIds.length) out(`Organizers: ${ctx.cfg.organizerUserIds.join(", ")}`);
}

async function cmdShare(args: Args): Promise<void> {
  const meetingId = str(args, "meeting");
  const joinWebUrl = str(args, "join-url");
  if (!meetingId && !joinWebUrl) throw new Error("share requires --meeting <id> or --join-url <url>");
  const ctx = createCtx();
  const organizerUserId = await resolveOrganizer(ctx, str(args, "organizer"));
  const req: ShareRequest = {
    organizerUserId,
    meetingId,
    joinWebUrl,
    recordingId: str(args, "recording"),
    driveItemId: str(args, "drive-item"),
    extraEmails: list(args, "extra"),
    dryRun: bool(args, "dry-run"),
  };
  const result = await shareRecordingWithAttendees(ctx.graph, ctx.cfg, req, log);
  if (bool(args, "json")) return printJson(result);
  printShareResult(result);
}

async function cmdAttendees(args: Args): Promise<void> {
  const ctx = createCtx();
  const organizerUserId = await resolveOrganizer(ctx, str(args, "organizer"));
  const meeting = await resolveMeeting(ctx, organizerUserId, str(args, "meeting"), str(args, "join-url"));
  const attended = await listAttendees(ctx.graph, organizerUserId, meeting.id, log);
  const wantCalendar = bool(args, "calendar") || ctx.cfg.includeCalendarInvitees;
  const invitees = wantCalendar ? await listCalendarInvitees(ctx.graph, organizerUserId, meeting, log) : [];
  const domains = internalDomainsFor(ctx, meeting);
  const attendees = mergeAttendees(attended, invitees).map((a) => ({
    ...a,
    isExternal: a.isExternal ?? (domains.length ? isExternal(a.email, domains) : undefined),
  }));
  if (bool(args, "json")) {
    return printJson({
      meeting: { id: meeting.id, subject: meeting.subject ?? null, joinWebUrl: meeting.joinWebUrl, organizerUserId },
      includeCalendarInvitees: wantCalendar,
      internalDomains: domains,
      attendees,
    });
  }
  out(`Meeting: ${meeting.subject ?? "(no subject)"} (${meeting.id})`);
  out(`${attendees.length} person(s) — ${attended.length} attended, ${invitees.length} calendar invitee(s)`);
  printAttendees(attendees);
}

async function cmdRecordings(args: Args): Promise<void> {
  const ctx = createCtx();
  const organizerUserId = await resolveOrganizer(ctx, str(args, "organizer"));
  const top = num(args, "top") ?? 25;
  const items = await listRecordingsFolder(ctx.graph, organizerUserId, ctx.cfg.recordingsFolder, top);
  if (bool(args, "json")) return printJson({ organizerUserId, folder: ctx.cfg.recordingsFolder, items });
  printRecordings(items, ctx.cfg.recordingsFolder);
}

/* --------------------------------------------------------------------- main */

async function main(): Promise<void> {
  loadDotEnv();
  const [command, ...rest] = process.argv.slice(2);
  const args = parseArgs(rest);
  if (!command || command === "help" || command === "--help" || command === "-h" || bool(args, "help")) {
    out(HELP);
    return;
  }
  switch (command) {
    case "login":
      return cmdLogin(args);
    case "whoami":
      return cmdWhoami(args);
    case "share":
      return cmdShare(args);
    case "attendees":
      return cmdAttendees(args);
    case "recordings":
      return cmdRecordings(args);
    default:
      throw new Error(`Unknown command "${command}". Run \`help\` to see the available commands.`);
  }
}

main().catch((err: unknown) => {
  const e = err as Error & { status?: number; code?: string };
  const detail = [e?.status ? `HTTP ${e.status}` : "", e?.code ?? ""].filter(Boolean).join(" ");
  process.stderr.write(`Error: ${e?.message ?? String(err)}${detail ? ` (${detail})` : ""}\n`);
  const msg = String(e?.message);
  if (/Missing required environment variable|AUTH_MODE|EXTERNAL_FALLBACK/i.test(msg)) {
    process.stderr.write(
      "Hint: set TENANT_ID, CLIENT_ID (and CLIENT_SECRET when AUTH_MODE=app) in your environment or a .env file — see .env.example.\n",
    );
  } else if (/token|login|sign[ -]?in|AADSTS|interaction_required/i.test(msg)) {
    process.stderr.write("Hint: run `npm run login` to (re)authenticate.\n");
  }
  process.exitCode = 1;
});
