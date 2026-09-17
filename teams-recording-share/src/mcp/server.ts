#!/usr/bin/env node
/**
 * MCP (stdio) server exposing the Teams recording sharing pipeline as tools.
 *
 * stdout is the JSON-RPC protocol channel, so **nothing** may be written to it apart from
 * protocol messages: all logging goes to stderr through `createLogger("mcp")`.
 *
 * The Graph client is created lazily on the first tool call so the server still starts (and can
 * answer `resources/list`, `tools/list`, ...) when the environment is not configured yet; tools
 * then return a helpful `isError` result telling the user what to set up.
 */
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import type { CallToolResult } from "@modelcontextprotocol/sdk/types.js";
import { z } from "zod";

import { loadConfig, loadDotEnv, type AppConfig } from "../config.js";
import { createLogger } from "../log.js";
import { tokenProviderFromConfig, type TokenProvider } from "../auth.js";
import { GraphClient } from "../graph/client.js";
import {
  findOnlineMeetingByJoinUrl,
  getOnlineMeeting,
  isExternal,
  listAttendees,
  listCalendarInvitees,
  listRecordings,
  mergeAttendees,
} from "../graph/meetings.js";
import { emailsWithAccess, listItemPermissions, listRecordingsFolder, pickRecordingItem } from "../graph/drive.js";
import { shareRecordingWithAttendees } from "../core/share.js";
import type { Attendee, DriveItem, OnlineMeeting } from "../graph/types.js";
import type { ShareRequest } from "../core/types.js";

const log = createLogger("mcp");

/* ------------------------------------------------------------------ context */

interface Ctx {
  cfg: AppConfig;
  tokens: TokenProvider;
  graph: GraphClient;
}

let ctx: Ctx | null = null;

const SETUP_HINT =
  "Set TENANT_ID and CLIENT_ID (plus CLIENT_SECRET when AUTH_MODE=app) in the environment or in a .env " +
  "file next to the server, then run `npm run login` to sign in (delegated mode).";

/** Build (once) the config + Graph client. Throws a user-facing error when configuration is missing. */
function getCtx(): Ctx {
  if (ctx) return ctx;
  let cfg: AppConfig;
  try {
    cfg = loadConfig();
  } catch (err) {
    throw new Error(`Configuration error: ${errorText(err)}. ${SETUP_HINT}`);
  }
  const tokens = tokenProviderFromConfig(cfg, {
    allowInteractive: false,
    onMessage: (message: string) => log.info("sign-in", { message }),
  });
  const graph = new GraphClient(tokens, { baseUrl: cfg.graphBaseUrl, logger: log });
  ctx = { cfg, tokens, graph };
  log.info("graph client ready", { authMode: cfg.authMode, baseUrl: cfg.graphBaseUrl });
  return ctx;
}

/* ------------------------------------------------------------------ results */

function errorText(err: unknown): string {
  if (err instanceof Error) {
    const e = err as Error & { status?: number; code?: string };
    const bits: string[] = [];
    if (typeof e.status === "number") bits.push(`HTTP ${e.status}`);
    if (e.code) bits.push(e.code);
    return bits.length ? `${err.message} (${bits.join(" ")})` : err.message;
  }
  return typeof err === "string" ? err : JSON.stringify(err);
}

function ok(result: unknown): CallToolResult {
  return {
    content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
    structuredContent: result as Record<string, unknown>,
  };
}

function fail(err: unknown): CallToolResult {
  let message = errorText(err);
  if (/token|login|sign[ -]?in|AADSTS|interaction_required|credential|cache/i.test(message) && !/npm run login/.test(message)) {
    message += " — run `npm run login` to (re)authenticate, or check your environment variables.";
  }
  log.error("tool failed", { error: message });
  return { isError: true, content: [{ type: "text", text: message }] };
}

/** Run a tool body, converting any throw into an MCP error result. */
async function run(name: string, fn: (c: Ctx) => Promise<unknown>): Promise<CallToolResult> {
  const started = Date.now();
  try {
    const c = getCtx();
    const result = await fn(c);
    log.info("tool ok", { tool: name, ms: Date.now() - started });
    return ok(result);
  } catch (err) {
    log.warn("tool error", { tool: name, ms: Date.now() - started });
    return fail(err);
  }
}

/* ------------------------------------------------------------------ helpers */

interface MeUser {
  id: string;
  displayName?: string;
  userPrincipalName?: string;
}

let meCache: MeUser | null = null;

async function getMe(c: Ctx): Promise<MeUser> {
  if (meCache) return meCache;
  if (c.cfg.authMode !== "delegated") throw new Error("`/me` is only available with AUTH_MODE=delegated.");
  meCache = await c.graph.get<MeUser>("/me", { query: { $select: "id,displayName,userPrincipalName" } });
  return meCache;
}

/**
 * Resolve the organizer whose OneDrive/meetings we act on.
 * Omitted (or the literal "me") means the signed-in user in delegated mode; in app mode an explicit
 * id is required unless exactly one ORGANIZER_USER_IDS entry is configured.
 */
async function resolveOrganizer(c: Ctx, given?: string): Promise<string> {
  const value = given?.trim();
  if (value && value !== "me") return value;
  if (c.cfg.authMode === "delegated") {
    const me = await getMe(c);
    if (!me.id) throw new Error("Could not resolve the signed-in user from `/me`.");
    return me.id;
  }
  if (c.cfg.organizerUserIds.length === 1) return c.cfg.organizerUserIds[0]!;
  throw new Error(
    "organizerUserId is required with AUTH_MODE=app (there is no signed-in user to default to). " +
      "Pass the organizer's user object id or UPN, or configure a single ORGANIZER_USER_IDS entry.",
  );
}

async function resolveMeeting(c: Ctx, userId: string, args: { meetingId?: string; joinWebUrl?: string }): Promise<OnlineMeeting> {
  if (args.meetingId) return getOnlineMeeting(c.graph, userId, args.meetingId);
  if (args.joinWebUrl) {
    const meeting = await findOnlineMeetingByJoinUrl(c.graph, userId, args.joinWebUrl);
    if (!meeting) throw new Error(`No online meeting found for joinWebUrl ${args.joinWebUrl}`);
    return meeting;
  }
  throw new Error("Provide either meetingId or joinWebUrl.");
}

/** Internal domains: configured value wins, else the organizer's UPN domain when we know it. */
function internalDomainsFor(c: Ctx, meeting?: OnlineMeeting): string[] {
  if (c.cfg.internalDomains.length) return c.cfg.internalDomains;
  const upn = meeting?.participants?.organizer?.upn ?? meCache?.userPrincipalName ?? undefined;
  const domain = upn?.split("@")[1]?.toLowerCase();
  return domain ? [domain] : [];
}

function markExternal(attendees: Attendee[], domains: string[]): Attendee[] {
  if (!domains.length) return attendees;
  return attendees.map((a) => ({ ...a, isExternal: a.isExternal ?? isExternal(a.email, domains) }));
}

function meetingSummary(meeting: OnlineMeeting, organizerUserId: string) {
  return {
    id: meeting.id,
    subject: meeting.subject ?? null,
    joinWebUrl: meeting.joinWebUrl,
    startDateTime: meeting.startDateTime,
    endDateTime: meeting.endDateTime,
    organizerUserId,
  };
}

function driveItemSummary(item: DriveItem) {
  return {
    id: item.id,
    name: item.name,
    webUrl: item.webUrl,
    createdDateTime: item.createdDateTime,
    size: item.size,
    mimeType: item.file?.mimeType,
  };
}

/** Non-secret view of the configuration, used by the `config` resource and `whoami`. */
function publicConfig(cfg: AppConfig): Record<string, unknown> {
  return {
    tenantId: cfg.tenantId,
    clientId: cfg.clientId,
    authMode: cfg.authMode,
    hasClientSecret: Boolean(cfg.clientSecret),
    clientStateConfigured: Boolean(cfg.clientState),
    tokenCachePath: cfg.tokenCachePath,
    delegatedScopes: cfg.delegatedScopes,
    organizerUserIds: cfg.organizerUserIds,
    publicBaseUrl: cfg.publicBaseUrl,
    port: cfg.port,
    shareRole: cfg.shareRole,
    sendInvitation: cfg.sendInvitation,
    invitationMessage: cfg.invitationMessage,
    externalFallback: cfg.externalFallback,
    fallbackLinkExpiryHours: cfg.fallbackLinkExpiryHours,
    includeCalendarInvitees: cfg.includeCalendarInvitees,
    alwaysInclude: cfg.alwaysInclude,
    internalDomains: cfg.internalDomains,
    recordingsFolder: cfg.recordingsFolder,
    driveLookupTimeoutMinutes: cfg.driveLookupTimeoutMinutes,
    stateDir: cfg.stateDir,
    graphBaseUrl: cfg.graphBaseUrl,
  };
}

/* ------------------------------------------------------------------- server */

const server = new McpServer(
  { name: "teams-recording-share", version: "0.1.0" },
  {
    instructions:
      "Tools for sharing Microsoft Teams meeting recordings with everyone who attended. " +
      "Identify a meeting by its Graph onlineMeeting id (`MSp...`) or by its Teams join URL. " +
      "`organizerUserId` defaults to the signed-in user in delegated mode. " +
      "Use `find_recording` / `list_recent_recordings` to locate the file, `list_meeting_attendees` to see " +
      "who would be granted access, then `share_recording` (start with `dryRun: true`).",
  },
);

const meetingRefShape = {
  organizerUserId: z
    .string()
    .optional()
    .describe("Organizer's Entra user object id or UPN. Defaults to the signed-in user (delegated mode)."),
  meetingId: z.string().optional().describe("Graph onlineMeeting id (base64 'MSp...' form)."),
  joinWebUrl: z.string().optional().describe("Teams join URL of the meeting; used when meetingId is not known."),
};

server.registerTool(
  "share_recording",
  {
    title: "Share a Teams recording with its attendees",
    description:
      "Grant read access to a Teams meeting recording (in the organizer's OneDrive) for everyone who attended, " +
      "guests included. Resolves the meeting, its recordings, the attendance list and the recording file, then " +
      "invites everyone who does not already have access, falling back to a sharing link for recipients that " +
      "cannot be invited directly. Use dryRun first to preview the plan. Re-running is safe: people who already " +
      "have access are left untouched.",
    inputSchema: {
      ...meetingRefShape,
      recordingId: z.string().optional().describe("Specific callRecording id; omit to process every recording of the meeting."),
      driveItemId: z.string().optional().describe("Explicit OneDrive item id, skipping the recording-file search."),
      extraEmails: z.array(z.string()).optional().describe("Extra email addresses to include beyond the attendees."),
      dryRun: z.boolean().optional().describe("Compute the plan without granting anything."),
    },
    annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: true, openWorldHint: true },
  },
  async (args) =>
    run("share_recording", async (c) => {
      if (!args.meetingId && !args.joinWebUrl) throw new Error("Provide either meetingId or joinWebUrl.");
      const organizerUserId = await resolveOrganizer(c, args.organizerUserId);
      const req: ShareRequest = {
        organizerUserId,
        meetingId: args.meetingId,
        joinWebUrl: args.joinWebUrl,
        recordingId: args.recordingId,
        driveItemId: args.driveItemId,
        extraEmails: args.extraEmails,
        dryRun: args.dryRun ?? false,
      };
      return await shareRecordingWithAttendees(c.graph, c.cfg, req, log);
    }),
);

server.registerTool(
  "list_meeting_attendees",
  {
    title: "List meeting attendees",
    description:
      "List the people who attended a Teams meeting (from its attendance reports), optionally merged with the " +
      "people invited on the calendar event. Each person is marked internal/external relative to the organizer's domain.",
    inputSchema: {
      ...meetingRefShape,
      includeCalendarInvitees: z
        .boolean()
        .optional()
        .describe("Also include people invited on the calendar event who did not join. Defaults to the server config."),
    },
    annotations: { readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: true },
  },
  async (args) =>
    run("list_meeting_attendees", async (c) => {
      const organizerUserId = await resolveOrganizer(c, args.organizerUserId);
      const meeting = await resolveMeeting(c, organizerUserId, args);
      const attended = await listAttendees(c.graph, organizerUserId, meeting.id, log);
      const wantCalendar = args.includeCalendarInvitees ?? c.cfg.includeCalendarInvitees;
      const invitees = wantCalendar ? await listCalendarInvitees(c.graph, organizerUserId, meeting, log) : [];
      const domains = internalDomainsFor(c, meeting);
      const attendees = markExternal(mergeAttendees(attended, invitees), domains);
      return {
        meeting: meetingSummary(meeting, organizerUserId),
        includeCalendarInvitees: wantCalendar,
        internalDomains: domains,
        counts: {
          total: attendees.length,
          attended: attended.length,
          calendarInvitees: invitees.length,
          external: attendees.filter((a) => a.isExternal).length,
        },
        attendees,
      };
    }),
);

server.registerTool(
  "find_recording",
  {
    title: "Find a meeting's recordings",
    description:
      "List the recordings of a Teams meeting and match each one to the corresponding file in the organizer's " +
      "OneDrive Recordings folder. Use the returned driveItem.id with get_recording_access or share_recording.",
    inputSchema: { ...meetingRefShape },
    annotations: { readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: true },
  },
  async (args) =>
    run("find_recording", async (c) => {
      const organizerUserId = await resolveOrganizer(c, args.organizerUserId);
      const meeting = await resolveMeeting(c, organizerUserId, args);
      const recordings = await listRecordings(c.graph, organizerUserId, meeting.id);
      let folderItems: DriveItem[] | null = null;
      const matched = [];
      for (const rec of recordings) {
        if (!folderItems) folderItems = await listRecordingsFolder(c.graph, organizerUserId, c.cfg.recordingsFolder, 50);
        const item = pickRecordingItem(folderItems, { subject: meeting.subject, createdDateTime: rec.createdDateTime });
        matched.push({
          recordingId: rec.id,
          createdDateTime: rec.createdDateTime,
          endDateTime: rec.endDateTime,
          driveItem: item ? driveItemSummary(item) : null,
        });
      }
      return {
        meeting: meetingSummary(meeting, organizerUserId),
        folder: c.cfg.recordingsFolder,
        count: matched.length,
        recordings: matched,
        warnings: matched.some((m) => !m.driveItem)
          ? ["Some recordings have no matching file yet — the upload to OneDrive can lag behind the meeting."]
          : [],
      };
    }),
);

server.registerTool(
  "list_recent_recordings",
  {
    title: "List recent recording files",
    description:
      "List the most recent files in the organizer's OneDrive Recordings folder (newest first). Handy when the " +
      "meeting id is unknown: pick a file and pass its id as driveItemId.",
    inputSchema: {
      organizerUserId: meetingRefShape.organizerUserId,
      top: z.number().int().min(1).max(200).optional().describe("How many items to return (default 25)."),
    },
    annotations: { readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: true },
  },
  async (args) =>
    run("list_recent_recordings", async (c) => {
      const organizerUserId = await resolveOrganizer(c, args.organizerUserId);
      const items = await listRecordingsFolder(c.graph, organizerUserId, c.cfg.recordingsFolder, args.top ?? 25);
      return {
        organizerUserId,
        folder: c.cfg.recordingsFolder,
        count: items.length,
        items: items.map(driveItemSummary),
      };
    }),
);

server.registerTool(
  "get_recording_access",
  {
    title: "Show who can access a recording",
    description:
      "List the permissions on a recording file in the organizer's OneDrive and the email addresses that currently " +
      "have access (direct grants, invitations and sharing links).",
    inputSchema: {
      organizerUserId: meetingRefShape.organizerUserId,
      driveItemId: z.string().describe("OneDrive item id of the recording file."),
    },
    annotations: { readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: true },
  },
  async (args) =>
    run("get_recording_access", async (c) => {
      const organizerUserId = await resolveOrganizer(c, args.organizerUserId);
      const permissions = await listItemPermissions(c.graph, organizerUserId, args.driveItemId);
      const emails = [...emailsWithAccess(permissions)].sort();
      return {
        organizerUserId,
        driveItemId: args.driveItemId,
        emails,
        counts: { emails: emails.length, permissions: permissions.length },
        links: permissions
          .filter((p) => p.link?.webUrl)
          .map((p) => ({ id: p.id, scope: p.link?.scope, type: p.link?.type, webUrl: p.link?.webUrl, expirationDateTime: p.expirationDateTime })),
        permissions,
      };
    }),
);

server.registerTool(
  "whoami",
  {
    title: "Show the signed-in account",
    description: "Report the authentication mode and, in delegated mode, the signed-in Microsoft 365 account.",
    inputSchema: {},
    annotations: { readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: true },
  },
  async () =>
    run("whoami", async (c) => {
      const account = c.tokens.account ? await c.tokens.account() : undefined;
      const base = {
        authMode: c.cfg.authMode,
        tenantId: c.cfg.tenantId,
        clientId: c.cfg.clientId,
        graphBaseUrl: c.cfg.graphBaseUrl,
        tokenCachePath: c.cfg.tokenCachePath,
        account: account ?? null,
        organizerUserIds: c.cfg.organizerUserIds,
      };
      if (c.cfg.authMode !== "delegated") return { ...base, user: null };
      const me = await getMe(c);
      return { ...base, user: { id: me.id, displayName: me.displayName, userPrincipalName: me.userPrincipalName } };
    }),
);

server.registerResource(
  "config",
  "teams-recording-share://config",
  {
    title: "Server configuration",
    description: "Non-secret view of the server configuration (auth mode, folders, sharing behaviour).",
    mimeType: "application/json",
  },
  async (uri) => {
    let body: Record<string, unknown>;
    try {
      body = { configured: true, ...publicConfig(getCtx().cfg) };
    } catch (err) {
      body = { configured: false, error: errorText(err) };
    }
    return { contents: [{ uri: uri.href, mimeType: "application/json", text: JSON.stringify(body, null, 2) }] };
  },
);

/* --------------------------------------------------------------------- main */

async function main(): Promise<void> {
  loadDotEnv();
  try {
    getCtx();
  } catch (err) {
    // Not fatal: the server still starts, tools explain what is missing.
    log.warn("starting without a usable configuration", { error: errorText(err) });
  }
  const transport = new StdioServerTransport();
  await server.connect(transport);
  log.info("mcp server listening on stdio", {
    tools: ["share_recording", "list_meeting_attendees", "find_recording", "list_recent_recordings", "get_recording_access", "whoami"],
  });
}

main().catch((err) => {
  log.error("fatal", { error: errorText(err) });
  process.exit(1);
});
