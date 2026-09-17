import test from "node:test";
import assert from "node:assert/strict";

import { loadConfig, type AppConfig } from "../src/config.js";
import { silentLogger } from "../src/log.js";
import { shareRecordingWithAttendees } from "../src/core/share.js";
import type { ShareResult } from "../src/core/types.js";
import type { GraphClient } from "../src/graph/client.js";

const ORGANIZER_OID = "880e1e61-af63-43c9-a48e-d6b63684c21c";
const THREAD_ID = "19:meeting_ZmY4YTJhNmUtY2Y3OS00ZGZlLWI1ODEtYjA2ZTZmZjA0YzQ0@thread.v2";
const MEETING_ID = Buffer.from(`1*${ORGANIZER_OID}*0**${THREAD_ID}`, "utf8").toString("base64");

const SUBJECT = "Quarterly Review";
const RECORDED_AT = "2025-01-15T10:05:00Z";
const ORGANIZER_UPN = "organizer@contoso.com";
const ALICE = "alice@contoso.com";
const BOB = "bob@contoso.com";
const DANA = "dana@partner.example";
const LINK_URL = "https://contoso-my.sharepoint.com/:v:/g/personal/organizer_contoso_com/fallback-link";

type Call = { method: string; path: string; body?: any };

/**
 * Minimal GraphClient stand-in: routes on the request path and replays canned Graph payloads.
 * Unrouted requests throw loudly so a pipeline change shows up as a clear failure.
 */
function fakeGraph() {
  const calls: Call[] = [];

  const meeting = {
    id: MEETING_ID,
    subject: SUBJECT,
    joinWebUrl: "https://teams.microsoft.com/l/meetup-join/19%3ameeting_x%40thread.v2/0",
    startDateTime: "2025-01-15T10:00:00Z",
    endDateTime: "2025-01-15T11:00:00Z",
    participants: { organizer: { identity: { user: { id: ORGANIZER_OID, displayName: "Olivia Organizer" } }, upn: ORGANIZER_UPN } },
  };

  const attendanceRecords = [
    { id: "r1", emailAddress: "Alice@Contoso.com", role: "Presenter", totalAttendanceInSeconds: 3000, identity: { id: "oid-alice", displayName: "Alice Adams" } },
    { id: "r2", emailAddress: BOB, role: "Attendee", totalAttendanceInSeconds: 2400, identity: { id: "oid-bob", displayName: "Bob Brown" } },
    { id: "r3", emailAddress: DANA, role: "Attendee", totalAttendanceInSeconds: 1800, identity: { id: "oid-dana", displayName: "Dana Doe" } },
    { id: "r4", emailAddress: null, role: "Attendee", totalAttendanceInSeconds: 120, identity: { displayName: "Anonymous joiner" } },
  ];

  const driveItem = {
    id: "item-1",
    name: `${SUBJECT}-20250115_100500UTC-Meeting Recording.mp4`,
    webUrl: "https://contoso-my.sharepoint.com/personal/organizer/Documents/Recordings/quarterly.mp4",
    createdDateTime: RECORDED_AT,
    size: 123456,
    file: { mimeType: "video/mp4" },
  };

  const permissions = [
    { id: "perm-owner", roles: ["owner"], grantedToV2: { siteUser: { email: ORGANIZER_UPN, loginName: `i:0#.f|membership|${ORGANIZER_UPN}` } } },
    { id: "perm-alice", roles: ["read"], grantedToV2: { siteUser: { email: "Alice@Contoso.com", loginName: `i:0#.f|membership|${ALICE}` } } },
  ];

  const routes: { method?: string; re: RegExp; handler: (p: string, body: any) => any }[] = [
    { re: /\/attendanceReports\/[^/?]+\/attendanceRecords/, handler: () => ({ value: attendanceRecords }) },
    { re: /\/attendanceReports(\?|$)/, handler: () => ({ value: [{ id: "report-1", totalParticipantCount: 4 }] }) },
    { re: /\/recordings\/[^/?]+(\?|$)/, handler: () => ({ id: "rec1", meetingId: MEETING_ID, createdDateTime: RECORDED_AT, endDateTime: "2025-01-15T11:00:00Z" }) },
    { re: /\/recordings(\?|$)/, handler: () => ({ value: [{ id: "rec1", meetingId: MEETING_ID, createdDateTime: RECORDED_AT, endDateTime: "2025-01-15T11:00:00Z" }] }) },
    { re: /\/calendarView/, handler: () => ({ value: [] }) },
    { re: /\/drive\/root:\/[^:]+:\/children/, handler: () => ({ value: [driveItem, { id: "item-2", name: "Standup-20250115_090000UTC-Meeting Recording.mp4", createdDateTime: "2025-01-15T09:00:00Z", file: { mimeType: "video/mp4" } }] }) },
    { re: /\/permissions\/[^/?]+\/grant/, handler: (_p, body) => ({
        value: (body?.recipients ?? []).map((r: any, i: number) => ({
          id: `perm-granted-${i}`,
          roles: body?.roles ?? ["read"],
          grantedToIdentitiesV2: [{ siteUser: { email: r.email } }],
          link: { scope: "users", type: "view", webUrl: LINK_URL },
        })),
      }) },
    { re: /\/items\/[^/?]+\/permissions(\?|$)/, handler: () => ({ value: permissions }) },
    { re: /\/items\/[^/?]+\/invite/, handler: (_p, body) => ({
        // Graph answers 207 by returning an `error` entry in place of the permission for each
        // recipient it could not add; brand new external guests fail under app-only auth.
        value: (body?.recipients ?? []).map((r: any, i: number) => {
          const email = String(r.email ?? "").toLowerCase();
          if (email.endsWith("@contoso.com")) {
            return { id: `perm-invited-${i}`, roles: body?.roles ?? ["read"], invitation: { email, signInRequired: true }, grantedToIdentitiesV2: [{ siteUser: { email } }] };
          }
          return { invitation: { email }, error: { code: "accountVerificationRequired", message: `The account ${email} needs to be verified before it can be invited.` } };
        }),
      }) },
    { re: /\/items\/[^/?]+\/createLink/, handler: (_p, body) => ({
        id: "perm-link",
        roles: ["read"],
        link: { scope: body?.scope ?? "users", type: body?.type ?? "view", webUrl: LINK_URL },
        expirationDateTime: body?.expirationDateTime ?? null,
      }) },
    { re: /\/items\/[^/?]+(\?|$)/, handler: () => driveItem },
    { re: /\/sendMail(\?|$)/, handler: () => ({}) },
    { re: /\/onlineMeetings\?|\/onlineMeetings(\?|$)/, handler: () => ({ value: [meeting] }) },
    { re: /\/onlineMeetings\/[^/?]+(\?|$)/, handler: () => meeting },
    { re: /^\/?(me|users\/[^/?]+)(\?|$)/, handler: () => ({ id: ORGANIZER_OID, displayName: "Olivia Organizer", userPrincipalName: ORGANIZER_UPN, mail: ORGANIZER_UPN }) },
  ];

  function route(method: string, p: string, body?: any) {
    calls.push({ method, path: p, body });
    const rel = p.replace(/^https?:\/\/[^/]+\/(v1\.0|beta)/, "");
    for (const r of routes) {
      if (r.method && r.method !== method) continue;
      if (r.re.test(rel)) return r.handler(rel, body);
    }
    throw new Error(`fakeGraph: no route for ${method} ${p}`);
  }

  const g = {
    calls,
    driveItem,
    get: async (p: string) => route("GET", p),
    getAll: async (p: string) => {
      const res: any = route("GET", p);
      return Array.isArray(res) ? res : (res?.value ?? []);
    },
    post: async (p: string, body: unknown) => route("POST", p, body),
    patch: async (p: string, body: unknown) => route("PATCH", p, body),
    delete: async (p: string) => {
      route("DELETE", p);
    },
  };
  return g;
}

function cfgFor(extra: NodeJS.ProcessEnv = {}): AppConfig {
  return loadConfig({
    TENANT_ID: "11111111-2222-3333-4444-555555555555",
    CLIENT_ID: "66666666-7777-8888-9999-000000000000",
    AUTH_MODE: "delegated",
    INTERNAL_DOMAINS: "contoso.com",
    EXTERNAL_FALLBACK: "users-link",
    SEND_INVITATION: "false",
    DRIVE_LOOKUP_TIMEOUT_MINUTES: "1",
    ...extra,
  });
}

const emails = (people: { email: string }[]) => people.map((p) => p.email.toLowerCase());

test("shares a recording: invites the missing attendees and falls back to a link for the external one", { timeout: 30_000 }, async () => {
  const g = fakeGraph();
  const result: ShareResult = await shareRecordingWithAttendees(
    g as unknown as GraphClient,
    cfgFor(),
    { organizerUserId: ORGANIZER_OID, meetingId: MEETING_ID },
    silentLogger,
  );

  assert.equal(result.dryRun, false);
  assert.equal(result.meeting.id, MEETING_ID);
  assert.equal(result.meeting.organizerUserId, ORGANIZER_OID);
  assert.equal(result.recordings.length, 1);

  const rec = result.recordings[0];
  assert.equal(rec.recordingId, "rec1");
  assert.equal(rec.driveItem?.id, "item-1", "the subject-matching mp4 in the Recordings folder must be picked");

  const granted = emails(rec.granted);
  const already = emails(rec.alreadyHadAccess);
  assert.ok(granted.includes(BOB), `expected ${BOB} to be granted, got ${granted.join(", ")}`);
  assert.ok(!granted.includes(ALICE), "Alice already had access and must not be re-granted");
  assert.ok(already.includes(ALICE), `expected ${ALICE} in alreadyHadAccess, got ${already.join(", ")}`);

  // The external guest is rejected by `invite` (accountVerificationRequired) and must be picked up
  // by the configured fallback: a users-scoped link granted to exactly that address.
  assert.ok(rec.fallbackLink, "a fallback sharing link should have been created for the failed invite");
  assert.equal(rec.fallbackLink.webUrl, LINK_URL);
  assert.equal(rec.fallbackLink.scope, "users");
  const reachedDana = [...granted, ...rec.fallbackLink.sentTo.map((e) => e.toLowerCase())];
  assert.ok(reachedDana.includes(DANA), `the external guest must end up on the fallback link, got ${reachedDana.join(", ")}`);
  const grantCall = g.calls.find((c) => c.method === "POST" && /\/grant(\?|$)/.test(c.path));
  assert.ok(grantCall, "the users-scoped link must be granted to the recipients invite could not add");
  assert.match(JSON.stringify(grantCall.body).toLowerCase(), new RegExp(DANA));
  const inviteCall = g.calls.find((c) => c.method === "POST" && /\/invite(\?|$)/.test(c.path));
  assert.ok(inviteCall, "missing attendees must be invited directly first");
  assert.ok(!JSON.stringify(inviteCall.body).toLowerCase().includes(ALICE), "Alice already had access; she must not be re-invited");

  // Attendance roll-up: every record with an email, deduped and lowercased; the anonymous one dropped.
  const attendeeEmails = result.attendees.map((a) => a.email);
  for (const e of [ALICE, BOB, DANA]) assert.ok(attendeeEmails.includes(e), `missing attendee ${e}`);
  assert.equal(new Set(attendeeEmails).size, attendeeEmails.length, "attendees must be deduped");
  for (const a of result.attendees) {
    assert.ok(a.email && a.email.includes("@"), "attendance records without an email must be skipped");
    assert.equal(a.email, a.email.toLowerCase());
  }
  assert.equal(result.attendees.find((a) => a.email === DANA)?.isExternal, true);
  assert.ok(!result.attendees.find((a) => a.email === ALICE)?.isExternal, "an internal attendee must not be marked external");
  assert.ok(result.attendees.some((a) => a.email === ORGANIZER_UPN), "the organizer should be part of the audience");
});

test("dry run changes nothing", { timeout: 30_000 }, async () => {
  const g = fakeGraph();
  const result = await shareRecordingWithAttendees(
    g as unknown as GraphClient,
    cfgFor(),
    { organizerUserId: ORGANIZER_OID, meetingId: MEETING_ID, dryRun: true },
    silentLogger,
  );

  assert.equal(result.dryRun, true);
  const writes = g.calls.filter((c) => c.method !== "GET");
  assert.deepEqual(writes, [], `dry run must not write: ${writes.map((w) => `${w.method} ${w.path}`).join(", ")}`);
  assert.equal(result.recordings.length, 1);
  assert.equal(result.recordings[0].driveItem?.id, "item-1");
  assert.ok(result.attendees.length >= 3);
});

test("extra emails are included in the audience", { timeout: 30_000 }, async () => {
  const g = fakeGraph();
  const result = await shareRecordingWithAttendees(
    g as unknown as GraphClient,
    cfgFor({ ALWAYS_INCLUDE: "Archive@Contoso.com" }),
    { organizerUserId: ORGANIZER_OID, meetingId: MEETING_ID, extraEmails: ["Extra.Person@Contoso.com"], dryRun: true },
    silentLogger,
  );
  const attendeeEmails = result.attendees.map((a) => a.email);
  assert.ok(attendeeEmails.includes("extra.person@contoso.com"), attendeeEmails.join(", "));
  assert.ok(attendeeEmails.includes("archive@contoso.com"), attendeeEmails.join(", "));
});

test("a meeting that cannot be resolved throws", { timeout: 30_000 }, async () => {
  const g = {
    get: async () => {
      const err: any = new Error("The requested meeting was not found.");
      err.status = 404;
      err.code = "itemNotFound";
      throw err;
    },
    getAll: async () => [],
    post: async () => ({}),
    patch: async () => ({}),
    delete: async () => undefined,
  };
  await assert.rejects(() =>
    shareRecordingWithAttendees(
      g as unknown as GraphClient,
      cfgFor(),
      { organizerUserId: ORGANIZER_OID, meetingId: MEETING_ID },
      silentLogger,
    ),
  );
});
