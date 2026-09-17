/**
 * Online meeting helpers: id decoding, notification resource parsing, meeting/recording lookups
 * and the attendee normalisation used by the sharing pipeline.
 */
import type { GraphClient } from "./client.js";
import type { Logger } from "../log.js";
import type {
  Attendee,
  AttendanceRecord,
  CalendarEvent,
  CallRecording,
  MeetingAttendanceReport,
  ODataCollection,
  OnlineMeeting,
} from "./types.js";

const DAY_MS = 24 * 60 * 60 * 1000;

/* ------------------------------------------------------------------ */
/* pure helpers                                                        */
/* ------------------------------------------------------------------ */

/**
 * Decode the base64 `MSp...` online meeting id.
 * Expected plaintext shape: `1*{organizerOid}*0**{threadId}`.
 * Accepts base64url as well. Returns null when the shape does not match.
 */
export function decodeOnlineMeetingId(id: string): { organizerId: string; threadId: string } | null {
  if (typeof id !== "string" || id.trim() === "") return null;
  let decoded: string;
  try {
    const b64 = id.trim().replace(/-/g, "+").replace(/_/g, "/");
    const padded = b64 + "=".repeat((4 - (b64.length % 4)) % 4);
    decoded = Buffer.from(padded, "base64").toString("utf8");
  } catch {
    return null;
  }
  if (!decoded || decoded.includes("�")) return null;
  const parts = decoded.split("*");
  // 1 * {organizerOid} * 0 * (empty) * {threadId}
  if (parts.length < 5) return null;
  if (parts[0] !== "1") return null;
  const organizerId = (parts[1] ?? "").trim();
  const threadId = parts.slice(4).join("*").trim();
  if (!organizerId || !threadId) return null;
  return { organizerId, threadId };
}

/** Pull `name('value')` or `name/value` out of an OData resource path. */
function resourceSegment(resource: string, name: string): string | undefined {
  const paren = new RegExp(`(?:^|/)${name}\\('([^']*)'\\)`, "i").exec(resource);
  if (paren && paren[1]) return paren[1];
  const slash = new RegExp(`(?:^|/)${name}/([^/?()'"]+)`, "i").exec(resource);
  if (slash && slash[1]) {
    const v = slash[1];
    // `users/{id}/onlineMeetings/getAllRecordings` is a subscription resource, not an entity id.
    if (/^getAll(Recordings|Transcripts)$/i.test(v)) return undefined;
    return v;
  }
  return undefined;
}

/**
 * Parse a change-notification `resource`, e.g.
 *  - `communications/onlineMeetings('MSp...')/recordings('MSp...')`
 *  - `users('{guid}')/onlineMeetings('MSp...')/transcripts('...')`
 * Leading slashes are tolerated. Returns null when no meeting id can be found.
 */
export function parseNotificationResource(
  resource: string,
): { userId?: string; meetingId?: string; recordingId?: string; transcriptId?: string } | null {
  if (typeof resource !== "string" || resource.trim() === "") return null;
  const s = resource.trim().replace(/^\/+/, "");
  const meetingId = resourceSegment(s, "onlineMeetings");
  if (!meetingId) return null;
  const out: { userId?: string; meetingId?: string; recordingId?: string; transcriptId?: string } = { meetingId };
  const userId = resourceSegment(s, "users");
  if (userId) out.userId = userId;
  const recordingId = resourceSegment(s, "recordings");
  if (recordingId) out.recordingId = recordingId;
  const transcriptId = resourceSegment(s, "transcripts");
  if (transcriptId) out.transcriptId = transcriptId;
  return out;
}

/** Lowercase / trim an address, tolerating `Name <a@b.com>`, `<a@b.com>` and `mailto:` forms. */
export function normaliseEmail(e: string | null | undefined): string | null {
  if (typeof e !== "string") return null;
  let v = e.trim();
  if (!v) return null;
  const angle = /<([^>]+)>\s*$/.exec(v);
  if (angle && angle[1]) v = angle[1].trim();
  v = v.replace(/^mailto:/i, "").replace(/^sip:/i, "").trim();
  v = v.replace(/^[<"'\s]+/, "").replace(/[>"'\s]+$/, "");
  v = v.toLowerCase();
  if (!v || v.indexOf("@") <= 0 || v.endsWith("@") || /\s/.test(v)) return null;
  return v;
}

/** True when the address' domain is not one of the internal domains (case-insensitive). */
export function isExternal(email: string, internalDomains: string[]): boolean {
  const addr = normaliseEmail(email);
  if (!addr) return true;
  const domain = addr.slice(addr.lastIndexOf("@") + 1);
  const internal = (internalDomains ?? [])
    .map((d) => (d ?? "").trim().toLowerCase().replace(/^@/, ""))
    .filter(Boolean);
  if (internal.length === 0) return false;
  return !internal.includes(domain);
}

/**
 * Merge attendee lists. First occurrence of an email wins for the scalar fields, but
 * `isExternal` is OR-ed and the highest `totalAttendanceInSeconds` is kept.
 */
export function mergeAttendees(...lists: Attendee[][]): Attendee[] {
  const byEmail = new Map<string, Attendee>();
  for (const list of lists) {
    for (const a of list ?? []) {
      const email = normaliseEmail(a?.email);
      if (!email) continue;
      const existing = byEmail.get(email);
      if (!existing) {
        byEmail.set(email, { ...a, email });
        continue;
      }
      if (a.isExternal || existing.isExternal) existing.isExternal = Boolean(a.isExternal || existing.isExternal);
      const secs = a.totalAttendanceInSeconds;
      if (typeof secs === "number" && (existing.totalAttendanceInSeconds ?? -1) < secs) {
        existing.totalAttendanceInSeconds = secs;
      }
      if (!existing.displayName && a.displayName) existing.displayName = a.displayName;
      if (!existing.role && a.role) existing.role = a.role;
      if (!existing.userId && a.userId) existing.userId = a.userId;
    }
  }
  return [...byEmail.values()];
}

/** Normalise a join URL for comparison: decoded, lowercased, no query/fragment, no trailing slashes. */
function normaliseJoinUrl(url: string | null | undefined): string | null {
  if (typeof url !== "string") return null;
  let v = url.trim();
  if (!v) return null;
  try {
    v = decodeURIComponent(v);
  } catch {
    /* keep the raw value when it is not valid percent-encoding */
  }
  v = v.split("#")[0] ?? v;
  v = v.split("?")[0] ?? v;
  v = v.replace(/\/+$/, "");
  return v.toLowerCase() || null;
}

/* ------------------------------------------------------------------ */
/* Graph calls                                                         */
/* ------------------------------------------------------------------ */

const seg = (v: string): string => encodeURIComponent(v);

export async function getOnlineMeeting(g: GraphClient, userId: string, meetingId: string): Promise<OnlineMeeting> {
  return g.get<OnlineMeeting>(`/users/${seg(userId)}/onlineMeetings/${seg(meetingId)}`);
}

export async function findOnlineMeetingByJoinUrl(
  g: GraphClient,
  userId: string,
  joinWebUrl: string,
): Promise<OnlineMeeting | null> {
  const escaped = joinWebUrl.replace(/'/g, "''");
  const res = await g.get<ODataCollection<OnlineMeeting>>(`/users/${seg(userId)}/onlineMeetings`, {
    query: { $filter: `JoinWebUrl eq '${escaped}'` },
  });
  return res?.value?.[0] ?? null;
}

export async function listRecordings(g: GraphClient, userId: string, meetingId: string): Promise<CallRecording[]> {
  return g.getAll<CallRecording>(`/users/${seg(userId)}/onlineMeetings/${seg(meetingId)}/recordings`);
}

export async function getRecording(
  g: GraphClient,
  userId: string,
  meetingId: string,
  recordingId: string,
): Promise<CallRecording> {
  return g.get<CallRecording>(
    `/users/${seg(userId)}/onlineMeetings/${seg(meetingId)}/recordings/${seg(recordingId)}`,
  );
}

/**
 * Union of every attendance report's records, deduped by email. Records without an email address
 * (anonymous / phone participants) are skipped and logged at debug level.
 */
export async function listAttendees(
  g: GraphClient,
  userId: string,
  meetingId: string,
  log?: Logger,
): Promise<Attendee[]> {
  const base = `/users/${seg(userId)}/onlineMeetings/${seg(meetingId)}`;
  const reports = await g.getAll<MeetingAttendanceReport>(`${base}/attendanceReports`);
  const byEmail = new Map<string, Attendee>();
  for (const report of reports) {
    if (!report?.id) continue;
    const records = await g.getAll<AttendanceRecord>(`${base}/attendanceReports/${seg(report.id)}/attendanceRecords`);
    for (const record of records) {
      const email = normaliseEmail(record?.emailAddress ?? record?.identity?.userPrincipalName);
      if (!email) {
        log?.debug("attendance record without an email address, skipped", {
          reportId: report.id,
          identityId: record?.identity?.id,
          displayName: record?.identity?.displayName,
        });
        continue;
      }
      const existing = byEmail.get(email);
      if (!existing) {
        byEmail.set(email, {
          email,
          displayName: record.identity?.displayName ?? undefined,
          source: "attendance",
          role: record.role ?? undefined,
          totalAttendanceInSeconds: record.totalAttendanceInSeconds ?? undefined,
          userId: record.identity?.id ?? undefined,
        });
        continue;
      }
      const secs = record.totalAttendanceInSeconds;
      if (typeof secs === "number" && (existing.totalAttendanceInSeconds ?? -1) < secs) {
        existing.totalAttendanceInSeconds = secs;
      }
      if (!existing.displayName && record.identity?.displayName) existing.displayName = record.identity.displayName;
      if (!existing.role && record.role) existing.role = record.role;
      if (!existing.userId && record.identity?.id) existing.userId = record.identity.id;
    }
  }
  return [...byEmail.values()];
}

/**
 * People invited on the organizer's calendar event backing this meeting.
 * Returns [] when no event matches the meeting's join URL.
 */
export async function listCalendarInvitees(
  g: GraphClient,
  userId: string,
  meeting: OnlineMeeting,
  log?: Logger,
): Promise<Attendee[]> {
  const target = normaliseJoinUrl(meeting?.joinWebUrl);
  if (!target) {
    log?.debug("meeting has no joinWebUrl, skipping calendar lookup", { meetingId: meeting?.id });
    return [];
  }
  const startMs = meeting.startDateTime ? Date.parse(meeting.startDateTime) : NaN;
  if (!Number.isFinite(startMs)) {
    log?.debug("meeting has no usable startDateTime, skipping calendar lookup", { meetingId: meeting?.id });
    return [];
  }
  const endMs = meeting.endDateTime ? Date.parse(meeting.endDateTime) : NaN;
  const windowStart = new Date(startMs - DAY_MS).toISOString();
  const windowEnd = new Date(Number.isFinite(endMs) ? endMs : startMs + DAY_MS).toISOString();

  const events = await g.getAll<CalendarEvent>(`/users/${seg(userId)}/calendarView`, {
    query: {
      startDateTime: windowStart,
      endDateTime: windowEnd,
      $select: "id,subject,start,end,isOnlineMeeting,onlineMeeting,attendees,organizer",
      $top: 50,
    },
  });

  const event = events.find((e) => normaliseJoinUrl(e?.onlineMeeting?.joinUrl) === target);
  if (!event) {
    log?.debug("no calendar event matched the meeting join url", { meetingId: meeting?.id, events: events.length });
    return [];
  }

  const organizerEmail = normaliseEmail(event.organizer?.emailAddress?.address);
  const out: Attendee[] = [];
  const seen = new Set<string>();
  for (const a of event.attendees ?? []) {
    if ((a?.type ?? "").toLowerCase() === "resource") continue;
    const email = normaliseEmail(a?.emailAddress?.address);
    if (!email) continue;
    if (organizerEmail && email === organizerEmail) continue; // added separately as source "organizer"
    if (seen.has(email)) continue;
    seen.add(email);
    out.push({
      email,
      displayName: a?.emailAddress?.name ?? undefined,
      source: "calendar",
      role: a?.type ?? undefined,
    });
  }
  return out;
}
