import test from "node:test";
import assert from "node:assert/strict";

import {
  decodeOnlineMeetingId,
  parseNotificationResource,
  normaliseEmail,
  isExternal,
  mergeAttendees,
} from "../src/graph/meetings.js";
import type { Attendee } from "../src/graph/types.js";

const ORGANIZER_OID = "880e1e61-af63-43c9-a48e-d6b63684c21c";
const THREAD_ID = "19:meeting_ZmY4YTJhNmUtY2Y3OS00ZGZlLWI1ODEtYjA2ZTZmZjA0YzQ0@thread.v2";
/** Graph online meeting ids are base64 of `1*{organizerOid}*0**{threadId}`. */
const MEETING_ID = Buffer.from(`1*${ORGANIZER_OID}*0**${THREAD_ID}`, "utf8").toString("base64");

test("decodeOnlineMeetingId returns the organizer oid and thread id", () => {
  const decoded = decodeOnlineMeetingId(MEETING_ID);
  assert.ok(decoded, "expected a decoded meeting id");
  assert.equal(decoded.organizerId, ORGANIZER_OID);
  assert.equal(decoded.threadId, THREAD_ID);
});

test("decodeOnlineMeetingId returns null for ids that are not meeting ids", () => {
  assert.equal(decodeOnlineMeetingId(""), null);
  assert.equal(decodeOnlineMeetingId("not a meeting id !!!"), null);
  // Valid base64, but the decoded payload has none of the expected structure.
  assert.equal(decodeOnlineMeetingId(Buffer.from("hello world").toString("base64")), null);
});

test("parseNotificationResource handles the communications/ form", () => {
  const parsed = parseNotificationResource(`communications/onlineMeetings('${MEETING_ID}')/recordings('REC123')`);
  assert.ok(parsed, "expected the resource to parse");
  assert.equal(parsed.meetingId, MEETING_ID);
  assert.equal(parsed.recordingId, "REC123");
  assert.ok(!parsed.userId, "communications/ form carries no user id");
});

test("parseNotificationResource handles the users('id') form", () => {
  const parsed = parseNotificationResource(`users('${ORGANIZER_OID}')/onlineMeetings('${MEETING_ID}')/recordings('REC123')`);
  assert.ok(parsed, "expected the resource to parse");
  assert.equal(parsed.userId, ORGANIZER_OID);
  assert.equal(parsed.meetingId, MEETING_ID);
  assert.equal(parsed.recordingId, "REC123");
});

test("parseNotificationResource tolerates a leading slash", () => {
  const parsed = parseNotificationResource(`/communications/onlineMeetings('${MEETING_ID}')/recordings('REC123')`);
  assert.ok(parsed, "expected the resource to parse");
  assert.equal(parsed.meetingId, MEETING_ID);
  assert.equal(parsed.recordingId, "REC123");
});

test("parseNotificationResource reports transcripts separately from recordings", () => {
  const parsed = parseNotificationResource(`communications/onlineMeetings('${MEETING_ID}')/transcripts('TR1')`);
  assert.ok(parsed, "expected the resource to parse");
  assert.equal(parsed.meetingId, MEETING_ID);
  assert.equal(parsed.transcriptId, "TR1");
  assert.ok(!parsed.recordingId, "a transcript resource is not a recording");
});

test("parseNotificationResource returns null for unrelated resources", () => {
  assert.equal(parseNotificationResource(""), null);
  assert.equal(parseNotificationResource("total nonsense"), null);
});

test("normaliseEmail trims, lowercases and rejects empties", () => {
  assert.equal(normaliseEmail("  Foo.Bar@Example.COM "), "foo.bar@example.com");
  assert.equal(normaliseEmail("already@lower.com"), "already@lower.com");
  assert.equal(normaliseEmail(""), null);
  assert.equal(normaliseEmail("   "), null);
  assert.equal(normaliseEmail(null), null);
  assert.equal(normaliseEmail(undefined), null);
});

test("isExternal compares the domain against the internal list, case-insensitively", () => {
  assert.equal(isExternal("alice@contoso.com", ["contoso.com"]), false);
  assert.equal(isExternal("Alice@CONTOSO.com", ["contoso.com"]), false);
  assert.equal(isExternal("alice@contoso.com", ["CONTOSO.COM"]), false);
  assert.equal(isExternal("dana@partner.example", ["contoso.com"]), true);
  assert.equal(isExternal("bob@fabrikam.com", ["contoso.com", "fabrikam.com"]), false);
});

test("mergeAttendees dedupes by email and keeps the first occurrence", () => {
  const attendance: Attendee[] = [
    { email: "alice@contoso.com", displayName: "Alice A", source: "attendance", role: "Presenter" },
    { email: "bob@contoso.com", source: "attendance" },
    { email: "bob@contoso.com", source: "attendance" }, // duplicate inside one list
  ];
  const calendar: Attendee[] = [
    { email: "alice@contoso.com", displayName: "Alice From Calendar", source: "calendar" },
    { email: "dana@partner.example", source: "calendar" },
  ];
  const extra: Attendee[] = [{ email: "ops@contoso.com", source: "extra" }];

  const merged = mergeAttendees(attendance, calendar, extra);
  assert.deepEqual(
    merged.map((a) => a.email),
    ["alice@contoso.com", "bob@contoso.com", "dana@partner.example", "ops@contoso.com"],
  );
  const alice = merged.find((a) => a.email === "alice@contoso.com");
  assert.equal(alice?.source, "attendance", "first occurrence wins");
  assert.equal(alice?.displayName, "Alice A");
});

test("mergeAttendees handles no lists and empty lists", () => {
  assert.deepEqual(mergeAttendees(), []);
  assert.deepEqual(mergeAttendees([], []), []);
});
