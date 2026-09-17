import test from "node:test";
import assert from "node:assert/strict";

import { pickRecordingItem, emailsWithAccess } from "../src/graph/drive.js";
import type { DriveItem, Permission } from "../src/graph/types.js";

/** Recording start used by every case below. */
const BASE = "2025-01-15T10:00:00Z";
const minutes = (n: number) => new Date(Date.parse(BASE) + n * 60_000).toISOString();

function mp4(id: string, name: string, createdDateTime: string): DriveItem {
  return { id, name, createdDateTime, size: 1024, file: { mimeType: "video/mp4" } };
}

test("pickRecordingItem prefers a subject match over a closer non-matching file", () => {
  const items: DriveItem[] = [
    mp4("near", "Some Other Call-20250115_100100UTC-Meeting Recording.mp4", minutes(1)),
    mp4("match", "Weekly Sync-20250115_100800UTC-Meeting Recording.mp4", minutes(8)),
  ];
  const picked = pickRecordingItem(items, { subject: "Weekly Sync", createdDateTime: BASE, toleranceMinutes: 15 });
  assert.equal(picked?.id, "match");
});

test("pickRecordingItem falls back to the closest mp4 in the window when no subject matches", () => {
  const items: DriveItem[] = [
    mp4("far", "Some Other Call-20250115_101200UTC-Meeting Recording.mp4", minutes(12)),
    mp4("near", "Another Call-20250115_100200UTC-Meeting Recording.mp4", minutes(2)),
  ];
  const picked = pickRecordingItem(items, { subject: "Weekly Sync", createdDateTime: BASE, toleranceMinutes: 15 });
  assert.equal(picked?.id, "near");
});

test("pickRecordingItem returns null when nothing is inside the tolerance window", () => {
  const items: DriveItem[] = [
    mp4("old", "Weekly Sync-20250115_090000UTC-Meeting Recording.mp4", minutes(-60)),
    mp4("new", "Weekly Sync-20250115_110000UTC-Meeting Recording.mp4", minutes(60)),
  ];
  const picked = pickRecordingItem(items, { subject: "Weekly Sync", createdDateTime: BASE, toleranceMinutes: 15 });
  assert.equal(picked, null);
});

test("pickRecordingItem returns null for an empty folder", () => {
  assert.equal(pickRecordingItem([], { subject: "Weekly Sync", createdDateTime: BASE, toleranceMinutes: 15 }), null);
});

test("pickRecordingItem ignores non-mp4 items even when the name and time match", () => {
  const items: DriveItem[] = [
    { id: "vtt", name: "Weekly Sync-20250115_100100UTC-Meeting Recording.vtt", createdDateTime: minutes(1), file: { mimeType: "text/vtt" } },
    { id: "folder", name: "Weekly Sync", createdDateTime: minutes(1), folder: { childCount: 0 } },
  ];
  assert.equal(pickRecordingItem(items, { subject: "Weekly Sync", createdDateTime: BASE, toleranceMinutes: 15 }), null);

  const withVideo = [...items, mp4("mp4", "Weekly Sync-20250115_100900UTC-Meeting Recording.mp4", minutes(9))];
  assert.equal(pickRecordingItem(withVideo, { subject: "Weekly Sync", createdDateTime: BASE, toleranceMinutes: 15 })?.id, "mp4");
});

test("pickRecordingItem matches subjects whose characters Teams strips from the file name", () => {
  // Teams removes / \ : ? * " < > | from the subject when it names the file.
  const items: DriveItem[] = [
    mp4("other", "Standup-20250115_100100UTC-Meeting Recording.mp4", minutes(1)),
    mp4("match", "Q3 Review Plan  Budget-20250115_100700UTC-Meeting Recording.mp4", minutes(7)),
  ];
  const picked = pickRecordingItem(items, {
    subject: "Q3 Review: Plan / Budget",
    createdDateTime: BASE,
    toleranceMinutes: 15,
  });
  assert.equal(picked?.id, "match");
});

test("pickRecordingItem matches subjects case-insensitively", () => {
  const items = [mp4("match", "weekly SYNC-20250115_100300UTC-Meeting Recording.mp4", minutes(3))];
  assert.equal(
    pickRecordingItem(items, { subject: "Weekly Sync", createdDateTime: BASE, toleranceMinutes: 15 })?.id,
    "match",
  );
});

test("emailsWithAccess collects emails from every permission shape, lowercased", () => {
  const perms: Permission[] = [
    { id: "p1", roles: ["read"], grantedToV2: { siteUser: { email: "Alice@Contoso.com", loginName: "i:0#.f|membership|alice@contoso.com" } } },
    { id: "p2", roles: ["read"], grantedToV2: { siteUser: { loginName: "i:0#.f|membership|Bob@Contoso.com" } } },
    { id: "p3", roles: ["read"], invitation: { email: "Carol@Partner.Example", signInRequired: true } },
    {
      id: "p4",
      roles: ["read"],
      grantedToIdentitiesV2: [
        { siteUser: { email: "Dave@Contoso.com" } },
        { siteUser: { loginName: "i:0#.f|membership|erin@contoso.com" } },
      ],
    },
    { id: "p5", roles: ["read"], link: { scope: "anonymous", type: "view", webUrl: "https://example.sharepoint.com/:v:/g/abc" } },
    { id: "p6", roles: ["owner"], grantedToV2: { user: { id: "oid-1", displayName: "Organizer" } } },
  ];

  const emails = emailsWithAccess(perms);
  assert.ok(emails instanceof Set);
  for (const e of ["alice@contoso.com", "bob@contoso.com", "carol@partner.example", "dave@contoso.com", "erin@contoso.com"]) {
    assert.ok(emails.has(e), `expected ${e} to have access`);
  }
  // The claims prefix must be stripped, not carried into the set.
  for (const e of emails) {
    assert.ok(!e.includes("i:0#.f|"), `claims prefix leaked into ${e}`);
    assert.equal(e, e.toLowerCase(), `${e} is not lowercased`);
  }
  assert.equal(emails.size, 5, [...emails].join(", "));
});

test("emailsWithAccess returns an empty set for no permissions", () => {
  assert.equal(emailsWithAccess([]).size, 0);
});
