/**
 * OneDrive (drive item) operations for Teams meeting recordings:
 * locating the recording file in the organizer's `Recordings` folder, reading who already has
 * access, inviting attendees, creating fallback sharing links and sending mail.
 */
import { GraphClient, GraphError } from "./client.js";
import type { Logger } from "../log.js";
import type { DriveItem, ODataCollection, Permission } from "./types.js";

/** SharePoint claims prefix that decorates `siteUser.loginName`. */
const CLAIMS_PREFIX = "i:0#.f|membership|";
/** Characters Teams / OneDrive strip out of a meeting subject when naming the recording file. */
const STRIPPED_NAME_CHARS = /[/\\:?*"<>|#%{}~&]/g;
const INVITE_BATCH_SIZE = 20;
const DEFAULT_TOLERANCE_MINUTES = 30;
const DEFAULT_POLL_INTERVAL_MS = 30_000;
const DEFAULT_TIMEOUT_MS = 20 * 60_000;

function enc(segment: string): string {
  return encodeURIComponent(segment);
}

/** Encode a folder path such as "Recordings" or "Recordings/2025" for the `/root:/<path>:` addressing form. */
function encodeFolderPath(folder: string): string {
  return folder
    .split("/")
    .map((s) => s.trim())
    .filter(Boolean)
    .map(enc)
    .join("/");
}

function driveItemsPath(userId: string, itemId: string): string {
  return `/users/${enc(userId)}/drive/items/${enc(itemId)}`;
}

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function chunk<T>(items: T[], size: number): T[][] {
  const out: T[][] = [];
  for (let i = 0; i < items.length; i += size) out.push(items.slice(i, i + size));
  return out;
}

/** List the children of the organizer's recordings folder, newest first. */
export async function listRecordingsFolder(
  g: GraphClient,
  userId: string,
  folder: string,
  top = 50,
): Promise<DriveItem[]> {
  const encoded = encodeFolderPath(folder);
  const path = encoded
    ? `/users/${enc(userId)}/drive/root:/${encoded}:/children`
    : `/users/${enc(userId)}/drive/root/children`;
  const page = await g.get<ODataCollection<DriveItem> | undefined>(path, {
    query: { $orderby: "createdDateTime desc", $top: top },
  });
  return page?.value ?? [];
}

/**
 * Normalise a meeting subject or file name for comparison: drop the characters OneDrive strips from
 * file names, lowercase, and collapse runs of whitespace (Teams also turns some of them into spaces).
 */
export function normaliseName(value: string): string {
  return value
    .replace(STRIPPED_NAME_CHARS, "")
    .replace(/[‘’]/g, "'")
    .replace(/[“”]/g, "")
    .replace(/[\s_]+/g, " ")
    .trim()
    .toLowerCase();
}

function isMp4(item: DriveItem): boolean {
  if (item.folder) return false;
  if (typeof item.name === "string" && /\.mp4$/i.test(item.name)) return true;
  return item.file?.mimeType === "video/mp4";
}

function timeOf(value: string | undefined | null): number | undefined {
  if (!value) return undefined;
  const t = Date.parse(value);
  return Number.isFinite(t) ? t : undefined;
}

/**
 * Pure matcher: pick the drive item that corresponds to a recording.
 *
 * 1. Prefer an .mp4 whose (normalised) name starts with the (normalised) meeting subject and whose
 *    createdDateTime is within `toleranceMinutes` of the recording's createdDateTime; ties broken by
 *    the smallest time delta.
 * 2. Otherwise the .mp4 closest in time within the tolerance.
 * 3. Otherwise null.
 */
export function pickRecordingItem(
  items: DriveItem[],
  opts: { subject?: string | null; createdDateTime?: string; toleranceMinutes?: number },
): DriveItem | null {
  const candidates = (items ?? []).filter(isMp4);
  if (candidates.length === 0) return null;

  const toleranceMs = (opts.toleranceMinutes ?? DEFAULT_TOLERANCE_MINUTES) * 60_000;
  const reference = timeOf(opts.createdDateTime);
  const subject = opts.subject ? normaliseName(opts.subject) : "";

  /** Time distance from the recording; undefined when either side has no usable timestamp. */
  const deltaOf = (item: DriveItem): number | undefined => {
    if (reference === undefined) return undefined;
    const created = timeOf(item.createdDateTime);
    if (created === undefined) return undefined;
    return Math.abs(created - reference);
  };
  const withinTolerance = (item: DriveItem): boolean => {
    const delta = deltaOf(item);
    // No timestamps to compare on: do not use the tolerance as a veto.
    if (delta === undefined) return reference === undefined;
    return delta <= toleranceMs;
  };

  const best = (pool: DriveItem[]): DriveItem | null => {
    let winner: DriveItem | null = null;
    let winnerScore = Number.POSITIVE_INFINITY;
    for (const item of pool) {
      // Prefer the smallest delta; with no reference time, prefer the newest file.
      const created = timeOf(item.createdDateTime);
      const score = deltaOf(item) ?? (created === undefined ? Number.MAX_SAFE_INTEGER : -created);
      if (winner === null || score < winnerScore) {
        winner = item;
        winnerScore = score;
      }
    }
    return winner;
  };

  if (subject) {
    const named = candidates.filter(
      (item) => typeof item.name === "string" && normaliseName(item.name).startsWith(subject) && withinTolerance(item),
    );
    const match = best(named);
    if (match) return match;
  }

  if (reference === undefined) return null;
  return best(candidates.filter(withinTolerance));
}

/** Poll the recordings folder until the matching file shows up (OneDrive lags behind the notification). */
export async function findRecordingDriveItem(
  g: GraphClient,
  userId: string,
  opts: {
    subject?: string | null;
    createdDateTime?: string;
    folder: string;
    toleranceMinutes?: number;
    timeoutMs?: number;
    pollIntervalMs?: number;
    log?: Logger;
  },
): Promise<DriveItem | null> {
  const timeoutMs = opts.timeoutMs ?? DEFAULT_TIMEOUT_MS;
  const pollIntervalMs = opts.pollIntervalMs ?? DEFAULT_POLL_INTERVAL_MS;
  const deadline = Date.now() + timeoutMs;
  let attempt = 0;

  for (;;) {
    attempt++;
    try {
      const items = await listRecordingsFolder(g, userId, opts.folder);
      const match = pickRecordingItem(items, {
        subject: opts.subject,
        createdDateTime: opts.createdDateTime,
        toleranceMinutes: opts.toleranceMinutes,
      });
      if (match) {
        opts.log?.info("recording file located", { itemId: match.id, name: match.name, attempt });
        return match;
      }
      opts.log?.debug("recording file not in folder yet", { folder: opts.folder, attempt, candidates: items.length });
    } catch (err) {
      // A missing folder or a transient failure is expected while the file is still being written;
      // an auth/permission problem will not fix itself, so surface it immediately.
      if (err instanceof GraphError && (err.status === 401 || err.status === 403)) throw err;
      opts.log?.warn("listing recordings folder failed", { error: (err as Error).message, attempt });
      if (timeoutMs === 0) throw err;
    }

    if (timeoutMs === 0) return null;
    const remaining = deadline - Date.now();
    if (remaining <= 0) break;
    await sleep(Math.min(pollIntervalMs, remaining));
  }
  opts.log?.warn("timed out waiting for the recording file", { folder: opts.folder, timeoutMs });
  return null;
}

export async function getDriveItem(g: GraphClient, userId: string, itemId: string): Promise<DriveItem> {
  return g.get<DriveItem>(driveItemsPath(userId, itemId));
}

export async function listItemPermissions(
  g: GraphClient,
  userId: string,
  itemId: string,
): Promise<Permission[]> {
  return g.getAll<Permission>(`${driveItemsPath(userId, itemId)}/permissions`);
}

/** Strip the SharePoint claims prefix from a loginName, if present. */
export function stripClaimsPrefix(loginName: string): string {
  return loginName.startsWith(CLAIMS_PREFIX) ? loginName.slice(CLAIMS_PREFIX.length) : loginName;
}

/**
 * External guests are stored as `alice_contoso.com#EXT#@tenant.onmicrosoft.com`; recover the
 * original address so an already-invited guest is not invited again.
 */
function guestUpnToEmail(value: string): string | undefined {
  const idx = value.toUpperCase().indexOf("#EXT#");
  if (idx < 0) return undefined;
  const local = value.slice(0, idx);
  const at = local.lastIndexOf("_");
  if (at <= 0 || at === local.length - 1) return undefined;
  return `${local.slice(0, at)}@${local.slice(at + 1)}`;
}

/** Pure: every email address that already appears in the item's permissions, lowercased. */
export function emailsWithAccess(perms: Permission[]): Set<string> {
  const out = new Set<string>();
  const add = (value: string | null | undefined): void => {
    if (!value) return;
    const cleaned = stripClaimsPrefix(String(value).trim());
    if (!cleaned.includes("@")) return;
    out.add(cleaned.toLowerCase());
    const guest = guestUpnToEmail(cleaned);
    if (guest) out.add(guest.toLowerCase());
  };

  for (const perm of perms ?? []) {
    if (!perm) continue;
    add(perm.invitation?.email);
    const identities = [
      perm.grantedToV2?.user,
      perm.grantedToV2?.siteUser,
      ...(perm.grantedToIdentitiesV2 ?? []).flatMap((i) => [i?.user, i?.siteUser]),
    ];
    for (const identity of identities) {
      if (!identity) continue;
      add((identity as { email?: string }).email);
      add((identity as { loginName?: string }).loginName);
      add(identity.userPrincipalName);
    }
  }
  return out;
}

export interface InviteResult {
  granted: string[];
  failed: { email: string; code?: string; message: string }[];
  permissions: Permission[];
}

/** A permission as returned inside an `invite` response, which may carry a per-recipient error. */
type InvitePermission = Permission & { error?: { code?: string; message?: string } };

function emailOfEntry(entry: InvitePermission | undefined): string | undefined {
  if (!entry) return undefined;
  const candidates = [
    entry.invitation?.email,
    entry.grantedToV2?.siteUser?.email,
    entry.grantedToV2?.user?.userPrincipalName,
    ...(entry.grantedToIdentitiesV2 ?? []).flatMap((i) => [i?.siteUser?.email, i?.user?.userPrincipalName]),
  ];
  for (const candidate of candidates) {
    if (candidate && candidate.includes("@")) return candidate;
  }
  return undefined;
}

/**
 * Invite people to the item in batches of 20, tolerating 207 Multi-Status responses where some
 * recipients fail (typically brand-new external guests under app-only auth).
 */
export async function inviteRecipients(
  g: GraphClient,
  userId: string,
  itemId: string,
  emails: string[],
  opts: { roles: string[]; sendInvitation: boolean; message?: string; requireSignIn?: boolean },
): Promise<InviteResult> {
  const result: InviteResult = { granted: [], failed: [], permissions: [] };
  const unique = [...new Set(emails.map((e) => e.trim().toLowerCase()).filter(Boolean))];
  if (unique.length === 0) return result;

  for (const batch of chunk(unique, INVITE_BATCH_SIZE)) {
    const body = {
      recipients: batch.map((email) => ({ email })),
      requireSignIn: opts.requireSignIn ?? true,
      sendInvitation: opts.sendInvitation,
      roles: opts.roles,
      ...(opts.message ? { message: opts.message } : {}),
    };
    let entries: InvitePermission[];
    try {
      const res = await g.post<ODataCollection<InvitePermission> | undefined>(
        `${driveItemsPath(userId, itemId)}/invite`,
        body,
      );
      entries = res?.value ?? [];
    } catch (err) {
      // The whole batch failed: report it per recipient instead of throwing.
      const code = err instanceof GraphError ? err.code : undefined;
      const message = (err as Error).message;
      for (const email of batch) result.failed.push({ email, code, message });
      continue;
    }

    for (let i = 0; i < batch.length; i++) {
      const entry = entries[i];
      // Prefer the address Graph echoes back; fall back to positional mapping.
      const email = (emailOfEntry(entry) ?? batch[i]).toLowerCase();
      if (!entry) {
        result.failed.push({ email, message: "Graph returned no result for this recipient" });
        continue;
      }
      if (entry.error) {
        result.failed.push({
          email,
          code: entry.error.code,
          message: entry.error.message ?? entry.error.code ?? "invite failed",
        });
        continue;
      }
      result.granted.push(email);
      result.permissions.push(entry);
    }
    // Extra entries beyond the batch size are unexpected; keep their permissions for the caller.
    for (let i = batch.length; i < entries.length; i++) {
      const entry = entries[i];
      if (!entry || entry.error) continue;
      result.permissions.push(entry);
      const email = emailOfEntry(entry);
      if (email) result.granted.push(email.toLowerCase());
    }
  }
  return result;
}

export async function createSharingLink(
  g: GraphClient,
  userId: string,
  itemId: string,
  opts: { scope: "users" | "organization" | "anonymous"; type?: "view"; expirationDateTime?: string },
): Promise<Permission> {
  const body = {
    type: opts.type ?? "view",
    scope: opts.scope,
    ...(opts.expirationDateTime ? { expirationDateTime: opts.expirationDateTime } : {}),
  };
  return g.post<Permission>(`${driveItemsPath(userId, itemId)}/createLink`, body);
}

/** Grant an existing (scope: "users") sharing link to specific people. */
export async function grantLinkToRecipients(
  g: GraphClient,
  userId: string,
  itemId: string,
  permissionId: string,
  emails: string[],
  roles: string[] = ["read"],
): Promise<Permission[]> {
  const unique = [...new Set(emails.map((e) => e.trim().toLowerCase()).filter(Boolean))];
  const out: Permission[] = [];
  for (const batch of chunk(unique, INVITE_BATCH_SIZE)) {
    const res = await g.post<ODataCollection<Permission> | undefined>(
      `${driveItemsPath(userId, itemId)}/permissions/${enc(permissionId)}/grant`,
      { recipients: batch.map((email) => ({ email })), roles },
    );
    if (res?.value) out.push(...res.value);
  }
  return out;
}

export async function sendMail(
  g: GraphClient,
  userId: string,
  msg: { to: string[]; subject: string; html: string },
): Promise<void> {
  const toRecipients = [...new Set(msg.to.map((e) => e.trim()).filter(Boolean))].map((address) => ({
    emailAddress: { address },
  }));
  if (toRecipients.length === 0) return;
  await g.post<void>(`/users/${enc(userId)}/sendMail`, {
    message: {
      subject: msg.subject,
      body: { contentType: "HTML", content: msg.html },
      toRecipients,
    },
    saveToSentItems: true,
  });
}
