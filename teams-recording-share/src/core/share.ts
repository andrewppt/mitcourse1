/**
 * The sharing pipeline: resolve a meeting and its recordings, work out who should have access,
 * then grant it on the recording file in the organizer's OneDrive.
 *
 * Per-person failures never abort the run: they are collected into the result. Only a meeting or
 * recording that cannot be resolved at all throws.
 */
import type { AppConfig } from "../config.js";
import type { Logger } from "../log.js";
import type { GraphClient } from "../graph/client.js";
import type { Attendee, DriveItem, OnlineMeeting, Permission } from "../graph/types.js";
import type { ShareRequest, ShareResult, ShareOutcomePerson } from "./types.js";
import {
  createSharingLink,
  emailsWithAccess,
  findRecordingDriveItem,
  getDriveItem,
  grantLinkToRecipients,
  inviteRecipients,
  listItemPermissions,
  sendMail,
} from "../graph/drive.js";
import {
  findOnlineMeetingByJoinUrl,
  getOnlineMeeting,
  getRecording,
  isExternal,
  listAttendees,
  listCalendarInvitees,
  listRecordings,
  mergeAttendees,
  normaliseEmail,
} from "../graph/meetings.js";

interface OrganizerUser {
  id: string;
  displayName?: string;
  userPrincipalName?: string;
  mail?: string | null;
}

type RecordingResult = ShareResult["recordings"][number];

const DRY_RUN_WARNING = "dry run: nothing was changed";

function errMessage(e: unknown): string {
  if (e instanceof Error) return e.message;
  if (typeof e === "string") return e;
  try {
    return JSON.stringify(e);
  } catch {
    return String(e);
  }
}

function errCode(e: unknown): string | undefined {
  const c = (e as { code?: unknown } | null | undefined)?.code;
  return typeof c === "string" && c ? c : undefined;
}

function person(a: Attendee): ShareOutcomePerson {
  return { email: a.email, displayName: a.displayName, source: a.source, isExternal: a.isExternal };
}

function escapeHtml(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function domainOf(email: string | null | undefined): string | null {
  const e = normaliseEmail(email);
  return e ? e.slice(e.lastIndexOf("@") + 1) : null;
}

export async function shareRecordingWithAttendees(
  g: GraphClient,
  cfg: AppConfig,
  req: ShareRequest,
  log: Logger,
): Promise<ShareResult> {
  const dryRun = Boolean(req.dryRun);
  const warnings: string[] = [];
  const addWarning = (w: string): void => {
    if (!warnings.includes(w)) warnings.push(w);
  };

  /* ---------------- organizer ---------------- */
  const requested = (req.organizerUserId ?? "").trim() || "me";
  const isMe = requested.toLowerCase() === "me";
  if (isMe && cfg.authMode !== "delegated") {
    throw new Error('organizerUserId "me" can only be resolved with AUTH_MODE=delegated; pass the organizer\'s user id or UPN');
  }

  const select = { $select: "id,userPrincipalName,mail,displayName" } as const;
  let organizer: OrganizerUser | null = null;
  try {
    organizer = isMe
      ? await g.get<OrganizerUser>("/me", { query: { ...select } })
      : await g.get<OrganizerUser>(`/users/${encodeURIComponent(requested)}`, { query: { ...select } });
  } catch (e) {
    if (isMe) throw new Error(`Could not resolve the signed-in user via GET /me: ${errMessage(e)}`);
    log.warn("could not read the organizer user object", { organizer: requested, error: errMessage(e) });
    addWarning(`Could not read the organizer user object (${errMessage(e)}); continuing with the id as given`);
  }

  const organizerUserId = organizer?.id || requested;
  const organizerEmail = normaliseEmail(organizer?.mail ?? organizer?.userPrincipalName);

  const internalDomains =
    cfg.internalDomains.length > 0
      ? cfg.internalDomains
      : [domainOf(organizer?.userPrincipalName ?? organizer?.mail)].filter((d): d is string => Boolean(d));
  if (internalDomains.length === 0) {
    addWarning("No internal domains configured and the organizer's domain could not be determined; nobody is marked external");
  }

  /* ---------------- meeting ---------------- */
  let meeting: OnlineMeeting | null = null;
  if (req.meetingId) {
    try {
      meeting = await getOnlineMeeting(g, organizerUserId, req.meetingId);
    } catch (e) {
      throw new Error(`Could not load online meeting "${req.meetingId}" for user ${organizerUserId}: ${errMessage(e)}`);
    }
  } else if (req.joinWebUrl) {
    try {
      meeting = await findOnlineMeetingByJoinUrl(g, organizerUserId, req.joinWebUrl);
    } catch (e) {
      throw new Error(`Could not look up the online meeting by join URL for user ${organizerUserId}: ${errMessage(e)}`);
    }
    if (!meeting) throw new Error(`No online meeting found for join URL "${req.joinWebUrl}" under user ${organizerUserId}`);
  } else {
    throw new Error("Either meetingId or joinWebUrl must be provided");
  }
  if (!meeting) throw new Error("Could not resolve the online meeting");

  const meetingInfo = {
    id: meeting.id,
    subject: meeting.subject ?? null,
    joinWebUrl: meeting.joinWebUrl,
    organizerUserId,
  };

  /* ---------------- recordings ---------------- */
  let recordings;
  if (req.recordingId) {
    try {
      recordings = [await getRecording(g, organizerUserId, meeting.id, req.recordingId)];
    } catch (e) {
      throw new Error(`Could not load recording "${req.recordingId}" of meeting ${meeting.id}: ${errMessage(e)}`);
    }
  } else {
    try {
      recordings = await listRecordings(g, organizerUserId, meeting.id);
    } catch (e) {
      throw new Error(`Could not list recordings of meeting ${meeting.id}: ${errMessage(e)}`);
    }
  }

  /* ---------------- attendees ---------------- */
  let attendance: Attendee[] = [];
  try {
    attendance = await listAttendees(g, organizerUserId, meeting.id, log);
  } catch (e) {
    log.warn("could not read attendance reports", { meetingId: meeting.id, error: errMessage(e) });
    addWarning(`Could not read attendance reports (${errMessage(e)})`);
  }

  const organizerAttendee: Attendee[] = organizerEmail
    ? [
        {
          email: organizerEmail,
          displayName: organizer?.displayName,
          source: "organizer",
          role: "Organizer",
          userId: organizer?.id,
        },
      ]
    : [];
  if (!organizerEmail) addWarning("The organizer's email address could not be determined");

  let calendar: Attendee[] = [];
  if (cfg.includeCalendarInvitees) {
    try {
      calendar = await listCalendarInvitees(g, organizerUserId, meeting, log);
    } catch (e) {
      log.warn("could not read calendar invitees", { meetingId: meeting.id, error: errMessage(e) });
      addWarning(`Could not read calendar invitees (${errMessage(e)})`);
    }
  }

  const extras: Attendee[] = [...cfg.alwaysInclude, ...(req.extraEmails ?? [])]
    .map((e) => normaliseEmail(e))
    .filter((e): e is string => Boolean(e))
    .map((email) => ({ email, source: "extra" as const }));

  const attendees = mergeAttendees(attendance, organizerAttendee, calendar, extras);
  for (const a of attendees) a.isExternal = isExternal(a.email, internalDomains);
  const byEmail = new Map(attendees.map((a) => [a.email, a]));

  if (recordings.length === 0) {
    log.info("no recordings found for meeting", { meetingId: meeting.id });
    return {
      dryRun,
      meeting: meetingInfo,
      recordings: [],
      attendees,
      warnings: [...warnings, "No recordings found for this meeting yet"],
    };
  }

  /* ---------------- per recording ---------------- */
  const results: RecordingResult[] = [];
  for (const recording of recordings) {
    const out: RecordingResult = {
      recordingId: recording.id,
      createdDateTime: recording.createdDateTime,
      granted: [],
      alreadyHadAccess: [],
      skipped: [],
      errors: [],
    };
    results.push(out);

    /* locate the OneDrive file */
    let item: DriveItem | null = null;
    try {
      item = req.driveItemId
        ? await getDriveItem(g, organizerUserId, req.driveItemId)
        : await findRecordingDriveItem(g, organizerUserId, {
            subject: meeting.subject,
            createdDateTime: recording.createdDateTime,
            folder: cfg.recordingsFolder,
            timeoutMs: dryRun ? 0 : cfg.driveLookupTimeoutMinutes * 60_000,
            log,
          });
    } catch (e) {
      log.error("drive item lookup failed", { recordingId: recording.id, error: errMessage(e) });
      out.errors.push({ message: `Could not locate the recording file in OneDrive: ${errMessage(e)}` });
      continue;
    }
    if (!item) {
      out.errors.push({
        message:
          `Recording file not found in OneDrive folder "${cfg.recordingsFolder}" of ${organizerUserId} ` +
          `(subject "${meeting.subject ?? ""}", recorded ${recording.createdDateTime ?? "unknown"}). ` +
          `Pass driveItemId to point at the file directly.`,
      });
      continue;
    }
    out.driveItem = { id: item.id, name: item.name, webUrl: item.webUrl };

    /* who already has access */
    let existing = new Set<string>();
    try {
      const perms: Permission[] = await listItemPermissions(g, organizerUserId, item.id);
      existing = emailsWithAccess(perms);
    } catch (e) {
      log.warn("could not list existing permissions", { itemId: item.id, error: errMessage(e) });
      out.errors.push({ message: `Could not list existing permissions on the file: ${errMessage(e)}` });
    }

    const toInvite: Attendee[] = [];
    for (const a of attendees) {
      if (organizerEmail && a.email === organizerEmail) {
        out.alreadyHadAccess.push(person(a)); // owns the file
        continue;
      }
      if (existing.has(a.email)) out.alreadyHadAccess.push(person(a));
      else toInvite.push(a);
    }

    if (toInvite.length === 0) {
      log.info("everyone already has access", { itemId: item.id, recordingId: recording.id });
      continue;
    }

    if (dryRun) {
      out.granted = toInvite.map(person);
      addWarning(DRY_RUN_WARNING);
      continue;
    }

    /* invite */
    let failed: { email: string; code?: string; message: string }[] = [];
    try {
      const res = await inviteRecipients(
        g,
        organizerUserId,
        item.id,
        toInvite.map((a) => a.email),
        {
          roles: [cfg.shareRole],
          sendInvitation: cfg.sendInvitation,
          message: cfg.invitationMessage,
          requireSignIn: true,
        },
      );
      for (const email of res.granted) {
        const a = byEmail.get(email);
        out.granted.push(a ? person(a) : { email, source: "extra" });
      }
      failed = res.failed;
    } catch (e) {
      log.error("invite call failed", { itemId: item.id, error: errMessage(e) });
      const code = errCode(e);
      const message = errMessage(e);
      failed = toInvite.map((a) => ({ email: a.email, code, message }));
      out.errors.push({ message: `Invite request failed: ${message}` });
    }

    for (const f of failed) {
      out.errors.push({ email: f.email, message: f.message });
    }
    const failedEmails = [...new Set(failed.map((f) => f.email))];
    if (failedEmails.length === 0) continue;

    const reasonFor = (email: string): string => {
      const f = failed.find((x) => x.email === email);
      const code = f?.code ? f.code : f?.message;
      return `external recipient could not be invited (${code ?? "unknown error"})`;
    };

    /* fallback */
    if (cfg.externalFallback === "skip") {
      for (const email of failedEmails) out.skipped.push({ email, reason: reasonFor(email) });
      continue;
    }

    if (cfg.externalFallback === "users-link") {
      let link: Permission | null = null;
      try {
        link = await createSharingLink(g, organizerUserId, item.id, { scope: "users", type: "view" });
      } catch (e) {
        log.error("could not create users-scoped sharing link", { itemId: item.id, error: errMessage(e) });
        out.errors.push({ message: `Could not create a sharing link: ${errMessage(e)}` });
        for (const email of failedEmails) out.skipped.push({ email, reason: `${reasonFor(email)}; sharing link failed: ${errMessage(e)}` });
        continue;
      }
      if (!link) {
        for (const email of failedEmails) out.skipped.push({ email, reason: `${reasonFor(email)}; no sharing link was returned` });
        continue;
      }
      if (link.link?.webUrl) {
        out.fallbackLink = { webUrl: link.link.webUrl, scope: link.link.scope ?? "users", sentTo: [] };
        if (link.expirationDateTime) out.fallbackLink.expirationDateTime = link.expirationDateTime;
      }
      try {
        await grantLinkToRecipients(g, organizerUserId, item.id, link.id, failedEmails, [cfg.shareRole]);
        for (const email of failedEmails) {
          const a = byEmail.get(email);
          out.granted.push(a ? person(a) : { email, source: "extra" });
        }
        if (out.fallbackLink) out.fallbackLink.sentTo = [...failedEmails];
      } catch (e) {
        log.error("could not grant the sharing link", { itemId: item.id, error: errMessage(e) });
        out.errors.push({ message: `Could not grant the sharing link: ${errMessage(e)}` });
        for (const email of failedEmails) out.skipped.push({ email, reason: `${reasonFor(email)}; link grant failed: ${errMessage(e)}` });
      }
      continue;
    }

    /* anonymous-link */
    let link: Permission | null = null;
    const expirationDateTime =
      cfg.fallbackLinkExpiryHours > 0
        ? new Date(Date.now() + cfg.fallbackLinkExpiryHours * 60 * 60 * 1000).toISOString()
        : undefined;
    try {
      link = await createSharingLink(g, organizerUserId, item.id, {
        scope: "anonymous",
        type: "view",
        expirationDateTime,
      });
    } catch (e) {
      log.error("could not create anonymous sharing link", { itemId: item.id, error: errMessage(e) });
      out.errors.push({ message: `Could not create an anonymous sharing link: ${errMessage(e)}` });
      for (const email of failedEmails) out.skipped.push({ email, reason: `${reasonFor(email)}; sharing link failed: ${errMessage(e)}` });
      continue;
    }

    const webUrl = link?.link?.webUrl;
    if (!link || !webUrl) {
      out.errors.push({ message: "The anonymous sharing link was created but carries no webUrl" });
      for (const email of failedEmails) out.skipped.push({ email, reason: `${reasonFor(email)}; sharing link had no URL` });
      continue;
    }
    out.fallbackLink = { webUrl, scope: link.link?.scope ?? "anonymous", sentTo: [] };
    if (link.expirationDateTime ?? expirationDateTime) {
      out.fallbackLink.expirationDateTime = link.expirationDateTime ?? expirationDateTime;
    }

    const subject = meeting.subject ?? "our meeting";
    try {
      await sendMail(g, organizerUserId, {
        to: failedEmails,
        subject: `Recording: ${subject}`,
        html:
          `<p>${escapeHtml(cfg.invitationMessage)}</p>` +
          `<p>Meeting: <strong>${escapeHtml(subject)}</strong></p>` +
          `<p><a href="${escapeHtml(webUrl)}">Open the recording</a></p>` +
          (expirationDateTime ? `<p>This link expires on ${escapeHtml(expirationDateTime)}.</p>` : ""),
      });
      out.fallbackLink.sentTo = failedEmails;
    } catch (e) {
      log.error("could not send the fallback link by email", { itemId: item.id, error: errMessage(e) });
      out.errors.push({ message: `Could not email the sharing link: ${errMessage(e)}` });
      for (const email of failedEmails) out.skipped.push({ email, reason: `${reasonFor(email)}; sending the link failed: ${errMessage(e)}` });
    }
  }

  log.info("share complete", {
    meetingId: meeting.id,
    dryRun,
    recordings: results.length,
    attendees: attendees.length,
  });

  return { dryRun, meeting: meetingInfo, recordings: results, attendees, warnings };
}
