import type { Attendee } from "../graph/types.js";

export interface ShareRequest {
  /** Organizer's user object id or UPN. Required: recordings live in the organizer's OneDrive. */
  organizerUserId: string;
  /** Graph onlineMeeting id (base64 "MSp..." form). Either this or joinWebUrl must be given. */
  meetingId?: string;
  joinWebUrl?: string;
  /** Specific callRecording id; when omitted, every recording of the meeting is processed. */
  recordingId?: string;
  /** Explicit OneDrive item id, skipping the file search (useful when auto-matching fails). */
  driveItemId?: string;
  /** Extra emails to include beyond attendees. */
  extraEmails?: string[];
  /** Compute the plan without granting anything. */
  dryRun?: boolean;
}

export interface ShareOutcomePerson { email: string; displayName?: string; source: Attendee["source"]; isExternal?: boolean; }

export interface ShareResult {
  dryRun: boolean;
  meeting: { id: string; subject?: string | null; joinWebUrl?: string; organizerUserId: string };
  recordings: {
    recordingId: string;
    createdDateTime?: string;
    driveItem?: { id: string; name: string; webUrl?: string };
    /** People granted direct access in this run. */
    granted: ShareOutcomePerson[];
    /** People who already had access (no change). */
    alreadyHadAccess: ShareOutcomePerson[];
    /** Fallback sharing link created for people `invite` could not add directly. */
    fallbackLink?: { webUrl: string; scope: string; sentTo: string[]; expirationDateTime?: string };
    skipped: { email: string; reason: string }[];
    errors: { email?: string; message: string }[];
  }[];
  attendees: Attendee[];
  warnings: string[];
}
