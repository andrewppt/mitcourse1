/** Subset of Microsoft Graph resource shapes used by this project. Field names match Graph exactly. */

export interface Identity { id?: string; displayName?: string; tenantId?: string; userPrincipalName?: string; }
export interface IdentitySet { user?: Identity; application?: Identity; device?: Identity; guest?: Identity; }

export interface OnlineMeeting {
  id: string;
  subject?: string | null;
  joinWebUrl?: string;
  startDateTime?: string;
  endDateTime?: string;
  participants?: { organizer?: { identity?: IdentitySet; upn?: string }; attendees?: { identity?: IdentitySet; upn?: string; role?: string }[] };
  joinMeetingIdSettings?: { joinMeetingId?: string };
}

export interface CallRecording {
  id: string;
  meetingId?: string;
  callId?: string;
  contentCorrelationId?: string;
  createdDateTime?: string;
  endDateTime?: string;
  recordingContentUrl?: string;
  meetingOrganizer?: IdentitySet;
}

export interface MeetingAttendanceReport {
  id: string;
  totalParticipantCount?: number;
  meetingStartDateTime?: string;
  meetingEndDateTime?: string;
}

export interface AttendanceRecord {
  id?: string;
  emailAddress?: string | null;
  totalAttendanceInSeconds?: number;
  role?: string;
  identity?: Identity & { "@odata.type"?: string };
  attendanceIntervals?: { joinDateTime?: string; leaveDateTime?: string; durationInSeconds?: number }[];
}

export interface CalendarEvent {
  id: string;
  subject?: string;
  start?: { dateTime: string; timeZone: string };
  end?: { dateTime: string; timeZone: string };
  isOnlineMeeting?: boolean;
  onlineMeeting?: { joinUrl?: string } | null;
  organizer?: { emailAddress?: { name?: string; address?: string } };
  attendees?: { type?: string; emailAddress?: { name?: string; address?: string }; status?: { response?: string } }[];
}

export interface DriveItem {
  id: string;
  name: string;
  webUrl?: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  size?: number;
  file?: { mimeType?: string };
  folder?: { childCount?: number };
  parentReference?: { driveId?: string; id?: string; path?: string };
}

export interface Permission {
  id: string;
  roles?: string[];
  grantedToV2?: { user?: Identity; siteUser?: Identity & { email?: string; loginName?: string }; group?: Identity; siteGroup?: Identity };
  grantedToIdentitiesV2?: { user?: Identity; siteUser?: Identity & { email?: string; loginName?: string } }[];
  invitation?: { email?: string; signInRequired?: boolean; invitedBy?: IdentitySet };
  link?: { type?: string; scope?: string; webUrl?: string; preventsDownload?: boolean };
  expirationDateTime?: string | null;
  hasPassword?: boolean;
}

export interface Subscription {
  id: string;
  resource: string;
  changeType: string;
  notificationUrl: string;
  lifecycleNotificationUrl?: string;
  expirationDateTime: string;
  clientState?: string;
  applicationId?: string;
  creatorId?: string;
}

export interface ChangeNotification {
  subscriptionId: string;
  subscriptionExpirationDateTime?: string;
  changeType: "created" | "updated" | "deleted";
  clientState?: string;
  /** e.g. "communications/onlineMeetings('MSp...')/recordings('...')" or "users('...')/onlineMeetings('...')/recordings('...')" */
  resource: string;
  resourceData?: { id?: string; "@odata.type"?: string; "@odata.id"?: string; [k: string]: unknown };
  tenantId?: string;
  /** Lifecycle notifications only */
  lifecycleEvent?: "reauthorizationRequired" | "subscriptionRemoved" | "missed";
}

export interface ODataCollection<T> { value: T[]; "@odata.nextLink"?: string; "@odata.deltaLink"?: string; }

/** Normalised person that should receive access. */
export interface Attendee {
  email: string;           // lowercased
  displayName?: string;
  /** Where this attendee came from. */
  source: "attendance" | "calendar" | "organizer" | "extra";
  role?: string;           // Organizer | Presenter | Attendee | ...
  totalAttendanceInSeconds?: number;
  /** Entra object id when known (tenant users). */
  userId?: string;
  /** True when the email domain is not one of the organizer's internal domains. */
  isExternal?: boolean;
}
