# Module contract (internal)

All modules are ESM TypeScript (`"type": "module"`, `moduleResolution: NodeNext`, so **relative imports must end in `.js`**).
Shared files that already exist and must NOT be modified by feature agents: `src/config.ts`, `src/log.ts`, `src/graph/types.ts`, `src/core/types.ts`.
If a shared type is missing something, add an *optional* field and note it in your report.

Verified Graph facts (do not re-research):
- Recording ready notification: subscribe `changeType: "created"` on
  - tenant-wide: `communications/onlineMeetings/getAllRecordings` (app permission `OnlineMeetingRecording.Read.All`)
  - per-organizer: `users/{userId}/onlineMeetings/getAllRecordings`
  Notification `resource` looks like `communications/onlineMeetings('MSp...')/recordings('MSp...')` or `users('guid')/onlineMeetings('...')/recordings('...')`.
  Max expiration 4320 minutes (3 days); `lifecycleNotificationUrl` required when expiration > 1 hour. Validation handshake: Graph GETs/POSTs `notificationUrl?validationToken=...`; respond 200 `text/plain` with the token within 10s. Notifications: respond 202 quickly, process async. Always check `clientState`.
- Online meeting id (`MSp...`) is base64 of `1*{organizerOid}*0**19:meeting_xxx@thread.v2`. Decode with base64 → split on `*`.
- `GET /users/{uid}/onlineMeetings/{mid}` ; `GET /users/{uid}/onlineMeetings?$filter=JoinWebUrl eq '{url}'`.
- `GET /users/{uid}/onlineMeetings/{mid}/recordings` and `/recordings/{rid}` → callRecording {id, meetingId, callId, createdDateTime, endDateTime, recordingContentUrl, meetingOrganizer}.
- `GET /users/{uid}/onlineMeetings/{mid}/attendanceReports` → then `/attendanceReports/{rid}/attendanceRecords` → {emailAddress, identity{id,displayName,tenantId}, role, totalAttendanceInSeconds}. Permission `OnlineMeetingArtifact.Read.All`. App-only calls to onlineMeetings/* need a Teams application access policy granted to the organizer.
- Recording file lands in organizer's OneDrive folder `Recordings`, name like `{Subject}-{yyyyMMdd_HHmmss}UTC-Meeting Recording.mp4`. List: `GET /users/{uid}/drive/root:/Recordings:/children?$orderby=createdDateTime desc&$top=50`.
- `GET /users/{uid}/drive/items/{itemId}/permissions` → permission {roles, grantedToV2.user{id,displayName}, grantedToV2.siteUser{email,loginName}, grantedToIdentitiesV2[], invitation.email, link{scope,type,webUrl}}.
- `POST /users/{uid}/drive/items/{itemId}/invite` body `{recipients:[{email}], requireSignIn:true, sendInvitation:bool, roles:["read"], message}` → 200 `{value:[permission]}` or **207 Multi-Status** with per-recipient `error` objects (e.g. code `accountVerificationRequired`). Constraint: *new* external guests cannot be invited with app-only auth (existing guests can); delegated auth as the organizer can invite them subject to the tenant's external-sharing policy.
- `POST /users/{uid}/drive/items/{itemId}/createLink` body `{type:"view", scope:"users"|"organization"|"anonymous", expirationDateTime?}` → permission with `link.webUrl`. (For `scope:"users"` the recipients are then granted via `invite`-like `recipients` — treat `users`-scope link as: create link, then `POST .../permissions/{permId}/grant` with `{recipients:[{email}], roles:["read"]}`.)
- `POST /users/{uid}/sendMail` body `{message:{subject, body:{contentType:"HTML", content}, toRecipients:[{emailAddress:{address}}]}, saveToSentItems:true}` → 202.
- Calendar invitees: `GET /users/{uid}/calendarView?startDateTime=...&endDateTime=...&$select=id,subject,start,end,isOnlineMeeting,onlineMeeting,attendees,organizer` → match `onlineMeeting.joinUrl` to the meeting's `joinWebUrl`.
- Subscriptions: `POST /subscriptions`, `PATCH /subscriptions/{id}` `{expirationDateTime}`, `DELETE /subscriptions/{id}`, `GET /subscriptions`.

## src/auth.ts  (Agent A)
```ts
export interface TokenProvider { getToken(): Promise<string>; /** account username when delegated */ account?: () => Promise<string | undefined> }
export function createAppTokenProvider(cfg: AppConfig): TokenProvider          // msal-node ConfidentialClientApplication, scope `${cfg.graphBaseUrl}/.default`
export function createDelegatedTokenProvider(cfg: AppConfig, opts?: { onDeviceCode?: (message: string) => void; allowInteractive?: boolean }): TokenProvider
  // msal-node PublicClientApplication with ICachePlugin persisting cfg.tokenCachePath (mode 0600, mkdir -p). acquireTokenSilent first; if it fails and allowInteractive, run device code; else throw a clear "run `npm run login`" error.
export function tokenProviderFromConfig(cfg: AppConfig, opts?): TokenProvider  // picks by cfg.authMode
export async function login(cfg: AppConfig, log: Logger): Promise<{ username?: string }> // forces device-code flow and persists cache
```
## src/graph/client.ts (Agent A)
```ts
export class GraphError extends Error { status: number; code?: string; body?: unknown; retryAfterMs?: number }
export interface RequestOptions { query?: Record<string, string | number | boolean | undefined>; headers?: Record<string,string>; version?: "v1.0" | "beta"; /** return raw Response instead of parsed JSON */ raw?: boolean }
export class GraphClient {
  constructor(tokens: TokenProvider, opts?: { baseUrl?: string; fetchImpl?: typeof fetch; logger?: Logger; maxRetries?: number });
  get<T>(path: string, opts?: RequestOptions): Promise<T>;
  post<T>(path: string, body: unknown, opts?: RequestOptions): Promise<T>;
  patch<T>(path: string, body: unknown, opts?: RequestOptions): Promise<T>;
  delete(path: string, opts?: RequestOptions): Promise<void>;
  /** Follows @odata.nextLink. */
  getAll<T>(path: string, opts?: RequestOptions): Promise<T[]>;
}
```
Rules: `path` may be relative (`/users/x`) or absolute (`https://graph.microsoft.com/v1.0/...` – used for nextLink). Retry 429/503/504 honoring `Retry-After` (max 4), exponential backoff otherwise. 204 → undefined. Non-2xx → GraphError with parsed `error.code`/`error.message`. 207 → return parsed body (callers inspect). Never log tokens.

## src/graph/drive.ts (Agent A)
```ts
export async function listRecordingsFolder(g: GraphClient, userId: string, folder: string, top?: number): Promise<DriveItem[]>
export function pickRecordingItem(items: DriveItem[], opts: { subject?: string | null; createdDateTime?: string; toleranceMinutes?: number }): DriveItem | null  // pure: prefer .mp4 whose name starts with subject (case/space-insensitive, Teams strips chars like / \ : ? * " < > |) and whose createdDateTime is within tolerance of recording createdDateTime; else closest-in-time mp4 within tolerance; else null
export async function findRecordingDriveItem(g, userId, opts: { subject?: string | null; createdDateTime?: string; folder: string; toleranceMinutes?: number; timeoutMs?: number; pollIntervalMs?: number; log?: Logger }): Promise<DriveItem | null> // polls until found or timeout
export async function getDriveItem(g, userId, itemId): Promise<DriveItem>
export async function listItemPermissions(g, userId, itemId): Promise<Permission[]>
export function emailsWithAccess(perms: Permission[]): Set<string>   // pure, lowercased; from grantedToV2.siteUser.email/loginName (strip "i:0#.f|membership|"), grantedToIdentitiesV2[*].siteUser.email, invitation.email, grantedToV2.user.email if present
export interface InviteResult { granted: string[]; failed: { email: string; code?: string; message: string }[]; permissions: Permission[] }
export async function inviteRecipients(g, userId, itemId, emails: string[], opts: { roles: string[]; sendInvitation: boolean; message?: string; requireSignIn?: boolean }): Promise<InviteResult> // batches of 20; handles 200 and 207; maps per-recipient errors by index/email
export async function createSharingLink(g, userId, itemId, opts: { scope: "users" | "organization" | "anonymous"; type?: "view"; expirationDateTime?: string }): Promise<Permission>
export async function grantLinkToRecipients(g, userId, itemId, permissionId: string, emails: string[], roles?: string[]): Promise<Permission[]>  // POST .../permissions/{id}/grant
export async function sendMail(g, userId, msg: { to: string[]; subject: string; html: string }): Promise<void>
```

## src/graph/meetings.ts (Agent B)
```ts
export function decodeOnlineMeetingId(id: string): { organizerId: string; threadId: string } | null
export function parseNotificationResource(resource: string): { userId?: string; meetingId?: string; recordingId?: string; transcriptId?: string } | null
export async function getOnlineMeeting(g, userId, meetingId): Promise<OnlineMeeting>
export async function findOnlineMeetingByJoinUrl(g, userId, joinWebUrl): Promise<OnlineMeeting | null>
export async function listRecordings(g, userId, meetingId): Promise<CallRecording[]>
export async function getRecording(g, userId, meetingId, recordingId): Promise<CallRecording>
export async function listAttendees(g, userId, meetingId, log?): Promise<Attendee[]>   // union of all attendanceReports' records; dedupe by email; skip records without email (log them); source "attendance"
export async function listCalendarInvitees(g, userId, meeting: OnlineMeeting, log?): Promise<Attendee[]> // source "calendar"; [] if no match
export function normaliseEmail(e: string | null | undefined): string | null
export function isExternal(email: string, internalDomains: string[]): boolean
export function mergeAttendees(...lists: Attendee[][]): Attendee[]   // pure, first occurrence wins, dedupe by email
```
## src/core/share.ts (Agent B)
```ts
export async function shareRecordingWithAttendees(g: GraphClient, cfg: AppConfig, req: ShareRequest, log: Logger): Promise<ShareResult>
```
Pipeline: resolve meeting (id or joinWebUrl) → recordings (one or all) → attendees (attendance + organizer + optional calendar invitees + cfg.alwaysInclude + req.extraEmails; mark isExternal using cfg.internalDomains or organizer's UPN domain) → for each recording: locate drive item (req.driveItemId, else findRecordingDriveItem with cfg.driveLookupTimeoutMinutes; timeout → error entry) → existing permissions → invite missing (unless dryRun) → for failures: per cfg.externalFallback create link (`users` → createSharingLink+grantLinkToRecipients; `anonymous` → createSharingLink + sendMail with link; `skip`) → build ShareResult. Never throw for per-person failures; throw only if meeting/recording cannot be resolved.

## src/mcp/server.ts + src/cli.ts (Agent C)
MCP (stdio) tools, each returning `content:[{type:"text", text: JSON}]` and `structuredContent`:
- `share_recording` (ShareRequest fields; organizerUserId optional → defaults to "me" resolved via `/me` when delegated)
- `list_meeting_attendees` {organizerUserId?, meetingId?|joinWebUrl, includeCalendarInvitees?}
- `find_recording` {organizerUserId?, meetingId?|joinWebUrl} → recordings + matched drive items
- `list_recent_recordings` {organizerUserId?, top?} → recent items in the Recordings folder
- `get_recording_access` {organizerUserId?, driveItemId} → who has access
- `whoami` → account / auth mode
CLI (`src/cli.ts`): `login`, `share --meeting <id>|--join-url <url> [--organizer <id>] [--dry-run] [--extra a@b.com]`, `attendees ...`, `whoami`. Uses `loadDotEnv()` then `loadConfig()`.

## src/service/* (Agent D)
- `store.ts`: JSON file state in cfg.stateDir: subscriptions {id, resource, expirationDateTime}, processed recording ids, last errors.
- `subscribe.ts` (script + exported fns): `ensureSubscriptions(g, cfg, store, log)` creates/renews `getAllRecordings` subscriptions (tenant-wide or per organizer), `renewSubscriptions`, `removeSubscriptions`. CLI: `npm run subscribe -- create|renew|list|delete`.
- `server.ts`: Express 5 app. `POST /webhook/notifications`: validationToken handshake; else verify clientState, 202 immediately, enqueue. `POST /webhook/lifecycle`: handshake; reauthorizationRequired → renew; subscriptionRemoved → recreate. `GET /healthz`. Worker: for each notification with `#Microsoft.Graph.callRecording`, parseNotificationResource → organizer from resource user id or decodeOnlineMeetingId → dedupe via store → `shareRecordingWithAttendees` → log result. Renewal timer every 6h. Exports `createApp(deps)` for tests.

## tests/ + docs (Agent E)
`node --test` via tsx. Unit tests for pure functions (decodeOnlineMeetingId, parseNotificationResource, pickRecordingItem, emailsWithAccess, mergeAttendees, isExternal, loadConfig) and for `createApp` handshake using supertest-free `fetch` against `app.listen(0)`. README.md (top-level) + docs/SETUP.md (Entra app registration, permissions tables for both auth modes, application access policy PowerShell, external sharing policy notes, ngrok/Azure deployment, MCP client config for Claude Desktop / Claude Code), `.env.example`.
