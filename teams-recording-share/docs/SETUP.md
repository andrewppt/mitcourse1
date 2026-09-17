# Setup

Everything you need to do in Microsoft 365 before `teams-recording-share` can do anything, in the
order you need to do it. Steps 1–4 need a Global Administrator (or an Application Administrator plus
a Teams Administrator and a SharePoint Administrator).

Two decisions up front:

- **Auth mode.** `delegated` (browser sign-in as a person; the CLI and MCP server) or `app`
  (client credentials, no user; the webhook service). You can register one app that does both.
- **Scope.** The webhook service can watch the whole tenant (one subscription on
  `communications/onlineMeetings/getAllRecordings`) or only named organizers
  (`users/{id}/onlineMeetings/getAllRecordings`, one subscription each). Tenant-wide is app-only and
  needs the Teams application access policy granted globally; per-organizer is the safer start.

## 1. Register the application

Entra admin center → **Identity** → **Applications** → **App registrations** → **New registration**.

| Field | Value |
| --- | --- |
| Name | `teams-recording-share` |
| Supported account types | **Accounts in this organizational directory only** (single tenant) |
| Redirect URI | platform **Mobile and desktop applications**, URI `http://localhost` |

Then, in the new registration:

- **Authentication** → confirm `http://localhost` is listed under **Mobile and desktop applications**.
  `npm run login` runs the authorization-code flow with PKCE: it starts a listener on a random
  loopback port, opens the browser, and receives the code on `http://localhost:<port>`. Entra accepts
  any port for a registered `http://localhost` redirect, so no port needs to be fixed (set
  `LOGIN_REDIRECT_PORT` if your policy requires one, and register that exact URI).
  **Allow public client flows** can stay **No**; it is only needed for `LOGIN_FLOW=device-code`,
  which many tenants disable.
- **Overview** → copy **Application (client) ID** → `CLIENT_ID`, and **Directory (tenant) ID** →
  `TENANT_ID`.
- Only if you will run `AUTH_MODE=app`: **Certificates & secrets** → **New client secret** → copy the
  *Value* immediately → `CLIENT_SECRET`. Note the expiry date; the service stops working when the
  secret expires.

## 2. API permissions

**API permissions** → **Add a permission** → **Microsoft Graph**.

### Delegated mode (`AUTH_MODE=delegated`)

Add these as **Delegated permissions**. The signed-in user must be the meeting organizer (or have
access to the organizer's OneDrive).

| Permission | Why |
| --- | --- |
| `OnlineMeetings.Read` | Read the meeting by id or join URL. |
| `OnlineMeetingRecording.Read.All` | List the meeting's recordings. |
| `OnlineMeetingArtifact.Read.All` | Read attendance reports and attendance records. |
| `Files.ReadWrite` | Find the recording file and manage its permissions (invite, links). |
| `Calendars.Read` | Match the calendar event to add invitees (`INCLUDE_CALENDAR_INVITEES`). |
| `Mail.Send` | Email the fallback link (`EXTERNAL_FALLBACK=anonymous-link`). |
| `User.Read` | Resolve the signed-in user (`whoami`, default organizer). |
| `offline_access` | Refresh tokens, so you sign in once rather than daily. |

`OnlineMeetingRecording.Read.All` and `OnlineMeetingArtifact.Read.All` require **admin consent** even
as delegated permissions. Click **Grant admin consent for \<tenant\>** after adding them.

### App-only mode (`AUTH_MODE=app`)

Add these as **Application permissions**, then **Grant admin consent** — application permissions
never work without it.

| Permission | Why |
| --- | --- |
| `OnlineMeetingRecording.Read.All` | Subscribe to `getAllRecordings` and read recordings. |
| `OnlineMeetingArtifact.Read.All` | Read attendance reports and records. |
| `OnlineMeetings.Read.All` | Read the meeting itself. |
| `Files.ReadWrite.All` | Read and share files in the organizers' OneDrive. |
| `Calendars.Read` | Calendar invitees (optional). |
| `Mail.Send` | Email the fallback link (optional). |
| `User.Read.All` | Resolve organizers by id or UPN, and their domain. |

`Files.ReadWrite.All`, `Calendars.Read`, `Mail.Send` and `User.Read.All` are tenant-wide application
permissions. If that is too broad for your organization, scope the file, mail and calendar ones to a
mail-enabled security group with an [application access policy for Exchange][exchange-policy] and
`Sites.Selected` style scoping — or simply run the delegated mode.

[exchange-policy]: https://learn.microsoft.com/graph/auth-limit-mailbox-access

## 3. Teams application access policy (app-only mode only)

App-only calls to `/users/{id}/onlineMeetings/...` are refused with **403** unless a Teams
application access policy grants your app access to that organizer. This is separate from Graph
permissions and is configured in PowerShell:

```powershell
Install-Module MicrosoftTeams -Scope CurrentUser   # once
Connect-MicrosoftTeams

New-CsApplicationAccessPolicy `
  -Identity Recording-Share-Policy `
  -AppIds "<client-id>" `
  -Description "Allows teams-recording-share to read meetings, recordings and attendance"

# Per organizer (object id or UPN) — repeat for each organizer in ORGANIZER_USER_IDS:
Grant-CsApplicationAccessPolicy `
  -PolicyName Recording-Share-Policy `
  -Identity "<organizer object id>"

# Or tenant-wide, which is what a tenant-wide getAllRecordings subscription needs:
Grant-CsApplicationAccessPolicy -PolicyName Recording-Share-Policy -Global
```

Useful checks:

```powershell
Get-CsApplicationAccessPolicy -Identity Recording-Share-Policy
Get-CsUserPolicyAssignment -Identity "<organizer upn>" -PolicyType ApplicationAccessPolicy
```

**Propagation takes up to 30 minutes.** Until then you will keep seeing 403s on `onlineMeetings`
even though everything is configured correctly. Wait before debugging further.

Delegated mode does not need this policy at all.

## 4. SharePoint / OneDrive external sharing

Granting a guest access to a recording is an ordinary OneDrive sharing operation, so the tenant's
sharing policy applies. Check, in the SharePoint admin center → **Policies** → **Sharing**:

- **External sharing for OneDrive** must be at least **New and existing guests** for guests to be
  invited at all. "Existing guests" only works for people already in the directory. "Only people in
  your organization" blocks every external grant this project makes.
- **"Anyone" links** must be enabled if you intend to use `EXTERNAL_FALLBACK=anonymous-link`. If they
  are off, keep the default `users-link`.
- **Default link type** does not matter — the project always asks for an explicit scope.
- **Link expiration** and **permission** defaults may override `FALLBACK_LINK_EXPIRY_HOURS`; a
  tenant-enforced maximum expiry wins.
- If the organizers are covered by a **sensitivity label** or a DLP policy that blocks external
  sharing, the invite fails per recipient and lands in `errors` in the result.

Guest accounts also have to be allowed at the directory level: Entra → **External Identities** →
**External collaboration settings** → guest invite restrictions.

## 5. Configure the project

```bash
cp .env.example .env
```

Fill in at least `TENANT_ID`, `CLIENT_ID`, `AUTH_MODE`, and — for the service — `CLIENT_SECRET`,
`PUBLIC_BASE_URL` and a random `CLIENT_STATE`:

```bash
node -e 'console.log(require("node:crypto").randomBytes(24).toString("base64url"))'
```

Then sign in (delegated mode only) and check the setup:

```bash
npm install
npm run login                 # opens a browser; sign in as an organizer
npx tsx src/cli.ts whoami
npx tsx src/cli.ts recordings # should list your recent recording files
```

### State directory

`STATE_DIR` (default `./data`) holds:

| File | Contents |
| --- | --- |
| `state.json` | Subscription ids and expiry, processed recording ids, recent errors. |
| `msal-cache.json` | Delegated token cache (mode 0600) unless `TOKEN_CACHE_PATH` points elsewhere. |

Both contain access to your tenant: keep the directory out of source control, and give it a
persistent volume in production. If `state.json` is lost, the service can re-share recordings it has
already processed (harmless but noisy) and will create fresh subscriptions.

## 6. Expose the webhook (service only)

Graph delivers notifications to `PUBLIC_BASE_URL/webhook/notifications` and lifecycle events to
`PUBLIC_BASE_URL/webhook/lifecycle`. The URL must be **public HTTPS with a valid certificate** —
localhost, self-signed certificates and IP allowlists that exclude Microsoft will not work.

When a subscription is created, Graph immediately calls the URL with `?validationToken=...` and
expects the token echoed back as `text/plain` within 10 seconds. If that fails, subscription
creation fails.

Development:

```bash
ngrok http 3978
# PUBLIC_BASE_URL=https://<something>.ngrok-free.app
```

The ngrok URL changes every restart on the free tier; update `PUBLIC_BASE_URL` and re-run
`npm run subscribe -- create` when it does.

Production: any host that gives you HTTPS and a stable hostname. Azure Container Apps (external
ingress on `PORT`) or Azure App Service (Linux, Node 22, `npm run build` then `node dist/service/server.js`)
keep the traffic inside the same tenant. Point `/healthz` at your platform's health probe.

## 7. Subscriptions

```bash
npm run subscribe -- create   # create or refresh the subscriptions for this configuration
npm run subscribe -- renew    # extend the expiry of the existing ones
npm run subscribe -- list     # what Graph currently has
npm run subscribe -- delete   # remove them
```

Facts worth knowing:

- Resource: `communications/onlineMeetings/getAllRecordings` when `ORGANIZER_USER_IDS` is empty,
  otherwise `users/{id}/onlineMeetings/getAllRecordings` per organizer. `changeType: created`.
- **Maximum lifetime is 4320 minutes (3 days).** Subscriptions must be renewed by `PATCH`ing a new
  `expirationDateTime` before they expire; a lapsed subscription cannot be revived, only recreated.
  The running service renews every 6 hours, so it tolerates a few hours of downtime but not three
  days.
- A `lifecycleNotificationUrl` is **required** for expirations over an hour. The service handles
  `reauthorizationRequired` by renewing and `subscriptionRemoved` by recreating.
- Every notification carries `clientState`; the service drops anything whose value does not match
  `CLIENT_STATE` exactly and never acts on it.
- Notifications are answered `202` immediately and processed on a queue; Graph retries only briefly,
  so slow processing must never block the response.

## Troubleshooting

| Symptom | Cause and fix |
| --- | --- |
| `403` on `/users/{id}/onlineMeetings/...` in app-only mode | Missing Teams application access policy for that organizer, or it has not propagated yet. Grant it (step 3) and wait up to 30 minutes. Confirm with `Get-CsUserPolicyAssignment`. |
| `403` immediately after granting admin consent | Consent not actually granted (check the green ticks on **API permissions**), or a token cached before the grant. Delete `msal-cache.json` and sign in again, or wait for the app token to expire. |
| `accountVerificationRequired` on invite | The recipient is a brand-new external guest and app-only auth cannot invite them. Run the share in delegated mode as the organizer, or let the fallback handle it (`EXTERNAL_FALLBACK=users-link`). Check the result's `fallbackLink`. |
| Invite fails for every external address | Tenant external sharing is off or set to "existing guests only" (step 4), or guest invitations are blocked in External Identities. |
| "recording file not found" / `driveItem` missing | The file had not appeared yet (raise `DRIVE_LOOKUP_TIMEOUT_MINUTES`), the folder is not called `Recordings` (set `RECORDINGS_FOLDER`), the subject match fell outside the time tolerance (pass `--drive-item <id>` explicitly), or it was a **channel meeting** — those recordings live in the channel's SharePoint library, which this project does not handle. |
| Attendees are missing from the result | Anonymous and dial-in joiners have no email in the attendance record and are skipped, or the meeting has no attendance report at all. Check `npx tsx src/cli.ts attendees --meeting <id>`. |
| No notifications ever arrive | `clientState` in `.env` differs from the one the subscription was created with (recreate it), the subscription expired (`npm run subscribe -- list`), or `PUBLIC_BASE_URL` is unreachable from the internet. Test with `curl -X POST "$PUBLIC_BASE_URL/webhook/notifications?validationToken=ping"` — it must return `ping`. |
| Subscription creation fails with "endpoint not valid" | The validation handshake did not return the token in plain text within 10s: service not running, wrong path, HTTP instead of HTTPS, or a proxy rewriting the response. |
| `Missing required environment variable ...` | `.env` not loaded — run from the project directory, or set the variables in the MCP client's `env` block. |
| `AUTH_MODE=app requires CLIENT_SECRET` | Set `CLIENT_SECRET`, or switch to `AUTH_MODE=delegated`. |
| `AADSTS50011` redirect URI mismatch at sign-in | Add `http://localhost` under **Mobile and desktop applications** on the app registration (not Web, not SPA). |
| `AADSTS7000218` / `AADSTS70002` client secret expected at sign-in | The redirect URI was registered as a **Web** platform; move it to **Mobile and desktop applications**. |
| Browser sign-in on a headless machine | Run `npm run login` on a machine with a browser and copy the `msal-cache.json` file over, or use `AUTH_MODE=app`. |
| Login keeps being asked for | The token cache is not persisting: check `TOKEN_CACHE_PATH`/`STATE_DIR` is writable, and that `offline_access` is in the delegated scopes. |
