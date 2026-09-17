# teams-recording-share

When a Teams meeting recording finishes processing into the organizer's OneDrive, this project
grants **read access to everyone who attended the meeting** — internal people and external guests
alike — instead of leaving the file visible only to the organizer.

What it does, in order:

1. Resolves the meeting (by online meeting id or join URL) and its recordings.
2. Reads the meeting's **attendance reports** to get the list of people who actually joined.
   Optionally adds the **calendar invitees** who were invited but did not join
   (`INCLUDE_CALENDAR_INVITEES=true`), plus the organizer and anything in `ALWAYS_INCLUDE`.
3. Finds the recording **file** in the organizer's OneDrive `Recordings` folder, matching on
   subject and creation time (retrying for a while, because the file appears a few minutes after
   the notification).
4. Reads the file's existing permissions and **invites only the people who do not already have
   access** — so re-running is safe and quiet.
5. For recipients OneDrive refuses to invite directly (typically brand-new external guests under
   app-only auth), falls back to a **sharing link**: a `users`-scoped link granted to exactly those
   addresses, or an anonymous link emailed to them, or nothing — your choice
   (`EXTERNAL_FALLBACK`).

Two ways to run it:

- **Automatic**: a small webhook service subscribed to Graph's "recording available" notifications.
  Every recording gets shared minutes after the meeting ends, with no one in the loop.
- **Manual / ad hoc**: an MCP server (use it from Claude Desktop or Claude Code) and a CLI, both
  running as you via device-code sign-in. Useful for one-off meetings, for backfilling, and when
  app-only auth cannot invite a guest.

Setup — app registration, permissions, the Teams application access policy, external sharing
settings — is in **[docs/SETUP.md](docs/SETUP.md)**.

## Architecture

```
  Microsoft Graph                       this project                    Microsoft 365
  ---------------                       ------------                    -------------

  subscription on                  +--------------------------+
  .../getAllRecordings             |  src/service/subscribe.ts|  POST /subscriptions
  (changeType: created) <----------|  create / renew / delete |------------> Graph
        |                          +--------------------------+
        |  "recording is ready"                 ^ every 6h (max lifetime 3 days)
        v                                       |
  POST <PUBLIC_BASE_URL>          +--------------------------+
    /webhook/notifications  ----> |  src/service/server.ts   |
    /webhook/lifecycle            |  validationToken echo    |
                                  |  clientState check       |
                                  |  202 + async queue       |
                                  |  dedupe via StateStore   |
                                  +------------+-------------+
                                               |
                                               v
                                  +--------------------------+
   MCP client (Claude)            |  src/core/share.ts       |
     src/mcp/server.ts  --------> |  the share pipeline      |
   terminal                       +------------+-------------+
     src/cli.ts        ---------> |            |
                                  |            | 1. meeting + recordings
                                  |            | 2. attendance records
                                  |            |    (+ calendar invitees)
                                  |            | 3. find the .mp4 in OneDrive
                                  |            | 4. read current permissions
                                  |            | 5. invite the missing people
                                  |            | 6. link fallback for guests
                                  +------------+-------------+
                                               |
                                               v
                                   organizer's OneDrive: /Recordings/<Subject>-...mp4
                                   permissions: read for every attendee
```

Auth is one of two modes (`AUTH_MODE`): `app` (client credentials, for the unattended service) or
`delegated` (device code as a signed-in person, for the CLI and MCP server). The pipeline itself is
identical in both.

## Quick start

Requires Node 22 (Node 20 works too) and an Entra app registration — see
[docs/SETUP.md](docs/SETUP.md) first, it takes about ten minutes.

```bash
npm install
cp .env.example .env      # fill in TENANT_ID and CLIENT_ID (AUTH_MODE=delegated to start)

npm run login             # device-code sign-in; prints a code to enter at microsoft.com/devicelogin
npx tsx src/cli.ts whoami # confirms who you are signed in as

# share the recording of one meeting, dry run first
npm run share -- --join-url "https://teams.microsoft.com/l/meetup-join/..." --dry-run
npm run share -- --join-url "https://teams.microsoft.com/l/meetup-join/..."

npm run mcp               # run the MCP server on stdio (usually started by the MCP client)
npm test                  # node --test
```

Other CLI commands (all accept `--json`):

| Command | What it does |
| --- | --- |
| `npx tsx src/cli.ts login` | Device-code sign-in; writes the token cache. |
| `npx tsx src/cli.ts whoami` | Auth mode and signed-in account. |
| `npx tsx src/cli.ts share --meeting <id>\|--join-url <url> [--organizer <id>] [--recording <id>] [--drive-item <id>] [--extra a@b.com,c@d.com] [--dry-run]` | Run the share pipeline. |
| `npx tsx src/cli.ts attendees --meeting <id>\|--join-url <url> [--calendar]` | Who attended, marked internal/external. |
| `npx tsx src/cli.ts recordings [--organizer <id>] [--top 25]` | Recent files in the `Recordings` folder. |

`npm run share -- ...` is shorthand for the `share` command.

## Use it from an MCP client

The server speaks MCP over stdio and writes nothing but protocol to stdout (logs go to stderr).
Sign in once with `npm run login` before wiring it up — the MCP server uses the same token cache.

**Claude Desktop** (`claude_desktop_config.json`: macOS
`~/Library/Application Support/Claude/claude_desktop_config.json`, Windows
`%APPDATA%\Claude\claude_desktop_config.json`):

```json
{
  "mcpServers": {
    "teams-recording-share": {
      "command": "npx",
      "args": ["tsx", "/absolute/path/to/teams-recording-share/src/mcp/server.ts"],
      "env": {
        "TENANT_ID": "00000000-0000-0000-0000-000000000000",
        "CLIENT_ID": "00000000-0000-0000-0000-000000000000",
        "AUTH_MODE": "delegated",
        "STATE_DIR": "/absolute/path/to/teams-recording-share/data"
      }
    }
  }
}
```

After `npm run build` you can use the compiled entry point instead, which starts faster:

```json
{
  "mcpServers": {
    "teams-recording-share": {
      "command": "node",
      "args": ["/absolute/path/to/teams-recording-share/dist/mcp/server.js"],
      "env": { "TENANT_ID": "...", "CLIENT_ID": "...", "AUTH_MODE": "delegated" }
    }
  }
}
```

**Claude Code**:

```bash
# compiled
claude mcp add teams-recording-share -- node /absolute/path/to/teams-recording-share/dist/mcp/server.js

# or straight from TypeScript
claude mcp add teams-recording-share -- npx tsx /absolute/path/to/teams-recording-share/src/mcp/server.ts
```

The server loads `.env` from its working directory, so either pass the variables in the client's
`env` block or start it from the project directory.

### Tools

| Tool | Arguments | Result |
| --- | --- | --- |
| `share_recording` | `meetingId` or `joinWebUrl`, `organizerUserId?`, `recordingId?`, `driveItemId?`, `extraEmails?`, `dryRun?` | Grants access to every attendee; returns granted / already had access / fallback link / skipped / errors. |
| `list_meeting_attendees` | `meetingId` or `joinWebUrl`, `organizerUserId?`, `includeCalendarInvitees?` | Attendance records (plus calendar invitees), each marked internal or external. |
| `find_recording` | `meetingId` or `joinWebUrl`, `organizerUserId?` | The meeting's recordings, each matched to its OneDrive file. |
| `list_recent_recordings` | `organizerUserId?`, `top?` | Newest files in the organizer's `Recordings` folder. |
| `get_recording_access` | `driveItemId`, `organizerUserId?` | Current permissions and the addresses that have access. |
| `whoami` | — | Auth mode and signed-in account. |

There is also a read-only resource, `teams-recording-share://config`, with a non-secret view of the
configuration.

In delegated mode `organizerUserId` defaults to the signed-in user. In app-only mode it is required.
Start with `dryRun: true` on `share_recording`: it performs no writes and shows exactly who would be
invited.

## Running the webhook service

The service subscribes to Graph, receives the "recording available" notification, and runs the same
pipeline automatically.

```bash
# 1. expose the port over HTTPS (development)
ngrok http 3978
#    put the https URL in .env as PUBLIC_BASE_URL, and set a random CLIENT_STATE

# 2. start the service (it also renews subscriptions every 6h)
npm run service

# 3. create the subscriptions (in another terminal)
npm run subscribe -- create     # also: renew | list | delete
```

Graph must be able to reach `PUBLIC_BASE_URL` over **public HTTPS** with a valid certificate before
the subscription can be created: it calls the endpoint with a `validationToken` and expects the token
echoed back within 10 seconds. The service answers on:

| Route | Purpose |
| --- | --- |
| `POST /webhook/notifications` | Change notifications (and the validation handshake). |
| `POST /webhook/lifecycle` | Lifecycle events: reauthorization required, subscription removed. |
| `GET /healthz` | Liveness probe. |
| `GET /status` | Subscriptions, recently processed recordings, recent errors. |

For production, `AUTH_MODE=app` with a client secret, and any container host will do — Azure
Container Apps (ingress on `PORT`, external, HTTPS terminated for you) or Azure App Service
(Linux, Node 22) are the least-friction options in the same tenant. Persist `STATE_DIR` on a volume
so the processed-recording list and subscription ids survive restarts; without it a restart can
re-share recordings it has already handled. Subscription and lifecycle details are in
[docs/SETUP.md](docs/SETUP.md#subscriptions).

## Limitations

- **Timing.** The notification fires when the recording is *ready*, which is typically a few minutes
  after the call ends — longer for long meetings. The file then still has to appear in OneDrive, so
  the pipeline polls for it for up to `DRIVE_LOOKUP_TIMEOUT_MINUTES` (default 20) before giving up.
- **Anonymous joiners cannot be granted access.** Attendance records for people who joined
  anonymously (or by phone) carry no email address. They are logged and skipped — there is nothing to
  invite.
- **Brand-new external guests cannot be invited by app-only auth.** Graph returns
  `accountVerificationRequired` for a guest who is not yet in the directory. Existing guests are
  fine. Either run the share in delegated mode as the organizer, or rely on the link fallback
  (`EXTERNAL_FALLBACK=users-link`, the default).
- **Your tenant's external sharing policy wins.** If SharePoint/OneDrive external sharing is set to
  "Only people in your organization", no guest can be granted access and no link will work,
  whatever this project does. Anonymous ("Anyone") links additionally have to be enabled for
  `EXTERNAL_FALLBACK=anonymous-link`.
- **Channel meetings are not covered.** Recordings of meetings held in a Teams channel land in that
  channel's SharePoint document library, not in the organizer's OneDrive, so the file lookup will not
  find them.
- **Attendance reports must exist.** They are produced for meetings that have them enabled; without
  an attendance report there is no attendee list, and only the organizer, `ALWAYS_INCLUDE` and any
  calendar invitees remain.
- **Recording folder name.** The lookup assumes the folder is called `Recordings`. Localised OneDrive
  tenants may name it differently — set `RECORDINGS_FOLDER`.
- One meeting's recordings only; there is no bulk backfill over a date range.
