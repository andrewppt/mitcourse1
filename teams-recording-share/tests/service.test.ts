import test, { after, before } from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import type { AddressInfo } from "node:net";

import { loadConfig, type AppConfig } from "../src/config.js";
import { silentLogger } from "../src/log.js";
import { createApp } from "../src/service/server.js";
import { StateStore } from "../src/service/store.js";

const ORGANIZER_OID = "880e1e61-af63-43c9-a48e-d6b63684c21c";
const THREAD_ID = "19:meeting_ZmY4YTJhNmUtY2Y3OS00ZGZlLWI1ODEtYjA2ZTZmZjA0YzQ0@thread.v2";
const MEETING_ID = Buffer.from(`1*${ORGANIZER_OID}*0**${THREAD_ID}`, "utf8").toString("base64");
const CLIENT_STATE = "test-client-state";

const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));

/** Poll `fn` until it returns a truthy value or the deadline passes. */
async function waitFor<T>(fn: () => T | undefined, timeoutMs = 2000, stepMs = 25): Promise<T | undefined> {
  const deadline = Date.now() + timeoutMs;
  for (;;) {
    const v = fn();
    if (v) return v;
    if (Date.now() > deadline) return undefined;
    await sleep(stepMs);
  }
}

/** The Graph client is never expected to be reached: the share step is faked. */
const fakeGraph = {
  get: async () => ({}),
  getAll: async () => [],
  post: async () => ({}),
  patch: async () => ({}),
  delete: async () => undefined,
};

let stateDir: string;
let cfg: AppConfig;
let created: any;
let server: import("node:http").Server;
let baseUrl: string;
const shareCalls: any[] = [];

const shareFake = async (...args: any[]) => {
  const req = args.find((a) => a && typeof a === "object" && ("organizerUserId" in a || "meetingId" in a || "recordingId" in a));
  shareCalls.push(req ?? args[0]);
  return {
    dryRun: false,
    meeting: { id: MEETING_ID, organizerUserId: req?.organizerUserId ?? ORGANIZER_OID },
    recordings: [],
    attendees: [],
    warnings: [],
  };
};

const notification = (overrides: Record<string, unknown> = {}) => ({
  value: [
    {
      subscriptionId: "sub-1",
      subscriptionExpirationDateTime: new Date(Date.now() + 3600_000).toISOString(),
      changeType: "created",
      clientState: CLIENT_STATE,
      resource: `communications/onlineMeetings('${MEETING_ID}')/recordings('rec1')`,
      resourceData: { "@odata.type": "#Microsoft.Graph.callRecording", id: "rec1" },
      tenantId: "11111111-2222-3333-4444-555555555555",
      ...overrides,
    },
  ],
});

const postJson = (pathname: string, body: unknown) =>
  fetch(`${baseUrl}${pathname}`, {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify(body),
  });

before(async () => {
  stateDir = fs.mkdtempSync(path.join(os.tmpdir(), "trs-service-"));
  cfg = loadConfig({
    TENANT_ID: "11111111-2222-3333-4444-555555555555",
    CLIENT_ID: "66666666-7777-8888-9999-000000000000",
    AUTH_MODE: "delegated",
    CLIENT_STATE,
    STATE_DIR: stateDir,
    PUBLIC_BASE_URL: "https://example.test",
  });
  const store = StateStore.fromConfig(cfg);
  created = createApp({ cfg, graph: fakeGraph as any, store, log: silentLogger, share: shareFake });
  const app = typeof created === "function" ? created : (created.app ?? created);
  server = await new Promise((resolve) => {
    const s = app.listen(0, () => resolve(s));
  });
  baseUrl = `http://127.0.0.1:${(server.address() as AddressInfo).port}`;
});

after(async () => {
  if (created && typeof created.close === "function") await created.close();
  else if (created && typeof created.stop === "function") await created.stop();
  if (server) await new Promise<void>((resolve) => server.close(() => resolve()));
  if (stateDir) fs.rmSync(stateDir, { recursive: true, force: true });
});

test("GET /healthz reports ok", async () => {
  const res = await fetch(`${baseUrl}/healthz`);
  assert.equal(res.status, 200);
  assert.match(await res.text(), /ok/i);
});

test("the validation handshake echoes the token as text/plain", async () => {
  const token = "validation-token-123";
  const res = await fetch(`${baseUrl}/webhook/notifications?validationToken=${token}`, { method: "POST" });
  assert.equal(res.status, 200);
  assert.match(res.headers.get("content-type") ?? "", /text\/plain/);
  assert.equal((await res.text()).trim(), token);
});

test("the lifecycle endpoint answers the same handshake", async () => {
  const token = "lifecycle-token-456";
  const res = await fetch(`${baseUrl}/webhook/lifecycle?validationToken=${token}`, { method: "POST" });
  assert.equal(res.status, 200);
  assert.equal((await res.text()).trim(), token);
});

test("a notification with the wrong clientState is ignored", async () => {
  const before = shareCalls.length;
  const res = await postJson("/webhook/notifications", notification({ clientState: "not-the-secret" }));
  assert.ok(res.status < 500, `unexpected status ${res.status}`);
  await res.arrayBuffer();
  await sleep(300);
  assert.equal(shareCalls.length, before, "the share pipeline must not run for an unverified notification");
});

test("a callRecording notification is accepted and shared for the decoded organizer", async () => {
  const res = await postJson("/webhook/notifications", notification());
  assert.equal(res.status, 202, "Graph expects an immediate 202");
  await res.arrayBuffer();

  const call = await waitFor(() => shareCalls[0], 2000);
  assert.ok(call, "the share pipeline was not invoked within 2s");
  assert.equal(call.organizerUserId, ORGANIZER_OID, "organizer must come from the decoded meeting id");
  assert.equal(call.meetingId, MEETING_ID);
  assert.equal(call.recordingId, "rec1");
});

test("a duplicate notification for the same recording is not reprocessed", async () => {
  const seen = shareCalls.length;
  const res = await postJson("/webhook/notifications", notification());
  assert.equal(res.status, 202);
  await res.arrayBuffer();
  await sleep(600);
  assert.equal(shareCalls.length, seen, "the same recording id must only be processed once");
});

test("notifications for other resource types are ignored", async () => {
  const seen = shareCalls.length;
  const res = await postJson(
    "/webhook/notifications",
    notification({
      resource: `communications/onlineMeetings('${MEETING_ID}')/transcripts('tr1')`,
      resourceData: { "@odata.type": "#Microsoft.Graph.callTranscript", id: "tr1" },
    }),
  );
  assert.ok(res.status < 500);
  await res.arrayBuffer();
  await sleep(300);
  assert.equal(shareCalls.length, seen, "only callRecording notifications trigger sharing");
});
