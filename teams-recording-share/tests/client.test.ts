import test from "node:test";
import assert from "node:assert/strict";

import { GraphClient, GraphError } from "../src/graph/client.js";
import { silentLogger } from "../src/log.js";

const BASE = "https://graph.test";
const tokens = { getToken: async () => "t" };

type Call = { url: string; method: string; headers: Headers; body?: string };

/** Builds a fetch stand-in that replays the given responses in order and records every call. */
function fakeFetch(responses: Response[]) {
  const calls: Call[] = [];
  const impl = (async (input: any, init: any = {}) => {
    const url = typeof input === "string" ? input : (input?.url ?? String(input));
    calls.push({
      url,
      method: (init.method ?? "GET").toUpperCase(),
      headers: new Headers(init.headers ?? {}),
      body: typeof init.body === "string" ? init.body : undefined,
    });
    const next = responses.shift();
    if (!next) throw new Error(`unexpected extra request: ${url}`);
    return next;
  }) as unknown as typeof fetch;
  return { impl, calls };
}

const json = (body: unknown, status = 200, headers: Record<string, string> = {}) =>
  new Response(JSON.stringify(body), { status, headers: { "content-type": "application/json", ...headers } });

const client = (impl: typeof fetch) => new GraphClient(tokens, { baseUrl: BASE, fetchImpl: impl, logger: silentLogger });

test("get returns parsed JSON and sends a bearer token", async () => {
  const { impl, calls } = fakeFetch([json({ id: "u1", displayName: "Alice" })]);
  const got = await client(impl).get<{ id: string }>("/users/u1");
  assert.deepEqual(got, { id: "u1", displayName: "Alice" });
  assert.equal(calls.length, 1);
  assert.equal(calls[0].method, "GET");
  assert.equal(calls[0].url, `${BASE}/v1.0/users/u1`);
  assert.equal(calls[0].headers.get("authorization"), "Bearer t");
});

test("query options are appended and undefined values dropped", async () => {
  const { impl, calls } = fakeFetch([json({ value: [] })]);
  await client(impl).get("/users/u1/drive/root:/Recordings:/children", {
    query: { $top: 50, $orderby: "createdDateTime desc", nothing: undefined },
  });
  const url = new URL(calls[0].url);
  assert.equal(url.searchParams.get("$top"), "50");
  assert.equal(url.searchParams.get("$orderby"), "createdDateTime desc");
  assert.ok(!url.searchParams.has("nothing"));
});

test("absolute paths are used verbatim (nextLink) and getAll follows them", async () => {
  const nextLink = `${BASE}/v1.0/users/u1/onlineMeetings/m1/attendanceReports?$skiptoken=abc`;
  const { impl, calls } = fakeFetch([
    json({ value: [{ id: "a" }, { id: "b" }], "@odata.nextLink": nextLink }),
    json({ value: [{ id: "c" }] }),
  ]);
  const all = await client(impl).getAll<{ id: string }>("/users/u1/onlineMeetings/m1/attendanceReports");
  assert.deepEqual(all.map((x) => x.id), ["a", "b", "c"]);
  assert.equal(calls.length, 2);
  assert.equal(calls[1].url, nextLink);
});

test("429 is retried after honouring Retry-After", async () => {
  const { impl, calls } = fakeFetch([
    json({ error: { code: "activityLimitReached", message: "throttled" } }, 429, { "retry-after": "0" }),
    json({ id: "u1" }),
  ]);
  const got = await client(impl).get<{ id: string }>("/users/u1");
  assert.deepEqual(got, { id: "u1" });
  assert.equal(calls.length, 2, "the request should have been retried exactly once");
});

test("a non-retryable error becomes a GraphError carrying status and code", async () => {
  const { impl, calls } = fakeFetch([
    json({ error: { code: "itemNotFound", message: "The resource could not be found." } }, 404),
  ]);
  await assert.rejects(
    () => client(impl).get("/users/u1/drive/items/missing"),
    (err: unknown) => {
      assert.ok(err instanceof GraphError, `expected a GraphError, got ${String(err)}`);
      assert.equal(err.status, 404);
      assert.equal(err.code, "itemNotFound");
      assert.match(err.message, /could not be found|itemNotFound/i);
      return true;
    },
  );
  assert.equal(calls.length, 1, "4xx other than 429 must not be retried");
});

test("207 Multi-Status bodies are returned to the caller instead of throwing", async () => {
  const body = {
    value: [
      { id: "p1", roles: ["read"], invitation: { email: "bob@contoso.com" } },
      { error: { code: "accountVerificationRequired", message: "The account needs to be verified." } },
    ],
  };
  const { impl } = fakeFetch([json(body, 207)]);
  const got = await client(impl).post<typeof body>("/users/u1/drive/items/i1/invite", {
    recipients: [{ email: "bob@contoso.com" }, { email: "dana@partner.example" }],
    roles: ["read"],
  });
  assert.deepEqual(got, body);
});

test("post sends a JSON body", async () => {
  const { impl, calls } = fakeFetch([json({ id: "p1" })]);
  await client(impl).post("/users/u1/drive/items/i1/createLink", { type: "view", scope: "users" });
  assert.equal(calls[0].method, "POST");
  assert.match(calls[0].headers.get("content-type") ?? "", /application\/json/);
  assert.deepEqual(JSON.parse(calls[0].body ?? "{}"), { type: "view", scope: "users" });
});

test("204 responses resolve without a body", async () => {
  const { impl, calls } = fakeFetch([new Response(null, { status: 204 })]);
  await client(impl).delete("/subscriptions/s1");
  assert.equal(calls[0].method, "DELETE");
});

test("the beta endpoint can be selected per request", async () => {
  const { impl, calls } = fakeFetch([json({ value: [] })]);
  await client(impl).get("/users/u1/onlineMeetings/m1/recordings", { version: "beta" });
  assert.equal(calls[0].url, `${BASE}/beta/users/u1/onlineMeetings/m1/recordings`);
});
