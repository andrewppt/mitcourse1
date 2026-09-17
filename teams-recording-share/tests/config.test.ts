import test from "node:test";
import assert from "node:assert/strict";
import path from "node:path";

import { loadConfig } from "../src/config.js";

const MINIMAL = {
  TENANT_ID: "11111111-2222-3333-4444-555555555555",
  CLIENT_ID: "66666666-7777-8888-9999-000000000000",
} as NodeJS.ProcessEnv;

test("loadConfig applies documented defaults in delegated mode", () => {
  const cfg = loadConfig({ ...MINIMAL });
  assert.equal(cfg.tenantId, MINIMAL.TENANT_ID);
  assert.equal(cfg.clientId, MINIMAL.CLIENT_ID);
  assert.equal(cfg.authMode, "delegated");
  assert.equal(cfg.clientSecret, undefined);
  assert.equal(cfg.port, 3978);
  assert.equal(cfg.clientState, "");
  assert.equal(cfg.shareRole, "read");
  assert.equal(cfg.sendInvitation, true);
  assert.equal(cfg.externalFallback, "users-link");
  assert.equal(cfg.fallbackLinkExpiryHours, 0);
  assert.equal(cfg.includeCalendarInvitees, false);
  assert.equal(cfg.recordingsFolder, "Recordings");
  assert.equal(cfg.driveLookupTimeoutMinutes, 20);
  assert.equal(cfg.stateDir, "./data");
  assert.equal(cfg.graphBaseUrl, "https://graph.microsoft.com");
  assert.equal(cfg.publicBaseUrl, undefined);
  assert.deepEqual(cfg.organizerUserIds, []);
  assert.deepEqual(cfg.alwaysInclude, []);
  assert.deepEqual(cfg.internalDomains, []);
  assert.equal(cfg.tokenCachePath, path.resolve("./data", "msal-cache.json"));
  assert.ok(cfg.delegatedScopes.includes("offline_access"));
  assert.ok(cfg.delegatedScopes.includes("OnlineMeetingRecording.Read.All"));
});

test("loadConfig requires TENANT_ID and CLIENT_ID", () => {
  assert.throws(() => loadConfig({ CLIENT_ID: "c" } as NodeJS.ProcessEnv), /TENANT_ID/);
  assert.throws(() => loadConfig({ TENANT_ID: "t" } as NodeJS.ProcessEnv), /CLIENT_ID/);
});

test("app mode requires a client secret", () => {
  assert.throws(() => loadConfig({ ...MINIMAL, AUTH_MODE: "app" }), /CLIENT_SECRET/);
  const cfg = loadConfig({ ...MINIMAL, AUTH_MODE: "app", CLIENT_SECRET: "s3cr3t" });
  assert.equal(cfg.authMode, "app");
  assert.equal(cfg.clientSecret, "s3cr3t");
});

test("AUTH_MODE is validated and case-insensitive", () => {
  assert.equal(loadConfig({ ...MINIMAL, AUTH_MODE: "Delegated" }).authMode, "delegated");
  assert.throws(() => loadConfig({ ...MINIMAL, AUTH_MODE: "certificate" }), /AUTH_MODE/);
});

test("ORGANIZER_USER_IDS accepts comma, semicolon and whitespace separated ids", () => {
  const cfg = loadConfig({ ...MINIMAL, ORGANIZER_USER_IDS: " a-1, b-2 ;c-3\nd-4  " });
  assert.deepEqual(cfg.organizerUserIds, ["a-1", "b-2", "c-3", "d-4"]);
  assert.deepEqual(loadConfig({ ...MINIMAL, ORGANIZER_USER_IDS: "" }).organizerUserIds, []);
});

test("list values that are addresses or domains are lowercased", () => {
  const cfg = loadConfig({
    ...MINIMAL,
    ALWAYS_INCLUDE: "Archive@Contoso.com, Legal@Contoso.com",
    INTERNAL_DOMAINS: "Contoso.COM fabrikam.com",
  });
  assert.deepEqual(cfg.alwaysInclude, ["archive@contoso.com", "legal@contoso.com"]);
  assert.deepEqual(cfg.internalDomains, ["contoso.com", "fabrikam.com"]);
});

test("booleans, numbers and URLs are parsed", () => {
  const cfg = loadConfig({
    ...MINIMAL,
    PORT: "8080",
    SEND_INVITATION: "false",
    INCLUDE_CALENDAR_INVITEES: "yes",
    FALLBACK_LINK_EXPIRY_HOURS: "72",
    DRIVE_LOOKUP_TIMEOUT_MINUTES: "5",
    PUBLIC_BASE_URL: "https://abc.ngrok.app///",
    RECORDINGS_FOLDER: "Aufzeichnungen",
    STATE_DIR: "/var/lib/trs",
    CLIENT_STATE: "shared-secret",
    GRAPH_BASE_URL: "https://graph.test",
  });
  assert.equal(cfg.port, 8080);
  assert.equal(cfg.sendInvitation, false);
  assert.equal(cfg.includeCalendarInvitees, true);
  assert.equal(cfg.fallbackLinkExpiryHours, 72);
  assert.equal(cfg.driveLookupTimeoutMinutes, 5);
  assert.equal(cfg.publicBaseUrl, "https://abc.ngrok.app");
  assert.equal(cfg.recordingsFolder, "Aufzeichnungen");
  assert.equal(cfg.stateDir, "/var/lib/trs");
  assert.equal(cfg.clientState, "shared-secret");
  assert.equal(cfg.graphBaseUrl, "https://graph.test");
  assert.equal(cfg.tokenCachePath, path.resolve("/var/lib/trs", "msal-cache.json"));
});

test("TOKEN_CACHE_PATH overrides the STATE_DIR default", () => {
  const cfg = loadConfig({ ...MINIMAL, STATE_DIR: "/var/lib/trs", TOKEN_CACHE_PATH: "/tmp/cache.json" });
  assert.equal(cfg.tokenCachePath, "/tmp/cache.json");
});

test("DELEGATED_SCOPES replaces the default scope set", () => {
  const cfg = loadConfig({ ...MINIMAL, DELEGATED_SCOPES: "User.Read offline_access" });
  assert.deepEqual(cfg.delegatedScopes, ["User.Read", "offline_access"]);
});

test("EXTERNAL_FALLBACK is validated", () => {
  assert.equal(loadConfig({ ...MINIMAL, EXTERNAL_FALLBACK: "skip" }).externalFallback, "skip");
  assert.equal(loadConfig({ ...MINIMAL, EXTERNAL_FALLBACK: "anonymous-link" }).externalFallback, "anonymous-link");
  assert.throws(() => loadConfig({ ...MINIMAL, EXTERNAL_FALLBACK: "carrier-pigeon" }), /EXTERNAL_FALLBACK/);
});
