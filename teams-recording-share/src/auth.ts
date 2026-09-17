/**
 * Token acquisition for Microsoft Graph, backed by msal-node.
 *
 *  - "app" mode:       ConfidentialClientApplication + client credentials (`<graphBaseUrl>/.default`).
 *  - "delegated" mode: PublicClientApplication + authorization code with PKCE on a localhost redirect
 *                      (a browser window; LOGIN_FLOW=browser, the default) or the device-code flow
 *                      (LOGIN_FLOW=device-code, only if your tenant allows it). The MSAL token cache is
 *                      persisted to `cfg.tokenCachePath` (mode 0600) so the CLI/MCP server only signs in once.
 *
 * Nothing in here ever writes to stdout and no token value is ever logged.
 */
import { spawn } from "node:child_process";
import { randomBytes } from "node:crypto";
import fs from "node:fs";
import fsp from "node:fs/promises";
import http from "node:http";
import type { AddressInfo } from "node:net";
import path from "node:path";
import {
  ConfidentialClientApplication,
  CryptoProvider,
  PublicClientApplication,
  type AccountInfo,
  type AuthenticationResult,
  type Configuration,
  type DeviceCodeRequest,
  type ICachePlugin,
  type TokenCacheContext,
} from "@azure/msal-node";
import type { AppConfig } from "./config.js";
import type { Logger } from "./log.js";

export interface TokenProvider {
  getToken(): Promise<string>;
  /** Account username when delegated. */
  account?: () => Promise<string | undefined>;
}

export interface DelegatedTokenOptions {
  /** Called with sign-in instructions that must be shown to the user (the URL to open, or the device code). Defaults to stderr. */
  onMessage?: (message: string) => void;
  /** Allow an interactive sign-in when the silent flow cannot produce a token. */
  allowInteractive?: boolean;
  logger?: Logger;
}

/** Refresh this long before the access token actually expires. */
const EXPIRY_SKEW_MS = 2 * 60 * 1000;

/** OIDC scopes MSAL manages itself; they must not be prefixed with the Graph resource URI. */
const RESERVED_SCOPES = new Set(["openid", "profile", "offline_access", "email"]);

type CachedToken = { token: string; expiresAt: number };

function isFresh(cached: CachedToken | undefined): cached is CachedToken {
  return !!cached && cached.expiresAt - EXPIRY_SKEW_MS > Date.now();
}

function expiresAtOf(result: AuthenticationResult): number {
  // MSAL always returns expiresOn for a successful token response; fall back to a conservative 5 minutes.
  return result.expiresOn ? result.expiresOn.getTime() : Date.now() + 5 * 60 * 1000;
}

function authorityFor(cfg: AppConfig): string {
  return `https://login.microsoftonline.com/${cfg.tenantId}`;
}

/**
 * Turn the configured short scope names into fully qualified Graph scopes.
 * `offline_access` / `openid` / `profile` are dropped: MSAL requests them itself and rejecting them
 * here keeps the resource of the request unambiguous.
 */
export function delegatedScopesFor(cfg: AppConfig): string[] {
  const base = cfg.graphBaseUrl.replace(/\/+$/, "");
  const out: string[] = [];
  for (const raw of cfg.delegatedScopes) {
    const scope = raw.trim();
    if (!scope) continue;
    if (RESERVED_SCOPES.has(scope.toLowerCase())) continue;
    const full = /^[a-z][a-z0-9+.-]*:\/\//i.test(scope) ? scope : `${base}/${scope}`;
    if (!out.includes(full)) out.push(full);
  }
  if (out.length === 0) out.push(`${base}/User.Read`);
  return out;
}

/** ICachePlugin that persists the MSAL cache as a 0600 file, creating parent directories as needed. */
export function createFileCachePlugin(cacheFile: string, log?: Logger): ICachePlugin {
  return {
    async beforeCacheAccess(cacheContext: TokenCacheContext): Promise<void> {
      try {
        if (!fs.existsSync(cacheFile)) return;
        const data = await fsp.readFile(cacheFile, "utf8");
        if (data.trim()) cacheContext.tokenCache.deserialize(data);
      } catch (err) {
        log?.warn("could not read token cache", { file: cacheFile, error: (err as Error).message });
      }
    },
    async afterCacheAccess(cacheContext: TokenCacheContext): Promise<void> {
      if (!cacheContext.hasChanged) return;
      try {
        await fsp.mkdir(path.dirname(path.resolve(cacheFile)), { recursive: true });
        await fsp.writeFile(cacheFile, cacheContext.tokenCache.serialize(), { encoding: "utf8", mode: 0o600 });
        // writeFile only applies `mode` when creating the file; enforce it for pre-existing files too.
        await fsp.chmod(cacheFile, 0o600).catch(() => undefined);
      } catch (err) {
        log?.warn("could not persist token cache", { file: cacheFile, error: (err as Error).message });
      }
    },
  };
}

function publicClientConfig(cfg: AppConfig, log?: Logger): Configuration {
  return {
    auth: { clientId: cfg.clientId, authority: authorityFor(cfg) },
    cache: { cachePlugin: createFileCachePlugin(cfg.tokenCachePath, log) },
  };
}

function createPublicClient(cfg: AppConfig, log?: Logger): PublicClientApplication {
  return new PublicClientApplication(publicClientConfig(cfg, log));
}

async function firstAccount(app: PublicClientApplication): Promise<AccountInfo | undefined> {
  const accounts = await app.getTokenCache().getAllAccounts();
  return accounts.length > 0 ? accounts[0] : undefined;
}

function deviceCodeRequest(scopes: string[], onMessage: (message: string) => void): DeviceCodeRequest {
  return {
    scopes,
    deviceCodeCallback: (response) => onMessage(response.message),
  };
}

const defaultMessagePrinter = (message: string): void => {
  process.stderr.write(message + "\n");
};

/** How long the browser sign-in waits for the redirect before giving up. */
const BROWSER_LOGIN_TIMEOUT_MS = 5 * 60 * 1000;

const LOGIN_DONE_HTML = `<!doctype html><html><head><meta charset="utf-8"><title>Signed in</title></head>
<body style="font-family:system-ui,sans-serif;margin:3rem"><h1>Signed in</h1><p>You can close this window and return to the terminal.</p></body></html>`;

const LOGIN_FAILED_HTML = (reason: string): string =>
  `<!doctype html><html><head><meta charset="utf-8"><title>Sign-in failed</title></head>
<body style="font-family:system-ui,sans-serif;margin:3rem"><h1>Sign-in failed</h1><p>${reason}</p></body></html>`;

function escapeHtml(s: string): string {
  return s.replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c] ?? c);
}

/** Best-effort: open the system browser. Never throws; the URL is always printed as a fallback. */
function tryOpenBrowser(url: string, log?: Logger): void {
  try {
    const [cmd, args] =
      process.platform === "win32"
        ? ["cmd", ["/c", "start", "", url]]
        : process.platform === "darwin"
          ? ["open", [url]]
          : ["xdg-open", [url]];
    const child = spawn(cmd, args, { stdio: "ignore", detached: true });
    child.on("error", (err) => log?.debug("could not open browser", { error: err.message }));
    child.unref();
  } catch (err) {
    log?.debug("could not open browser", { error: (err as Error).message });
  }
}

/**
 * Authorization-code + PKCE sign-in for a public client: start a one-shot HTTP listener on
 * 127.0.0.1, send the user to the Entra authorize endpoint with `http://localhost:{port}` as the
 * redirect, and exchange the returned code. Entra accepts any port for a registered
 * `http://localhost` redirect URI (RFC 8252 loopback rule), so a random free port is fine.
 */
async function acquireTokenByBrowser(
  app: PublicClientApplication,
  scopes: string[],
  opts: { onMessage: (message: string) => void; port?: number; log?: Logger },
): Promise<AuthenticationResult> {
  const { verifier, challenge } = await new CryptoProvider().generatePkceCodes();
  const state = randomBytes(16).toString("hex");

  const server = http.createServer();
  await new Promise<void>((resolve, reject) => {
    server.once("error", reject);
    server.listen(opts.port ?? 0, "127.0.0.1", () => resolve());
  });
  const port = (server.address() as AddressInfo).port;
  const redirectUri = `http://localhost:${port}`;

  const authUrl = await app.getAuthCodeUrl({
    scopes,
    redirectUri,
    codeChallenge: challenge,
    codeChallengeMethod: "S256",
    state,
    prompt: "select_account",
  });

  const code = new Promise<string>((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error(`Browser sign-in timed out after ${BROWSER_LOGIN_TIMEOUT_MS / 60000} minutes`)), BROWSER_LOGIN_TIMEOUT_MS);
    server.on("request", (req, res) => {
      const url = new URL(req.url ?? "/", redirectUri);
      if (url.pathname !== "/") { res.writeHead(404).end(); return; }
      const err = url.searchParams.get("error");
      const returnedState = url.searchParams.get("state");
      const returnedCode = url.searchParams.get("code");
      const fail = (reason: string) => {
        res.writeHead(400, { "Content-Type": "text/html; charset=utf-8" }).end(LOGIN_FAILED_HTML(escapeHtml(reason)));
        clearTimeout(timer);
        reject(new Error(reason));
      };
      if (err) return fail(`${err}: ${url.searchParams.get("error_description") ?? ""}`.trim());
      if (returnedState !== state) return fail("State mismatch in the sign-in redirect; please retry");
      if (!returnedCode) return fail("No authorization code in the sign-in redirect");
      res.writeHead(200, { "Content-Type": "text/html; charset=utf-8" }).end(LOGIN_DONE_HTML);
      clearTimeout(timer);
      resolve(returnedCode);
    });
  });

  opts.onMessage(`Open this URL in your browser to sign in (it should open automatically):\n${authUrl}`);
  tryOpenBrowser(authUrl, opts.log);

  try {
    const authCode = await code;
    return await app.acquireTokenByCode({ code: authCode, scopes, redirectUri, codeVerifier: verifier });
  } finally {
    server.close();
  }
}

/** Run the interactive flow selected by `cfg.loginFlow`. */
async function acquireTokenInteractively(
  cfg: AppConfig,
  app: PublicClientApplication,
  scopes: string[],
  onMessage: (message: string) => void,
  log?: Logger,
): Promise<AuthenticationResult> {
  if (cfg.loginFlow === "device-code") {
    const result = await app.acquireTokenByDeviceCode(deviceCodeRequest(scopes, onMessage));
    if (!result?.accessToken) throw new Error("Device-code sign-in did not complete");
    return result;
  }
  return acquireTokenByBrowser(app, scopes, { onMessage, port: cfg.loginRedirectPort, log });
}

/** Client-credentials token provider (daemon / webhook service). */
export function createAppTokenProvider(cfg: AppConfig): TokenProvider {
  if (!cfg.clientSecret) throw new Error("AUTH_MODE=app requires CLIENT_SECRET");
  const app = new ConfidentialClientApplication({
    auth: { clientId: cfg.clientId, authority: authorityFor(cfg), clientSecret: cfg.clientSecret },
  });
  const scopes = [`${cfg.graphBaseUrl.replace(/\/+$/, "")}/.default`];
  let cached: CachedToken | undefined;
  let inflight: Promise<string> | undefined;

  const acquire = async (): Promise<string> => {
    const result = await app.acquireTokenByClientCredential({ scopes });
    if (!result?.accessToken) throw new Error("Client-credentials flow returned no access token");
    cached = { token: result.accessToken, expiresAt: expiresAtOf(result) };
    return cached.token;
  };

  return {
    async getToken(): Promise<string> {
      if (isFresh(cached)) return cached.token;
      if (!inflight) inflight = acquire().finally(() => { inflight = undefined; });
      return inflight;
    },
  };
}

/** Delegated token provider using the on-disk MSAL cache (CLI / MCP server). */
export function createDelegatedTokenProvider(cfg: AppConfig, opts: DelegatedTokenOptions = {}): TokenProvider {
  const log = opts.logger;
  const app = createPublicClient(cfg, log);
  const scopes = delegatedScopesFor(cfg);
  const onMessage = opts.onMessage ?? defaultMessagePrinter;
  let cached: CachedToken | undefined;
  let inflight: Promise<string> | undefined;

  const acquire = async (): Promise<string> => {
    let result: AuthenticationResult | null = null;
    const account = await firstAccount(app);
    if (account) {
      try {
        result = await app.acquireTokenSilent({ account, scopes });
      } catch (err) {
        log?.debug("silent token acquisition failed", { error: (err as Error).message });
        result = null;
      }
    }
    if (!result?.accessToken) {
      if (!opts.allowInteractive) {
        throw new Error(
          `No usable delegated token in ${cfg.tokenCachePath}. Sign in first: run \`npm run login\`.`,
        );
      }
      result = await acquireTokenInteractively(cfg, app, scopes, onMessage, log);
    }
    if (!result?.accessToken) throw new Error("Interactive sign-in returned no access token");
    cached = { token: result.accessToken, expiresAt: expiresAtOf(result) };
    return cached.token;
  };

  return {
    async getToken(): Promise<string> {
      if (isFresh(cached)) return cached.token;
      if (!inflight) inflight = acquire().finally(() => { inflight = undefined; });
      return inflight;
    },
    async account(): Promise<string | undefined> {
      return (await firstAccount(app))?.username;
    },
  };
}

/** Pick the provider matching `cfg.authMode`. */
export function tokenProviderFromConfig(cfg: AppConfig, opts: DelegatedTokenOptions = {}): TokenProvider {
  return cfg.authMode === "app" ? createAppTokenProvider(cfg) : createDelegatedTokenProvider(cfg, opts);
}

/** Force an interactive sign-in and persist the resulting cache. Used by `npm run login`. */
export async function login(cfg: AppConfig, log: Logger): Promise<{ username?: string }> {
  if (cfg.authMode === "app") {
    log.warn("AUTH_MODE=app does not use an interactive sign-in; verifying the client secret instead");
    await createAppTokenProvider(cfg).getToken();
    return {};
  }
  const app = createPublicClient(cfg, log);
  const scopes = delegatedScopesFor(cfg);
  log.info("starting sign-in", { flow: cfg.loginFlow, scopes, cache: cfg.tokenCachePath });
  const result = await acquireTokenInteractively(cfg, app, scopes, (message) => process.stderr.write(message + "\n"), log);
  if (!result?.accessToken) throw new Error("Sign-in did not complete");
  const username = result.account?.username;
  log.info("signed in", { username, cache: cfg.tokenCachePath });
  return { username };
}
