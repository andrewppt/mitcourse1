/**
 * Token acquisition for Microsoft Graph, backed by msal-node.
 *
 *  - "app" mode:       ConfidentialClientApplication + client credentials (`<graphBaseUrl>/.default`).
 *  - "delegated" mode: PublicClientApplication + device code, with the MSAL token cache persisted
 *                      to `cfg.tokenCachePath` (mode 0600) so the CLI/MCP server only signs in once.
 *
 * Nothing in here ever writes to stdout and no token value is ever logged.
 */
import fs from "node:fs";
import fsp from "node:fs/promises";
import path from "node:path";
import {
  ConfidentialClientApplication,
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
  /** Called with the device-code message that must be shown to the user. Defaults to stderr. */
  onDeviceCode?: (message: string) => void;
  /** Allow an interactive device-code sign-in when the silent flow cannot produce a token. */
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

const defaultDeviceCodePrinter = (message: string): void => {
  process.stderr.write(message + "\n");
};

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

/** Device-code token provider using the on-disk MSAL cache (CLI / MCP server). */
export function createDelegatedTokenProvider(cfg: AppConfig, opts: DelegatedTokenOptions = {}): TokenProvider {
  const log = opts.logger;
  const app = createPublicClient(cfg, log);
  const scopes = delegatedScopesFor(cfg);
  const onDeviceCode = opts.onDeviceCode ?? defaultDeviceCodePrinter;
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
      result = await app.acquireTokenByDeviceCode(deviceCodeRequest(scopes, onDeviceCode));
    }
    if (!result?.accessToken) throw new Error("Device-code flow returned no access token");
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

/** Force a device-code sign-in and persist the resulting cache. Used by `npm run login`. */
export async function login(cfg: AppConfig, log: Logger): Promise<{ username?: string }> {
  if (cfg.authMode === "app") {
    log.warn("AUTH_MODE=app does not use an interactive sign-in; verifying the client secret instead");
    await createAppTokenProvider(cfg).getToken();
    return {};
  }
  const app = createPublicClient(cfg, log);
  const scopes = delegatedScopesFor(cfg);
  log.info("starting device-code sign-in", { scopes, cache: cfg.tokenCachePath });
  const result = await app.acquireTokenByDeviceCode(
    deviceCodeRequest(scopes, (message) => log.info(message)),
  );
  if (!result?.accessToken) throw new Error("Device-code sign-in did not complete");
  const username = result.account?.username;
  log.info("signed in", { username, cache: cfg.tokenCachePath });
  return { username };
}
