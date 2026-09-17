/**
 * Tiny Microsoft Graph HTTP client built on global `fetch`.
 *
 * Handles: bearer auth, JSON bodies, throttling (429) and transient server errors (503/504) with
 * `Retry-After` support, `@odata.nextLink` paging, and turning Graph's error envelope into GraphError.
 * Tokens are never logged.
 */
import type { TokenProvider } from "../auth.js";
import type { Logger } from "../log.js";
import type { ODataCollection } from "./types.js";

export class GraphError extends Error {
  readonly status: number;
  readonly code?: string;
  readonly body?: unknown;
  readonly retryAfterMs?: number;

  constructor(
    message: string,
    status: number,
    opts: { code?: string; body?: unknown; retryAfterMs?: number } = {},
  ) {
    super(message);
    this.name = "GraphError";
    this.status = status;
    this.code = opts.code;
    this.body = opts.body;
    this.retryAfterMs = opts.retryAfterMs;
  }
}

export interface RequestOptions {
  query?: Record<string, string | number | boolean | undefined>;
  headers?: Record<string, string>;
  version?: "v1.0" | "beta";
  /** Return the raw Response instead of parsed JSON. */
  raw?: boolean;
}

export interface GraphClientOptions {
  baseUrl?: string;
  fetchImpl?: typeof fetch;
  logger?: Logger;
  maxRetries?: number;
}

const RETRYABLE_STATUS = new Set([429, 503, 504]);
const DEFAULT_MAX_RETRIES = 4;

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

/** `Retry-After` is either delta-seconds or an HTTP date. Returns milliseconds, or undefined. */
export function parseRetryAfter(value: string | null, now: number = Date.now()): number | undefined {
  if (!value) return undefined;
  const trimmed = value.trim();
  if (/^\d+$/.test(trimmed)) return Number(trimmed) * 1000;
  const when = Date.parse(trimmed);
  if (Number.isFinite(when)) return Math.max(0, when - now);
  return undefined;
}

function encodeQuery(query: Record<string, string | number | boolean | undefined> | undefined): string {
  if (!query) return "";
  const parts: string[] = [];
  for (const [key, value] of Object.entries(query)) {
    if (value === undefined) continue;
    parts.push(`${encodeURIComponent(key)}=${encodeURIComponent(String(value))}`);
  }
  return parts.join("&");
}

function isAbsolute(pathOrUrl: string): boolean {
  return /^https?:\/\//i.test(pathOrUrl);
}

export class GraphClient {
  private readonly tokens: TokenProvider;
  private readonly baseUrl: string;
  private readonly fetchImpl: typeof fetch;
  private readonly log?: Logger;
  private readonly maxRetries: number;

  constructor(tokens: TokenProvider, opts: GraphClientOptions = {}) {
    this.tokens = tokens;
    this.baseUrl = (opts.baseUrl ?? "https://graph.microsoft.com").replace(/\/+$/, "");
    this.fetchImpl = opts.fetchImpl ?? ((input, init) => fetch(input, init));
    this.log = opts.logger;
    this.maxRetries = opts.maxRetries ?? DEFAULT_MAX_RETRIES;
  }

  /**
   * Absolute URLs (e.g. an `@odata.nextLink`) are used verbatim so their already-encoded query is
   * never encoded a second time; relative paths are resolved against `<baseUrl>/<version>`.
   */
  private buildUrl(pathOrUrl: string, opts: RequestOptions | undefined): string {
    const query = encodeQuery(opts?.query);
    if (isAbsolute(pathOrUrl)) {
      if (!query) return pathOrUrl;
      return pathOrUrl + (pathOrUrl.includes("?") ? "&" : "?") + query;
    }
    const version = opts?.version ?? "v1.0";
    const rel = pathOrUrl.startsWith("/") ? pathOrUrl : `/${pathOrUrl}`;
    const url = `${this.baseUrl}/${version}${rel}`;
    if (!query) return url;
    return url + (url.includes("?") ? "&" : "?") + query;
  }

  private async request(
    method: string,
    pathOrUrl: string,
    body: unknown,
    opts: RequestOptions = {},
  ): Promise<unknown> {
    const url = this.buildUrl(pathOrUrl, opts);
    const hasBody = body !== undefined && body !== null;
    const logPath = isAbsolute(pathOrUrl) ? pathOrUrl.split("?")[0] : pathOrUrl;

    let lastError: unknown;
    for (let attempt = 0; attempt <= this.maxRetries; attempt++) {
      const token = await this.tokens.getToken();
      const headers: Record<string, string> = {
        Authorization: `Bearer ${token}`,
        Accept: "application/json",
        ...(opts.headers ?? {}),
      };
      if (hasBody) headers["Content-Type"] = headers["Content-Type"] ?? "application/json";

      this.log?.debug("graph request", { method, path: logPath, attempt });

      let res: Response;
      try {
        res = await this.fetchImpl(url, {
          method,
          headers,
          body: hasBody ? JSON.stringify(body) : undefined,
        });
      } catch (err) {
        // Network-level failure: retry with backoff, otherwise surface it.
        lastError = err;
        if (attempt < this.maxRetries) {
          await sleep(backoffMs(attempt));
          continue;
        }
        throw err;
      }

      if (RETRYABLE_STATUS.has(res.status) && attempt < this.maxRetries) {
        const retryAfterMs = parseRetryAfter(res.headers.get("retry-after")) ?? backoffMs(attempt);
        this.log?.debug("graph retry", { method, path: logPath, status: res.status, retryAfterMs });
        await res.text().catch(() => undefined);
        await sleep(retryAfterMs);
        continue;
      }

      this.log?.debug("graph response", { method, path: logPath, status: res.status });

      if (!res.ok && res.status !== 207) {
        throw await toGraphError(res);
      }
      if (opts.raw) return res;
      return parseBody(res);
    }
    // Only reachable if every attempt was a network failure.
    throw lastError instanceof Error ? lastError : new Error(String(lastError));
  }

  async get<T>(path: string, opts?: RequestOptions): Promise<T> {
    return (await this.request("GET", path, undefined, opts)) as T;
  }

  async post<T>(path: string, body: unknown, opts?: RequestOptions): Promise<T> {
    return (await this.request("POST", path, body, opts)) as T;
  }

  async patch<T>(path: string, body: unknown, opts?: RequestOptions): Promise<T> {
    return (await this.request("PATCH", path, body, opts)) as T;
  }

  async delete(path: string, opts?: RequestOptions): Promise<void> {
    await this.request("DELETE", path, undefined, opts);
  }

  /** GET a collection, following `@odata.nextLink` until exhausted. */
  async getAll<T>(path: string, opts?: RequestOptions): Promise<T[]> {
    const out: T[] = [];
    let page = await this.get<ODataCollection<T> | undefined>(path, opts);
    let guard = 0;
    while (page) {
      if (Array.isArray(page.value)) out.push(...page.value);
      const next = page["@odata.nextLink"];
      if (!next || ++guard > 1000) break;
      // nextLink is absolute and fully encoded: pass it through without the original query.
      page = await this.get<ODataCollection<T> | undefined>(next, {
        headers: opts?.headers,
        version: opts?.version,
      });
    }
    return out;
  }
}

/** Jittered exponential backoff: ~0.5s, 1s, 2s, 4s (+/- 25%). */
function backoffMs(attempt: number): number {
  const base = 500 * Math.pow(2, attempt);
  const jitter = base * 0.25 * (Math.random() * 2 - 1);
  return Math.max(100, Math.round(base + jitter));
}

async function parseBody(res: Response): Promise<unknown> {
  if (res.status === 204 || res.status === 205) return undefined;
  const text = await res.text().catch(() => "");
  if (!text) return undefined;
  const contentType = res.headers.get("content-type") ?? "";
  if (contentType.includes("json") || text.startsWith("{") || text.startsWith("[")) {
    try {
      return JSON.parse(text);
    } catch {
      return text;
    }
  }
  return text;
}

async function toGraphError(res: Response): Promise<GraphError> {
  const body = await parseBody(res).catch(() => undefined);
  let code: string | undefined;
  let message: string | undefined;
  if (body && typeof body === "object") {
    const err = (body as { error?: unknown }).error;
    if (typeof err === "string") {
      code = err;
      message = (body as { error_description?: string }).error_description;
    } else if (err && typeof err === "object") {
      const e = err as { code?: unknown; message?: unknown };
      if (typeof e.code === "string") code = e.code;
      if (typeof e.message === "string") message = e.message;
    }
  } else if (typeof body === "string" && body) {
    message = body;
  }
  return new GraphError(
    `Graph ${res.status}${code ? ` ${code}` : ""}: ${message ?? res.statusText ?? "request failed"}`,
    res.status,
    { code, body, retryAfterMs: parseRetryAfter(res.headers.get("retry-after")) },
  );
}
