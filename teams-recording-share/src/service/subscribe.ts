/**
 * Create / renew / list / delete the Microsoft Graph change subscriptions that tell us
 * when a Teams meeting recording is ready.
 *
 * Tenant-wide:   communications/onlineMeetings/getAllRecordings
 * Per organizer: users/{id}/onlineMeetings/getAllRecordings
 *
 * Also usable as a script:  npm run subscribe -- create|renew|list|delete
 */
import { pathToFileURL } from "node:url";
import type { AppConfig } from "../config.js";
import { loadConfig, loadDotEnv } from "../config.js";
import { createLogger, type Logger } from "../log.js";
import type { Subscription } from "../graph/types.js";
import { GraphClient, GraphError } from "../graph/client.js";
import { tokenProviderFromConfig } from "../auth.js";
import { StateStore, type SubscriptionRecord } from "./store.js";

/** Graph caps `getAllRecordings` subscriptions at 4320 minutes; stay a little under. */
export const SUBSCRIPTION_MINUTES = 4230;
/** Renew when the remaining lifetime drops below this. */
export const RENEW_WITHIN_MS = 12 * 60 * 60 * 1000;

export function notificationUrl(cfg: AppConfig): string {
  return `${cfg.publicBaseUrl}/webhook/notifications`;
}
export function lifecycleUrl(cfg: AppConfig): string {
  return `${cfg.publicBaseUrl}/webhook/lifecycle`;
}

/** Resources we want a `created` subscription on. */
export function desiredResources(cfg: AppConfig): string[] {
  return cfg.organizerUserIds.length
    ? cfg.organizerUserIds.map((id) => `users/${id}/onlineMeetings/getAllRecordings`)
    : ["communications/onlineMeetings/getAllRecordings"];
}

/** Throw a clear error when the service is not configured well enough to receive notifications. */
export function assertSubscriptionConfig(cfg: AppConfig): void {
  if (!cfg.publicBaseUrl) {
    throw new Error(
      "PUBLIC_BASE_URL is not set. Graph must be able to reach this service over the public internet " +
        "(e.g. an ngrok https URL or your Azure app URL) before subscriptions can be created.",
    );
  }
  if (!/^https:\/\//i.test(cfg.publicBaseUrl)) {
    throw new Error(`PUBLIC_BASE_URL must be an https URL, got "${cfg.publicBaseUrl}"`);
  }
  if (!cfg.clientState) {
    throw new Error("CLIENT_STATE is empty. Set it to a random secret; it is echoed in every notification and verified by the webhook.");
  }
}

function expirationFromNow(minutes = SUBSCRIPTION_MINUTES): string {
  return new Date(Date.now() + minutes * 60_000).toISOString();
}

/** Graph echoes resources back with slightly different casing/slashes; compare forgivingly. */
function sameResource(a: string | undefined, b: string | undefined): boolean {
  const norm = (s: string | undefined) => (s ?? "").replace(/^\/+|\/+$/g, "").toLowerCase();
  return norm(a) === norm(b);
}
function sameUrl(a: string | undefined, b: string | undefined): boolean {
  const norm = (s: string | undefined) => (s ?? "").replace(/\/+$/, "").toLowerCase();
  return norm(a) === norm(b);
}

function toRecord(sub: Subscription): SubscriptionRecord {
  return {
    id: sub.id,
    resource: sub.resource,
    expirationDateTime: sub.expirationDateTime,
    notificationUrl: sub.notificationUrl,
    lifecycleNotificationUrl: sub.lifecycleNotificationUrl,
    changeType: sub.changeType,
    updatedAt: new Date().toISOString(),
  };
}

/** Every subscription Graph currently holds for this app. */
export async function listSubscriptions(g: GraphClient): Promise<Subscription[]> {
  return g.getAll<Subscription>("/subscriptions");
}

/**
 * Make sure a subscription exists for every desired resource.
 * Existing ones (same resource + notificationUrl) are reused, and renewed when they expire soon.
 */
export async function ensureSubscriptions(
  g: GraphClient,
  cfg: AppConfig,
  store: StateStore,
  log: Logger,
  opts: { renewWithinMs?: number } = {},
): Promise<SubscriptionRecord[]> {
  assertSubscriptionConfig(cfg);
  const renewWithinMs = opts.renewWithinMs ?? RENEW_WITHIN_MS;
  const notifyUrl = notificationUrl(cfg);
  const lifeUrl = lifecycleUrl(cfg);

  let existing: Subscription[] = [];
  try {
    existing = await listSubscriptions(g);
  } catch (err) {
    log.warn("could not list existing subscriptions; will try to create", { error: errorMessage(err) });
  }

  const out: SubscriptionRecord[] = [];
  for (const resource of desiredResources(cfg)) {
    const match = existing.find((s) => sameResource(s.resource, resource) && sameUrl(s.notificationUrl, notifyUrl));
    if (match) {
      const msLeft = Date.parse(match.expirationDateTime) - Date.now();
      if (Number.isFinite(msLeft) && msLeft > renewWithinMs) {
        log.info("subscription already active", { id: match.id, resource, expirationDateTime: match.expirationDateTime });
        const rec = toRecord(match);
        store.upsertSubscription(rec);
        out.push(rec);
        continue;
      }
      try {
        const renewed = await g.patch<Subscription>(`/subscriptions/${match.id}`, { expirationDateTime: expirationFromNow() });
        const rec = toRecord({ ...match, ...renewed });
        store.upsertSubscription(rec);
        out.push(rec);
        log.info("subscription renewed", { id: rec.id, resource, expirationDateTime: rec.expirationDateTime });
        continue;
      } catch (err) {
        log.warn("renew failed; recreating subscription", { id: match.id, resource, error: errorMessage(err) });
        store.removeSubscription(match.id);
      }
    }

    const created = await g.post<Subscription>("/subscriptions", {
      changeType: "created",
      notificationUrl: notifyUrl,
      lifecycleNotificationUrl: lifeUrl,
      resource,
      expirationDateTime: expirationFromNow(),
      clientState: cfg.clientState,
    });
    const rec = toRecord(created);
    store.upsertSubscription(rec);
    out.push(rec);
    log.info("subscription created", { id: rec.id, resource, expirationDateTime: rec.expirationDateTime });
  }
  return out;
}

/**
 * Extend every subscription we know about. Anything Graph no longer has (404/410) is
 * dropped from the store; `ensureSubscriptions` then recreates whatever is still desired.
 */
export async function renewSubscriptions(
  g: GraphClient,
  cfg: AppConfig,
  store: StateStore,
  log: Logger,
): Promise<SubscriptionRecord[]> {
  assertSubscriptionConfig(cfg);
  for (const rec of store.getSubscriptions()) {
    try {
      const renewed = await g.patch<Subscription>(`/subscriptions/${rec.id}`, { expirationDateTime: expirationFromNow() });
      const next = toRecord({ ...(rec as unknown as Subscription), ...renewed });
      store.upsertSubscription(next);
      log.info("subscription renewed", { id: next.id, resource: next.resource, expirationDateTime: next.expirationDateTime });
    } catch (err) {
      const status = err instanceof GraphError ? err.status : undefined;
      if (status === 404 || status === 410) {
        log.warn("subscription gone; dropping from state", { id: rec.id, resource: rec.resource });
        store.removeSubscription(rec.id);
      } else {
        log.error("subscription renew failed", { id: rec.id, resource: rec.resource, error: errorMessage(err) });
        store.recordError(`renew ${rec.id}: ${errorMessage(err)}`);
      }
    }
  }
  // Recreate anything missing and cover resources added to the config since the last run.
  return ensureSubscriptions(g, cfg, store, log);
}

/** Renew one subscription by id (used by the `reauthorizationRequired` lifecycle event). */
export async function renewSubscription(
  g: GraphClient,
  store: StateStore,
  subscriptionId: string,
  log: Logger,
): Promise<SubscriptionRecord | null> {
  try {
    const renewed = await g.patch<Subscription>(`/subscriptions/${subscriptionId}`, { expirationDateTime: expirationFromNow() });
    const known = store.getSubscriptions().find((s) => s.id === subscriptionId);
    const rec: SubscriptionRecord = {
      id: renewed?.id ?? subscriptionId,
      resource: renewed?.resource ?? known?.resource ?? "",
      expirationDateTime: renewed?.expirationDateTime ?? expirationFromNow(),
      notificationUrl: renewed?.notificationUrl ?? known?.notificationUrl,
      lifecycleNotificationUrl: renewed?.lifecycleNotificationUrl ?? known?.lifecycleNotificationUrl,
      changeType: renewed?.changeType ?? known?.changeType,
      updatedAt: new Date().toISOString(),
    };
    store.upsertSubscription(rec);
    log.info("subscription reauthorized", { id: rec.id, expirationDateTime: rec.expirationDateTime });
    return rec;
  } catch (err) {
    log.error("reauthorization renew failed", { id: subscriptionId, error: errorMessage(err) });
    store.recordError(`reauthorize ${subscriptionId}: ${errorMessage(err)}`);
    return null;
  }
}

/** Delete every subscription recorded in the store. */
export async function removeSubscriptions(g: GraphClient, store: StateStore, log: Logger): Promise<string[]> {
  const removed: string[] = [];
  for (const rec of store.getSubscriptions()) {
    try {
      await g.delete(`/subscriptions/${rec.id}`);
      log.info("subscription deleted", { id: rec.id, resource: rec.resource });
    } catch (err) {
      const status = err instanceof GraphError ? err.status : undefined;
      if (status !== 404 && status !== 410) {
        log.error("subscription delete failed", { id: rec.id, error: errorMessage(err) });
        store.recordError(`delete ${rec.id}: ${errorMessage(err)}`);
        continue;
      }
      log.warn("subscription already gone", { id: rec.id });
    }
    store.removeSubscription(rec.id);
    removed.push(rec.id);
  }
  return removed;
}

export function errorMessage(err: unknown): string {
  if (err instanceof GraphError) return `${err.status}${err.code ? ` ${err.code}` : ""}: ${err.message}`;
  return err instanceof Error ? err.message : String(err);
}

// ---------------------------------------------------------------------------
// Script mode: npm run subscribe -- create|renew|list|delete
// ---------------------------------------------------------------------------

export async function main(argv: string[] = process.argv.slice(2)): Promise<number> {
  const command = (argv[0] ?? "create").toLowerCase();
  if (!["create", "ensure", "renew", "list", "delete", "remove"].includes(command)) {
    process.stderr.write("Usage: npm run subscribe -- create|renew|list|delete\n");
    return 2;
  }
  loadDotEnv();
  const cfg = loadConfig();
  const log = createLogger("subscribe");
  const graph = new GraphClient(tokenProviderFromConfig(cfg), { baseUrl: cfg.graphBaseUrl, logger: log });
  const store = new StateStore(cfg.stateDir);
  store.load();

  const print = (v: unknown) => process.stdout.write(JSON.stringify(v, null, 2) + "\n");

  switch (command) {
    case "create":
    case "ensure": {
      print(await ensureSubscriptions(graph, cfg, store, log));
      return 0;
    }
    case "renew": {
      print(await renewSubscriptions(graph, cfg, store, log));
      return 0;
    }
    case "list": {
      const remote = await listSubscriptions(graph);
      print({ desired: desiredResources(cfg), stored: store.getSubscriptions(), graph: remote });
      return 0;
    }
    case "delete":
    case "remove": {
      print({ deleted: await removeSubscriptions(graph, store, log) });
      return 0;
    }
    /* c8 ignore next 3 */
    default:
      process.stderr.write("Usage: npm run subscribe -- create|renew|list|delete\n");
      return 2;
  }
}

const invokedDirectly = process.argv[1] !== undefined && import.meta.url === pathToFileURL(process.argv[1]).href;
if (invokedDirectly) {
  main().then(
    (code) => {
      process.exitCode = code;
    },
    (err: unknown) => {
      process.stderr.write(`${errorMessage(err)}\n`);
      process.exitCode = 1;
    },
  );
}
