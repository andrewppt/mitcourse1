/**
 * Express 5 webhook service.
 *
 * Microsoft Graph calls:
 *   POST /webhook/notifications   – "a recording is ready" change notifications
 *   POST /webhook/lifecycle       – reauthorizationRequired / subscriptionRemoved / missed
 * Both endpoints implement the validation handshake (`?validationToken=...` → 200 text/plain).
 *
 * Notifications are acknowledged with 202 straight away (Graph gives us ~3s) and handled
 * afterwards on a small in-process queue, one recording at a time.
 */
import type { Express, NextFunction, Request, Response } from "express";
import express from "express";
import { pathToFileURL } from "node:url";
import type { AppConfig } from "../config.js";
import { loadConfig, loadDotEnv } from "../config.js";
import { createLogger, type Logger } from "../log.js";
import type { ChangeNotification } from "../graph/types.js";
import { GraphClient } from "../graph/client.js";
import { tokenProviderFromConfig } from "../auth.js";
import { decodeOnlineMeetingId, parseNotificationResource } from "../graph/meetings.js";
import { shareRecordingWithAttendees } from "../core/share.js";
import { StateStore } from "./store.js";
import { ensureSubscriptions, errorMessage, renewSubscription, renewSubscriptions } from "./subscribe.js";

export interface ServerDeps {
  cfg: AppConfig;
  graph: GraphClient;
  store: StateStore;
  log: Logger;
  /** Injectable for tests; defaults to the real sharing pipeline. */
  share?: typeof shareRecordingWithAttendees;
  /** Called once each accepted notification has been handled (tests use it to await the queue). */
  onNotification?: (n: ChangeNotification) => void;
}

/** Every 6 hours; Graph subscriptions live at most 3 days. */
export const RENEW_INTERVAL_MS = 6 * 60 * 60 * 1000;

function isRecordingNotification(n: ChangeNotification): boolean {
  const odataType = String(n.resourceData?.["@odata.type"] ?? "").trim();
  if (/callrecording$/i.test(odataType)) return true;
  return (n.resource ?? "").toLowerCase().includes("/recordings(");
}

/** Sequential, in-process work queue. One recording is shared at a time. */
class NotificationQueue {
  private chain: Promise<void> = Promise.resolve();
  private pending = 0;

  push(task: () => Promise<void>): void {
    this.pending += 1;
    this.chain = this.chain.then(task).catch(() => undefined).finally(() => {
      this.pending -= 1;
    });
  }

  get size(): number {
    return this.pending;
  }

  /** Resolves when everything queued so far has finished. */
  async drain(): Promise<void> {
    await this.chain;
  }
}

export function createApp(deps: ServerDeps): Express {
  const { cfg, graph, store, log } = deps;
  const share = deps.share ?? shareRecordingWithAttendees;
  const queue = new NotificationQueue();
  const app = express();

  app.disable("x-powered-by");
  // Exposed for tests: `drainApp(app)` waits for queued notifications to finish.
  app.locals.queue = queue;
  app.use(express.json({ limit: "1mb", type: ["application/json", "application/*+json"] }));

  /** Graph's endpoint validation: echo the token back as text/plain within 10s. */
  const handshake = (req: Request, res: Response): boolean => {
    const token = req.query.validationToken;
    if (token === undefined) return false;
    const value = Array.isArray(token) ? String(token[0]) : String(token);
    res.status(200).type("text/plain").send(value);
    return true;
  };

  const notifications = (body: unknown): ChangeNotification[] => {
    const value = (body as { value?: unknown } | undefined)?.value;
    return Array.isArray(value) ? (value as ChangeNotification[]) : [];
  };

  /** Notifications that do not carry our secret are not ours. */
  const verified = (list: ChangeNotification[]): ChangeNotification[] =>
    list.filter((n) => {
      if (n?.clientState === cfg.clientState) return true;
      log.warn("dropping notification with bad clientState", { subscriptionId: n?.subscriptionId, resource: n?.resource });
      return false;
    });

  async function handleRecording(n: ChangeNotification): Promise<void> {
    if (!isRecordingNotification(n)) {
      log.debug("ignoring non-recording notification", { resource: n.resource, type: n.resourceData?.["@odata.type"] });
      return;
    }
    const parsed = parseNotificationResource(n.resource ?? "");
    const meetingId = parsed?.meetingId;
    const recordingId = parsed?.recordingId ?? (typeof n.resourceData?.id === "string" ? n.resourceData.id : undefined);
    if (!meetingId || !recordingId) {
      log.error("could not parse recording notification resource", { resource: n.resource });
      store.recordError(`unparsable notification resource: ${n.resource}`);
      return;
    }
    const organizerUserId = parsed?.userId ?? decodeOnlineMeetingId(meetingId)?.organizerId;
    if (!organizerUserId) {
      log.error("no organizer could be derived from notification", { resource: n.resource, meetingId });
      store.recordError(`no organizer for meeting ${meetingId}`);
      return;
    }
    if (store.isProcessed(recordingId)) {
      log.info("recording already processed; skipping", { recordingId, meetingId });
      return;
    }

    log.info("sharing recording", { recordingId, meetingId, organizerUserId });
    try {
      const result = await share(graph, cfg, { organizerUserId, meetingId, recordingId }, log);
      let granted = 0;
      let skipped = 0;
      let errors = 0;
      let driveItemId: string | undefined;
      for (const rec of result.recordings ?? []) {
        granted += rec.granted?.length ?? 0;
        skipped += rec.skipped?.length ?? 0;
        errors += rec.errors?.length ?? 0;
        driveItemId ??= rec.driveItem?.id;
      }
      store.markProcessed(recordingId, { meetingId, driveItemId, granted, skipped, errors, at: new Date().toISOString() });
      log.info("recording shared", { recordingId, meetingId, driveItemId, granted, skipped, errors });
    } catch (err) {
      const msg = errorMessage(err);
      log.error("sharing recording failed", { recordingId, meetingId, organizerUserId, error: msg });
      store.recordError(`share ${recordingId}: ${msg}`);
    }
  }

  app.post("/webhook/notifications", (req: Request, res: Response) => {
    if (handshake(req, res)) return;
    const accepted = verified(notifications(req.body));
    res.status(202).json({ accepted: accepted.length });
    for (const n of accepted) {
      queue.push(async () => {
        try {
          await handleRecording(n);
        } finally {
          deps.onNotification?.(n);
        }
      });
    }
  });

  app.post("/webhook/lifecycle", (req: Request, res: Response) => {
    if (handshake(req, res)) return;
    const accepted = verified(notifications(req.body));
    res.status(202).json({ accepted: accepted.length });
    for (const n of accepted) {
      queue.push(async () => {
        try {
          switch (n.lifecycleEvent) {
            case "reauthorizationRequired":
              log.info("lifecycle: reauthorizationRequired", { subscriptionId: n.subscriptionId });
              await renewSubscription(graph, store, n.subscriptionId, log);
              break;
            case "subscriptionRemoved":
              log.warn("lifecycle: subscriptionRemoved; recreating", { subscriptionId: n.subscriptionId });
              store.removeSubscription(n.subscriptionId);
              await ensureSubscriptions(graph, cfg, store, log);
              break;
            case "missed":
              log.warn("lifecycle: missed notifications", { subscriptionId: n.subscriptionId, resource: n.resource });
              break;
            default:
              log.warn("lifecycle: unknown event", { subscriptionId: n.subscriptionId, lifecycleEvent: n.lifecycleEvent });
          }
        } catch (err) {
          const msg = errorMessage(err);
          log.error("lifecycle handling failed", { subscriptionId: n.subscriptionId, lifecycleEvent: n.lifecycleEvent, error: msg });
          store.recordError(`lifecycle ${n.lifecycleEvent ?? "?"} ${n.subscriptionId}: ${msg}`);
        } finally {
          deps.onNotification?.(n);
        }
      });
    }
  });

  // Graph also probes with GET during endpoint validation.
  app.get("/webhook/notifications", (req: Request, res: Response) => {
    if (handshake(req, res)) return;
    res.status(405).json({ error: "use POST" });
  });
  app.get("/webhook/lifecycle", (req: Request, res: Response) => {
    if (handshake(req, res)) return;
    res.status(405).json({ error: "use POST" });
  });

  app.get("/healthz", (_req: Request, res: Response) => {
    res.json({ ok: true, subscriptions: store.getSubscriptions().length, processed: store.processedCount() });
  });

  app.get("/status", (_req: Request, res: Response) => {
    res.json({
      ok: true,
      authMode: cfg.authMode,
      publicBaseUrl: cfg.publicBaseUrl ?? null,
      organizerUserIds: cfg.organizerUserIds,
      queued: queue.size,
      subscriptions: store.getSubscriptions().map((s) => ({
        id: s.id,
        resource: s.resource,
        expirationDateTime: s.expirationDateTime,
        notificationUrl: s.notificationUrl,
      })),
      processed: store.listProcessed(25),
      errors: store.listErrors(10),
    });
  });

  app.use((err: unknown, _req: Request, res: Response, next: NextFunction) => {
    if (res.headersSent) return next(err);
    const msg = err instanceof Error ? err.message : String(err);
    log.warn("bad request", { error: msg });
    res.status(400).json({ error: msg });
  });

  return app;
}

/** Exposed so tests (and `start`) can wait for the queue; `createApp` keeps its own instance. */
export async function drainApp(app: Express): Promise<void> {
  const queue = (app as unknown as { locals?: { queue?: { drain(): Promise<void> } } }).locals?.queue;
  await queue?.drain();
}

/** Periodic renewal so subscriptions never lapse (unref'd: it must not hold the process open). */
export function startRenewalTimer(deps: ServerDeps, intervalMs = RENEW_INTERVAL_MS): NodeJS.Timeout {
  const timer = setInterval(() => {
    renewSubscriptions(deps.graph, deps.cfg, deps.store, deps.log).catch((err: unknown) => {
      deps.log.error("scheduled renewal failed", { error: errorMessage(err) });
    });
  }, intervalMs);
  timer.unref();
  return timer;
}

export async function start(): Promise<void> {
  loadDotEnv();
  const cfg = loadConfig();
  const log = createLogger("service");
  const store = new StateStore(cfg.stateDir);
  store.load();
  const graph = new GraphClient(tokenProviderFromConfig(cfg), { baseUrl: cfg.graphBaseUrl, logger: log });
  const deps: ServerDeps = { cfg, graph, store, log };
  const app = createApp(deps);

  try {
    await ensureSubscriptions(graph, cfg, store, log);
  } catch (err) {
    log.error("could not ensure subscriptions at startup (service still starting)", { error: errorMessage(err) });
  }

  const timer = startRenewalTimer(deps);
  const server = app.listen(cfg.port, () => {
    const base = cfg.publicBaseUrl ?? `http://localhost:${cfg.port}`;
    log.info("webhook service listening", {
      port: cfg.port,
      url: `${base}/webhook/notifications`,
      health: `http://localhost:${cfg.port}/healthz`,
    });
  });

  let shuttingDown = false;
  const shutdown = (signal: string) => {
    if (shuttingDown) return;
    shuttingDown = true;
    log.info("shutting down", { signal });
    clearInterval(timer);
    server.close(() => {
      store.save();
      process.exit(0);
    });
    setTimeout(() => process.exit(0), 10_000).unref();
  };
  process.on("SIGINT", () => shutdown("SIGINT"));
  process.on("SIGTERM", () => shutdown("SIGTERM"));
}

const invokedDirectly = process.argv[1] !== undefined && import.meta.url === pathToFileURL(process.argv[1]).href;
if (invokedDirectly) {
  start().catch((err: unknown) => {
    process.stderr.write(`${errorMessage(err)}\n`);
    process.exitCode = 1;
  });
}
