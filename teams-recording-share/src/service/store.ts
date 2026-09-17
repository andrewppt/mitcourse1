/**
 * Tiny JSON-file state store for the webhook service.
 *
 * Holds the Graph subscriptions we own, the recording ids we have already processed
 * (so a redelivered notification is a no-op) and the last few errors for `/status`.
 *
 * Everything is kept in memory and flushed to `${cfg.stateDir}/state.json` with an
 * atomic temp-file + rename, so a crash mid-write cannot corrupt the file.
 */
import fs from "node:fs";
import path from "node:path";
import type { AppConfig } from "../config.js";

/** A subscription we created (mirrors the fields of a Graph subscription we care about). */
export interface SubscriptionRecord {
  id: string;
  resource: string;
  expirationDateTime: string;
  notificationUrl?: string;
  lifecycleNotificationUrl?: string;
  changeType?: string;
  /** ISO timestamp of the last create/renew we performed. */
  updatedAt?: string;
}

/** What happened when we shared one recording. */
export interface ProcessedSummary {
  meetingId?: string;
  driveItemId?: string;
  granted: number;
  skipped: number;
  errors: number;
  /** ISO timestamp. */
  at: string;
}

export interface ProcessedRecord extends ProcessedSummary {
  recordingId: string;
}

export interface ErrorRecord {
  message: string;
  at: string;
}

export interface StateShape {
  version: number;
  subscriptions: SubscriptionRecord[];
  processed: ProcessedRecord[];
  errors: ErrorRecord[];
}

const MAX_PROCESSED = 500;
const MAX_ERRORS = 50;

function emptyState(): StateShape {
  return { version: 1, subscriptions: [], processed: [], errors: [] };
}

export class StateStore {
  readonly file: string;
  private state: StateShape = emptyState();
  /** recordingId -> index into `state.processed`, rebuilt on load. */
  private processedIds = new Set<string>();

  constructor(private readonly stateDir: string) {
    this.file = path.resolve(stateDir, "state.json");
  }

  /** Convenience factory used by the service entry points. */
  static fromConfig(cfg: AppConfig): StateStore {
    const store = new StateStore(cfg.stateDir);
    store.load();
    return store;
  }

  /** Read the state file if it exists. A missing or unreadable file starts from empty state. */
  load(): void {
    try {
      const raw = fs.readFileSync(this.file, "utf8");
      const parsed = JSON.parse(raw) as Partial<StateShape>;
      this.state = {
        version: typeof parsed.version === "number" ? parsed.version : 1,
        subscriptions: Array.isArray(parsed.subscriptions) ? parsed.subscriptions : [],
        processed: Array.isArray(parsed.processed) ? parsed.processed : [],
        errors: Array.isArray(parsed.errors) ? parsed.errors : [],
      };
    } catch {
      this.state = emptyState();
    }
    this.processedIds = new Set(this.state.processed.map((p) => p.recordingId));
  }

  /** Atomically persist the state (mkdir -p, write temp file, rename). */
  save(): void {
    fs.mkdirSync(this.stateDir, { recursive: true });
    const tmp = `${this.file}.${process.pid}.${Date.now()}.tmp`;
    fs.writeFileSync(tmp, JSON.stringify(this.state, null, 2), { encoding: "utf8", mode: 0o600 });
    fs.renameSync(tmp, this.file);
  }

  getSubscriptions(): SubscriptionRecord[] {
    return this.state.subscriptions.map((s) => ({ ...s }));
  }

  upsertSubscription(rec: SubscriptionRecord): void {
    const next: SubscriptionRecord = { ...rec, updatedAt: rec.updatedAt ?? new Date().toISOString() };
    const i = this.state.subscriptions.findIndex((s) => s.id === rec.id);
    if (i >= 0) this.state.subscriptions[i] = { ...this.state.subscriptions[i], ...next };
    else this.state.subscriptions.push(next);
    this.save();
  }

  removeSubscription(id: string): void {
    const before = this.state.subscriptions.length;
    this.state.subscriptions = this.state.subscriptions.filter((s) => s.id !== id);
    if (this.state.subscriptions.length !== before) this.save();
  }

  isProcessed(recordingId: string): boolean {
    return this.processedIds.has(recordingId);
  }

  markProcessed(recordingId: string, summary: ProcessedSummary): void {
    const rec: ProcessedRecord = { recordingId, ...summary };
    const i = this.state.processed.findIndex((p) => p.recordingId === recordingId);
    if (i >= 0) this.state.processed.splice(i, 1);
    this.state.processed.push(rec);
    if (this.state.processed.length > MAX_PROCESSED) {
      const dropped = this.state.processed.splice(0, this.state.processed.length - MAX_PROCESSED);
      for (const d of dropped) this.processedIds.delete(d.recordingId);
    }
    this.processedIds.add(recordingId);
    this.save();
  }

  /** Most recent first. */
  listProcessed(limit = 50): ProcessedRecord[] {
    const n = Math.max(0, limit);
    return this.state.processed.slice(-n).reverse().map((p) => ({ ...p }));
  }

  processedCount(): number {
    return this.state.processed.length;
  }

  recordError(msg: string): void {
    this.state.errors.push({ message: msg, at: new Date().toISOString() });
    if (this.state.errors.length > MAX_ERRORS) this.state.errors.splice(0, this.state.errors.length - MAX_ERRORS);
    this.save();
  }

  /** Most recent first. */
  listErrors(limit = 20): ErrorRecord[] {
    const n = Math.max(0, limit);
    return this.state.errors.slice(-n).reverse().map((e) => ({ ...e }));
  }
}
