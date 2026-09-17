/** Minimal structured logger. MCP servers must never write to stdout (it is the protocol channel), so everything goes to stderr. */
export type Logger = {
  info: (msg: string, meta?: Record<string, unknown>) => void;
  warn: (msg: string, meta?: Record<string, unknown>) => void;
  error: (msg: string, meta?: Record<string, unknown>) => void;
  debug: (msg: string, meta?: Record<string, unknown>) => void;
};

export function createLogger(scope: string, level: "debug" | "info" | "warn" | "error" = (process.env.LOG_LEVEL as any) || "info"): Logger {
  const order = { debug: 0, info: 1, warn: 2, error: 3 } as const;
  const emit = (lvl: keyof typeof order, msg: string, meta?: Record<string, unknown>) => {
    if (order[lvl] < order[level]) return;
    const line = JSON.stringify({ ts: new Date().toISOString(), level: lvl, scope, msg, ...(meta ?? {}) });
    process.stderr.write(line + "\n");
  };
  return {
    debug: (m, meta) => emit("debug", m, meta),
    info: (m, meta) => emit("info", m, meta),
    warn: (m, meta) => emit("warn", m, meta),
    error: (m, meta) => emit("error", m, meta),
  };
}

export const silentLogger: Logger = { info() {}, warn() {}, error() {}, debug() {} };
