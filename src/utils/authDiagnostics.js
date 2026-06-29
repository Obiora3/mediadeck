const AUTH_DIAGNOSTICS_KEY = "msp_auth_diagnostics";
const AUTH_DIAGNOSTICS_MAX = 80;

const readEntries = () => {
  try {
    const raw = localStorage.getItem(AUTH_DIAGNOSTICS_KEY);
    const parsed = raw ? JSON.parse(raw) : [];
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
};

const writeEntries = (entries = []) => {
  try {
    localStorage.setItem(AUTH_DIAGNOSTICS_KEY, JSON.stringify(entries.slice(-AUTH_DIAGNOSTICS_MAX)));
  } catch {
    // Diagnostics should never interrupt normal auth flow.
  }
};

const summarizeError = (error) => {
  if (!error) return null;
  return {
    name: error?.name || "",
    message: error?.message || String(error),
    status: error?.status || error?.code || "",
  };
};

export const summarizeAuthSession = (session) => {
  if (!session) return { hasSession: false };
  const expiresAtMs = session.expires_at ? Number(session.expires_at) * 1000 : null;
  return {
    hasSession: true,
    userId: session.user?.id || "",
    email: session.user?.email || "",
    agencyId: session.user?.user_metadata?.agency_id || "",
    expiresAt: expiresAtMs ? new Date(expiresAtMs).toISOString() : "",
    expiresInSeconds: expiresAtMs ? Math.round((expiresAtMs - Date.now()) / 1000) : null,
  };
};

export const recordAuthDiagnostic = (event, details = {}) => {
  const entry = {
    id: `${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 8)}`,
    at: new Date().toISOString(),
    event: String(event || "auth_event"),
    origin: typeof window !== "undefined" ? window.location.origin : "",
    path: typeof window !== "undefined" ? `${window.location.pathname}${window.location.search}` : "",
    visibilityState: typeof document !== "undefined" ? document.visibilityState : "",
    online: typeof navigator !== "undefined" ? navigator.onLine : null,
    details: {
      ...details,
      error: summarizeError(details?.error),
    },
    flushedAt: null,
  };
  const entries = [...readEntries(), entry].slice(-AUTH_DIAGNOSTICS_MAX);
  writeEntries(entries);
  if (import.meta.env.DEV) {
    console.info("[auth diagnostic]", entry.event, entry.details);
  }
  return entry;
};

export const flushAuthDiagnosticsToAudit = async ({ agencyId, actor, createAuditEvent, limit = 25 } = {}) => {
  if (!agencyId || !actor?.id || typeof createAuditEvent !== "function") return 0;
  const entries = readEntries();
  const pending = entries.filter((entry) => !entry.flushedAt).slice(-limit);
  if (!pending.length) return 0;

  let flushed = 0;
  const flushedAt = new Date().toISOString();
  const flushedIds = new Set();

  for (const entry of pending) {
    try {
      await createAuditEvent({
        agencyId,
        recordType: "auth",
        recordId: actor.id,
        action: `auth_${String(entry.event || "event").replace(/[^a-z0-9_]+/gi, "_").toLowerCase()}`,
        actor,
        note: `Auth diagnostic: ${entry.event}`,
        metadata: {
          at: entry.at,
          origin: entry.origin,
          path: entry.path,
          visibilityState: entry.visibilityState,
          online: entry.online,
          ...entry.details,
        },
      });
      flushed += 1;
      flushedIds.add(entry.id);
    } catch (error) {
      console.error("Failed to flush auth diagnostic:", error);
      break;
    }
  }

  if (flushedIds.size) {
    writeEntries(entries.map((entry) => flushedIds.has(entry.id) ? { ...entry, flushedAt } : entry));
  }
  return flushed;
};
