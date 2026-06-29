const CACHE_PREFIX = "msp_workspace_cache_v1";
const CACHE_MAX_AGE_MS = 1000 * 60 * 60 * 24 * 14;

const makeKey = (agencyId, slice) =>
  agencyId && slice ? `${CACHE_PREFIX}:${agencyId}:${slice}` : "";

export const getCachedWorkspaceSlice = (agencyId, slice, fallback = null) => {
  const key = makeKey(agencyId, slice);
  if (!key) return fallback;
  try {
    const raw = localStorage.getItem(key);
    if (!raw) return fallback;
    const cached = JSON.parse(raw);
    const cachedAt = cached?.cachedAt ? new Date(cached.cachedAt).getTime() : 0;
    if (!cachedAt || Date.now() - cachedAt > CACHE_MAX_AGE_MS) return fallback;
    return cached?.data ?? fallback;
  } catch {
    return fallback;
  }
};

export const setCachedWorkspaceSlice = (agencyId, slice, data) => {
  const key = makeKey(agencyId, slice);
  if (!key) return;
  try {
    localStorage.setItem(key, JSON.stringify({
      cachedAt: new Date().toISOString(),
      data,
    }));
  } catch {
    // Local storage can be full, especially with large MPO schedules.
  }
};

export const hydrateWorkspaceFromCache = (agencyId, userId = "") => ({
  settings: getCachedWorkspaceSlice(agencyId, "settings", null),
  members: getCachedWorkspaceSlice(agencyId, "members", null),
  notifications: userId ? getCachedWorkspaceSlice(agencyId, `notifications:${userId}`, null) : null,
  vendors: getCachedWorkspaceSlice(agencyId, "vendors", null),
  clients: getCachedWorkspaceSlice(agencyId, "clients", null),
  campaigns: getCachedWorkspaceSlice(agencyId, "campaigns", null),
  rates: getCachedWorkspaceSlice(agencyId, "rates", null),
  mpos: getCachedWorkspaceSlice(agencyId, "mpos", null),
  receivables: getCachedWorkspaceSlice(agencyId, "receivables", null),
});
