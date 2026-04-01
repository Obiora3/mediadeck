const ROLE_LABELS = { admin: "Admin", planner: "Planner", buyer: "Buyer", finance: "Finance", viewer: "Viewer" };
const normalizeRole = (role = "viewer") => String(role || "viewer").trim().toLowerCase();
const formatRoleLabel = (role) => ROLE_LABELS[normalizeRole(role)] || "Viewer";
const makeUserRole = (user) => user ? ({ ...user, role: normalizeRole(user.role || "viewer"), roleLabel: formatRoleLabel(user.role || "viewer") }) : null;

export const store = {
  get: (k, d = null) => { try { const v = localStorage.getItem(k); return v ? JSON.parse(v) : d; } catch { return d; } },
  set: (k, v) => { try { localStorage.setItem(k, JSON.stringify(v)); } catch {} },
  del: (k) => { try { localStorage.removeItem(k); } catch {} },
};

export const uid = () => Math.random().toString(36).slice(2) + Date.now().toString(36);
export const themeKeyForUser = (userId) => userId ? `msp_theme_${userId}` : "msp_theme";
const signatureKeyForUser = (userId) => userId ? `msp_signature_${userId}` : "msp_signature";
const agencyContactKey = (agencyId) => agencyId ? `msp_agency_contact_${agencyId}` : "msp_agency_contact";

export const getStoredUserSignature = (userId) => store.get(signatureKeyForUser(userId), "") || "";
export const setStoredUserSignature = (userId, value = "") => {
  if (!userId) return;
  if (value) store.set(signatureKeyForUser(userId), value);
  else store.del(signatureKeyForUser(userId));
};
export const getStoredAgencyContact = (agencyId) => store.get(agencyContactKey(agencyId), { email: "", phone: "" }) || { email: "", phone: "" };
export const setStoredAgencyContact = (agencyId, contact = {}) => {
  if (!agencyId) return;
  store.set(agencyContactKey(agencyId), { email: contact?.email || "", phone: contact?.phone || "" });
};

export const getDefaultTheme = (userId = null) => {
  const saved = store.get(themeKeyForUser(userId), null);
  if (saved === "light" || saved === "dark") return saved;
  return window?.matchMedia?.("(prefers-color-scheme: dark)")?.matches ? "dark" : "light";
};

export const saveSession = (user) => {
  if (!user) {
    store.del("msp_session");
    return null;
  }
  const now = Date.now();
  const safeUser = makeUserRole(user);
  const session = {
    ...safeUser,
    sessionStartedAt: safeUser.sessionStartedAt || now,
    lastActiveAt: now,
  };
  store.set("msp_session", session);
  return session;
};

export const touchSession = () => {
  const current = store.get("msp_session");
  if (!current) return null;
  return saveSession(current);
};

export const sessionExpired = () => false;
