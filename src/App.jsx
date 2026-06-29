import { useState, useEffect, useCallback, useRef, useMemo } from "react";
import { usePwaInstall } from "./hooks/usePwaInstall";
import { supabase } from "./lib/supabase";
import Badge from "./components/Badge";
import Empty from "./components/Empty";
import Toast from "./components/Toast";
import Modal from "./components/Modal";
import Confirm from "./components/Confirm";
import PwaInstallPrompt from "./components/PwaInstallPrompt";
import { GlobalStyle, Btn, Field, AttachmentField, Card, Stat } from "./components/ui/primitives";
import AuthPage from "./components/auth/AuthPage";
import Sidebar from "./components/layout/Sidebar";
import TopRightNotificationsButton from "./components/layout/TopRightNotificationsButton";
import Dashboard from "./pages/Dashboard";
import VendorsPage from "./pages/VendorsPage";
import ClientsPage from "./pages/ClientsPage";
import CampaignsPage from "./pages/CampaignsPage";
import RatesPage from "./pages/RatesPage";
import ReportsPage from "./pages/ReportsPage";
import SettingsPage from "./pages/SettingsPage";
import FinancePage from "./pages/FinancePage";
import MPOPage from "./pages/MPOPage";
import PrintPreview from "./components/mpo/PrintPreview";
import { buildCSV } from "./utils/export";
import { themeKeyForUser, setStoredUserSignature, getDefaultTheme } from "./utils/session";
import { recordAuthDiagnostic, summarizeAuthSession, flushAuthDiagnosticsToAudit } from "./utils/authDiagnostics";
import { hydrateWorkspaceFromCache, setCachedWorkspaceSlice } from "./utils/workspaceCache";
import { activeOnly, archivedOnly, isArchived, archiveRecord, restoreRecord, pctWithin } from "./utils/records";
import { fmt, fmtN } from "./utils/formatters";
import {
  ROLE_LABELS,
  normalizeRole,
  formatRoleLabel,
  PERMISSIONS,
  hasPermission,
  readOnlyMessage,
} from "./constants/roles";
import {
  MPO_STATUS_OPTIONS,
  MPO_STATUS_LABELS,
  MPO_EXECUTION_STATUS_OPTIONS,
  MPO_INVOICE_STATUS_OPTIONS,
  MPO_PROOF_STATUS_OPTIONS,
  MPO_PAYMENT_STATUS_OPTIONS,
  MPO_RECON_STATUS_OPTIONS,
  toIsoInput,
  toIsoOrNull,
  fmtDateTime,
  getExecutionHealthColor,
  getExecutionHealthLabel,
  getAllowedMpoStatusTargets,
  mpoStatusNeedsNote,
  getMpoWorkflowMeta,
  isMpoAwaitingUser,
  getWorkflowActionLabel,
  getWorkflowActionVariant,
  getQuickWorkflowActions,
  canEditMpoContent,
  MPO_STATUS_COLORS,
} from "./constants/mpoWorkflow";
import { DEFAULT_SESSION_HOURS, DEFAULT_APP_SETTINGS, mergeAppSettings } from "./constants/appDefaults";
import { loadAppUserFromSupabase, buildFallbackAppUser, persistSignatureForUser, updateProfileInSupabase } from "./services/users";
import {
  normalizeAgencyCode,
  findExistingAgencyByName,
  ensureAgencyForUser,
  updateAgencyInSupabase,
  fetchAgencyMembersFromSupabase,
  updateAgencyMemberRoleInSupabase,
} from "./services/agencies";
import { changePasswordInSupabase } from "./services/auth";
import {
  uploadMpoAttachmentAndGetUrl,
  fetchMposFromSupabase,
  createMpoInSupabase,
  updateMpoInSupabase,
  archiveMpoInSupabase,
  restoreMpoInSupabase,
  updateMpoStatusInSupabase,
  updateMpoExecutionInSupabase,
  generateNextMpoNoFromSupabase,
} from "./services/mpos";
import { fetchVendorsFromSupabase } from "./services/vendors";
import { fetchClientsFromSupabase } from "./services/clients";
import { fetchCampaignsFromSupabase } from "./services/campaigns";
import { fetchRatesFromSupabase } from "./services/rates";
import {
  createAuditEventInSupabase,
  fetchAuditEventsForRecord,
  fetchAuditEventsForAgency,
  fetchNotificationsFromSupabase,
  markNotificationReadInSupabase,
  markAllNotificationsReadInSupabase,
  createNotificationForUserInSupabase,
  notifyMpoWorkflowTransition,
  notifyExecutionUpdate,
} from "./services/notifications";
import { fetchAppSettingsFromSupabase, saveAppSettingsToSupabase } from "./services/settings";
import {
  getReceivablesSyncMeta,
  getDaysPastDue,
  normalizeReceivableRecord,
  normalizePaymentEntry,
  getStoredReceivables,
  setStoredReceivables,
  fetchReceivablesFromSupabase,
  insertReceivableInSupabase,
  updateReceivableInSupabase,
  deleteReceivableInSupabase,
  insertReceivablePaymentInSupabase,
  updateReceivableStatusInSupabase,
  buildReceivableFromMpo,
} from "./services/receivables";

/* ── LOCAL HELPERS LEFT IN APP FOR THIS PHASE ───────────────────────────── */
const store = {
  get: (k, d = null) => { try { const v = localStorage.getItem(k); return v ? JSON.parse(v) : d; } catch { return d; } },
  set: (k, v) => { try { localStorage.setItem(k, JSON.stringify(v)); } catch {} },
  del: (k) => { try { localStorage.removeItem(k); } catch {} },
};
const uid = () => Math.random().toString(36).slice(2) + Date.now().toString(36);
const looksLikeUuid = (value = "") =>
  /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(
    String(value || "").trim()
  );
const APP_VERSION = "2.3";
const getAppSettings = () => mergeAppSettings(store.get("msp_app_settings", {}));
const roundMoneyValue = (value, settings = getAppSettings()) => {
  const num = Number(value) || 0;
  return settings?.roundToWholeNaira ? Math.round(num) : Math.round(num * 100) / 100;
};

const formatAuditTimestamp = (value) => {
  if (!value) return "—";
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? String(value) : date.toLocaleString();
};
const makeUserRole = (user) => user ? ({ ...user, role: normalizeRole(user.role || "viewer"), roleLabel: formatRoleLabel(user.role || "viewer") }) : null;
const downloadJSON = (filename, payload) => {
  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
};
const fileToDataUrl = (file) => new Promise((resolve, reject) => {
  try {
    if (!file) return resolve("");
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || ""));
    reader.onerror = () => reject(new Error("Failed to read selected signature file."));
    reader.readAsDataURL(file);
  } catch (error) {
    reject(error);
  }
});
const hashPassword = async (password) => {
  const source = new TextEncoder().encode(String(password || ""));
  const digest = await window.crypto.subtle.digest("SHA-256", source);
  return Array.from(new Uint8Array(digest)).map(b => b.toString(16).padStart(2, "0")).join("");
};
const verifyPassword = async (user, candidate) => {
  if (!user) return false;
  if (user.passwordHash) return (await hashPassword(candidate)) === user.passwordHash;
  return user.password === candidate;
};
const legacyUserNeedsHashUpgrade = (user) => Boolean(user?.password && !user?.passwordHash);
const downloadTextFile = (filename, content, mimeType = "text/plain;charset=utf-8") => {
  const blob = new Blob([content], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
};
const downloadRateTemplate = () => {
  const csv = [
    ["Vendor","Media","Programme","Timebelt","Duration","Rate","Discount","Commission","Notes"],
    ["Example FM","Radio","Morning Drive","06:00-09:00","30","150000","10","5","Prime time rate"],
    ["Example TV","Television","News at 9","21:00-21:30","45","450000","0","7.5","Headline bulletin"],
  ].map(row => row.map(v => `"${String(v).replace(/"/g, '""')}"`).join(",")).join("\n");

  downloadTextFile("mediadesk-rate-template.csv", csv, "text/csv;charset=utf-8");
};

/* ── MAIN APP ───────────────────────────────────────────── */
export default function App() {

  const [user, setUser] = useState(null);
  const [authUser, setAuthUser] = useState(null);
  const [page, setPage] = useState(() => {
    const p = new URLSearchParams(window.location.search).get("page");
    const valid = ["dashboard","vendors","clients","campaigns","rates","finance","mpo","reports","settings"];
    return valid.includes(p) ? p : "dashboard";
  });
  const [collapsed, setCollapsed] = useState(false);
  const [theme, setTheme] = useState(() => getDefaultTheme());
  const [appSettings, _setAppSettings] = useState(() => getAppSettings());
  useEffect(() => {
    setTheme(getDefaultTheme(user?.id || null));
  }, [user?.id]);

  useEffect(() => {
    const root = document.getElementById("root");
    const html = document.documentElement;
    const body = document.body;

    const previous = {
      htmlWidth: html.style.width,
      htmlMargin: html.style.margin,
      htmlPadding: html.style.padding,
      bodyWidth: body.style.width,
      bodyMargin: body.style.margin,
      bodyPadding: body.style.padding,
      bodyOverflowX: body.style.overflowX,
      rootWidth: root?.style.width || "",
      rootMaxWidth: root?.style.maxWidth || "",
      rootMargin: root?.style.margin || "",
      rootPadding: root?.style.padding || "",
    };

    html.style.width = "100%";
    html.style.margin = "0";
    html.style.padding = "0";

    body.style.width = "100%";
    body.style.margin = "0";
    body.style.padding = "0";
    body.style.overflowX = "hidden";

    if (root) {
      root.style.width = "100%";
      root.style.maxWidth = "100%";
      root.style.margin = "0";
      root.style.padding = "0";
    }

    return () => {
      html.style.width = previous.htmlWidth;
      html.style.margin = previous.htmlMargin;
      html.style.padding = previous.htmlPadding;

      body.style.width = previous.bodyWidth;
      body.style.margin = previous.bodyMargin;
      body.style.padding = previous.bodyPadding;
      body.style.overflowX = previous.bodyOverflowX;

      if (root) {
        root.style.width = previous.rootWidth;
        root.style.maxWidth = previous.rootMaxWidth;
        root.style.margin = previous.rootMargin;
        root.style.padding = previous.rootPadding;
      }
    };
  }, []);

  const setAppSettings = useCallback((value) => {
    _setAppSettings(prev => {
      const next = typeof value === "function" ? value(prev) : value;
      store.set("msp_app_settings", next);
      if (user?.agencyId) {
        setCachedWorkspaceSlice(user.agencyId, "settings", next);
        saveAppSettingsToSupabase(user.agencyId, next).catch(error => console.error("Failed to persist app settings:", error));
      }
      return next;
    });
  }, [user?.agencyId]);
  const toggleTheme = () => setTheme(t => {
    const n = t === "light" ? "dark" : "light";
    store.set(themeKeyForUser(user?.id || null), n);
    return n;
  });

  const [vendors, setVendors] = useState([]);
  const [clients, setClients] = useState([]);
  const [campaigns, setCampaigns] = useState([]);
  const [rates, setRates] = useState([]);
  const [mpos, setMpos] = useState([]);
  const [receivables, setReceivables] = useState([]);
  const [receivablesSync, setReceivablesSync] = useState(() => getReceivablesSyncMeta("local"));
  const [members, setMembers] = useState([]);
  const [notifications, setNotifications] = useState([]);
  const [alertsOpen, setAlertsOpen] = useState(false);
  const [settingsOpenSection, setSettingsOpenSection] = useState(null);

// MPOs now come from Supabase

  const [authReady, setAuthReady] = useState(false);
  const [sessionExpired, setSessionExpired] = useState(false);
  const { canInstall, showBanner, isInstalling, install, dismissBanner } = usePwaInstall();
  const mpoBulkImportRef = useRef(false);
  const refreshMposRef = useRef(async () => {});
  const mpoRefreshTimerRef = useRef(null);
  const mpoRefreshInFlightRef = useRef(false);
  const mpoRefreshQueuedRef = useRef(false);
  const authDiagnosticsFlushRef = useRef(false);
  const currentUserRef = useRef(null);

  useEffect(() => {
    currentUserRef.current = user;
  }, [user]);

  const flushAuthDiagnosticsForCurrentUser = useCallback(async () => {
    const currentUser = currentUserRef.current;
    if (!currentUser?.agencyId || !currentUser?.id || authDiagnosticsFlushRef.current) return 0;
    authDiagnosticsFlushRef.current = true;
    try {
      return await flushAuthDiagnosticsToAudit({
        agencyId: currentUser.agencyId,
        actor: currentUser,
        createAuditEvent: createAuditEventInSupabase,
      });
    } catch (error) {
      console.error("Failed to flush auth diagnostics:", error);
      return 0;
    } finally {
      authDiagnosticsFlushRef.current = false;
    }
  }, []);

  const resetWorkspaceState = useCallback(() => {
    setPage("dashboard");
    setUser(null);
    setAuthUser(null);
    setVendors([]);
    setClients([]);
    setCampaigns([]);
    setRates([]);
    setMpos([]);
    setReceivables([]);
    setReceivablesSync(getReceivablesSyncMeta("local"));
    setMembers([]);
    setNotifications([]);
    setAlertsOpen(false);
    setSettingsOpenSection(null);
    _setAppSettings(getAppSettings());
  }, []);

  useEffect(() => {
    let mounted = true;
    const timeout = setTimeout(() => {
      if (mounted) setAuthReady(true);
    }, 6000);

    const bootstrapAuth = async () => {
      recordAuthDiagnostic("bootstrap_start");
      try {
        const { data } = await supabase.auth.getSession();
        if (!mounted) return;
        recordAuthDiagnostic("bootstrap_session", summarizeAuthSession(data?.session));
        setAuthUser(data?.session?.user || null);
      } catch (error) {
        console.error("Failed to bootstrap auth:", error);
        recordAuthDiagnostic("bootstrap_error", { error });
      } finally {
        if (mounted) setAuthReady(true);
        clearTimeout(timeout);
      }
    };

    bootstrapAuth();

    const {
      data: { subscription },
    } = supabase.auth.onAuthStateChange((event, session) => {
      if (!mounted) return;
      recordAuthDiagnostic("state_change", {
        authEvent: event,
        ...summarizeAuthSession(session),
      });
      flushAuthDiagnosticsForCurrentUser();
      if (event === "SIGNED_OUT") {
        setAuthUser(null);
      } else if (event === "TOKEN_REFRESHED" || event === "USER_UPDATED" || event === "SIGNED_IN" || event === "PASSWORD_RECOVERY") {
        if (session?.user) setAuthUser(session.user);
      } else if (event === "INITIAL_SESSION") {
        setAuthUser(session?.user || null);
      }
      setAuthReady(true);
      clearTimeout(timeout);
    });

    return () => {
      mounted = false;
      clearTimeout(timeout);
      subscription.unsubscribe();
    };
  }, [flushAuthDiagnosticsForCurrentUser]);

  useEffect(() => {
    let active = true;

    const hydrateUser = async () => {
      if (!authUser?.id) {
        if (active) {
          if (user?.id) {
            recordAuthDiagnostic("auth_user_missing_reset", {
              previousUserId: user.id,
              previousEmail: user.email || "",
              previousAgencyId: user.agencyId || "",
            });
          }
          setSessionExpired(prev => {
            // Only show "session expired" if user was previously logged in
            return user !== null ? true : prev;
          });
          resetWorkspaceState();
          setAuthReady(true);
        }
        return;
      }

      try {
        let agencyId = authUser.user_metadata?.agency_id || null;
        try {
          agencyId = await ensureAgencyForUser(authUser);
        } catch (agencyError) {
          console.error("Failed to ensure agency for user:", agencyError);
          recordAuthDiagnostic("ensure_agency_error", {
            userId: authUser.id,
            email: authUser.email || "",
            error: agencyError,
          });
        }
        const appUser = await loadAppUserFromSupabase(authUser);

        if (!active) return;

        setSessionExpired(false);
        recordAuthDiagnostic("hydrate_user_success", {
          userId: authUser.id,
          email: authUser.email || "",
          agencyId: appUser?.agencyId || agencyId || "",
          usedFallbackUser: !appUser,
        });
        setUser(
          appUser
            ? { ...appUser, agencyId: appUser.agencyId || agencyId || null }
            : buildFallbackAppUser(authUser, agencyId)
        );
      } catch (error) {
        console.error("Failed to load user:", error);
        if (!active) return;
        setSessionExpired(false);
        recordAuthDiagnostic("hydrate_user_error", {
          userId: authUser.id,
          email: authUser.email || "",
          error,
        });
        setUser(buildFallbackAppUser(authUser));
      }
    };

    hydrateUser();

    return () => {
      active = false;
    };
  }, [authUser?.id, resetWorkspaceState]);

  useEffect(() => {
    if (!user?.agencyId || !user?.id) return;
    flushAuthDiagnosticsForCurrentUser();
  }, [user?.agencyId, user?.id, flushAuthDiagnosticsForCurrentUser]);

  useEffect(() => {
    if (!user?.agencyId) return;
    const cached = hydrateWorkspaceFromCache(user.agencyId, user.id || "");
    let hydratedCount = 0;

    if (cached.settings) {
      _setAppSettings(cached.settings);
      store.set("msp_app_settings", cached.settings);
      hydratedCount += 1;
    }
    if (Array.isArray(cached.members)) {
      setMembers(cached.members);
      hydratedCount += 1;
    }
    if (Array.isArray(cached.notifications)) {
      setNotifications(cached.notifications);
      hydratedCount += 1;
    }
    if (Array.isArray(cached.vendors)) {
      setVendors(cached.vendors);
      hydratedCount += 1;
    }
    if (Array.isArray(cached.clients)) {
      setClients(cached.clients);
      hydratedCount += 1;
    }
    if (Array.isArray(cached.campaigns)) {
      setCampaigns(cached.campaigns);
      hydratedCount += 1;
    }
    if (Array.isArray(cached.rates)) {
      setRates(cached.rates);
      hydratedCount += 1;
    }
    if (Array.isArray(cached.mpos)) {
      setMpos(cached.mpos);
      hydratedCount += 1;
    }

    const cachedReceivables = Array.isArray(cached.receivables) ? cached.receivables : getStoredReceivables(user.agencyId);
    if (Array.isArray(cachedReceivables) && cachedReceivables.length) {
      setReceivables(cachedReceivables);
      setReceivablesSync(getReceivablesSyncMeta("supabase"));
      hydratedCount += 1;
    }

    if (hydratedCount) {
      recordAuthDiagnostic("workspace_cache_hydrated", {
        agencyId: user.agencyId,
        userId: user.id || "",
        sliceCount: hydratedCount,
      });
    }
  }, [user?.agencyId, user?.id]);

  useEffect(() => {
    if (!user?.agencyId) {
      _setAppSettings(getAppSettings());
      return;
    }

    let active = true;
    const loadWorkspaceSettings = async () => {
      try {
        const settings = await fetchAppSettingsFromSupabase(user.agencyId);
        if (!active) return;
        _setAppSettings(settings);
        store.set("msp_app_settings", settings);
        setCachedWorkspaceSlice(user.agencyId, "settings", settings);
      } catch (e) {
        console.error("Failed to load app settings:", e);
      }
    };

    loadWorkspaceSettings();
    return () => { active = false; };
  }, [user?.agencyId]);

  useEffect(() => {
    if (!user?.agencyId) {
      setMembers([]);
      return;
    }

    let active = true;
    const loadMembers = async () => {
      try {
        const rows = await fetchAgencyMembersFromSupabase(user.agencyId);
        if (active) {
          setMembers(rows);
          setCachedWorkspaceSlice(user.agencyId, "members", rows);
        }
      } catch (e) {
        console.error("Failed to load members:", e);
      }
    };

    loadMembers();
    return () => { active = false; };
  }, [user?.agencyId]);

  useEffect(() => {
    if (!user?.agencyId || !user?.id) {
      setNotifications([]);
      return;
    }

    let active = true;
    const loadNotifications = async () => {
      try {
        const rows = await fetchNotificationsFromSupabase(user.id, user.agencyId);
        if (active) {
          setNotifications(rows);
          setCachedWorkspaceSlice(user.agencyId, `notifications:${user.id}`, rows);
        }
      } catch (e) {
        console.error("Failed to load notifications:", e);
      }
    };

    loadNotifications();
    return () => { active = false; };
  }, [user?.agencyId, user?.id]);

  useEffect(() => {
    if (!user?.agencyId) {
      setVendors([]);
      return;
    }

    let active = true;
    const loadVendors = async () => {
      try {
        const rows = await fetchVendorsFromSupabase(user.agencyId);
        if (active) {
          setVendors(rows);
          setCachedWorkspaceSlice(user.agencyId, "vendors", rows);
        }
      } catch (e) {
        console.error("Failed to load vendors:", e);
      }
    };

    loadVendors();
    return () => { active = false; };
  }, [user?.agencyId]);

  useEffect(() => {
    if (!user?.agencyId) {
      setClients([]);
      return;
    }

    let active = true;
    const loadClients = async () => {
      try {
        const rows = await fetchClientsFromSupabase(user.agencyId);
        if (active) {
          setClients(rows);
          setCachedWorkspaceSlice(user.agencyId, "clients", rows);
        }
      } catch (e) {
        console.error("Failed to load clients:", e);
      }
    };

    loadClients();
    return () => { active = false; };
  }, [user?.agencyId]);

  useEffect(() => {
    if (!user?.agencyId) {
      setCampaigns([]);
      return;
    }

    let active = true;
    const loadCampaigns = async () => {
      try {
        const rows = await fetchCampaignsFromSupabase(user.agencyId);
        if (active) {
          setCampaigns(rows);
          setCachedWorkspaceSlice(user.agencyId, "campaigns", rows);
        }
      } catch (e) {
        console.error("Failed to load campaigns:", e);
      }
    };

    loadCampaigns();
    return () => { active = false; };
  }, [user?.agencyId]);

  useEffect(() => {
    if (!user?.agencyId) {
      setRates([]);
      return;
    }

    let active = true;
    const loadRates = async () => {
      try {
        const rows = await fetchRatesFromSupabase(user.agencyId);
        if (active) {
          setRates(rows);
          setCachedWorkspaceSlice(user.agencyId, "rates", rows);
        }
      } catch (e) {
        console.error("Failed to load rates:", e);
      }
    };

    loadRates();
    return () => { active = false; };
  }, [user?.agencyId]);

  useEffect(() => {
    if (!user?.agencyId) {
      setMpos([]);
      return;
    }

    let active = true;
    const loadMpos = async () => {
      try {
        const rows = await fetchMposFromSupabase(user.agencyId);
        if (active) {
          setMpos(rows);
          setCachedWorkspaceSlice(user.agencyId, "mpos", rows);
        }
      } catch (e) {
        console.error("Failed to load MPOs:", e);
      }
    };

    loadMpos();
    return () => { active = false; };
  }, [user?.agencyId]);

  useEffect(() => {
    if (!user?.agencyId) {
      setReceivables([]);
      setReceivablesSync(getReceivablesSyncMeta("local"));
      return;
    }

    let active = true;
    const loadReceivables = async () => {
      try {
        const rows = await fetchReceivablesFromSupabase(user.agencyId);
        if (!active) return;
        setReceivables(rows);
        setReceivablesSync(getReceivablesSyncMeta("supabase"));
        setCachedWorkspaceSlice(user.agencyId, "receivables", rows);
      } catch (error) {
        console.error("Failed to load receivables from Supabase:", error);
        if (!active) return;
        setReceivables(getStoredReceivables(user.agencyId));
        setReceivablesSync(getReceivablesSyncMeta("local", error));
      }
    };

    loadReceivables();
    return () => { active = false; };
  }, [user?.agencyId]);

  useEffect(() => {
    if (!user?.agencyId) return;
    setStoredReceivables(user.agencyId, receivables);
    setCachedWorkspaceSlice(user.agencyId, "receivables", receivables);
  }, [user?.agencyId, receivables]);

  useEffect(() => {
    if (!user?.agencyId) return;

    const refreshUser = async () => {
      if (!authUser?.id) return;
      try {
        const agencyId = await ensureAgencyForUser(authUser);
        const refreshed = await loadAppUserFromSupabase(authUser);
        if (!refreshed) return;
        setUser(prev => {
          if (!prev) return prev;
          return { ...prev, ...refreshed, agencyId: refreshed.agencyId || agencyId || prev.agencyId || null };
        });
      } catch (error) {
        console.error("Failed to refresh user:", error);
      }
    };
    const refreshVendors = async () => {
      try {
        const rows = await fetchVendorsFromSupabase(user.agencyId);
        setVendors(rows);
        setCachedWorkspaceSlice(user.agencyId, "vendors", rows);
      } catch (error) { console.error("Realtime vendors refresh failed:", error); }
    };
    const refreshClients = async () => {
      try {
        const rows = await fetchClientsFromSupabase(user.agencyId);
        setClients(rows);
        setCachedWorkspaceSlice(user.agencyId, "clients", rows);
      } catch (error) { console.error("Realtime clients refresh failed:", error); }
    };
    const refreshCampaigns = async () => {
      try {
        const rows = await fetchCampaignsFromSupabase(user.agencyId);
        setCampaigns(rows);
        setCachedWorkspaceSlice(user.agencyId, "campaigns", rows);
      } catch (error) { console.error("Realtime campaigns refresh failed:", error); }
    };
    const refreshRates = async () => {
      try {
        const rows = await fetchRatesFromSupabase(user.agencyId);
        setRates(rows);
        setCachedWorkspaceSlice(user.agencyId, "rates", rows);
      } catch (error) { console.error("Realtime rates refresh failed:", error); }
    };
    const refreshMpos = async (options = {}) => {
      const { immediate = false } = options || {};
      const runRefresh = async () => {
        if (mpoRefreshInFlightRef.current) {
          mpoRefreshQueuedRef.current = true;
          return;
        }
        mpoRefreshInFlightRef.current = true;
        try {
          const rows = await fetchMposFromSupabase(user.agencyId);
          setMpos(rows);
          setCachedWorkspaceSlice(user.agencyId, "mpos", rows);
        } catch (error) {
          console.error("Realtime MPO refresh failed:", error);
        } finally {
          mpoRefreshInFlightRef.current = false;
          if (mpoRefreshQueuedRef.current) {
            mpoRefreshQueuedRef.current = false;
            await refreshMpos({ immediate: true });
          }
        }
      };

      if (mpoRefreshTimerRef.current) {
        clearTimeout(mpoRefreshTimerRef.current);
        mpoRefreshTimerRef.current = null;
      }

      if (immediate) {
        await runRefresh();
        return;
      }

      mpoRefreshTimerRef.current = setTimeout(() => {
        mpoRefreshTimerRef.current = null;
        runRefresh();
      }, 300);
    };
    refreshMposRef.current = refreshMpos;
    const refreshMembers = async () => {
      try {
        const rows = await fetchAgencyMembersFromSupabase(user.agencyId);
        setMembers(rows);
        setCachedWorkspaceSlice(user.agencyId, "members", rows);
      } catch (error) { console.error("Realtime members refresh failed:", error); }
    };
    const refreshReceivables = async () => {
      try {
        const rows = await fetchReceivablesFromSupabase(user.agencyId);
        setReceivables(rows);
        setReceivablesSync(getReceivablesSyncMeta("supabase"));
        setCachedWorkspaceSlice(user.agencyId, "receivables", rows);
      } catch (error) {
        console.error("Realtime receivables refresh failed:", error);
        setReceivablesSync(getReceivablesSyncMeta("local", error));
      }
    };
    const refreshNotifications = async () => {
      try {
        const rows = await fetchNotificationsFromSupabase(user.id, user.agencyId);
        setNotifications(rows);
        setCachedWorkspaceSlice(user.agencyId, `notifications:${user.id}`, rows);
      } catch (error) { console.error("Realtime notifications refresh failed:", error); }
    };
    const refreshSettings = async () => {
      try {
        const settings = await fetchAppSettingsFromSupabase(user.agencyId);
        _setAppSettings(settings);
        setCachedWorkspaceSlice(user.agencyId, "settings", settings);
      } catch (error) {
        console.error("Realtime settings refresh failed:", error);
      }
    };

    let channel = supabase
      .channel(`agency-live-${user.agencyId}`)
      .on("postgres_changes", { event: "*", schema: "public", table: "vendors", filter: `agency_id=eq.${user.agencyId}` }, refreshVendors)
      .on("postgres_changes", { event: "*", schema: "public", table: "clients", filter: `agency_id=eq.${user.agencyId}` }, refreshClients)
      .on("postgres_changes", { event: "*", schema: "public", table: "campaigns", filter: `agency_id=eq.${user.agencyId}` }, refreshCampaigns)
      .on("postgres_changes", { event: "*", schema: "public", table: "rates", filter: `agency_id=eq.${user.agencyId}` }, refreshRates)
      .on("postgres_changes", { event: "*", schema: "public", table: "mpos", filter: `agency_id=eq.${user.agencyId}` }, async () => { if (mpoBulkImportRef.current) return; await refreshMpos(); })
      .on("postgres_changes", { event: "*", schema: "public", table: "app_settings", filter: `agency_id=eq.${user.agencyId}` }, refreshSettings)
      .on("postgres_changes", { event: "*", schema: "public", table: "agencies", filter: `id=eq.${user.agencyId}` }, async () => { await refreshUser(); await refreshSettings(); })
      .on("postgres_changes", { event: "*", schema: "public", table: "profiles", filter: `agency_id=eq.${user.agencyId}` }, async () => { await refreshUser(); await refreshMembers(); })
      .on("postgres_changes", { event: "*", schema: "public", table: "notifications", filter: `recipient_user_id=eq.${user.id}` }, refreshNotifications);

    if (receivablesSync.mode === "supabase") {
      channel = channel
        .on("postgres_changes", { event: "*", schema: "public", table: "receivables", filter: `agency_id=eq.${user.agencyId}` }, refreshReceivables)
        .on("postgres_changes", { event: "*", schema: "public", table: "receivable_payments", filter: `agency_id=eq.${user.agencyId}` }, refreshReceivables);
    }

    channel.subscribe();

    return () => {
      refreshMposRef.current = async () => {};
      if (mpoRefreshTimerRef.current) {
        clearTimeout(mpoRefreshTimerRef.current);
        mpoRefreshTimerRef.current = null;
      }
      mpoRefreshQueuedRef.current = false;
      mpoRefreshInFlightRef.current = false;
      supabase.removeChannel(channel).catch(() => {});
    };
  }, [user?.agencyId, user?.id, authUser?.id, receivablesSync.mode]);

  const handleMpoBulkImportStateChange = (isActive) => {
    mpoBulkImportRef.current = !!isActive;
  };

  const requestMpoRefresh = async () => {
    await refreshMposRef.current?.({ immediate: true });
  };

  const unreadNotifications = useMemo(
    () => notifications.filter(notification => !notification.readAt).length,
    [notifications]
  );

  const openNotificationsSettings = () => {
    setSettingsOpenSection({ section: "notifications", key: Date.now() });
    setPage("settings");
  };

  const workspaceAlerts = (notifications || []).slice(0, 12).map(notification => ({
    id: notification.id,
    icon: notification.category === "finance" ? "💳" : notification.category === "reconciliation" ? "📑" : notification.category === "proof" ? "📎" : "🔔",
    title: notification.title || "Notification",
    message: notification.message || "Open settings to view this alert.",
    page: notification.linkPage || "settings",
    isUnread: !notification.readAt,
  }));

  const handleMarkNotificationRead = async (notificationId) => {
    try {
      const updated = await markNotificationReadInSupabase(notificationId);
      if (!updated) return;
      setNotifications(items => items.map(item => item.id === notificationId ? updated : item));
    } catch (error) {
      console.error("Failed to mark notification as read:", error);
    }
  };

  const handleMarkAllNotificationsRead = async () => {
    try {
      await markAllNotificationsReadInSupabase(user?.id, user?.agencyId);
      const timestamp = new Date().toISOString();
      setNotifications(items => items.map(item => item.readAt ? item : ({ ...item, readAt: timestamp })));
    } catch (error) {
      console.error("Failed to mark all notifications as read:", error);
    }
  };

  const switchReceivablesToLocalFallback = useCallback((error) => {
    setReceivablesSync(getReceivablesSyncMeta("local", error));
  }, []);

  const upsertLocalReceivable = useCallback((record) => {
    const normalized = normalizeReceivableRecord(record);
    setReceivables(items => [normalized, ...items.filter(item => item.id !== normalized.id)].map(normalizeReceivableRecord));
    return normalized;
  }, []);

  const removeLocalReceivable = useCallback((receivableId) => {
    setReceivables(items => items.filter(item => item.id !== receivableId));
    return true;
  }, []);

  const logLocalReceivablePayment = useCallback((receivableId, paymentInput) => {
    let updated = null;
    setReceivables(items => items.map(item => {
      if (item.id !== receivableId) return item;
      const normalized = normalizeReceivableRecord(item);
      const payment = normalizePaymentEntry(paymentInput);
      updated = normalizeReceivableRecord({
        ...normalized,
        status: normalized.status === "draft" ? "issued" : normalized.status,
        payments: [payment, ...(normalized.payments || [])],
        lastPaymentAt: payment.receivedAt,
        lastFollowUpAt: payment.receivedAt,
        updatedAt: new Date().toISOString(),
      });
      return updated;
    }).map(normalizeReceivableRecord));
    return updated;
  }, []);

  const updateLocalReceivableStatus = useCallback((receivableId, updates = {}) => {
    let updated = null;
    setReceivables(items => items.map(item => {
      if (item.id !== receivableId) return item;
      updated = normalizeReceivableRecord({
        ...item,
        ...updates,
        updatedAt: new Date().toISOString(),
      });
      return updated;
    }).map(normalizeReceivableRecord));
    return updated;
  }, []);

  const handleSaveReceivableRecord = useCallback(async (record) => {
    const normalized = normalizeReceivableRecord(record);
    const shouldUseSupabase = !!user?.agencyId && !!user?.id && receivablesSync.mode === "supabase";
    if (!shouldUseSupabase) return upsertLocalReceivable(normalized);

    const existsInState = receivables.some(item => item.id === normalized.id);
    const previousRecord = receivables.find(item => item.id === normalized.id) || null;
    const optimistic = normalizeReceivableRecord({ ...normalized, updatedAt: new Date().toISOString() });
    setReceivables(items => [optimistic, ...items.filter(item => item.id !== optimistic.id)].map(normalizeReceivableRecord));

    (async () => {
      try {
        const saved = existsInState && looksLikeUuid(optimistic.id)
          ? await updateReceivableInSupabase(optimistic.id, optimistic)
          : await insertReceivableInSupabase(user.agencyId, user.id, optimistic);
        setReceivables(items => [
          saved,
          ...items.filter(item => item.id !== optimistic.id && item.id !== saved.id),
        ].map(normalizeReceivableRecord));
      } catch (error) {
        console.error("Failed to save receivable in Supabase:", error);
        switchReceivablesToLocalFallback(error);
        setReceivables(items => {
          const withoutOptimistic = items.filter(item => item.id !== optimistic.id);
          return previousRecord
            ? [previousRecord, ...withoutOptimistic].map(normalizeReceivableRecord)
            : withoutOptimistic.map(normalizeReceivableRecord);
        });
      }
    })();

    return optimistic;
  }, [user?.agencyId, user?.id, receivablesSync.mode, receivables, switchReceivablesToLocalFallback, upsertLocalReceivable]);

  const handleRemoveReceivableRecord = useCallback(async (receivableId) => {
    const shouldUseSupabase = !!user?.agencyId && receivablesSync.mode === "supabase" && looksLikeUuid(receivableId);
    if (!shouldUseSupabase) return removeLocalReceivable(receivableId);

    const previousRecord = receivables.find(item => item.id === receivableId) || null;
    setReceivables(items => items.filter(item => item.id !== receivableId));

    deleteReceivableInSupabase(receivableId).catch(error => {
      console.error("Failed to remove receivable from Supabase:", error);
      switchReceivablesToLocalFallback(error);
      if (previousRecord) {
        setReceivables(items => [previousRecord, ...items.filter(item => item.id !== receivableId)].map(normalizeReceivableRecord));
      }
    });

    return true;
  }, [user?.agencyId, receivablesSync.mode, receivables, switchReceivablesToLocalFallback, removeLocalReceivable]);

  const handleLogReceivablePayment = useCallback(async (receivableId, paymentInput) => {
    const current = receivables.find(item => item.id === receivableId);
    if (!current) throw new Error("Receivable not found.");
    const shouldUseSupabase = !!user?.agencyId && !!user?.id && receivablesSync.mode === "supabase" && looksLikeUuid(receivableId);
    if (!shouldUseSupabase) return logLocalReceivablePayment(receivableId, paymentInput);

    const payment = normalizePaymentEntry(paymentInput);
    const optimistic = normalizeReceivableRecord({
      ...current,
      payments: [payment, ...(current.payments || [])],
      lastPaymentAt: payment.receivedAt,
      lastFollowUpAt: payment.receivedAt,
      updatedAt: new Date().toISOString(),
    });
    setReceivables(items => items.map(item => item.id === receivableId ? optimistic : item).map(normalizeReceivableRecord));

    insertReceivablePaymentInSupabase(user.agencyId, user.id, receivableId, paymentInput, current)
      .then(saved => {
        if (saved) setReceivables(items => items.map(item => item.id === saved.id ? saved : item).map(normalizeReceivableRecord));
      })
      .catch(error => {
        console.error("Failed to save receivable payment in Supabase:", error);
        switchReceivablesToLocalFallback(error);
        setReceivables(items => items.map(item => item.id === receivableId ? current : item).map(normalizeReceivableRecord));
      });

    return optimistic;
  }, [user?.agencyId, user?.id, receivablesSync.mode, receivables, logLocalReceivablePayment, switchReceivablesToLocalFallback]);

  const handleUpdateReceivableStatus = useCallback(async (receivableId, updates = {}) => {
    const shouldUseSupabase = !!user?.agencyId && receivablesSync.mode === "supabase" && looksLikeUuid(receivableId);
    if (!shouldUseSupabase) return updateLocalReceivableStatus(receivableId, updates);

    const current = receivables.find(item => item.id === receivableId);
    const optimistic = normalizeReceivableRecord({ ...(current || {}), ...updates, id: receivableId, updatedAt: new Date().toISOString() });
    setReceivables(items => items.map(item => item.id === receivableId ? optimistic : item).map(normalizeReceivableRecord));

    updateReceivableStatusInSupabase(receivableId, updates)
      .then(saved => {
        setReceivables(items => items.map(item => item.id === receivableId ? saved : item).map(normalizeReceivableRecord));
      })
      .catch(error => {
        console.error("Failed to update receivable status in Supabase:", error);
        switchReceivablesToLocalFallback(error);
        if (current) setReceivables(items => items.map(item => item.id === receivableId ? current : item).map(normalizeReceivableRecord));
      });

    return optimistic;
  }, [user?.agencyId, receivablesSync.mode, receivables, switchReceivablesToLocalFallback, updateLocalReceivableStatus]);

  const handleLogout = async () => {
    recordAuthDiagnostic("manual_sign_out_clicked", {
      userId: user?.id || "",
      email: user?.email || "",
      agencyId: user?.agencyId || "",
    });
    await flushAuthDiagnosticsForCurrentUser();
    setSessionExpired(false);
    resetWorkspaceState();

    try {
      await supabase.auth.signOut({ scope: "local" });
    } catch (e) {
      console.error("Failed to sign out:", e);
    }
  };

  const handleUserUpdate = (u) => setUser(prev => ({ ...prev, ...u }));

  useEffect(() => {
    if (page !== "settings" && settingsOpenSection) {
      setSettingsOpenSection(null);
    }
  }, [page, settingsOpenSection]);

  const pp = useMemo(() => ({
    vendors, clients, campaigns, rates, mpos, receivables,
    notifications, unreadNotifications,
    setVendors, setClients, setCampaigns, setRates, setMpos, setReceivables,
  }), [vendors, clients, campaigns, rates, mpos, receivables, notifications, unreadNotifications]);

if (!authReady) {
  return (
    <>
      <GlobalStyle theme={theme} />
      <div
        style={{
          minHeight: "100vh",
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
          background: "var(--bg)",
          color: "var(--text)",
          fontFamily: "var(--font-heading)",
          fontWeight: 700,
        }}
      >
        Loading...
      </div>
    </>
  );
}

if (!user) {
  return (
    <>
      <GlobalStyle theme={theme} />
      <AuthPage sessionExpired={sessionExpired} />
      <PwaInstallPrompt show={showBanner} onInstall={install} onDismiss={dismissBanner} isInstalling={isInstalling} />
    </>
  );
}

  return (
    <>
      <GlobalStyle theme={theme} />
      <div className="app-shell" style={{ display: "flex", minHeight: "100vh", width: "100%", maxWidth: "100%", overflowX: "hidden" }}>
        <Sidebar page={page} setPage={setPage} user={user} onLogout={handleLogout} collapsed={collapsed} setCollapsed={setCollapsed} theme={theme} toggleTheme={toggleTheme} unreadNotifications={unreadNotifications} canInstall={canInstall} onInstall={install} />
        <main className="app-main" style={{ flex: 1, minWidth: 0, width: "100%", overflowY: "auto", overflowX: "hidden", paddingTop: "calc(env(titlebar-area-height, 0px) + 28px)", paddingRight: 28, paddingBottom: 52, paddingLeft: 28, position: "relative", boxSizing: "border-box" }}>
          <TopRightNotificationsButton count={unreadNotifications} onClick={() => setAlertsOpen(true)} />
          {alertsOpen && (
            <Modal title="Workspace Alerts" onClose={() => setAlertsOpen(false)} width={560}>
              <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
                {workspaceAlerts.length === 0 ? (
                  <Empty icon="🔔" title="No alerts right now" sub="Your workspace looks clear." />
                ) : workspaceAlerts.map(alert => (
                  <Card key={alert.id} style={{ padding: 16 }}>
                    <div style={{ display: "flex", gap: 12, alignItems: "flex-start" }}>
                      <div style={{ width: 40, height: 40, borderRadius: 10, background: "var(--bg3)", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 18, flexShrink: 0 }}>{alert.icon}</div>
                      <div style={{ flex: 1 }}>
                        <div style={{ fontFamily: "var(--font-heading)", fontWeight: 700, fontSize: 14 }}>{alert.title}</div>
                        <div style={{ fontSize: 13, color: "var(--text2)", marginTop: 4 }}>{alert.message}</div>
                      </div>
                      {alert.isUnread ? <Badge color="accent">New</Badge> : null}
                    </div>
                  </Card>
                ))}
                <div style={{ display: "flex", justifyContent: "flex-end" }}>
                  <Btn variant="ghost" onClick={() => { setAlertsOpen(false); openNotificationsSettings(); }}>Open Inbox</Btn>
                </div>
              </div>
            </Modal>
          )}
          {page === "dashboard"  && <Dashboard user={user} {...pp} setPage={setPage} onOpenNotifications={openNotificationsSettings} />}
          {page === "vendors"    && <VendorsPage {...pp} user={user} />}
          {page === "clients"    && <ClientsPage {...pp} user={user} />}
          {page === "campaigns"  && <CampaignsPage {...pp} user={user} />}
          {page === "rates"      && <RatesPage {...pp} user={user} />}
          {page === "finance"    && <FinancePage user={user} vendors={vendors} clients={clients} campaigns={campaigns} mpos={mpos} receivables={receivables} receivablesMeta={receivablesSync} onSaveReceivable={handleSaveReceivableRecord} onRemoveReceivable={handleRemoveReceivableRecord} onLogReceivablePayment={handleLogReceivablePayment} onUpdateReceivableStatus={handleUpdateReceivableStatus} />}
          {page === "mpo"        && <MPOPage {...pp} user={user} appSettings={appSettings} onBulkImportStateChange={handleMpoBulkImportStateChange} requestMpoRefresh={requestMpoRefresh} />}
          {page === "reports"    && <ReportsPage {...pp} activeOnly={activeOnly} fmtN={fmtN} MPO_STATUS_LABELS={MPO_STATUS_LABELS} PrintPreview={PrintPreview} buildCSV={buildCSV} />}
          {page === "settings"   && <SettingsPage user={user} onUserUpdate={handleUserUpdate} onLogout={handleLogout} appSettings={appSettings} setAppSettings={setAppSettings} vendors={vendors} clients={clients} campaigns={campaigns} rates={rates} mpos={mpos} receivables={receivables} members={members} setMembers={setMembers} notifications={notifications} setNotifications={setNotifications} unreadNotifications={unreadNotifications} onMarkNotificationRead={handleMarkNotificationRead} onMarkAllNotificationsRead={handleMarkAllNotificationsRead} initialSectionRequest={settingsOpenSection} />}
        </main>
      </div>
      <PwaInstallPrompt show={showBanner} onInstall={install} onDismiss={dismissBanner} isInstalling={isInstalling} />
    </>
  );
}
