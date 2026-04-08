import { useState, useEffect, useRef } from "react";
import Badge from "../components/Badge";
import Empty from "../components/Empty";
import Toast from "../components/Toast";
import Confirm from "../components/Confirm";
import { Btn, Field, Card, Stat } from "../components/ui/primitives";
import { themeKeyForUser, setStoredUserSignature, getDefaultTheme } from "../utils/session";
import { fmtN } from "../utils/formatters";
import {
  ROLE_LABELS,
  normalizeRole,
  formatRoleLabel,
  hasPermission,
  readOnlyMessage,
} from "../constants/roles";
import { DEFAULT_SESSION_HOURS, DEFAULT_APP_SETTINGS } from "../constants/appDefaults";
import { persistSignatureForUser, updateProfileInSupabase } from "../services/users";
import { updateAgencyInSupabase, updateAgencyMemberRoleInSupabase } from "../services/agencies";
import { changePasswordInSupabase } from "../services/auth";
import { createAuditEventInSupabase, fetchAuditEventsForAgency } from "../services/notifications";

const APP_VERSION = "2.3";
const store = {
  get: (k, d = null) => { try { const v = localStorage.getItem(k); return v ? JSON.parse(v) : d; } catch { return d; } },
  set: (k, v) => { try { localStorage.setItem(k, JSON.stringify(v)); } catch {} },
  del: (k) => { try { localStorage.removeItem(k); } catch {} },
};


const formatAuditTimestamp = (value) => {
  if (!value) return "—";
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? String(value) : date.toLocaleString();
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

const SettingsPage = ({ user, onUserUpdate, onLogout, appSettings, setAppSettings, vendors, clients, campaigns, rates, mpos, receivables = [], members, setMembers, notifications, unreadNotifications, onMarkNotificationRead, onMarkAllNotificationsRead, initialSectionRequest = null }) => {
  const [toast, setToast] = useState(null);
  const backupInputRef = useRef(null);
  const [confirm, setConfirm] = useState(null);
  const [section, setSection] = useState("profile"); // profile | security | agency | team | activity | data
  const canManageWorkspace = hasPermission(user, "manageWorkspace");
  const canManageMembers = hasPermission(user, "manageMembers");
  const canDanger = hasPermission(user, "manageDangerZone");
  const [memberRoles, setMemberRoles] = useState({});
  const [savingMemberId, setSavingMemberId] = useState(null);
  const [activityItems, setActivityItems] = useState([]);
  const [activityLoading, setActivityLoading] = useState(false);
  const [activityFilter, setActivityFilter] = useState("all");
  const [notificationFilter, setNotificationFilter] = useState("all");

  // Profile form
  const [pf, setPf] = useState({ name: user.name || "", title: user.title || "", email: user.email || "", phone: user.phone || "", signatureDataUrl: user.signatureDataUrl || "" });
  const up = k => v => setPf(p => ({ ...p, [k]: v }));

  // Agency form
  const [af, setAf] = useState({ agency: user.agency || "", address: user.agencyAddress || "", email: user.agencyEmail || "", phone: user.agencyPhone || "" });
  const [docSettings, setDocSettings] = useState({
    vatRate: String(appSettings?.vatRate ?? 7.5),
    sessionHours: String(appSettings?.sessionHours ?? DEFAULT_SESSION_HOURS),
    roundToWholeNaira: !!appSettings?.roundToWholeNaira,
    mpoTerms: (appSettings?.mpoTerms || DEFAULT_APP_SETTINGS.mpoTerms).join("\n"),
  });
  const udp = k => v => setDocSettings(s => ({ ...s, [k]: v }));
  const ua = k => v => setAf(a => ({ ...a, [k]: v }));

  // Password form
  const [sf, setSf] = useState({ current: "", newPw: "", confirm: "" });
  const us = k => v => setSf(s => ({ ...s, [k]: v }));
  const signatureInputRef = useRef(null);

  const handleSignatureUpload = async (file) => {
    if (!file) return;
    try {
      const dataUrl = await fileToDataUrl(file);
      setPf(p => ({ ...p, signatureDataUrl: dataUrl }));
      onUserUpdate({ ...user, signatureDataUrl: dataUrl });
      await persistSignatureForUser(user, dataUrl);
      setToast({ msg: "Signature uploaded and saved.", type: "success" });
    } catch (error) {
      setToast({ msg: error.message || "Failed to upload signature.", type: "error" });
    }
  };

  const handleSignatureRemove = async () => {
    try {
      setPf(p => ({ ...p, signatureDataUrl: "" }));
      onUserUpdate({ ...user, signatureDataUrl: "" });
      await persistSignatureForUser(user, "");
      setToast({ msg: "Signature removed.", type: "success" });
    } catch (error) {
      setToast({ msg: error.message || "Failed to remove signature.", type: "error" });
    }
  };

  useEffect(() => {
    setPf({ name: user.name || "", title: user.title || "", email: user.email || "", phone: user.phone || "", signatureDataUrl: user.signatureDataUrl || "" });
    setAf({ agency: user.agency || "", address: user.agencyAddress || "", email: user.agencyEmail || "", phone: user.agencyPhone || "" });
  }, [user?.id, user?.name, user?.title, user?.email, user?.phone, user?.agency, user?.agencyAddress, user?.signatureDataUrl, user?.agencyEmail, user?.agencyPhone]);

  useEffect(() => {
    setDocSettings({
      vatRate: String(appSettings?.vatRate ?? 7.5),
      sessionHours: String(appSettings?.sessionHours ?? DEFAULT_SESSION_HOURS),
      roundToWholeNaira: !!appSettings?.roundToWholeNaira,
      mpoTerms: (appSettings?.mpoTerms || DEFAULT_APP_SETTINGS.mpoTerms).join("\n"),
    });
  }, [appSettings?.vatRate, appSettings?.sessionHours, appSettings?.roundToWholeNaira, JSON.stringify(appSettings?.mpoTerms || [])]);

  useEffect(() => {
    setMemberRoles(Object.fromEntries((members || []).map(member => [member.id, normalizeRole(member.role)])));
  }, [JSON.stringify((members || []).map(member => ({ id: member.id, role: member.role })))]);

  useEffect(() => {
    if (section !== "activity" || !user?.agencyId) return;
    let active = true;
    setActivityLoading(true);
    fetchAuditEventsForAgency(user.agencyId, activityFilter)
      .then(rows => { if (active) setActivityItems(rows); })
      .catch(error => { if (active) setToast({ msg: error.message || "Failed to load workspace activity.", type: "error" }); })
      .finally(() => { if (active) setActivityLoading(false); });
    return () => { active = false; };
  }, [section, user?.agencyId, activityFilter]);

  const saveMemberRole = async (memberId) => {
    if (!canManageMembers) return setToast({ msg: readOnlyMessage(user), type: "error" });
    const nextRole = normalizeRole(memberRoles[memberId]);
    const targetMember = (members || []).find(member => member.id === memberId);
    try {
      setSavingMemberId(memberId);
      const savedRole = await updateAgencyMemberRoleInSupabase(memberId, nextRole);
      setMembers(prev => prev.map(member => member.id === memberId ? { ...member, role: savedRole, roleLabel: formatRoleLabel(savedRole) } : member));
      if (memberId === user.id) onUserUpdate({ ...user, role: savedRole, roleLabel: formatRoleLabel(savedRole) });
      createAuditEventInSupabase({
        agencyId: user.agencyId,
        recordType: "member",
        recordId: memberId,
        action: "role_changed",
        actor: user,
        note: `${targetMember?.name || targetMember?.email || "Member"} is now ${formatRoleLabel(savedRole)}.`,
        metadata: { email: targetMember?.email || "", role: savedRole },
      }).catch(error => console.error("Failed to write audit event:", error));
      setToast({ msg: "Member role updated.", type: "success" });
    } catch (error) {
      setToast({ msg: error.message || "Failed to update member role.", type: "error" });
    } finally {
      setSavingMemberId(null);
    }
  };

  const saveProfile = async () => {
    if (!pf.name || !pf.email) return setToast({ msg: "Name and email are required.", type: "error" });
    try {
      const updated = await updateProfileInSupabase(user, pf);
      onUserUpdate(updated);
      createAuditEventInSupabase({
        agencyId: user.agencyId,
        recordType: "profile",
        recordId: user.id,
        action: "updated",
        actor: user,
        metadata: { email: updated.email || "", title: updated.title || "" },
      }).catch(error => console.error("Failed to write audit event:", error));
      setToast({ msg: pf.email !== user.email ? "Profile updated. Check your email if confirmation is required for the new address." : "Profile updated successfully!", type: "success" });
    } catch (error) {
      setToast({ msg: error.message || "Failed to update profile.", type: "error" });
    }
  };

  const saveAgency = async () => {
    if (!canManageWorkspace) return setToast({ msg: readOnlyMessage(user), type: "error" });
    if (!af.agency) return setToast({ msg: "Agency name is required.", type: "error" });
    try {
      const updated = await updateAgencyInSupabase(user.agencyId, af);
      onUserUpdate(updated);
      createAuditEventInSupabase({
        agencyId: user.agencyId,
        recordType: "agency",
        recordId: user.agencyId,
        action: "updated",
        actor: user,
        metadata: { name: updated.agency || af.agency || "", address: updated.agencyAddress || af.address || "", email: updated.agencyEmail || af.email || "", phone: updated.agencyPhone || af.phone || "" },
      }).catch(error => console.error("Failed to write audit event:", error));
      setToast({ msg: "Agency details updated!", type: "success" });
    } catch (error) {
      setToast({ msg: error.message || "Failed to update agency details.", type: "error" });
    }
  };

  const exportBackup = () => {
    const payload = {
      version: APP_VERSION,
      exportedAt: new Date().toISOString(),
      user: { ...user },
      data: { vendors, clients, campaigns, rates, mpos, receivables, notifications, appSettings, localPreferences: { theme: store.get(themeKeyForUser(user?.id || null), getDefaultTheme(user?.id || null)), signatureCached: !!user?.signatureDataUrl } },
    };
    downloadJSON(`mediadesk-backup-${new Date().toISOString().slice(0,10)}.json`, payload);
    setToast({ msg: "Agency backup exported successfully.", type: "success" });
  };

  const handleBackupImport = async (file) => {
    if (!file) return;
    setToast({ msg: "Snapshot import is disabled for cloud workspaces to protect your live Supabase data. Use CSV tools for structured imports instead.", type: "error" });
  };

  const saveDocumentSettings = async () => {
    if (!canManageWorkspace) return setToast({ msg: readOnlyMessage(user), type: "error" });
    const vatRate = parseFloat(docSettings.vatRate);
    const sessionHours = parseFloat(docSettings.sessionHours);
    if (!Number.isFinite(vatRate) || vatRate < 0 || vatRate > 100) return setToast({ msg: "Enter a valid VAT rate between 0 and 100.", type: "error" });
    if (!Number.isFinite(sessionHours) || sessionHours < 1 || sessionHours > 48) return setToast({ msg: "Session hours must be between 1 and 48.", type: "error" });
    try {
      const nextSettings = {
        ...appSettings,
        vatRate,
        sessionHours,
        roundToWholeNaira: !!docSettings.roundToWholeNaira,
        mpoTerms: docSettings.mpoTerms.split("\n").map(t => t.trim()).filter(Boolean),
      };
      setAppSettings(nextSettings);
      createAuditEventInSupabase({
        agencyId: user.agencyId,
        recordType: "workspace",
        recordId: user.agencyId,
        action: "settings_updated",
        actor: user,
        metadata: { vatRate: nextSettings.vatRate, sessionHours: nextSettings.sessionHours, roundToWholeNaira: nextSettings.roundToWholeNaira },
      }).catch(error => console.error("Failed to write audit event:", error));
      setToast({ msg: "Document and billing settings updated.", type: "success" });
    } catch (error) {
      setToast({ msg: error.message || "Failed to update workspace settings.", type: "error" });
    }
  };

  const changePassword = async () => {
    if (!sf.current || !sf.newPw || !sf.confirm) return setToast({ msg: "All password fields required.", type: "error" });
    if (sf.newPw.length < 6) return setToast({ msg: "New password must be at least 6 characters.", type: "error" });
    if (sf.newPw !== sf.confirm) return setToast({ msg: "New passwords do not match.", type: "error" });
    try {
      await changePasswordInSupabase(sf.newPw);
      setSf({ current: "", newPw: "", confirm: "" });
      createAuditEventInSupabase({
        agencyId: user.agencyId,
        recordType: "security",
        recordId: user.id,
        action: "password_changed",
        actor: user,
      }).catch(error => console.error("Failed to write audit event:", error));
      setToast({ msg: "Password changed successfully!", type: "success" });
    } catch (error) {
      setToast({ msg: error.message || "Failed to change password.", type: "error" });
    }
  };

  const tabs = [
    { id: "profile",  icon: "👤", label: "My Profile" },
    ...(canManageWorkspace ? [{ id: "agency", icon: "🏢", label: "Agency" }] : []),
    { id: "security", icon: "🔒", label: "Security" },
    { id: "notifications", icon: "🔔", label: `Notifications${unreadNotifications ? ` (${unreadNotifications})` : ""}` },
    { id: "activity", icon: "🕘", label: "Activity" },
    ...(canManageMembers ? [{ id: "team", icon: "🛡️", label: "Team & Roles" }] : []),
    ...(canDanger ? [{ id: "data", icon: "🗄️", label: "Data Management" }] : []),
  ];

  useEffect(() => {
    const requestedSection = initialSectionRequest?.section;
    if (requestedSection && tabs.find(tab => tab.id === requestedSection)) {
      setSection(requestedSection);
    }
  }, [initialSectionRequest?.key]);

  useEffect(() => {
    if (!tabs.find(tab => tab.id === section)) setSection("profile");
  }, [section, canManageWorkspace, canManageMembers, canDanger]);

  const InfoRow = ({ label, value, accent }) => (
    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "10px 0", borderBottom: "1px solid var(--border)", fontSize: 13 }}>
      <span style={{ color: "var(--text3)", fontWeight: 600 }}>{label}</span>
      <span style={{ fontWeight: 600, color: accent || "var(--text)" }}>{value || "—"}</span>
    </div>
  );

  return (
    <div className="fade">
      {toast && <Toast msg={toast.msg} type={toast.type} onDone={() => setToast(null)} />}
      {confirm && <Confirm msg={confirm.msg} onYes={confirm.onYes} onNo={() => setConfirm(null)} danger />}

      {/* Header */}
      <div style={{ marginBottom: 28 }}>
        <h1 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 800, fontSize: 24 }}>⚙️ System Settings</h1>
        <p style={{ color: "var(--text2)", marginTop: 3, fontSize: 13 }}>Manage your account, agency details, and platform preferences</p>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "200px 1fr", gap: 20, alignItems: "start" }}>

        {/* Tab sidebar */}
        <div style={{ background: "var(--bg2)", border: "1px solid var(--border)", borderRadius: 14, overflow: "hidden" }}>
          {tabs.map(t => (
            <button key={t.id} onClick={() => setSection(t.id)}
              style={{ display: "flex", alignItems: "center", gap: 10, width: "100%", padding: "12px 16px", border: "none", borderBottom: "1px solid var(--border)", background: section === t.id ? "rgba(139,92,246,.12)" : "transparent", color: section === t.id ? "var(--purple)" : "var(--text2)", fontFamily: "'Syne',sans-serif", fontWeight: section === t.id ? 700 : 400, fontSize: 13, cursor: "pointer", textAlign: "left", transition: "all .14s", borderLeft: section === t.id ? "3px solid var(--purple)" : "3px solid transparent" }}
              onMouseEnter={e => { if (section !== t.id) e.currentTarget.style.background = "var(--bg3)"; }}
              onMouseLeave={e => { if (section !== t.id) e.currentTarget.style.background = "transparent"; }}>
              <span style={{ fontSize: 16 }}>{t.icon}</span>{t.label}
            </button>
          ))}
          {/* Sign out button at bottom of tab panel */}
          <button onClick={() => setConfirm({ msg: "Sign out of MediaDesk Pro?", onYes: onLogout })}
            style={{ display: "flex", alignItems: "center", gap: 10, width: "100%", padding: "12px 16px", border: "none", background: "transparent", color: "var(--red)", fontFamily: "'Syne',sans-serif", fontWeight: 400, fontSize: 13, cursor: "pointer", textAlign: "left", transition: "all .14s" }}
            onMouseEnter={e => e.currentTarget.style.background = "rgba(239,68,68,.07)"}
            onMouseLeave={e => e.currentTarget.style.background = "transparent"}>
            <span style={{ fontSize: 16 }}>⎋</span> Sign Out
          </button>
        </div>

        {/* Panel content */}
        <div>

          {/* ── PROFILE ── */}
          {section === "profile" && (
            <div className="fade">
              {/* Avatar card */}
              <Card style={{ marginBottom: 16, padding: "20px 24px" }}>
                <div style={{ display: "flex", alignItems: "center", gap: 18 }}>
                  <div style={{ width: 64, height: 64, background: "linear-gradient(135deg,var(--accent),var(--purple))", borderRadius: "50%", display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "'Syne',sans-serif", fontWeight: 800, fontSize: 26, color: "#000", flexShrink: 0 }}>{user.name?.[0]?.toUpperCase() || "U"}</div>
                  <div>
                    <div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 800, fontSize: 20 }}>{user.name}</div>
                    <div style={{ color: "var(--text2)", fontSize: 13, marginTop: 2 }}>{user.title || "No title set"} · {user.agency}</div>
                    <div style={{ color: "var(--text3)", fontSize: 12, marginTop: 2 }}>{user.email}</div>
                  </div>
                </div>
              </Card>

              <Card>
                <h2 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 16, marginBottom: 18 }}>Edit Profile</h2>
                <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 }}>
                    <Field label="Full Name" value={pf.name} onChange={up("name")} placeholder="Jane Okafor" required />
                    <Field label="Job Title" value={pf.title} onChange={up("title")} placeholder="Media Buyer" />
                  </div>
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 }}>
                    <Field label="Email Address" type="email" value={pf.email} onChange={up("email")} placeholder="you@agency.com" required note="Used for login" />
                    <Field label="Phone Number" type="tel" value={pf.phone} onChange={up("phone")} placeholder="+234 800 000 0000" note="Used as Prepared By contact on MPOs" />
                  </div>
                  <div style={{ display: "flex", flexDirection: "column", gap: 10, padding: "12px 14px", borderRadius: 10, background: "var(--bg3)", border: "1px solid var(--border)" }}>
                    <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, flexWrap: "wrap" }}>
                      <div>
                        <div style={{ fontSize: 11, fontWeight: 700, color: "var(--text3)", textTransform: "uppercase", letterSpacing: ".08em" }}>Signature</div>
                        <div style={{ fontSize: 12, color: "var(--text2)", marginTop: 4 }}>Upload your signature image to show on MPO signatories.</div>
                      </div>
                      <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                        <Btn variant="ghost" size="sm" onClick={() => signatureInputRef.current?.click()}>Upload Signature</Btn>
                        {pf.signatureDataUrl ? <Btn variant="danger" size="sm" onClick={handleSignatureRemove}>Remove</Btn> : null}
                      </div>
                    </div>
                    <input ref={signatureInputRef} type="file" accept="image/*" style={{ display: "none" }} onChange={e => handleSignatureUpload(e.target.files?.[0])} />
                    {pf.signatureDataUrl ? (
                      <img src={pf.signatureDataUrl} alt="Signature preview" style={{ maxHeight: 62, maxWidth: 220, objectFit: "contain", background: "#fff", borderRadius: 8, border: "1px solid var(--border)" }} />
                    ) : (
                      <div style={{ fontSize: 12, color: "var(--text3)" }}>No signature uploaded yet.</div>
                    )}
                  </div>
                  <div style={{ background: "rgba(240,165,0,.07)", border: "1px solid rgba(240,165,0,.18)", borderRadius: 9, padding: "10px 14px", fontSize: 12, color: "var(--text2)" }}>
                    💡 Your <strong style={{ color: "var(--accent)" }}>name, title, and signature</strong> are automatically used in the "Prepared By" field on every MPO you generate.
                  </div>
                  <div style={{ display: "flex", justifyContent: "flex-end" }}>
                    <Btn onClick={saveProfile} icon="✓">Save Profile</Btn>
                  </div>
                </div>
              </Card>
            </div>
          )}

          {/* ── AGENCY ── */}
          {section === "agency" && (
            <div className="fade">
              <Card style={{ marginBottom: 16 }}>
                <div style={{ display: "flex", gap: 14, alignItems: "center", marginBottom: 10 }}>
                  <div style={{ width: 48, height: 48, background: "rgba(59,126,245,.12)", border: "1px solid rgba(59,126,245,.25)", borderRadius: 12, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 22, flexShrink: 0 }}>🏢</div>
                  <div><div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 17 }}>{user.agency}</div><div style={{ fontSize: 12, color: "var(--text2)", marginTop: 2 }}>{user.agencyAddress || "No address set"}</div><div style={{ display: "flex", gap: 12, flexWrap: "wrap", marginTop: 6, fontSize: 12, color: "var(--text2)" }}><span>✉️ {user.agencyEmail || "No agency email"}</span><span>📞 {user.agencyPhone || "No agency phone"}</span></div></div>
                </div>
                <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, background: "var(--bg3)", border: "1px solid var(--border)", borderRadius: 10, padding: "12px 14px" }}>
                  <div>
                    <div style={{ fontSize: 11, color: "var(--text3)", textTransform: "uppercase", letterSpacing: ".08em", fontWeight: 700 }}>Agency Invite Code</div>
                    <div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 800, fontSize: 18, marginTop: 4 }}>{user.agencyCode || "Not available yet"}</div>
                  </div>
                  {user.agencyCode && <Btn size="sm" onClick={() => navigator.clipboard?.writeText(user.agencyCode)} icon="⧉">Copy Code</Btn>}
                </div>
              </Card>
              <Card style={{ marginBottom: 16 }}>
                <h2 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 16, marginBottom: 18 }}>Agency Details</h2>
                <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
                  <Field label="Agency / Company Name" value={af.agency} onChange={ua("agency")} placeholder="Apex Media Ltd" required />
                  <Field label="Agency Address" value={af.address} onChange={ua("address")} placeholder="5, Craig Street, Ogudu GRA, Lagos" note="Used as the footer address on all MPO documents" />
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 }}>
                    <Field label="Agency Email" type="email" value={af.email} onChange={ua("email")} placeholder="hello@agency.com" />
                    <Field label="Agency Phone" type="tel" value={af.phone} onChange={ua("phone")} placeholder="+234 800 000 0000" />
                  </div>
                  <div style={{ display: "flex", justifyContent: "flex-end" }}>
                    {canManageWorkspace && <Btn onClick={saveAgency} icon="✓">Save Agency Details</Btn>}
                  </div>
                </div>
              </Card>
              <Card>
                <h2 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 16, marginBottom: 18 }}>Document & Billing Defaults</h2>
                <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 }}>
                    <Field label="VAT Rate (%)" type="number" value={docSettings.vatRate} onChange={udp("vatRate")} placeholder="7.5" note="Applied to new MPOs" />
                    <Field label="Session Timeout (hours)" type="number" value={docSettings.sessionHours} onChange={udp("sessionHours")} placeholder="8" note="Automatic logout after inactivity" />
                  </div>
                  <div style={{ display: "flex", alignItems: "center", gap: 10, padding: "10px 12px", background: "var(--bg3)", borderRadius: 10, border: "1px solid var(--border)" }}>
                    <input id="round-naira" type="checkbox" checked={docSettings.roundToWholeNaira} onChange={e => udp("roundToWholeNaira")(e.target.checked)} style={{ accentColor: "var(--accent)", width: 16, height: 16 }} />
                    <label htmlFor="round-naira" style={{ fontSize: 13, color: "var(--text2)" }}>Round new MPO totals to the nearest naira</label>
                  </div>
                  <Field label="Standard MPO Terms" value={docSettings.mpoTerms} onChange={udp("mpoTerms")} rows={7} placeholder="One term per line" note="Each line becomes a numbered contract term on new MPOs." />
                  <div style={{ display: "flex", justifyContent: "flex-end" }}>
                    {canManageWorkspace && <Btn onClick={saveDocumentSettings} icon="✓">Save Document Defaults</Btn>}
                  </div>
                </div>
              </Card>
            </div>
          )}

          {section === "team" && (
            <div className="fade">
              <Card>
                <h2 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 16, marginBottom: 12 }}>Team Roles & Permissions</h2>
                <p style={{ color: "var(--text2)", fontSize: 13, marginBottom: 16 }}>Admins can control what each user in the agency can edit.</p>
                <div style={{ display: "grid", gap: 10 }}>
                  {(members || []).length === 0 ? (
                    <div style={{ color: "var(--text2)", fontSize: 13 }}>No team members loaded yet.</div>
                  ) : (members || []).map(member => (
                    <div key={member.id} style={{ display: "grid", gridTemplateColumns: "1.4fr 1fr auto", gap: 12, alignItems: "center", padding: "12px 14px", background: "var(--bg3)", border: "1px solid var(--border)", borderRadius: 10 }}>
                      <div>
                        <div style={{ fontWeight: 700, fontSize: 13 }}>{member.name}{member.id === user.id ? " (You)" : ""}</div>
                        <div style={{ fontSize: 12, color: "var(--text2)", marginTop: 3 }}>{member.email}</div>
                        {member.title && <div style={{ fontSize: 11, color: "var(--text3)", marginTop: 2 }}>{member.title}</div>}
                      </div>
                      <Field value={memberRoles[member.id] || member.role || "viewer"} onChange={value => setMemberRoles(prev => ({ ...prev, [member.id]: normalizeRole(value) }))} options={Object.entries(ROLE_LABELS).map(([value, label]) => ({ value, label }))} />
                      <Btn size="sm" onClick={() => saveMemberRole(member.id)} loading={savingMemberId === member.id}>Save Role</Btn>
                    </div>
                  ))}
                </div>
                <div style={{ marginTop: 16, padding: "12px 14px", background: "rgba(59,126,245,.06)", border: "1px solid rgba(59,126,245,.16)", borderRadius: 10, fontSize: 12, color: "var(--text2)", lineHeight: 1.6 }}>
                  <strong>Role guide:</strong> Admin = full control. Planner = clients/campaigns/MPO creation. Buyer = vendors/rates/MPO creation. Finance = MPO status updates and reporting. Viewer = read-only.
                </div>
              </Card>
            </div>
          )}

          {/* ── SECURITY ── */}
          {section === "security" && (
            <div className="fade">
              <Card style={{ marginBottom: 16 }}>
                <h2 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 16, marginBottom: 6 }}>Account Info</h2>
                <InfoRow label="Registered Email" value={user.email} />
                <InfoRow label="Role" value={formatRoleLabel(user.role)} />
                <InfoRow label="Account Created" value={user.createdAt ? new Date(user.createdAt).toLocaleDateString("en-NG", { day: "numeric", month: "long", year: "numeric" }) : "—"} />
                <InfoRow label="User ID" value={user.id?.slice(0, 12) + "…"} accent="var(--text3)" />
              </Card>
              <Card>
                <h2 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 16, marginBottom: 18 }}>Change Password</h2>
                <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
                  <Field label="Current Password" type="password" value={sf.current} onChange={us("current")} placeholder="••••••••" />
                  <Field label="New Password" type="password" value={sf.newPw} onChange={us("newPw")} placeholder="••••••••" note="Minimum 6 characters" />
                  <Field label="Confirm New Password" type="password" value={sf.confirm} onChange={us("confirm")} placeholder="••••••••" />
                  <div style={{ display: "flex", justifyContent: "flex-end" }}>
                    <Btn onClick={changePassword} icon="🔒">Update Password</Btn>
                  </div>
                </div>
              </Card>
            </div>
          )}

          {/* ── NOTIFICATIONS ── */}
          {section === "notifications" && (
            <div className="fade">
              <Card>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 12, flexWrap: "wrap", marginBottom: 16 }}>
                  <div>
                    <h2 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 16 }}>Notifications Inbox</h2>
                    <div style={{ fontSize: 12, color: "var(--text2)", marginTop: 4 }}>Alerts for approvals, finance updates, proof of airing, and reconciliation milestones.</div>
                  </div>
                  <div style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
                    <div style={{ minWidth: 220 }}>
                      <Field value={notificationFilter} onChange={setNotificationFilter} options={[
                        { value: "all", label: "All Notifications" },
                        { value: "unread", label: "Unread Only" },
                        { value: "mpo", label: "MPO Workflow" },
                        { value: "finance", label: "Finance" },
                        { value: "proof", label: "Proof & Reconciliation" },
                        { value: "workspace", label: "Workspace" },
                      ]} />
                    </div>
                    <Btn variant="ghost" size="sm" onClick={onMarkAllNotificationsRead} disabled={!unreadNotifications}>Mark all as read</Btn>
                  </div>
                </div>
                {(() => {
                  const filteredNotifications = (notifications || []).filter(notification => {
                    if (notificationFilter === "unread" && notification.readAt) return false;
                    if (notificationFilter !== "all" && notificationFilter !== "unread" && (notification.category || "workspace") !== notificationFilter) return false;
                    return true;
                  });
                  if (!filteredNotifications.length) return <Empty icon="🔔" title="No notifications" sub="You're all caught up for this workspace." />;
                  return (
                    <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                      {filteredNotifications.map(notification => (
                        <div key={notification.id} style={{ border: "1px solid var(--border)", borderRadius: 10, padding: "12px 14px", background: notification.readAt ? "var(--bg3)" : "rgba(240,165,0,.07)" }}>
                          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 10, flexWrap: "wrap" }}>
                            <div>
                              <div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 14 }}>{notification.title}</div>
                              <div style={{ fontSize: 12, color: "var(--text2)", marginTop: 4 }}>{notification.message || "Open the related workspace page for more details."}</div>
                              <div style={{ fontSize: 11, color: "var(--text3)", marginTop: 8 }}>{notification.actorName || "System"} · {formatRoleLabel(notification.actorRole || "viewer")} · {formatAuditTimestamp(notification.createdAt)}</div>
                            </div>
                            <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                              {!notification.readAt ? <Badge color="accent">Unread</Badge> : <Badge color="gray">Read</Badge>}
                              {!notification.readAt ? <Btn variant="ghost" size="sm" onClick={() => onMarkNotificationRead(notification.id)}>Mark as read</Btn> : null}
                            </div>
                          </div>
                        </div>
                      ))}
                    </div>
                  );
                })()}
              </Card>
            </div>
          )}

          {/* ── ACTIVITY ── */}
          {section === "activity" && (
            <div className="fade">
              <Card>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 12, flexWrap: "wrap", marginBottom: 16 }}>
                  <div>
                    <h2 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 16 }}>Workspace Activity</h2>
                    <div style={{ fontSize: 12, color: "var(--text2)", marginTop: 4 }}>Approval actions, MPO workflow changes, and key workspace edits appear here.</div>
                  </div>
                  <div style={{ minWidth: 220 }}>
                    <Field value={activityFilter} onChange={setActivityFilter} options={[
                      { value: "all", label: "All Activity" },
                      { value: "mpo", label: "MPO Workflow" },
                      { value: "member", label: "Team & Roles" },
                      { value: "agency", label: "Agency" },
                      { value: "workspace", label: "Workspace Settings" },
                      { value: "profile", label: "Profiles" },
                      { value: "security", label: "Security" },
                    ]} />
                  </div>
                </div>
                {activityLoading ? (
                  <div style={{ color: "var(--text2)", fontSize: 13 }}>Loading activity…</div>
                ) : activityItems.length ? (
                  <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                    {activityItems.map(event => (
                      <div key={event.id} style={{ border: "1px solid var(--border)", borderRadius: 10, padding: "12px 14px", background: "var(--bg3)" }}>
                        <div style={{ display: "flex", justifyContent: "space-between", gap: 12, flexWrap: "wrap" }}>
                          <div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 14 }}>
                            {(event.recordType || "workspace").toUpperCase()} · {(event.action || "updated").replace(/_/g, " ").toUpperCase()}
                          </div>
                          <div style={{ fontSize: 12, color: "var(--text3)" }}>{formatAuditTimestamp(event.createdAt)}</div>
                        </div>
                        <div style={{ fontSize: 12, color: "var(--text2)", marginTop: 4 }}>
                          {event.actorName || "System"} · {formatRoleLabel(event.actorRole || "viewer")}
                        </div>
                        {event.note ? <div style={{ marginTop: 8, fontSize: 13 }}>{event.note}</div> : null}
                        {event.metadata && Object.keys(event.metadata || {}).length ? (
                          <div style={{ marginTop: 10, display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(180px,1fr))", gap: 8 }}>
                            {Object.entries(event.metadata).map(([key, value]) => (
                              <div key={key} style={{ background: "var(--bg2)", border: "1px solid var(--border)", borderRadius: 8, padding: "8px 10px", fontSize: 12 }}>
                                <div style={{ color: "var(--text3)", textTransform: "uppercase", letterSpacing: ".04em", fontSize: 10, fontWeight: 700 }}>{key.replace(/_/g, " ")}</div>
                                <div style={{ marginTop: 4, wordBreak: "break-word" }}>{Array.isArray(value) ? value.join(", ") : String(value ?? "—")}</div>
                              </div>
                            ))}
                          </div>
                        ) : null}
                      </div>
                    ))}
                  </div>
                ) : <Empty icon="🕘" title="No activity yet" sub="Actions will start appearing here once your team begins working." />}
              </Card>
            </div>
          )}

          {/* ── DATA ── */}
          {section === "data" && (
            <div className="fade">
              <Card style={{ marginBottom: 14 }}>
                <h2 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 16, marginBottom: 6 }}>Workspace Data & Maintenance</h2>
                <p style={{ color: "var(--text2)", fontSize: 13, lineHeight: 1.6, marginBottom: 16 }}>Your operational records now live in your Supabase workspace. This page is for exporting a snapshot, reviewing what is synced, and cleaning up browser-only preferences on this device.</p>
                <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 16 }}>
                  <Btn variant="blue" icon="⬇" onClick={exportBackup}>Export Workspace Snapshot</Btn>
                  <Btn variant="secondary" icon="⬆" onClick={() => backupInputRef.current?.click()}>Import Snapshot</Btn>
                  <input ref={backupInputRef} type="file" accept="application/json" style={{ display: "none" }} onChange={e => handleBackupImport(e.target.files?.[0])} />
                </div>
                <div style={{ background: "var(--bg3)", borderRadius: 10, border: "1px solid var(--border)", padding: "12px 14px", marginBottom: 14 }}>
                  <div style={{ fontSize: 12, color: "var(--text2)", lineHeight: 1.6 }}>Snapshot export includes vendors, clients, campaigns, rates, MPOs, notifications, and document settings for review. Direct browser snapshot import is intentionally disabled for cloud workspaces to prevent overwriting live Supabase data accidentally.</div>
                </div>
                {[
                  [vendors, "Vendors"], [clients, "Clients & Brands"],
                  [campaigns, "Campaigns"], [rates, "Media Rates"], [mpos, "MPOs"], [notifications, "Notifications"],
                ].map(([data, label]) => (
                  <div key={label} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "10px 0", borderBottom: "1px solid var(--border)", fontSize: 13 }}>
                    <span style={{ color: "var(--text2)" }}>{label}</span>
                    <span style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, color: (data?.length || 0) > 0 ? "var(--accent)" : "var(--text3)" }}>{data?.length || 0} record{(data?.length || 0) !== 1 ? "s" : ""}</span>
                  </div>
                ))}
              </Card>
              <Card style={{ marginBottom: 14 }}>
                <h2 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 16, marginBottom: 6 }}>Browser-only Preferences</h2>
                <p style={{ color: "var(--text2)", fontSize: 13, marginBottom: 16, lineHeight: 1.6 }}>These controls affect only this browser. They do not delete live Supabase records for your agency.</p>
                <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(220px,1fr))", gap: 10 }}>
                  <div style={{ padding: "12px 14px", background: "var(--bg3)", borderRadius: 10, border: "1px solid var(--border)" }}>
                    <div style={{ fontSize: 11, color: "var(--text3)", textTransform: "uppercase", letterSpacing: ".08em", fontWeight: 700 }}>Theme Preference</div>
                    <div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, marginTop: 6 }}>{store.get(themeKeyForUser(user?.id || null), getDefaultTheme(user?.id || null))}</div>
                    <div style={{ fontSize: 11, color: "var(--text3)", marginTop: 4 }}>Stored per user on this device.</div>
                  </div>
                  <div style={{ padding: "12px 14px", background: "var(--bg3)", borderRadius: 10, border: "1px solid var(--border)" }}>
                    <div style={{ fontSize: 11, color: "var(--text3)", textTransform: "uppercase", letterSpacing: ".08em", fontWeight: 700 }}>Signature Cache</div>
                    <div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, marginTop: 6 }}>{user?.signatureDataUrl ? "Available" : "Not saved"}</div>
                    <div style={{ fontSize: 11, color: "var(--text3)", marginTop: 4 }}>Used to keep the signatory visible between sessions.</div>
                  </div>
                </div>
              </Card>
              <Card style={{ border: "1px solid rgba(239,68,68,.25)", background: "rgba(239,68,68,.04)" }}>
                <h2 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 16, marginBottom: 6, color: "var(--red)" }}>⚠️ Local Cleanup</h2>
                <p style={{ color: "var(--text2)", fontSize: 13, marginBottom: 16, lineHeight: 1.6 }}>Use these only when this browser has stale preferences. Your Supabase workspace records remain intact.</p>
                <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                  {[
                    { label: "Clear MPO Draft Cache", desc: "Remove the autosaved draft on this browser only.", onYes: () => { store.del("msp_mpo_draft"); setToast({ msg: "Local MPO draft cache cleared.", type: "success" }); setConfirm(null); } },
                    { label: "Reset My Theme Preference", desc: "Forget this user’s local light/dark mode on this browser.", onYes: () => { store.del(themeKeyForUser(user?.id || null)); setToast({ msg: "Local theme preference reset. Reload to apply system default.", type: "success" }); setConfirm(null); } },
                    { label: "Clear My Saved Signature Cache", desc: "Removes only the browser cache copy. Your saved profile value stays in Supabase metadata if available.", onYes: () => { setStoredUserSignature(user?.id, ""); setToast({ msg: "Local signature cache cleared.", type: "success" }); setConfirm(null); } },
                  ].map(({ label, desc, onYes }) => (
                    <div key={label} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "12px 14px", background: "var(--bg3)", borderRadius: 10, border: "1px solid var(--border2)" }}>
                      <div><div style={{ fontSize: 13, fontWeight: 600 }}>{label}</div><div style={{ fontSize: 11, color: "var(--text3)", marginTop: 2 }}>{desc}</div></div>
                      {canDanger && <Btn variant="danger" size="sm" onClick={() => setConfirm({ msg: `${label}? This only affects this browser.`, onYes })}>Clear</Btn>}
                    </div>
                  ))}
                </div>
              </Card>
            </div>
          )}

        </div>
      </div>
    </div>
  );
};






/* ── BUDGETING & BILLING ─────────────────────────────────── */

export default SettingsPage;
