import { useEffect, useMemo, useState } from "react";
import { supabase } from "../../lib/supabase";
import { normalizeAgencyCode, findExistingAgencyByName } from "../../services/agencies";

/* ── ICONS ─────────────────────────────────────────────────── */
const EyeIcon = ({ crossed = false }) => (
  <svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <path d="M2 12s3.5-6 10-6 10 6 10 6-3.5 6-10 6S2 12 2 12Z" />
    <circle cx="12" cy="12" r="3" />
    {crossed && <path d="M4 4l16 16" />}
  </svg>
);

const GoogleIcon = () => (
  <svg width="18" height="18" viewBox="0 0 18 18" fill="none">
    <path d="M17.64 9.2a10.34 10.34 0 0 0-.16-1.84H9v3.48h4.84a4.14 4.14 0 0 1-1.8 2.72v2.26h2.92C16.68 12.01 17.64 9.69 17.64 9.2Z" fill="#4285F4"/>
    <path d="M9 18c2.43 0 4.47-.8 5.96-2.18l-2.92-2.26c-.81.54-1.84.86-3.04.86-2.34 0-4.32-1.58-5.03-3.71H.99v2.34A8.997 8.997 0 0 0 9 18Z" fill="#34A853"/>
    <path d="M3.97 10.71A5.41 5.41 0 0 1 3.69 9c0-.59.1-1.17.28-1.71V4.95H.99A9 9 0 0 0 0 9c0 1.45.35 2.82.99 4.05l2.98-2.34Z" fill="#FBBC05"/>
    <path d="M9 3.58c1.32 0 2.51.45 3.44 1.35l2.58-2.58C13.46.89 11.43 0 9 0A8.997 8.997 0 0 0 .99 4.95L3.97 7.29C4.68 5.16 6.66 3.58 9 3.58Z" fill="#EA4335"/>
  </svg>
);

/* ── PILL INPUT ─────────────────────────────────────────────── */
const PillInput = ({ label, type = "text", value, onChange, placeholder, required, autoComplete, note }) => (
  <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
    {label && (
      <label style={{ fontSize: 13, fontWeight: 600, color: "#64748b", fontFamily: "'Inter', sans-serif" }}>
        {label}{required && <span style={{ color: "#ef4444", marginLeft: 3 }}>*</span>}
      </label>
    )}
    <input
      type={type}
      value={value}
      onChange={(e) => onChange(e.target.value)}
      placeholder={placeholder}
      required={required}
      autoComplete={autoComplete}
      style={{
        background: "#eef0f8",
        border: "none",
        borderRadius: 999,
        padding: "13px 22px",
        fontSize: 14,
        color: "#0f172a",
        fontFamily: "'Inter', sans-serif",
        outline: "none",
        width: "100%",
      }}
    />
    {note && <span style={{ fontSize: 11, color: "#94a3b8", paddingLeft: 8, fontFamily: "'Inter', sans-serif" }}>{note}</span>}
  </div>
);

/* ── PILL PASSWORD ──────────────────────────────────────────── */
const PillPassword = ({ label, value, onChange, placeholder, required, visible, onToggle, autoComplete }) => (
  <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
    {label && (
      <label style={{ fontSize: 13, fontWeight: 600, color: "#64748b", fontFamily: "'Inter', sans-serif" }}>
        {label}{required && <span style={{ color: "#ef4444", marginLeft: 3 }}>*</span>}
      </label>
    )}
    <div style={{ position: "relative" }}>
      <input
        type={visible ? "text" : "password"}
        value={value}
        onChange={(e) => onChange(e.target.value)}
        placeholder={placeholder}
        required={required}
        autoComplete={autoComplete}
        style={{
          background: "#eef0f8",
          border: "none",
          borderRadius: 999,
          padding: "13px 52px 13px 22px",
          fontSize: 14,
          color: "#0f172a",
          fontFamily: "'Inter', sans-serif",
          outline: "none",
          width: "100%",
        }}
      />
      <button
        type="button"
        onClick={onToggle}
        aria-label={`${visible ? "Hide" : "Show"} password`}
        style={{
          position: "absolute", right: 18, top: "50%", transform: "translateY(-50%)",
          border: "none", background: "none", color: "#94a3b8", cursor: "pointer",
          display: "flex", alignItems: "center", padding: 0, lineHeight: 0,
        }}
      >
        <EyeIcon crossed={visible} />
      </button>
    </div>
  </div>
);

/* ── RIGHT PANEL ────────────────────────────────────────────── */
const RightPanel = ({ mode }) => {
  const copy = mode === "register"
    ? { title: "Run your media buying Smarter.", sub: "MPOs · Clients · Campaigns · Reports\nall in one place." }
    : { title: "Run your media buying Smarter.", sub: "MPOs · Clients · Campaigns · Reports\nall in one place." };

  return (
    <div style={{
      flex: "0 0 44%",
      background: "linear-gradient(150deg, #f0a500 0%, #c97d00 100%)",
      position: "relative",
      overflow: "hidden",
      display: "flex",
      flexDirection: "column",
      alignItems: "center",
      justifyContent: "center",
      padding: "40px 28px",
      gap: 20,
    }}>
      {/* Decorative blobs */}
      <div style={{ position: "absolute", top: "-18%", right: "-12%", width: "65%", paddingTop: "65%", borderRadius: "50%", background: "rgba(255,255,255,.09)", pointerEvents: "none" }} />
      <div style={{ position: "absolute", bottom: "-12%", left: "-18%", width: "70%", paddingTop: "70%", borderRadius: "50%", background: "rgba(255,255,255,.07)", pointerEvents: "none" }} />
      <div style={{ position: "absolute", top: "38%", left: "5%", width: "40%", paddingTop: "40%", borderRadius: "50%", background: "rgba(255,255,255,.05)", pointerEvents: "none" }} />

      {/* Copy */}
      <div style={{ zIndex: 1, textAlign: "center" }}>
        <h3 style={{
          margin: "0 0 10px", color: "#fff",
          fontFamily: "'Plus Jakarta Sans', sans-serif",
          fontWeight: 900, fontSize: "clamp(22px, 2.4vw, 28px)", lineHeight: 1.2,
        }}>
          {copy.title}
        </h3>
        <p style={{
          margin: 0, color: "rgba(255,255,255,.75)",
          fontSize: 13, lineHeight: 1.65,
          fontFamily: "'Inter', sans-serif", whiteSpace: "pre-line",
        }}>
          {copy.sub}
        </p>
      </div>

      {/* Stat cards */}
      <div style={{ zIndex: 1, width: "100%", maxWidth: 280, position: "relative" }}>
        {/* Card 1 */}
        <div style={{
          background: "rgba(255,255,255,.18)",
          backdropFilter: "blur(8px)",
          borderRadius: 18,
          padding: "18px 20px",
          marginBottom: 0,
          position: "relative",
          zIndex: 2,
        }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 10 }}>
            <span style={{ fontSize: 16 }}>📋</span>
            <span style={{ color: "#fff", fontWeight: 700, fontSize: 13, fontFamily: "'Plus Jakarta Sans', sans-serif" }}>Active MPOs</span>
          </div>
          <div style={{ color: "#fff", fontWeight: 900, fontSize: 40, lineHeight: 1, fontFamily: "'Plus Jakarta Sans', sans-serif" }}>24</div>
          <div style={{ color: "rgba(255,255,255,.65)", fontSize: 12, marginTop: 5, fontFamily: "'Inter', sans-serif" }}>across all clients</div>

        </div>

        {/* Card 2 — offset right */}
        <div style={{
          background: "rgba(255,255,255,.13)",
          backdropFilter: "blur(8px)",
          borderRadius: 18,
          padding: "26px 20px 18px",
          marginTop: 10,
          marginLeft: 24,
        }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 10 }}>
            <span style={{ fontSize: 16 }}>✅</span>
            <span style={{ color: "#fff", fontWeight: 700, fontSize: 13, fontFamily: "'Plus Jakarta Sans', sans-serif" }}>Execution Rate</span>
          </div>
          <div style={{ color: "#fff", fontWeight: 900, fontSize: 40, lineHeight: 1, fontFamily: "'Plus Jakarta Sans', sans-serif" }}>99.99%</div>
          {/* Progress segments */}
          <div style={{ marginTop: 10, display: "flex", gap: 3 }}>
            {Array.from({ length: 9 }).map((_, i) => (
              <div key={i} style={{
                flex: 1, height: 3, borderRadius: 99,
                background: "rgba(255,255,255,.75)",
              }} />
            ))}
          </div>
        </div>
      </div>
    </div>
  );
};

/* ── MAIN ───────────────────────────────────────────────────── */
const AuthPage = ({ onLogin, sessionExpired = false }) => {
  const [mode, setMode] = useState("login");
  const [f, setF] = useState({
    name: "", email: "", password: "", agency: "",
    agencyCode: "", agencyMode: "create", title: "", phone: "", confirm: "",
  });
  const [err, setErr] = useState("");
  const [info, setInfo] = useState("");
  const [loading, setLoading] = useState(false);
  const [resetLoading, setResetLoading] = useState(false);
  const [existingAgencyMatch, setExistingAgencyMatch] = useState(null);
  const [agencyCheckLoading, setAgencyCheckLoading] = useState(false);
  const [visiblePasswords, setVisiblePasswords] = useState({ password: false, confirm: false });
  const [rememberMe, setRememberMe] = useState(false);

  const isRecoveryMode = mode === "recovery";
  const isRegisterMode = mode === "register";

  const heading = useMemo(() => {
    if (mode === "login") return "Welcome back";
    if (mode === "recovery") return "New password";
    return "Create account";
  }, [mode]);

  const subheading = useMemo(() => {
    if (mode === "login") return "Sign in to continue to your workspace.";
    if (mode === "recovery") return "Set a new password to regain access.";
    return "Register your profile and set up your agency workspace.";
  }, [mode]);

  const passwordLabel = useMemo(() => isRecoveryMode ? "New Password" : "Password", [isRecoveryMode]);

  const u = (k) => (v) => { setErr(""); setInfo(""); setF((p) => ({ ...p, [k]: v })); };
  const togglePw = (key) => setVisiblePasswords((p) => ({ ...p, [key]: !p[key] }));

  useEffect(() => {
    const prev = { html: document.documentElement.style.backgroundColor, body: document.body.style.backgroundColor };
    const root = document.getElementById("root");
    const prevRoot = root?.style.backgroundColor || "";
    const bg = "#eaedf5";
    document.documentElement.style.backgroundColor = bg;
    document.body.style.backgroundColor = bg;
    if (root) root.style.backgroundColor = bg;
    return () => {
      document.documentElement.style.backgroundColor = prev.html;
      document.body.style.backgroundColor = prev.body;
      if (root) root.style.backgroundColor = prevRoot;
    };
  }, []);

  useEffect(() => {
    const hash = window.location.hash || "";
    const params = new URLSearchParams(hash.startsWith("#") ? hash.slice(1) : hash);
    if (params.get("type") === "recovery") { setMode("recovery"); setInfo("Enter your new password to complete the reset."); }
    const { data: { subscription } } = supabase.auth.onAuthStateChange((event) => {
      if (event === "PASSWORD_RECOVERY") { setMode("recovery"); setInfo("Enter your new password to complete the reset."); }
    });
    return () => subscription.unsubscribe();
  }, []);

  useEffect(() => {
    let active = true;
    const check = async () => {
      if (mode !== "register" || f.agencyMode !== "create" || !f.agency.trim()) {
        if (active) { setExistingAgencyMatch(null); setAgencyCheckLoading(false); }
        return;
      }
      setAgencyCheckLoading(true);
      try {
        const match = await findExistingAgencyByName(f.agency.trim());
        if (active) setExistingAgencyMatch(match || null);
      } catch { if (active) setExistingAgencyMatch(null); }
      finally { if (active) setAgencyCheckLoading(false); }
    };
    const t = setTimeout(check, 250);
    return () => { active = false; clearTimeout(t); };
  }, [mode, f.agencyMode, f.agency]);

  const handleForgotPassword = async () => {
    setErr(""); setInfo("");
    if (!f.email.trim()) { setErr("Enter your email first, then click Forgot password."); return; }
    setResetLoading(true);
    try {
      const redirectTo = `${window.location.origin}${window.location.pathname}`;
      const { error } = await supabase.auth.resetPasswordForEmail(f.email.trim(), { redirectTo });
      if (error) throw error;
      setInfo("Reset email sent. Check your inbox.");
    } catch (e) { setErr(e.message || "Failed to send reset email."); }
    finally { setResetLoading(false); }
  };

  const handleGoogleSignIn = async () => {
    setErr("");
    try {
      const { error } = await supabase.auth.signInWithOAuth({ provider: "google", options: { redirectTo: window.location.origin } });
      if (error) throw error;
    } catch (e) { setErr(e.message || "Google sign-in failed."); }
  };

  const submit = async () => {
    setErr(""); setInfo("");
    if (mode === "register") {
      if (!f.name || !f.email || !f.password) return setErr("Full name, email, and password are required.");
      if (f.agencyMode === "create" && !f.agency.trim()) return setErr("Agency name is required.");
      if (f.agencyMode === "create" && existingAgencyMatch) return setErr("Agency already exists — contact admin.");
      if (f.agencyMode === "join" && !normalizeAgencyCode(f.agencyCode)) return setErr("Invite code is required.");
      if (f.password !== f.confirm) return setErr("Passwords do not match.");
      if (f.password.length < 6) return setErr("Password must be at least 6 characters.");
    }
    if (mode === "recovery") {
      if (!f.password) return setErr("New password is required.");
      if (f.password.length < 6) return setErr("Password must be at least 6 characters.");
      if (f.password !== f.confirm) return setErr("Passwords do not match.");
    }
    setLoading(true);
    try {
      if (mode === "login") {
        const { error } = await supabase.auth.signInWithPassword({ email: f.email, password: f.password });
        if (error) throw error;
        if (typeof onLogin === "function") onLogin();
      } else if (mode === "recovery") {
        const { error } = await supabase.auth.updateUser({ password: f.password });
        if (error) throw error;
        setInfo("Password updated. Sign in with your new password.");
        setMode("login");
        setF((p) => ({ ...p, password: "", confirm: "" }));
        window.history.replaceState({}, document.title, window.location.pathname + window.location.search);
      } else {
        const { data, error } = await supabase.auth.signUp({
          email: f.email, password: f.password,
          options: { data: { full_name: f.name, title: f.title || "", phone: f.phone || "", agency_name: f.agencyMode === "create" ? f.agency.trim() : "", agency_code: f.agencyMode === "join" ? normalizeAgencyCode(f.agencyCode) : "", agency_mode: f.agencyMode } },
        });
        if (error) throw error;
        if (!data.user) throw new Error("Signup failed.");
        const { error: profileError } = await supabase.from("profiles").update({ full_name: f.name, title: f.title || "", phone: f.phone || "" }).eq("id", data.user.id);
        if (profileError) throw profileError;
        if (!data.session) { setInfo("Account created! Check your email to confirm, then sign in."); setMode("login"); }
      }
    } catch (e) { setErr(e.message || "Something went wrong."); }
    finally { setLoading(false); }
  };

  const switchMode = (next) => {
    setMode(next); setErr(""); setInfo("");
    setVisiblePasswords({ password: false, confirm: false });
    if (next !== "register") { setExistingAgencyMatch(null); setAgencyCheckLoading(false); }
    if (next === "login") setF((p) => ({ ...p, password: "", confirm: "" }));
  };

  const submitLabel = mode === "login" ? "Sign In"
    : mode === "recovery" ? "Update Password"
    : existingAgencyMatch && f.agencyMode === "create" ? "Agency Already Exists"
    : "Create Account";

  /* ── shared input style for inline inputs (agency/invite code) ── */
  const pillInlineInput = {
    background: "#eef0f8", border: "none", borderRadius: 999,
    padding: "13px 22px", fontSize: 14, color: "#0f172a",
    fontFamily: "'Inter', sans-serif", outline: "none", width: "100%",
  };

  return (
    <div style={{
      minHeight: "100vh", width: "100%",
      background: "#eaedf5",
      display: "flex", alignItems: "center", justifyContent: "center",
      padding: "32px 20px", boxSizing: "border-box",
    }}>
      <div style={{
        width: "100%", maxWidth: 900,
        borderRadius: 26,
        overflow: "hidden",
        display: "flex",
        background: "#fff",
        boxShadow: "0 24px 64px rgba(100,100,160,.14)",
        minHeight: 560,
      }}>
        {/* ── LEFT PANEL ── */}
        <div style={{
          flex: 1,
          background: "#fff",
          padding: "40px 48px 32px",
          display: "flex",
          flexDirection: "column",
          overflowY: "auto",
          maxHeight: "92vh",
        }}>
          {/* Logo pill */}
          <div style={{ marginBottom: 32, alignSelf: "flex-start" }}>
            <div style={{
              display: "inline-flex", alignItems: "center", gap: 8,
              border: "1.5px solid #e2e8f0", borderRadius: 999,
              padding: "7px 16px 7px 7px", background: "#fff",
            }}>
              <div style={{
                width: 30, height: 30, borderRadius: "50%",
                background: "#f0a500",
                display: "flex", alignItems: "center", justifyContent: "center",
                fontSize: 14, lineHeight: 1, flexShrink: 0,
              }}>
                📡
              </div>
              <span style={{ fontFamily: "'Plus Jakarta Sans', sans-serif", fontWeight: 800, fontSize: 14, color: "#0f172a" }}>
                MediaDesk Pro
              </span>
            </div>
          </div>

          {/* Session expired */}
          {sessionExpired && (
            <div style={{ background: "rgba(240,165,0,.08)", border: "1px solid rgba(240,165,0,.3)", borderRadius: 12, padding: "11px 16px", marginBottom: 20, fontSize: 13, color: "#92610a", lineHeight: 1.5, fontFamily: "'Inter', sans-serif" }}>
              Your session expired. Sign in again to continue.
            </div>
          )}

          {/* Heading */}
          <h2 style={{ margin: "0 0 6px", fontFamily: "'Plus Jakarta Sans', sans-serif", fontWeight: 900, fontSize: "clamp(26px, 3vw, 34px)", color: "#0f172a", letterSpacing: "-.03em", lineHeight: 1.1 }}>
            {heading}
          </h2>
          <p style={{ margin: "0 0 26px", color: "#94a3b8", fontSize: 14, lineHeight: 1.5, fontFamily: "'Inter', sans-serif" }}>
            {subheading}
          </p>

          {/* Form */}
          <div style={{ display: "flex", flexDirection: "column", gap: 14, flex: 1 }}>

            {/* Register-only fields */}
            {isRegisterMode && (
              <>
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
                  <PillInput label="Full Name" value={f.name} onChange={u("name")} placeholder="Jane Okafor" required />
                  <PillInput label="Job Title" value={f.title} onChange={u("title")} placeholder="Media Buyer" />
                </div>

                {/* Agency mode */}
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 6, background: "#f5f7fc", borderRadius: 999, padding: 4 }}>
                  {[
                    { key: "create", label: "New Agency" },
                    { key: "join", label: "Join With Code" },
                  ].map((opt) => {
                    const active = f.agencyMode === opt.key;
                    return (
                      <button
                        key={opt.key}
                        type="button"
                        onClick={() => { setErr(""); setInfo(""); if (opt.key === "create") setExistingAgencyMatch(null); setF((p) => ({ ...p, agencyMode: opt.key })); }}
                        style={{
                          border: "none",
                          background: active ? "#f0a500" : "transparent",
                          color: active ? "#000" : "#64748b",
                          borderRadius: 999,
                          padding: "9px 12px",
                          fontWeight: 700, fontSize: 12,
                          cursor: "pointer",
                          transition: "all .18s",
                          fontFamily: "'Plus Jakarta Sans', sans-serif",
                        }}
                      >
                        {opt.label}
                      </button>
                    );
                  })}
                </div>

                {f.agencyMode === "create" ? (
                  <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                    <label style={{ fontSize: 13, fontWeight: 600, color: "#64748b", fontFamily: "'Inter', sans-serif" }}>
                      Agency Name <span style={{ color: "#ef4444" }}>*</span>
                    </label>
                    <input
                      type="text"
                      value={f.agency}
                      onChange={(e) => { setErr(""); setInfo(""); u("agency")(e.target.value); }}
                      placeholder="Apex Media Ltd"
                      style={pillInlineInput}
                    />
                    <span style={{ fontSize: 11, color: "#94a3b8", paddingLeft: 8, fontFamily: "'Inter', sans-serif" }}>
                      {agencyCheckLoading ? "Checking…" : "Only for brand-new workspaces."}
                    </span>
                    {existingAgencyMatch && (
                      <div style={{ background: "rgba(239,68,68,.06)", border: "1px solid rgba(239,68,68,.2)", borderRadius: 12, padding: "10px 14px", fontSize: 13, color: "#dc2626", lineHeight: 1.5, fontFamily: "'Inter', sans-serif" }}>
                        <strong>{existingAgencyMatch.name}</strong> already exists. Contact your admin instead.
                      </div>
                    )}
                  </div>
                ) : (
                  <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                    <label style={{ fontSize: 13, fontWeight: 600, color: "#64748b", fontFamily: "'Inter', sans-serif" }}>
                      Invite Code <span style={{ color: "#ef4444" }}>*</span>
                    </label>
                    <input
                      type="text"
                      value={f.agencyCode}
                      onChange={(e) => { setErr(""); setInfo(""); u("agencyCode")(e.target.value); }}
                      placeholder="QVT-7K4P"
                      style={pillInlineInput}
                    />
                    <span style={{ fontSize: 11, color: "#94a3b8", paddingLeft: 8, fontFamily: "'Inter', sans-serif" }}>Ask your agency admin for the code.</span>
                  </div>
                )}

                <PillInput label="Phone" type="tel" value={f.phone} onChange={u("phone")} placeholder="+234 800 000 0000" />
              </>
            )}

            {/* Email */}
            <PillInput
              label="Email"
              type="email"
              value={f.email}
              onChange={u("email")}
              placeholder="Enter your email"
              required
              autoComplete="email"
            />

            {/* Password */}
            <PillPassword
              label={passwordLabel}
              value={f.password}
              onChange={u("password")}
              visible={visiblePasswords.password}
              onToggle={() => togglePw("password")}
              autoComplete={mode === "login" ? "current-password" : "new-password"}
              placeholder="••••••••"
              required
            />

            {(isRegisterMode || isRecoveryMode) && (
              <PillPassword
                label="Confirm Password"
                value={f.confirm}
                onChange={u("confirm")}
                visible={visiblePasswords.confirm}
                onToggle={() => togglePw("confirm")}
                autoComplete="new-password"
                placeholder="••••••••"
              />
            )}

            {/* Login: remember + forgot */}
            {mode === "login" && (
              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
                <label style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13, color: "#64748b", cursor: "pointer", fontFamily: "'Inter', sans-serif" }}>
                  <input type="checkbox" checked={rememberMe} onChange={(e) => setRememberMe(e.target.checked)} style={{ accentColor: "#f0a500", width: 14, height: 14 }} />
                  Remember me
                </label>
                <button
                  type="button"
                  onClick={handleForgotPassword}
                  disabled={resetLoading || loading}
                  style={{ background: "none", border: "none", color: "#f0a500", fontWeight: 600, fontSize: 13, cursor: "pointer", fontFamily: "'Inter', sans-serif", padding: 0 }}
                >
                  {resetLoading ? "Sending…" : "Forgot password?"}
                </button>
              </div>
            )}

            {/* Recovery back link */}
            {isRecoveryMode && (
              <button
                type="button"
                onClick={() => switchMode("login")}
                style={{ background: "none", border: "none", color: "#f0a500", fontWeight: 600, fontSize: 13, cursor: "pointer", fontFamily: "'Inter', sans-serif", padding: 0, textAlign: "left" }}
              >
                ← Back to sign in
              </button>
            )}

            {/* Banners */}
            {info && (
              <div style={{ background: "rgba(34,197,94,.07)", border: "1px solid rgba(34,197,94,.2)", borderRadius: 12, padding: "11px 16px", color: "#166534", fontSize: 13, lineHeight: 1.5, fontFamily: "'Inter', sans-serif" }}>
                {info}
              </div>
            )}
            {err && (
              <div style={{ background: "rgba(239,68,68,.06)", border: "1px solid rgba(239,68,68,.2)", borderRadius: 12, padding: "11px 16px", color: "#dc2626", fontSize: 13, lineHeight: 1.5, fontFamily: "'Inter', sans-serif" }}>
                {err}
              </div>
            )}

            {/* Primary CTA */}
            <button
              type="button"
              onClick={submit}
              disabled={loading || (mode === "register" && f.agencyMode === "create" && !!existingAgencyMatch)}
              style={{
                width: "100%",
                background: loading ? "rgba(240,165,0,.55)" : "#f0a500",
                color: "#000",
                border: "none",
                borderRadius: 999,
                padding: "15px 20px",
                fontSize: 15,
                fontWeight: 800,
                cursor: loading ? "wait" : "pointer",
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                gap: 8,
                fontFamily: "'Plus Jakarta Sans', sans-serif",
                opacity: (mode === "register" && f.agencyMode === "create" && !!existingAgencyMatch) ? 0.5 : 1,
                transition: "opacity .18s, background .18s",
                marginTop: 4,
              }}
            >
              {loading && <span style={{ animation: "spin 1s linear infinite", display: "inline-block", fontSize: 14 }}>⟳</span>}
              {submitLabel}
            </button>

            {/* Google sign-in */}
            {mode === "login" && (
              <button
                type="button"
                onClick={handleGoogleSignIn}
                style={{
                  width: "100%",
                  background: "#fff",
                  color: "#0f172a",
                  border: "1.5px solid #e2e8f0",
                  borderRadius: 999,
                  padding: "13px 20px",
                  fontSize: 14,
                  fontWeight: 600,
                  cursor: "pointer",
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "center",
                  gap: 10,
                  fontFamily: "'Plus Jakarta Sans', sans-serif",
                  transition: "background .18s",
                }}
                onMouseEnter={(e) => { e.currentTarget.style.background = "#f8fafc"; }}
                onMouseLeave={(e) => { e.currentTarget.style.background = "#fff"; }}
              >
                <GoogleIcon />
                Sign In with Google
              </button>
            )}

            {/* Footer row */}
            <div style={{ marginTop: "auto", borderTop: "1px solid #f1f5f9", paddingTop: 20, display: "flex", justifyContent: "space-between", alignItems: "center", gap: 12 }}>
              <span style={{ color: "#94a3b8", fontSize: 13, fontFamily: "'Inter', sans-serif" }}>
                {mode === "login" ? "Don't have an account? " : mode === "recovery" ? "Back to " : "Already have an account? "}
                <button
                  type="button"
                  onClick={() => switchMode(mode === "login" ? "register" : "login")}
                  style={{ background: "none", border: "none", color: "#f0a500", fontWeight: 700, fontSize: 13, cursor: "pointer", padding: 0, fontFamily: "'Inter', sans-serif" }}
                >
                  {mode === "login" ? "Sign up" : "Sign In"}
                </button>
              </span>
              <span style={{ color: "#c0c8d8", fontSize: 13, fontFamily: "'Inter', sans-serif", whiteSpace: "nowrap" }}>Terms &amp; Privacy</span>
            </div>
          </div>
        </div>

        {/* ── RIGHT PANEL ── */}
        <RightPanel mode={mode} />
      </div>
    </div>
  );
};

export default AuthPage;
