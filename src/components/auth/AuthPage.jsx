import { useEffect, useMemo, useState } from "react";
import { supabase } from "../../lib/supabase";
import { normalizeAgencyCode, findExistingAgencyByName } from "../../services/agencies";
import { Card, Field, Btn } from "../ui/primitives";

/* ── AUTH ───────────────────────────────────────────────── */
const AUTH_BG = "#f8fafc";

const AuthPage = ({ onLogin }) => {
  const [mode, setMode] = useState("login");
  const [f, setF] = useState({
    name: "",
    email: "",
    password: "",
    agency: "",
    agencyCode: "",
    agencyMode: "create",
    title: "",
    phone: "",
    confirm: "",
  });
  const [err, setErr] = useState("");
  const [info, setInfo] = useState("");
  const [loading, setLoading] = useState(false);
  const [resetLoading, setResetLoading] = useState(false);
  const [existingAgencyMatch, setExistingAgencyMatch] = useState(null);
  const [agencyCheckLoading, setAgencyCheckLoading] = useState(false);

  const isRecoveryMode = mode === "recovery";
  const passwordLabel = useMemo(() => (isRecoveryMode ? "New Password" : "Password"), [isRecoveryMode]);

  const u = (k) => (v) => {
    setErr("");
    setInfo("");
    setF((p) => ({ ...p, [k]: v }));
  };

  useEffect(() => {
    const previousHtmlBg = document.documentElement.style.backgroundColor;
    const previousBodyBg = document.body.style.backgroundColor;
    document.documentElement.style.backgroundColor = AUTH_BG;
    document.body.style.backgroundColor = AUTH_BG;

    return () => {
      document.documentElement.style.backgroundColor = previousHtmlBg;
      document.body.style.backgroundColor = previousBodyBg;
    };
  }, []);

  useEffect(() => {
    const hash = window.location.hash || "";
    const params = new URLSearchParams(hash.startsWith("#") ? hash.slice(1) : hash);
    const type = params.get("type");
    if (type === "recovery") {
      setMode("recovery");
      setErr("");
      setInfo("Enter your new password to complete the reset.");
    }

    const { data: { subscription } } = supabase.auth.onAuthStateChange((event) => {
      if (event === "PASSWORD_RECOVERY") {
        setMode("recovery");
        setErr("");
        setInfo("Enter your new password to complete the reset.");
      }
    });

    return () => {
      subscription.unsubscribe();
    };
  }, []);

  useEffect(() => {
    let active = true;
    const checkAgency = async () => {
      if (mode !== "register" || f.agencyMode !== "create" || !f.agency.trim()) {
        if (active) {
          setExistingAgencyMatch(null);
          setAgencyCheckLoading(false);
        }
        return;
      }

      setAgencyCheckLoading(true);
      try {
        const match = await findExistingAgencyByName(f.agency.trim());
        if (active) setExistingAgencyMatch(match || null);
      } catch (error) {
        console.error("Failed to validate agency name during signup:", error);
        if (active) setExistingAgencyMatch(null);
      } finally {
        if (active) setAgencyCheckLoading(false);
      }
    };

    const timer = setTimeout(checkAgency, 250);
    return () => {
      active = false;
      clearTimeout(timer);
    };
  }, [mode, f.agencyMode, f.agency]);

  const handleForgotPassword = async () => {
    setErr("");
    setInfo("");

    if (!f.email.trim()) {
      setErr("Enter your email address first, then click Forgot password.");
      return;
    }

    setResetLoading(true);

    try {
      const redirectTo = `${window.location.origin}${window.location.pathname}`;
      const { error } = await supabase.auth.resetPasswordForEmail(f.email.trim(), { redirectTo });
      if (error) throw error;

      setInfo("Password reset email sent. Check your inbox and open the link to set a new password.");
    } catch (e) {
      setErr(e.message || "Failed to send password reset email.");
    } finally {
      setResetLoading(false);
    }
  };

  const submit = async () => {
    setErr("");
    setInfo("");

    if (mode === "register") {
      if (!f.name || !f.email || !f.password) {
        return setErr("Full name, email, and password are required.");
      }
      if (f.agencyMode === "create" && !f.agency.trim()) {
        return setErr("Agency name is required when creating a new agency.");
      }
      if (f.agencyMode === "create" && existingAgencyMatch) {
        return setErr("Agency already existing contact admin.");
      }
      if (f.agencyMode === "join" && !normalizeAgencyCode(f.agencyCode)) {
        return setErr("Agency invite code is required when joining an existing agency.");
      }
      if (f.password !== f.confirm) {
        return setErr("Passwords do not match.");
      }
      if (f.password.length < 6) {
        return setErr("Password must be at least 6 characters.");
      }
    }

    if (mode === "recovery") {
      if (!f.password) {
        return setErr("New password is required.");
      }
      if (f.password.length < 6) {
        return setErr("Password must be at least 6 characters.");
      }
      if (f.password !== f.confirm) {
        return setErr("Passwords do not match.");
      }
    }

    setLoading(true);

    try {
      if (mode === "login") {
        const { error } = await supabase.auth.signInWithPassword({
          email: f.email,
          password: f.password,
        });

        if (error) throw error;
      } else if (mode === "recovery") {
        const { error } = await supabase.auth.updateUser({
          password: f.password,
        });

        if (error) throw error;

        setInfo("Password updated successfully. Sign in with your new password.");
        setMode("login");
        setF((p) => ({ ...p, password: "", confirm: "" }));
        window.history.replaceState({}, document.title, window.location.pathname + window.location.search);
      } else {
        const { data, error } = await supabase.auth.signUp({
          email: f.email,
          password: f.password,
          options: {
            data: {
              full_name: f.name,
              title: f.title || "",
              phone: f.phone || "",
              agency_name: f.agencyMode === "create" ? f.agency.trim() : "",
              agency_code: f.agencyMode === "join" ? normalizeAgencyCode(f.agencyCode) : "",
              agency_mode: f.agencyMode,
            },
          },
        });

        if (error) throw error;

        if (!data.user) {
          throw new Error("Signup failed. No user was returned.");
        }

        const { error: profileError } = await supabase
          .from("profiles")
          .update({
            full_name: f.name,
            title: f.title || "",
            phone: f.phone || "",
          })
          .eq("id", data.user.id);

        if (profileError) throw profileError;

        if (!data.session) {
          setInfo("Account created. Check your email to confirm your account, then sign in to finish joining your agency.");
          setMode("login");
        }
      }
    } catch (e) {
      setErr(e.message || "Something went wrong.");
    } finally {
      setLoading(false);
    }
  };

  const switchMode = (nextMode) => {
    setMode(nextMode);
    setErr("");
    setInfo("");
    if (nextMode !== "register") {
      setExistingAgencyMatch(null);
      setAgencyCheckLoading(false);
    }
    if (nextMode === "login") {
      setF((p) => ({ ...p, password: "", confirm: "" }));
    }
  };

  const lightAuthThemeVars = {
    "--bg": AUTH_BG,
    "--bg2": "#ffffff",
    "--bg3": "#f8fafc",
    "--text": "#0f172a",
    "--text2": "#475569",
    "--text3": "#64748b",
    "--border": "rgba(15,23,42,.10)",
    "--border2": "rgba(15,23,42,.14)",
    "--accent": "#d97706",
    "--red": "#dc2626",
    "--green": "#16a34a",
    "--shadow": "0 20px 50px rgba(15,23,42,.10)",
  };

  return (
    <div
      style={{
        ...lightAuthThemeVars,
        minHeight: "100vh",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        background: AUTH_BG,
        padding: 20,
        color: "var(--text)",
      }}
    >
      <div style={{ width: "100%", maxWidth: 420 }}>
        <div style={{ textAlign: "center", marginBottom: 36 }}>
          <div
            style={{
              display: "inline-flex",
              alignItems: "center",
              gap: 11,
              background: "#ffffff",
              border: "1px solid rgba(15,23,42,.10)",
              borderRadius: 14,
              padding: "11px 18px",
              marginBottom: 22,
              boxShadow: "0 10px 30px rgba(15,23,42,.06)",
            }}
          >
            <div
              style={{
                width: 34,
                height: 34,
                background: "var(--accent)",
                borderRadius: 9,
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                fontSize: 17,
              }}
            >
              📡
            </div>
            <div>
              <div
                style={{
                  fontFamily: "'Syne',sans-serif",
                  fontWeight: 800,
                  fontSize: 17,
                }}
              >
                MediaDesk Pro
              </div>
              <div style={{ fontSize: 10, color: "var(--text3)" }}>
                MEDIA SCHEDULE PLATFORM
              </div>
            </div>
          </div>

          <h1
            style={{
              fontFamily: "'Syne',sans-serif",
              fontWeight: 800,
              fontSize: 26,
              letterSpacing: "-.03em",
            }}
          >
            {mode === "login" ? "Welcome back" : mode === "recovery" ? "Reset password" : "Create account"}
          </h1>
          <p style={{ color: "var(--text2)", marginTop: 7, fontSize: 14 }}>
            {mode === "login"
              ? "Sign in to your agency workspace"
              : mode === "recovery"
                ? "Set a new password for your account"
                : "Create an account and either create a brand-new agency or join an existing one"}
          </p>
        </div>

        <div
          style={{
            background: "#ffffff",
            border: "1px solid rgba(15,23,42,.08)",
            borderRadius: 22,
            padding: 20,
            boxShadow: "0 24px 60px rgba(15,23,42,.10)",
          }}
        >
          <Card style={{ background: "transparent", boxShadow: "none", border: "none", padding: 0 }}>
            <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
              {mode === "register" && (
                <>
                  <Field
                    label="Full Name"
                    value={f.name}
                    onChange={u("name")}
                    placeholder="Jane Okafor"
                    required
                  />
                  <Field
                    label="Job Title"
                    value={f.title}
                    onChange={u("title")}
                    placeholder="Media Buyer / Account Executive"
                  />
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
                    <button
                      type="button"
                      onClick={() => {
                        setErr("");
                        setInfo("");
                        setExistingAgencyMatch(null);
                        setF((p) => ({ ...p, agencyMode: "create", agencyCode: p.agencyCode || "" }));
                      }}
                      style={{ padding: "11px 12px", borderRadius: 10, border: f.agencyMode === "create" ? "1px solid var(--accent)" : "1px solid var(--border)", background: f.agencyMode === "create" ? "rgba(217,119,6,.10)" : "#ffffff", color: f.agencyMode === "create" ? "var(--accent)" : "var(--text2)", cursor: "pointer", fontWeight: 600 }}
                    >
                      Create New Agency
                    </button>
                    <button
                      type="button"
                      onClick={() => {
                        setErr("");
                        setInfo("");
                        setF((p) => ({ ...p, agencyMode: "join", agency: p.agency || "" }));
                      }}
                      style={{ padding: "11px 12px", borderRadius: 10, border: f.agencyMode === "join" ? "1px solid var(--accent)" : "1px solid var(--border)", background: f.agencyMode === "join" ? "rgba(217,119,6,.10)" : "#ffffff", color: f.agencyMode === "join" ? "var(--accent)" : "var(--text2)", cursor: "pointer", fontWeight: 600 }}
                    >
                      Join With Invite Code
                    </button>
                  </div>
                  {f.agencyMode === "create" ? (
                    <>
                      <Field
                        label="Agency Name"
                        value={f.agency}
                        onChange={(value) => {
                          setErr("");
                          setInfo("");
                          u("agency")(value);
                        }}
                        placeholder="Apex Media Ltd"
                        required
                        note={agencyCheckLoading ? "Checking whether this agency already exists..." : "Only use Create New Agency for a brand-new workspace."}
                      />
                      {existingAgencyMatch ? (
                        <div style={{ background: "rgba(239,68,68,.08)", border: "1px solid rgba(239,68,68,.28)", borderRadius: 10, padding: "11px 13px", display: "flex", flexDirection: "column", gap: 8 }}>
                          <div style={{ fontSize: 12, color: "var(--red)", fontWeight: 700 }}>This agency already exists</div>
                          <div style={{ fontSize: 12, color: "var(--text2)", lineHeight: 1.5 }}>
                            <strong style={{ color: "var(--text)" }}>{existingAgencyMatch.name}</strong> already has a workspace. Agency already existing contact admin.
                          </div>
                        </div>
                      ) : null}
                    </>
                  ) : (
                    <Field
                      label="Agency Invite Code"
                      value={f.agencyCode}
                      onChange={(value) => {
                        setErr("");
                        setInfo("");
                        u("agencyCode")(value);
                      }}
                      placeholder="QVT-7K4P"
                      required
                      note="Ask your agency admin for the invite code."
                    />
                  )}
                  <Field
                    label="Phone Number"
                    type="tel"
                    value={f.phone}
                    onChange={u("phone")}
                    placeholder="+234 800 000 0000"
                  />
                </>
              )}

              <Field
                label="Email"
                type="email"
                value={f.email}
                onChange={u("email")}
                placeholder="you@agency.com"
                required
              />

              <Field
                label={passwordLabel}
                type="password"
                value={f.password}
                onChange={u("password")}
                placeholder="••••••••"
                required
              />

              {(mode === "register" || mode === "recovery") && (
                <Field
                  label="Confirm Password"
                  type="password"
                  value={f.confirm}
                  onChange={u("confirm")}
                  placeholder="••••••••"
                />
              )}

              {mode === "login" && (
                <div style={{ display: "flex", justifyContent: "flex-end", marginTop: -4 }}>
                  <button
                    type="button"
                    onClick={handleForgotPassword}
                    disabled={resetLoading || loading}
                    style={{
                      background: "none",
                      border: "none",
                      color: "var(--accent)",
                      fontWeight: 600,
                      cursor: resetLoading || loading ? "wait" : "pointer",
                      fontSize: 12,
                      padding: 0,
                    }}
                  >
                    {resetLoading ? "Sending reset link..." : "Forgot password?"}
                  </button>
                </div>
              )}

              {info && (
                <div
                  style={{
                    background: "rgba(34,197,94,.1)",
                    border: "1px solid rgba(34,197,94,.3)",
                    borderRadius: 8,
                    padding: "9px 13px",
                    color: "#15803d",
                    fontSize: 12,
                  }}
                >
                  {info}
                </div>
              )}

              {err && (
                <div
                  style={{
                    background: "rgba(239,68,68,.1)",
                    border: "1px solid rgba(239,68,68,.3)",
                    borderRadius: 8,
                    padding: "9px 13px",
                    color: "var(--red)",
                    fontSize: 12,
                  }}
                >
                  {err}
                </div>
              )}

              <Btn
                size="lg"
                onClick={submit}
                loading={loading || (mode === "register" && f.agencyMode === "create" && agencyCheckLoading)}
                disabled={mode === "register" && f.agencyMode === "create" && !!existingAgencyMatch}
                style={{ width: "100%", justifyContent: "center", marginTop: 4 }}
              >
                {mode === "login"
                  ? "Sign In →"
                  : mode === "recovery"
                    ? "Update Password →"
                    : (existingAgencyMatch && f.agencyMode === "create" ? "Already Existing Contact Admin" : "Create Account →")}
              </Btn>

              <p style={{ textAlign: "center", color: "var(--text2)", fontSize: 13 }}>
                {mode === "login"
                  ? "No account? "
                  : mode === "recovery"
                    ? "Back to "
                    : "Have an account? "}
                <button
                  onClick={() => {
                    if (mode === "login") {
                      switchMode("register");
                    } else {
                      switchMode("login");
                    }
                  }}
                  style={{
                    background: "none",
                    border: "none",
                    color: "var(--accent)",
                    fontWeight: 600,
                    cursor: "pointer",
                    fontSize: 13,
                  }}
                >
                  {mode === "login" ? "Register" : "Sign In"}
                </button>
              </p>
            </div>
          </Card>
        </div>
      </div>
    </div>
  );
};

export default AuthPage;
