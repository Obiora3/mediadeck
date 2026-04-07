import { useEffect, useMemo, useState } from "react";
import { supabase } from "../../lib/supabase";
import { normalizeAgencyCode, findExistingAgencyByName } from "../../services/agencies";
import { Field, Btn } from "../ui/primitives";

/* ── AUTH ───────────────────────────────────────────────── */
const AUTH_BG = "#f5f7fb";


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
  const isRegisterMode = mode === "register";
  const passwordLabel = useMemo(() => (isRecoveryMode ? "New Password" : "Password"), [isRecoveryMode]);
  const heading = useMemo(() => {
    if (mode === "login") return "Welcome back";
    if (mode === "recovery") return "Create a new password";
    return "Create your workspace account";
  }, [mode]);
  const subheading = useMemo(() => {
    if (mode === "login") return "Sign in to continue managing schedules, reports, and agency operations.";
    if (mode === "recovery") return "Set a new password to regain secure access to your workspace.";
    return "Register your profile, then create a new agency or join an existing one with an invite code.";
  }, [mode]);

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

    const {
      data: { subscription },
    } = supabase.auth.onAuthStateChange((event) => {
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
        if (typeof onLogin === "function") onLogin();
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
    "--bg3": "#eef2ff",
    "--text": "#0f172a",
    "--text2": "#475569",
    "--text3": "#64748b",
    "--border": "rgba(15,23,42,.10)",
    "--border2": "rgba(15,23,42,.14)",
    "--accent": "#d97706",
    "--accent2": "#7c3aed",
    "--red": "#dc2626",
    "--green": "#16a34a",
    "--shadow": "0 20px 50px rgba(15,23,42,.10)",
  };

  return (
    <div
      style={{
        ...lightAuthThemeVars,
        minHeight: "100vh",
        width: "100%",
        background: AUTH_BG,
        color: "var(--text)",
        padding: "clamp(72px, 10vw, 112px) clamp(24px, 5vw, 40px) clamp(96px, 12vw, 140px)",
        boxSizing: "border-box",
        display: "flex",
        alignItems: "flex-start",
        justifyContent: "center",
      }}
    >
      <section
        style={{
          width: "100%",
          maxWidth: 468,
          margin: "0 auto",
          paddingBottom: 24,
        }}
      >
        <div
          style={{
            width: "100%",
            borderRadius: 28,
            background: "#ffffff",
            border: "1px solid rgba(15,23,42,.08)",
            boxShadow: "0 25px 70px rgba(15,23,42,.10)",
            padding: "clamp(28px, 3.5vw, 36px)",
          }}
        >
          <div
            style={{
              display: "flex",
              flexDirection: "column",
              alignItems: "center",
              textAlign: "center",
              marginBottom: 22,
            }}
          >
            <div
              style={{
                display: "inline-flex",
                alignItems: "center",
                gap: 16,
                padding: "16px 22px",
                borderRadius: 22,
                background: "#ffffff",
                border: "1px solid rgba(15,23,42,.08)",
                boxShadow: "0 12px 30px rgba(15,23,42,.05)",
                marginBottom: 18,
                width: "100%",
                maxWidth: 440,
                justifyContent: "flex-start",
              }}
            >
              <div
                style={{
                  width: 56,
                  height: 56,
                  borderRadius: 14,
                  background: "#d97706",
                  color: "#ffffff",
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "center",
                  fontSize: 26,
                  lineHeight: 1,
                  flexShrink: 0,
                  boxShadow: "inset 0 1px 0 rgba(255,255,255,.18)",
                }}
              >
                📡
              </div>
              <div style={{ textAlign: "left", minWidth: 0 }}>
                <div
                  style={{
                    fontFamily: "'Syne',sans-serif",
                    fontWeight: 800,
                    fontSize: 18,
                    letterSpacing: "-.03em",
                    lineHeight: 1.05,
                    color: "#0f172a",
                  }}
                >
                  MediaDesk Pro
                </div>
                <div
                  style={{
                    marginTop: 8,
                    fontSize: 11,
                    fontWeight: 700,
                    letterSpacing: ".08em",
                    textTransform: "uppercase",
                    color: "#64748b",
                  }}
                >
                  Media Schedule Platform
                </div>
              </div>
            </div>

            <div
              style={{
                display: "inline-flex",
                alignItems: "center",
                justifyContent: "center",
                padding: "7px 12px",
                borderRadius: 999,
                background: "rgba(124,58,237,.08)",
                color: "#6d28d9",
                fontSize: 11,
                fontWeight: 800,
                letterSpacing: ".08em",
                textTransform: "uppercase",
                marginBottom: 14,
              }}
            >
              Secure access
            </div>

            <h2
              style={{
                margin: 0,
                maxWidth: 320,
                fontFamily: "'Syne',sans-serif",
                fontWeight: 800,
                fontSize: 20,
                lineHeight: 1.15,
                letterSpacing: "-.03em",
                textAlign: "center",
              }}
            >
              {heading}
            </h2>
            <p
              style={{
                margin: "10px 0 0",
                maxWidth: 340,
                color: "var(--text2)",
                fontSize: 14,
                lineHeight: 1.6,
                textAlign: "center",
              }}
            >
              {subheading}
            </p>
          </div>

          <div
            style={{
              display: "grid",
              gridTemplateColumns: "repeat(2, minmax(0, 1fr))",
              gap: 10,
              marginBottom: 18,
            }}
          >
            {[
              { key: "login", label: "Sign In" },
              { key: "register", label: "Register" },
            ].map((item) => {
              const active = mode === item.key;
              return (
                <button
                  key={item.key}
                  type="button"
                  onClick={() => switchMode(item.key)}
                  style={{
                    border: active ? "1px solid rgba(124,58,237,.32)" : "1px solid rgba(15,23,42,.08)",
                    background: active ? "linear-gradient(135deg, rgba(124,58,237,.10), rgba(217,119,6,.08))" : "rgba(248,250,252,.85)",
                    color: active ? "#5b21b6" : "var(--text2)",
                    borderRadius: 14,
                    padding: "12px 10px",
                    fontWeight: 700,
                    cursor: "pointer",
                    transition: "all .2s ease",
                  }}
                >
                  {item.label}
                </button>
              );
            })}
          </div>

          <div style={{ display: "flex", flexDirection: "column", gap: 14, paddingInline: 2 }}>
            {isRegisterMode && (
              <>
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
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
                    placeholder="Media Buyer"
                  />
                </div>

                <div
                  style={{
                    display: "grid",
                    gridTemplateColumns: "1fr 1fr",
                    gap: 10,
                    padding: 8,
                    borderRadius: 16,
                    background: "rgba(15,23,42,.03)",
                    border: "1px solid rgba(15,23,42,.06)",
                  }}
                >
                  <button
                    type="button"
                    onClick={() => {
                      setErr("");
                      setInfo("");
                      setExistingAgencyMatch(null);
                      setF((p) => ({ ...p, agencyMode: "create", agencyCode: p.agencyCode || "" }));
                    }}
                    style={{
                      padding: "13px 12px",
                      borderRadius: 14,
                      border: f.agencyMode === "create" ? "1px solid rgba(217,119,6,.28)" : "1px solid transparent",
                      background: f.agencyMode === "create" ? "rgba(217,119,6,.10)" : "transparent",
                      color: f.agencyMode === "create" ? "#b45309" : "var(--text2)",
                      cursor: "pointer",
                      fontWeight: 700,
                    }}
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
                    style={{
                      padding: "13px 12px",
                      borderRadius: 14,
                      border: f.agencyMode === "join" ? "1px solid rgba(124,58,237,.24)" : "1px solid transparent",
                      background: f.agencyMode === "join" ? "rgba(124,58,237,.08)" : "transparent",
                      color: f.agencyMode === "join" ? "#6d28d9" : "var(--text2)",
                      cursor: "pointer",
                      fontWeight: 700,
                    }}
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
                      note={agencyCheckLoading ? "Checking whether this agency already exists..." : "Only use this when creating a brand-new workspace."}
                    />
                    {existingAgencyMatch ? (
                      <div
                        style={{
                          background: "rgba(239,68,68,.07)",
                          border: "1px solid rgba(239,68,68,.24)",
                          borderRadius: 16,
                          padding: "13px 14px",
                          display: "flex",
                          flexDirection: "column",
                          gap: 8,
                        }}
                      >
                        <div style={{ fontSize: 11, color: "var(--red)", fontWeight: 800, textTransform: "uppercase", letterSpacing: ".05em" }}>
                          Agency already exists
                        </div>
                        <div style={{ fontSize: 13, color: "var(--text2)", lineHeight: 1.6 }}>
                          <strong style={{ color: "var(--text)" }}>{existingAgencyMatch.name}</strong> already has a workspace. Contact your admin for access instead of creating another agency.
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

            {(isRegisterMode || isRecoveryMode) && (
              <Field
                label="Confirm Password"
                type="password"
                value={f.confirm}
                onChange={u("confirm")}
                placeholder="••••••••"
              />
            )}

            {mode === "recovery" && (
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10, marginTop: -2 }}>
                <div style={{ fontSize: 12, color: "var(--text3)" }}>Open the reset link from your email, then set your new password here.</div>
                <button
                  type="button"
                  onClick={() => switchMode("login")}
                  style={{
                    background: "none",
                    border: "none",
                    color: "#7c3aed",
                    fontWeight: 700,
                    cursor: "pointer",
                    fontSize: 12,
                    padding: 0,
                  }}
                >
                  Back to sign in
                </button>
              </div>
            )}

            {mode === "login" && (
              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, marginTop: -2 }}>
                <div style={{ fontSize: 11, color: "var(--text3)" }}>Use the email linked to your agency workspace.</div>
                <button
                  type="button"
                  onClick={handleForgotPassword}
                  disabled={resetLoading || loading}
                  style={{
                    background: "none",
                    border: "none",
                    color: "#7c3aed",
                    fontWeight: 700,
                    cursor: resetLoading || loading ? "wait" : "pointer",
                    fontSize: 11,
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
                  background: "rgba(34,197,94,.10)",
                  border: "1px solid rgba(34,197,94,.24)",
                  borderRadius: 16,
                  padding: "12px 14px",
                  color: "#166534",
                  fontSize: 13,
                  lineHeight: 1.5,
                }}
              >
                {info}
              </div>
            )}

            {err && (
              <div
                style={{
                  background: "rgba(239,68,68,.08)",
                  border: "1px solid rgba(239,68,68,.24)",
                  borderRadius: 16,
                  padding: "12px 14px",
                  color: "var(--red)",
                  fontSize: 13,
                  lineHeight: 1.5,
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
              style={{
                width: "100%",
                justifyContent: "center",
                marginTop: 6,
                minHeight: 52,
                borderRadius: 16,
                fontWeight: 800,
                letterSpacing: ".01em",
              }}
            >
              {mode === "login"
                ? "Sign In"
                : mode === "recovery"
                  ? "Update Password"
                  : existingAgencyMatch && f.agencyMode === "create"
                    ? "Already Existing Contact Admin"
                    : "Create Account"}
            </Btn>

            <div
              style={{
                display: "grid",
                gridTemplateColumns: "1fr auto 1fr",
                alignItems: "center",
                gap: 12,
                marginTop: 4,
                color: "var(--text3)",
                fontSize: 11,
              }}
            >
              <span style={{ height: 1, background: "rgba(15,23,42,.08)" }} />
              <span>Quick switch</span>
              <span style={{ height: 1, background: "rgba(15,23,42,.08)" }} />
            </div>

            <p style={{ textAlign: "center", color: "var(--text2)", fontSize: 13, margin: 0 }}>
              {mode === "login" ? "No account yet? " : mode === "recovery" ? "Back to " : "Already have an account? "}
              <button
                type="button"
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
                  color: "#7c3aed",
                  fontWeight: 800,
                  cursor: "pointer",
                  fontSize: 13,
                  padding: 0,
                }}
              >
                {mode === "login" ? "Register" : "Sign In"}
              </button>
            </p>
          </div>
        </div>
      </section>
    </div>
  );
};

export default AuthPage;
