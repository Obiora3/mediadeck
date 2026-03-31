import { useEffect, useState } from "react";
import { supabase } from "../../lib/supabase";
import { normalizeAgencyCode, findExistingAgencyByName } from "../../services/agencies";
import { Card, Field, Btn } from "../ui/primitives";

/* ── AUTH ───────────────────────────────────────────────── */
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
  const [existingAgencyMatch, setExistingAgencyMatch] = useState(null);
  const [agencyCheckLoading, setAgencyCheckLoading] = useState(false);

  const u = (k) => (v) => setF((p) => ({ ...p, [k]: v }));

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

  const switchMode = (nextMode) => {
    setMode(nextMode);
    setErr("");
    setInfo("");
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

    if (mode === "forgot") {
      if (!f.email) {
        return setErr("Email is required.");
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
        // App-level auth listener will load the user and agency.
      } else if (mode === "forgot") {
        const redirectTo = `${window.location.origin}${window.location.pathname}`;
        const { error } = await supabase.auth.resetPasswordForEmail(f.email, {
          redirectTo,
        });

        if (error) throw error;

        setInfo("Password reset link sent. Check your email to continue.");
        setMode("login");
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
          setInfo(
            "Account created. Check your email to confirm your account, then sign in to finish joining your agency."
          );
          setMode("login");
        }
        // If a session exists immediately, App-level auth listener will complete agency join/load.
      }
    } catch (e) {
      setErr(e.message || "Something went wrong.");
    } finally {
      setLoading(false);
    }
  };

  return (
    <div
      style={{
        minHeight: "100vh",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        background:
          "radial-gradient(ellipse 80% 50% at 50% -10%,rgba(240,165,0,.09) 0%,transparent 70%),var(--bg)",
        padding: 20,
      }}
    >
      <div style={{ width: "100%", maxWidth: 420 }}>
        <div style={{ textAlign: "center", marginBottom: 36 }}>
          <div
            style={{
              display: "inline-flex",
              alignItems: "center",
              gap: 11,
              background: "var(--bg2)",
              border: "1px solid var(--border2)",
              borderRadius: 14,
              padding: "11px 18px",
              marginBottom: 22,
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
            {mode === "login"
              ? "Welcome back"
              : mode === "forgot"
                ? "Reset password"
                : "Create account"}
          </h1>
          <p style={{ color: "var(--text2)", marginTop: 7, fontSize: 14 }}>
            {mode === "login"
              ? "Sign in to your agency workspace"
              : mode === "forgot"
                ? "Enter your email and we will send you a password reset link"
                : "Create an account and either create a brand-new agency or join an existing one"}
          </p>
        </div>

        <Card>
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
                    style={{ padding: "11px 12px", borderRadius: 10, border: f.agencyMode === "create" ? "1px solid var(--accent)" : "1px solid var(--border)", background: f.agencyMode === "create" ? "rgba(240,165,0,.12)" : "var(--bg3)", color: f.agencyMode === "create" ? "var(--accent)" : "var(--text2)", cursor: "pointer", fontWeight: 600 }}
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
                    style={{ padding: "11px 12px", borderRadius: 10, border: f.agencyMode === "join" ? "1px solid var(--accent)" : "1px solid var(--border)", background: f.agencyMode === "join" ? "rgba(240,165,0,.12)" : "var(--bg3)", color: f.agencyMode === "join" ? "var(--accent)" : "var(--text2)", cursor: "pointer", fontWeight: 600 }}
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
              onChange={(value) => {
                setErr("");
                setInfo("");
                u("email")(value);
              }}
              placeholder="you@agency.com"
              required
            />

            {mode !== "forgot" && (
              <Field
                label="Password"
                type="password"
                value={f.password}
                onChange={(value) => {
                  setErr("");
                  setInfo("");
                  u("password")(value);
                }}
                placeholder="••••••••"
                required
              />
            )}

            {mode === "register" && (
              <Field
                label="Confirm Password"
                type="password"
                value={f.confirm}
                onChange={(value) => {
                  setErr("");
                  setInfo("");
                  u("confirm")(value);
                }}
                placeholder="••••••••"
              />
            )}

            {info && (
              <div
                style={{
                  background: "rgba(34,197,94,.1)",
                  border: "1px solid rgba(34,197,94,.25)",
                  borderRadius: 8,
                  padding: "9px 13px",
                  color: "var(--text)",
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
                : mode === "forgot"
                  ? "Send Reset Link →"
                  : (existingAgencyMatch && f.agencyMode === "create" ? "Already Existing Contact Admin" : "Create Account →")}
            </Btn>

            {mode === "login" && (
              <div style={{ display: "flex", justifyContent: "flex-end", marginTop: -2 }}>
                <button
                  type="button"
                  onClick={() => switchMode("forgot")}
                  style={{
                    background: "none",
                    border: "none",
                    color: "var(--accent)",
                    fontWeight: 600,
                    cursor: "pointer",
                    fontSize: 13,
                    padding: 0,
                  }}
                >
                  Forgot password?
                </button>
              </div>
            )}

            <p style={{ textAlign: "center", color: "var(--text2)", fontSize: 13 }}>
              {mode === "login"
                ? "No account? "
                : mode === "forgot"
                  ? "Remembered your password? "
                  : "Have an account? "}
              <button
                onClick={() => {
                  switchMode(mode === "login" ? "register" : "login");
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
  );
};

export default AuthPage;
