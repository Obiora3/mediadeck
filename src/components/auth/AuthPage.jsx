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

  const submit = async () => {
    setErr("");

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

    setLoading(true);

    try {
      if (mode === "login") {
        const { error } = await supabase.auth.signInWithPassword({
          email: f.email,
          password: f.password,
        });

        if (error) throw error;
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
          setErr(
            "Account created. Check your email to confirm your account, then sign in to finish joining your agency."
          );
          setMode("login");
        }
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
        position: "fixed",
        inset: 0,
        zIndex: 9999,
        overflowY: "auto",
        overflowX: "hidden",
        minHeight: "100dvh",
        boxSizing: "border-box",
        paddingTop: "max(28px, calc(env(safe-area-inset-top, 0px) + 20px))",
        paddingRight: 20,
        paddingBottom: "max(28px, calc(env(safe-area-inset-bottom, 0px) + 20px))",
        paddingLeft: 20,
        color: "var(--text)",
        background:
          "radial-gradient(ellipse 80% 50% at 50% -10%,rgba(240,165,0,.12) 0%,transparent 70%),#f8fafc",

        "--bg": "#f8fafc",
        "--bg2": "#ffffff",
        "--bg3": "#f1f5f9",
        "--text": "#0f172a",
        "--text2": "#334155",
        "--text3": "#64748b",
        "--border": "rgba(15,23,42,.10)",
        "--border2": "rgba(15,23,42,.14)",
        "--accent": "#f0a500",
        "--red": "#ef4444",
        "--green": "#16a34a",
      }}
    >
      <div style={{ width: "100%", maxWidth: 420, margin: "0 auto 24px" }}>
        <div style={{ textAlign: "center", marginBottom: 28 }}>
          <div
            style={{
              display: "inline-flex",
              alignItems: "center",
              gap: 11,
              background: "var(--bg2)",
              border: "1px solid var(--border2)",
              borderRadius: 14,
              padding: "11px 18px",
              marginBottom: 20,
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
              lineHeight: 1.15,
            }}
          >
            {mode === "login" ? "Welcome back" : "Create account"}
          </h1>
          <p style={{ color: "var(--text2)", marginTop: 7, fontSize: 14, lineHeight: 1.45 }}>
            {mode === "login"
              ? "Sign in to your agency workspace"
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
                    onClick={() => { setErr(""); setExistingAgencyMatch(null); setF(p => ({ ...p, agencyMode: "create", agencyCode: p.agencyCode || "" })); }}
                    style={{ padding: "11px 12px", borderRadius: 10, border: f.agencyMode === "create" ? "1px solid var(--accent)" : "1px solid var(--border)", background: f.agencyMode === "create" ? "rgba(240,165,0,.12)" : "var(--bg3)", color: f.agencyMode === "create" ? "var(--accent)" : "var(--text2)", cursor: "pointer", fontWeight: 600, lineHeight: 1.3 }}
                  >
                    Create New Agency
                  </button>
                  <button
                    type="button"
                    onClick={() => { setErr(""); setF(p => ({ ...p, agencyMode: "join", agency: p.agency || "" })); }}
                    style={{ padding: "11px 12px", borderRadius: 10, border: f.agencyMode === "join" ? "1px solid var(--accent)" : "1px solid var(--border)", background: f.agencyMode === "join" ? "rgba(240,165,0,.12)" : "var(--bg3)", color: f.agencyMode === "join" ? "var(--accent)" : "var(--text2)", cursor: "pointer", fontWeight: 600, lineHeight: 1.3 }}
                  >
                    Join With Invite Code
                  </button>
                </div>
                {f.agencyMode === "create" ? (
                  <>
                    <Field
                      label="Agency Name"
                      value={f.agency}
                      onChange={value => {
                        setErr("");
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
                    onChange={value => {
                      setErr("");
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
              label="Password"
              type="password"
              value={f.password}
              onChange={u("password")}
              placeholder="••••••••"
              required
            />

            {mode === "register" && (
              <Field
                label="Confirm Password"
                type="password"
                value={f.confirm}
                onChange={u("confirm")}
                placeholder="••••••••"
              />
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
                  lineHeight: 1.45,
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
              {mode === "login" ? "Sign In →" : (existingAgencyMatch && f.agencyMode === "create" ? "Already Existing Contact Admin" : "Create Account →")}
            </Btn>

            <p style={{ textAlign: "center", color: "var(--text2)", fontSize: 13, lineHeight: 1.45 }}>
              {mode === "login" ? "No account? " : "Have an account? "}
              <button
                onClick={() => {
                  setMode(mode === "login" ? "register" : "login");
                  setErr("");
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
