/* ── STYLES ─────────────────────────────────────────────── */
const GlobalStyle = ({ theme }) => (
  <style>{`
    @import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;600;700;800&family=DM+Sans:ital,wght@0,300;0,400;0,500;0,600;1,400&display=swap');
    *,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
    :root{
      ${theme === "dark" ? `
      --bg:#07090f;--bg2:#0e1118;--bg3:#141824;--bg4:#1c2233;
      --border:rgba(255,255,255,0.07);--border2:rgba(255,255,255,0.13);
      --text:#e8ecf4;--text2:#8b93a7;--text3:#4f576b;
      ` : `
      --bg:#f0f2f7;--bg2:#ffffff;--bg3:#f5f7fc;--bg4:#e8ecf4;
      --border:rgba(15,23,42,0.08);--border2:rgba(15,23,42,0.14);
      --text:#0f172a;--text2:#475467;--text3:#667085;
      `}
      --accent:#f0a500;--blue:#3b7ef5;--green:#16a34a;
      --red:#ef4444;--purple:#8b5cf6;--teal:#0d9488;--orange:#f97316;
    }
    html,body,#root{height:100%;font-family:'DM Sans',sans-serif;background:var(--bg);color:var(--text)}
    h1,h2,h3,h4,h5,h6,strong,th{color:var(--text)}
    p,span,label,td,li,small{color:inherit}
    input::placeholder,textarea::placeholder{color:var(--text3);opacity:1}
    select,option,input,textarea{color:var(--text)}
    *{scrollbar-width:thin;scrollbar-color:var(--bg4) transparent}
    ::-webkit-scrollbar{width:5px}::-webkit-scrollbar-track{background:transparent}::-webkit-scrollbar-thumb{background:var(--bg4);border-radius:3px}
    input,select,textarea,button{font-family:inherit}
    .fade{animation:fadeIn .25s ease}
    @keyframes fadeIn{from{opacity:0;transform:translateY(6px)}to{opacity:1;transform:translateY(0)}}
    @keyframes spin{to{transform:rotate(360deg)}}
    .spin{animation:spin 1s linear infinite}
    textarea{resize:vertical}
    @media print{
      body{background:#fff!important;color:#000!important}
      .no-print{display:none!important}
      .print-area{display:block!important}
    }
  `}</style>
);

/* ── UI PRIMITIVES ──────────────────────────────────────── */
const Btn = ({ children, variant = "primary", size = "md", onClick, type = "button", disabled, style, icon, loading }) => {
  const sz = size === "sm" ? { padding: "5px 13px", fontSize: 12 } : size === "lg" ? { padding: "13px 28px", fontSize: 15 } : { padding: "9px 18px", fontSize: 13 };
  const vars = {
    primary:  { background: "var(--accent)", color: "#000" },
    secondary:{ background: "var(--bg4)", color: "var(--text)", border: "1px solid var(--border2)" },
    danger:   { background: "rgba(239,68,68,.15)", color: "var(--red)", border: "1px solid rgba(239,68,68,.3)" },
    ghost:    { background: "transparent", color: "var(--text2)", border: "1px solid var(--border)" },
    success:  { background: "rgba(34,197,94,.15)", color: "var(--green)", border: "1px solid rgba(34,197,94,.3)" },
    blue:     { background: "rgba(59,126,245,.15)", color: "var(--blue)", border: "1px solid rgba(59,126,245,.3)" },
    purple:   { background: "rgba(139,92,246,.15)", color: "var(--purple)", border: "1px solid rgba(139,92,246,.3)" },
  };
  return (
    <button type={type} onClick={onClick} disabled={disabled || loading}
      style={{ display: "inline-flex", alignItems: "center", gap: 7, fontFamily: "'Syne',sans-serif", fontWeight: 600, border: "none", borderRadius: 9, cursor: (disabled || loading) ? "not-allowed" : "pointer", transition: "all .18s", opacity: (disabled || loading) ? .55 : 1, outline: "none", whiteSpace: "nowrap", ...sz, ...vars[variant], ...style }}
      onMouseEnter={e => { if (!disabled && !loading) e.currentTarget.style.filter = "brightness(1.18)"; }}
      onMouseLeave={e => e.currentTarget.style.filter = ""}>
      {loading ? <span className="spin" style={{ fontSize: 14 }}>⟳</span> : icon && <span style={{ fontSize: size === "sm" ? 13 : 16 }}>{icon}</span>}
      {children}
    </button>
  );
};

const Field = ({ label, value, onChange, type = "text", placeholder, required, options, note, error, rows }) => (
  <div style={{ display: "flex", flexDirection: "column", gap: 5 }}>
    {label && <label style={{ fontSize: 11, fontWeight: 600, color: "var(--text3)", textTransform: "uppercase", letterSpacing: ".08em" }}>{label}{required && <span style={{ color: "var(--accent)", marginLeft: 3 }}>*</span>}</label>}
    {options ? (
      <select value={value} onChange={e => onChange(e.target.value)}
        style={{ background: "var(--bg3)", border: `1px solid ${error ? "var(--red)" : "var(--border2)"}`, borderRadius: 8, padding: "9px 13px", color: value ? "var(--text)" : "var(--text3)", fontSize: 13, outline: "none", cursor: "pointer" }}>
        <option value="">{placeholder || "Select…"}</option>
        {options.map(o => <option key={o.value ?? o} value={o.value ?? o}>{o.label ?? o}</option>)}
      </select>
    ) : rows ? (
      <textarea value={value} onChange={e => onChange(e.target.value)} placeholder={placeholder} rows={rows}
        style={{ background: "var(--bg3)", border: `1px solid ${error ? "var(--red)" : "var(--border2)"}`, borderRadius: 8, padding: "9px 13px", color: "var(--text)", fontSize: 13, outline: "none", lineHeight: 1.5 }}
        onFocus={e => e.target.style.borderColor = "var(--accent)"}
        onBlur={e => e.target.style.borderColor = error ? "var(--red)" : "var(--border2)"} />
    ) : (
      <input type={type} value={value} onChange={e => onChange(e.target.value)} placeholder={placeholder} required={required}
        style={{ background: "var(--bg3)", border: `1px solid ${error ? "var(--red)" : "var(--border2)"}`, borderRadius: 8, padding: "9px 13px", color: "var(--text)", fontSize: 13, outline: "none", width: "100%" }}
        onFocus={e => e.target.style.borderColor = "var(--accent)"}
        onBlur={e => e.target.style.borderColor = error ? "var(--red)" : "var(--border2)"} />
    )}
    {error && <span style={{ fontSize: 11, color: "var(--red)" }}>{error}</span>}
    {note && !error && <span style={{ fontSize: 11, color: "var(--text3)" }}>{note}</span>}
  </div>
);

const AttachmentField = ({ label, url, onUrlChange, onFileSelected, uploading, accept = ".pdf,image/*,.doc,.docx,.xls,.xlsx,.ppt,.pptx" }) => (
  <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
    <Field label={`${label} Link`} value={url} onChange={onUrlChange} placeholder="Paste a link or upload a file below" />
    <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
      <input
        type="file"
        accept={accept}
        onChange={e => {
          const file = e.target.files?.[0] || null;
          if (file) onFileSelected(file);
          e.target.value = "";
        }}
        style={{ fontSize: 12, color: "var(--text2)", maxWidth: "100%" }}
      />
      {uploading && <span style={{ fontSize: 12, color: "var(--blue)", fontWeight: 600 }}>Uploading…</span>}
      {url && (
        <a href={url} target="_blank" rel="noreferrer" style={{ fontSize: 12, color: "var(--accent)", textDecoration: "none" }}>
          Open current file
        </a>
      )}
    </div>
  </div>
);

const Card = ({ children, style, glow, hoverable = true }) => (
  <div
    style={{ background: "var(--bg2)", color: "var(--text)", border: `1px solid ${glow ? "rgba(240,165,0,.25)" : "var(--border)"}`, borderRadius: 14, padding: 22, boxShadow: glow ? "0 0 28px rgba(240,165,0,.07)" : "0 6px 18px rgba(15,23,42,.08)", transition: hoverable ? "transform .16s ease, box-shadow .16s ease, border-color .16s ease" : "none", minWidth: 0, overflowWrap: "anywhere", wordBreak: "break-word", ...style }}
    onMouseEnter={e => {
      if (!hoverable) return;
      e.currentTarget.style.transform = "translateY(-2px)";
      e.currentTarget.style.boxShadow = glow ? "0 10px 28px rgba(240,165,0,.14)" : "0 12px 28px rgba(15,23,42,.12)";
      e.currentTarget.style.borderColor = "rgba(240,165,0,.22)";
    }}
    onMouseLeave={e => {
      if (!hoverable) return;
      e.currentTarget.style.transform = "translateY(0)";
      e.currentTarget.style.boxShadow = glow ? "0 0 28px rgba(240,165,0,.07)" : "0 6px 18px rgba(15,23,42,.08)";
      e.currentTarget.style.borderColor = glow ? "rgba(240,165,0,.25)" : "var(--border)";
    }}
  >
    {children}
  </div>
);





const Stat = ({ label, value, sub, color = "var(--accent)", icon, valueSize = "clamp(18px, 2.25vw, 26px)" }) => (
  <Card hoverable style={{ position: "relative", overflow: "hidden", minWidth: 0, padding: "18px 20px" }}>
    <div style={{ position: "absolute", top: -16, right: -12, fontSize: 72, opacity: .04, pointerEvents: "none" }}>{icon}</div>
    <div style={{ fontSize: 11, color: "var(--text2)", fontWeight: 800, textTransform: "uppercase", letterSpacing: ".06em", marginBottom: 10, lineHeight: 1.2 }}>{label}</div>
    <div
      title={String(value ?? "")}
      style={{
        fontSize: valueSize,
        lineHeight: 1.08,
        fontWeight: 800,
        fontFamily: "Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif",
        fontVariantNumeric: "tabular-nums",
        letterSpacing: "-0.04em",
        color,
        minWidth: 0,
        maxWidth: "100%",
        overflow: "hidden",
        textOverflow: "ellipsis",
        whiteSpace: "nowrap",
      }}
    >
      {value}
    </div>
    {sub && <div style={{ fontSize: 11, color: "var(--text2)", opacity: 0.92, marginTop: 10, lineHeight: 1.45, overflowWrap: "anywhere" }}>{sub}</div>}
  </Card>
);




export { GlobalStyle, Btn, Field, AttachmentField, Card, Stat };
