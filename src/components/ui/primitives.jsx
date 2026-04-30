/* ── STYLES ─────────────────────────────────────────────── */
const GlobalStyle = ({ theme }) => (
  <style>{`
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@500;600;700;800&family=Inter:wght@400;500;600;700;800&display=swap');
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
      --font-body:'Inter',sans-serif;--font-heading:'Plus Jakarta Sans',sans-serif;
    }
    html,body,#root{height:100%;width:100%;font-family:var(--font-body);background:var(--bg);color:var(--text);overflow-x:hidden}
    #root{min-height:100dvh;text-align:initial;border:0}
    h1,h2,h3,h4,h5,h6,strong,th{color:var(--text);font-family:var(--font-heading)}
    p,span,label,td,li,small{color:inherit}
    input::placeholder,textarea::placeholder{color:var(--text3);opacity:1}
    select,option,input,textarea{color:var(--text)}
    img,svg,canvas,video{max-width:100%}
    *{scrollbar-width:thin;scrollbar-color:var(--bg4) transparent}
    ::-webkit-scrollbar{width:5px}::-webkit-scrollbar-track{background:transparent}::-webkit-scrollbar-thumb{background:var(--bg4);border-radius:3px}
    input,select,textarea,button{font-family:inherit}
    input,select,textarea,button{max-width:100%}
    .app-shell{display:flex;min-height:100vh;width:100%;max-width:100%;overflow-x:hidden}
    .app-main{flex:1;min-width:0;width:100%;overflow-y:auto;overflow-x:hidden;padding:calc(env(titlebar-area-height, 0px) + 28px) 28px 52px;position:relative;box-sizing:border-box}
    .desktop-sidebar{display:flex}
    .mobile-topbar,.mobile-bottom-nav{display:none}
    .modal-backdrop{position:fixed;inset:0;background:rgba(0,0,0,.72);z-index:1000;display:flex;align-items:center;justify-content:center;padding:20px;backdrop-filter:blur(5px)}
    .modal-panel{background:var(--bg2);border:1px solid var(--border2);border-radius:18px;width:100%;max-height:92vh;overflow-y:auto;box-shadow:0 28px 80px rgba(0,0,0,.65)}
    .modal-body{padding:22px}
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
    @media (max-width: 860px){
      html,body,#root{height:auto;min-height:100dvh}
      .app-shell{display:block!important;min-height:100dvh!important;padding-bottom:78px!important;overflow-x:hidden!important}
      .app-main{width:100%!important;min-width:0!important;padding:72px 14px 92px!important;overflow:visible!important}
      .desktop-sidebar{display:none!important}
      .mobile-topbar{position:fixed;top:0;left:0;right:0;height:58px;z-index:70;display:flex;align-items:center;gap:10px;padding:9px 64px 9px 12px;background:var(--bg2);background:color-mix(in srgb,var(--bg2) 94%,transparent);border-bottom:1px solid var(--border);backdrop-filter:blur(16px);box-shadow:0 10px 30px rgba(0,0,0,.12)}
      .mobile-bottom-nav{position:fixed;left:0;right:0;bottom:0;z-index:75;display:flex;gap:4px;overflow-x:auto;padding:8px 10px calc(8px + env(safe-area-inset-bottom));background:var(--bg2);background:color-mix(in srgb,var(--bg2) 96%,transparent);border-top:1px solid var(--border);backdrop-filter:blur(16px);box-shadow:0 -14px 30px rgba(0,0,0,.14)}
      .mobile-bottom-nav::-webkit-scrollbar{display:none}
      .top-alert-button{top:9px!important;right:12px!important;width:40px!important;height:40px!important;font-size:17px!important;z-index:90!important}
      .fade{max-width:100%}
      .app-main h1{font-size:22px!important;line-height:1.15!important}
      .app-main h2,.app-main h3{line-height:1.2}
      .app-main [style*="justify-content: space-between"]{row-gap:10px}
      .app-main [style*="grid-template-columns: 1fr 1fr"],
      .app-main [style*="grid-template-columns: 1fr 340px"],
      .app-main [style*="grid-template-columns: 1fr auto"],
      .app-main [style*="grid-template-columns: 1fr 180px"],
      .app-main [style*="grid-template-columns: 1fr 2fr"],
      .app-main [style*="grid-template-columns: 2fr"],
      .app-main [style*="grid-template-columns: 2.5fr"],
      .app-main [style*="grid-template-columns: 3fr"],
      .app-main [style*="grid-template-columns: 1.2fr"],
      .app-main [style*="grid-template-columns: 1.15fr"],
      .app-main [style*="grid-template-columns: 1.12fr"],
      .app-main [style*="grid-template-columns: 1.3fr"],
      .app-main [style*="grid-template-columns: 1.4fr"],
      .app-main [style*="grid-template-columns: 1.5fr"],
      .app-main [style*="grid-template-columns: 200px"],
      .app-main [style*="grid-template-columns: auto 1fr auto"]{grid-template-columns:1fr!important}
      .app-main [style*="grid-template-columns: repeat(4"],
      .app-main [style*="grid-template-columns: repeat(5"],
      .app-main [style*="grid-template-columns: repeat(6"]{grid-template-columns:repeat(2,minmax(0,1fr))!important}
      .app-main [style*="grid-template-columns: repeat(3"]{grid-template-columns:1fr!important}
      .app-main [style*="minmax(300px"]{grid-template-columns:1fr!important}
      .app-main [style*="min-width: 220px"],
      .app-main [style*="min-width: 280px"],
      .app-main [style*="min-width: 300px"]{min-width:0!important}
      .app-main [style*="max-height: calc(100vh - 260px)"]{max-height:none!important}
      .app-main [style*="padding: 22px"]{padding:16px!important}
      .modal-backdrop{align-items:flex-end!important;padding:10px!important}
      .modal-panel{max-width:100%!important;max-height:92dvh!important;border-radius:16px 16px 0 0}
      .modal-body{padding:16px!important}
      table{font-size:12px}
    }
    @media (max-width: 520px){
      .app-main{padding-left:12px!important;padding-right:12px!important}
      .mobile-bottom-nav{gap:2px;padding-left:8px;padding-right:8px}
      .mobile-bottom-nav button{min-width:64px!important}
      .app-main [style*="grid-template-columns: repeat(2"]{grid-template-columns:1fr!important}
      .app-main [style*="gap: 20px"]{gap:12px!important}
      .app-main [style*="gap: 18px"]{gap:14px!important}
      .app-main [style*="gap: 14px"]{gap:12px!important}
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
      style={{ display: "inline-flex", alignItems: "center", gap: 7, fontFamily: "var(--font-heading)", fontWeight: 700, letterSpacing: "-0.01em", border: "none", borderRadius: 9, cursor: (disabled || loading) ? "not-allowed" : "pointer", transition: "all .18s", opacity: (disabled || loading) ? .55 : 1, outline: "none", whiteSpace: "nowrap", ...sz, ...vars[variant], ...style }}
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
        fontFamily: "var(--font-heading)",
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
