import { formatRoleLabel } from "../../constants/roles";

/* ── SIDEBAR ────────────────────────────────────────────── */
const Sidebar = ({ page, setPage, user, onLogout, collapsed, setCollapsed, theme, toggleTheme, unreadNotifications = 0, canInstall = false, onInstall }) => {
  const nav = [
    { id: "dashboard", icon: "◈", label: "Dashboard" },
    { id: "vendors",   icon: "🏢", label: "Vendors" },
    { id: "clients",   icon: "👥", label: "Clients & Brands" },
    { id: "campaigns", icon: "📢", label: "Campaigns" },
    { id: "rates",     icon: "💰", label: "Media Rates" },
    { id: "finance",   icon: "💳", label: "Finance" },
    { id: "mpo",       icon: "📄", label: "MPO Generator" },
    { id: "reports",   icon: "📊", label: "Reports" },
    { id: "settings",  icon: "⚙️", label: "Settings", badge: unreadNotifications },
  ];
  return (
    <div style={{ width: collapsed ? 64 : 230, minHeight: "100vh", background: "var(--bg2)", borderRight: "1px solid var(--border)", display: "flex", flexDirection: "column", transition: "width .22s ease", overflow: "hidden", flexShrink: 0, position: "sticky", top: 0, height: "100vh" }}>
      <div style={{ padding: collapsed ? "18px 14px" : "18px 18px", borderBottom: "1px solid var(--border)", display: "flex", alignItems: "center", gap: 11, minHeight: 68 }}>
        <div style={{ width: 34, height: 34, background: "var(--accent)", borderRadius: 9, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 17, flexShrink: 0 }}>📡</div>
        {!collapsed && <div><div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 800, fontSize: 15 }}>MediaDesk</div><div style={{ fontSize: 9, color: "var(--text3)" }}>PRO PLATFORM · {formatRoleLabel(user?.role)}</div></div>}
      </div>
      <nav style={{ flex: 1, padding: "10px 6px", display: "flex", flexDirection: "column", gap: 2 }}>
        {nav.map(n => {
          const active = page === n.id;
          return (
            <button key={n.id} onClick={() => setPage(n.id)}
              style={{ display: "flex", alignItems: "center", gap: 11, padding: collapsed ? "9px 10px" : "9px 13px", borderRadius: 9, border: "none", background: active ? "rgba(240,165,0,.12)" : "transparent", color: active ? "var(--accent)" : "var(--text2)", cursor: "pointer", fontSize: 13, fontWeight: active ? 600 : 400, textAlign: "left", transition: "all .14s", whiteSpace: "nowrap", justifyContent: collapsed ? "center" : "flex-start", borderLeft: active ? "2px solid var(--accent)" : "2px solid transparent" }}
              onMouseEnter={e => { if (!active) e.currentTarget.style.background = "var(--bg3)"; }}
              onMouseLeave={e => { if (!active) e.currentTarget.style.background = "transparent"; }}>
              <span style={{ fontSize: 17, flexShrink: 0 }}>{n.icon}</span>
              {!collapsed && <><span>{n.label}</span>{n.badge ? <span style={{ marginLeft: "auto", background: "var(--accent)", color: "#111", minWidth: 18, height: 18, borderRadius: 999, display: "inline-flex", alignItems: "center", justifyContent: "center", fontSize: 10, fontWeight: 800, padding: "0 6px" }}>{n.badge > 99 ? "99+" : n.badge}</span> : null}</>}
            </button>
          );
        })}
      </nav>

      {canInstall && (
        <button onClick={onInstall}
          title="Install MediaDeck app"
          style={{ margin: "4px 6px 2px", padding: "7px", background: "rgba(240,165,0,.08)", border: "1px solid rgba(240,165,0,.25)", borderRadius: 7, color: "var(--accent)", cursor: "pointer", fontSize: 14, display: "flex", alignItems: "center", justifyContent: collapsed ? "center" : "flex-start", gap: 8, width: "calc(100% - 12px)" }}>
          ⬇
          {!collapsed && <span style={{ fontSize: 11, fontWeight: 600 }}>Install App</span>}
        </button>
      )}
      <button onClick={() => toggleTheme && toggleTheme()}
        title={theme === "light" ? "Switch to Dark Mode" : "Switch to Light Mode"}
        style={{ margin: "4px 6px 2px", padding: "7px", background: theme === "dark" ? "rgba(240,165,0,.12)" : "rgba(59,126,245,.1)", border: `1px solid ${theme === "dark" ? "rgba(240,165,0,.3)" : "rgba(59,126,245,.25)"}`, borderRadius: 7, color: theme === "dark" ? "var(--accent)" : "var(--blue)", cursor: "pointer", fontSize: 14, display: "flex", alignItems: "center", justifyContent: collapsed ? "center" : "flex-start", gap: 8, width: "calc(100% - 12px)" }}>
        {theme === "light" ? "🌙" : "☀️"}
        {!collapsed && <span style={{ fontSize: 11, fontWeight: 600 }}>{theme === "light" ? "Dark Mode" : "Light Mode"}</span>}
      </button>
      <button onClick={() => setCollapsed(!collapsed)} style={{ margin: "4px 6px 6px", padding: "7px", background: "var(--bg3)", border: "1px solid var(--border)", borderRadius: 7, color: "var(--text3)", cursor: "pointer", fontSize: 12, display: "flex", alignItems: "center", justifyContent: "center" }}>{collapsed ? "▶" : "◀"}</button>

      {/* User strip */}
      <div style={{ padding: "10px 10px 14px", borderTop: "1px solid var(--border)", display: "flex", alignItems: "center", gap: 9 }}>
        <div onClick={() => setPage("settings")} style={{ width: 32, height: 32, background: "linear-gradient(135deg,var(--accent),var(--purple))", borderRadius: "50%", display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 700, fontSize: 13, flexShrink: 0, color: "#000", cursor: "pointer" }} title="My Profile">{user.name?.[0]?.toUpperCase() || "U"}</div>
        {!collapsed && <><div style={{ flex: 1, overflow: "hidden" }}><div style={{ fontSize: 12, fontWeight: 600, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{user.name}</div><div style={{ fontSize: 10, color: "var(--text3)", overflow: "hidden", textOverflow: "ellipsis" }}>{user.title ? user.title : user.agency}</div></div><button onClick={onLogout} title="Sign Out" style={{ background: "none", border: "none", color: "var(--text3)", cursor: "pointer", fontSize: 15, flexShrink: 0 }}>⎋</button></>}
      </div>
    </div>
  );
};



export default Sidebar;
