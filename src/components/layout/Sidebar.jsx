import { formatRoleLabel } from "../../constants/roles";

const Sidebar = ({
  page,
  setPage,
  user,
  onLogout,
  collapsed,
  setCollapsed,
  theme,
  toggleTheme,
  unreadNotifications = 0,
  canInstall = false,
  onInstall,
}) => {
  const nav = [
    { id: "dashboard", icon: "◈", label: "Dashboard" },
    { id: "vendors", icon: "🏢", label: "Vendors" },
    { id: "clients", icon: "👥", label: "Clients & Brands" },
    { id: "campaigns", icon: "📢", label: "Campaigns" },
    { id: "rates", icon: "💰", label: "Media Rates" },
    { id: "finance", icon: "💳", label: "Finance" },
    { id: "mpo", icon: "📄", label: "MPO Generator" },
    { id: "reports", icon: "📊", label: "Reports" },
    { id: "settings", icon: "⚙️", label: "Settings", badge: unreadNotifications },
  ];
  const activeNav = nav.find((item) => item.id === page) || nav[0];
  const userInitial = user.name?.[0]?.toUpperCase() || "U";

  return (
    <>
      <aside
        className="desktop-sidebar"
        style={{
          width: collapsed ? 64 : 230,
          minHeight: "100vh",
          background: "var(--bg2)",
          borderRight: "1px solid var(--border)",
          flexDirection: "column",
          transition: "width .22s ease",
          overflow: "hidden",
          flexShrink: 0,
          position: "sticky",
          top: 0,
          height: "100vh",
        }}
      >
        <div
          style={{
            padding: collapsed ? "18px 14px" : "18px 18px",
            borderBottom: "1px solid var(--border)",
            display: "flex",
            alignItems: "center",
            gap: 11,
            minHeight: 68,
            WebkitAppRegion: "drag",
            appRegion: "drag",
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
              fontSize: 13,
              fontFamily: "var(--font-heading)",
              fontWeight: 900,
              color: "#111",
              flexShrink: 0,
            }}
          >
            MD
          </div>
          {!collapsed && (
            <div>
              <div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 800, fontSize: 15 }}>MediaDesk</div>
              <div style={{ fontSize: 9, color: "var(--text3)" }}>PRO PLATFORM / {formatRoleLabel(user?.role)}</div>
            </div>
          )}
        </div>

        <nav aria-label="Primary navigation" style={{ flex: 1, padding: "10px 6px", display: "flex", flexDirection: "column", gap: 2 }}>
          {nav.map((item) => {
            const active = page === item.id;
            return (
              <button
                key={item.id}
                type="button"
                onClick={() => setPage(item.id)}
                aria-current={active ? "page" : undefined}
                title={collapsed ? item.label : undefined}
                style={{
                  display: "flex",
                  alignItems: "center",
                  gap: 11,
                  padding: collapsed ? "9px 10px" : "9px 13px",
                  borderRadius: 9,
                  border: "none",
                  background: active ? "rgba(240,165,0,.12)" : "transparent",
                  color: active ? "var(--accent)" : "var(--text2)",
                  cursor: "pointer",
                  fontSize: 13,
                  fontWeight: active ? 600 : 400,
                  textAlign: "left",
                  transition: "all .14s",
                  whiteSpace: "nowrap",
                  justifyContent: collapsed ? "center" : "flex-start",
                  borderLeft: active ? "2px solid var(--accent)" : "2px solid transparent",
                }}
                onMouseEnter={(event) => {
                  if (!active) event.currentTarget.style.background = "var(--bg3)";
                }}
                onMouseLeave={(event) => {
                  if (!active) event.currentTarget.style.background = "transparent";
                }}
              >
                <span style={{ width: 18, flexShrink: 0, fontSize: 16, textAlign: "center" }}>{item.icon}</span>
                {!collapsed && (
                  <>
                    <span>{item.label}</span>
                    {item.badge ? (
                      <span
                        style={{
                          marginLeft: "auto",
                          background: "var(--accent)",
                          color: "#111",
                          minWidth: 18,
                          height: 18,
                          borderRadius: 999,
                          display: "inline-flex",
                          alignItems: "center",
                          justifyContent: "center",
                          fontSize: 10,
                          fontWeight: 800,
                          padding: "0 6px",
                        }}
                      >
                        {item.badge > 99 ? "99+" : item.badge}
                      </span>
                    ) : null}
                  </>
                )}
              </button>
            );
          })}
        </nav>

        {canInstall && (
          <button
            type="button"
            onClick={onInstall}
            title="Install MediaDeck app"
            style={{
              margin: "4px 6px 2px",
              padding: "7px",
              background: "rgba(240,165,0,.08)",
              border: "1px solid rgba(240,165,0,.25)",
              borderRadius: 7,
              color: "var(--accent)",
              cursor: "pointer",
              fontSize: 14,
              display: "flex",
              alignItems: "center",
              justifyContent: collapsed ? "center" : "flex-start",
              gap: 8,
              width: "calc(100% - 12px)",
            }}
          >
            <span style={{ fontWeight: 900 }}>+</span>
            {!collapsed && <span style={{ fontSize: 11, fontWeight: 600 }}>Install App</span>}
          </button>
        )}
        <button
          type="button"
          onClick={() => toggleTheme && toggleTheme()}
          title={theme === "light" ? "Switch to Dark Mode" : "Switch to Light Mode"}
          style={{
            margin: "4px 6px 2px",
            padding: "7px",
            background: theme === "dark" ? "rgba(240,165,0,.12)" : "rgba(59,126,245,.1)",
            border: `1px solid ${theme === "dark" ? "rgba(240,165,0,.3)" : "rgba(59,126,245,.25)"}`,
            borderRadius: 7,
            color: theme === "dark" ? "var(--accent)" : "var(--blue)",
            cursor: "pointer",
            fontSize: 14,
            display: "flex",
            alignItems: "center",
            justifyContent: collapsed ? "center" : "flex-start",
            gap: 8,
            width: "calc(100% - 12px)",
          }}
        >
          {theme === "light" ? "Dark" : "Light"}
          {!collapsed && <span style={{ fontSize: 11, fontWeight: 600 }}>{theme === "light" ? "Dark Mode" : "Light Mode"}</span>}
        </button>
        <button
          type="button"
          onClick={() => setCollapsed(!collapsed)}
          style={{
            margin: "4px 6px 6px",
            padding: "7px",
            background: "var(--bg3)",
            border: "1px solid var(--border)",
            borderRadius: 7,
            color: "var(--text3)",
            cursor: "pointer",
            fontSize: 12,
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
          }}
        >
          {collapsed ? ">" : "<"}
        </button>

        <div style={{ padding: "10px 10px 14px", borderTop: "1px solid var(--border)", display: "flex", alignItems: "center", gap: 9 }}>
          <div
            onClick={() => setPage("settings")}
            style={{
              width: 32,
              height: 32,
              background: "linear-gradient(135deg,var(--accent),var(--purple))",
              borderRadius: "50%",
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              fontWeight: 700,
              fontSize: 13,
              flexShrink: 0,
              color: "#000",
              cursor: "pointer",
            }}
            title="My Profile"
          >
            {userInitial}
          </div>
          {!collapsed && (
            <>
              <div style={{ flex: 1, overflow: "hidden" }}>
                <div style={{ fontSize: 12, fontWeight: 600, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{user.name}</div>
                <div style={{ fontSize: 10, color: "var(--text3)", overflow: "hidden", textOverflow: "ellipsis" }}>
                  {user.title ? user.title : user.agency}
                </div>
              </div>
              <button type="button" onClick={onLogout} title="Sign Out" style={{ background: "none", border: "none", color: "var(--text3)", cursor: "pointer", fontSize: 13, flexShrink: 0 }}>
                Out
              </button>
            </>
          )}
        </div>
      </aside>

      <header className="mobile-topbar">
        <button
          type="button"
          onClick={() => setPage("dashboard")}
          aria-label="Go to dashboard"
          style={{
            width: 38,
            height: 38,
            borderRadius: 10,
            border: "none",
            background: "var(--accent)",
            color: "#111",
            fontFamily: "var(--font-heading)",
            fontWeight: 900,
            cursor: "pointer",
            flexShrink: 0,
          }}
        >
          MD
        </button>
        <div style={{ minWidth: 0, flex: 1 }}>
          <div style={{ fontFamily: "var(--font-heading)", fontWeight: 800, fontSize: 14, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
            {activeNav.label}
          </div>
          <div style={{ color: "var(--text3)", fontSize: 10, fontWeight: 700, textTransform: "uppercase", letterSpacing: ".06em", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
            {formatRoleLabel(user?.role)} / {user?.agency || "Workspace"}
          </div>
        </div>
        <button
          type="button"
          onClick={() => setPage("settings")}
          title="My Profile"
          style={{
            width: 34,
            height: 34,
            borderRadius: 999,
            border: "1px solid var(--border2)",
            background: "linear-gradient(135deg,var(--accent),var(--purple))",
            color: "#000",
            display: "inline-flex",
            alignItems: "center",
            justifyContent: "center",
            fontWeight: 800,
            cursor: "pointer",
            flexShrink: 0,
          }}
        >
          {userInitial}
        </button>
      </header>

      <nav className="mobile-bottom-nav" aria-label="Mobile navigation">
        {nav.map((item) => {
          const active = page === item.id;
          return (
            <button
              key={item.id}
              type="button"
              onClick={() => setPage(item.id)}
              aria-current={active ? "page" : undefined}
              style={{
                minWidth: 74,
                minHeight: 54,
                border: `1px solid ${active ? "rgba(240,165,0,.35)" : "var(--border)"}`,
                borderRadius: 12,
                background: active ? "rgba(240,165,0,.12)" : "var(--bg3)",
                color: active ? "var(--accent)" : "var(--text2)",
                display: "flex",
                flexDirection: "column",
                alignItems: "center",
                justifyContent: "center",
                gap: 4,
                fontSize: 10,
                fontWeight: active ? 800 : 700,
                cursor: "pointer",
                position: "relative",
              }}
            >
              <span style={{ fontSize: 15, lineHeight: 1 }}>{item.icon}</span>
              <span style={{ maxWidth: "100%", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", padding: "0 4px" }}>{item.label.replace(" & Brands", "")}</span>
              {item.badge ? (
                <span
                  style={{
                    position: "absolute",
                    top: 4,
                    right: 5,
                    minWidth: 16,
                    height: 16,
                    borderRadius: 999,
                    background: "var(--accent)",
                    color: "#111",
                    fontSize: 9,
                    fontWeight: 900,
                    display: "inline-flex",
                    alignItems: "center",
                    justifyContent: "center",
                    padding: "0 4px",
                  }}
                >
                  {item.badge > 99 ? "99+" : item.badge}
                </span>
              ) : null}
            </button>
          );
        })}
        <button
          type="button"
          onClick={() => toggleTheme && toggleTheme()}
          style={{
            minWidth: 74,
            minHeight: 54,
            border: "1px solid var(--border)",
            borderRadius: 12,
            background: "var(--bg3)",
            color: "var(--text2)",
            display: "flex",
            flexDirection: "column",
            alignItems: "center",
            justifyContent: "center",
            gap: 4,
            fontSize: 10,
            fontWeight: 700,
            cursor: "pointer",
          }}
        >
          <span style={{ fontSize: 15, lineHeight: 1 }}>{theme === "light" ? "🌙" : "☀️"}</span>
          <span>{theme === "light" ? "Dark" : "Light"}</span>
        </button>
        {canInstall && (
          <button
            type="button"
            onClick={onInstall}
            style={{
              minWidth: 74,
              minHeight: 54,
              border: "1px solid rgba(240,165,0,.25)",
              borderRadius: 12,
              background: "rgba(240,165,0,.08)",
              color: "var(--accent)",
              display: "flex",
              flexDirection: "column",
              alignItems: "center",
              justifyContent: "center",
              gap: 4,
              fontSize: 10,
              fontWeight: 800,
              cursor: "pointer",
            }}
          >
            <span style={{ fontFamily: "var(--font-heading)", fontSize: 15, fontWeight: 900 }}>+</span>
            <span>Install</span>
          </button>
        )}
      </nav>
    </>
  );
};

export default Sidebar;
