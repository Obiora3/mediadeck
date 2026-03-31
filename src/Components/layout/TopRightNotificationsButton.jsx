const TopRightNotificationsButton = ({ count = 0, onClick }) => (
  <button
    onClick={onClick}
    title="Workspace alerts"
    style={{
      position: "fixed",
      top: 18,
      right: 22,
      zIndex: 80,
      width: 46,
      height: 46,
      borderRadius: 999,
      border: "1px solid var(--border2)",
      background: "var(--bg2)",
      color: "var(--text)",
      boxShadow: "0 12px 28px rgba(0,0,0,.18)",
      cursor: "pointer",
      display: "flex",
      alignItems: "center",
      justifyContent: "center",
      fontSize: 19,
      transition: "transform .16s ease, box-shadow .16s ease",
    }}
    onMouseEnter={e => { e.currentTarget.style.transform = "translateY(-2px)"; e.currentTarget.style.boxShadow = "0 16px 34px rgba(0,0,0,.22)"; }}
    onMouseLeave={e => { e.currentTarget.style.transform = "translateY(0)"; e.currentTarget.style.boxShadow = "0 12px 28px rgba(0,0,0,.18)"; }}
  >
    🔔
    {count > 0 && (
      <span style={{
        position: "absolute",
        top: -4,
        right: -4,
        minWidth: 20,
        height: 20,
        borderRadius: 999,
        padding: "0 6px",
        background: "var(--accent)",
        color: "#111",
        display: "inline-flex",
        alignItems: "center",
        justifyContent: "center",
        fontSize: 10,
        fontWeight: 800,
        fontFamily: "'Syne',sans-serif",
      }}>
        {count > 99 ? "99+" : count}
      </span>
    )}
  </button>
);


export default TopRightNotificationsButton;
