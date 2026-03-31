import React from "react";

export default function Btn({
  children,
  variant = "primary",
  size = "md",
  onClick,
  type = "button",
  disabled,
  style,
  icon,
  loading,
}) {
  const sz =
    size === "sm"
      ? { padding: "5px 13px", fontSize: 12 }
      : size === "lg"
      ? { padding: "13px 28px", fontSize: 15 }
      : { padding: "9px 18px", fontSize: 13 };

  const vars = {
    primary: { background: "var(--accent)", color: "#000" },
    secondary: {
      background: "var(--bg4)",
      color: "var(--text)",
      border: "1px solid var(--border2)",
    },
    danger: {
      background: "rgba(239,68,68,.15)",
      color: "var(--red)",
      border: "1px solid rgba(239,68,68,.3)",
    },
    ghost: {
      background: "transparent",
      color: "var(--text2)",
      border: "1px solid var(--border)",
    },
    success: {
      background: "rgba(34,197,94,.15)",
      color: "var(--green)",
      border: "1px solid rgba(34,197,94,.3)",
    },
    blue: {
      background: "rgba(59,126,245,.15)",
      color: "var(--blue)",
      border: "1px solid rgba(59,126,245,.3)",
    },
    purple: {
      background: "rgba(139,92,246,.15)",
      color: "var(--purple)",
      border: "1px solid rgba(139,92,246,.3)",
    },
  };

  return (
    <button
      type={type}
      onClick={onClick}
      disabled={disabled || loading}
      style={{
        display: "inline-flex",
        alignItems: "center",
        gap: 7,
        fontFamily: "'Syne',sans-serif",
        fontWeight: 600,
        border: "none",
        borderRadius: 9,
        cursor: disabled || loading ? "not-allowed" : "pointer",
        transition: "all .18s",
        opacity: disabled || loading ? 0.55 : 1,
        outline: "none",
        whiteSpace: "nowrap",
        ...sz,
        ...vars[variant],
        ...style,
      }}
      onMouseEnter={(e) => {
        if (!disabled && !loading) e.currentTarget.style.filter = "brightness(1.18)";
      }}
      onMouseLeave={(e) => {
        e.currentTarget.style.filter = "";
      }}
    >
      {loading ? (
        <span className="spin" style={{ fontSize: 14 }}>
          ⟳
        </span>
      ) : (
        icon && <span style={{ fontSize: size === "sm" ? 13 : 16 }}>{icon}</span>
      )}
      {children}
    </button>
  );
}
