import React from "react";

export default function Card({ children, style, glow, hoverable = true }) {
  return (
    <div
      style={{
        background: "var(--bg2)",
        border: `1px solid ${glow ? "rgba(240,165,0,.25)" : "var(--border)"}`,
        borderRadius: 14,
        padding: 22,
        boxShadow: glow ? "0 0 28px rgba(240,165,0,.07)" : "0 2px 12px rgba(0,0,0,.3)",
        transition: hoverable ? "transform .16s ease, box-shadow .16s ease, border-color .16s ease" : "none",
        ...style,
      }}
      onMouseEnter={(e) => {
        if (!hoverable) return;
        e.currentTarget.style.transform = "translateY(-2px)";
        e.currentTarget.style.boxShadow = glow
          ? "0 10px 28px rgba(240,165,0,.14)"
          : "0 10px 24px rgba(0,0,0,.14)";
        e.currentTarget.style.borderColor = "rgba(240,165,0,.22)";
      }}
      onMouseLeave={(e) => {
        if (!hoverable) return;
        e.currentTarget.style.transform = "translateY(0)";
        e.currentTarget.style.boxShadow = glow
          ? "0 0 28px rgba(240,165,0,.07)"
          : "0 2px 12px rgba(0,0,0,.3)";
        e.currentTarget.style.borderColor = glow ? "rgba(240,165,0,.25)" : "var(--border)";
      }}
    >
      {children}
    </div>
  );
}
