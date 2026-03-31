import React from "react";

export default function Badge({ children, color = "accent" }) {
  const map = {
    accent: ["rgba(240,165,0,.15)", "var(--accent)"],
    green: ["rgba(34,197,94,.15)", "var(--green)"],
    blue: ["rgba(59,126,245,.15)", "var(--blue)"],
    red: ["rgba(239,68,68,.15)", "var(--red)"],
    purple: ["rgba(139,92,246,.15)", "var(--purple)"],
    teal: ["rgba(20,184,166,.15)", "var(--teal)"],
    orange: ["rgba(249,115,22,.15)", "var(--orange)"],
  };

  const [bg, fg] = map[color] || map.accent;

  return (
    <span
      style={{
        background: bg,
        color: fg,
        padding: "3px 10px",
        borderRadius: 99,
        fontSize: 10,
        fontWeight: 600,
        fontFamily: "'Syne',sans-serif",
        whiteSpace: "normal",
        maxWidth: "100%",
        textAlign: "center",
        lineHeight: 1.25,
      }}
    >
      {children}
    </span>
  );
}
