import React, { useEffect } from "react";

export default function Toast({ msg, type, onDone }) {
  useEffect(() => {
    const t = setTimeout(() => {
      if (typeof onDone === "function") onDone();
    }, 3500);
    return () => clearTimeout(t);
  }, [onDone]);

  const bg =
    type === "error"
      ? "#ef4444"
      : type === "success"
      ? "#22c55e"
      : "var(--bg4)";

  const color =
    type === "error" || type === "success" ? "#fff" : "var(--text)";

  return (
    <div
      style={{
        position: "fixed",
        bottom: 22,
        right: 22,
        zIndex: 9999,
        background: bg,
        color,
        padding: "11px 20px",
        borderRadius: 11,
        fontWeight: 600,
        fontSize: 13,
        boxShadow: "0 8px 28px rgba(0,0,0,.4)",
        animation: "fadeIn .3s ease",
        maxWidth: 320,
      }}
    >
      {msg}
    </div>
  );
}
