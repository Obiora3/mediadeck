import React from "react";

export default function Empty({
  icon = "📭",
  title = "Nothing here yet",
  sub = "",
}) {
  return (
    <div
      style={{
        padding: 28,
        border: "1px dashed var(--border2)",
        borderRadius: 16,
        textAlign: "center",
        background: "var(--bg3)",
      }}
    >
      <div style={{ fontSize: 30, marginBottom: 10 }}>{icon}</div>
      <div
        style={{
          fontFamily: "'Syne',sans-serif",
          fontWeight: 700,
          fontSize: 16,
          marginBottom: 6,
        }}
      >
        {title}
      </div>
      {sub ? (
        <div
          style={{
            color: "var(--text2)",
            fontSize: 13,
            maxWidth: 520,
            margin: "0 auto",
            lineHeight: 1.5,
          }}
        >
          {sub}
        </div>
      ) : null}
    </div>
  );
}
