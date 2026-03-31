import React from "react";
import Card from "./Card";

export default function StatCard({
  label,
  value,
  sub,
  color = "var(--accent)",
  icon,
}) {
  return (
    <Card hoverable style={{ position: "relative", overflow: "hidden", minWidth: 0 }}>
      <div
        style={{
          position: "absolute",
          top: -16,
          right: -12,
          fontSize: 72,
          opacity: 0.04,
          pointerEvents: "none",
        }}
      >
        {icon}
      </div>

      <div
        style={{
          fontSize: 11,
          color: "var(--text3)",
          fontWeight: 600,
          textTransform: "uppercase",
          letterSpacing: ".08em",
          marginBottom: 7,
        }}
      >
        {label}
      </div>

      <div
        style={{
          fontSize: "clamp(22px, 3vw, 30px)",
          lineHeight: 1.05,
          fontWeight: 800,
          fontFamily: "'Syne',sans-serif",
          color,
          minWidth: 0,
          overflowWrap: "anywhere",
          wordBreak: "break-word",
        }}
      >
        {value}
      </div>

      {sub ? (
        <div
          style={{
            fontSize: 12,
            color: "var(--text2)",
            marginTop: 6,
            overflowWrap: "anywhere",
          }}
        >
          {sub}
        </div>
      ) : null}
    </Card>
  );
}