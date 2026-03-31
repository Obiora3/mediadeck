import React from "react";

export default function Modal({ title, children, onClose, width = 540 }) {
  return (
    <div
      style={{
        position: "fixed",
        inset: 0,
        background: "rgba(0,0,0,.72)",
        zIndex: 1000,
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        padding: 20,
        backdropFilter: "blur(5px)",
      }}
      onClick={(e) => e.target === e.currentTarget && onClose()}
    >
      <div
        className="fade"
        style={{
          background: "var(--bg2)",
          border: "1px solid var(--border2)",
          borderRadius: 18,
          width: "100%",
          maxWidth: width,
          maxHeight: "92vh",
          overflowY: "auto",
          boxShadow: "0 28px 80px rgba(0,0,0,.65)",
        }}
      >
        <div
          style={{
            display: "flex",
            alignItems: "center",
            justifyContent: "space-between",
            padding: "18px 22px",
            borderBottom: "1px solid var(--border)",
            position: "sticky",
            top: 0,
            background: "var(--bg2)",
            zIndex: 1,
          }}
        >
          <h3
            style={{
              margin: 0,
              fontFamily: "'Syne',sans-serif",
              fontSize: 18,
              fontWeight: 700,
            }}
          >
            {title}
          </h3>
          <button
            type="button"
            onClick={onClose}
            style={{
              border: "none",
              background: "transparent",
              color: "var(--text2)",
              fontSize: 22,
              cursor: "pointer",
              lineHeight: 1,
            }}
            aria-label="Close modal"
          >
            ×
          </button>
        </div>
        <div style={{ padding: 22 }}>{children}</div>
      </div>
    </div>
  );
}
