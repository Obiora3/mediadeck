import React from "react";

export default function Field({
  label,
  value,
  onChange,
  type = "text",
  placeholder,
  required,
  options,
  note,
  error,
  rows,
}) {
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 5 }}>
      {label && (
        <label
          style={{
            fontSize: 11,
            fontWeight: 600,
            color: "var(--text3)",
            textTransform: "uppercase",
            letterSpacing: ".08em",
          }}
        >
          {label}
          {required && <span style={{ color: "var(--accent)", marginLeft: 3 }}>*</span>}
        </label>
      )}

      {options ? (
        <select
          value={value}
          onChange={(e) => onChange(e.target.value)}
          style={{
            background: "var(--bg3)",
            border: `1px solid ${error ? "var(--red)" : "var(--border2)"}`,
            borderRadius: 8,
            padding: "9px 13px",
            color: value ? "var(--text)" : "var(--text3)",
            fontSize: 13,
            outline: "none",
            cursor: "pointer",
          }}
        >
          <option value="">{placeholder || "Select…"}</option>
          {options.map((o) => (
            <option key={o.value ?? o} value={o.value ?? o}>
              {o.label ?? o}
            </option>
          ))}
        </select>
      ) : rows ? (
        <textarea
          value={value}
          onChange={(e) => onChange(e.target.value)}
          placeholder={placeholder}
          rows={rows}
          style={{
            background: "var(--bg3)",
            border: `1px solid ${error ? "var(--red)" : "var(--border2)"}`,
            borderRadius: 8,
            padding: "9px 13px",
            color: "var(--text)",
            fontSize: 13,
            outline: "none",
            lineHeight: 1.5,
          }}
          onFocus={(e) => (e.target.style.borderColor = "var(--accent)")}
          onBlur={(e) =>
            (e.target.style.borderColor = error ? "var(--red)" : "var(--border2)")
          }
        />
      ) : (
        <input
          type={type}
          value={value}
          onChange={(e) => onChange(e.target.value)}
          placeholder={placeholder}
          required={required}
          style={{
            background: "var(--bg3)",
            border: `1px solid ${error ? "var(--red)" : "var(--border2)"}`,
            borderRadius: 8,
            padding: "9px 13px",
            color: "var(--text)",
            fontSize: 13,
            outline: "none",
            width: "100%",
          }}
          onFocus={(e) => (e.target.style.borderColor = "var(--accent)")}
          onBlur={(e) =>
            (e.target.style.borderColor = error ? "var(--red)" : "var(--border2)")
          }
        />
      )}

      {error && <span style={{ fontSize: 11, color: "var(--red)" }}>{error}</span>}
      {note && !error && <span style={{ fontSize: 11, color: "var(--text3)" }}>{note}</span>}
    </div>
  );
}
