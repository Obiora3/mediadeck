import { Btn } from "./ui/primitives";

export default function PwaInstallPrompt({ show, onInstall, onDismiss, isInstalling }) {
  if (!show) return null;

  return (
    <div
      role="dialog"
      aria-label="Install MediaDeck"
      className="fade"
      style={{
        position: "fixed",
        left: 22,
        bottom: 22,
        zIndex: 9998,
        width: "min(380px, calc(100vw - 44px))",
        background: "var(--bg2)",
        color: "var(--text)",
        border: "1px solid var(--border2)",
        borderRadius: 12,
        boxShadow: "0 18px 48px rgba(0,0,0,.28)",
        padding: 16,
      }}
    >
      <div style={{ display: "flex", alignItems: "flex-start", gap: 12 }}>
        <div
          aria-hidden="true"
          style={{
            width: 42,
            height: 42,
            borderRadius: 10,
            background: "linear-gradient(135deg, var(--accent), var(--blue))",
            display: "grid",
            placeItems: "center",
            color: "#07101f",
            fontFamily: "var(--font-heading)",
            fontWeight: 900,
            fontSize: 15,
            flexShrink: 0,
          }}
        >
          MD
        </div>
        <div style={{ minWidth: 0, flex: 1 }}>
          <div style={{ fontFamily: "var(--font-heading)", fontSize: 14, fontWeight: 800 }}>
            Install MediaDeck
          </div>
          <div style={{ marginTop: 4, color: "var(--text2)", fontSize: 12, lineHeight: 1.45 }}>
            Add MediaDeck to this device for quicker access in a standalone app window.
          </div>
        </div>
      </div>
      <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginTop: 14 }}>
        <Btn variant="ghost" size="sm" onClick={onDismiss}>
          Later
        </Btn>
        <Btn size="sm" onClick={onInstall} loading={isInstalling}>
          Install
        </Btn>
      </div>
    </div>
  );
}
