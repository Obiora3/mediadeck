import React from "react";
import Modal from "./Modal";
import Btn from "./Btn";

export default function Confirm({ msg, onYes, onNo, danger = true }) {
  return (
    <Modal title="Confirm Action" onClose={onNo} width={380}>
      <p style={{ color: "var(--text2)", marginBottom: 22, lineHeight: 1.6 }}>
        {msg}
      </p>

      <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
        <Btn variant="ghost" onClick={onNo}>
          Cancel
        </Btn>

        <Btn variant={danger ? "danger" : "success"} onClick={onYes}>
          Confirm
        </Btn>
      </div>
    </Modal>
  );
}