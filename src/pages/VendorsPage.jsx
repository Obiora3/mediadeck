import { useState } from "react";
import Badge from "../components/Badge";
import Empty from "../components/Empty";
import Toast from "../components/Toast";
import Modal from "../components/Modal";
import Confirm from "../components/Confirm";
import { activeOnly, archivedOnly, isArchived } from "../utils/records";
import { fmtN } from "../utils/formatters";
import { hasPermission, readOnlyMessage, formatRoleLabel, isAdmin, adminOnlyMessage } from "../constants/roles";
import { createVendorInSupabase, updateVendorInSupabase, archiveVendorInSupabase, restoreVendorInSupabase, deleteVendorInSupabase } from "../services/vendors";
import { Card, Field, Btn } from "../components/ui/primitives";

/* ── VENDORS ────────────────────────────────────────────── */
const VendorsPage = ({ vendors, setVendors, user }) => {
  const [modal, setModal] = useState(null);
  const [search, setSearch] = useState("");
  const [viewMode, setViewMode] = useState("active");
  const [toast, setToast] = useState(null);
  const [confirm, setConfirm] = useState(null);
  const blank = { name: "", type: "", contact: "", email: "", phone: "", location: "", rate: "", discount: "", commission: "", notes: "" };
  const [f, setF] = useState(blank);
  const u = k => v => setF(p => ({ ...p, [k]: v }));
  const mediaTypes = ["Television","Radio","Print","Digital/Online","Out-of-Home (OOH)","Cinema","Podcast","Social Media"];
  const canManage = hasPermission(user, "manageVendors");
  const canDelete = isAdmin(user);
  const save = async () => {
    if (!canManage) return setToast({ msg: readOnlyMessage(user), type: "error" });
  if (!f.name || !f.type) {
    return setToast({ msg: "Name and type required.", type: "error" });
  }

  try {
    if (modal === "add") {
      const newVendor = await createVendorInSupabase(user.agencyId, user.id, f);
      setVendors(v => [newVendor, ...v]);
    } else {
      const updatedVendor = await updateVendorInSupabase(modal.id, f);
      setVendors(v => v.map(x => x.id === modal.id ? updatedVendor : x));
    }

    setToast({
      msg: modal === "add" ? "Vendor added." : "Vendor updated.",
      type: "success",
    });
    setModal(null);
    setF(blank);
  } catch (e) {
    setToast({
      msg: e.message || "Failed to save vendor.",
      type: "error",
    });
  }
};

const del = async (id) => {
  if (!canManage) return setToast({ msg: readOnlyMessage(user), type: "error" });
  try {
    const archivedVendor = await archiveVendorInSupabase(id);
    setVendors(v => v.map(x => x.id === id ? archivedVendor : x));
    setToast({ msg: "Vendor archived.", type: "success" });
    setConfirm(null);
  } catch (e) {
    setToast({
      msg: e.message || "Failed to archive vendor.",
      type: "error",
    });
  }
};
const restore = async (id) => {
  if (!canManage) return setToast({ msg: readOnlyMessage(user), type: "error" });
  try {
    const restoredVendor = await restoreVendorInSupabase(id);
    setVendors(v => v.map(x => x.id === id ? restoredVendor : x));
    setToast({ msg: "Vendor restored.", type: "success" });
  } catch (e) {
    setToast({ msg: e.message || "Failed to restore vendor.", type: "error" });
  }
};
const hardDelete = async (id, name) => {
  if (!canDelete) return setToast({ msg: adminOnlyMessage(user), type: "error" });
  try {
    await deleteVendorInSupabase(id);
    setVendors(v => v.filter(x => x.id !== id));
    setToast({ msg: `Vendor "${name || "record"}" deleted permanently.`, type: "success" });
    setConfirm(null);
  } catch (e) {
    setToast({ msg: e.message || "Failed to delete vendor.", type: "error" });
  }
};
  const visible = viewMode === "archived" ? archivedOnly(vendors) : viewMode === "all" ? vendors : activeOnly(vendors);
  const filtered = visible.filter(v => `${v.name} ${v.type}`.toLowerCase().includes(search.toLowerCase()));
  const typeIcon = t => ({ Television: "📺", Radio: "📻", Print: "📰", "Digital/Online": "💻", "Out-of-Home (OOH)": "🪧", Cinema: "🎬", Podcast: "🎙", "Social Media": "📱" }[t] || "📡");
  const typeColor = t => ({ Television: "accent", Radio: "blue", Print: "green", "Digital/Online": "purple", "Out-of-Home (OOH)": "teal", Cinema: "red", Podcast: "orange", "Social Media": "purple" }[t] || "accent");
  return (
    <div className="fade">
      {toast && <Toast msg={toast.msg} type={toast.type} onDone={() => setToast(null)} />}
      {confirm && <Confirm msg={confirm.msg} onYes={confirm.onYes} onNo={() => setConfirm(null)} />}
      {!canManage && <Card style={{ marginBottom: 14, background: "rgba(59,126,245,.06)", border: "1px solid rgba(59,126,245,.18)" }}><div style={{ fontSize: 13, color: "var(--text2)" }}>You have read-only access to Clients as <strong>{formatRoleLabel(user?.role)}</strong>.</div></Card>}
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 24, flexWrap: "wrap", gap: 10 }}>
        <div><h1 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 800, fontSize: 24 }}>Vendors</h1><p style={{ color: "var(--text2)", marginTop: 3, fontSize: 13 }}>Media houses & suppliers</p></div>
        {canManage && <Btn icon="+" onClick={() => { setF(blank); setModal("add"); }}>Add Vendor</Btn>}
      </div>
      <div style={{ display: "flex", gap: 10, marginBottom: 18, flexWrap: "wrap" }}><div style={{ flex: 1, minWidth: 180 }}><Field value={search} onChange={setSearch} placeholder="Search vendors…" /></div><Field value={viewMode} onChange={setViewMode} options={[{value:"active",label:"Active"},{value:"archived",label:"Archived"},{value:"all",label:"All"}]} /></div>
      {filtered.length === 0 ? <Card><Empty icon="🏢" title="No vendors found" sub="Add your first media house" /></Card> :
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill,minmax(300px,1fr))", gap: 14 }}>
          {filtered.map(v => (
            <Card key={v.id}>
              <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", marginBottom: 12 }}>
                <div style={{ display: "flex", alignItems: "center", gap: 11 }}>
                  <div style={{ width: 42, height: 42, background: "var(--bg3)", borderRadius: 11, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 20 }}>{typeIcon(v.type)}</div>
                  <div><div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 14 }}>{v.name}</div><div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginTop: 3 }}><Badge color={typeColor(v.type)}>{v.type}</Badge>{isArchived(v) && <Badge color="red">Archived</Badge>}</div></div>
                </div>
                <div style={{ display: "flex", gap: 5 }}>
                  {canManage && <Btn variant="ghost" size="sm" onClick={() => { setF({ name: v.name, type: v.type, contact: v.contact||"", email: v.email||"", phone: v.phone||"", location: v.location||"", rate: v.rate||"", discount: v.discount||"", commission: v.commission||"", notes: v.notes||"" }); setModal(v); }}>✏️</Btn>}
                  {canManage && (isArchived(v) ? <Btn variant="success" size="sm" onClick={() => restore(v.id)}>↩</Btn> : <Btn variant="danger" size="sm" onClick={() => setConfirm({ msg: `Archive "${v.name}"? Existing reports and MPO links will stay intact.`, onYes: () => del(v.id) })}>🗄</Btn>)}
                  {canDelete && <Btn variant="danger" size="sm" onClick={() => setConfirm({ msg: `Delete "${v.name}" permanently? This cannot be undone.`, onYes: () => hardDelete(v.id, v.name) })}>🗑</Btn>}
                </div>
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 7 }}>
                {[["Rate/Spot", v.rate ? fmtN(v.rate) : "—"], ["Vol Disc.", v.discount ? `${v.discount}%` : "—"], ["Comm.", v.commission ? `${v.commission}%` : "—"], ["Contact", v.contact || "—"]].map(([l, val]) => (
                  <div key={l} style={{ background: "var(--bg3)", borderRadius: 7, padding: "7px 11px" }}>
                    <div style={{ fontSize: 9, color: "var(--text3)", fontWeight: 600, textTransform: "uppercase", letterSpacing: ".07em" }}>{l}</div>
                    <div style={{ fontSize: 12, fontWeight: 600, marginTop: 2, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{val}</div>
                  </div>
                ))}
              </div>
              {v.notes && <div style={{ marginTop: 10, fontSize: 11, color: "var(--text2)", borderTop: "1px solid var(--border)", paddingTop: 9 }}>{v.notes}</div>}
            </Card>
          ))}
        </div>}
      {modal !== null && (
        <Modal title={modal === "add" ? "Add Vendor" : `Edit: ${modal.name}`} onClose={() => setModal(null)} width={560}>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 }}>
            <div style={{ gridColumn: "1/-1" }}><Field label="Vendor Name" value={f.name} onChange={u("name")} placeholder="NTA, Channels TV…" required /></div>
            <Field label="Media Type" value={f.type} onChange={u("type")} options={mediaTypes} placeholder="Select type" required />
            <Field label="Contact Person" value={f.contact} onChange={u("contact")} placeholder="Name" />
            <Field label="Email" type="email" value={f.email} onChange={u("email")} placeholder="vendor@example.com" />
            <Field label="Phone" value={f.phone} onChange={u("phone")} placeholder="+234 800 000 0000" />
            <div style={{ gridColumn: "1/-1" }}><Field label="Location / City" value={f.location} onChange={u("location")} placeholder="Lagos, Abuja, Port Harcourt…" note="State or city where vendor operates" /></div>
            <Field label="Default Rate/Spot (₦)" type="number" value={f.rate} onChange={u("rate")} placeholder="0" />
            <Field label="Volume Discount (%)" type="number" value={f.discount} onChange={u("discount")} placeholder="e.g. 27" />
            <div style={{ gridColumn: "1/-1" }}><Field label="Agency Commission (%)" type="number" value={f.commission} onChange={u("commission")} placeholder="e.g. 15" /></div>
            <div style={{ gridColumn: "1/-1" }}><Field label="Notes" value={f.notes} onChange={u("notes")} placeholder="Additional notes…" /></div>
          </div>
          <div style={{ display: "flex", gap: 10, justifyContent: "flex-end", marginTop: 20 }}>
            <Btn variant="ghost" onClick={() => setModal(null)}>Cancel</Btn>
            <Btn onClick={save}>{modal === "add" ? "Add Vendor" : "Save Changes"}</Btn>
          </div>
        </Modal>
      )}
    </div>
  );
};


export default VendorsPage;
