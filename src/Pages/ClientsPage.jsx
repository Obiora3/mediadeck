import { useState } from "react";
import Badge from "../components/Badge";
import Empty from "../components/Empty";
import Toast from "../components/Toast";
import Modal from "../components/Modal";
import Confirm from "../components/Confirm";
import { activeOnly, archivedOnly, isArchived } from "../utils/records";
import { hasPermission, readOnlyMessage, formatRoleLabel } from "../constants/roles";
import { createClientInSupabase, updateClientInSupabase, archiveClientInSupabase, restoreClientInSupabase } from "../services/clients";
import { createAuditEventInSupabase } from "../services/notifications";
import { Card, Field, Btn } from "../components/ui/primitives";

/* ── CLIENTS ────────────────────────────────────────────── */
const ClientsPage = ({ clients, setClients, user }) => {
  const canManage = hasPermission(user, "manageClients");
  const [modal, setModal] = useState(null); const [search, setSearch] = useState(""); const [viewMode, setViewMode] = useState("active"); const [toast, setToast] = useState(null); const [confirm, setConfirm] = useState(null);
  const blank = { name: "", industry: "", contact: "", email: "", phone: "", address: "", brands: "" };
  const [f, setF] = useState(blank); const u = k => v => setF(p => ({ ...p, [k]: v }));
  const industries = ["Healthcare/Pharma","FMCG","Banking/Finance","Telecoms","Government/NGO","Education","Real Estate","Energy/Oil & Gas","Retail","Technology","Media/Entertainment","Food & Beverage","Automotive","Other"];
  const save = async () => {
    if (!canManage) return setToast({ msg: readOnlyMessage(user), type: "error" });
    if (!f.name) return setToast({ msg: "Client name required.", type: "error" });
    const duplicate = clients.find(c => c.id !== modal?.id && !isArchived(c) && c.name.trim().toLowerCase() === f.name.trim().toLowerCase());
    if (duplicate) return setToast({ msg: "A client with this name already exists.", type: "error" });
    try {
      if (modal === "add") {
        const newClient = await createClientInSupabase(user.agencyId, user.id, f);
        setClients(v => [newClient, ...v]);
        createAuditEventInSupabase({ agencyId: user.agencyId, recordType: "client", recordId: newClient.id, action: "created", actor: user, metadata: { name: newClient.name || "" } }).catch(error => console.error("Failed to write audit event:", error));
      } else {
        const updatedClient = await updateClientInSupabase(modal.id, f);
        setClients(v => v.map(x => x.id === modal.id ? updatedClient : x));
        createAuditEventInSupabase({ agencyId: user.agencyId, recordType: "client", recordId: modal.id, action: "updated", actor: user, metadata: { name: updatedClient.name || "" } }).catch(error => console.error("Failed to write audit event:", error));
      }
      setToast({ msg: modal === "add" ? "Client added." : "Client updated.", type: "success" });
      setModal(null);
      setF(blank);
    } catch (e) {
      setToast({ msg: e.message || "Failed to save client.", type: "error" });
    }
  };
  const del = async id => {
    if (!canManage) return setToast({ msg: readOnlyMessage(user), type: "error" });
    if (!canManage) return setToast({ msg: readOnlyMessage(user), type: "error" });
    try {
      const archivedClient = await archiveClientInSupabase(id);
      setClients(v => v.map(x => x.id === id ? archivedClient : x));
      createAuditEventInSupabase({ agencyId: user.agencyId, recordType: "client", recordId: id, action: "archived", actor: user, metadata: { name: archivedClient.name || "" } }).catch(error => console.error("Failed to write audit event:", error));
      setToast({ msg: "Client archived.", type: "success" });
      setConfirm(null);
    } catch (e) {
      setToast({ msg: e.message || "Failed to archive client.", type: "error" });
    }
  };
  const restore = async id => {
    if (!canManage) return setToast({ msg: readOnlyMessage(user), type: "error" });
    if (!canManage) return setToast({ msg: readOnlyMessage(user), type: "error" });
    try {
      const restoredClient = await restoreClientInSupabase(id);
      setClients(v => v.map(x => x.id === id ? restoredClient : x));
      createAuditEventInSupabase({ agencyId: user.agencyId, recordType: "client", recordId: id, action: "restored", actor: user, metadata: { name: restoredClient.name || "" } }).catch(error => console.error("Failed to write audit event:", error));
      setToast({ msg: "Client restored.", type: "success" });
    } catch (e) {
      setToast({ msg: e.message || "Failed to restore client.", type: "error" });
    }
  };
  const visible = viewMode === "archived" ? archivedOnly(clients) : viewMode === "all" ? clients : activeOnly(clients);
  const filtered = visible.filter(c => `${c.name} ${c.industry}`.toLowerCase().includes(search.toLowerCase()));
  return (
    <div className="fade">
      {toast && <Toast msg={toast.msg} type={toast.type} onDone={() => setToast(null)} />}
      {confirm && <Confirm msg={confirm.msg} onYes={confirm.onYes} onNo={() => setConfirm(null)} />}
      {!canManage && <Card style={{ marginBottom: 14, background: "rgba(59,126,245,.06)", border: "1px solid rgba(59,126,245,.18)" }}><div style={{ fontSize: 13, color: "var(--text2)" }}>You have read-only access to Campaigns as <strong>{formatRoleLabel(user?.role)}</strong>.</div></Card>}
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 24, flexWrap: "wrap", gap: 10 }}>
        <div><h1 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 800, fontSize: 24 }}>Clients & Brands</h1><p style={{ color: "var(--text2)", marginTop: 3, fontSize: 13 }}>Manage your client portfolio</p></div>
        {canManage && <Btn icon="+" onClick={() => { setF(blank); setModal("add"); }}>Add Client</Btn>}
      </div>
      <div style={{ display: "flex", gap: 10, marginBottom: 18, flexWrap: "wrap" }}><div style={{ flex: 1, minWidth: 180 }}><Field value={search} onChange={setSearch} placeholder="Search clients…" /></div><Field value={viewMode} onChange={setViewMode} options={[{value:"active",label:"Active"},{value:"archived",label:"Archived"},{value:"all",label:"All"}]} /></div>
      {filtered.length === 0 ? <Card><Empty icon="👥" title="No clients yet" sub="Add your first client" /></Card> :
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill,minmax(280px,1fr))", gap: 14 }}>
          {filtered.map(c => {
            const brands = c.brands ? c.brands.split(",").map(b => b.trim()).filter(Boolean) : [];
            return (
              <Card key={c.id}>
                <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", marginBottom: 12 }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 11 }}>
                    <div style={{ width: 42, height: 42, background: `hsl(${c.name.charCodeAt(0)*7},40%,18%)`, borderRadius: 11, display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "'Syne',sans-serif", fontWeight: 800, fontSize: 17, color: `hsl(${c.name.charCodeAt(0)*7},70%,65%)` }}>{c.name[0]?.toUpperCase()}</div>
                    <div><div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 14 }}>{c.name}</div><div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginTop: 3 }}>{c.industry && <Badge color="blue">{c.industry}</Badge>}{isArchived(c) && <Badge color="red">Archived</Badge>}</div></div>
                  </div>
                  <div style={{ display: "flex", gap: 5 }}>
                    <Btn variant="ghost" size="sm" onClick={() => { setF({ name: c.name, industry: c.industry||"", contact: c.contact||"", email: c.email||"", phone: c.phone||"", address: c.address||"", brands: c.brands||"" }); setModal(c); }}>✏️</Btn>
                    {isArchived(c) ? <Btn variant="success" size="sm" onClick={() => restore(c.id)}>↩</Btn> : <Btn variant="danger" size="sm" onClick={() => setConfirm({ msg: `Archive "${c.name}"? Campaign history will be retained.`, onYes: () => del(c.id) })}>🗄</Btn>}
                  </div>
                </div>
                {brands.length > 0 && <div style={{ display: "flex", flexWrap: "wrap", gap: 5, marginBottom: 10 }}>{brands.map(b => <Badge key={b} color="green">{b}</Badge>)}</div>}
                {(c.email || c.phone) && <div style={{ fontSize: 11, color: "var(--text2)", borderTop: "1px solid var(--border)", paddingTop: 9 }}>{c.email && <div>📧 {c.email}</div>}{c.phone && <div style={{ marginTop: 2 }}>📞 {c.phone}</div>}</div>}
              </Card>
            );
          })}
        </div>}
      {modal !== null && (
        <Modal title={modal === "add" ? "Add Client" : `Edit: ${modal.name}`} onClose={() => setModal(null)} width={540}>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 }}>
            <div style={{ gridColumn: "1/-1" }}><Field label="Client / Organization Name" value={f.name} onChange={u("name")} placeholder="Breakthrough Action Nigeria" required /></div>
            <Field label="Industry" value={f.industry} onChange={u("industry")} options={industries} placeholder="Select industry" />
            <Field label="Contact Person" value={f.contact} onChange={u("contact")} placeholder="Primary contact" />
            <Field label="Email" type="email" value={f.email} onChange={u("email")} placeholder="client@org.com" />
            <Field label="Phone" value={f.phone} onChange={u("phone")} placeholder="+234 800 000 0000" />
            <div style={{ gridColumn: "1/-1" }}><Field label="Brands (comma-separated)" value={f.brands} onChange={u("brands")} placeholder="Brand A, Brand B" note="Separate multiple brands with commas" /></div>
            <div style={{ gridColumn: "1/-1" }}><Field label="Address" value={f.address} onChange={u("address")} placeholder="Physical address" /></div>
          </div>
          <div style={{ display: "flex", gap: 10, justifyContent: "flex-end", marginTop: 20 }}>
            <Btn variant="ghost" onClick={() => setModal(null)}>Cancel</Btn>
            <Btn onClick={save}>{modal === "add" ? "Add Client" : "Save Changes"}</Btn>
          </div>
        </Modal>
      )}
    </div>
  );
};


export default ClientsPage;
