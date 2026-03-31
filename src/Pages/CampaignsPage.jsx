import { useState } from "react";
import Badge from "../components/Badge";
import Empty from "../components/Empty";
import Toast from "../components/Toast";
import Modal from "../components/Modal";
import Confirm from "../components/Confirm";
import { activeOnly, archivedOnly, isArchived } from "../utils/records";
import { fmtN } from "../utils/formatters";
import { hasPermission, readOnlyMessage, formatRoleLabel } from "../constants/roles";
import { createCampaignInSupabase, updateCampaignInSupabase, archiveCampaignInSupabase, restoreCampaignInSupabase } from "../services/campaigns";
import { createAuditEventInSupabase } from "../services/notifications";
import { Card, Field, Btn } from "../components/ui/primitives";

/* ── CAMPAIGNS ──────────────────────────────────────────── */
const CampaignsPage = ({ campaigns, setCampaigns, clients, user }) => {
  const canManage = hasPermission(user, "manageCampaigns");
  const [modal, setModal] = useState(null); const [search, setSearch] = useState(""); const [filterStatus, setFilterStatus] = useState("all"); const [viewMode, setViewMode] = useState("active"); const [toast, setToast] = useState(null); const [confirm, setConfirm] = useState(null);
  const blank = { name: "", clientId: "", brand: "", objective: "", startDate: "", endDate: "", budget: "", status: "planning", medium: "", notes: "", materialList: [] };
  const [f, setF] = useState(blank); const u = k => v => setF(p => ({ ...p, [k]: v }));
  const save = async () => {
    if (!canManage) return setToast({ msg: readOnlyMessage(user), type: "error" });
    if (!f.name || !f.clientId) return setToast({ msg: "Campaign name and client required.", type: "error" });
    if (f.startDate && f.endDate && new Date(f.endDate) < new Date(f.startDate)) return setToast({ msg: "End date cannot be before start date.", type: "error" });
    if (f.budget && (parseFloat(f.budget) || 0) < 0) return setToast({ msg: "Budget cannot be negative.", type: "error" });
    const duplicate = campaigns.find(c => c.id !== modal?.id && !isArchived(c) && c.clientId === f.clientId && c.name.trim().toLowerCase() === f.name.trim().toLowerCase());
    if (duplicate) return setToast({ msg: "A campaign with this client and title already exists.", type: "error" });
    try {
      if (modal === "add") {
        const newCampaign = await createCampaignInSupabase(user.agencyId, user.id, f);
        setCampaigns(v => [newCampaign, ...v]);
        createAuditEventInSupabase({ agencyId: user.agencyId, recordType: "campaign", recordId: newCampaign.id, action: "created", actor: user, metadata: { name: newCampaign.name || "", status: newCampaign.status || "planning" } }).catch(error => console.error("Failed to write audit event:", error));
      } else {
        const updatedCampaign = await updateCampaignInSupabase(modal.id, f);
        setCampaigns(v => v.map(x => x.id === modal.id ? updatedCampaign : x));
        createAuditEventInSupabase({ agencyId: user.agencyId, recordType: "campaign", recordId: modal.id, action: "updated", actor: user, metadata: { name: updatedCampaign.name || "", status: updatedCampaign.status || "planning" } }).catch(error => console.error("Failed to write audit event:", error));
      }
      setToast({ msg: modal === "add" ? "Campaign created." : "Campaign updated.", type: "success" }); setModal(null); setF(blank);
    } catch (e) {
      setToast({ msg: e.message || "Failed to save campaign.", type: "error" });
    }
  };
  const del = async id => {
    if (!canManage) return setToast({ msg: readOnlyMessage(user), type: "error" });
    if (!canManage) return setToast({ msg: readOnlyMessage(user), type: "error" });
    try {
      const archived = await archiveCampaignInSupabase(id);
      setCampaigns(v => v.map(x => x.id === id ? archived : x));
      createAuditEventInSupabase({ agencyId: user.agencyId, recordType: "campaign", recordId: id, action: "archived", actor: user, metadata: { name: archived.name || "" } }).catch(error => console.error("Failed to write audit event:", error));
      setToast({ msg: "Campaign archived.", type: "success" }); setConfirm(null);
    } catch (e) {
      setToast({ msg: e.message || "Failed to archive campaign.", type: "error" });
    }
  };
  const restore = async id => {
    if (!canManage) return setToast({ msg: readOnlyMessage(user), type: "error" });
    if (!canManage) return setToast({ msg: readOnlyMessage(user), type: "error" });
    try {
      const restored = await restoreCampaignInSupabase(id);
      setCampaigns(v => v.map(x => x.id === id ? restored : x));
      createAuditEventInSupabase({ agencyId: user.agencyId, recordType: "campaign", recordId: id, action: "restored", actor: user, metadata: { name: restored.name || "" } }).catch(error => console.error("Failed to write audit event:", error));
      setToast({ msg: "Campaign restored.", type: "success" });
    } catch (e) {
      setToast({ msg: e.message || "Failed to restore campaign.", type: "error" });
    }
  };
  const clientOpts = activeOnly(clients).map(c => ({ value: c.id, label: c.name }));
  const clientBrands = f.clientId ? (clients.find(c => c.id === f.clientId)?.brands || "").split(",").map(b => b.trim()).filter(Boolean) : [];
  const visible = viewMode === "archived" ? archivedOnly(campaigns) : viewMode === "all" ? campaigns : activeOnly(campaigns);
  const filtered = visible.filter(c => {
    const cl = clients.find(x => x.id === c.clientId);
    return `${c.name} ${cl?.name || ""}`.toLowerCase().includes(search.toLowerCase()) && (filterStatus === "all" || c.status === filterStatus);
  });
  const getDur = c => { if (!c.startDate || !c.endDate) return "—"; const d = Math.ceil((new Date(c.endDate) - new Date(c.startDate)) / 86400000); return d > 0 ? `${d}d` : "—"; };
  const sc = { planning: "accent", active: "green", paused: "purple", completed: "blue" };
  return (
    <div className="fade">
      {toast && <Toast msg={toast.msg} type={toast.type} onDone={() => setToast(null)} />}
      {confirm && <Confirm msg={confirm.msg} onYes={confirm.onYes} onNo={() => setConfirm(null)} />}
      {!canManage && <Card style={{ marginBottom: 14, background: "rgba(59,126,245,.06)", border: "1px solid rgba(59,126,245,.18)" }}><div style={{ fontSize: 13, color: "var(--text2)" }}>You have read-only access to Vendors as <strong>{formatRoleLabel(user?.role)}</strong>.</div></Card>}
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 24, flexWrap: "wrap", gap: 10 }}>
        <div><h1 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 800, fontSize: 24 }}>Campaigns</h1><p style={{ color: "var(--text2)", marginTop: 3, fontSize: 13 }}>Track all advertising campaigns</p></div>
        {canManage && <Btn icon="+" onClick={() => { setF(blank); setModal("add"); }}>New Campaign</Btn>}
      </div>
      <div style={{ display: "flex", gap: 10, marginBottom: 18, flexWrap: "wrap" }}>
        <div style={{ flex: 1, minWidth: 180 }}><Field value={search} onChange={setSearch} placeholder="Search…" /></div>
        <Field value={filterStatus} onChange={setFilterStatus} options={[{value:"all",label:"All Status"},{value:"planning",label:"Planning"},{value:"active",label:"Active"},{value:"paused",label:"Paused"},{value:"completed",label:"Completed"}]} />
        <Field value={viewMode} onChange={setViewMode} options={[{value:"active",label:"Active"},{value:"archived",label:"Archived"},{value:"all",label:"All"}]} />
      </div>
      {filtered.length === 0 ? <Card><Empty icon="📢" title="No campaigns found" sub="Create your first campaign" /></Card> :
        <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
          {filtered.map(c => {
            const cl = clients.find(x => x.id === c.clientId);
            return (
              <Card key={c.id} style={{ padding: "14px 18px" }}>
                <div style={{ display: "flex", alignItems: "center", gap: 14, flexWrap: "wrap" }}>
                  <div style={{ width: 44, height: 44, background: "var(--bg3)", border: "1px solid var(--border2)", borderRadius: 11, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 20, flexShrink: 0 }}>📢</div>
                  <div style={{ flex: 1, minWidth: 160 }}>
                    <div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 15 }}>{c.name}</div>
                    <div style={{ fontSize: 12, color: "var(--text2)", marginTop: 2 }}>{cl?.name || "—"}{c.brand && ` · ${c.brand}`}{c.medium && ` · ${c.medium}`}</div>
                  </div>
                  <div style={{ display: "flex", gap: 18, flexWrap: "wrap", alignItems: "center" }}>
                    <div style={{ textAlign: "center" }}><div style={{ fontSize: 10, color: "var(--text3)", fontWeight: 600, textTransform: "uppercase" }}>Duration</div><div style={{ fontSize: 13, fontWeight: 600, marginTop: 2 }}>{getDur(c)}</div></div>
                    <div style={{ textAlign: "center" }}><div style={{ fontSize: 10, color: "var(--text3)", fontWeight: 600, textTransform: "uppercase" }}>Budget</div><div style={{ fontSize: 13, fontWeight: 600, marginTop: 2, color: "var(--green)" }}>{fmtN(c.budget)}</div></div>
                    <div style={{ textAlign: "center" }}><div style={{ fontSize: 10, color: "var(--text3)", fontWeight: 600, textTransform: "uppercase" }}>Period</div><div style={{ fontSize: 11, marginTop: 2 }}>{c.startDate || "—"} → {c.endDate || "—"}</div></div>
                    <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}><Badge color={sc[c.status] || "accent"}>{c.status}</Badge>{isArchived(c) && <Badge color="red">Archived</Badge>}</div>
                  </div>
                  <div style={{ display: "flex", gap: 6 }}>
                    {canManage && <Btn variant="ghost" size="sm" onClick={() => { setF({ name: c.name, clientId: c.clientId, brand: c.brand||"", objective: c.objective||"", startDate: c.startDate||"", endDate: c.endDate||"", budget: c.budget||"", status: c.status||"planning", medium: c.medium||"", notes: c.notes||"", materialList: Array.isArray(c.materialList) ? c.materialList : [] }); setModal(c); }}>✏️</Btn>}
                    {canManage && (isArchived(c) ? <Btn variant="success" size="sm" onClick={() => restore(c.id)}>↩</Btn> : <Btn variant="danger" size="sm" onClick={() => setConfirm({ msg: `Archive "${c.name}"? Linked MPO history will remain available.`, onYes: () => del(c.id) })}>🗄</Btn>)}
                  </div>
                </div>
                {c.objective && <div style={{ marginTop: 8, fontSize: 12, color: "var(--text2)", borderTop: "1px solid var(--border)", paddingTop: 8 }}>🎯 {c.objective}</div>}
                {c.materialList && c.materialList.length > 0 && <div style={{ marginTop: 6, fontSize: 11, color: "var(--text3)" }}>🎬 {c.materialList.length} material{c.materialList.length!==1?"s":""}: {c.materialList.slice(0,2).join(", ")}{c.materialList.length>2?` +${c.materialList.length-2} more`:""}</div>}
              </Card>
            );
          })}
        </div>}
      {modal !== null && (
        <Modal title={modal === "add" ? "New Campaign" : "Edit Campaign"} onClose={() => setModal(null)} width={580}>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 }}>
            <div style={{ gridColumn: "1/-1" }}><Field label="Campaign Title" value={f.name} onChange={u("name")} placeholder="TB Awareness Campaign Q4 2024" required /></div>
            <Field label="Client" value={f.clientId} onChange={v => { u("clientId")(v); u("brand")(""); }} options={clientOpts} placeholder="Select client" required />
            <Field label="Brand" value={f.brand} onChange={u("brand")} options={clientBrands.length ? clientBrands.map(b => ({ value: b, label: b })) : undefined} placeholder="Enter brand name" />
            <Field label="Campaign Objective" value={f.objective} onChange={u("objective")} placeholder="Awareness, Sales…" />
            <Field label="Medium" value={f.medium} onChange={u("medium")} options={["Television","Radio","Print","Digital","Multi-Platform","OOH"]} />
            <Field label="Start Date" type="date" value={f.startDate} onChange={u("startDate")} />
            <Field label="End Date" type="date" value={f.endDate} onChange={u("endDate")} />
            <Field label="Total Budget (₦)" type="number" value={f.budget} onChange={u("budget")} placeholder="0" />
            <Field label="Status" value={f.status} onChange={u("status")} options={[{value:"planning",label:"Planning"},{value:"active",label:"Active"},{value:"paused",label:"Paused"},{value:"completed",label:"Completed"}]} />
            <div style={{ gridColumn: "1/-1" }}>
              <div style={{ fontSize: 11, fontWeight: 600, color: "var(--text3)", textTransform: "uppercase", letterSpacing: ".08em", marginBottom: 6 }}>Campaign Materials</div>
              <div style={{ display: "flex", flexDirection: "column", gap: 7, marginBottom: 8 }}>
                {(f.materialList||[]).map((mat, mi) => (
                  <div key={mi} style={{ display: "flex", gap: 7, alignItems: "center" }}>
                    <input value={mat} onChange={e => { const ml = [...(f.materialList||[])]; ml[mi] = e.target.value; u("materialList")(ml); }}
                      placeholder="e.g. TB Thematic English 30secs (MP4)"
                      style={{ flex: 1, background: "var(--bg3)", border: "1px solid var(--border2)", borderRadius: 8, padding: "8px 12px", color: "var(--text)", fontSize: 13, outline: "none" }}
                      onFocus={e => e.target.style.borderColor="var(--accent)"}
                      onBlur={e => e.target.style.borderColor="var(--border2)"} />
                    <button onClick={() => { const ml = (f.materialList||[]).filter((_,i)=>i!==mi); u("materialList")(ml); }}
                      style={{ background: "rgba(239,68,68,.12)", border: "1px solid rgba(239,68,68,.3)", color: "var(--red)", borderRadius: 7, width: 32, height: 32, cursor: "pointer", fontSize: 14, flexShrink: 0 }}>×</button>
                  </div>
                ))}
                <button onClick={() => u("materialList")([...(f.materialList||[]), ""])}
                  style={{ display: "flex", alignItems: "center", gap: 6, background: "none", border: "1px dashed var(--border2)", borderRadius: 8, padding: "7px 14px", cursor: "pointer", fontSize: 12, color: "var(--text2)", width: "100%" }}>
                  <span style={{ fontSize: 16 }}>+</span> Add Material
                </button>
              </div>
              <div style={{ fontSize: 11, color: "var(--text3)" }}>Each material will be available as a dropdown when generating MPOs</div>
            </div>
            <div style={{ gridColumn: "1/-1" }}><Field label="Notes" value={f.notes} onChange={u("notes")} placeholder="Additional notes…" /></div>
          </div>
          <div style={{ display: "flex", gap: 10, justifyContent: "flex-end", marginTop: 20 }}>
            <Btn variant="ghost" onClick={() => setModal(null)}>Cancel</Btn>
            <Btn onClick={save}>{modal === "add" ? "Create Campaign" : "Save Changes"}</Btn>
          </div>
        </Modal>
      )}
    </div>
  );
};


export default CampaignsPage;
