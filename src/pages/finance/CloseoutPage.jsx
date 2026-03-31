import { useState } from "react";
import Empty from "../../components/Empty";
import Toast from "../../components/Toast";
import Btn from "../../components/Btn";
import Field from "../../components/Field";
import Badge from "../../components/Badge";
import Card from "../../components/Card";
import { MPO_PROOF_STATUS_OPTIONS, MPO_RECON_STATUS_OPTIONS } from "../../constants/mpoWorkflow";

const buildCSV = (rows = [], headers = []) => {
  const esc = (value) => `"${String(value ?? "").replace(/"/g, '""')}"`;
  return [headers, ...rows].map(row => row.map(esc).join(",")).join("\n");
};

const Stat = ({ label, value, sub, color = "var(--accent)", icon, valueSize = "clamp(18px, 2.25vw, 26px)" }) => (
  <Card hoverable style={{ position: "relative", overflow: "hidden", minWidth: 0, padding: "18px 20px" }}>
    <div style={{ position: "absolute", top: -16, right: -12, fontSize: 72, opacity: .04, pointerEvents: "none" }}>{icon}</div>
    <div style={{ fontSize: 11, color: "var(--text2)", fontWeight: 800, textTransform: "uppercase", letterSpacing: ".06em", marginBottom: 10, lineHeight: 1.2 }}>{label}</div>
    <div
      title={String(value ?? "")}
      style={{
        fontSize: valueSize,
        lineHeight: 1.08,
        fontWeight: 800,
        fontFamily: "Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif",
        fontVariantNumeric: "tabular-nums",
        letterSpacing: "-0.04em",
        color,
        minWidth: 0,
        maxWidth: "100%",
        overflow: "hidden",
        textOverflow: "ellipsis",
        whiteSpace: "nowrap",
      }}
    >
      {value}
    </div>
    {sub && <div style={{ fontSize: 11, color: "var(--text2)", opacity: 0.92, marginTop: 10, lineHeight: 1.45, overflowWrap: "anywhere" }}>{sub}</div>}
  </Card>
);

const CloseoutPage = ({ vendors, clients, campaigns, mpos, onOpenReceivables, activeOnly, fmtN, buildCSV, PrintPreview }) => {
  const [toast, setToast] = useState(null);
  const [preview, setPreview] = useState(null);
  const [filters, setFilters] = useState({ clientId: "", proofStatus: "", reconciliationStatus: "", closeoutStage: "", search: "" });

  const liveClients = activeOnly(clients);
  const liveCampaigns = activeOnly(campaigns);
  const liveMpos = activeOnly(mpos);
  const updateFilter = (key) => (value) => setFilters(prev => ({ ...prev, [key]: value }));
  const stageRank = { live: 1, reconciling: 2, billed: 3, collecting: 4, closed: 5, exception: 6 };

  const getCloseoutStage = (mpo) => {
    const proof = String(mpo?.proofStatus || "pending").toLowerCase();
    const recon = String(mpo?.reconciliationStatus || "not_started").toLowerCase();
    const invoice = String(mpo?.invoiceStatus || "pending").toLowerCase();
    const payment = String(mpo?.paymentStatus || "unpaid").toLowerCase();
    const status = String(mpo?.status || "draft").toLowerCase();
    if ([proof, recon, invoice, payment].includes("disputed")) return "exception";
    if (status === "closed" || (payment === "paid" && recon === "completed")) return "closed";
    if (payment === "processing" || (payment === "unpaid" && recon === "completed" && invoice !== "pending")) return "collecting";
    if (invoice !== "pending" && recon === "completed") return "billed";
    if (status === "aired" || recon === "in_progress" || recon === "ready" || proof !== "pending") return "reconciling";
    return "live";
  };

  const getCloseoutLabel = (stage) => ({
    live: "Live",
    reconciling: "Reconciling",
    billed: "Billed",
    collecting: "Collecting",
    closed: "Closed",
    exception: "Exception",
  }[stage] || stage);

  const getCloseoutColor = (stage) => ({
    live: "blue",
    reconciling: "purple",
    billed: "accent",
    collecting: "orange",
    closed: "green",
    exception: "red",
  }[stage] || "gray");

  const filteredMpos = liveMpos.filter(mpo => {
    const campaign = liveCampaigns.find(row => row.id === mpo.campaignId);
    const client = liveClients.find(row => row.id === campaign?.clientId);
    const stage = getCloseoutStage(mpo);
    if (filters.clientId && (campaign?.clientId || "") !== filters.clientId) return false;
    if (filters.proofStatus && String(mpo.proofStatus || "pending") !== filters.proofStatus) return false;
    if (filters.reconciliationStatus && String(mpo.reconciliationStatus || "not_started") !== filters.reconciliationStatus) return false;
    if (filters.closeoutStage && stage !== filters.closeoutStage) return false;
    const term = filters.search.trim().toLowerCase();
    if (!term) return true;
    return [mpo.mpoNo, mpo.vendorName, mpo.clientName, mpo.campaignName, mpo.brand, mpo.invoiceNo, mpo.paymentReference, campaign?.name, client?.name, stage]
      .some(value => String(value || "").toLowerCase().includes(term));
  });

  const closeoutRows = filteredMpos.map(mpo => {
    const campaign = liveCampaigns.find(row => row.id === mpo.campaignId);
    const client = liveClients.find(row => row.id === campaign?.clientId);
    const planned = Number(mpo.plannedSpotsExecution ?? mpo.totalSpots) || 0;
    const aired = Number(mpo.airedSpots) || 0;
    const missed = Number(mpo.missedSpots) || Math.max(planned - aired, 0);
    const makegood = Number(mpo.makegoodSpots) || 0;
    const deliveryPct = planned > 0 ? (aired / planned) * 100 : (aired > 0 ? 100 : 0);
    const variance = aired - planned;
    const value = Number(mpo.reconciledAmount) || Number(mpo.invoiceAmount) || Number(mpo.grandTotal) || Number(mpo.netVal) || 0;
    const stage = getCloseoutStage(mpo);
    const proof = String(mpo.proofStatus || "pending");
    const recon = String(mpo.reconciliationStatus || "not_started");
    const invoice = String(mpo.invoiceStatus || "pending");
    const payment = String(mpo.paymentStatus || "unpaid");
    const completeness = [
      mpo.dispatchStatus && mpo.dispatchStatus !== "pending",
      proof !== "pending",
      aired > 0 || planned > 0,
      recon === "completed",
      invoice !== "pending",
      payment === "paid" || String(mpo.status || "").toLowerCase() === "closed",
    ].filter(Boolean).length / 6 * 100;
    const readyToBill = recon === "completed" && ["received", "approved"].includes(proof);
    const readyToClose = readyToBill && (payment === "paid" || String(mpo.status || "").toLowerCase() === "closed");
    return {
      id: mpo.id,
      mpoNo: mpo.mpoNo || "—",
      campaign: campaign?.name || mpo.campaignName || "—",
      client: client?.name || mpo.clientName || "—",
      vendor: mpo.vendorName || "—",
      brand: campaign?.brand || mpo.brand || "—",
      planned,
      aired,
      missed,
      makegood,
      deliveryPct,
      variance,
      proofStatus: proof,
      reconciliationStatus: recon,
      invoiceStatus: invoice,
      paymentStatus: payment,
      reconciledAmount: value,
      stage,
      stageLabel: getCloseoutLabel(stage),
      stageColor: getCloseoutColor(stage),
      completeness,
      readyToBill,
      readyToClose,
      invoiceNo: mpo.invoiceNo || "—",
      paymentReference: mpo.paymentReference || "—",
      proofReceivedAt: mpo.proofReceivedAt ? new Date(mpo.proofReceivedAt).toLocaleDateString("en-NG") : "—",
      paidAt: mpo.paidAt ? new Date(mpo.paidAt).toLocaleDateString("en-NG") : "—",
      notes: mpo.reconciliationNotes || "",
    };
  }).sort((a, b) => (stageRank[a.stage] || 0) - (stageRank[b.stage] || 0) || b.reconciledAmount - a.reconciledAmount);

  const stageSummary = ["live", "reconciling", "billed", "collecting", "closed", "exception"].map(stage => {
    const rows = closeoutRows.filter(row => row.stage === stage);
    return { stage, label: getCloseoutLabel(stage), color: getCloseoutColor(stage), count: rows.length, value: rows.reduce((sum, row) => sum + row.reconciledAmount, 0) };
  });

  const makegoodRows = closeoutRows.filter(row => row.missed > 0 || row.makegood > 0 || row.proofStatus === "disputed" || row.reconciliationStatus === "disputed");
  const billingRows = closeoutRows.filter(row => row.readyToBill || row.stage === "collecting" || row.stage === "closed");
  const collectionsRows = closeoutRows.filter(row => ["billed", "collecting", "closed"].includes(row.stage));
  const totalValue = closeoutRows.reduce((sum, row) => sum + row.reconciledAmount, 0);
  const readyToBillCount = closeoutRows.filter(row => row.readyToBill).length;
  const proofPendingCount = closeoutRows.filter(row => row.proofStatus === "pending").length;
  const exceptionCount = closeoutRows.filter(row => row.stage === "exception").length;
  const closedCount = closeoutRows.filter(row => row.stage === "closed").length;
  const openMakegoodCount = makegoodRows.filter(row => row.missed > row.makegood || row.proofStatus === "disputed" || row.reconciliationStatus === "disputed").length;
  const collectionsExposure = closeoutRows.filter(row => row.stage !== "closed").reduce((sum, row) => sum + row.reconciledAmount, 0);
  const avgDelivery = closeoutRows.length ? closeoutRows.reduce((sum, row) => sum + Math.min(row.deliveryPct, 100), 0) / closeoutRows.length : 0;

  const exportView = (title, headers, rows) => {
    if (!rows.length) {
      setToast({ msg: `No data to export for ${title}.`, type: "error" });
      return;
    }
    const esc = v => String(v ?? "").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");
    const htmlRows = rows.map(row => `<tr>${row.map((cell, i) => `<td style="padding:6px 9px;border:1px solid #ddd;font-size:10px;${i===0?"font-weight:600":""}">${esc(cell)}</td>`).join("")}</tr>`).join("");
    const html = `<!DOCTYPE html><html><head><meta charset="utf-8"><title>${title}</title>
      <style>
        body{font-family:Arial,sans-serif;padding:24px;color:#000;margin:0}
        h1{font-size:18px;margin-bottom:4px;color:#0A1F44}
        p{font-size:11px;color:#555;margin-bottom:16px}
        table{width:100%;border-collapse:collapse;margin-top:8px}
        th{background:#0A1F44;color:#fff;padding:7px 9px;font-size:10px;text-align:left;border:1px solid #0A1F44}
        tr:nth-child(even){background:#F5F7FA}
      </style>
    </head><body>
      <h1>${title}</h1>
      <p>Generated ${new Date().toLocaleString("en-NG")}</p>
      <table><tr>${headers.map(h => `<th>${esc(h)}</th>`).join("")}</tr>${htmlRows}</table>
    </body></html>`;
    const csv = buildCSV(rows, headers);
    setPreview({ html, csv, title });
  };

  const exportCloseoutPack = (row) => {
    const completenessColor = row.completeness >= 100 ? "#16a34a" : row.completeness >= 70 ? "#f59e0b" : "#ef4444";
    const html = `<!DOCTYPE html><html><head><meta charset="utf-8"><title>Closeout Pack - ${row.mpoNo}</title>
      <style>
        body{font-family:Arial,sans-serif;padding:28px;color:#111;margin:0;background:#fff}
        h1{font-size:22px;margin:0 0 4px;color:#0A1F44}
        h2{font-size:14px;margin:0 0 10px;color:#0A1F44;text-transform:uppercase;letter-spacing:.08em}
        .muted{font-size:12px;color:#666}
        .grid{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:12px;margin:18px 0}
        .card{border:1px solid #d6dde8;border-radius:12px;padding:14px 16px;background:#f8fafc}
        .label{font-size:10px;color:#667085;text-transform:uppercase;letter-spacing:.08em}
        .value{font-size:16px;font-weight:700;margin-top:6px}
        table{width:100%;border-collapse:collapse;margin-top:8px}
        th,td{border:1px solid #d6dde8;padding:8px 10px;text-align:left;font-size:11px}
        th{background:#0A1F44;color:#fff}
        .pill{display:inline-block;padding:4px 8px;border-radius:999px;font-size:11px;font-weight:700;background:#eef2ff;color:#312e81}
      </style>
    </head><body>
      <h1>Campaign Closeout Pack</h1>
      <div class="muted">Generated ${new Date().toLocaleString("en-NG")}</div>
      <div class="grid">
        <div class="card"><div class="label">MPO</div><div class="value">${row.mpoNo}</div><div class="muted">${row.campaign}</div></div>
        <div class="card"><div class="label">Client / Brand</div><div class="value">${row.client}</div><div class="muted">${row.brand}</div></div>
        <div class="card"><div class="label">Vendor</div><div class="value">${row.vendor}</div><div class="muted">Closeout stage: ${row.stageLabel}</div></div>
      </div>
      <div class="grid">
        <div class="card"><div class="label">Delivery</div><div class="value">${row.aired} / ${row.planned} spots</div><div class="muted">${row.deliveryPct.toFixed(1)}% delivered</div></div>
        <div class="card"><div class="label">Make-good</div><div class="value">${row.makegood} spots</div><div class="muted">Missed: ${row.missed}</div></div>
        <div class="card"><div class="label">Reconciled Amount</div><div class="value">${fmtN(row.reconciledAmount)}</div><div class="muted">Invoice: ${row.invoiceNo}</div></div>
      </div>
      <div class="card" style="margin-bottom:18px">
        <h2>Closure Health</h2>
        <div class="muted">Completeness</div>
        <div style="margin-top:8px;height:10px;border-radius:999px;background:#e5e7eb;overflow:hidden"><div style="width:${Math.min(row.completeness,100)}%;height:100%;background:${completenessColor}"></div></div>
        <div class="muted" style="margin-top:8px">${row.completeness.toFixed(0)}% complete · Proof ${row.proofStatus} · Reconciliation ${row.reconciliationStatus} · Payment ${row.paymentStatus}</div>
      </div>
      <h2>Closeout Checklist</h2>
      <table>
        <tr><th>Checkpoint</th><th>Status</th><th>Detail</th></tr>
        <tr><td>Proof of airing</td><td>${row.proofStatus}</td><td>Received: ${row.proofReceivedAt}</td></tr>
        <tr><td>Reconciliation</td><td>${row.reconciliationStatus}</td><td>${row.notes || "No reconciliation notes recorded."}</td></tr>
        <tr><td>Billing readiness</td><td>${row.readyToBill ? "Ready" : "Pending"}</td><td>Invoice ${row.invoiceStatus}</td></tr>
        <tr><td>Collection / payment</td><td>${row.paymentStatus}</td><td>Reference: ${row.paymentReference} · Paid at: ${row.paidAt}</td></tr>
        <tr><td>Final status</td><td>${row.readyToClose ? "Ready to close" : row.stageLabel}</td><td>Workflow stage aligned to current MPO finance fields.</td></tr>
      </table>
    </body></html>`;
    setPreview({ html, csv: "", title: `Closeout Pack - ${row.mpoNo}` });
  };

  const filterInputStyle = {
    background: "var(--bg3)",
    border: "1px solid var(--border2)",
    borderRadius: 8,
    padding: "9px 13px",
    color: "var(--text)",
    fontSize: 13,
    outline: "none",
    width: "100%",
  };

  return (
    <div className="fade">
      {toast && <Toast msg={toast.msg} type={toast.type} onDone={() => setToast(null)} />}
      {preview && <PrintPreview html={preview.html} csv={preview.csv} pdfBytes={preview.pdfBytes} title={preview.title} onClose={() => setPreview(null)} />}

      <div style={{ marginBottom: 24 }}>
        <h1 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 800, fontSize: 24 }}>Campaign Reconciliation & Performance Closeout</h1>
        <p style={{ color: "var(--text2)", marginTop: 3, fontSize: 13 }}>Manage proof of airing, variance, make-goods, billing readiness, and campaign closure from one end-of-flight workspace.</p>
      </div>

      <Card style={{ marginBottom: 20 }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, flexWrap: "wrap", marginBottom: 16 }}>
          <div>
            <h2 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 16 }}>Closeout Filters</h2>
            <p style={{ color: "var(--text2)", fontSize: 12, marginTop: 3 }}>Slice execution and closure signals by client, proof, reconciliation, stage, and keyword.</p>
          </div>
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
            <Btn variant="ghost" size="sm" onClick={() => setFilters({ clientId: "", proofStatus: "", reconciliationStatus: "", closeoutStage: "", search: "" })}>Reset</Btn>
            <Btn variant="blue" size="sm" icon="⬇" onClick={() => exportView("Closeout Summary", ["Metric", "Value"], [["MPO Value in Closeout", totalValue], ["Ready to Bill", readyToBillCount], ["Proof Pending", proofPendingCount], ["Open Make-good Items", openMakegoodCount], ["Exceptions", exceptionCount], ["Closed", closedCount], ["Collections Exposure", collectionsExposure], ["Average Delivery", `${avgDelivery.toFixed(1)}%`]])}>Export Summary</Btn>
          </div>
        </div>

        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(180px,1fr))", gap: 12 }}>
          <Field label="Client" value={filters.clientId} onChange={updateFilter("clientId")} options={liveClients.map(client => ({ value: client.id, label: client.name }))} placeholder="All Clients" />
          <Field label="Proof Status" value={filters.proofStatus} onChange={updateFilter("proofStatus")} options={MPO_PROOF_STATUS_OPTIONS} placeholder="All Proof States" />
          <Field label="Reconciliation" value={filters.reconciliationStatus} onChange={updateFilter("reconciliationStatus")} options={[...MPO_RECON_STATUS_OPTIONS, { value: "disputed", label: "Disputed" }]} placeholder="All Reconciliation States" />
          <Field label="Closeout Stage" value={filters.closeoutStage} onChange={updateFilter("closeoutStage")} options={[{ value: "live", label: "Live" }, { value: "reconciling", label: "Reconciling" }, { value: "billed", label: "Billed" }, { value: "collecting", label: "Collecting" }, { value: "closed", label: "Closed" }, { value: "exception", label: "Exception" }]} placeholder="All Stages" />
          <div>
            <label style={{ fontSize: 11, fontWeight: 600, color: "var(--text3)", textTransform: "uppercase", letterSpacing: ".08em", marginBottom: 5, display: "block" }}>Search</label>
            <input value={filters.search} onChange={e => updateFilter("search")(e.target.value)} placeholder="MPO, campaign, vendor, invoice..." style={filterInputStyle} />
          </div>
        </div>
      </Card>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(180px,1fr))", gap: 14, marginBottom: 22 }}>
        <Stat icon="📦" label="Closeout Value" value={fmtN(totalValue)} sub={`${closeoutRows.length} MPOs`} color="var(--accent)" valueSize="clamp(16px, 1.7vw, 22px)" />
        <Stat icon="🧾" label="Ready to Bill" value={readyToBillCount} sub="Reconciled delivery ready" color="var(--blue)" valueSize="clamp(16px, 1.7vw, 22px)" />
        <Stat icon="📎" label="Proof Pending" value={proofPendingCount} sub="Awaiting proof of airing" color="var(--purple)" valueSize="clamp(16px, 1.7vw, 22px)" />
        <Stat icon="🔁" label="Open Make-goods" value={openMakegoodCount} sub="Missed or disputed delivery" color="var(--orange)" valueSize="clamp(16px, 1.7vw, 22px)" />
        <Stat icon="🚨" label="Exceptions" value={exceptionCount} sub="Disputed closeout items" color="var(--red)" valueSize="clamp(16px, 1.7vw, 22px)" />
        <Stat icon="✅" label="Closed" value={closedCount} sub={`${avgDelivery.toFixed(1)}% avg delivery`} color="var(--green)" valueSize="clamp(16px, 1.7vw, 22px)" />
      </div>

      <Card style={{ marginBottom: 18 }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, flexWrap: "wrap", marginBottom: 14 }}>
          <div>
            <h2 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 15 }}>Closeout Pipeline</h2>
            <p style={{ color: "var(--text2)", fontSize: 12, marginTop: 3 }}>Track MPOs through live delivery, reconciliation, billing, collection, and final closure.</p>
          </div>
          <Btn variant="ghost" size="sm" onClick={() => exportView("Closeout Pipeline", ["Stage", "Count", "Value"], stageSummary.map(row => [row.label, row.count, row.value]))}>Export Pipeline</Btn>
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(160px,1fr))", gap: 12 }}>
          {stageSummary.map(row => (
            <div key={row.stage} style={{ padding: "12px 14px", borderRadius: 12, background: "var(--bg3)", border: "1px solid var(--border)" }}>
              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8 }}>
                <div style={{ fontSize: 12, color: "var(--text3)", textTransform: "uppercase", fontWeight: 700 }}>{row.label}</div>
                <Badge color={row.color}>{row.count}</Badge>
              </div>
              <div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 800, fontSize: 20, marginTop: 10 }}>{fmtN(row.value)}</div>
              <div style={{ fontSize: 11, color: "var(--text2)", marginTop: 4 }}>{row.count === 1 ? "1 MPO" : `${row.count} MPOs`}</div>
            </div>
          ))}
        </div>
      </Card>

      <div style={{ display: "grid", gridTemplateColumns: "1.12fr .88fr", gap: 18, alignItems: "start" }}>
        <Card>
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, flexWrap: "wrap", marginBottom: 14 }}>
            <div>
              <h2 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 15 }}>Execution Variance & Billing Readiness</h2>
              <p style={{ color: "var(--text2)", fontSize: 12, marginTop: 3 }}>Compare planned vs aired spots and identify which MPOs are ready for billing or closure.</p>
            </div>
            <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
              <Btn variant="ghost" size="sm" onClick={() => setPage && setPage("mpo")}>Open MPO Workspace</Btn>
              <Btn variant="ghost" size="sm" onClick={() => exportView("Execution Variance and Billing Readiness", ["MPO No.", "Campaign", "Client", "Vendor", "Planned", "Aired", "Missed", "Make-good", "Delivery %", "Proof", "Reconciliation", "Invoice", "Payment", "Stage", "Ready to Bill", "Ready to Close", "Reconciled Amount"], closeoutRows.map(row => [row.mpoNo, row.campaign, row.client, row.vendor, row.planned, row.aired, row.missed, row.makegood, `${row.deliveryPct.toFixed(1)}%`, row.proofStatus, row.reconciliationStatus, row.invoiceStatus, row.paymentStatus, row.stageLabel, row.readyToBill ? "Yes" : "No", row.readyToClose ? "Yes" : "No", row.reconciledAmount]))}>Export Matrix</Btn>
            </div>
          </div>
          {closeoutRows.length === 0 ? <Empty icon="📑" title="No closeout rows yet" sub="As campaigns air and proofs arrive, closeout signals will appear here." /> : (
            <div style={{ overflowX: "auto" }}>
              <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 1120 }}>
                <thead>
                  <tr style={{ background: "var(--bg3)" }}>
                    {["MPO / Campaign", "Delivery", "Proof / Recon", "Billing", "Closeout", "Pack"].map(header => <th key={header} style={{ padding: "8px 10px", textAlign: "left", fontSize: 10, fontWeight: 600, textTransform: "uppercase", color: "var(--text3)", borderBottom: "1px solid var(--border)" }}>{header}</th>)}
                  </tr>
                </thead>
                <tbody>
                  {closeoutRows.slice(0, 18).map((row, index) => {
                    const deliveryColor = row.deliveryPct >= 100 ? "var(--green)" : row.deliveryPct >= 80 ? "var(--orange)" : "var(--red)";
                    return (
                      <tr key={row.id} style={{ borderBottom: "1px solid var(--border)", background: index % 2 === 0 ? "transparent" : "rgba(255,255,255,.01)" }}>
                        <td style={{ padding: "9px 10px", fontSize: 12, fontWeight: 700 }}>{row.mpoNo}<div style={{ fontSize: 11, color: "var(--text3)", fontWeight: 400, marginTop: 3 }}>{row.campaign} · {row.client}</div><div style={{ fontSize: 11, color: "var(--text3)", marginTop: 3 }}>{row.vendor}</div></td>
                        <td style={{ padding: "9px 10px", fontSize: 12 }}><div style={{ display: "flex", alignItems: "center", gap: 8 }}><div style={{ flex: 1, minWidth: 84, height: 8, borderRadius: 999, background: "var(--bg4)", overflow: "hidden" }}><div style={{ width: `${Math.min(row.deliveryPct, 100)}%`, height: "100%", background: deliveryColor }} /></div><span style={{ color: deliveryColor, fontWeight: 700 }}>{row.deliveryPct.toFixed(1)}%</span></div><div style={{ fontSize: 11, color: "var(--text3)", marginTop: 4 }}>Planned {row.planned} · Aired {row.aired} · Missed {row.missed} · Make-good {row.makegood}</div></td>
                        <td style={{ padding: "9px 10px", fontSize: 12 }}><div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}><Badge color={row.proofStatus === "received" ? "green" : row.proofStatus === "partial" ? "orange" : row.proofStatus === "disputed" ? "red" : "accent"}>{row.proofStatus}</Badge><Badge color={row.reconciliationStatus === "completed" ? "green" : row.reconciliationStatus === "ready" ? "blue" : row.reconciliationStatus === "disputed" ? "red" : "purple"}>{row.reconciliationStatus}</Badge></div><div style={{ fontSize: 11, color: "var(--text3)", marginTop: 5 }}>{row.proofReceivedAt}</div></td>
                        <td style={{ padding: "9px 10px", fontSize: 12 }}><div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}><Badge color={row.invoiceStatus === "approved" ? "green" : row.invoiceStatus === "received" ? "blue" : row.invoiceStatus === "disputed" ? "red" : "accent"}>{row.invoiceStatus}</Badge><Badge color={row.paymentStatus === "paid" ? "green" : row.paymentStatus === "processing" ? "blue" : row.paymentStatus === "disputed" ? "red" : "orange"}>{row.paymentStatus}</Badge></div><div style={{ fontSize: 11, color: "var(--text3)", marginTop: 5 }}>{fmtN(row.reconciledAmount)}</div></td>
                        <td style={{ padding: "9px 10px", fontSize: 12 }}><Badge color={row.stageColor}>{row.stageLabel}</Badge><div style={{ fontSize: 11, color: "var(--text3)", marginTop: 6 }}>{row.readyToBill ? "Ready to bill" : "Billing pending"} · {row.readyToClose ? "Ready to close" : "Still open"}</div><div style={{ marginTop: 6, height: 6, background: "var(--bg4)", borderRadius: 999, overflow: "hidden" }}><div style={{ width: `${Math.min(row.completeness, 100)}%`, height: "100%", background: row.completeness >= 100 ? "var(--green)" : row.completeness >= 70 ? "var(--orange)" : "var(--red)" }} /></div></td>
                        <td style={{ padding: "9px 10px", fontSize: 12 }}><Btn variant="ghost" size="sm" onClick={() => exportCloseoutPack(row)}>Export Pack</Btn></td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
              {closeoutRows.length > 18 ? <div style={{ marginTop: 10, fontSize: 12, color: "var(--text3)" }}>Showing 18 of {closeoutRows.length} rows in-app. Use export for the full matrix.</div> : null}
            </div>
          )}
        </Card>

        <div style={{ display: "flex", flexDirection: "column", gap: 18 }}>
          <Card>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, marginBottom: 14 }}>
              <h2 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 15 }}>Make-good & Exception Tracker</h2>
              <Btn variant="ghost" size="sm" onClick={() => exportView("Make-good and Exception Tracker", ["MPO No.", "Campaign", "Vendor", "Missed", "Make-good", "Proof", "Reconciliation", "Stage", "Notes"], makegoodRows.map(row => [row.mpoNo, row.campaign, row.vendor, row.missed, row.makegood, row.proofStatus, row.reconciliationStatus, row.stageLabel, row.notes]))}>Export Tracker</Btn>
            </div>
            {makegoodRows.length === 0 ? <Empty icon="🔁" title="No make-good issues" sub="Missed spots and disputes will surface here for follow-up." /> : (
              <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                {makegoodRows.slice(0, 6).map(row => (
                  <div key={row.id} style={{ padding: "11px 12px", borderRadius: 10, background: "var(--bg3)", border: "1px solid var(--border)" }}>
                    <div style={{ display: "flex", justifyContent: "space-between", gap: 10, alignItems: "center" }}>
                      <div>
                        <div style={{ fontSize: 13, fontWeight: 700 }}>{row.mpoNo}</div>
                        <div style={{ fontSize: 11, color: "var(--text3)", marginTop: 2 }}>{row.campaign} · {row.vendor}</div>
                      </div>
                      <Badge color={row.missed > row.makegood || row.stage === "exception" ? "red" : "orange"}>{row.stageLabel}</Badge>
                    </div>
                    <div style={{ display: "grid", gridTemplateColumns: "repeat(3,minmax(0,1fr))", gap: 8, marginTop: 10 }}>
                      <div style={{ background: "var(--bg2)", borderRadius: 8, padding: "8px 10px" }}><div style={{ fontSize: 10, color: "var(--text3)", textTransform: "uppercase" }}>Missed</div><div style={{ fontSize: 12, fontWeight: 700, marginTop: 4 }}>{row.missed}</div></div>
                      <div style={{ background: "var(--bg2)", borderRadius: 8, padding: "8px 10px" }}><div style={{ fontSize: 10, color: "var(--text3)", textTransform: "uppercase" }}>Make-good</div><div style={{ fontSize: 12, fontWeight: 700, marginTop: 4 }}>{row.makegood}</div></div>
                      <div style={{ background: "var(--bg2)", borderRadius: 8, padding: "8px 10px" }}><div style={{ fontSize: 10, color: "var(--text3)", textTransform: "uppercase" }}>Proof / Recon</div><div style={{ fontSize: 12, fontWeight: 700, marginTop: 4 }}>{row.proofStatus} / {row.reconciliationStatus}</div></div>
                    </div>
                  </div>
                ))}
              </div>
            )}
          </Card>

          <Card>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, marginBottom: 14 }}>
              <h2 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 15 }}>Billing & Collections Watchlist</h2>
              <Btn variant="ghost" size="sm" onClick={() => exportView("Billing and Collections Watchlist", ["MPO No.", "Client", "Campaign", "Stage", "Invoice Status", "Payment Status", "Ready to Bill", "Amount", "Payment Ref", "Paid At"], collectionsRows.map(row => [row.mpoNo, row.client, row.campaign, row.stageLabel, row.invoiceStatus, row.paymentStatus, row.readyToBill ? "Yes" : "No", row.reconciledAmount, row.paymentReference, row.paidAt]))}>Export Watchlist</Btn>
            </div>
            {billingRows.length === 0 ? <Empty icon="💼" title="No billing items yet" sub="Completed reconciliation will move rows into this watchlist." /> : (
              <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                {billingRows.slice(0, 6).map(row => (
                  <div key={row.id} style={{ padding: "11px 12px", borderRadius: 10, background: "var(--bg3)", border: "1px solid var(--border)" }}>
                    <div style={{ display: "flex", justifyContent: "space-between", gap: 10, alignItems: "center" }}>
                      <div>
                        <div style={{ fontSize: 13, fontWeight: 700 }}>{row.client}</div>
                        <div style={{ fontSize: 11, color: "var(--text3)", marginTop: 2 }}>{row.mpoNo} · {row.campaign}</div>
                      </div>
                      <Badge color={row.stageColor}>{row.stageLabel}</Badge>
                    </div>
                    <div style={{ marginTop: 10, display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10 }}>
                      <div>
                        <div style={{ fontSize: 11, color: "var(--text3)" }}>Exposure</div>
                        <div style={{ fontSize: 14, fontWeight: 800, marginTop: 3 }}>{fmtN(row.reconciledAmount)}</div>
                      </div>
                      <div style={{ textAlign: "right" }}>
                        <div style={{ fontSize: 11, color: "var(--text3)" }}>{row.readyToBill ? "Ready to bill" : "Needs follow-up"}</div>
                        <div style={{ fontSize: 12, fontWeight: 700, marginTop: 3 }}>{row.paymentStatus === "paid" ? row.paidAt : row.invoiceStatus}</div>
                      </div>
                    </div>
                  </div>
                ))}
              </div>
            )}
          </Card>
        </div>
      </div>
    </div>
  );
};



export default CloseoutPage;
