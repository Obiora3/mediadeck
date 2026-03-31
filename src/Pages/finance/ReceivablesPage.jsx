import { useState } from "react";
import Empty from "../../components/Empty";
import Toast from "../../components/Toast";
import Modal from "../../components/Modal";
import Btn from "../../components/Btn";
import Field from "../../components/Field";
import Badge from "../../components/Badge";
import Card from "../../components/Card";


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

const ReceivablesPage = ({ user, clients, campaigns, mpos, receivables, receivablesMeta, onSaveReceivable, onRemoveReceivable, onLogReceivablePayment, onUpdateReceivableStatus, onOpenCloseout, activeOnly, fmtN, buildCSV, PrintPreview, uid, isoToday, addDaysToIso, normalizeReceivableRecord, getDaysPastDue, formatIsoDate, buildReceivableFromMpo }) => {
  const [toast, setToast] = useState(null);
  const [preview, setPreview] = useState(null);
  const [filters, setFilters] = useState({ clientId: "", status: "", dueBucket: "", search: "" });
  const [formModal, setFormModal] = useState(null);
  const [paymentModal, setPaymentModal] = useState(null);
  const [formBusy, setFormBusy] = useState(false);
  const [paymentBusy, setPaymentBusy] = useState(false);

  const liveClients = activeOnly(clients);
  const liveCampaigns = activeOnly(campaigns);
  const liveMpos = activeOnly(mpos);
  const updateFilter = (key) => (value) => setFilters(prev => ({ ...prev, [key]: value }));

  const statusOptions = [
    { value: "draft", label: "Draft" },
    { value: "issued", label: "Issued" },
    { value: "part_paid", label: "Part Paid" },
    { value: "paid", label: "Paid" },
    { value: "overdue", label: "Overdue" },
    { value: "disputed", label: "Disputed" },
    { value: "write_off", label: "Written Off" },
  ];
  const stageOptions = [
    { value: "invoicing", label: "Invoicing" },
    { value: "follow_up", label: "Follow-up" },
    { value: "promise_to_pay", label: "Promise to Pay" },
    { value: "escalated", label: "Escalated" },
    { value: "resolved", label: "Resolved" },
  ];
  const paymentChannelOptions = [
    { value: "bank_transfer", label: "Bank Transfer" },
    { value: "cash", label: "Cash" },
    { value: "cheque", label: "Cheque" },
    { value: "card", label: "Card" },
    { value: "other", label: "Other" },
  ];

  const getStatusLabel = (status) => (statusOptions.find(option => option.value === status)?.label || status || "—");
  const getStageLabel = (stage) => (stageOptions.find(option => option.value === stage)?.label || stage || "—");
  const getStatusColor = (status) => ({ draft: "accent", issued: "blue", part_paid: "orange", paid: "green", overdue: "red", disputed: "purple", write_off: "gray" }[status] || "accent");
  const getAgingBucket = (daysPastDue, balance) => {
    if ((Number(balance) || 0) <= 0) return "paid";
    if (daysPastDue <= 0) return "current";
    if (daysPastDue <= 30) return "1_30";
    if (daysPastDue <= 60) return "31_60";
    if (daysPastDue <= 90) return "61_90";
    return "90_plus";
  };
  const getAgingLabel = (bucket) => ({ current: "Current", "1_30": "1-30 Days", "31_60": "31-60 Days", "61_90": "61-90 Days", "90_plus": "90+ Days", paid: "Paid" }[bucket] || bucket);

  const buildBlankForm = () => ({
    id: uid(),
    mpoId: "",
    clientId: "",
    campaignId: "",
    invoiceNo: "",
    invoiceDate: isoToday(),
    dueDate: addDaysToIso(isoToday(), 30),
    grossAmount: "",
    status: "draft",
    collectionStage: "invoicing",
    owner: user?.name || "",
    notes: "",
    payments: [],
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
  });

  const receivableRows = (receivables || []).map(record => {
    const normalized = normalizeReceivableRecord(record);
    const campaign = liveCampaigns.find(item => item.id === normalized.campaignId) || liveCampaigns.find(item => item.id === liveMpos.find(mpo => mpo.id === normalized.mpoId)?.campaignId);
    const client = liveClients.find(item => item.id === normalized.clientId) || liveClients.find(item => item.id === campaign?.clientId);
    const mpo = liveMpos.find(item => item.id === normalized.mpoId);
    const daysPastDue = getDaysPastDue(normalized.dueDate, normalized.balance);
    const agingBucket = getAgingBucket(daysPastDue, normalized.balance);
    return {
      ...normalized,
      clientName: client?.name || mpo?.clientName || "—",
      campaignName: campaign?.name || mpo?.campaignName || "—",
      brand: campaign?.brand || mpo?.brand || "—",
      mpoNo: mpo?.mpoNo || "—",
      daysPastDue,
      agingBucket,
      collectionProgress: normalized.grossAmount > 0 ? (normalized.amountReceived / normalized.grossAmount) * 100 : 0,
      paymentCount: normalized.payments.length,
      latestPayment: normalized.payments[0] || null,
    };
  }).sort((a, b) => new Date(b.updatedAt || b.createdAt || 0) - new Date(a.updatedAt || a.createdAt || 0));

  const filteredRows = receivableRows.filter(row => {
    if (filters.clientId && row.clientId !== filters.clientId) return false;
    if (filters.status && row.status !== filters.status) return false;
    if (filters.dueBucket && row.agingBucket !== filters.dueBucket) return false;
    const term = filters.search.trim().toLowerCase();
    if (!term) return true;
    return [row.invoiceNo, row.clientName, row.campaignName, row.brand, row.mpoNo, row.owner, row.notes, row.status, row.collectionStage]
      .some(value => String(value || "").toLowerCase().includes(term));
  });

  const billingCandidates = liveMpos.filter(mpo => {
    const recon = String(mpo?.reconciliationStatus || "not_started").toLowerCase();
    const proof = String(mpo?.proofStatus || "pending").toLowerCase();
    return recon === "completed" && ["received", "approved"].includes(proof) && !receivableRows.some(row => row.mpoId === mpo.id);
  }).map(mpo => {
    const campaign = liveCampaigns.find(item => item.id === mpo.campaignId);
    const client = liveClients.find(item => item.id === campaign?.clientId);
    const amount = Number(mpo.reconciledAmount) || Number(mpo.invoiceAmount) || Number(mpo.grandTotal) || Number(mpo.netVal) || 0;
    return {
      id: mpo.id,
      mpoNo: mpo.mpoNo || "—",
      campaignId: campaign?.id || mpo.campaignId || "",
      campaignName: campaign?.name || mpo.campaignName || "—",
      clientId: client?.id || campaign?.clientId || "",
      clientName: client?.name || mpo.clientName || "—",
      vendorName: mpo.vendorName || "—",
      amount,
      mpo,
      campaign,
      client,
    };
  }).sort((a, b) => b.amount - a.amount);

  const totalGross = filteredRows.reduce((sum, row) => sum + row.grossAmount, 0);
  const totalReceived = filteredRows.reduce((sum, row) => sum + row.amountReceived, 0);
  const totalBalance = filteredRows.reduce((sum, row) => sum + row.balance, 0);
  const overdueBalance = filteredRows.filter(row => row.agingBucket !== "current" && row.agingBucket !== "paid").reduce((sum, row) => sum + row.balance, 0);
  const paidValue = filteredRows.filter(row => row.status === "paid").reduce((sum, row) => sum + row.grossAmount, 0);
  const collectionRate = totalGross > 0 ? (totalReceived / totalGross) * 100 : 0;
  const overdueCount = filteredRows.filter(row => row.status === "overdue").length;
  const disputedCount = filteredRows.filter(row => row.status === "disputed").length;
  const agingSummary = ["current", "1_30", "31_60", "61_90", "90_plus"].map(bucket => ({
    bucket,
    label: getAgingLabel(bucket),
    balance: filteredRows.filter(row => row.agingBucket === bucket).reduce((sum, row) => sum + row.balance, 0),
    count: filteredRows.filter(row => row.agingBucket === bucket).length,
  }));

  const exportView = (title, headers, rows) => {
    if (!rows.length) {
      setToast({ msg: `No data to export for ${title}.`, type: "error" });
      return;
    }
    const esc = (value) => String(value ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
    const htmlRows = rows.map(row => `<tr>${row.map((cell, index) => `<td style="padding:6px 9px;border:1px solid #ddd;font-size:10px;${index===0?"font-weight:600":""}">${esc(cell)}</td>`).join("")}</tr>`).join("");
    const html = `<!DOCTYPE html><html><head><meta charset="utf-8"><title>${title}</title>
      <style>
        body{font-family:Arial,sans-serif;padding:24px;color:#111;margin:0}
        h1{font-size:18px;margin-bottom:4px;color:#0A1F44}
        p{font-size:11px;color:#555;margin-bottom:16px}
        table{width:100%;border-collapse:collapse;margin-top:8px}
        th{background:#0A1F44;color:#fff;padding:7px 9px;font-size:10px;text-align:left;border:1px solid #0A1F44}
        tr:nth-child(even){background:#F5F7FA}
      </style>
    </head><body>
      <h1>${title}</h1>
      <p>Generated ${new Date().toLocaleString("en-NG")}</p>
      <table><tr>${headers.map(header => `<th>${esc(header)}</th>`).join("")}</tr>${htmlRows}</table>
    </body></html>`;
    const csv = buildCSV(rows, headers);
    setPreview({ html, csv, title });
  };

  const exportStatement = (row) => {
    const paymentRows = row.payments.length
      ? row.payments.map(payment => `<tr><td>${formatIsoDate(payment.receivedAt)}</td><td>${payment.reference || "—"}</td><td>${payment.channel || "—"}</td><td>${fmtN(payment.amount)}</td><td>${payment.note || "—"}</td></tr>`).join("")
      : `<tr><td colspan="5" style="text-align:center;color:#667085">No payment logged yet.</td></tr>`;
    const html = `<!DOCTYPE html><html><head><meta charset="utf-8"><title>Receivable Statement - ${row.invoiceNo}</title>
      <style>
        body{font-family:Arial,sans-serif;padding:28px;color:#111;margin:0;background:#fff}
        h1{font-size:22px;margin:0 0 4px;color:#0A1F44}
        h2{font-size:14px;margin:18px 0 10px;color:#0A1F44;text-transform:uppercase;letter-spacing:.08em}
        .muted{font-size:12px;color:#666}
        .grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:12px;margin:18px 0}
        .card{border:1px solid #d6dde8;border-radius:12px;padding:14px 16px;background:#f8fafc}
        .label{font-size:10px;color:#667085;text-transform:uppercase;letter-spacing:.08em}
        .value{font-size:16px;font-weight:700;margin-top:6px}
        table{width:100%;border-collapse:collapse;margin-top:8px}
        th,td{border:1px solid #d6dde8;padding:8px 10px;text-align:left;font-size:11px}
        th{background:#0A1F44;color:#fff}
      </style>
    </head><body>
      <h1>Client Receivable Statement</h1>
      <div class="muted">Generated ${new Date().toLocaleString("en-NG")}</div>
      <div class="grid">
        <div class="card"><div class="label">Client</div><div class="value">${row.clientName}</div></div>
        <div class="card"><div class="label">Invoice No.</div><div class="value">${row.invoiceNo}</div></div>
        <div class="card"><div class="label">Invoice / Due</div><div class="value">${formatIsoDate(row.invoiceDate)} / ${formatIsoDate(row.dueDate)}</div></div>
        <div class="card"><div class="label">Status</div><div class="value">${getStatusLabel(row.status)}</div></div>
      </div>
      <div class="grid">
        <div class="card"><div class="label">Gross Amount</div><div class="value">${fmtN(row.grossAmount)}</div></div>
        <div class="card"><div class="label">Received</div><div class="value">${fmtN(row.amountReceived)}</div></div>
        <div class="card"><div class="label">Outstanding</div><div class="value">${fmtN(row.balance)}</div></div>
        <div class="card"><div class="label">Collection Stage</div><div class="value">${getStageLabel(row.collectionStage)}</div></div>
      </div>
      <h2>Context</h2>
      <table>
        <tr><th>MPO</th><td>${row.mpoNo}</td><th>Campaign</th><td>${row.campaignName}</td></tr>
        <tr><th>Brand</th><td>${row.brand}</td><th>Owner</th><td>${row.owner || "—"}</td></tr>
        <tr><th>Last Follow-up</th><td>${formatIsoDate(row.lastFollowUpAt)}</td><th>Notes</th><td>${row.notes || "—"}</td></tr>
      </table>
      <h2>Payment History</h2>
      <table>
        <tr><th>Date</th><th>Reference</th><th>Channel</th><th>Amount</th><th>Note</th></tr>
        ${paymentRows}
      </table>
    </body></html>`;
    const csvRows = row.payments.length
      ? row.payments.map(payment => [row.invoiceNo, row.clientName, formatIsoDate(payment.receivedAt), payment.reference, payment.channel, payment.amount, payment.note])
      : [[row.invoiceNo, row.clientName, "", "", "", 0, "No payment logged yet"]];
    setPreview({ html, csv: buildCSV(csvRows, ["Invoice No.", "Client", "Payment Date", "Reference", "Channel", "Amount", "Note"]), title: `Statement - ${row.invoiceNo}` });
  };

  const openManualForm = () => setFormModal(buildBlankForm());
  const openCandidateForm = (candidate) => {
    const next = buildReceivableFromMpo({ mpo: candidate.mpo, campaign: candidate.campaign, client: candidate.client, owner: user?.name || "" });
    setFormModal({ ...next, grossAmount: String(next.grossAmount) });
  };
  const openEditForm = (row) => setFormModal({ ...row, grossAmount: String(row.grossAmount || 0) });

  const handleFormMpoChange = (mpoId) => {
    const mpo = liveMpos.find(item => item.id === mpoId);
    if (!mpo) return setFormModal(prev => ({ ...prev, mpoId: "" }));
    const campaign = liveCampaigns.find(item => item.id === mpo.campaignId);
    const client = liveClients.find(item => item.id === campaign?.clientId);
    const grossAmount = Number(mpo.reconciledAmount) || Number(mpo.invoiceAmount) || Number(mpo.grandTotal) || Number(mpo.netVal) || 0;
    setFormModal(prev => ({
      ...prev,
      mpoId,
      clientId: client?.id || prev.clientId,
      campaignId: campaign?.id || prev.campaignId,
      invoiceNo: prev.invoiceNo || mpo.invoiceNo || `AR-${String(mpo.mpoNo || uid()).replace(/\s+/g, "-")}`,
      grossAmount: grossAmount ? String(grossAmount) : prev.grossAmount,
      notes: prev.notes || `Created from ${mpo.mpoNo || "MPO"}`,
      status: prev.status === "draft" ? "issued" : prev.status,
    }));
  };

  const saveReceivable = async () => {
    if (!formModal || !onSaveReceivable) return;
    const grossAmount = Number(formModal.grossAmount) || 0;
    if (!formModal.clientId) return setToast({ msg: "Client is required.", type: "error" });
    if (!formModal.invoiceNo.trim()) return setToast({ msg: "Invoice number is required.", type: "error" });
    if (!formModal.invoiceDate || !formModal.dueDate) return setToast({ msg: "Invoice and due dates are required.", type: "error" });
    if (grossAmount <= 0) return setToast({ msg: "Gross amount must be greater than zero.", type: "error" });
    const existing = receivables.find(item => item.id === formModal.id);
    const record = normalizeReceivableRecord({
      ...(existing || {}),
      ...formModal,
      grossAmount,
      owner: formModal.owner || user?.name || "",
      payments: existing?.payments || formModal.payments || [],
      createdAt: existing?.createdAt || formModal.createdAt || new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    });
    try {
      setFormBusy(true);
      await onSaveReceivable(record);
      setFormModal(null);
      setToast({ msg: existing ? "Receivable updated." : "Receivable created.", type: "success" });
    } catch (error) {
      console.error("Failed to save receivable:", error);
      setToast({ msg: error?.message || "Failed to save receivable.", type: "error" });
    } finally {
      setFormBusy(false);
    }
  };

  const removeReceivable = async (recordId) => {
    if (!onRemoveReceivable) return;
    try {
      await onRemoveReceivable(recordId);
      setToast({ msg: "Receivable removed.", type: "success" });
    } catch (error) {
      console.error("Failed to remove receivable:", error);
      setToast({ msg: error?.message || "Failed to remove receivable.", type: "error" });
    }
  };

  const openPaymentEntry = (row) => {
    if (row.balance <= 0) return setToast({ msg: "This receivable is already fully collected.", type: "error" });
    setPaymentModal({ receivableId: row.id, amount: String(row.balance), receivedAt: isoToday(), reference: "", channel: "bank_transfer", note: "" });
  };

  const savePayment = async () => {
    if (!paymentModal || !onLogReceivablePayment) return;
    const amount = Number(paymentModal.amount) || 0;
    if (amount <= 0) return setToast({ msg: "Payment amount must be greater than zero.", type: "error" });
    const record = receivables.find(item => item.id === paymentModal.receivableId);
    if (!record) return setToast({ msg: "Receivable not found.", type: "error" });
    const normalized = normalizeReceivableRecord(record);
    if (amount > normalized.balance) return setToast({ msg: "Payment cannot exceed the outstanding balance.", type: "error" });
    const payment = normalizePaymentEntry({
      amount,
      receivedAt: paymentModal.receivedAt || isoToday(),
      reference: paymentModal.reference || "",
      channel: paymentModal.channel || "bank_transfer",
      note: paymentModal.note || "",
    });
    try {
      setPaymentBusy(true);
      await onLogReceivablePayment(normalized.id, payment);
      setPaymentModal(null);
      setToast({ msg: "Payment logged successfully.", type: "success" });
    } catch (error) {
      console.error("Failed to log payment:", error);
      setToast({ msg: error?.message || "Failed to log payment.", type: "error" });
    } finally {
      setPaymentBusy(false);
    }
  };

  const updateQuickStatus = async (row, status) => {
    if (!onUpdateReceivableStatus) return;
    try {
      await onUpdateReceivableStatus(row.id, {
        status,
        collectionStage: status === "paid" ? "resolved" : row.collectionStage,
        lastFollowUpAt: status === row.status ? row.lastFollowUpAt : isoToday(),
      });
      setToast({ msg: `Status updated to ${getStatusLabel(status)}.`, type: "success" });
    } catch (error) {
      console.error("Failed to update receivable status:", error);
      setToast({ msg: error?.message || "Failed to update status.", type: "error" });
    }
  };

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 16, flexWrap: "wrap" }}>
        <div>
          <h1 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 800, fontSize: 24 }}>Client Receivables & Collections</h1>
          <p style={{ marginTop: 6, color: "var(--text3)", maxWidth: 860, lineHeight: 1.6 }}>Move collections out of inference and into a real ledger. Create invoice records, log partial payments, track overdue balances, and export client-ready statements from one workspace.</p>
        </div>
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
          <Btn variant="secondary" onClick={openManualForm}>New Manual Invoice</Btn>
          <Btn variant="ghost" onClick={() => onOpenCloseout && onOpenCloseout()}>Open Closeout</Btn>
          <Btn variant="blue" icon="⬇" onClick={() => exportView("Receivables Ledger", ["Invoice No.", "Client", "Campaign", "MPO", "Invoice Date", "Due Date", "Status", "Collection Stage", "Gross Amount", "Received", "Outstanding", "Days Past Due", "Owner", "Notes"], filteredRows.map(row => [row.invoiceNo, row.clientName, row.campaignName, row.mpoNo, row.invoiceDate, row.dueDate, getStatusLabel(row.status), getStageLabel(row.collectionStage), row.grossAmount, row.amountReceived, row.balance, row.daysPastDue, row.owner, row.notes]))}>Export Ledger</Btn>
        </div>
      </div>

      <Card style={{ padding: 18 }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, marginBottom: 14, flexWrap: "wrap" }}>
          <h2 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 16 }}>Ledger Filters</h2>
          <div style={{ fontSize: 12, color: "var(--text3)" }}>{receivablesMeta?.mode === "supabase" ? "Live Supabase mode with realtime sync across invoices and payment logs." : receivablesMeta?.message || "Workspace local backup mode is active for receivables."}</div>
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(4,minmax(0,1fr))", gap: 12 }}>
          <Field label="Client" value={filters.clientId} onChange={updateFilter("clientId")} options={liveClients.map(client => ({ value: client.id, label: client.name }))} placeholder="All Clients" />
          <Field label="Status" value={filters.status} onChange={updateFilter("status")} options={statusOptions} placeholder="All Statuses" />
          <Field label="Aging Bucket" value={filters.dueBucket} onChange={updateFilter("dueBucket")} options={[{ value: "current", label: "Current" }, { value: "1_30", label: "1-30 Days" }, { value: "31_60", label: "31-60 Days" }, { value: "61_90", label: "61-90 Days" }, { value: "90_plus", label: "90+ Days" }, { value: "paid", label: "Paid" }]} placeholder="All Buckets" />
          <Field label="Search" value={filters.search} onChange={updateFilter("search")} placeholder="Invoice no., client, MPO, note…" />
        </div>
      </Card>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(6,minmax(0,1fr))", gap: 14 }}>
        <Stat icon="🧾" label="Gross Ledger" value={fmtN(totalGross)} sub={`${filteredRows.length} invoices`} color="var(--accent)" valueSize="clamp(16px, 1.7vw, 22px)" />
        <Stat icon="💳" label="Collected" value={fmtN(totalReceived)} sub={`${collectionRate.toFixed(1)}% collection`} color="var(--green)" valueSize="clamp(16px, 1.7vw, 22px)" />
        <Stat icon="⏳" label="Outstanding" value={fmtN(totalBalance)} sub={`${filteredRows.filter(row => row.balance > 0).length} open`} color="var(--orange)" valueSize="clamp(16px, 1.7vw, 22px)" />
        <Stat icon="🚨" label="Overdue" value={fmtN(overdueBalance)} sub={`${overdueCount} overdue`} color="var(--red)" valueSize="clamp(16px, 1.7vw, 22px)" />
        <Stat icon="✅" label="Paid Value" value={fmtN(paidValue)} sub={`${filteredRows.filter(row => row.status === "paid").length} settled`} color="var(--blue)" valueSize="clamp(16px, 1.7vw, 22px)" />
        <Stat icon="⚠️" label="Disputes" value={String(disputedCount)} sub="Need escalation" color="var(--purple)" valueSize="clamp(16px, 1.7vw, 22px)" />
      </div>

      <Card>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, marginBottom: 14 }}>
          <h2 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 15 }}>Aging Summary</h2>
          <Btn variant="ghost" size="sm" onClick={() => exportView("Receivable Aging Summary", ["Bucket", "Count", "Outstanding Balance"], agingSummary.map(row => [row.label, row.count, row.balance]))}>Export Aging</Btn>
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(5,minmax(0,1fr))", gap: 12 }}>
          {agingSummary.map(bucket => (
            <div key={bucket.bucket} style={{ padding: 14, borderRadius: 12, background: "var(--bg3)", border: "1px solid var(--border)" }}>
              <div style={{ fontSize: 10, color: "var(--text3)", textTransform: "uppercase", letterSpacing: ".08em" }}>{bucket.label}</div>
              <div style={{ fontSize: 18, fontWeight: 800, marginTop: 8 }}>{fmtN(bucket.balance)}</div>
              <div style={{ fontSize: 12, color: "var(--text3)", marginTop: 6 }}>{bucket.count} invoice{bucket.count === 1 ? "" : "s"}</div>
            </div>
          ))}
        </div>
      </Card>

      <div style={{ display: "grid", gridTemplateColumns: "1.5fr .9fr", gap: 18, alignItems: "start" }}>
        <Card>
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, marginBottom: 14 }}>
            <h2 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 15 }}>Receivables Ledger</h2>
            <Btn variant="ghost" size="sm" onClick={() => exportView("Collections Watchlist", ["Invoice No.", "Client", "Status", "Stage", "Due Date", "Days Past Due", "Outstanding", "Owner", "Last Follow-up"], filteredRows.filter(row => row.balance > 0).map(row => [row.invoiceNo, row.clientName, getStatusLabel(row.status), getStageLabel(row.collectionStage), row.dueDate, row.daysPastDue, row.balance, row.owner, row.lastFollowUpAt]))}>Export Watchlist</Btn>
          </div>
          {filteredRows.length === 0 ? <Empty icon="💵" title="No receivable records yet" sub="Create a manual invoice or promote reconciled closeout items into the ledger." /> : (
            <div style={{ overflowX: "auto" }}>
              <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 1240 }}>
                <thead>
                  <tr style={{ background: "var(--bg3)" }}>
                    {["Invoice / Client", "Campaign / MPO", "Amounts", "Status", "Aging", "Payments", "Actions"].map(header => <th key={header} style={{ padding: "8px 10px", textAlign: "left", fontSize: 10, fontWeight: 600, textTransform: "uppercase", color: "var(--text3)", borderBottom: "1px solid var(--border)" }}>{header}</th>)}
                  </tr>
                </thead>
                <tbody>
                  {filteredRows.map((row, index) => (
                    <tr key={row.id} style={{ borderBottom: "1px solid var(--border)", background: index % 2 === 0 ? "transparent" : "rgba(255,255,255,.01)" }}>
                      <td style={{ padding: "9px 10px", fontSize: 12, fontWeight: 700 }}>
                        {row.invoiceNo}
                        <div style={{ fontSize: 11, color: "var(--text3)", fontWeight: 400, marginTop: 3 }}>{row.clientName}</div>
                        <div style={{ fontSize: 11, color: "var(--text3)", marginTop: 3 }}>Invoice {formatIsoDate(row.invoiceDate)} · Due {formatIsoDate(row.dueDate)}</div>
                      </td>
                      <td style={{ padding: "9px 10px", fontSize: 12 }}>
                        <div style={{ fontWeight: 700 }}>{row.campaignName}</div>
                        <div style={{ fontSize: 11, color: "var(--text3)", marginTop: 3 }}>{row.brand}</div>
                        <div style={{ fontSize: 11, color: "var(--text3)", marginTop: 3 }}>{row.mpoNo !== "—" ? row.mpoNo : "Manual record"}</div>
                      </td>
                      <td style={{ padding: "9px 10px", fontSize: 12 }}>
                        <div style={{ fontWeight: 700 }}>{fmtN(row.grossAmount)}</div>
                        <div style={{ fontSize: 11, color: "var(--text3)", marginTop: 3 }}>Received {fmtN(row.amountReceived)}</div>
                        <div style={{ fontSize: 11, color: row.balance > 0 ? "var(--orange)" : "var(--green)", marginTop: 3 }}>Outstanding {fmtN(row.balance)}</div>
                      </td>
                      <td style={{ padding: "9px 10px", fontSize: 12 }}>
                        <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                          <Badge color={getStatusColor(row.status)}>{getStatusLabel(row.status)}</Badge>
                          <Badge color={row.collectionStage === "resolved" ? "green" : row.collectionStage === "escalated" ? "red" : row.collectionStage === "promise_to_pay" ? "blue" : "accent"}>{getStageLabel(row.collectionStage)}</Badge>
                        </div>
                        <div style={{ fontSize: 11, color: "var(--text3)", marginTop: 5 }}>Owner: {row.owner || "—"}</div>
                      </td>
                      <td style={{ padding: "9px 10px", fontSize: 12 }}>
                        <div style={{ fontWeight: 700, color: row.daysPastDue > 0 ? "var(--red)" : "var(--green)" }}>{row.daysPastDue > 0 ? `${row.daysPastDue} days late` : getAgingLabel(row.agingBucket)}</div>
                        <div style={{ marginTop: 6, height: 6, background: "var(--bg4)", borderRadius: 999, overflow: "hidden" }}><div style={{ width: `${Math.min(row.collectionProgress, 100)}%`, height: "100%", background: row.balance <= 0 ? "var(--green)" : row.daysPastDue > 0 ? "var(--red)" : "var(--blue)" }} /></div>
                        <div style={{ fontSize: 11, color: "var(--text3)", marginTop: 5 }}>{row.collectionProgress.toFixed(1)}% collected</div>
                      </td>
                      <td style={{ padding: "9px 10px", fontSize: 12 }}>
                        <div style={{ fontWeight: 700 }}>{row.paymentCount} payment{row.paymentCount === 1 ? "" : "s"}</div>
                        <div style={{ fontSize: 11, color: "var(--text3)", marginTop: 3 }}>{row.latestPayment ? `${formatIsoDate(row.latestPayment.receivedAt)} · ${fmtN(row.latestPayment.amount)}` : "No payment logged"}</div>
                        <div style={{ fontSize: 11, color: "var(--text3)", marginTop: 3 }}>Follow-up {formatIsoDate(row.lastFollowUpAt)}</div>
                      </td>
                      <td style={{ padding: "9px 10px", fontSize: 12 }}>
                        <div style={{ display: "flex", flexDirection: "column", gap: 8, alignItems: "flex-start" }}>
                          <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                            <Btn variant="ghost" size="sm" onClick={() => openEditForm(row)}>Edit</Btn>
                            <Btn variant="secondary" size="sm" onClick={() => openPaymentEntry(row)}>Log Payment</Btn>
                            <Btn variant="ghost" size="sm" onClick={() => exportStatement(row)}>Statement</Btn>
                          </div>
                          <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                            {row.status !== "paid" ? <Btn variant="success" size="sm" onClick={() => updateQuickStatus(row, "paid")}>Mark Paid</Btn> : null}
                            {row.status !== "disputed" ? <Btn variant="purple" size="sm" onClick={() => updateQuickStatus(row, "disputed")}>Dispute</Btn> : null}
                            <Btn variant="danger" size="sm" onClick={() => removeReceivable(row.id)}>Delete</Btn>
                          </div>
                        </div>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </Card>

        <div style={{ display: "flex", flexDirection: "column", gap: 18 }}>
          <Card>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, marginBottom: 14 }}>
              <h2 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 15 }}>Ready to Raise Invoice</h2>
              <Btn variant="ghost" size="sm" onClick={() => onOpenCloseout && onOpenCloseout()}>Review Closeout</Btn>
            </div>
            {billingCandidates.length === 0 ? <Empty icon="📦" title="No billing candidates" sub="Completed reconciliation items without a ledger record will surface here." /> : (
              <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                {billingCandidates.slice(0, 6).map(candidate => (
                  <div key={candidate.id} style={{ padding: "11px 12px", borderRadius: 10, background: "var(--bg3)", border: "1px solid var(--border)" }}>
                    <div style={{ display: "flex", justifyContent: "space-between", gap: 10, alignItems: "center" }}>
                      <div>
                        <div style={{ fontSize: 13, fontWeight: 700 }}>{candidate.clientName}</div>
                        <div style={{ fontSize: 11, color: "var(--text3)", marginTop: 2 }}>{candidate.mpoNo} · {candidate.campaignName}</div>
                      </div>
                      <Badge color="green">Ready</Badge>
                    </div>
                    <div style={{ display: "flex", justifyContent: "space-between", gap: 10, alignItems: "center", marginTop: 10 }}>
                      <div>
                        <div style={{ fontSize: 11, color: "var(--text3)" }}>Billing value</div>
                        <div style={{ fontSize: 14, fontWeight: 800, marginTop: 3 }}>{fmtN(candidate.amount)}</div>
                      </div>
                      <Btn variant="secondary" size="sm" onClick={() => openCandidateForm(candidate)}>Create AR</Btn>
                    </div>
                  </div>
                ))}
              </div>
            )}
          </Card>

          <Card>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, marginBottom: 14 }}>
              <h2 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 15 }}>Overdue Focus</h2>
              <Btn variant="ghost" size="sm" onClick={() => exportView("Overdue Receivables", ["Invoice No.", "Client", "Campaign", "Due Date", "Days Past Due", "Outstanding", "Owner", "Collection Stage"], filteredRows.filter(row => row.daysPastDue > 0 && row.balance > 0).map(row => [row.invoiceNo, row.clientName, row.campaignName, row.dueDate, row.daysPastDue, row.balance, row.owner, getStageLabel(row.collectionStage)]))}>Export Overdue</Btn>
            </div>
            {filteredRows.filter(row => row.daysPastDue > 0 && row.balance > 0).length === 0 ? <Empty icon="🕒" title="No overdue invoices" sub="Current unpaid items will roll in here once they pass due date." /> : (
              <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                {filteredRows.filter(row => row.daysPastDue > 0 && row.balance > 0).slice(0, 6).map(row => (
                  <div key={row.id} style={{ padding: "11px 12px", borderRadius: 10, background: "var(--bg3)", border: "1px solid var(--border)" }}>
                    <div style={{ display: "flex", justifyContent: "space-between", gap: 10, alignItems: "center" }}>
                      <div>
                        <div style={{ fontSize: 13, fontWeight: 700 }}>{row.invoiceNo}</div>
                        <div style={{ fontSize: 11, color: "var(--text3)", marginTop: 2 }}>{row.clientName} · {row.campaignName}</div>
                      </div>
                      <Badge color="red">{row.daysPastDue}d late</Badge>
                    </div>
                    <div style={{ marginTop: 10, display: "grid", gridTemplateColumns: "repeat(2,minmax(0,1fr))", gap: 8 }}>
                      <div style={{ background: "var(--bg2)", borderRadius: 8, padding: "8px 10px" }}><div style={{ fontSize: 10, color: "var(--text3)", textTransform: "uppercase" }}>Outstanding</div><div style={{ fontSize: 12, fontWeight: 700, marginTop: 4 }}>{fmtN(row.balance)}</div></div>
                      <div style={{ background: "var(--bg2)", borderRadius: 8, padding: "8px 10px" }}><div style={{ fontSize: 10, color: "var(--text3)", textTransform: "uppercase" }}>Owner / Stage</div><div style={{ fontSize: 12, fontWeight: 700, marginTop: 4 }}>{row.owner || "—"} · {getStageLabel(row.collectionStage)}</div></div>
                    </div>
                  </div>
                ))}
              </div>
            )}
          </Card>
        </div>
      </div>

      {formModal && (
        <Modal title={receivables.some(item => item.id === formModal.id) ? "Edit Receivable" : "New Receivable"} onClose={() => setFormModal(null)} width={720}>
          <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(2,minmax(0,1fr))", gap: 12 }}>
              <Field label="Source MPO" value={formModal.mpoId || ""} onChange={handleFormMpoChange} options={liveMpos.map(mpo => ({ value: mpo.id, label: `${mpo.mpoNo || "MPO"} · ${mpo.clientName || mpo.campaignName || mpo.vendorName || "—"}` }))} placeholder="Optional" />
              <Field label="Client" value={formModal.clientId} onChange={value => setFormModal(prev => ({ ...prev, clientId: value }))} options={liveClients.map(client => ({ value: client.id, label: client.name }))} placeholder="Select Client" />
              <Field label="Campaign" value={formModal.campaignId} onChange={value => setFormModal(prev => ({ ...prev, campaignId: value }))} options={liveCampaigns.filter(campaign => !formModal.clientId || campaign.clientId === formModal.clientId).map(campaign => ({ value: campaign.id, label: campaign.name }))} placeholder="Optional" />
              <Field label="Invoice No." value={formModal.invoiceNo} onChange={value => setFormModal(prev => ({ ...prev, invoiceNo: value }))} placeholder="e.g. INV-2026-014" />
              <Field label="Invoice Date" type="date" value={formModal.invoiceDate} onChange={value => setFormModal(prev => ({ ...prev, invoiceDate: value, dueDate: prev.dueDate || addDaysToIso(value, 30) }))} />
              <Field label="Due Date" type="date" value={formModal.dueDate} onChange={value => setFormModal(prev => ({ ...prev, dueDate: value }))} />
              <Field label="Gross Amount" type="number" value={formModal.grossAmount} onChange={value => setFormModal(prev => ({ ...prev, grossAmount: value }))} />
              <Field label="Status" value={formModal.status} onChange={value => setFormModal(prev => ({ ...prev, status: value }))} options={statusOptions} />
              <Field label="Collection Stage" value={formModal.collectionStage} onChange={value => setFormModal(prev => ({ ...prev, collectionStage: value }))} options={stageOptions} />
              <Field label="Owner" value={formModal.owner} onChange={value => setFormModal(prev => ({ ...prev, owner: value }))} placeholder="Collection owner" />
              <div style={{ gridColumn: "1 / -1" }}>
                <Field label="Notes" value={formModal.notes} onChange={value => setFormModal(prev => ({ ...prev, notes: value }))} rows={4} placeholder="Add invoice or collection notes…" />
              </div>
            </div>
            <div style={{ display: "flex", justifyContent: "flex-end", gap: 10 }}>
              <Btn variant="ghost" onClick={() => setFormModal(null)}>Cancel</Btn>
              <Btn onClick={saveReceivable} loading={formBusy}>Save Receivable</Btn>
            </div>
          </div>
        </Modal>
      )}

      {paymentModal && (
        <Modal title="Log Payment" onClose={() => setPaymentModal(null)} width={560}>
          <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(2,minmax(0,1fr))", gap: 12 }}>
              <Field label="Amount Received" type="number" value={paymentModal.amount} onChange={value => setPaymentModal(prev => ({ ...prev, amount: value }))} />
              <Field label="Payment Date" type="date" value={paymentModal.receivedAt} onChange={value => setPaymentModal(prev => ({ ...prev, receivedAt: value }))} />
              <Field label="Reference" value={paymentModal.reference} onChange={value => setPaymentModal(prev => ({ ...prev, reference: value }))} placeholder="Transfer ref / receipt no." />
              <Field label="Channel" value={paymentModal.channel} onChange={value => setPaymentModal(prev => ({ ...prev, channel: value }))} options={paymentChannelOptions} />
              <div style={{ gridColumn: "1 / -1" }}>
                <Field label="Note" value={paymentModal.note} onChange={value => setPaymentModal(prev => ({ ...prev, note: value }))} rows={3} placeholder="Optional note for the payment entry" />
              </div>
            </div>
            <div style={{ display: "flex", justifyContent: "flex-end", gap: 10 }}>
              <Btn variant="ghost" onClick={() => setPaymentModal(null)}>Cancel</Btn>
              <Btn onClick={savePayment} loading={paymentBusy}>Save Payment</Btn>
            </div>
          </div>
        </Modal>
      )}

      {preview && <PrintPreview html={preview.html} csv={preview.csv} pdfBytes={preview.pdfBytes} title={preview.title} onClose={() => setPreview(null)} />}
      {toast && <Toast {...toast} onDone={() => setToast(null)} />}
    </div>
  );
};

/* ── MAIN APP ───────────────────────────────────────────── */

export default ReceivablesPage;
