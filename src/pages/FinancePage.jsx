import { useState } from "react";
import Badge from "../components/Badge";
import Empty from "../components/Empty";
import Toast from "../components/Toast";
import Modal from "../components/Modal";
import PrintPreview from "../components/mpo/PrintPreview";
import { Btn, Field, Card, Stat } from "../components/ui/primitives";
import { activeOnly } from "../utils/records";
import { formatNairaExportValue, formatExportRowsWithCurrency } from "../utils/export";
import { fmt, fmtN } from "../utils/formatters";
import { MPO_PAYMENT_STATUS_OPTIONS, MPO_RECON_STATUS_OPTIONS, MPO_PROOF_STATUS_OPTIONS } from "../constants/mpoWorkflow";
import { getDaysPastDue, normalizeReceivableRecord, buildReceivableFromMpo } from "../services/receivables";

const uid = () => Math.random().toString(36).slice(2) + Date.now().toString(36);
const isoToday = () => new Date().toISOString().slice(0, 10);
const addDaysToIso = (isoDate, days = 0) => {
  const base = isoDate ? new Date(`${isoDate}T00:00:00`) : new Date();
  if (Number.isNaN(base.getTime())) return isoToday();
  base.setDate(base.getDate() + (Number(days) || 0));
  return base.toISOString().slice(0, 10);
};
const formatIsoDate = (value) => {
  if (!value) return "—";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return String(value);
  return date.toLocaleDateString("en-NG", { day: "numeric", month: "short", year: "numeric" });
};
const buildCSV = (rows, headers) => {
  const esc = (v) => `"${String(v ?? "").replace(/"/g, '""')}"`;
  return [headers.map(esc).join(","), ...rows.map((r) => r.map(esc).join(","))] .join("\n");
};

const BudgetingPage = ({ vendors, clients, campaigns, mpos }) => {
  const [toast, setToast] = useState(null);
  const [preview, setPreview] = useState(null);
  const [budgetFilters, setBudgetFilters] = useState({ clientId: "", status: "", paymentStatus: "", search: "" });

  const liveClients = activeOnly(clients);
  const liveCampaigns = activeOnly(campaigns);
  const liveMpos = activeOnly(mpos);

  const updateBudgetFilter = (key) => (value) => setBudgetFilters(prev => ({ ...prev, [key]: value }));

  const filteredCampaigns = liveCampaigns.filter(campaign => {
    if (budgetFilters.clientId && campaign.clientId !== budgetFilters.clientId) return false;
    if (budgetFilters.status && (campaign.status || "planning") !== budgetFilters.status) return false;
    const term = budgetFilters.search.trim().toLowerCase();
    if (!term) return true;
    const client = liveClients.find(row => row.id === campaign.clientId);
    return [campaign.name, campaign.brand, campaign.medium, campaign.status, client?.name].some(value => String(value || "").toLowerCase().includes(term));
  });

  const filteredCampaignIds = new Set(filteredCampaigns.map(campaign => campaign.id));

  const filteredMpos = liveMpos.filter(mpo => {
    if (budgetFilters.paymentStatus && (mpo.paymentStatus || "unpaid") !== budgetFilters.paymentStatus) return false;
    if (budgetFilters.clientId) {
      const campaign = liveCampaigns.find(row => row.id === mpo.campaignId);
      if ((campaign?.clientId || "") !== budgetFilters.clientId) return false;
    }
    if (filteredCampaignIds.size && mpo.campaignId && !filteredCampaignIds.has(mpo.campaignId)) return false;
    const term = budgetFilters.search.trim().toLowerCase();
    if (!term) return true;
    return [mpo.mpoNo, mpo.vendorName, mpo.clientName, mpo.brand, mpo.campaignName, mpo.invoiceNo, mpo.paymentReference].some(value => String(value || "").toLowerCase().includes(term));
  });

  const campaignFinanceRows = filteredCampaigns.map(campaign => {
    const client = liveClients.find(row => row.id === campaign.clientId);
    const campaignMpos = filteredMpos.filter(mpo => mpo.campaignId === campaign.id);
    const budget = parseFloat(campaign.budget) || 0;
    const committed = campaignMpos.reduce((sum, mpo) => sum + (parseFloat(mpo.grandTotal) || parseFloat(mpo.netVal) || 0), 0);
    const invoiced = campaignMpos.reduce((sum, mpo) => sum + (parseFloat(mpo.invoiceAmount) || parseFloat(mpo.grandTotal) || parseFloat(mpo.netVal) || 0), 0);
    const paid = campaignMpos.reduce((sum, mpo) => sum + ((mpo.paymentStatus || "unpaid") === "paid" ? (parseFloat(mpo.reconciledAmount) || parseFloat(mpo.invoiceAmount) || parseFloat(mpo.grandTotal) || parseFloat(mpo.netVal) || 0) : 0), 0);
    const outstanding = campaignMpos.reduce((sum, mpo) => sum + ((mpo.paymentStatus || "unpaid") !== "paid" ? (parseFloat(mpo.reconciledAmount) || parseFloat(mpo.invoiceAmount) || parseFloat(mpo.grandTotal) || parseFloat(mpo.netVal) || 0) : 0), 0);
    const utilization = budget > 0 ? (committed / budget) * 100 : 0;
    return {
      id: campaign.id,
      campaign: campaign.name || "Unnamed Campaign",
      client: client?.name || "—",
      brand: campaign.brand || "—",
      medium: campaign.medium || "—",
      status: campaign.status || "planning",
      budget,
      committed,
      available: budget - committed,
      invoiced,
      paid,
      outstanding,
      utilization,
      mpoCount: campaignMpos.length,
    };
  }).sort((a, b) => b.committed - a.committed);

  const clientBudgetRows = Object.values(campaignFinanceRows.reduce((acc, row) => {
    const key = row.client || "Unknown Client";
    acc[key] = acc[key] || { client: key, campaigns: 0, budget: 0, committed: 0, available: 0, invoiced: 0, paid: 0, outstanding: 0 };
    acc[key].campaigns += 1;
    acc[key].budget += row.budget;
    acc[key].committed += row.committed;
    acc[key].available += row.available;
    acc[key].invoiced += row.invoiced;
    acc[key].paid += row.paid;
    acc[key].outstanding += row.outstanding;
    return acc;
  }, {})).sort((a, b) => b.committed - a.committed);

  const today = Date.now();
  const invoiceAgingRows = filteredMpos.map(mpo => {
    const value = parseFloat(mpo.reconciledAmount) || parseFloat(mpo.invoiceAmount) || parseFloat(mpo.grandTotal) || parseFloat(mpo.netVal) || 0;
    const invoiceTs = mpo.invoiceReceivedAt ? new Date(mpo.invoiceReceivedAt).getTime() : (mpo.date ? new Date(mpo.date).getTime() : 0);
    const ageDays = invoiceTs ? Math.max(0, Math.floor((today - invoiceTs) / 86400000)) : 0;
    return {
      mpoNo: mpo.mpoNo || "—",
      vendor: mpo.vendorName || "—",
      client: mpo.clientName || "—",
      campaign: mpo.campaignName || "—",
      invoiceNo: mpo.invoiceNo || "—",
      invoiceStatus: mpo.invoiceStatus || "pending",
      paymentStatus: mpo.paymentStatus || "unpaid",
      paymentReference: mpo.paymentReference || "—",
      invoiceDate: mpo.invoiceReceivedAt ? new Date(mpo.invoiceReceivedAt).toLocaleDateString() : "—",
      paidAt: mpo.paidAt ? new Date(mpo.paidAt).toLocaleDateString() : "—",
      ageDays,
      value,
      outstanding: (mpo.paymentStatus || "unpaid") === "paid" ? 0 : value,
    };
  }).sort((a, b) => b.outstanding - a.outstanding || b.ageDays - a.ageDays);

  const monthlyCashflow = Object.values(filteredMpos.reduce((acc, mpo) => {
    const baseDate = mpo.invoiceReceivedAt || mpo.date || mpo.createdAt;
    const parsed = baseDate ? new Date(baseDate) : null;
    const monthKey = parsed && !Number.isNaN(parsed.getTime())
      ? `${parsed.getFullYear()}-${String(parsed.getMonth() + 1).padStart(2, "0")}`
      : "Undated";
    acc[monthKey] = acc[monthKey] || { period: monthKey, budget: 0, committed: 0, invoiced: 0, paid: 0, outstanding: 0 };
    const campaign = liveCampaigns.find(row => row.id === mpo.campaignId);
    const value = parseFloat(mpo.grandTotal) || parseFloat(mpo.netVal) || 0;
    const invoiced = parseFloat(mpo.invoiceAmount) || value;
    const paid = (mpo.paymentStatus || "unpaid") === "paid" ? (parseFloat(mpo.reconciledAmount) || invoiced || value) : 0;
    acc[monthKey].committed += value;
    acc[monthKey].invoiced += invoiced;
    acc[monthKey].paid += paid;
    acc[monthKey].outstanding += ((mpo.paymentStatus || "unpaid") === "paid" ? 0 : (parseFloat(mpo.reconciledAmount) || invoiced || value));
    acc[monthKey].budget += parseFloat(campaign?.budget) || 0;
    return acc;
  }, {})).sort((a, b) => a.period.localeCompare(b.period));

  const totalBudget = campaignFinanceRows.reduce((sum, row) => sum + row.budget, 0);
  const totalCommitted = campaignFinanceRows.reduce((sum, row) => sum + row.committed, 0);
  const totalInvoiced = campaignFinanceRows.reduce((sum, row) => sum + row.invoiced, 0);
  const totalPaid = campaignFinanceRows.reduce((sum, row) => sum + row.paid, 0);
  const totalOutstanding = campaignFinanceRows.reduce((sum, row) => sum + row.outstanding, 0);
  const availableBudget = totalBudget - totalCommitted;
  const overBudgetCount = campaignFinanceRows.filter(row => row.budget > 0 && row.committed > row.budget).length;
  const averageUtilization = campaignFinanceRows.length ? campaignFinanceRows.reduce((sum, row) => sum + Math.min(row.utilization, 100), 0) / campaignFinanceRows.length : 0;
  const invoiceReceivedCount = invoiceAgingRows.filter(row => row.invoiceStatus !== "pending").length;
  const paidCount = invoiceAgingRows.filter(row => row.paymentStatus === "paid").length;
  const outstandingCount = invoiceAgingRows.filter(row => row.outstanding > 0).length;

  const agingBuckets = [
    { label: "0-30 Days", test: days => days <= 30 },
    { label: "31-60 Days", test: days => days >= 31 && days <= 60 },
    { label: "61-90 Days", test: days => days >= 61 && days <= 90 },
    { label: "90+ Days", test: days => days > 90 },
  ].map(bucket => {
    const rows = invoiceAgingRows.filter(row => row.outstanding > 0 && bucket.test(row.ageDays));
    return { label: bucket.label, count: rows.length, value: rows.reduce((sum, row) => sum + row.outstanding, 0) };
  });

  const exportBudgetView = (title, headers, rows) => {
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

  const budgetInputStyle = {
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
        <h1 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 800, fontSize: 24 }}>Budgeting & Invoice / Payment Phase</h1>
        <p style={{ color: "var(--text2)", marginTop: 3, fontSize: 13 }}>Track budget allocation, MPO commitments, invoice exposure, and payment close-out from a single finance view.</p>
      </div>

      <Card style={{ marginBottom: 20 }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, flexWrap: "wrap", marginBottom: 16 }}>
          <div>
            <h2 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 16 }}>Finance Filters</h2>
            <p style={{ color: "var(--text2)", fontSize: 12, marginTop: 3 }}>Filter budget and billing insights by client, campaign status, payment stage, and keywords.</p>
          </div>
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
            <Btn variant="ghost" size="sm" onClick={() => setBudgetFilters({ clientId: "", status: "", paymentStatus: "", search: "" })}>Reset</Btn>
            <Btn variant="blue" size="sm" icon="⬇" onClick={() => exportBudgetView("Budgeting Summary", ["Metric", "Value"], [["Campaign Budget Pool", formatNairaExportValue(totalBudget)], ["Committed Spend", formatNairaExportValue(totalCommitted)], ["Invoiced", formatNairaExportValue(totalInvoiced)], ["Paid", formatNairaExportValue(totalPaid)], ["Outstanding", formatNairaExportValue(totalOutstanding)], ["Available Budget", formatNairaExportValue(availableBudget)], ["Over-budget Campaigns", overBudgetCount], ["Average Utilization", `${averageUtilization.toFixed(1)}%`]])}>Export Summary</Btn>
          </div>
        </div>

        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(180px,1fr))", gap: 12 }}>
          <Field label="Client" value={budgetFilters.clientId} onChange={updateBudgetFilter("clientId")} options={liveClients.map(client => ({ value: client.id, label: client.name }))} placeholder="All Clients" />
          <Field label="Campaign Status" value={budgetFilters.status} onChange={updateBudgetFilter("status")} options={[{ value: "planning", label: "Planning" }, { value: "active", label: "Active" }, { value: "paused", label: "Paused" }, { value: "completed", label: "Completed" }]} placeholder="All Statuses" />
          <Field label="Payment Status" value={budgetFilters.paymentStatus} onChange={updateBudgetFilter("paymentStatus")} options={MPO_PAYMENT_STATUS_OPTIONS} placeholder="All Payment States" />
          <div>
            <label style={{ fontSize: 11, fontWeight: 600, color: "var(--text3)", textTransform: "uppercase", letterSpacing: ".08em", marginBottom: 5, display: "block" }}>Search</label>
            <input value={budgetFilters.search} onChange={e => updateBudgetFilter("search")(e.target.value)} placeholder="Campaign, vendor, MPO, invoice..." style={budgetInputStyle} />
          </div>
        </div>
      </Card>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(180px,1fr))", gap: 14, marginBottom: 22 }}>
        <Stat icon="🎯" label="Budget Pool" value={fmtN(totalBudget)} sub={`${campaignFinanceRows.length} campaigns`} color="var(--accent)" valueSize="clamp(16px, 1.7vw, 22px)" />
        <Stat icon="🧾" label="Committed" value={fmtN(totalCommitted)} sub={`${invoiceReceivedCount} invoices logged`} color="var(--purple)" valueSize="clamp(16px, 1.7vw, 22px)" />
        <Stat icon="💳" label="Paid" value={fmtN(totalPaid)} sub={`${paidCount} MPOs paid`} color="var(--green)" valueSize="clamp(16px, 1.7vw, 22px)" />
        <Stat icon="⏳" label="Outstanding" value={fmtN(totalOutstanding)} sub={`${outstandingCount} items awaiting payment`} color="var(--red)" valueSize="clamp(16px, 1.7vw, 22px)" />
        <Stat icon="📉" label="Available Budget" value={fmtN(availableBudget)} sub={`${overBudgetCount} over budget`} color="var(--blue)" valueSize="clamp(16px, 1.7vw, 22px)" />
        <Stat icon="📈" label="Avg Utilization" value={`${averageUtilization.toFixed(1)}%`} sub="Across filtered campaigns" color="var(--teal)" valueSize="clamp(16px, 1.7vw, 22px)" />
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1.15fr .95fr", gap: 18, alignItems: "start" }}>
        <Card>
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, flexWrap: "wrap", marginBottom: 14 }}>
            <div>
              <h2 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 15 }}>Campaign Budget Tracker</h2>
              <p style={{ color: "var(--text2)", fontSize: 12, marginTop: 3 }}>Budget vs commitments, invoicing, and payment delivery per campaign.</p>
            </div>
            <Btn variant="ghost" size="sm" onClick={() => exportBudgetView("Campaign Budget Tracker", ["Campaign", "Client", "Brand", "Medium", "Status", "Budget", "Committed", "Available", "Invoiced", "Paid", "Outstanding", "Utilization %", "MPO Count"], formatExportRowsWithCurrency(campaignFinanceRows.map(row => [row.campaign, row.client, row.brand, row.medium, row.status, row.budget, row.committed, row.available, row.invoiced, row.paid, row.outstanding, `${row.utilization.toFixed(1)}%`, row.mpoCount]), [5, 6, 7, 8, 9, 10]))}>Export Tracker</Btn>
          </div>
          {campaignFinanceRows.length === 0 ? <Empty icon="🎯" title="No campaign budgets found" sub="Create campaigns and issue MPOs to populate this tracker." /> : (
            <div style={{ overflowX: "auto" }}>
              <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 860 }}>
                <thead>
                  <tr style={{ background: "var(--bg3)" }}>
                    {["Campaign", "Client", "Budget", "Committed", "Available", "Invoiced", "Paid", "Outstanding", "Utilization"].map(header => <th key={header} style={{ padding: "8px 10px", textAlign: "left", fontSize: 10, fontWeight: 600, textTransform: "uppercase", color: "var(--text3)", borderBottom: "1px solid var(--border)" }}>{header}</th>)}
                  </tr>
                </thead>
                <tbody>
                  {campaignFinanceRows.map((row, index) => {
                    const util = row.utilization > 100 ? "var(--red)" : row.utilization > 80 ? "var(--orange)" : "var(--green)";
                    return (
                      <tr key={row.id} style={{ borderBottom: "1px solid var(--border)", background: index % 2 === 0 ? "transparent" : "rgba(255,255,255,.01)" }}>
                        <td style={{ padding: "9px 10px", fontSize: 12, fontWeight: 700 }}>
                          {row.campaign}
                          <div style={{ fontSize: 11, color: "var(--text3)", fontWeight: 400, marginTop: 3 }}>{row.brand} · {row.medium}</div>
                        </td>
                        <td style={{ padding: "9px 10px", fontSize: 12, color: "var(--text2)" }}>{row.client}</td>
                        <td style={{ padding: "9px 10px", fontSize: 12 }}>{fmtN(row.budget)}</td>
                        <td style={{ padding: "9px 10px", fontSize: 12, color: "var(--accent)", fontWeight: 700 }}>{fmtN(row.committed)}</td>
                        <td style={{ padding: "9px 10px", fontSize: 12, color: row.available < 0 ? "var(--red)" : "var(--green)", fontWeight: 700 }}>{fmtN(row.available)}</td>
                        <td style={{ padding: "9px 10px", fontSize: 12 }}>{fmtN(row.invoiced)}</td>
                        <td style={{ padding: "9px 10px", fontSize: 12, color: "var(--green)", fontWeight: 700 }}>{fmtN(row.paid)}</td>
                        <td style={{ padding: "9px 10px", fontSize: 12, color: row.outstanding > 0 ? "var(--red)" : "var(--text2)", fontWeight: row.outstanding > 0 ? 700 : 400 }}>{fmtN(row.outstanding)}</td>
                        <td style={{ padding: "9px 10px", fontSize: 12 }}>
                          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                            <div style={{ flex: 1, minWidth: 84, height: 8, borderRadius: 999, background: "var(--bg4)", overflow: "hidden" }}>
                              <div style={{ width: `${Math.min(row.utilization, 100)}%`, height: "100%", background: util }} />
                            </div>
                            <span style={{ color: util, fontWeight: 700 }}>{row.utilization.toFixed(1)}%</span>
                          </div>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          )}
        </Card>

        <div style={{ display: "flex", flexDirection: "column", gap: 18 }}>
          <Card>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, marginBottom: 14 }}>
              <h2 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 15 }}>Invoice Aging</h2>
              <Btn variant="ghost" size="sm" onClick={() => exportBudgetView("Invoice Aging", ["Bucket", "Count", "Outstanding"], formatExportRowsWithCurrency(agingBuckets.map(row => [row.label, row.count, row.value]), [2]))}>Export Aging</Btn>
            </div>
            <div style={{ display: "grid", gap: 10 }}>
              {agingBuckets.map(bucket => (
                <div key={bucket.label} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, padding: "11px 12px", borderRadius: 10, background: "var(--bg3)", border: "1px solid var(--border)" }}>
                  <div>
                    <div style={{ fontSize: 13, fontWeight: 700 }}>{bucket.label}</div>
                    <div style={{ fontSize: 11, color: "var(--text3)", marginTop: 2 }}>{bucket.count} invoice{bucket.count !== 1 ? "s" : ""}</div>
                  </div>
                  <div style={{ fontSize: 13, fontWeight: 700, color: bucket.value > 0 ? "var(--red)" : "var(--text2)" }}>{fmtN(bucket.value)}</div>
                </div>
              ))}
            </div>
          </Card>

          <Card>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, marginBottom: 14 }}>
              <h2 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 15 }}>Client Budget Coverage</h2>
              <Btn variant="ghost" size="sm" onClick={() => exportBudgetView("Client Budget Coverage", ["Client", "Campaigns", "Budget", "Committed", "Available", "Invoiced", "Paid", "Outstanding"], formatExportRowsWithCurrency(clientBudgetRows.map(row => [row.client, row.campaigns, row.budget, row.committed, row.available, row.invoiced, row.paid, row.outstanding]), [2, 3, 4, 5, 6, 7]))}>Export Clients</Btn>
            </div>
            {clientBudgetRows.length === 0 ? <Empty icon="👥" title="No client budgets yet" sub="Link campaigns to clients to build the coverage summary." /> : (
              <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                {clientBudgetRows.slice(0, 6).map(row => (
                  <div key={row.client} style={{ padding: "11px 12px", borderRadius: 10, background: "var(--bg3)", border: "1px solid var(--border)" }}>
                    <div style={{ display: "flex", justifyContent: "space-between", gap: 10, alignItems: "center" }}>
                      <div>
                        <div style={{ fontSize: 13, fontWeight: 700 }}>{row.client}</div>
                        <div style={{ fontSize: 11, color: "var(--text3)", marginTop: 2 }}>{row.campaigns} campaign{row.campaigns !== 1 ? "s" : ""}</div>
                      </div>
                      <Badge color={row.outstanding > 0 ? "orange" : "green"}>{fmtN(row.outstanding)} outstanding</Badge>
                    </div>
                    <div style={{ display: "grid", gridTemplateColumns: "repeat(3,minmax(0,1fr))", gap: 8, marginTop: 10 }}>
                      <div style={{ background: "var(--bg2)", borderRadius: 8, padding: "8px 10px" }}><div style={{ fontSize: 10, color: "var(--text3)", textTransform: "uppercase" }}>Budget</div><div style={{ fontSize: 12, fontWeight: 700, marginTop: 4 }}>{fmtN(row.budget)}</div></div>
                      <div style={{ background: "var(--bg2)", borderRadius: 8, padding: "8px 10px" }}><div style={{ fontSize: 10, color: "var(--text3)", textTransform: "uppercase" }}>Committed</div><div style={{ fontSize: 12, fontWeight: 700, marginTop: 4 }}>{fmtN(row.committed)}</div></div>
                      <div style={{ background: "var(--bg2)", borderRadius: 8, padding: "8px 10px" }}><div style={{ fontSize: 10, color: "var(--text3)", textTransform: "uppercase" }}>Available</div><div style={{ fontSize: 12, fontWeight: 700, marginTop: 4, color: row.available < 0 ? "var(--red)" : "var(--green)" }}>{fmtN(row.available)}</div></div>
                    </div>
                  </div>
                ))}
              </div>
            )}
          </Card>
        </div>
      </div>

      <Card style={{ marginTop: 18, marginBottom: 18 }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, flexWrap: "wrap", marginBottom: 14 }}>
          <div>
            <h2 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 15 }}>Invoice & Payment Register</h2>
            <p style={{ color: "var(--text2)", fontSize: 12, marginTop: 3 }}>Monitor every MPO through invoice receipt, payment processing, and close-out.</p>
          </div>
          <Btn variant="ghost" size="sm" onClick={() => exportBudgetView("Invoice and Payment Register", ["MPO No.", "Vendor", "Client", "Campaign", "Invoice No.", "Invoice Date", "Invoice Status", "Payment Status", "Payment Ref", "Paid At", "Age (Days)", "Outstanding"], formatExportRowsWithCurrency(invoiceAgingRows.map(row => [row.mpoNo, row.vendor, row.client, row.campaign, row.invoiceNo, row.invoiceDate, row.invoiceStatus, row.paymentStatus, row.paymentReference, row.paidAt, row.ageDays, row.outstanding]), [11]))}>Export Register</Btn>
        </div>
        {invoiceAgingRows.length === 0 ? <Empty icon="💳" title="No invoice rows yet" sub="Record invoice and payment details in MPO execution to populate this register." /> : (
          <div style={{ overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 980 }}>
              <thead>
                <tr style={{ background: "var(--bg3)" }}>
                  {["MPO", "Vendor / Client", "Invoice", "Status", "Payment Ref", "Age", "Outstanding"].map(header => <th key={header} style={{ padding: "8px 10px", textAlign: "left", fontSize: 10, fontWeight: 600, textTransform: "uppercase", color: "var(--text3)", borderBottom: "1px solid var(--border)" }}>{header}</th>)}
                </tr>
              </thead>
              <tbody>
                {invoiceAgingRows.slice(0, 18).map((row, index) => (
                  <tr key={`${row.mpoNo}-${index}`} style={{ borderBottom: "1px solid var(--border)", background: index % 2 === 0 ? "transparent" : "rgba(255,255,255,.01)" }}>
                    <td style={{ padding: "9px 10px", fontSize: 12, fontWeight: 700 }}>{row.mpoNo}<div style={{ fontSize: 11, color: "var(--text3)", fontWeight: 400, marginTop: 3 }}>{row.campaign}</div></td>
                    <td style={{ padding: "9px 10px", fontSize: 12 }}>{row.vendor}<div style={{ fontSize: 11, color: "var(--text3)", marginTop: 3 }}>{row.client}</div></td>
                    <td style={{ padding: "9px 10px", fontSize: 12 }}>{row.invoiceNo}<div style={{ fontSize: 11, color: "var(--text3)", marginTop: 3 }}>{row.invoiceDate}</div></td>
                    <td style={{ padding: "9px 10px", fontSize: 12 }}><div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}><Badge color={row.invoiceStatus === "approved" ? "green" : row.invoiceStatus === "received" ? "blue" : row.invoiceStatus === "disputed" ? "red" : "accent"}>{row.invoiceStatus}</Badge><Badge color={row.paymentStatus === "paid" ? "green" : row.paymentStatus === "processing" ? "blue" : row.paymentStatus === "disputed" ? "red" : "orange"}>{row.paymentStatus}</Badge></div></td>
                    <td style={{ padding: "9px 10px", fontSize: 12 }}>{row.paymentReference}<div style={{ fontSize: 11, color: "var(--text3)", marginTop: 3 }}>{row.paidAt}</div></td>
                    <td style={{ padding: "9px 10px", fontSize: 12 }}>{row.ageDays}d</td>
                    <td style={{ padding: "9px 10px", fontSize: 12, fontWeight: 700, color: row.outstanding > 0 ? "var(--red)" : "var(--green)" }}>{fmtN(row.outstanding)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
            {invoiceAgingRows.length > 18 ? <div style={{ marginTop: 10, fontSize: 12, color: "var(--text3)" }}>Showing 18 of {invoiceAgingRows.length} rows in-app. Use export for the full register.</div> : null}
          </div>
        )}
      </Card>

      <Card>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, flexWrap: "wrap", marginBottom: 14 }}>
          <div>
            <h2 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 15 }}>Monthly Budget vs Cashflow</h2>
            <p style={{ color: "var(--text2)", fontSize: 12, marginTop: 3 }}>Watch commitments, invoicing, and cash outflow by period.</p>
          </div>
          <Btn variant="ghost" size="sm" onClick={() => exportBudgetView("Monthly Budget vs Cashflow", ["Period", "Budget", "Committed", "Invoiced", "Paid", "Outstanding"], formatExportRowsWithCurrency(monthlyCashflow.map(row => [row.period, row.budget, row.committed, row.invoiced, row.paid, row.outstanding]), [1, 2, 3, 4, 5]))}>Export Cashflow</Btn>
        </div>
        {monthlyCashflow.length === 0 ? <Empty icon="📅" title="No monthly cashflow yet" sub="Invoices and MPOs will populate the monthly cashflow view." /> : (
          <div style={{ overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 760 }}>
              <thead>
                <tr style={{ background: "var(--bg3)" }}>
                  {["Period", "Budget", "Committed", "Invoiced", "Paid", "Outstanding"].map(header => <th key={header} style={{ padding: "8px 10px", textAlign: "left", fontSize: 10, fontWeight: 600, textTransform: "uppercase", color: "var(--text3)", borderBottom: "1px solid var(--border)" }}>{header}</th>)}
                </tr>
              </thead>
              <tbody>
                {monthlyCashflow.map((row, index) => (
                  <tr key={row.period} style={{ borderBottom: "1px solid var(--border)", background: index % 2 === 0 ? "transparent" : "rgba(255,255,255,.01)" }}>
                    <td style={{ padding: "9px 10px", fontSize: 12, fontWeight: 700 }}>{row.period}</td>
                    <td style={{ padding: "9px 10px", fontSize: 12 }}>{fmtN(row.budget)}</td>
                    <td style={{ padding: "9px 10px", fontSize: 12, color: "var(--accent)", fontWeight: 700 }}>{fmtN(row.committed)}</td>
                    <td style={{ padding: "9px 10px", fontSize: 12 }}>{fmtN(row.invoiced)}</td>
                    <td style={{ padding: "9px 10px", fontSize: 12, color: "var(--green)", fontWeight: 700 }}>{fmtN(row.paid)}</td>
                    <td style={{ padding: "9px 10px", fontSize: 12, color: row.outstanding > 0 ? "var(--red)" : "var(--text2)", fontWeight: row.outstanding > 0 ? 700 : 400 }}>{fmtN(row.outstanding)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </Card>
    </div>
  );
};


const CloseoutPage = ({ vendors, clients, campaigns, mpos, onOpenReceivables }) => {
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
        <div class="card"><div class="label">Reconciled Amount</div><div class="value">${formatNairaExportValue(row.reconciledAmount)}</div><div class="muted">Invoice: ${row.invoiceNo}</div></div>
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
            <Btn variant="blue" size="sm" icon="⬇" onClick={() => exportView("Closeout Summary", ["Metric", "Value"], [["MPO Value in Closeout", formatNairaExportValue(totalValue)], ["Ready to Bill", readyToBillCount], ["Proof Pending", proofPendingCount], ["Open Make-good Items", openMakegoodCount], ["Exceptions", exceptionCount], ["Closed", closedCount], ["Collections Exposure", formatNairaExportValue(collectionsExposure)], ["Average Delivery", `${avgDelivery.toFixed(1)}%`]])}>Export Summary</Btn>
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
          <Btn variant="ghost" size="sm" onClick={() => exportView("Closeout Pipeline", ["Stage", "Count", "Value"], formatExportRowsWithCurrency(stageSummary.map(row => [row.label, row.count, row.value]), [2]))}>Export Pipeline</Btn>
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
              <Btn variant="ghost" size="sm" onClick={() => exportView("Execution Variance and Billing Readiness", ["MPO No.", "Campaign", "Client", "Vendor", "Planned", "Aired", "Missed", "Make-good", "Delivery %", "Proof", "Reconciliation", "Invoice", "Payment", "Stage", "Ready to Bill", "Ready to Close", "Reconciled Amount"], formatExportRowsWithCurrency(closeoutRows.map(row => [row.mpoNo, row.campaign, row.client, row.vendor, row.planned, row.aired, row.missed, row.makegood, `${row.deliveryPct.toFixed(1)}%`, row.proofStatus, row.reconciliationStatus, row.invoiceStatus, row.paymentStatus, row.stageLabel, row.readyToBill ? "Yes" : "No", row.readyToClose ? "Yes" : "No", row.reconciledAmount]), [16]))}>Export Matrix</Btn>
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
              <Btn variant="ghost" size="sm" onClick={() => exportView("Billing and Collections Watchlist", ["MPO No.", "Client", "Campaign", "Stage", "Invoice Status", "Payment Status", "Ready to Bill", "Amount", "Payment Ref", "Paid At"], formatExportRowsWithCurrency(collectionsRows.map(row => [row.mpoNo, row.client, row.campaign, row.stageLabel, row.invoiceStatus, row.paymentStatus, row.readyToBill ? "Yes" : "No", row.reconciledAmount, row.paymentReference, row.paidAt]), [7]))}>Export Watchlist</Btn>
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


const ReceivablesPage = ({ user, clients, campaigns, mpos, receivables, receivablesMeta, onSaveReceivable, onRemoveReceivable, onLogReceivablePayment, onUpdateReceivableStatus, onOpenCloseout }) => {
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
      ? row.payments.map(payment => `<tr><td>${formatIsoDate(payment.receivedAt)}</td><td>${payment.reference || "—"}</td><td>${payment.channel || "—"}</td><td>${formatNairaExportValue(payment.amount)}</td><td>${payment.note || "—"}</td></tr>`).join("")
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
        <div class="card"><div class="label">Gross Amount</div><div class="value">${formatNairaExportValue(row.grossAmount)}</div></div>
        <div class="card"><div class="label">Received</div><div class="value">${formatNairaExportValue(row.amountReceived)}</div></div>
        <div class="card"><div class="label">Outstanding</div><div class="value">${formatNairaExportValue(row.balance)}</div></div>
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
      ? row.payments.map(payment => [row.invoiceNo, row.clientName, formatIsoDate(payment.receivedAt), payment.reference, payment.channel, formatNairaExportValue(payment.amount), payment.note])
      : [[row.invoiceNo, row.clientName, "", "", "", formatNairaExportValue(0), "No payment logged yet"]];
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
          <Btn variant="blue" icon="⬇" onClick={() => exportView("Receivables Ledger", ["Invoice No.", "Client", "Campaign", "MPO", "Invoice Date", "Due Date", "Status", "Collection Stage", "Gross Amount", "Received", "Outstanding", "Days Past Due", "Owner", "Notes"], formatExportRowsWithCurrency(filteredRows.map(row => [row.invoiceNo, row.clientName, row.campaignName, row.mpoNo, row.invoiceDate, row.dueDate, getStatusLabel(row.status), getStageLabel(row.collectionStage), row.grossAmount, row.amountReceived, row.balance, row.daysPastDue, row.owner, row.notes]), [8, 9, 10]))}>Export Ledger</Btn>
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
          <Btn variant="ghost" size="sm" onClick={() => exportView("Receivable Aging Summary", ["Bucket", "Count", "Outstanding Balance"], formatExportRowsWithCurrency(agingSummary.map(row => [row.label, row.count, row.balance]), [2]))}>Export Aging</Btn>
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
            <Btn variant="ghost" size="sm" onClick={() => exportView("Collections Watchlist", ["Invoice No.", "Client", "Status", "Stage", "Due Date", "Days Past Due", "Outstanding", "Owner", "Last Follow-up"], formatExportRowsWithCurrency(filteredRows.filter(row => row.balance > 0).map(row => [row.invoiceNo, row.clientName, getStatusLabel(row.status), getStageLabel(row.collectionStage), row.dueDate, row.daysPastDue, row.balance, row.owner, row.lastFollowUpAt]), [6]))}>Export Watchlist</Btn>
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
              <Btn variant="ghost" size="sm" onClick={() => exportView("Overdue Receivables", ["Invoice No.", "Client", "Campaign", "Due Date", "Days Past Due", "Outstanding", "Owner", "Collection Stage"], formatExportRowsWithCurrency(filteredRows.filter(row => row.daysPastDue > 0 && row.balance > 0).map(row => [row.invoiceNo, row.clientName, row.campaignName, row.dueDate, row.daysPastDue, row.balance, row.owner, getStageLabel(row.collectionStage)]), [5]))}>Export Overdue</Btn>
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



/* ── SETTINGS ───────────────────────────────────────────── */
const FinancePage = ({ user, vendors, clients, campaigns, mpos, receivables, receivablesMeta, onSaveReceivable, onRemoveReceivable, onLogReceivablePayment, onUpdateReceivableStatus }) => {
  const [tab, setTab] = useState("budgeting");

  const tabs = [
    { id: "budgeting", icon: "🧮", label: "Budgeting", sub: "Budget control, invoice and payment status" },
    { id: "closeout", icon: "🧾", label: "Closeout", sub: "Reconciliation, proof, make-goods and billing readiness" },
    { id: "receivables", icon: "💵", label: "Receivables", sub: "Client invoicing, collections and cash recovery" },
  ];

  const activeTab = tabs.find(item => item.id === tab) || tabs[0];

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 18, minWidth: 0 }}>
      <Card style={{ padding: 18 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 16, flexWrap: "wrap" }}>
          <div style={{ minWidth: 0, flex: "1 1 320px" }}>
            <h1 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 800, fontSize: 24, margin: 0 }}>Finance</h1>
            <p style={{ marginTop: 6, marginBottom: 0, color: "var(--text3)", lineHeight: 1.6, fontSize: 13, maxWidth: 860 }}>Manage budget control, vendor settlement, reconciliation closeout, and client collections in one finance workspace.</p>
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap", minWidth: 0 }}>
            <Badge color="accent">{activeTab.label}</Badge>
            <span style={{ fontSize: 11, color: "var(--text3)", maxWidth: 360, lineHeight: 1.45 }}>{activeTab.sub}</span>
          </div>
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(180px,1fr))", gap: 10, marginTop: 16 }}>
          {tabs.map(item => {
            const active = item.id === tab;
            return (
              <button
                key={item.id}
                onClick={() => setTab(item.id)}
                style={{
                  border: active ? "1px solid rgba(240,165,0,.38)" : "1px solid var(--border)",
                  background: active ? "rgba(240,165,0,.12)" : "var(--bg3)",
                  color: active ? "var(--accent)" : "var(--text2)",
                  borderRadius: 12,
                  padding: "12px 14px",
                  textAlign: "left",
                  cursor: "pointer",
                  minWidth: 0,
                }}
              >
                <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 6, minWidth: 0 }}>
                  <span style={{ fontSize: 16, flexShrink: 0 }}>{item.icon}</span>
                  <span style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 14, minWidth: 0, overflowWrap: "anywhere" }}>{item.label}</span>
                </div>
                <div style={{ fontSize: 11, lineHeight: 1.45, color: active ? "var(--text)" : "var(--text3)", overflowWrap: "anywhere" }}>{item.sub}</div>
              </button>
            );
          })}
        </div>
      </Card>

      {tab === "budgeting" && <BudgetingPage vendors={vendors} clients={clients} campaigns={campaigns} mpos={mpos} />}
      {tab === "closeout" && <CloseoutPage vendors={vendors} clients={clients} campaigns={campaigns} mpos={mpos} onOpenReceivables={() => setTab("receivables")} />}
      {tab === "receivables" && <ReceivablesPage user={user} clients={clients} campaigns={campaigns} mpos={mpos} receivables={receivables} receivablesMeta={receivablesMeta} onSaveReceivable={onSaveReceivable} onRemoveReceivable={onRemoveReceivable} onLogReceivablePayment={onLogReceivablePayment} onUpdateReceivableStatus={onUpdateReceivableStatus} onOpenCloseout={() => setTab("closeout")} />}
    </div>
  );
};

export default FinancePage;
