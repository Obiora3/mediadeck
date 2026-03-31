import { useState } from "react";
import Empty from "../../components/Empty";
import Toast from "../../components/Toast";
import Btn from "../../components/Btn";
import Field from "../../components/Field";
import Badge from "../../components/Badge";
import Card from "../../components/Card";
import Stat from "../../components/Statcard";
import { MPO_PAYMENT_STATUS_OPTIONS } from "../../constants/mpoWorkflow";

const BudgetingPage = ({ vendors, clients, campaigns, mpos, activeOnly, fmtN, buildCSV, PrintPreview }) => {
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
            <Btn variant="blue" size="sm" icon="⬇" onClick={() => exportBudgetView("Budgeting Summary", ["Metric", "Value"], [["Campaign Budget Pool", totalBudget], ["Committed Spend", totalCommitted], ["Invoiced", totalInvoiced], ["Paid", totalPaid], ["Outstanding", totalOutstanding], ["Available Budget", availableBudget], ["Over-budget Campaigns", overBudgetCount], ["Average Utilization", `${averageUtilization.toFixed(1)}%`]])}>Export Summary</Btn>
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
            <Btn variant="ghost" size="sm" onClick={() => exportBudgetView("Campaign Budget Tracker", ["Campaign", "Client", "Brand", "Medium", "Status", "Budget", "Committed", "Available", "Invoiced", "Paid", "Outstanding", "Utilization %", "MPO Count"], campaignFinanceRows.map(row => [row.campaign, row.client, row.brand, row.medium, row.status, row.budget, row.committed, row.available, row.invoiced, row.paid, row.outstanding, `${row.utilization.toFixed(1)}%`, row.mpoCount]))}>Export Tracker</Btn>
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
              <Btn variant="ghost" size="sm" onClick={() => exportBudgetView("Invoice Aging", ["Bucket", "Count", "Outstanding"], agingBuckets.map(row => [row.label, row.count, row.value]))}>Export Aging</Btn>
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
              <Btn variant="ghost" size="sm" onClick={() => exportBudgetView("Client Budget Coverage", ["Client", "Campaigns", "Budget", "Committed", "Available", "Invoiced", "Paid", "Outstanding"], clientBudgetRows.map(row => [row.client, row.campaigns, row.budget, row.committed, row.available, row.invoiced, row.paid, row.outstanding]))}>Export Clients</Btn>
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
          <Btn variant="ghost" size="sm" onClick={() => exportBudgetView("Invoice and Payment Register", ["MPO No.", "Vendor", "Client", "Campaign", "Invoice No.", "Invoice Date", "Invoice Status", "Payment Status", "Payment Ref", "Paid At", "Age (Days)", "Outstanding"], invoiceAgingRows.map(row => [row.mpoNo, row.vendor, row.client, row.campaign, row.invoiceNo, row.invoiceDate, row.invoiceStatus, row.paymentStatus, row.paymentReference, row.paidAt, row.ageDays, row.outstanding]))}>Export Register</Btn>
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
          <Btn variant="ghost" size="sm" onClick={() => exportBudgetView("Monthly Budget vs Cashflow", ["Period", "Budget", "Committed", "Invoiced", "Paid", "Outstanding"], monthlyCashflow.map(row => [row.period, row.budget, row.committed, row.invoiced, row.paid, row.outstanding]))}>Export Cashflow</Btn>
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



export default BudgetingPage;
