import { useState } from "react";
import Empty from "../components/Empty";
import Toast from "../components/Toast";
import { Btn, Field, Card } from "../components/ui/primitives";
import { MPO_STATUS_OPTIONS } from "../constants/mpoWorkflow";
import { formatNairaExportValue, formatExportRowsWithCurrency } from "../utils/export";

export default function ReportsPage({ vendors, clients, campaigns, rates, mpos, activeOnly, fmtN, MPO_STATUS_LABELS, PrintPreview, buildCSV }) {
  const [toast, setToast] = useState(null);
  const [preview, setPreview] = useState(null);
  const [filters, setFilters] = useState({
    startDate: "",
    endDate: "",
    vendorId: "",
    clientId: "",
    campaignId: "",
    medium: "",
    mpoStatus: "",
    paymentStatus: "",
    reconciliationStatus: "",
    search: "",
  });

  const updateFilter = (key) => (value) => setFilters(prev => ({ ...prev, [key]: value }));

  const liveVendors = activeOnly(vendors);
  const liveClients = activeOnly(clients);
  const liveCampaigns = activeOnly(campaigns);
  const liveRates = activeOnly(rates);
  const liveMpos = activeOnly(mpos);

  const parseMpoTimestamp = (mpo) => {
    if (mpo?.date) {
      const t = new Date(mpo.date).getTime();
      if (!Number.isNaN(t)) return t;
    }
    if (mpo?.createdAt) return mpo.createdAt;
    if (mpo?.updatedAt) return mpo.updatedAt;
    return 0;
  };

  const mediumForMpo = (mpo) => {
    const campaign = liveCampaigns.find(c => c.id === mpo.campaignId);
    const vendor = liveVendors.find(v => v.id === mpo.vendorId);
    return (mpo.medium || campaign?.medium || vendor?.type || "Unknown").trim();
  };

  const matchSearch = (mpo) => {
    const term = filters.search.trim().toLowerCase();
    if (!term) return true;
    return [
      mpo.mpoNo,
      mpo.vendorName,
      mpo.clientName,
      mpo.brand,
      mpo.campaignName,
      mediumForMpo(mpo),
      mpo.status,
      mpo.paymentStatus,
      mpo.reconciliationStatus,
    ].some(value => String(value || "").toLowerCase().includes(term));
  };

  const filteredMpos = liveMpos.filter(mpo => {
    const mpoTs = parseMpoTimestamp(mpo);
    const startTs = filters.startDate ? new Date(`${filters.startDate}T00:00:00`).getTime() : null;
    const endTs = filters.endDate ? new Date(`${filters.endDate}T23:59:59`).getTime() : null;

    if (startTs && mpoTs && mpoTs < startTs) return false;
    if (endTs && mpoTs && mpoTs > endTs) return false;
    if (filters.vendorId && mpo.vendorId !== filters.vendorId) return false;
    if (filters.clientId) {
      const campaign = liveCampaigns.find(c => c.id === mpo.campaignId);
      if ((campaign?.clientId || "") !== filters.clientId) return false;
    }
    if (filters.campaignId && mpo.campaignId !== filters.campaignId) return false;
    if (filters.medium && mediumForMpo(mpo) !== filters.medium) return false;
    if (filters.mpoStatus && (mpo.status || "draft") !== filters.mpoStatus) return false;
    if (filters.paymentStatus && (mpo.paymentStatus || "unpaid") !== filters.paymentStatus) return false;
    if (filters.reconciliationStatus && (mpo.reconciliationStatus || "not_started") !== filters.reconciliationStatus) return false;
    if (!matchSearch(mpo)) return false;
    return true;
  });

  const filteredCampaignIds = new Set(filteredMpos.map(m => m.campaignId).filter(Boolean));
  const filteredVendorIds = new Set(filteredMpos.map(m => m.vendorId).filter(Boolean));
  const filteredClientIds = new Set(
    filteredMpos
      .map(m => liveCampaigns.find(c => c.id === m.campaignId)?.clientId)
      .filter(Boolean)
  );

  const selectedCampaigns = liveCampaigns.filter(c => {
    if (filters.campaignId && c.id !== filters.campaignId) return false;
    if (filters.clientId && c.clientId !== filters.clientId) return false;
    if (filters.medium && (c.medium || "") !== filters.medium) return false;
    if (!filters.campaignId && !filters.clientId && !filters.medium && filteredCampaignIds.size) {
      return filteredCampaignIds.has(c.id);
    }
    return true;
  });

  const selectedRates = liveRates.filter(rate => {
    if (filters.vendorId && rate.vendorId !== filters.vendorId) return false;
    if (filters.medium && (rate.mediaType || "") !== filters.medium) return false;
    if (filters.search) {
      const term = filters.search.toLowerCase();
      const vendorName = liveVendors.find(v => v.id === rate.vendorId)?.name || "";
      if (![vendorName, rate.programme, rate.timeBelt, rate.mediaType, rate.notes].some(value => String(value || "").toLowerCase().includes(term))) return false;
    }
    return true;
  });

  const totalMpoValue = filteredMpos.reduce((sum, mpo) => sum + (parseFloat(mpo.grandTotal) || parseFloat(mpo.netVal) || 0), 0);
  const totalGrossValue = filteredMpos.reduce((sum, mpo) => sum + (parseFloat(mpo.totalGross) || 0), 0);
  const paidValue = filteredMpos.reduce((sum, mpo) => sum + ((mpo.paymentStatus || "unpaid") === "paid" ? (parseFloat(mpo.reconciledAmount) || parseFloat(mpo.invoiceAmount) || parseFloat(mpo.grandTotal) || 0) : 0), 0);
  const outstandingValue = filteredMpos.reduce((sum, mpo) => sum + ((mpo.paymentStatus || "unpaid") !== "paid" ? (parseFloat(mpo.reconciledAmount) || parseFloat(mpo.invoiceAmount) || parseFloat(mpo.grandTotal) || 0) : 0), 0);
  const reconciledValue = filteredMpos.reduce((sum, mpo) => sum + ((mpo.reconciliationStatus || "not_started") === "completed" ? (parseFloat(mpo.reconciledAmount) || parseFloat(mpo.grandTotal) || 0) : 0), 0);
  const budgetPool = selectedCampaigns.reduce((sum, campaign) => sum + (parseFloat(campaign.budget) || 0), 0);
  const plannedSpots = filteredMpos.reduce((sum, mpo) => sum + (parseFloat(mpo.plannedSpotsExecution ?? mpo.totalSpots) || 0), 0);
  const airedSpots = filteredMpos.reduce((sum, mpo) => sum + (parseFloat(mpo.airedSpots) || 0), 0);
  const pendingApprovals = filteredMpos.filter(mpo => ["draft", "submitted", "reviewed"].includes(mpo.status || "draft")).length;
  const unpaidCount = filteredMpos.filter(mpo => (mpo.paymentStatus || "unpaid") !== "paid").length;
  const reconciliationPendingCount = filteredMpos.filter(mpo => (mpo.reconciliationStatus || "not_started") !== "completed").length;

  const spendByVendor = Object.values(filteredMpos.reduce((acc, mpo) => {
    const key = mpo.vendorId || mpo.vendorName || "unknown";
    if (!acc[key]) {
      const vendor = liveVendors.find(v => v.id === mpo.vendorId);
      acc[key] = {
        vendor: mpo.vendorName || vendor?.name || "Unknown Vendor",
        medium: mediumForMpo(mpo),
        mpoCount: 0,
        spots: 0,
        gross: 0,
        net: 0,
        paid: 0,
        outstanding: 0,
      };
    }
    acc[key].mpoCount += 1;
    acc[key].spots += parseFloat(mpo.totalSpots) || 0;
    acc[key].gross += parseFloat(mpo.totalGross) || 0;
    const value = parseFloat(mpo.grandTotal) || parseFloat(mpo.netVal) || 0;
    acc[key].net += value;
    if ((mpo.paymentStatus || "unpaid") === "paid") acc[key].paid += value;
    else acc[key].outstanding += value;
    return acc;
  }, {})).sort((a, b) => b.net - a.net);

  const spendByClient = Object.values(filteredMpos.reduce((acc, mpo) => {
    const campaign = liveCampaigns.find(c => c.id === mpo.campaignId);
    const client = liveClients.find(c => c.id === campaign?.clientId);
    const key = campaign?.clientId || mpo.clientName || "unknown";
    if (!acc[key]) {
      acc[key] = {
        client: mpo.clientName || client?.name || "Unknown Client",
        campaignCount: 0,
        mpoCount: 0,
        budget: 0,
        spend: 0,
        variance: 0,
      };
    }
    acc[key].mpoCount += 1;
    acc[key].spend += parseFloat(mpo.grandTotal) || parseFloat(mpo.netVal) || 0;
    return acc;
  }, {})).map(entry => {
    const campaignsForClient = selectedCampaigns.filter(campaign => {
      const client = liveClients.find(c => c.id === campaign.clientId);
      return (client?.name || "") === entry.client || campaign.clientId === liveClients.find(c => c.name === entry.client)?.id;
    });
    entry.campaignCount = campaignsForClient.length;
    entry.budget = campaignsForClient.reduce((sum, campaign) => sum + (parseFloat(campaign.budget) || 0), 0);
    entry.variance = entry.budget - entry.spend;
    return entry;
  }).sort((a, b) => b.spend - a.spend);

  const campaignBudgetControl = selectedCampaigns.map(campaign => {
    const client = liveClients.find(c => c.id === campaign.clientId);
    const campaignMpos = filteredMpos.filter(mpo => mpo.campaignId === campaign.id);
    const spend = campaignMpos.reduce((sum, mpo) => sum + (parseFloat(mpo.grandTotal) || parseFloat(mpo.netVal) || 0), 0);
    const gross = campaignMpos.reduce((sum, mpo) => sum + (parseFloat(mpo.totalGross) || 0), 0);
    const budget = parseFloat(campaign.budget) || 0;
    const utilization = budget > 0 ? (spend / budget) * 100 : 0;
    return {
      campaign: campaign.name || "Unnamed Campaign",
      client: client?.name || "—",
      brand: campaign.brand || "—",
      medium: campaign.medium || "—",
      status: campaign.status || "draft",
      budget,
      gross,
      spend,
      variance: budget - spend,
      mpoCount: campaignMpos.length,
      utilization,
    };
  }).sort((a, b) => b.spend - a.spend);

  const financeTracker = filteredMpos.map(mpo => ({
    mpoNo: mpo.mpoNo || "—",
    vendor: mpo.vendorName || "—",
    client: mpo.clientName || "—",
    campaign: mpo.campaignName || "—",
    value: parseFloat(mpo.grandTotal) || parseFloat(mpo.netVal) || 0,
    invoiceStatus: mpo.invoiceStatus || "pending",
    invoiceNo: mpo.invoiceNo || "—",
    paymentStatus: mpo.paymentStatus || "unpaid",
    paymentReference: mpo.paymentReference || "—",
    paidAt: mpo.paidAt ? new Date(mpo.paidAt).toLocaleDateString() : "—",
    outstanding: (mpo.paymentStatus || "unpaid") === "paid" ? 0 : (parseFloat(mpo.reconciledAmount) || parseFloat(mpo.invoiceAmount) || parseFloat(mpo.grandTotal) || 0),
  })).sort((a, b) => b.outstanding - a.outstanding);

  const reconciliationControl = filteredMpos.map(mpo => ({
    mpoNo: mpo.mpoNo || "—",
    vendor: mpo.vendorName || "—",
    campaign: mpo.campaignName || "—",
    planned: parseFloat(mpo.plannedSpotsExecution ?? mpo.totalSpots) || 0,
    aired: parseFloat(mpo.airedSpots) || 0,
    missed: parseFloat(mpo.missedSpots) || 0,
    makegood: parseFloat(mpo.makegoodSpots) || 0,
    proofStatus: mpo.proofStatus || "pending",
    reconciliationStatus: mpo.reconciliationStatus || "not_started",
    reconciledAmount: parseFloat(mpo.reconciledAmount) || parseFloat(mpo.grandTotal) || 0,
    notes: mpo.reconciliationNotes || "—",
  })).sort((a, b) => b.reconciledAmount - a.reconciledAmount);

  const statusPipeline = [
    "draft",
    "submitted",
    "reviewed",
    "approved",
    "sent",
    "aired",
    "reconciled",
    "closed",
    "rejected",
  ].map(status => {
    const rows = filteredMpos.filter(mpo => (mpo.status || "draft") === status);
    return {
      status: MPO_STATUS_LABELS[status] || status,
      count: rows.length,
      value: rows.reduce((sum, mpo) => sum + (parseFloat(mpo.grandTotal) || parseFloat(mpo.netVal) || 0), 0),
      paid: rows.reduce((sum, mpo) => sum + ((mpo.paymentStatus || "unpaid") === "paid" ? (parseFloat(mpo.grandTotal) || parseFloat(mpo.netVal) || 0) : 0), 0),
    };
  });

  const rateCardSnapshot = selectedRates.map(rate => {
    const vendor = liveVendors.find(v => v.id === rate.vendorId);
    const discount = parseFloat(rate.discount) || 0;
    const commission = parseFloat(rate.commission) || 0;
    const rateValue = parseFloat(rate.ratePerSpot) || 0;
    const netRate = rateValue * (1 - discount / 100) * (1 - commission / 100);
    return {
      vendor: vendor?.name || "—",
      programme: rate.programme || "—",
      timeBelt: rate.timeBelt || "—",
      medium: rate.mediaType || vendor?.type || "—",
      duration: `${rate.duration || "30"}"`,
      rate: rateValue,
      discount,
      commission,
      netRate,
      notes: rate.notes || "—",
    };
  }).sort((a, b) => b.netRate - a.netRate);

  const filteredMpoRegister = filteredMpos.map(mpo => ({
    mpoNo: mpo.mpoNo || "—",
    date: mpo.date || "—",
    vendor: mpo.vendorName || "—",
    client: mpo.clientName || "—",
    brand: mpo.brand || "—",
    campaign: mpo.campaignName || "—",
    medium: mediumForMpo(mpo),
    status: mpo.status || "draft",
    paymentStatus: mpo.paymentStatus || "unpaid",
    reconciliationStatus: mpo.reconciliationStatus || "not_started",
    totalSpots: parseFloat(mpo.totalSpots) || 0,
    gross: parseFloat(mpo.totalGross) || 0,
    net: parseFloat(mpo.netVal) || 0,
    grandTotal: parseFloat(mpo.grandTotal) || parseFloat(mpo.netVal) || 0,
  })).sort((a, b) => (new Date(b.date).getTime() || 0) - (new Date(a.date).getTime() || 0));

  const summaryCards = [
    { label: "Filtered MPO Value", value: totalMpoValue, sub: `${filteredMpos.length} MPOs`, color: "var(--accent)", icon: "📄" },
    { label: "Paid Value", value: paidValue, sub: `${filteredMpos.filter(m => (m.paymentStatus || "unpaid") === "paid").length} paid`, color: "var(--green)", icon: "✅" },
    { label: "Outstanding Exposure", value: outstandingValue, sub: `${unpaidCount} unpaid`, color: "var(--red)", icon: "💸" },
    { label: "Reconciled Value", value: reconciledValue, sub: `${filteredMpos.filter(m => (m.reconciliationStatus || "not_started") === "completed").length} completed`, color: "var(--blue)", icon: "🧾" },
    { label: "Pending Approvals", value: pendingApprovals, sub: "Draft / submitted / reviewed", color: "var(--purple)", icon: "⏳" },
    { label: "Aired vs Planned Spots", value: `${airedSpots}/${plannedSpots || 0}`, sub: plannedSpots ? `${Math.round((airedSpots / plannedSpots) * 100)}% delivery` : "No planned spots", color: "var(--teal)", icon: "📡" },
    { label: "Campaign Budget Pool", value: budgetPool, sub: `${selectedCampaigns.length} campaigns`, color: "var(--orange)", icon: "🎯" },
    { label: "Gross MPO Value", value: totalGrossValue, sub: `${reconciliationPendingCount} reconciliation pending`, color: "var(--text)", icon: "📊" },
  ];

  const buildSectionExport = (title, headers, rows, descriptor = "") => {
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
        @media print{body{padding:10px}}
      </style>
    </head><body>
      <h1>${title}</h1>
      <p>Generated ${new Date().toLocaleString("en-NG")}${descriptor ? ` · ${descriptor}` : ""}</p>
      <table><tr>${headers.map(h => `<th>${esc(h)}</th>`).join("")}</tr>${htmlRows}</table>
    </body></html>`;
    const csv = buildCSV(rows, headers);
    setPreview({ html, csv, title });
  };

  const filterDescriptorParts = [
    filters.startDate ? `From ${filters.startDate}` : "",
    filters.endDate ? `To ${filters.endDate}` : "",
    filters.vendorId ? `Vendor: ${liveVendors.find(v => v.id === filters.vendorId)?.name || "Selected"}` : "",
    filters.clientId ? `Client: ${liveClients.find(c => c.id === filters.clientId)?.name || "Selected"}` : "",
    filters.campaignId ? `Campaign: ${liveCampaigns.find(c => c.id === filters.campaignId)?.name || "Selected"}` : "",
    filters.medium ? `Medium: ${filters.medium}` : "",
    filters.mpoStatus ? `Status: ${MPO_STATUS_LABELS[filters.mpoStatus] || filters.mpoStatus}` : "",
    filters.paymentStatus ? `Payment: ${filters.paymentStatus}` : "",
    filters.reconciliationStatus ? `Reconciliation: ${filters.reconciliationStatus}` : "",
  ].filter(Boolean);
  const filterDescriptor = filterDescriptorParts.join(" · ");

  const sectionConfigs = [
    {
      id: "spend-vendor",
      title: "Spend by Vendor",
      icon: "🏢",
      desc: "See which media owners are taking the largest share of spend and exposure.",
      headers: ["Vendor", "Medium", "MPO Count", "Spots", "Gross Value", "Net Value", "Paid", "Outstanding"],
      rows: spendByVendor.map(row => [row.vendor, row.medium, row.mpoCount, row.spots, row.gross, row.net, row.paid, row.outstanding]),
      currencyColumns: [4, 5, 6, 7],
    },
    {
      id: "spend-client",
      title: "Spend by Client",
      icon: "👥",
      desc: "Track client value, budget coverage, and campaign concentration.",
      headers: ["Client", "Campaigns", "MPO Count", "Budget Pool", "Spend", "Budget Variance"],
      rows: spendByClient.map(row => [row.client, row.campaignCount, row.mpoCount, row.budget, row.spend, row.variance]),
      currencyColumns: [3, 4, 5],
    },
    {
      id: "campaign-budget",
      title: "Campaign Budget Control",
      icon: "🎯",
      desc: "Compare campaign budgets against MPO spend and utilization.",
      headers: ["Campaign", "Client", "Brand", "Medium", "Status", "Budget", "Gross Value", "MPO Spend", "Variance", "Utilization %", "MPO Count"],
      rows: campaignBudgetControl.map(row => [row.campaign, row.client, row.brand, row.medium, row.status, row.budget, row.gross, row.spend, row.variance, `${row.utilization.toFixed(1)}%`, row.mpoCount]),
      currencyColumns: [5, 6, 7, 8],
    },
    {
      id: "finance-tracker",
      title: "Finance Tracker",
      icon: "💸",
      desc: "Follow invoice, payment, and outstanding exposure on every MPO.",
      headers: ["MPO No.", "Vendor", "Client", "Campaign", "Value", "Invoice Status", "Invoice No.", "Payment Status", "Payment Ref", "Paid At", "Outstanding"],
      rows: financeTracker.map(row => [row.mpoNo, row.vendor, row.client, row.campaign, row.value, row.invoiceStatus, row.invoiceNo, row.paymentStatus, row.paymentReference, row.paidAt, row.outstanding]),
      currencyColumns: [4, 10],
    },
    {
      id: "reconciliation-control",
      title: "Reconciliation Control",
      icon: "🧾",
      desc: "Watch proof of airing, delivery variance, and reconciliation status.",
      headers: ["MPO No.", "Vendor", "Campaign", "Planned Spots", "Aired Spots", "Missed", "Makegood", "Proof Status", "Reconciliation", "Reconciled Amount", "Notes"],
      rows: reconciliationControl.map(row => [row.mpoNo, row.vendor, row.campaign, row.planned, row.aired, row.missed, row.makegood, row.proofStatus, row.reconciliationStatus, row.reconciledAmount, row.notes]),
      currencyColumns: [9],
    },
    {
      id: "pipeline",
      title: "Status Pipeline",
      icon: "📈",
      desc: "Measure MPO flow through draft, approval, airing, and closeout.",
      headers: ["Status", "Count", "Value", "Paid Value"],
      rows: statusPipeline.map(row => [row.status, row.count, row.value, row.paid]),
      currencyColumns: [2, 3],
    },
    {
      id: "rate-snapshot",
      title: "Rate Card Snapshot",
      icon: "💰",
      desc: "Review effective net rates after discount and commission.",
      headers: ["Vendor", "Programme", "Time Belt", "Medium", "Duration", "Rate/Spot", "Disc %", "Comm %", "Net Rate", "Notes"],
      rows: rateCardSnapshot.map(row => [row.vendor, row.programme, row.timeBelt, row.medium, row.duration, row.rate, `${row.discount}%`, `${row.commission}%`, row.netRate, row.notes]),
      currencyColumns: [5, 8],
    },
    {
      id: "mpo-register",
      title: "Filtered MPO Register",
      icon: "📄",
      desc: "A clean register of all MPOs matching your filters.",
      headers: ["MPO No.", "Date", "Vendor", "Client", "Brand", "Campaign", "Medium", "Status", "Payment Status", "Reconciliation", "Total Spots", "Gross", "Net", "Grand Total"],
      rows: filteredMpoRegister.map(row => [row.mpoNo, row.date, row.vendor, row.client, row.brand, row.campaign, row.medium, row.status, row.paymentStatus, row.reconciliationStatus, row.totalSpots, row.gross, row.net, row.grandTotal]),
      currencyColumns: [11, 12, 13],
    },
  ];

  const mediumOptions = Array.from(new Set([
    ...liveCampaigns.map(c => c.medium).filter(Boolean),
    ...liveRates.map(r => r.mediaType).filter(Boolean),
    ...liveMpos.map(m => mediumForMpo(m)).filter(Boolean),
  ])).sort().map(value => ({ value, label: value }));

  const resetFilters = () => setFilters({
    startDate: "",
    endDate: "",
    vendorId: "",
    clientId: "",
    campaignId: "",
    medium: "",
    mpoStatus: "",
    paymentStatus: "",
    reconciliationStatus: "",
    search: "",
  });

  const exportExecutiveSummary = () => {
    const headers = ["Metric", "Value", "Context"];
    const currencySummaryLabels = new Set([
      "Filtered MPO Value",
      "Paid Value",
      "Outstanding Exposure",
      "Reconciled Value",
      "Campaign Budget Pool",
      "Gross MPO Value",
    ]);
    const rows = summaryCards.map(card => [
      card.label,
      currencySummaryLabels.has(card.label) && typeof card.value === "number"
        ? formatNairaExportValue(card.value)
        : card.value,
      card.sub,
    ]);
    buildSectionExport("Executive Finance Summary", headers, rows, filterDescriptor);
  };

  const financeInputStyle = {
    background: "var(--bg3)",
    border: "1px solid var(--border2)",
    borderRadius: 8,
    padding: "9px 13px",
    color: "var(--text)",
    fontSize: 13,
    outline: "none",
    width: "100%",
  };
  const formatCount = (value) => Number(value || 0).toLocaleString("en-NG");
  const countSummaryLabels = new Set([
    "Pending Approvals",
  ]);
  const sectionCountColumns = {
    "spend-vendor": new Set([2, 3]),
    "spend-client": new Set([1, 2]),
    "campaign-budget": new Set([10]),
    "reconciliation": new Set([3, 4, 5, 6]),
    "status-pipeline": new Set([1]),
    "mpo-register": new Set([10]),
  };

  return (
    <div className="fade">
      {toast && <Toast msg={toast.msg} type={toast.type} onDone={() => setToast(null)} />}
      {preview && <PrintPreview html={preview.html} csv={preview.csv} pdfBytes={preview.pdfBytes} title={preview.title} onClose={() => setPreview(null)} />}

      <div style={{ marginBottom: 24 }}>
        <h1 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 800, fontSize: 24 }}>Reports & Finance Control</h1>
        <p style={{ color: "var(--text2)", marginTop: 3, fontSize: 13 }}>Filter, monitor, and export spend, payment, reconciliation, and campaign control views from one workspace.</p>
      </div>

      <Card style={{ marginBottom: 20 }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, flexWrap: "wrap", marginBottom: 16 }}>
          <div>
            <h2 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 16 }}>Control Filters</h2>
            <p style={{ color: "var(--text2)", fontSize: 12, marginTop: 3 }}>Narrow the reporting scope by dates, ownership, status, and finance state.</p>
          </div>
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
            <Btn variant="ghost" size="sm" onClick={resetFilters}>Reset Filters</Btn>
            <Btn variant="blue" size="sm" icon="⬇" onClick={exportExecutiveSummary}>Export Summary</Btn>
          </div>
        </div>

        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(180px,1fr))", gap: 12 }}>
          <div><label style={{ fontSize: 11, fontWeight: 600, color: "var(--text3)", textTransform: "uppercase", letterSpacing: ".08em", marginBottom: 5, display: "block" }}>Start Date</label><input type="date" value={filters.startDate} onChange={e => updateFilter("startDate")(e.target.value)} style={financeInputStyle} /></div>
          <div><label style={{ fontSize: 11, fontWeight: 600, color: "var(--text3)", textTransform: "uppercase", letterSpacing: ".08em", marginBottom: 5, display: "block" }}>End Date</label><input type="date" value={filters.endDate} onChange={e => updateFilter("endDate")(e.target.value)} style={financeInputStyle} /></div>
          <Field label="Vendor" value={filters.vendorId} onChange={updateFilter("vendorId")} options={liveVendors.map(v => ({ value: v.id, label: v.name }))} placeholder="All Vendors" />
          <Field label="Client" value={filters.clientId} onChange={updateFilter("clientId")} options={liveClients.map(c => ({ value: c.id, label: c.name }))} placeholder="All Clients" />
          <Field label="Campaign" value={filters.campaignId} onChange={updateFilter("campaignId")} options={liveCampaigns.map(c => ({ value: c.id, label: c.name }))} placeholder="All Campaigns" />
          <Field label="Medium" value={filters.medium} onChange={updateFilter("medium")} options={mediumOptions} placeholder="All Media" />
          <Field label="MPO Status" value={filters.mpoStatus} onChange={updateFilter("mpoStatus")} options={MPO_STATUS_OPTIONS.map(option => ({ value: option.value, label: option.label }))} placeholder="All Statuses" />
          <Field label="Payment Status" value={filters.paymentStatus} onChange={updateFilter("paymentStatus")} options={[
            { value: "unpaid", label: "Unpaid" },
            { value: "processing", label: "Processing" },
            { value: "paid", label: "Paid" },
          ]} placeholder="All Payment States" />
          <Field label="Reconciliation" value={filters.reconciliationStatus} onChange={updateFilter("reconciliationStatus")} options={[
            { value: "not_started", label: "Not Started" },
            { value: "in_progress", label: "In Progress" },
            { value: "completed", label: "Completed" },
            { value: "disputed", label: "Disputed" },
          ]} placeholder="All Reconciliation States" />
          <div style={{ gridColumn: "1 / -1" }}>
            <label style={{ fontSize: 11, fontWeight: 600, color: "var(--text3)", textTransform: "uppercase", letterSpacing: ".08em", marginBottom: 5, display: "block" }}>Search</label>
            <input value={filters.search} onChange={e => updateFilter("search")(e.target.value)} placeholder="Search MPO no., vendor, client, brand, campaign, medium, or finance status…" style={financeInputStyle} />
          </div>
        </div>

        {filterDescriptor ? (
          <div style={{ marginTop: 14, fontSize: 12, color: "var(--text2)", padding: "10px 12px", borderRadius: 10, background: "var(--bg3)", border: "1px solid var(--border)" }}>
            <strong style={{ color: "var(--text)" }}>Active filters:</strong> {filterDescriptor}
          </div>
        ) : null}
      </Card>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(220px,1fr))", gap: 14, marginBottom: 24 }}>
        {summaryCards.map(card => {
          const displayValue = typeof card.value === "number"
            ? (countSummaryLabels.has(card.label) ? formatCount(card.value) : fmtN(card.value))
            : card.value;
          return (
            <Card key={card.label} hoverable style={{ position: "relative", overflow: "hidden", padding: 18, minHeight: 148, display: "flex", flexDirection: "column", justifyContent: "space-between" }}>
              <div style={{ position: "absolute", top: -14, right: -10, fontSize: 64, opacity: .05 }}>{card.icon}</div>
              <div style={{ fontSize: 11, color: "var(--text3)", fontWeight: 600, textTransform: "uppercase", letterSpacing: ".08em", marginBottom: 10, paddingRight: 28 }}>{card.label}</div>
              <div
                title={displayValue}
                style={{
                  fontSize: typeof card.value === "number" ? "clamp(1.7rem, 2vw, 2.35rem)" : "clamp(1.35rem, 1.8vw, 1.9rem)",
                  fontWeight: 800,
                  fontFamily: "Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif",
                  fontVariantNumeric: "tabular-nums lining-nums",
                  letterSpacing: "-0.04em",
                  lineHeight: 1.05,
                  color: card.color,
                  whiteSpace: "nowrap",
                  overflow: "hidden",
                  textOverflow: "ellipsis",
                  maxWidth: "100%",
                }}
              >
                {displayValue}
              </div>
              <div style={{ fontSize: 12, color: "var(--text2)", marginTop: 10, lineHeight: 1.45 }}>{card.sub}</div>
            </Card>
          );
        })}
      </div>

      {filteredMpos.length === 0 ? (
        <Card>
          <Empty icon="📊" title="No reporting data for this filter set" sub="Adjust your filters or add more MPO / campaign activity to see finance control insights." />
        </Card>
      ) : (
        <div style={{ display: "flex", flexDirection: "column", gap: 18 }}>
          {sectionConfigs.map(section => (
            <Card key={section.id}>
              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, flexWrap: "wrap", marginBottom: 14 }}>
                <div>
                  <h2 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 16 }}>{section.icon} {section.title}</h2>
                  <p style={{ color: "var(--text2)", fontSize: 12, marginTop: 3 }}>{section.desc}</p>
                </div>
                <Btn variant="blue" size="sm" icon="⬇" onClick={() => buildSectionExport(section.title, section.headers, formatExportRowsWithCurrency(section.rows, section.currencyColumns || []), filterDescriptor)}>
                  Export Section
                </Btn>
              </div>

              {section.rows.length === 0 ? (
                <Empty icon={section.icon} title={`No data for ${section.title}`} sub="Nothing matches the current filters." />
              ) : (
                <div style={{ overflowX: "auto" }}>
                  <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 760 }}>
                    <thead>
                      <tr style={{ background: "var(--bg3)" }}>
                        {section.headers.map(header => (
                          <th key={header} style={{ padding: "8px 10px", textAlign: "left", fontSize: 10, fontWeight: 600, textTransform: "uppercase", letterSpacing: ".07em", color: "var(--text3)", borderBottom: "1px solid var(--border)", whiteSpace: "nowrap" }}>
                            {header}
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {section.rows.slice(0, 12).map((row, rowIndex) => (
                        <tr
                          key={`${section.id}-${rowIndex}`}
                          style={{ borderBottom: "1px solid var(--border)", background: rowIndex % 2 === 0 ? "transparent" : "rgba(255,255,255,.01)" }}
                          onMouseEnter={e => e.currentTarget.style.background = "var(--bg3)"}
                          onMouseLeave={e => e.currentTarget.style.background = rowIndex % 2 === 0 ? "transparent" : "rgba(255,255,255,.01)"}
                        >
                          {row.map((cell, cellIndex) => {
                            const isCountColumn = sectionCountColumns[section.id]?.has(cellIndex);
                            const display = typeof cell === "number"
                              ? (isCountColumn ? formatCount(cell) : (Math.abs(cell) >= 1000 ? fmtN(cell) : cell))
                              : cell;
                            return (
                              <td key={cellIndex} style={{ padding: "8px 10px", fontSize: 12, color: cellIndex === 0 ? "var(--text)" : "var(--text2)", fontWeight: cellIndex === 0 ? 600 : 400, whiteSpace: cellIndex < 2 ? "nowrap" : "normal" }}>
                                {display}
                              </td>
                            );
                          })}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                  {section.rows.length > 12 ? (
                    <div style={{ marginTop: 10, fontSize: 12, color: "var(--text3)" }}>
                      Showing 12 of {section.rows.length} rows in-app. Use <strong style={{ color: "var(--text)" }}>Export Section</strong> for the full dataset.
                    </div>
                  ) : null}
                </div>
              )}
            </Card>
          ))}
        </div>
      )}
    </div>
  );
};
