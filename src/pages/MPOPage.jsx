import { useEffect, useRef, useState } from "react";
import Badge from "../components/Badge";
import Empty from "../components/Empty";
import Toast from "../components/Toast";
import Modal from "../components/Modal";
import Confirm from "../components/Confirm";
import { Btn, Field, AttachmentField, Card, Stat } from "../components/ui/primitives";
import PrintPreview from "../components/mpo/PrintPreview";
import DailyCalendar from "../components/mpo/DailyCalendar";
import { buildCSV, sanitizeMPOForExport, buildMPOHTML } from "../utils/export";
import { activeOnly, archivedOnly, isArchived } from "../utils/records";
import { fmtN } from "../utils/formatters";
import { formatRoleLabel, hasPermission, readOnlyMessage } from "../constants/roles";
import {
  MPO_STATUS_OPTIONS,
  MPO_STATUS_LABELS,
  MPO_EXECUTION_STATUS_OPTIONS,
  MPO_INVOICE_STATUS_OPTIONS,
  MPO_PROOF_STATUS_OPTIONS,
  MPO_PAYMENT_STATUS_OPTIONS,
  MPO_RECON_STATUS_OPTIONS,
  toIsoInput,
  toIsoOrNull,
  getExecutionHealthColor,
  getExecutionHealthLabel,
  getAllowedMpoStatusTargets,
  mpoStatusNeedsNote,
  getMpoWorkflowMeta,
  isMpoAwaitingUser,
  getWorkflowActionLabel,
  getWorkflowActionVariant,
  getQuickWorkflowActions,
  canEditMpoContent,
  MPO_STATUS_COLORS,
} from "../constants/mpoWorkflow";
import { DEFAULT_APP_SETTINGS } from "../constants/appDefaults";
import {
  uploadMpoAttachmentAndGetUrl,
  createMpoInSupabase,
  updateMpoInSupabase,
  archiveMpoInSupabase,
  restoreMpoInSupabase,
  updateMpoStatusInSupabase,
  updateMpoExecutionInSupabase,
  generateNextMpoNoFromSupabase,
} from "../services/mpos";
import {
  createAuditEventInSupabase,
  fetchAuditEventsForRecord,
  notifyMpoWorkflowTransition,
  notifyExecutionUpdate,
} from "../services/notifications";

const store = {
  get: (k, d = null) => { try { const v = localStorage.getItem(k); return v ? JSON.parse(v) : d; } catch { return d; } },
  set: (k, v) => { try { localStorage.setItem(k, JSON.stringify(v)); } catch {} },
  del: (k) => { try { localStorage.removeItem(k); } catch {} },
};
const uid = () => Math.random().toString(36).slice(2) + Date.now().toString(36);
const collapseDaysToCounts = (days = []) =>
  (days || []).reduce((acc, day) => {
    const key = Number(day);
    if (!key) return acc;
    acc[key] = (acc[key] || 0) + 1;
    return acc;
  }, {});

const expandCountsToDays = (dayCounts = {}) =>
  Object.entries(dayCounts || {})
    .flatMap(([day, count]) =>
      Array.from({ length: Math.max(0, Number(count) || 0) }, () => Number(day))
    )
    .sort((a, b) => a - b);

const totalCountFromDayCounts = (dayCounts = {}) =>
  Object.values(dayCounts || {}).reduce((sum, count) => sum + (Number(count) || 0), 0);
const roundMoneyValue = (value, settings = {}) => {
  const num = Number(value) || 0;
  return settings?.roundToWholeNaira ? Math.round(num) : Math.round(num * 100) / 100;
};
const formatAuditTimestamp = (value) => {
  if (!value) return "—";
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? String(value) : date.toLocaleString();
};

const normalizePdfText = (value) => String(value ?? "")
  .replace(/[\u2018\u2019]/g, "'")
  .replace(/[\u201C\u201D]/g, '"')
  .replace(/[\u2013\u2014]/g, '-')
  .replace(/\u2026/g, '...')
  .replace(/[^\x20-\x7E\n]/g, ' ');

const escapePdfText = (value) => normalizePdfText(value)
  .replace(/\\/g, '\\\\')
  .replace(/\(/g, '\\(')
  .replace(/\)/g, '\\)');

const wrapPdfText = (value, maxChars = 90) => {
  const text = normalizePdfText(value).replace(/\s+/g, ' ').trim();
  if (!text) return [''];
  const words = text.split(' ');
  const lines = [];
  let current = '';
  words.forEach((word) => {
    const next = current ? `${current} ${word}` : word;
    if (next.length <= maxChars) {
      current = next;
    } else {
      if (current) lines.push(current);
      current = word.length > maxChars ? `${word.slice(0, maxChars - 1)}-` : word;
    }
  });
  if (current) lines.push(current);
  return lines.length ? lines : [''];
};

const truncatePdfCell = (value, maxChars) => {
  const text = normalizePdfText(value).replace(/\s+/g, ' ').trim();
  if (text.length <= maxChars) return text;
  return `${text.slice(0, Math.max(0, maxChars - 3)).trim()}...`;
};

const buildMpoPdfBytes = (mpo) => {
  const pageWidth = 595.28;
  const pageHeight = 841.89;
  const marginX = 42;
  const topY = 804;
  const bottomY = 50;
  const lineHeight = 14;
  const pages = [];
  let commands = [];
  let cursorY = topY;

  const pushPage = () => {
    if (commands.length) pages.push(commands.join('\n'));
    commands = [];
    cursorY = topY;
  };

  const ensureSpace = (linesNeeded = 1) => {
    if (cursorY - (linesNeeded * lineHeight) < bottomY) pushPage();
  };

  const addTextLine = (text, options = {}) => {
    const size = options.size || 10;
    const x = options.x ?? marginX;
    const y = options.y ?? cursorY;
    commands.push(`BT /F1 ${size} Tf 1 0 0 1 ${x.toFixed(2)} ${y.toFixed(2)} Tm (${escapePdfText(text)}) Tj ET`);
    if (!options.absolute) cursorY = y - (options.lineHeight || lineHeight);
  };

  const addWrappedBlock = (text, options = {}) => {
    const lines = wrapPdfText(text, options.maxChars || 90);
    ensureSpace(lines.length);
    lines.forEach((line) => addTextLine(line, options));
  };

  const addSpacer = (amount = 8) => {
    cursorY -= amount;
  };

  const addSectionTitle = (text) => {
    ensureSpace(2);
    addTextLine(text.toUpperCase(), { size: 11 });
    addSpacer(2);
  };

  const addLabelValue = (label, value) => {
    const line = `${label}: ${value || '-'}`;
    addWrappedBlock(line, { size: 10, maxChars: 92 });
  };

  const addTableHeader = () => {
    ensureSpace(2);
    addTextLine('Programme                      Time Belt            Material                  Dur   Spots   Rate        Gross', { size: 9 });
    addTextLine('----------------------------------------------------------------------------------------------------------', { size: 9, lineHeight: 12 });
  };

  const currency = (value) => {
    const num = Number(value) || 0;
    try {
      return `N${num.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
    } catch (_error) {
      return `N${num.toFixed(2)}`;
    }
  };

  pushPage();
  addTextLine('MEDIA PURCHASE ORDER (MPO)', { size: 18 });
  addSpacer(4);
  addLabelValue('MPO Number', mpo.mpoNo || 'Draft');
  addLabelValue('Date', mpo.date || '-');
  addLabelValue('Period', `${(mpo.months || []).length ? mpo.months.join(', ') : (mpo.month || '-')} ${mpo.year || ''}`.trim());
  addLabelValue('Client', mpo.clientName || '-');
  addLabelValue('Brand', mpo.brand || '-');
  addLabelValue('Campaign', mpo.campaignName || '-');
  addLabelValue('Vendor', mpo.vendorName || '-');
  addLabelValue('Medium', mpo.medium || '-');
  addLabelValue('Transmit Instruction', mpo.transmitMsg || '-');

  addSpacer(6);
  addSectionTitle('Spot Schedule');
  addTableHeader();

  (mpo.spots || []).forEach((spot) => {
    ensureSpace(2);
    const row = [
      truncatePdfCell(spot.programme || '-', 28).padEnd(28, ' '),
      truncatePdfCell(spot.timeBelt || spot.wd || '-', 18).padEnd(18, ' '),
      truncatePdfCell(spot.material || '-', 24).padEnd(24, ' '),
      String(spot.duration || '-').padStart(4, ' '),
      String(spot.spots || '0').padStart(6, ' '),
      truncatePdfCell(currency(spot.ratePerSpot || 0), 11).padStart(11, ' '),
      truncatePdfCell(currency((Number(spot.spots) || 0) * (Number(spot.ratePerSpot) || 0)), 12).padStart(12, ' '),
    ].join(' ');
    addTextLine(row, { size: 9, lineHeight: 12 });
    if (spot.scheduleMonth || (spot.calendarDays || []).length) {
      const metaParts = [];
      if (spot.scheduleMonth) metaParts.push(`Month: ${spot.scheduleMonth}`);
      if ((spot.calendarDays || []).length) metaParts.push(`Days: ${(spot.calendarDays || []).join(', ')}`);
      addWrappedBlock(`   ${metaParts.join(' | ')}`, { size: 8, maxChars: 110 });
    }
  });

  addSpacer(6);
  addSectionTitle('Cost Summary');
  [
    ['Total Spots', mpo.totalSpots],
    ['Total Gross Value', currency(mpo.totalGross)],
    ['Discount', currency(mpo.discAmt)],
    ['Net After Discount', currency(mpo.lessDisc)],
    ['Agency Commission', currency(mpo.commAmt)],
    ['After Commission', currency(mpo.afterComm)],
    [mpo.surchLabel || 'Surcharge', currency(mpo.surchAmt || 0)],
    ['Net Value', currency(mpo.netVal)],
    [`VAT (${mpo.vatPct || 7.5}%)`, currency(mpo.vatAmt)],
    ['Total Amount Payable', currency(mpo.grandTotal)],
  ].forEach(([label, value]) => addLabelValue(label, value));

  addSpacer(6);
  addSectionTitle('Prepared By / Signatories');
  addLabelValue('Prepared By', mpo.preparedBy || '-');
  addLabelValue('Prepared Title', mpo.preparedTitle || '-');
  addLabelValue('Prepared Contact', mpo.preparedContact || '-');
  addLabelValue('Signed By', mpo.signedBy || '-');
  addLabelValue('Signed Title', mpo.signedTitle || '-');
  addLabelValue('Agency Address', mpo.agencyAddress || '-');
  addLabelValue('Agency Email', mpo.agencyEmail || '-');
  addLabelValue('Agency Phone', mpo.agencyPhone || '-');

  addSpacer(10);
  addWrappedBlock('This PDF was generated directly as a document file from the MPO data.', { size: 8, maxChars: 100 });

  if (commands.length) pages.push(commands.join('\n'));

  const objects = [];
  const addObject = (body) => {
    objects.push(body);
    return objects.length;
  };

  const fontId = addObject('<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>');
  const pageIds = [];
  const contentIds = [];
  pages.forEach((pageContent) => {
    const stream = `${pageContent}\nBT /F1 8 Tf 1 0 0 1 ${marginX.toFixed(2)} 24 Tm (Page ${pageIds.length + 1} of ${pages.length}) Tj ET`;
    const contentId = addObject(`<< /Length ${stream.length} >>\nstream\n${stream}\nendstream`);
    contentIds.push(contentId);
    const pageId = addObject(`<< /Type /Page /Parent 0 0 R /MediaBox [0 0 ${pageWidth} ${pageHeight}] /Resources << /Font << /F1 ${fontId} 0 R >> >> /Contents ${contentId} 0 R >>`);
    pageIds.push(pageId);
  });

  const pagesId = addObject(`<< /Type /Pages /Kids [${pageIds.map((id) => `${id} 0 R`).join(' ')}] /Count ${pageIds.length} >>`);
  pageIds.forEach((pageId) => {
    objects[pageId - 1] = objects[pageId - 1].replace('/Parent 0 0 R', `/Parent ${pagesId} 0 R`);
  });
  const catalogId = addObject(`<< /Type /Catalog /Pages ${pagesId} 0 R >>`);

  let pdf = '%PDF-1.4\n';
  const offsets = [0];
  objects.forEach((body, index) => {
    offsets.push(pdf.length);
    pdf += `${index + 1} 0 obj\n${body}\nendobj\n`;
  });
  const xrefStart = pdf.length;
  pdf += `xref\n0 ${objects.length + 1}\n`;
  pdf += '0000000000 65535 f \n';
  offsets.slice(1).forEach((offset) => {
    pdf += `${String(offset).padStart(10, '0')} 00000 n \n`;
  });
  pdf += `trailer\n<< /Size ${objects.length + 1} /Root ${catalogId} 0 R >>\nstartxref\n${xrefStart}\n%%EOF`;

  return new TextEncoder().encode(pdf);
};

export default function MPOPage({ vendors, clients, campaigns, rates, mpos, setMpos, user, appSettings }) {
  const canManage = hasPermission(user, "manageMpos");
  const canManageStatus = hasPermission(user, "manageMpoStatus") || canManage;
  const [view, setView] = useState("list");
  const [editId, setEditId] = useState(null);
  const [step, setStep] = useState(1);
  const [viewMode, setViewMode] = useState("active");
  const [toast, setToast] = useState(null);
  const [confirm, setConfirm] = useState(null);
  const [preview, setPreview] = useState(null);
  const [historyModal, setHistoryModal] = useState(null);
  const [statusModal, setStatusModal] = useState(null);
  const [executionModal, setExecutionModal] = useState(null);
  const [executionUploading, setExecutionUploading] = useState({ signedMpo: false, invoice: false, proof: false });
  const [searchTerm, setSearchTerm] = useState("");
  const [statusFilter, setStatusFilter] = useState("all");
  const workflowPanelStorageKey = `msp_mpo_workflow_panel_${user?.id || "guest"}`;
  const [workflowPanelOpen, setWorkflowPanelOpen] = useState(() => store.get(workflowPanelStorageKey, true));

  const VAT_RATE = parseFloat(appSettings?.vatRate) || 7.5;
  const [surcharge, setSurcharge] = useState({ pct: "", label: "" });
  const [vendorPickerOpen, setVendorPickerOpen] = useState(false);
  const [vendorSearch, setVendorSearch] = useState("");

  // Auto-generate MPO number: BRA001-JAN2025
  const genMpoNo = (brand, existingMpos) => {
    const prefix = (brand || "MPO").replace(/\s+/g,"").toUpperCase().slice(0,3);
    const monthNames = ["JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"];
    const now = new Date();
    const mon = monthNames[now.getMonth()];
    const yr = String(now.getFullYear()).slice(-2);
    const count = existingMpos.filter(m => (m.mpoNo||"").startsWith(prefix)).length + 1;
    return `${prefix}${String(count).padStart(3,"0")}-${mon}${yr}`;
  };

  const blankMPO = (brand = "", existingMpos = []) => ({
    campaignId: "", vendorId: "",
    mpoNo: existingMpos && existingMpos.length ? genMpoNo(brand, existingMpos) : "Pending auto-number on save",
    date: new Date().toISOString().slice(0, 10),
    month: "", months: [], year: new Date().getFullYear().toString(),
    medium: "",
    signedBy: "", signedTitle: "",
    preparedBy: user?.name || "", preparedContact: user?.phone || user?.email || "",
    preparedTitle: user?.title || "",
    preparedSignature: user?.signatureDataUrl || "",
    signedSignature: "",
    agencyAddress: user?.agencyAddress || "5, Craig Street, Ogudu GRA, Lagos",
    agencyEmail: user?.agencyEmail || "",
    agencyPhone: user?.agencyPhone || "",
    transmitMsg: "", status: "draft"
  });

  const [mpoData, setMpoData] = useState(() => blankMPO("", mpos));
  const [spots, setSpots] = useState([]);
  const [spotModal, setSpotModal] = useState(null);

  // Daily schedule — calendar mode
  // calRows: array of { id, programme, timeBelt, material, duration, rateId, customRate, dayCounts: { [day]: count } }
  const [dailyMode, setDailyMode] = useState(false);
  const blankCalRow = () => ({ id: uid(), programme: "", timeBelt: "", material: "", materialCustom: "", duration: "30", rateId: "", customRate: "", dayCounts: {} });
  // Multi-month: calData maps "Month Year" -> calRows array
  const [calData, setCalData] = useState({});
  const [activeCalMonth, setActiveCalMonth] = useState("");
  // Keep a stable first blank row per month so the first programme selection updates immediately.
  const calMonthFallbackRowsRef = useRef({});
  // Legacy single-month calRows still used for single-month mode
  const [calRows, setCalRows] = useState([blankCalRow()]);

  const getStableCalRows = (monthLabel) => {
    if (calData[monthLabel]) return calData[monthLabel];
    if (!calMonthFallbackRowsRef.current[monthLabel]) {
      calMonthFallbackRowsRef.current[monthLabel] = [blankCalRow()];
    }
    return calMonthFallbackRowsRef.current[monthLabel];
  };

  const setStableCalRows = (monthLabel, updater) => {
    setCalData(prev => {
      const baseRows = prev[monthLabel] || calMonthFallbackRowsRef.current[monthLabel] || [blankCalRow()];
      const nextRows = typeof updater === "function" ? updater(baseRows) : updater;
      calMonthFallbackRowsRef.current[monthLabel] = nextRows;
      return { ...prev, [monthLabel]: nextRows };
    });
  };

  const blankSpot = { programme: "", wd: "", timeBelt: "", material: "", materialCustom: "", duration: "30", rateId: "", customRate: "", spots: "", calendarDays: [], calendarDayCounts: {} };
  const [spotForm, setSpotForm] = useState(blankSpot);
  const [editSpotId, setEditSpotId] = useState(null);
  const draftKey = "msp_mpo_draft";
  const [hasSavedDraft, setHasSavedDraft] = useState(() => Boolean(store.get(draftKey)));
  const upd = k => v => setMpoData(m => ({ ...m, [k]: v }));
  const updS = k => v => setSpotForm(f => ({ ...f, [k]: v }));
  const months = ["January","February","March","April","May","June","July","August","September","October","November","December"];
  const campaign = campaigns.find(c => c.id === mpoData.campaignId);
  const client = clients.find(c => c.id === campaign?.clientId);
  const vendor = vendors.find(v => v.id === mpoData.vendorId);
  const vendorRates = activeOnly(rates).filter(r => r.vendorId === mpoData.vendorId);
  const totalSpots = spots.reduce((s, r) => s + (parseFloat(r.spots) || 0), 0);
  const totalGross = spots.reduce((s, r) => s + (parseFloat(r.spots) || 0) * (parseFloat(r.ratePerSpot) || 0), 0);
  const discPct = vendor ? (parseFloat(vendor.discount) || 0) / 100 : 0;
  const commPct = vendor ? (parseFloat(vendor.commission) || 0) / 100 : 0;
  const discAmt = roundMoneyValue(totalGross * discPct, appSettings);
  const lessDisc = roundMoneyValue(totalGross - discAmt, appSettings);
  const commAmt = roundMoneyValue(lessDisc * commPct, appSettings);
  const afterComm = roundMoneyValue(lessDisc - commAmt, appSettings);
  const surchPct = (parseFloat(surcharge.pct) || 0) / 100;
  const surchAmt = roundMoneyValue(afterComm * surchPct, appSettings);
  const netVal = roundMoneyValue(afterComm + surchAmt, appSettings);
  const vatAmt = roundMoneyValue(netVal * VAT_RATE / 100, appSettings);
  const grandTotal = roundMoneyValue(netVal + vatAmt, appSettings);
  const rateOptions = vendorRates.map(r => ({ value: r.id, label: `${r.programme || "Unnamed"} – ${fmtN(r.ratePerSpot)}` }));
  const filteredVendors = vendors.filter(v => {
    const q = vendorSearch.trim().toLowerCase();
    if (!q) return true;
    return [v.name, v.type, v.contactName, v.email, v.phone]
      .filter(Boolean)
      .some(value => String(value).toLowerCase().includes(q));
  });

  useEffect(() => {
    setWorkflowPanelOpen(store.get(`msp_mpo_workflow_panel_${user?.id || "guest"}`, true));
  }, [user?.id]);
  useEffect(() => {
    store.set(workflowPanelStorageKey, workflowPanelOpen);
  }, [workflowPanelStorageKey, workflowPanelOpen]);

  // When campaign changes, auto-generate MPO number with brand
  useEffect(() => {
    if (campaign?.brand && !editId) {
      setMpoData(m => ({ ...m, mpoNo: genMpoNo(campaign.brand, mpos) }));
    }
  }, [mpoData.campaignId]);

  useEffect(() => {
    if (view !== "form") return;
    const payload = { editId, step, mpoData, spots, surcharge, savedAt: Date.now() };
    store.set(draftKey, payload);
    setHasSavedDraft(true);
  }, [view, editId, step, mpoData, spots, surcharge]);

  useEffect(() => {
    const handler = (e) => {
      if (view === "form") {
        e.preventDefault();
        e.returnValue = "";
      }
    };
    window.addEventListener("beforeunload", handler);
    return () => window.removeEventListener("beforeunload", handler);
  }, [view]);

  const resumeSavedDraft = () => {
    const draft = store.get(draftKey);
    if (!draft) return setToast({ msg: "No saved draft found.", type: "error" });
    setEditId(draft.editId || null);
    setStep(draft.step || 1);
    setMpoData({ ...blankMPO("", mpos), ...(draft.mpoData || {}), preparedSignature: draft?.mpoData?.preparedSignature || user?.signatureDataUrl || "", signedSignature: draft?.mpoData?.signedSignature || "" });
    setSpots(draft.spots || []);
    setSurcharge(draft.surcharge || { pct: "", label: "" });
    setView("form");
    setToast({ msg: "Saved MPO draft restored.", type: "success" });
  };

  const clearSavedDraft = () => {
    store.del(draftKey);
    setHasSavedDraft(false);
  };

  const refreshHistoryModal = async (mpoId, title) => {
    try {
      const events = await fetchAuditEventsForRecord(user.agencyId, "mpo", mpoId);
      setHistoryModal({ mpoId, title, events });
    } catch (error) {
      setToast({ msg: error.message || "Failed to load MPO history.", type: "error" });
    }
  };

  const openMpoHistory = async (mpo) => {
    await refreshHistoryModal(mpo.id, `MPO History — ${mpo.mpoNo || mpo.id}`);
  };

  const openExecutionModal = (mpo) => {
    setExecutionModal({
      mpoId: mpo.id,
      mpoNo: mpo.mpoNo || mpo.id,
      dispatchStatus: mpo.dispatchStatus || "pending",
      dispatchedAt: toIsoInput(mpo.dispatchedAt),
      dispatchContact: mpo.dispatchContact || mpo.vendorName || "",
      dispatchNote: mpo.dispatchNote || "",
      signedMpoUrl: mpo.signedMpoUrl || "",
      invoiceStatus: mpo.invoiceStatus || "pending",
      invoiceNo: mpo.invoiceNo || "",
      invoiceAmount: String(mpo.invoiceAmount ?? mpo.grandTotal ?? ""),
      invoiceReceivedAt: toIsoInput(mpo.invoiceReceivedAt),
      invoiceUrl: mpo.invoiceUrl || "",
      proofStatus: mpo.proofStatus || "pending",
      proofUrl: mpo.proofUrl || "",
      proofReceivedAt: toIsoInput(mpo.proofReceivedAt),
      plannedSpotsExecution: String(mpo.plannedSpotsExecution ?? mpo.totalSpots ?? 0),
      airedSpots: String(mpo.airedSpots ?? 0),
      missedSpots: String(mpo.missedSpots ?? 0),
      makegoodSpots: String(mpo.makegoodSpots ?? 0),
      reconciliationStatus: mpo.reconciliationStatus || "not_started",
      reconciliationNotes: mpo.reconciliationNotes || "",
      reconciledAmount: String(mpo.reconciledAmount ?? mpo.grandTotal ?? 0),
      paymentStatus: mpo.paymentStatus || "unpaid",
      paymentReference: mpo.paymentReference || "",
      paidAt: toIsoInput(mpo.paidAt),
    });
  };

  const uploadExecutionAttachment = async (kind, file) => {
    if (!file) return;
    if (!executionModal?.mpoId) {
      setToast({ msg: "Save or open an existing MPO before uploading attachments.", type: "error" });
      return;
    }
    if (!user?.agencyId) {
      setToast({ msg: "No agency found for this workspace.", type: "error" });
      return;
    }
    const keyMap = {
      signedMpo: "signedMpoUrl",
      invoice: "invoiceUrl",
      proof: "proofUrl",
    };
    const fieldName = keyMap[kind] || "signedMpoUrl";
    setExecutionUploading(state => ({ ...state, [kind]: true }));
    try {
      const url = await uploadMpoAttachmentAndGetUrl({
        agencyId: user.agencyId,
        mpoId: executionModal.mpoId,
        kind,
        file,
      });
      setExecutionModal(modal => ({ ...modal, [fieldName]: url }));
      setToast({ msg: `${kind === "signedMpo" ? "Signed MPO" : kind === "invoice" ? "Invoice" : "Proof"} uploaded. Save execution to persist it.`, type: "success" });
    } catch (e) {
      setToast({ msg: e.message || "Upload failed.", type: "error" });
    } finally {
      setExecutionUploading(state => ({ ...state, [kind]: false }));
    }
  };

  const saveExecutionModal = async () => {
    if (!executionModal?.mpoId) return;
    try {
      const patch = {
        dispatchStatus: executionModal.dispatchStatus,
        dispatchedAt: toIsoOrNull(executionModal.dispatchedAt),
        dispatchedBy: user?.id || null,
        dispatchContact: executionModal.dispatchContact,
        dispatchNote: executionModal.dispatchNote,
        signedMpoUrl: executionModal.signedMpoUrl,
        invoiceStatus: executionModal.invoiceStatus,
        invoiceNo: executionModal.invoiceNo,
        invoiceAmount: executionModal.invoiceAmount,
        invoiceReceivedAt: toIsoOrNull(executionModal.invoiceReceivedAt),
        invoiceUrl: executionModal.invoiceUrl,
        proofStatus: executionModal.proofStatus,
        proofUrl: executionModal.proofUrl,
        proofReceivedAt: toIsoOrNull(executionModal.proofReceivedAt),
        plannedSpotsExecution: executionModal.plannedSpotsExecution,
        airedSpots: executionModal.airedSpots,
        missedSpots: executionModal.missedSpots,
        makegoodSpots: executionModal.makegoodSpots,
        reconciliationStatus: executionModal.reconciliationStatus,
        reconciliationNotes: executionModal.reconciliationNotes,
        reconciledAmount: executionModal.reconciledAmount,
        paymentStatus: executionModal.paymentStatus,
        paymentReference: executionModal.paymentReference,
        paidAt: toIsoOrNull(executionModal.paidAt),
      };
      const updated = await updateMpoExecutionInSupabase(executionModal.mpoId, patch);
      setMpos(ms => ms.map(x => x.id === executionModal.mpoId ? updated : x));
      await createAuditEventInSupabase({
        agencyId: user.agencyId,
        recordType: "mpo",
        recordId: executionModal.mpoId,
        action: "execution_updated",
        actor: user,
        note: `Execution & reconciliation updated for ${executionModal.mpoNo || executionModal.mpoId}.`,
        metadata: {
          mpoNo: executionModal.mpoNo || "",
          dispatchStatus: patch.dispatchStatus,
          invoiceStatus: patch.invoiceStatus,
          proofStatus: patch.proofStatus,
          reconciliationStatus: patch.reconciliationStatus,
          paymentStatus: patch.paymentStatus,
        },
      });
      await notifyExecutionUpdate({ agencyId: user.agencyId, mpo: executionModal, actor: user, patch });
      setExecutionModal(null);
      setToast({ msg: "Execution & reconciliation saved.", type: "success" });
    } catch (error) {
      setToast({ msg: error.message || "Failed to save execution details.", type: "error" });
    }
  };

  const applyMpoStatusChange = async (mpo, nextStatus, note = "") => {
    try {
      const updated = await updateMpoStatusInSupabase(mpo.id, nextStatus);
      setMpos(ms => ms.map(x => x.id === mpo.id ? updated : x));
      await createAuditEventInSupabase({
        agencyId: user.agencyId,
        recordType: "mpo",
        recordId: mpo.id,
        action: "status_changed",
        actor: user,
        note,
        metadata: {
          mpoNo: mpo.mpoNo || "",
          fromStatus: mpo.status || "draft",
          toStatus: nextStatus,
        },
      });
      await notifyMpoWorkflowTransition({ agencyId: user.agencyId, mpo, nextStatus, actor: user, note });
      setStatusModal(null);
      if (historyModal?.mpoId === mpo.id) {
        await refreshHistoryModal(mpo.id, historyModal.title || `MPO History — ${mpo.mpoNo || mpo.id}`);
      }
      setToast({ msg: `MPO moved to ${MPO_STATUS_LABELS[nextStatus] || nextStatus}.`, type: "success" });
    } catch (error) {
      setToast({ msg: error.message || "Failed to update MPO status.", type: "error" });
    }
  };

  const requestMpoStatusChange = (mpo, nextStatus) => {
    const current = String(mpo?.status || "draft").toLowerCase();
    const target = String(nextStatus || current).toLowerCase();
    if (target === current) return;
    const allowedTargets = getAllowedMpoStatusTargets(user, mpo);
    if (!allowedTargets.includes(target)) {
      setToast({ msg: `Your role (${formatRoleLabel(user?.role)}) cannot move this MPO from ${MPO_STATUS_LABELS[current] || current} to ${MPO_STATUS_LABELS[target] || target}.`, type: "error" });
      return;
    }
    const nextOwner = getMpoWorkflowMeta({ ...mpo, status: target });
    if (mpoStatusNeedsNote(target)) {
      setStatusModal({
        mpo,
        nextStatus: target,
        note: "",
        actionLabel: getWorkflowActionLabel(current, target),
        noteLabel: target === "rejected" ? "Request changes note" : "Workflow note",
        helperText: nextOwner?.hint || "",
        nextOwnerLabel: nextOwner?.label || "Next team member",
      });
      return;
    }
    applyMpoStatusChange(mpo, target, "");
  };

  const restoreMPO = async (id) => {
    if (!canManage) return setToast({ msg: readOnlyMessage(user), type: "error" });
    try {
      const restored = await restoreMpoInSupabase(id);
      setMpos(m => m.map(x => x.id === id ? restored : x));
      createAuditEventInSupabase({
        agencyId: user.agencyId,
        recordType: "mpo",
        recordId: id,
        action: "restored",
        actor: user,
        metadata: { mpoNo: restored.mpoNo || "", status: restored.status || "draft" },
      }).catch(error => console.error("Failed to write MPO audit event:", error));
      setToast({ msg: "MPO restored.", type: "success" });
    } catch (e) {
      setToast({ msg: e.message || "Failed to restore MPO.", type: "error" });
    }
  };

  const openNew = () => {
    const blank = blankMPO("", mpos);
    setMpoData(blank); setSpots([]); setStep(1); setEditId(null); setView("form");
    calMonthFallbackRowsRef.current = {};
    setDailyMode(false); setCalRows([blankCalRow()]); setCalData({}); setActiveCalMonth("");
    clearSavedDraft();
  };
  const openEdit = (mpo) => {
    if (!canEditMpoContent(user, mpo)) {
      setToast({ msg: `You can only edit MPOs in Draft or Rejected status. Current status: ${MPO_STATUS_LABELS[mpo?.status || "draft"] || (mpo?.status || "draft")}.`, type: "error" });
      return;
    }
    setMpoData({ campaignId: mpo.campaignId||"", vendorId: mpo.vendorId||"", mpoNo: mpo.mpoNo||"", date: mpo.date||"", month: mpo.month||"", months: mpo.months||[], year: mpo.year||"", medium: mpo.medium||"", signedBy: mpo.signedBy||"", signedTitle: mpo.signedTitle||"", preparedBy: mpo.preparedBy||user?.name||"", preparedContact: mpo.preparedContact||user?.phone||user?.email||"", preparedTitle: mpo.preparedTitle||user?.title||"", preparedSignature: mpo.preparedSignature||user?.signatureDataUrl||"", signedSignature: mpo.signedSignature||"", agencyAddress: mpo.agencyAddress||user?.agencyAddress||"", transmitMsg: mpo.transmitMsg||"", status: mpo.status||"draft" });
    setSurcharge({ pct: mpo.surchPct ? String((mpo.surchPct||0)*100) : "", label: mpo.surchLabel||"" });
    setSpots(mpo.spots || []); setEditId(mpo.id); setStep(1); setView("form");
    calMonthFallbackRowsRef.current = {};
    setDailyMode(false); setCalRows([blankCalRow()]); setCalData({}); setActiveCalMonth("");
    clearSavedDraft();
  };

  // Add spot rows from calendar schedule — supports single or multi-month
  const addFromCalendarSchedule = (rowsToAdd, monthLabel) => {
    const sourceRows = rowsToAdd || calRows;
    const newSpots = [];
    sourceRows.forEach(row => {
      const calendarDays = expandCountsToDays(row.dayCounts || {});
      if (!row.programme || calendarDays.length === 0) return;
      const rate = vendorRates.find(r => r.id === row.rateId);
      const ratePerSpot = parseFloat(row.customRate) || parseFloat(rate?.ratePerSpot) || 0;
      const matFinal = row.material === "__custom__" ? (row.materialCustom || "") : (row.material || "");
      newSpots.push({
        id: uid(),
        programme: row.programme,
        wd: "",
        timeBelt: row.timeBelt,
        material: matFinal,
        duration: row.duration,
        rateId: row.rateId,
        ratePerSpot,
        spots: calendarDays.length,
        calendarDays,
        scheduleMonth: monthLabel || mpoData.month
      });
    });
    if (!newSpots.length) return setToast({ msg: "Fill Programme and add at least one spot on at least one date per row.", type: "error" });
    setSpots(s => [...s, ...newSpots]);
    if (!rowsToAdd) { setDailyMode(false); setCalRows([blankCalRow()]); }
    setToast({ msg: `${newSpots.length} spot row(s) added from calendar.`, type: "success" });
  };

  const addSpot = () => {
    if (!spotForm.programme) return;
    const rate = vendorRates.find(r => r.id === spotForm.rateId);
    const ratePerSpot = parseFloat(spotForm.customRate) || parseFloat(rate?.ratePerSpot) || 0;
    const calDays = spotForm.calendarDays?.length
      ? spotForm.calendarDays
      : expandCountsToDays(spotForm.calendarDayCounts || {});
    const spotsCount = calDays.length > 0 ? calDays.length : (parseFloat(spotForm.spots) || 0);
    if (!spotsCount) return setToast({ msg: "Select at least one airing date or enter a spot count.", type: "error" });
    const newSpot = { id: uid(), ...spotForm, ratePerSpot, spots: String(spotsCount), calendarDays: calDays, calendarDayCounts: collapseDaysToCounts(calDays) };
    if (editSpotId) { setSpots(s => s.map(x => x.id === editSpotId ? newSpot : x)); setEditSpotId(null); }
    else setSpots(s => [...s, newSpot]);
    setSpotForm(blankSpot); setSpotModal(null);
  };

  const saveMPO = async () => {
    if (!canManage) return setToast({ msg: readOnlyMessage(user), type: "error" });
    const existingEditingMpo = editId ? mpos.find(m => m.id === editId) : null;
    if (existingEditingMpo && !canEditMpoContent(user, existingEditingMpo)) {
      return setToast({ msg: `You can only edit MPOs in Draft or Rejected status. Current status: ${MPO_STATUS_LABELS[existingEditingMpo.status || "draft"] || (existingEditingMpo.status || "draft")}.`, type: "error" });
    }
    if (!mpoData.campaignId) return setToast({ msg: "Please select a campaign.", type: "error" });
    if (!mpoData.vendorId) return setToast({ msg: "Please select a vendor.", type: "error" });
    if (!spots.length) return setToast({ msg: "Add at least one spot row before saving this MPO.", type: "error" });
    if (editId && !mpoData.mpoNo) return setToast({ msg: "MPO number is required.", type: "error" });
    const duplicate = editId && mpoData.mpoNo ? mpos.find(m => m.id !== editId && !isArchived(m) && (m.mpoNo || "").trim().toLowerCase() === mpoData.mpoNo.trim().toLowerCase()) : null;
    if (duplicate) return setToast({ msg: "This MPO number already exists.", type: "error" });
    const campaignBudget = parseFloat(campaign?.budget) || 0;
    if (campaignBudget > 0 && grandTotal > campaignBudget) return setToast({ msg: "This MPO total is above the campaign budget. Reduce spots or update the campaign budget first.", type: "error" });

    try {
      const generatedMpoNo = editId ? mpoData.mpoNo : await generateNextMpoNoFromSupabase(campaign?.brand || mpoData.brand || "MPO");
      const existingExec = editId ? (mpos.find(m => m.id === editId) || {}) : {};
      const record = { id: editId || uid(), ...mpoData, preparedSignature: mpoData.preparedSignature || user?.signatureDataUrl || "", signedSignature: mpoData.signedSignature || "", agencyEmail: mpoData.agencyEmail || user?.agencyEmail || "", agencyPhone: mpoData.agencyPhone || user?.agencyPhone || "", mpoNo: generatedMpoNo, vendorName: vendor?.name || "", clientName: client?.name || "", campaignName: campaign?.name || "", brand: campaign?.brand || "", medium: mpoData.medium || campaign?.medium || "", months: mpoData.months || [], spots, totalSpots, totalGross, discPct, discAmt, lessDisc, commPct, commAmt: commAmt, afterComm, surchPct, surchAmt, surchLabel: surcharge.label, netVal, vatPct: VAT_RATE, vatAmt, grandTotal, terms: appSettings?.mpoTerms || DEFAULT_APP_SETTINGS.mpoTerms, roundToWholeNaira: !!appSettings?.roundToWholeNaira,
        dispatchStatus: existingExec.dispatchStatus || "pending",
        dispatchedAt: existingExec.dispatchedAt || null,
        dispatchedBy: existingExec.dispatchedBy || null,
        dispatchContact: existingExec.dispatchContact || "",
        dispatchNote: existingExec.dispatchNote || "",
        signedMpoUrl: existingExec.signedMpoUrl || "",
        invoiceStatus: existingExec.invoiceStatus || "pending",
        invoiceNo: existingExec.invoiceNo || "",
        invoiceAmount: existingExec.invoiceAmount ?? 0,
        invoiceReceivedAt: existingExec.invoiceReceivedAt || null,
        invoiceUrl: existingExec.invoiceUrl || "",
        proofStatus: existingExec.proofStatus || "pending",
        proofUrl: existingExec.proofUrl || "",
        proofReceivedAt: existingExec.proofReceivedAt || null,
        plannedSpotsExecution: existingExec.plannedSpotsExecution ?? totalSpots,
        airedSpots: existingExec.airedSpots ?? 0,
        missedSpots: existingExec.missedSpots ?? 0,
        makegoodSpots: existingExec.makegoodSpots ?? 0,
        reconciliationStatus: existingExec.reconciliationStatus || "not_started",
        reconciliationNotes: existingExec.reconciliationNotes || "",
        reconciledAmount: existingExec.reconciledAmount ?? grandTotal,
        paymentStatus: existingExec.paymentStatus || "unpaid",
        paymentReference: existingExec.paymentReference || "",
        paidAt: existingExec.paidAt || null,
        createdAt: editId ? (mpos.find(m => m.id === editId)?.createdAt || Date.now()) : Date.now(), updatedAt: Date.now() };

      let saved;
      if (editId) {
        saved = { ...(await updateMpoInSupabase(editId, record)), preparedSignature: record.preparedSignature, signedSignature: record.signedSignature };
        setMpos(m => m.map(x => x.id === editId ? saved : x));
        createAuditEventInSupabase({
          agencyId: user.agencyId,
          recordType: "mpo",
          recordId: editId,
          action: "updated",
          actor: user,
          metadata: { mpoNo: saved.mpoNo || generatedMpoNo, status: saved.status || "draft", grandTotal: saved.grandTotal || grandTotal },
        }).catch(error => console.error("Failed to write MPO audit event:", error));
      } else {
        saved = { ...(await createMpoInSupabase(user.agencyId, user.id, record)), preparedSignature: record.preparedSignature, signedSignature: record.signedSignature };
        setMpos(m => [saved, ...m]);
        createAuditEventInSupabase({
          agencyId: user.agencyId,
          recordType: "mpo",
          recordId: saved.id,
          action: "created",
          actor: user,
          metadata: { mpoNo: saved.mpoNo || generatedMpoNo, status: saved.status || "draft", grandTotal: saved.grandTotal || grandTotal },
        }).catch(error => console.error("Failed to write MPO audit event:", error));
      }
      clearSavedDraft();
      setToast({ msg: editId ? "MPO updated!" : `MPO ${generatedMpoNo} saved!`, type: "success" });
      setView("list");
    } catch (e) {
      setToast({ msg: e.message || "Failed to save MPO.", type: "error" });
    }
  };

  const delMPO = async (id) => {
    if (!canManage) return setToast({ msg: readOnlyMessage(user), type: "error" });
    try {
      const archived = await archiveMpoInSupabase(id);
      setMpos(m => m.map(x => x.id === id ? archived : x));
      createAuditEventInSupabase({
        agencyId: user.agencyId,
        recordType: "mpo",
        recordId: id,
        action: "archived",
        actor: user,
        metadata: { mpoNo: archived.mpoNo || "", status: archived.status || "draft" },
      }).catch(error => console.error("Failed to write MPO audit event:", error));
      setConfirm(null);
      setToast({ msg: "MPO archived.", type: "success" });
    } catch (e) {
      setToast({ msg: e.message || "Failed to archive MPO.", type: "error" });
    }
  };

  const openPreview = (mpo) => {
    const safeMpo = sanitizeMPOForExport({ ...mpo, preparedSignature: mpo?.preparedSignature || user?.signatureDataUrl || "", signedSignature: mpo?.signedSignature || "", agencyEmail: mpo?.agencyEmail || user?.agencyEmail || "", agencyPhone: mpo?.agencyPhone || user?.agencyPhone || "" });
    const html = buildMPOHTML(safeMpo);
    const pdfBytes = buildMpoPdfBytes(safeMpo);
    const csvHeaders = ["Programme","WD","Time Belt","Material","Duration","Rate/Spot","Spots","Gross Value"];
    const csvRows = (mpo.spots || []).map(s => [s.programme, s.wd, s.timeBelt, s.material, s.duration+'"', s.ratePerSpot, s.spots, (parseFloat(s.spots||0))*(parseFloat(s.ratePerSpot||0))]);
    csvRows.push([]);
    csvRows.push(["","","","","","Total Spots","",mpo.totalSpots]);
    csvRows.push(["","","","","","Gross Value","",mpo.totalGross]);
    csvRows.push(["","","","","","Discount","",-Math.round(mpo.discAmt||0)]);
    csvRows.push(["","","","","","Net after Disc","",Math.round(mpo.lessDisc||0)]);
    csvRows.push(["","","","","","Agency Commission","",-Math.round(mpo.commAmt||0)]);
    csvRows.push(["","","","","","Net Value","",Math.round(mpo.netVal||0)]);
    csvRows.push(["","","","","",`VAT (${mpo.vatPct || 7.5}%)`,"",Math.round(mpo.vatAmt||0)]);
    csvRows.push(["","","","","","TOTAL PAYABLE","",Math.round(mpo.grandTotal||0)]);
    const csv = buildCSV(csvRows, csvHeaders);
    setPreview({ html, csv, title: `MPO — ${mpo.mpoNo || "Draft"} | ${mpo.vendorName} | ${mpo.month} ${mpo.year}` });
  };

  const statusColors = MPO_STATUS_COLORS;
  const visibleMpos = (viewMode === "archived" ? archivedOnly(mpos) : viewMode === "all" ? mpos : activeOnly(mpos)).filter(m => {
    const q = `${m.mpoNo || ""} ${m.vendorName || ""} ${m.clientName || ""} ${m.brand || ""}`.toLowerCase();
    return q.includes(searchTerm.toLowerCase()) && (statusFilter === "all" || (m.status || "draft") === statusFilter);
  });
  const workflowStats = {
    myQueue: visibleMpos.filter(m => isMpoAwaitingUser(user, m)).length,
    pendingReview: visibleMpos.filter(m => ["submitted", "reviewed"].includes(String(m.status || "draft").toLowerCase())).length,
    readyToSend: visibleMpos.filter(m => String(m.status || "draft").toLowerCase() === "approved").length,
    needsChanges: visibleMpos.filter(m => String(m.status || "draft").toLowerCase() === "rejected").length,
  };
  const myWorkflowQueue = [...visibleMpos]
    .filter(m => isMpoAwaitingUser(user, m))
    .sort((a, b) => (b.updatedAt || b.createdAt || 0) - (a.updatedAt || a.createdAt || 0))
    .slice(0, 6);
  const workflowLaneCounts = MPO_STATUS_OPTIONS.reduce((acc, option) => {
    acc[option.value] = visibleMpos.filter(m => String(m.status || "draft").toLowerCase() === option.value).length;
    return acc;
  }, {});
  const topQuickQueue = myWorkflowQueue.slice(0, 3);
  const visibleMpoCards = [...visibleMpos].sort((a, b) => b.createdAt - a.createdAt);

  if (view === "list") return (
    <div className="fade">
      {toast && <Toast msg={toast.msg} type={toast.type} onDone={() => setToast(null)} />}
      {confirm && <Confirm msg={confirm.msg} onYes={confirm.onYes} onNo={() => setConfirm(null)} />}
      {preview && <PrintPreview html={preview.html} csv={preview.csv} title={preview.title} onClose={() => setPreview(null)} />}
      {historyModal && (
        <Modal title={historyModal.title || "MPO History"} onClose={() => setHistoryModal(null)} width={760}>
          {historyModal.events?.length ? (
            <div style={{ display: "flex", flexDirection: "column", gap: 10, maxHeight: "70vh", overflowY: "auto" }}>
              {historyModal.events.map(event => (
                <div key={event.id} style={{ border: "1px solid var(--border)", borderRadius: 10, padding: "12px 14px", background: "var(--bg3)" }}>
                  <div style={{ display: "flex", justifyContent: "space-between", gap: 12, flexWrap: "wrap" }}>
                    <div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 14 }}>
                      {(event.action || "updated").replace(/_/g, " ").toUpperCase()}
                    </div>
                    <div style={{ fontSize: 12, color: "var(--text3)" }}>{formatAuditTimestamp(event.createdAt)}</div>
                  </div>
                  <div style={{ fontSize: 12, color: "var(--text2)", marginTop: 4 }}>
                    {event.actorName || "System"} · {formatRoleLabel(event.actorRole || "viewer")}
                  </div>
                  {event.note ? <div style={{ marginTop: 8, fontSize: 13 }}>{event.note}</div> : null}
                  {event.metadata && Object.keys(event.metadata || {}).length ? (
                    <div style={{ marginTop: 10, display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(180px,1fr))", gap: 8 }}>
                      {Object.entries(event.metadata).map(([key, value]) => (
                        <div key={key} style={{ background: "var(--bg2)", border: "1px solid var(--border)", borderRadius: 8, padding: "8px 10px", fontSize: 12 }}>
                          <div style={{ color: "var(--text3)", textTransform: "uppercase", letterSpacing: ".04em", fontSize: 10, fontWeight: 700 }}>{key.replace(/_/g, " ")}</div>
                          <div style={{ marginTop: 4, wordBreak: "break-word" }}>{Array.isArray(value) ? value.join(", ") : String(value ?? "—")}</div>
                        </div>
                      ))}
                    </div>
                  ) : null}
                </div>
              ))}
            </div>
          ) : <Empty icon="🕘" title="No history yet" sub="This MPO has no audit events yet." />}
        </Modal>
      )}
      {executionModal && (
        <Modal title={`Execution & Reconciliation — ${executionModal.mpoNo || "MPO"}`} onClose={() => setExecutionModal(null)} width={920}>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(240px,1fr))", gap: 14 }}>
            <Card style={{ padding: 14 }}>
              <div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, marginBottom: 10 }}>Dispatch & Documents</div>
              <div style={{ display: "grid", gap: 10 }}>
                <Field label="Dispatch Status" value={executionModal.dispatchStatus} onChange={value => setExecutionModal(m => ({ ...m, dispatchStatus: value }))} options={MPO_EXECUTION_STATUS_OPTIONS} />
                <Field label="Dispatched At" type="datetime-local" value={executionModal.dispatchedAt} onChange={value => setExecutionModal(m => ({ ...m, dispatchedAt: value }))} />
                <Field label="Vendor Contact Used" value={executionModal.dispatchContact} onChange={value => setExecutionModal(m => ({ ...m, dispatchContact: value }))} placeholder="Email, phone, or contact name" />
                <AttachmentField label="Signed MPO" url={executionModal.signedMpoUrl} onUrlChange={value => setExecutionModal(m => ({ ...m, signedMpoUrl: value }))} onFileSelected={file => uploadExecutionAttachment("signedMpo", file)} uploading={executionUploading.signedMpo} />
                <Field label="Dispatch Note" rows={3} value={executionModal.dispatchNote} onChange={value => setExecutionModal(m => ({ ...m, dispatchNote: value }))} placeholder="Who sent it, what was shared, any follow-up note..." />
              </div>
            </Card>
            <Card style={{ padding: 14 }}>
              <div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, marginBottom: 10 }}>Invoice & Payment</div>
              <div style={{ display: "grid", gap: 10 }}>
                <Field label="Invoice Status" value={executionModal.invoiceStatus} onChange={value => setExecutionModal(m => ({ ...m, invoiceStatus: value }))} options={MPO_INVOICE_STATUS_OPTIONS} />
                <Field label="Invoice Number" value={executionModal.invoiceNo} onChange={value => setExecutionModal(m => ({ ...m, invoiceNo: value }))} />
                <Field label="Invoice Amount" type="number" value={executionModal.invoiceAmount} onChange={value => setExecutionModal(m => ({ ...m, invoiceAmount: value }))} />
                <Field label="Invoice Received At" type="datetime-local" value={executionModal.invoiceReceivedAt} onChange={value => setExecutionModal(m => ({ ...m, invoiceReceivedAt: value }))} />
                <AttachmentField label="Invoice" url={executionModal.invoiceUrl} onUrlChange={value => setExecutionModal(m => ({ ...m, invoiceUrl: value }))} onFileSelected={file => uploadExecutionAttachment("invoice", file)} uploading={executionUploading.invoice} />
                <Field label="Payment Status" value={executionModal.paymentStatus} onChange={value => setExecutionModal(m => ({ ...m, paymentStatus: value }))} options={MPO_PAYMENT_STATUS_OPTIONS} />
                <Field label="Payment Reference" value={executionModal.paymentReference} onChange={value => setExecutionModal(m => ({ ...m, paymentReference: value }))} />
                <Field label="Paid At" type="datetime-local" value={executionModal.paidAt} onChange={value => setExecutionModal(m => ({ ...m, paidAt: value }))} />
              </div>
            </Card>
            <Card style={{ padding: 14, gridColumn: "1 / -1" }}>
              <div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, marginBottom: 10 }}>Proof of Airing & Reconciliation</div>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(180px,1fr))", gap: 10 }}>
                <Field label="Proof Status" value={executionModal.proofStatus} onChange={value => setExecutionModal(m => ({ ...m, proofStatus: value }))} options={MPO_PROOF_STATUS_OPTIONS} />
                <AttachmentField label="Proof of Airing" url={executionModal.proofUrl} onUrlChange={value => setExecutionModal(m => ({ ...m, proofUrl: value }))} onFileSelected={file => uploadExecutionAttachment("proof", file)} uploading={executionUploading.proof} />
                <Field label="Proof Received At" type="datetime-local" value={executionModal.proofReceivedAt} onChange={value => setExecutionModal(m => ({ ...m, proofReceivedAt: value }))} />
                <Field label="Planned Spots" type="number" value={executionModal.plannedSpotsExecution} onChange={value => setExecutionModal(m => ({ ...m, plannedSpotsExecution: value }))} />
                <Field label="Aired Spots" type="number" value={executionModal.airedSpots} onChange={value => setExecutionModal(m => ({ ...m, airedSpots: value }))} />
                <Field label="Missed Spots" type="number" value={executionModal.missedSpots} onChange={value => setExecutionModal(m => ({ ...m, missedSpots: value }))} />
                <Field label="Make-good Spots" type="number" value={executionModal.makegoodSpots} onChange={value => setExecutionModal(m => ({ ...m, makegoodSpots: value }))} />
                <Field label="Reconciliation Status" value={executionModal.reconciliationStatus} onChange={value => setExecutionModal(m => ({ ...m, reconciliationStatus: value }))} options={MPO_RECON_STATUS_OPTIONS} />
                <Field label="Reconciled Amount" type="number" value={executionModal.reconciledAmount} onChange={value => setExecutionModal(m => ({ ...m, reconciledAmount: value }))} />
              </div>
              <div style={{ marginTop: 10 }}>
                <Field label="Reconciliation Notes" rows={4} value={executionModal.reconciliationNotes} onChange={value => setExecutionModal(m => ({ ...m, reconciliationNotes: value }))} placeholder="Delivered spots, discrepancies, make-goods, invoice notes, approvals..." />
              </div>
              <div style={{ display: "flex", justifyContent: "flex-end", gap: 10, marginTop: 14 }}>
                <Btn variant="ghost" onClick={() => setExecutionModal(null)}>Cancel</Btn>
                <Btn onClick={saveExecutionModal} loading={executionUploading.signedMpo || executionUploading.invoice || executionUploading.proof}>Save Execution</Btn>
              </div>
            </Card>
          </div>
        </Modal>
      )}
      {statusModal && (
        <Modal title={statusModal.actionLabel || `Move MPO to ${MPO_STATUS_LABELS[statusModal.nextStatus] || statusModal.nextStatus}`} onClose={() => setStatusModal(null)} width={560}>
          <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
            <div style={{ fontSize: 13, color: "var(--text2)" }}>
              {statusModal.mpo?.mpoNo || "MPO"} will move from <strong>{MPO_STATUS_LABELS[statusModal.mpo?.status || "draft"] || (statusModal.mpo?.status || "draft")}</strong> to <strong>{MPO_STATUS_LABELS[statusModal.nextStatus] || statusModal.nextStatus}</strong>.
            </div>
            <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
              <Badge color={getMpoWorkflowMeta(statusModal.mpo).color}>{getMpoWorkflowMeta(statusModal.mpo).label}</Badge>
              <Badge color={getMpoWorkflowMeta({ ...statusModal.mpo, status: statusModal.nextStatus }).color}>Next up: {statusModal.nextOwnerLabel || getMpoWorkflowMeta({ ...statusModal.mpo, status: statusModal.nextStatus }).label}</Badge>
            </div>
            <div style={{ fontSize: 12, color: "var(--text2)", background: "var(--bg3)", borderRadius: 10, padding: "10px 12px", border: "1px solid var(--border)" }}>
              {statusModal.helperText || getMpoWorkflowMeta({ ...statusModal.mpo, status: statusModal.nextStatus }).hint}
            </div>
            <Field
              label={statusModal.noteLabel || (statusModal.nextStatus === "rejected" ? "Rejection note" : "Approval / workflow note")}
              rows={4}
              value={statusModal.note || ""}
              onChange={value => setStatusModal(modal => ({ ...modal, note: value }))}
              placeholder={statusModal.nextStatus === "rejected" ? "What should be corrected before this MPO can move forward?" : "Add context for the next team member..."}
              required
            />
            <div style={{ display: "flex", justifyContent: "flex-end", gap: 10 }}>
              <Btn variant="ghost" onClick={() => setStatusModal(null)}>Cancel</Btn>
              <Btn variant={getWorkflowActionVariant(statusModal.nextStatus)} onClick={() => {
                if (mpoStatusNeedsNote(statusModal.nextStatus) && !String(statusModal.note || "").trim()) {
                  setToast({ msg: "A note is required for this workflow action.", type: "error" });
                  return;
                }
                applyMpoStatusChange(statusModal.mpo, statusModal.nextStatus, statusModal.note || "");
              }}>{statusModal.actionLabel || "Confirm Status Change"}</Btn>
            </div>
          </div>
        </Modal>
      )}
      <style>{`
        @media (max-width: 1520px) {
          .mpo-list-card-grid {
            grid-template-columns: 72px minmax(320px,1fr) !important;
          }
          .mpo-list-card-main,
          .mpo-list-card-side,
          .mpo-list-card-actions {
            grid-column: 2 / -1 !important;
          }
          .mpo-list-card-side {
            justify-self: stretch !important;
          }
          .mpo-list-card-summary {
            justify-content: space-between !important;
          }
        }
        @media (max-width: 980px) {
          .mpo-list-card-grid {
            grid-template-columns: 1fr !important;
          }
          .mpo-list-card-icon,
          .mpo-list-card-main,
          .mpo-list-card-side,
          .mpo-list-card-actions {
            grid-column: auto !important;
          }
          .mpo-list-card-main {
            text-align: left !important;
          }
          .mpo-list-card-main-badges {
            justify-content: flex-start !important;
          }
          .mpo-list-card-summary {
            grid-template-columns: 1fr !important;
            justify-content: stretch !important;
          }
          .mpo-list-card-actions {
            justify-content: flex-start !important;
          }
        }
      `}</style>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 24, flexWrap: "wrap", gap: 10 }}>
        <div><h1 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 800, fontSize: 24 }}>MPO Generator</h1><p style={{ color: "var(--text2)", marginTop: 3, fontSize: 13 }}>Create, manage & export Media Purchase Orders</p></div>
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center" }}><Field value={searchTerm} onChange={setSearchTerm} placeholder="Search MPO, vendor, client..." /><Field value={statusFilter} onChange={setStatusFilter} options={[{value:"all",label:"All Statuses"}, ...MPO_STATUS_OPTIONS.map(o => ({ value: o.value, label: o.label }))]} /><Field value={viewMode} onChange={setViewMode} options={[{value:"active",label:"Active"},{value:"archived",label:"Archived"},{value:"all",label:"All"}]} />{canManage && <Btn icon="+" onClick={openNew}>New MPO</Btn>}</div>
      </div>
      {hasSavedDraft && <Card style={{ marginBottom: 14, padding: "14px 18px", background: "rgba(59,126,245,.08)", border: "1px solid rgba(59,126,245,.22)" }}><div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 12, flexWrap: "wrap" }}><div><div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 14 }}>Saved MPO draft found</div><div style={{ fontSize: 12, color: "var(--text2)", marginTop: 3 }}>Your last in-progress MPO was autosaved locally. You can resume it or clear it.</div></div><div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}><Btn variant="blue" size="sm" onClick={resumeSavedDraft}>Resume Draft</Btn><Btn variant="ghost" size="sm" onClick={clearSavedDraft}>Clear Draft</Btn></div></div></Card>}
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(170px,1fr))", gap: 12, marginBottom: 14 }}>
        <Stat icon="⏳" label="My Queue" value={workflowStats.myQueue} sub="MPOs currently waiting on your role" color="var(--blue)" />
        <Stat icon="🧾" label="Pending Review" value={workflowStats.pendingReview} sub="Submitted and reviewed MPOs in the approval lane" color="var(--purple)" />
        <Stat icon="📤" label="Ready to Send" value={workflowStats.readyToSend} sub="Approved MPOs waiting for dispatch" color="var(--teal)" />
        <Stat icon="🛠" label="Needs Changes" value={workflowStats.needsChanges} sub="Rejected MPOs waiting for revision" color="var(--red)" />
      </div>
      <Card style={{ marginBottom: 14, padding: "16px 18px" }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, flexWrap: "wrap", marginBottom: workflowPanelOpen ? 14 : 0 }}>
          <div>
            <div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 15 }}>Approvals Automation</div>
            <div style={{ fontSize: 12, color: "var(--text2)", marginTop: 4 }}>Role-based queue, waiting-on visibility, and one-click workflow actions.</div>
          </div>
          <Btn variant="ghost" size="sm" onClick={() => setWorkflowPanelOpen(open => !open)}>{workflowPanelOpen ? "Hide" : "Show"} queue</Btn>
        </div>
        {workflowPanelOpen && (
          <div style={{ display: "grid", gridTemplateColumns: "1.3fr .9fr", gap: 14 }}>
            <div style={{ border: "1px solid var(--border)", borderRadius: 12, padding: 14, background: "var(--bg3)" }}>
              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 10 }}>
                <div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 14 }}>My approval / action queue</div>
                <Badge color="blue">{workflowStats.myQueue} item{workflowStats.myQueue !== 1 ? "s" : ""}</Badge>
              </div>
              {topQuickQueue.length === 0 ? (
                <div style={{ fontSize: 12, color: "var(--text2)" }}>Nothing is currently waiting on your role.</div>
              ) : (
                <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                  {topQuickQueue.map(mpo => {
                    const workflowMeta = getMpoWorkflowMeta(mpo);
                    const quickActions = getQuickWorkflowActions(user, mpo).slice(0, 2);
                    return (
                      <div key={mpo.id} style={{ border: "1px solid var(--border)", borderRadius: 10, padding: "12px 13px", background: "var(--bg2)" }}>
                        <div style={{ display: "flex", justifyContent: "space-between", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
                          <div style={{ fontWeight: 700, fontSize: 13 }}>{mpo.mpoNo || "MPO"} · {mpo.vendorName || "Vendor"}</div>
                          <Badge color={workflowMeta.color}>Waiting on you</Badge>
                        </div>
                        <div style={{ fontSize: 12, color: "var(--text2)", marginTop: 5 }}>{workflowMeta.hint}</div>
                        <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginTop: 10 }}>
                          {quickActions.map(action => (
                            <Btn key={action.value} size="sm" variant={action.variant} onClick={() => requestMpoStatusChange(mpo, action.value)}>{action.label}</Btn>
                          ))}
                          <Btn size="sm" variant="ghost" onClick={() => openMpoHistory(mpo)}>History</Btn>
                        </div>
                      </div>
                    );
                  })}
                </div>
              )}
            </div>
            <div style={{ border: "1px solid var(--border)", borderRadius: 12, padding: 14, background: "var(--bg3)" }}>
              <div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 14, marginBottom: 10 }}>Workflow lanes</div>
              <div style={{ display: "grid", gap: 8 }}>
                {MPO_STATUS_OPTIONS.map(option => (
                  <div key={option.value} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "9px 10px", borderRadius: 9, background: "var(--bg2)", border: "1px solid var(--border)" }}>
                    <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                      <Badge color={statusColors[option.value] || "accent"}>{option.label}</Badge>
                      <span style={{ fontSize: 12, color: "var(--text2)" }}>{getMpoWorkflowMeta({ status: option.value }).label}</span>
                    </div>
                    <strong style={{ fontFamily: "'Syne',sans-serif" }}>{workflowLaneCounts[option.value] || 0}</strong>
                  </div>
                ))}
              </div>
            </div>
          </div>
        )}
      </Card>
      {visibleMpos.length === 0 ? <Card><Empty icon="📄" title="No MPOs yet" sub="Create your first Media Purchase Order" /></Card> :
        <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
          {visibleMpoCards.map(m => (
            <Card key={m.id} style={{ padding: "18px 26px" }}>
              <div
                className="mpo-list-card-grid"
                style={{
                  display: "grid",
                  gridTemplateColumns: "72px minmax(460px,1fr) minmax(320px,430px)",
                  columnGap: 18,
                  rowGap: 18,
                  alignItems: "start",
                }}
              >
                <div
                  className="mpo-list-card-icon"
                  style={{
                    width: 64,
                    height: 64,
                    background: "rgba(240,165,0,.08)",
                    border: "1px solid rgba(240,165,0,.22)",
                    borderRadius: 16,
                    display: "flex",
                    alignItems: "center",
                    justifyContent: "center",
                    fontSize: 28,
                    color: "var(--accent)",
                    gridRow: "1 / span 2",
                    alignSelf: "start",
                    justifySelf: "center",
                  }}
                >
                  📄
                </div>

                <div
                  className="mpo-list-card-main"
                  style={{
                    minWidth: 0,
                    textAlign: "center",
                    paddingTop: 2,
                  }}
                >
                  <div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 800, fontSize: 16, lineHeight: 1.2 }}>
                    {m.mpoNo || "MPO"} <span style={{ color: "var(--text3)", fontSize: 14 }}>— {m.vendorName}</span>
                  </div>
                  <div style={{ fontSize: 13, color: "var(--text2)", marginTop: 8, lineHeight: 1.45 }}>
                    {m.clientName} {m.brand && `· ${m.brand}`} · {(m.months || []).length > 1 ? (m.months || []).join("-") : m.month} {m.year} · {m.totalSpots} spots
                  </div>
                  <div
                    className="mpo-list-card-main-badges"
                    style={{ display: "flex", gap: 8, flexWrap: "wrap", marginTop: 14, justifyContent: "center" }}
                  >
                    <Badge color={statusColors[m.status || "draft"] || "accent"}>{MPO_STATUS_LABELS[m.status || "draft"] || (m.status || "draft")}</Badge>
                    <Badge color={getMpoWorkflowMeta(m).color}>Waiting on: {getMpoWorkflowMeta(m).label}</Badge>
                    {isMpoAwaitingUser(user, m) ? <Badge color="blue">My Queue</Badge> : null}
                    <Badge color={getExecutionHealthColor(m)}>{getExecutionHealthLabel(m)}</Badge>
                    <Badge color="blue">Invoice: {(MPO_INVOICE_STATUS_OPTIONS.find(o => o.value === (m.invoiceStatus || "pending"))?.label || m.invoiceStatus || "pending")}</Badge>
                    <Badge color="purple">Proof: {(MPO_PROOF_STATUS_OPTIONS.find(o => o.value === (m.proofStatus || "pending"))?.label || m.proofStatus || "pending")}</Badge>
                    <Badge color="green">Payment: {(MPO_PAYMENT_STATUS_OPTIONS.find(o => o.value === (m.paymentStatus || "unpaid"))?.label || m.paymentStatus || "unpaid")}</Badge>
                    {isArchived(m) ? <Badge color="red">Archived</Badge> : null}
                  </div>
                  <div style={{ fontSize: 12, color: "var(--text3)", marginTop: 14, lineHeight: 1.5 }}>
                    {getMpoWorkflowMeta(m).hint}
                  </div>
                </div>

                <div
                  className="mpo-list-card-side"
                  style={{
                    minWidth: 0,
                    justifySelf: "end",
                    width: "100%",
                    maxWidth: 420,
                  }}
                >
                  <div
                    className="mpo-list-card-summary"
                    style={{
                      display: "grid",
                      gridTemplateColumns: "minmax(90px,1fr) minmax(120px,1fr) minmax(150px,170px)",
                      gap: 18,
                      alignItems: "center",
                      justifyContent: "end",
                    }}
                  >
                    <div style={{ textAlign: "center" }}>
                      <div style={{ fontSize: 10, color: "var(--text3)", fontWeight: 700, textTransform: "uppercase", letterSpacing: ".04em" }}>Gross</div>
                      <div style={{ fontSize: 14, fontWeight: 700, marginTop: 8, color: "var(--text2)" }}>{fmtN(m.totalGross)}</div>
                    </div>
                    <div style={{ textAlign: "center" }}>
                      <div style={{ fontSize: 10, color: "var(--text3)", fontWeight: 700, textTransform: "uppercase", letterSpacing: ".04em" }}>Net Payable</div>
                      <div style={{ fontSize: 14, fontWeight: 800, marginTop: 8, color: "var(--accent)" }}>{fmtN(m.grandTotal || m.netVal)}</div>
                    </div>
                    <div style={{ minWidth: 0 }}>
                      {canManageStatus ? (
                        <Field
                          value={m.status || "draft"}
                          onChange={v => requestMpoStatusChange(m, v)}
                          options={[
                            { value: m.status || "draft", label: MPO_STATUS_LABELS[m.status || "draft"] || (m.status || "draft") },
                            ...getAllowedMpoStatusTargets(user, m).map(value => ({ value, label: MPO_STATUS_LABELS[value] || value })),
                          ].filter((option, index, arr) => arr.findIndex(item => item.value === option.value) === index)}
                        />
                      ) : (
                        <div style={{ paddingTop: 8, textAlign: "center" }}>
                          <Badge color={statusColors[m.status || "draft"] || "accent"}>{(m.status || "draft").toUpperCase()}</Badge>
                        </div>
                      )}
                    </div>
                  </div>
                </div>

                <div
                  className="mpo-list-card-actions"
                  style={{
                    gridColumn: "2 / -1",
                    display: "flex",
                    gap: 10,
                    flexWrap: "wrap",
                    alignItems: "center",
                    justifyContent: "flex-start",
                  }}
                >
                  {getQuickWorkflowActions(user, m).slice(0, 2).map(action => (
                    <Btn key={action.value} variant={action.variant} size="sm" onClick={() => requestMpoStatusChange(m, action.value)}>{action.label}</Btn>
                  ))}
                  <Btn variant="ghost" size="sm" onClick={() => openEdit(m)} icon="✏️">Edit</Btn>
                  <Btn variant="ghost" size="sm" onClick={() => openExecutionModal(m)} icon="📦">Execution</Btn>
                  <Btn variant="ghost" size="sm" onClick={() => openMpoHistory(m)} icon="🕘">History</Btn>
                  <Btn variant="blue" size="sm" onClick={() => openPreview(m)} icon="⬇">Preview & Export</Btn>
                  {isArchived(m) ? <Btn variant="success" size="sm" onClick={() => restoreMPO(m.id)}>↩</Btn> : <Btn variant="danger" size="sm" onClick={() => setConfirm({ msg: `Archive MPO "${m.mpoNo || m.id}"?`, onYes: () => delMPO(m.id) })}>🗄</Btn>}
                </div>
              </div>
            </Card>
          ))}
        </div>}
    </div>
  );

  // FORM VIEW
  const stepLabels = ["Campaign & Vendor", "Spot Schedule", "Costing", "Signatories & Save"];
  return (
    <div className="fade">
      {toast && <Toast msg={toast.msg} type={toast.type} onDone={() => setToast(null)} />}
      {vendorPickerOpen && (
        <Modal title="Search Media House / Vendor" onClose={() => setVendorPickerOpen(false)} width={720}>
          <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
            <Field
              label="Search vendor"
              value={vendorSearch}
              onChange={setVendorSearch}
              placeholder="Type vendor name, type, phone, or email"
            />
            <div style={{ maxHeight: "55vh", overflowY: "auto", border: "1px solid var(--border)", borderRadius: 12, background: "var(--bg3)" }}>
              {filteredVendors.length === 0 ? (
                <div style={{ padding: "18px 16px", fontSize: 13, color: "var(--text2)" }}>
                  No vendor matched “{vendorSearch}”.
                </div>
              ) : (
                filteredVendors.map(item => {
                  const isSelected = item.id === mpoData.vendorId;
                  return (
                    <button
                      key={item.id}
                      type="button"
                      onClick={() => {
                        setMpoData(data => ({ ...data, vendorId: item.id }));
                        setVendorPickerOpen(false);
                      }}
                      style={{
                        width: "100%",
                        textAlign: "left",
                        padding: "13px 15px",
                        border: "none",
                        borderBottom: "1px solid var(--border)",
                        background: isSelected ? "rgba(240,165,0,.12)" : "transparent",
                        cursor: "pointer",
                      }}
                    >
                      <div style={{ display: "flex", justifyContent: "space-between", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
                        <div style={{ fontWeight: 700, fontSize: 13, color: isSelected ? "var(--accent)" : "var(--text)" }}>
                          {isSelected ? "✓ " : ""}{item.name}
                        </div>
                        <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                          {item.type ? <Badge color="blue">{item.type}</Badge> : null}
                          {item.discount ? <Badge color="green">Disc {item.discount}%</Badge> : null}
                          {item.commission ? <Badge color="purple">Comm {item.commission}%</Badge> : null}
                        </div>
                      </div>
                      <div style={{ marginTop: 6, fontSize: 12, color: "var(--text2)", display: "flex", gap: 12, flexWrap: "wrap" }}>
                        {item.contactName ? <span>Contact: {item.contactName}</span> : null}
                        {item.phone ? <span>{item.phone}</span> : null}
                        {item.email ? <span>{item.email}</span> : null}
                      </div>
                    </button>
                  );
                })
              )}
            </div>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
              <div style={{ fontSize: 12, color: "var(--text2)" }}>
                {filteredVendors.length} vendor{filteredVendors.length !== 1 ? "s" : ""} found
              </div>
              <Btn variant="ghost" onClick={() => setVendorPickerOpen(false)}>Close</Btn>
            </div>
          </div>
        </Modal>
      )}
      {spotModal && (
        <Modal title={editSpotId ? "Edit Spot Row" : "Add Spot Row"} onClose={() => { setSpotModal(null); setEditSpotId(null); setSpotForm(blankSpot); }} width={620}>

          {/* Vendor Rate Card Picker */}
          {vendorRates.length > 0 && (
            <div style={{ marginBottom: 18 }}>
              <div style={{ fontSize: 11, fontWeight: 600, color: "var(--text3)", textTransform: "uppercase", letterSpacing: ".06em", marginBottom: 10 }}>
                📺 Select from {vendor?.name || "Vendor"} Rate Cards — click a row to auto-fill
              </div>
              <div style={{ border: "1px solid var(--border2)", borderRadius: 10, overflow: "hidden", maxHeight: 220, overflowY: "auto" }}>
                <table style={{ width: "100%", borderCollapse: "collapse" }}>
                  <thead>
                    <tr style={{ background: "var(--bg3)", position: "sticky", top: 0 }}>
                      {["Programme","Time Belt","Duration","Rate/Spot","Disc %","Net Rate"].map(h => (
                        <th key={h} style={{ padding: "7px 10px", textAlign: "left", fontSize: 10, fontWeight: 600, textTransform: "uppercase", color: "var(--text3)", borderBottom: "1px solid var(--border)", whiteSpace: "nowrap" }}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {vendorRates.map((r, i) => {
                      const selected = spotForm.rateId === r.id;
                      const disc = parseFloat(r.discount) || 0;
                      const net = (parseFloat(r.ratePerSpot) || 0) * (1 - disc / 100);
                      return (
                        <tr key={r.id}
                          onClick={() => {
                            setSpotForm(f => ({
                              ...f,
                              rateId: r.id,
                              programme: r.programme || f.programme,
                              timeBelt: r.timeBelt || f.timeBelt,
                              duration: r.duration || f.duration,
                              customRate: "",
                            }));
                          }}
                          style={{ cursor: "pointer", background: selected ? "rgba(240,165,0,.13)" : i % 2 === 0 ? "transparent" : "rgba(255,255,255,.02)", borderBottom: "1px solid var(--border)", transition: "background .12s" }}
                          onMouseEnter={e => { if (!selected) e.currentTarget.style.background = "var(--bg3)"; }}
                          onMouseLeave={e => { if (!selected) e.currentTarget.style.background = i % 2 === 0 ? "transparent" : "rgba(255,255,255,.02)"; }}>
                          <td style={{ padding: "8px 10px", fontWeight: selected ? 700 : 500, fontSize: 12, color: selected ? "var(--accent)" : "var(--text)" }}>
                            {selected && <span style={{ marginRight: 5 }}>✓</span>}{r.programme || "—"}
                          </td>
                          <td style={{ padding: "8px 10px", fontSize: 12, color: "var(--text2)" }}>{r.timeBelt || "—"}</td>
                          <td style={{ padding: "8px 10px", fontSize: 12, color: "var(--text2)" }}>{r.duration ? `${r.duration}"` : "—"}</td>
                          <td style={{ padding: "8px 10px", fontSize: 12, color: "var(--accent)", fontWeight: 600 }}>{fmtN(r.ratePerSpot)}</td>
                          <td style={{ padding: "8px 10px", fontSize: 12, color: disc > 0 ? "var(--green)" : "var(--text3)" }}>{disc > 0 ? `${disc}%` : "—"}</td>
                          <td style={{ padding: "8px 10px", fontSize: 12, fontWeight: 700, color: "var(--green)" }}>{fmtN(net)}</td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
              {spotForm.rateId && (
                <div style={{ marginTop: 8, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                  <span style={{ fontSize: 11, color: "var(--green)" }}>✓ Rate card applied — fields auto-filled below</span>
                  <button onClick={() => setSpotForm(f => ({ ...f, rateId: "", customRate: "" }))} style={{ fontSize: 11, color: "var(--text3)", background: "none", border: "none", cursor: "pointer", textDecoration: "underline" }}>Clear selection</button>
                </div>
              )}
            </div>
          )}

          {vendorRates.length === 0 && (
            <div style={{ marginBottom: 16, background: "rgba(240,165,0,.07)", border: "1px solid rgba(240,165,0,.2)", borderRadius: 9, padding: "10px 14px", fontSize: 12, color: "var(--text2)" }}>
              ⚠️ No rate cards found for this vendor. Go to <strong style={{ color: "var(--accent)" }}>Media Rates</strong> to add them, or fill in the fields manually below.
            </div>
          )}

          {/* Manual / override fields */}
          <div style={{ borderTop: vendorRates.length > 0 ? "1px solid var(--border)" : "none", paddingTop: vendorRates.length > 0 ? 16 : 0 }}>
            <div style={{ fontSize: 11, fontWeight: 600, color: "var(--text3)", textTransform: "uppercase", letterSpacing: ".06em", marginBottom: 12 }}>
              {vendorRates.length > 0 ? "✏️ Review / Override Details" : "✏️ Spot Details"}
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 13 }}>
              <Field label="Programme" value={spotForm.programme} onChange={updS("programme")} placeholder="NTA News, SuperStory…" required />
              <Field label="Fixed Time / Time Belt" value={spotForm.timeBelt} onChange={updS("timeBelt")} placeholder="08:45–09:00" />
              <Field label="Duration (secs)" type="number" value={spotForm.duration} onChange={updS("duration")} placeholder="30" />
              <Field label="Day of Week" value={spotForm.wd} onChange={updS("wd")} options={["Mon","Tue","Wed","Thu","Fri","Sat","Sun","Daily","Weekdays","Weekends"]} />
              <div style={{ gridColumn: "1/-1" }}>
                {campaign?.materialList?.length > 0 ? (
                  <div style={{ display: "flex", flexDirection: "column", gap: 5 }}>
                    <label style={{ fontSize: 11, fontWeight: 600, color: "var(--text3)", textTransform: "uppercase", letterSpacing: ".08em" }}>Material / Spot Name</label>
                    <select value={spotForm.material} onChange={e => updS("material")(e.target.value)}
                      style={{ background: "var(--bg3)", border: "1px solid var(--border2)", borderRadius: 8, padding: "9px 13px", color: spotForm.material ? "var(--text)" : "var(--text3)", fontSize: 13, outline: "none", cursor: "pointer" }}>
                      <option value="">— Select material —</option>
                      {(campaign.materialList||[]).map((m,i) => <option key={i} value={m}>{m}</option>)}
                      <option value="__custom__">Custom…</option>
                    </select>
                    {spotForm.material === "__custom__" && (
                      <input value={spotForm.materialCustom || ""} onChange={e => updS("materialCustom")(e.target.value)} placeholder="Type material name" style={{ background: "var(--bg3)", border: "1px solid var(--border2)", borderRadius: 8, padding: "9px 13px", color: "var(--text)", fontSize: 13, outline: "none" }} />
                    )}
                  </div>
                ) : (
                  <Field label="Material / Spot Name" value={spotForm.material} onChange={updS("material")} placeholder="SM Thematic English 30secs (MP4)" />
                )}
              </div>
              <div style={{ gridColumn: "1/-1" }}>
                <Field label="Rate per Spot (₦)" type="number"
                  value={spotForm.customRate || (vendorRates.find(r => r.id === spotForm.rateId)?.ratePerSpot || "")}
                  onChange={updS("customRate")}
                  note={spotForm.rateId && !spotForm.customRate ? `From rate card: ${fmtN(vendorRates.find(r => r.id === spotForm.rateId)?.ratePerSpot)}` : "Enter rate manually or select from rate card above"}
                  placeholder="0" />
              </div>
            </div>
          </div>

          {/* Calendar day picker */}
          <div style={{ marginTop: 16, borderTop: "1px solid var(--border)", paddingTop: 16 }}>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 10 }}>
              <div style={{ fontSize: 12, fontWeight: 600, color: "var(--text2)", textTransform: "uppercase", letterSpacing: ".06em" }}>
                📅 Select Airing Dates — <span style={{ color: "var(--accent)" }}>{(spotForm.calendarDays||[]).length} day{(spotForm.calendarDays||[]).length !== 1 ? "s" : ""} selected</span>
              </div>
              <div style={{ display: "flex", gap: 6 }}>
                <button onClick={() => { updS("calendarDays")(Array.from({length:31},(_,i)=>i+1)); updS("spots")("31"); }} style={{ padding: "3px 10px", fontSize: 11, background: "var(--bg3)", border: "1px solid var(--border2)", borderRadius: 6, color: "var(--text2)", cursor: "pointer" }}>All</button>
                <button onClick={() => { updS("calendarDays")([]); updS("spots")(""); }} style={{ padding: "3px 10px", fontSize: 11, background: "var(--bg3)", border: "1px solid var(--border2)", borderRadius: 6, color: "var(--text2)", cursor: "pointer" }}>Clear</button>
              </div>
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(11, 1fr)", gap: 4 }}>
              {Array.from({length: 31}, (_, i) => i + 1).map(d => {
                const sel = (spotForm.calendarDays || []).includes(d);
                return (
                  <button key={d} onClick={() => {
                    const cur = spotForm.calendarDays || [];
                    const next = sel ? cur.filter(x => x !== d) : [...cur, d].sort((a,b)=>a-b);
                    updS("calendarDays")(next);
                    updS("spots")(String(next.length));
                  }}
                  style={{ padding: "6px 2px", border: `1px solid ${sel ? "var(--accent)" : "var(--border2)"}`, borderRadius: 6, background: sel ? "rgba(240,165,0,.18)" : "var(--bg3)", color: sel ? "var(--accent)" : "var(--text3)", fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 12, cursor: "pointer", transition: "all .1s" }}>
                    {d}
                  </button>
                );
              })}
            </div>
          </div>

          <div style={{ marginTop: 14, display: "flex", alignItems: "center", justifyContent: "space-between", background: "var(--bg3)", borderRadius: 9, padding: "10px 14px" }}>
            <span style={{ fontSize: 13, color: "var(--text2)" }}>Number of Spots (auto-counted from dates)</span>
            <span style={{ fontFamily: "'Syne',sans-serif", fontWeight: 800, fontSize: 18, color: "var(--accent)" }}>{(spotForm.calendarDays||[]).length || spotForm.spots || 0}</span>
          </div>

          <div style={{ display: "flex", gap: 10, justifyContent: "flex-end", marginTop: 16 }}>
            <Btn variant="ghost" onClick={() => { setSpotModal(null); setEditSpotId(null); setSpotForm(blankSpot); }}>Cancel</Btn>
            <Btn onClick={addSpot}>{editSpotId ? "Save Changes" : "Add Row"}</Btn>
          </div>
        </Modal>
      )}

      <div style={{ display: "flex", alignItems: "center", gap: 14, marginBottom: 22 }}>
        <Btn variant="ghost" size="sm" onClick={() => { if (window.confirm("Leave this MPO form? Your latest draft has been autosaved and can be restored later.")) setView("list"); }}>← All MPOs</Btn>
        <h1 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 800, fontSize: 22 }}>{editId ? "Edit MPO" : "New MPO"}</h1>
        {editId && <Badge color="accent">Editing</Badge>}
      </div>

      {/* Step bar */}
      <div style={{ display: "flex", gap: 0, marginBottom: 24, background: "var(--bg2)", border: "1px solid var(--border)", borderRadius: 11, overflow: "hidden" }}>
        {stepLabels.map((s, i) => (
          <button key={i} onClick={() => setStep(i + 1)}
            style={{ flex: 1, padding: "11px 6px", border: "none", background: step === i + 1 ? "var(--accent)" : "transparent", color: step === i + 1 ? "#000" : step > i + 1 ? "var(--green)" : "var(--text2)", fontFamily: "'Syne',sans-serif", fontWeight: 600, fontSize: 12, cursor: "pointer", borderRight: i < 3 ? "1px solid var(--border)" : "none", transition: "all .18s" }}>
            <span style={{ display: "block", fontSize: 9, opacity: .75, marginBottom: 1 }}>STEP {i + 1}</span>{s}
          </button>
        ))}
      </div>

      {/* Step 1 */}
      {step === 1 && (
        <div className="fade">
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 18, marginBottom: 18 }}>
            <Card>
              <h3 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, marginBottom: 15, fontSize: 14 }}>Campaign Details</h3>
              <div style={{ display: "flex", flexDirection: "column", gap: 13 }}>
                <Field label="Campaign" value={mpoData.campaignId} onChange={upd("campaignId")} options={campaigns.map(c => ({ value: c.id, label: c.name }))} placeholder="Select campaign" />
                {campaign && <div style={{ background: "var(--bg3)", borderRadius: 8, padding: 11, fontSize: 12 }}><div style={{ color: "var(--text2)" }}>Client: <strong style={{ color: "var(--text)" }}>{client?.name}</strong></div><div style={{ color: "var(--text2)", marginTop: 3 }}>Brand: <strong style={{ color: "var(--text)" }}>{campaign.brand || "—"}</strong></div>{campaign.materialList && campaign.materialList.length > 0 && <div style={{ color: "var(--text2)", marginTop: 3 }}>🎬 <strong style={{ color: "var(--teal)" }}>{campaign.materialList.length} material{campaign.materialList.length!==1?"s":""} available</strong></div>}</div>}
                <div>
                  <div style={{ fontSize: 11, fontWeight: 600, color: "var(--text3)", marginBottom: 5, textTransform: "uppercase", letterSpacing: ".06em" }}>MPO Number (auto-generated)</div>
                  <div style={{ background: "var(--bg3)", border: "1px solid var(--border2)", borderRadius: 9, padding: "10px 13px", fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 15, color: "var(--accent)", letterSpacing: ".04em" }}>{mpoData.mpoNo || "—"}</div>
                  <div style={{ fontSize: 10, color: "var(--text3)", marginTop: 4 }}>Format: Brand prefix + sequence + Month/Year</div>
                </div>
                <Field label="MPO Date" type="date" value={mpoData.date} onChange={upd("date")} />
                <Field label="Status" value={mpoData.status} onChange={upd("status")} options={[{value:"draft",label:"Draft"},{value:"sent",label:"Sent"},{value:"approved",label:"Approved"},{value:"rejected",label:"Rejected"}]} />
              </div>
            </Card>
            <Card>
              <h3 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, marginBottom: 15, fontSize: 14 }}>Vendor & Period</h3>
              <div style={{ display: "flex", flexDirection: "column", gap: 13 }}>
                <div style={{ display: "grid", gridTemplateColumns: "1fr auto", gap: 10, alignItems: "end" }}>
                  <Field label="Media House / Vendor" value={mpoData.vendorId} onChange={upd("vendorId")} options={vendors.map(v => ({ value: v.id, label: v.name }))} placeholder="Select vendor" />
                  <Btn variant="ghost" onClick={() => setVendorPickerOpen(true)} icon="🔎">Search Vendor</Btn>
                </div>
                <div style={{ marginTop: -3, fontSize: 11, color: "var(--text3)" }}>
                  Use the search button to quickly find a media house or vendor by name, type, phone, or email.
                </div>
                {vendor && <div style={{ background: "var(--bg3)", borderRadius: 8, padding: 11, fontSize: 12 }}><Badge color="blue">{vendor.type}</Badge><div style={{ marginTop: 6, color: "var(--text2)" }}>Vol Disc: <strong style={{ color: "var(--accent)" }}>{vendor.discount || 0}%</strong> · Comm: <strong style={{ color: "var(--accent)" }}>{vendor.commission || 0}%</strong></div></div>}
                <div style={{ display: "grid", gridTemplateColumns: "3fr 1fr", gap: 11, alignItems: "start" }}>
                  <div>
                    <div style={{ fontSize: 11, fontWeight: 600, color: "var(--text3)", textTransform: "uppercase", letterSpacing: ".08em", marginBottom: 6 }}>Campaign Month(s)</div>
                    <div style={{ display: "flex", flexWrap: "wrap", gap: 5 }}>
                      {months.map(m => {
                        const key = `${m} ${mpoData.year || new Date().getFullYear()}`;
                        const selected = (mpoData.months || []).includes(m) || mpoData.month === m;
                        return (
                          <button key={m} onClick={() => {
                            const cur = mpoData.months?.length ? [...mpoData.months] : (mpoData.month ? [mpoData.month] : []);
                            const next = cur.includes(m) ? cur.filter(x => x !== m) : [...cur, m];
                            const sorted = months.filter(x => next.includes(x));
                            setMpoData(d => ({ ...d, months: sorted, month: sorted[0] || "" }));
                          }}
                          style={{ padding: "4px 11px", borderRadius: 20, fontSize: 11, fontWeight: 700, cursor: "pointer", fontFamily: "'Syne',sans-serif", border: selected ? "2px solid var(--accent)" : "1px solid var(--border2)", background: selected ? "var(--accent)" : "var(--bg3)", color: selected ? "#000" : "var(--text2)", transition: "all .12s" }}>
                            {m.slice(0,3)}
                          </button>
                        );
                      })}
                    </div>
                    {(mpoData.months || []).length > 0 && (
                      <div style={{ marginTop: 6, fontSize: 10, color: "var(--green)" }}>
                        ✓ {(mpoData.months||[]).join(", ")}
                      </div>
                    )}
                  </div>
                  <Field label="Year" value={mpoData.year} onChange={upd("year")} placeholder="2025" />
                </div>
                <Field label="Medium" value={mpoData.medium} onChange={upd("medium")} options={["Television","Radio","Print","Digital","OOH","Multi-Platform"]} />
                <Field label="Transmit Instruction" value={mpoData.transmitMsg} onChange={upd("transmitMsg")} placeholder={`PLEASE TRANSMIT SPOTS ON ${vendor?.name || "VENDOR"} AS SCHEDULED`} rows={2} />
              </div>
            </Card>
          </div>
          <div style={{ display: "flex", justifyContent: "flex-end" }}><Btn onClick={() => setStep(2)}>Next: Spot Schedule →</Btn></div>
        </div>
      )}

      {/* Step 2 */}
      {step === 2 && (
        <div className="fade">
          <Card style={{ marginBottom: 18 }}>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 15, flexWrap: "wrap", gap: 10 }}>
              <div>
                <h3 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 14 }}>Spot Schedule</h3>
                {vendor && vendorRates.length > 0 && (
                  <div style={{ fontSize: 11, color: "var(--green)", marginTop: 3 }}>
                    ✓ {vendorRates.length} rate card{vendorRates.length !== 1 ? "s" : ""} from <strong>{vendor.name}</strong> — select when adding spots
                  </div>
                )}
                {vendor && vendorRates.length === 0 && (
                  <div style={{ fontSize: 11, color: "var(--text3)", marginTop: 3 }}>
                    ⚠️ No rate cards for <strong>{vendor.name}</strong> — add them in Media Rates
                  </div>
                )}
              </div>
              <div style={{ display: "flex", gap: 8 }}>
                <Btn variant={dailyMode ? "primary" : "ghost"} size="sm" onClick={() => setDailyMode(true)}>📅 Daily Schedule</Btn>
                <Btn variant={!dailyMode ? "secondary" : "ghost"} size="sm" onClick={() => setDailyMode(false)}>➕ Single Row</Btn>
              </div>
            </div>

            {/* ── Daily Calendar Schedule — Multi-Month ── */}
            {dailyMode && (() => {
              const selectedMonths = mpoData.months?.length ? mpoData.months : (mpoData.month ? [mpoData.month] : []);
              if (!selectedMonths.length) return (
                <div style={{ background: "rgba(240,165,0,.1)", border: "1px solid rgba(240,165,0,.3)", borderRadius: 8, padding: "13px 16px", fontSize: 12, color: "var(--accent)", marginBottom: 12 }}>
                  ⚠️ Please select at least one <strong>Campaign Month</strong> in Step 1 before using the Daily Calendar.
                </div>
              );
              const curCalMonth = activeCalMonth || selectedMonths[0];
              return (
                <div>
                  {selectedMonths.length > 1 && (
                    <div style={{ display: "flex", gap: 0, marginBottom: 16, background: "var(--bg3)", border: "1px solid var(--border2)", borderRadius: 10, overflow: "hidden" }}>
                      {selectedMonths.map(m => (
                        <button key={m} onClick={() => setActiveCalMonth(m)}
                          style={{ flex: 1, padding: "9px 8px", border: "none", background: curCalMonth === m ? "var(--accent)" : "transparent", color: curCalMonth === m ? "#000" : "var(--text2)", fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 12, cursor: "pointer", borderRight: "1px solid var(--border)", transition: "all .15s" }}>
                          {m.slice(0,3)} <span style={{ fontSize: 10, opacity: .7 }}>{mpoData.year?.slice(-2)}</span>
                          {(calData[m] || []).some(r => totalCountFromDayCounts(r.dayCounts || {}) > 0) && (
                            <span style={{ marginLeft: 4, background: "rgba(34,197,94,.25)", color: "var(--green)", borderRadius: 10, padding: "1px 5px", fontSize: 9 }}>✓</span>
                          )}
                        </button>
                      ))}
                    </div>
                  )}
                  <DailyCalendar
                    month={curCalMonth} year={mpoData.year}
                    calRows={getStableCalRows(curCalMonth)}
                    setCalRows={fn => setStableCalRows(curCalMonth, fn)}
                    vendorRates={vendorRates} fmtN={fmtN}
                    blankCalRow={blankCalRow}
                    campaignMaterials={campaign?.materialList || []}
                    onAdd={() => {
                      const rows = getStableCalRows(curCalMonth);
                      addFromCalendarSchedule(rows, `${curCalMonth} ${mpoData.year}`);
                      setStableCalRows(curCalMonth, [blankCalRow()]);
                    }}
                  />
                  {selectedMonths.length > 1 && (
                    <div style={{ marginTop: 10, padding: "10px 14px", background: "var(--bg3)", borderRadius: 9, border: "1px solid var(--border)", fontSize: 11, color: "var(--text2)" }}>
                      💡 Switch months using the tabs above to schedule spots across multiple months. Each month is saved independently.
                    </div>
                  )}
                </div>
              );
            })()}

            {/* Single row modal trigger */}
            {!dailyMode && (
              <div style={{ marginBottom: 14, display: "flex", justifyContent: "flex-end" }}>
                {canManage && <Btn size="sm" icon="+" onClick={() => { setSpotForm(blankSpot); setEditSpotId(null); setSpotModal(true); }}>Add Single Spot Row</Btn>}
              </div>
            )}

            {/* Spots table */}
            {spots.length === 0 ? <Empty icon="📋" title="No spots added" sub="Use Daily Schedule or Single Row to add spots" /> :
              <div style={{ overflowX: "auto" }}>
                <table style={{ width: "100%", borderCollapse: "collapse" }}>
                  <thead><tr style={{ background: "var(--bg3)" }}>{["Month","Programme","WD","Time Belt","Material","Dur","Rate/Spot","Spots","Gross",""].map(h => <th key={h} style={{ padding: "7px 9px", textAlign: "left", fontSize: 9, fontWeight: 600, textTransform: "uppercase", color: "var(--text3)", borderBottom: "1px solid var(--border)", whiteSpace: "nowrap" }}>{h}</th>)}</tr></thead>
                  <tbody>
                    {spots.map((r, i) => (
                      <tr key={r.id} style={{ borderBottom: "1px solid var(--border)", background: i % 2 === 0 ? "transparent" : "rgba(255,255,255,.01)" }}>
                        <td style={{ padding: "7px 9px" }}><span style={{ fontSize: 9, background: "rgba(59,126,245,.12)", color: "var(--blue)", padding: "2px 6px", borderRadius: 8, fontWeight: 700, whiteSpace: "nowrap" }}>{r.scheduleMonth || mpoData.month || "—"}</span></td>
                        <td style={{ padding: "7px 9px", fontWeight: 600, fontSize: 12 }}>{r.programme}</td>
                        <td style={{ padding: "7px 9px", fontSize: 12 }}><Badge color="blue">{r.wd}</Badge></td>
                        <td style={{ padding: "7px 9px", color: "var(--text2)", fontSize: 12 }}>{r.timeBelt}</td>
                        <td style={{ padding: "7px 9px", color: "var(--text2)", maxWidth: 120, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", fontSize: 11 }}>{r.material}</td>
                        <td style={{ padding: "7px 9px", fontSize: 12 }}>{r.duration}"</td>
                        <td style={{ padding: "7px 9px", color: "var(--accent)", fontSize: 12 }}>{fmtN(r.ratePerSpot)}</td>
                        <td style={{ padding: "7px 9px", fontWeight: 700, fontSize: 12 }}>{r.spots}</td>
                        <td style={{ padding: "7px 9px", color: "var(--green)", fontWeight: 600, fontSize: 12 }}>{fmtN((parseFloat(r.spots) || 0) * (parseFloat(r.ratePerSpot) || 0))}</td>
                        <td style={{ padding: "7px 9px" }}>
                          <div style={{ display: "flex", gap: 5 }}>
                            {canManage && <Btn variant="ghost" size="sm" onClick={() => {
                              setSpotForm({
                                programme: r.programme,
                                wd: r.wd,
                                timeBelt: r.timeBelt,
                                material: r.material,
                                materialCustom: "",
                                duration: r.duration,
                                rateId: r.rateId || "",
                                customRate: r.ratePerSpot || "",
                                spots: r.spots,
                                calendarDays: r.calendarDays || [],
                                calendarDayCounts: collapseDaysToCounts(r.calendarDays || []),
                              });
                              setEditSpotId(r.id);
                              setSpotModal(true);
                            }}>✏️</Btn>}
                            <Btn variant="danger" size="sm" onClick={() => setSpots(s => s.filter(x => x.id !== r.id))}>×</Btn>
                          </div>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                  <tfoot><tr style={{ background: "var(--bg3)", borderTop: "2px solid var(--border2)" }}><td colSpan={7} style={{ padding: "7px 9px", fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 12 }}>TOTAL</td><td style={{ padding: "7px 9px", fontWeight: 700, color: "var(--blue)", fontSize: 12 }}>{totalSpots}</td><td style={{ padding: "7px 9px", fontWeight: 700, color: "var(--green)", fontSize: 12 }}>{fmtN(totalGross)}</td><td></td></tr></tfoot>
                </table>
              </div>}
          </Card>
          <div style={{ display: "flex", justifyContent: "space-between" }}>
            <Btn variant="ghost" onClick={() => setStep(1)}>← Back</Btn>
            <Btn onClick={() => setStep(3)}>Next: Costing →</Btn>
          </div>
        </div>
      )}

      {/* Step 3 */}
      {step === 3 && (
        <div className="fade">
          <div style={{ display: "grid", gridTemplateColumns: "1fr 340px", gap: 18, marginBottom: 18 }}>
            <Card>
              <h3 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, marginBottom: 15, fontSize: 14 }}>Costing Summary</h3>
              <div style={{ background: "rgba(240,165,0,.07)", border: "1px solid rgba(240,165,0,.18)", borderRadius: 8, padding: "8px 13px", fontSize: 12, color: "var(--text2)", marginBottom: 16 }}>
                💡 VAT is applied automatically using your current document settings.
              </div>
              {/* Surcharge input */}
              <div style={{ marginBottom: 14, background: "var(--bg3)", border: "1px solid var(--border2)", borderRadius: 10, padding: "12px 15px" }}>
                <div style={{ fontSize: 11, fontWeight: 700, color: "var(--text3)", textTransform: "uppercase", letterSpacing: ".06em", marginBottom: 10 }}>Surcharge (Optional)</div>
                <div style={{ display: "grid", gridTemplateColumns: "1fr 2fr", gap: 10 }}>
                  <div>
                    <label style={{ fontSize: 11, color: "var(--text3)", fontWeight: 600, display: "block", marginBottom: 4 }}>Surcharge %</label>
                    <input type="number" value={surcharge.pct} onChange={e => setSurcharge(s => ({ ...s, pct: e.target.value }))}
                      placeholder="e.g. 10"
                      style={{ width: "100%", background: "var(--bg2)", border: "1px solid var(--border2)", borderRadius: 8, padding: "8px 12px", color: "var(--text)", fontSize: 13, outline: "none" }} />
                  </div>
                  <div>
                    <label style={{ fontSize: 11, color: "var(--text3)", fontWeight: 600, display: "block", marginBottom: 4 }}>Surcharge Label</label>
                    <input value={surcharge.label} onChange={e => setSurcharge(s => ({ ...s, label: e.target.value }))}
                      placeholder="e.g. Production Surcharge, Agency Fee"
                      style={{ width: "100%", background: "var(--bg2)", border: "1px solid var(--border2)", borderRadius: 8, padding: "8px 12px", color: "var(--text)", fontSize: 13, outline: "none" }} />
                  </div>
                </div>
                {surchPct > 0 && <div style={{ marginTop: 8, fontSize: 11, color: "var(--orange)" }}>⚡ Surcharge adds {fmtN(surchAmt)} to the net cost</div>}
              </div>

              <div style={{ marginTop: 0, border: "1px solid var(--border)", borderRadius: 11, overflow: "hidden" }}>
                {[
                  ["Total Paid Spots", totalSpots, "var(--text)", false],
                  ["Total Gross Value", fmtN(totalGross), "var(--accent)", true],
                  ...(discPct > 0 ? [[`Volume Discount (${(discPct * 100).toFixed(0)}%)`, `(${fmtN(discAmt)})`, "var(--red)", false], ["Less Discount", fmtN(lessDisc), "var(--text)", false]] : []),
                  ...(commPct > 0 ? [[`Agency Commission (${(commPct * 100).toFixed(0)}%)`, `(${fmtN(commAmt)})`, "var(--red)", false], ["After Commission", fmtN(afterComm), "var(--text)", false]] : []),
                  ...(surchPct > 0 ? [[surcharge.label || `Surcharge (${(surchPct * 100).toFixed(0)}%)`, `+${fmtN(surchAmt)}`, "var(--orange)", false]] : []),
                  ["Net Value", fmtN(netVal), "var(--green)", true],
                  [`VAT (${VAT_RATE}%)`, fmtN(vatAmt), "var(--text)", false]
                ].map(([l, v, c, bold], i) => (
                  <div key={l} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "11px 15px", background: i % 2 === 0 ? "var(--bg3)" : "transparent", borderBottom: "1px solid var(--border)" }}>
                    <span style={{ fontSize: 13, color: "var(--text2)" }}>{l}</span>
                    <span style={{ fontFamily: "'Syne',sans-serif", fontWeight: bold ? 800 : 600, fontSize: bold ? 17 : 13, color: c }}>{v}</span>
                  </div>
                ))}
                <div style={{ padding: "14px 15px", background: "var(--accent)", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                  <span style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, color: "#000", fontSize: 13 }}>TOTAL AMOUNT PAYABLE</span>
                  <span style={{ fontFamily: "'Syne',sans-serif", fontWeight: 800, fontSize: 20, color: "#000" }}>{fmtN(grandTotal)}</span>
                </div>
              </div>
            </Card>
            <Card>
              <h3 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, marginBottom: 15, fontSize: 14 }}>MPO Summary</h3>
              {[["Client", client?.name || "—"], ["Brand", campaign?.brand || "—"], ["Campaign", campaign?.name || "—"], ["Vendor", vendor?.name || "—"], ["MPO No.", mpoData.mpoNo || "—"], ["Period", `${(mpoData.months||[]).length > 1 ? (mpoData.months||[]).join(", ") : (mpoData.month || "—")} ${mpoData.year}`], ["Total Spots", totalSpots], ["Gross", fmtN(totalGross)], ["Net Payable", fmtN(grandTotal)]].map(([l, v]) => (
                <div key={l} style={{ display: "flex", justifyContent: "space-between", padding: "7px 0", borderBottom: "1px solid var(--border)", fontSize: 12 }}>
                  <span style={{ color: "var(--text3)", fontWeight: 600 }}>{l}</span>
                  <span style={{ fontWeight: 600 }}>{v}</span>
                </div>
              ))}
            </Card>
          </div>
          <div style={{ display: "flex", justifyContent: "space-between" }}>
            <Btn variant="ghost" onClick={() => setStep(2)}>← Back</Btn>
            <Btn onClick={() => setStep(4)}>Next: Signatories →</Btn>
          </div>
        </div>
      )}

      {/* Step 4 */}
      {step === 4 && (
        <div className="fade">
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 18, marginBottom: 18 }}>
            <Card>
              <h3 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, marginBottom: 15, fontSize: 14 }}>Signatories</h3>
              <div style={{ display: "flex", flexDirection: "column", gap: 13 }}>
                <Field label="Signed By (Name)" value={mpoData.signedBy} onChange={upd("signedBy")} placeholder="Nkechi Eluma" />
                <Field label="Signed By (Title)" value={mpoData.signedTitle} onChange={upd("signedTitle")} placeholder="Head Buying and Compliance" />
                {user?.signatureDataUrl && (
                  <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
                    <Btn variant="ghost" size="sm" onClick={() => setMpoData(m => ({ ...m, signedSignature: user.signatureDataUrl }))}>Use My Uploaded Signature</Btn>
                    {mpoData.signedSignature ? <Btn variant="danger" size="sm" onClick={() => setMpoData(m => ({ ...m, signedSignature: "" }))}>Clear Signature</Btn> : null}
                  </div>
                )}
                <div style={{ background: "var(--bg3)", border: "1px solid var(--border2)", borderRadius: 10, padding: "13px 15px" }}>
                  <div style={{ fontSize: 10, fontWeight: 600, color: "var(--text3)", textTransform: "uppercase", letterSpacing: ".06em", marginBottom: 10 }}>Prepared By (from your account)</div>
                  <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                    <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13 }}><span style={{ color: "var(--text3)" }}>Name</span><span style={{ fontWeight: 600 }}>{user?.name || "—"}</span></div>
                    <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13 }}><span style={{ color: "var(--text3)" }}>Title</span><span style={{ fontWeight: 600 }}>{user?.title || "—"}</span></div>
                    <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13 }}><span style={{ color: "var(--text3)" }}>Contact</span><span style={{ fontWeight: 600 }}>{user?.email || "—"}</span></div>
                  </div>
                  {user?.signatureDataUrl && (
                    <div style={{ marginTop: 10 }}>
                      <div style={{ fontSize: 11, color: "var(--text3)", marginBottom: 6 }}>Prepared signature</div>
                      <img src={user.signatureDataUrl} alt="Prepared signature" style={{ maxHeight: 52, maxWidth: 220, objectFit: "contain", background: "#fff", borderRadius: 8, border: "1px solid var(--border)" }} />
                    </div>
                  )}
                  <div style={{ marginTop: 9, fontSize: 11, color: "var(--text3)" }}>To update, edit your profile or re-register with updated details.</div>
                </div>
                <Field label="Agency Address" value={mpoData.agencyAddress} onChange={upd("agencyAddress")} placeholder="5, Craig Street, Ogudu GRA, Lagos" />
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 }}>
                  <Field label="Agency Email" type="email" value={mpoData.agencyEmail || ""} onChange={upd("agencyEmail")} placeholder="hello@agency.com" />
                  <Field label="Agency Phone" type="tel" value={mpoData.agencyPhone || ""} onChange={upd("agencyPhone")} placeholder="+234 800 000 0000" />
                </div>
              </div>
            </Card>
            <Card style={{ background: "linear-gradient(135deg,rgba(240,165,0,.07),rgba(59,126,245,.04))", border: "1px solid rgba(240,165,0,.18)" }}>
              <h3 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, marginBottom: 15, fontSize: 14 }}>Ready to Save</h3>
              <div style={{ display: "flex", flexDirection: "column", gap: 8, fontSize: 13 }}>
                <div>📋 <strong>{campaign?.name || "No campaign selected"}</strong></div>
                <div>🏢 <strong>{vendor?.name || "No vendor selected"}</strong></div>
                <div>👥 {client?.name || "—"}{campaign?.brand && ` · ${campaign.brand}`}</div>
                <div>📅 {mpoData.month} {mpoData.year} · MPO: {mpoData.mpoNo || "—"}</div>
                <div>📋 {totalSpots} spots · {spots.length} rows</div>
                <div style={{ marginTop: 10, padding: "13px 15px", background: "rgba(240,165,0,.1)", borderRadius: 10, border: "1px solid rgba(240,165,0,.2)" }}>
                  <div style={{ fontSize: 11, color: "var(--text3)", marginBottom: 3 }}>TOTAL AMOUNT PAYABLE</div>
                  <div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 800, fontSize: 22, color: "var(--accent)" }}>{fmtN(grandTotal)}</div>
                </div>
                <div style={{ marginTop: 6, fontSize: 11, color: "var(--text2)", lineHeight: 1.5 }}>After saving, you can download this MPO as PDF, Word document, or Excel/CSV from the MPO list.</div>
              </div>
            </Card>
          </div>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
            <Btn variant="ghost" onClick={() => setStep(3)}>← Back</Btn>
            <div style={{ display: "flex", gap: 10 }}>
              <Btn variant="secondary" onClick={() => { if (window.confirm("Leave this MPO form? Your latest draft has been autosaved and can be restored later.")) setView("list"); }}>Cancel</Btn>
              {canManage && <Btn variant="success" icon="✓" onClick={saveMPO}>{editId ? "Save Changes" : "Save MPO"}</Btn>}
            </div>
          </div>
        </div>
      )}
    </div>
  );

}
