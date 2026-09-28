import { memo, useMemo, useRef, useState } from "react";
import Badge from "../components/Badge";
import Empty from "../components/Empty";
import Toast from "../components/Toast";
import { Btn, Card, Field, Stat } from "../components/ui/primitives";
import { buildCSV, loadBrowserScript } from "../utils/export";
import { activeOnly } from "../utils/records";
import { fmtN } from "../utils/formatters";
import { hasPermission, readOnlyMessage } from "../constants/roles";
import {
  MPO_INVOICE_STATUS_OPTIONS,
  MPO_PAYMENT_STATUS_OPTIONS,
  MPO_RECON_STATUS_OPTIONS,
} from "../constants/mpoWorkflow";
import { updateMpoExecutionInSupabase } from "../services/mpos";
import { createAuditEventInSupabase, notifyExecutionUpdate } from "../services/notifications";

const MONTHS = [
  "january",
  "february",
  "march",
  "april",
  "may",
  "june",
  "july",
  "august",
  "september",
  "october",
  "november",
  "december",
];

const HEADER_ALIASES = {
  mpoNo: ["mpo no", "mpo number", "mpo", "media purchase order", "order no", "release order", "ro no", "po no"],
  vendor: ["vendor", "station", "media owner", "media house", "channel", "publisher", "supplier"],
  client: ["client", "advertiser", "customer"],
  brand: ["brand", "product"],
  campaign: ["campaign", "campaign name", "campaign title", "flight", "activity"],
  programme: ["programme", "program", "show", "placement", "program title", "programme title"],
  material: ["material", "copy", "spot name", "creative", "commercial", "asset", "ad title"],
  date: ["date", "aired date", "air date", "transmission date", "broadcast date", "log date"],
  day: ["day", "air day", "broadcast day"],
  month: ["month", "air month", "broadcast month"],
  year: ["year"],
  time: ["aired time", "air time", "time aired", "actual time", "broadcast time", "tx time", "transmission time", "time"],
  timeBand: ["time band", "timeband", "time belt", "timebelt", "booked time", "programme time", "slot"],
  duration: ["duration", "dur", "length", "sec", "seconds"],
  spots: ["spots", "spot count", "aired spots", "count", "qty", "quantity", "no of spots", "number of spots"],
  invoiceStatus: ["invoice status", "invoice", "billing status"],
  paymentStatus: ["payment status", "payment", "paid status"],
  invoiceAmount: ["invoice amount", "billed amount", "amount invoiced"],
  amountPaid: ["amount paid", "paid amount", "received amount", "payment amount"],
  grossAmount: ["gross amount", "gross", "total gross"],
  netAmount: ["net amount", "net", "payable", "total payable"],
  discount: ["discount", "discount applied", "disc", "discount amount"],
};

const invoiceLabelMap = Object.fromEntries(MPO_INVOICE_STATUS_OPTIONS.map(option => [option.value, option.label]));
const paymentLabelMap = Object.fromEntries(MPO_PAYMENT_STATUS_OPTIONS.map(option => [option.value, option.label]));
const reconLabelMap = Object.fromEntries(MPO_RECON_STATUS_OPTIONS.map(option => [option.value, option.label]));

const uid = () => Math.random().toString(36).slice(2) + Date.now().toString(36);

const normalizeWhitespace = (value = "") => String(value ?? "").replace(/\s+/g, " ").trim();

const normalizeText = (value = "") =>
  normalizeWhitespace(value)
    .toLowerCase()
    .replace(/&/g, " and ")
    .replace(/[^a-z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();

const compactText = (value = "") => normalizeText(value).replace(/\s+/g, "");

const normalizeHeader = (value = "") => normalizeText(value).replace(/\b(no|number)\b/g, "no").trim();

const parseNumber = (value) => {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  const cleaned = String(value ?? "").replace(/[^\d.-]/g, "");
  if (!cleaned || cleaned === "-" || cleaned === ".") return 0;
  const parsed = Number(cleaned);
  return Number.isFinite(parsed) ? parsed : 0;
};

const normalizeCell = (value) => {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value.toISOString().slice(0, 10);
  }
  return normalizeWhitespace(value);
};

const tokensFor = (value = "") => normalizeText(value).split(" ").filter(token => token.length >= 2);

const textSimilarity = (left, right) => {
  const a = normalizeText(left);
  const b = normalizeText(right);
  if (!a || !b) return 0;
  if (a === b) return 1;
  if (a.includes(b) || b.includes(a)) return 0.85;
  const at = new Set(tokensFor(a));
  const bt = new Set(tokensFor(b));
  if (!at.size || !bt.size) return 0;
  let overlap = 0;
  at.forEach(token => {
    if (bt.has(token)) overlap += 1;
  });
  return overlap / Math.max(at.size, bt.size);
};

const getScheduledSpotCount = (spot = {}) => {
  if (Array.isArray(spot.calendarDays) && spot.calendarDays.length) return spot.calendarDays.length;
  if (Array.isArray(spot.ad) && spot.ad.length) return spot.ad.length;
  return parseNumber(spot.spots);
};

const getPaidSpotCount = (spot = {}) => {
  const scheduled = getScheduledSpotCount(spot);
  if (spot.isComplimentary || Number(spot.ratePerSpot || 0) <= 0) return 0;
  const bonus = Math.max(0, Math.min(Number(spot.bonusSpots) || 0, scheduled));
  return Math.max(0, scheduled - bonus);
};

const monthIndexFromText = (value = "") => {
  const raw = normalizeText(value);
  if (!raw) return -1;
  const numeric = Number(raw);
  if (Number.isInteger(numeric) && numeric >= 1 && numeric <= 12) return numeric - 1;
  return MONTHS.findIndex(month => month === raw || month.slice(0, 3) === raw.slice(0, 3));
};

const parseReportDate = (value, fallbackYear = "") => {
  const raw = normalizeWhitespace(value);
  if (!raw) return null;

  if (/^\d+(\.\d+)?$/.test(raw)) {
    const serial = Number(raw);
    if (serial > 20000 && serial < 80000) {
      const date = new Date(Math.round((serial - 25569) * 86400 * 1000));
      if (!Number.isNaN(date.getTime())) {
        return {
          date,
          day: date.getDate(),
          monthIndex: date.getMonth(),
          year: date.getFullYear(),
          hasYear: true,
          iso: date.toISOString().slice(0, 10),
        };
      }
    }
  }

  const ymd = raw.match(/\b(20\d{2}|19\d{2})[-/.](\d{1,2})[-/.](\d{1,2})\b/);
  if (ymd) {
    const date = new Date(Number(ymd[1]), Number(ymd[2]) - 1, Number(ymd[3]));
    if (!Number.isNaN(date.getTime())) {
      return { date, day: date.getDate(), monthIndex: date.getMonth(), year: date.getFullYear(), hasYear: true, iso: date.toISOString().slice(0, 10) };
    }
  }

  const dmy = raw.match(/\b(\d{1,2})[-/.](\d{1,2})(?:[-/.](\d{2,4}))?\b/);
  if (dmy) {
    let first = Number(dmy[1]);
    let second = Number(dmy[2]);
    let day = first;
    let month = second;
    if (second > 12 && first <= 12) {
      day = second;
      month = first;
    }
    let year = dmy[3] ? Number(dmy[3]) : Number(fallbackYear) || new Date().getFullYear();
    if (year < 100) year += 2000;
    const date = new Date(year, month - 1, day);
    if (!Number.isNaN(date.getTime())) {
      return { date, day, monthIndex: month - 1, year, hasYear: Boolean(dmy[3]), iso: date.toISOString().slice(0, 10) };
    }
  }

  const named = raw.match(/\b(\d{1,2})\s+([A-Za-z]+)\s*(\d{2,4})?\b/) || raw.match(/\b([A-Za-z]+)\s+(\d{1,2})\s*(\d{2,4})?\b/);
  if (named) {
    const firstIsMonth = Number.isNaN(Number(named[1]));
    const day = Number(firstIsMonth ? named[2] : named[1]);
    const monthIndex = monthIndexFromText(firstIsMonth ? named[1] : named[2]);
    let year = named[3] ? Number(named[3]) : Number(fallbackYear) || new Date().getFullYear();
    if (year < 100) year += 2000;
    if (day && monthIndex >= 0) {
      const date = new Date(year, monthIndex, day);
      if (!Number.isNaN(date.getTime())) {
        return { date, day, monthIndex, year, hasYear: Boolean(named[3]), iso: date.toISOString().slice(0, 10) };
      }
    }
  }

  const parsed = new Date(raw);
  if (!Number.isNaN(parsed.getTime())) {
    return {
      date: parsed,
      day: parsed.getDate(),
      monthIndex: parsed.getMonth(),
      year: parsed.getFullYear(),
      hasYear: /\b(20\d{2}|19\d{2})\b/.test(raw),
      iso: parsed.toISOString().slice(0, 10),
    };
  }

  return null;
};

const parseTimeToMinutes = (value = "") => {
  const raw = normalizeWhitespace(value).toLowerCase();
  if (!raw) return null;

  const compact = raw.match(/\b([01]?\d|2[0-3])([0-5]\d)\b/);
  if (compact) return Number(compact[1]) * 60 + Number(compact[2]);

  const clock = raw.match(/\b(\d{1,2})(?::([0-5]\d))?\s*(am|pm)?\b/);
  if (!clock) return null;

  let hour = Number(clock[1]);
  const minute = Number(clock[2] || 0);
  const meridian = clock[3];
  if (meridian === "pm" && hour < 12) hour += 12;
  if (meridian === "am" && hour === 12) hour = 0;
  if (hour > 23 || minute > 59) return null;
  return hour * 60 + minute;
};

const parseTimeRange = (value = "") => {
  const raw = normalizeWhitespace(value).toLowerCase();
  if (!raw) return null;
  const matches = Array.from(raw.matchAll(/\b(\d{1,2})(?::([0-5]\d))?\s*(am|pm)?\b/g))
    .map(match => parseTimeToMinutes(match[0]))
    .filter(value => value !== null);
  if (matches.length >= 2) return { start: matches[0], end: matches[1] };
  const compactMatches = Array.from(raw.matchAll(/\b([01]?\d|2[0-3])([0-5]\d)\b/g))
    .map(match => Number(match[1]) * 60 + Number(match[2]));
  if (compactMatches.length >= 2) return { start: compactMatches[0], end: compactMatches[1] };
  return null;
};

const minutesInRange = (minutes, range) => {
  if (minutes === null || !range) return false;
  if (range.start <= range.end) return minutes >= range.start && minutes <= range.end;
  return minutes >= range.start || minutes <= range.end;
};

const formatTimeBandCheck = (row, spot) => {
  const rowTime = row.time || row.timeBand;
  const spotBand = spot?.timeBelt || "";
  const rowText = normalizeText(rowTime);
  const spotText = normalizeText(spotBand);
  const actual = parseTimeToMinutes(rowTime);
  const bookedRange = parseTimeRange(spotBand);

  if (actual !== null && bookedRange) {
    return minutesInRange(actual, bookedRange)
      ? { status: "in", label: "In band" }
      : { status: "out", label: "Out of band" };
  }

  if (rowText && spotText && (rowText.includes(spotText) || spotText.includes(rowText))) {
    return { status: "in", label: "Band text matched" };
  }

  return { status: "unknown", label: actual === null ? "No aired time" : "No booked range" };
};

const detectDelimiter = (text = "") => {
  const firstLine = String(text).split(/\r?\n/).find(line => line.trim()) || "";
  const candidates = [",", "\t", ";", "|"];
  return candidates
    .map(delimiter => ({ delimiter, count: firstLine.split(delimiter).length }))
    .sort((a, b) => b.count - a.count)[0]?.delimiter || ",";
};

const parseDelimitedTextToMatrix = (text = "") => {
  const delimiter = detectDelimiter(text);
  const rows = [];
  let row = [];
  let cell = "";
  let quoted = false;

  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    const next = text[index + 1];

    if (char === '"') {
      if (quoted && next === '"') {
        cell += '"';
        index += 1;
      } else {
        quoted = !quoted;
      }
      continue;
    }

    if (!quoted && char === delimiter) {
      row.push(cell);
      cell = "";
      continue;
    }

    if (!quoted && (char === "\n" || char === "\r")) {
      if (char === "\r" && next === "\n") index += 1;
      row.push(cell);
      rows.push(row);
      row = [];
      cell = "";
      continue;
    }

    cell += char;
  }

  row.push(cell);
  rows.push(row);
  return rows.filter(item => item.some(cellValue => normalizeWhitespace(cellValue)));
};

const headerKnownCount = (row = []) =>
  row.reduce((count, header) => {
    const normalized = normalizeHeader(header);
    const matched = Object.values(HEADER_ALIASES).some(aliases =>
      aliases.some(alias => normalized === normalizeHeader(alias))
    );
    return count + (matched ? 1 : 0);
  }, 0);

const getRawValue = (raw = {}, aliases = []) => {
  const entries = Object.entries(raw).map(([key, value]) => [normalizeHeader(key), value]);
  for (const alias of aliases) {
    const normalizedAlias = normalizeHeader(alias);
    const exact = entries.find(([key, value]) => key === normalizedAlias && normalizeWhitespace(value));
    if (exact) return exact[1];
  }
  for (const alias of aliases) {
    const normalizedAlias = normalizeHeader(alias);
    const partial = entries.find(([key, value]) => {
      if (!normalizeWhitespace(value)) return false;
      if (normalizedAlias.length <= 3) return key === normalizedAlias;
      return key.includes(normalizedAlias) || normalizedAlias.includes(key);
    });
    if (partial) return partial[1];
  }
  return "";
};

const normalizeReportRow = (raw = {}, index = 0, source = "") => {
  const yearValue = getRawValue(raw, HEADER_ALIASES.year);
  const dateValue = getRawValue(raw, HEADER_ALIASES.date);
  const dayValue = getRawValue(raw, HEADER_ALIASES.day);
  const monthValue = getRawValue(raw, HEADER_ALIASES.month);
  const inferredDate = dateValue || ([dayValue, monthValue, yearValue].filter(Boolean).join(" "));
  const airedDate = parseReportDate(inferredDate, yearValue);
  const airCount = Math.max(1, Math.round(parseNumber(getRawValue(raw, HEADER_ALIASES.spots)) || 1));

  const row = {
    id: uid(),
    rowNumber: index + 1,
    source,
    raw,
    mpoNo: normalizeWhitespace(getRawValue(raw, HEADER_ALIASES.mpoNo)),
    vendor: normalizeWhitespace(getRawValue(raw, HEADER_ALIASES.vendor)),
    client: normalizeWhitespace(getRawValue(raw, HEADER_ALIASES.client)),
    brand: normalizeWhitespace(getRawValue(raw, HEADER_ALIASES.brand)),
    campaign: normalizeWhitespace(getRawValue(raw, HEADER_ALIASES.campaign)),
    programme: normalizeWhitespace(getRawValue(raw, HEADER_ALIASES.programme)),
    material: normalizeWhitespace(getRawValue(raw, HEADER_ALIASES.material)),
    dateRaw: normalizeWhitespace(inferredDate),
    airedDate,
    time: normalizeWhitespace(getRawValue(raw, HEADER_ALIASES.time)),
    timeBand: normalizeWhitespace(getRawValue(raw, HEADER_ALIASES.timeBand)),
    duration: normalizeWhitespace(getRawValue(raw, HEADER_ALIASES.duration)),
    airCount,
    invoiceStatus: normalizeWhitespace(getRawValue(raw, HEADER_ALIASES.invoiceStatus)),
    paymentStatus: normalizeWhitespace(getRawValue(raw, HEADER_ALIASES.paymentStatus)),
    invoiceAmount: parseNumber(getRawValue(raw, HEADER_ALIASES.invoiceAmount)),
    amountPaid: parseNumber(getRawValue(raw, HEADER_ALIASES.amountPaid)),
    grossAmount: parseNumber(getRawValue(raw, HEADER_ALIASES.grossAmount)),
    netAmount: parseNumber(getRawValue(raw, HEADER_ALIASES.netAmount)),
    discount: parseNumber(getRawValue(raw, HEADER_ALIASES.discount)),
  };

  row.hasMeaning = [
    row.mpoNo,
    row.vendor,
    row.client,
    row.brand,
    row.campaign,
    row.programme,
    row.material,
    row.dateRaw,
    row.time,
    row.timeBand,
  ].some(Boolean);

  return row;
};

const parseMatrixRows = (matrix = [], source = "") => {
  const normalizedMatrix = matrix
    .map(row => (row || []).map(normalizeCell))
    .filter(row => row.some(Boolean));

  if (!normalizedMatrix.length) return [];

  const detectedHeaderIndex = normalizedMatrix.findIndex(row => headerKnownCount(row) >= 2);
  const headerIndex = detectedHeaderIndex >= 0 ? detectedHeaderIndex : 0;
  const headers = normalizedMatrix[headerIndex].map((header, index) => normalizeWhitespace(header) || `Column ${index + 1}`);

  return normalizedMatrix.slice(headerIndex + 1)
    .map((cells, index) => {
      const raw = {};
      headers.forEach((header, cellIndex) => {
        raw[header] = normalizeCell(cells[cellIndex] ?? "");
      });
      return normalizeReportRow(raw, index, source);
    })
    .filter(row => row.hasMeaning);
};

const loadSheetJS = () =>
  loadBrowserScript(
    "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js",
    () => window.XLSX
  );

const parseMonitoringFile = async (file) => {
  const fileName = file?.name || "monitoring-report";
  if (!file) return [];

  if (/\.(xlsx|xls)$/i.test(fileName)) {
    const XLSX = await loadSheetJS();
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
    return workbook.SheetNames.flatMap(sheetName => {
      const sheet = workbook.Sheets[sheetName];
      const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: "" });
      return parseMatrixRows(matrix, sheetName);
    });
  }

  const text = await file.text();
  return parseMatrixRows(parseDelimitedTextToMatrix(text), fileName);
};

const mpoPeriodMonths = (mpo = {}) => {
  const labels = [
    ...(Array.isArray(mpo.months) ? mpo.months : []),
    mpo.month,
    ...(mpo.spots || []).map(spot => spot.scheduleMonth),
  ].filter(Boolean);
  return new Set(labels.map(monthIndexFromText).filter(index => index >= 0));
};

const dateMatchesMpoPeriod = (row, mpo) => {
  if (!row?.airedDate) return { known: false, matches: true };
  const months = mpoPeriodMonths(mpo);
  const year = Number(mpo.year) || 0;
  const monthOk = !months.size || months.has(row.airedDate.monthIndex);
  const yearOk = !row.airedDate.hasYear || !year || row.airedDate.year === year;
  return { known: true, matches: monthOk && yearOk };
};

const dateMatchesSpot = (row, spot = {}, mpo = {}) => {
  if (!row?.airedDate) return { known: false, matches: true };
  const spotMonthIndex = monthIndexFromText(spot.scheduleMonth || mpo.month);
  const monthOk = spotMonthIndex < 0 || spotMonthIndex === row.airedDate.monthIndex;
  const year = Number(mpo.year) || 0;
  const yearOk = !row.airedDate.hasYear || !year || row.airedDate.year === year;
  const days = Array.isArray(spot.calendarDays) ? spot.calendarDays.map(Number) : [];
  const dayOk = !days.length || days.includes(Number(row.airedDate.day));
  return { known: true, matches: monthOk && yearOk && dayOk };
};

const scoreMpoForRow = (row, mpo, lookup) => {
  let score = 0;
  const reasons = [];

  if (row.mpoNo) {
    const rowNo = compactText(row.mpoNo);
    const mpoNo = compactText(mpo.mpoNo);
    if (rowNo && mpoNo && rowNo === mpoNo) {
      score += 95;
      reasons.push("MPO no exact");
    } else if (rowNo && mpoNo && (rowNo.includes(mpoNo) || mpoNo.includes(rowNo))) {
      score += 75;
      reasons.push("MPO no partial");
    } else {
      score -= 18;
    }
  }

  const vendorName = mpo.vendorName || lookup.vendors.get(mpo.vendorId)?.name || "";
  const campaignName = mpo.campaignName || lookup.campaigns.get(mpo.campaignId)?.name || "";
  const clientName = mpo.clientName || lookup.clients.get(lookup.campaigns.get(mpo.campaignId)?.clientId)?.name || "";

  const vendorScore = textSimilarity(row.vendor, vendorName);
  if (vendorScore) {
    score += vendorScore * 35;
    if (vendorScore >= 0.6) reasons.push("vendor");
  }

  const campaignScore = Math.max(
    textSimilarity(row.campaign, campaignName),
    textSimilarity(row.brand, mpo.brand),
    textSimilarity(row.client, clientName)
  );
  if (campaignScore) {
    score += campaignScore * 28;
    if (campaignScore >= 0.6) reasons.push("campaign");
  }

  const period = dateMatchesMpoPeriod(row, mpo);
  if (period.known) score += period.matches ? 12 : -10;

  return { score, reasons };
};

const scoreSpotForRow = (row, spot, mpo) => {
  let score = 0;
  const reasons = [];

  const programmeScore = textSimilarity(row.programme, spot.programme);
  if (programmeScore) {
    score += programmeScore * 34;
    if (programmeScore >= 0.6) reasons.push("programme");
  }

  const materialScore = textSimilarity(row.material, spot.material);
  if (materialScore) {
    score += materialScore * 32;
    if (materialScore >= 0.6) reasons.push("material");
  }

  const rowDuration = parseNumber(row.duration);
  const spotDuration = parseNumber(spot.duration);
  if (rowDuration && spotDuration) score += Math.abs(rowDuration - spotDuration) <= 1 ? 8 : -4;

  const dateCheck = dateMatchesSpot(row, spot, mpo);
  if (dateCheck.known) {
    score += dateCheck.matches ? 18 : -18;
    if (dateCheck.matches) reasons.push("date");
  }

  const timeCheck = formatTimeBandCheck(row, spot);
  if (timeCheck.status === "in") {
    score += 24;
    reasons.push("time band");
  } else if (timeCheck.status === "out") {
    score -= 30;
  } else {
    score += 3;
  }

  return { score, reasons, dateCheck, timeCheck };
};

const findBestMatch = (row, mpos, lookup) => {
  let best = null;

  mpos.forEach(mpo => {
    const mpoScore = scoreMpoForRow(row, mpo, lookup);
    let bestSpot = null;

    (mpo.spots || []).forEach(spot => {
      const spotScore = scoreSpotForRow(row, spot, mpo);
      if (!bestSpot || spotScore.score > bestSpot.score) {
        bestSpot = { spot, ...spotScore };
      }
    });

    const totalScore = mpoScore.score + (bestSpot?.score || 0);
    const candidate = {
      mpo,
      score: totalScore,
      mpoScore: mpoScore.score,
      spot: bestSpot?.score >= 18 ? bestSpot.spot : null,
      spotScore: bestSpot?.score || 0,
      spotDetails: bestSpot,
      reasons: [...mpoScore.reasons, ...(bestSpot?.reasons || [])],
    };

    if (!best || candidate.score > best.score) best = candidate;
  });

  const exactMpo = row.mpoNo && compactText(row.mpoNo) === compactText(best?.mpo?.mpoNo);
  const accepted = Boolean(best && (exactMpo || (best.score >= 58 && (best.mpoScore >= 24 || best.spotScore >= 45))));
  return accepted ? best : null;
};

const deriveReceivableForMpo = (mpo, receivables = []) =>
  receivables.find(item => item.mpoId === mpo.id)
  || receivables.find(item => item.invoiceNo && mpo.invoiceNo && compactText(item.invoiceNo) === compactText(mpo.invoiceNo))
  || null;

const deriveFinancials = (mpo, receivables = []) => {
  const receivable = deriveReceivableForMpo(mpo, receivables);
  const invoiceBase = Number(mpo.invoiceAmount) || Number(mpo.reconciledAmount) || Number(mpo.grandTotal) || Number(mpo.netVal) || 0;
  const amountPaid = receivable
    ? Number(receivable.amountReceived) || 0
    : (mpo.paymentStatus === "paid" ? invoiceBase : 0);
  const outstanding = receivable
    ? Number(receivable.balance) || 0
    : Math.max(invoiceBase - amountPaid, 0);

  return {
    gross: Number(mpo.totalGross) || 0,
    net: Number(mpo.netVal) || Number(mpo.grandTotal) || 0,
    discountPct: (Number(mpo.discPct) || 0) * 100,
    discountAmount: Number(mpo.discAmt) || 0,
    invoiceBase,
    invoiceStatus: mpo.invoiceStatus || "pending",
    paymentStatus: mpo.paymentStatus || "unpaid",
    amountPaid,
    outstanding,
  };
};

const createComplianceResult = (mpo, receivables) => {
  const spotRows = (mpo.spots || []).map((spot, index) => ({
    id: spot.id || `${mpo.id}-spot-${index}`,
    spot,
    planned: getScheduledSpotCount(spot),
    paid: getPaidSpotCount(spot),
    airedInBand: 0,
    outsideBand: 0,
    unverifiedBand: 0,
    matchedRows: [],
  }));

  return {
    mpo,
    financials: deriveFinancials(mpo, receivables),
    spotRows,
    reportRows: [],
    unallocatedRows: [],
    bestScoreTotal: 0,
    bestScoreCount: 0,
  };
};

const finalizeResult = (result) => {
  const plannedFromSpots = result.spotRows.reduce((sum, row) => sum + row.planned, 0);
  const planned = plannedFromSpots || Number(result.mpo.plannedSpotsExecution ?? result.mpo.totalSpots) || 0;
  const airedInBandRaw = result.spotRows.reduce((sum, row) => sum + row.airedInBand, 0);
  const outsideBand = result.spotRows.reduce((sum, row) => sum + row.outsideBand, 0);
  const unverifiedBand = result.spotRows.reduce((sum, row) => sum + row.unverifiedBand, 0);
  const deliveredInBand = Math.min(airedInBandRaw, planned);
  const missed = Math.max(planned - deliveredInBand, 0);
  const extra = Math.max(airedInBandRaw - planned, 0);
  const deliveryPct = planned > 0 ? (deliveredInBand / planned) * 100 : 0;
  const confidence = result.bestScoreCount > 0 ? result.bestScoreTotal / result.bestScoreCount : 0;
  let complianceStatus = "pending";

  if (result.reportRows.length) {
    if (missed <= 0 && outsideBand <= 0 && unverifiedBand <= 0) complianceStatus = "completed";
    else if (deliveredInBand > 0) complianceStatus = "exceptions";
    else complianceStatus = "exceptions";
  }

  return {
    ...result,
    planned,
    airedInBandRaw,
    deliveredInBand,
    outsideBand,
    unverifiedBand,
    missed,
    extra,
    deliveryPct,
    confidence,
    complianceStatus,
  };
};

const buildComplianceResults = ({ reportRows, mpos, vendors, clients, campaigns, receivables }) => {
  const liveMpos = activeOnly(mpos);
  const lookup = {
    vendors: new Map((vendors || []).map(item => [item.id, item])),
    clients: new Map((clients || []).map(item => [item.id, item])),
    campaigns: new Map((campaigns || []).map(item => [item.id, item])),
  };
  const resultMap = new Map(liveMpos.map(mpo => [mpo.id, createComplianceResult(mpo, receivables)]));
  const unmatchedRows = [];

  reportRows.forEach(row => {
    const match = findBestMatch(row, liveMpos, lookup);
    if (!match?.mpo) {
      unmatchedRows.push({ row, reason: "No confident MPO match" });
      return;
    }

    const result = resultMap.get(match.mpo.id);
    const timeCheck = match.spot ? formatTimeBandCheck(row, match.spot) : { status: "unknown", label: "No spot line" };
    const matchedRow = {
      ...row,
      matchScore: match.score,
      matchReasons: match.reasons,
      matchedSpotId: match.spot?.id || "",
      timeCheck,
    };

    result.reportRows.push(matchedRow);
    result.bestScoreTotal += match.score;
    result.bestScoreCount += 1;

    if (!match.spot) {
      result.unallocatedRows.push(matchedRow);
      return;
    }

    const spotRow = result.spotRows.find(item => item.id === match.spot.id);
    if (!spotRow) {
      result.unallocatedRows.push(matchedRow);
      return;
    }

    spotRow.matchedRows.push(matchedRow);
    if (timeCheck.status === "in") spotRow.airedInBand += row.airCount;
    else if (timeCheck.status === "out") spotRow.outsideBand += row.airCount;
    else spotRow.unverifiedBand += row.airCount;
  });

  const results = Array.from(resultMap.values()).map(finalizeResult);
  return { results, unmatchedRows };
};

const statusBadgeColor = (status) => {
  if (status === "completed") return "green";
  if (status === "exceptions") return "orange";
  if (status === "matched") return "blue";
  return "purple";
};

const statusLabel = (status) => {
  if (status === "completed") return "Complete";
  if (status === "exceptions") return "Needs Review";
  if (status === "matched") return "Matched";
  return "Pending";
};

const buildPatchForResult = (result, user) => {
  const mpo = result.mpo;
  const planned = result.planned || Number(mpo.plannedSpotsExecution ?? mpo.totalSpots) || 0;
  const delivered = result.deliveredInBand || 0;
  const missed = Math.max(planned - delivered, 0);
  const ratio = planned > 0 ? Math.min(delivered / planned, 1) : 0;
  const baseAmount = Number(mpo.invoiceAmount) || Number(mpo.grandTotal) || Number(mpo.netVal) || 0;
  const reconciliationStatus =
    planned > 0 && missed <= 0 && result.outsideBand <= 0 && result.unverifiedBand <= 0
      ? "completed"
      : delivered > 0
        ? "in_progress"
        : "not_started";

  const proofStatus =
    result.reportRows.length === 0
      ? (mpo.proofStatus || "pending")
      : missed <= 0 && result.outsideBand <= 0
        ? "received"
        : "partial";

  const noteLine = [
    `Compliance upload ${new Date().toLocaleDateString("en-NG")}`,
    `planned ${planned}`,
    `in-band ${delivered}`,
    `missed ${missed}`,
    `out-of-band ${result.outsideBand}`,
    `unverified ${result.unverifiedBand}`,
  ].join(" | ");

  const existingNotes = normalizeWhitespace(mpo.reconciliationNotes || "");

  return {
    dispatchStatus: mpo.dispatchStatus || "pending",
    dispatchedAt: mpo.dispatchedAt || null,
    dispatchedBy: mpo.dispatchedBy || user?.id || null,
    dispatchContact: mpo.dispatchContact || "",
    dispatchNote: mpo.dispatchNote || "",
    signedMpoUrl: mpo.signedMpoUrl || "",
    invoiceStatus: mpo.invoiceStatus || "pending",
    invoiceNo: mpo.invoiceNo || "",
    invoiceAmount: Number(mpo.invoiceAmount) || baseAmount,
    invoiceReceivedAt: mpo.invoiceReceivedAt || null,
    invoiceUrl: mpo.invoiceUrl || "",
    proofStatus,
    proofUrl: mpo.proofUrl || "",
    proofReceivedAt: mpo.proofReceivedAt || (result.reportRows.length ? new Date().toISOString() : null),
    plannedSpotsExecution: planned,
    airedSpots: delivered,
    missedSpots: missed,
    makegoodSpots: Number(mpo.makegoodSpots) || 0,
    reconciliationStatus,
    reconciliationNotes: existingNotes ? `${existingNotes}\n${noteLine}` : noteLine,
    reconciledAmount: Math.round(baseAmount * ratio * 100) / 100,
    paymentStatus: mpo.paymentStatus || "unpaid",
    paymentReference: mpo.paymentReference || "",
    paidAt: mpo.paidAt || null,
  };
};

function CompliancePage({ user, vendors = [], clients = [], campaigns = [], mpos = [], receivables = [], setMpos }) {
  const [toast, setToast] = useState(null);
  const [reportRows, setReportRows] = useState([]);
  const [fileName, setFileName] = useState("");
  const [isParsing, setIsParsing] = useState(false);
  const [applyingIds, setApplyingIds] = useState([]);
  const [selectedMpoId, setSelectedMpoId] = useState("");
  const [filters, setFilters] = useState({
    search: "",
    status: "",
    vendorId: "",
    campaignId: "",
  });
  const fileInputRef = useRef(null);

  const canApply = hasPermission(user, "manageMpos") || hasPermission(user, "manageMpoStatus");

  const { results, unmatchedRows } = useMemo(
    () => buildComplianceResults({ reportRows, mpos, vendors, clients, campaigns, receivables }),
    [reportRows, mpos, vendors, clients, campaigns, receivables]
  );

  const filteredResults = useMemo(() => {
    const term = normalizeText(filters.search);
    return results.filter(result => {
      if (filters.vendorId && result.mpo.vendorId !== filters.vendorId) return false;
      if (filters.campaignId && result.mpo.campaignId !== filters.campaignId) return false;
      if (filters.status && result.complianceStatus !== filters.status) return false;
      if (!term) return true;
      return [
        result.mpo.mpoNo,
        result.mpo.vendorName,
        result.mpo.clientName,
        result.mpo.brand,
        result.mpo.campaignName,
        result.mpo.invoiceNo,
      ].some(value => normalizeText(value).includes(term));
    });
  }, [results, filters]);

  const selectedResult = useMemo(() => {
    if (selectedMpoId) return results.find(result => result.mpo.id === selectedMpoId) || null;
    return filteredResults.find(result => result.reportRows.length) || filteredResults[0] || results[0] || null;
  }, [results, filteredResults, selectedMpoId]);

  const summary = useMemo(() => {
    const matched = results.filter(result => result.reportRows.length);
    return {
      parsedRows: reportRows.length,
      matchedMpos: matched.length,
      planned: matched.reduce((sum, result) => sum + result.planned, 0),
      delivered: matched.reduce((sum, result) => sum + result.deliveredInBand, 0),
      missed: matched.reduce((sum, result) => sum + result.missed, 0),
      outstanding: results.reduce((sum, result) => sum + result.financials.outstanding, 0),
    };
  }, [results, reportRows.length]);

  const handleFileSelected = async (file) => {
    if (!file) return;
    setFileName(file.name || "monitoring-report");
    setIsParsing(true);
    try {
      const rows = await parseMonitoringFile(file);
      setReportRows(rows);
      setSelectedMpoId("");
      setToast({
        msg: rows.length ? `Parsed ${rows.length} monitoring row${rows.length === 1 ? "" : "s"}.` : "No monitoring rows found in that file.",
        type: rows.length ? "success" : "error",
      });
    } catch (error) {
      setReportRows([]);
      setToast({ msg: error.message || "Failed to parse monitoring report.", type: "error" });
    } finally {
      setIsParsing(false);
      if (fileInputRef.current) fileInputRef.current.value = "";
    }
  };

  const applyResult = async (result) => {
    if (!result?.mpo?.id) return;
    if (!canApply) {
      setToast({ msg: readOnlyMessage(user), type: "error" });
      return;
    }
    if (!result.reportRows.length) {
      setToast({ msg: "No monitoring rows matched this MPO.", type: "error" });
      return;
    }

    setApplyingIds(ids => [...new Set([...ids, result.mpo.id])]);
    try {
      const patch = buildPatchForResult(result, user);
      const updated = await updateMpoExecutionInSupabase(result.mpo.id, patch);
      setMpos(items => items.map(item => item.id === result.mpo.id ? updated : item));
      await createAuditEventInSupabase({
        agencyId: user.agencyId,
        recordType: "mpo",
        recordId: result.mpo.id,
        action: "compliance_reconciled",
        actor: user,
        note: `Compliance upload reconciled ${result.mpo.mpoNo || result.mpo.id}.`,
        metadata: {
          mpoNo: result.mpo.mpoNo || "",
          plannedSpots: patch.plannedSpotsExecution,
          airedSpots: patch.airedSpots,
          missedSpots: patch.missedSpots,
          outOfBandSpots: result.outsideBand,
          unverifiedBandSpots: result.unverifiedBand,
          fileName,
        },
      });
      notifyExecutionUpdate({ agencyId: user.agencyId, mpo: result.mpo, actor: user, patch })
        .catch(error => console.error("Failed to send compliance notification:", error));
      setToast({ msg: `Compliance saved for ${result.mpo.mpoNo || "MPO"}.`, type: "success" });
    } catch (error) {
      setToast({ msg: error.message || "Failed to apply compliance reconciliation.", type: "error" });
    } finally {
      setApplyingIds(ids => ids.filter(id => id !== result.mpo.id));
    }
  };

  const applyAllMatched = async () => {
    const candidates = filteredResults.filter(result => result.reportRows.length);
    if (!candidates.length) {
      setToast({ msg: "No matched MPOs to apply.", type: "error" });
      return;
    }
    for (const result of candidates) {
      await applyResult(result);
    }
  };

  const exportResults = () => {
    const headers = [
      "MPO No.",
      "Vendor",
      "Campaign",
      "Planned Spots",
      "Aired In Timeband",
      "Missed Spots",
      "Out Of Band",
      "Gross Amount",
      "Net Amount",
      "Discount Applied",
      "Invoice Status",
      "Payment Status",
      "Amount Paid",
      "Outstanding",
      "Compliance Status",
    ];
    const rows = filteredResults.map(result => [
      result.mpo.mpoNo || "",
      result.mpo.vendorName || "",
      result.mpo.campaignName || "",
      result.planned,
      result.deliveredInBand,
      result.missed,
      result.outsideBand,
      result.financials.gross,
      result.financials.net,
      result.financials.discountAmount || `${result.financials.discountPct.toFixed(2)}%`,
      invoiceLabelMap[result.financials.invoiceStatus] || result.financials.invoiceStatus,
      paymentLabelMap[result.financials.paymentStatus] || result.financials.paymentStatus,
      result.financials.amountPaid,
      result.financials.outstanding,
      statusLabel(result.complianceStatus),
    ]);
    const blob = new Blob([buildCSV(rows, headers)], { type: "text/csv;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `compliance-reconciliation-${new Date().toISOString().slice(0, 10)}.csv`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  };

  const inputStyle = {
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

      <div style={{ marginBottom: 24, display: "flex", justifyContent: "space-between", gap: 16, alignItems: "flex-start", flexWrap: "wrap" }}>
        <div>
          <h1 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 800, fontSize: 24 }}>Compliance</h1>
          <p style={{ color: "var(--text2)", marginTop: 3, fontSize: 13 }}>Upload monitoring reports, reconcile delivery, and update MPO execution controls.</p>
        </div>
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
          <Btn variant="ghost" onClick={() => fileInputRef.current?.click()} loading={isParsing} icon="^">Upload Report</Btn>
          <Btn variant="blue" onClick={exportResults} disabled={!filteredResults.length} icon="v">Export View</Btn>
          <Btn onClick={applyAllMatched} disabled={!canApply || !filteredResults.some(result => result.reportRows.length)} loading={applyingIds.length > 0} icon="OK">
            Apply Matched
          </Btn>
        </div>
      </div>

      <input
        ref={fileInputRef}
        type="file"
        accept=".csv,.tsv,.txt,.xlsx,.xls"
        style={{ display: "none" }}
        onChange={event => handleFileSelected(event.target.files?.[0])}
      />

      <Card style={{ marginBottom: 20 }}>
        <div style={{ display: "grid", gridTemplateColumns: "1.2fr repeat(4,minmax(150px,1fr))", gap: 12, alignItems: "end" }}>
          <div>
            <label style={{ fontSize: 11, fontWeight: 600, color: "var(--text3)", textTransform: "uppercase", letterSpacing: ".08em", marginBottom: 5, display: "block" }}>Monitoring File</label>
            <button
              type="button"
              onClick={() => fileInputRef.current?.click()}
              style={{
                ...inputStyle,
                minHeight: 42,
                cursor: "pointer",
                display: "flex",
                alignItems: "center",
                justifyContent: "space-between",
                gap: 10,
                textAlign: "left",
              }}
            >
              <span style={{ color: fileName ? "var(--text)" : "var(--text3)", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                {fileName || "Select CSV, TSV, XLS or XLSX"}
              </span>
              <span style={{ color: "var(--accent)", fontWeight: 800 }}>Browse</span>
            </button>
          </div>
          <Field label="Vendor" value={filters.vendorId} onChange={value => setFilters(prev => ({ ...prev, vendorId: value }))} placeholder="All Vendors" options={activeOnly(vendors).map(vendor => ({ value: vendor.id, label: vendor.name }))} />
          <Field label="Campaign" value={filters.campaignId} onChange={value => setFilters(prev => ({ ...prev, campaignId: value }))} placeholder="All Campaigns" options={activeOnly(campaigns).map(campaign => ({ value: campaign.id, label: campaign.name }))} />
          <Field label="Compliance" value={filters.status} onChange={value => setFilters(prev => ({ ...prev, status: value }))} placeholder="All States" options={[
            { value: "completed", label: "Complete" },
            { value: "exceptions", label: "Needs Review" },
            { value: "pending", label: "Pending" },
          ]} />
          <div>
            <label style={{ fontSize: 11, fontWeight: 600, color: "var(--text3)", textTransform: "uppercase", letterSpacing: ".08em", marginBottom: 5, display: "block" }}>Search</label>
            <input value={filters.search} onChange={event => setFilters(prev => ({ ...prev, search: event.target.value }))} placeholder="MPO, vendor, client, campaign" style={inputStyle} />
          </div>
        </div>
      </Card>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(170px,1fr))", gap: 14, marginBottom: 20 }}>
        <Stat icon="Rows" label="Parsed Rows" value={summary.parsedRows} sub={fileName || "No file uploaded"} color="var(--accent)" valueSize="clamp(18px,1.8vw,24px)" />
        <Stat icon="MPO" label="Matched MPOs" value={summary.matchedMpos} sub={`${unmatchedRows.length} unmatched rows`} color="var(--blue)" valueSize="clamp(18px,1.8vw,24px)" />
        <Stat icon="Air" label="Aired In Timeband" value={`${summary.delivered}/${summary.planned || 0}`} sub={summary.planned ? `${Math.round((summary.delivered / summary.planned) * 100)}% delivery` : "No planned spots matched"} color="var(--green)" valueSize="clamp(18px,1.8vw,24px)" />
        <Stat icon="Miss" label="Missed Spots" value={summary.missed} sub="From matched MPOs" color="var(--orange)" valueSize="clamp(18px,1.8vw,24px)" />
        <Stat icon="AR" label="Outstanding" value={fmtN(summary.outstanding)} sub="Across all active MPOs" color="var(--red)" valueSize="clamp(15px,1.5vw,20px)" />
      </div>

      {filteredResults.length === 0 ? (
        <Card>
          <Empty icon="📄" title="No MPOs available for compliance" sub="Create MPOs first, then upload a monitoring report for reconciliation." />
        </Card>
      ) : (
        <div style={{ display: "grid", gridTemplateColumns: "minmax(0,1.35fr) minmax(340px,.85fr)", gap: 18, alignItems: "start" }}>
          <Card>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 12, flexWrap: "wrap", marginBottom: 14 }}>
              <div>
                <h2 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 16 }}>MPO Reconciliation</h2>
                <p style={{ color: "var(--text2)", fontSize: 12, marginTop: 3 }}>{filteredResults.length} MPO{filteredResults.length === 1 ? "" : "s"} in view.</p>
              </div>
              {!canApply ? <Badge color="purple">Read Only</Badge> : null}
            </div>

            <div style={{ overflowX: "auto" }}>
              <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 1080 }}>
                <thead>
                  <tr style={{ background: "var(--bg3)" }}>
                    {["MPO", "Vendor / Campaign", "Spots", "Gross", "Net", "Discount Applied", "Invoice", "Paid", "Outstanding", "Compliance", ""].map(header => (
                      <th key={header} style={{ padding: "8px 10px", textAlign: "left", fontSize: 10, fontWeight: 700, textTransform: "uppercase", letterSpacing: ".07em", color: "var(--text3)", borderBottom: "1px solid var(--border)", whiteSpace: "nowrap" }}>
                        {header}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {filteredResults.map(result => {
                    const isSelected = selectedResult?.mpo.id === result.mpo.id;
                    const applying = applyingIds.includes(result.mpo.id);
                    return (
                      <tr
                        key={result.mpo.id}
                        style={{ borderBottom: "1px solid var(--border)", background: isSelected ? "rgba(240,165,0,.08)" : "transparent" }}
                      >
                        <td style={{ padding: "9px 10px", fontSize: 12, color: "var(--text)", fontWeight: 800, whiteSpace: "nowrap" }}>
                          {result.mpo.mpoNo || "No MPO No."}
                          <div style={{ fontSize: 10, color: "var(--text3)", fontWeight: 600 }}>{reconLabelMap[result.mpo.reconciliationStatus || "not_started"] || "Not Started"}</div>
                        </td>
                        <td style={{ padding: "9px 10px", fontSize: 12, color: "var(--text2)", minWidth: 210 }}>
                          <div style={{ color: "var(--text)", fontWeight: 700 }}>{result.mpo.vendorName || "Unknown Vendor"}</div>
                          <div>{result.mpo.campaignName || result.mpo.brand || "No campaign"}</div>
                        </td>
                        <td style={{ padding: "9px 10px", fontSize: 12, whiteSpace: "nowrap" }}>
                          <strong>{result.deliveredInBand}</strong> / {result.planned}
                          <div style={{ color: result.missed ? "var(--orange)" : "var(--text3)", fontSize: 10 }}>{result.missed} missed</div>
                        </td>
                        <td style={{ padding: "9px 10px", fontSize: 12, color: "var(--text2)", whiteSpace: "nowrap" }}>{fmtN(result.financials.gross)}</td>
                        <td style={{ padding: "9px 10px", fontSize: 12, color: "var(--text2)", whiteSpace: "nowrap" }}>{fmtN(result.financials.net)}</td>
                        <td style={{ padding: "9px 10px", fontSize: 12, color: "var(--text2)", whiteSpace: "nowrap" }}>
                          {result.financials.discountAmount ? fmtN(result.financials.discountAmount) : `${result.financials.discountPct.toFixed(2)}%`}
                        </td>
                        <td style={{ padding: "9px 10px", fontSize: 12, whiteSpace: "nowrap" }}>
                          <Badge color={result.financials.invoiceStatus === "approved" || result.financials.invoiceStatus === "received" ? "blue" : result.financials.invoiceStatus === "disputed" ? "red" : "purple"}>
                            {invoiceLabelMap[result.financials.invoiceStatus] || result.financials.invoiceStatus}
                          </Badge>
                        </td>
                        <td style={{ padding: "9px 10px", fontSize: 12, color: "var(--green)", whiteSpace: "nowrap" }}>{fmtN(result.financials.amountPaid)}</td>
                        <td style={{ padding: "9px 10px", fontSize: 12, color: result.financials.outstanding > 0 ? "var(--red)" : "var(--green)", fontWeight: 700, whiteSpace: "nowrap" }}>{fmtN(result.financials.outstanding)}</td>
                        <td style={{ padding: "9px 10px", fontSize: 12, whiteSpace: "nowrap" }}>
                          <Badge color={statusBadgeColor(result.complianceStatus)}>{statusLabel(result.complianceStatus)}</Badge>
                          <div style={{ color: "var(--text3)", fontSize: 10, marginTop: 4 }}>{result.reportRows.length} report row{result.reportRows.length === 1 ? "" : "s"}</div>
                        </td>
                        <td style={{ padding: "9px 10px", textAlign: "right", whiteSpace: "nowrap" }}>
                          <div style={{ display: "inline-flex", gap: 6 }}>
                            <Btn variant="ghost" size="sm" onClick={() => setSelectedMpoId(result.mpo.id)}>Review</Btn>
                            <Btn size="sm" disabled={!canApply || !result.reportRows.length} loading={applying} onClick={() => applyResult(result)}>Apply</Btn>
                          </div>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </Card>

          <Card>
            {!selectedResult ? (
              <Empty icon="📄" title="No MPO selected" sub="Select an MPO to inspect spot-level compliance." />
            ) : (
              <>
                <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", gap: 12, marginBottom: 14 }}>
                  <div>
                    <h2 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 16 }}>{selectedResult.mpo.mpoNo || "MPO Detail"}</h2>
                    <p style={{ color: "var(--text2)", fontSize: 12, marginTop: 3 }}>{selectedResult.mpo.vendorName || "Unknown Vendor"} / {selectedResult.mpo.campaignName || selectedResult.mpo.brand || "No campaign"}</p>
                  </div>
                  <Badge color={statusBadgeColor(selectedResult.complianceStatus)}>{statusLabel(selectedResult.complianceStatus)}</Badge>
                </div>

                <div style={{ display: "grid", gridTemplateColumns: "repeat(3,minmax(0,1fr))", gap: 10, marginBottom: 14 }}>
                  <div style={{ background: "var(--bg3)", border: "1px solid var(--border)", borderRadius: 10, padding: 10 }}>
                    <div style={{ fontSize: 10, color: "var(--text3)", textTransform: "uppercase", fontWeight: 800 }}>In Band</div>
                    <div style={{ fontFamily: "var(--font-heading)", fontWeight: 800, fontSize: 20, color: "var(--green)" }}>{selectedResult.deliveredInBand}</div>
                  </div>
                  <div style={{ background: "var(--bg3)", border: "1px solid var(--border)", borderRadius: 10, padding: 10 }}>
                    <div style={{ fontSize: 10, color: "var(--text3)", textTransform: "uppercase", fontWeight: 800 }}>Out Band</div>
                    <div style={{ fontFamily: "var(--font-heading)", fontWeight: 800, fontSize: 20, color: "var(--orange)" }}>{selectedResult.outsideBand}</div>
                  </div>
                  <div style={{ background: "var(--bg3)", border: "1px solid var(--border)", borderRadius: 10, padding: 10 }}>
                    <div style={{ fontSize: 10, color: "var(--text3)", textTransform: "uppercase", fontWeight: 800 }}>Missed</div>
                    <div style={{ fontFamily: "var(--font-heading)", fontWeight: 800, fontSize: 20, color: selectedResult.missed ? "var(--red)" : "var(--green)" }}>{selectedResult.missed}</div>
                  </div>
                </div>

                <div style={{ overflowX: "auto", maxHeight: "calc(100vh - 360px)" }}>
                  <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 720 }}>
                    <thead>
                      <tr style={{ background: "var(--bg3)" }}>
                        {["Programme / Material", "Booked Band", "Planned", "Aired", "Out", "Missed", "Status"].map(header => (
                          <th key={header} style={{ padding: "8px 9px", textAlign: "left", fontSize: 10, fontWeight: 700, textTransform: "uppercase", letterSpacing: ".07em", color: "var(--text3)", borderBottom: "1px solid var(--border)", whiteSpace: "nowrap" }}>
                            {header}
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {selectedResult.spotRows.map(row => {
                        const delivered = Math.min(row.airedInBand, row.planned);
                        const missed = Math.max(row.planned - delivered, 0);
                        const rowStatus = missed <= 0 && row.outsideBand <= 0 && row.unverifiedBand <= 0 ? "Aired" : row.matchedRows.length ? "Review" : "Pending";
                        return (
                          <tr key={row.id} style={{ borderBottom: "1px solid var(--border)" }}>
                            <td style={{ padding: "8px 9px", fontSize: 12, color: "var(--text2)", minWidth: 210 }}>
                              <div style={{ color: "var(--text)", fontWeight: 700 }}>{row.spot.programme || "Untitled programme"}</div>
                              <div>{row.spot.material || "No material"}</div>
                            </td>
                            <td style={{ padding: "8px 9px", fontSize: 12, color: "var(--text2)", whiteSpace: "nowrap" }}>{row.spot.timeBelt || "Any"}</td>
                            <td style={{ padding: "8px 9px", fontSize: 12 }}>{row.planned}</td>
                            <td style={{ padding: "8px 9px", fontSize: 12, color: "var(--green)", fontWeight: 800 }}>{delivered}</td>
                            <td style={{ padding: "8px 9px", fontSize: 12, color: row.outsideBand ? "var(--orange)" : "var(--text3)" }}>{row.outsideBand}</td>
                            <td style={{ padding: "8px 9px", fontSize: 12, color: missed ? "var(--red)" : "var(--green)", fontWeight: 700 }}>{missed}</td>
                            <td style={{ padding: "8px 9px", fontSize: 12 }}>
                              <Badge color={rowStatus === "Aired" ? "green" : rowStatus === "Review" ? "orange" : "purple"}>{rowStatus}</Badge>
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>

                {selectedResult.unallocatedRows.length ? (
                  <div style={{ marginTop: 12, padding: 12, borderRadius: 10, border: "1px solid rgba(249,115,22,.25)", background: "rgba(249,115,22,.08)", color: "var(--orange)", fontSize: 12, lineHeight: 1.45 }}>
                    {selectedResult.unallocatedRows.length} matched monitoring row{selectedResult.unallocatedRows.length === 1 ? "" : "s"} reached this MPO but did not land on a specific spot line.
                  </div>
                ) : null}
              </>
            )}
          </Card>
        </div>
      )}

      {unmatchedRows.length ? (
        <Card style={{ marginTop: 18 }}>
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, flexWrap: "wrap", marginBottom: 12 }}>
            <div>
              <h2 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 16 }}>Unmatched Monitoring Rows</h2>
              <p style={{ color: "var(--text2)", fontSize: 12, marginTop: 3 }}>Rows that need MPO number, vendor, campaign, or spot detail review.</p>
            </div>
            <Badge color="orange">{unmatchedRows.length} unmatched</Badge>
          </div>
          <div style={{ overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 860 }}>
              <thead>
                <tr style={{ background: "var(--bg3)" }}>
                  {["Source", "MPO No.", "Vendor", "Campaign", "Programme", "Material", "Date", "Time", "Spots", "Reason"].map(header => (
                    <th key={header} style={{ padding: "8px 10px", textAlign: "left", fontSize: 10, fontWeight: 700, textTransform: "uppercase", letterSpacing: ".07em", color: "var(--text3)", borderBottom: "1px solid var(--border)", whiteSpace: "nowrap" }}>{header}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {unmatchedRows.slice(0, 25).map(({ row, reason }) => (
                  <tr key={row.id} style={{ borderBottom: "1px solid var(--border)" }}>
                    <td style={{ padding: "8px 10px", fontSize: 12, color: "var(--text3)" }}>{row.source || fileName || "Report"}</td>
                    <td style={{ padding: "8px 10px", fontSize: 12 }}>{row.mpoNo || "-"}</td>
                    <td style={{ padding: "8px 10px", fontSize: 12 }}>{row.vendor || "-"}</td>
                    <td style={{ padding: "8px 10px", fontSize: 12 }}>{row.campaign || row.brand || "-"}</td>
                    <td style={{ padding: "8px 10px", fontSize: 12 }}>{row.programme || "-"}</td>
                    <td style={{ padding: "8px 10px", fontSize: 12 }}>{row.material || "-"}</td>
                    <td style={{ padding: "8px 10px", fontSize: 12 }}>{row.airedDate?.iso || row.dateRaw || "-"}</td>
                    <td style={{ padding: "8px 10px", fontSize: 12 }}>{row.time || row.timeBand || "-"}</td>
                    <td style={{ padding: "8px 10px", fontSize: 12 }}>{row.airCount}</td>
                    <td style={{ padding: "8px 10px", fontSize: 12, color: "var(--orange)" }}>{reason}</td>
                  </tr>
                ))}
              </tbody>
            </table>
            {unmatchedRows.length > 25 ? (
              <div style={{ marginTop: 10, fontSize: 12, color: "var(--text3)" }}>Showing 25 of {unmatchedRows.length} unmatched rows.</div>
            ) : null}
          </div>
        </Card>
      ) : null}
    </div>
  );
}

export default memo(CompliancePage);
