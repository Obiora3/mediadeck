// ===============================
// NUMBER & CURRENCY
// ===============================

export function fmt(value, digits = 0) {
  return new Intl.NumberFormat("en-NG", {
    minimumFractionDigits: digits,
    maximumFractionDigits: digits,
  }).format(Number(value || 0));
}

export function fmtN(value, digits = 0) {
  return `₦${fmt(value, digits)}`;
}

export function formatCurrency(value) {
  return new Intl.NumberFormat("en-NG", {
    style: "currency",
    currency: "NGN",
    maximumFractionDigits: 0,
  }).format(Number(value || 0));
}

export function formatCompactCurrency(value) {
  const num = Number(value || 0);

  if (num >= 1000000000) return `₦${(num / 1000000000).toFixed(1).replace(/\.0$/, "")}B`;
  if (num >= 1000000) return `₦${(num / 1000000).toFixed(1).replace(/\.0$/, "")}M`;
  if (num >= 1000) return `₦${(num / 1000).toFixed(1).replace(/\.0$/, "")}K`;

  return formatCurrency(num);
}

export function formatNumber(value) {
  return new Intl.NumberFormat("en-NG").format(Number(value || 0));
}

export function formatNumberFixed(value, digits = 2) {
  return Number(value || 0).toFixed(digits);
}

export function formatNumberFixed2(value) {
  return Number(value || 0).toFixed(2);
}

export function formatPercent(value, digits = 0) {
  return `${Number(value || 0).toFixed(digits)}%`;
}

// ===============================
// DATE & TIME
// ===============================

export function formatDate(value) {
  if (!value) return "—";
  const d = new Date(value);
  if (Number.isNaN(d.getTime())) return "—";

  return d.toLocaleDateString("en-GB", {
    day: "numeric",
    month: "short",
    year: "numeric",
  });
}

export function formatDateOnly(value) {
  return formatDate(value);
}

export function formatDateTime(value) {
  if (!value) return "—";
  const d = new Date(value);
  if (Number.isNaN(d.getTime())) return "—";

  return d.toLocaleString("en-GB", {
    day: "numeric",
    month: "short",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit",
  });
}

export function formatTime(value) {
  if (!value) return "—";
  const d = new Date(value);
  if (Number.isNaN(d.getTime())) return value;

  return d.toLocaleTimeString("en-GB", {
    hour: "2-digit",
    minute: "2-digit",
  });
}

export function formatRelativeTime(value) {
  if (!value) return "—";

  const then = new Date(value).getTime();
  if (Number.isNaN(then)) return "—";

  const diff = Date.now() - then;
  const mins = Math.floor(diff / 60000);
  const hrs = Math.floor(diff / 3600000);
  const days = Math.floor(diff / 86400000);

  if (mins < 1) return "just now";
  if (mins < 60) return `${mins} min ago`;
  if (hrs < 24) return `${hrs} hr ago`;
  if (days < 30) return `${days} day${days === 1 ? "" : "s"} ago`;

  return formatDate(value);
}

// ===============================
// CALENDAR / DATE HELPERS
// ===============================

export function getCurrentYear() {
  return new Date().getFullYear();
}

export function getCurrentMonth() {
  return new Date().getMonth();
}

export function getCurrentDate() {
  return new Date();
}

export function getMonthLabel(monthIndex) {
  return MONTH_NAMES_SHORT[monthIndex] || "—";
}

export function getDayLabel(dayIndex) {
  return DAY_LABELS_SHORT[dayIndex] || "—";
}

export function getMonthName(monthIndex) {
  return MONTH_NAMES_FULL[monthIndex] || "—";
}

export function getMonthNameShort(monthIndex) {
  return MONTH_NAMES_SHORT[monthIndex] || "—";
}

export function getDaysInMonth(year, monthIndex) {
  return new Date(year, monthIndex + 1, 0).getDate();
}

export function startOfMonth(year, monthIndex) {
  return new Date(year, monthIndex, 1);
}

export function endOfMonth(year, monthIndex) {
  return new Date(year, monthIndex + 1, 0);
}

export function isSameDay(a, b) {
  if (!a || !b) return false;
  const da = new Date(a);
  const db = new Date(b);
  return (
    da.getFullYear() === db.getFullYear() &&
    da.getMonth() === db.getMonth() &&
    da.getDate() === db.getDate()
  );
}

export function toIsoDate(value) {
  if (!value) return "";
  const d = new Date(value);
  if (Number.isNaN(d.getTime())) return "";
  return d.toISOString().slice(0, 10);
}

export function resolveMonthIndex(value) {
  if (value === null || value === undefined || value === "") return -1;

  if (typeof value === "number" && Number.isInteger(value)) {
    if (value >= 0 && value <= 11) return value;
    if (value >= 1 && value <= 12) return value - 1;
    return -1;
  }

  const raw = String(value).trim();
  if (!raw) return -1;

  const numeric = Number(raw);
  if (!Number.isNaN(numeric)) {
    if (numeric >= 0 && numeric <= 11) return numeric;
    if (numeric >= 1 && numeric <= 12) return numeric - 1;
  }

  const lower = raw.toLowerCase();

  const fullIndex = MONTH_NAMES_FULL.findIndex(
    (month) => month.toLowerCase() === lower
  );
  if (fullIndex !== -1) return fullIndex;

  const shortIndex = MONTH_NAMES_SHORT.findIndex(
    (month) => month.toLowerCase() === lower
  );
  if (shortIndex !== -1) return shortIndex;

  return -1;
}

export function resolveMonthName(value) {
  const idx = resolveMonthIndex(value);
  return idx >= 0 ? MONTH_NAMES_FULL[idx] : "—";
}

// ===============================
// LABELS & HELPERS
// ===============================

export function formatRoleLabel(role) {
  if (!role) return "—";

  return String(role)
    .replace(/_/g, " ")
    .replace(/\b\w/g, (m) => m.toUpperCase());
}

export function formatRole(role) {
  return formatRoleLabel(role);
}

export function formatAuditTimestamp(value) {
  return formatDateTime(value);
}

// ===============================
// CONSTANTS
// ===============================

export const MONTH_NAMES_FULL = [
  "January",
  "February",
  "March",
  "April",
  "May",
  "June",
  "July",
  "August",
  "September",
  "October",
  "November",
  "December",
];

export const MONTH_NAMES_SHORT = [
  "Jan",
  "Feb",
  "Mar",
  "Apr",
  "May",
  "Jun",
  "Jul",
  "Aug",
  "Sep",
  "Oct",
  "Nov",
  "Dec",
];

export const DAY_LABELS_SHORT = [
  "Sun",
  "Mon",
  "Tue",
  "Wed",
  "Thu",
  "Fri",
  "Sat",
];