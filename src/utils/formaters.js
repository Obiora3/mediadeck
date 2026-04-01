// ===============================
// NUMBER & CURRENCY FORMATTERS
// ===============================

export function formatCurrency(value) {
  return new Intl.NumberFormat("en-NG", {
    style: "currency",
    currency: "NGN",
    maximumFractionDigits: 0,
  }).format(value || 0);
}

export function formatNumber(value) {
  return new Intl.NumberFormat("en-NG").format(value || 0);
}

// ===============================
// DATE FORMATTERS
// ===============================

export function formatDate(date) {
  if (!date) return "-";

  try {
    return new Date(date).toLocaleDateString("en-GB", {
      day: "2-digit",
      month: "short",
      year: "numeric",
    });
  } catch {
    return "-";
  }
}

export function formatDateTime(date) {
  if (!date) return "-";

  try {
    return new Date(date).toLocaleString("en-GB", {
      day: "2-digit",
      month: "short",
      year: "numeric",
      hour: "2-digit",
      minute: "2-digit",
    });
  } catch {
    return "-";
  }
}

// ===============================
// RELATIVE TIME (e.g. "2 hrs ago")
// ===============================

export function formatRelativeTime(date) {
  if (!date) return "-";

  const now = new Date();
  const past = new Date(date);
  const diffMs = now - past;

  const minutes = Math.floor(diffMs / (1000 * 60));
  const hours = Math.floor(diffMs / (1000 * 60 * 60));
  const days = Math.floor(diffMs / (1000 * 60 * 60 * 24));

  if (minutes < 1) return "just now";
  if (minutes < 60) return `${minutes} min ago`;
  if (hours < 24) return `${hours} hr ago`;
  return `${days} day(s) ago`;
}

// ===============================
// ROLE LABEL FORMATTER
// ===============================

export function formatRole(role) {
  if (!role) return "-";

  return role
    .replace(/_/g, " ")
    .toLowerCase()
    .replace(/\b\w/g, (c) => c.toUpperCase());
}

// ===============================
// MONTH + DAY LABELS
// ===============================

export const MONTHS = [
  "Jan", "Feb", "Mar", "Apr", "May", "Jun",
  "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"
];

export const DAYS = [
  "Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"
];