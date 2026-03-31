import { hasPermission, normalizeRole } from "./roles";
import { isArchived } from "../utils/records";

export const MPO_STATUS_OPTIONS = [
  { value: "draft", label: "Draft" },
  { value: "submitted", label: "Submitted" },
  { value: "reviewed", label: "Reviewed" },
  { value: "approved", label: "Approved" },
  { value: "sent", label: "Sent to Vendor" },
  { value: "aired", label: "Aired" },
  { value: "reconciled", label: "Reconciled" },
  { value: "closed", label: "Closed" },
  { value: "rejected", label: "Rejected" },
];
export const MPO_STATUS_LABELS = Object.fromEntries(MPO_STATUS_OPTIONS.map(option => [option.value, option.label]));
export const MPO_EXECUTION_STATUS_OPTIONS = [{ value: "pending", label: "Pending Dispatch" }, { value: "sent", label: "Sent to Vendor" }, { value: "confirmed", label: "Vendor Confirmed" }];
export const MPO_INVOICE_STATUS_OPTIONS = [{ value: "pending", label: "Pending Invoice" }, { value: "received", label: "Invoice Received" }, { value: "approved", label: "Invoice Approved" }, { value: "disputed", label: "Invoice Disputed" }];
export const MPO_PROOF_STATUS_OPTIONS = [{ value: "pending", label: "Pending Proof" }, { value: "partial", label: "Partial Proof" }, { value: "received", label: "Proof Received" }, { value: "disputed", label: "Proof Disputed" }];
export const MPO_PAYMENT_STATUS_OPTIONS = [{ value: "unpaid", label: "Unpaid" }, { value: "processing", label: "Processing" }, { value: "paid", label: "Paid" }, { value: "disputed", label: "Disputed" }];
export const MPO_RECON_STATUS_OPTIONS = [{ value: "not_started", label: "Not Started" }, { value: "in_progress", label: "In Progress" }, { value: "ready", label: "Ready for Review" }, { value: "completed", label: "Completed" }];
export const toIsoInput = (value) => value ? String(value).slice(0, 16) : "";
export const toIsoOrNull = (value) => value ? new Date(value).toISOString() : null;
export const fmtDateTime = (value) => value ? new Date(value).toLocaleString('en-NG') : "—";
export const getExecutionHealthColor = (mpo) => {
  if ((mpo.reconciliationStatus || 'not_started') === 'completed') return 'green';
  if ((mpo.invoiceStatus || 'pending') === 'disputed' || (mpo.proofStatus || 'pending') === 'disputed') return 'red';
  if ((mpo.dispatchStatus || 'pending') === 'confirmed') return 'blue';
  if ((mpo.dispatchStatus || 'pending') === 'sent') return 'purple';
  return 'gray';
};
export const getExecutionHealthLabel = (mpo) => {
  if ((mpo.reconciliationStatus || 'not_started') === 'completed') return 'Reconciled';
  if ((mpo.invoiceStatus || 'pending') === 'disputed' || (mpo.proofStatus || 'pending') === 'disputed') return 'Disputed';
  if ((mpo.dispatchStatus || 'pending') === 'confirmed') return 'Confirmed';
  if ((mpo.dispatchStatus || 'pending') === 'sent') return 'Dispatched';
  return 'Pending';
};
export const MPO_WORKFLOW_TRANSITIONS = {
  admin: {
    draft: ["submitted", "approved", "rejected"],
    submitted: ["reviewed", "approved", "rejected"],
    reviewed: ["approved", "rejected"],
    approved: ["sent", "reconciled", "closed", "rejected"],
    rejected: ["draft", "submitted"],
    sent: ["aired", "reconciled", "closed"],
    aired: ["reconciled", "closed"],
    reconciled: ["closed"],
    closed: [],
  },
  planner: {
    draft: ["submitted"],
    rejected: ["submitted"],
  },
  buyer: {
    draft: ["submitted"],
    rejected: ["submitted"],
    approved: ["sent"],
    sent: ["aired"],
    aired: ["reconciled"],
    reconciled: ["closed"],
  },
  finance: {
    submitted: ["reviewed", "rejected"],
    reviewed: ["approved", "rejected"],
    approved: ["reconciled"],
  },
  viewer: {},
};
export const getAllowedMpoStatusTargets = (user, mpo) => {
  const role = normalizeRole(user?.role);
  const current = String(mpo?.status || "draft").toLowerCase();
  return MPO_WORKFLOW_TRANSITIONS[role]?.[current] || [];
};
export const mpoStatusNeedsNote = (status) => ["submitted", "reviewed", "approved", "rejected", "closed"].includes(String(status || "").toLowerCase());
export const MPO_WAITING_OWNER = {
  draft: { label: "Planner / Buyer", roles: ["planner", "buyer", "admin"], color: "accent", hint: "Complete the MPO and submit it for finance review." },
  submitted: { label: "Finance Review", roles: ["finance", "admin"], color: "blue", hint: "Finance should review rates, controls, and support notes." },
  reviewed: { label: "Admin Approval", roles: ["admin"], color: "purple", hint: "Awaiting final leadership approval before dispatch." },
  approved: { label: "Buyer / Planner Dispatch", roles: ["buyer", "planner", "admin"], color: "teal", hint: "Send the approved MPO to the vendor and confirm dispatch." },
  rejected: { label: "Planner / Buyer Revision", roles: ["planner", "buyer", "admin"], color: "red", hint: "Apply requested changes, update the MPO, and resubmit." },
  sent: { label: "Buyer / Planner Follow-up", roles: ["buyer", "planner", "admin"], color: "orange", hint: "Monitor airing and collect proofs from the vendor." },
  aired: { label: "Finance Reconciliation", roles: ["finance", "admin"], color: "blue", hint: "Reconcile proof, invoice, and final payable." },
  reconciled: { label: "Finance / Admin Close-out", roles: ["finance", "admin"], color: "green", hint: "Record final payment and close the MPO." },
  closed: { label: "Completed", roles: [], color: "green", hint: "This MPO has completed the workflow." },
};
export const getMpoWorkflowMeta = (mpo) => {
  const current = String(mpo?.status || "draft").toLowerCase();
  return MPO_WAITING_OWNER[current] || MPO_WAITING_OWNER.draft;
};
export const isMpoAwaitingUser = (user, mpo) => {
  if (!user || isArchived(mpo)) return false;
  const current = String(mpo?.status || "draft").toLowerCase();
  if (current === "closed") return false;
  return getMpoWorkflowMeta(mpo).roles.includes(normalizeRole(user?.role));
};
export const getWorkflowActionLabel = (currentStatus, targetStatus) => {
  const current = String(currentStatus || "draft").toLowerCase();
  const target = String(targetStatus || "").toLowerCase();
  if (target === "submitted") return current === "rejected" ? "Resubmit MPO" : "Submit for Review";
  if (target === "reviewed") return "Mark Reviewed";
  if (target === "approved") return "Approve MPO";
  if (target === "rejected") return ["submitted", "reviewed", "approved"].includes(current) ? "Request Changes" : "Reject MPO";
  if (target === "sent") return "Send to Vendor";
  if (target === "aired") return "Mark as Aired";
  if (target === "reconciled") return current === "approved" ? "Move to Reconciliation" : "Mark Reconciled";
  if (target === "closed") return "Close MPO";
  return MPO_STATUS_LABELS[target] || target;
};
export const getWorkflowActionVariant = (targetStatus) => {
  const target = String(targetStatus || "").toLowerCase();
  if (target === "approved" || target === "closed") return "success";
  if (target === "reviewed" || target === "reconciled") return "blue";
  if (target === "sent") return "purple";
  if (target === "aired") return "secondary";
  if (target === "rejected") return "danger";
  return "ghost";
};
export const getQuickWorkflowActions = (user, mpo) => {
  const current = String(mpo?.status || "draft").toLowerCase();
  return getAllowedMpoStatusTargets(user, mpo).map(target => ({
    value: target,
    label: getWorkflowActionLabel(current, target),
    variant: getWorkflowActionVariant(target),
  }));
};
export const canEditMpoContent = (user, mpo) => {
  const role = normalizeRole(user?.role);
  if (role === "admin") return true;
  if (!hasPermission(user, "manageMpos")) return false;
  return ["draft", "rejected"].includes(String(mpo?.status || "draft").toLowerCase());
};
export const MPO_STATUS_COLORS = { draft: "accent", submitted: "blue", reviewed: "purple", approved: "green", sent: "teal", aired: "orange", reconciled: "blue", closed: "green", rejected: "red" };
