import { supabase } from "../lib/supabase";
import { normalizeRole } from "../constants/roles";
import { MPO_STATUS_LABELS } from "../constants/mpoWorkflow";

const auditEventSelect = `
  id,
  agency_id,
  record_type,
  record_id,
  action,
  actor_id,
  actor_name,
  actor_role,
  note,
  metadata,
  created_at
`;

const notificationSelect = `
  id,
  agency_id,
  recipient_user_id,
  category,
  title,
  message,
  record_type,
  record_id,
  link_page,
  actor_id,
  actor_name,
  actor_role,
  metadata,
  read_at,
  created_at
`;

const mapAuditEventFromSupabase = (row) => ({
  id: row.id,
  agencyId: row.agency_id,
  recordType: row.record_type || "workspace",
  recordId: row.record_id || null,
  action: row.action || "updated",
  actorId: row.actor_id || null,
  actorName: row.actor_name || "System",
  actorRole: row.actor_role || "system",
  note: row.note || "",
  metadata: row.metadata || {},
  createdAt: row.created_at || null,
});

export const createAuditEventInSupabase = async ({ agencyId, recordType, recordId = null, action, actor, note = "", metadata = {} }) => {
  if (!agencyId || !recordType || !action) return null;
  const payload = {
    agency_id: agencyId,
    record_type: String(recordType),
    record_id: recordId || null,
    action: String(action),
    actor_id: actor?.id || null,
    actor_name: actor?.name || actor?.email || "System",
    actor_role: normalizeRole(actor?.role || "system"),
    note: String(note || "").trim() || null,
    metadata: metadata && typeof metadata === "object" ? metadata : {},
  };
  const { data, error } = await supabase
    .from("audit_events")
    .insert([payload])
    .select(auditEventSelect)
    .single();
  if (error) {
    if ((error.message || "").toLowerCase().includes("audit_events")) {
      throw new Error('Missing table "audit_events". Run the provided SQL patch in Supabase SQL Editor.');
    }
    throw error;
  }
  return mapAuditEventFromSupabase(data);
};

export const fetchAuditEventsForRecord = async (agencyId, recordType, recordId) => {
  if (!agencyId || !recordType || !recordId) return [];
  const { data, error } = await supabase
    .from("audit_events")
    .select(auditEventSelect)
    .eq("agency_id", agencyId)
    .eq("record_type", recordType)
    .eq("record_id", recordId)
    .order("created_at", { ascending: false });
  if (error) {
    if ((error.message || "").toLowerCase().includes("audit_events")) {
      throw new Error('Missing table "audit_events". Run the provided SQL patch in Supabase SQL Editor.');
    }
    throw error;
  }
  return (data || []).map(mapAuditEventFromSupabase);
};

export const fetchAuditEventsForAgency = async (agencyId, filter = "all", limit = 80) => {
  if (!agencyId) return [];
  let query = supabase
    .from("audit_events")
    .select(auditEventSelect)
    .eq("agency_id", agencyId)
    .order("created_at", { ascending: false })
    .limit(limit);
  if (filter && filter !== "all") query = query.eq("record_type", filter);
  const { data, error } = await query;
  if (error) {
    if ((error.message || "").toLowerCase().includes("audit_events")) {
      throw new Error('Missing table "audit_events". Run the provided SQL patch in Supabase SQL Editor.');
    }
    throw error;
  }
  return (data || []).map(mapAuditEventFromSupabase);
};

const mapNotificationFromSupabase = (row) => ({
  id: row.id,
  agencyId: row.agency_id,
  recipientUserId: row.recipient_user_id,
  category: row.category || "workspace",
  title: row.title || "Notification",
  message: row.message || "",
  recordType: row.record_type || null,
  recordId: row.record_id || null,
  linkPage: row.link_page || "dashboard",
  actorId: row.actor_id || null,
  actorName: row.actor_name || "System",
  actorRole: row.actor_role || "system",
  metadata: row.metadata || {},
  readAt: row.read_at || null,
  createdAt: row.created_at || null,
});

export const fetchNotificationsFromSupabase = async (userId, agencyId, limit = 50) => {
  if (!userId || !agencyId) return [];
  const { data, error } = await supabase
    .from("notifications")
    .select(notificationSelect)
    .eq("agency_id", agencyId)
    .eq("recipient_user_id", userId)
    .order("created_at", { ascending: false })
    .limit(limit);
  if (error) {
    if ((error.message || "").toLowerCase().includes("notifications")) {
      throw new Error('Missing table "notifications". Run the provided SQL patch in Supabase SQL Editor.');
    }
    throw error;
  }
  return (data || []).map(mapNotificationFromSupabase);
};

export const markNotificationReadInSupabase = async (notificationId) => {
  if (!notificationId) return null;
  const { data, error } = await supabase
    .from("notifications")
    .update({ is_read: true, read_at: new Date().toISOString() })
    .eq("id", notificationId)
    .select(notificationSelect)
    .single();
  if (error) throw error;
  return mapNotificationFromSupabase(data);
};

export const markAllNotificationsReadInSupabase = async (userId, agencyId) => {
  if (!userId || !agencyId) return 0;
  const { data, error } = await supabase
    .from("notifications")
    .update({ is_read: true, read_at: new Date().toISOString() })
    .eq("agency_id", agencyId)
    .eq("recipient_user_id", userId)
    .is("read_at", null)
    .select("id");
  if (error) throw error;
  return (data || []).length;
};

export const createNotificationForUserInSupabase = async ({ agencyId, recipientUserId, category = "workspace", title, message = "", recordType = null, recordId = null, linkPage = "dashboard", actor = null, metadata = {} }) => {
  if (!agencyId || !recipientUserId || !title) return null;
  const payload = {
    agency_id: agencyId,
    recipient_user_id: recipientUserId,
    category: String(category || "workspace"),
    title: String(title),
    message: String(message || "").trim() || null,
    record_type: recordType || null,
    record_id: recordId || null,
    link_page: linkPage || "dashboard",
    actor_id: actor?.id || null,
    actor_name: actor?.name || actor?.email || "System",
    actor_role: normalizeRole(actor?.role || "system"),
    metadata: metadata && typeof metadata === "object" ? metadata : {},
  };
  const { data, error } = await supabase.from("notifications").insert([payload]).select(notificationSelect).single();
  if (error) {
    if ((error.message || "").toLowerCase().includes("notifications")) {
      throw new Error('Missing table "notifications". Run the provided SQL patch in Supabase SQL Editor.');
    }
    throw error;
  }
  return mapNotificationFromSupabase(data);
};

const createNotificationsForRolesInSupabase = async ({ agencyId, roles = [], excludeUserId = null, category = "workspace", title, message = "", recordType = null, recordId = null, linkPage = "dashboard", actor = null, metadata = {} }) => {
  if (!agencyId || !title) return [];
  let query = supabase.from("profiles").select("id, role").eq("agency_id", agencyId);
  const normalizedRoles = (roles || []).map(normalizeRole).filter(Boolean);
  if (normalizedRoles.length) query = query.in("role", normalizedRoles);
  const { data: recipients, error: recipientsError } = await query;
  if (recipientsError) throw recipientsError;
  const payload = (recipients || [])
    .filter(row => row.id && row.id !== excludeUserId)
    .map(row => ({
      agency_id: agencyId,
      recipient_user_id: row.id,
      category: String(category || "workspace"),
      title: String(title),
      message: String(message || "").trim() || null,
      record_type: recordType || null,
      record_id: recordId || null,
      link_page: linkPage || "dashboard",
      actor_id: actor?.id || null,
      actor_name: actor?.name || actor?.email || "System",
      actor_role: normalizeRole(actor?.role || "system"),
      metadata: metadata && typeof metadata === "object" ? metadata : {},
    }));
  if (!payload.length) return [];
  const { data, error } = await supabase.from("notifications").insert(payload).select(notificationSelect);
  if (error) {
    if ((error.message || "").toLowerCase().includes("notifications")) {
      throw new Error('Missing table "notifications". Run the provided SQL patch in Supabase SQL Editor.');
    }
    throw error;
  }
  return (data || []).map(mapNotificationFromSupabase);
};

export const notifyMpoWorkflowTransition = async ({ agencyId, mpo, nextStatus, actor, note = "" }) => {
  const status = String(nextStatus || "").toLowerCase();
  const mpoNo = mpo?.mpoNo || "MPO";
  const titleMap = {
    submitted: `${mpoNo} submitted for review`,
    reviewed: `${mpoNo} reviewed by Finance`,
    approved: `${mpoNo} approved`,
    rejected: `${mpoNo} rejected`,
    sent: `${mpoNo} sent to vendor`,
    aired: `${mpoNo} marked as aired`,
    reconciled: `${mpoNo} marked as reconciled`,
    closed: `${mpoNo} closed`,
  };
  const recipientRolesMap = {
    submitted: ["finance", "admin"],
    reviewed: ["admin", "planner", "buyer"],
    approved: ["admin", "planner", "buyer", "finance"],
    rejected: ["admin", "planner", "buyer"],
    sent: ["admin", "finance"],
    aired: ["admin", "finance"],
    reconciled: ["admin", "finance", "planner", "buyer"],
    closed: ["admin", "finance", "planner", "buyer"],
  };
  const title = titleMap[status] || `${mpoNo} updated`;
  const message = note?.trim() || `${actor?.name || "A teammate"} moved ${mpoNo} to ${MPO_STATUS_LABELS[status] || status}.`;
  return createNotificationsForRolesInSupabase({
    agencyId,
    roles: recipientRolesMap[status] || ["admin"],
    excludeUserId: actor?.id || null,
    category: "mpo",
    title,
    message,
    recordType: "mpo",
    recordId: mpo?.id || null,
    linkPage: "mpo",
    actor,
    metadata: { mpoNo, fromStatus: mpo?.status || "draft", toStatus: status },
  });
};

export const notifyExecutionUpdate = async ({ agencyId, mpo, actor, patch }) => {
  const mpoNo = mpo?.mpoNo || "MPO";
  const tasks = [];
  if ((patch?.invoiceStatus || "").toLowerCase() === "received") {
    tasks.push(createNotificationsForRolesInSupabase({
      agencyId,
      roles: ["finance", "admin"],
      excludeUserId: actor?.id || null,
      category: "finance",
      title: `${mpoNo} invoice received`,
      message: `${actor?.name || "A teammate"} uploaded or recorded an invoice for ${mpoNo}.`,
      recordType: "mpo",
      recordId: mpo?.id || null,
      linkPage: "mpo",
      actor,
      metadata: { mpoNo, invoiceStatus: patch.invoiceStatus, paymentStatus: patch.paymentStatus || "unpaid" },
    }));
  }
  if ((patch?.proofStatus || "").toLowerCase() === "received") {
    tasks.push(createNotificationsForRolesInSupabase({
      agencyId,
      roles: ["admin", "planner", "buyer", "finance"],
      excludeUserId: actor?.id || null,
      category: "proof",
      title: `${mpoNo} proof received`,
      message: `${actor?.name || "A teammate"} uploaded proof of airing for ${mpoNo}.`,
      recordType: "mpo",
      recordId: mpo?.id || null,
      linkPage: "mpo",
      actor,
      metadata: { mpoNo, proofStatus: patch.proofStatus },
    }));
  }
  if (["ready", "completed"].includes(String(patch?.reconciliationStatus || "").toLowerCase())) {
    tasks.push(createNotificationsForRolesInSupabase({
      agencyId,
      roles: ["finance", "admin", "planner", "buyer"],
      excludeUserId: actor?.id || null,
      category: "reconciliation",
      title: `${mpoNo} reconciliation ${String(patch.reconciliationStatus).toLowerCase() === "completed" ? "completed" : "ready for review"}`,
      message: `${actor?.name || "A teammate"} updated reconciliation for ${mpoNo}.`,
      recordType: "mpo",
      recordId: mpo?.id || null,
      linkPage: "mpo",
      actor,
      metadata: { mpoNo, reconciliationStatus: patch.reconciliationStatus, paymentStatus: patch.paymentStatus || "unpaid" },
    }));
  }
  if ((patch?.paymentStatus || "").toLowerCase() === "paid") {
    tasks.push(createNotificationsForRolesInSupabase({
      agencyId,
      roles: ["admin", "planner", "buyer", "finance"],
      excludeUserId: actor?.id || null,
      category: "finance",
      title: `${mpoNo} marked as paid`,
      message: `${actor?.name || "A teammate"} recorded payment for ${mpoNo}.`,
      recordType: "mpo",
      recordId: mpo?.id || null,
      linkPage: "mpo",
      actor,
      metadata: { mpoNo, paymentStatus: patch.paymentStatus, paymentReference: patch.paymentReference || "" },
    }));
  }
  if (!tasks.length) return [];
  return Promise.allSettled(tasks);
};
