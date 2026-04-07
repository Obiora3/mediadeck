import { supabase } from "../lib/supabase";

const store = {
  get: (k, d = null) => {
    try {
      const v = localStorage.getItem(k);
      return v ? JSON.parse(v) : d;
    } catch {
      return d;
    }
  },
  set: (k, v) => {
    try {
      localStorage.setItem(k, JSON.stringify(v));
    } catch {}
  },
  del: (k) => {
    try {
      localStorage.removeItem(k);
    } catch {}
  },
};

const uid = () => Math.random().toString(36).slice(2) + Date.now().toString(36);

const receivablesKeyForAgency = (agencyId) =>
  agencyId ? `msp_receivables_${agencyId}` : "msp_receivables";

const receivableListSelect = `
  id,
  agency_id,
  mpo_id,
  client_id,
  campaign_id,
  invoice_no,
  invoice_date,
  due_date,
  gross_amount,
  status,
  collection_stage,
  owner,
  source,
  notes,
  last_payment_at,
  last_follow_up_at,
  created_at,
  updated_at
`;

const receivableDetailSelect = `*, payments:receivable_payments(*)`;

export const isoToday = () => new Date().toISOString().slice(0, 10);

const addDaysToIso = (isoDate = isoToday(), days = 0) => {
  const base = isoDate ? new Date(`${isoDate}T00:00:00`) : new Date();
  if (Number.isNaN(base.getTime())) return isoToday();
  base.setDate(base.getDate() + Number(days || 0));
  return base.toISOString().slice(0, 10);
};

export const formatIsoDate = (isoDate) => {
  if (!isoDate) return "—";
  const parsed = new Date(`${isoDate}T00:00:00`);
  if (Number.isNaN(parsed.getTime())) return "—";
  return parsed.toLocaleDateString("en-NG");
};

export const looksLikeUuid = (value = "") =>
  /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(
    String(value || "").trim()
  );

const isLikelyMissingReceivablesSchema = (error) => {
  const msg = `${error?.message || ""} ${error?.details || ""} ${error?.hint || ""}`.toLowerCase();
  return (
    ["42p01", "pgrst205", "pgrst116"].includes(String(error?.code || "").toLowerCase()) ||
    ["receivables", "receivable_payments", "could not find", "relation", "schema cache", "column"].some(
      (token) => msg.includes(token)
    )
  );
};

export const getReceivablesSyncMeta = (mode = "local", error = null) => ({
  mode,
  message:
    mode === "supabase"
      ? "Native Supabase receivables sync is active with realtime updates."
      : isLikelyMissingReceivablesSchema(error)
        ? "Supabase receivables tables were not found, so the app is using the workspace local backup ledger."
        : "Supabase receivables sync is unavailable right now, so the app is using the workspace local backup ledger.",
  error: error?.message || "",
});

export const getDaysPastDue = (isoDate, balance = 0) => {
  if (!isoDate || (Number(balance) || 0) <= 0) return 0;
  const due = new Date(`${isoDate}T00:00:00`);
  const today = new Date(`${isoToday()}T00:00:00`);
  if (Number.isNaN(due.getTime()) || Number.isNaN(today.getTime())) return 0;
  const diff = Math.floor((today.getTime() - due.getTime()) / (1000 * 60 * 60 * 24));
  return Math.max(diff, 0);
};

export const normalizePaymentEntry = (payment = {}) => ({
  id: payment.id || uid(),
  amount: Number(payment.amount) || 0,
  receivedAt: payment.receivedAt || payment.date || isoToday(),
  reference: payment.reference || "",
  channel: payment.channel || "bank_transfer",
  note: payment.note || "",
  createdAt: payment.createdAt || new Date().toISOString(),
});

const deriveReceivableStatus = (record = {}) => {
  const requested = String(record.status || "").toLowerCase();
  const grossAmount = Number(record.grossAmount ?? record.amount ?? 0) || 0;
  const payments = Array.isArray(record.payments) ? record.payments.map(normalizePaymentEntry) : [];
  const amountReceived =
    record.amountReceived != null
      ? Number(record.amountReceived) || 0
      : payments.reduce((sum, payment) => sum + (Number(payment.amount) || 0), 0);

  const balance = Math.max(grossAmount - amountReceived, 0);

  if (requested === "write_off") return "write_off";
  if (requested === "disputed") return "disputed";
  if (requested === "draft") return "draft";
  if (balance <= 0 && grossAmount > 0) return "paid";
  if (amountReceived > 0) return getDaysPastDue(record.dueDate, balance) > 0 ? "overdue" : "part_paid";
  if (getDaysPastDue(record.dueDate, balance) > 0 && grossAmount > 0) return "overdue";
  return requested || (grossAmount > 0 ? "issued" : "draft");
};

export const normalizeReceivableRecord = (record = {}) => {
  const payments = Array.isArray(record.payments)
    ? record.payments
        .map(normalizePaymentEntry)
        .sort((a, b) => new Date(b.receivedAt || b.createdAt || 0) - new Date(a.receivedAt || a.createdAt || 0))
    : [];

  const grossAmount = Number(record.grossAmount ?? record.amount ?? 0) || 0;
  const amountReceived =
    record.amountReceived != null
      ? Number(record.amountReceived) || 0
      : payments.reduce((sum, payment) => sum + (Number(payment.amount) || 0), 0);

  const balance = Math.max(grossAmount - amountReceived, 0);
  const status = deriveReceivableStatus({ ...record, grossAmount, amountReceived, payments });

  return {
    id: record.id || uid(),
    mpoId: record.mpoId || "",
    clientId: record.clientId || "",
    campaignId: record.campaignId || "",
    invoiceNo: record.invoiceNo || "",
    invoiceDate: record.invoiceDate || isoToday(),
    dueDate: record.dueDate || addDaysToIso(record.invoiceDate || isoToday(), 30),
    grossAmount,
    amountReceived,
    balance,
    status,
    collectionStage: record.collectionStage || (status === "paid" ? "resolved" : "invoicing"),
    owner: record.owner || "",
    source: record.source || (record.mpoId ? "mpo" : "manual"),
    notes: record.notes || "",
    payments,
    lastPaymentAt: record.lastPaymentAt || payments[0]?.receivedAt || "",
    lastFollowUpAt: record.lastFollowUpAt || "",
    createdAt: record.createdAt || new Date().toISOString(),
    updatedAt: record.updatedAt || new Date().toISOString(),
  };
};

export const getStoredReceivables = (agencyId) => {
  if (!agencyId) return [];
  return (store.get(receivablesKeyForAgency(agencyId), []) || []).map(normalizeReceivableRecord);
};

export const setStoredReceivables = (agencyId, items = []) => {
  if (!agencyId) return;
  store.set(receivablesKeyForAgency(agencyId), (items || []).map(normalizeReceivableRecord));
};

const mapReceivablePaymentRowFromSupabase = (payment = {}) =>
  normalizePaymentEntry({
    id: payment.id,
    amount: payment.amount,
    receivedAt: payment.received_at,
    reference: payment.reference,
    channel: payment.channel,
    note: payment.note,
    createdAt: payment.created_at,
  });

const mapReceivableRowFromSupabase = (record = {}) =>
  normalizeReceivableRecord({
    id: record.id,
    mpoId: record.mpo_id,
    clientId: record.client_id,
    campaignId: record.campaign_id,
    invoiceNo: record.invoice_no,
    invoiceDate: record.invoice_date,
    dueDate: record.due_date,
    grossAmount: record.gross_amount,
    status: record.status,
    collectionStage: record.collection_stage,
    owner: record.owner,
    source: record.source,
    notes: record.notes,
    payments: Array.isArray(record.payments) ? record.payments.map(mapReceivablePaymentRowFromSupabase) : [],
    lastPaymentAt: record.last_payment_at,
    lastFollowUpAt: record.last_follow_up_at,
    createdAt: record.created_at,
    updatedAt: record.updated_at,
  });

const buildReceivablePayloadForSupabase = (record = {}, agencyId, userId, mode = "insert") => {
  const normalized = normalizeReceivableRecord(record);

  const payload = {
    agency_id: agencyId,
    mpo_id: normalized.mpoId || null,
    client_id: normalized.clientId || null,
    campaign_id: normalized.campaignId || null,
    invoice_no: normalized.invoiceNo || null,
    invoice_date: normalized.invoiceDate || isoToday(),
    due_date: normalized.dueDate || addDaysToIso(normalized.invoiceDate || isoToday(), 30),
    gross_amount: Number(normalized.grossAmount) || 0,
    status: normalized.status,
    collection_stage: normalized.collectionStage || (normalized.status === "paid" ? "resolved" : "invoicing"),
    owner: normalized.owner || "",
    source: normalized.source || (normalized.mpoId ? "mpo" : "manual"),
    notes: normalized.notes || "",
    last_payment_at: normalized.lastPaymentAt || null,
    last_follow_up_at: normalized.lastFollowUpAt || null,
    updated_at: new Date().toISOString(),
  };

  if (mode === "insert") {
    payload.created_by = userId || null;
    if (looksLikeUuid(normalized.id)) payload.id = normalized.id;
  }

  return Object.fromEntries(Object.entries(payload).filter(([, value]) => value !== undefined));
};

export const fetchReceivablesFromSupabase = async (agencyId) => {
  if (!agencyId) return [];
  const { data, error } = await supabase
    .from("receivables")
    .select(receivableListSelect)
    .eq("agency_id", agencyId)
    .order("updated_at", { ascending: false });

  if (error) throw error;
  return (data || []).map(mapReceivableRowFromSupabase);
};

export const fetchReceivableByIdFromSupabase = async (receivableId) => {
  const { data, error } = await supabase
    .from("receivables")
    .select(receivableDetailSelect)
    .eq("id", receivableId)
    .maybeSingle();

  if (error) throw error;
  return data ? mapReceivableRowFromSupabase(data) : null;
};

export const insertReceivableInSupabase = async (agencyId, userId, record) => {
  const payload = buildReceivablePayloadForSupabase(record, agencyId, userId, "insert");
  const { data, error } = await supabase
    .from("receivables")
    .insert([payload])
    .select(receivableDetailSelect)
    .single();

  if (error) throw error;
  return mapReceivableRowFromSupabase(data);
};

export const updateReceivableInSupabase = async (receivableId, record) => {
  const payload = buildReceivablePayloadForSupabase(record, undefined, undefined, "update");
  delete payload.agency_id;

  const { data, error } = await supabase
    .from("receivables")
    .update(payload)
    .eq("id", receivableId)
    .select(receivableDetailSelect)
    .single();

  if (error) throw error;
  return mapReceivableRowFromSupabase(data);
};

export const deleteReceivableInSupabase = async (receivableId) => {
  const { error } = await supabase.from("receivables").delete().eq("id", receivableId);
  if (error) throw error;
  return true;
};

export const insertReceivablePaymentInSupabase = async (
  agencyId,
  userId,
  receivableId,
  paymentInput = {},
  currentRecord = {}
) => {
  const payment = normalizePaymentEntry(paymentInput);

  const { error } = await supabase.from("receivable_payments").insert([
    {
      agency_id: agencyId,
      receivable_id: receivableId,
      amount: payment.amount,
      received_at: payment.receivedAt,
      reference: payment.reference || null,
      channel: payment.channel || "bank_transfer",
      note: payment.note || null,
      created_by: userId || null,
    },
  ]);

  if (error) throw error;

  const nextRecord = normalizeReceivableRecord({
    ...currentRecord,
    payments: [payment, ...((currentRecord?.payments || []).map(normalizePaymentEntry))],
    lastPaymentAt: payment.receivedAt,
    lastFollowUpAt: payment.receivedAt,
    updatedAt: new Date().toISOString(),
  });

  const { error: updateError } = await supabase
    .from("receivables")
    .update({
      status: nextRecord.status,
      collection_stage: nextRecord.status === "paid" ? "resolved" : nextRecord.collectionStage,
      last_payment_at: payment.receivedAt,
      last_follow_up_at: payment.receivedAt,
      updated_at: new Date().toISOString(),
    })
    .eq("id", receivableId);

  if (updateError) throw updateError;

  return await fetchReceivableByIdFromSupabase(receivableId);
};

export const updateReceivableStatusInSupabase = async (receivableId, updates = {}) => {
  const payload = {
    status: updates.status,
    collection_stage: updates.collectionStage,
    last_follow_up_at: updates.lastFollowUpAt ?? undefined,
    updated_at: new Date().toISOString(),
  };

  const { data, error } = await supabase
    .from("receivables")
    .update(Object.fromEntries(Object.entries(payload).filter(([, value]) => value !== undefined)))
    .eq("id", receivableId)
    .select(receivableDetailSelect)
    .single();

  if (error) throw error;
  return mapReceivableRowFromSupabase(data);
};

export const buildReceivableFromMpo = ({ mpo, campaign, client, owner = "" }) => {
  const grossAmount =
    Number(mpo?.reconciledAmount) ||
    Number(mpo?.invoiceAmount) ||
    Number(mpo?.grandTotal) ||
    Number(mpo?.netVal) ||
    0;

  const invoiceDate = isoToday();

  return normalizeReceivableRecord({
    id: uid(),
    mpoId: mpo?.id || "",
    clientId: client?.id || campaign?.clientId || "",
    campaignId: campaign?.id || mpo?.campaignId || "",
    invoiceNo: mpo?.invoiceNo || `AR-${String(mpo?.mpoNo || uid()).replace(/\s+/g, "-")}`,
    invoiceDate,
    dueDate: addDaysToIso(invoiceDate, 30),
    grossAmount,
    status: grossAmount > 0 ? "issued" : "draft",
    collectionStage: "invoicing",
    owner,
    source: "mpo",
    notes: mpo?.reconciliationNotes
      ? `Closeout note: ${mpo.reconciliationNotes}`
      : `Created from ${mpo?.mpoNo || "MPO"}`,
    payments: [],
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
  });
};