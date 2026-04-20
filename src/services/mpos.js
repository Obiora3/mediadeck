import { supabase } from "../lib/supabase";

const DEFAULT_MPO_TERMS = [
  "Make-goods or low quality reproductions are unacceptable.",
  "Transmission/broadcast of wrong material is totally unacceptable and could be considered as non-compliance. Kindly contact the agency if material is bad or damaged.",
  "Payment will be made within 30 days after submission of invoice and COT, with compliance based on the client's tracking or monitoring report.",
  "Any change in rate must be communicated within 90 days from the expiration of this contract period.",
  "Please note that on no account should there be any change to this flighting without approval from the agency.",
  "Please acknowledge receipt and implementation of this order.",
  "As agreed, 1 complimentary spot will be run for every 4 paid spots to support the campaign objective.",
];

const MPO_ATTACHMENTS_BUCKET = "mpo-attachments";
const sanitizeAttachmentFileName = (name = "file") => name.replace(/[^a-zA-Z0-9._-]+/g, "_");

const MPO_SPOT_META_PREFIX = "__MSP_META__";
const encodeMpoSpotMeta = (spot = {}) => {
  const payload = {
    wd: spot?.wd || "",
    bonusSpots: Number(spot?.bonusSpots) || 0,
    paidSpots: Number(spot?.paidSpots) || 0,
    isComplimentary: !!spot?.isComplimentary,
  };
  return `${MPO_SPOT_META_PREFIX}${JSON.stringify(payload)}`;
};
const decodeMpoSpotMeta = (rawWd = "") => {
  const value = String(rawWd || "");
  if (!value.startsWith(MPO_SPOT_META_PREFIX)) {
    return {
      wd: value,
      bonusSpots: 0,
      paidSpots: 0,
      isComplimentary: false,
    };
  }
  try {
    const parsed = JSON.parse(value.slice(MPO_SPOT_META_PREFIX.length));
    return {
      wd: parsed?.wd || "",
      bonusSpots: Number(parsed?.bonusSpots) || 0,
      paidSpots: Number(parsed?.paidSpots) || 0,
      isComplimentary: !!parsed?.isComplimentary,
    };
  } catch {
    return {
      wd: value,
      bonusSpots: 0,
      paidSpots: 0,
      isComplimentary: false,
    };
  }
};

const mpoParentSelect = `
  id,
  agency_id,
  campaign_id,
  vendor_id,
  mpo_no,
  issue_date,
  month,
  months,
  year,
  medium,
  signed_by,
  signed_title,
  prepared_by,
  prepared_contact,
  prepared_title,
  agency_address,
  transmit_msg,
  status,
  vendor_name,
  client_name,
  campaign_name,
  brand,
  total_spots,
  total_gross,
  disc_pct,
  disc_amt,
  less_disc,
  comm_pct,
  comm_amt,
  after_comm,
  surch_pct,
  surch_amt,
  surch_label,
  net_val,
  vat_pct,
  vat_amt,
  grand_total,
  terms,
  round_to_whole_naira,
  dispatch_status,
  dispatched_at,
  dispatched_by,
  dispatch_contact,
  dispatch_note,
  signed_mpo_url,
  invoice_status,
  invoice_no,
  invoice_amount,
  invoice_received_at,
  invoice_url,
  proof_status,
  proof_url,
  proof_received_at,
  planned_spots,
  aired_spots,
  missed_spots,
  makegood_spots,
  reconciliation_status,
  reconciliation_notes,
  reconciled_amount,
  payment_status,
  payment_reference,
  paid_at,
  is_archived,
  created_at,
  updated_at
`;

const mpoSpotSelect = `
  id,
  mpo_id,
  programme,
  wd,
  time_belt,
  material,
  duration,
  rate_id,
  rate_per_spot,
  spots,
  calendar_days,
  schedule_month,
  sort_order
`;

export const uploadMpoAttachmentAndGetUrl = async ({ agencyId, mpoId, kind, file }) => {
  if (!agencyId) throw new Error("No agency found for this workspace.");
  if (!mpoId) throw new Error("No MPO found for this upload.");
  if (!file) throw new Error("No file selected.");
  const safeName = sanitizeAttachmentFileName(file.name || "file");
  const path = `${agencyId}/${mpoId}/${kind}/${Date.now()}_${safeName}`;
  const { error } = await supabase.storage
    .from(MPO_ATTACHMENTS_BUCKET)
    .upload(path, file, { cacheControl: "3600", upsert: false });
  if (error) throw error;
  const { data } = supabase.storage.from(MPO_ATTACHMENTS_BUCKET).getPublicUrl(path);
  return data?.publicUrl || "";
};

const mapMpoSpotFromSupabase = (data) => {
  const meta = decodeMpoSpotMeta(data.wd || "");
  return ({
    id: data.id,
  agencyId: data.agency_id || "",
    programme: data.programme || "",
    wd: meta.wd || "",
    timeBelt: data.time_belt || "",
    material: data.material || "",
    duration: data.duration ? String(data.duration) : "30",
    rateId: data.rate_id || "",
    ratePerSpot: data.rate_per_spot ?? 0,
    spots: String(data.spots ?? 0),
    bonusSpots: meta.bonusSpots ?? 0,
    paidSpots: meta.paidSpots ?? 0,
    isComplimentary: !!meta.isComplimentary,
    calendarDays: Array.isArray(data.calendar_days) ? data.calendar_days : [],
    scheduleMonth: data.schedule_month || "",
  });
};

const mapMpoFromSupabase = (data, spots = []) => ({
  id: data.id,
  campaignId: data.campaign_id || "",
  vendorId: data.vendor_id || "",
  mpoNo: data.mpo_no || "",
  date: data.issue_date || "",
  month: data.month || "",
  months: Array.isArray(data.months) ? data.months : [],
  year: data.year || String(new Date().getFullYear()),
  medium: data.medium || "",
  signedBy: data.signed_by || "",
  signedTitle: data.signed_title || "",
  preparedBy: data.prepared_by || "",
  preparedContact: data.prepared_contact || "",
  preparedTitle: data.prepared_title || "",
  agencyAddress: data.agency_address || "",
  transmitMsg: data.transmit_msg || "",
  status: data.status || "draft",
  vendorName: data.vendor_name || "",
  clientName: data.client_name || "",
  campaignName: data.campaign_name || "",
  brand: data.brand || "",
  spots,
  totalSpots: data.total_spots ?? 0,
  totalGross: data.total_gross ?? 0,
  discPct: data.disc_pct ?? 0,
  discAmt: data.disc_amt ?? 0,
  lessDisc: data.less_disc ?? 0,
  commPct: data.comm_pct ?? 0,
  commAmt: data.comm_amt ?? 0,
  afterComm: data.after_comm ?? 0,
  surchPct: data.surch_pct ?? 0,
  surchAmt: data.surch_amt ?? 0,
  surchLabel: data.surch_label || "",
  netVal: data.net_val ?? 0,
  vatPct: data.vat_pct ?? 0,
  vatAmt: data.vat_amt ?? 0,
  grandTotal: data.grand_total ?? 0,
  terms: Array.isArray(data.terms) ? data.terms : DEFAULT_MPO_TERMS,
  roundToWholeNaira: !!data.round_to_whole_naira,
  dispatchStatus: data.dispatch_status || "pending",
  dispatchedAt: data.dispatched_at || null,
  dispatchedBy: data.dispatched_by || "",
  dispatchContact: data.dispatch_contact || "",
  dispatchNote: data.dispatch_note || "",
  signedMpoUrl: data.signed_mpo_url || "",
  invoiceStatus: data.invoice_status || "pending",
  invoiceNo: data.invoice_no || "",
  invoiceAmount: data.invoice_amount ?? 0,
  invoiceReceivedAt: data.invoice_received_at || null,
  invoiceUrl: data.invoice_url || "",
  proofStatus: data.proof_status || "pending",
  proofUrl: data.proof_url || "",
  proofReceivedAt: data.proof_received_at || null,
  plannedSpotsExecution: data.planned_spots ?? (data.total_spots ?? 0),
  airedSpots: data.aired_spots ?? 0,
  missedSpots: data.missed_spots ?? 0,
  makegoodSpots: data.makegood_spots ?? 0,
  reconciliationStatus: data.reconciliation_status || "not_started",
  reconciliationNotes: data.reconciliation_notes || "",
  reconciledAmount: data.reconciled_amount ?? 0,
  paymentStatus: data.payment_status || "unpaid",
  paymentReference: data.payment_reference || "",
  paidAt: data.paid_at || null,
  archivedAt: data.is_archived ? (data.updated_at ? new Date(data.updated_at).getTime() : Date.now()) : null,
  createdAt: data.created_at ? new Date(data.created_at).getTime() : Date.now(),
  updatedAt: data.updated_at ? new Date(data.updated_at).getTime() : Date.now(),
});

const mapMpoSpotToSupabase = (spot, index = 0) => ({
  programme: spot.programme || null,
  wd: encodeMpoSpotMeta(spot),
  time_belt: spot.timeBelt || null,
  material: spot.material || null,
  duration: spot.duration ? String(spot.duration) : "30",
  rate_id: spot.rateId || null,
  rate_per_spot: spot.ratePerSpot ? Number(spot.ratePerSpot) : 0,
  spots: spot.spots ? Number(spot.spots) : 0,
  calendar_days: Array.isArray(spot.calendarDays) ? spot.calendarDays.map((d) => Number(d)) : [],
  schedule_month: spot.scheduleMonth || null,
  sort_order: index,
});

const MPO_SPOTS_PAGE_SIZE = 1000;
const MPO_PARENTS_PAGE_SIZE = 200;

const fetchMpoParentsForAgency = async (agencyId) => {
  if (!agencyId) return [];

  const allRows = [];
  let from = 0;

  while (true) {
    const to = from + MPO_PARENTS_PAGE_SIZE - 1;
    const { data, error } = await supabase
      .from("mpos")
      .select(mpoParentSelect)
      .eq("agency_id", agencyId)
      .order("created_at", { ascending: false })
      .range(from, to);

    if (error) throw error;

    const pageRows = data || [];
    allRows.push(...pageRows);

    if (pageRows.length < MPO_PARENTS_PAGE_SIZE) break;
    from += MPO_PARENTS_PAGE_SIZE;
  }

  return allRows;
};

const fetchMpoSpotsForIds = async (mpoIds = []) => {
  if (!mpoIds.length) return [];

  const allRows = [];
  const chunkSize = 100;

  for (let chunkStart = 0; chunkStart < mpoIds.length; chunkStart += chunkSize) {
    const chunkIds = mpoIds.slice(chunkStart, chunkStart + chunkSize);
    let from = 0;

    while (true) {
      const to = from + MPO_SPOTS_PAGE_SIZE - 1;
      const { data, error } = await supabase
        .from("mpo_spots")
        .select(mpoSpotSelect)
        .in("mpo_id", chunkIds)
        .order("mpo_id", { ascending: true })
        .order("sort_order", { ascending: true })
        .range(from, to);

      if (error) throw error;

      const pageRows = data || [];
      allRows.push(...pageRows);

      if (pageRows.length < MPO_SPOTS_PAGE_SIZE) break;
      from += MPO_SPOTS_PAGE_SIZE;
    }
  }

  return allRows;
};

export const fetchMappedMpoById = async (mpoId) => {
  const { data: parent, error: parentError } = await supabase
    .from("mpos")
    .select(mpoParentSelect)
    .eq("id", mpoId)
    .single();
  if (parentError) throw parentError;

  const { data: spotRows, error: spotError } = await supabase
    .from("mpo_spots")
    .select(mpoSpotSelect)
    .eq("mpo_id", mpoId)
    .order("sort_order", { ascending: true });
  if (spotError) throw spotError;

  return mapMpoFromSupabase(parent, (spotRows || []).map(mapMpoSpotFromSupabase));
};

export const fetchMappedMpoByAgencyAndNo = async (agencyId, mpoNo) => {
  if (!agencyId || !mpoNo) return null;

  const { data: parent, error: parentError } = await supabase
    .from("mpos")
    .select(mpoParentSelect)
    .eq("agency_id", agencyId)
    .eq("mpo_no", mpoNo)
    .maybeSingle();

  if (parentError) throw parentError;
  if (!parent) return null;
  return mapMpoFromSupabase(parent, []);
};

const deleteMpoAttachmentsFromSupabase = async (agencyId, mpoId) => {
  if (!agencyId || !mpoId) return;

  const attachmentKinds = ["signedMpo", "invoice", "proof"];

  await Promise.all(
    attachmentKinds.map(async (kind) => {
      const folder = `${agencyId}/${mpoId}/${kind}`;
      const { data: files, error: listError } = await supabase.storage
        .from(MPO_ATTACHMENTS_BUCKET)
        .list(folder, { limit: 1000 });

      if (listError) return;

      const filePaths = (files || [])
        .filter((item) => item && item.name && !item.id?.endsWith?.("/"))
        .map((item) => `${folder}/${item.name}`);

      if (!filePaths.length) return;

      await supabase.storage.from(MPO_ATTACHMENTS_BUCKET).remove(filePaths);
    })
  );
};

export const fetchMposFromSupabase = async (agencyId) => {
  if (!agencyId) return [];

  const mpoRows = await fetchMpoParentsForAgency(agencyId);
  const ids = (mpoRows || []).map((row) => row.id);
  const spotsRows = await fetchMpoSpotsForIds(ids);

  const spotsByMpo = spotsRows.reduce((acc, row) => {
    (acc[row.mpo_id] ||= []).push(mapMpoSpotFromSupabase(row));
    return acc;
  }, {});

  return (mpoRows || []).map((row) => mapMpoFromSupabase(row, spotsByMpo[row.id] || []));
};

export const createMpoInSupabase = async (agencyId, userId, record) => {
  if (!agencyId) throw new Error("No agency found for this user.");
  const parentPayload = {
    agency_id: agencyId,
    created_by: userId,
    campaign_id: record.campaignId || null,
    vendor_id: record.vendorId || null,
    mpo_no: record.mpoNo,
    issue_date: record.date || null,
    month: record.month || null,
    months: Array.isArray(record.months) ? record.months : [],
    year: record.year || null,
    medium: record.medium || null,
    signed_by: record.signedBy || null,
    signed_title: record.signedTitle || null,
    prepared_by: record.preparedBy || null,
    prepared_contact: record.preparedContact || null,
    prepared_title: record.preparedTitle || null,
    agency_address: record.agencyAddress || null,
    transmit_msg: record.transmitMsg || null,
    status: record.status || "draft",
    vendor_name: record.vendorName || null,
    client_name: record.clientName || null,
    campaign_name: record.campaignName || null,
    brand: record.brand || null,
    total_spots: Number(record.totalSpots) || 0,
    total_gross: Number(record.totalGross) || 0,
    disc_pct: Number(record.discPct) || 0,
    disc_amt: Number(record.discAmt) || 0,
    less_disc: Number(record.lessDisc) || 0,
    comm_pct: Number(record.commPct) || 0,
    comm_amt: Number(record.commAmt) || 0,
    after_comm: Number(record.afterComm) || 0,
    surch_pct: Number(record.surchPct) || 0,
    surch_amt: Number(record.surchAmt) || 0,
    surch_label: record.surchLabel || null,
    net_val: Number(record.netVal) || 0,
    vat_pct: Number(record.vatPct) || 0,
    vat_amt: Number(record.vatAmt) || 0,
    grand_total: Number(record.grandTotal) || 0,
    terms: Array.isArray(record.terms) ? record.terms : DEFAULT_MPO_TERMS,
    round_to_whole_naira: !!record.roundToWholeNaira,
    dispatch_status: record.dispatchStatus || "pending",
    dispatched_at: record.dispatchedAt || null,
    dispatched_by: record.dispatchedBy || null,
    dispatch_contact: record.dispatchContact || null,
    dispatch_note: record.dispatchNote || null,
    signed_mpo_url: record.signedMpoUrl || null,
    invoice_status: record.invoiceStatus || "pending",
    invoice_no: record.invoiceNo || null,
    invoice_amount: Number(record.invoiceAmount) || 0,
    invoice_received_at: record.invoiceReceivedAt || null,
    invoice_url: record.invoiceUrl || null,
    proof_status: record.proofStatus || "pending",
    proof_url: record.proofUrl || null,
    proof_received_at: record.proofReceivedAt || null,
    planned_spots: Number(record.plannedSpotsExecution ?? record.totalSpots) || 0,
    aired_spots: Number(record.airedSpots) || 0,
    missed_spots: Number(record.missedSpots) || 0,
    makegood_spots: Number(record.makegoodSpots) || 0,
    reconciliation_status: record.reconciliationStatus || "not_started",
    reconciliation_notes: record.reconciliationNotes || null,
    reconciled_amount: Number(record.reconciledAmount) || 0,
    payment_status: record.paymentStatus || "unpaid",
    payment_reference: record.paymentReference || null,
    paid_at: record.paidAt || null,
    is_archived: false,
  };

  const { data: parent, error: parentError } = await supabase.from("mpos").insert([parentPayload]).select(mpoParentSelect).single();
  if (parentError) throw parentError;

  let mappedSpots = [];
  if ((record.spots || []).length) {
    const payload = record.spots.map((spot, index) => ({ mpo_id: parent.id, ...mapMpoSpotToSupabase(spot, index) }));
    const { data: insertedSpots, error: spotsError } = await supabase.from("mpo_spots").insert(payload).select(mpoSpotSelect).order("sort_order", { ascending: true });
    if (spotsError) throw spotsError;
    mappedSpots = (insertedSpots || []).map(mapMpoSpotFromSupabase);
  }

  return mapMpoFromSupabase(parent, mappedSpots);
};

export const updateMpoInSupabase = async (mpoId, record) => {
  const parentPayload = {
    campaign_id: record.campaignId || null,
    vendor_id: record.vendorId || null,
    mpo_no: record.mpoNo,
    issue_date: record.date || null,
    month: record.month || null,
    months: Array.isArray(record.months) ? record.months : [],
    year: record.year || null,
    medium: record.medium || null,
    signed_by: record.signedBy || null,
    signed_title: record.signedTitle || null,
    prepared_by: record.preparedBy || null,
    prepared_contact: record.preparedContact || null,
    prepared_title: record.preparedTitle || null,
    agency_address: record.agencyAddress || null,
    transmit_msg: record.transmitMsg || null,
    status: record.status || "draft",
    vendor_name: record.vendorName || null,
    client_name: record.clientName || null,
    campaign_name: record.campaignName || null,
    brand: record.brand || null,
    total_spots: Number(record.totalSpots) || 0,
    total_gross: Number(record.totalGross) || 0,
    disc_pct: Number(record.discPct) || 0,
    disc_amt: Number(record.discAmt) || 0,
    less_disc: Number(record.lessDisc) || 0,
    comm_pct: Number(record.commPct) || 0,
    comm_amt: Number(record.commAmt) || 0,
    after_comm: Number(record.afterComm) || 0,
    surch_pct: Number(record.surchPct) || 0,
    surch_amt: Number(record.surchAmt) || 0,
    surch_label: record.surchLabel || null,
    net_val: Number(record.netVal) || 0,
    vat_pct: Number(record.vatPct) || 0,
    vat_amt: Number(record.vatAmt) || 0,
    grand_total: Number(record.grandTotal) || 0,
    terms: Array.isArray(record.terms) ? record.terms : DEFAULT_MPO_TERMS,
    round_to_whole_naira: !!record.roundToWholeNaira,
    dispatch_status: record.dispatchStatus || "pending",
    dispatched_at: record.dispatchedAt || null,
    dispatched_by: record.dispatchedBy || null,
    dispatch_contact: record.dispatchContact || null,
    dispatch_note: record.dispatchNote || null,
    signed_mpo_url: record.signedMpoUrl || null,
    invoice_status: record.invoiceStatus || "pending",
    invoice_no: record.invoiceNo || null,
    invoice_amount: Number(record.invoiceAmount) || 0,
    invoice_received_at: record.invoiceReceivedAt || null,
    invoice_url: record.invoiceUrl || null,
    proof_status: record.proofStatus || "pending",
    proof_url: record.proofUrl || null,
    proof_received_at: record.proofReceivedAt || null,
    planned_spots: Number(record.plannedSpotsExecution ?? record.totalSpots) || 0,
    aired_spots: Number(record.airedSpots) || 0,
    missed_spots: Number(record.missedSpots) || 0,
    makegood_spots: Number(record.makegoodSpots) || 0,
    reconciliation_status: record.reconciliationStatus || "not_started",
    reconciliation_notes: record.reconciliationNotes || null,
    reconciled_amount: Number(record.reconciledAmount) || 0,
    payment_status: record.paymentStatus || "unpaid",
    payment_reference: record.paymentReference || null,
    paid_at: record.paidAt || null,
  };

  const { data: parent, error: parentError } = await supabase.from("mpos").update(parentPayload).eq("id", mpoId).select(mpoParentSelect).single();
  if (parentError) throw parentError;
  const { error: deleteError } = await supabase.from("mpo_spots").delete().eq("mpo_id", mpoId);
  if (deleteError) throw deleteError;

  let mappedSpots = [];
  if ((record.spots || []).length) {
    const payload = record.spots.map((spot, index) => ({ mpo_id: mpoId, ...mapMpoSpotToSupabase(spot, index) }));
    const { data: insertedSpots, error: spotsError } = await supabase.from("mpo_spots").insert(payload).select(mpoSpotSelect).order("sort_order", { ascending: true });
    if (spotsError) throw spotsError;
    mappedSpots = (insertedSpots || []).map(mapMpoSpotFromSupabase);
  }
  return mapMpoFromSupabase(parent, mappedSpots);
};

export const archiveMpoInSupabase = async (mpoId) => {
  const { data, error } = await supabase.from("mpos").update({ is_archived: true }).eq("id", mpoId).select(mpoParentSelect).single();
  if (error) throw error;
  const { data: spotRows, error: spotError } = await supabase.from("mpo_spots").select(mpoSpotSelect).eq("mpo_id", mpoId).order("sort_order", { ascending: true });
  if (spotError) throw spotError;
  return mapMpoFromSupabase(data, (spotRows || []).map(mapMpoSpotFromSupabase));
};

export const restoreMpoInSupabase = async (mpoId) => {
  const { data, error } = await supabase.from("mpos").update({ is_archived: false }).eq("id", mpoId).select(mpoParentSelect).single();
  if (error) throw error;
  const { data: spotRows, error: spotError } = await supabase.from("mpo_spots").select(mpoSpotSelect).eq("mpo_id", mpoId).order("sort_order", { ascending: true });
  if (spotError) throw spotError;
  return mapMpoFromSupabase(data, (spotRows || []).map(mapMpoSpotFromSupabase));
};

export const updateMpoStatusInSupabase = async (mpoId, status) => {
  const { data, error } = await supabase.from("mpos").update({ status }).eq("id", mpoId).select(mpoParentSelect).single();
  if (error) throw error;
  const { data: spotRows, error: spotError } = await supabase.from("mpo_spots").select(mpoSpotSelect).eq("mpo_id", mpoId).order("sort_order", { ascending: true });
  if (spotError) throw spotError;
  return mapMpoFromSupabase(data, (spotRows || []).map(mapMpoSpotFromSupabase));
};

export const updateMpoExecutionInSupabase = async (mpoId, patch) => {
  const payload = {
    dispatch_status: patch.dispatchStatus || "pending",
    dispatched_at: patch.dispatchedAt || null,
    dispatched_by: patch.dispatchedBy || null,
    dispatch_contact: patch.dispatchContact || null,
    dispatch_note: patch.dispatchNote || null,
    signed_mpo_url: patch.signedMpoUrl || null,
    invoice_status: patch.invoiceStatus || "pending",
    invoice_no: patch.invoiceNo || null,
    invoice_amount: Number(patch.invoiceAmount) || 0,
    invoice_received_at: patch.invoiceReceivedAt || null,
    invoice_url: patch.invoiceUrl || null,
    proof_status: patch.proofStatus || "pending",
    proof_url: patch.proofUrl || null,
    proof_received_at: patch.proofReceivedAt || null,
    planned_spots: Number(patch.plannedSpotsExecution) || 0,
    aired_spots: Number(patch.airedSpots) || 0,
    missed_spots: Number(patch.missedSpots) || 0,
    makegood_spots: Number(patch.makegoodSpots) || 0,
    reconciliation_status: patch.reconciliationStatus || "not_started",
    reconciliation_notes: patch.reconciliationNotes || null,
    reconciled_amount: Number(patch.reconciledAmount) || 0,
    payment_status: patch.paymentStatus || "unpaid",
    payment_reference: patch.paymentReference || null,
    paid_at: patch.paidAt || null,
  };
  const { data, error } = await supabase.from("mpos").update(payload).eq("id", mpoId).select(mpoParentSelect).single();
  if (error) throw error;
  const { data: spotRows, error: spotError } = await supabase.from("mpo_spots").select(mpoSpotSelect).eq("mpo_id", mpoId).order("sort_order", { ascending: true });
  if (spotError) throw spotError;
  return mapMpoFromSupabase(data, (spotRows || []).map(mapMpoSpotFromSupabase));
};

export const generateNextMpoNoFromSupabase = async (brand = "MPO") => {
  const { data, error } = await supabase.rpc("generate_next_mpo_no", { p_brand: brand || "MPO" });
  if (error) {
    if ((error.message || "").toLowerCase().includes("generate_next_mpo_no")) {
      throw new Error('Missing database function "generate_next_mpo_no". Run the provided SQL patch in Supabase SQL Editor.');
    }
    throw error;
  }
  return data;
};


export const deleteMpoInSupabase = async (mpoId) => {
  if (!mpoId) throw new Error("No MPO selected for deletion.");

  const existing = await fetchMappedMpoById(mpoId);

  const { error: spotDeleteError } = await supabase
    .from("mpo_spots")
    .delete()
    .eq("mpo_id", mpoId);
  if (spotDeleteError) throw spotDeleteError;

  const { error: parentDeleteError } = await supabase
    .from("mpos")
    .delete()
    .eq("id", mpoId);
  if (parentDeleteError) throw parentDeleteError;

  await deleteMpoAttachmentsFromSupabase(existing.agencyId || existing.agency_id || "", mpoId);

  return existing;
};
