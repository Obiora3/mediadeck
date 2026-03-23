import { useState, useEffect, useCallback, useRef } from "react";
import { supabase } from "./lib/supabase";

/* ── STORAGE ────────────────────────────────────────────── */
const store = {
  get: (k, d = null) => { try { const v = localStorage.getItem(k); return v ? JSON.parse(v) : d; } catch { return d; } },
  set: (k, v) => { try { localStorage.setItem(k, JSON.stringify(v)); } catch {} },
  del: (k) => { try { localStorage.removeItem(k); } catch {} },
};
const uid = () => Math.random().toString(36).slice(2) + Date.now().toString(36);
const fmt = (n, d = 0) => (parseFloat(n) || 0).toLocaleString("en-NG", { minimumFractionDigits: d, maximumFractionDigits: d });
const fmtN = (n) => `₦${fmt(n)}`;
const normalizeAgencyName = (value = "") => value.trim().replace(/\s+/g, " ").toLowerCase();
const normalizeAgencyCode = (value = "") => value.trim().toUpperCase().replace(/\s+/g, "");
const MPO_ATTACHMENTS_BUCKET = "mpo-attachments";
const sanitizeAttachmentFileName = (name = "file") => name.replace(/[^a-zA-Z0-9._-]+/g, "_");
const uploadMpoAttachmentAndGetUrl = async ({ agencyId, mpoId, kind, file }) => {
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

const loadAppUserFromSupabase = async (authUser) => {
  if (!authUser) return null;

  const { data: profile } = await supabase
    .from("profiles")
    .select("*")
    .eq("id", authUser.id)
    .single();

  let agencyName = "";
  let agencyCode = "";
  let agencyAddress = "";
  let agencyEmail = "";
  let agencyPhone = "";
  let agencyId = profile?.agency_id || null;
  if (profile?.agency_id) {
    let agency = null;
    const primary = await supabase
      .from("agencies")
      .select("name, agency_code, address, email, phone")
      .eq("id", profile.agency_id)
      .maybeSingle();

    if (!primary.error) {
      agency = primary.data;
    } else {
      const fallback = await supabase
        .from("agencies")
        .select("name, agency_code, address")
        .eq("id", profile.agency_id)
        .maybeSingle();
      agency = fallback.data || null;
    }

    agencyName = agency?.name || "";
    agencyCode = agency?.agency_code || "";
    agencyAddress = agency?.address || "";
    agencyEmail = agency?.email || "";
    agencyPhone = agency?.phone || "";
  }

  const storedSignature = getStoredUserSignature(authUser.id);
  const storedAgencyContact = getStoredAgencyContact(agencyId);

  return {
    id: authUser.id,
    name:
      profile?.full_name ||
      authUser.user_metadata?.full_name ||
      authUser.email ||
      "User",
    email: authUser.email || "",
    title: profile?.title || authUser.user_metadata?.title || "",
    phone: profile?.phone || authUser.user_metadata?.phone || "",
    agency: agencyName || authUser.user_metadata?.agency_name || "My Agency",
    agencyId,
    agencyCode,
    agencyAddress,
    agencyEmail: agencyEmail || authUser.user_metadata?.agency_email || storedAgencyContact.email || "",
    agencyPhone: agencyPhone || authUser.user_metadata?.agency_phone || storedAgencyContact.phone || "",
    signatureDataUrl: authUser.user_metadata?.signature_data_url || storedSignature || "",
    role: normalizeRole(profile?.role || "admin"),
    roleLabel: formatRoleLabel(profile?.role || "admin"),
  };
};

const ensureAgencyForUser = async (authUser) => {
  if (!authUser) return null;

  const { data: profile, error: profileError } = await supabase
    .from("profiles")
    .select("agency_id")
    .eq("id", authUser.id)
    .single();

  if (profileError) throw profileError;
  if (profile?.agency_id) return profile.agency_id;

  const joinCode = normalizeAgencyCode(authUser.user_metadata?.agency_code || "");
  if (joinCode) {
    const { data, error } = await supabase.rpc("join_my_agency_by_code", {
      p_code: joinCode,
    });

    if (error) {
      if ((error.message || "").toLowerCase().includes("join_my_agency_by_code")) {
        throw new Error('Missing database function "join_my_agency_by_code". Run the provided SQL patch in Supabase SQL Editor.');
      }
      throw error;
    }

    return data;
  }

  const agencyName = (authUser.user_metadata?.agency_name || "").trim();
  if (!agencyName) throw new Error("No agency invite code or agency name found for this user.");

  const { data, error } = await supabase.rpc("create_my_agency_with_code", {
    p_name: agencyName,
  });

  if (error) {
    if ((error.message || "").toLowerCase().includes("create_my_agency_with_code")) {
      throw new Error('Missing database function "create_my_agency_with_code". Run the provided SQL patch in Supabase SQL Editor.');
    }
    throw error;
  }

  if (Array.isArray(data)) return data[0]?.agency_id || null;
  if (data && typeof data === "object" && "agency_id" in data) return data.agency_id;
  return data;
};

const fetchVendorsFromSupabase = async () => {
  const { data, error } = await supabase
    .from("vendors")
    .select("*")
    .order("created_at", { ascending: false });

  if (error) throw error;

  return (data || []).map(v => ({
    id: v.id,
    name: v.name || "",
    type: v.type || "",
    contact: v.contact || "",
    email: v.email || "",
    phone: v.phone || "",
    location: v.location || "",
    rate: v.default_rate ?? "",
    discount: v.discount ?? "",
    commission: v.commission ?? "",
    notes: v.notes || "",
    archivedAt: v.is_archived ? (v.updated_at ? new Date(v.updated_at).getTime() : Date.now()) : null,
    createdAt: v.created_at ? new Date(v.created_at).getTime() : Date.now(),
  }));
};

const createVendorInSupabase = async (agencyId, userId, form) => {
  if (!agencyId) throw new Error("No agency found for this user.");

  const { data, error } = await supabase
    .from("vendors")
    .insert([{ agency_id: agencyId, created_by: userId, name: form.name, type: form.type, contact: form.contact || null, email: form.email || null, phone: form.phone || null, location: form.location || null, default_rate: form.rate ? Number(form.rate) : null, discount: form.discount ? Number(form.discount) : 0, commission: form.commission ? Number(form.commission) : 0, notes: form.notes || null, is_archived: false }])
    .select()
    .single();

  if (error) throw error;

  return { id: data.id, name: data.name || "", type: data.type || "", contact: data.contact || "", email: data.email || "", phone: data.phone || "", location: data.location || "", rate: data.default_rate ?? "", discount: data.discount ?? "", commission: data.commission ?? "", notes: data.notes || "", archivedAt: data.is_archived ? (data.updated_at ? new Date(data.updated_at).getTime() : Date.now()) : null, createdAt: data.created_at ? new Date(data.created_at).getTime() : Date.now() };
};

const updateVendorInSupabase = async (vendorId, form) => {
  const { data, error } = await supabase
    .from("vendors")
    .update({ name: form.name, type: form.type, contact: form.contact || null, email: form.email || null, phone: form.phone || null, location: form.location || null, default_rate: form.rate ? Number(form.rate) : null, discount: form.discount ? Number(form.discount) : 0, commission: form.commission ? Number(form.commission) : 0, notes: form.notes || null })
    .eq("id", vendorId)
    .select()
    .single();

  if (error) throw error;

  return { id: data.id, name: data.name || "", type: data.type || "", contact: data.contact || "", email: data.email || "", phone: data.phone || "", location: data.location || "", rate: data.default_rate ?? "", discount: data.discount ?? "", commission: data.commission ?? "", notes: data.notes || "", archivedAt: data.is_archived ? (data.updated_at ? new Date(data.updated_at).getTime() : Date.now()) : null, createdAt: data.created_at ? new Date(data.created_at).getTime() : Date.now() };
};

const archiveVendorInSupabase = async (vendorId) => {
  const { data, error } = await supabase
    .from("vendors")
    .update({ is_archived: true })
    .eq("id", vendorId)
    .select()
    .single();
  if (error) throw error;
  return { id: data.id, name: data.name || "", type: data.type || "", contact: data.contact || "", email: data.email || "", phone: data.phone || "", location: data.location || "", rate: data.default_rate ?? "", discount: data.discount ?? "", commission: data.commission ?? "", notes: data.notes || "", archivedAt: data.updated_at ? new Date(data.updated_at).getTime() : Date.now(), createdAt: data.created_at ? new Date(data.created_at).getTime() : Date.now() };
};

const restoreVendorInSupabase = async (vendorId) => {
  const { data, error } = await supabase
    .from("vendors")
    .update({ is_archived: false })
    .eq("id", vendorId)
    .select()
    .single();
  if (error) throw error;
  return { id: data.id, name: data.name || "", type: data.type || "", contact: data.contact || "", email: data.email || "", phone: data.phone || "", location: data.location || "", rate: data.default_rate ?? "", discount: data.discount ?? "", commission: data.commission ?? "", notes: data.notes || "", archivedAt: null, createdAt: data.created_at ? new Date(data.created_at).getTime() : Date.now() };
};

const fetchClientsFromSupabase = async () => {
  const { data, error } = await supabase
    .from("clients")
    .select("*")
    .order("created_at", { ascending: false });
  if (error) throw error;
  return (data || []).map(c => ({ id: c.id, name: c.name || "", industry: c.industry || "", contact: c.contact || "", email: c.email || "", phone: c.phone || "", address: c.address || "", brands: c.brands || "", archivedAt: c.is_archived ? (c.updated_at ? new Date(c.updated_at).getTime() : Date.now()) : null, createdAt: c.created_at ? new Date(c.created_at).getTime() : Date.now(), updatedAt: c.updated_at ? new Date(c.updated_at).getTime() : Date.now() }));
};

const createClientInSupabase = async (agencyId, userId, form) => {
  if (!agencyId) throw new Error("No agency found for this user.");
  const { data, error } = await supabase
    .from("clients")
    .insert([{ agency_id: agencyId, created_by: userId, name: form.name.trim(), industry: form.industry || null, contact: form.contact || null, email: form.email || null, phone: form.phone || null, address: form.address || null, brands: form.brands || null, is_archived: false }])
    .select()
    .single();
  if (error) throw error;
  return { id: data.id, name: data.name || "", industry: data.industry || "", contact: data.contact || "", email: data.email || "", phone: data.phone || "", address: data.address || "", brands: data.brands || "", archivedAt: null, createdAt: data.created_at ? new Date(data.created_at).getTime() : Date.now(), updatedAt: data.updated_at ? new Date(data.updated_at).getTime() : Date.now() };
};

const updateClientInSupabase = async (clientId, form) => {
  const { data, error } = await supabase
    .from("clients")
    .update({ name: form.name.trim(), industry: form.industry || null, contact: form.contact || null, email: form.email || null, phone: form.phone || null, address: form.address || null, brands: form.brands || null })
    .eq("id", clientId)
    .select()
    .single();
  if (error) throw error;
  return { id: data.id, name: data.name || "", industry: data.industry || "", contact: data.contact || "", email: data.email || "", phone: data.phone || "", address: data.address || "", brands: data.brands || "", archivedAt: data.is_archived ? (data.updated_at ? new Date(data.updated_at).getTime() : Date.now()) : null, createdAt: data.created_at ? new Date(data.created_at).getTime() : Date.now(), updatedAt: data.updated_at ? new Date(data.updated_at).getTime() : Date.now() };
};

const archiveClientInSupabase = async (clientId) => {
  const { data, error } = await supabase
    .from("clients")
    .update({ is_archived: true })
    .eq("id", clientId)
    .select()
    .single();
  if (error) throw error;
  return { id: data.id, name: data.name || "", industry: data.industry || "", contact: data.contact || "", email: data.email || "", phone: data.phone || "", address: data.address || "", brands: data.brands || "", archivedAt: data.updated_at ? new Date(data.updated_at).getTime() : Date.now(), createdAt: data.created_at ? new Date(data.created_at).getTime() : Date.now(), updatedAt: data.updated_at ? new Date(data.updated_at).getTime() : Date.now() };
};

const restoreClientInSupabase = async (clientId) => {
  const { data, error } = await supabase
    .from("clients")
    .update({ is_archived: false })
    .eq("id", clientId)
    .select()
    .single();
  if (error) throw error;
  return { id: data.id, name: data.name || "", industry: data.industry || "", contact: data.contact || "", email: data.email || "", phone: data.phone || "", address: data.address || "", brands: data.brands || "", archivedAt: null, createdAt: data.created_at ? new Date(data.created_at).getTime() : Date.now(), updatedAt: data.updated_at ? new Date(data.updated_at).getTime() : Date.now() };
};

const fetchCampaignsFromSupabase = async () => {
  const { data, error } = await supabase
    .from("campaigns")
    .select("*")
    .order("created_at", { ascending: false });
  if (error) throw error;
  return (data || []).map(c => ({
    id: c.id,
    name: c.name || "",
    clientId: c.client_id || "",
    brand: c.brand || "",
    objective: c.objective || "",
    startDate: c.start_date || "",
    endDate: c.end_date || "",
    budget: c.budget ?? "",
    status: c.status || "planning",
    medium: c.medium || "",
    notes: c.notes || "",
    materialList: Array.isArray(c.material_list) ? c.material_list : [],
    archivedAt: c.is_archived ? (c.updated_at ? new Date(c.updated_at).getTime() : Date.now()) : null,
    createdAt: c.created_at ? new Date(c.created_at).getTime() : Date.now(),
    updatedAt: c.updated_at ? new Date(c.updated_at).getTime() : Date.now(),
  }));
};

const createCampaignInSupabase = async (agencyId, userId, form) => {
  if (!agencyId) throw new Error("No agency found for this user.");
  const { data, error } = await supabase
    .from("campaigns")
    .insert([{
      agency_id: agencyId,
      client_id: form.clientId,
      created_by: userId,
      name: form.name.trim(),
      brand: form.brand || null,
      objective: form.objective || null,
      start_date: form.startDate || null,
      end_date: form.endDate || null,
      budget: form.budget ? Number(form.budget) : 0,
      status: form.status || "planning",
      medium: form.medium || null,
      notes: form.notes || null,
      material_list: (form.materialList || []).map(x => x.trim()).filter(Boolean),
      is_archived: false,
    }])
    .select()
    .single();
  if (error) throw error;
  return {
    id: data.id,
    name: data.name || "",
    clientId: data.client_id || "",
    brand: data.brand || "",
    objective: data.objective || "",
    startDate: data.start_date || "",
    endDate: data.end_date || "",
    budget: data.budget ?? "",
    status: data.status || "planning",
    medium: data.medium || "",
    notes: data.notes || "",
    materialList: Array.isArray(data.material_list) ? data.material_list : [],
    archivedAt: null,
    createdAt: data.created_at ? new Date(data.created_at).getTime() : Date.now(),
    updatedAt: data.updated_at ? new Date(data.updated_at).getTime() : Date.now(),
  };
};

const updateCampaignInSupabase = async (campaignId, form) => {
  const { data, error } = await supabase
    .from("campaigns")
    .update({
      client_id: form.clientId,
      name: form.name.trim(),
      brand: form.brand || null,
      objective: form.objective || null,
      start_date: form.startDate || null,
      end_date: form.endDate || null,
      budget: form.budget ? Number(form.budget) : 0,
      status: form.status || "planning",
      medium: form.medium || null,
      notes: form.notes || null,
      material_list: (form.materialList || []).map(x => x.trim()).filter(Boolean),
    })
    .eq("id", campaignId)
    .select()
    .single();
  if (error) throw error;
  return {
    id: data.id,
    name: data.name || "",
    clientId: data.client_id || "",
    brand: data.brand || "",
    objective: data.objective || "",
    startDate: data.start_date || "",
    endDate: data.end_date || "",
    budget: data.budget ?? "",
    status: data.status || "planning",
    medium: data.medium || "",
    notes: data.notes || "",
    materialList: Array.isArray(data.material_list) ? data.material_list : [],
    archivedAt: data.is_archived ? (data.updated_at ? new Date(data.updated_at).getTime() : Date.now()) : null,
    createdAt: data.created_at ? new Date(data.created_at).getTime() : Date.now(),
    updatedAt: data.updated_at ? new Date(data.updated_at).getTime() : Date.now(),
  };
};

const archiveCampaignInSupabase = async (campaignId) => {
  const { data, error } = await supabase
    .from("campaigns")
    .update({ is_archived: true })
    .eq("id", campaignId)
    .select()
    .single();
  if (error) throw error;
  return {
    id: data.id,
    name: data.name || "",
    clientId: data.client_id || "",
    brand: data.brand || "",
    objective: data.objective || "",
    startDate: data.start_date || "",
    endDate: data.end_date || "",
    budget: data.budget ?? "",
    status: data.status || "planning",
    medium: data.medium || "",
    notes: data.notes || "",
    materialList: Array.isArray(data.material_list) ? data.material_list : [],
    archivedAt: data.updated_at ? new Date(data.updated_at).getTime() : Date.now(),
    createdAt: data.created_at ? new Date(data.created_at).getTime() : Date.now(),
    updatedAt: data.updated_at ? new Date(data.updated_at).getTime() : Date.now(),
  };
};

const restoreCampaignInSupabase = async (campaignId) => {
  const { data, error } = await supabase
    .from("campaigns")
    .update({ is_archived: false })
    .eq("id", campaignId)
    .select()
    .single();
  if (error) throw error;
  return {
    id: data.id,
    name: data.name || "",
    clientId: data.client_id || "",
    brand: data.brand || "",
    objective: data.objective || "",
    startDate: data.start_date || "",
    endDate: data.end_date || "",
    budget: data.budget ?? "",
    status: data.status || "planning",
    medium: data.medium || "",
    notes: data.notes || "",
    materialList: Array.isArray(data.material_list) ? data.material_list : [],
    archivedAt: null,
    createdAt: data.created_at ? new Date(data.created_at).getTime() : Date.now(),
    updatedAt: data.updated_at ? new Date(data.updated_at).getTime() : Date.now(),
  };
};

const mapRateFromSupabase = (data) => ({
  id: data.id,
  vendorId: data.vendor_id || "",
  mediaType: data.media_type || "",
  programme: data.programme || "",
  timeBelt: data.time_belt || "",
  duration: data.duration ? String(data.duration) : "30",
  ratePerSpot: data.rate_per_spot ?? "",
  discount: data.discount ?? "",
  commission: data.commission ?? "",
  vat: data.vat_rate ?? "0",
  notes: data.notes || "",
  campaignId: data.campaign_id || "",
  clientId: data.client_id || "",
  archivedAt: data.is_archived ? (data.updated_at ? new Date(data.updated_at).getTime() : Date.now()) : null,
  createdAt: data.created_at ? new Date(data.created_at).getTime() : Date.now(),
  updatedAt: data.updated_at ? new Date(data.updated_at).getTime() : Date.now(),
});

const fetchRatesFromSupabase = async () => {
  const { data, error } = await supabase
    .from("rates")
    .select("*")
    .order("created_at", { ascending: false });
  if (error) throw error;
  return (data || []).map(mapRateFromSupabase);
};

const createRatesInSupabase = async (agencyId, userId, hdr, validRows) => {
  if (!agencyId) throw new Error("No agency found for this user.");
  const payload = validRows.map(row => ({
    agency_id: agencyId,
    created_by: userId,
    vendor_id: hdr.vendorId || null,
    media_type: hdr.mediaType || null,
    programme: row.programme || null,
    time_belt: row.timeBelt || null,
    duration: row.duration ? Number(row.duration) : 30,
    rate_per_spot: row.ratePerSpot ? Number(row.ratePerSpot) : 0,
    discount: hdr.discount ? Number(hdr.discount) : 0,
    commission: hdr.commission ? Number(hdr.commission) : 0,
    vat_rate: 0,
    notes: hdr.notes || null,
    campaign_id: null,
    client_id: null,
    is_archived: false,
  }));

  const { data, error } = await supabase
    .from("rates")
    .insert(payload)
    .select();
  if (error) throw error;
  return (data || []).map(mapRateFromSupabase);
};

const updateRateInSupabase = async (rateId, hdr, row) => {
  const { data, error } = await supabase
    .from("rates")
    .update({
      vendor_id: hdr.vendorId || null,
      media_type: hdr.mediaType || null,
      programme: row.programme || null,
      time_belt: row.timeBelt || null,
      duration: row.duration ? Number(row.duration) : 30,
      rate_per_spot: row.ratePerSpot ? Number(row.ratePerSpot) : 0,
      discount: hdr.discount ? Number(hdr.discount) : 0,
      commission: hdr.commission ? Number(hdr.commission) : 0,
      notes: hdr.notes || null,
    })
    .eq("id", rateId)
    .select()
    .single();
  if (error) throw error;
  return mapRateFromSupabase(data);
};

const archiveRateInSupabase = async (rateId) => {
  const { data, error } = await supabase
    .from("rates")
    .update({ is_archived: true })
    .eq("id", rateId)
    .select()
    .single();
  if (error) throw error;
  return mapRateFromSupabase(data);
};

const restoreRateInSupabase = async (rateId) => {
  const { data, error } = await supabase
    .from("rates")
    .update({ is_archived: false })
    .eq("id", rateId)
    .select()
    .single();
  if (error) throw error;
  return mapRateFromSupabase(data);
};

const importRatesInSupabase = async (agencyId, userId, newRates) => {
  if (!agencyId) throw new Error("No agency found for this user.");
  const payload = newRates.map(r => ({
    agency_id: agencyId,
    created_by: userId,
    vendor_id: r.vendorId || null,
    media_type: r.mediaType || null,
    programme: r.programme || null,
    time_belt: r.timeBelt || null,
    duration: r.duration ? Number(r.duration) : 30,
    rate_per_spot: r.ratePerSpot ? Number(r.ratePerSpot) : 0,
    discount: r.discount ? Number(r.discount) : 0,
    commission: r.commission ? Number(r.commission) : 0,
    vat_rate: r.vat ? Number(r.vat) : 0,
    notes: r.notes || null,
    campaign_id: r.campaignId || null,
    client_id: r.clientId || null,
    is_archived: false,
  }));

  const { data, error } = await supabase
    .from("rates")
    .insert(payload)
    .select();
  if (error) throw error;
  return (data || []).map(mapRateFromSupabase);
};


const mapMpoSpotFromSupabase = (data) => ({
  id: data.id,
  programme: data.programme || "",
  wd: data.wd || "",
  timeBelt: data.time_belt || "",
  material: data.material || "",
  duration: data.duration ? String(data.duration) : "30",
  rateId: data.rate_id || "",
  ratePerSpot: data.rate_per_spot ?? 0,
  spots: String(data.spots ?? 0),
  calendarDays: Array.isArray(data.calendar_days) ? data.calendar_days : [],
  scheduleMonth: data.schedule_month || "",
});

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
  terms: Array.isArray(data.terms) ? data.terms : DEFAULT_APP_SETTINGS.mpoTerms,
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
  wd: spot.wd || null,
  time_belt: spot.timeBelt || null,
  material: spot.material || null,
  duration: spot.duration ? String(spot.duration) : "30",
  rate_id: spot.rateId || null,
  rate_per_spot: spot.ratePerSpot ? Number(spot.ratePerSpot) : 0,
  spots: spot.spots ? Number(spot.spots) : 0,
  calendar_days: Array.isArray(spot.calendarDays) ? spot.calendarDays.map(d => Number(d)) : [],
  schedule_month: spot.scheduleMonth || null,
  sort_order: index,
});

const fetchMposFromSupabase = async () => {
  const { data: mpoRows, error: mpoError } = await supabase
    .from("mpos")
    .select("*")
    .order("created_at", { ascending: false });
  if (mpoError) throw mpoError;

  const ids = (mpoRows || []).map(row => row.id);
  let spotsRows = [];
  if (ids.length) {
    const { data, error } = await supabase
      .from("mpo_spots")
      .select("*")
      .in("mpo_id", ids)
      .order("sort_order", { ascending: true });
    if (error) throw error;
    spotsRows = data || [];
  }

  const spotsByMpo = spotsRows.reduce((acc, row) => {
    (acc[row.mpo_id] ||= []).push(mapMpoSpotFromSupabase(row));
    return acc;
  }, {});

  return (mpoRows || []).map(row => mapMpoFromSupabase(row, spotsByMpo[row.id] || []));
};

const createMpoInSupabase = async (agencyId, userId, record) => {
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
    terms: Array.isArray(record.terms) ? record.terms : DEFAULT_APP_SETTINGS.mpoTerms,
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

  const { data: parent, error: parentError } = await supabase
    .from("mpos")
    .insert([parentPayload])
    .select()
    .single();
  if (parentError) throw parentError;

  let mappedSpots = [];
  if ((record.spots || []).length) {
    const payload = record.spots.map((spot, index) => ({
      mpo_id: parent.id,
      ...mapMpoSpotToSupabase(spot, index),
    }));
    const { data: insertedSpots, error: spotsError } = await supabase
      .from("mpo_spots")
      .insert(payload)
      .select()
      .order("sort_order", { ascending: true });
    if (spotsError) throw spotsError;
    mappedSpots = (insertedSpots || []).map(mapMpoSpotFromSupabase);
  }

  return mapMpoFromSupabase(parent, mappedSpots);
};

const updateMpoInSupabase = async (mpoId, record) => {
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
    terms: Array.isArray(record.terms) ? record.terms : DEFAULT_APP_SETTINGS.mpoTerms,
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

  const { data: parent, error: parentError } = await supabase
    .from("mpos")
    .update(parentPayload)
    .eq("id", mpoId)
    .select()
    .single();
  if (parentError) throw parentError;

  const { error: deleteError } = await supabase.from("mpo_spots").delete().eq("mpo_id", mpoId);
  if (deleteError) throw deleteError;

  let mappedSpots = [];
  if ((record.spots || []).length) {
    const payload = record.spots.map((spot, index) => ({
      mpo_id: mpoId,
      ...mapMpoSpotToSupabase(spot, index),
    }));
    const { data: insertedSpots, error: spotsError } = await supabase
      .from("mpo_spots")
      .insert(payload)
      .select()
      .order("sort_order", { ascending: true });
    if (spotsError) throw spotsError;
    mappedSpots = (insertedSpots || []).map(mapMpoSpotFromSupabase);
  }

  return mapMpoFromSupabase(parent, mappedSpots);
};

const archiveMpoInSupabase = async (mpoId) => {
  const { data, error } = await supabase
    .from("mpos")
    .update({ is_archived: true })
    .eq("id", mpoId)
    .select()
    .single();
  if (error) throw error;
  const { data: spotRows, error: spotError } = await supabase
    .from("mpo_spots")
    .select("*")
    .eq("mpo_id", mpoId)
    .order("sort_order", { ascending: true });
  if (spotError) throw spotError;
  return mapMpoFromSupabase(data, (spotRows || []).map(mapMpoSpotFromSupabase));
};

const restoreMpoInSupabase = async (mpoId) => {
  const { data, error } = await supabase
    .from("mpos")
    .update({ is_archived: false })
    .eq("id", mpoId)
    .select()
    .single();
  if (error) throw error;
  const { data: spotRows, error: spotError } = await supabase
    .from("mpo_spots")
    .select("*")
    .eq("mpo_id", mpoId)
    .order("sort_order", { ascending: true });
  if (spotError) throw spotError;
  return mapMpoFromSupabase(data, (spotRows || []).map(mapMpoSpotFromSupabase));
};

const updateMpoStatusInSupabase = async (mpoId, status) => {
  const { data, error } = await supabase
    .from("mpos")
    .update({ status })
    .eq("id", mpoId)
    .select()
    .single();
  if (error) throw error;
  const { data: spotRows, error: spotError } = await supabase
    .from("mpo_spots")
    .select("*")
    .eq("mpo_id", mpoId)
    .order("sort_order", { ascending: true });
  if (spotError) throw spotError;
  return mapMpoFromSupabase(data, (spotRows || []).map(mapMpoSpotFromSupabase));
};

const updateMpoExecutionInSupabase = async (mpoId, patch) => {
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
  const { data, error } = await supabase
    .from("mpos")
    .update(payload)
    .eq("id", mpoId)
    .select()
    .single();
  if (error) throw error;
  const { data: spotRows, error: spotError } = await supabase
    .from("mpo_spots")
    .select("*")
    .eq("mpo_id", mpoId)
    .order("sort_order", { ascending: true });
  if (spotError) throw spotError;
  return mapMpoFromSupabase(data, (spotRows || []).map(mapMpoSpotFromSupabase));
};

const APP_VERSION = "2.3";
const isArchived = (item) => Boolean(item?.archivedAt);
const activeOnly = (items = []) => items.filter(item => !isArchived(item));
const archivedOnly = (items = []) => items.filter(isArchived);
const archiveRecord = (item, user) => ({ ...item, archivedAt: Date.now(), archivedBy: user?.id || "system", updatedAt: Date.now() });
const restoreRecord = (item) => ({ ...item, archivedAt: null, archivedBy: null, updatedAt: Date.now() });
const pctWithin = (value) => {
  if (value === "" || value === null || value === undefined) return true;
  const n = parseFloat(value);
  return Number.isFinite(n) && n >= 0 && n <= 100;
};
const ROLE_LABELS = { admin: "Admin", planner: "Planner", buyer: "Buyer", finance: "Finance", viewer: "Viewer" };
const normalizeRole = (role = "viewer") => String(role || "viewer").trim().toLowerCase();
const formatRoleLabel = (role) => ROLE_LABELS[normalizeRole(role)] || "Viewer";
const PERMISSIONS = {
  admin: { manageVendors: true, manageClients: true, manageCampaigns: true, manageRates: true, manageMpos: true, manageMpoStatus: true, manageWorkspace: true, manageMembers: true, manageDangerZone: true },
  planner: { manageVendors: false, manageClients: true, manageCampaigns: true, manageRates: false, manageMpos: true, manageMpoStatus: false, manageWorkspace: false, manageMembers: false, manageDangerZone: false },
  buyer: { manageVendors: true, manageClients: false, manageCampaigns: false, manageRates: true, manageMpos: true, manageMpoStatus: false, manageWorkspace: false, manageMembers: false, manageDangerZone: false },
  finance: { manageVendors: false, manageClients: false, manageCampaigns: false, manageRates: false, manageMpos: false, manageMpoStatus: true, manageWorkspace: false, manageMembers: false, manageDangerZone: false },
  viewer: { manageVendors: false, manageClients: false, manageCampaigns: false, manageRates: false, manageMpos: false, manageMpoStatus: false, manageWorkspace: false, manageMembers: false, manageDangerZone: false },
};
const hasPermission = (user, key) => !!PERMISSIONS[normalizeRole(user?.role)]?.[key];
const readOnlyMessage = (user) => `Your role (${formatRoleLabel(user?.role)}) is read-only for this action.`;
const MPO_STATUS_OPTIONS = [
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
const MPO_STATUS_LABELS = Object.fromEntries(MPO_STATUS_OPTIONS.map(option => [option.value, option.label]));
const MPO_EXECUTION_STATUS_OPTIONS = [{ value: "pending", label: "Pending Dispatch" }, { value: "sent", label: "Sent to Vendor" }, { value: "confirmed", label: "Vendor Confirmed" }];
const MPO_INVOICE_STATUS_OPTIONS = [{ value: "pending", label: "Pending Invoice" }, { value: "received", label: "Invoice Received" }, { value: "approved", label: "Invoice Approved" }, { value: "disputed", label: "Invoice Disputed" }];
const MPO_PROOF_STATUS_OPTIONS = [{ value: "pending", label: "Pending Proof" }, { value: "partial", label: "Partial Proof" }, { value: "received", label: "Proof Received" }, { value: "disputed", label: "Proof Disputed" }];
const MPO_PAYMENT_STATUS_OPTIONS = [{ value: "unpaid", label: "Unpaid" }, { value: "processing", label: "Processing" }, { value: "paid", label: "Paid" }, { value: "disputed", label: "Disputed" }];
const MPO_RECON_STATUS_OPTIONS = [{ value: "not_started", label: "Not Started" }, { value: "in_progress", label: "In Progress" }, { value: "ready", label: "Ready for Review" }, { value: "completed", label: "Completed" }];
const toIsoInput = (value) => value ? String(value).slice(0, 16) : "";
const toIsoOrNull = (value) => value ? new Date(value).toISOString() : null;
const fmtDateTime = (value) => value ? new Date(value).toLocaleString('en-NG') : "—";
const getExecutionHealthColor = (mpo) => {
  if ((mpo.reconciliationStatus || 'not_started') === 'completed') return 'green';
  if ((mpo.invoiceStatus || 'pending') === 'disputed' || (mpo.proofStatus || 'pending') === 'disputed') return 'red';
  if ((mpo.dispatchStatus || 'pending') === 'confirmed') return 'blue';
  if ((mpo.dispatchStatus || 'pending') === 'sent') return 'purple';
  return 'gray';
};
const getExecutionHealthLabel = (mpo) => {
  if ((mpo.reconciliationStatus || 'not_started') === 'completed') return 'Reconciled';
  if ((mpo.invoiceStatus || 'pending') === 'disputed' || (mpo.proofStatus || 'pending') === 'disputed') return 'Disputed';
  if ((mpo.dispatchStatus || 'pending') === 'confirmed') return 'Confirmed';
  if ((mpo.dispatchStatus || 'pending') === 'sent') return 'Dispatched';
  return 'Pending';
};
const MPO_WORKFLOW_TRANSITIONS = {
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
const getAllowedMpoStatusTargets = (user, mpo) => {
  const role = normalizeRole(user?.role);
  const current = String(mpo?.status || "draft").toLowerCase();
  return MPO_WORKFLOW_TRANSITIONS[role]?.[current] || [];
};
const mpoStatusNeedsNote = (status) => ["submitted", "reviewed", "approved", "rejected", "closed"].includes(String(status || "").toLowerCase());

const MPO_WAITING_OWNER = {
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
const getMpoWorkflowMeta = (mpo) => {
  const current = String(mpo?.status || "draft").toLowerCase();
  return MPO_WAITING_OWNER[current] || MPO_WAITING_OWNER.draft;
};
const isMpoAwaitingUser = (user, mpo) => {
  if (!user || isArchived(mpo)) return false;
  const current = String(mpo?.status || "draft").toLowerCase();
  if (current === "closed") return false;
  return getMpoWorkflowMeta(mpo).roles.includes(normalizeRole(user?.role));
};
const getWorkflowActionLabel = (currentStatus, targetStatus) => {
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
const getWorkflowActionVariant = (targetStatus) => {
  const target = String(targetStatus || "").toLowerCase();
  if (target === "approved" || target === "closed") return "success";
  if (target === "reviewed" || target === "reconciled") return "blue";
  if (target === "sent") return "purple";
  if (target === "aired") return "secondary";
  if (target === "rejected") return "danger";
  return "ghost";
};
const getQuickWorkflowActions = (user, mpo) => {
  const current = String(mpo?.status || "draft").toLowerCase();
  return getAllowedMpoStatusTargets(user, mpo).map(target => ({
    value: target,
    label: getWorkflowActionLabel(current, target),
    variant: getWorkflowActionVariant(target),
  }));
};
const canEditMpoContent = (user, mpo) => {
  const role = normalizeRole(user?.role);
  if (role === "admin") return true;
  if (!hasPermission(user, "manageMpos")) return false;
  return ["draft", "rejected"].includes(String(mpo?.status || "draft").toLowerCase());
};
const formatAuditTimestamp = (value) => {
  if (!value) return "—";
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? String(value) : date.toLocaleString();
};
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
const createAuditEventInSupabase = async ({ agencyId, recordType, recordId = null, action, actor, note = "", metadata = {} }) => {
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
    .select()
    .single();
  if (error) {
    if ((error.message || "").toLowerCase().includes("audit_events")) {
      throw new Error('Missing table "audit_events". Run the provided SQL patch in Supabase SQL Editor.');
    }
    throw error;
  }
  return mapAuditEventFromSupabase(data);
};
const fetchAuditEventsForRecord = async (agencyId, recordType, recordId) => {
  if (!agencyId || !recordType || !recordId) return [];
  const { data, error } = await supabase
    .from("audit_events")
    .select("*")
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
const fetchAuditEventsForAgency = async (agencyId, filter = "all", limit = 80) => {
  if (!agencyId) return [];
  let query = supabase
    .from("audit_events")
    .select("*")
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
const fetchNotificationsFromSupabase = async (userId, agencyId, limit = 50) => {
  if (!userId || !agencyId) return [];
  const { data, error } = await supabase
    .from("notifications")
    .select("*")
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
const markNotificationReadInSupabase = async (notificationId) => {
  if (!notificationId) return null;
  const { data, error } = await supabase
    .from("notifications")
    .update({ is_read: true, read_at: new Date().toISOString() })
    .eq("id", notificationId)
    .select()
    .single();
  if (error) throw error;
  return mapNotificationFromSupabase(data);
};
const markAllNotificationsReadInSupabase = async (userId, agencyId) => {
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
const createNotificationForUserInSupabase = async ({ agencyId, recipientUserId, category = "workspace", title, message = "", recordType = null, recordId = null, linkPage = "dashboard", actor = null, metadata = {} }) => {
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
  const { data, error } = await supabase.from("notifications").insert([payload]).select().single();
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
  const { data, error } = await supabase.from("notifications").insert(payload).select();
  if (error) {
    if ((error.message || "").toLowerCase().includes("notifications")) {
      throw new Error('Missing table "notifications". Run the provided SQL patch in Supabase SQL Editor.');
    }
    throw error;
  }
  return (data || []).map(mapNotificationFromSupabase);
};
const notifyMpoWorkflowTransition = async ({ agencyId, mpo, nextStatus, actor, note = "" }) => {
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
const notifyExecutionUpdate = async ({ agencyId, mpo, actor, patch }) => {
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
const makeUserRole = (user) => user ? ({ ...user, role: normalizeRole(user.role || "viewer"), roleLabel: formatRoleLabel(user.role || "viewer") }) : null;
const downloadJSON = (filename, payload) => {
  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
};

const fileToDataUrl = (file) => new Promise((resolve, reject) => {
  try {
    if (!file) return resolve("");
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || ""));
    reader.onerror = () => reject(new Error("Failed to read selected signature file."));
    reader.readAsDataURL(file);
  } catch (error) {
    reject(error);
  }
});
const themeKeyForUser = (userId) => userId ? `msp_theme_${userId}` : "msp_theme";
const signatureKeyForUser = (userId) => userId ? `msp_signature_${userId}` : "msp_signature";
const agencyContactKey = (agencyId) => agencyId ? `msp_agency_contact_${agencyId}` : "msp_agency_contact";
const getStoredUserSignature = (userId) => store.get(signatureKeyForUser(userId), "") || "";
const setStoredUserSignature = (userId, value = "") => {
  if (!userId) return;
  if (value) store.set(signatureKeyForUser(userId), value);
  else store.del(signatureKeyForUser(userId));
};
const getStoredAgencyContact = (agencyId) => store.get(agencyContactKey(agencyId), { email: "", phone: "" }) || { email: "", phone: "" };
const setStoredAgencyContact = (agencyId, contact = {}) => {
  if (!agencyId) return;
  store.set(agencyContactKey(agencyId), { email: contact?.email || "", phone: contact?.phone || "" });
};
const persistSignatureForUser = async (currentUser, signatureDataUrl = "") => {
  setStoredUserSignature(currentUser?.id, signatureDataUrl || "");
  try {
    const { error } = await supabase.auth.updateUser({ data: { signature_data_url: signatureDataUrl || "" } });
    if (error) throw error;
  } catch (error) {
    console.error("Failed to persist signature metadata:", error);
  }
  return signatureDataUrl || "";
};
const loadBrowserScript = (src, readyCheck) => new Promise((resolve, reject) => {
  try {
    const ready = typeof readyCheck === "function" ? readyCheck() : readyCheck;
    if (ready) return resolve(ready);
    const existing = document.querySelector(`script[src="${src}"]`);
    if (existing) {
      existing.addEventListener("load", () => resolve(typeof readyCheck === "function" ? readyCheck() : readyCheck), { once: true });
      existing.addEventListener("error", () => reject(new Error(`Failed to load ${src}`)), { once: true });
      return;
    }
    const script = document.createElement("script");
    script.src = src;
    script.async = true;
    script.onload = () => resolve(typeof readyCheck === "function" ? readyCheck() : readyCheck);
    script.onerror = () => reject(new Error(`Failed to load ${src}`));
    document.head.appendChild(script);
  } catch (error) {
    reject(error);
  }
});
const loadPreviewPdfLibraries = async () => {
  await loadBrowserScript("https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js", () => window.html2canvas);
  await loadBrowserScript("https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js", () => window.jspdf?.jsPDF);
  return { html2canvas: window.html2canvas, jsPDF: window.jspdf.jsPDF };
};

const DEFAULT_SESSION_HOURS = 8;
const DEFAULT_APP_SETTINGS = {
  vatRate: 7.5,
  roundToWholeNaira: false,
  sessionHours: DEFAULT_SESSION_HOURS,
  mpoTerms: [
    "Make-goods or low quality reproductions are unacceptable.",
    "Transmission/broadcast of wrong material is totally unacceptable and could be considered as non-compliance. Kindly contact the agency if material is bad or damaged.",
    "Payment will be made within 30 days after submission of invoice and COT, with compliance based on the client's tracking or monitoring report.",
    "Any change in rate must be communicated within 90 days from the expiration of this contract period.",
    "Please note that on no account should there be any change to this flighting without approval from the agency.",
    "Please acknowledge receipt and implementation of this order.",
    "As agreed, 1 complimentary spot will be run for every 4 paid spots to support the campaign objective."
  ],
};
const mergeAppSettings = (saved = {}) => {
  const terms = Array.isArray(saved?.mpoTerms) && saved.mpoTerms.length ? saved.mpoTerms : DEFAULT_APP_SETTINGS.mpoTerms;
  return { ...DEFAULT_APP_SETTINGS, ...(saved || {}), mpoTerms: terms };
};
const getAppSettings = () => mergeAppSettings(store.get("msp_app_settings", {}));

const fetchAppSettingsFromSupabase = async (agencyId) => {
  if (!agencyId) return getAppSettings();
  const { data, error } = await supabase
    .from("app_settings")
    .select("*")
    .eq("agency_id", agencyId)
    .maybeSingle();
  if (error) throw error;
  return mergeAppSettings({
    vatRate: data?.vat_rate ?? DEFAULT_APP_SETTINGS.vatRate,
    roundToWholeNaira: data?.round_to_whole_naira ?? DEFAULT_APP_SETTINGS.roundToWholeNaira,
    sessionHours: data?.session_hours ?? DEFAULT_APP_SETTINGS.sessionHours,
    mpoTerms: Array.isArray(data?.mpo_terms) ? data.mpo_terms : DEFAULT_APP_SETTINGS.mpoTerms,
  });
};

const saveAppSettingsToSupabase = async (agencyId, settings) => {
  if (!agencyId) throw new Error("No agency found for this workspace.");
  const payload = {
    agency_id: agencyId,
    vat_rate: Number(settings?.vatRate) || DEFAULT_APP_SETTINGS.vatRate,
    round_to_whole_naira: !!settings?.roundToWholeNaira,
    session_hours: Math.max(1, Number(settings?.sessionHours) || DEFAULT_APP_SETTINGS.sessionHours),
    mpo_terms: Array.isArray(settings?.mpoTerms) ? settings.mpoTerms : DEFAULT_APP_SETTINGS.mpoTerms,
  };
  const { data, error } = await supabase
    .from("app_settings")
    .upsert(payload, { onConflict: "agency_id" })
    .select()
    .single();
  if (error) throw error;
  return mergeAppSettings({
    vatRate: data?.vat_rate,
    roundToWholeNaira: data?.round_to_whole_naira,
    sessionHours: data?.session_hours,
    mpoTerms: Array.isArray(data?.mpo_terms) ? data.mpo_terms : DEFAULT_APP_SETTINGS.mpoTerms,
  });
};

const updateProfileInSupabase = async (currentUser, form) => {
  const payload = {
    full_name: form.name?.trim() || currentUser.name || "",
    title: form.title || "",
    phone: form.phone || "",
  };
  const nextSignature = form.signatureDataUrl || "";
  setStoredUserSignature(currentUser?.id, nextSignature);
  const authPayload = {
    data: {
      full_name: payload.full_name,
      title: payload.title,
      phone: payload.phone,
      agency_name: currentUser.agency || "",
      agency_code: currentUser.agencyCode || "",
      agency_email: currentUser.agencyEmail || "",
      agency_phone: currentUser.agencyPhone || "",
      signature_data_url: nextSignature,
    },
  };
  if (form.email && form.email !== currentUser.email) authPayload.email = form.email;
  const { error: authError } = await supabase.auth.updateUser(authPayload);
  if (authError) throw authError;
  const { error: profileError } = await supabase
    .from("profiles")
    .update({ ...payload, email: form.email || currentUser.email || null })
    .eq("id", currentUser.id);
  if (profileError) throw profileError;
  return {
    ...currentUser,
    name: payload.full_name,
    title: payload.title,
    phone: payload.phone,
    email: form.email || currentUser.email || "",
    signatureDataUrl: nextSignature,
  };
};

const updateAgencyInSupabase = async (agencyId, form) => {
  if (!agencyId) throw new Error("No agency found for this workspace.");

  const fallbackContact = { email: form.email || "", phone: form.phone || "" };
  let data = null;
  const primary = await supabase
    .from("agencies")
    .update({
      name: form.agency?.trim() || "My Agency",
      address: form.address || null,
      email: form.email || null,
      phone: form.phone || null,
    })
    .eq("id", agencyId)
    .select("id, name, address, agency_code, email, phone")
    .maybeSingle();

  if (!primary.error) {
    data = primary.data;
  } else {
    const fallback = await supabase
      .from("agencies")
      .update({
        name: form.agency?.trim() || "My Agency",
        address: form.address || null,
      })
      .eq("id", agencyId)
      .select("id, name, address, agency_code")
      .single();
    if (fallback.error) throw fallback.error;
    data = fallback.data;
  }

  setStoredAgencyContact(agencyId, fallbackContact);
  try {
    await supabase.auth.updateUser({
      data: {
        agency_name: data?.name || form.agency || "My Agency",
        agency_code: data?.agency_code || "",
        agency_email: data?.email || fallbackContact.email || "",
        agency_phone: data?.phone || fallbackContact.phone || "",
      },
    });
  } catch (error) {
    console.error("Failed to update agency contact metadata:", error);
  }

  return {
    agency: data?.name || form.agency || "My Agency",
    agencyAddress: data?.address || "",
    agencyCode: data?.agency_code || "",
    agencyEmail: data?.email || fallbackContact.email || "",
    agencyPhone: data?.phone || fallbackContact.phone || "",
  };
};

const fetchAgencyMembersFromSupabase = async (agencyId) => {
  if (!agencyId) return [];
  const { data, error } = await supabase
    .from("profiles")
    .select("id, email, full_name, title, phone, role, agency_id, created_at")
    .eq("agency_id", agencyId)
    .order("full_name", { ascending: true });
  if (error) throw error;
  return (data || []).map(member => ({
    id: member.id,
    email: member.email || "",
    name: member.full_name || member.email || "User",
    title: member.title || "",
    phone: member.phone || "",
    role: normalizeRole(member.role || "viewer"),
    roleLabel: formatRoleLabel(member.role || "viewer"),
    createdAt: member.created_at ? new Date(member.created_at).getTime() : null,
  }));
};

const updateAgencyMemberRoleInSupabase = async (targetUserId, nextRole) => {
  const role = normalizeRole(nextRole);
  const { error } = await supabase.rpc("admin_update_user_role", {
    p_target_user_id: targetUserId,
    p_new_role: role,
  });
  if (error) {
    if ((error.message || "").toLowerCase().includes("admin_update_user_role")) {
      throw new Error('Missing database function "admin_update_user_role". Run the provided SQL patch in Supabase SQL Editor.');
    }
    throw error;
  }
  return role;
};

const changePasswordInSupabase = async (newPassword) => {
  const { error } = await supabase.auth.updateUser({ password: newPassword });
  if (error) throw error;
};

const generateNextMpoNoFromSupabase = async (brand = "MPO") => {
  const { data, error } = await supabase.rpc("generate_next_mpo_no", { p_brand: brand || "MPO" });
  if (error) {
    if ((error.message || "").toLowerCase().includes("generate_next_mpo_no")) {
      throw new Error('Missing database function "generate_next_mpo_no". Run the provided SQL patch in Supabase SQL Editor.');
    }
    throw error;
  }
  return data;
};
const roundMoneyValue = (value, settings = getAppSettings()) => {
  const n = parseFloat(value) || 0;
  return settings?.roundToWholeNaira ? Math.round(n) : Math.round(n * 100) / 100;
};
const escapeHtml = (value) => String(value ?? "")
  .replace(/&/g, "&amp;")
  .replace(/</g, "&lt;")
  .replace(/>/g, "&gt;")
  .replace(/"/g, "&quot;")
  .replace(/'/g, "&#39;");
const safeText = (value) => String(value ?? "").replace(/[\u0000-\u001F\u007F]/g, " ").trim();
const sanitizeMPOForExport = (mpo) => ({
  ...mpo,
  mpoNo: escapeHtml(safeText(mpo?.mpoNo)),
  date: escapeHtml(safeText(mpo?.date)),
  month: escapeHtml(safeText(mpo?.month)),
  year: escapeHtml(safeText(mpo?.year)),
  vendorName: escapeHtml(safeText(mpo?.vendorName)),
  clientName: escapeHtml(safeText(mpo?.clientName)),
  brand: escapeHtml(safeText(mpo?.brand)),
  campaignName: escapeHtml(safeText(mpo?.campaignName)),
  agencyAddress: escapeHtml(safeText(mpo?.agencyAddress)),
  agencyEmail: escapeHtml(safeText(mpo?.agencyEmail)),
  agencyPhone: escapeHtml(safeText(mpo?.agencyPhone)),
  signedBy: escapeHtml(safeText(mpo?.signedBy)),
  signedTitle: escapeHtml(safeText(mpo?.signedTitle)),
  signedSignature: mpo?.signedSignature || "",
  preparedBy: escapeHtml(safeText(mpo?.preparedBy)),
  preparedContact: escapeHtml(safeText(mpo?.preparedContact)),
  preparedTitle: escapeHtml(safeText(mpo?.preparedTitle)),
  preparedSignature: mpo?.preparedSignature || "",
  medium: escapeHtml(safeText(mpo?.medium)),
  surchLabel: escapeHtml(safeText(mpo?.surchLabel)),
  transmitMsg: escapeHtml(safeText(mpo?.transmitMsg)),
  terms: (Array.isArray(mpo?.terms) ? mpo.terms : DEFAULT_APP_SETTINGS.mpoTerms).map(t => escapeHtml(safeText(t))),
  spots: (mpo?.spots || []).map((s) => ({
    ...s,
    programme: escapeHtml(safeText(s?.programme)),
    wd: escapeHtml(safeText(s?.wd)),
    timeBelt: escapeHtml(safeText(s?.timeBelt)),
    material: escapeHtml(safeText(s?.material)),
    duration: escapeHtml(safeText(s?.duration)),
    scheduleMonth: escapeHtml(safeText(s?.scheduleMonth)),
  })),
});
const getDefaultTheme = (userId = null) => {
  const saved = store.get(themeKeyForUser(userId), null);
  if (saved === "light" || saved === "dark") return saved;
  return window?.matchMedia?.("(prefers-color-scheme: dark)")?.matches ? "dark" : "light";
};
const sessionDurationMs = (settings = getAppSettings()) => {
  const hours = Math.max(1, parseFloat(settings?.sessionHours) || DEFAULT_SESSION_HOURS);
  return hours * 60 * 60 * 1000;
};
const saveSession = (user, settings = getAppSettings()) => {
  if (!user) {
    store.del("msp_session");
    return null;
  }
  const now = Date.now();
  const safeUser = makeUserRole(user);
  const session = {
    ...safeUser,
    sessionStartedAt: safeUser.sessionStartedAt || now,
    lastActiveAt: now,
    sessionExpiresAt: now + sessionDurationMs(settings),
  };
  store.set("msp_session", session);
  return session;
};
const touchSession = (settings = getAppSettings()) => {
  const current = store.get("msp_session");
  if (!current) return null;
  return saveSession(current, settings);
};
const sessionExpired = (session) => !session || ((session.sessionExpiresAt || 0) <= Date.now());
const hashPassword = async (password) => {
  const source = new TextEncoder().encode(String(password || ""));
  const digest = await window.crypto.subtle.digest("SHA-256", source);
  return Array.from(new Uint8Array(digest)).map(b => b.toString(16).padStart(2, "0")).join("");
};
const verifyPassword = async (user, candidate) => {
  if (!user) return false;

  if (user.passwordHash) return (await hashPassword(candidate)) === user.passwordHash;
  return user.password === candidate;
};
const legacyUserNeedsHashUpgrade = (user) => Boolean(user?.password && !user?.passwordHash);
const downloadTextFile = (filename, content, mimeType = "text/plain;charset=utf-8") => {
  const blob = new Blob([content], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
};
const downloadRateTemplate = () => {
  const csv = [
    ["Vendor","Media","Programme","Timebelt","Duration","Rate","Discount","Commission","Notes"],
    ["Example FM","Radio","Morning Drive","06:00-09:00","30","150000","10","5","Prime time rate"],
    ["Example TV","Television","News at 9","21:00-21:30","45","450000","0","7.5","Headline bulletin"],
  ].map(row => row.map(v => `"${String(v).replace(/"/g, '""')}"`).join(",")).join("\n");
  downloadTextFile("mediadesk-rate-template.csv", csv, "text/csv;charset=utf-8");
};

const MPO_STATUS_COLORS = { draft: "accent", submitted: "blue", reviewed: "purple", approved: "green", sent: "teal", aired: "orange", reconciled: "blue", closed: "green", rejected: "red" };

/* ── STYLES ─────────────────────────────────────────────── */
const GlobalStyle = ({ theme }) => (
  <style>{`
    @import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;600;700;800&family=DM+Sans:ital,wght@0,300;0,400;0,500;0,600;1,400&display=swap');
    *,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
    :root{
      ${theme === "dark" ? `
      --bg:#07090f;--bg2:#0e1118;--bg3:#141824;--bg4:#1c2233;
      --border:rgba(255,255,255,0.07);--border2:rgba(255,255,255,0.13);
      --text:#e8ecf4;--text2:#8b93a7;--text3:#4f576b;
      ` : `
      --bg:#f0f2f7;--bg2:#ffffff;--bg3:#f5f7fc;--bg4:#e8ecf4;
      --border:rgba(0,0,0,0.08);--border2:rgba(0,0,0,0.14);
      --text:#111827;--text2:#4b5563;--text3:#9ca3af;
      `}
      --accent:#f0a500;--blue:#3b7ef5;--green:#16a34a;
      --red:#ef4444;--purple:#8b5cf6;--teal:#0d9488;--orange:#f97316;
    }
    html,body,#root{height:100%;font-family:'DM Sans',sans-serif;background:var(--bg);color:var(--text)}
    *{scrollbar-width:thin;scrollbar-color:var(--bg4) transparent}
    ::-webkit-scrollbar{width:5px}::-webkit-scrollbar-track{background:transparent}::-webkit-scrollbar-thumb{background:var(--bg4);border-radius:3px}
    input,select,textarea,button{font-family:inherit}
    .fade{animation:fadeIn .25s ease}
    @keyframes fadeIn{from{opacity:0;transform:translateY(6px)}to{opacity:1;transform:translateY(0)}}
    @keyframes spin{to{transform:rotate(360deg)}}
    .spin{animation:spin 1s linear infinite}
    textarea{resize:vertical}
    @media print{
      body{background:#fff!important;color:#000!important}
      .no-print{display:none!important}
      .print-area{display:block!important}
    }
  `}</style>
);

/* ── UI PRIMITIVES ──────────────────────────────────────── */
const Btn = ({ children, variant = "primary", size = "md", onClick, type = "button", disabled, style, icon, loading }) => {
  const sz = size === "sm" ? { padding: "5px 13px", fontSize: 12 } : size === "lg" ? { padding: "13px 28px", fontSize: 15 } : { padding: "9px 18px", fontSize: 13 };
  const vars = {
    primary:  { background: "var(--accent)", color: "#000" },
    secondary:{ background: "var(--bg4)", color: "var(--text)", border: "1px solid var(--border2)" },
    danger:   { background: "rgba(239,68,68,.15)", color: "var(--red)", border: "1px solid rgba(239,68,68,.3)" },
    ghost:    { background: "transparent", color: "var(--text2)", border: "1px solid var(--border)" },
    success:  { background: "rgba(34,197,94,.15)", color: "var(--green)", border: "1px solid rgba(34,197,94,.3)" },
    blue:     { background: "rgba(59,126,245,.15)", color: "var(--blue)", border: "1px solid rgba(59,126,245,.3)" },
    purple:   { background: "rgba(139,92,246,.15)", color: "var(--purple)", border: "1px solid rgba(139,92,246,.3)" },
  };
  return (
    <button type={type} onClick={onClick} disabled={disabled || loading}
      style={{ display: "inline-flex", alignItems: "center", gap: 7, fontFamily: "'Syne',sans-serif", fontWeight: 600, border: "none", borderRadius: 9, cursor: (disabled || loading) ? "not-allowed" : "pointer", transition: "all .18s", opacity: (disabled || loading) ? .55 : 1, outline: "none", whiteSpace: "nowrap", ...sz, ...vars[variant], ...style }}
      onMouseEnter={e => { if (!disabled && !loading) e.currentTarget.style.filter = "brightness(1.18)"; }}
      onMouseLeave={e => e.currentTarget.style.filter = ""}>
      {loading ? <span className="spin" style={{ fontSize: 14 }}>⟳</span> : icon && <span style={{ fontSize: size === "sm" ? 13 : 16 }}>{icon}</span>}
      {children}
    </button>
  );
};

const Field = ({ label, value, onChange, type = "text", placeholder, required, options, note, error, rows }) => (
  <div style={{ display: "flex", flexDirection: "column", gap: 5 }}>
    {label && <label style={{ fontSize: 11, fontWeight: 600, color: "var(--text3)", textTransform: "uppercase", letterSpacing: ".08em" }}>{label}{required && <span style={{ color: "var(--accent)", marginLeft: 3 }}>*</span>}</label>}
    {options ? (
      <select value={value} onChange={e => onChange(e.target.value)}
        style={{ background: "var(--bg3)", border: `1px solid ${error ? "var(--red)" : "var(--border2)"}`, borderRadius: 8, padding: "9px 13px", color: value ? "var(--text)" : "var(--text3)", fontSize: 13, outline: "none", cursor: "pointer" }}>
        <option value="">{placeholder || "Select…"}</option>
        {options.map(o => <option key={o.value ?? o} value={o.value ?? o}>{o.label ?? o}</option>)}
      </select>
    ) : rows ? (
      <textarea value={value} onChange={e => onChange(e.target.value)} placeholder={placeholder} rows={rows}
        style={{ background: "var(--bg3)", border: `1px solid ${error ? "var(--red)" : "var(--border2)"}`, borderRadius: 8, padding: "9px 13px", color: "var(--text)", fontSize: 13, outline: "none", lineHeight: 1.5 }}
        onFocus={e => e.target.style.borderColor = "var(--accent)"}
        onBlur={e => e.target.style.borderColor = error ? "var(--red)" : "var(--border2)"} />
    ) : (
      <input type={type} value={value} onChange={e => onChange(e.target.value)} placeholder={placeholder} required={required}
        style={{ background: "var(--bg3)", border: `1px solid ${error ? "var(--red)" : "var(--border2)"}`, borderRadius: 8, padding: "9px 13px", color: "var(--text)", fontSize: 13, outline: "none", width: "100%" }}
        onFocus={e => e.target.style.borderColor = "var(--accent)"}
        onBlur={e => e.target.style.borderColor = error ? "var(--red)" : "var(--border2)"} />
    )}
    {error && <span style={{ fontSize: 11, color: "var(--red)" }}>{error}</span>}
    {note && !error && <span style={{ fontSize: 11, color: "var(--text3)" }}>{note}</span>}
  </div>
);

const AttachmentField = ({ label, url, onUrlChange, onFileSelected, uploading, accept = ".pdf,image/*,.doc,.docx,.xls,.xlsx,.ppt,.pptx" }) => (
  <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
    <Field label={`${label} Link`} value={url} onChange={onUrlChange} placeholder="Paste a link or upload a file below" />
    <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
      <input
        type="file"
        accept={accept}
        onChange={e => {
          const file = e.target.files?.[0] || null;
          if (file) onFileSelected(file);
          e.target.value = "";
        }}
        style={{ fontSize: 12, color: "var(--text2)", maxWidth: "100%" }}
      />
      {uploading && <span style={{ fontSize: 12, color: "var(--blue)", fontWeight: 600 }}>Uploading…</span>}
      {url && (
        <a href={url} target="_blank" rel="noreferrer" style={{ fontSize: 12, color: "var(--accent)", textDecoration: "none" }}>
          Open current file
        </a>
      )}
    </div>
  </div>
);

const Card = ({ children, style, glow, hoverable = true }) => (
  <div
    style={{ background: "var(--bg2)", border: `1px solid ${glow ? "rgba(240,165,0,.25)" : "var(--border)"}`, borderRadius: 14, padding: 22, boxShadow: glow ? "0 0 28px rgba(240,165,0,.07)" : "0 2px 12px rgba(0,0,0,.3)", transition: hoverable ? "transform .16s ease, box-shadow .16s ease, border-color .16s ease" : "none", ...style }}
    onMouseEnter={e => {
      if (!hoverable) return;
      e.currentTarget.style.transform = "translateY(-2px)";
      e.currentTarget.style.boxShadow = glow ? "0 10px 28px rgba(240,165,0,.14)" : "0 10px 24px rgba(0,0,0,.14)";
      e.currentTarget.style.borderColor = "rgba(240,165,0,.22)";
    }}
    onMouseLeave={e => {
      if (!hoverable) return;
      e.currentTarget.style.transform = "translateY(0)";
      e.currentTarget.style.boxShadow = glow ? "0 0 28px rgba(240,165,0,.07)" : "0 2px 12px rgba(0,0,0,.3)";
      e.currentTarget.style.borderColor = glow ? "rgba(240,165,0,.25)" : "var(--border)";
    }}
  >
    {children}
  </div>
);

const Badge = ({ children, color = "accent" }) => {
  const map = { accent: ["rgba(240,165,0,.15)", "var(--accent)"], green: ["rgba(34,197,94,.15)", "var(--green)"], blue: ["rgba(59,126,245,.15)", "var(--blue)"], red: ["rgba(239,68,68,.15)", "var(--red)"], purple: ["rgba(139,92,246,.15)", "var(--purple)"], teal: ["rgba(20,184,166,.15)", "var(--teal)"], orange: ["rgba(249,115,22,.15)", "var(--orange)"] };
  const [bg, fg] = map[color] || map.accent;
  return <span style={{ background: bg, color: fg, padding: "3px 10px", borderRadius: 99, fontSize: 11, fontWeight: 600, fontFamily: "'Syne',sans-serif", whiteSpace: "nowrap" }}>{children}</span>;
};

const Modal = ({ title, children, onClose, width = 540 }) => (
  <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,.72)", zIndex: 1000, display: "flex", alignItems: "center", justifyContent: "center", padding: 20, backdropFilter: "blur(5px)" }}
    onClick={e => e.target === e.currentTarget && onClose()}>
    <div className="fade" style={{ background: "var(--bg2)", border: "1px solid var(--border2)", borderRadius: 18, width: "100%", maxWidth: width, maxHeight: "92vh", overflowY: "auto", boxShadow: "0 28px 80px rgba(0,0,0,.65)" }}>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "18px 22px", borderBottom: "1px solid var(--border)", position: "sticky", top: 0, background: "var(--bg2)", zIndex: 1 }}>
        <h3 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 17 }}>{title}</h3>
        <button onClick={onClose} style={{ background: "var(--bg4)", border: "none", color: "var(--text2)", width: 30, height: 30, borderRadius: 7, fontSize: 17, cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center" }}>×</button>
      </div>
      <div style={{ padding: 22 }}>{children}</div>
    </div>
  </div>
);

const Toast = ({ msg, type, onDone }) => {
  useEffect(() => { const t = setTimeout(onDone, 3500); return () => clearTimeout(t); }, []);
  const bg = type === "error" ? "#ef4444" : type === "success" ? "#22c55e" : "var(--bg4)";
  return <div style={{ position: "fixed", bottom: 22, right: 22, zIndex: 9999, background: bg, color: type === "error" || type === "success" ? "#fff" : "var(--text)", padding: "11px 20px", borderRadius: 11, fontWeight: 600, fontSize: 13, boxShadow: "0 8px 28px rgba(0,0,0,.4)", animation: "fadeIn .3s ease", maxWidth: 320 }}>{msg}</div>;
};

const Empty = ({ icon, title, sub }) => (
  <div style={{ textAlign: "center", padding: "52px 16px" }}>
    <div style={{ fontSize: 44, marginBottom: 14, opacity: .35 }}>{icon}</div>
    <div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 17, marginBottom: 7 }}>{title}</div>
    <div style={{ color: "var(--text2)", fontSize: 13 }}>{sub}</div>
  </div>
);

const Stat = ({ label, value, sub, color = "var(--accent)", icon }) => (
  <Card hoverable style={{ position: "relative", overflow: "hidden", minWidth: 0 }}>
    <div style={{ position: "absolute", top: -16, right: -12, fontSize: 72, opacity: .04, pointerEvents: "none" }}>{icon}</div>
    <div style={{ fontSize: 11, color: "var(--text3)", fontWeight: 600, textTransform: "uppercase", letterSpacing: ".08em", marginBottom: 7 }}>{label}</div>
    <div style={{ fontSize: "clamp(22px, 3vw, 30px)", lineHeight: 1.05, fontWeight: 800, fontFamily: "'Syne',sans-serif", color, minWidth: 0, overflowWrap: "anywhere", wordBreak: "break-word" }}>{value}</div>
    {sub && <div style={{ fontSize: 12, color: "var(--text2)", marginTop: 6, overflowWrap: "anywhere" }}>{sub}</div>}
  </Card>
);

const Confirm = ({ msg, onYes, onNo, danger = true }) => (
  <Modal title="Confirm Action" onClose={onNo} width={380}>
    <p style={{ color: "var(--text2)", marginBottom: 22, lineHeight: 1.6 }}>{msg}</p>
    <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
      <Btn variant="ghost" onClick={onNo}>Cancel</Btn>
      <Btn variant={danger ? "danger" : "success"} onClick={onYes}>Confirm</Btn>
    </div>
  </Modal>
);

/* ── AUTH ───────────────────────────────────────────────── */
const AuthPage = ({ onLogin }) => {
  const [mode, setMode] = useState("login");
  const [f, setF] = useState({
    name: "",
    email: "",
    password: "",
    agency: "",
    agencyCode: "",
    agencyMode: "create",
    title: "",
    phone: "",
    confirm: "",
  });
  const [err, setErr] = useState("");
  const [loading, setLoading] = useState(false);

  const u = (k) => (v) => setF((p) => ({ ...p, [k]: v }));

  const submit = async () => {
    setErr("");

    if (mode === "register") {
      if (!f.name || !f.email || !f.password) {
        return setErr("Full name, email, and password are required.");
      }
      if (f.agencyMode === "create" && !f.agency.trim()) {
        return setErr("Agency name is required when creating a new agency.");
      }
      if (f.agencyMode === "join" && !normalizeAgencyCode(f.agencyCode)) {
        return setErr("Agency invite code is required when joining an existing agency.");
      }
      if (f.password !== f.confirm) {
        return setErr("Passwords do not match.");
      }
      if (f.password.length < 6) {
        return setErr("Password must be at least 6 characters.");
      }
    }

    setLoading(true);

    try {
      if (mode === "login") {
        const { error } = await supabase.auth.signInWithPassword({
          email: f.email,
          password: f.password,
        });

        if (error) throw error;
        // App-level auth listener will load the user and agency.
      } else {
        const { data, error } = await supabase.auth.signUp({
          email: f.email,
          password: f.password,
          options: {
            data: {
              full_name: f.name,
              title: f.title || "",
              phone: f.phone || "",
              agency_name: f.agencyMode === "create" ? f.agency.trim() : "",
              agency_code: f.agencyMode === "join" ? normalizeAgencyCode(f.agencyCode) : "",
              agency_mode: f.agencyMode,
            },
          },
        });

        if (error) throw error;

        if (!data.user) {
          throw new Error("Signup failed. No user was returned.");
        }

        const { error: profileError } = await supabase
          .from("profiles")
          .update({
            full_name: f.name,
            title: f.title || "",
            phone: f.phone || "",
          })
          .eq("id", data.user.id);

        if (profileError) throw profileError;

        if (!data.session) {
          setErr(
            "Account created. Check your email to confirm your account, then sign in to finish joining your agency."
          );
          setMode("login");
        }
        // If a session exists immediately, App-level auth listener will complete agency join/load.
      }
    } catch (e) {
      setErr(e.message || "Something went wrong.");
    } finally {
      setLoading(false);
    }
  };

  return (
    <div
      style={{
        minHeight: "100vh",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        background:
          "radial-gradient(ellipse 80% 50% at 50% -10%,rgba(240,165,0,.09) 0%,transparent 70%),var(--bg)",
        padding: 20,
      }}
    >
      <div style={{ width: "100%", maxWidth: 420 }}>
        <div style={{ textAlign: "center", marginBottom: 36 }}>
          <div
            style={{
              display: "inline-flex",
              alignItems: "center",
              gap: 11,
              background: "var(--bg2)",
              border: "1px solid var(--border2)",
              borderRadius: 14,
              padding: "11px 18px",
              marginBottom: 22,
            }}
          >
            <div
              style={{
                width: 34,
                height: 34,
                background: "var(--accent)",
                borderRadius: 9,
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                fontSize: 17,
              }}
            >
              📡
            </div>
            <div>
              <div
                style={{
                  fontFamily: "'Syne',sans-serif",
                  fontWeight: 800,
                  fontSize: 17,
                }}
              >
                MediaDesk Pro
              </div>
              <div style={{ fontSize: 10, color: "var(--text3)" }}>
                MEDIA SCHEDULE PLATFORM
              </div>
            </div>
          </div>

          <h1
            style={{
              fontFamily: "'Syne',sans-serif",
              fontWeight: 800,
              fontSize: 26,
              letterSpacing: "-.03em",
            }}
          >
            {mode === "login" ? "Welcome back" : "Create account"}
          </h1>
          <p style={{ color: "var(--text2)", marginTop: 7, fontSize: 14 }}>
            {mode === "login"
              ? "Sign in to your agency workspace"
              : "Create an account and either create or join an agency"}
          </p>
        </div>

        <Card>
          <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
            {mode === "register" && (
              <>
                <Field
                  label="Full Name"
                  value={f.name}
                  onChange={u("name")}
                  placeholder="Jane Okafor"
                  required
                />
                <Field
                  label="Job Title"
                  value={f.title}
                  onChange={u("title")}
                  placeholder="Media Buyer / Account Executive"
                />
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
                  <button
                    type="button"
                    onClick={() => setF(p => ({ ...p, agencyMode: "create", agencyCode: p.agencyCode || "" }))}
                    style={{ padding: "11px 12px", borderRadius: 10, border: f.agencyMode === "create" ? "1px solid var(--accent)" : "1px solid var(--border)", background: f.agencyMode === "create" ? "rgba(240,165,0,.12)" : "var(--bg3)", color: f.agencyMode === "create" ? "var(--accent)" : "var(--text2)", cursor: "pointer", fontWeight: 600 }}
                  >
                    Create New Agency
                  </button>
                  <button
                    type="button"
                    onClick={() => setF(p => ({ ...p, agencyMode: "join", agency: p.agency || "" }))}
                    style={{ padding: "11px 12px", borderRadius: 10, border: f.agencyMode === "join" ? "1px solid var(--accent)" : "1px solid var(--border)", background: f.agencyMode === "join" ? "rgba(240,165,0,.12)" : "var(--bg3)", color: f.agencyMode === "join" ? "var(--accent)" : "var(--text2)", cursor: "pointer", fontWeight: 600 }}
                  >
                    Join With Invite Code
                  </button>
                </div>
                {f.agencyMode === "create" ? (
                  <Field
                    label="Agency Name"
                    value={f.agency}
                    onChange={u("agency")}
                    placeholder="Apex Media Ltd"
                    required
                    note="You will get an invite code to share with teammates after signup."
                  />
                ) : (
                  <Field
                    label="Agency Invite Code"
                    value={f.agencyCode}
                    onChange={u("agencyCode")}
                    placeholder="QVT-7K4P"
                    required
                    note="Ask your agency admin for the invite code."
                  />
                )}
                <Field
                  label="Phone Number"
                  type="tel"
                  value={f.phone}
                  onChange={u("phone")}
                  placeholder="+234 800 000 0000"
                />
              </>
            )}

            <Field
              label="Email"
              type="email"
              value={f.email}
              onChange={u("email")}
              placeholder="you@agency.com"
              required
            />

            <Field
              label="Password"
              type="password"
              value={f.password}
              onChange={u("password")}
              placeholder="••••••••"
              required
            />

            {mode === "register" && (
              <Field
                label="Confirm Password"
                type="password"
                value={f.confirm}
                onChange={u("confirm")}
                placeholder="••••••••"
              />
            )}

            {err && (
              <div
                style={{
                  background: "rgba(239,68,68,.1)",
                  border: "1px solid rgba(239,68,68,.3)",
                  borderRadius: 8,
                  padding: "9px 13px",
                  color: "var(--red)",
                  fontSize: 12,
                }}
              >
                {err}
              </div>
            )}

            <Btn
              size="lg"
              onClick={submit}
              loading={loading}
              style={{ width: "100%", justifyContent: "center", marginTop: 4 }}
            >
              {mode === "login" ? "Sign In →" : "Create Account →"}
            </Btn>

            <p style={{ textAlign: "center", color: "var(--text2)", fontSize: 13 }}>
              {mode === "login" ? "No account? " : "Have an account? "}
              <button
                onClick={() => {
                  setMode(mode === "login" ? "register" : "login");
                  setErr("");
                }}
                style={{
                  background: "none",
                  border: "none",
                  color: "var(--accent)",
                  fontWeight: 600,
                  cursor: "pointer",
                  fontSize: 13,
                }}
              >
                {mode === "login" ? "Register" : "Sign In"}
              </button>
            </p>
          </div>
        </Card>
      </div>
    </div>
  );
};

/* ── SIDEBAR ────────────────────────────────────────────── */
const Sidebar = ({ page, setPage, user, onLogout, collapsed, setCollapsed, theme, toggleTheme, unreadNotifications = 0 }) => {
  const nav = [
    { id: "dashboard", icon: "◈", label: "Dashboard" },
    { id: "vendors",   icon: "🏢", label: "Vendors" },
    { id: "clients",   icon: "👥", label: "Clients & Brands" },
    { id: "campaigns", icon: "📢", label: "Campaigns" },
    { id: "rates",     icon: "💰", label: "Media Rates" },
    { id: "mpo",       icon: "📄", label: "MPO Generator" },
    { id: "reports",   icon: "📊", label: "Reports" },
    { id: "settings",  icon: "⚙️", label: "Settings", badge: unreadNotifications },
  ];
  return (
    <div style={{ width: collapsed ? 64 : 230, minHeight: "100vh", background: "var(--bg2)", borderRight: "1px solid var(--border)", display: "flex", flexDirection: "column", transition: "width .22s ease", overflow: "hidden", flexShrink: 0, position: "sticky", top: 0, height: "100vh" }}>
      <div style={{ padding: collapsed ? "18px 14px" : "18px 18px", borderBottom: "1px solid var(--border)", display: "flex", alignItems: "center", gap: 11, minHeight: 68 }}>
        <div style={{ width: 34, height: 34, background: "var(--accent)", borderRadius: 9, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 17, flexShrink: 0 }}>📡</div>
        {!collapsed && <div><div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 800, fontSize: 15 }}>MediaDesk</div><div style={{ fontSize: 9, color: "var(--text3)" }}>PRO PLATFORM · {formatRoleLabel(user?.role)}</div></div>}
      </div>
      <nav style={{ flex: 1, padding: "10px 6px", display: "flex", flexDirection: "column", gap: 2 }}>
        {nav.map(n => {
          const active = page === n.id;
          return (
            <button key={n.id} onClick={() => setPage(n.id)}
              style={{ display: "flex", alignItems: "center", gap: 11, padding: collapsed ? "9px 10px" : "9px 13px", borderRadius: 9, border: "none", background: active ? "rgba(240,165,0,.12)" : "transparent", color: active ? "var(--accent)" : "var(--text2)", cursor: "pointer", fontSize: 13, fontWeight: active ? 600 : 400, textAlign: "left", transition: "all .14s", whiteSpace: "nowrap", justifyContent: collapsed ? "center" : "flex-start", borderLeft: active ? "2px solid var(--accent)" : "2px solid transparent" }}
              onMouseEnter={e => { if (!active) e.currentTarget.style.background = "var(--bg3)"; }}
              onMouseLeave={e => { if (!active) e.currentTarget.style.background = "transparent"; }}>
              <span style={{ fontSize: 17, flexShrink: 0 }}>{n.icon}</span>
              {!collapsed && <><span>{n.label}</span>{n.badge ? <span style={{ marginLeft: "auto", background: "var(--accent)", color: "#111", minWidth: 18, height: 18, borderRadius: 999, display: "inline-flex", alignItems: "center", justifyContent: "center", fontSize: 10, fontWeight: 800, padding: "0 6px" }}>{n.badge > 99 ? "99+" : n.badge}</span> : null}</>}
            </button>
          );
        })}
      </nav>

      <button onClick={() => toggleTheme && toggleTheme()}
        title={theme === "light" ? "Switch to Dark Mode" : "Switch to Light Mode"}
        style={{ margin: "4px 6px 2px", padding: "7px", background: theme === "dark" ? "rgba(240,165,0,.12)" : "rgba(59,126,245,.1)", border: `1px solid ${theme === "dark" ? "rgba(240,165,0,.3)" : "rgba(59,126,245,.25)"}`, borderRadius: 7, color: theme === "dark" ? "var(--accent)" : "var(--blue)", cursor: "pointer", fontSize: 14, display: "flex", alignItems: "center", justifyContent: collapsed ? "center" : "flex-start", gap: 8, width: "calc(100% - 12px)" }}>
        {theme === "light" ? "🌙" : "☀️"}
        {!collapsed && <span style={{ fontSize: 11, fontWeight: 600 }}>{theme === "light" ? "Dark Mode" : "Light Mode"}</span>}
      </button>
      <button onClick={() => setCollapsed(!collapsed)} style={{ margin: "4px 6px 6px", padding: "7px", background: "var(--bg3)", border: "1px solid var(--border)", borderRadius: 7, color: "var(--text3)", cursor: "pointer", fontSize: 12, display: "flex", alignItems: "center", justifyContent: "center" }}>{collapsed ? "▶" : "◀"}</button>

      {/* User strip */}
      <div style={{ padding: "10px 10px 14px", borderTop: "1px solid var(--border)", display: "flex", alignItems: "center", gap: 9 }}>
        <div onClick={() => setPage("settings")} style={{ width: 32, height: 32, background: "linear-gradient(135deg,var(--accent),var(--purple))", borderRadius: "50%", display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 700, fontSize: 13, flexShrink: 0, color: "#000", cursor: "pointer" }} title="My Profile">{user.name?.[0]?.toUpperCase() || "U"}</div>
        {!collapsed && <><div style={{ flex: 1, overflow: "hidden" }}><div style={{ fontSize: 12, fontWeight: 600, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{user.name}</div><div style={{ fontSize: 10, color: "var(--text3)", overflow: "hidden", textOverflow: "ellipsis" }}>{user.title ? user.title : user.agency}</div></div><button onClick={onLogout} title="Sign Out" style={{ background: "none", border: "none", color: "var(--text3)", cursor: "pointer", fontSize: 15, flexShrink: 0 }}>⎋</button></>}
      </div>
    </div>
  );
};


const TopRightNotificationsButton = ({ count = 0, onClick }) => (
  <button
    onClick={onClick}
    title="Workspace alerts"
    style={{
      position: "fixed",
      top: 18,
      right: 22,
      zIndex: 80,
      width: 46,
      height: 46,
      borderRadius: 999,
      border: "1px solid var(--border2)",
      background: "var(--bg2)",
      color: "var(--text)",
      boxShadow: "0 12px 28px rgba(0,0,0,.18)",
      cursor: "pointer",
      display: "flex",
      alignItems: "center",
      justifyContent: "center",
      fontSize: 19,
      transition: "transform .16s ease, box-shadow .16s ease",
    }}
    onMouseEnter={e => { e.currentTarget.style.transform = "translateY(-2px)"; e.currentTarget.style.boxShadow = "0 16px 34px rgba(0,0,0,.22)"; }}
    onMouseLeave={e => { e.currentTarget.style.transform = "translateY(0)"; e.currentTarget.style.boxShadow = "0 12px 28px rgba(0,0,0,.18)"; }}
  >
    🔔
    {count > 0 && (
      <span style={{
        position: "absolute",
        top: -4,
        right: -4,
        minWidth: 20,
        height: 20,
        borderRadius: 999,
        padding: "0 6px",
        background: "var(--accent)",
        color: "#111",
        display: "inline-flex",
        alignItems: "center",
        justifyContent: "center",
        fontSize: 10,
        fontWeight: 800,
        fontFamily: "'Syne',sans-serif",
      }}>
        {count > 99 ? "99+" : count}
      </span>
    )}
  </button>
);

/* ── DASHBOARD ──────────────────────────────────────────── */
const Dashboard = ({ user, vendors, clients, campaigns, rates, mpos, notifications, unreadNotifications, setPage, onOpenNotifications }) => {
  const liveVendors = activeOnly(vendors);
  const liveClients = activeOnly(clients);
  const liveCampaigns = activeOnly(campaigns);
  const liveMpos = activeOnly(mpos);
  const totalBudget = liveCampaigns.reduce((s, c) => s + (parseFloat(c.budget) || 0), 0);
  const totalMPOValue = liveMpos.reduce((s, m) => s + (m.netVal || 0), 0);
  const pendingApprovals = liveMpos.filter(m => ["draft", "submitted", "reviewed"].includes(m.status || "draft")).length;
  const unreconciledCount = liveMpos.filter(m => (m.reconciliationStatus || "not_started") !== "completed").length;
  const pendingPaymentCount = liveMpos.filter(m => ["received", "approved", "disputed"].includes(m.invoiceStatus || "pending") && (m.paymentStatus || "unpaid") !== "paid").length;
  const pendingProofCount = liveMpos.filter(m => !["received"].includes(m.proofStatus || "pending")).length;
  const recent = [...liveMpos].sort((a, b) => b.createdAt - a.createdAt).slice(0, 4);
  const myWorkflowQueue = [...liveMpos].filter(m => isMpoAwaitingUser(user, m)).sort((a, b) => (b.updatedAt || b.createdAt || 0) - (a.updatedAt || a.createdAt || 0)).slice(0, 4);
  const readyForDispatch = liveMpos.filter(m => String(m.status || "draft").toLowerCase() === "approved").length;
  const needsRevision = liveMpos.filter(m => String(m.status || "draft").toLowerCase() === "rejected").length;
  const recentNotifications = (notifications || []).slice(0, 5);
  const topVendors = Object.values(liveMpos.reduce((acc, mpo) => {
    const key = mpo.vendorName || "Unknown Vendor";
    acc[key] = acc[key] || { name: key, spend: 0, count: 0 };
    acc[key].spend += parseFloat(mpo.netVal) || 0;
    acc[key].count += 1;
    return acc;
  }, {})).sort((a, b) => b.spend - a.spend).slice(0, 4);
  const budgetUtilization = totalBudget > 0 ? Math.min(100, Math.round((totalMPOValue / totalBudget) * 100)) : 0;
  return (
    <div className="fade">
      <div style={{ marginBottom: 28 }}>
        <h1 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 800, fontSize: 26, letterSpacing: "-.03em" }}>Good {new Date().getHours() < 12 ? "morning" : new Date().getHours() < 17 ? "afternoon" : "evening"}, <span style={{ color: "var(--accent)" }}>{user.name?.split(" ")[0]}</span> 👋</h1>
        <p style={{ color: "var(--text2)", marginTop: 5 }}>{user.agency} — Media Schedule Platform</p>
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(160px,1fr))", gap: 14, marginBottom: 24 }}>
        <Stat icon="🏢" label="Vendors" value={liveVendors.length} sub={`${archivedOnly(vendors).length} archived`} />
        <Stat icon="👥" label="Clients" value={liveClients.length} sub="Brands & orgs" color="var(--blue)" />
        <Stat icon="📢" label="Campaigns" value={liveCampaigns.length} sub={`${liveCampaigns.filter(c => c.status === "active").length} active`} color="var(--green)" />
        <Stat icon="📄" label="MPOs Issued" value={liveMpos.length} sub={`${pendingApprovals} pending approval`} color="var(--purple)" />
        <Stat icon="💰" label="MPO Value" value={`₦${(totalMPOValue / 1e6).toFixed(1)}M`} sub={`Budget pool ${fmtN(totalBudget)}`} color="var(--teal)" />
        <Stat icon="🔔" label="Unread Alerts" value={unreadNotifications} sub={`${pendingPaymentCount} awaiting payment`} color="var(--orange)" />
        <Stat icon="✅" label="My Queue" value={myWorkflowQueue.length} sub={`${readyForDispatch} approved · ${needsRevision} need changes`} color="var(--blue)" />
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "1.2fr .95fr", gap: 18 }}>
        <div style={{ display: "flex", flexDirection: "column", gap: 18 }}>
          <Card>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 18 }}>
              <h2 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 15 }}>Operational Watchlist</h2>
              <Btn variant="ghost" size="sm" onClick={() => setPage("mpo")}>Open MPO Workspace →</Btn>
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(180px,1fr))", gap: 12 }}>
              <div style={{ padding: "14px 16px", borderRadius: 12, background: "var(--bg3)", border: "1px solid var(--border)" }}>
                <div style={{ fontSize: 12, color: "var(--text3)", marginBottom: 6 }}>Pending Approvals</div>
                <div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 800, fontSize: 24 }}>{pendingApprovals}</div>
                <div style={{ fontSize: 12, color: "var(--text2)", marginTop: 6 }}>Draft, submitted, and reviewed MPOs</div>
              </div>
              <div style={{ padding: "14px 16px", borderRadius: 12, background: "var(--bg3)", border: "1px solid var(--border)" }}>
                <div style={{ fontSize: 12, color: "var(--text3)", marginBottom: 6 }}>Needs Reconciliation</div>
                <div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 800, fontSize: 24 }}>{unreconciledCount}</div>
                <div style={{ fontSize: 12, color: "var(--text2)", marginTop: 6 }}>Execution not fully reconciled yet</div>
              </div>
              <div style={{ padding: "14px 16px", borderRadius: 12, background: "var(--bg3)", border: "1px solid var(--border)" }}>
                <div style={{ fontSize: 12, color: "var(--text3)", marginBottom: 6 }}>Awaiting Payment</div>
                <div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 800, fontSize: 24 }}>{pendingPaymentCount}</div>
                <div style={{ fontSize: 12, color: "var(--text2)", marginTop: 6 }}>Invoices received/approved but not paid</div>
              </div>
              <div style={{ padding: "14px 16px", borderRadius: 12, background: "var(--bg3)", border: "1px solid var(--border)" }}>
                <div style={{ fontSize: 12, color: "var(--text3)", marginBottom: 6 }}>Awaiting Proof</div>
                <div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 800, fontSize: 24 }}>{pendingProofCount}</div>
                <div style={{ fontSize: 12, color: "var(--text2)", marginTop: 6 }}>Proof of airing still outstanding</div>
              </div>
            </div>
            <div style={{ marginTop: 16 }}>
              <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12, color: "var(--text2)", marginBottom: 6 }}>
                <span>Budget Utilization</span>
                <strong style={{ color: "var(--text)" }}>{budgetUtilization}%</strong>
              </div>
              <div style={{ height: 10, background: "var(--bg2)", borderRadius: 999, overflow: "hidden", border: "1px solid var(--border)" }}>
                <div style={{ width: `${budgetUtilization}%`, height: "100%", background: "linear-gradient(90deg,var(--accent),var(--purple))" }} />
              </div>
            </div>
          </Card>
          <Card>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 18 }}>
              <h2 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 15 }}>Recent MPOs</h2>
              <Btn variant="ghost" size="sm" onClick={() => setPage("mpo")}>View all →</Btn>
            </div>
            {recent.length === 0 ? <Empty icon="📄" title="No MPOs yet" sub="Generate your first MPO to see it here" /> :
              <div style={{ display: "flex", flexDirection: "column", gap: 9 }}>
                {recent.map(m => (
                  <div key={m.id} style={{ display: "flex", alignItems: "center", gap: 12, padding: "11px 14px", background: "var(--bg3)", borderRadius: 10, border: "1px solid var(--border)" }}>
                    <div style={{ width: 38, height: 38, background: "rgba(240,165,0,.12)", borderRadius: 9, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 17, flexShrink: 0 }}>📄</div>
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <div style={{ fontWeight: 600, fontSize: 13, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{m.mpoNo || "MPO"} — {m.vendorName}</div>
                      <div style={{ fontSize: 11, color: "var(--text2)" }}>{m.clientName} · {m.month} {m.year}</div>
                    </div>
                    <div style={{ textAlign: "right", flexShrink: 0 }}>
                      <div style={{ fontWeight: 700, color: "var(--accent)", fontSize: 13 }}>{fmtN(m.netVal)}</div>
                      <Badge color={m.status === "approved" ? "green" : m.status === "sent" ? "blue" : "accent"}>{m.status || "draft"}</Badge>
                    </div>
                  </div>
                ))}
              </div>}
          </Card>
        </div>
        <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
          <Card>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 14 }}>
              <h2 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 15 }}>My Workflow Queue</h2>
              <Btn variant="ghost" size="sm" onClick={() => setPage("mpo")}>Open MPOs →</Btn>
            </div>
            {myWorkflowQueue.length === 0 ? <Empty icon="✅" title="Nothing waiting on you" sub="Approvals and dispatch work assigned to your role will appear here." /> : (
              <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                {myWorkflowQueue.map(mpo => {
                  const workflowMeta = getMpoWorkflowMeta(mpo);
                  return (
                    <button key={mpo.id} onClick={() => setPage("mpo")} style={{ textAlign: "left", border: "1px solid var(--border)", background: "var(--bg3)", borderRadius: 10, padding: "10px 12px", cursor: "pointer" }}>
                      <div style={{ display: "flex", justifyContent: "space-between", gap: 8, alignItems: "center" }}>
                        <div style={{ fontSize: 13, fontWeight: 700 }}>{mpo.mpoNo || "MPO"} · {mpo.vendorName || "Vendor"}</div>
                        <Badge color={workflowMeta.color}>Waiting on you</Badge>
                      </div>
                      <div style={{ fontSize: 12, color: "var(--text2)", marginTop: 5 }}>{workflowMeta.hint}</div>
                      <div style={{ fontSize: 11, color: "var(--text3)", marginTop: 7 }}>{MPO_STATUS_LABELS[mpo.status || "draft"] || (mpo.status || "draft")} · {mpo.clientName || "No client"}</div>
                    </button>
                  );
                })}
              </div>
            )}
          </Card>
          <Card>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 14 }}>
              <h2 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 15 }}>Notifications</h2>
              <Btn variant="ghost" size="sm" onClick={onOpenNotifications}>Open inbox →</Btn>
            </div>
            {recentNotifications.length === 0 ? <Empty icon="🔔" title="No alerts yet" sub="Workflow notifications will appear here." /> : (
              <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                {recentNotifications.map(notification => (
                  <button key={notification.id} onClick={onOpenNotifications} style={{ textAlign: "left", border: "1px solid var(--border)", background: notification.readAt ? "var(--bg3)" : "rgba(240,165,0,.07)", borderRadius: 10, padding: "10px 12px", cursor: "pointer" }}>
                    <div style={{ display: "flex", justifyContent: "space-between", gap: 8 }}>
                      <div style={{ fontSize: 13, fontWeight: 700 }}>{notification.title}</div>
                      {!notification.readAt ? <Badge color="accent">New</Badge> : null}
                    </div>
                    <div style={{ fontSize: 12, color: "var(--text2)", marginTop: 5 }}>{notification.message || "Open settings notifications to review this alert."}</div>
                    <div style={{ fontSize: 11, color: "var(--text3)", marginTop: 7 }}>{formatAuditTimestamp(notification.createdAt)}</div>
                  </button>
                ))}
              </div>
            )}
          </Card>
          <Card>
            <h2 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 15, marginBottom: 14 }}>Top Vendor Spend</h2>
            {topVendors.length === 0 ? <Empty icon="💼" title="No spend yet" sub="Vendor spend distribution appears once MPOs are issued." /> : (
              <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                {topVendors.map((vendor, idx) => (
                  <div key={vendor.name} style={{ display: "flex", justifyContent: "space-between", gap: 10, alignItems: "center", borderBottom: idx === topVendors.length - 1 ? "none" : "1px solid var(--border)", paddingBottom: 10 }}>
                    <div>
                      <div style={{ fontWeight: 700, fontSize: 13 }}>{vendor.name}</div>
                      <div style={{ fontSize: 11, color: "var(--text2)" }}>{vendor.count} MPO{vendor.count !== 1 ? "s" : ""}</div>
                    </div>
                    <div style={{ fontWeight: 700, color: "var(--accent)" }}>{fmtN(vendor.spend)}</div>
                  </div>
                ))}
              </div>
            )}
          </Card>
          <Card style={{ background: "linear-gradient(135deg,rgba(240,165,0,.1),rgba(139,92,246,.07))", border: "1px solid rgba(240,165,0,.18)" }}>
            <div style={{ fontSize: 22, marginBottom: 9 }}>📡</div>
            <div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 14, marginBottom: 4 }}>{user.agency}</div>
            <div style={{ fontSize: 11, color: "var(--text2)", marginBottom: 8 }}>{user.email}</div>
            <div style={{ display: "inline-flex", alignItems: "center", gap: 8, padding: "6px 10px", borderRadius: 999, background: "rgba(15,23,42,.08)", fontSize: 11, fontWeight: 700 }}>
              <span>{formatRoleLabel(user.role)}</span>
              {unreadNotifications > 0 ? <span>• {unreadNotifications} unread</span> : <span>• Inbox clear</span>}
            </div>
          </Card>
        </div>
      </div>
    </div>
  );
};

/* ── VENDORS ────────────────────────────────────────────── */
const VendorsPage = ({ vendors, setVendors, user }) => {
  const [modal, setModal] = useState(null);
  const [search, setSearch] = useState("");
  const [viewMode, setViewMode] = useState("active");
  const [toast, setToast] = useState(null);
  const [confirm, setConfirm] = useState(null);
  const blank = { name: "", type: "", contact: "", email: "", phone: "", location: "", rate: "", discount: "", commission: "", notes: "" };
  const [f, setF] = useState(blank);
  const u = k => v => setF(p => ({ ...p, [k]: v }));
  const mediaTypes = ["Television","Radio","Print","Digital/Online","Out-of-Home (OOH)","Cinema","Podcast","Social Media"];
  const canManage = hasPermission(user, "manageVendors");
  const save = async () => {
    if (!canManage) return setToast({ msg: readOnlyMessage(user), type: "error" });
  if (!f.name || !f.type) {
    return setToast({ msg: "Name and type required.", type: "error" });
  }

  try {
    if (modal === "add") {
      const newVendor = await createVendorInSupabase(user.agencyId, user.id, f);
      setVendors(v => [newVendor, ...v]);
    } else {
      const updatedVendor = await updateVendorInSupabase(modal.id, f);
      setVendors(v => v.map(x => x.id === modal.id ? updatedVendor : x));
    }

    setToast({
      msg: modal === "add" ? "Vendor added." : "Vendor updated.",
      type: "success",
    });
    setModal(null);
    setF(blank);
  } catch (e) {
    setToast({
      msg: e.message || "Failed to save vendor.",
      type: "error",
    });
  }
};

const del = async (id) => {
  if (!canManage) return setToast({ msg: readOnlyMessage(user), type: "error" });
  try {
    const archivedVendor = await archiveVendorInSupabase(id);
    setVendors(v => v.map(x => x.id === id ? archivedVendor : x));
    setToast({ msg: "Vendor archived.", type: "success" });
    setConfirm(null);
  } catch (e) {
    setToast({
      msg: e.message || "Failed to archive vendor.",
      type: "error",
    });
  }
};
const restore = async (id) => {
  if (!canManage) return setToast({ msg: readOnlyMessage(user), type: "error" });
  try {
    const restoredVendor = await restoreVendorInSupabase(id);
    setVendors(v => v.map(x => x.id === id ? restoredVendor : x));
    setToast({ msg: "Vendor restored.", type: "success" });
  } catch (e) {
    setToast({ msg: e.message || "Failed to restore vendor.", type: "error" });
  }
};
  const visible = viewMode === "archived" ? archivedOnly(vendors) : viewMode === "all" ? vendors : activeOnly(vendors);
  const filtered = visible.filter(v => `${v.name} ${v.type}`.toLowerCase().includes(search.toLowerCase()));
  const typeIcon = t => ({ Television: "📺", Radio: "📻", Print: "📰", "Digital/Online": "💻", "Out-of-Home (OOH)": "🪧", Cinema: "🎬", Podcast: "🎙", "Social Media": "📱" }[t] || "📡");
  const typeColor = t => ({ Television: "accent", Radio: "blue", Print: "green", "Digital/Online": "purple", "Out-of-Home (OOH)": "teal", Cinema: "red", Podcast: "orange", "Social Media": "purple" }[t] || "accent");
  return (
    <div className="fade">
      {toast && <Toast msg={toast.msg} type={toast.type} onDone={() => setToast(null)} />}
      {confirm && <Confirm msg={confirm.msg} onYes={confirm.onYes} onNo={() => setConfirm(null)} />}
      {!canManage && <Card style={{ marginBottom: 14, background: "rgba(59,126,245,.06)", border: "1px solid rgba(59,126,245,.18)" }}><div style={{ fontSize: 13, color: "var(--text2)" }}>You have read-only access to Clients as <strong>{formatRoleLabel(user?.role)}</strong>.</div></Card>}
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 24, flexWrap: "wrap", gap: 10 }}>
        <div><h1 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 800, fontSize: 24 }}>Vendors</h1><p style={{ color: "var(--text2)", marginTop: 3, fontSize: 13 }}>Media houses & suppliers</p></div>
        {canManage && <Btn icon="+" onClick={() => { setF(blank); setModal("add"); }}>Add Vendor</Btn>}
      </div>
      <div style={{ display: "flex", gap: 10, marginBottom: 18, flexWrap: "wrap" }}><div style={{ flex: 1, minWidth: 180 }}><Field value={search} onChange={setSearch} placeholder="Search vendors…" /></div><Field value={viewMode} onChange={setViewMode} options={[{value:"active",label:"Active"},{value:"archived",label:"Archived"},{value:"all",label:"All"}]} /></div>
      {filtered.length === 0 ? <Card><Empty icon="🏢" title="No vendors found" sub="Add your first media house" /></Card> :
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill,minmax(300px,1fr))", gap: 14 }}>
          {filtered.map(v => (
            <Card key={v.id}>
              <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", marginBottom: 12 }}>
                <div style={{ display: "flex", alignItems: "center", gap: 11 }}>
                  <div style={{ width: 42, height: 42, background: "var(--bg3)", borderRadius: 11, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 20 }}>{typeIcon(v.type)}</div>
                  <div><div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 14 }}>{v.name}</div><div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginTop: 3 }}><Badge color={typeColor(v.type)}>{v.type}</Badge>{isArchived(v) && <Badge color="red">Archived</Badge>}</div></div>
                </div>
                <div style={{ display: "flex", gap: 5 }}>
                  {canManage && <Btn variant="ghost" size="sm" onClick={() => { setF({ name: v.name, type: v.type, contact: v.contact||"", email: v.email||"", phone: v.phone||"", location: v.location||"", rate: v.rate||"", discount: v.discount||"", commission: v.commission||"", notes: v.notes||"" }); setModal(v); }}>✏️</Btn>}
                  {canManage && (isArchived(v) ? <Btn variant="success" size="sm" onClick={() => restore(v.id)}>↩</Btn> : <Btn variant="danger" size="sm" onClick={() => setConfirm({ msg: `Archive "${v.name}"? Existing reports and MPO links will stay intact.`, onYes: () => del(v.id) })}>🗄</Btn>)}
                </div>
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 7 }}>
                {[["Rate/Spot", v.rate ? fmtN(v.rate) : "—"], ["Vol Disc.", v.discount ? `${v.discount}%` : "—"], ["Comm.", v.commission ? `${v.commission}%` : "—"], ["Contact", v.contact || "—"]].map(([l, val]) => (
                  <div key={l} style={{ background: "var(--bg3)", borderRadius: 7, padding: "7px 11px" }}>
                    <div style={{ fontSize: 9, color: "var(--text3)", fontWeight: 600, textTransform: "uppercase", letterSpacing: ".07em" }}>{l}</div>
                    <div style={{ fontSize: 12, fontWeight: 600, marginTop: 2, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{val}</div>
                  </div>
                ))}
              </div>
              {v.notes && <div style={{ marginTop: 10, fontSize: 11, color: "var(--text2)", borderTop: "1px solid var(--border)", paddingTop: 9 }}>{v.notes}</div>}
            </Card>
          ))}
        </div>}
      {modal !== null && (
        <Modal title={modal === "add" ? "Add Vendor" : `Edit: ${modal.name}`} onClose={() => setModal(null)} width={560}>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 }}>
            <div style={{ gridColumn: "1/-1" }}><Field label="Vendor Name" value={f.name} onChange={u("name")} placeholder="NTA, Channels TV…" required /></div>
            <Field label="Media Type" value={f.type} onChange={u("type")} options={mediaTypes} placeholder="Select type" required />
            <Field label="Contact Person" value={f.contact} onChange={u("contact")} placeholder="Name" />
            <Field label="Email" type="email" value={f.email} onChange={u("email")} placeholder="vendor@example.com" />
            <Field label="Phone" value={f.phone} onChange={u("phone")} placeholder="+234 800 000 0000" />
            <div style={{ gridColumn: "1/-1" }}><Field label="Location / City" value={f.location} onChange={u("location")} placeholder="Lagos, Abuja, Port Harcourt…" note="State or city where vendor operates" /></div>
            <Field label="Default Rate/Spot (₦)" type="number" value={f.rate} onChange={u("rate")} placeholder="0" />
            <Field label="Volume Discount (%)" type="number" value={f.discount} onChange={u("discount")} placeholder="e.g. 27" />
            <div style={{ gridColumn: "1/-1" }}><Field label="Agency Commission (%)" type="number" value={f.commission} onChange={u("commission")} placeholder="e.g. 15" /></div>
            <div style={{ gridColumn: "1/-1" }}><Field label="Notes" value={f.notes} onChange={u("notes")} placeholder="Additional notes…" /></div>
          </div>
          <div style={{ display: "flex", gap: 10, justifyContent: "flex-end", marginTop: 20 }}>
            <Btn variant="ghost" onClick={() => setModal(null)}>Cancel</Btn>
            <Btn onClick={save}>{modal === "add" ? "Add Vendor" : "Save Changes"}</Btn>
          </div>
        </Modal>
      )}
    </div>
  );
};

/* ── CLIENTS ────────────────────────────────────────────── */
const ClientsPage = ({ clients, setClients, user }) => {
  const canManage = hasPermission(user, "manageClients");
  const [modal, setModal] = useState(null); const [search, setSearch] = useState(""); const [viewMode, setViewMode] = useState("active"); const [toast, setToast] = useState(null); const [confirm, setConfirm] = useState(null);
  const blank = { name: "", industry: "", contact: "", email: "", phone: "", address: "", brands: "" };
  const [f, setF] = useState(blank); const u = k => v => setF(p => ({ ...p, [k]: v }));
  const industries = ["Healthcare/Pharma","FMCG","Banking/Finance","Telecoms","Government/NGO","Education","Real Estate","Energy/Oil & Gas","Retail","Technology","Media/Entertainment","Food & Beverage","Automotive","Other"];
  const save = async () => {
    if (!canManage) return setToast({ msg: readOnlyMessage(user), type: "error" });
    if (!f.name) return setToast({ msg: "Client name required.", type: "error" });
    const duplicate = clients.find(c => c.id !== modal?.id && !isArchived(c) && c.name.trim().toLowerCase() === f.name.trim().toLowerCase());
    if (duplicate) return setToast({ msg: "A client with this name already exists.", type: "error" });
    try {
      if (modal === "add") {
        const newClient = await createClientInSupabase(user.agencyId, user.id, f);
        setClients(v => [newClient, ...v]);
        createAuditEventInSupabase({ agencyId: user.agencyId, recordType: "client", recordId: newClient.id, action: "created", actor: user, metadata: { name: newClient.name || "" } }).catch(error => console.error("Failed to write audit event:", error));
      } else {
        const updatedClient = await updateClientInSupabase(modal.id, f);
        setClients(v => v.map(x => x.id === modal.id ? updatedClient : x));
        createAuditEventInSupabase({ agencyId: user.agencyId, recordType: "client", recordId: modal.id, action: "updated", actor: user, metadata: { name: updatedClient.name || "" } }).catch(error => console.error("Failed to write audit event:", error));
      }
      setToast({ msg: modal === "add" ? "Client added." : "Client updated.", type: "success" });
      setModal(null);
      setF(blank);
    } catch (e) {
      setToast({ msg: e.message || "Failed to save client.", type: "error" });
    }
  };
  const del = async id => {
    if (!canManage) return setToast({ msg: readOnlyMessage(user), type: "error" });
    if (!canManage) return setToast({ msg: readOnlyMessage(user), type: "error" });
    try {
      const archivedClient = await archiveClientInSupabase(id);
      setClients(v => v.map(x => x.id === id ? archivedClient : x));
      createAuditEventInSupabase({ agencyId: user.agencyId, recordType: "client", recordId: id, action: "archived", actor: user, metadata: { name: archivedClient.name || "" } }).catch(error => console.error("Failed to write audit event:", error));
      setToast({ msg: "Client archived.", type: "success" });
      setConfirm(null);
    } catch (e) {
      setToast({ msg: e.message || "Failed to archive client.", type: "error" });
    }
  };
  const restore = async id => {
    if (!canManage) return setToast({ msg: readOnlyMessage(user), type: "error" });
    if (!canManage) return setToast({ msg: readOnlyMessage(user), type: "error" });
    try {
      const restoredClient = await restoreClientInSupabase(id);
      setClients(v => v.map(x => x.id === id ? restoredClient : x));
      createAuditEventInSupabase({ agencyId: user.agencyId, recordType: "client", recordId: id, action: "restored", actor: user, metadata: { name: restoredClient.name || "" } }).catch(error => console.error("Failed to write audit event:", error));
      setToast({ msg: "Client restored.", type: "success" });
    } catch (e) {
      setToast({ msg: e.message || "Failed to restore client.", type: "error" });
    }
  };
  const visible = viewMode === "archived" ? archivedOnly(clients) : viewMode === "all" ? clients : activeOnly(clients);
  const filtered = visible.filter(c => `${c.name} ${c.industry}`.toLowerCase().includes(search.toLowerCase()));
  return (
    <div className="fade">
      {toast && <Toast msg={toast.msg} type={toast.type} onDone={() => setToast(null)} />}
      {confirm && <Confirm msg={confirm.msg} onYes={confirm.onYes} onNo={() => setConfirm(null)} />}
      {!canManage && <Card style={{ marginBottom: 14, background: "rgba(59,126,245,.06)", border: "1px solid rgba(59,126,245,.18)" }}><div style={{ fontSize: 13, color: "var(--text2)" }}>You have read-only access to Campaigns as <strong>{formatRoleLabel(user?.role)}</strong>.</div></Card>}
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 24, flexWrap: "wrap", gap: 10 }}>
        <div><h1 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 800, fontSize: 24 }}>Clients & Brands</h1><p style={{ color: "var(--text2)", marginTop: 3, fontSize: 13 }}>Manage your client portfolio</p></div>
        {canManage && <Btn icon="+" onClick={() => { setF(blank); setModal("add"); }}>Add Client</Btn>}
      </div>
      <div style={{ display: "flex", gap: 10, marginBottom: 18, flexWrap: "wrap" }}><div style={{ flex: 1, minWidth: 180 }}><Field value={search} onChange={setSearch} placeholder="Search clients…" /></div><Field value={viewMode} onChange={setViewMode} options={[{value:"active",label:"Active"},{value:"archived",label:"Archived"},{value:"all",label:"All"}]} /></div>
      {filtered.length === 0 ? <Card><Empty icon="👥" title="No clients yet" sub="Add your first client" /></Card> :
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill,minmax(280px,1fr))", gap: 14 }}>
          {filtered.map(c => {
            const brands = c.brands ? c.brands.split(",").map(b => b.trim()).filter(Boolean) : [];
            return (
              <Card key={c.id}>
                <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", marginBottom: 12 }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 11 }}>
                    <div style={{ width: 42, height: 42, background: `hsl(${c.name.charCodeAt(0)*7},40%,18%)`, borderRadius: 11, display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "'Syne',sans-serif", fontWeight: 800, fontSize: 17, color: `hsl(${c.name.charCodeAt(0)*7},70%,65%)` }}>{c.name[0]?.toUpperCase()}</div>
                    <div><div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 14 }}>{c.name}</div><div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginTop: 3 }}>{c.industry && <Badge color="blue">{c.industry}</Badge>}{isArchived(c) && <Badge color="red">Archived</Badge>}</div></div>
                  </div>
                  <div style={{ display: "flex", gap: 5 }}>
                    <Btn variant="ghost" size="sm" onClick={() => { setF({ name: c.name, industry: c.industry||"", contact: c.contact||"", email: c.email||"", phone: c.phone||"", address: c.address||"", brands: c.brands||"" }); setModal(c); }}>✏️</Btn>
                    {isArchived(c) ? <Btn variant="success" size="sm" onClick={() => restore(c.id)}>↩</Btn> : <Btn variant="danger" size="sm" onClick={() => setConfirm({ msg: `Archive "${c.name}"? Campaign history will be retained.`, onYes: () => del(c.id) })}>🗄</Btn>}
                  </div>
                </div>
                {brands.length > 0 && <div style={{ display: "flex", flexWrap: "wrap", gap: 5, marginBottom: 10 }}>{brands.map(b => <Badge key={b} color="green">{b}</Badge>)}</div>}
                {(c.email || c.phone) && <div style={{ fontSize: 11, color: "var(--text2)", borderTop: "1px solid var(--border)", paddingTop: 9 }}>{c.email && <div>📧 {c.email}</div>}{c.phone && <div style={{ marginTop: 2 }}>📞 {c.phone}</div>}</div>}
              </Card>
            );
          })}
        </div>}
      {modal !== null && (
        <Modal title={modal === "add" ? "Add Client" : `Edit: ${modal.name}`} onClose={() => setModal(null)} width={540}>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 }}>
            <div style={{ gridColumn: "1/-1" }}><Field label="Client / Organization Name" value={f.name} onChange={u("name")} placeholder="Breakthrough Action Nigeria" required /></div>
            <Field label="Industry" value={f.industry} onChange={u("industry")} options={industries} placeholder="Select industry" />
            <Field label="Contact Person" value={f.contact} onChange={u("contact")} placeholder="Primary contact" />
            <Field label="Email" type="email" value={f.email} onChange={u("email")} placeholder="client@org.com" />
            <Field label="Phone" value={f.phone} onChange={u("phone")} placeholder="+234 800 000 0000" />
            <div style={{ gridColumn: "1/-1" }}><Field label="Brands (comma-separated)" value={f.brands} onChange={u("brands")} placeholder="Brand A, Brand B" note="Separate multiple brands with commas" /></div>
            <div style={{ gridColumn: "1/-1" }}><Field label="Address" value={f.address} onChange={u("address")} placeholder="Physical address" /></div>
          </div>
          <div style={{ display: "flex", gap: 10, justifyContent: "flex-end", marginTop: 20 }}>
            <Btn variant="ghost" onClick={() => setModal(null)}>Cancel</Btn>
            <Btn onClick={save}>{modal === "add" ? "Add Client" : "Save Changes"}</Btn>
          </div>
        </Modal>
      )}
    </div>
  );
};

/* ── CAMPAIGNS ──────────────────────────────────────────── */
const CampaignsPage = ({ campaigns, setCampaigns, clients, user }) => {
  const canManage = hasPermission(user, "manageCampaigns");
  const [modal, setModal] = useState(null); const [search, setSearch] = useState(""); const [filterStatus, setFilterStatus] = useState("all"); const [viewMode, setViewMode] = useState("active"); const [toast, setToast] = useState(null); const [confirm, setConfirm] = useState(null);
  const blank = { name: "", clientId: "", brand: "", objective: "", startDate: "", endDate: "", budget: "", status: "planning", medium: "", notes: "", materialList: [] };
  const [f, setF] = useState(blank); const u = k => v => setF(p => ({ ...p, [k]: v }));
  const save = async () => {
    if (!canManage) return setToast({ msg: readOnlyMessage(user), type: "error" });
    if (!f.name || !f.clientId) return setToast({ msg: "Campaign name and client required.", type: "error" });
    if (f.startDate && f.endDate && new Date(f.endDate) < new Date(f.startDate)) return setToast({ msg: "End date cannot be before start date.", type: "error" });
    if (f.budget && (parseFloat(f.budget) || 0) < 0) return setToast({ msg: "Budget cannot be negative.", type: "error" });
    const duplicate = campaigns.find(c => c.id !== modal?.id && !isArchived(c) && c.clientId === f.clientId && c.name.trim().toLowerCase() === f.name.trim().toLowerCase());
    if (duplicate) return setToast({ msg: "A campaign with this client and title already exists.", type: "error" });
    try {
      if (modal === "add") {
        const newCampaign = await createCampaignInSupabase(user.agencyId, user.id, f);
        setCampaigns(v => [newCampaign, ...v]);
        createAuditEventInSupabase({ agencyId: user.agencyId, recordType: "campaign", recordId: newCampaign.id, action: "created", actor: user, metadata: { name: newCampaign.name || "", status: newCampaign.status || "planning" } }).catch(error => console.error("Failed to write audit event:", error));
      } else {
        const updatedCampaign = await updateCampaignInSupabase(modal.id, f);
        setCampaigns(v => v.map(x => x.id === modal.id ? updatedCampaign : x));
        createAuditEventInSupabase({ agencyId: user.agencyId, recordType: "campaign", recordId: modal.id, action: "updated", actor: user, metadata: { name: updatedCampaign.name || "", status: updatedCampaign.status || "planning" } }).catch(error => console.error("Failed to write audit event:", error));
      }
      setToast({ msg: modal === "add" ? "Campaign created." : "Campaign updated.", type: "success" }); setModal(null); setF(blank);
    } catch (e) {
      setToast({ msg: e.message || "Failed to save campaign.", type: "error" });
    }
  };
  const del = async id => {
    if (!canManage) return setToast({ msg: readOnlyMessage(user), type: "error" });
    if (!canManage) return setToast({ msg: readOnlyMessage(user), type: "error" });
    try {
      const archived = await archiveCampaignInSupabase(id);
      setCampaigns(v => v.map(x => x.id === id ? archived : x));
      createAuditEventInSupabase({ agencyId: user.agencyId, recordType: "campaign", recordId: id, action: "archived", actor: user, metadata: { name: archived.name || "" } }).catch(error => console.error("Failed to write audit event:", error));
      setToast({ msg: "Campaign archived.", type: "success" }); setConfirm(null);
    } catch (e) {
      setToast({ msg: e.message || "Failed to archive campaign.", type: "error" });
    }
  };
  const restore = async id => {
    if (!canManage) return setToast({ msg: readOnlyMessage(user), type: "error" });
    if (!canManage) return setToast({ msg: readOnlyMessage(user), type: "error" });
    try {
      const restored = await restoreCampaignInSupabase(id);
      setCampaigns(v => v.map(x => x.id === id ? restored : x));
      createAuditEventInSupabase({ agencyId: user.agencyId, recordType: "campaign", recordId: id, action: "restored", actor: user, metadata: { name: restored.name || "" } }).catch(error => console.error("Failed to write audit event:", error));
      setToast({ msg: "Campaign restored.", type: "success" });
    } catch (e) {
      setToast({ msg: e.message || "Failed to restore campaign.", type: "error" });
    }
  };
  const clientOpts = activeOnly(clients).map(c => ({ value: c.id, label: c.name }));
  const clientBrands = f.clientId ? (clients.find(c => c.id === f.clientId)?.brands || "").split(",").map(b => b.trim()).filter(Boolean) : [];
  const visible = viewMode === "archived" ? archivedOnly(campaigns) : viewMode === "all" ? campaigns : activeOnly(campaigns);
  const filtered = visible.filter(c => {
    const cl = clients.find(x => x.id === c.clientId);
    return `${c.name} ${cl?.name || ""}`.toLowerCase().includes(search.toLowerCase()) && (filterStatus === "all" || c.status === filterStatus);
  });
  const getDur = c => { if (!c.startDate || !c.endDate) return "—"; const d = Math.ceil((new Date(c.endDate) - new Date(c.startDate)) / 86400000); return d > 0 ? `${d}d` : "—"; };
  const sc = { planning: "accent", active: "green", paused: "purple", completed: "blue" };
  return (
    <div className="fade">
      {toast && <Toast msg={toast.msg} type={toast.type} onDone={() => setToast(null)} />}
      {confirm && <Confirm msg={confirm.msg} onYes={confirm.onYes} onNo={() => setConfirm(null)} />}
      {!canManage && <Card style={{ marginBottom: 14, background: "rgba(59,126,245,.06)", border: "1px solid rgba(59,126,245,.18)" }}><div style={{ fontSize: 13, color: "var(--text2)" }}>You have read-only access to Vendors as <strong>{formatRoleLabel(user?.role)}</strong>.</div></Card>}
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 24, flexWrap: "wrap", gap: 10 }}>
        <div><h1 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 800, fontSize: 24 }}>Campaigns</h1><p style={{ color: "var(--text2)", marginTop: 3, fontSize: 13 }}>Track all advertising campaigns</p></div>
        {canManage && <Btn icon="+" onClick={() => { setF(blank); setModal("add"); }}>New Campaign</Btn>}
      </div>
      <div style={{ display: "flex", gap: 10, marginBottom: 18, flexWrap: "wrap" }}>
        <div style={{ flex: 1, minWidth: 180 }}><Field value={search} onChange={setSearch} placeholder="Search…" /></div>
        <Field value={filterStatus} onChange={setFilterStatus} options={[{value:"all",label:"All Status"},{value:"planning",label:"Planning"},{value:"active",label:"Active"},{value:"paused",label:"Paused"},{value:"completed",label:"Completed"}]} />
        <Field value={viewMode} onChange={setViewMode} options={[{value:"active",label:"Active"},{value:"archived",label:"Archived"},{value:"all",label:"All"}]} />
      </div>
      {filtered.length === 0 ? <Card><Empty icon="📢" title="No campaigns found" sub="Create your first campaign" /></Card> :
        <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
          {filtered.map(c => {
            const cl = clients.find(x => x.id === c.clientId);
            return (
              <Card key={c.id} style={{ padding: "14px 18px" }}>
                <div style={{ display: "flex", alignItems: "center", gap: 14, flexWrap: "wrap" }}>
                  <div style={{ width: 44, height: 44, background: "var(--bg3)", border: "1px solid var(--border2)", borderRadius: 11, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 20, flexShrink: 0 }}>📢</div>
                  <div style={{ flex: 1, minWidth: 160 }}>
                    <div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 15 }}>{c.name}</div>
                    <div style={{ fontSize: 12, color: "var(--text2)", marginTop: 2 }}>{cl?.name || "—"}{c.brand && ` · ${c.brand}`}{c.medium && ` · ${c.medium}`}</div>
                  </div>
                  <div style={{ display: "flex", gap: 18, flexWrap: "wrap", alignItems: "center" }}>
                    <div style={{ textAlign: "center" }}><div style={{ fontSize: 10, color: "var(--text3)", fontWeight: 600, textTransform: "uppercase" }}>Duration</div><div style={{ fontSize: 13, fontWeight: 600, marginTop: 2 }}>{getDur(c)}</div></div>
                    <div style={{ textAlign: "center" }}><div style={{ fontSize: 10, color: "var(--text3)", fontWeight: 600, textTransform: "uppercase" }}>Budget</div><div style={{ fontSize: 13, fontWeight: 600, marginTop: 2, color: "var(--green)" }}>{fmtN(c.budget)}</div></div>
                    <div style={{ textAlign: "center" }}><div style={{ fontSize: 10, color: "var(--text3)", fontWeight: 600, textTransform: "uppercase" }}>Period</div><div style={{ fontSize: 11, marginTop: 2 }}>{c.startDate || "—"} → {c.endDate || "—"}</div></div>
                    <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}><Badge color={sc[c.status] || "accent"}>{c.status}</Badge>{isArchived(c) && <Badge color="red">Archived</Badge>}</div>
                  </div>
                  <div style={{ display: "flex", gap: 6 }}>
                    {canManage && <Btn variant="ghost" size="sm" onClick={() => { setF({ name: c.name, clientId: c.clientId, brand: c.brand||"", objective: c.objective||"", startDate: c.startDate||"", endDate: c.endDate||"", budget: c.budget||"", status: c.status||"planning", medium: c.medium||"", notes: c.notes||"", materialList: Array.isArray(c.materialList) ? c.materialList : [] }); setModal(c); }}>✏️</Btn>}
                    {canManage && (isArchived(c) ? <Btn variant="success" size="sm" onClick={() => restore(c.id)}>↩</Btn> : <Btn variant="danger" size="sm" onClick={() => setConfirm({ msg: `Archive "${c.name}"? Linked MPO history will remain available.`, onYes: () => del(c.id) })}>🗄</Btn>)}
                  </div>
                </div>
                {c.objective && <div style={{ marginTop: 8, fontSize: 12, color: "var(--text2)", borderTop: "1px solid var(--border)", paddingTop: 8 }}>🎯 {c.objective}</div>}
                {c.materialList && c.materialList.length > 0 && <div style={{ marginTop: 6, fontSize: 11, color: "var(--text3)" }}>🎬 {c.materialList.length} material{c.materialList.length!==1?"s":""}: {c.materialList.slice(0,2).join(", ")}{c.materialList.length>2?` +${c.materialList.length-2} more`:""}</div>}
              </Card>
            );
          })}
        </div>}
      {modal !== null && (
        <Modal title={modal === "add" ? "New Campaign" : "Edit Campaign"} onClose={() => setModal(null)} width={580}>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 }}>
            <div style={{ gridColumn: "1/-1" }}><Field label="Campaign Title" value={f.name} onChange={u("name")} placeholder="TB Awareness Campaign Q4 2024" required /></div>
            <Field label="Client" value={f.clientId} onChange={v => { u("clientId")(v); u("brand")(""); }} options={clientOpts} placeholder="Select client" required />
            <Field label="Brand" value={f.brand} onChange={u("brand")} options={clientBrands.length ? clientBrands.map(b => ({ value: b, label: b })) : undefined} placeholder="Enter brand name" />
            <Field label="Campaign Objective" value={f.objective} onChange={u("objective")} placeholder="Awareness, Sales…" />
            <Field label="Medium" value={f.medium} onChange={u("medium")} options={["Television","Radio","Print","Digital","Multi-Platform","OOH"]} />
            <Field label="Start Date" type="date" value={f.startDate} onChange={u("startDate")} />
            <Field label="End Date" type="date" value={f.endDate} onChange={u("endDate")} />
            <Field label="Total Budget (₦)" type="number" value={f.budget} onChange={u("budget")} placeholder="0" />
            <Field label="Status" value={f.status} onChange={u("status")} options={[{value:"planning",label:"Planning"},{value:"active",label:"Active"},{value:"paused",label:"Paused"},{value:"completed",label:"Completed"}]} />
            <div style={{ gridColumn: "1/-1" }}>
              <div style={{ fontSize: 11, fontWeight: 600, color: "var(--text3)", textTransform: "uppercase", letterSpacing: ".08em", marginBottom: 6 }}>Campaign Materials</div>
              <div style={{ display: "flex", flexDirection: "column", gap: 7, marginBottom: 8 }}>
                {(f.materialList||[]).map((mat, mi) => (
                  <div key={mi} style={{ display: "flex", gap: 7, alignItems: "center" }}>
                    <input value={mat} onChange={e => { const ml = [...(f.materialList||[])]; ml[mi] = e.target.value; u("materialList")(ml); }}
                      placeholder="e.g. TB Thematic English 30secs (MP4)"
                      style={{ flex: 1, background: "var(--bg3)", border: "1px solid var(--border2)", borderRadius: 8, padding: "8px 12px", color: "var(--text)", fontSize: 13, outline: "none" }}
                      onFocus={e => e.target.style.borderColor="var(--accent)"}
                      onBlur={e => e.target.style.borderColor="var(--border2)"} />
                    <button onClick={() => { const ml = (f.materialList||[]).filter((_,i)=>i!==mi); u("materialList")(ml); }}
                      style={{ background: "rgba(239,68,68,.12)", border: "1px solid rgba(239,68,68,.3)", color: "var(--red)", borderRadius: 7, width: 32, height: 32, cursor: "pointer", fontSize: 14, flexShrink: 0 }}>×</button>
                  </div>
                ))}
                <button onClick={() => u("materialList")([...(f.materialList||[]), ""])}
                  style={{ display: "flex", alignItems: "center", gap: 6, background: "none", border: "1px dashed var(--border2)", borderRadius: 8, padding: "7px 14px", cursor: "pointer", fontSize: 12, color: "var(--text2)", width: "100%" }}>
                  <span style={{ fontSize: 16 }}>+</span> Add Material
                </button>
              </div>
              <div style={{ fontSize: 11, color: "var(--text3)" }}>Each material will be available as a dropdown when generating MPOs</div>
            </div>
            <div style={{ gridColumn: "1/-1" }}><Field label="Notes" value={f.notes} onChange={u("notes")} placeholder="Additional notes…" /></div>
          </div>
          <div style={{ display: "flex", gap: 10, justifyContent: "flex-end", marginTop: 20 }}>
            <Btn variant="ghost" onClick={() => setModal(null)}>Cancel</Btn>
            <Btn onClick={save}>{modal === "add" ? "Create Campaign" : "Save Changes"}</Btn>
          </div>
        </Modal>
      )}
    </div>
  );
};

/* ── RATES ──────────────────────────────────────────────── */
const blankRateRow = () => ({ _id: uid(), programme: "", timeBelt: "", duration: "30", ratePerSpot: "" });

/* Excel upload helper — uses SheetJS from CDN */
const loadSheetJS = () => new Promise((resolve, reject) => {
  if (window.XLSX) return resolve(window.XLSX);
  const s = document.createElement("script");
  s.src = "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";
  s.onload = () => resolve(window.XLSX);
  s.onerror = reject;
  document.head.appendChild(s);
});

const normaliseExcelRow = (row, vendors) => {
  const get = (...keys) => {
    for (const k of keys) {
      const found = Object.keys(row).find(rk => rk.trim().toLowerCase() === k.toLowerCase());
      if (found !== undefined && row[found] !== undefined && row[found] !== "") return String(row[found]).trim();
    }
    return "";
  };
  const vendorName = get("vendor", "vendor name", "station", "media owner");
  const matchedVendor = vendors.find(v => v.name.toLowerCase() === vendorName.toLowerCase());
  return {
    _vendorName: vendorName,
    vendorId:    matchedVendor?.id || "",
    mediaType:   get("media", "media type", "type", "medium"),
    programme:   get("programme", "program", "slot", "programme / slot", "programme/slot"),
    timeBelt:    get("timebelt", "time belt", "time", "belt", "daypart"),
    duration:    get("duration", "dur", "duration (secs)", "secs") || "30",
    ratePerSpot: get("rate", "rate per spot", "rate/spot", "cost", "amount", "price"),
    discount:    get("discount", "disc", "volume discount", "disc%", "volume discount (%)") || "0",
    commission:  get("commission", "comm", "agency commission", "comm%", "agency commission (%)") || "0",
    vat: "0", notes: get("notes", "note", "remarks"),
    campaignId: "", clientId: "",
  };
};

const ExcelImportModal = ({ vendors, onImport, onClose }) => {
  const [step, setStep]         = useState("upload");
  const [rows, setRows]         = useState([]);
  const [errors, setErrors]     = useState([]);
  const [fileName, setFileName] = useState("");
  const [loading, setLoading]   = useState(false);
  const [selected, setSelected] = useState([]);
  const fileRef                 = useRef();

  const handleFile = async (file) => {
    if (!file) return;
    setLoading(true); setFileName(file.name);
    try {
      const XLSX = await loadSheetJS();
      const buf  = await file.arrayBuffer();
      const wb   = XLSX.read(buf, { type: "array" });
      const ws   = wb.Sheets[wb.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json(ws, { defval: "" });
      if (!json.length) { setErrors(["The sheet appears to be empty."]); setLoading(false); return; }
      const parsed = json.map((r, i) => ({ ...normaliseExcelRow(r, vendors), _rowIdx: i }));
      const errs = [];
      parsed.forEach((r, i) => {
        if (!r._vendorName) errs.push("Row " + (i + 2) + ": Missing Vendor name.");
        if (!r.ratePerSpot || isNaN(parseFloat(r.ratePerSpot))) errs.push("Row " + (i + 2) + ": Invalid or missing Rate.");
        if (!r.vendorId) errs.push("Row " + (i + 2) + ": Vendor \"" + r._vendorName + "\" not found — will import without vendor link.");
      });
      setErrors(errs); setRows(parsed); setSelected(parsed.map((_, i) => i)); setStep("preview");
    } catch(e) { setErrors(["Failed to parse file. Ensure it is a valid .xlsx or .xls file."]); }
    setLoading(false);
  };

  const toggleRow = i => setSelected(s => s.includes(i) ? s.filter(x => x !== i) : [...s, i]);
  const toggleAll = () => setSelected(selected.length === rows.length ? [] : rows.map((_, i) => i));

  const confirmImport = () => {
    const toImport = rows.filter((_, i) => selected.includes(i)).map(r => ({
      id: uid(), createdAt: Date.now(),
      vendorId: r.vendorId, mediaType: r.mediaType, programme: r.programme,
      timeBelt: r.timeBelt, duration: r.duration || "30",
      ratePerSpot: r.ratePerSpot, discount: r.discount || "0",
      commission: r.commission || "0", vat: "0",
      notes: r.notes, campaignId: "", clientId: "",
    }));
    onImport(toImport); setStep("done");
  };

  const hardErrors = errors.filter(e => e.includes("Missing Vendor name") || e.includes("Invalid or missing Rate"));
  const warnings   = errors.filter(e => !hardErrors.includes(e));

  return (
    <Modal title="Import Rate Cards from Excel" onClose={onClose} width={860}>
      {step === "upload" && (
        <div>
          <div style={{ background: "var(--bg3)", borderRadius: 12, padding: 16, marginBottom: 20, border: "1px solid var(--border2)" }}>
            <div style={{ fontSize: 12, fontWeight: 700, color: "var(--accent)", fontFamily: "\'Syne\',sans-serif", marginBottom: 10 }}>Required Excel Columns</div>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill,minmax(175px,1fr))", gap: 8 }}>
              {[{col:"Vendor",req:true,note:"Must match vendor in system"},{col:"Media",req:false,note:"e.g. Television, Radio"},{col:"Type",req:false,note:"Media type / category"},{col:"Programme",req:false,note:"Show or slot name"},{col:"Timebelt",req:false,note:"e.g. 21:00-21:30"},{col:"Duration",req:false,note:"In seconds, e.g. 30"},{col:"Rate",req:true,note:"Rate per spot in N"},{col:"Discounts",req:false,note:"Volume discount %"}].map(({ col, req, note }) => (
                <div key={col} style={{ background: "var(--bg4)", borderRadius: 8, padding: "8px 11px", border: req ? "1px solid rgba(240,165,0,.3)" : "1px solid var(--border)" }}>
                  <div style={{ fontSize: 11, fontWeight: 700, color: req ? "var(--accent)" : "var(--text)", marginBottom: 2 }}>{col}{req && <span style={{ color: "var(--red)", marginLeft: 3 }}>*</span>}</div>
                  <div style={{ fontSize: 10, color: "var(--text3)" }}>{note}</div>
                </div>
              ))}
            </div>
            <div style={{ marginTop: 10, display: "flex", gap: 10, justifyContent: "space-between", alignItems: "center", flexWrap: "wrap" }}><div style={{ fontSize: 11, color: "var(--text3)" }}>Column headers are flexible. First row must be headers.</div><Btn variant="ghost" size="sm" onClick={downloadRateTemplate}>⬇ Download Template</Btn></div>
          </div>
          <div onClick={() => fileRef.current?.click()}
            onDragOver={e => { e.preventDefault(); e.currentTarget.style.borderColor = "var(--accent)"; }}
            onDragLeave={e => { e.currentTarget.style.borderColor = "var(--border2)"; }}
            onDrop={e => { e.preventDefault(); e.currentTarget.style.borderColor = "var(--border2)"; handleFile(e.dataTransfer.files[0]); }}
            style={{ background: "var(--bg3)", border: "2px dashed var(--border2)", borderRadius: 14, padding: "38px 24px", textAlign: "center", cursor: "pointer", transition: "all .2s" }}>
            <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }} onChange={e => handleFile(e.target.files[0])} />
            {loading
              ? <div><div style={{ fontSize: 36, marginBottom: 10 }}>&#x23F3;</div><div style={{ fontFamily: "\'Syne\',sans-serif", fontWeight: 700 }}>Parsing file...</div></div>
              : <><div style={{ fontSize: 48, marginBottom: 10 }}>&#x1F4CA;</div><div style={{ fontFamily: "\'Syne\',sans-serif", fontWeight: 700, fontSize: 16, marginBottom: 6 }}>Drop your Excel file here</div><div style={{ color: "var(--text2)", fontSize: 13, marginBottom: 14 }}>or click to browse</div><Btn variant="secondary" size="sm">Choose File</Btn></>}
          </div>
          {errors.length > 0 && <div style={{ marginTop: 12, background: "rgba(239,68,68,.08)", border: "1px solid rgba(239,68,68,.25)", borderRadius: 10, padding: 12 }}>{errors.map((e, i) => <div key={i} style={{ fontSize: 12, color: "var(--red)" }}>Warning: {e}</div>)}</div>}
        </div>
      )}
      {step === "preview" && (
        <div>
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 12, flexWrap: "wrap", gap: 8 }}>
            <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center" }}>
              <Badge color="blue">{fileName}</Badge>
              <Badge color="green">{rows.length} rows parsed</Badge>
              <Badge color="accent">{selected.length} selected</Badge>
              {warnings.length > 0 && <Badge color="orange">{warnings.length} warning{warnings.length > 1 ? "s" : ""}</Badge>}
              {hardErrors.length > 0 && <Badge color="red">{hardErrors.length} error{hardErrors.length > 1 ? "s" : ""}</Badge>}
            </div>
            <Btn variant="ghost" size="sm" onClick={() => { setStep("upload"); setRows([]); setErrors([]); setSelected([]); }}>Re-upload</Btn>
          </div>
          {warnings.length > 0 && <div style={{ background: "rgba(249,115,22,.08)", border: "1px solid rgba(249,115,22,.25)", borderRadius: 10, padding: "9px 12px", marginBottom: 10 }}>{warnings.map((w, i) => <div key={i} style={{ fontSize: 11, color: "var(--orange)" }}>Warning: {w}</div>)}</div>}
          {hardErrors.length > 0 && <div style={{ background: "rgba(239,68,68,.08)", border: "1px solid rgba(239,68,68,.25)", borderRadius: 10, padding: "9px 12px", marginBottom: 10 }}>{hardErrors.map((e, i) => <div key={i} style={{ fontSize: 11, color: "var(--red)" }}>Error: {e}</div>)}</div>}
          <div style={{ overflowX: "auto", maxHeight: 360, overflowY: "auto", borderRadius: 10, border: "1px solid var(--border)" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 700, fontSize: 12 }}>
              <thead style={{ position: "sticky", top: 0, background: "var(--bg3)", zIndex: 1 }}>
                <tr>
                  <th style={{ padding: "9px 12px", textAlign: "center", borderBottom: "1px solid var(--border)", width: 36 }}>
                    <input type="checkbox" checked={selected.length === rows.length} onChange={toggleAll} style={{ accentColor: "var(--accent)", cursor: "pointer", width: 14, height: 14 }} />
                  </th>
                  {["Vendor","Media Type","Programme","Time Belt","Dur","Rate","Disc%","Comm%","Status"].map(h => (
                    <th key={h} style={{ padding: "9px 12px", textAlign: "left", fontSize: 10, fontWeight: 600, textTransform: "uppercase", color: "var(--text3)", borderBottom: "1px solid var(--border)", whiteSpace: "nowrap" }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {rows.map((r, i) => {
                  const isSelected = selected.includes(i);
                  const hasError   = hardErrors.some(e => e.startsWith("Row " + (i + 2) + ":"));
                  const hasWarn    = warnings.some(w => w.startsWith("Row " + (i + 2) + ":"));
                  return (
                    <tr key={i} style={{ borderBottom: "1px solid var(--border)", opacity: isSelected ? 1 : 0.4 }}>
                      <td style={{ padding: "9px 12px", textAlign: "center" }}><input type="checkbox" checked={isSelected} onChange={() => toggleRow(i)} style={{ accentColor: "var(--accent)", cursor: "pointer", width: 14, height: 14 }} /></td>
                      <td style={{ padding: "9px 12px", fontWeight: 600 }}>
                        {r._vendorName || "—"}
                        {r.vendorId && <div style={{ fontSize: 10, color: "var(--green)" }}>Linked</div>}
                        {!r.vendorId && r._vendorName && <div style={{ fontSize: 10, color: "var(--orange)" }}>Not matched</div>}
                      </td>
                      <td style={{ padding: "9px 12px", color: "var(--text2)" }}>{r.mediaType || "—"}</td>
                      <td style={{ padding: "9px 12px", color: "var(--text2)" }}>{r.programme || "—"}</td>
                      <td style={{ padding: "9px 12px", color: "var(--text2)" }}>{r.timeBelt || "—"}</td>
                      <td style={{ padding: "9px 12px", color: "var(--text2)" }}>{r.duration || "30"}</td>
                      <td style={{ padding: "9px 12px", fontWeight: 600, color: r.ratePerSpot ? "var(--accent)" : "var(--red)" }}>{r.ratePerSpot || "MISSING"}</td>
                      <td style={{ padding: "9px 12px", color: "var(--red)" }}>{r.discount || 0}%</td>
                      <td style={{ padding: "9px 12px", color: "var(--red)" }}>{r.commission || 0}%</td>
                      <td style={{ padding: "9px 12px" }}>{hasError ? <Badge color="red">Error</Badge> : hasWarn ? <Badge color="orange">Warning</Badge> : <Badge color="green">Ready</Badge>}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginTop: 16, flexWrap: "wrap", gap: 10 }}>
            <div style={{ fontSize: 12, color: "var(--text2)" }}>{selected.length === 0 ? "No rows selected." : selected.length + " of " + rows.length + " rows will be imported."}</div>
            <div style={{ display: "flex", gap: 10 }}>
              <Btn variant="ghost" onClick={onClose}>Cancel</Btn>
              <Btn onClick={confirmImport} disabled={selected.length === 0} icon="save">Import {selected.length} Rate Card{selected.length !== 1 ? "s" : ""}</Btn>
            </div>
          </div>
        </div>
      )}
      {step === "done" && (
        <div style={{ textAlign: "center", padding: "36px 16px" }}>
          <div style={{ fontSize: 56, marginBottom: 14 }}>&#x2705;</div>
          <div style={{ fontFamily: "\'Syne\',sans-serif", fontWeight: 800, fontSize: 22, marginBottom: 8 }}>Import Complete!</div>
          <div style={{ color: "var(--text2)", fontSize: 14, marginBottom: 24 }}>{selected.length} rate card{selected.length !== 1 ? "s" : ""} from {fileName} saved.</div>
          <Btn onClick={onClose}>Close</Btn>
        </div>
      )}
    </Modal>
  );
};


const RatesPage = ({ rates, setRates, vendors, clients, campaigns, user }) => {
  const canManage = hasPermission(user, "manageRates");
  const [modal, setModal]         = useState(null);
  const [importModal, setImportModal] = useState(false);
  const [search, setSearch]       = useState("");
  const [filterV, setFilterV]     = useState("");
  const [viewMode, setViewMode]   = useState("active");
  const [toast, setToast]         = useState(null);
  const [confirm, setConfirm]     = useState(null);

  /* shared vendor-level fields */
  const blankHeader = { vendorId: "", mediaType: "", discount: "", commission: "", notes: "" };
  const [hdr, setHdr]   = useState(blankHeader);
  const uh = k => v => setHdr(p => ({ ...p, [k]: v }));

  /* multi-row programme list */
  const [rows, setRows] = useState([blankRateRow()]);
  const updRow = (id, k, v) => setRows(rs => rs.map(r => r._id === id ? { ...r, [k]: v } : r));
  const addRow = () => setRows(rs => [...rs, blankRateRow()]);
  const delRow = id => setRows(rs => rs.length > 1 ? rs.filter(r => r._id !== id) : rs);

  /* editing a single existing rate — keep header + one row */
  const [editId, setEditId] = useState(null);

  /* auto-fill discount/commission from vendor */
  useEffect(() => {
    if (hdr.vendorId) {
      const v = vendors.find(x => x.id === hdr.vendorId);
      if (v) setHdr(p => ({ ...p, discount: p.discount || v.discount || "", commission: p.commission || v.commission || "", mediaType: p.mediaType || v.type || "" }));
    }
  }, [hdr.vendorId]);

  const calcNet = r => {
    const rate = parseFloat(r.ratePerSpot) || 0;
    return rate * (1 - (parseFloat(hdr.discount || r.discount) || 0) / 100) * (1 - (parseFloat(hdr.commission || r.commission) || 0) / 100);
  };
  const calcNetR = (r, disc, comm) => {
    const rate = parseFloat(r.ratePerSpot) || 0;
    return rate * (1 - (parseFloat(disc) || 0) / 100) * (1 - (parseFloat(comm) || 0) / 100);
  };

  const openAdd = () => {
    if (!canManage) return setToast({ msg: readOnlyMessage(user), type: "error" });
    setHdr(blankHeader); setRows([blankRateRow()]); setEditId(null); setModal("add");
  };
  const openEdit = r => {
    setHdr({ vendorId: r.vendorId, mediaType: r.mediaType || "", discount: r.discount || "", commission: r.commission || "", notes: r.notes || "" });
    setRows([{ _id: uid(), programme: r.programme || "", timeBelt: r.timeBelt || "", duration: r.duration || "30", ratePerSpot: r.ratePerSpot || "" }]);
    setEditId(r.id); setModal("edit");
  };

  const save = async () => {
    if (!canManage) return setToast({ msg: readOnlyMessage(user), type: "error" });
    if (!hdr.vendorId) return setToast({ msg: "Please select a vendor.", type: "error" });
    if (!pctWithin(hdr.discount) || !pctWithin(hdr.commission)) return setToast({ msg: "Discount and commission must be between 0 and 100.", type: "error" });
    const validRows = rows.filter(r => r.programme && r.ratePerSpot);
    if (!validRows.length) return setToast({ msg: "Add at least one programme with a rate.", type: "error" });
    if (validRows.some(r => (parseFloat(r.ratePerSpot) || 0) <= 0)) return setToast({ msg: "Every saved rate must be greater than zero.", type: "error" });
    const duplicateRow = validRows.find(row => rates.find(existing => existing.id !== editId && !isArchived(existing) && existing.vendorId === hdr.vendorId && (existing.programme || "").trim().toLowerCase() === row.programme.trim().toLowerCase() && (existing.timeBelt || "") === (row.timeBelt || "") && String(existing.duration || "30") === String(row.duration || "30")));
    if (duplicateRow) return setToast({ msg: `A matching rate card already exists for ${duplicateRow.programme}.`, type: "error" });

    try {
      if (modal === "edit" && editId) {
        const row = validRows[0];
        const updatedRate = await updateRateInSupabase(editId, hdr, row);
        setRates(v => v.map(x => x.id === editId ? updatedRate : x));
        createAuditEventInSupabase({ agencyId: user.agencyId, recordType: "rate", recordId: editId, action: "updated", actor: user, metadata: { vendorId: updatedRate.vendorId || hdr.vendorId || "", programme: updatedRate.programme || row.programme || "" } }).catch(error => console.error("Failed to write audit event:", error));
        setToast({ msg: "Rate updated.", type: "success" });
      } else {
        const createdRates = await createRatesInSupabase(user.agencyId, user.id, hdr, validRows);
        setRates(v => [...createdRates, ...v]);
        createAuditEventInSupabase({ agencyId: user.agencyId, recordType: "rate", recordId: null, action: "created", actor: user, note: `${createdRates.length} rate card${createdRates.length !== 1 ? "s" : ""} added.`, metadata: { count: createdRates.length, vendorId: hdr.vendorId || "" } }).catch(error => console.error("Failed to write audit event:", error));
        setToast({ msg: `${createdRates.length} rate card${createdRates.length !== 1 ? "s" : ""} added.`, type: "success" });
      }
      setModal(null);
    } catch (e) {
      setToast({ msg: e.message || "Failed to save rate.", type: "error" });
    }
  };

  const del = async id => {
    if (!canManage) return setToast({ msg: readOnlyMessage(user), type: "error" });
    if (!canManage) return setToast({ msg: readOnlyMessage(user), type: "error" });
    try {
      const archived = await archiveRateInSupabase(id);
      setRates(v => v.map(x => x.id === id ? archived : x));
      createAuditEventInSupabase({ agencyId: user.agencyId, recordType: "rate", recordId: id, action: "archived", actor: user, metadata: { programme: archived.programme || "", vendorId: archived.vendorId || "" } }).catch(error => console.error("Failed to write audit event:", error));
      setToast({ msg: "Rate card archived.", type: "success" });
      setConfirm(null);
    } catch (e) {
      setToast({ msg: e.message || "Failed to archive rate.", type: "error" });
    }
  };
  const restore = async id => {
    if (!canManage) return setToast({ msg: readOnlyMessage(user), type: "error" });
    if (!canManage) return setToast({ msg: readOnlyMessage(user), type: "error" });
    try {
      const restored = await restoreRateInSupabase(id);
      setRates(v => v.map(x => x.id === id ? restored : x));
      createAuditEventInSupabase({ agencyId: user.agencyId, recordType: "rate", recordId: id, action: "restored", actor: user, metadata: { programme: restored.programme || "", vendorId: restored.vendorId || "" } }).catch(error => console.error("Failed to write audit event:", error));
      setToast({ msg: "Rate card restored.", type: "success" });
    } catch (e) {
      setToast({ msg: e.message || "Failed to restore rate.", type: "error" });
    }
  };

  const handleExcelImport = async (newRates) => {
    try {
      const imported = await importRatesInSupabase(user.agencyId, user.id, newRates);
      setRates(v => [...imported, ...v]);
      setToast({ msg: `${imported.length} rate card${imported.length !== 1 ? "s" : ""} imported successfully!`, type: "success" });
    } catch (e) {
      setToast({ msg: e.message || "Failed to import rates.", type: "error" });
    }
  };

  const visibleRates = viewMode === "archived" ? archivedOnly(rates) : viewMode === "all" ? rates : activeOnly(rates);
  const filtered = visibleRates.filter(r => {
    const vn = vendors.find(v => v.id === r.vendorId)?.name || "";
    return `${vn} ${r.programme || ""}`.toLowerCase().includes(search.toLowerCase()) && (!filterV || r.vendorId === filterV);
  });

  const inputSt = { background: "var(--bg2)", border: "1px solid var(--border2)", borderRadius: 7, padding: "7px 10px", color: "var(--text)", fontSize: 12, outline: "none", width: "100%" };

  return (
    <div className="fade">
      {toast && <Toast msg={toast.msg} type={toast.type} onDone={() => setToast(null)} />}
      {confirm && <Confirm msg={confirm.msg} onYes={confirm.onYes} onNo={() => setConfirm(null)} />}

      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 24, flexWrap: "wrap", gap: 10 }}>
        <div><h1 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 800, fontSize: 24 }}>Media Rates</h1><p style={{ color: "var(--text2)", marginTop: 3, fontSize: 13 }}>Rate cards across media owners</p></div>
        <div style={{ display: "flex", gap: 10 }}>
          <Btn variant="blue" icon="&#x1F4CA;" onClick={() => setImportModal(true)}>Import from Excel</Btn>
          <Btn icon="+" onClick={openAdd}>Add Rate Card</Btn>
        </div>
      </div>

      <div style={{ display: "flex", gap: 10, marginBottom: 18, flexWrap: "wrap" }}>
        <div style={{ flex: 1, minWidth: 180 }}><Field value={search} onChange={setSearch} placeholder="Search by vendor or programme…" /></div>
        <Field value={filterV} onChange={setFilterV} options={activeOnly(vendors).map(v => ({ value: v.id, label: v.name }))} placeholder="All Vendors" />
        <Field value={viewMode} onChange={setViewMode} options={[{value:"active",label:"Active"},{value:"archived",label:"Archived"},{value:"all",label:"All"}]} />
      </div>

      {filtered.length === 0 ? <Card><Empty icon="💰" title="No rate cards" sub="Add media rates to use in MPO generation" /></Card> :
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 800 }}>
            <thead><tr style={{ background: "var(--bg3)" }}>
              {["Vendor","Programme","Time Belt","Type","Dur","Rate/Spot","Disc%","Comm%","Net Rate",""].map(h => <th key={h} style={{ padding: "7px 9px", textAlign: "left", fontSize: 9, fontWeight: 600, textTransform: "uppercase", letterSpacing: ".07em", color: "var(--text3)", borderBottom: "1px solid var(--border)", whiteSpace: "nowrap" }}>{h}</th>)}
            </tr></thead>
            <tbody>
              {filtered.map((r, i) => {
                const vn  = vendors.find(v => v.id === r.vendorId)?.name || "—";
                const net = calcNetR(r, r.discount, r.commission);
                return (
                  <tr key={r.id} style={{ borderBottom: "1px solid var(--border)", background: i % 2 === 0 ? "transparent" : "rgba(255,255,255,.01)" }}
                    onMouseEnter={e => e.currentTarget.style.background = "var(--bg3)"}
                    onMouseLeave={e => e.currentTarget.style.background = i % 2 === 0 ? "transparent" : "rgba(255,255,255,.01)"}>
                    <td style={{ padding: "8px 9px", fontWeight: 600, fontSize: 12 }}>{vn}</td>
                    <td style={{ padding: "8px 9px", color: "var(--text2)", fontSize: 12 }}>{r.programme || "—"}</td>
                    <td style={{ padding: "8px 9px", color: "var(--text2)", fontSize: 12 }}>{r.timeBelt || "—"}</td>
                    <td style={{ padding: "8px 9px", fontSize: 12 }}><div style={{ display: "flex", gap: 5, flexWrap: "wrap" }}><Badge color="blue">{r.mediaType || "—"}</Badge>{isArchived(r) && <Badge color="red">Archived</Badge>}</div></td>
                    <td style={{ padding: "8px 9px", color: "var(--text2)", fontSize: 12 }}>{r.duration}"</td>
                    <td style={{ padding: "8px 9px", fontWeight: 600, fontSize: 12, color: "var(--accent)" }}>{fmtN(r.ratePerSpot)}</td>
                    <td style={{ padding: "11px 12px", color: "var(--red)" }}>{r.discount || 0}%</td>
                    <td style={{ padding: "11px 12px", color: "var(--red)" }}>{r.commission || 0}%</td>
                    <td style={{ padding: "8px 9px", fontWeight: 700, color: "var(--green)", fontSize: 12 }}>{fmtN(net)}</td>
                    <td style={{ padding: "11px 12px" }}>
                      <div style={{ display: "flex", gap: 5 }}>
                        {canManage && <Btn variant="ghost" size="sm" onClick={() => openEdit(r)}>✏️</Btn>}
                        {canManage && (isArchived(r) ? <Btn variant="success" size="sm" onClick={() => restore(r.id)}>↩</Btn> : <Btn variant="danger" size="sm" onClick={() => setConfirm({ msg: `Archive this rate card? Existing MPOs will still retain their saved values.`, onYes: () => del(r.id) })}>🗄</Btn>)}
                      </div>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>}

      {/* Excel Import Modal */}
      {importModal && (
        <ExcelImportModal vendors={vendors} onImport={handleExcelImport} onClose={() => setImportModal(false)} />
      )}

      {/* Add / Edit Modal */}
      {modal !== null && (
        <Modal title={modal === "edit" ? "Edit Rate Card" : "Add Rate Cards"} onClose={() => setModal(null)} width={780}>

          {/* ── Vendor-level header ── */}
          <div style={{ background: "var(--bg3)", borderRadius: 12, padding: "14px 16px", marginBottom: 18, border: "1px solid var(--border2)" }}>
            <div style={{ fontSize: 11, fontWeight: 700, color: "var(--accent)", textTransform: "uppercase", letterSpacing: ".07em", marginBottom: 12, fontFamily: "'Syne',sans-serif" }}>Vendor Details</div>
            <div style={{ display: "grid", gridTemplateColumns: "2fr 1.5fr 1fr 1fr", gap: 12 }}>
              <Field label="Vendor" value={hdr.vendorId} onChange={uh("vendorId")} options={vendors.map(v => ({ value: v.id, label: v.name }))} placeholder="Select vendor" required note="Disc & Comm auto-filled" />
              <Field label="Media Type" value={hdr.mediaType} onChange={uh("mediaType")} options={["Television","Radio","Print","Digital/Online","Out-of-Home (OOH)","Cinema"]} />
              <Field label="Volume Discount (%)" type="number" value={hdr.discount} onChange={uh("discount")} placeholder="0" note="e.g. 27 for 27%" />
              <Field label="Agency Commission (%)" type="number" value={hdr.commission} onChange={uh("commission")} placeholder="0" note="e.g. 15 for 15%" />
            </div>
            <div style={{ marginTop: 12 }}>
              <Field label="Notes" value={hdr.notes} onChange={uh("notes")} placeholder="Additional notes…" />
            </div>
          </div>

          {/* ── Programme rows ── */}
          <div style={{ marginBottom: 10 }}>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 10 }}>
              <div style={{ fontSize: 11, fontWeight: 700, color: "var(--text2)", textTransform: "uppercase", letterSpacing: ".07em", fontFamily: "'Syne',sans-serif" }}>
                Programmes / Slots — {modal === "edit" ? "1 entry" : `${rows.length} row${rows.length !== 1 ? "s" : ""}`}
              </div>
              {modal !== "edit" && (
                <button onClick={addRow} style={{ display: "flex", alignItems: "center", gap: 6, background: "none", border: "1px dashed var(--border2)", borderRadius: 7, padding: "5px 12px", cursor: "pointer", fontSize: 12, color: "var(--text2)" }}>
                  <span style={{ fontSize: 16 }}>+</span> Add Row
                </button>
              )}
            </div>

            {/* column headers */}
            <div style={{ display: "grid", gridTemplateColumns: "2.5fr 1.5fr 80px 1.5fr 36px", gap: 8, marginBottom: 6, padding: "0 4px" }}>
              {["Programme / Slot","Time Belt","Dur (s)","Rate per Spot (₦)",""].map(h => (
                <div key={h} style={{ fontSize: 10, fontWeight: 600, color: "var(--text3)", textTransform: "uppercase", letterSpacing: ".07em" }}>{h}</div>
              ))}
            </div>

            <div style={{ display: "flex", flexDirection: "column", gap: 8, maxHeight: 340, overflowY: "auto" }}>
              {rows.map((row, ri) => {
                const net = parseFloat(row.ratePerSpot) ? fmtN(parseFloat(row.ratePerSpot) * (1 - (parseFloat(hdr.discount)||0)/100) * (1 - (parseFloat(hdr.commission)||0)/100)) : null;
                return (
                  <div key={row._id} style={{ display: "grid", gridTemplateColumns: "2.5fr 1.5fr 80px 1.5fr 36px", gap: 8, alignItems: "center", background: "var(--bg3)", borderRadius: 9, padding: "10px 12px", border: "1px solid var(--border)" }}>
                    <div>
                      <input value={row.programme} onChange={e => updRow(row._id, "programme", e.target.value)}
                        placeholder="e.g. NTA Network News" style={inputSt}
                        onFocus={e => e.target.style.borderColor="var(--accent)"}
                        onBlur={e => e.target.style.borderColor="var(--border2)"} />
                    </div>
                    <div>
                      <input value={row.timeBelt} onChange={e => updRow(row._id, "timeBelt", e.target.value)}
                        placeholder="e.g. 21:00–21:30" style={inputSt}
                        onFocus={e => e.target.style.borderColor="var(--accent)"}
                        onBlur={e => e.target.style.borderColor="var(--border2)"} />
                    </div>
                    <div>
                      <input type="number" value={row.duration} onChange={e => updRow(row._id, "duration", e.target.value)}
                        placeholder="30" style={inputSt}
                        onFocus={e => e.target.style.borderColor="var(--accent)"}
                        onBlur={e => e.target.style.borderColor="var(--border2)"} />
                    </div>
                    <div>
                      <input type="number" value={row.ratePerSpot} onChange={e => updRow(row._id, "ratePerSpot", e.target.value)}
                        placeholder="0" style={inputSt}
                        onFocus={e => e.target.style.borderColor="var(--accent)"}
                        onBlur={e => e.target.style.borderColor="var(--border2)"} />
                      {net && <div style={{ fontSize: 10, color: "var(--green)", marginTop: 2 }}>Net: {net}</div>}
                    </div>
                    <div style={{ display: "flex", justifyContent: "center" }}>
                      {rows.length > 1 ? (
                        <button onClick={() => delRow(row._id)} style={{ background: "none", border: "none", cursor: "pointer", color: "var(--red)", fontSize: 18, lineHeight: 1, opacity: .7 }}>×</button>
                      ) : (
                        <span style={{ fontSize: 12, color: "var(--text3)", opacity: .4 }}>#{ri+1}</span>
                      )}
                    </div>
                  </div>
                );
              })}
            </div>

            {/* live summary */}
            {rows.filter(r => r.ratePerSpot).length > 0 && modal !== "edit" && (
              <div style={{ marginTop: 12, background: "rgba(240,165,0,.07)", border: "1px solid rgba(240,165,0,.18)", borderRadius: 9, padding: "10px 14px", display: "flex", gap: 20, flexWrap: "wrap", alignItems: "center" }}>
                <span style={{ fontSize: 11, color: "var(--text3)" }}>Summary:</span>
                <span style={{ fontSize: 12, fontWeight: 600, color: "var(--accent)" }}>{rows.filter(r => r.programme && r.ratePerSpot).length} valid row{rows.filter(r => r.programme && r.ratePerSpot).length !== 1 ? "s" : ""} ready to save</span>
                {hdr.discount > 0 && <span style={{ fontSize: 11, color: "var(--text2)" }}>Vol Disc: <strong style={{ color:"var(--red)" }}>{hdr.discount}%</strong></span>}
                {hdr.commission > 0 && <span style={{ fontSize: 11, color: "var(--text2)" }}>Comm: <strong style={{ color:"var(--red)" }}>{hdr.commission}%</strong></span>}
              </div>
            )}
          </div>

          <div style={{ display: "flex", gap: 10, justifyContent: "flex-end", marginTop: 18, paddingTop: 16, borderTop: "1px solid var(--border)" }}>
            <Btn variant="ghost" onClick={() => setModal(null)}>Cancel</Btn>
            <Btn onClick={save} icon="💾">{modal === "edit" ? "Save Changes" : `Add ${rows.filter(r=>r.programme&&r.ratePerSpot).length || ""} Rate Card${rows.filter(r=>r.programme&&r.ratePerSpot).length !== 1 ? "s" : ""}`}</Btn>
          </div>
        </Modal>
      )}
    </div>
  );
};

/* ── EXPORT HELPERS ─────────────────────────────────────── */
const buildCSV = (rows, headers) => {
  const esc = v => `"${String(v ?? "").replace(/"/g,'""')}"`;
  return [headers.map(esc).join(","), ...rows.map(r => r.map(esc).join(","))].join("\n");
};

/* ── Lazy-load a CDN script ── */
/* ══════════════════════════════════════════════════════════════════
   PURE JS PDF BUILDER — zero CDN, zero canvas, zero dependencies.
   Generates a valid PDF entirely in-browser from MPO data.
   ══════════════════════════════════════════════════════════════════ */
/* ══════════════════════════════════════════════════════════════════
   PURE-JS PDF BUILDER  —  mirrors buildMPOHTML exactly
   No CDN · No canvas · No network · Works in any sandbox
   ══════════════════════════════════════════════════════════════════ */
const buildMPOPdf = (mpo) => {
  /* ── shared helpers ─────────────────────────────────────── */
  const esc  = s => String(s??'').replace(/\\/g,'\\\\').replace(/\(/g,'\\(').replace(/\)/g,'\\)');
  const fmtN = n => Number(n||0).toLocaleString('en-NG',{minimumFractionDigits:2,maximumFractionDigits:2});
  const clip = (s,n) => String(s||'').slice(0,n);
  const MN   = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];
  const FULL = ['January','February','March','April','May','June','July','August','September','October','November','December'];
  const DAYS = ['SU','M','T','W','TH','FR','SA'];

  /* month index */
  const rawM = String(mpo.month||'').trim().toUpperCase();
  let mIdx   = MN.indexOf(rawM.slice(0,3));
  if (mIdx<0) mIdx = FULL.findIndex(n=>n.toUpperCase()===rawM);
  const yr   = parseInt(mpo.year)||new Date().getFullYear();
  const dim  = mIdx>=0 ? new Date(yr,mIdx+1,0).getDate() : 31;
  const mLbl = mIdx>=0 ? MN[mIdx]+'-'+String(yr).slice(-2) : rawM;
  const getDN= d => DAYS[new Date(yr,mIdx,d).getDay()];

  /* expand spots to aired dates */
  const WDM={MON:1,TUE:2,WED:3,THU:4,FRI:5,SAT:6,SUN:0};
  const sWD = (mpo.spots||[]).map(s=>{
    let ad=[];
    if (s.calendarDays&&s.calendarDays.length) { ad=s.calendarDays.map(Number); }
    else if (s.wd) {
      const k=s.wd.toUpperCase();
      const set=k==='DAILY'?[0,1,2,3,4,5,6]:k==='WEEKDAYS'?[1,2,3,4,5]:k==='WEEKENDS'?[0,6]:WDM[k]!==undefined?[WDM[k]]:[];
      for(let d=1;d<=dim;d++) if(mIdx>=0&&set.includes(new Date(yr,mIdx,d).getDay())) ad.push(d);
    }
    return {...s,ad};
  });
  const dNums = Array.from({length:dim},(_,i)=>i+1);

  /* costing */
  const costLines = sWD.map(s=>({ programme:s.programme||'', material:s.material||'', duration:s.duration||'', cnt:s.ad.length||parseInt(s.spots)||0, rate:parseFloat(s.ratePerSpot)||0 }));
  costLines.forEach(l=>l.gross=l.cnt*l.rate);
  const subTotal  = costLines.reduce((a,l)=>a+l.gross,0);
  const vdPct     = parseFloat(mpo.discPct)||0;
  const vdAmt     = subTotal*vdPct;
  const afterDisc = subTotal-vdAmt;
  const cPct      = parseFloat(mpo.commPct)||0;
  const cAmt      = afterDisc*cPct;
  const afterComm = afterDisc-cAmt;
  const spPct     = parseFloat(mpo.surchPct)||0;
  const spAmt     = afterComm*spPct;
  const netAmt    = afterComm+spAmt;
  const vatRate   = (parseFloat(mpo.vatPct) || 7.5) / 100;
  const vatAmt    = netAmt * vatRate;
  const totalPayable = netAmt+vatAmt;

  /* ── PDF engine ──────────────────────────────────────────── */
  const PW=595,PH=842;
  const parts=['%PDF-1.4\n'], xref=[];
  const addObj=c=>{ xref.push(parts.reduce((a,b)=>a+b.length,0)); const id=xref.length; parts.push(`${id} 0 obj\n${c}\nendobj\n`); return id; };

  /* fonts — Helvetica built-in, no embed needed */
  const fR=addObj(`<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>`);
  const fB=addObj(`<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>`);
  const RES=`<< /Font << /Hr ${fR} 0 R /Hb ${fB} 0 R >> >>`;

  /* stream helpers */
  const pages=[];
  let   ops=[];
  const o  = t => ops.push(t);
  const BT = (x,y,sz,bold,R,G,B,txt)=>{
    o(`BT /${bold?'Hb':'Hr'} ${sz} Tf ${(R/255).toFixed(3)} ${(G/255).toFixed(3)} ${(B/255).toFixed(3)} rg ${x.toFixed(2)} ${(PH-y).toFixed(2)} Td (${esc(txt)}) Tj ET`);
  };
  // right-aligned: estimate glyph width at ~0.52*size per char for Helvetica
  const BTR=(rx,y,sz,bold,R,G,B,txt)=>{
    const w=String(txt).length*sz*0.52; BT(rx-w,y,sz,bold,R,G,B,txt);
  };
  // center-aligned
  const BTC=(cx,y,sz,bold,R,G,B,txt)=>{
    const w=String(txt).length*sz*0.52; BT(cx-w/2,y,sz,bold,R,G,B,txt);
  };
  const LN=(x1,y1,x2,y2,R=160,G=160,B=160,w=0.4)=>
    o(`${(R/255).toFixed(3)} ${(G/255).toFixed(3)} ${(B/255).toFixed(3)} RG ${w} w ${x1.toFixed(2)} ${(PH-y1).toFixed(2)} m ${x2.toFixed(2)} ${(PH-y2).toFixed(2)} l S`);
  const RECT=(x,y,w,h,fr,fg,fb,sr,sg,sb)=>{
    const yb=(PH-y-h).toFixed(2);
    if(fr!=null) o(`${(fr/255).toFixed(3)} ${(fg/255).toFixed(3)} ${(fb/255).toFixed(3)} rg ${x.toFixed(2)} ${yb} ${w.toFixed(2)} ${h.toFixed(2)} re f`);
    if(sr!=null) o(`${(sr/255).toFixed(3)} ${(sg/255).toFixed(3)} ${(sb/255).toFixed(3)} RG 0.4 w ${x.toFixed(2)} ${yb} ${w.toFixed(2)} ${h.toFixed(2)} re S`);
  };
  const newPage=()=>{ pages.push(ops.join('\n')); ops=[]; };
  const checkY=(need,margin=42)=>{ if(y>PH-margin-need){ newPage(); y=MT; } };

  const ML=28,MR=28,MT=32,MB=32,CW=PW-ML-MR;
  let y=MT;

  /* ══ PAGE 1 ══════════════════════════════════════════════ */

  /* --- header bar --- */
  RECT(ML,y,CW,16, 26,58,107, null,null,null);
  BTC(PW/2, y+11, 11,true,255,255,255, 'MEDIA PURCHASE ORDER');
  y+=20;

  /* agency address */
  BTC(PW/2, y+6, 7,false,80,80,80, mpo.agencyAddress || mpo.agency || '5, Craig Street, Ogudu GRA, Lagos');
  y+=12; LN(ML,y,PW-MR,y,26,58,107,0.7); y+=6;

  /* --- info grid --- */
  const LI=[
    ['CLIENT:',    (mpo.clientName||'—').toUpperCase()],
    ['BRAND:',     (mpo.brand||'—').toUpperCase()],
    ['CAMPAIGN:',  clip(mpo.campaignName||'—',36)],
    ['MEDIUM:',    (mpo.medium||'—').toUpperCase()],
    ['VENDOR:',    clip(mpo.vendorName||'—',32)],
  ];
  const RI=[
    ['MPO No:',   mpo.mpoNo||'—'],
    ['Date:',     mpo.date||'—'],
    ['Period:',   `${mpo.month||''} ${mpo.year||''}`],
    ['Status:',   (mpo.status||'DRAFT').toUpperCase()],
    ['Prepared:', clip(mpo.preparedBy||'—',24)],
  ];
  const iy=y;
  LI.forEach(([l,v],i)=>{ BT(ML,iy+i*9+7,6.5,true,100,100,100,l); BT(ML+36,iy+i*9+7,6.5,false,0,0,0,v); });
  RI.forEach(([l,v],i)=>{ BT(PW/2+4,iy+i*9+7,6.5,true,100,100,100,l); BT(PW/2+36,iy+i*9+7,6.5,false,0,0,0,v); });
  y=iy+LI.length*9+12;

  /* transmit bar */
  RECT(ML,y,CW,13,26,58,107,null,null,null);
  const tx=`PLEASE TRANSMIT SPOTS ON ${(mpo.vendorName||'VENDOR').toUpperCase()} AS SCHEDULED`;
  BTC(PW/2,y+9,7,true,255,255,255,tx);
  y+=17;

  /* ══ CALENDAR GRID ════════════════════════════════════════ */
  /* Group spots by time belt */
  const order=[],groups={};
  sWD.forEach(s=>{ const k=(s.timeBelt||'GENERAL').trim(); if(!groups[k]){groups[k]=[];order.push(k);} groups[k].push(s); });

  /* column widths: Month(20) | Belt(26) | Prog(32) | d1..dN(each 8) | #Spots(14) | Material(30) */
  const dW=7.5;  // per-day column width
  const fixedW=20+26+32+14+30;
  const tableW=fixedW+dim*dW;
  const tScale = tableW>CW ? CW/tableW : 1; // scale down if too wide
  const cMonth =20*tScale, cBelt=26*tScale, cProg=32*tScale, cDay=dW*tScale, cSpots=14*tScale, cMat=30*tScale;
  const tx0=ML, tx1=tx0+cMonth, tx2=tx1+cBelt, tx3=tx2+cProg;
  const txSpots=tx3+dim*cDay, txMat=txSpots+cSpots;
  const RH=9, HH=10;

  /* header row */
  checkY(HH+RH*2);
  RECT(ML,y,CW,HH,26,58,107,null,null,null);
  BT(tx0+1,y+7,5,true,255,255,255,'MONTH');
  BT(tx1+1,y+7,5,true,255,255,255,'Time Belt');
  BT(tx2+1,y+7,5,true,255,255,255,'Programme');
  dNums.forEach(d=>BT(tx3+(d-1)*cDay+1,y+7,4.5,true,255,255,255,String(d)));
  BT(txSpots+1,y+7,5,true,255,255,255,'#');
  BT(txMat+1,y+7,5,true,255,255,255,'Material');
  y+=HH;

  /* day-of-week sub-header */
  RECT(ML,y,CW,7,238,244,252,null,null,null);
  if(mIdx>=0) dNums.forEach(d=>BT(tx3+(d-1)*cDay+1,y+5,4,false,80,80,80,getDN(d)));
  y+=7;

  /* data rows */
  let grandPaid=0;
  order.forEach(belt=>{
    const rows=groups[belt];
    let bTotal=0;
    rows.forEach((s,si)=>{
      const cnt=s.ad.length||parseInt(s.spots)||0;
      bTotal+=cnt; grandPaid+=cnt;
      checkY(RH+2);
      const bg=si%2===0?[252,252,255]:[245,248,253];
      RECT(ML,y,CW,RH,...bg,null,null,null);
      LN(ML,y+RH,ML+CW,y+RH,200,200,200,0.3);
      /* month cell (only first row of belt group) */
      if(si===0){
        BT(tx0+1,y+6,5,true,26,58,107,mLbl);
        BT(tx1+1,y+6,5,false,40,40,40,clip(belt,8));
      }
      BT(tx2+1,y+6,5,false,20,20,20,clip(s.programme||'',12));
      /* day dots */
      dNums.forEach(d=>{
        if(s.ad.includes(d)){
          RECT(tx3+(d-1)*cDay+1,y+2,cDay-2,RH-4,26,58,107,null,null,null);
          BT(tx3+(d-1)*cDay+1.5,y+7,4,true,255,255,255,'1');
        }
      });
      /* spot count */
      RECT(txSpots,y,cSpots,RH,220,232,244,null,null,null);
      BTC(txSpots+cSpots/2,y+6,6,true,26,58,107,String(cnt));
      /* material */
      BT(txMat+1,y+6,4.5,false,60,60,60,clip(s.material||'',14));
      /* row border */
      LN(ML,y,ML+CW,y,200,200,200,0.2);
      y+=RH;
    });
    /* belt subtotal if multiple rows */
    if(rows.length>1){
      RECT(ML,y,CW,8,220,232,244,null,null,null);
      BTR(txSpots+cSpots-1,y+5.5,6,true,26,58,107,String(bTotal));
      y+=8;
    }
  });

  /* grand total row */
  RECT(ML,y,CW,10,26,58,107,null,null,null);
  BT(ML+2,y+7,6,true,255,255,255,'GRAND TOTAL');
  BTR(txSpots+cSpots-1,y+7,7,true,255,255,255,String(grandPaid));
  y+=14;

  /* ══ COSTING TABLE ════════════════════════════════════════ */
  checkY(60);
  BTC(PW/2,y+7,9,true,26,58,107,'C  O  S  T  I  N  G');
  y+=12; LN(ML,y,PW-MR,y,26,58,107,0.6); y+=5;

  /* per-spot cost lines */
  const cH=10;
  RECT(ML,y,CW,cH,26,58,107,null,null,null);
  [['Programme',50],['Material',60],['Duration',20],['Spots',18],['Rate/Spot (N)',42],['Total (N)',42]].reduce((x,[h,w])=>{
    BT(x+1,y+7,5.5,true,255,255,255,h); return x+w;
  },ML);
  y+=cH;

  costLines.forEach((l,ri)=>{
    checkY(cH+2);
    RECT(ML,y,CW,cH,ri%2===0?252:246,ri%2===0?252:248,ri%2===0?255:252,210,210,210);
    let cx=ML;
    [[clip(l.programme,16),50,false],[clip(l.material,20),60,false],[String(l.duration)+'s',20,false],
     [String(l.cnt),18,true],[fmtN(l.rate),42,true],[fmtN(l.gross),42,true]].forEach(([v,w,b])=>{
      BT(cx+1,y+7,5.5,b,20,20,20,v); cx+=w;
    });
    y+=cH;
  });

  /* costing summary */
  const summaryRows=[
    ['Sub Total',                                          fmtN(subTotal),    false,false],
    ...(vdPct>0?[
      [`Volume Discount (${Math.round(vdPct*100)}%)`,     `- ${fmtN(vdAmt)}`,false,false],
      ['Less Discount',                                    fmtN(afterDisc),   false,false],
    ]:[]),
    ...(cPct>0?[
      [`Agency Commission (${Math.round(cPct*100)}%)`,    `- ${fmtN(cAmt)}`, false,false],
      ['Less Commission',                                  fmtN(afterComm),   false,false],
    ]:[]),
    ...(spPct>0?[
      [mpo.surchLabel||`Surcharge (${Math.round(spPct*100)}%)`, `+ ${fmtN(spAmt)}`,false,false],
      ['Net After Surcharge',                              fmtN(netAmt),      false,false],
    ]:[]),
    [`VAT (${parseFloat(mpo.vatPct) || 7.5}%)`,                    fmtN(vatAmt),      false,false],
    ['TOTAL AMOUNT PAYABLE',                              fmtN(totalPayable),true, true],
  ];
  y+=4;
  summaryRows.forEach(([label,val,bold,tot])=>{
    checkY(11);
    const rh=11;
    if(tot) RECT(ML,y,CW,rh,26,58,107,null,null,null);
    else    RECT(ML,y,CW,rh,248,250,255,215,215,215);
    const [cr,cg,cb]=tot?[255,255,255]:bold?[0,0,0]:[55,55,55];
    BT(ML+4,y+7.5,7,bold||tot,cr,cg,cb,label);
    BTR(PW-MR-3,y+7.5,7,bold||tot,cr,cg,cb,String(val));
    y+=rh;
  });
  y+=8;

  /* ══ CONTRACT TERMS ═══════════════════════════════════════ */
  checkY(30);
  BT(ML,y+6,7.5,true,0,0,0,'Contract Terms & Conditions');
  y+=12;
  const terms = Array.isArray(mpo.terms) && mpo.terms.length ? mpo.terms : DEFAULT_APP_SETTINGS.mpoTerms;
  terms.forEach((t,i)=>{
    checkY(10);
    BT(ML,y+6,6,false,50,50,50,`${i+1}.  ${clip(t,90)}`);
    y+=9;
  });
  y+=8;

  /* ══ SIGNATURES ═══════════════════════════════════════════ */
  checkY(38);
  const sy=y+22;
  LN(ML,sy,ML+52,sy,80,80,80,0.5);
  LN(ML+60,sy,ML+112,sy,80,80,80,0.5);
  LN(ML+120,sy,PW-MR,sy,80,80,80,0.5);
  BT(ML,    sy+5,6,true, 0,0,0,'For (Media House / Supplier)');
  BT(ML+60, sy+5,6,true, 0,0,0,`SIGNED BY: ${clip((mpo.signedBy||'').toUpperCase(),20)}`);
  BT(ML+120,sy+5,6,true, 0,0,0,`PREPARED BY: ${clip((mpo.preparedBy||'').toUpperCase(),18)}`);
  BT(ML,    sy+11,5.5,false,100,100,100,'Name / Signature / Official Stamp');
  BT(ML+60, sy+11,5.5,false,100,100,100,clip(mpo.signedTitle||'',22));
  BT(ML+120,sy+11,5.5,false,100,100,100,clip(mpo.preparedTitle||'',22));
  BT(ML+60, sy+16,5.5,false,100,100,100,clip(mpo.preparedContact||'',24));

  newPage(); // flush last page

  /* ══ ASSEMBLE PDF FILE ════════════════════════════════════ */
  const sids = pages.map(ps=>addObj(`<< /Length ${ps.length} >>\nstream\n${ps}\nendstream`));
  const kids = sids.map(sid=>addObj(`<< /Type /Page /MediaBox [0 0 ${PW} ${PH}] /Contents ${sid} 0 R /Resources ${RES} >>`));
  const pagesId=addObj(`<< /Type /Pages /Kids [${kids.map(i=>`${i} 0 R`).join(' ')}] /Count ${kids.length} >>`);
  /* re-emit pages with /Parent */
  const fkids=kids.map((_,pi)=>addObj(`<< /Type /Page /Parent ${pagesId} 0 R /MediaBox [0 0 ${PW} ${PH}] /Contents ${sids[pi]} 0 R /Resources ${RES} >>`));
  const catId=addObj(`<< /Type /Catalog /Pages ${pagesId} 0 R >>`);

  const body=parts.join('');
  const xpos=body.length;
  const n=xref.length+1;
  const xs=`xref\n0 ${n}\n0000000000 65535 f \n`+xref.map(o=>o.toString().padStart(10,'0')+' 00000 n ').join('\n')+'\n';
  const tr=`trailer\n<< /Size ${n} /Root ${catId} 0 R >>\nstartxref\n${xpos}\n%%EOF`;
  const full=body+xs+tr;
  const bytes=new Uint8Array(full.length);
  for(let i=0;i<full.length;i++) bytes[i]=full.charCodeAt(i)&0xff;
  return bytes;
};

const PrintPreview = ({ html, csv, pdfBytes, title, onClose }) => {
  const [tab, setTab]       = useState("preview");
  const [copied, setCopied] = useState(false);
  const [pdfBusy, setPdfBusy] = useState(false);
  const iframeRef           = useRef(null);

  const handlePrint = () => {
    iframeRef.current?.contentWindow?.postMessage("print-mpo", "*");
  };

  const safeName = t => (t || "MPO").replace(/[^a-z0-9\-_. ]/gi, "_").slice(0, 80);

  const downloadFallbackPdf = () => {
    if (!pdfBytes) return;
    const blob = new Blob([pdfBytes], { type: "application/pdf" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `${safeName(title)}.pdf`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1500);
  };

  const handleDownloadPDF = async () => {
    if (pdfBusy) return;
    setPdfBusy(true);
    try {
      const iframe = iframeRef.current;
      const doc = iframe?.contentDocument;
      const body = doc?.body;
      const root = doc?.documentElement;
      if (!doc || !body || !root) {
        downloadFallbackPdf();
        return;
      }

      const { html2canvas, jsPDF } = await loadPreviewPdfLibraries();
      await new Promise(resolve => setTimeout(resolve, 160));
      const targetWidth = Math.max(body.scrollWidth, root.scrollWidth, iframe.clientWidth || 0);
      const targetHeight = Math.max(body.scrollHeight, root.scrollHeight, iframe.clientHeight || 0);
      const canvas = await html2canvas(body, {
        scale: 2,
        useCORS: true,
        backgroundColor: "#ffffff",
        windowWidth: targetWidth,
        windowHeight: targetHeight,
        width: targetWidth,
        height: targetHeight,
        scrollX: 0,
        scrollY: 0,
      });

      const pdf = new jsPDF("p", "mm", "a4");
      const pageWidth = pdf.internal.pageSize.getWidth();
      const pageHeight = pdf.internal.pageSize.getHeight();
      const imgData = canvas.toDataURL("image/png");
      const imgWidth = pageWidth;
      const imgHeight = (canvas.height * imgWidth) / canvas.width;
      let heightLeft = imgHeight;
      let position = 0;

      pdf.addImage(imgData, "PNG", 0, position, imgWidth, imgHeight, undefined, "FAST");
      heightLeft -= pageHeight;
      while (heightLeft > 0) {
        position = heightLeft - imgHeight;
        pdf.addPage();
        pdf.addImage(imgData, "PNG", 0, position, imgWidth, imgHeight, undefined, "FAST");
        heightLeft -= pageHeight;
      }
      pdf.save(`${safeName(title)}.pdf`);
    } catch (error) {
      console.error("Exact preview PDF export failed:", error);
      downloadFallbackPdf();
    } finally {
      setPdfBusy(false);
    }
  };

  const handleDownloadHTML = () => {
    const clean = html.replace(/<script>[\s\S]*?window\.addEventListener[\s\S]*?<\/script>/,"");
    const blob  = new Blob([clean], { type: "text/html;charset=utf-8" });
    const url   = URL.createObjectURL(blob);
    const a     = document.createElement("a");
    a.href = url; a.download = `${safeName(title)}.html`;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 1500);
  };

  const handleDownloadCSV = () => {
    if (!csv) return;
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement("a");
    a.href = url; a.download = `${safeName(title)}.csv`;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 1500);
  };

  const handleCopy = () => {
    if (!csv) return;
    navigator.clipboard.writeText(csv).then(() => {
      setCopied(true); setTimeout(() => setCopied(false), 2500);
    }).catch(() => {
      const ta = document.getElementById("csv-ta");
      if (ta) { ta.select(); document.execCommand("copy"); setCopied(true); setTimeout(() => setCopied(false), 2500); }
    });
  };

  const btnStyle = (bg, color, border) => ({
    display:"flex", alignItems:"center", gap:6, padding:"7px 16px",
    background:bg, color, border:`1px solid ${border}`, borderRadius:8,
    fontFamily:"'Syne',sans-serif", fontWeight:700, fontSize:12, cursor:"pointer"
  });

  return (
    <div style={{ position:"fixed", inset:0, zIndex:2000, background:"rgba(0,0,0,.88)", display:"flex", flexDirection:"column", backdropFilter:"blur(6px)" }}>
      <div style={{ background:"#0e1118", borderBottom:"1px solid rgba(255,255,255,.1)", padding:"11px 18px", display:"flex", alignItems:"center", gap:10, flexShrink:0, flexWrap:"wrap" }}>
        <div style={{ fontFamily:"'Syne',sans-serif", fontWeight:700, fontSize:14, flex:1, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{title}</div>
        <div style={{ display:"flex", gap:0, background:"#141824", borderRadius:7, overflow:"hidden", border:"1px solid rgba(255,255,255,.1)" }}>
          {[ ["preview","📄 Preview / Print"], ["csv","📊 CSV / Excel"] ].map(([t,l]) => (
            <button key={t} onClick={() => setTab(t)} style={{ padding:"7px 14px", border:"none", background:tab===t?"#f0a500":"transparent", color:tab===t?"#000":"#8b93a7", fontFamily:"'Syne',sans-serif", fontWeight:600, fontSize:12, cursor:"pointer" }}>{l}</button>
          ))}
        </div>
        {tab === "preview" && (<>
          <button onClick={handleDownloadPDF} disabled={pdfBusy} style={{ ...btnStyle("rgba(34,197,94,.15)","#22c55e","rgba(34,197,94,.35)"), opacity: pdfBusy ? 0.65 : 1, cursor: pdfBusy ? "wait" : "pointer" }}>{pdfBusy ? "⏳ Building Exact PDF…" : "⬇ Download Exact PDF"}</button>
          <button onClick={handlePrint} style={btnStyle("#0A1F44","#D4870A","#D4870A")}>🖨 Print</button>
          <button onClick={handleDownloadHTML} style={btnStyle("rgba(139,92,246,.15)","#a78bfa","rgba(139,92,246,.4)")}>⬇ Download HTML</button>
        </>)}
        {tab === "csv" && csv && (<>
          <button onClick={handleDownloadCSV} style={btnStyle("rgba(34,197,94,.15)","#22c55e","rgba(34,197,94,.4)")}>⬇ Download CSV</button>
          <button onClick={handleCopy} style={btnStyle(copied?"rgba(34,197,94,.25)":"rgba(255,255,255,.05)", copied?"#22c55e":"#8b93a7", copied?"rgba(34,197,94,.4)":"rgba(255,255,255,.12)")}>{copied ? "✓ Copied!" : "📋 Copy to Clipboard"}</button>
        </>)}
        <button onClick={onClose} style={btnStyle("rgba(239,68,68,.15)","#ef4444","rgba(239,68,68,.3)")}>✕ Close</button>
      </div>
      {tab === "preview" && (
        <div style={{ flex:1, overflow:"auto", background:"#ccc", display:"flex", justifyContent:"center", padding:24 }}>
          <iframe
            ref={iframeRef}
            srcDoc={html}
            style={{ width:"100%", maxWidth:900, border:"none", boxShadow:"0 8px 40px rgba(0,0,0,.5)", background:"#fff", minHeight:700 }}
            title="MPO Preview"
            onLoad={e => { try { e.target.style.height = e.target.contentDocument.body.scrollHeight + 60 + "px"; } catch {} }}
          />
        </div>
      )}
      {tab === "csv" && csv && (
        <div style={{ flex:1, overflow:"hidden", display:"flex", flexDirection:"column", padding:20, gap:10 }}>
          <div style={{ background:"rgba(34,197,94,.08)", border:"1px solid rgba(34,197,94,.2)", borderRadius:9, padding:"11px 15px", fontSize:13, color:"#22c55e" }}>
            💡 Click <strong>Download CSV</strong> to save, then open in Excel or Google Sheets. Or <strong>Copy to Clipboard</strong> and paste directly.
          </div>
          <textarea id="csv-ta" readOnly value={csv}
            style={{ flex:1, background:"#0e1118", border:"1px solid rgba(255,255,255,.1)", borderRadius:9, padding:14, color:"#8b93a7", fontFamily:"monospace", fontSize:11, resize:"none", outline:"none", lineHeight:1.6 }} />
        </div>
      )}
    </div>
  );
};

const buildMPOHTML = (mpo) => {
  const {
    mpoNo, date, month, year, vendorName, clientName, brand, campaignName,
    agencyAddress, agencyEmail, agencyPhone, signedBy, signedTitle, signedSignature, preparedBy, preparedContact, preparedTitle, preparedSignature,
    spots, discPct, commPct, surchPct, surchLabel, terms = DEFAULT_APP_SETTINGS.mpoTerms, vatPct = 7.5,
  } = mpo;

  const fmt = n => Number(n||0).toLocaleString("en-NG",{minimumFractionDigits:2,maximumFractionDigits:2});


  const MN   = ["JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"];
  const FULL = ["January","February","March","April","May","June","July","August","September","October","November","December"];
  const DAY  = ["SU","M","T","W","TH","FR","SA"];
  const yr   = parseInt(year) || new Date().getFullYear();

  const resolveMIdx = m => {
    const u = (m||"").trim().toUpperCase();
    let i = MN.indexOf(u.slice(0,3));
    if (i < 0) i = FULL.findIndex(n => n.toUpperCase() === u);
    return i;
  };

  /* Determine which months to render */
  const allMonths = (mpo.months && mpo.months.length > 0) ? mpo.months : (month ? [month] : []);

  /* Build one calendar section for a given month and its spots */
  const buildMonthBlock = (monthName, monthSpots) => {
    const mIdx = resolveMIdx(monthName);
    const dim  = mIdx >= 0 ? new Date(yr, mIdx+1, 0).getDate() : 31;
    const getDN = d => DAY[new Date(yr, mIdx, d).getDay()];
    const monthLabel = mIdx >= 0 ? MN[mIdx]+"-"+String(yr).slice(-2) : (monthName||"").toUpperCase().slice(0,6);

    const WD = {MON:1,TUE:2,WED:3,THU:4,FRI:5,SAT:6,SUN:0};
    const sWD = monthSpots.map(s => {
      let ad = [];
      if (s.calendarDays && s.calendarDays.length) {
        ad = s.calendarDays.map(Number);
      } else if (s.wd) {
        const k = s.wd.toUpperCase();
        const set = k==="DAILY"?[0,1,2,3,4,5,6]:k==="WEEKDAYS"?[1,2,3,4,5]:k==="WEEKENDS"?[0,6]:WD[k]!==undefined?[WD[k]]:[];
        for(let d=1;d<=dim;d++) if(mIdx>=0 && set.includes(new Date(yr,mIdx,d).getDay())) ad.push(d);
      }
      return {...s, ad};
    });

    const dNums   = Array.from({length:dim},(_,i)=>i+1);
    const dateRow = dNums.map(d =>
      '<th style="background:#e8f0f8;color:#000;font-size:6.5px;padding:1px 0;text-align:center;border:1px solid #aaa;min-width:14px;width:14px;font-weight:700">'+d+'</th>').join("");
    const dayRow  = mIdx >= 0
      ? dNums.map(d => '<td style="background:#eef3fa;font-size:6px;padding:1px 0;text-align:center;border:1px solid #aaa;font-weight:500;color:#444">'+getDN(d)+'</td>').join("")
      : dNums.map(()=>'<td style="border:1px solid #aaa"></td>').join("");

    const order = [], groups = {};
    sWD.forEach(s => {
      const k = (s.timeBelt||"GENERAL").trim();
      if (!groups[k]) { groups[k]=[]; order.push(k); }
      groups[k].push(s);
    });

    let calRowsHtml = "";
    let grandPaid   = 0;

    order.forEach(belt => {
      const rows = groups[belt];
      let bTotal = 0;
      rows.forEach((s, si) => {
        const cnt = s.ad.length || parseInt(s.spots)||0;
        bTotal    += cnt;
        grandPaid += cnt;
        const cells = dNums.map(d =>
          '<td style="text-align:center;font-size:7.5px;padding:1px 0;border:1px solid #ddd;font-weight:700">'+(s.ad.includes(d)?"1":"")+'</td>'
        ).join("");
        const isFirst = si === 0;
        const monthTd = isFirst
          ? '<td rowspan="'+rows.length+'" style="font-size:7.5px;padding:2px 3px;border:1px solid #aaa;font-weight:700;text-align:center;vertical-align:middle;white-space:nowrap;background:#f5f8fd">'+monthLabel+'</td><td rowspan="'+rows.length+'" style="font-size:7.5px;padding:2px 5px;border:1px solid #aaa;font-weight:700;vertical-align:middle;white-space:nowrap;background:#f5f8fd">'+belt+'</td>'
          : "";
        calRowsHtml += '<tr>'+monthTd+'<td style="font-size:7.5px;padding:2px 5px;border:1px solid #aaa;white-space:nowrap">'+(s.programme||"")+'</td>'+cells+'<td style="text-align:center;font-weight:700;font-size:8px;padding:2px 3px;border:1px solid #aaa;background:#dce8f4">'+cnt+'</td><td style="font-size:7.5px;padding:2px 5px;border:1px solid #aaa;white-space:nowrap">'+(s.material||"")+'</td></tr>';
      });
      if (rows.length > 1) {
        calRowsHtml += '<tr style="background:#dce8f4"><td colspan="3" style="border:1px solid #aaa;padding:2px 4px;font-weight:700;font-size:7.5px;text-align:right"></td>'+dNums.map(()=>'<td style="border:1px solid #aaa"></td>').join("")+'<td style="text-align:center;font-weight:800;font-size:9px;border:1px solid #aaa;padding:2px 3px;background:#c4d8ee">'+bTotal+'</td><td style="border:1px solid #aaa"></td></tr>';
      }
    });

    const grandRow = '<tr style="background:#fff"><td colspan="'+(dim+5)+'" style="border:1px solid #aaa;padding:3px 6px;font-weight:800;font-size:10px;text-align:right">'+grandPaid+'</td></tr>';

    const headerHtml =
      '<div style="font-family:Arial,Helvetica,sans-serif;font-size:8.5px;font-weight:700;color:#1a3a6b;text-align:center;margin:10px 0 4px;letter-spacing:2px;text-transform:uppercase;border-bottom:1px solid #c0d0e8;padding-bottom:3px">'+monthLabel+' SCHEDULE</div>' +
      '<div class="cal-wrap"><table class="cal"><thead><tr>' +
      '<th style="background:#1a3a6b;color:#fff;font-size:7.5px;padding:3px;text-align:center;min-width:40px">MONTH</th>' +
      '<th style="background:#1a3a6b;color:#fff;font-size:7.5px;padding:3px 5px;text-align:left;min-width:70px">Time Belt</th>' +
      '<th style="background:#1a3a6b;color:#fff;font-size:7.5px;padding:3px 5px;text-align:left;min-width:80px">Programme</th>' +
      '<th colspan="'+dim+'" style="background:#1a3a6b;color:#fff;font-size:8px;padding:3px;text-align:center;letter-spacing:5px">SCHEDULE</th>' +
      '<th style="background:#1a3a6b;color:#fff;font-size:7px;padding:3px 2px;text-align:center;min-width:40px;line-height:1.4">NO OF<br>SPOTS</th>' +
      '<th style="background:#1a3a6b;color:#fff;font-size:7px;padding:3px 4px;text-align:left;min-width:90px;line-height:1.4">MATERIAL TITLE/<br>SPECIFICATION</th>' +
      '</tr><tr>' +
      '<td style="background:#dce8f4;font-size:7px;padding:2px 3px;text-align:center;font-weight:700;border:1px solid #aaa">DATES&#8594;</td>' +
      '<td style="background:#dce8f4;border:1px solid #aaa"></td>' +
      '<td style="background:#dce8f4;border:1px solid #aaa"></td>' +
      dateRow +
      '<td style="background:#dce8f4;border:1px solid #aaa"></td>' +
      '<td style="background:#dce8f4;border:1px solid #aaa"></td>' +
      '</tr><tr>' +
      '<td style="background:#eef3fa;border:1px solid #aaa"></td>' +
      '<td style="background:#eef3fa;border:1px solid #aaa"></td>' +
      '<td style="background:#eef3fa;border:1px solid #aaa"></td>' +
      dayRow +
      '<td style="background:#eef3fa;border:1px solid #aaa"></td>' +
      '<td style="background:#eef3fa;border:1px solid #aaa"></td>' +
      '</tr></thead><tbody>' + calRowsHtml + grandRow + '</tbody></table></div>';

    return { html: headerHtml, sWD };
  };

  /* Accumulate all months */
  let allCalendarHTML = "";
  let allSWD = [];

  if (allMonths.length === 0) {
    allCalendarHTML = '<div style="padding:10px;color:#888;font-style:italic;text-align:center">No schedule months configured.</div>';
  } else {
    allMonths.forEach(monthName => {
      const monthSpots = (spots||[]).filter(s => {
        if (!s.scheduleMonth) return true;
        const sm = (s.scheduleMonth||"").toLowerCase();
        const mn = (monthName||"").toLowerCase();
        return sm.startsWith(mn.slice(0,3)) || sm.includes(mn);
      });
      if (monthSpots.length === 0) return;
      const { html, sWD } = buildMonthBlock(monthName, monthSpots);
      allCalendarHTML += html;
      allSWD = allSWD.concat(sWD);
    });
    if (!allCalendarHTML) {
      allCalendarHTML = '<div style="padding:10px;color:#888;font-style:italic;text-align:center">No spots scheduled yet.</div>';
    }
  }

  const firstDur = allSWD.length > 0 ? (allSWD[0].duration||"30")+"SECS" : "30SECS";

  const costLines = allSWD.map(s => {
    const cnt  = s.ad ? (s.ad.length || parseInt(s.spots)||0) : (parseInt(s.spots)||0);
    const rate = parseFloat(s.ratePerSpot)||0;
    return { programme: s.programme||"", material: s.material||"", duration: s.duration||"",
             cnt, rate, gross: cnt * rate };
  });

  const subTotal   = costLines.reduce((a,l)=>a+l.gross, 0);
  const vdPct      = parseFloat(discPct)||0;
  const vdAmt      = subTotal * vdPct;
  const afterDisc  = subTotal - vdAmt;
  const cPct       = parseFloat(commPct)||0;
  const cAmt       = afterDisc * cPct;
  const afterComm  = afterDisc - cAmt;
  const VAT_RATE   = (parseFloat(vatPct) || 7.5) / 100;
  const spPct      = parseFloat(surchPct)||0;
  const spAmt      = afterComm * spPct;
  const netAmt     = afterComm + spAmt;
  const vatAmt     = netAmt * VAT_RATE;
  const totalPayable = netAmt + vatAmt;

  const periodLabel = allMonths.length > 1
    ? allMonths.map(m => (m||"").toUpperCase().slice(0,3)).join("/") + " " + yr
    : (allMonths[0]||month||"").toUpperCase() + " " + yr;

  const LOGO_SRC = `data:image/jpeg;base64,/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAIBAQEBAQIBAQECAgICAgQDAgICAgUEBAMEBgUGBgYFBgYGBwkIBgcJBwYGCAsICQoKCgoKBggLDAsKDAkKCgr/2wBDAQICAgICAgUDAwUKBwYHCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgr/wAARCACcAMgDASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwD9/KKKKACq+qatpeiafNq2s6jBaWtvGZLi5uZljjiQDJZmYgKB6k4rxP8Abf8A2/Pgn+w34GGu/EDUPtuuXsbf2H4Zs3Bub1h/Ef8AnnGD1c8dhk1+J/7Yv/BRX9pf9s3xBcP8QfGlxYeHDOXsPCWlTtFZQLn5d6jBmcD+N898AV4ea59hMsfJ8U+y6er6fmfrnh54O8ScfJYpfuMLe3tZJvm7qEdOZrvdRT0vfQ/V/wDaL/4LffsU/A26m0Pw14luvG+qQsyPb+GYw8COM8NO2E68ZXNfIfxM/wCDj343alcywfCv4D+H9Jtm/wBRPq97Jczr16qm1D271+cTjGABgAcAdqhlHHFfHV+Jc0xD92SivJfq7s/p/KPAbw+yamlWoyxE19qpJ2+UY8sbeqfqfadx/wAHAH/BQtpmaDVvBkaE5VP+EV3bR6Z83mtnwf8A8HFP7anh+f8A4rHwd4M19dwJX+z5LPI9Mo7V8Gv1qtd9VOO1YwzfM00/ay+89bF+GXAE6bj/AGbSS8o2f3qz/E/YP4Hf8HI3wb8RXUGlfHz4N6r4bZ9qy6lo1wLyAE9TsOHCj86+7v2f/wBqz9nz9qPw7/wk3wK+KmleIIVQNcQWk+Li3yBxJC2HTrjJGM9Ca/mFm/pWv8PPib8Q/hB4ttfH3ws8ban4e1qycNbanpN40Mye2VPzD1ByD0Ir2cHxJi6btWXMvuf+R+X8SeA3DWOpSnlU5Yep0TbnB+TT95eqk7dmf1S0V+XX/BNn/gvZp/ja6074L/ts39pp+qSlbfT/AB0kYitrp84VbpR8sLHp5gwhPJC1+oVtc295Al1azJJHIoaORGBVlIyCCOoI719fhMbh8dS56Tv37r1P5g4l4VzrhLH/AFXMafK/syWsZLvF9fTddUh9FFFdZ86FFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFeF/t8/tv+Bv2H/gtc+PNaEV9rt6rQeGtDMmGvLnHBbuI16sfTgcmvbNX1bTtB0q51vWLyO2tLOB57q4lbCxRopZnY9gACT9K/no/4KDftf8Aif8AbN/aM1b4j39zImh2c0lj4V04uSltZI5CvjpvkxvY+4HavA4gzb+y8J7nxy0Xl3fy/M/YvBfw2/4iFxI/rSf1TD2lV6c137tNP+9Z3fSKdrOx5l8bPjJ8R/j98RtS+K/xV8TT6rrWqzmS5uJm4Qfwxxr0SNRwFHAArkX61PP9wfWu5/Zq/Zn+Kn7WPxY0/wCEPwk0Y3N/eNuuLmQEQWUA+9NK38Kj8yeBX5fTVbE1kleUpP1bbP8AQfFTyzIstlKTjRoUY+UYwjFfckl0PPVilnlS3giZ5JHCRxopZnY9FAHJJ7AcmvpT4C/8Eg/26v2g7C31zRvhK+gaVc4aLU/FU4s1ZCMhxE37xl9wtfrH+w//AMEnv2b/ANjrTrTxDLocPijxoIR9r8S6vCJPKc8kW0bDbCuccj5jjk19SBQvQV9xl/CfuKWKlr2X6v8Ay+8/knjP6SMvrEsPw7QTitPa1U9fOME1ZdnJ69Yo/G3Sv+DcP9pO5shJrPxq8I20/wDFFFHPIo+h2iuK+KH/AAb5/tx+EbCTUvB134X8UCPJS1sNTME7D6TKF/AHtX7kUV674ayvlsk18z80p+PXiDGtz1KlOa/ldNJf+S2f4n8uHxk+BXxi+APiZvBvxo+GuseGdRBIS31eyaLzgP4o2PyyDHPyk8Vxr9K/qU+MfwK+Ef7QHg248A/GP4f6Z4h0q4Qq1rqVsH2E/wASN96Nv9pSDX4xf8FW/wDgjbq37IFhcfHj4CXF5rHw/M4GpWVyxkutD3Hhmb/lpBnjeeVJG71rwcwyCtgoOpTfNFfev8z9k4H8Z8q4orwwOPh7DES0jreE32TesW+id77KTeh8By/dA9Tg1+of/BDb/gqpqPhXW9I/Yn+P+vmbSLx/s/gXXL64JNlKfuWDs3/LNukZJ+U/L0Ix+XkhwB9aak81tMlxbTvFJG4aOWJyrIwIIYEcgggEEdCK87BYurgqyqQ+a7rsfc8W8NZdxXlVTA4tb6xl1hLpJenVdVdH9ZAORkUV8hf8EZ/25739s39lu3tfHWppP408GtHpniKQgBrtAv7i7wO7oMN/tK3rX17X6Th69PE0Y1YbM/g3OcpxmRZpVwGKVqlN2fn2a8mrNeTCiiitjzAooooAKKKKACiigkDqaACik3D/APUKN6jk5A9SKAFopsc8MpxHKrH0DA0UBe58ef8ABbn9oa4+Cv7Gt94T0W/MOqeN7tdIhMb4cW5G6cjvjaNp/wB6vw2kAAAAwAeBX6L/APBxL8SbjVvjj4I+FsN3m30jQJb6aHH3ZppNob8UU1+dEn3fxr8n4pxTxGcSj0hZL83+LP8ARr6PeQUsl8NKFe3v4mUqsn8+WPy5Yp/NjDFLcFYoInkd2CpHGuWdicBQO5JOAO5Nfvj/AMEp/wBibRv2Qv2a9Lk1vQ4ovGviW0jvvFN2QC6M3zR2wb+7GpAx/ezX5Q/8EmPgPY/H79uTwhoWuWIuNL0OZ9b1KJk3K62wDRqw9DJsFfv6q7VxXu8HYCLjLFzWvwx/V/p95+S/Se4xrwrYfhvDytFpVatuurUIvyVnJrvyvoLRRRX3Z/H4UUUUAFU/EPh/RPFeh3fhrxJpVvfWF/bvBeWd1GHjmiYFWRlPBBBIxVyihpMabi7rc/nR/wCCrP7Dd5+w7+09feFNGtpD4Q8Qh9T8IXDchbcvh7Ynu0THb7qUPevmOTpX7uf8HBXwD074n/sRy/FGG0T+0/AeqRX0M+35vs0h8qZM9cHcpx6jNfhJIPlx6GvzrNsHHBY2UY/C9V8/+Cf3H4acT1uKuE6Veu71ad6c33cbWb83FpvzufXv/BEH9pi4/Z7/AG7PD2h6hqZh0TxyDoWpRs5CGWTm3cjuRKAo/wB81/QYDkZFfyg+GPE2oeCvEuneMdKuWhudJv4b2CVDgo0Th8j/AL5r+qH4X+MoviJ8NvD/AI/gi2JrmiWmoIn90TwpKB/4/XvcNV3KjOk+jv8Af/wx+MePWUQw+a4XMYL+LFxl6wtZ/dK3yN2iiivpz8CCignFfK37Zf8AwVo/Z0/ZRkuvCGlXq+LvF8GVbQtIuBstm9J5hlYz/sjLe1dmBy/G5niFQwtNzk+i/N9EvNnn5nmuXZNhXiMbVVOC6v8AJLdvyWp9TySJEpeRgABkkngD1rw344f8FH/2P/gHJPp/i34u2N7qVu22TSNEb7XcBv7pCcKfqRX5I/tP/wDBTf8Aas/aovJ7fXPHE3h3w/Ix8nw14bne3hC+ksgIkmPXliByRjFeE6SeZD6kE+/Wv1bJvC3nSnmVW392H6yf6L5n4hxD41ezcoZRQTS+3Uvr6RTX4v5H6XfFz/gvvdtNPp/wO+CKBOkGo+I705Jz1MMXb/gVeC+O/wDgr1+3F4+nkgtPiLZ+H7aYY8jQdLjjZP8Adkfc4/OvlhPvj61atv8AXr9a/RMBwXwxgbcmGi33l7z/APJr/gfkmZ+IXGWaN+1xk4p9IPkX/ktn97Z6X4t/ap/aX8eFT4x+Pni3UNn3fP1uUY/74IrLtvHPjm5tENz4612Qkc79cuTn85K5WtbTwfscfHb+tfS08Hg6MeWnTil2SSX4Hx1bMMfXnz1KspN9XJt/e2fbX/BFTWtb1D9pXXIdR1y/uUHhVyEub6WVQfOTnDMRmiq3/BEkEftM65kf8yo//o5KK/nXxNjGPFMkl9iP5H9ZeDU5z4Ji5O79pP8ANHh3/BeZ2b9v67RmJC+D9L2jPTPnZr4wcZWvu7/g4J8ETaD+2No/jNmJTX/B0AXnobeR0P8A6HXwi33T9K/ljPIuOcV0/wCZ/jqf7P8AhHVp1/DPKpQd17GK+cfdf3NNH6Af8G8Gn2cn7UPivVHA8+LwiY4zn+FpkJ/UCv2Nr8PP+CEXxOsPAf7dNr4c1NyqeKvD91p8DZwomXEq59zsIH1r9wxyM19/wlOMsoSXST/zP42+klhq1HxLnUmvdnSpuPok4v8AFMKKKK+nPwIKKKKAKPiPxL4e8H6Jc+JfFeuWmm6dZx+Zd319OsUUKZxuZ2ICjnqasafqOn6tZRalpd9Dc208YkguIJA6SIRkMrDgg+or8sv+C4P7ecHirWn/AGOPhjqxaz0ydZPGt3C3yzTgBkswQfmCZDP/ALWB2NfN37Fn/BTP9oL9je/i0bS9VfxD4RL/AOkeFtVuGaOIEjLW7nJgbrwPkJPI71+gYHw+zPH5JHGQklUlqoPS8ejv0b3Selrao/LMy8VMnyviOWX1It0o6SqR1tPqrdUtm1re+jP12/4KR6bZat+wj8VrPUEDR/8ACFXr8/3lTcv6gV/NGCWhVj1Kgn8q/oY0j9pr4D/8FVv2YvFnwT+FfxFfw14i1/QpbW50vVI1+12ZYDLBA2Jo88FkPTsK/FP9s79gT9or9hnxl/wjfxi8Ll9MuJCmkeJ9OVnsL8f7DkfI+OsbYYe45r8Z4vy3H4LFqNek4uKs7rbX+rPZn9wfR34kyHG5TXoUMVGUqk1KEU90opNru+63VtUeHXADW0qnoYmB/I1/Tl+wPqF7qf7FfwsvNQkLyt4F00MzdSBAqj9AK/mPaCa6U2luMyTDy4x/tN8o/Uiv6kv2XvB8nw//AGbvAXgme0EEul+DtNtp4gfuyLaxhx/31urk4YT9tUfkj0fH+pBZbgYdXOb+Sir/AJo7uqfiDxDofhPRLrxL4m1a3sNPsYGmvLy7lCRwxqMlmY8AAUeIvEOieE9Cu/E3iTU4bKwsLd57y7uHCpDGoyzMT0AFfil/wVA/4KmeJP2wfEFx8J/hLe3OmfDbT7nBQ/JNrkqE/vpcHiEHlIz1+83YD9R4e4exfEGM9nT0gvil0S/Vvoj+N+K+K8DwrgPbVfeqS+CHWT/RLq/u1PQf+Cjf/BZ7xd8Vr7Ufgv8AsrapPo3hdd9vqPiiFil3qY6MICOYYj/eHzN6gV8BwSyTSSSyyM7O25mdiSxOckk9T71VTrVi0BO4Aelf0Pk2T4DJsKqGFhZdX1b7t9f06H8ocQZ9mfEONeJxs+Z9F9mK7RXRfi+pft/9SKvaWyqJCzADjkn617H+xt/wT9/aA/bN1Et8P9FGnaBbybL7xRqqMlpGc8omBmVx6LnHcivsHxV4K/4JNf8ABInT4Lv9pHxBF49+If2YTRaLJape3G4gkFLPPlwqTjDynPQiuDPONMmyBunUlz1F9iO69Xsvz8j1OHfDziDiiCqUo+zov7c9E/8ACt5fLTzPjz4F/sVftP8A7RQjvfhX8IdUvLB2GNWuY/s9pj1EsmAw/wB3NfVPwt/4IP8Axn1dIb/4rfFjRtCUgmS002B7uRT2G47Vr56/ae/4OZf2lPGs03h39lD4aaN4B0YLsttR1eFb/UdvTIQEQxcYxw2OeteBfsn/ALcX7ZHx9/bz+Fcnxh/ah8ca5Fd+OrFLmxl8QzQWcqmTO1raApCRx0KV+YZh4oZ/iZNYVRpR9OZ/e9P/ACVH7HlPgxwvg4p4yU68ut3yx+Sjr98mfrj4Q/4IW/s76XFG3i/4i+JtVlXmTyZI7dG/AAkfnXeaP/wR6/Ym0hVU+ENZucZybrX5m3fgMCvqSivmq3GHFFd3ni5/J8v5WPscP4fcFYZWhgKb/wAUeb/0q55T8DP2Kv2df2cvE1x4w+Engb+zdQubQ2005vJJCYiQxXDEjqBRXq1FeHisZisdW9riKjnLvJtv72fS4HL8DllBUMJSjThvaKUVd7uyPz0/4OEPgfP4t+APh342aXZh5fCmr+RqDogyLa4G0EnrgOBx71+PxHY1/TB8dfhD4a+Pfwh8Q/B7xdHnT/EOly2czhQTEWX5ZB7q2GH0r+cf40/CHxn8Bfiprvwg+IGmva6roGoyWtwrjiQKfklU91dNrg+jV+W8Y4GVLGRxKXuzVn6r/Nfkz++foxcXUcx4YrZDVl+9w0nKK705u+n+Gd79uaPcz/hn8QfEXwm+Imh/E/wjP5WqeH9Uhv7BwcfvInDAH2PKn2Jr+jz9m745eFP2kfgj4c+NPg2dWste02O4MYbJglxiSJvQq4ZcH0r+ac9T9a+wv+CU/wDwU31D9i7xnL8N/ifd3N18O9buA9yigyPpFycD7RGo5KEffQdcAjkc8/DObQy/EunVdoT/AAfR+nR/I9vx98N8TxrkdPHZdDmxWGvaK3nTfxRXeSa5or/Elqz9zqKxvAPxB8FfFHwjYePPh74ms9Y0fU4BNY6hYTCSOVD3BHQ9iDyDwQK2a/UU1JXWx/n1Up1KNR06iaknZp6NNbprowr52/4KXftqad+xf+z3deI9NuEbxVrm+x8K2pwT55X5pyD1WMEMfcqO9fRNfnl/wVs/4Jq/tRftT+PYvjN8LvGtrr9rpmnC2sPBlyRbPaIPmcwuTtkZ25O7BOAO1e7w5h8txGcUo4+ajSTu77O2yv0Te97K19T5bi/FZxhMgrSyym51mrK26vvJLdtLZK7vbSx+TGo6tqes6rca7rF/LdXl3O893c3Ehd5pHYszsx5JJJJJ9aejB1DDvWh8Qvht8QPhN4ouPBXxM8G6joWrWrET2Gp2xikX3GeGH+0CQfWsm2k2tsJ4PT61/T9KcJQTg04va23yP4zrQqQqONRNSW6e9/O5r+F/FHiLwbr9r4m8K69eaZqFnKJLW+sLlopYWHQqykEV+jf7Kv8AwVO+Fn7R/gNv2V/+CjXh3TNUsdVQWieJby1X7Ncg8L9qA/1MgOMTJjnB+WvzUqxbyb1KsckV5Od8PZXxDhXRxcL9pfaXo+3dPRnvcM8W55wjj44rL6ri003G7s7emz7SVmujPvD4kf8ABAHVfC37Vngbxn8Adcj1/wCFmpeJrW61e3vLlXn0u0V/NIDjieJgoCuOfmGc9a/XdQEQADAA4A7V+L3/AATr/wCCq3jn9lG6t/hj8VpbvxB4BkkCxxvKZLjRh0LQZzuj7mLp3XB4P6A/t3/8FB/hz8Ev2LJfjp8LfF1jq134ttTZeCZbabImnkUgygdR5SkswI4IAPWv57xvAGO4dzT6tShzRqySjJbPy8mtW0/O10f1/Lxsj4h5JTxeY1ffwkHzRfxa2u30leySkkr6X94+M/8AguB/wUWn8beJLn9j34M+JGXSNKm2+Nr6zk+W8uRyLMMOqR9XxwW4/hNfnNZf6n/gVQ3d5d6hdy3+oXUk888jSTzytlpHYksxJ6kkkk+pqay/1P8AwKv3jI8qw+T4KOGo9Fq+76t/1otD+UOJs6xWfY2eLrvVvRdIx6Jen4u76k6dfwr7e/4Jcf8ABK3Wf2qrmH40fGaC40/4f29x/otsMpNrjqfmVD1SEHguOScgdM15/wD8EvP2B9V/bU+Mq3fiazni8C+HJo5vEd4h2/aWzuSzRv7z4+YjouehIr91PDvh3Q/CWhWnhnw1pVvY6fYW6QWdnaxBI4Y1GFRVHAAFfLca8YTyuP1DBP8Aev4pfyp9F/ef4Lzat9l4dcA086mszzGN6EX7sX9trq/7q/F6bJ3r+DvBXhP4deFbPwV4G8PWmlaTp1uIbHT7GARxQoOgVR/PqTya/mN/4Ktyyzf8FJ/jbLNKzsPiFeoGdiSFG0BeewHQV/UM/wBw/Sv5d/8Agqt/ykk+N3/ZRb/+a1+ISlKcnKTu2f0coQpwUYqyWiS2R4BXsn/BO/8A5Ps+Ef8A2Pth/wCjK8br2T/gnf8A8n2fCP8A7H2w/wDRlIFuf1SUUUUGgUUUUAFfBf8AwWc/4Jx3X7Rfg4/tFfBzQRN408O2W3UrC2i/eavYpk4GPvSx8lR1IyPSvvSggMMEVyY7BUMww0qFVaP8H0aPpOEuKc14Mz6jm2XytUpvZ7Si/ijLupLR9t1qkz+XCRGSRkcEEMQQRyKhm4cEelfrt/wVC/4Iz23xQudR/aB/ZO0K3tPEMrPca74TgAji1NyctNAOFjmPJK9HPIwTz+THizwz4i8G6/deFvFuh3emanYTGG9sL63aKaBx/C6MAVP1r8lzHKsVldfkqrTo+j/rsf6ScEeIXD3iDlCxWXztUSXPSbXPB+a6x7SWj8ndL1P9kj9vj9pD9i7xD9v+EXjJzpU0m7UPDWpZmsLrtkxk/I/+2hB9cjiv0j+A/wDwcOfs9eK9OhtPj18P9Y8K6n8qyzacv220du7AjDqv+8K/HV+tRv2rqy/OswwEVGnO8ez1X/A+R4HGfhVwVxjVeIx2H5az/wCXlN8k3620l6yTfY/oGsf+CwP/AATuv7ZLmP8AaQ0uMP8AwzWsysv1BTiuQ+KH/BdP9gD4f2Rl0Tx7qPiefJC2+haVI3I6ZZwoA96/COVm2/ePX1qGUkgEnNez/rXmMo2UYr5P/M/LY/Ry4LoVeadevJduaC/FQufsl4P/AOCin/BNv/gqDE/wZ/aX+HieFtUluGj8Pz+IpY1k+bAVobyPiGQ9NhOD718xft5/8Eg/iv8AstW9z8TvhNcz+MPAyESPcQxbr3TkPIMyIPnQf89F4xgkCvgK5JUqynBz1r7a/wCCbn/BZT4mfsrXtr8Jvj3e3/i74cSr5KRXDedeaOCeTEzcyRcnMTE/7OOQftuDvE7NcirKnWlzU76p7fd9l+a+aZ+LeMn0UMi4gws8ZkScasV8O8tF9mT+L/BP/t2SdkfMkUglQMOvenxuY3DCv0i/bb/4Jj/Cr9oD4bj9s/8A4J6Xdpqdjq1ub688N6OQ0F6uSXktVH+rlBzuh9QQADkH837i3ntJ3tbqB4pY3KyRSoVZGBwQQeQQeMGv6xyHiDL+IcGsRhZeq6r1/R7M/wAvuKeFM44RzOWCx8Gmm7Ozs7Oz32a6p6rqWVOQGHejxBqGv6x4ctvDc2tXcljYTyT2WnvcMYYZJABIyIThWYAZI64GahtpP+WZP0qavclCFWNpI+ap1KlGfNF2OTIIOCOR1rY8G+Gtd8Za5Y+EfC+nPealqd7Ha2FrGMmWZ2Cov5kfQc1B4gsRBOLqNflk+97GvuT/AIIJfsxx/FX9oq/+OniLTPN0vwJbg2DSJlW1GYEJ26om5vbNfN5tj4ZLgquJqa8i0830XzZ9hkmW1OIsfQwlLT2kkm+yWsn8ldn6hfsTfsveGP2Q/wBnTQPg5oECm6t7cXGuXnG67v5ADPISOo3fKvoqKK9ZoHAxRX80YjEVsVXlWqu8pNtvzZ/YGEwtDA4WGHoxtCCSS7JaIR/uH6V/Lv8A8FVv+Uknxu/7KLf/AM1r+oh/uH6V/Lv/AMFVv+Uknxu/7KLf/wA1rE2keAV7J/wTv/5Ps+Ef/Y+2H/oyvG69k/4J3/8AJ9nwj/7H2w/9GUErc/qkooooNAooooAKKKKAAjNeD/te/wDBOb9mT9s3T3n+Jvg5bXXli2WvinSQIb6L0DMBiVR/dcH8K94orKtQo4mm6dWKkn0Z6GV5tmeSY2OMy+tKlVjtKLaf4dH1T0fU/FX9pL/ggV+1N8MrqfU/gjqlh470sMTDBG4tb1U7Bkc7WP8AutXx38Rf2ffjp8J74ab8Svg94l0SYk7V1DRZkDAHBIO0gj3r+m7APUVHdWdrewtbXlsksbjDRyoGU/UHivmMRwjgqkr0ZOHluv8AP8T9/wAi+krxZgaSpZnh4YlL7WtOb9Wk4/dFH8sF0v2dzHcERt/dkO0/kan0zwx4k8QTx2mg+HdQvpZWCxpZ2MkpY+nyqa/pwvvgP8ENTmNzqXwd8K3EhOS8/h21ck/Ux1r+HfA3gvwipTwp4R0vTARgjT9PigyP+AKK5IcITUta2n+H/gn0mJ+kzhp0n7LLHzedVW/Cmfz2fBf/AIJT/t3fHx4X8LfAbU9MspXx/aXiPFjCB/e/efMy+6qa/QL9kP8A4N4vhH4BvLTxl+1X4rPjC9iKyDw3YBodPDcHbK335h2K8Ke+a/SbA9KK9nCcO4DDNSleb89vu/zufl3Evjfxln9OVKhKOGg/+fd+Zr/G7tf9u8pneFfCPhfwN4etfCfg3w/Z6XpllEIrSw0+2WGGFB0VVUAAV+b3/BZ7/gm/pKaXqP7YvwT0RbaeE+d450m2TCTKSB9uRR0cEjzAPvD5uoOf0yqtrOj6Z4g0m50PWrCK6tLyB4bm2nTcksbAhlYdwQSK+1yLOcVkOYQxNB6LRrpKPVP9Oz1P5/4nyDB8U5ZPC4rVu7jLdxl0l/n3V0fzMglTkH6VajcSKGFe2/8ABRj9ke6/Y7/aU1TwHYROfDup51HwvO2T/ojsf3JPdo2yh9gD3rwu2fa+0ng1/UOBxlDH4WGIou8ZpNfP9e5/GGZZficsx1TCYhWnTbTXmv0e6fVC6nbrc2Mkbdl3DHtzX7jf8Ed/2ez8Av2G/C7apYeTq/ixG17VNwG7/SOYVyOSBCIyAem8ivxs+A/wp1T45fGfwv8ACDRiwn8Ra1BZ+YoBMcbNmR8HrtQM2Pav6K/Deg6d4W8PWHhjR7dYbTTrOK1tYkGAkcaBFUewAFfmHinj1ChQwkXrJuT9Fovxb+4/Z/BTLJVMTicdNaQSjF+ctZfckvvLtFFFfix/Qwj/AHD9K/l3/wCCq3/KST43f9lFv/5rX9RD/cP0r+Xf/gqt/wApJPjd/wBlFv8A+a0EyPAK9k/4J3/8n2fCP/sfbD/0ZXjdeyf8E7/+T7PhH/2Pth/6MoJW5/VJRRRQaBRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQB8yf8ABTX9ga6/bo+HOiaV4V1zT9J8RaBqRmsdS1CJmQwSLtliO3nBwrfVRXxT/wAQ9/7Rf/Ra/CP/AID3H+FfrjRX0+V8YZ9lGEWGw1RKCu0nFO19XufGZzwBwxn2PljMXSbqNJNqTV7aLRPe2h8E/wDBPD/gkR4y/ZQ+Pi/Gj4q+OdF1s2Glyw6RBpkMgMdxIQDI28dkyBjnJNfe1FFeXm2b4/O8V9YxcuaVktrKy8ke3keQ5Zw7gvquBhywu3q222+7eoUUUV5h7AjDKkD0r8Pf27/+CFH/AAUC+PX7ZnxN+NPw78L+GJ9B8UeL7rUdIluPEqRSNBJgrvQp8p4PFfuHRgelAmrn88f/ABDm/wDBTT/oTPCX/hWR/wDxFej/ALIX/BBX/goX8Hf2pPAHxW8beFPDEWkeHvFVpf6lJb+JkkdYY3yxVQnzHHav3VwPQUYHpQFkFFFFAwooooAKKKKACvCf2vP+Ckv7IP7C2u6N4b/aY+Jj6Bea/ZTXelRLpc9x58UTKsjfukbGCyjB9a92r8av+Dl7wkvj/wDbD/Z68ASag1omvafd6c10ibjD5+oWkW/b327s474oE9Efr/4F8ceFPiX4N0v4geBtbg1LR9ZsYrzTb+1kDRzwyKGVgR7H8DxXlv7XH/BQT9k79huDSZf2lvinDoD640n9mWy2stxNMqD5n2RKzBR03EYzxX5m/wDBOn9vzxd/wSR+KHjr/gnt+3neXMGgeF0u7/wfqioZAoXdIsUOeWhuV+aPsrkqcV8mftif8NIft9/DL4hf8FXvi7M2l+GrTxPaeG/BmjHJRbdpGAgi7bYlIZ3/AI5Xb0FArn9FPwh+LHgf46fDPRPi98NdVa+0HxDp6Xuk3jQtGZoX+621gGXPoRmukrwD/glgSf8Agnb8HST/AMyJZf8AoJr3+go4/wDaA+NnhT9nH4LeJfjp45tb2fSPC2ky6hqEOnQiSd4oxkhFJAZvQZFfBZ/4OiP+CfSgeZ4H+JSEqDh/DsIPP/bevqD/AIKu/wDKOP4yf9iJefyFfm5/wR2/bP8A+CU/wN/YvsfAf7XereC4PGKa9eTzJrnhT7ZP9ncp5R8zyXyMA4GeKBM+yv2R/wDgvZ+xt+2X+0FoX7N3wu8M+N7bXfEKXLWM2raNFFbjyIWlfeyysR8qnHB5r7dHIzXyR+yd+2R/wSW+O3xns/BH7Kd/4GuvGa2s1xZLo/hEWtwkSriRlkMK7eDg4POa+t6AQEhQWPYV8MftO/8ABf79jb9k/wCO3iX9n34leDvHk2s+F71bbULjTNEjltnZo1kBRzKMjDjsOc19zSf6tv8AdNfjR8JfBnhDx/8A8HOHxA8MeOvCunazp0trqUklhqtklxCzrp9oVYo4IJGTg470Az6h+D//AAcd/wDBOH4q+J4PDOp+IfEvhQ3MojjvvEuieXbgnGCzxu+we5GBX3V4c8R6B4v0Gz8U+FtYttQ03ULZLixvrOYSRTxMMq6sOCCO9fN37ZX/AASk/Y5/aq+Dmp+D5/gj4c0HXUsZToHiPQdIhs7qxuNpKfPGo3RlgAyNlSM8Z5r5a/4Nofjz441X4YfEL9k3xvqkt2nw610Po/nOWNtBM7pJCueiCWNmUZ43GgNT9Q6CcUUjdPxH86Bnzx+xx/wUr+Bf7bXxP8c/Cb4WaB4js9T+H90bfW31rT1hikYTPDmJldt43IeeOMV9EV+Sf/Bvd/ye9+09/wBhx/8A043FfrZQJaoKr6vqun6DpVzrerXaQWtnbvPczyNhY40UszE9gACasV8J/wDBwP8AtjN+zF+xFd+AfDGsNb+J/iTcNoumrCwEiWm3ddyjuAIyEz6yCgZ2f7HP/BaH9kT9tv48XX7PXwph8SWmtQ2txcWk+uaYkFvfLC+1/JcOSxx8wGBlea+ua/nz+L/7IPxG/wCCSXw8/Zf/AG9vDVjeLrk0puPHtu0jAJdSss8VqR0QNaGSEjpvQk9hX72/Cr4j+GvjB8NNA+Kng6+S50rxFpFvqNhNG2Q0U0YdefocfhQJO5v0UUUDCiiigAr8gv8Ag4a/5SDfst/9fn/uWsq/X2vD/wBp/wD4J1/sq/th/E3wd8Xvjx4M1LUtd8Byb/DVzZeI72ySA+dHP88cEqJL88SH5w3QjoTQJ6o4H/gpb/wSf+DP/BSLTfD194p1q48OeItAv0VPEWmwq08unNJme1YHrkZKMfuNyOprwv8A4Ly/Bv4d/s+/8Ec7f4M/Cjw9Fpfh/wAO+INDstMs4+dsaSEbmbq7scsznlmYk8mv0iAA4Fea/tYfsk/A/wDbX+EU3wN/aE8P3up+HJ7+C8ktbDWLixkM0Lboz5tu6OACeRnB70BY4r/glf8A8o7Pg7/2Ill/6Ca9/rmfg38IvA3wF+FuhfBv4aadNaaB4b02Ox0m2ubyS4kjgQYVWklZnc+7EmumoGfPf/BV3/lHH8ZP+xEvP5Cvym/4JY23/BEs/snWcn7d8HhVvHza1d+cdZa6877L8nk/6s7cYziv22+M/wAIPAnx++FevfBj4nabNeeH/EmnSWOrW1veSW7yQP8AeVZImV0PupBr5AX/AINzP+CUqKFX4NeJwAMD/i5Wtf8AyVQJmf8AsQ6l/wAELdH/AGjdIj/Ypk8JQ/EO9trm30pNJN350kXllpQBJ8uAqknPpX3tXyl+zh/wRX/4J9fsofGXSPj58FPhnrth4m0JZxpt3e+N9UvI4/OiMT5inuGjbKsQMqcdRzX1b0oBCSf6tv8AdNfil4a+NXws/Z+/4OUfiH8TPjL41svD2g28WoW82qahJtiSWTT7XYpPqdpxX7XEBgQe9fJPx+/4Igf8E6/2mvi/rvx1+L/ws12+8SeI7lZ9WurXxzqlrHLIqLGCIobhUT5VXhQOmaAZxv7Zf/BeH9iX4H/BrVtS+D/xX0/xr4unsXj0HSNHy6LMwKrLNIRtSNCdxzyQMAc1wH/BuL+yn8R/hj8EvFv7UXxa0q4sNT+KWppc6baXcBjlexjLMLhlbkCWR2ZQcfKAe4r3H4O/8EOP+CYnwN1+DxV4V/ZotdRvrW4We0l8U6xeassMg6MqXcrqCPp1r6zgghtYUtraFY441CxxooAUAYAAHQUBr1H0jdPxH86WggHrQM/JP/g3u/5Pe/ae/wCw4/8A6cbiv1srxX9mT/gn1+y7+yB8QPF/xO+BPg3UNN1nx1dG48R3F54gvLxZ3MrSnYk8rrEN7scIAOcdBXtVAlohHYIpYkDHc1+AP/BRj4j/ALQn/BSb/gqPqtv+zD8NX8e6f8JXjs9B0TyDLaulrOGuZ5gHTfHLcjaRuBKoFziv321nSrXXNJutFvWlWG7t3hlaCZo3CspU7WUgqcHgg5HavF/2Pv8AgnP+yd+wnd+INR/Zw8B3umXfieSN9ZvNT8QXmozTbCxVQ91LIyLlmOFIBJyaAep+XH7V/ij/AIL5ftl/A+/+AHxv/Ya8KP4evZIZt2i+F5re6tZIWDxvDI9/IqEYx905BI717/8A8G2H7V+s+LPgv4n/AGLviTfTLr/w21J5dHtbziWPTZZGDwYPP7m4DrjqBJ6AV+nJGa8F+HH/AATS/ZE+Ef7Uur/tj/DnwNqek+OddkuH1W6tfE18LO4M4Hm5s/N8jDEbsBMbssOTmgLWZ71RRRQM/9k=`;

  const vdRow   = vdPct > 0 ? '<tr class="sum"><td colspan="5" style="text-align:right;font-weight:700">Volume Discount ('+Math.round(vdPct*100)+'%)</td><td style="text-align:right;color:#b00">- &#8358; '+fmt(vdAmt)+'</td></tr><tr class="sum"><td colspan="5" style="text-align:right;font-weight:700">Less Discount</td><td style="text-align:right;font-weight:700">&#8358; '+fmt(afterDisc)+'</td></tr>' : "";
  const cRow    = cPct > 0  ? '<tr class="sum"><td colspan="5" style="text-align:right;font-weight:700">Agency Commission ('+Math.round(cPct*100)+'%)</td><td style="text-align:right;color:#b00">- &#8358; '+fmt(cAmt)+'</td></tr><tr class="sum"><td colspan="5" style="text-align:right;font-weight:700">Less Commission</td><td style="text-align:right;font-weight:700">&#8358; '+fmt(afterComm)+'</td></tr>' : "";
  const spRow   = spPct > 0 ? '<tr class="sum"><td colspan="5" style="text-align:right;font-weight:700">'+(surchLabel||("Surcharge ("+Math.round(spPct*100)+"%)"))+' </td><td style="text-align:right;color:#b25400">+ &#8358; '+fmt(spAmt)+'</td></tr><tr class="sum"><td colspan="5" style="text-align:right;font-weight:700">Net After Surcharge</td><td style="text-align:right;font-weight:700">&#8358; '+fmt(netAmt)+'</td></tr>' : "";
  const costBodyRows = costLines.map(l => '<tr><td>'+l.programme+'</td><td>'+l.material+'</td><td style="text-align:center">'+l.duration+'secs</td><td style="text-align:center;font-weight:700">'+l.cnt+'</td><td style="text-align:right">'+fmt(l.rate)+'</td><td style="text-align:right;font-weight:700">'+fmt(l.gross)+'</td></tr>').join("");
  const termsRows = terms.map((t,i) => '<tr><td class="n">'+(i+1)+'</td><td class="t">'+t+'</td></tr>').join("");

  return `<!DOCTYPE html><html><head><meta charset="utf-8"><title>MPO ${mpoNo}</title>
  <style>
    *{box-sizing:border-box;margin:0;padding:0}
    body{font-family:Arial,Helvetica,sans-serif;font-size:9px;color:#111;background:#fff;padding:8mm 10mm}
    @media print{body{padding:5mm 7mm}@page{size:A4;margin:15mm}}
    .logo-wrap{text-align:center;margin:0;padding:0;line-height:1;border:none;background:none}
    .logo-wrap img{max-height:60px;max-width:120px;object-fit:contain;display:block;border:none;outline:none;margin:0 auto;padding:0;background:transparent}
    .agency-addr{text-align:center;font-size:7.5px;color:#7b0000;font-weight:700;text-transform:uppercase;margin:0;padding:0;letter-spacing:.3px;white-space:nowrap;line-height:1.2}
    .header-wrap{text-align:center;margin-bottom:4px;border:none;background:none}
    .header-logo-col{display:flex;flex-direction:column;align-items:center;gap:8px;margin:0;padding:0;border:none;background:none}
    .sub-header-wrap{display:flex;align-items:flex-start;gap:12mm;margin:4mm 0 6px 0;width:100%}
    .rec-box{border:0.5pt solid #000;border-radius:3mm;background:#fff;padding:2.5mm;box-sizing:border-box}
    .rec-box-left{width:48mm;min-height:19mm;flex-shrink:0}
    .rec-box-right{width:62mm;min-height:26mm;flex-shrink:0;margin-left:auto}
    .rec-line{line-height:1.2;color:#000;font-size:6pt;font-family:Arial,Helvetica,sans-serif}
    .rec-line-bold{font-weight:700;font-size:6.5pt;text-transform:uppercase;font-family:Arial,Helvetica,sans-serif;color:#000}
    .det-line{line-height:1.15;color:#000;font-size:5.5pt;font-family:Arial,Helvetica,sans-serif;margin-bottom:0.5mm}
    .det-lbl{font-weight:700;text-transform:uppercase;font-family:Arial,Helvetica,sans-serif}
    .det-val{font-weight:400;font-family:Arial,Helvetica,sans-serif}
    .tx-bar{background:#1a3a6b;color:#fff;text-align:center;font-size:9px;font-weight:700;padding:5px;margin-bottom:8px;letter-spacing:.8px}
    .cal-wrap{overflow-x:auto;margin-bottom:8px}
    .cal{border-collapse:collapse;font-size:8px;width:100%;table-layout:auto}
    .cal th,.cal td{border:1px solid #aaa;padding:2px 3px;white-space:nowrap}
    .cal thead th{background:#1a3a6b;color:#fff;font-size:7.5px;text-align:center}
    .costing-title{text-align:center;font-weight:800;font-size:10px;letter-spacing:6px;margin:12px 0 5px;color:#1a3a6b;text-transform:uppercase}
    .cost{width:100%;border-collapse:collapse;font-size:8.5px;margin-bottom:8px}
    .cost th{background:#1a3a6b;color:#fff;padding:5px 8px;text-align:left;font-size:8px}
    .cost td{padding:4px 8px;border:1px solid #ddd}
    .cost .sum td{border:none;padding:3px 8px}
    .cost .payable td{border:2px solid #888;padding:5px 8px;font-weight:700;font-size:10.5px;background:#f5f8fd}
    .terms-title{font-weight:700;text-decoration:underline;font-size:9px;margin:10px 0 3px;font-style:italic}
    .terms{width:100%;border-collapse:collapse}
    .terms td{padding:3px 5px;font-size:8px;line-height:1.55;vertical-align:top;border:1px solid #e0e0e0}
    .terms .n{width:16px;font-weight:700;text-align:right;border-right:none;white-space:nowrap;color:#222}
    .terms .t{border-left:none}
    .sig{width:100%;border-collapse:collapse;margin-top:24px}
    .sig td{width:33.3%;vertical-align:bottom;padding:0 8px 0 0;font-size:8px}
    .sig .dots{font-size:9px;color:#444;margin-bottom:3px;letter-spacing:1px}
    .sig .role{font-weight:700;font-size:8.5px;margin-bottom:1px}
    .sig .sub{color:#555;font-size:7.5px}
  </style></head><body>

  <div class="header-wrap">
    <table style="margin:0 auto;border-collapse:collapse;border:none">
      <tr><td style="text-align:center;padding:0;border:none">
        <img src="${LOGO_SRC}" alt="QVT Media" style="max-height:60px;max-width:120px;display:block;margin:0 auto;border:none;outline:none">
      </td></tr>
      <tr><td style="text-align:center;padding-top:8px;border:none">
        <span style="font-size:7.5px;color:#7b0000;font-weight:700;text-transform:uppercase;letter-spacing:.3px;white-space:nowrap">${(agencyAddress||"5, CRAIG STREET, OGUDU GRA, LAGOS").replace(/\n/g," | ")} &nbsp;|&nbsp; TEL: ${agencyPhone||preparedContact||"+234 800 000 0000"}${agencyEmail ? ` &nbsp;|&nbsp; EMAIL: ${agencyEmail}` : ""}</span>
      </td></tr>
    </table>
  </div>

  <div class="sub-header-wrap">
    <!-- LEFT BOX: Recipient -->
    <div class="rec-box rec-box-left">
      <div class="rec-line rec-line-bold">THE COMMERCIAL MANAGER,</div>
      <div class="rec-line">${(vendorName||"—").toUpperCase()} MEDIA SALES,</div>
      <div class="rec-line">LAGOS.</div>
    </div>
    <!-- RIGHT BOX: Client Details -->
    <div class="rec-box rec-box-right">
      <div class="det-line"><span class="det-lbl">CLIENT NAME: </span><span class="det-val">${(clientName||"—").toUpperCase()}</span></div>
      <div class="det-line"><span class="det-lbl">BRAND: </span><span class="det-val">${(brand||"—").toUpperCase()}</span></div>
      <div class="det-line"><span class="det-lbl">MEDIA PURCHASE ORDER No: </span><span class="det-val">${mpoNo||"—"}</span></div>
      <div class="det-line"><span class="det-lbl">MEDIUM: </span><span class="det-val">${(mpo.medium||"RADIO").toUpperCase()}</span></div>
      <div class="det-line"><span class="det-lbl">CAMPAIGN TITLE: </span><span class="det-val">${(campaignName||"—").toUpperCase()}</span></div>
      <div class="det-line"><span class="det-lbl">PERIOD: </span><span class="det-val">${periodLabel}</span></div>
      <div class="det-line"><span class="det-lbl">DATE: </span><span class="det-val">${date||""}</span></div>
    </div>
  </div>

  <div class="tx-bar">PLEASE TRANSMIT ${firstDur} SPOTS ON ${(vendorName||"").toUpperCase()} AS SCHEDULED</div>

  ${allCalendarHTML}

  <div class="costing-title">C &nbsp; O &nbsp; S &nbsp; T &nbsp; I &nbsp; N &nbsp; G</div>
  <table class="cost">
    <thead><tr>
      <th style="text-align:left;min-width:100px">PROGRAMME</th>
      <th style="text-align:left;min-width:110px">MATERIAL</th>
      <th style="text-align:center;min-width:55px">DURATION</th>
      <th style="text-align:center;min-width:60px">NO OF SPOTS</th>
      <th style="text-align:right;min-width:90px">RATE/SPOT (&#8358;)</th>
      <th style="text-align:right;min-width:110px">TOTAL AMOUNT (&#8358;)</th>
    </tr></thead>
    <tbody>${costBodyRows}</tbody>
    <tfoot>
      <tr class="sum"><td colspan="5" style="text-align:right;font-weight:700;border-top:2px solid #888">Sub Total</td><td style="text-align:right;font-weight:700;border-top:2px solid #888">&#8358; ${fmt(subTotal)}</td></tr>
      ${vdRow}${cRow}${spRow}
      <tr class="sum"><td colspan="5" style="text-align:right;font-weight:700">VAT (${parseFloat(vatPct) || 7.5}%)</td><td style="text-align:right;font-weight:700">&#8358; ${fmt(vatAmt)}</td></tr>
      <tr class="payable"><td colspan="5" style="text-align:right;letter-spacing:.5px">Total Amount Payable &#8596;</td><td style="text-align:right;color:#1a3a6b;font-size:11px">&#8358; ${fmt(totalPayable)}</td></tr>
    </tfoot>
  </table>

  <div class="terms-title">Contract Terms &amp; Condition</div>
  <table class="terms">${termsRows}</table>

  <table class="sig"><tr>
    <td><div class="dots">&#8230;&#8230;&#8230;&#8230;&#8230;&#8230;&#8230;&#8230;&#8230;&#8230;&#8230;&#8230;&#8230;</div><div class="role">NAME/SIGNATURE/DATE &amp; OFFICIAL STAMP</div><div class="sub">For (Media House / Third party supplier)</div></td>
    <td>${signedSignature ? `<div style="height:42px;display:flex;align-items:flex-end;margin:0 0 2px;"><img src="${signedSignature}" alt="Signed signature" style="max-height:40px;max-width:160px;object-fit:contain"></div>` : `<div style="height:42px;"></div>`}<div class="dots" style="margin-bottom:4px;">&#8230;&#8230;&#8230;&#8230;&#8230;&#8230;&#8230;&#8230;&#8230;&#8230;&#8230;&#8230;</div><div class="role">SIGNED BY: ${(signedBy||"").toUpperCase()}</div><div>${preparedContact||""}</div><div class="sub">${(signedTitle||"").toUpperCase()}</div></td>
    <td>${preparedSignature ? `<div style="height:42px;display:flex;align-items:flex-end;margin:0 0 2px;"><img src="${preparedSignature}" alt="Prepared signature" style="max-height:40px;max-width:160px;object-fit:contain"></div>` : `<div style="height:42px;"></div>`}<div class="dots" style="margin-bottom:4px;">&#8230;&#8230;&#8230;&#8230;&#8230;&#8230;&#8230;&#8230;&#8230;&#8230;&#8230;&#8230;</div><div class="role">PREPARED BY: ${(preparedBy||"").toUpperCase()}</div><div>${preparedContact||""}</div><div class="sub">${(preparedTitle||"").toUpperCase()}</div></td>
  </tr></table>

  <script>
    window.addEventListener('message', function(e) {
      if (e.data === 'print-mpo') { window.focus(); window.print(); }
    });
  </script>
  </body></html>`;
};


/* ── DAILY CALENDAR COMPONENT ──────────────────────────────────────── */
const DailyCalendar = ({ month, year, calRows, setCalRows, vendorRates, fmtN, blankCalRow, onAdd, campaignMaterials }) => {
  const MONTH_NAMES = ["JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"];
  const DAY_LABELS  = ["Sun","Mon","Tue","Wed","Thu","Fri","Sat"];

  const mIdx      = (() => {
    const m = (month||"").toUpperCase();
    // Try 3-letter abbreviation first (JAN, FEB…), then full name (JANUARY…)
    const short = MONTH_NAMES.indexOf(m.slice(0,3));
    return short >= 0 ? short : MONTH_NAMES.findIndex(n => m.startsWith(n));
  })();
  const yr        = parseInt(year) || new Date().getFullYear();
  const validMonth = mIdx >= 0;
  const dim       = validMonth ? new Date(yr, mIdx + 1, 0).getDate() : 31;
  const firstDOW  = validMonth ? new Date(yr, mIdx, 1).getDay() : 0;

  // Build week grid: array of 7-slot rows, null = empty cell
  const cells = [];
  for (let i = 0; i < firstDOW; i++) cells.push(null);
  for (let d = 1; d <= dim; d++) cells.push(d);
  while (cells.length % 7 !== 0) cells.push(null);
  const weeks = [];
  for (let i = 0; i < cells.length; i += 7) weeks.push(cells.slice(i, i + 7));

  const updRow = (id, key, val) =>
    setCalRows(rows => rows.map(r => r.id === id ? { ...r, [key]: val } : r));

  const toggleDate = (rowId, d) =>
    setCalRows(rows => rows.map(r => {
      if (r.id !== rowId) return r;
      const next = new Set(r.selectedDates);
      next.has(d) ? next.delete(d) : next.add(d);
      return { ...r, selectedDates: next };
    }));

  const quickSelect = (rowId, wdNums) =>
    setCalRows(rows => rows.map(r => {
      if (r.id !== rowId) return r;
      const next = new Set();
      if (wdNums === "all") {
        for (let d = 1; d <= dim; d++) next.add(d);
      } else if (wdNums === "clear") {
        /* empty */
      } else {
        for (let d = 1; d <= dim; d++)
          if (validMonth && wdNums.includes(new Date(yr, mIdx, d).getDay())) next.add(d);
      }
      return { ...r, selectedDates: next };
    }));

  const inp = (extra = {}) => ({
    background: "var(--bg2)", border: "1px solid var(--border2)", borderRadius: 6,
    padding: "5px 8px", color: "var(--text)", fontSize: 11, outline: "none",
    width: "100%", ...extra
  });

  if (!validMonth) return (
    <div style={{ background: "rgba(240,165,0,.1)", border: "1px solid rgba(240,165,0,.3)",
        borderRadius: 8, padding: "14px 16px", fontSize: 12, color: "var(--accent)", marginBottom: 14 }}>
      ⚠️ Please set <strong>Month &amp; Year</strong> in Step 1 before using the Daily Calendar.
    </div>
  );

  return (
    <div className="fade" style={{ marginBottom: 12 }}>
      {/* Month header */}
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 14 }}>
        <div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 800, fontSize: 15, letterSpacing: 1 }}>
          📅 {MONTH_NAMES[mIdx]} {yr}
        </div>
        <div style={{ fontSize: 11, color: "var(--text3)" }}>
          Click dates on the calendar to toggle airing days per spot row
        </div>
      </div>

      {calRows.map((row, ri) => {
        const rateObj = vendorRates.find(r => r.id === row.rateId);
        const rateVal = parseFloat(row.customRate) || parseFloat(rateObj?.ratePerSpot) || 0;
        const spotCount = row.selectedDates.size;
        const gross = rateVal * spotCount;

        return (
          <div key={row.id} style={{ border: "1px solid var(--border2)", borderRadius: 10,
              marginBottom: 16, background: "var(--bg2)", overflow: "hidden" }}>

            {/* Row header bar */}
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between",
                background: "var(--bg3)", padding: "10px 14px", borderBottom: "1px solid var(--border)" }}>
              <span style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 12,
                  color: "var(--text2)", textTransform: "uppercase", letterSpacing: 1 }}>
                Spot Row {ri + 1}
              </span>
              <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                {spotCount > 0 && (
                  <span style={{ fontSize: 11, background: "rgba(240,165,0,.18)", color: "var(--accent)",
                      padding: "3px 10px", borderRadius: 12, fontWeight: 700 }}>
                    {spotCount} spot{spotCount !== 1 ? "s" : ""}{rateVal > 0 ? ` · ₦${fmtN(gross)}` : ""}
                  </span>
                )}
                {calRows.length > 1 && (
                  <button onClick={() => setCalRows(r => r.filter(x => x.id !== row.id))}
                    style={{ background: "none", border: "none", cursor: "pointer",
                        color: "var(--text3)", fontSize: 18, lineHeight: 1, padding: "0 2px" }}>×</button>
                )}
              </div>
            </div>

            <div style={{ padding: "12px 14px" }}>
              {/* Programme selector from vendor rate cards */}
              {vendorRates.length > 0 && (
                <div style={{ marginBottom: 12 }}>
                  <div style={{ fontSize: 10, fontWeight: 700, color: "var(--text3)", textTransform: "uppercase", marginBottom: 5, letterSpacing: .5 }}>
                    Select Programme from Rate Card
                  </div>
                  <select
                    value={row.rateId}
                    onChange={e => {
                      const picked = vendorRates.find(r => r.id === e.target.value);
                      if (picked) {
                        setCalRows(rs => rs.map(r => r.id !== row.id ? r : {
                          ...r,
                          rateId:      picked.id,
                          programme:   picked.programme || r.programme,
                          timeBelt:    picked.timeBelt  || r.timeBelt,
                          duration:    picked.duration  || r.duration,
                          customRate:  "",
                        }));
                      } else {
                        updRow(row.id, "rateId", "");
                      }
                    }}
                    style={{ ...inp(), cursor: "pointer", background: row.rateId ? "rgba(240,165,0,.08)" : "var(--bg2)", border: row.rateId ? "1px solid rgba(240,165,0,.4)" : "1px solid var(--border2)" }}>
                    <option value="">— Select programme / rate card —</option>
                    {vendorRates.map(r => (
                      <option key={r.id} value={r.id}>
                        {r.programme}{r.timeBelt ? ` · ${r.timeBelt}` : ""} · {r.duration}s · ₦{fmtN(r.ratePerSpot)}
                      </option>
                    ))}
                  </select>
                  {row.rateId && (
                    <div style={{ marginTop: 4, display: "flex", alignItems: "center", justifyContent: "space-between" }}>
                      <span style={{ fontSize: 10, color: "var(--green)" }}>✓ Rate card applied — fields auto-filled below</span>
                      <button onClick={() => setCalRows(rs => rs.map(r => r.id !== row.id ? r : { ...r, rateId: "", programme: "", timeBelt: "", duration: "30", customRate: "" }))}
                        style={{ fontSize: 10, color: "var(--text3)", background: "none", border: "none", cursor: "pointer", textDecoration: "underline" }}>Clear</button>
                    </div>
                  )}
                </div>
              )}

              {/* Spot detail fields */}
              <div style={{ display: "grid", gridTemplateColumns: "2fr 1.4fr 1.4fr 64px 1.5fr",
                  gap: 8, marginBottom: 12 }}>
                <div>
                  <div style={{ fontSize: 10, fontWeight: 700, color: "var(--text3)",
                      textTransform: "uppercase", marginBottom: 4, letterSpacing: .5 }}>Programme</div>
                  <input value={row.programme} onChange={e => updRow(row.id, "programme", e.target.value)}
                    placeholder="e.g. NTA Network News" style={inp()} />
                </div>
                <div>
                  <div style={{ fontSize: 10, fontWeight: 700, color: "var(--text3)",
                      textTransform: "uppercase", marginBottom: 4, letterSpacing: .5 }}>Time Belt</div>
                  <input value={row.timeBelt} onChange={e => updRow(row.id, "timeBelt", e.target.value)}
                    placeholder="e.g. 9PM–10PM" style={inp()} />
                </div>
                <div>
                  <div style={{ fontSize: 10, fontWeight: 700, color: "var(--text3)",
                      textTransform: "uppercase", marginBottom: 4, letterSpacing: .5 }}>Material</div>
                  {campaignMaterials && campaignMaterials.length > 0 ? (
                    <>
                      <select value={row.material} onChange={e => updRow(row.id, "material", e.target.value)}
                        style={{ ...inp(), cursor: "pointer" }}>
                        <option value="">— Select —</option>
                        {campaignMaterials.map((m, mi) => <option key={mi} value={m}>{m}</option>)}
                        <option value="__custom__">Custom…</option>
                      </select>
                      {row.material === "__custom__" && (
                        <input value={row.materialCustom || ""} onChange={e => updRow(row.id, "materialCustom", e.target.value)}
                          placeholder="Type material" style={{ ...inp(), marginTop: 4 }} />
                      )}
                    </>
                  ) : (
                    <input value={row.material} onChange={e => updRow(row.id, "material", e.target.value)}
                      placeholder="e.g. 30s TVC" style={inp()} />
                  )}
                </div>
                <div>
                  <div style={{ fontSize: 10, fontWeight: 700, color: "var(--text3)",
                      textTransform: "uppercase", marginBottom: 4, letterSpacing: .5 }}>Secs</div>
                  <input type="number" value={row.duration}
                    onChange={e => updRow(row.id, "duration", e.target.value)}
                    placeholder="30" style={inp({ width: 60 })} />
                </div>
                <div>
                  <div style={{ fontSize: 10, fontWeight: 700, color: "var(--text3)",
                      textTransform: "uppercase", marginBottom: 4, letterSpacing: .5 }}>Rate / Spot (₦)</div>
                  {(vendorRates.length === 0 || !row.rateId) ? (
                    <input type="number" value={row.customRate}
                      onChange={e => { updRow(row.id, "customRate", e.target.value); updRow(row.id, "rateId", ""); }}
                      placeholder="Enter rate" style={inp()} />
                  ) : (
                    <div style={{ ...inp(), background: "rgba(240,165,0,.08)", border: "1px solid rgba(240,165,0,.3)", color: "var(--accent)", fontWeight: 700 }}>
                      ₦{fmtN(vendorRates.find(r => r.id === row.rateId)?.ratePerSpot || 0)}
                    </div>
                  )}
                </div>
              </div>

              {/* Quick-select chips */}
              <div style={{ display: "flex", gap: 5, flexWrap: "wrap", marginBottom: 12 }}>
                {[
                  { label: "All",        act: () => quickSelect(row.id, "all") },
                  { label: "Weekdays",   act: () => quickSelect(row.id, [1,2,3,4,5]) },
                  { label: "M·W·F",      act: () => quickSelect(row.id, [1,3,5]) },
                  { label: "T·Th",       act: () => quickSelect(row.id, [2,4]) },
                  { label: "Weekends",   act: () => quickSelect(row.id, [0,6]) },
                  { label: "Mon only",   act: () => quickSelect(row.id, [1]) },
                  { label: "Clear",      act: () => quickSelect(row.id, "clear"), danger: true },
                ].map(({ label, act, danger }) => (
                  <button key={label} onClick={act} style={{
                    padding: "3px 10px", borderRadius: 20, fontSize: 10, fontWeight: 600,
                    cursor: "pointer", border: "1px solid var(--border2)",
                    background: danger ? "rgba(220,50,50,.1)" : "var(--bg3)",
                    color: danger ? "#e05555" : "var(--text2)", transition: "all .1s"
                  }}>{label}</button>
                ))}
              </div>

              {/* ── CALENDAR GRID ── */}
              <div style={{ userSelect: "none" }}>
                {/* Day-of-week headers */}
                <div style={{ display: "grid", gridTemplateColumns: "repeat(7, 1fr)",
                    gap: 3, marginBottom: 4 }}>
                  {DAY_LABELS.map((d, i) => (
                    <div key={d} style={{
                      textAlign: "center", fontSize: 10, fontWeight: 700, letterSpacing: .5,
                      color: i === 0 || i === 6 ? "var(--accent)" : "var(--text3)",
                      padding: "4px 0"
                    }}>{d}</div>
                  ))}
                </div>

                {/* Weeks */}
                {weeks.map((wk, wi) => (
                  <div key={wi} style={{ display: "grid", gridTemplateColumns: "repeat(7, 1fr)", gap: 3, marginBottom: 3 }}>
                    {wk.map((d, di) => {
                      if (!d) return <div key={`e${di}`} />;
                      const isActive = row.selectedDates.has(d);
                      const isWE = di === 0 || di === 6;
                      const dotw = validMonth ? new Date(yr, mIdx, d).getDay() : di;
                      const isWkend = dotw === 0 || dotw === 6;
                      return (
                        <div key={d} onClick={() => toggleDate(row.id, d)}
                          style={{
                            aspectRatio: "1", display: "flex", alignItems: "center",
                            justifyContent: "center", borderRadius: 8, cursor: "pointer",
                            border: isActive ? "2px solid var(--accent)" : "1px solid var(--border2)",
                            background: isActive ? "var(--accent)" : isWkend ? "rgba(240,165,0,.05)" : "var(--bg3)",
                            color: isActive ? "#000" : isWkend ? "var(--accent)" : "var(--text)",
                            fontWeight: isActive ? 800 : isWkend ? 600 : 400,
                            fontSize: 12, transition: "all .1s",
                            boxShadow: isActive ? "0 2px 8px rgba(240,165,0,.35)" : "none"
                          }}>
                          {d}
                        </div>
                      );
                    })}
                  </div>
                ))}
              </div>

            </div>
          </div>
        );
      })}

      {/* Footer actions */}
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginTop: 8 }}>
        <button onClick={() => setCalRows(r => [...r, blankCalRow()])}
          style={{ background: "none", border: "1px dashed var(--border2)", borderRadius: 8,
              padding: "7px 16px", cursor: "pointer", fontSize: 12, color: "var(--text2)",
              display: "flex", alignItems: "center", gap: 6 }}>
          <span style={{ fontSize: 18, lineHeight: 1 }}>+</span> Add Spot Row
        </button>
        <button onClick={onAdd}
          style={{ background: "var(--accent)", color: "#000", border: "none", borderRadius: 8,
              padding: "9px 22px", cursor: "pointer", fontWeight: 700, fontSize: 13,
              fontFamily: "'Syne',sans-serif" }}>
          ✓ Add to Schedule
        </button>
      </div>
    </div>
  );
};

/* ── MPO GENERATOR ──────────────────────────────────────── */
const MPOPage = ({ vendors, clients, campaigns, rates, mpos, setMpos, user, appSettings }) => {
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
  const [workflowPanelOpen, setWorkflowPanelOpen] = useState(true);

  const VAT_RATE = parseFloat(appSettings?.vatRate) || 7.5;
  const [surcharge, setSurcharge] = useState({ pct: "", label: "" });

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
  // calRows: array of { id, programme, timeBelt, material, duration, rateId, customRate, selectedDates: Set<number> }
  const [dailyMode, setDailyMode] = useState(false);
  const blankCalRow = () => ({ id: uid(), programme: "", timeBelt: "", material: "", duration: "30", rateId: "", customRate: "", selectedDates: new Set() });
  // Multi-month: calData maps "Month Year" -> calRows array
  const [calData, setCalData] = useState({});
  const [activeCalMonth, setActiveCalMonth] = useState("");
  // Legacy single-month calRows still used for single-month mode
  const [calRows, setCalRows] = useState([blankCalRow()]);

  const blankSpot = { programme: "", wd: "", timeBelt: "", material: "", duration: "30", rateId: "", customRate: "", spots: "", calendarDays: [] };
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
    setDailyMode(false); setCalRows([blankCalRow()]); setCalData({}); setActiveCalMonth("");
    clearSavedDraft();
  };

  // Add spot rows from calendar schedule — supports single or multi-month
  const addFromCalendarSchedule = (rowsToAdd, monthLabel) => {
    const sourceRows = rowsToAdd || calRows;
    const newSpots = [];
    sourceRows.forEach(row => {
      if (!row.programme || row.selectedDates.size === 0) return;
      const rate = vendorRates.find(r => r.id === row.rateId);
      const ratePerSpot = parseFloat(row.customRate) || parseFloat(rate?.ratePerSpot) || 0;
      const calendarDays = Array.from(row.selectedDates).sort((a,b)=>a-b);
      const matFinal = row.material === "__custom__" ? (row.materialCustom || "") : (row.material || "");
      newSpots.push({ id: uid(), programme: row.programme, wd: "", timeBelt: row.timeBelt,
        material: matFinal, duration: row.duration, rateId: row.rateId,
        ratePerSpot, spots: calendarDays.length, calendarDays,
        scheduleMonth: monthLabel || mpoData.month });
    });
    if (!newSpots.length) return setToast({ msg: "Fill Programme and select at least one date per row.", type: "error" });
    setSpots(s => [...s, ...newSpots]);
    if (!rowsToAdd) { setDailyMode(false); setCalRows([blankCalRow()]); }
    setToast({ msg: `${newSpots.length} spot row(s) added from calendar.`, type: "success" });
  };

  const addSpot = () => {
    if (!spotForm.programme) return;
    const rate = vendorRates.find(r => r.id === spotForm.rateId);
    const ratePerSpot = parseFloat(spotForm.customRate) || parseFloat(rate?.ratePerSpot) || 0;
    const calDays = spotForm.calendarDays || [];
    const spotsCount = calDays.length > 0 ? calDays.length : (parseFloat(spotForm.spots) || 0);
    if (!spotsCount) return setToast({ msg: "Select at least one airing date or enter a spot count.", type: "error" });
    const newSpot = { id: uid(), ...spotForm, ratePerSpot, spots: String(spotsCount), calendarDays: calDays };
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
    const pdfBytes = buildMPOPdf(safeMpo);
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
    setPreview({ html, csv, pdfBytes, title: `MPO — ${mpo.mpoNo || "Draft"} | ${mpo.vendorName} | ${mpo.month} ${mpo.year}` });
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
      {preview && <PrintPreview html={preview.html} csv={preview.csv} pdfBytes={preview.pdfBytes} title={preview.title} onClose={() => setPreview(null)} />}
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
            <Card key={m.id} style={{ padding: "14px 18px" }}>
              <div style={{ display: "flex", alignItems: "center", gap: 14, flexWrap: "wrap" }}>
                <div style={{ width: 44, height: 44, background: "rgba(240,165,0,.1)", border: "1px solid rgba(240,165,0,.2)", borderRadius: 11, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 20, flexShrink: 0 }}>📄</div>
                <div style={{ flex: 1, minWidth: 160 }}>
                  <div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 15 }}>{m.mpoNo || "MPO"} <span style={{ color: "var(--text3)", fontSize: 13 }}>— {m.vendorName}</span></div>
                  <div style={{ fontSize: 12, color: "var(--text2)", marginTop: 2 }}>{m.clientName} {m.brand && `· ${m.brand}`} · {(m.months||[]).length > 1 ? (m.months||[]).join("-") : m.month} {m.year} · {m.totalSpots} spots</div>
                  <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginTop: 6 }}>
                    <Badge color={statusColors[m.status || "draft"] || "accent"}>{MPO_STATUS_LABELS[m.status || "draft"] || (m.status || "draft")}</Badge>
                    <Badge color={getMpoWorkflowMeta(m).color}>Waiting on: {getMpoWorkflowMeta(m).label}</Badge>
                    {isMpoAwaitingUser(user, m) ? <Badge color="blue">My Queue</Badge> : null}
                    <Badge color={getExecutionHealthColor(m)}>{getExecutionHealthLabel(m)}</Badge>
                    <Badge color="blue">Invoice: {(MPO_INVOICE_STATUS_OPTIONS.find(o => o.value === (m.invoiceStatus || "pending"))?.label || m.invoiceStatus || "pending")}</Badge>
                    <Badge color="purple">Proof: {(MPO_PROOF_STATUS_OPTIONS.find(o => o.value === (m.proofStatus || "pending"))?.label || m.proofStatus || "pending")}</Badge>
                    <Badge color="green">Payment: {(MPO_PAYMENT_STATUS_OPTIONS.find(o => o.value === (m.paymentStatus || "unpaid"))?.label || m.paymentStatus || "unpaid")}</Badge>
                  </div>
                  <div style={{ fontSize: 11, color: "var(--text3)", marginTop: 7 }}>{getMpoWorkflowMeta(m).hint}</div>{isArchived(m) && <div style={{ marginTop: 4 }}><Badge color="red">Archived</Badge></div>}
                </div>
                <div style={{ display: "flex", gap: 14, flexWrap: "wrap", alignItems: "center" }}>
                  <div style={{ textAlign: "center" }}><div style={{ fontSize: 10, color: "var(--text3)", fontWeight: 600, textTransform: "uppercase" }}>Gross</div><div style={{ fontSize: 13, fontWeight: 600, marginTop: 2, color: "var(--text2)" }}>{fmtN(m.totalGross)}</div></div>
                  <div style={{ textAlign: "center" }}><div style={{ fontSize: 10, color: "var(--text3)", fontWeight: 600, textTransform: "uppercase" }}>Net Payable</div><div style={{ fontSize: 14, fontWeight: 700, marginTop: 2, color: "var(--accent)" }}>{fmtN(m.grandTotal || m.netVal)}</div></div>
                  {canManageStatus ? <Field value={m.status || "draft"} onChange={v => requestMpoStatusChange(m, v)} options={[{ value: m.status || "draft", label: MPO_STATUS_LABELS[m.status || "draft"] || (m.status || "draft") }, ...getAllowedMpoStatusTargets(user, m).map(value => ({ value, label: MPO_STATUS_LABELS[value] || value }))].filter((option, index, arr) => arr.findIndex(item => item.value === option.value) === index)} /> : <div style={{ minWidth: 110 }}><Badge color={statusColors[m.status || "draft"] || "accent"}>{(m.status || "draft").toUpperCase()}</Badge></div>}
                </div>
                <div style={{ display: "flex", gap: 6, flexShrink: 0, flexWrap: "wrap" }}>
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
                <Field label="Media House / Vendor" value={mpoData.vendorId} onChange={upd("vendorId")} options={vendors.map(v => ({ value: v.id, label: v.name }))} placeholder="Select vendor" />
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
              const getRows = m => calData[m] || [blankCalRow()];
              const setRows = (m, fn) => setCalData(d => ({ ...d, [m]: typeof fn === "function" ? fn(d[m] || [blankCalRow()]) : fn }));
              return (
                <div>
                  {selectedMonths.length > 1 && (
                    <div style={{ display: "flex", gap: 0, marginBottom: 16, background: "var(--bg3)", border: "1px solid var(--border2)", borderRadius: 10, overflow: "hidden" }}>
                      {selectedMonths.map(m => (
                        <button key={m} onClick={() => setActiveCalMonth(m)}
                          style={{ flex: 1, padding: "9px 8px", border: "none", background: curCalMonth === m ? "var(--accent)" : "transparent", color: curCalMonth === m ? "#000" : "var(--text2)", fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 12, cursor: "pointer", borderRight: "1px solid var(--border)", transition: "all .15s" }}>
                          {m.slice(0,3)} <span style={{ fontSize: 10, opacity: .7 }}>{mpoData.year?.slice(-2)}</span>
                          {(calData[m] || []).some(r => r.selectedDates?.size > 0) && (
                            <span style={{ marginLeft: 4, background: "rgba(34,197,94,.25)", color: "var(--green)", borderRadius: 10, padding: "1px 5px", fontSize: 9 }}>✓</span>
                          )}
                        </button>
                      ))}
                    </div>
                  )}
                  <DailyCalendar
                    month={curCalMonth} year={mpoData.year}
                    calRows={getRows(curCalMonth)}
                    setCalRows={fn => setRows(curCalMonth, fn)}
                    vendorRates={vendorRates} fmtN={fmtN}
                    blankCalRow={blankCalRow}
                    campaignMaterials={campaign?.materialList || []}
                    onAdd={() => {
                      const rows = getRows(curCalMonth);
                      addFromCalendarSchedule(rows, `${curCalMonth} ${mpoData.year}`);
                      setRows(curCalMonth, [blankCalRow()]);
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
                  <thead><tr style={{ background: "var(--bg3)" }}>{["Month","Programme","WD","Time Belt","Material","Dur","Rate/Spot","Spots","Gross",""].map(h => <th key={h} style={{ padding: "8px 11px", textAlign: "left", fontSize: 10, fontWeight: 600, textTransform: "uppercase", color: "var(--text3)", borderBottom: "1px solid var(--border)", whiteSpace: "nowrap" }}>{h}</th>)}</tr></thead>
                  <tbody>
                    {spots.map((r, i) => (
                      <tr key={r.id} style={{ borderBottom: "1px solid var(--border)", background: i % 2 === 0 ? "transparent" : "rgba(255,255,255,.01)" }}>
                        <td style={{ padding: "9px 11px" }}><span style={{ fontSize: 10, background: "rgba(59,126,245,.12)", color: "var(--blue)", padding: "2px 7px", borderRadius: 8, fontWeight: 700, whiteSpace: "nowrap" }}>{r.scheduleMonth || mpoData.month || "—"}</span></td>
                        <td style={{ padding: "9px 11px", fontWeight: 600 }}>{r.programme}</td>
                        <td style={{ padding: "9px 11px" }}><Badge color="blue">{r.wd}</Badge></td>
                        <td style={{ padding: "9px 11px", color: "var(--text2)" }}>{r.timeBelt}</td>
                        <td style={{ padding: "9px 11px", color: "var(--text2)", maxWidth: 140, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{r.material}</td>
                        <td style={{ padding: "9px 11px" }}>{r.duration}"</td>
                        <td style={{ padding: "9px 11px", color: "var(--accent)" }}>{fmtN(r.ratePerSpot)}</td>
                        <td style={{ padding: "9px 11px", fontWeight: 700 }}>{r.spots}</td>
                        <td style={{ padding: "9px 11px", color: "var(--green)", fontWeight: 600 }}>{fmtN((parseFloat(r.spots) || 0) * (parseFloat(r.ratePerSpot) || 0))}</td>
                        <td style={{ padding: "9px 11px" }}>
                          <div style={{ display: "flex", gap: 5 }}>
                            {canManage && <Btn variant="ghost" size="sm" onClick={() => { setSpotForm({ programme: r.programme, wd: r.wd, timeBelt: r.timeBelt, material: r.material, duration: r.duration, rateId: r.rateId||"", customRate: r.ratePerSpot||"", spots: r.spots, calendarDays: r.calendarDays||[] }); setEditSpotId(r.id); setSpotModal(true); }}>✏️</Btn>}
                            <Btn variant="danger" size="sm" onClick={() => setSpots(s => s.filter(x => x.id !== r.id))}>×</Btn>
                          </div>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                  <tfoot><tr style={{ background: "var(--bg3)", borderTop: "2px solid var(--border2)" }}><td colSpan={7} style={{ padding: "9px 11px", fontFamily: "'Syne',sans-serif", fontWeight: 700 }}>TOTAL</td><td style={{ padding: "9px 11px", fontWeight: 700, color: "var(--blue)" }}>{totalSpots}</td><td style={{ padding: "9px 11px", fontWeight: 700, color: "var(--green)" }}>{fmtN(totalGross)}</td><td></td></tr></tfoot>
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
};

/* ── SETTINGS ───────────────────────────────────────────── */
const SettingsPage = ({ user, onUserUpdate, onLogout, appSettings, setAppSettings, vendors, clients, campaigns, rates, mpos, members, setMembers, notifications, unreadNotifications, onMarkNotificationRead, onMarkAllNotificationsRead }) => {
  const [toast, setToast] = useState(null);
  const backupInputRef = useRef(null);
  const [confirm, setConfirm] = useState(null);
  const [section, setSection] = useState("profile"); // profile | security | agency | team | activity | data
  const canManageWorkspace = hasPermission(user, "manageWorkspace");
  const canManageMembers = hasPermission(user, "manageMembers");
  const canDanger = hasPermission(user, "manageDangerZone");
  const [memberRoles, setMemberRoles] = useState({});
  const [savingMemberId, setSavingMemberId] = useState(null);
  const [activityItems, setActivityItems] = useState([]);
  const [activityLoading, setActivityLoading] = useState(false);
  const [activityFilter, setActivityFilter] = useState("all");
  const [notificationFilter, setNotificationFilter] = useState("all");

  // Profile form
  const [pf, setPf] = useState({ name: user.name || "", title: user.title || "", email: user.email || "", phone: user.phone || "", signatureDataUrl: user.signatureDataUrl || "" });
  const up = k => v => setPf(p => ({ ...p, [k]: v }));

  // Agency form
  const [af, setAf] = useState({ agency: user.agency || "", address: user.agencyAddress || "", email: user.agencyEmail || "", phone: user.agencyPhone || "" });
  const [docSettings, setDocSettings] = useState({
    vatRate: String(appSettings?.vatRate ?? 7.5),
    sessionHours: String(appSettings?.sessionHours ?? DEFAULT_SESSION_HOURS),
    roundToWholeNaira: !!appSettings?.roundToWholeNaira,
    mpoTerms: (appSettings?.mpoTerms || DEFAULT_APP_SETTINGS.mpoTerms).join("\n"),
  });
  const udp = k => v => setDocSettings(s => ({ ...s, [k]: v }));
  const ua = k => v => setAf(a => ({ ...a, [k]: v }));

  // Password form
  const [sf, setSf] = useState({ current: "", newPw: "", confirm: "" });
  const us = k => v => setSf(s => ({ ...s, [k]: v }));
  const signatureInputRef = useRef(null);

  const handleSignatureUpload = async (file) => {
    if (!file) return;
    try {
      const dataUrl = await fileToDataUrl(file);
      setPf(p => ({ ...p, signatureDataUrl: dataUrl }));
      onUserUpdate({ ...user, signatureDataUrl: dataUrl });
      await persistSignatureForUser(user, dataUrl);
      setToast({ msg: "Signature uploaded and saved.", type: "success" });
    } catch (error) {
      setToast({ msg: error.message || "Failed to upload signature.", type: "error" });
    }
  };

  const handleSignatureRemove = async () => {
    try {
      setPf(p => ({ ...p, signatureDataUrl: "" }));
      onUserUpdate({ ...user, signatureDataUrl: "" });
      await persistSignatureForUser(user, "");
      setToast({ msg: "Signature removed.", type: "success" });
    } catch (error) {
      setToast({ msg: error.message || "Failed to remove signature.", type: "error" });
    }
  };

  useEffect(() => {
    setPf({ name: user.name || "", title: user.title || "", email: user.email || "", phone: user.phone || "", signatureDataUrl: user.signatureDataUrl || "" });
    setAf({ agency: user.agency || "", address: user.agencyAddress || "", email: user.agencyEmail || "", phone: user.agencyPhone || "" });
  }, [user?.id, user?.name, user?.title, user?.email, user?.phone, user?.agency, user?.agencyAddress, user?.signatureDataUrl, user?.agencyEmail, user?.agencyPhone]);

  useEffect(() => {
    setDocSettings({
      vatRate: String(appSettings?.vatRate ?? 7.5),
      sessionHours: String(appSettings?.sessionHours ?? DEFAULT_SESSION_HOURS),
      roundToWholeNaira: !!appSettings?.roundToWholeNaira,
      mpoTerms: (appSettings?.mpoTerms || DEFAULT_APP_SETTINGS.mpoTerms).join("\n"),
    });
  }, [appSettings?.vatRate, appSettings?.sessionHours, appSettings?.roundToWholeNaira, JSON.stringify(appSettings?.mpoTerms || [])]);

  useEffect(() => {
    setMemberRoles(Object.fromEntries((members || []).map(member => [member.id, normalizeRole(member.role)])));
  }, [JSON.stringify((members || []).map(member => ({ id: member.id, role: member.role })))]);

  useEffect(() => {
    if (section !== "activity" || !user?.agencyId) return;
    let active = true;
    setActivityLoading(true);
    fetchAuditEventsForAgency(user.agencyId, activityFilter)
      .then(rows => { if (active) setActivityItems(rows); })
      .catch(error => { if (active) setToast({ msg: error.message || "Failed to load workspace activity.", type: "error" }); })
      .finally(() => { if (active) setActivityLoading(false); });
    return () => { active = false; };
  }, [section, user?.agencyId, activityFilter]);

  const saveMemberRole = async (memberId) => {
    if (!canManageMembers) return setToast({ msg: readOnlyMessage(user), type: "error" });
    const nextRole = normalizeRole(memberRoles[memberId]);
    const targetMember = (members || []).find(member => member.id === memberId);
    try {
      setSavingMemberId(memberId);
      const savedRole = await updateAgencyMemberRoleInSupabase(memberId, nextRole);
      setMembers(prev => prev.map(member => member.id === memberId ? { ...member, role: savedRole, roleLabel: formatRoleLabel(savedRole) } : member));
      if (memberId === user.id) onUserUpdate({ ...user, role: savedRole, roleLabel: formatRoleLabel(savedRole) });
      createAuditEventInSupabase({
        agencyId: user.agencyId,
        recordType: "member",
        recordId: memberId,
        action: "role_changed",
        actor: user,
        note: `${targetMember?.name || targetMember?.email || "Member"} is now ${formatRoleLabel(savedRole)}.`,
        metadata: { email: targetMember?.email || "", role: savedRole },
      }).catch(error => console.error("Failed to write audit event:", error));
      setToast({ msg: "Member role updated.", type: "success" });
    } catch (error) {
      setToast({ msg: error.message || "Failed to update member role.", type: "error" });
    } finally {
      setSavingMemberId(null);
    }
  };

  const saveProfile = async () => {
    if (!pf.name || !pf.email) return setToast({ msg: "Name and email are required.", type: "error" });
    try {
      const updated = await updateProfileInSupabase(user, pf);
      onUserUpdate(updated);
      createAuditEventInSupabase({
        agencyId: user.agencyId,
        recordType: "profile",
        recordId: user.id,
        action: "updated",
        actor: user,
        metadata: { email: updated.email || "", title: updated.title || "" },
      }).catch(error => console.error("Failed to write audit event:", error));
      setToast({ msg: pf.email !== user.email ? "Profile updated. Check your email if confirmation is required for the new address." : "Profile updated successfully!", type: "success" });
    } catch (error) {
      setToast({ msg: error.message || "Failed to update profile.", type: "error" });
    }
  };

  const saveAgency = async () => {
    if (!canManageWorkspace) return setToast({ msg: readOnlyMessage(user), type: "error" });
    if (!af.agency) return setToast({ msg: "Agency name is required.", type: "error" });
    try {
      const updated = await updateAgencyInSupabase(user.agencyId, af);
      onUserUpdate(updated);
      createAuditEventInSupabase({
        agencyId: user.agencyId,
        recordType: "agency",
        recordId: user.agencyId,
        action: "updated",
        actor: user,
        metadata: { name: updated.agency || af.agency || "", address: updated.agencyAddress || af.address || "", email: updated.agencyEmail || af.email || "", phone: updated.agencyPhone || af.phone || "" },
      }).catch(error => console.error("Failed to write audit event:", error));
      setToast({ msg: "Agency details updated!", type: "success" });
    } catch (error) {
      setToast({ msg: error.message || "Failed to update agency details.", type: "error" });
    }
  };

  const exportBackup = () => {
    const payload = {
      version: APP_VERSION,
      exportedAt: new Date().toISOString(),
      user: { ...user },
      data: { vendors, clients, campaigns, rates, mpos, notifications, appSettings, localPreferences: { theme: store.get(themeKeyForUser(user?.id || null), getDefaultTheme(user?.id || null)), signatureCached: !!user?.signatureDataUrl } },
    };
    downloadJSON(`mediadesk-backup-${new Date().toISOString().slice(0,10)}.json`, payload);
    setToast({ msg: "Agency backup exported successfully.", type: "success" });
  };

  const handleBackupImport = async (file) => {
    if (!file) return;
    setToast({ msg: "Snapshot import is disabled for cloud workspaces to protect your live Supabase data. Use CSV tools for structured imports instead.", type: "error" });
  };

  const saveDocumentSettings = async () => {
    if (!canManageWorkspace) return setToast({ msg: readOnlyMessage(user), type: "error" });
    const vatRate = parseFloat(docSettings.vatRate);
    const sessionHours = parseFloat(docSettings.sessionHours);
    if (!Number.isFinite(vatRate) || vatRate < 0 || vatRate > 100) return setToast({ msg: "Enter a valid VAT rate between 0 and 100.", type: "error" });
    if (!Number.isFinite(sessionHours) || sessionHours < 1 || sessionHours > 48) return setToast({ msg: "Session hours must be between 1 and 48.", type: "error" });
    try {
      const nextSettings = {
        ...appSettings,
        vatRate,
        sessionHours,
        roundToWholeNaira: !!docSettings.roundToWholeNaira,
        mpoTerms: docSettings.mpoTerms.split("\n").map(t => t.trim()).filter(Boolean),
      };
      setAppSettings(nextSettings);
      createAuditEventInSupabase({
        agencyId: user.agencyId,
        recordType: "workspace",
        recordId: user.agencyId,
        action: "settings_updated",
        actor: user,
        metadata: { vatRate: nextSettings.vatRate, sessionHours: nextSettings.sessionHours, roundToWholeNaira: nextSettings.roundToWholeNaira },
      }).catch(error => console.error("Failed to write audit event:", error));
      setToast({ msg: "Document and billing settings updated.", type: "success" });
    } catch (error) {
      setToast({ msg: error.message || "Failed to update workspace settings.", type: "error" });
    }
  };

  const changePassword = async () => {
    if (!sf.current || !sf.newPw || !sf.confirm) return setToast({ msg: "All password fields required.", type: "error" });
    if (sf.newPw.length < 6) return setToast({ msg: "New password must be at least 6 characters.", type: "error" });
    if (sf.newPw !== sf.confirm) return setToast({ msg: "New passwords do not match.", type: "error" });
    try {
      await changePasswordInSupabase(sf.newPw);
      setSf({ current: "", newPw: "", confirm: "" });
      createAuditEventInSupabase({
        agencyId: user.agencyId,
        recordType: "security",
        recordId: user.id,
        action: "password_changed",
        actor: user,
      }).catch(error => console.error("Failed to write audit event:", error));
      setToast({ msg: "Password changed successfully!", type: "success" });
    } catch (error) {
      setToast({ msg: error.message || "Failed to change password.", type: "error" });
    }
  };

  const tabs = [
    { id: "profile",  icon: "👤", label: "My Profile" },
    ...(canManageWorkspace ? [{ id: "agency", icon: "🏢", label: "Agency" }] : []),
    { id: "security", icon: "🔒", label: "Security" },
    { id: "notifications", icon: "🔔", label: `Notifications${unreadNotifications ? ` (${unreadNotifications})` : ""}` },
    { id: "activity", icon: "🕘", label: "Activity" },
    ...(canManageMembers ? [{ id: "team", icon: "🛡️", label: "Team & Roles" }] : []),
    ...(canDanger ? [{ id: "data", icon: "🗄️", label: "Data Management" }] : []),
  ];

  useEffect(() => {
    if (!tabs.find(tab => tab.id === section)) setSection("profile");
  }, [section, canManageWorkspace, canManageMembers, canDanger]);

  const InfoRow = ({ label, value, accent }) => (
    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "10px 0", borderBottom: "1px solid var(--border)", fontSize: 13 }}>
      <span style={{ color: "var(--text3)", fontWeight: 600 }}>{label}</span>
      <span style={{ fontWeight: 600, color: accent || "var(--text)" }}>{value || "—"}</span>
    </div>
  );

  return (
    <div className="fade">
      {toast && <Toast msg={toast.msg} type={toast.type} onDone={() => setToast(null)} />}
      {confirm && <Confirm msg={confirm.msg} onYes={confirm.onYes} onNo={() => setConfirm(null)} danger />}

      {/* Header */}
      <div style={{ marginBottom: 28 }}>
        <h1 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 800, fontSize: 24 }}>⚙️ System Settings</h1>
        <p style={{ color: "var(--text2)", marginTop: 3, fontSize: 13 }}>Manage your account, agency details, and platform preferences</p>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "200px 1fr", gap: 20, alignItems: "start" }}>

        {/* Tab sidebar */}
        <div style={{ background: "var(--bg2)", border: "1px solid var(--border)", borderRadius: 14, overflow: "hidden" }}>
          {tabs.map(t => (
            <button key={t.id} onClick={() => setSection(t.id)}
              style={{ display: "flex", alignItems: "center", gap: 10, width: "100%", padding: "12px 16px", border: "none", borderBottom: "1px solid var(--border)", background: section === t.id ? "rgba(139,92,246,.12)" : "transparent", color: section === t.id ? "var(--purple)" : "var(--text2)", fontFamily: "'Syne',sans-serif", fontWeight: section === t.id ? 700 : 400, fontSize: 13, cursor: "pointer", textAlign: "left", transition: "all .14s", borderLeft: section === t.id ? "3px solid var(--purple)" : "3px solid transparent" }}
              onMouseEnter={e => { if (section !== t.id) e.currentTarget.style.background = "var(--bg3)"; }}
              onMouseLeave={e => { if (section !== t.id) e.currentTarget.style.background = "transparent"; }}>
              <span style={{ fontSize: 16 }}>{t.icon}</span>{t.label}
            </button>
          ))}
          {/* Sign out button at bottom of tab panel */}
          <button onClick={() => setConfirm({ msg: "Sign out of MediaDesk Pro?", onYes: onLogout })}
            style={{ display: "flex", alignItems: "center", gap: 10, width: "100%", padding: "12px 16px", border: "none", background: "transparent", color: "var(--red)", fontFamily: "'Syne',sans-serif", fontWeight: 400, fontSize: 13, cursor: "pointer", textAlign: "left", transition: "all .14s" }}
            onMouseEnter={e => e.currentTarget.style.background = "rgba(239,68,68,.07)"}
            onMouseLeave={e => e.currentTarget.style.background = "transparent"}>
            <span style={{ fontSize: 16 }}>⎋</span> Sign Out
          </button>
        </div>

        {/* Panel content */}
        <div>

          {/* ── PROFILE ── */}
          {section === "profile" && (
            <div className="fade">
              {/* Avatar card */}
              <Card style={{ marginBottom: 16, padding: "20px 24px" }}>
                <div style={{ display: "flex", alignItems: "center", gap: 18 }}>
                  <div style={{ width: 64, height: 64, background: "linear-gradient(135deg,var(--accent),var(--purple))", borderRadius: "50%", display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "'Syne',sans-serif", fontWeight: 800, fontSize: 26, color: "#000", flexShrink: 0 }}>{user.name?.[0]?.toUpperCase() || "U"}</div>
                  <div>
                    <div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 800, fontSize: 20 }}>{user.name}</div>
                    <div style={{ color: "var(--text2)", fontSize: 13, marginTop: 2 }}>{user.title || "No title set"} · {user.agency}</div>
                    <div style={{ color: "var(--text3)", fontSize: 12, marginTop: 2 }}>{user.email}</div>
                  </div>
                </div>
              </Card>

              <Card>
                <h2 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 16, marginBottom: 18 }}>Edit Profile</h2>
                <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 }}>
                    <Field label="Full Name" value={pf.name} onChange={up("name")} placeholder="Jane Okafor" required />
                    <Field label="Job Title" value={pf.title} onChange={up("title")} placeholder="Media Buyer" />
                  </div>
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 }}>
                    <Field label="Email Address" type="email" value={pf.email} onChange={up("email")} placeholder="you@agency.com" required note="Used for login" />
                    <Field label="Phone Number" type="tel" value={pf.phone} onChange={up("phone")} placeholder="+234 800 000 0000" note="Used as Prepared By contact on MPOs" />
                  </div>
                  <div style={{ display: "flex", flexDirection: "column", gap: 10, padding: "12px 14px", borderRadius: 10, background: "var(--bg3)", border: "1px solid var(--border)" }}>
                    <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, flexWrap: "wrap" }}>
                      <div>
                        <div style={{ fontSize: 11, fontWeight: 700, color: "var(--text3)", textTransform: "uppercase", letterSpacing: ".08em" }}>Signature</div>
                        <div style={{ fontSize: 12, color: "var(--text2)", marginTop: 4 }}>Upload your signature image to show on MPO signatories.</div>
                      </div>
                      <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                        <Btn variant="ghost" size="sm" onClick={() => signatureInputRef.current?.click()}>Upload Signature</Btn>
                        {pf.signatureDataUrl ? <Btn variant="danger" size="sm" onClick={handleSignatureRemove}>Remove</Btn> : null}
                      </div>
                    </div>
                    <input ref={signatureInputRef} type="file" accept="image/*" style={{ display: "none" }} onChange={e => handleSignatureUpload(e.target.files?.[0])} />
                    {pf.signatureDataUrl ? (
                      <img src={pf.signatureDataUrl} alt="Signature preview" style={{ maxHeight: 62, maxWidth: 220, objectFit: "contain", background: "#fff", borderRadius: 8, border: "1px solid var(--border)" }} />
                    ) : (
                      <div style={{ fontSize: 12, color: "var(--text3)" }}>No signature uploaded yet.</div>
                    )}
                  </div>
                  <div style={{ background: "rgba(240,165,0,.07)", border: "1px solid rgba(240,165,0,.18)", borderRadius: 9, padding: "10px 14px", fontSize: 12, color: "var(--text2)" }}>
                    💡 Your <strong style={{ color: "var(--accent)" }}>name, title, and signature</strong> are automatically used in the "Prepared By" field on every MPO you generate.
                  </div>
                  <div style={{ display: "flex", justifyContent: "flex-end" }}>
                    <Btn onClick={saveProfile} icon="✓">Save Profile</Btn>
                  </div>
                </div>
              </Card>
            </div>
          )}

          {/* ── AGENCY ── */}
          {section === "agency" && (
            <div className="fade">
              <Card style={{ marginBottom: 16 }}>
                <div style={{ display: "flex", gap: 14, alignItems: "center", marginBottom: 10 }}>
                  <div style={{ width: 48, height: 48, background: "rgba(59,126,245,.12)", border: "1px solid rgba(59,126,245,.25)", borderRadius: 12, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 22, flexShrink: 0 }}>🏢</div>
                  <div><div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 17 }}>{user.agency}</div><div style={{ fontSize: 12, color: "var(--text2)", marginTop: 2 }}>{user.agencyAddress || "No address set"}</div><div style={{ display: "flex", gap: 12, flexWrap: "wrap", marginTop: 6, fontSize: 12, color: "var(--text2)" }}><span>✉️ {user.agencyEmail || "No agency email"}</span><span>📞 {user.agencyPhone || "No agency phone"}</span></div></div>
                </div>
                <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, background: "var(--bg3)", border: "1px solid var(--border)", borderRadius: 10, padding: "12px 14px" }}>
                  <div>
                    <div style={{ fontSize: 11, color: "var(--text3)", textTransform: "uppercase", letterSpacing: ".08em", fontWeight: 700 }}>Agency Invite Code</div>
                    <div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 800, fontSize: 18, marginTop: 4 }}>{user.agencyCode || "Not available yet"}</div>
                  </div>
                  {user.agencyCode && <Btn size="sm" onClick={() => navigator.clipboard?.writeText(user.agencyCode)} icon="⧉">Copy Code</Btn>}
                </div>
              </Card>
              <Card style={{ marginBottom: 16 }}>
                <h2 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 16, marginBottom: 18 }}>Agency Details</h2>
                <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
                  <Field label="Agency / Company Name" value={af.agency} onChange={ua("agency")} placeholder="Apex Media Ltd" required />
                  <Field label="Agency Address" value={af.address} onChange={ua("address")} placeholder="5, Craig Street, Ogudu GRA, Lagos" note="Used as the footer address on all MPO documents" />
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 }}>
                    <Field label="Agency Email" type="email" value={af.email} onChange={ua("email")} placeholder="hello@agency.com" />
                    <Field label="Agency Phone" type="tel" value={af.phone} onChange={ua("phone")} placeholder="+234 800 000 0000" />
                  </div>
                  <div style={{ display: "flex", justifyContent: "flex-end" }}>
                    {canManageWorkspace && <Btn onClick={saveAgency} icon="✓">Save Agency Details</Btn>}
                  </div>
                </div>
              </Card>
              <Card>
                <h2 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 16, marginBottom: 18 }}>Document & Billing Defaults</h2>
                <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 }}>
                    <Field label="VAT Rate (%)" type="number" value={docSettings.vatRate} onChange={udp("vatRate")} placeholder="7.5" note="Applied to new MPOs" />
                    <Field label="Session Timeout (hours)" type="number" value={docSettings.sessionHours} onChange={udp("sessionHours")} placeholder="8" note="Automatic logout after inactivity" />
                  </div>
                  <div style={{ display: "flex", alignItems: "center", gap: 10, padding: "10px 12px", background: "var(--bg3)", borderRadius: 10, border: "1px solid var(--border)" }}>
                    <input id="round-naira" type="checkbox" checked={docSettings.roundToWholeNaira} onChange={e => udp("roundToWholeNaira")(e.target.checked)} style={{ accentColor: "var(--accent)", width: 16, height: 16 }} />
                    <label htmlFor="round-naira" style={{ fontSize: 13, color: "var(--text2)" }}>Round new MPO totals to the nearest naira</label>
                  </div>
                  <Field label="Standard MPO Terms" value={docSettings.mpoTerms} onChange={udp("mpoTerms")} rows={7} placeholder="One term per line" note="Each line becomes a numbered contract term on new MPOs." />
                  <div style={{ display: "flex", justifyContent: "flex-end" }}>
                    {canManageWorkspace && <Btn onClick={saveDocumentSettings} icon="✓">Save Document Defaults</Btn>}
                  </div>
                </div>
              </Card>
            </div>
          )}

          {section === "team" && (
            <div className="fade">
              <Card>
                <h2 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 16, marginBottom: 12 }}>Team Roles & Permissions</h2>
                <p style={{ color: "var(--text2)", fontSize: 13, marginBottom: 16 }}>Admins can control what each user in the agency can edit.</p>
                <div style={{ display: "grid", gap: 10 }}>
                  {(members || []).length === 0 ? (
                    <div style={{ color: "var(--text2)", fontSize: 13 }}>No team members loaded yet.</div>
                  ) : (members || []).map(member => (
                    <div key={member.id} style={{ display: "grid", gridTemplateColumns: "1.4fr 1fr auto", gap: 12, alignItems: "center", padding: "12px 14px", background: "var(--bg3)", border: "1px solid var(--border)", borderRadius: 10 }}>
                      <div>
                        <div style={{ fontWeight: 700, fontSize: 13 }}>{member.name}{member.id === user.id ? " (You)" : ""}</div>
                        <div style={{ fontSize: 12, color: "var(--text2)", marginTop: 3 }}>{member.email}</div>
                        {member.title && <div style={{ fontSize: 11, color: "var(--text3)", marginTop: 2 }}>{member.title}</div>}
                      </div>
                      <Field value={memberRoles[member.id] || member.role || "viewer"} onChange={value => setMemberRoles(prev => ({ ...prev, [member.id]: normalizeRole(value) }))} options={Object.entries(ROLE_LABELS).map(([value, label]) => ({ value, label }))} />
                      <Btn size="sm" onClick={() => saveMemberRole(member.id)} loading={savingMemberId === member.id}>Save Role</Btn>
                    </div>
                  ))}
                </div>
                <div style={{ marginTop: 16, padding: "12px 14px", background: "rgba(59,126,245,.06)", border: "1px solid rgba(59,126,245,.16)", borderRadius: 10, fontSize: 12, color: "var(--text2)", lineHeight: 1.6 }}>
                  <strong>Role guide:</strong> Admin = full control. Planner = clients/campaigns/MPO creation. Buyer = vendors/rates/MPO creation. Finance = MPO status updates and reporting. Viewer = read-only.
                </div>
              </Card>
            </div>
          )}

          {/* ── SECURITY ── */}
          {section === "security" && (
            <div className="fade">
              <Card style={{ marginBottom: 16 }}>
                <h2 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 16, marginBottom: 6 }}>Account Info</h2>
                <InfoRow label="Registered Email" value={user.email} />
                <InfoRow label="Role" value={formatRoleLabel(user.role)} />
                <InfoRow label="Account Created" value={user.createdAt ? new Date(user.createdAt).toLocaleDateString("en-NG", { day: "numeric", month: "long", year: "numeric" }) : "—"} />
                <InfoRow label="User ID" value={user.id?.slice(0, 12) + "…"} accent="var(--text3)" />
              </Card>
              <Card>
                <h2 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 16, marginBottom: 18 }}>Change Password</h2>
                <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
                  <Field label="Current Password" type="password" value={sf.current} onChange={us("current")} placeholder="••••••••" />
                  <Field label="New Password" type="password" value={sf.newPw} onChange={us("newPw")} placeholder="••••••••" note="Minimum 6 characters" />
                  <Field label="Confirm New Password" type="password" value={sf.confirm} onChange={us("confirm")} placeholder="••••••••" />
                  <div style={{ display: "flex", justifyContent: "flex-end" }}>
                    <Btn onClick={changePassword} icon="🔒">Update Password</Btn>
                  </div>
                </div>
              </Card>
            </div>
          )}

          {/* ── NOTIFICATIONS ── */}
          {section === "notifications" && (
            <div className="fade">
              <Card>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 12, flexWrap: "wrap", marginBottom: 16 }}>
                  <div>
                    <h2 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 16 }}>Notifications Inbox</h2>
                    <div style={{ fontSize: 12, color: "var(--text2)", marginTop: 4 }}>Alerts for approvals, finance updates, proof of airing, and reconciliation milestones.</div>
                  </div>
                  <div style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
                    <div style={{ minWidth: 220 }}>
                      <Field value={notificationFilter} onChange={setNotificationFilter} options={[
                        { value: "all", label: "All Notifications" },
                        { value: "unread", label: "Unread Only" },
                        { value: "mpo", label: "MPO Workflow" },
                        { value: "finance", label: "Finance" },
                        { value: "proof", label: "Proof & Reconciliation" },
                        { value: "workspace", label: "Workspace" },
                      ]} />
                    </div>
                    <Btn variant="ghost" size="sm" onClick={onMarkAllNotificationsRead} disabled={!unreadNotifications}>Mark all as read</Btn>
                  </div>
                </div>
                {(() => {
                  const filteredNotifications = (notifications || []).filter(notification => {
                    if (notificationFilter === "unread" && notification.readAt) return false;
                    if (notificationFilter !== "all" && notificationFilter !== "unread" && (notification.category || "workspace") !== notificationFilter) return false;
                    return true;
                  });
                  if (!filteredNotifications.length) return <Empty icon="🔔" title="No notifications" sub="You're all caught up for this workspace." />;
                  return (
                    <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                      {filteredNotifications.map(notification => (
                        <div key={notification.id} style={{ border: "1px solid var(--border)", borderRadius: 10, padding: "12px 14px", background: notification.readAt ? "var(--bg3)" : "rgba(240,165,0,.07)" }}>
                          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 10, flexWrap: "wrap" }}>
                            <div>
                              <div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 14 }}>{notification.title}</div>
                              <div style={{ fontSize: 12, color: "var(--text2)", marginTop: 4 }}>{notification.message || "Open the related workspace page for more details."}</div>
                              <div style={{ fontSize: 11, color: "var(--text3)", marginTop: 8 }}>{notification.actorName || "System"} · {formatRoleLabel(notification.actorRole || "viewer")} · {formatAuditTimestamp(notification.createdAt)}</div>
                            </div>
                            <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                              {!notification.readAt ? <Badge color="accent">Unread</Badge> : <Badge color="gray">Read</Badge>}
                              {!notification.readAt ? <Btn variant="ghost" size="sm" onClick={() => onMarkNotificationRead(notification.id)}>Mark as read</Btn> : null}
                            </div>
                          </div>
                        </div>
                      ))}
                    </div>
                  );
                })()}
              </Card>
            </div>
          )}

          {/* ── ACTIVITY ── */}
          {section === "activity" && (
            <div className="fade">
              <Card>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 12, flexWrap: "wrap", marginBottom: 16 }}>
                  <div>
                    <h2 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 16 }}>Workspace Activity</h2>
                    <div style={{ fontSize: 12, color: "var(--text2)", marginTop: 4 }}>Approval actions, MPO workflow changes, and key workspace edits appear here.</div>
                  </div>
                  <div style={{ minWidth: 220 }}>
                    <Field value={activityFilter} onChange={setActivityFilter} options={[
                      { value: "all", label: "All Activity" },
                      { value: "mpo", label: "MPO Workflow" },
                      { value: "member", label: "Team & Roles" },
                      { value: "agency", label: "Agency" },
                      { value: "workspace", label: "Workspace Settings" },
                      { value: "profile", label: "Profiles" },
                      { value: "security", label: "Security" },
                    ]} />
                  </div>
                </div>
                {activityLoading ? (
                  <div style={{ color: "var(--text2)", fontSize: 13 }}>Loading activity…</div>
                ) : activityItems.length ? (
                  <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                    {activityItems.map(event => (
                      <div key={event.id} style={{ border: "1px solid var(--border)", borderRadius: 10, padding: "12px 14px", background: "var(--bg3)" }}>
                        <div style={{ display: "flex", justifyContent: "space-between", gap: 12, flexWrap: "wrap" }}>
                          <div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 14 }}>
                            {(event.recordType || "workspace").toUpperCase()} · {(event.action || "updated").replace(/_/g, " ").toUpperCase()}
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
                ) : <Empty icon="🕘" title="No activity yet" sub="Actions will start appearing here once your team begins working." />}
              </Card>
            </div>
          )}

          {/* ── DATA ── */}
          {section === "data" && (
            <div className="fade">
              <Card style={{ marginBottom: 14 }}>
                <h2 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 16, marginBottom: 6 }}>Workspace Data & Maintenance</h2>
                <p style={{ color: "var(--text2)", fontSize: 13, lineHeight: 1.6, marginBottom: 16 }}>Your operational records now live in your Supabase workspace. This page is for exporting a snapshot, reviewing what is synced, and cleaning up browser-only preferences on this device.</p>
                <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 16 }}>
                  <Btn variant="blue" icon="⬇" onClick={exportBackup}>Export Workspace Snapshot</Btn>
                  <Btn variant="secondary" icon="⬆" onClick={() => backupInputRef.current?.click()}>Import Snapshot</Btn>
                  <input ref={backupInputRef} type="file" accept="application/json" style={{ display: "none" }} onChange={e => handleBackupImport(e.target.files?.[0])} />
                </div>
                <div style={{ background: "var(--bg3)", borderRadius: 10, border: "1px solid var(--border)", padding: "12px 14px", marginBottom: 14 }}>
                  <div style={{ fontSize: 12, color: "var(--text2)", lineHeight: 1.6 }}>Snapshot export includes vendors, clients, campaigns, rates, MPOs, notifications, and document settings for review. Direct browser snapshot import is intentionally disabled for cloud workspaces to prevent overwriting live Supabase data accidentally.</div>
                </div>
                {[
                  [vendors, "Vendors"], [clients, "Clients & Brands"],
                  [campaigns, "Campaigns"], [rates, "Media Rates"], [mpos, "MPOs"], [notifications, "Notifications"],
                ].map(([data, label]) => (
                  <div key={label} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "10px 0", borderBottom: "1px solid var(--border)", fontSize: 13 }}>
                    <span style={{ color: "var(--text2)" }}>{label}</span>
                    <span style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, color: (data?.length || 0) > 0 ? "var(--accent)" : "var(--text3)" }}>{data?.length || 0} record{(data?.length || 0) !== 1 ? "s" : ""}</span>
                  </div>
                ))}
              </Card>
              <Card style={{ marginBottom: 14 }}>
                <h2 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 16, marginBottom: 6 }}>Browser-only Preferences</h2>
                <p style={{ color: "var(--text2)", fontSize: 13, marginBottom: 16, lineHeight: 1.6 }}>These controls affect only this browser. They do not delete live Supabase records for your agency.</p>
                <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(220px,1fr))", gap: 10 }}>
                  <div style={{ padding: "12px 14px", background: "var(--bg3)", borderRadius: 10, border: "1px solid var(--border)" }}>
                    <div style={{ fontSize: 11, color: "var(--text3)", textTransform: "uppercase", letterSpacing: ".08em", fontWeight: 700 }}>Theme Preference</div>
                    <div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, marginTop: 6 }}>{store.get(themeKeyForUser(user?.id || null), getDefaultTheme(user?.id || null))}</div>
                    <div style={{ fontSize: 11, color: "var(--text3)", marginTop: 4 }}>Stored per user on this device.</div>
                  </div>
                  <div style={{ padding: "12px 14px", background: "var(--bg3)", borderRadius: 10, border: "1px solid var(--border)" }}>
                    <div style={{ fontSize: 11, color: "var(--text3)", textTransform: "uppercase", letterSpacing: ".08em", fontWeight: 700 }}>Signature Cache</div>
                    <div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, marginTop: 6 }}>{user?.signatureDataUrl ? "Available" : "Not saved"}</div>
                    <div style={{ fontSize: 11, color: "var(--text3)", marginTop: 4 }}>Used to keep the signatory visible between sessions.</div>
                  </div>
                </div>
              </Card>
              <Card style={{ border: "1px solid rgba(239,68,68,.25)", background: "rgba(239,68,68,.04)" }}>
                <h2 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 16, marginBottom: 6, color: "var(--red)" }}>⚠️ Local Cleanup</h2>
                <p style={{ color: "var(--text2)", fontSize: 13, marginBottom: 16, lineHeight: 1.6 }}>Use these only when this browser has stale preferences. Your Supabase workspace records remain intact.</p>
                <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                  {[
                    { label: "Clear MPO Draft Cache", desc: "Remove the autosaved draft on this browser only.", onYes: () => { store.del("msp_mpo_draft"); setToast({ msg: "Local MPO draft cache cleared.", type: "success" }); setConfirm(null); } },
                    { label: "Reset My Theme Preference", desc: "Forget this user’s local light/dark mode on this browser.", onYes: () => { store.del(themeKeyForUser(user?.id || null)); setToast({ msg: "Local theme preference reset. Reload to apply system default.", type: "success" }); setConfirm(null); } },
                    { label: "Clear My Saved Signature Cache", desc: "Removes only the browser cache copy. Your saved profile value stays in Supabase metadata if available.", onYes: () => { setStoredUserSignature(user?.id, ""); setToast({ msg: "Local signature cache cleared.", type: "success" }); setConfirm(null); } },
                  ].map(({ label, desc, onYes }) => (
                    <div key={label} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "12px 14px", background: "var(--bg3)", borderRadius: 10, border: "1px solid var(--border2)" }}>
                      <div><div style={{ fontSize: 13, fontWeight: 600 }}>{label}</div><div style={{ fontSize: 11, color: "var(--text3)", marginTop: 2 }}>{desc}</div></div>
                      {canDanger && <Btn variant="danger" size="sm" onClick={() => setConfirm({ msg: `${label}? This only affects this browser.`, onYes })}>Clear</Btn>}
                    </div>
                  ))}
                </div>
              </Card>
            </div>
          )}

        </div>
      </div>
    </div>
  );
};



const ReportsPage = ({ vendors, clients, campaigns, rates, mpos }) => {
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
    },
    {
      id: "spend-client",
      title: "Spend by Client",
      icon: "👥",
      desc: "Track client value, budget coverage, and campaign concentration.",
      headers: ["Client", "Campaigns", "MPO Count", "Budget Pool", "Spend", "Budget Variance"],
      rows: spendByClient.map(row => [row.client, row.campaignCount, row.mpoCount, row.budget, row.spend, row.variance]),
    },
    {
      id: "campaign-budget",
      title: "Campaign Budget Control",
      icon: "🎯",
      desc: "Compare campaign budgets against MPO spend and utilization.",
      headers: ["Campaign", "Client", "Brand", "Medium", "Status", "Budget", "Gross Value", "MPO Spend", "Variance", "Utilization %", "MPO Count"],
      rows: campaignBudgetControl.map(row => [row.campaign, row.client, row.brand, row.medium, row.status, row.budget, row.gross, row.spend, row.variance, `${row.utilization.toFixed(1)}%`, row.mpoCount]),
    },
    {
      id: "finance-tracker",
      title: "Finance Tracker",
      icon: "💸",
      desc: "Follow invoice, payment, and outstanding exposure on every MPO.",
      headers: ["MPO No.", "Vendor", "Client", "Campaign", "Value", "Invoice Status", "Invoice No.", "Payment Status", "Payment Ref", "Paid At", "Outstanding"],
      rows: financeTracker.map(row => [row.mpoNo, row.vendor, row.client, row.campaign, row.value, row.invoiceStatus, row.invoiceNo, row.paymentStatus, row.paymentReference, row.paidAt, row.outstanding]),
    },
    {
      id: "reconciliation-control",
      title: "Reconciliation Control",
      icon: "🧾",
      desc: "Watch proof of airing, delivery variance, and reconciliation status.",
      headers: ["MPO No.", "Vendor", "Campaign", "Planned Spots", "Aired Spots", "Missed", "Makegood", "Proof Status", "Reconciliation", "Reconciled Amount", "Notes"],
      rows: reconciliationControl.map(row => [row.mpoNo, row.vendor, row.campaign, row.planned, row.aired, row.missed, row.makegood, row.proofStatus, row.reconciliationStatus, row.reconciledAmount, row.notes]),
    },
    {
      id: "pipeline",
      title: "Status Pipeline",
      icon: "📈",
      desc: "Measure MPO flow through draft, approval, airing, and closeout.",
      headers: ["Status", "Count", "Value", "Paid Value"],
      rows: statusPipeline.map(row => [row.status, row.count, row.value, row.paid]),
    },
    {
      id: "rate-snapshot",
      title: "Rate Card Snapshot",
      icon: "💰",
      desc: "Review effective net rates after discount and commission.",
      headers: ["Vendor", "Programme", "Time Belt", "Medium", "Duration", "Rate/Spot", "Disc %", "Comm %", "Net Rate", "Notes"],
      rows: rateCardSnapshot.map(row => [row.vendor, row.programme, row.timeBelt, row.medium, row.duration, row.rate, `${row.discount}%`, `${row.commission}%`, row.netRate, row.notes]),
    },
    {
      id: "mpo-register",
      title: "Filtered MPO Register",
      icon: "📄",
      desc: "A clean register of all MPOs matching your filters.",
      headers: ["MPO No.", "Date", "Vendor", "Client", "Brand", "Campaign", "Medium", "Status", "Payment Status", "Reconciliation", "Total Spots", "Gross", "Net", "Grand Total"],
      rows: filteredMpoRegister.map(row => [row.mpoNo, row.date, row.vendor, row.client, row.brand, row.campaign, row.medium, row.status, row.paymentStatus, row.reconciliationStatus, row.totalSpots, row.gross, row.net, row.grandTotal]),
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
    const rows = summaryCards.map(card => [card.label, typeof card.value === "number" ? card.value : card.value, card.sub]);
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
        {summaryCards.map(card => (
          <Card key={card.label} hoverable style={{ position: "relative", overflow: "hidden", padding: 18 }}>
            <div style={{ position: "absolute", top: -14, right: -10, fontSize: 64, opacity: .05 }}>{card.icon}</div>
            <div style={{ fontSize: 11, color: "var(--text3)", fontWeight: 600, textTransform: "uppercase", letterSpacing: ".08em", marginBottom: 7 }}>{card.label}</div>
            <div style={{ fontSize: typeof card.value === "number" ? 24 : 22, fontWeight: 800, fontFamily: "'Syne',sans-serif", color: card.color }}>
              {typeof card.value === "number" ? fmtN(card.value) : card.value}
            </div>
            <div style={{ fontSize: 12, color: "var(--text2)", marginTop: 4 }}>{card.sub}</div>
          </Card>
        ))}
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
                <Btn variant="blue" size="sm" icon="⬇" onClick={() => buildSectionExport(section.title, section.headers, section.rows, filterDescriptor)}>
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
                            const display = typeof cell === "number" && Math.abs(cell) >= 1000 ? fmtN(cell) : cell;
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


/* ── MAIN APP ───────────────────────────────────────────── */
export default function App() {
  const [user, setUser] = useState(null);
  const [authUser, setAuthUser] = useState(null);
  const [page, setPage] = useState("dashboard");
  const [collapsed, setCollapsed] = useState(false);
  const [theme, setTheme] = useState(() => getDefaultTheme());
  const [appSettings, _setAppSettings] = useState(() => getAppSettings());
  useEffect(() => {
    setTheme(getDefaultTheme(user?.id || null));
  }, [user?.id]);

  const setAppSettings = useCallback((value) => {
    _setAppSettings(prev => {
      const next = typeof value === "function" ? value(prev) : value;
      store.set("msp_app_settings", next);
      if (user?.agencyId) {
        saveAppSettingsToSupabase(user.agencyId, next).catch(error => console.error("Failed to persist app settings:", error));
      }
      return next;
    });
  }, [user?.agencyId]);
  const toggleTheme = () => setTheme(t => {
    const n = t === "light" ? "dark" : "light";
    store.set(themeKeyForUser(user?.id || null), n);
    return n;
  });

  const [vendors, setVendors] = useState([]);
  const [clients, setClients] = useState([]);
  const [campaigns, setCampaigns] = useState([]);
  const [rates, setRates] = useState([]);
  const [mpos, setMpos] = useState([]);
  const [members, setMembers] = useState([]);
  const [notifications, setNotifications] = useState([]);
  const [alertsOpen, setAlertsOpen] = useState(false);

// MPOs now come from Supabase

  const [authReady, setAuthReady] = useState(false);

  useEffect(() => {
    let mounted = true;
    const timeout = setTimeout(() => {
      if (mounted) setAuthReady(true);
    }, 5000);

    const bootstrapAuth = async () => {
      try {
        const { data } = await supabase.auth.getSession();
        if (!mounted) return;
        setAuthUser(data?.session?.user || null);
      } catch (e) {
        console.error("Failed to bootstrap auth:", e);
      } finally {
        if (mounted) setAuthReady(true);
        clearTimeout(timeout);
      }
    };

    bootstrapAuth();

    const { data: { subscription } } = supabase.auth.onAuthStateChange((_event, session) => {
      if (!mounted) return;
      setAuthUser(session?.user || null);
      setAuthReady(true);
      clearTimeout(timeout);
    });

    return () => {
      mounted = false;
      clearTimeout(timeout);
      subscription.unsubscribe();
    };
  }, []);

  useEffect(() => {
    let active = true;

    const loadUser = async () => {
      if (!authUser?.id) {
        if (active) setUser(null);
        return;
      }

      try {
        await ensureAgencyForUser(authUser);
        const appUser = await loadAppUserFromSupabase(authUser);
        if (active) setUser(appUser);
      } catch (e) {
        console.error("Failed to load user:", e);
        if (active) {
          setUser({
            id: authUser.id,
            name: authUser.user_metadata?.full_name || authUser.email || "User",
            email: authUser.email || "",
            title: authUser.user_metadata?.title || "",
            phone: authUser.user_metadata?.phone || "",
            agency: authUser.user_metadata?.agency_name || "My Agency",
            agencyId: null,
            agencyCode: authUser.user_metadata?.agency_code || "",
            role: "admin",
          });
        }
      }
    };

    loadUser();
    return () => { active = false; };
  }, [authUser?.id]);

  useEffect(() => {
    if (!user?.agencyId) {
      _setAppSettings(getAppSettings());
      return;
    }

    let active = true;
    const loadWorkspaceSettings = async () => {
      try {
        const settings = await fetchAppSettingsFromSupabase(user.agencyId);
        if (!active) return;
        _setAppSettings(settings);
        store.set("msp_app_settings", settings);
      } catch (e) {
        console.error("Failed to load app settings:", e);
      }
    };

    loadWorkspaceSettings();
    return () => { active = false; };
  }, [user?.agencyId]);

  useEffect(() => {
    if (!user?.agencyId) {
      setMembers([]);
      return;
    }

    let active = true;
    const loadMembers = async () => {
      try {
        const rows = await fetchAgencyMembersFromSupabase(user.agencyId);
        if (active) setMembers(rows);
      } catch (e) {
        console.error("Failed to load members:", e);
      }
    };

    loadMembers();
    return () => { active = false; };
  }, [user?.agencyId]);

  useEffect(() => {
    if (!user?.agencyId || !user?.id) {
      setNotifications([]);
      return;
    }

    let active = true;
    const loadNotifications = async () => {
      try {
        const rows = await fetchNotificationsFromSupabase(user.id, user.agencyId);
        if (active) setNotifications(rows);
      } catch (e) {
        console.error("Failed to load notifications:", e);
      }
    };

    loadNotifications();
    return () => { active = false; };
  }, [user?.agencyId, user?.id]);

  useEffect(() => {
    if (!user?.agencyId) {
      setVendors([]);
      return;
    }

    let active = true;
    const loadVendors = async () => {
      try {
        const rows = await fetchVendorsFromSupabase();
        if (active) setVendors(rows);
      } catch (e) {
        console.error("Failed to load vendors:", e);
      }
    };

    loadVendors();
    return () => { active = false; };
  }, [user?.agencyId]);

  useEffect(() => {
    if (!user?.agencyId) {
      setClients([]);
      return;
    }

    let active = true;
    const loadClients = async () => {
      try {
        const rows = await fetchClientsFromSupabase();
        if (active) setClients(rows);
      } catch (e) {
        console.error("Failed to load clients:", e);
      }
    };

    loadClients();
    return () => { active = false; };
  }, [user?.agencyId]);

  useEffect(() => {
    if (!user?.agencyId) {
      setCampaigns([]);
      return;
    }

    let active = true;
    const loadCampaigns = async () => {
      try {
        const rows = await fetchCampaignsFromSupabase();
        if (active) setCampaigns(rows);
      } catch (e) {
        console.error("Failed to load campaigns:", e);
      }
    };

    loadCampaigns();
    return () => { active = false; };
  }, [user?.agencyId]);

  useEffect(() => {
    if (!user?.agencyId) {
      setRates([]);
      return;
    }

    let active = true;
    const loadRates = async () => {
      try {
        const rows = await fetchRatesFromSupabase();
        if (active) setRates(rows);
      } catch (e) {
        console.error("Failed to load rates:", e);
      }
    };

    loadRates();
    return () => { active = false; };
  }, [user?.agencyId]);

  useEffect(() => {
    if (!user?.agencyId) {
      setMpos([]);
      return;
    }

    let active = true;
    const loadMpos = async () => {
      try {
        const rows = await fetchMposFromSupabase();
        if (active) setMpos(rows);
      } catch (e) {
        console.error("Failed to load MPOs:", e);
      }
    };

    loadMpos();
    return () => { active = false; };
  }, [user?.agencyId]);

  useEffect(() => {
    if (!user?.agencyId) return;

    const refreshUser = async () => {
      if (!authUser?.id) return;
      try {
        const refreshed = await loadAppUserFromSupabase(authUser);
        setUser(prev => prev ? { ...prev, ...refreshed } : refreshed);
      } catch (error) {
        console.error("Failed to refresh user:", error);
      }
    };
    const refreshVendors = async () => {
      try { setVendors(await fetchVendorsFromSupabase()); } catch (error) { console.error("Realtime vendors refresh failed:", error); }
    };
    const refreshClients = async () => {
      try { setClients(await fetchClientsFromSupabase()); } catch (error) { console.error("Realtime clients refresh failed:", error); }
    };
    const refreshCampaigns = async () => {
      try { setCampaigns(await fetchCampaignsFromSupabase()); } catch (error) { console.error("Realtime campaigns refresh failed:", error); }
    };
    const refreshRates = async () => {
      try { setRates(await fetchRatesFromSupabase()); } catch (error) { console.error("Realtime rates refresh failed:", error); }
    };
    const refreshMpos = async () => {
      try { setMpos(await fetchMposFromSupabase()); } catch (error) { console.error("Realtime MPO refresh failed:", error); }
    };
    const refreshMembers = async () => {
      try { setMembers(await fetchAgencyMembersFromSupabase(user.agencyId)); } catch (error) { console.error("Realtime members refresh failed:", error); }
    };
    const refreshNotifications = async () => {
      try { setNotifications(await fetchNotificationsFromSupabase(user.id, user.agencyId)); } catch (error) { console.error("Realtime notifications refresh failed:", error); }
    };
    const refreshSettings = async () => {
      try {
        const settings = await fetchAppSettingsFromSupabase(user.agencyId);
        _setAppSettings(settings);
      } catch (error) {
        console.error("Realtime settings refresh failed:", error);
      }
    };

    const channel = supabase
      .channel(`agency-live-${user.agencyId}`)
      .on("postgres_changes", { event: "*", schema: "public", table: "vendors", filter: `agency_id=eq.${user.agencyId}` }, refreshVendors)
      .on("postgres_changes", { event: "*", schema: "public", table: "clients", filter: `agency_id=eq.${user.agencyId}` }, refreshClients)
      .on("postgres_changes", { event: "*", schema: "public", table: "campaigns", filter: `agency_id=eq.${user.agencyId}` }, refreshCampaigns)
      .on("postgres_changes", { event: "*", schema: "public", table: "rates", filter: `agency_id=eq.${user.agencyId}` }, refreshRates)
      .on("postgres_changes", { event: "*", schema: "public", table: "mpos", filter: `agency_id=eq.${user.agencyId}` }, refreshMpos)
      .on("postgres_changes", { event: "*", schema: "public", table: "mpo_spots" }, refreshMpos)
      .on("postgres_changes", { event: "*", schema: "public", table: "app_settings", filter: `agency_id=eq.${user.agencyId}` }, refreshSettings)
      .on("postgres_changes", { event: "*", schema: "public", table: "agencies", filter: `id=eq.${user.agencyId}` }, async () => { await refreshUser(); await refreshSettings(); })
      .on("postgres_changes", { event: "*", schema: "public", table: "profiles", filter: `agency_id=eq.${user.agencyId}` }, async () => { await refreshUser(); await refreshMembers(); })
      .on("postgres_changes", { event: "*", schema: "public", table: "notifications", filter: `recipient_user_id=eq.${user.id}` }, refreshNotifications)
      .subscribe();

    return () => {
      supabase.removeChannel(channel).catch(() => {});
    };
  }, [user?.agencyId, user?.id, authUser?.id]);

  const unreadNotifications = notifications.filter(notification => !notification.readAt).length;

  const workspaceAlerts = (notifications || []).slice(0, 12).map(notification => ({
    id: notification.id,
    icon: notification.category === "finance" ? "💳" : notification.category === "reconciliation" ? "📑" : notification.category === "proof" ? "📎" : "🔔",
    title: notification.title || "Notification",
    message: notification.message || "Open settings to view this alert.",
    page: notification.linkPage || "settings",
    isUnread: !notification.readAt,
  }));

  const handleMarkNotificationRead = async (notificationId) => {
    try {
      const updated = await markNotificationReadInSupabase(notificationId);
      if (!updated) return;
      setNotifications(items => items.map(item => item.id === notificationId ? updated : item));
    } catch (error) {
      console.error("Failed to mark notification as read:", error);
    }
  };

  const handleMarkAllNotificationsRead = async () => {
    try {
      await markAllNotificationsReadInSupabase(user?.id, user?.agencyId);
      const timestamp = new Date().toISOString();
      setNotifications(items => items.map(item => item.readAt ? item : ({ ...item, readAt: timestamp })));
    } catch (error) {
      console.error("Failed to mark all notifications as read:", error);
    }
  };

  const handleLogin = () => {};

  const handleLogout = async () => {
    setPage("dashboard");
    setUser(null);
    setAuthUser(null);
    setVendors([]);
    setClients([]);
    setCampaigns([]);
    setRates([]);
    setMpos([]);
    setMembers([]);
    _setAppSettings(getAppSettings());
    try {
      await supabase.auth.signOut({ scope: "local" });
    } catch (e) {
      console.error("Failed to sign out:", e);
    }
  };

  const handleUserUpdate = (u) => setUser(prev => ({ ...prev, ...u }));

if (!authReady) {
  return (
    <>
      <GlobalStyle theme={theme} />
      <div
        style={{
          minHeight: "100vh",
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
          background: "var(--bg)",
          color: "var(--text)",
          fontFamily: "'Syne',sans-serif",
          fontWeight: 700,
        }}
      >
        Loading...
      </div>
    </>
  );
}

if (!user) {
  return (
    <>
      <GlobalStyle theme={theme} />
      <AuthPage onLogin={handleLogin} />
    </>
  );
}

  const pp = { vendors, clients, campaigns, rates, mpos, notifications, unreadNotifications, setVendors, setClients, setCampaigns, setRates, setMpos };

  return (
    <>
      <GlobalStyle theme={theme} />
      <div style={{ display: "flex", minHeight: "100vh" }}>
        <Sidebar page={page} setPage={setPage} user={user} onLogout={handleLogout} collapsed={collapsed} setCollapsed={setCollapsed} theme={theme} toggleTheme={toggleTheme} unreadNotifications={unreadNotifications} />
        <main style={{ flex: 1, overflowY: "auto", padding: "28px 28px 52px", position: "relative" }}>
          <TopRightNotificationsButton count={unreadNotifications} onClick={() => setAlertsOpen(true)} />
          {alertsOpen && (
            <Modal title="Workspace Alerts" onClose={() => setAlertsOpen(false)} width={560}>
              <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
                {workspaceAlerts.length === 0 ? (
                  <Empty icon="🔔" title="No alerts right now" sub="Your workspace looks clear." />
                ) : workspaceAlerts.map(alert => (
                  <Card key={alert.id} style={{ padding: 16 }}>
                    <div style={{ display: "flex", gap: 12, alignItems: "flex-start" }}>
                      <div style={{ width: 40, height: 40, borderRadius: 10, background: "var(--bg3)", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 18, flexShrink: 0 }}>{alert.icon}</div>
                      <div style={{ flex: 1 }}>
                        <div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 14 }}>{alert.title}</div>
                        <div style={{ fontSize: 13, color: "var(--text2)", marginTop: 4 }}>{alert.message}</div>
                      </div>
                      {alert.isUnread ? <Badge color="accent">New</Badge> : null}
                    </div>
                  </Card>
                ))}
                <div style={{ display: "flex", justifyContent: "flex-end" }}>
                  <Btn variant="ghost" onClick={() => { setAlertsOpen(false); setPage("settings"); }}>Open Inbox</Btn>
                </div>
              </div>
            </Modal>
          )}
          {page === "dashboard"  && <Dashboard user={user} {...pp} setPage={setPage} onOpenNotifications={() => setPage("settings")} />}
          {page === "vendors"    && <VendorsPage {...pp} user={user} />}
          {page === "clients"    && <ClientsPage {...pp} user={user} />}
          {page === "campaigns"  && <CampaignsPage {...pp} user={user} />}
          {page === "rates"      && <RatesPage {...pp} user={user} />}
          {page === "mpo"        && <MPOPage {...pp} user={user} appSettings={appSettings} />}
          {page === "reports"    && <ReportsPage {...pp} />}
          {page === "settings"   && <SettingsPage user={user} onUserUpdate={handleUserUpdate} onLogout={handleLogout} appSettings={appSettings} setAppSettings={setAppSettings} vendors={vendors} clients={clients} campaigns={campaigns} rates={rates} mpos={mpos} members={members} setMembers={setMembers} notifications={notifications} unreadNotifications={unreadNotifications} onMarkNotificationRead={handleMarkNotificationRead} onMarkAllNotificationsRead={handleMarkAllNotificationsRead} />}
        </main>
      </div>
    </>
  );
}
