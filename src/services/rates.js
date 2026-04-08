import { supabase } from '../lib/supabase';
import { ensureVendorExistsInSupabase, normalizeVendorName } from './vendors';

const normRateText = (value) => String(value ?? '').trim().replace(/\s+/g, ' ').toLowerCase();
const normRateTimeBelt = (value) => normRateText(value).replace(/\s*[-–—]\s*/g, '-');
const normRateDuration = (value) => String(value ?? '30').trim() || '30';
const normalizeImportVendorName = (value) => normalizeVendorName(value);
const autoCreatedVendorNote = 'Auto-created from media rates import.';

const makeRateDuplicateKey = ({ vendorId = '', mediaType = '', programme = '', timeBelt = '', duration = '30' }) => (
  [
    vendorId ? `id:${vendorId}` : 'id:',
    normRateText(mediaType),
    normRateText(programme),
    normRateTimeBelt(timeBelt),
    normRateDuration(duration),
  ].join('|')
);

const rateListSelect = `
  id,
  agency_id,
  vendor_id,
  media_type,
  programme,
  time_belt,
  duration,
  rate_per_spot,
  discount,
  commission,
  vat_rate,
  notes,
  campaign_id,
  client_id,
  is_archived,
  created_at,
  updated_at
`;

const fetchExistingActiveRatesForAgency = async (agencyId, excludeRateId = null) => {
  let query = supabase
    .from('rates')
    .select('id, vendor_id, media_type, programme, time_belt, duration')
    .eq('agency_id', agencyId)
    .eq('is_archived', false);

  if (excludeRateId) query = query.neq('id', excludeRateId);

  const { data, error } = await query;
  if (error) throw error;
  return data || [];
};

const mapRateFromSupabase = (data) => ({
  id: data.id,
  vendorId: data.vendor_id || '',
  mediaType: data.media_type || '',
  programme: data.programme || '',
  timeBelt: data.time_belt || '',
  duration: data.duration ? String(data.duration) : '30',
  ratePerSpot: data.rate_per_spot ?? '',
  discount: data.discount ?? '',
  commission: data.commission ?? '',
  vat: data.vat_rate ?? '0',
  notes: data.notes || '',
  campaignId: data.campaign_id || '',
  clientId: data.client_id || '',
  archivedAt: data.is_archived ? (data.updated_at ? new Date(data.updated_at).getTime() : Date.now()) : null,
  createdAt: data.created_at ? new Date(data.created_at).getTime() : Date.now(),
  updatedAt: data.updated_at ? new Date(data.updated_at).getTime() : Date.now(),
});

export const fetchRatesFromSupabase = async (agencyId) => {
  if (!agencyId) return [];

  const { data, error } = await supabase
    .from('rates')
    .select(rateListSelect)
    .eq('agency_id', agencyId)
    .order('created_at', { ascending: false });

  if (error) throw error;
  return (data || []).map(mapRateFromSupabase);
};

export const createRatesInSupabase = async (agencyId, userId, hdr, validRows) => {
  if (!agencyId) throw new Error('No agency found for this user.');

  const existingRates = await fetchExistingActiveRatesForAgency(agencyId);
  const existingKeys = new Set(
    existingRates.map((row) => makeRateDuplicateKey({
      vendorId: row.vendor_id || '',
      mediaType: row.media_type || '',
      programme: row.programme || '',
      timeBelt: row.time_belt || '',
      duration: row.duration ? String(row.duration) : '30',
    }))
  );

  const seenInBatch = new Set();
  for (const row of validRows) {
    const key = makeRateDuplicateKey({
      vendorId: hdr.vendorId || '',
      mediaType: hdr.mediaType || '',
      programme: row.programme || '',
      timeBelt: row.timeBelt || '',
      duration: row.duration ? String(row.duration) : '30',
    });

    if (existingKeys.has(key) || seenInBatch.has(key)) {
      throw new Error(`Duplicate rate card already exists for ${row.programme || 'this programme'}.`);
    }
    seenInBatch.add(key);
  }

  const payload = validRows.map((row) => ({
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
    .from('rates')
    .insert(payload)
    .select(rateListSelect);

  if (error) throw error;
  return (data || []).map(mapRateFromSupabase);
};

export const updateRateInSupabase = async (rateId, hdr, row) => {
  const { data: currentRate, error: currentError } = await supabase
    .from('rates')
    .select('agency_id')
    .eq('id', rateId)
    .single();

  if (currentError) throw currentError;

  const existingRates = await fetchExistingActiveRatesForAgency(currentRate.agency_id, rateId);
  const existingKeys = new Set(
    existingRates.map((existing) => makeRateDuplicateKey({
      vendorId: existing.vendor_id || '',
      mediaType: existing.media_type || '',
      programme: existing.programme || '',
      timeBelt: existing.time_belt || '',
      duration: existing.duration ? String(existing.duration) : '30',
    }))
  );

  const nextKey = makeRateDuplicateKey({
    vendorId: hdr.vendorId || '',
    mediaType: hdr.mediaType || '',
    programme: row.programme || '',
    timeBelt: row.timeBelt || '',
    duration: row.duration ? String(row.duration) : '30',
  });

  if (existingKeys.has(nextKey)) {
    throw new Error(`Duplicate rate card already exists for ${row.programme || 'this programme'}.`);
  }

  const { data, error } = await supabase
    .from('rates')
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
    .eq('id', rateId)
    .select(rateListSelect)
    .single();

  if (error) throw error;
  return mapRateFromSupabase(data);
};

export const archiveRateInSupabase = async (rateId) => {
  const { data, error } = await supabase
    .from('rates')
    .update({ is_archived: true })
    .eq('id', rateId)
    .select(rateListSelect)
    .single();

  if (error) throw error;
  return mapRateFromSupabase(data);
};

export const restoreRateInSupabase = async (rateId) => {
  const { data, error } = await supabase
    .from('rates')
    .update({ is_archived: false })
    .eq('id', rateId)
    .select(rateListSelect)
    .single();

  if (error) throw error;
  return mapRateFromSupabase(data);
};

export const importRatesInSupabase = async (agencyId, userId, newRates) => {
  if (!agencyId) throw new Error('No agency found for this user.');

  const existingRates = await fetchExistingActiveRatesForAgency(agencyId);
  const existingKeys = new Set(
    existingRates.map((row) => makeRateDuplicateKey({
      vendorId: row.vendor_id || '',
      mediaType: row.media_type || '',
      programme: row.programme || '',
      timeBelt: row.time_belt || '',
      duration: row.duration ? String(row.duration) : '30',
    }))
  );

  const vendorCache = new Map();
  const createdVendorsMap = new Map();
  const preparedRates = [];

  for (const rawRate of newRates) {
    let vendorId = rawRate.vendorId || '';
    const vendorName = rawRate._vendorName || rawRate.vendorName || '';

    if (!vendorId && normalizeImportVendorName(vendorName)) {
      const cacheKey = normalizeImportVendorName(vendorName);
      if (!vendorCache.has(cacheKey)) {
        vendorCache.set(
          cacheKey,
          ensureVendorExistsInSupabase(agencyId, userId, vendorName, {
            type: rawRate.mediaType || '',
            discount: rawRate.discount || '',
            commission: rawRate.commission || '',
            notes: rawRate.notes || autoCreatedVendorNote,
          })
        );
      }

      const ensuredVendor = await vendorCache.get(cacheKey);
      vendorId = ensuredVendor?.id || '';
      if (ensuredVendor?.id) {
        createdVendorsMap.set(ensuredVendor.id, ensuredVendor);
      }
    }

    preparedRates.push({
      ...rawRate,
      vendorId,
    });
  }

  const seenInBatch = new Set();
  const duplicateRows = [];
  const uniqueRates = [];

  for (const r of preparedRates) {
    const key = makeRateDuplicateKey({
      vendorId: r.vendorId || '',
      mediaType: r.mediaType || '',
      programme: r.programme || '',
      timeBelt: r.timeBelt || '',
      duration: r.duration ? String(r.duration) : '30',
    });

    if (existingKeys.has(key) || seenInBatch.has(key)) {
      duplicateRows.push(r);
      continue;
    }

    seenInBatch.add(key);
    uniqueRates.push(r);
  }

  if (!uniqueRates.length) {
    throw new Error('All selected rows are duplicates of existing active rate cards.');
  }

  const payload = uniqueRates.map((r) => ({
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
    .from('rates')
    .insert(payload)
    .select(rateListSelect);

  if (error) throw error;
  return {
    insertedRates: (data || []).map(mapRateFromSupabase),
    duplicateRows,
    createdVendors: Array.from(createdVendorsMap.values()),
  };
};


export const deleteRateInSupabase = async (rateId) => {
  const { error } = await supabase
    .from('rates')
    .delete()
    .eq('id', rateId);

  if (error) throw error;
  return rateId;
};
