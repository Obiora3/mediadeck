import { supabase } from '../lib/supabase';

export const normalizeVendorName = (value) => String(value ?? '')
  .trim()
  .replace(/\s+/g, ' ')
  .toLowerCase();

const vendorListSelect = `
  id,
  agency_id,
  name,
  type,
  contact,
  email,
  phone,
  location,
  default_rate,
  discount,
  commission,
  notes,
  is_archived,
  created_at,
  updated_at
`;

const mapVendorFromSupabase = (v) => ({
  id: v.id,
  name: v.name || '',
  type: v.type || '',
  contact: v.contact || '',
  email: v.email || '',
  phone: v.phone || '',
  location: v.location || '',
  rate: v.default_rate ?? '',
  discount: v.discount ?? '',
  commission: v.commission ?? '',
  notes: v.notes || '',
  archivedAt: v.is_archived ? (v.updated_at ? new Date(v.updated_at).getTime() : Date.now()) : null,
  createdAt: v.created_at ? new Date(v.created_at).getTime() : Date.now(),
});

export const fetchVendorsFromSupabase = async (agencyId) => {
  if (!agencyId) return [];

  const { data, error } = await supabase
    .from('vendors')
    .select(vendorListSelect)
    .eq('agency_id', agencyId)
    .order('created_at', { ascending: false });

  if (error) throw error;

  return (data || []).map(mapVendorFromSupabase);
};

export const createVendorInSupabase = async (agencyId, userId, form) => {
  if (!agencyId) throw new Error('No agency found for this user.');

  const { data, error } = await supabase
    .from('vendors')
    .insert([{
      agency_id: agencyId,
      created_by: userId,
      name: form.name,
      type: form.type,
      contact: form.contact || null,
      email: form.email || null,
      phone: form.phone || null,
      location: form.location || null,
      default_rate: form.rate ? Number(form.rate) : null,
      discount: form.discount ? Number(form.discount) : 0,
      commission: form.commission ? Number(form.commission) : 0,
      notes: form.notes || null,
      is_archived: false,
    }])
    .select(vendorListSelect)
    .single();

  if (error) throw error;

  return mapVendorFromSupabase(data);
};

export const updateVendorInSupabase = async (vendorId, form) => {
  const { data, error } = await supabase
    .from('vendors')
    .update({
      name: form.name,
      type: form.type,
      contact: form.contact || null,
      email: form.email || null,
      phone: form.phone || null,
      location: form.location || null,
      default_rate: form.rate ? Number(form.rate) : null,
      discount: form.discount ? Number(form.discount) : 0,
      commission: form.commission ? Number(form.commission) : 0,
      notes: form.notes || null,
    })
    .eq('id', vendorId)
    .select(vendorListSelect)
    .single();

  if (error) throw error;

  return mapVendorFromSupabase(data);
};

export const archiveVendorInSupabase = async (vendorId) => {
  const { data, error } = await supabase
    .from('vendors')
    .update({ is_archived: true })
    .eq('id', vendorId)
    .select(vendorListSelect)
    .single();

  if (error) throw error;

  return mapVendorFromSupabase(data);
};

export const restoreVendorInSupabase = async (vendorId) => {
  const { data, error } = await supabase
    .from('vendors')
    .update({ is_archived: false })
    .eq('id', vendorId)
    .select(vendorListSelect)
    .single();

  if (error) throw error;

  return mapVendorFromSupabase(data);
};


export const findVendorByNameInSupabase = async (agencyId, vendorName, options = {}) => {
  const normalizedName = normalizeVendorName(vendorName);
  if (!agencyId || !normalizedName) return null;

  const includeArchived = options.includeArchived !== false;

  let query = supabase
    .from('vendors')
    .select(vendorListSelect)
    .eq('agency_id', agencyId);

  if (!includeArchived) {
    query = query.eq('is_archived', false);
  }

  const { data, error } = await query;
  if (error) throw error;

  const match = (data || []).find((vendor) => normalizeVendorName(vendor.name) === normalizedName);
  return match ? mapVendorFromSupabase(match) : null;
};

export const ensureVendorExistsInSupabase = async (agencyId, userId, vendorName, defaults = {}) => {
  const normalizedName = normalizeVendorName(vendorName);
  if (!agencyId || !normalizedName) return null;

  const existingAny = await findVendorByNameInSupabase(agencyId, vendorName, { includeArchived: true });
  if (existingAny) {
    if (existingAny.archivedAt) {
      const { data, error } = await supabase
        .from('vendors')
        .update({
          is_archived: false,
          type: existingAny.type || defaults.type || '',
          discount: existingAny.discount !== '' && existingAny.discount !== null ? Number(existingAny.discount) : (defaults.discount ? Number(defaults.discount) : 0),
          commission: existingAny.commission !== '' && existingAny.commission !== null ? Number(existingAny.commission) : (defaults.commission ? Number(defaults.commission) : 0),
          notes: existingAny.notes || defaults.notes || 'Auto-created from media rates import.',
        })
        .eq('id', existingAny.id)
        .select(vendorListSelect)
        .single();

      if (error) throw error;
      return mapVendorFromSupabase(data);
    }

    return existingAny;
  }

  return createVendorInSupabase(agencyId, userId, {
    name: String(vendorName ?? '').trim(),
    type: defaults.type || '',
    contact: defaults.contact || '',
    email: defaults.email || '',
    phone: defaults.phone || '',
    location: defaults.location || '',
    rate: defaults.rate || '',
    discount: defaults.discount || '',
    commission: defaults.commission || '',
    notes: defaults.notes || 'Auto-created from media rates import.',
  });
};


export const deleteVendorInSupabase = async (vendorId) => {
  const { error } = await supabase
    .from('vendors')
    .delete()
    .eq('id', vendorId);

  if (error) throw error;
  return vendorId;
};
