import { supabase } from '../lib/supabase';

const clientListSelect = `
  id,
  agency_id,
  name,
  industry,
  contact,
  email,
  phone,
  address,
  brands,
  is_archived,
  created_at,
  updated_at
`;

const mapClientFromSupabase = (c) => ({
  id: c.id,
  name: c.name || '',
  industry: c.industry || '',
  contact: c.contact || '',
  email: c.email || '',
  phone: c.phone || '',
  address: c.address || '',
  brands: c.brands || '',
  archivedAt: c.is_archived ? (c.updated_at ? new Date(c.updated_at).getTime() : Date.now()) : null,
  createdAt: c.created_at ? new Date(c.created_at).getTime() : Date.now(),
  updatedAt: c.updated_at ? new Date(c.updated_at).getTime() : Date.now(),
});

export const fetchClientsFromSupabase = async (agencyId) => {
  if (!agencyId) return [];

  const { data, error } = await supabase
    .from('clients')
    .select(clientListSelect)
    .eq('agency_id', agencyId)
    .order('created_at', { ascending: false });

  if (error) throw error;

  return (data || []).map(mapClientFromSupabase);
};

export const createClientInSupabase = async (agencyId, userId, form) => {
  if (!agencyId) throw new Error('No agency found for this user.');

  const { data, error } = await supabase
    .from('clients')
    .insert([{
      agency_id: agencyId,
      created_by: userId,
      name: form.name.trim(),
      industry: form.industry || null,
      contact: form.contact || null,
      email: form.email || null,
      phone: form.phone || null,
      address: form.address || null,
      brands: form.brands || null,
      is_archived: false,
    }])
    .select(clientListSelect)
    .single();

  if (error) throw error;

  return mapClientFromSupabase(data);
};

export const updateClientInSupabase = async (clientId, form) => {
  const { data, error } = await supabase
    .from('clients')
    .update({
      name: form.name.trim(),
      industry: form.industry || null,
      contact: form.contact || null,
      email: form.email || null,
      phone: form.phone || null,
      address: form.address || null,
      brands: form.brands || null,
    })
    .eq('id', clientId)
    .select(clientListSelect)
    .single();

  if (error) throw error;

  return mapClientFromSupabase(data);
};

export const archiveClientInSupabase = async (clientId) => {
  const { data, error } = await supabase
    .from('clients')
    .update({ is_archived: true })
    .eq('id', clientId)
    .select(clientListSelect)
    .single();

  if (error) throw error;

  return mapClientFromSupabase(data);
};

export const restoreClientInSupabase = async (clientId) => {
  const { data, error } = await supabase
    .from('clients')
    .update({ is_archived: false })
    .eq('id', clientId)
    .select(clientListSelect)
    .single();

  if (error) throw error;

  return mapClientFromSupabase(data);
};


export const deleteClientInSupabase = async (clientId) => {
  const { error } = await supabase
    .from('clients')
    .delete()
    .eq('id', clientId);

  if (error) throw error;
  return clientId;
};
