import { supabase } from '../lib/supabase';

export const fetchClientsFromSupabase = async () => {
  const { data, error } = await supabase
    .from('clients')
    .select('*')
    .order('created_at', { ascending: false });

  if (error) throw error;

  return (data || []).map((c) => ({
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
  }));
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
    .select()
    .single();

  if (error) throw error;

  return {
    id: data.id,
    name: data.name || '',
    industry: data.industry || '',
    contact: data.contact || '',
    email: data.email || '',
    phone: data.phone || '',
    address: data.address || '',
    brands: data.brands || '',
    archivedAt: null,
    createdAt: data.created_at ? new Date(data.created_at).getTime() : Date.now(),
    updatedAt: data.updated_at ? new Date(data.updated_at).getTime() : Date.now(),
  };
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
    .select()
    .single();

  if (error) throw error;

  return {
    id: data.id,
    name: data.name || '',
    industry: data.industry || '',
    contact: data.contact || '',
    email: data.email || '',
    phone: data.phone || '',
    address: data.address || '',
    brands: data.brands || '',
    archivedAt: data.is_archived ? (data.updated_at ? new Date(data.updated_at).getTime() : Date.now()) : null,
    createdAt: data.created_at ? new Date(data.created_at).getTime() : Date.now(),
    updatedAt: data.updated_at ? new Date(data.updated_at).getTime() : Date.now(),
  };
};

export const archiveClientInSupabase = async (clientId) => {
  const { data, error } = await supabase
    .from('clients')
    .update({ is_archived: true })
    .eq('id', clientId)
    .select()
    .single();

  if (error) throw error;

  return {
    id: data.id,
    name: data.name || '',
    industry: data.industry || '',
    contact: data.contact || '',
    email: data.email || '',
    phone: data.phone || '',
    address: data.address || '',
    brands: data.brands || '',
    archivedAt: data.updated_at ? new Date(data.updated_at).getTime() : Date.now(),
    createdAt: data.created_at ? new Date(data.created_at).getTime() : Date.now(),
    updatedAt: data.updated_at ? new Date(data.updated_at).getTime() : Date.now(),
  };
};

export const restoreClientInSupabase = async (clientId) => {
  const { data, error } = await supabase
    .from('clients')
    .update({ is_archived: false })
    .eq('id', clientId)
    .select()
    .single();

  if (error) throw error;

  return {
    id: data.id,
    name: data.name || '',
    industry: data.industry || '',
    contact: data.contact || '',
    email: data.email || '',
    phone: data.phone || '',
    address: data.address || '',
    brands: data.brands || '',
    archivedAt: null,
    createdAt: data.created_at ? new Date(data.created_at).getTime() : Date.now(),
    updatedAt: data.updated_at ? new Date(data.updated_at).getTime() : Date.now(),
  };
};
