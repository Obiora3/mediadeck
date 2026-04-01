import { supabase } from '../lib/supabase';

export const fetchCampaignsFromSupabase = async () => {
  const { data, error } = await supabase
    .from('campaigns')
    .select('*')
    .order('created_at', { ascending: false });

  if (error) throw error;

  return (data || []).map((c) => ({
    id: c.id,
    name: c.name || '',
    clientId: c.client_id || '',
    brand: c.brand || '',
    objective: c.objective || '',
    startDate: c.start_date || '',
    endDate: c.end_date || '',
    budget: c.budget ?? '',
    status: c.status || 'planning',
    medium: c.medium || '',
    notes: c.notes || '',
    materialList: Array.isArray(c.material_list) ? c.material_list : [],
    archivedAt: c.is_archived ? (c.updated_at ? new Date(c.updated_at).getTime() : Date.now()) : null,
    createdAt: c.created_at ? new Date(c.created_at).getTime() : Date.now(),
    updatedAt: c.updated_at ? new Date(c.updated_at).getTime() : Date.now(),
  }));
};

export const createCampaignInSupabase = async (agencyId, userId, form) => {
  if (!agencyId) throw new Error('No agency found for this user.');

  const { data, error } = await supabase
    .from('campaigns')
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
      status: form.status || 'planning',
      medium: form.medium || null,
      notes: form.notes || null,
      material_list: (form.materialList || []).map((x) => x.trim()).filter(Boolean),
      is_archived: false,
    }])
    .select()
    .single();

  if (error) throw error;

  return {
    id: data.id,
    name: data.name || '',
    clientId: data.client_id || '',
    brand: data.brand || '',
    objective: data.objective || '',
    startDate: data.start_date || '',
    endDate: data.end_date || '',
    budget: data.budget ?? '',
    status: data.status || 'planning',
    medium: data.medium || '',
    notes: data.notes || '',
    materialList: Array.isArray(data.material_list) ? data.material_list : [],
    archivedAt: null,
    createdAt: data.created_at ? new Date(data.created_at).getTime() : Date.now(),
    updatedAt: data.updated_at ? new Date(data.updated_at).getTime() : Date.now(),
  };
};

export const updateCampaignInSupabase = async (campaignId, form) => {
  const { data, error } = await supabase
    .from('campaigns')
    .update({
      client_id: form.clientId,
      name: form.name.trim(),
      brand: form.brand || null,
      objective: form.objective || null,
      start_date: form.startDate || null,
      end_date: form.endDate || null,
      budget: form.budget ? Number(form.budget) : 0,
      status: form.status || 'planning',
      medium: form.medium || null,
      notes: form.notes || null,
      material_list: (form.materialList || []).map((x) => x.trim()).filter(Boolean),
    })
    .eq('id', campaignId)
    .select()
    .single();

  if (error) throw error;

  return {
    id: data.id,
    name: data.name || '',
    clientId: data.client_id || '',
    brand: data.brand || '',
    objective: data.objective || '',
    startDate: data.start_date || '',
    endDate: data.end_date || '',
    budget: data.budget ?? '',
    status: data.status || 'planning',
    medium: data.medium || '',
    notes: data.notes || '',
    materialList: Array.isArray(data.material_list) ? data.material_list : [],
    archivedAt: data.is_archived ? (data.updated_at ? new Date(data.updated_at).getTime() : Date.now()) : null,
    createdAt: data.created_at ? new Date(data.created_at).getTime() : Date.now(),
    updatedAt: data.updated_at ? new Date(data.updated_at).getTime() : Date.now(),
  };
};

export const archiveCampaignInSupabase = async (campaignId) => {
  const { data, error } = await supabase
    .from('campaigns')
    .update({ is_archived: true })
    .eq('id', campaignId)
    .select()
    .single();

  if (error) throw error;

  return {
    id: data.id,
    name: data.name || '',
    clientId: data.client_id || '',
    brand: data.brand || '',
    objective: data.objective || '',
    startDate: data.start_date || '',
    endDate: data.end_date || '',
    budget: data.budget ?? '',
    status: data.status || 'planning',
    medium: data.medium || '',
    notes: data.notes || '',
    materialList: Array.isArray(data.material_list) ? data.material_list : [],
    archivedAt: data.updated_at ? new Date(data.updated_at).getTime() : Date.now(),
    createdAt: data.created_at ? new Date(data.created_at).getTime() : Date.now(),
    updatedAt: data.updated_at ? new Date(data.updated_at).getTime() : Date.now(),
  };
};

export const restoreCampaignInSupabase = async (campaignId) => {
  const { data, error } = await supabase
    .from('campaigns')
    .update({ is_archived: false })
    .eq('id', campaignId)
    .select()
    .single();

  if (error) throw error;

  return {
    id: data.id,
    name: data.name || '',
    clientId: data.client_id || '',
    brand: data.brand || '',
    objective: data.objective || '',
    startDate: data.start_date || '',
    endDate: data.end_date || '',
    budget: data.budget ?? '',
    status: data.status || 'planning',
    medium: data.medium || '',
    notes: data.notes || '',
    materialList: Array.isArray(data.material_list) ? data.material_list : [],
    archivedAt: null,
    createdAt: data.created_at ? new Date(data.created_at).getTime() : Date.now(),
    updatedAt: data.updated_at ? new Date(data.updated_at).getTime() : Date.now(),
  };
};
