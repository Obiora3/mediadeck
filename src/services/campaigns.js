import { supabase } from '../lib/supabase';

const campaignListSelect = `
  id,
  agency_id,
  client_id,
  name,
  brand,
  objective,
  start_date,
  end_date,
  budget,
  status,
  medium,
  notes,
  material_list,
  is_archived,
  created_at,
  updated_at
`;

const mapCampaignFromSupabase = (c) => ({
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
});

export const fetchCampaignsFromSupabase = async (agencyId) => {
  if (!agencyId) return [];

  const { data, error } = await supabase
    .from('campaigns')
    .select(campaignListSelect)
    .eq('agency_id', agencyId)
    .order('created_at', { ascending: false });

  if (error) throw error;

  return (data || []).map(mapCampaignFromSupabase);
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
    .select(campaignListSelect)
    .single();

  if (error) throw error;

  return mapCampaignFromSupabase(data);
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
    .select(campaignListSelect)
    .single();

  if (error) throw error;

  return mapCampaignFromSupabase(data);
};

export const archiveCampaignInSupabase = async (campaignId) => {
  const { data, error } = await supabase
    .from('campaigns')
    .update({ is_archived: true })
    .eq('id', campaignId)
    .select(campaignListSelect)
    .single();

  if (error) throw error;

  return mapCampaignFromSupabase(data);
};

export const restoreCampaignInSupabase = async (campaignId) => {
  const { data, error } = await supabase
    .from('campaigns')
    .update({ is_archived: false })
    .eq('id', campaignId)
    .select(campaignListSelect)
    .single();

  if (error) throw error;

  return mapCampaignFromSupabase(data);
};
