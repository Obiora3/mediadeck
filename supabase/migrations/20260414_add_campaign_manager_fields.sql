alter table public.campaigns
  add column if not exists campaign_manager_name text,
  add column if not exists campaign_manager_email text,
  add column if not exists campaign_manager_phone text;
