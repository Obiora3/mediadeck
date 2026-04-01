import { supabase } from "../lib/supabase";
import { DEFAULT_APP_SETTINGS, mergeAppSettings } from "../constants/appDefaults";

export const fetchAppSettingsFromSupabase = async (agencyId) => {
  if (!agencyId) return mergeAppSettings();
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

export const saveAppSettingsToSupabase = async (agencyId, settings) => {
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
