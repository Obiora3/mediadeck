import { supabase } from "../lib/supabase";

export const changePasswordInSupabase = async (newPassword) => {
  const { error } = await supabase.auth.updateUser({ password: newPassword });
  if (error) throw error;
};
