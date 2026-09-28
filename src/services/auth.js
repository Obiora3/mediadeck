import { supabase } from "../lib/supabase";

export const changePasswordInSupabase = async (newPassword) => {
  const { error } = await supabase.auth.updateUser({ password: newPassword });
  if (error) throw error;
};

export const changePasswordWithCurrentPasswordInSupabase = async ({ email, currentPassword, newPassword }) => {
  if (!email) throw new Error("No account email is available for reauthentication.");

  const { error: signInError } = await supabase.auth.signInWithPassword({
    email,
    password: currentPassword,
  });

  if (signInError) {
    throw new Error("Current password is incorrect.");
  }

  await changePasswordInSupabase(newPassword);
};
