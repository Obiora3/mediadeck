import { supabase } from "../lib/supabase";
import { getStoredAgencyContact, getStoredUserSignature, setStoredUserSignature } from "../utils/session";

const ROLE_LABELS = { admin: "Admin", planner: "Planner", buyer: "Buyer", finance: "Finance", viewer: "Viewer" };
const normalizeRole = (role = "viewer") => String(role || "viewer").trim().toLowerCase();
const formatRoleLabel = (role) => ROLE_LABELS[normalizeRole(role)] || "Viewer";

export const loadAppUserFromSupabase = async (authUser) => {
  if (!authUser) return null;

  let profile = null;
  try {
    const { data, error } = await supabase
      .from("profiles")
      .select("*")
      .eq("id", authUser.id)
      .maybeSingle();
    if (!error) profile = data || null;
  } catch (error) {
    console.error("Profile load failed:", error);
  }

  let agencyName = "";
  let agencyCode = "";
  let agencyAddress = "";
  let agencyEmail = "";
  let agencyPhone = "";
  const agencyId = profile?.agency_id || null;

  if (profile?.agency_id) {
    try {
      const primary = await supabase
        .from("agencies")
        .select("id, name, agency_code, address, created_by, created_at, updated_at")
        .eq("id", profile.agency_id)
        .maybeSingle();
      const agency = primary.data || null;
      agencyName = agency?.name || "";
      agencyCode = agency?.agency_code || "";
      agencyAddress = agency?.address || "";
      agencyEmail = agency?.email || "";
      agencyPhone = agency?.phone || "";
    } catch (error) {
      console.error("Agency load failed:", error);
    }
  }

  const storedSignature = getStoredUserSignature(authUser.id);
  const storedAgencyContact = getStoredAgencyContact(agencyId);
  const resolvedRole = normalizeRole(profile?.role || authUser.user_metadata?.role || "viewer");

  return {
    id: authUser.id,
    name: profile?.full_name || authUser.user_metadata?.full_name || authUser.email || "User",
    email: authUser.email || "",
    title: profile?.title || authUser.user_metadata?.title || "",
    phone: profile?.phone || authUser.user_metadata?.phone || "",
    agency: agencyName || authUser.user_metadata?.agency_name || "My Agency",
    agencyId,
    agencyCode,
    agencyAddress,
    agencyEmail: agencyEmail || authUser.user_metadata?.agency_email || storedAgencyContact.email || "",
    agencyPhone: agencyPhone || authUser.user_metadata?.agency_phone || storedAgencyContact.phone || "",
    signatureDataUrl: authUser.user_metadata?.signature_data_url || storedSignature || "",
    role: resolvedRole,
    roleLabel: formatRoleLabel(resolvedRole),
    profileMissing: !profile,
  };
};

export const persistSignatureForUser = async (currentUser, signatureDataUrl = "") => {
  setStoredUserSignature(currentUser?.id, signatureDataUrl || "");
  try {
    const { error } = await supabase.auth.updateUser({ data: { signature_data_url: signatureDataUrl || "" } });
    if (error) throw error;
  } catch (error) {
    console.error("Failed to persist signature metadata:", error);
  }
  return signatureDataUrl || "";
};

export const updateProfileInSupabase = async (currentUser, form) => {
  const payload = {
    full_name: form.name?.trim() || currentUser.name || "",
    title: form.title || "",
    phone: form.phone || "",
  };
  const nextSignature = form.signatureDataUrl || "";
  setStoredUserSignature(currentUser?.id, nextSignature);
  const authPayload = {
    data: {
      full_name: payload.full_name,
      title: payload.title,
      phone: payload.phone,
      agency_name: currentUser.agency || "",
      agency_code: currentUser.agencyCode || "",
      agency_email: currentUser.agencyEmail || "",
      agency_phone: currentUser.agencyPhone || "",
      signature_data_url: nextSignature,
    },
  };
  if (form.email && form.email !== currentUser.email) authPayload.email = form.email;
  const { error: authError } = await supabase.auth.updateUser(authPayload);
  if (authError) throw authError;
  const { error: profileError } = await supabase
    .from("profiles")
    .update({ ...payload, email: form.email || currentUser.email || null })
    .eq("id", currentUser.id);
  if (profileError) throw profileError;
  return {
    ...currentUser,
    name: payload.full_name,
    title: payload.title,
    phone: payload.phone,
    email: form.email || currentUser.email || "",
    signatureDataUrl: nextSignature,
  };
};
