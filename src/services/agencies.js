import { supabase } from "../lib/supabase";
import { setStoredAgencyContact } from "../utils/session";

const ROLE_LABELS = { admin: "Admin", planner: "Planner", buyer: "Buyer", finance: "Finance", viewer: "Viewer" };
const normalizeRole = (role = "viewer") => String(role || "viewer").trim().toLowerCase();
const formatRoleLabel = (role) => ROLE_LABELS[normalizeRole(role)] || "Viewer";

const normalizeAgencyName = (value = "") => value.trim().replace(/\s+/g, " ").toLowerCase();
export const normalizeAgencyCode = (value = "") => value.trim().toUpperCase().replace(/\s+/g, "");

export const findExistingAgencyByName = async (name = "") => {
  const normalized = normalizeAgencyName(name);
  if (!normalized) return null;
  try {
    const { data, error } = await supabase
      .from("agencies")
      .select("id, name, agency_code")
      .ilike("name", `%${name.trim()}%`)
      .limit(10);
    if (error) throw error;
    return (data || []).find((agency) => normalizeAgencyName(agency.name || "") === normalized) || null;
  } catch (error) {
    console.error("Failed to check for existing agency name:", error);
    return null;
  }
};

export const ensureAgencyForUser = async (authUser) => {
  if (!authUser) return null;

  const { data: profile, error: profileError } = await supabase
    .from("profiles")
    .select("agency_id")
    .eq("id", authUser.id)
    .single();

  if (profileError) throw profileError;
  if (profile?.agency_id) return profile.agency_id;

  const joinCode = normalizeAgencyCode(authUser.user_metadata?.agency_code || "");
  if (joinCode) {
    const { data, error } = await supabase.rpc("join_my_agency_by_code", { p_code: joinCode });
    if (error) {
      if ((error.message || "").toLowerCase().includes("join_my_agency_by_code")) {
        throw new Error('Missing database function "join_my_agency_by_code". Run the provided SQL patch in Supabase SQL Editor.');
      }
      throw error;
    }
    return data;
  }

  const agencyName = (authUser.user_metadata?.agency_name || "").trim();
  if (!agencyName) throw new Error("No agency invite code or agency name found for this user.");

  const existingAgency = await findExistingAgencyByName(agencyName);
  if (existingAgency) {
    if ((authUser.user_metadata?.agency_mode || "") === "create") {
      throw new Error("Agency already existing contact admin.");
    }
    return existingAgency.id;
  }

  const { data, error } = await supabase.rpc("create_my_agency_with_code", { p_name: agencyName });
  if (error) {
    if ((error.message || "").toLowerCase().includes("create_my_agency_with_code")) {
      throw new Error('Missing database function "create_my_agency_with_code". Run the provided SQL patch in Supabase SQL Editor.');
    }
    throw error;
  }

  if (Array.isArray(data)) return data[0]?.agency_id || null;
  if (data && typeof data === "object" && "agency_id" in data) return data.agency_id;
  return data;
};

export const updateAgencyInSupabase = async (agencyId, form) => {
  if (!agencyId) throw new Error("No agency found for this workspace.");

  const fallbackContact = { email: form.email || "", phone: form.phone || "" };
  const primary = await supabase
    .from("agencies")
    .update({
      name: form.agency?.trim() || "My Agency",
      address: form.address || null,
    })
    .eq("id", agencyId)
    .select("id, name, address, agency_code, created_by, created_at, updated_at")
    .maybeSingle();

  if (primary.error) throw primary.error;
  const data = primary.data;

  setStoredAgencyContact(agencyId, fallbackContact);
  try {
    await supabase.auth.updateUser({
      data: {
        agency_name: data?.name || form.agency || "My Agency",
        agency_code: data?.agency_code || "",
        agency_email: data?.email || fallbackContact.email || "",
        agency_phone: data?.phone || fallbackContact.phone || "",
      },
    });
  } catch (error) {
    console.error("Failed to update agency contact metadata:", error);
  }

  return {
    agency: data?.name || form.agency || "My Agency",
    agencyAddress: data?.address || "",
    agencyCode: data?.agency_code || "",
    agencyEmail: data?.email || fallbackContact.email || "",
    agencyPhone: data?.phone || fallbackContact.phone || "",
  };
};

export const fetchAgencyMembersFromSupabase = async (agencyId) => {
  if (!agencyId) return [];
  const { data, error } = await supabase
    .from("profiles")
    .select("id, email, full_name, title, phone, role, agency_id, created_at")
    .eq("agency_id", agencyId)
    .order("full_name", { ascending: true });
  if (error) throw error;
  return (data || []).map((member) => ({
    id: member.id,
    email: member.email || "",
    name: member.full_name || member.email || "User",
    title: member.title || "",
    phone: member.phone || "",
    role: normalizeRole(member.role || "viewer"),
    roleLabel: formatRoleLabel(member.role || "viewer"),
    createdAt: member.created_at ? new Date(member.created_at).getTime() : null,
  }));
};

export const updateAgencyMemberRoleInSupabase = async (targetUserId, nextRole) => {
  const role = normalizeRole(nextRole);
  const { error } = await supabase.rpc("admin_update_user_role", {
    p_target_user_id: targetUserId,
    p_new_role: role,
  });
  if (error) {
    if ((error.message || "").toLowerCase().includes("admin_update_user_role")) {
      throw new Error('Missing database function "admin_update_user_role". Run the provided SQL patch in Supabase SQL Editor.');
    }
    throw error;
  }
  return role;
};
