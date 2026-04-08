export const ROLE_LABELS = { admin: "Admin", planner: "Planner", buyer: "Buyer", finance: "Finance", viewer: "Viewer" };

export const normalizeRole = (role = "viewer") => String(role || "viewer").trim().toLowerCase();
export const formatRoleLabel = (role) => ROLE_LABELS[normalizeRole(role)] || "Viewer";

export const PERMISSIONS = {
  admin: { manageVendors: true, manageClients: true, manageCampaigns: true, manageRates: true, manageMpos: true, manageMpoStatus: true, manageWorkspace: true, manageMembers: true, manageDangerZone: true },
  planner: { manageVendors: false, manageClients: true, manageCampaigns: true, manageRates: false, manageMpos: true, manageMpoStatus: false, manageWorkspace: false, manageMembers: false, manageDangerZone: false },
  buyer: { manageVendors: true, manageClients: false, manageCampaigns: false, manageRates: true, manageMpos: true, manageMpoStatus: false, manageWorkspace: false, manageMembers: false, manageDangerZone: false },
  finance: { manageVendors: false, manageClients: false, manageCampaigns: false, manageRates: false, manageMpos: false, manageMpoStatus: true, manageWorkspace: false, manageMembers: false, manageDangerZone: false },
  viewer: { manageVendors: false, manageClients: false, manageCampaigns: false, manageRates: false, manageMpos: false, manageMpoStatus: false, manageWorkspace: false, manageMembers: false, manageDangerZone: false },
};

export const hasPermission = (user, key) => !!PERMISSIONS[normalizeRole(user?.role)]?.[key];
export const readOnlyMessage = (user) => `Your role (${formatRoleLabel(user?.role)}) is read-only for this action.`;

export const isAdmin = (user) => normalizeRole(user?.role) === "admin";
export const adminOnlyMessage = (user) => `Only admins can permanently delete records. Your role is ${formatRoleLabel(user?.role)}.`;
