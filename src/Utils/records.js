export const isArchived = (item) => Boolean(item?.archivedAt);
export const activeOnly = (items = []) => items.filter(item => !isArchived(item));
export const archivedOnly = (items = []) => items.filter(isArchived);
export const archiveRecord = (item, user) => ({ ...item, archivedAt: Date.now(), archivedBy: user?.id || "system", updatedAt: Date.now() });
export const restoreRecord = (item) => ({ ...item, archivedAt: null, archivedBy: null, updatedAt: Date.now() });
export const pctWithin = (value) => {
  if (value === "" || value === null || value === undefined) return true;
  const n = parseFloat(value);
  return Number.isFinite(n) && n >= 0 && n <= 100;
};
