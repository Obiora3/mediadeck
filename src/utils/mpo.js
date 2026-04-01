export const buildProgrammeCostLines = (spots = []) => {
  const grouped = new Map();

  (spots || []).forEach((spot) => {
    const programme = String(spot?.programme || '').trim() || 'Untitled Programme';
    const duration = String(spot?.duration || '').trim() || '';
    const rate = parseFloat(spot?.ratePerSpot) || 0;
    const cnt = Array.isArray(spot?.ad) && spot.ad.length
      ? spot.ad.length
      : Array.isArray(spot?.calendarDays) && spot.calendarDays.length
        ? spot.calendarDays.length
        : (parseFloat(spot?.spots) || 0);

    const key = programme.toLowerCase();
    if (!grouped.has(key)) grouped.set(key, { programme, duration, cnt: 0, rate });

    const entry = grouped.get(key);
    entry.cnt += cnt;
    if (!entry.rate && rate) entry.rate = rate;
    if (!entry.duration && duration) entry.duration = duration;
  });

  return Array.from(grouped.values()).map((line) => ({
    ...line,
    gross: line.cnt * line.rate,
  }));
};
