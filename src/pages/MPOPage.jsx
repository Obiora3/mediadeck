import { useEffect, useRef, useState } from "react";
import Badge from "../components/Badge";
import Empty from "../components/Empty";
import Toast from "../components/Toast";
import Modal from "../components/Modal";
import Confirm from "../components/Confirm";
import { Btn, Field, AttachmentField, Card, Stat } from "../components/ui/primitives";
import PrintPreview from "../components/mpo/PrintPreview";
import { buildCSV, sanitizeMPOForExport, buildMPOHTML } from "../utils/export";
import { activeOnly, archivedOnly, isArchived } from "../utils/records";
import { fmtN } from "../utils/formatters";
import { formatRoleLabel, hasPermission, readOnlyMessage } from "../constants/roles";
import {
  MPO_STATUS_OPTIONS,
  MPO_STATUS_LABELS,
  MPO_EXECUTION_STATUS_OPTIONS,
  MPO_INVOICE_STATUS_OPTIONS,
  MPO_PROOF_STATUS_OPTIONS,
  MPO_PAYMENT_STATUS_OPTIONS,
  MPO_RECON_STATUS_OPTIONS,
  toIsoInput,
  toIsoOrNull,
  getExecutionHealthColor,
  getExecutionHealthLabel,
  getAllowedMpoStatusTargets,
  mpoStatusNeedsNote,
  getMpoWorkflowMeta,
  isMpoAwaitingUser,
  getWorkflowActionLabel,
  getWorkflowActionVariant,
  getQuickWorkflowActions,
  canEditMpoContent,
  MPO_STATUS_COLORS,
} from "../constants/mpoWorkflow";
import { DEFAULT_APP_SETTINGS } from "../constants/appDefaults";
import {
  uploadMpoAttachmentAndGetUrl,
  createMpoInSupabase,
  updateMpoInSupabase,
  archiveMpoInSupabase,
  restoreMpoInSupabase,
  updateMpoStatusInSupabase,
  updateMpoExecutionInSupabase,
  generateNextMpoNoFromSupabase,
  fetchMappedMpoById,
  fetchMappedMpoByAgencyAndNo,
} from "../services/mpos";
import {
  createAuditEventInSupabase,
  fetchAuditEventsForRecord,
  notifyMpoWorkflowTransition,
  notifyExecutionUpdate,
} from "../services/notifications";
import { ensureVendorExistsInSupabase } from "../services/vendors";

const store = {
  get: (k, d = null) => { try { const v = localStorage.getItem(k); return v ? JSON.parse(v) : d; } catch { return d; } },
  set: (k, v) => { try { localStorage.setItem(k, JSON.stringify(v)); } catch {} },
  del: (k) => { try { localStorage.removeItem(k); } catch {} },
};
const uid = () => Math.random().toString(36).slice(2) + Date.now().toString(36);
const collapseDaysToCounts = (days = []) =>
  (days || []).reduce((acc, day) => {
    const key = Number(day);
    if (!key) return acc;
    acc[key] = (acc[key] || 0) + 1;
    return acc;
  }, {});

const expandCountsToDays = (dayCounts = {}) =>
  Object.entries(dayCounts || {})
    .flatMap(([day, count]) =>
      Array.from({ length: Math.max(0, Number(count) || 0) }, () => Number(day))
    )
    .sort((a, b) => a - b);

const totalCountFromDayCounts = (dayCounts = {}) =>
  Object.values(dayCounts || {}).reduce((sum, count) => sum + (Number(count) || 0), 0);

const isComplimentarySpot = (spot = {}) =>
  Boolean(spot?.isComplimentary || Number(spot?.ratePerSpot || 0) <= 0);

const getSpotBonusCount = (spot = {}) => {
  if (isComplimentarySpot(spot)) return Number(spot?.spots) || 0;
  const raw = Number(spot?.bonusSpots) || 0;
  const total = Array.isArray(spot?.calendarDays) && spot.calendarDays.length
    ? spot.calendarDays.length
    : Array.isArray(spot?.ad) && spot.ad.length
      ? spot.ad.length
      : (parseFloat(spot?.spots) || 0);
  return Math.max(0, Math.min(raw, total));
};

const getPaidSpotCount = (spot = {}) => {
  const total = Array.isArray(spot?.calendarDays) && spot.calendarDays.length
    ? spot.calendarDays.length
    : Array.isArray(spot?.ad) && spot.ad.length
      ? spot.ad.length
      : (parseFloat(spot?.spots) || 0);
  return Math.max(0, total - getSpotBonusCount(spot));
};

const normalizeBonusAdjustedSpot = (baseSpot = {}, totalSpotsInput = 0, bonusSpotsInput = 0) => {
  const totalSpots = Math.max(0, Number(totalSpotsInput) || 0);
  const requestedBonus = baseSpot?.isComplimentary
    ? totalSpots
    : Math.max(0, Math.min(Number(bonusSpotsInput) || 0, totalSpots));
  const paidSpots = Math.max(0, totalSpots - requestedBonus);
  const calendarDays = Array.isArray(baseSpot?.calendarDays) ? [...baseSpot.calendarDays] : [];
  return {
    ...baseSpot,
    id: baseSpot?.id || uid(),
    spots: String(totalSpots),
    bonusSpots: requestedBonus,
    paidSpots,
    ratePerSpot: baseSpot?.isComplimentary ? 0 : (Number(baseSpot?.ratePerSpot) || 0),
    customRate: baseSpot?.isComplimentary ? "" : (baseSpot?.customRate || ""),
    calendarDays,
    calendarDayCounts: collapseDaysToCounts(calendarDays),
  };
};

const CALENDAR_MONTHS = ["January","February","March","April","May","June","July","August","September","October","November","December"];
const getMonthIndexFromLabel = (monthLabel = "") => {
  const normalized = String(monthLabel || "").trim().toLowerCase();
  const direct = CALENDAR_MONTHS.findIndex(month => month.toLowerCase() === normalized);
  if (direct >= 0) return direct;
  const short = CALENDAR_MONTHS.findIndex(month => month.slice(0, 3).toLowerCase() === normalized.slice(0, 3));
  return short >= 0 ? short : new Date().getMonth();
};
const getDaysInMonth = (monthLabel = "", yearValue = "") => {
  const monthIndex = getMonthIndexFromLabel(monthLabel);
  const year = Number(yearValue) || new Date().getFullYear();
  return new Date(year, monthIndex + 1, 0).getDate();
};
const getCalendarWeekdayLabel = (monthLabel = "", yearValue = "", dayNumber = 1) => {
  const monthIndex = getMonthIndexFromLabel(monthLabel);
  const year = Number(yearValue) || new Date().getFullYear();
  const date = new Date(year, monthIndex, Number(dayNumber) || 1);
  return date.toLocaleDateString("en-NG", { weekday: "short" });
};
const getCalendarDayOfWeek = (monthLabel = "", yearValue = "", dayNumber = 1) => {
  const monthIndex = getMonthIndexFromLabel(monthLabel);
  const year = Number(yearValue) || new Date().getFullYear();
  return new Date(year, monthIndex, Number(dayNumber) || 1).getDay();
};
const roundMoneyValue = (value, settings = {}) => {
  const num = Number(value) || 0;
  return settings?.roundToWholeNaira ? Math.round(num) : Math.round(num * 100) / 100;
};
const formatAuditTimestamp = (value) => {
  if (!value) return "—";
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? String(value) : date.toLocaleString();
};

const normalizePdfText = (value) => String(value ?? "")
  .replace(/[\u2018\u2019]/g, "'")
  .replace(/[\u201C\u201D]/g, '"')
  .replace(/[\u2013\u2014]/g, '-')
  .replace(/\u2026/g, '...')
  .replace(/[^\x20-\x7E\n]/g, ' ');

const escapePdfText = (value) => normalizePdfText(value)
  .replace(/\\/g, '\\\\')
  .replace(/\(/g, '\\(')
  .replace(/\)/g, '\\)');

const wrapPdfText = (value, maxChars = 90) => {
  const text = normalizePdfText(value).replace(/\s+/g, ' ').trim();
  if (!text) return [''];
  const words = text.split(' ');
  const lines = [];
  let current = '';
  words.forEach((word) => {
    const next = current ? `${current} ${word}` : word;
    if (next.length <= maxChars) {
      current = next;
    } else {
      if (current) lines.push(current);
      current = word.length > maxChars ? `${word.slice(0, maxChars - 1)}-` : word;
    }
  });
  if (current) lines.push(current);
  return lines.length ? lines : [''];
};

const truncatePdfCell = (value, maxChars) => {
  const text = normalizePdfText(value).replace(/\s+/g, ' ').trim();
  if (text.length <= maxChars) return text;
  return `${text.slice(0, Math.max(0, maxChars - 3)).trim()}...`;
};


const InlineDailyCalendar = ({ month, year, calRows = [], setCalRows, vendorRates = [], fmtN, blankCalRow, campaignMaterials = [], onAdd }) => {
  const dayTotal = (dayCounts = {}) => Object.values(dayCounts || {}).reduce((sum, value) => sum + (Number(value) || 0), 0);
  const daysInMonth = getDaysInMonth(month, year);

  const rateOptions = (vendorRates || []).map((rate) => ({
    value: rate.id,
    label: `${rate.programme || "Untitled"}${rate.timeBelt ? ` · ${rate.timeBelt}` : ""}${rate.duration ? ` · ${rate.duration}"` : ""}`,
  }));

  const updateRow = (rowId, updater) => {
    setCalRows((rows) =>
      (rows || []).map((row) => {
        if (row.id !== rowId) return row;
        const next = typeof updater === "function" ? updater(row) : { ...row, ...updater };
        const autoMatched = applyAutoMatchedVendorRate(next, vendorRates || []);
        const nextTotal = dayTotal(autoMatched.dayCounts || {});
        const nextBonus = autoMatched.isComplimentary
          ? nextTotal
          : Math.max(0, Math.min(Number(autoMatched.bonusSpots) || 0, nextTotal));
        return {
          ...autoMatched,
          bonusSpots: nextTotal > 0 ? String(nextBonus) : "",
          customRate: autoMatched.isComplimentary ? "" : autoMatched.customRate,
        };
      })
    );
  };

  const removeRow = (rowId) => {
    setCalRows((rows) => {
      const next = (rows || []).filter((row) => row.id !== rowId);
      return next.length ? next : [blankCalRow()];
    });
  };

  const setDayCount = (rowId, day, nextCount) => {
    updateRow(rowId, (row) => {
      const dayCounts = { ...(row.dayCounts || {}) };
      const normalized = Math.max(0, Number(nextCount) || 0);
      if (normalized <= 0) delete dayCounts[day];
      else dayCounts[day] = normalized;
      return { ...row, dayCounts };
    });
  };

  const toggleDay = (rowId, day) => {
    updateRow(rowId, (row) => {
      const dayCounts = { ...(row.dayCounts || {}) };
      if (dayCounts[day]) delete dayCounts[day];
      else dayCounts[day] = 1;
      return { ...row, dayCounts };
    });
  };

  const applyWeekdaySelection = (rowId, weekdayValues = [], mode = "replace") => {
    updateRow(rowId, (row) => {
      const nextCounts = mode === "append" ? { ...(row.dayCounts || {}) } : {};
      for (let day = 1; day <= daysInMonth; day += 1) {
        const dayOfWeek = getCalendarDayOfWeek(month, year, day);
        if (weekdayValues.includes(dayOfWeek)) {
          nextCounts[day] = Math.max(1, Number(nextCounts[day]) || 1);
        } else if (mode === "replace") {
          delete nextCounts[day];
        }
      }
      return { ...row, dayCounts: nextCounts };
    });
  };

  const selectProgramme = (rowId, rateId) => {
    const rate = (vendorRates || []).find((item) => item.id === rateId);
    updateRow(rowId, (row) => ({
      ...row,
      rateId: rateId || "",
      programme: rate?.programme || row.programme || "",
      timeBelt: rate?.timeBelt || row.timeBelt || "",
      duration: rate?.duration || row.duration || "30",
      customRate: row.isComplimentary ? "" : (rate?.ratePerSpot ? String(rate.ratePerSpot) : row.customRate || ""),
    }));
  };

  const inputStyle = {
    background: "var(--bg3)",
    border: "1px solid var(--border2)",
    borderRadius: 8,
    padding: "8px 11px",
    color: "var(--text)",
    fontSize: 13,
    outline: "none",
    width: "100%",
  };

  const weekdayQuickOptions = [
    { label: "Weekdays", days: [1,2,3,4,5] },
    { label: "Weekends", days: [0,6] },
    { label: "Mon", days: [1] },
    { label: "Tue", days: [2] },
    { label: "Wed", days: [3] },
    { label: "Thu", days: [4] },
    { label: "Fri", days: [5] },
    { label: "Sat", days: [6] },
    { label: "Sun", days: [0] },
  ];

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
      {(calRows || []).map((row, index) => {
        const totalScheduledSpots = dayTotal(row.dayCounts);
        const bonusValue = row.isComplimentary
          ? totalScheduledSpots
          : Math.max(0, Math.min(Number(row.bonusSpots) || 0, totalScheduledSpots));
        const paidValue = Math.max(0, totalScheduledSpots - bonusValue);

        return (
          <div key={row.id} style={{ border: "1px solid var(--border)", borderRadius: 12, background: "var(--bg2)", padding: 14 }}>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, flexWrap: "wrap", marginBottom: 12 }}>
              <div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 12 }}>
                Schedule Row {index + 1}
              </div>
              <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
                {totalScheduledSpots > 0 ? <Badge color="blue">{totalScheduledSpots} spot{totalScheduledSpots !== 1 ? "s" : ""}</Badge> : null}
                {bonusValue > 0 ? <Badge color="purple">Bonus {bonusValue}</Badge> : null}
                {(calRows || []).length > 1 ? <Btn variant="danger" size="sm" onClick={() => removeRow(row.id)}>Remove</Btn> : null}
              </div>
            </div>

            <div style={{ display: "grid", gridTemplateColumns: "repeat(2,minmax(0,1fr))", gap: 12 }}>
              {rateOptions.length > 0 ? (
                <div style={{ width: "100%", overflowX: "auto" }}>
                  <label style={{ fontSize: 11, fontWeight: 600, color: "var(--text3)", textTransform: "uppercase", letterSpacing: ".08em", marginBottom: 5, display: "block" }}>Select Programme</label>
                  <select
                    value={row.rateId || ""}
                    onChange={(e) => selectProgramme(row.id, e.target.value)}
                    style={{ ...inputStyle, cursor: "pointer" }}
                  >
                    <option value="">— Select programme from rate card —</option>
                    {rateOptions.map((option) => <option key={option.value} value={option.value}>{option.label}</option>)}
                  </select>
                </div>
              ) : <div />}
              <div>
                <label style={{ fontSize: 11, fontWeight: 600, color: "var(--text3)", textTransform: "uppercase", letterSpacing: ".08em", marginBottom: 5, display: "block" }}>Programme</label>
                <input
                  value={row.programme || ""}
                  onChange={(e) => updateRow(row.id, { programme: e.target.value })}
                  placeholder="NTA News, SuperStory…"
                  style={inputStyle}
                />
              </div>

              <div>
                <label style={{ fontSize: 11, fontWeight: 600, color: "var(--text3)", textTransform: "uppercase", letterSpacing: ".08em", marginBottom: 5, display: "block" }}>Weekday (WD)</label>
                <input value={row.wd || ""} onChange={(e) => updateRow(row.id, { wd: e.target.value.toUpperCase() })} placeholder="MON / TUE / DAILY" style={inputStyle} />
              </div>
              <div>
                <label style={{ fontSize: 11, fontWeight: 600, color: "var(--text3)", textTransform: "uppercase", letterSpacing: ".08em", marginBottom: 5, display: "block" }}>Fixed Time / Time Belt</label>
                <input value={row.timeBelt || ""} onChange={(e) => updateRow(row.id, { timeBelt: e.target.value })} placeholder="08:45–09:00" style={inputStyle} />
              </div>

              <div>
                <label style={{ fontSize: 11, fontWeight: 600, color: "var(--text3)", textTransform: "uppercase", letterSpacing: ".08em", marginBottom: 5, display: "block" }}>Duration (secs)</label>
                <input value={row.duration || "30"} onChange={(e) => updateRow(row.id, { duration: e.target.value })} placeholder="30" type="number" style={inputStyle} />
              </div>
              <div>
                <label style={{ fontSize: 11, fontWeight: 600, color: "var(--text3)", textTransform: "uppercase", letterSpacing: ".08em", marginBottom: 5, display: "block" }}>Rate per Spot (₦)</label>
                <input
                  type="number"
                  value={row.isComplimentary ? "0" : (row.customRate || ((vendorRates || []).find((item) => item.id === row.rateId)?.ratePerSpot || ""))}
                  onChange={(e) => updateRow(row.id, { customRate: e.target.value })}
                  placeholder="0"
                  style={inputStyle}
                />
              </div>

              <div style={{ gridColumn: "1 / -1" }}>
                {campaignMaterials.length > 0 ? (
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 180px", gap: 10 }}>
                    <div>
                      <label style={{ fontSize: 11, fontWeight: 600, color: "var(--text3)", textTransform: "uppercase", letterSpacing: ".08em", marginBottom: 5, display: "block" }}>Material / Spot Name</label>
                      <select value={row.material || ""} onChange={(e) => updateRow(row.id, { material: e.target.value })} style={{ ...inputStyle, cursor: "pointer" }}>
                        <option value="">— Select material —</option>
                        {campaignMaterials.map((material, i) => <option key={i} value={material}>{material}</option>)}
                        <option value="__custom__">Custom…</option>
                      </select>
                    </div>
                    {row.material === "__custom__" ? (
                      <div>
                        <label style={{ fontSize: 11, fontWeight: 600, color: "var(--text3)", textTransform: "uppercase", letterSpacing: ".08em", marginBottom: 5, display: "block" }}>Custom Material</label>
                        <input value={row.materialCustom || ""} onChange={(e) => updateRow(row.id, { materialCustom: e.target.value })} placeholder="Type material name" style={inputStyle} />
                      </div>
                    ) : <div />}
                  </div>
                ) : (
                  <div>
                    <label style={{ fontSize: 11, fontWeight: 600, color: "var(--text3)", textTransform: "uppercase", letterSpacing: ".08em", marginBottom: 5, display: "block" }}>Material / Spot Name</label>
                    <input value={row.material || ""} onChange={(e) => updateRow(row.id, { material: e.target.value })} placeholder="SM Thematic English 30secs (MP4)" style={inputStyle} />
                  </div>
                )}
              </div>

              <div style={{ gridColumn: "1 / -1", display: "grid", gridTemplateColumns: "repeat(3,minmax(0,1fr))", gap: 10, background: "var(--bg3)", border: "1px solid var(--border2)", borderRadius: 10, padding: "10px 12px" }}>
                <div>
                  <div style={{ fontSize: 10, fontWeight: 700, color: "var(--text3)", textTransform: "uppercase", marginBottom: 4 }}>Total Scheduled</div>
                  <div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 800, fontSize: 18, color: "var(--accent)" }}>{totalScheduledSpots}</div>
                </div>
                <div>
                  <div style={{ fontSize: 10, fontWeight: 700, color: "var(--text3)", textTransform: "uppercase", marginBottom: 4 }}>Bonus Spots</div>
                  <input
                    type="number"
                    min="0"
                    max={totalScheduledSpots}
                    value={row.isComplimentary ? String(totalScheduledSpots) : (row.bonusSpots || "")}
                    onChange={(e) => updateRow(row.id, { bonusSpots: e.target.value })}
                    disabled={!!row.isComplimentary}
                    style={{ width: "100%", background: "var(--bg2)", border: "1px solid var(--border2)", borderRadius: 8, padding: "8px 10px", color: "var(--text)", fontSize: 13, outline: "none", opacity: row.isComplimentary ? 0.65 : 1 }}
                  />
                </div>
                <div>
                  <div style={{ fontSize: 10, fontWeight: 700, color: "var(--text3)", textTransform: "uppercase", marginBottom: 4 }}>Paid Spots</div>
                  <div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 800, fontSize: 18, color: "var(--green)" }}>{paidValue}</div>
                </div>
              </div>

              <div style={{ gridColumn: "1 / -1", display: "flex", alignItems: "center", gap: 10, padding: "10px 12px", background: "var(--bg3)", border: "1px solid var(--border2)", borderRadius: 8 }}>
                <input
                  id={`cal-complimentary-${row.id}`}
                  type="checkbox"
                  checked={!!row.isComplimentary}
                  onChange={(e) => updateRow(row.id, (current) => ({
                    isComplimentary: e.target.checked,
                    customRate: e.target.checked ? "" : current.customRate,
                    bonusSpots: e.target.checked ? String(dayTotal(current.dayCounts || {})) : current.bonusSpots,
                  }))}
                  style={{ accentColor: "var(--accent)", width: 16, height: 16 }}
                />
                <label htmlFor={`cal-complimentary-${row.id}`} style={{ fontSize: 13, color: "var(--text2)", cursor: "pointer" }}>
                  Mark this scheduled row as complimentary / bonus spots
                </label>
              </div>
            </div>

            <div style={{ marginTop: 14 }}>
              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, marginBottom: 10, flexWrap: "wrap" }}>
                <div style={{ fontSize: 12, fontWeight: 600, color: "var(--text2)", textTransform: "uppercase", letterSpacing: ".06em" }}>
                  Select Airing Dates
                </div>
                <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                  {weekdayQuickOptions.map((option) => (
                    <button
                      key={`${row.id}-${option.label}`}
                      type="button"
                      onClick={() => applyWeekdaySelection(row.id, option.days)}
                      style={{ padding: "4px 10px", fontSize: 11, background: "var(--bg3)", border: "1px solid var(--border2)", borderRadius: 999, color: "var(--text2)", cursor: "pointer", fontWeight: 700 }}
                    >
                      {option.label}
                    </button>
                  ))}
                  <button onClick={() => updateRow(row.id, { dayCounts: Object.fromEntries(Array.from({ length: daysInMonth }, (_, i) => [i + 1, 1])) })} style={{ padding: "4px 10px", fontSize: 11, background: "var(--bg3)", border: "1px solid var(--border2)", borderRadius: 999, color: "var(--text2)", cursor: "pointer", fontWeight: 700 }}>All</button>
                  <button onClick={() => updateRow(row.id, { dayCounts: {} })} style={{ padding: "4px 10px", fontSize: 11, background: "var(--bg3)", border: "1px solid var(--border2)", borderRadius: 999, color: "var(--text2)", cursor: "pointer", fontWeight: 700 }}>Clear</button>
                </div>
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(7, minmax(0, 1fr))", gap: 6 }}>
                {Array.from({ length: daysInMonth }, (_, i) => i + 1).map((day) => {
                  const count = Number(row.dayCounts?.[day] || 0);
                  const selected = count > 0;
                  const weekdayLabel = getCalendarWeekdayLabel(month, year, day);
                  return (
                    <div
                      key={day}
                      style={{
                        border: `1px solid ${selected ? "var(--accent)" : "var(--border2)"}`,
                        borderRadius: 8,
                        background: selected ? "rgba(240,165,0,.12)" : "var(--bg3)",
                        padding: 6,
                        minHeight: 80,
                        display: "flex",
                        flexDirection: "column",
                        justifyContent: "space-between",
                        gap: 6,
                      }}
                    >
                      <button
                        type="button"
                        onClick={() => toggleDay(row.id, day)}
                        style={{
                          border: "none",
                          background: "transparent",
                          color: selected ? "var(--accent)" : "var(--text3)",
                          cursor: "pointer",
                          textAlign: "left",
                          padding: 0,
                          display: "flex",
                          flexDirection: "column",
                          alignItems: "flex-start",
                          gap: 2,
                        }}
                        title={`Select ${weekdayLabel} ${day} only`}
                      >
                        <span style={{ fontSize: 10, fontWeight: 800, letterSpacing: ".06em", textTransform: "uppercase" }}>
                          {weekdayLabel}
                        </span>
                        <span style={{ fontFamily: "'Syne',sans-serif", fontWeight: 800, fontSize: 20, lineHeight: 1 }}>
                          {day}
                        </span>
                      </button>
                      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 4 }}>
                        <button
                          type="button"
                          onClick={() => setDayCount(row.id, day, count - 1)}
                          disabled={!selected}
                          style={{
                            width: 24,
                            height: 24,
                            borderRadius: 6,
                            border: "1px solid var(--border2)",
                            background: selected ? "var(--bg2)" : "var(--bg3)",
                            color: selected ? "var(--text)" : "var(--text3)",
                            cursor: selected ? "pointer" : "not-allowed",
                            opacity: selected ? 1 : 0.55,
                            fontWeight: 700,
                          }}
                        >
                          −
                        </button>
                        <button
                          type="button"
                          onClick={() => toggleDay(row.id, day)}
                          style={{
                            minWidth: 32,
                            border: "none",
                            background: "transparent",
                            color: selected ? "var(--accent)" : "var(--text3)",
                            fontFamily: "'Syne',sans-serif",
                            fontWeight: 800,
                            fontSize: 14,
                            cursor: "pointer",
                            textAlign: "center",
                            padding: 0,
                          }}
                        >
                          {count || 0}
                        </button>
                        <button
                          type="button"
                          onClick={() => setDayCount(row.id, day, count + 1 || 1)}
                          style={{
                            width: 24,
                            height: 24,
                            borderRadius: 6,
                            border: "1px solid var(--border2)",
                            background: "var(--bg2)",
                            color: "var(--text)",
                            cursor: "pointer",
                            fontWeight: 700,
                          }}
                        >
                          +
                        </button>
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
          </div>
        );
      })}

      <div style={{ display: "flex", justifyContent: "space-between", gap: 10, flexWrap: "wrap" }}>
        <Btn variant="ghost" size="sm" onClick={() => setCalRows((rows) => [...(rows || []), blankCalRow()])}>+ Add Another Schedule Row</Btn>
        <Btn variant="secondary" size="sm" onClick={onAdd}>Add to Schedule</Btn>
      </div>
    </div>
  );
};


const buildMpoPdfBytes = (mpo) => {
  const pageWidth = 595.28;
  const pageHeight = 841.89;
  const marginX = 42;
  const topY = 804;
  const bottomY = 50;
  const lineHeight = 14;
  const pages = [];
  let commands = [];
  let cursorY = topY;

  const pushPage = () => {
    if (commands.length) pages.push(commands.join('\n'));
    commands = [];
    cursorY = topY;
  };

  const ensureSpace = (linesNeeded = 1) => {
    if (cursorY - (linesNeeded * lineHeight) < bottomY) pushPage();
  };

  const addTextLine = (text, options = {}) => {
    const size = options.size || 10;
    const x = options.x ?? marginX;
    const y = options.y ?? cursorY;
    commands.push(`BT /F1 ${size} Tf 1 0 0 1 ${x.toFixed(2)} ${y.toFixed(2)} Tm (${escapePdfText(text)}) Tj ET`);
    if (!options.absolute) cursorY = y - (options.lineHeight || lineHeight);
  };

  const addWrappedBlock = (text, options = {}) => {
    const lines = wrapPdfText(text, options.maxChars || 90);
    ensureSpace(lines.length);
    lines.forEach((line) => addTextLine(line, options));
  };

  const addSpacer = (amount = 8) => {
    cursorY -= amount;
  };

  const addSectionTitle = (text) => {
    ensureSpace(2);
    addTextLine(text.toUpperCase(), { size: 11 });
    addSpacer(2);
  };

  const addLabelValue = (label, value) => {
    const line = `${label}: ${value || '-'}`;
    addWrappedBlock(line, { size: 10, maxChars: 92 });
  };

  const addTableHeader = () => {
    ensureSpace(2);
    addTextLine('Programme                      Time Belt            Material                  Dur   Spots   Rate        Gross', { size: 9 });
    addTextLine('----------------------------------------------------------------------------------------------------------', { size: 9, lineHeight: 12 });
  };

  const currency = (value) => {
    const num = Number(value) || 0;
    try {
      return `N${num.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
    } catch (_error) {
      return `N${num.toFixed(2)}`;
    }
  };

  pushPage();
  addTextLine('MEDIA PURCHASE ORDER (MPO)', { size: 18 });
  addSpacer(4);
  addLabelValue('MPO Number', mpo.mpoNo || 'Draft');
  addLabelValue('Date', mpo.date || '-');
  addLabelValue('Period', `${(mpo.months || []).length ? mpo.months.join(', ') : (mpo.month || '-')} ${mpo.year || ''}`.trim());
  addLabelValue('Client', mpo.clientName || '-');
  addLabelValue('Brand', mpo.brand || '-');
  addLabelValue('Campaign', mpo.campaignName || '-');
  addLabelValue('Vendor', mpo.vendorName || '-');
  addLabelValue('Medium', mpo.medium || '-');
  addLabelValue('Transmit Instruction', mpo.transmitMsg || '-');

  addSpacer(6);
  addSectionTitle('Spot Schedule');
  addTableHeader();

  (mpo.spots || []).forEach((spot) => {
    ensureSpace(2);
    const row = [
      truncatePdfCell(spot.programme || '-', 28).padEnd(28, ' '),
      truncatePdfCell(spot.timeBelt || spot.wd || '-', 18).padEnd(18, ' '),
      truncatePdfCell(spot.material || '-', 24).padEnd(24, ' '),
      String(spot.duration || '-').padStart(4, ' '),
      String(spot.spots || '0').padStart(6, ' '),
      truncatePdfCell(currency(spot.ratePerSpot || 0), 11).padStart(11, ' '),
      truncatePdfCell(currency((Number(spot.spots) || 0) * (Number(spot.ratePerSpot) || 0)), 12).padStart(12, ' '),
    ].join(' ');
    addTextLine(row, { size: 9, lineHeight: 12 });
    if (spot.scheduleMonth || (spot.calendarDays || []).length) {
      const metaParts = [];
      if (spot.scheduleMonth) metaParts.push(`Month: ${spot.scheduleMonth}`);
      if ((spot.calendarDays || []).length) metaParts.push(`Days: ${(spot.calendarDays || []).join(', ')}`);
      addWrappedBlock(`   ${metaParts.join(' | ')}`, { size: 8, maxChars: 110 });
    }
  });

  addSpacer(6);
  addSectionTitle('Cost Summary');
  [
    ['Total Spots', mpo.totalSpots],
    ['Total Gross Value', currency(mpo.totalGross)],
    ['Discount', currency(mpo.discAmt)],
    ['Net After Discount', currency(mpo.lessDisc)],
    ['Agency Commission', currency(mpo.commAmt)],
    ['After Commission', currency(mpo.afterComm)],
    [mpo.surchLabel || 'Surcharge', currency(mpo.surchAmt || 0)],
    ['Net Value', currency(mpo.netVal)],
    [`VAT (${mpo.vatPct || 7.5}%)`, currency(mpo.vatAmt)],
    ['Total Amount Payable', currency(mpo.grandTotal)],
  ].forEach(([label, value]) => addLabelValue(label, value));

  addSpacer(6);
  addSectionTitle('Prepared By / Signatories');
  addLabelValue('Prepared By', mpo.preparedBy || '-');
  addLabelValue('Prepared Title', mpo.preparedTitle || '-');
  addLabelValue('Prepared Contact', mpo.preparedContact || '-');
  addLabelValue('Signed By', mpo.signedBy || '-');
  addLabelValue('Signed Title', mpo.signedTitle || '-');
  addLabelValue('Agency Address', mpo.agencyAddress || '-');
  addLabelValue('Agency Email', mpo.agencyEmail || '-');
  addLabelValue('Agency Phone', mpo.agencyPhone || '-');

  addSpacer(10);
  addWrappedBlock('This PDF was generated directly as a document file from the MPO data.', { size: 8, maxChars: 100 });

  if (commands.length) pages.push(commands.join('\n'));

  const objects = [];
  const addObject = (body) => {
    objects.push(body);
    return objects.length;
  };

  const fontId = addObject('<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>');
  const pageIds = [];
  const contentIds = [];
  pages.forEach((pageContent) => {
    const stream = `${pageContent}\nBT /F1 8 Tf 1 0 0 1 ${marginX.toFixed(2)} 24 Tm (Page ${pageIds.length + 1} of ${pages.length}) Tj ET`;
    const contentId = addObject(`<< /Length ${stream.length} >>\nstream\n${stream}\nendstream`);
    contentIds.push(contentId);
    const pageId = addObject(`<< /Type /Page /Parent 0 0 R /MediaBox [0 0 ${pageWidth} ${pageHeight}] /Resources << /Font << /F1 ${fontId} 0 R >> >> /Contents ${contentId} 0 R >>`);
    pageIds.push(pageId);
  });

  const pagesId = addObject(`<< /Type /Pages /Kids [${pageIds.map((id) => `${id} 0 R`).join(' ')}] /Count ${pageIds.length} >>`);
  pageIds.forEach((pageId) => {
    objects[pageId - 1] = objects[pageId - 1].replace('/Parent 0 0 R', `/Parent ${pagesId} 0 R`);
  });
  const catalogId = addObject(`<< /Type /Catalog /Pages ${pagesId} 0 R >>`);

  let pdf = '%PDF-1.4\n';
  const offsets = [0];
  objects.forEach((body, index) => {
    offsets.push(pdf.length);
    pdf += `${index + 1} 0 obj\n${body}\nendobj\n`;
  });
  const xrefStart = pdf.length;
  pdf += `xref\n0 ${objects.length + 1}\n`;
  pdf += '0000000000 65535 f \n';
  offsets.slice(1).forEach((offset) => {
    pdf += `${String(offset).padStart(10, '0')} 00000 n \n`;
  });
  pdf += `trailer\n<< /Size ${objects.length + 1} /Root ${catalogId} 0 R >>\nstartxref\n${xrefStart}\n%%EOF`;

  return new TextEncoder().encode(pdf);
};

const PLAN_IMPORT_MONTHS = ["January","February","March","April","May","June","July","August","September","October","November","December"];
const PLAN_IMPORT_WEEKDAYS = ["sun","mon","tue","wed","thu","fri","sat"];
const loadSheetJS = () => new Promise((resolve, reject) => {
  if (window.XLSX) return resolve(window.XLSX);
  const existing = document.querySelector('script[data-sheetjs="true"]');
  if (existing) {
    existing.addEventListener("load", () => resolve(window.XLSX));
    existing.addEventListener("error", reject);
    return;
  }
  const s = document.createElement("script");
  s.src = "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";
  s.async = true;
  s.dataset.sheetjs = "true";
  s.onload = () => resolve(window.XLSX);
  s.onerror = () => reject(new Error("Failed to load spreadsheet parser."));
  document.head.appendChild(s);
});
const planText = (value) => String(value ?? "").replace(/\s+/g, " ").trim();
const planTextLower = (value) => planText(value).toLowerCase();
const planCellNumber = (value) => {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  const raw = planText(value);
  if (!raw) return 0;
  const cleaned = raw.replace(/[₦,\s]/g, "");
  const parsed = Number(cleaned);
  return Number.isFinite(parsed) ? parsed : 0;
};
const normalizeDurationValue = (value, fallback = "30") => {
  const raw = String(value ?? "").trim();
  if (!raw) return fallback;
  const cleaned = raw
    .replace(/\bseconds?\b/gi, "")
    .replace(/\bsecs?\b/gi, "")
    .replace(/["']/g, "")
    .trim();
  if (!cleaned) return fallback;
  const numeric = cleaned.match(/\d+(?:\.\d+)?/);
  return numeric ? numeric[0] : cleaned;
};
const planCellHasValue = (value) => {
  if (typeof value === "number" && Number.isFinite(value)) return true;
  return planText(value) !== "";
};
const extractPlanPercent = (value) => {
  const textValue = planText(value);
  if (!textValue) return null;
  const match = textValue.match(/(-?\d+(?:\.\d+)?)\s*%/);
  if (match) return Number(match[1]) || 0;
  return null;
};
const parseCompositePlanFooterPercents = (rows = [], startIndex = 0) => {
  const result = { discountPct: null, commissionPct: null };
  for (let rowIndex = Math.max(0, startIndex); rowIndex < rows.length; rowIndex += 1) {
    const row = rows[rowIndex] || [];
    const combined = row.map(planText).filter(Boolean).join(' | ');
    if (!combined) continue;
    if (result.discountPct === null && /volume discount|discount/i.test(combined)) {
      const pct = extractPlanPercent(combined);
      if (pct !== null) result.discountPct = pct;
    }
    if (result.commissionPct === null && /agency commission|commission/i.test(combined)) {
      const pct = extractPlanPercent(combined);
      if (pct !== null) result.commissionPct = pct;
    }
    if (result.discountPct !== null && result.commissionPct !== null) break;
  }
  return result;
};
const normalizeMediaPlanVendorName = (value) => planText(value).toLowerCase().replace(/[^a-z0-9]+/g, " ").replace(/\s+/g, " ").trim();
const isPlanWeekday = (value) => PLAN_IMPORT_WEEKDAYS.includes(planTextLower(value).slice(0, 3));
const looksLikePlanTimeBelt = (value) => /\d{1,2}:\d{2}/.test(planText(value));

const looksLikePlanHeader = (value) => {
  const textValue = planText(value);
  if (!textValue) return false;
  if (looksLikePlanTimeBelt(textValue) || isPlanWeekday(textValue)) return false;
  const letters = textValue.replace(/[^A-Za-z]+/g, "");
  if (!letters) return false;
  const uppercaseRatio = letters.split("").filter(ch => ch === ch.toUpperCase()).length / letters.length;
  if (uppercaseRatio >= 0.75) return true;
  return textValue.split(" ").length <= 6 && !/\d/.test(textValue) && /[A-Za-z]{3,}/.test(textValue);
};

const getPlanLeadCell = (row = [], maxScanCol = 10) => {
  const limit = Math.min(row.length, Math.max(1, maxScanCol));
  for (let index = 0; index < limit; index += 1) {
    const textValue = planText(row[index]);
    if (textValue) return { text: textValue, index };
  }
  return { text: "", index: -1 };
};

const GENERIC_MEDIA_PLAN_REGION_HEADERS = new Set([
  "network",
  "lagos",
  "abuja",
  "ibadan",
  "kano",
  "kaduna",
  "port harcourt",
  "ph",
  "oyo",
  "edo",
  "enugu",
  "rivers",
  "north",
  "south",
  "east",
  "west",
  "north east",
  "north west",
  "north central",
  "south east",
  "south south",
  "south west",
  "regional",
  "regionals",
  "regional stations",
  "terrestial tv",
  "terrestrial tv",
  "radio",
  "television",
  "tv",
]);

const isGenericRegionHeader = (value = "") => {
  const normalized = normalizeMediaPlanVendorName(value);
  if (!normalized) return false;
  if (GENERIC_MEDIA_PLAN_REGION_HEADERS.has(normalized)) return true;
  return (
    looksLikePlanHeader(value) &&
    normalized.split(" ").length <= 3 &&
    !/(fm|tv|channel|radio|info|max|ait|nta|stv|silverbird|wazobia|bond|kiss|raypower|brila|beat|cool|classic|rhythm|inspiration|arewa|unity|rockcity|lagos talks|treasure|ltv|itv|etv|ebs|bcos|rahma|liberty|galaxy|artv|ctv)/i.test(value)
  );
};

const isLikelyVendorHeaderText = (value = "") => {
  const textValue = planText(value);
  if (!textValue || !looksLikePlanHeader(textValue) || isGenericRegionHeader(textValue)) return false;
  return (
    /(fm|tv|channel|radio|info|max|ait|nta|stv|silverbird|wazobia|bond|kiss|raypower|brila|beat|cool|classic|rhythm|inspiration|arewa|unity|rockcity|lagos talks|treasure|ltv|itv|etv|ebs|bcos|rahma|liberty|galaxy|artv|ctv)/i.test(textValue) ||
    normalizeMediaPlanVendorName(textValue).split(" ").length <= 4
  );
};

const rowHasMarkedSchedule = (row = [], layout = {}) => {
  for (let col = layout.scheduleStartCol; col <= layout.scheduleEndCol; col += 1) {
    if (planCellNumber(row[col]) > 0) return true;
  }
  return false;
};

const rowHasPlanSummaryValues = (row = [], layout = {}) => {
  const candidateCols = [
    layout.totalSpotsCol,
    layout.bonusSpotsCol,
    layout.plannedSpotsCol,
    layout.grossValueCol,
  ].filter((col) => Number.isInteger(col) && col >= 0);
  return candidateCols.some((col) => planCellNumber(row[col]) > 0);
};

const rowLooksLikeScheduleEntry = (row = [], layout = {}) => {
  const wdValue = planText(row[layout.wdCol]);
  const timeValue = planText(row[layout.timeCol]);
  const materialValue = planText(row[layout.materialCol]);
  const durationValue = planText(row[layout.durationCol]);
  const rateValue = planCellNumber(row[layout.rateCol]);
  return Boolean(
    isPlanWeekday(wdValue) ||
    looksLikePlanTimeBelt(timeValue) ||
    materialValue ||
    durationValue ||
    rateValue > 0 ||
    rowHasMarkedSchedule(row, layout) ||
    rowHasPlanSummaryValues(row, layout)
  );
};

const isPlanRowEmpty = (row = []) =>
  !(row || []).some((cell) => planText(cell));

const shouldStopCompositePlanParsingAtRow = (rows = [], rowIndex = 0, layout = {}, blankLookahead = 18) => {
  const row = rows[rowIndex] || [];
  if (!isPlanRowEmpty(row)) return false;

  let sawMeaningfulGap = false;
  for (let nextIndex = rowIndex + 1; nextIndex < rows.length && nextIndex <= rowIndex + blankLookahead; nextIndex += 1) {
    const nextRow = rows[nextIndex] || [];
    if (isPlanRowEmpty(nextRow)) continue;

    sawMeaningfulGap = true;
    if (isCompositePlanTotalRow(nextRow, layout)) return true;
    if (rowLooksLikeScheduleEntry(nextRow, layout)) return false;

    const leadCell = getPlanLeadCell(nextRow, Math.max(8, (layout?.rateCol || 0) + 2));
    const leadText = leadCell.text;
    if (!leadText) continue;
    if (isLikelyVendorHeaderText(leadText) || isGenericRegionHeader(leadText) || looksLikePlanHeader(leadText)) {
      return false;
    }
  }

  return !sawMeaningfulGap;
};

const isCompositePlanTotalRow = (row = [], layout = {}) => {
  const cells = (row || []).map(planText).filter(Boolean);
  if (!cells.length) return false;

  const summaryText = cells.join(" | ").toLowerCase();
  const leadText = planText(row?.[layout?.programmeCol]);
  const wdText = planText(row?.[layout?.wdCol]);
  const timeText = planText(row?.[layout?.timeCol]);
  const materialText = planText(row?.[layout?.materialCol]);
  const durationText = planText(row?.[layout?.durationCol]);

  const leadingCells = (row || [])
    .slice(0, Math.max(8, (layout?.scheduleStartCol || 0)))
    .map(planText)
    .filter(Boolean);

  if (/(grand total|overall total|campaign total|sub total|subtotal|total net|total gross|net total|gross total)/i.test(summaryText)) {
    return true;
  }

  if (/^(total|totals|grand total|sub total|subtotal)$/i.test(leadText)) {
    return true;
  }

  if (
    leadingCells.length <= 3 &&
    leadingCells.some((value) => /^(total|totals|grand total|sub total|subtotal)$/i.test(value))
  ) {
    return true;
  }

  const descriptiveCellCount = [leadText, wdText, timeText, materialText, durationText].filter(Boolean).length;
  const scheduleMarkCount = (() => {
    let count = 0;
    for (let col = layout?.scheduleStartCol || 0; col <= (layout?.scheduleEndCol || -1); col += 1) {
      if (planCellNumber(row?.[col]) > 0) count += 1;
    }
    return count;
  })();

  const summaryCols = [
    layout?.rateCol,
    layout?.totalSpotsCol,
    layout?.bonusSpotsCol,
    layout?.plannedSpotsCol,
    layout?.grossValueCol,
    layout?.discountPctCol,
    layout?.discountValueCol,
    layout?.commissionPctCol,
    layout?.commissionValueCol,
    layout?.netValueCol,
  ].filter((value) => Number.isInteger(value) && value >= 0);

  const summaryValueCount = summaryCols.reduce((count, col) => count + (planCellNumber(row?.[col]) > 0 ? 1 : 0), 0);

  if (
    descriptiveCellCount === 0 &&
    (scheduleMarkCount >= 5 || summaryValueCount >= 3)
  ) {
    return true;
  }

  if (
    descriptiveCellCount <= 1 &&
    !looksLikePlanTimeBelt(timeText) &&
    !isPlanWeekday(wdText) &&
    !materialText &&
    scheduleMarkCount >= 8
  ) {
    return true;
  }

  return false;
};

const isCompositePlanSummaryOnlyRow = (params = {}) => {
  const {
    programmeValue = "",
    wdValue = "",
    timeValue = "",
    materialValue = "",
    durationValue = "",
    rateValue = 0,
    scheduledFromGrid = 0,
    totalSpots = 0,
    importedGross = 0,
    importedNet = 0,
    bonusSpots = 0,
  } = params || {};

  const clean = (value = "") => {
    const text = planText(value).trim();
    if (!text) return "";
    if (/^(—|-|–|n\/a|na)$/i.test(text)) return "";
    return text;
  };

  const hasProgramme = !!clean(programmeValue);
  const hasWeekday = isPlanWeekday(wdValue);
  const hasTime = looksLikePlanTimeBelt(timeValue);
  const hasMaterial = !!clean(materialValue);
  const hasDuration = !!clean(durationValue);
  const hasRate = Number(rateValue) > 0;
  const hasScheduleGrid = Number(scheduledFromGrid) > 0;
  const hasSummaryOnlyTotals = Number(totalSpots) > 0 || Number(importedGross) > 0 || Number(importedNet) > 0 || Number(bonusSpots) > 0;

  if (
    hasSummaryOnlyTotals &&
    !hasScheduleGrid &&
    !hasWeekday &&
    !hasTime &&
    !hasMaterial &&
    !hasDuration &&
    !hasRate
  ) {
    return true;
  }

  if (
    hasSummaryOnlyTotals &&
    !hasScheduleGrid &&
    hasProgramme &&
    !hasWeekday &&
    !hasTime &&
    !hasMaterial &&
    !hasDuration &&
    !hasRate
  ) {
    return true;
  }

  return false;
};

const findNextScheduleRowOffset = (rows = [], startIndex = 0, layout = {}, lookahead = 10) => {
  for (let offset = 1; offset <= lookahead; offset += 1) {
    const row = rows[startIndex + offset];
    if (!row) break;
    if (rowLooksLikeScheduleEntry(row, layout)) return offset;
  }
  return -1;
};

const getPlanRowSummaryText = (row = [], from = 0, to = row.length - 1) =>
  row.slice(from, to + 1).map(planText).filter(Boolean).join(" | ");

const countPlanRowPatternHits = (row = [], pattern) =>
  (row || []).reduce((count, cell) => count + (pattern.test(planTextLower(cell)) ? 1 : 0), 0);

const getCompositeHeaderRowScore = (row = []) => {
  const rowText = getPlanRowSummaryText(row).toLowerCase();
  if (!rowText) return 0;
  let score = 0;
  if (/(spot rate|rate per spot|spotrate|rate)/i.test(rowText)) score += 4;
  if (/(material title|material specification|material|specification|creative)/i.test(rowText)) score += 4;
  if (/(targeted time|time belt|timebelt|time slot|air time|slot)/i.test(rowText)) score += 4;
  if (/(^|\W)(wd|weekday)(\W|$)/i.test(rowText)) score += 3;
  if (/(spot duration|duration|secs|seconds)/i.test(rowText)) score += 3;
  if (/(television|radio|print|programme|program)/i.test(rowText)) score += 3;
  if (/(bonus|gross|disc|commission|net)/i.test(rowText)) score += 2;
  return score;
};

const findBestNearbyRowIndex = (rows = [], fromIndex = 0, toIndex = 0, scorer = () => 0, preferClosestTo = null) => {
  let best = { index: -1, score: -Infinity, distance: Infinity };
  for (let i = Math.max(0, fromIndex); i <= Math.min(rows.length - 1, toIndex); i += 1) {
    const score = scorer(rows[i] || [], i);
    const distance = preferClosestTo === null ? 0 : Math.abs(i - preferClosestTo);
    if (score > best.score || (score === best.score && distance < best.distance)) {
      best = { index: i, score, distance };
    }
  }
  return best.index;
};

const buildColumnHeaderTextReader = (rows = [], fromIndex = 0, toIndex = 0) => {
  const startIndex = Math.max(0, fromIndex);
  const endIndex = Math.min(rows.length - 1, Math.max(startIndex, toIndex));
  return (colIndex) =>
    rows
      .slice(startIndex, endIndex + 1)
      .map((row) => planTextLower(row?.[colIndex]))
      .filter(Boolean)
      .join(" | ");
};

const findColumnByHeaderPatterns = (columnCount = 0, getHeaderText = () => "", patterns = []) => {
  for (let index = 0; index < columnCount; index += 1) {
    const headerText = getHeaderText(index);
    if (!headerText) continue;
    if (patterns.some((pattern) => pattern.test(headerText))) return index;
  }
  return -1;
};

const findAllColumnsByHeaderPatterns = (columnCount = 0, getHeaderText = () => "", patterns = []) => {
  const matches = [];
  for (let index = 0; index < columnCount; index += 1) {
    const headerText = getHeaderText(index);
    if (!headerText) continue;
    if (patterns.some((pattern) => pattern.test(headerText))) matches.push(index);
  }
  return matches;
};

const scorePlanColumnByDataPattern = (rows = [], columnIndex = -1, predicate = () => false, fromIndex = 0, toIndex = 0, sampleLimit = 80) => {
  if (columnIndex < 0) return 0;
  let score = 0;
  let seen = 0;
  for (let rowIndex = Math.max(0, fromIndex); rowIndex <= Math.min(rows.length - 1, toIndex); rowIndex += 1) {
    const raw = planText(rows[rowIndex]?.[columnIndex]);
    if (!raw) continue;
    seen += 1;
    if (predicate(raw)) score += 1;
    if (seen >= sampleLimit) break;
  }
  return score;
};

const inferCompositePlanMedium = (rows = [], indices = []) => {
  const summary = indices
    .map((index) => getPlanRowSummaryText(rows[index] || []))
    .filter(Boolean)
    .join(" | ")
    .toLowerCase();
  if (/radio/.test(summary)) return "Radio";
  if (/(television|terrestial|terrestrial| tv\b)/.test(summary)) return "Television";
  if (/print/.test(summary)) return "Print";
  return "";
};

const detectCompositePlanLayout = (rows = []) => {
  const headerCandidates = rows
    .map((row, index) => ({ index, score: getCompositeHeaderRowScore(row) }))
    .filter((item) => item.score >= 8)
    .sort((a, b) => (b.score - a.score) || (a.index - b.index));

  const metadataHeaderRow = headerCandidates[0]?.index ?? -1;
  if (metadataHeaderRow === -1) {
    throw new Error("Could not detect the media plan header row.");
  }

  const headerSearchFrom = Math.max(0, metadataHeaderRow - 3);
  const headerSearchTo = Math.min(rows.length - 1, metadataHeaderRow + 4);

  const weekdayHeaderRowIndex = findBestNearbyRowIndex(
    rows,
    metadataHeaderRow - 1,
    headerSearchTo,
    (row) => countPlanRowPatternHits(row, /^(sun|mon|tue|wed|thu|fri|sat)$/i),
    metadataHeaderRow + 2
  );

  const dateHeaderRowIndex = findBestNearbyRowIndex(
    rows,
    metadataHeaderRow,
    headerSearchTo,
    (row) => (row || []).reduce((count, cell) => {
      const value = planCellNumber(cell);
      return count + ((value >= 1 && value <= 31) ? 1 : 0);
    }, 0),
    weekdayHeaderRowIndex >= 0 ? weekdayHeaderRowIndex - 1 : metadataHeaderRow + 1
  );

  const monthHeaderRowIndex = findBestNearbyRowIndex(
    rows,
    headerSearchFrom,
    Math.max(headerSearchFrom, (dateHeaderRowIndex >= 0 ? dateHeaderRowIndex : metadataHeaderRow)),
    (row) => countPlanRowPatternHits(row, /^(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)/i),
    dateHeaderRowIndex >= 0 ? dateHeaderRowIndex - 1 : metadataHeaderRow - 1
  );

  const wdHeaderRowIndex = findBestNearbyRowIndex(
    rows,
    headerSearchFrom,
    headerSearchTo,
    (row) => countPlanRowPatternHits(row, /^(wd|weekday)$/i),
    metadataHeaderRow
  );

  const metadataRow = rows[metadataHeaderRow] || [];
  const monthHeaderRow = rows[Math.max(0, monthHeaderRowIndex)] || [];
  const wdHeaderRow = rows[Math.max(0, wdHeaderRowIndex)] || [];
  const dateHeaderRow = rows[Math.max(0, dateHeaderRowIndex)] || [];
  const weekdayHeaderRow = rows[Math.max(0, weekdayHeaderRowIndex)] || [];

  const columnCount = Math.max(
    metadataRow.length,
    monthHeaderRow.length,
    wdHeaderRow.length,
    dateHeaderRow.length,
    weekdayHeaderRow.length
  );

  const getHeaderText = buildColumnHeaderTextReader(
    rows,
    Math.max(0, Math.min(
      metadataHeaderRow - 1,
      wdHeaderRowIndex >= 0 ? wdHeaderRowIndex : metadataHeaderRow,
      monthHeaderRowIndex >= 0 ? monthHeaderRowIndex : metadataHeaderRow
    )),
    Math.max(metadataHeaderRow, weekdayHeaderRowIndex >= 0 ? weekdayHeaderRowIndex : metadataHeaderRow)
  );

  const programmeCol = findColumnByHeaderPatterns(columnCount, getHeaderText, [
    /(television|radio|print)/i,
    /(programme|program)/i,
  ]);

  const wdCol = findColumnByHeaderPatterns(columnCount, getHeaderText, [
    /(^|\W)(wd|weekday)(\W|$)/i,
  ]);

  const timeHeaderCols = Array.from({ length: columnCount }, (_, index) => index).filter((index) =>
    /(targeted time|time belt|timebelt|time slot|air time|slot)/i.test(getHeaderText(index))
  );
  const timeCol = timeHeaderCols.find((index) => index > (wdCol >= 0 ? wdCol : (programmeCol >= 0 ? programmeCol : 1)))
    ?? timeHeaderCols[0]
    ?? -1;

  let materialCol = findColumnByHeaderPatterns(columnCount, getHeaderText, [
    /(material title|material specification|material|specification|creative|requirement)/i,
  ]);
  if (materialCol >= 0 && timeCol >= 0 && materialCol <= timeCol) {
    materialCol = findColumnByHeaderPatterns(columnCount, (index) => index > timeCol ? getHeaderText(index) : "", [
      /(material title|material specification|material|specification|creative|requirement)/i,
    ]);
  }

  let durationCol = findColumnByHeaderPatterns(columnCount, getHeaderText, [
    /(spot duration|duration|secs|seconds)/i,
  ]);
  if (durationCol >= 0 && materialCol >= 0 && durationCol <= materialCol) {
    durationCol = findColumnByHeaderPatterns(columnCount, (index) => index > materialCol ? getHeaderText(index) : "", [
      /(spot duration|duration|secs|seconds)/i,
    ]);
  }

  const rateCol = findColumnByHeaderPatterns(columnCount, getHeaderText, [
    /(spot rate|rate per spot|spotrate|rate)/i,
  ]);

  const raCol = findColumnByHeaderPatterns(columnCount, getHeaderText, [
    /(^|\W)ra(\W|$)/i,
    /rate analysis/i,
  ]);

  let scheduleStartCol = -1;
  for (let index = 0; index < columnCount; index += 1) {
    const dateCell = planCellNumber(dateHeaderRow[index]);
    if (dateCell >= 1 && dateCell <= 31) {
      scheduleStartCol = index;
      break;
    }
  }
  if (scheduleStartCol === -1) {
    for (let index = 0; index < columnCount; index += 1) {
      const weekdayCell = planText(weekdayHeaderRow[index]);
      if (isPlanWeekday(weekdayCell)) {
        scheduleStartCol = index;
        break;
      }
    }
  }
  if (scheduleStartCol === -1) {
    scheduleStartCol = findColumnByHeaderPatterns(columnCount, getHeaderText, [/(week day|weekdays|schedule)/i]);
  }
  if (scheduleStartCol === -1) {
    throw new Error("Could not detect the schedule matrix in this media plan.");
  }

  let summaryStartCol = columnCount;
  for (let index = scheduleStartCol; index < columnCount; index += 1) {
    const headerText = getHeaderText(index);
    if (/(bonus|planned|gross|disc|discount|comm|commission|net|no of spots|spot total|total spots|value|cost of)/i.test(headerText)) {
      summaryStartCol = index;
      break;
    }
  }
  if (summaryStartCol === columnCount) {
    for (let index = scheduleStartCol; index < columnCount; index += 1) {
      const dateCell = planCellNumber(dateHeaderRow[index]);
      const weekdayCell = planText(weekdayHeaderRow[index]);
      if (!(dateCell >= 1 && dateCell <= 31) && !isPlanWeekday(weekdayCell) && getHeaderText(index)) {
        summaryStartCol = index;
        break;
      }
    }
  }

  const summaryHeaderStartRow = Math.max(0, Math.min(
    metadataHeaderRow,
    monthHeaderRowIndex >= 0 ? monthHeaderRowIndex : metadataHeaderRow,
    dateHeaderRowIndex >= 0 ? dateHeaderRowIndex : metadataHeaderRow,
    weekdayHeaderRowIndex >= 0 ? weekdayHeaderRowIndex : metadataHeaderRow
  ) - 1);
  const summaryHeaderEndRow = Math.min(rows.length - 1, Math.max(
    metadataHeaderRow,
    monthHeaderRowIndex >= 0 ? monthHeaderRowIndex : metadataHeaderRow,
    dateHeaderRowIndex >= 0 ? dateHeaderRowIndex : metadataHeaderRow,
    weekdayHeaderRowIndex >= 0 ? weekdayHeaderRowIndex : metadataHeaderRow
  ) + 1);

  const summaryHeaderTextByCol = Array.from({ length: columnCount }, (_, index) => {
    const parts = [];
    for (let rowIndex = summaryHeaderStartRow; rowIndex <= summaryHeaderEndRow; rowIndex += 1) {
      const raw = planTextLower(rows[rowIndex]?.[index]);
      if (raw) parts.push(raw);
    }
    return parts.join(" | ");
  });

  const findSummaryCol = (predicate) => {
    for (let index = summaryStartCol; index < columnCount; index += 1) {
      const headerText = [getHeaderText(index), summaryHeaderTextByCol[index]].filter(Boolean).join(" | ");
      if (predicate(headerText, index)) return index;
    }
    return -1;
  };

  const totalSpotsCol = findSummaryCol((value) =>
    /(no\s*of\s*spots|total\s*burst\s*spots|total\s*spots|burst\s*spots|scheduled\s*spots)/i.test(value)
  );
  const bonusSpotsCol = findSummaryCol((value) =>
    /(bonus|free|foc|complimentary)\s*(spots?)?/i.test(value) && !/(cost\s*of|bonus\s*cost|value)/i.test(value)
  );
  const bonusCostCol = findSummaryCol((value) =>
    /(cost\s*of\s*bonus\s*spots|bonus\s*cost|bonus\s*value|value\s*of\s*bonus)/i.test(value)
  );
  const plannedSpotsCol = findSummaryCol((value) =>
    /(planned\s*spots|paid\s*spots|booked\s*spots)/i.test(value)
  );
  const grossValueCol = findSummaryCol((value) =>
    /(total\s*gross\s*value|gross\s*value|gross\s*cost)/i.test(value)
  );
  const discountPctCol = findSummaryCol((value) =>
    /((disc|discount)\s*%)/i.test(value)
  );
  const discountValueCol = findSummaryCol((value) =>
    /((disc|discount)\s*value)/i.test(value)
  );
  const commissionPctCol = findSummaryCol((value) =>
    /((agency\s*)?(comm|commission)\s*%)/i.test(value)
  );
  const commissionValueCol = findSummaryCol((value) =>
    /((agency\s*)?(comm|commission)\s*value)/i.test(value)
  );
  const netValueCol = findSummaryCol((value) =>
    /net\s*value/i.test(value)
  );

  const medium = inferCompositePlanMedium(rows, [
    metadataHeaderRow,
    monthHeaderRowIndex,
    dateHeaderRowIndex,
    weekdayHeaderRowIndex,
  ].filter((value) => Number.isInteger(value) && value >= 0));

  return {
    metadataHeaderRow,
    dataStartRow: Math.max(metadataHeaderRow + 1, weekdayHeaderRowIndex >= 0 ? weekdayHeaderRowIndex + 1 : metadataHeaderRow + 1),
    programmeCol: programmeCol >= 0 ? programmeCol : 1,
    wdCol: wdCol >= 0 ? wdCol : Math.max(0, (programmeCol >= 0 ? programmeCol + 1 : 2)),
    timeCol: timeCol >= 0 ? timeCol : Math.max(0, (wdCol >= 0 ? wdCol + 1 : 3)),
    materialCol: materialCol >= 0 ? materialCol : Math.max(0, (timeCol >= 0 ? timeCol + 1 : 4)),
    raCol,
    durationCol: durationCol >= 0 ? durationCol : Math.max(0, (materialCol >= 0 ? materialCol + 1 : 5)),
    rateCol: rateCol >= 0 ? rateCol : Math.max(0, (durationCol >= 0 ? durationCol + 1 : 6)),
    scheduleStartCol,
    scheduleEndCol: Math.max(scheduleStartCol, (summaryStartCol === columnCount ? columnCount - 1 : summaryStartCol - 1)),
    monthHeaderRow,
    monthHeaderRowIndex,
    dateHeaderRow,
    dateHeaderRowIndex,
    weekdayHeaderRow,
    weekdayHeaderRowIndex,
    totalSpotsCol,
    bonusSpotsCol,
    bonusCostCol,
    plannedSpotsCol,
    grossValueCol,
    discountPctCol,
    discountValueCol,
    commissionPctCol,
    commissionValueCol,
    netValueCol,
    medium,
  };
};

const PLAN_MONTH_NAMES = ["january","february","march","april","may","june","july","august","september","october","november","december"];
const normalizePlanMonthName = (value = "") => {
  const raw = String(value || "").trim().toLowerCase();
  const direct = PLAN_MONTH_NAMES.find(name => name === raw);
  if (direct) return direct.charAt(0).toUpperCase() + direct.slice(1);
  const short = PLAN_MONTH_NAMES.find(name => name.slice(0, 3) === raw.slice(0, 3));
  return short ? short.charAt(0).toUpperCase() + short.slice(1) : "";
};


const buildCompositePlanScheduleColumnMap = (layout, yearValue = "") => {
  const scheduleColumns = {};
  const fallbackYear = Number(yearValue) || new Date().getFullYear();
  const monthHeader = layout?.monthHeaderRow || [];
  const dateHeader = layout?.dateHeaderRow || [];

  let currentMonthIndex = -1;
  let previousDayNumber = null;

  const firstExplicitMonthIndex = (() => {
    for (let col = layout.scheduleStartCol; col <= layout.scheduleEndCol; col += 1) {
      const label = normalizePlanMonthName(monthHeader[col]);
      if (label) return PLAN_MONTH_NAMES.indexOf(label.toLowerCase());
    }
    return -1;
  })();
  if (firstExplicitMonthIndex >= 0) currentMonthIndex = firstExplicitMonthIndex;

  for (let col = layout.scheduleStartCol; col <= layout.scheduleEndCol; col += 1) {
    const dayNumber = planCellNumber(dateHeader[col]);
    if (!(dayNumber >= 1 && dayNumber <= 31)) continue;

    const explicitMonth = normalizePlanMonthName(monthHeader[col]);
    if (explicitMonth) {
      currentMonthIndex = PLAN_MONTH_NAMES.indexOf(explicitMonth.toLowerCase());
    } else if (currentMonthIndex === -1) {
      currentMonthIndex = 0;
    } else if (previousDayNumber !== null && dayNumber < previousDayNumber) {
      currentMonthIndex = (currentMonthIndex + 1) % 12;
    }

    const actualDate = new Date(fallbackYear, Math.max(0, currentMonthIndex), dayNumber);
    scheduleColumns[col] = {
      col,
      weekday: actualDate.toLocaleDateString("en-NG", { weekday: "short" }),
      dayNumber: actualDate.getDate(),
      monthLabel: actualDate.toLocaleDateString("en-NG", { month: "long" }),
      isoDate: actualDate.toISOString().slice(0, 10),
      date: actualDate,
    };

    previousDayNumber = dayNumber;
  }

  return scheduleColumns;
};

const detectCompositePlanMonthSegments = (layout, rows = []) => {
  const monthHeader = layout?.monthHeaderRow || [];
  const segments = [];
  for (let col = layout.scheduleStartCol; col <= layout.scheduleEndCol; col += 1) {
    const monthLabel = normalizePlanMonthName(monthHeader[col]);
    if (!monthLabel) continue;
    segments.push({ label: monthLabel, startCol: col, endCol: layout.scheduleEndCol });
  }
  if (!segments.length) {
    return [{ label: "", startCol: layout.scheduleStartCol, endCol: layout.scheduleEndCol }];
  }
  return segments.map((segment, index) => ({
    ...segment,
    endCol: (segments[index + 1]?.startCol || (layout.scheduleEndCol + 1)) - 1,
  }));
};

const getMonthKeyFromScheduleLabel = (scheduleLabel = "", fallbackMonth = "") => {
  const normalized = normalizePlanMonthName(scheduleLabel);
  if (normalized) return normalized;
  return normalizePlanMonthName(fallbackMonth) || String(fallbackMonth || "").trim();
};

const buildCalendarDraftRowsFromSpots = (spots = [], fallbackMonth = "") => {
  const grouped = {};
  (spots || []).forEach((spot) => {
    const monthKey = getMonthKeyFromScheduleLabel(spot?.scheduleMonth || "", fallbackMonth || "");
    if (!grouped[monthKey]) grouped[monthKey] = [];
    grouped[monthKey].push({
      id: uid(),
      programme: spot?.programme || "",
      timeBelt: spot?.timeBelt || "",
      material: spot?.material || "",
      materialCustom: "",
      duration: normalizeDurationValue(spot?.duration, "30"),
      rateId: spot?.rateId || "",
      customRate: spot?.isComplimentary ? "" : (spot?.customRate || (spot?.ratePerSpot ? String(spot.ratePerSpot) : "")),
      dayCounts: collapseDaysToCounts(Array.isArray(spot?.calendarDays) ? spot.calendarDays : []),
      isComplimentary: !!spot?.isComplimentary,
      bonusSpots: spot?.bonusSpots ?? getSpotBonusCount(spot),
    });
  });
  return grouped;
};

const findBestMediaPlanVendorMatch = (vendorName, vendors = []) => {
  const normalized = normalizeMediaPlanVendorName(vendorName);
  if (!normalized) return { vendor: null, matchType: "none" };
  const exact = (vendors || []).find(item => normalizeMediaPlanVendorName(item.name) === normalized);
  if (exact) return { vendor: exact, matchType: "exact" };
  const fuzzyCandidates = (vendors || []).filter(item => {
    const candidate = normalizeMediaPlanVendorName(item.name);
    return candidate && (candidate.includes(normalized) || normalized.includes(candidate));
  });
  if (fuzzyCandidates.length === 1) return { vendor: fuzzyCandidates[0], matchType: "fuzzy" };
  return { vendor: null, matchType: "none" };
};

const buildMediaPlanVendorDisplayName = (vendorName = "", regionName = "", mediumLabel = "") => {
  const base = planText(vendorName);
  const region = planText(regionName);
  if (!base) return region || "";
  if (!region) return base;
  const normalizedBase = normalizeMediaPlanVendorName(base);
  const normalizedRegion = normalizeMediaPlanVendorName(region);
  if (!normalizedRegion || normalizedBase.includes(normalizedRegion)) return base;
  return `${base} ${region}`;
};

const findBestMediaPlanVendorMatchFromCandidates = (candidateLabels = [], vendors = []) => {
  const uniqueLabels = [...new Set((candidateLabels || []).map(planText).filter(Boolean))];
  let best = { vendor: null, matchType: "none" };
  for (const label of uniqueLabels) {
    const match = findBestMediaPlanVendorMatch(label, vendors);
    if (match.vendor && match.matchType === "exact") return match;
    if (match.vendor && best.matchType === "none") best = match;
  }
  return best;
};

const parseCompositeMediaPlanRows = (sheetRows = [], vendors = [], sourceSheet = "", yearValue = "") => {
  const layout = detectCompositePlanLayout(sheetRows);
  const scheduleColumnMap = buildCompositePlanScheduleColumnMap(layout, yearValue);
  const footerPercents = parseCompositePlanFooterPercents(sheetRows, layout.dataStartRow);
  const parsedRows = [];
  const warnings = [];
  let currentRegion = "";
  let currentVendor = "";
  let lastProgramme = "";
  const mediumLabel = layout?.medium || "";

  const inferVendorFromScheduleRow = (programmeValue = "") => {
    const candidates = [
      planText(programmeValue),
      planText(programmeValue).split(",").slice(0, 2).join(", "),
      planText(programmeValue).split(",").find(part => /(fm|tv|channel|info|max|wazobia|kiss|nta|ait|stv|bond|raypower|brila|beat|cool|classic|rhythm|inspiration|arewa|unity|rockcity|lagos talks|treasure|ltv|itv|etv|ebs|bcos|rahma|liberty|galaxy|artv|ctv)/i.test(part)) || "",
    ].map(planText).filter(Boolean);

    for (const candidate of candidates) {
      const match = findBestMediaPlanVendorMatch(candidate, vendors);
      if (match.vendor?.name) return match.vendor.name;
    }
    return "";
  };

  const distributeBonusAcrossSlices = (slices = [], totalBonus = 0) => {
    const normalizedBonus = Math.max(0, Number(totalBonus) || 0);
    if (!normalizedBonus || !slices.length) return slices.map(() => 0);
    const totalScheduled = slices.reduce((sum, slice) => sum + (slice?.scheduledFromGrid || 0), 0);
    if (!totalScheduled) return slices.map(() => 0);

    const rawAllocations = slices.map((slice) => normalizedBonus * ((slice?.scheduledFromGrid || 0) / totalScheduled));
    const floors = rawAllocations.map((value, index) => Math.min(slices[index]?.scheduledFromGrid || 0, Math.floor(value)));
    let remaining = normalizedBonus - floors.reduce((sum, value) => sum + value, 0);

    const order = rawAllocations
      .map((value, index) => ({
        index,
        remainder: value - floors[index],
        capacity: Math.max(0, (slices[index]?.scheduledFromGrid || 0) - floors[index]),
      }))
      .sort((a, b) => (b.remainder - a.remainder) || ((slices[b.index]?.scheduledFromGrid || 0) - (slices[a.index]?.scheduledFromGrid || 0)));

    while (remaining > 0) {
      const next = order.find((item) => item.capacity > 0);
      if (!next) break;
      floors[next.index] += 1;
      next.capacity -= 1;
      remaining -= 1;
    }

    return floors;
  };

  for (let rowIndex = layout.dataStartRow; rowIndex < sheetRows.length; rowIndex += 1) {
    const row = sheetRows[rowIndex] || [];

    if (shouldStopCompositePlanParsingAtRow(sheetRows, rowIndex, layout)) {
      break;
    }

    const leadCell = getPlanLeadCell(row, Math.max(8, (layout.rateCol || 0) + 2));
    const leadText = leadCell.text;

    if (isCompositePlanTotalRow(row, layout)) {
      break;
    }

    const isScheduleRow = rowLooksLikeScheduleEntry(row, layout);

    if (!isScheduleRow) {
      if (!leadText) continue;
      const nextScheduleOffset = findNextScheduleRowOffset(sheetRows, rowIndex, layout);
      if (isGenericRegionHeader(leadText)) {
        currentRegion = leadText;
        currentVendor = "";
        lastProgramme = "";
        continue;
      }
      if (isLikelyVendorHeaderText(leadText) || (nextScheduleOffset > 0 && !isGenericRegionHeader(leadText) && normalizeMediaPlanVendorName(leadText).split(" ").length <= 5)) {
        currentVendor = leadText;
        lastProgramme = "";
        continue;
      }
      if (looksLikePlanHeader(leadText)) {
        currentRegion = leadText;
        currentVendor = "";
        lastProgramme = "";
      }
      continue;
    }

    const programmeValue = planText(row[layout.programmeCol]) || (leadCell.index === layout.programmeCol ? leadText : "");
    let wdValue = planText(row[layout.wdCol]);
    let timeValue = planText(row[layout.timeCol]);

    if (!looksLikePlanTimeBelt(timeValue) && looksLikePlanTimeBelt(row?.[layout.timeCol + 1])) {
      timeValue = planText(row?.[layout.timeCol + 1]);
    }
    if (!isPlanWeekday(wdValue) && isPlanWeekday(row?.[layout.wdCol - 1])) {
      wdValue = planText(row?.[layout.wdCol - 1]);
    }
    if (
      looksLikePlanTimeBelt(wdValue) &&
      (!looksLikePlanTimeBelt(timeValue) || isPlanWeekday(timeValue)) &&
      (isPlanWeekday(timeValue) || !timeValue)
    ) {
      const swappedWd = planText(timeValue);
      timeValue = wdValue;
      wdValue = swappedWd;
    }

    const materialValue = planText(row[layout.materialCol]);
    const rateValue = planCellNumber(row[layout.rateCol]);
    const ratePerSpot = Math.max(0, rateValue);
    const durationValue = planText(row[layout.durationCol]);
    const importedGross = Math.max(0, planCellNumber(row[layout.grossValueCol]));
    const programme = programmeValue || lastProgramme;

    if (!programme) continue;
    if (programmeValue) lastProgramme = programmeValue;

    const vendorName = currentVendor || inferVendorFromScheduleRow(programme) || "";
    const regionName = currentRegion;
    const vendorDisplayName = buildMediaPlanVendorDisplayName(vendorName, regionName, mediumLabel);

    const monthlySliceMap = new Map();
    let scheduledFromGrid = 0;
    let scheduleWeekdayMismatchCount = 0;

    for (let col = layout.scheduleStartCol; col <= layout.scheduleEndCol; col += 1) {
      const count = planCellNumber(row[col]);
      if (count <= 0) continue;
      const columnMeta = scheduleColumnMap[col];
      if (!columnMeta) continue;
      const monthKey = columnMeta.monthLabel || "";
      if (!monthlySliceMap.has(monthKey)) {
        monthlySliceMap.set(monthKey, {
          scheduleMonth: monthKey,
          dayCounts: {},
          scheduledFromGrid: 0,
        });
      }
      const slice = monthlySliceMap.get(monthKey);
      slice.dayCounts[columnMeta.dayNumber] = (slice.dayCounts[columnMeta.dayNumber] || 0) + count;
      slice.scheduledFromGrid += count;
      scheduledFromGrid += count;

      if (wdValue && isPlanWeekday(wdValue) && columnMeta.weekday && columnMeta.weekday.toLowerCase() !== planText(wdValue).slice(0, 3).toLowerCase()) {
        scheduleWeekdayMismatchCount += count;
      }
    }

    const monthlySlices = Array.from(monthlySliceMap.values()).filter((slice) => slice.scheduledFromGrid > 0);
    const totalSpots = scheduledFromGrid > 0
      ? scheduledFromGrid
      : Math.max(
          0,
          planCellNumber(row[layout.totalSpotsCol]),
          planCellNumber(row[layout.plannedSpotsCol])
        );

    const bonusFromCountColumn = Math.max(0, planCellNumber(row[layout.bonusSpotsCol]));
    const bonusCostValue = Math.max(0, planCellNumber(row[layout.bonusCostCol]));
    const summaryTotalSpots = Math.max(0, planCellNumber(row[layout.totalSpotsCol]));
    const summaryPlannedOrPaidSpots = Math.max(0, planCellNumber(row[layout.plannedSpotsCol]));
    const plannedOrPaidSpotsFromSummary = Math.max(
      0,
      summaryPlannedOrPaidSpots,
      Math.max(0, summaryTotalSpots - bonusFromCountColumn)
    );
    const inferredBonusFromCost =
      ratePerSpot > 0 && bonusCostValue > 0
        ? Math.max(0, Math.min(totalSpots, Math.round(bonusCostValue / ratePerSpot)))
        : 0;

    const inferredPaidSpotsFromGross =
      ratePerSpot > 0 && importedGross > 0
        ? Math.max(0, Math.min(totalSpots, Math.round(importedGross / ratePerSpot)))
        : -1;

    const inferredBonusFromGross =
      inferredPaidSpotsFromGross >= 0
        ? Math.max(0, Math.min(totalSpots, totalSpots - inferredPaidSpotsFromGross))
        : 0;

    const isRadioImport = String(mediumLabel || "").toLowerCase() === "radio";
    const withinOne = (a, b) => Math.abs((Number(a) || 0) - (Number(b) || 0)) <= 1;

    let bonusSpots = 0;

    if (isRadioImport) {
      // Radio sheets: the explicit Bonus Spots cell is the primary source of truth.
      // The earlier issue came from the bonus column collapsing into another summary column.
      // Once the radio bonus column is identified correctly, keep the workbook bonus value.
      if (bonusFromCountColumn > 0) {
        bonusSpots = bonusFromCountColumn;
      } else if (inferredBonusFromCost > 0 && inferredBonusFromGross > 0 && withinOne(inferredBonusFromCost, inferredBonusFromGross)) {
        bonusSpots = Math.max(inferredBonusFromCost, inferredBonusFromGross);
      } else if (inferredBonusFromCost > 0 && importedGross <= 0) {
        bonusSpots = inferredBonusFromCost;
      } else if (inferredBonusFromGross > 0 && bonusCostValue <= 0) {
        bonusSpots = inferredBonusFromGross;
      } else {
        bonusSpots = 0;
      }
    } else {
      if (bonusFromCountColumn > 0) {
        bonusSpots = bonusFromCountColumn;
      } else if (summaryPlannedOrPaidSpots > 0 && totalSpots > summaryPlannedOrPaidSpots) {
        bonusSpots = Math.max(0, totalSpots - summaryPlannedOrPaidSpots);
      } else if (plannedOrPaidSpotsFromSummary > 0 && totalSpots > plannedOrPaidSpotsFromSummary) {
        bonusSpots = Math.max(0, totalSpots - plannedOrPaidSpotsFromSummary);
      } else if (inferredBonusFromGross > 0) {
        bonusSpots = inferredBonusFromGross;
      } else if (inferredBonusFromCost > 0) {
        bonusSpots = inferredBonusFromCost;
      }
    }

    bonusSpots = Math.max(0, Math.min(bonusSpots, totalSpots));
    const hasImportedDiscountPct = planCellHasValue(row[layout.discountPctCol]);
    const hasImportedCommissionPct = planCellHasValue(row[layout.commissionPctCol]);
    const discountPct = hasImportedDiscountPct ? (planCellNumber(row[layout.discountPctCol]) || 0) : 0;
    const discountValue = planCellNumber(row[layout.discountValueCol]);
    const commissionPct = hasImportedCommissionPct ? (planCellNumber(row[layout.commissionPctCol]) || 0) : 0;
    const commissionValue = planCellNumber(row[layout.commissionValueCol]);
    const importedNet = planCellNumber(row[layout.netValueCol]);

    if (isCompositePlanSummaryOnlyRow({
      programmeValue,
      wdValue,
      timeValue,
      materialValue,
      durationValue,
      rateValue,
      scheduledFromGrid,
      totalSpots,
      importedGross,
      importedNet,
      bonusSpots,
    })) {
      continue;
    }

    const computedGross = Math.max(0, totalSpots - bonusSpots) * ratePerSpot;
    const rowErrors = [];
    const rowWarnings = [];

    if (!vendorName) rowErrors.push("No vendor header was detected above this row.");
    if (!programme) rowErrors.push("Programme name is missing.");
    if (!ratePerSpot) rowErrors.push("Rate per spot is missing.");
    if (!totalSpots) rowErrors.push("No scheduled spots were detected.");
    if (scheduledFromGrid > 0 && totalSpots !== scheduledFromGrid) {
      rowWarnings.push(`Calendar marks (${scheduledFromGrid}) differ from total spots (${totalSpots}).`);
    }
    if (scheduleWeekdayMismatchCount > 0) {
      rowWarnings.push(`Detected ${scheduleWeekdayMismatchCount} schedule mark(s) whose weekday does not match the WD column.`);
    }
    if (bonusSpots > totalSpots) {
      rowErrors.push("Bonus spots exceed scheduled spots.");
    }
    if (importedGross && Math.abs(importedGross - computedGross) > 1) {
      rowWarnings.push(`Gross value mismatch: workbook ${fmtN(importedGross)} vs computed ${fmtN(computedGross)}.`);
    }
    if (!bonusFromCountColumn && inferredBonusFromCost > 0) {
      rowWarnings.push(`Bonus spots inferred from cost-of-bonus value (${fmtN(bonusCostValue)}).`);
    }

    const match = findBestMediaPlanVendorMatchFromCandidates([vendorDisplayName, vendorName], vendors);
    const slicesToPersist = monthlySlices.length
      ? monthlySlices
      : [{ scheduleMonth: "", dayCounts: {}, scheduledFromGrid: totalSpots }];

    const sliceBonusAllocations = distributeBonusAcrossSlices(slicesToPersist, bonusSpots);

    slicesToPersist.forEach((slice, sliceIndex) => {
      const sliceSpots = slice.scheduledFromGrid || (sliceIndex === 0 ? totalSpots : 0);
      if (!sliceSpots) return;
      const sliceBonus = Math.min(sliceBonusAllocations[sliceIndex] || 0, sliceSpots);
      const slicePaid = Math.max(0, sliceSpots - sliceBonus);
      parsedRows.push({
        key: `${normalizeMediaPlanVendorName(vendorDisplayName || vendorName || programme)}-${rowIndex}-${slice.scheduleMonth || "all"}-${sliceIndex}`,
        sourceSheet: sourceSheet || "",
        sourceRow: rowIndex + 1,
        vendorName: vendorDisplayName || vendorName || match.vendor?.name || "",
        rawVendorName: vendorName || "",
        regionName,
        programme,
        weekday: wdValue,
        timeBelt: timeValue,
        material: materialValue,
        duration: normalizeDurationValue(durationValue, "30"),
        ratePerSpot,
        totalSpots: sliceSpots,
        bonusSpots: sliceBonus,
        paidSpots: slicePaid,
        dayCounts: slice.dayCounts || {},
        scheduleMonth: slice.scheduleMonth || "",
        importedGross: importedGross > 0 ? importedGross : (slicePaid * ratePerSpot),
        discountPct,
        hasImportedDiscountPct,
        discountValue,
        commissionPct,
        hasImportedCommissionPct,
        commissionValue,
        importedNet,
        vendorId: match.vendor?.id || "",
        vendorMatchType: match.matchType,
        medium: mediumLabel,
        errors: rowErrors,
        warnings: rowWarnings,
      });
    });

    warnings.push(...rowErrors.map(message => `Row ${rowIndex + 1}: ${message}`));
    warnings.push(...rowWarnings.map(message => `Row ${rowIndex + 1}: ${message}`));
  }

  const grouped = Object.values(parsedRows.reduce((acc, row) => {
    const key = normalizeMediaPlanVendorName(row.vendorName || buildMediaPlanVendorDisplayName(row.rawVendorName || "", row.regionName || "", row.medium || mediumLabel)) || `unmatched-${row.sourceRow}`;
    if (!acc[key]) {
      acc[key] = {
        id: key,
        vendorName: row.vendorName || "Unknown Vendor",
        regionName: row.regionName || "",
        vendorId: row.vendorId || "",
        vendorMatchType: row.vendorMatchType || "none",
        medium: row.medium || mediumLabel || "",
        rows: [],
        warnings: [],
        errors: [],
      };
    }
    acc[key].rows.push(row);
    acc[key].warnings.push(...row.warnings.map(message => `Row ${row.sourceRow}: ${message}`));
    acc[key].errors.push(...row.errors.map(message => `Row ${row.sourceRow}: ${message}`));
    return acc;
  }, {})).map(group => {
    const totals = group.rows.reduce((sum, row) => {
      sum.totalSpots += row.totalSpots;
      sum.totalBonusSpots += row.bonusSpots;
      sum.totalPaidSpots += row.paidSpots;
      sum.totalGross += row.importedGross || (row.paidSpots * row.ratePerSpot);
      return sum;
    }, { totalSpots: 0, totalBonusSpots: 0, totalPaidSpots: 0, totalGross: 0 });

    const explicitDiscPcts = group.rows
      .filter(row => row.hasImportedDiscountPct)
      .map(row => Number(row.discountPct) || 0);
    const explicitCommPcts = group.rows
      .filter(row => row.hasImportedCommissionPct)
      .map(row => Number(row.commissionPct) || 0);
    const uniqueDiscPcts = [...new Set(explicitDiscPcts.map(value => Number(value.toFixed(4))))];
    const uniqueCommPcts = [...new Set(explicitCommPcts.map(value => Number(value.toFixed(4))))];
    const hasImportedDiscountPct = explicitDiscPcts.length > 0 || footerPercents.discountPct !== null;
    const hasImportedCommissionPct = explicitCommPcts.length > 0 || footerPercents.commissionPct !== null;
    const discountPct = uniqueDiscPcts.length === 1 ? uniqueDiscPcts[0] : (footerPercents.discountPct ?? 0);
    const commissionPct = uniqueCommPcts.length === 1 ? uniqueCommPcts[0] : (footerPercents.commissionPct ?? 0);
    const discountValue = group.rows.length ? totals.totalGross * (discountPct / 100) : 0;
    const lessDisc = totals.totalGross - discountValue;
    const commissionValue = lessDisc * (commissionPct / 100);
    const importedNet = lessDisc - commissionValue;
    return {
      ...group,
      totalSpots: totals.totalSpots,
      totalBonusSpots: totals.totalBonusSpots,
      totalPaidSpots: totals.totalPaidSpots,
      totalGross: totals.totalGross,
      discountValue,
      commissionValue,
      importedNet,
      discountPct,
      commissionPct,
      hasImportedDiscountPct,
      hasImportedCommissionPct,
      blockedRowCount: group.rows.filter(row => (row.errors || []).length).length,
      hasBlockingErrors: group.rows.some(row => (row.errors || []).length > 0),
    };
  }).sort((a, b) => b.totalGross - a.totalGross);

  return {
    layout,
    parsedRows,
    blockedRows: parsedRows.filter(row => (row.errors || []).length > 0),
    groups: grouped,
    warnings,
  };
};

const normalizeProgrammeForMatch = (value = "") => normalizeMediaPlanVendorName(value).replace(/\([^)]*\)/g, " ").replace(/\s+/g, " ").trim();
const normalizeTimeBeltForMatch = (value = "") =>
  String(planText(value) || "")
    .toLowerCase()
    .replace(/[–—]/g, "-")
    .replace(/\s*-\s*/g, "-")
    .replace(/\s+/g, " ")
    .trim();
const applyAutoMatchedVendorRate = (row = {}, vendorRates = []) => {
  if (!row) return row;
  if (row.isComplimentary) return { ...row, rateId: "", customRate: "" };
  if (String(row.customRate ?? "").trim()) return row;
  const matchedRate = findMatchingVendorRateForImportRow(row, vendorRates);
  if (!matchedRate) return { ...row, rateId: "" };
  return {
    ...row,
    rateId: matchedRate.id || "",
    timeBelt: row.timeBelt || matchedRate.timeBelt || "",
    duration: row.duration || (matchedRate.duration ? String(matchedRate.duration) : row.duration) || "30",
  };
};
const findMatchingVendorRateForImportRow = (row = {}, vendorRates = []) => {
  const targetProgramme = normalizeProgrammeForMatch(row?.programme || "");
  const targetTime = normalizeTimeBeltForMatch(row?.timeBelt || "");
  const targetDuration = String(row?.duration || "").trim();
  if (!targetProgramme && !targetTime) return null;

  const exactTimeCandidates = (vendorRates || []).map((rate) => {
    const rateTime = normalizeTimeBeltForMatch(rate?.timeBelt || "");
    const rateDuration = String(rate?.duration || "").trim();
    const rateProgramme = normalizeProgrammeForMatch(rate?.programme || "");
    const exactTime = !!targetTime && !!rateTime && rateTime === targetTime;
    const exactProgramme = !!targetProgramme && !!rateProgramme && rateProgramme === targetProgramme;
    return { rate, exactTime, exactProgramme, rateTime, rateDuration };
  }).filter((item) => item.exactTime);

  if (exactTimeCandidates.length === 1) return exactTimeCandidates[0].rate || null;
  if (exactTimeCandidates.length > 1) {
    const exactDurationTimeMatches = exactTimeCandidates.filter(item => targetDuration && item.rateDuration && item.rateDuration === targetDuration);
    if (exactDurationTimeMatches.length === 1) return exactDurationTimeMatches[0].rate || null;

    const distinctRates = [...new Set(exactTimeCandidates.map(item => Number(item?.rate?.ratePerSpot) || 0))];
    if (distinctRates.length === 1) return exactTimeCandidates[0].rate || null;

    const exactProgrammeTimeMatches = exactTimeCandidates.filter(item => item.exactProgramme);
    if (exactProgrammeTimeMatches.length === 1) return exactProgrammeTimeMatches[0].rate || null;

    return exactTimeCandidates[0]?.rate || null;
  }

  const candidates = (vendorRates || []).map((rate) => {
    const rateProgramme = normalizeProgrammeForMatch(rate?.programme || "");
    const rateTime = normalizeTimeBeltForMatch(rate?.timeBelt || "");
    const rateDuration = String(rate?.duration || "").trim();
    const exactProgramme = !!rateProgramme && rateProgramme === targetProgramme;
    const fuzzyProgramme = !!rateProgramme && !exactProgramme && (rateProgramme.includes(targetProgramme) || targetProgramme.includes(rateProgramme));
    const exactTime = !!targetTime && !!rateTime && rateTime === targetTime;
    let score = 0;
    if (exactTime) score += 120;
    if (exactProgramme) score += 100;
    else if (fuzzyProgramme) score += 60;
    if (targetDuration && rateDuration && targetDuration === rateDuration) score += 10;
    return { rate, score, exactProgramme, fuzzyProgramme, exactTime, rateProgramme, rateTime, rateDuration };
  }).filter(item => item.exactProgramme || item.fuzzyProgramme);

  if (!candidates.length) return null;

  if (targetTime) {
    const exactTimeMatches = candidates.filter(item => item.exactTime);
    if (exactTimeMatches.length === 1) return exactTimeMatches[0].rate || null;
    if (exactTimeMatches.length > 1) {
      const exactProgrammeTimeMatches = exactTimeMatches.filter(item => item.exactProgramme);
      if (exactProgrammeTimeMatches.length === 1) return exactProgrammeTimeMatches[0].rate || null;
      const exactDurationTimeMatches = exactTimeMatches.filter(item => targetDuration && item.rateDuration && item.rateDuration === targetDuration);
      if (exactDurationTimeMatches.length === 1) return exactDurationTimeMatches[0].rate || null;
      const sortedTimeMatches = [...exactTimeMatches].sort((a, b) => b.score - a.score);
      if (sortedTimeMatches.length && (sortedTimeMatches.length === 1 || sortedTimeMatches[0].score > sortedTimeMatches[1].score)) {
        return sortedTimeMatches[0].rate || null;
      }
      return sortedTimeMatches[0]?.rate || null;
    }
  }

  const exactProgrammeMatches = candidates.filter(item => item.exactProgramme);
  if (exactProgrammeMatches.length === 1) return exactProgrammeMatches[0].rate || null;
  if (exactProgrammeMatches.length > 1) {
    const exactDurationMatches = exactProgrammeMatches.filter(item => targetDuration && item.rateDuration && item.rateDuration === targetDuration);
    if (exactDurationMatches.length === 1) return exactDurationMatches[0].rate || null;
    const exactTimeMatches = exactProgrammeMatches.filter(item => item.exactTime);
    if (exactTimeMatches.length === 1) return exactTimeMatches[0].rate || null;
    const sortedExact = [...exactProgrammeMatches].sort((a, b) => b.score - a.score);
    if (sortedExact.length && (sortedExact.length === 1 || sortedExact[0].score > sortedExact[1].score)) return sortedExact[0].rate || null;
    return sortedExact[0]?.rate || null;
  }

  const fuzzyProgrammeMatches = [...candidates].sort((a, b) => b.score - a.score);
  if (!fuzzyProgrammeMatches.length) return null;
  if (fuzzyProgrammeMatches.length === 1) return fuzzyProgrammeMatches[0].rate || null;
  if (fuzzyProgrammeMatches[0].score > fuzzyProgrammeMatches[1].score) return fuzzyProgrammeMatches[0].rate || null;
  return fuzzyProgrammeMatches[0]?.rate || null;
};

const buildImportedMpoRecord = ({ group, campaign, client, vendor, rates = [], user, appSettings, months, year, mpoNo, importBatchId = "", sourceFileName = "" }) => {
  const selectedMonths = Array.isArray(months) ? months : [];
  const scheduleLabel = selectedMonths.length ? `${selectedMonths.join(", ")} ${year}`.trim() : `${year}`.trim();
  const roundMoney = (value) => roundMoneyValue(value, appSettings);
  const vendorRatesForImport = activeOnly(rates).filter(rate => rate.vendorId === vendor?.id);
  const matchedRates = group.rows.map((row) => findMatchingVendorRateForImportRow(row, vendorRatesForImport));
  const normalizedRows = group.rows.map((row, rowIndex) => {
    const matchedRate = matchedRates[rowIndex];
    const resolvedProgramme = matchedRate?.programme || row.programme;
    const resolvedTimeBelt = matchedRate?.timeBelt || row.timeBelt;
    const resolvedDuration = matchedRate?.duration ? String(matchedRate.duration) : (row.duration || "30");
    const resolvedRatePerSpot = Number(matchedRate?.ratePerSpot) || Number(row.ratePerSpot) || 0;
    return normalizeBonusAdjustedSpot({
      id: uid(),
      programme: resolvedProgramme,
      wd: row.weekday,
      timeBelt: resolvedTimeBelt,
      material: row.material,
      duration: resolvedDuration,
      rateId: matchedRate?.id || "",
      ratePerSpot: resolvedRatePerSpot,
      customRate: resolvedRatePerSpot ? String(resolvedRatePerSpot) : "",
      scheduleMonth: row.scheduleMonth ? `${row.scheduleMonth} ${year}`.trim() : scheduleLabel,
      calendarDays: expandCountsToDays(row.dayCounts || {}),
      isComplimentary: false,
      sourceSheet: row.sourceSheet || "",
      sourceRow: row.sourceRow,
      importVendorLabel: row.vendorName || group.vendorName || "",
      importRegionLabel: row.regionName || group.regionName || "",
      importWarnings: [...(row.errors || []), ...(row.warnings || [])],
      importBatchId,
      sourceFileName,
    }, row.totalSpots, row.bonusSpots);
  });

  const totalSpots = normalizedRows.reduce((sum, row) => sum + (Number(row.spots) || 0), 0);
  const totalBonusSpots = normalizedRows.reduce((sum, row) => sum + getSpotBonusCount(row), 0);
  const totalPaidSpots = normalizedRows.reduce((sum, row) => sum + getPaidSpotCount(row), 0);
  const totalGross = roundMoney(normalizedRows.reduce((sum, row) => sum + (getPaidSpotCount(row) * (Number(row.ratePerSpot) || 0)), 0));
  const matchedRateDiscounts = matchedRates
    .map((rate) => (rate && planCellHasValue(rate.discount) ? Number(rate.discount) || 0 : null))
    .filter((value) => value !== null);
  const uniqueMatchedRateDiscounts = [...new Set(matchedRateDiscounts.map((value) => Number(value.toFixed(4))))];
  const matchedRateDiscountPct = uniqueMatchedRateDiscounts.length === 1 ? uniqueMatchedRateDiscounts[0] / 100 : null;
  const vendorDiscountPct = planCellHasValue(vendor?.discount) ? ((Number(vendor.discount) || 0) / 100) : 0;
  const importedDiscountPct = group.hasImportedDiscountPct ? ((Number(group.discountPct) || 0) / 100) : null;
  const mediumForDiscount = String(group.medium || campaign?.medium || '').toLowerCase();
  const preferVendorDiscount = mediumForDiscount === 'radio';
  const discountPct = preferVendorDiscount
    ? vendorDiscountPct
    : (matchedRateDiscountPct ?? importedDiscountPct ?? vendorDiscountPct);
  const discountValue = roundMoney(totalGross * discountPct);
  const lessDisc = roundMoney(totalGross - discountValue);
  const commPct = group.hasImportedCommissionPct ? ((Number(group.commissionPct) || 0) / 100) : (vendor?.commission ? Number(vendor.commission) / 100 : 0);
  const commAmt = roundMoney(lessDisc * commPct);
  const afterComm = roundMoney(lessDisc - commAmt);
  const netVal = Number(group.importedNet) > 0 ? roundMoney(group.importedNet) : afterComm;
  const vatPct = Number(appSettings?.vatRate) || 7.5;
  const vatAmt = roundMoney(netVal * vatPct / 100);
  const grandTotal = roundMoney(netVal + vatAmt);

  return {
    id: uid(),
    campaignId: campaign?.id || "",
    vendorId: vendor?.id || "",
    mpoNo,
    date: new Date().toISOString().slice(0, 10),
    month: selectedMonths[0] || "",
    months: selectedMonths,
    year: String(year || new Date().getFullYear()),
    medium: group.medium || campaign?.medium || "",
    signedBy: "",
    signedTitle: "",
    preparedBy: user?.name || "",
    preparedContact: user?.phone || user?.email || "",
    preparedTitle: user?.title || "",
    preparedSignature: user?.signatureDataUrl || "",
    signedSignature: "",
    agencyAddress: user?.agencyAddress || "5, Craig Street, Ogudu GRA, Lagos",
    agencyEmail: user?.agencyEmail || "",
    agencyPhone: user?.agencyPhone || "",
    transmitMsg: `PLEASE TRANSMIT SPOTS ON ${vendor?.name || group.vendorName} AS SCHEDULED`,
    status: "draft",
    vendorName: vendor?.name || group.vendorName || "",
    clientName: client?.name || "",
    campaignName: campaign?.name || "",
    brand: campaign?.brand || "",
    spots: normalizedRows,
    totalSpots,
    totalBonusSpots,
    totalPaidSpots,
    totalGross,
    discPct: discountPct,
    discAmt: discountValue,
    lessDisc,
    commPct,
    commAmt,
    afterComm,
    surchPct: 0,
    surchAmt: 0,
    surchLabel: "",
    netVal,
    vatPct,
    vatAmt,
    grandTotal,
    terms: appSettings?.mpoTerms || DEFAULT_APP_SETTINGS.mpoTerms,
    roundToWholeNaira: !!appSettings?.roundToWholeNaira,
    importBatchId,
    sourceFileName,
    importSummary: {
      source: "media_plan_upload",
      vendorLabel: group.vendorName || "",
      regionLabel: group.regionName || "",
      rowCount: group.rows.length,
      warningCount: group.warnings.length,
      blockedRowCount: group.blockedRowCount || 0,
      discountSource: (String(group.medium || campaign?.medium || '').toLowerCase() === 'radio') ? 'vendor' : 'programme_or_import',
    },
  };
};

const MediaPlanImportModal = ({ vendors = [], clients = [], campaigns = [], rates = [], user, appSettings, setMpos, setVendors = () => {}, onClose, onToast, onBulkImportStateChange = () => {}, requestMpoRefresh = async () => {} }) => {
  const [step, setStep] = useState("setup");
  const [campaignId, setCampaignId] = useState("");
  const [year, setYear] = useState(String(new Date().getFullYear()));
  const [selectedMonths, setSelectedMonths] = useState([]);
  const [fileName, setFileName] = useState("");
  const [loading, setLoading] = useState(false);
  const [saving, setSaving] = useState(false);
  const [importWarnings, setImportWarnings] = useState([]);
  const [parsedRows, setParsedRows] = useState([]);
  const [blockedRows, setBlockedRows] = useState([]);
  const [groups, setGroups] = useState([]);
  const [groupSelections, setGroupSelections] = useState({});
  const [vendorAssignments, setVendorAssignments] = useState({});
  const [expandedGroupId, setExpandedGroupId] = useState("");
  const [summary, setSummary] = useState({ parsedRows: 0, parsedSpots: 0, parsedBonusSpots: 0, vendorGroups: 0, blockedRows: 0 });
  const [autoCreateMissingVendors, setAutoCreateMissingVendors] = useState(true);
  const [importProgress, setImportProgress] = useState({ current: 0, total: 0, currentVendor: "", phase: "" });
  const fileRef = useRef(null);

  const isTransientImportError = (error) => {
    const message = String(error?.message || error || "").toLowerCase();
    return message.includes("gateway timeout")
      || message.includes("upstream request timeout")
      || message.includes("timeout")
      || message.includes("504");
  };

  const waitForImportRetry = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

  const runImportStepWithRetry = async (task, attempts = 3) => {
    let lastError = null;
    for (let attempt = 1; attempt <= attempts; attempt += 1) {
      try {
        return await task();
      } catch (error) {
        lastError = error;
        if (attempt >= attempts || !isTransientImportError(error)) throw error;
        await waitForImportRetry(400 * attempt);
      }
    }
    throw lastError;
  };

  const isDuplicateMpoNumberError = (error) => {
    const message = String(error?.message || error || "").toLowerCase();
    const code = String(error?.code || "").toLowerCase();
    return code === "23505"
      || message.includes("mpos_agency_mpo_no_unique")
      || message.includes("duplicate key value");
  };

  const createImportedDraftWithRecovery = async ({ group, campaign, client, vendor, importBatchId }) => {
    let lastError = null;

    for (let attempt = 1; attempt <= 3; attempt += 1) {
      const mpoNo = await runImportStepWithRetry(
        () => generateNextMpoNoFromSupabase(campaign?.brand || "MPO")
      );

      const record = buildImportedMpoRecord({
        group,
        campaign,
        client,
        vendor,
        rates,
        user,
        appSettings,
        months: selectedMonths,
        year,
        mpoNo,
        importBatchId,
        sourceFileName: fileName || "",
      });

      try {
        const saved = await runImportStepWithRetry(
          () => createMpoInSupabase(user?.agencyId, user?.id, record)
        );
        return { saved, mpoNo };
      } catch (error) {
        lastError = error;

        if (isTransientImportError(error) || isDuplicateMpoNumberError(error)) {
          try {
            const existing = await fetchMappedMpoByAgencyAndNo(user?.agencyId, mpoNo);
            if (existing) return { saved: existing, mpoNo };
          } catch (lookupError) {
            console.error("Failed to look up import MPO after create retry/conflict:", lookupError);
          }
        }

        if (attempt >= 3 || !isDuplicateMpoNumberError(error)) {
          throw error;
        }
      }
    }

    throw lastError;
  };

  const toggleMonth = (monthValue) => {
    setSelectedMonths((current) => current.includes(monthValue)
      ? current.filter((item) => item !== monthValue)
      : PLAN_IMPORT_MONTHS.filter((item) => [...current, monthValue].includes(item)));
  };


  const handleFile = async (file) => {
    if (!file) return;
    setLoading(true);
    setFileName(file.name || "media-plan.xlsx");
    try {
      const XLSX = await loadSheetJS();
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, {
        type: "array",
        raw: true,
        cellDates: false,
        cellFormula: true,
      });

      const parseCandidates = [];
      const parseFailures = [];

      for (const sheetName of workbook.SheetNames || []) {
        try {
          const sheet = workbook.Sheets[sheetName];
          const sheetRows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: "" });
          const parsed = parseCompositeMediaPlanRows(sheetRows, activeOnly(vendors), sheetName || "Sheet", year);
          if ((parsed?.parsedRows || []).length > 0) {
            parseCandidates.push({
              sheetName,
              parsed,
              score:
                ((parsed?.groups || []).length * 20)
                + ((parsed?.parsedRows || []).length * 5)
                - ((parsed?.blockedRows || []).length * 3)
                - ((parsed?.warnings || []).length * 0.1),
            });
          }
        } catch (error) {
          parseFailures.push(`${sheetName}: ${error?.message || "Could not parse this sheet."}`);
        }
      }

      if (!parseCandidates.length) {
        throw new Error(
          parseFailures.length
            ? `Failed to parse this workbook. ${parseFailures.slice(0, 3).join(" | ")}`
            : "Failed to parse this media plan workbook."
        );
      }

      parseCandidates.sort((a, b) => b.score - a.score);
      const best = parseCandidates[0].parsed;

      setParsedRows(best.parsedRows || []);
      setBlockedRows(best.blockedRows || []);
      setImportWarnings([
        ...(best.warnings || []),
        ...(parseFailures.length ? [`Skipped ${parseFailures.length} sheet(s): ${parseFailures.slice(0, 2).join(" | ")}`] : []),
      ]);
      setGroups(best.groups || []);
      setSummary({
        parsedRows: best.parsedRows.length,
        parsedSpots: (best.parsedRows || []).reduce((sum, row) => sum + (Number(row.totalSpots) || 0), 0),
        parsedBonusSpots: (best.parsedRows || []).reduce((sum, row) => sum + (Number(row.bonusSpots) || 0), 0),
        vendorGroups: best.groups.length,
        blockedRows: (best.blockedRows || []).length,
      });
      setGroupSelections(Object.fromEntries((best.groups || []).map((group) => [group.id, true])));
      setVendorAssignments(Object.fromEntries((best.groups || []).map((group) => [group.id, group.vendorId || ""])));
      setExpandedGroupId(best.groups[0]?.id || "");
      setStep("review");
    } catch (error) {
      console.error("Failed to parse media plan:", error);
      setParsedRows([]);
      setBlockedRows([]);
      setImportWarnings([error?.message || "Failed to parse this media plan workbook."]);
      setGroups([]);
      setSummary({ parsedRows: 0, parsedSpots: 0, parsedBonusSpots: 0, vendorGroups: 0, blockedRows: 0 });
      setStep("review");
    } finally {
      setLoading(false);
    }
  };


  const downloadExceptionReport = () => {
    const exceptionRows = (blockedRows || []).map((row) => [
      fileName || "",
      row.sourceSheet || "",
      row.sourceRow || "",
      row.vendorName || "",
      row.regionName || "",
      row.programme || "",
      row.weekday || "",
      row.timeBelt || "",
      row.material || "",
      row.duration || "",
      row.ratePerSpot || 0,
      row.totalSpots || 0,
      row.bonusSpots || 0,
      row.paidSpots || 0,
      (row.errors || []).join(" | "),
      (row.warnings || []).join(" | "),
    ]);
    if (!exceptionRows.length) {
      onToast?.({ msg: "No blocked rows to export.", type: "error" });
      return;
    }
    downloadMediaPlanImportCsv(
      `${(fileName || "media-plan").replace(/\.[^.]+$/, "")}-import-exceptions.csv`,
      ["Source File","Sheet","Row","Vendor Label","Region","Programme","WD","Time","Material","Duration","Rate","Scheduled Spots","Bonus Spots","Paid Spots","Errors","Warnings"],
      exceptionRows
    );
  };

  const handleSave = async () => {
    const campaign = campaigns.find((item) => item.id === campaignId);
    if (!campaign) {
      onToast?.({ msg: "Select a campaign before saving MPO drafts.", type: "error" });
      return;
    }
    if (!selectedMonths.length) {
      onToast?.({ msg: "Select at least one campaign month for the imported MPO drafts.", type: "error" });
      return;
    }

    const activeGroups = groups.filter((group) => groupSelections[group.id]);
    if (!activeGroups.length) {
      onToast?.({ msg: "Select at least one detected vendor group to import.", type: "error" });
      return;
    }

    const blockedGroups = activeGroups.filter((group) => group.hasBlockingErrors);
    if (blockedGroups.length) {
      onToast?.({ msg: `${blockedGroups.length} selected group${blockedGroups.length !== 1 ? "s have" : " has"} blocking row errors. Export the exception report and fix those rows before saving.`, type: "error" });
      return;
    }

    try {
      setSaving(true);
      onBulkImportStateChange(true);
      setImportProgress({ current: 0, total: activeGroups.length, currentVendor: "", phase: "Preparing import…" });
      const createdRecords = [];
      const client = campaign ? clients?.find?.((item) => item.id === campaign.clientId) : null;
      const importBatchId = `mpi_${Date.now().toString(36)}`;
      const workingAssignments = { ...vendorAssignments };
      const refreshedVendors = [...activeOnly(vendors)];
      const autoCreatedVendors = [];

      if (autoCreateMissingVendors) {
        for (const group of activeGroups.filter((item) => !workingAssignments[item.id] && item.vendorName)) {
          const ensuredVendor = await ensureVendorExistsInSupabase(user?.agencyId, user?.id, group.vendorName, {
            type: group.medium || campaign?.medium || "Television",
            discount: group.discountPct || "",
            commission: group.commissionPct || "",
            notes: `Auto-created from media plan import (${fileName || "media plan"}).`,
          });
          if (ensuredVendor?.id) {
            workingAssignments[group.id] = ensuredVendor.id;
            if (!refreshedVendors.some((item) => item.id === ensuredVendor.id)) {
              refreshedVendors.unshift(ensuredVendor);
              autoCreatedVendors.push(ensuredVendor);
            }
          }
        }
      }

      const missingVendorGroups = activeGroups.filter((group) => !workingAssignments[group.id]);
      if (missingVendorGroups.length) {
        onToast?.({ msg: `Assign a vendor to ${missingVendorGroups.length} group${missingVendorGroups.length !== 1 ? "s" : ""} before saving.`, type: "error" });
        setVendorAssignments(workingAssignments);
        onBulkImportStateChange(false);
        setSaving(false);
        return;
      }

      for (let groupIndex = 0; groupIndex < activeGroups.length; groupIndex += 1) {
        const group = activeGroups[groupIndex];
        const vendor = refreshedVendors.find((item) => item.id === workingAssignments[group.id]);
        setImportProgress({
          current: Math.min(groupIndex + 1, activeGroups.length),
          total: activeGroups.length,
          currentVendor: group.vendorName || vendor?.name || "",
          phase: vendor ? "Creating draft MPO…" : "Resolving vendor…",
        });
        if (!vendor) continue;
        const { saved, mpoNo } = await createImportedDraftWithRecovery({
          group,
          campaign,
          client,
          vendor,
          importBatchId,
        });
        createAuditEventInSupabase({
          agencyId: user?.agencyId,
          recordType: "mpo",
          recordId: saved.id,
          action: "created",
          actor: user,
          note: `Imported from media plan: ${group.vendorName || vendor?.name || "Vendor"}`,
          metadata: {
            mpoNo: saved.mpoNo || mpoNo,
            status: "draft",
            importSource: "media_plan_upload",
            importBatchId,
            sourceFileName: fileName || "",
            rowCount: group.rows.length,
            vendorLabel: group.vendorName || "",
            regionLabel: group.regionName || "",
            warningCount: group.warnings.length,
          },
        }).catch((error) => console.error("Failed to write MPO import audit event:", error));
        createdRecords.push(saved);
      }
      if (createdRecords.length) {
        setMpos((items) => [...createdRecords, ...items]);
      }
      if (autoCreatedVendors.length) {
        setVendors?.((items) => {
          const existing = Array.isArray(items) ? items : [];
          const map = new Map(existing.map((item) => [item.id, item]));
          autoCreatedVendors.forEach((vendorItem) => map.set(vendorItem.id, vendorItem));
          return Array.from(map.values());
        });
      }
      setVendorAssignments(workingAssignments);
      setImportProgress({
        current: createdRecords.length,
        total: activeGroups.length,
        currentVendor: "",
        phase: "Completed",
      });
      onToast?.({
        msg: `${createdRecords.length} MPO draft${createdRecords.length !== 1 ? "s" : ""} created from ${fileName || "media plan"}${autoCreatedVendors.length ? ` · ${autoCreatedVendors.length} vendor${autoCreatedVendors.length !== 1 ? "s" : ""} auto-created` : ""}.`,
        type: "success",
      });
      onClose?.();
    } catch (error) {
      console.error("Failed to create MPO drafts from media plan:", error);
      onToast?.({ msg: error?.message || "Failed to create MPO drafts from this media plan.", type: "error" });
    } finally {
      onBulkImportStateChange(false);
      setSaving(false);
      setImportProgress((state) => ({ ...state, currentVendor: state.phase === "Completed" ? state.currentVendor : "", phase: state.phase === "Completed" ? state.phase : "" }));
    }
  };

  const activeVendors = activeOnly(vendors);

  return (
    <Modal title="Import Media Plan → Draft MPOs" onClose={onClose} width={1080}>
      <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
          <Badge color={step === "setup" ? "accent" : "blue"}>1. Setup</Badge>
          <Badge color={step === "review" ? "accent" : "blue"}>2. Review & Save</Badge>
        </div>

        {step === "setup" && (
          <>
            <Card style={{ padding: 16 }}>
              <div style={{ display: "grid", gridTemplateColumns: "1.25fr .75fr", gap: 14 }}>
                <Field
                  label="Campaign"
                  value={campaignId}
                  onChange={setCampaignId}
                  options={campaigns.map((campaign) => ({ value: campaign.id, label: campaign.name }))}
                  placeholder="Select campaign"
                />
                <Field
                  label="Year"
                  value={year}
                  onChange={setYear}
                  placeholder="2026"
                />
              </div>
              <div style={{ marginTop: 14 }}>
                <div style={{ fontSize: 11, fontWeight: 700, color: "var(--text3)", textTransform: "uppercase", letterSpacing: ".08em", marginBottom: 8 }}>
                  Campaign Month(s)
                </div>
                <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                  {PLAN_IMPORT_MONTHS.map((monthValue) => {
                    const selected = selectedMonths.includes(monthValue);
                    return (
                      <button
                        key={monthValue}
                        type="button"
                        onClick={() => toggleMonth(monthValue)}
                        style={{
                          padding: "6px 12px",
                          borderRadius: 999,
                          border: selected ? "2px solid var(--accent)" : "1px solid var(--border2)",
                          background: selected ? "rgba(240,165,0,.14)" : "var(--bg3)",
                          color: selected ? "var(--accent)" : "var(--text2)",
                          fontFamily: "'Syne',sans-serif",
                          fontWeight: 700,
                          fontSize: 11,
                          cursor: "pointer",
                        }}
                      >
                        {monthValue.slice(0, 3)}
                      </button>
                    );
                  })}
                </div>
              </div>
            </Card>

            <Card style={{ padding: 16 }}>
              <div style={{ display: "flex", justifyContent: "space-between", gap: 12, alignItems: "center", flexWrap: "wrap" }}>
                <div>
                  <div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 15 }}>Upload Composite Media Plan</div>
                  <div style={{ marginTop: 4, fontSize: 12, color: "var(--text2)" }}>
                    The importer will detect vendor blocks, translate schedule rows into MPO-ready rows, keep bonus spots on the same line, and create draft MPOs after review.
                  </div>
                </div>
                <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
                  <Btn variant="ghost" onClick={() => fileRef.current?.click()}>Choose file</Btn>
                  <input
                    ref={fileRef}
                    type="file"
                    accept=".xlsx,.xls"
                    onChange={(event) => handleFile(event.target.files?.[0])}
                    style={{ display: "none" }}
                  />
                </div>
              </div>
              {fileName ? (
                <div style={{ marginTop: 10, fontSize: 12, color: "var(--green)" }}>Loaded file: {fileName}</div>
              ) : null}
              {loading ? <div style={{ marginTop: 10, fontSize: 12, color: "var(--text2)" }}>Parsing workbook…</div> : null}
            </Card>
          </>
        )}

        {step === "review" && (
          <>
            <Card style={{ padding: 16 }}>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(180px,1fr))", gap: 12 }}>
                <Stat icon="📄" label="Parsed Rows" value={summary.parsedRows} sub="Executable schedule rows detected" color="var(--accent)" />
                <Stat icon="🔢" label="Parsed Spots" value={summary.parsedSpots || 0} sub="Total scheduled spots detected from the workbook" color="var(--teal)" />
                <Stat icon="🎁" label="Bonus Spots" value={summary.parsedBonusSpots || 0} sub="Total bonus spots detected from the workbook" color="var(--purple)" />
                <Stat icon="🏢" label="Vendor Groups" value={summary.vendorGroups} sub="Draft MPO candidates detected" color="var(--blue)" />
                <Stat icon="⚠️" label="Warnings" value={importWarnings.length} sub="Rows that need manual review" color="var(--orange)" />
                <Stat icon="⛔" label="Blocked Rows" value={summary.blockedRows || 0} sub="Rows with hard-stop import errors" color="var(--red)" />
                <Stat icon="✅" label="Ready Groups" value={groups.filter((group) => vendorAssignments[group.id] && !group.hasBlockingErrors).length} sub="Groups with vendor mappings and no blockers" color="var(--green)" />
              </div>
              {saving && importProgress.total > 0 ? (
                <div style={{ marginTop: 14, border: "1px solid rgba(59,126,245,.28)", background: "rgba(59,126,245,.08)", borderRadius: 12, padding: "14px 16px" }}>
                  <div style={{ display: "flex", justifyContent: "space-between", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
                    <div>
                      <div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 14 }}>
                        Creating {Math.min(importProgress.current || 0, importProgress.total || 0)} of {importProgress.total}…
                      </div>
                      <div style={{ marginTop: 4, fontSize: 12, color: "var(--text2)" }}>
                        {importProgress.phase || "Creating draft MPOs…"}{importProgress.currentVendor ? ` · Current vendor: ${importProgress.currentVendor}` : ""}
                      </div>
                    </div>
                    <Badge color="blue">
                      {importProgress.total > 0 ? `${Math.round(((Math.min(importProgress.current || 0, importProgress.total || 0)) / importProgress.total) * 100)}%` : "0%"}
                    </Badge>
                  </div>
                  <div style={{ marginTop: 12, height: 10, borderRadius: 999, background: "rgba(255,255,255,.55)", overflow: "hidden" }}>
                    <div
                      style={{
                        width: `${importProgress.total > 0 ? ((Math.min(importProgress.current || 0, importProgress.total || 0) / importProgress.total) * 100) : 0}%`,
                        height: "100%",
                        background: "linear-gradient(90deg, var(--blue), var(--accent))",
                        transition: "width .25s ease",
                      }}
                    />
                  </div>
                </div>
              ) : null}

              {importWarnings.length ? (
                <div style={{ marginTop: 14, border: "1px solid var(--border)", background: "var(--bg3)", borderRadius: 10, padding: "12px 14px", maxHeight: 170, overflowY: "auto" }}>
                  <div style={{ fontSize: 11, fontWeight: 700, color: "var(--text3)", textTransform: "uppercase", letterSpacing: ".08em", marginBottom: 6 }}>Import warnings</div>
                  <div style={{ display: "flex", flexDirection: "column", gap: 5 }}>
                    {importWarnings.slice(0, 30).map((warning, index) => <div key={index} style={{ fontSize: 12, color: "var(--text2)" }}>{warning}</div>)}
                    {importWarnings.length > 30 ? <div style={{ fontSize: 11, color: "var(--text3)" }}>Showing first 30 warnings…</div> : null}
                  </div>
                </div>
              ) : null}
              <div style={{ marginTop: 14, display: "flex", justifyContent: "space-between", gap: 10, flexWrap: "wrap", alignItems: "center" }}>
                <label style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 12, color: "var(--text2)" }}>
                  <input type="checkbox" checked={autoCreateMissingVendors} onChange={(event) => setAutoCreateMissingVendors(event.target.checked)} style={{ accentColor: "var(--accent)", width: 16, height: 16 }} />
                  Auto-create unmatched vendors during save
                </label>
                <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                  <Btn variant="ghost" size="sm" onClick={downloadExceptionReport}>Download Exception Report</Btn>
                  <Btn variant="ghost" size="sm" onClick={() => setStep("setup")}>Back to Setup</Btn>
                </div>
              </div>
            </Card>

            <div style={{ display: "flex", flexDirection: "column", gap: 12, maxHeight: "56vh", overflowY: "auto", paddingRight: 4 }}>
              {groups.length === 0 ? (
                <Empty icon="📁" title="No MPO draft candidates detected" sub="Try another workbook or return to setup and upload again." />
              ) : groups.map((group, groupIndex) => {
                const assignedVendorId = vendorAssignments[group.id] || "";
                const expanded = expandedGroupId === group.id;
                return (
                  <Card key={`${group.id || "group"}-${groupIndex}`} style={{ padding: 16 }}>
                    <div style={{ display: "grid", gridTemplateColumns: "auto 1fr auto", gap: 14, alignItems: "flex-start" }}>
                      <input
                        type="checkbox"
                        checked={!!groupSelections[group.id]}
                        onChange={(event) => setGroupSelections((state) => ({ ...state, [group.id]: event.target.checked }))}
                        style={{ marginTop: 4, width: 16, height: 16, accentColor: "var(--accent)" }}
                      />
                      <div>
                        <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center" }}>
                          <div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 15 }}>{group.vendorName}</div>
                          {group.regionName ? <Badge color="blue">{group.regionName}</Badge> : null}
                          <Badge color={group.vendorMatchType === "exact" ? "green" : group.vendorMatchType === "fuzzy" ? "orange" : "red"}>
                            {group.vendorMatchType === "exact" ? "Exact vendor match" : group.vendorMatchType === "fuzzy" ? "Fuzzy vendor match" : "Vendor needs mapping"}
                          </Badge>
                          {group.hasBlockingErrors ? <Badge color="red">Blocked rows: {group.blockedRowCount}</Badge> : null}
                        </div>
                        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(140px,1fr))", gap: 10, marginTop: 10 }}>
                          <div style={{ background: "var(--bg3)", borderRadius: 10, padding: "10px 12px" }}>
                            <div style={{ fontSize: 10, color: "var(--text3)", textTransform: "uppercase" }}>Rows</div>
                            <div style={{ fontWeight: 800, fontSize: 16, marginTop: 4 }}>{group.rows.length}</div>
                          </div>
                          <div style={{ background: "var(--bg3)", borderRadius: 10, padding: "10px 12px" }}>
                            <div style={{ fontSize: 10, color: "var(--text3)", textTransform: "uppercase" }}>Scheduled Spots</div>
                            <div style={{ fontWeight: 800, fontSize: 16, marginTop: 4 }}>{group.totalSpots}</div>
                          </div>
                          <div style={{ background: "var(--bg3)", borderRadius: 10, padding: "10px 12px" }}>
                            <div style={{ fontSize: 10, color: "var(--text3)", textTransform: "uppercase" }}>Bonus Spots</div>
                            <div style={{ fontWeight: 800, fontSize: 16, marginTop: 4 }}>{group.totalBonusSpots}</div>
                          </div>
                          <div style={{ background: "var(--bg3)", borderRadius: 10, padding: "10px 12px" }}>
                            <div style={{ fontSize: 10, color: "var(--text3)", textTransform: "uppercase" }}>Paid Spots</div>
                            <div style={{ fontWeight: 800, fontSize: 16, marginTop: 4 }}>{group.totalPaidSpots}</div>
                          </div>
                          <div style={{ background: "var(--bg3)", borderRadius: 10, padding: "10px 12px" }}>
                            <div style={{ fontSize: 10, color: "var(--text3)", textTransform: "uppercase" }}>Imported Gross</div>
                            <div style={{ fontWeight: 800, fontSize: 16, marginTop: 4 }}>{fmtN(group.totalGross)}</div>
                          </div>
                        </div>
                        <div style={{ marginTop: 12 }}>
                          <Field
                            label="Vendor mapping"
                            value={assignedVendorId}
                            onChange={(value) => setVendorAssignments((state) => ({ ...state, [group.id]: value }))}
                            options={activeVendors.map((vendorItem) => ({ value: vendorItem.id, label: vendorItem.name }))}
                            placeholder={autoCreateMissingVendors && !assignedVendorId ? "Will auto-create on save if unmatched" : "Select matched vendor"}
                          />
                        </div>
                        {group.errors.length ? (
                          <div style={{ marginTop: 10, fontSize: 12, color: "var(--red)" }}>
                            {group.errors.slice(0, 2).join(" · ")}
                          </div>
                        ) : null}
                        {group.warnings.length ? (
                          <div style={{ marginTop: 6, fontSize: 12, color: "var(--orange)" }}>
                            {group.warnings.slice(0, 3).join(" · ")}
                          </div>
                        ) : null}
                      </div>
                      <Btn variant="ghost" size="sm" onClick={() => setExpandedGroupId((current) => current === group.id ? "" : group.id)}>
                        {expanded ? "Hide rows" : "Show rows"}
                      </Btn>
                    </div>
                    {expanded ? (
                      <div style={{ marginTop: 14, overflowX: "auto" }}>
                        <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 860 }}>
                          <thead>
                            <tr style={{ background: "var(--bg3)" }}>
                              {["Row", "Programme", "WD", "Time", "Material", "Rate", "Total", "Bonus", "Paid", "Gross", "Issues"].map((heading) => (
                                <th key={heading} style={{ padding: "7px 8px", textAlign: "left", fontSize: 10, fontWeight: 700, textTransform: "uppercase", color: "var(--text3)", borderBottom: "1px solid var(--border)" }}>{heading}</th>
                              ))}
                            </tr>
                          </thead>
                          <tbody>
                            {group.rows.map((row, index) => (
                              <tr key={`${row.key}-${index}`} style={{ borderBottom: "1px solid var(--border)", background: index % 2 === 0 ? "transparent" : "rgba(255,255,255,.01)" }}>
                                <td style={{ padding: "7px 8px", fontSize: 12 }}>{row.sourceRow}</td>
                                <td style={{ padding: "7px 8px", fontSize: 12, fontWeight: 700 }}>{row.programme}</td>
                                <td style={{ padding: "7px 8px", fontSize: 12 }}>{row.weekday || "—"}</td>
                                <td style={{ padding: "7px 8px", fontSize: 12 }}>{row.timeBelt || "—"}</td>
                                <td style={{ padding: "7px 8px", fontSize: 12 }}>{row.material || "—"}</td>
                                <td style={{ padding: "7px 8px", fontSize: 12 }}>{fmtN(row.ratePerSpot || 0)}</td>
                                <td style={{ padding: "7px 8px", fontSize: 12 }}>{row.totalSpots}</td>
                                <td style={{ padding: "7px 8px", fontSize: 12, color: row.bonusSpots > 0 ? "var(--purple)" : "var(--text2)", fontWeight: row.bonusSpots > 0 ? 700 : 400 }}>{row.bonusSpots}</td>
                                <td style={{ padding: "7px 8px", fontSize: 12, color: "var(--green)", fontWeight: 700 }}>{row.paidSpots}</td>
                                <td style={{ padding: "7px 8px", fontSize: 12 }}>{fmtN(row.importedGross || 0)}</td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    ) : null}
                  </Card>
                );
              })}
            </div>
          </>
        )}

        <div style={{ display: "flex", justifyContent: "space-between", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
          <Btn variant="ghost" onClick={() => {
            if (step === "review") {
              setStep("setup");
              return;
            }
            onClose?.();
          }}>
            {step === "review" ? "← Back to setup" : "Close"}
          </Btn>
          <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
            {step === "setup" ? (
              <Btn
                variant="blue"
                onClick={() => fileRef.current?.click()}
                loading={loading}
              >
                {loading ? "Parsing…" : "Upload workbook"}
              </Btn>
            ) : (
              <Btn variant="blue" onClick={handleSave} loading={saving}>
                {saving ? "Creating MPO drafts…" : "Create Draft MPOs"}
              </Btn>
            )}
          </div>
        </div>
      </div>
    </Modal>
  );
};

export default function MPOPage({ vendors, clients, campaigns, rates, mpos, setMpos, setVendors, user, appSettings, onBulkImportStateChange, requestMpoRefresh }) {
  const canManage = hasPermission(user, "manageMpos");
  const canManageStatus = hasPermission(user, "manageMpoStatus") || canManage;
  const [view, setView] = useState("list");
  const [editId, setEditId] = useState(null);
  const [step, setStep] = useState(1);
  const [viewMode, setViewMode] = useState("active");
  const [toast, setToast] = useState(null);
  const [confirm, setConfirm] = useState(null);
  const [preview, setPreview] = useState(null);
  const [historyModal, setHistoryModal] = useState(null);
  const [statusModal, setStatusModal] = useState(null);
  const [executionModal, setExecutionModal] = useState(null);
  const [executionUploading, setExecutionUploading] = useState({ signedMpo: false, invoice: false, proof: false });
  const [searchTerm, setSearchTerm] = useState("");
  const [statusFilter, setStatusFilter] = useState("all");
  const [campaignFilter, setCampaignFilter] = useState("all");
  const [vendorFilter, setVendorFilter] = useState("all");
  const [clientFilter, setClientFilter] = useState("all");
  const [selectedMpoIds, setSelectedMpoIds] = useState([]);
  const [deleteProgress, setDeleteProgress] = useState(null);
  const workflowPanelStorageKey = `msp_mpo_workflow_panel_${user?.id || "guest"}`;
  const [workflowPanelOpen, setWorkflowPanelOpen] = useState(() => store.get(workflowPanelStorageKey, true));
  const [mediaPlanImportOpen, setMediaPlanImportOpen] = useState(false);
  const mpoListScrollRef = useRef(null);

  const VAT_RATE = parseFloat(appSettings?.vatRate) || 7.5;
  const [surcharge, setSurcharge] = useState({ pct: "", label: "" });
  const [vendorPickerOpen, setVendorPickerOpen] = useState(false);
  const [vendorSearch, setVendorSearch] = useState("");

  // Auto-generate MPO number: BRA001-JAN2025
  const genMpoNo = (brand, existingMpos) => {
    const prefix = (brand || "MPO").replace(/\s+/g,"").toUpperCase().slice(0,3);
    const monthNames = ["JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"];
    const now = new Date();
    const mon = monthNames[now.getMonth()];
    const yr = String(now.getFullYear()).slice(-2);
    const count = existingMpos.filter(m => (m.mpoNo||"").startsWith(prefix)).length + 1;
    return `${prefix}${String(count).padStart(3,"0")}-${mon}${yr}`;
  };

  const blankMPO = (brand = "", existingMpos = []) => ({
    campaignId: "", vendorId: "",
    mpoNo: existingMpos && existingMpos.length ? genMpoNo(brand, existingMpos) : "Pending auto-number on save",
    date: new Date().toISOString().slice(0, 10),
    month: "", months: [], year: new Date().getFullYear().toString(),
    medium: "",
    signedBy: "", signedTitle: "",
    preparedBy: user?.name || "", preparedContact: user?.phone || user?.email || "",
    preparedTitle: user?.title || "",
    preparedSignature: user?.signatureDataUrl || "",
    signedSignature: "",
    agencyAddress: user?.agencyAddress || "5, Craig Street, Ogudu GRA, Lagos",
    agencyEmail: user?.agencyEmail || "",
    agencyPhone: user?.agencyPhone || "",
    transmitMsg: "", status: "draft",
    discountOverridePct: ""
  });

  const [mpoData, setMpoData] = useState(() => blankMPO("", mpos));
  const [spots, setSpots] = useState([]);
  const [spotModal, setSpotModal] = useState(null);

  // Daily schedule — calendar mode
  // calRows: array of { id, programme, timeBelt, material, duration, rateId, customRate, dayCounts: { [day]: count } }
  const [dailyMode, setDailyMode] = useState(false);
  const blankCalRow = () => ({ id: uid(), programme: "", wd: "", timeBelt: "", material: "", materialCustom: "", duration: "30", rateId: "", customRate: "", dayCounts: {}, isComplimentary: false, bonusSpots: "" });
  // Multi-month: calData maps "Month Year" -> calRows array
  const [calData, setCalData] = useState({});
  const [activeCalMonth, setActiveCalMonth] = useState("");
  // Keep a stable first blank row per month so the first programme selection updates immediately.
  const calMonthFallbackRowsRef = useRef({});
  // Legacy single-month calRows still used for single-month mode
  const [calRows, setCalRows] = useState([blankCalRow()]);

  const getStableCalRows = (monthLabel) => {
    if (calData[monthLabel]) return calData[monthLabel];
    if (!calMonthFallbackRowsRef.current[monthLabel]) {
      calMonthFallbackRowsRef.current[monthLabel] = [blankCalRow()];
    }
    return calMonthFallbackRowsRef.current[monthLabel];
  };

  const setStableCalRows = (monthLabel, updater) => {
    setCalData(prev => {
      const baseRows = prev[monthLabel] || calMonthFallbackRowsRef.current[monthLabel] || [blankCalRow()];
      const nextRows = typeof updater === "function" ? updater(baseRows) : updater;
      calMonthFallbackRowsRef.current[monthLabel] = nextRows;
      return { ...prev, [monthLabel]: nextRows };
    });
  };

  const blankSpot = { programme: "", wd: "", timeBelt: "", material: "", materialCustom: "", duration: "30", rateId: "", customRate: "", spots: "", calendarDays: [], calendarDayCounts: {}, isComplimentary: false, bonusSpots: "" };
  const [spotForm, setSpotForm] = useState(blankSpot);
  const [editSpotId, setEditSpotId] = useState(null);
  const legacyDraftKey = "msp_mpo_draft";
  const draftCollectionKey = `msp_mpo_drafts_${user?.id || "guest"}`;
  const sortDraftsByRecent = (drafts = []) => [...drafts].sort((a, b) => (b?.savedAt || 0) - (a?.savedAt || 0));
  const migrateLegacyDraftIfNeeded = () => {
    const existingDrafts = store.get(draftCollectionKey, []);
    if (Array.isArray(existingDrafts) && existingDrafts.length) return sortDraftsByRecent(existingDrafts);
    const legacyDraft = store.get(legacyDraftKey);
    if (!legacyDraft) return [];
    const migratedDraft = {
      id: legacyDraft.id || uid(),
      ...legacyDraft,
      createdAt: legacyDraft.createdAt || legacyDraft.savedAt || Date.now(),
      savedAt: legacyDraft.savedAt || Date.now(),
    };
    const migrated = [migratedDraft];
    store.set(draftCollectionKey, migrated);
    store.del(legacyDraftKey);
    return migrated;
  };
  const [savedDrafts, setSavedDrafts] = useState(() => migrateLegacyDraftIfNeeded());
  const [activeDraftId, setActiveDraftId] = useState(null);
  const upd = k => v => setMpoData(m => ({ ...m, [k]: v }));
  const months = ["January","February","March","April","May","June","July","August","September","October","November","December"];
  const campaign = campaigns.find(c => c.id === mpoData.campaignId);
  const client = clients.find(c => c.id === campaign?.clientId);
  const vendor = vendors.find(v => v.id === mpoData.vendorId);
  const vendorRates = activeOnly(rates).filter(r => r.vendorId === mpoData.vendorId);
  const updS = k => v => setSpotForm(f => applyAutoMatchedVendorRate({ ...f, [k]: v }, vendorRates));
  const extractMonthNameFromScheduleLabel = (value = "") => {
    const raw = String(value || "").trim();
    if (!raw) return "";
    const direct = months.find(monthName => raw.toLowerCase() === monthName.toLowerCase());
    if (direct) return direct;
    const firstWord = raw.split(/\s+/)[0] || "";
    const firstMatch = months.find(monthName => firstWord.toLowerCase() === monthName.toLowerCase());
    if (firstMatch) return firstMatch;
    return months.find(monthName => raw.toLowerCase().includes(monthName.toLowerCase())) || "";
  };
  const deriveMonthsFromSpotRows = (spotRows = [], fallbackMonths = []) => {
    const monthSet = new Set();
    (spotRows || []).forEach((spot) => {
      const monthName = extractMonthNameFromScheduleLabel(spot?.scheduleMonth);
      if (monthName) monthSet.add(monthName);
    });
    if (!monthSet.size) {
      (fallbackMonths || []).forEach((monthName) => {
        const normalized = extractMonthNameFromScheduleLabel(monthName);
        if (normalized) monthSet.add(normalized);
      });
    }
    return months.filter((monthName) => monthSet.has(monthName));
  };

  const buildCalDataFromSpots = (spotRows = []) => {
    const next = {};
    (spotRows || []).forEach((spot) => {
      const monthName = extractMonthNameFromScheduleLabel(spot?.scheduleMonth) || mpoData.month || mpoData.months?.[0] || "";
      if (!monthName) return;
      const calendarDays = Array.isArray(spot?.calendarDays) && spot.calendarDays.length
        ? spot.calendarDays
        : expandCountsToDays(spot?.calendarDayCounts || {});
      next[monthName] ||= [];
      next[monthName].push({
        ...blankCalRow(),
        id: spot?.id || uid(),
        programme: spot?.programme || "",
        wd: spot?.wd || "",
        timeBelt: spot?.timeBelt || "",
        material: spot?.material || "",
        materialCustom: "",
        duration: normalizeDurationValue(spot?.duration, "30"),
        rateId: spot?.rateId || "",
        customRate: spot?.isComplimentary ? "" : (spot?.customRate || (spot?.ratePerSpot ? String(spot.ratePerSpot) : "")),
        dayCounts: collapseDaysToCounts(calendarDays),
        isComplimentary: !!spot?.isComplimentary,
        bonusSpots: String((spot?.bonusSpots ?? getSpotBonusCount(spot)) || ""),
      });
    });
    return next;
  };

  const buildSpotsFromCalData = (calendarData = {}) => {
    const rebuilt = [];
    (mpoData.months?.length ? mpoData.months : Object.keys(calendarData || {})).forEach((monthName) => {
      const rows = calendarData?.[monthName] || [];
      rows.forEach((row) => {
        const calendarDays = expandCountsToDays(row?.dayCounts || {});
        if (!row?.programme || calendarDays.length === 0) return;
        const rate = vendorRates.find(r => r.id === row.rateId);
        const selectedRate = parseFloat(row.customRate) || parseFloat(rate?.ratePerSpot) || 0;
        const ratePerSpot = row.isComplimentary ? 0 : selectedRate;
        const matFinal = row.material === "__custom__" ? (row.materialCustom || "") : (row.material || "");
        const totalScheduledSpots = calendarDays.length;
        const bonusSpots = row.isComplimentary
          ? totalScheduledSpots
          : Math.max(0, Math.min(Number(row.bonusSpots) || 0, totalScheduledSpots));

        rebuilt.push(normalizeBonusAdjustedSpot({
          id: row.id || uid(),
          programme: row.programme,
          wd: row.wd || "",
          timeBelt: row.timeBelt,
          material: matFinal,
          duration: normalizeDurationValue(row.duration, "30"),
          rateId: row.rateId,
          ratePerSpot,
          customRate: row.isComplimentary ? "" : row.customRate,
          scheduleMonth: `${monthName} ${mpoData.year}`.trim(),
          calendarDays,
          isComplimentary: !!row.isComplimentary,
        }, totalScheduledSpots, bonusSpots));
      });
    });
    return rebuilt;
  };

  const updateCalendarRowField = (monthName, rowId, patch) => {
    setStableCalRows(monthName, (rows) => (rows || []).map((row) => {
      if (row.id !== rowId) return row;
      const totalScheduledSpots = totalCountFromDayCounts(row?.dayCounts || {});
      const nextPatch = typeof patch === "function" ? patch(row, totalScheduledSpots) : patch;
      const merged = applyAutoMatchedVendorRate({ ...row, ...nextPatch }, vendorRates);
      const nextTotal = totalCountFromDayCounts(merged?.dayCounts || {});
      const normalizedBonus = merged.isComplimentary
        ? nextTotal
        : Math.max(0, Math.min(Number(merged.bonusSpots) || 0, nextTotal));
      return {
        ...merged,
        bonusSpots: nextTotal > 0 ? String(normalizedBonus) : "",
        customRate: merged.isComplimentary ? "" : merged.customRate,
      };
    }));
  };

  const calendarEditHasRows = Object.values(calData || {}).some((rows) =>
    (rows || []).some((row) => row?.programme || totalCountFromDayCounts(row?.dayCounts || {}) > 0)
  );
  const effectiveSpots = dailyMode && calendarEditHasRows ? buildSpotsFromCalData(calData) : spots;

  const totalSpots = effectiveSpots.reduce((s, r) => s + (parseFloat(r.spots) || 0), 0);
  const totalBonusSpots = effectiveSpots.reduce((s, r) => s + getSpotBonusCount(r), 0);
  const totalPaidSpots = effectiveSpots.reduce((s, r) => s + getPaidSpotCount(r), 0);
  const totalGross = effectiveSpots.reduce((s, r) => s + getPaidSpotCount(r) * (parseFloat(r.ratePerSpot) || 0), 0);
  const vendorDiscPct = vendor ? (parseFloat(vendor.discount) || 0) / 100 : 0;
  const hasDiscountOverride = String(mpoData?.discountOverridePct ?? "").trim() !== "";
  const parsedDiscountOverride = parseFloat(mpoData?.discountOverridePct);
  const discountOverridePct = Number.isFinite(parsedDiscountOverride) ? Math.max(0, parsedDiscountOverride) / 100 : null;
  const discPct = hasDiscountOverride ? (discountOverridePct ?? 0) : vendorDiscPct;
  const commPct = vendor ? (parseFloat(vendor.commission) || 0) / 100 : 0;
  const discAmt = roundMoneyValue(totalGross * discPct, appSettings);
  const lessDisc = roundMoneyValue(totalGross - discAmt, appSettings);
  const commAmt = roundMoneyValue(lessDisc * commPct, appSettings);
  const afterComm = roundMoneyValue(lessDisc - commAmt, appSettings);
  const surchPct = (parseFloat(surcharge.pct) || 0) / 100;
  const surchAmt = roundMoneyValue(afterComm * surchPct, appSettings);
  const netVal = roundMoneyValue(afterComm + surchAmt, appSettings);
  const vatAmt = roundMoneyValue(netVal * VAT_RATE / 100, appSettings);
  const grandTotal = roundMoneyValue(netVal + vatAmt, appSettings);
  const rateOptions = vendorRates.map(r => ({ value: r.id, label: `${r.programme || "Unnamed"} – ${fmtN(r.ratePerSpot)}` }));
  const filteredVendors = vendors.filter(v => {
    const q = vendorSearch.trim().toLowerCase();
    if (!q) return true;
    return [v.name, v.type, v.contactName, v.email, v.phone]
      .filter(Boolean)
      .some(value => String(value).toLowerCase().includes(q));
  });

  useEffect(() => {
    setWorkflowPanelOpen(store.get(`msp_mpo_workflow_panel_${user?.id || "guest"}`, true));
  }, [user?.id]);
  useEffect(() => {
    store.set(workflowPanelStorageKey, workflowPanelOpen);
  }, [workflowPanelStorageKey, workflowPanelOpen]);

  useEffect(() => {
    setSelectedMpoIds(ids => ids.filter(id => mpos.some(item => item.id === id)));
  }, [mpos]);

  // When campaign changes, auto-generate MPO number with brand
  useEffect(() => {
    if (campaign?.brand && !editId) {
      setMpoData(m => ({ ...m, mpoNo: genMpoNo(campaign.brand, mpos) }));
    }
  }, [mpoData.campaignId]);

  useEffect(() => {
    const handler = (e) => {
      if (view === "form") {
        e.preventDefault();
        e.returnValue = "";
      }
    };
    window.addEventListener("beforeunload", handler);
    return () => window.removeEventListener("beforeunload", handler);
  }, [view]);

  const syncSavedDrafts = (updater) => {
    setSavedDrafts(prev => {
      const nextDrafts = typeof updater === "function" ? updater(prev || []) : updater;
      const normalized = sortDraftsByRecent(Array.isArray(nextDrafts) ? nextDrafts : []);
      store.set(draftCollectionKey, normalized);
      return normalized;
    });
  };

  const buildDraftTitle = (draft) => {
    const draftCampaign = campaigns.find(c => c.id === draft?.mpoData?.campaignId);
    const draftVendor = vendors.find(v => v.id === draft?.mpoData?.vendorId);
    const draftBrand = draftCampaign?.brand || draft?.mpoData?.brand || "";
    if (draft?.mpoData?.mpoNo && draft?.editId) return draft.mpoData.mpoNo;
    if (draftCampaign?.name) return draftCampaign.name;
    if (draftVendor?.name && draftBrand) return `${draftBrand} · ${draftVendor.name}`;
    if (draftVendor?.name) return draftVendor.name;
    if (draftBrand) return draftBrand;
    return "Untitled MPO Draft";
  };

  const buildDraftSubtitle = (draft) => {
    const draftCampaign = campaigns.find(c => c.id === draft?.mpoData?.campaignId);
    const draftClient = clients.find(c => c.id === draftCampaign?.clientId);
    const draftVendor = vendors.find(v => v.id === draft?.mpoData?.vendorId);
    const bits = [
      draftClient?.name || draft?.mpoData?.clientName || "",
      draftVendor?.name || draft?.mpoData?.vendorName || "",
      draft?.spots?.length ? `${draft.spots.length} row${draft.spots.length !== 1 ? "s" : ""}` : "",
      draft?.step ? `Step ${draft.step}` : "",
    ].filter(Boolean);
    return bits.join(" · ");
  };

  const hasMeaningfulDraftContent = (nextMpoData, nextSpots, nextSurcharge, nextCalData = {}, nextDailyMode = false) => {
    return Boolean(
      nextMpoData?.campaignId ||
      nextMpoData?.vendorId ||
      nextMpoData?.signedBy ||
      nextMpoData?.signedTitle ||
      nextMpoData?.transmitMsg ||
      nextMpoData?.month ||
      (nextMpoData?.months || []).length ||
      nextMpoData?.medium ||
      (nextSpots || []).length ||
      Object.keys(nextCalData || {}).length ||
      nextDailyMode ||
      nextSurcharge?.pct ||
      nextSurcharge?.label
    );
  };

  const removeSavedDraft = (draftId, options = {}) => {
    if (!draftId) return;
    syncSavedDrafts(prev => prev.filter(draft => draft.id !== draftId));
    if (activeDraftId === draftId) setActiveDraftId(null);
    if (!options.quiet) setToast({ msg: "Draft removed.", type: "success" });
  };

  useEffect(() => {
    setSavedDrafts(migrateLegacyDraftIfNeeded());
    setActiveDraftId(null);
  }, [draftCollectionKey]);

  useEffect(() => {
    if (view !== "form") return;
    if (!hasMeaningfulDraftContent(mpoData, spots, surcharge, calData, dailyMode)) return;

    const nextDraftId = activeDraftId || uid();
    if (!activeDraftId) setActiveDraftId(nextDraftId);

    syncSavedDrafts(prev => {
      const existingDraft = prev.find(draft => draft.id === nextDraftId);
      const payload = {
        id: nextDraftId,
        editId,
        step,
        mpoData,
        spots,
        calData,
        activeCalMonth,
        dailyMode,
        surcharge,
        createdAt: existingDraft?.createdAt || Date.now(),
        savedAt: Date.now(),
      };
      return [payload, ...prev.filter(draft => draft.id !== nextDraftId)];
    });
  }, [view, activeDraftId, editId, step, mpoData, spots, calData, activeCalMonth, dailyMode, surcharge]);

  const resumeSavedDraft = (draftId) => {
    const draft = savedDrafts.find(item => item.id === draftId);
    if (!draft) return setToast({ msg: "Draft not found.", type: "error" });
    setActiveDraftId(draft.id);
    setEditId(draft.editId || null);
    setStep(draft.step || 1);
    setMpoData({ ...blankMPO("", mpos), ...(draft.mpoData || {}), preparedSignature: draft?.mpoData?.preparedSignature || user?.signatureDataUrl || "", signedSignature: draft?.mpoData?.signedSignature || "" });
    const restoredSpots = draft.spots || [];
    const restoredCalData = draft.calData && Object.keys(draft.calData || {}).length ? draft.calData : buildCalDataFromSpots(restoredSpots);
    setSpots(restoredSpots);
    setCalData(restoredCalData);
    calMonthFallbackRowsRef.current = restoredCalData || {};
    setActiveCalMonth(draft.activeCalMonth || Object.keys(restoredCalData || {})[0] || "");
    setDailyMode(typeof draft.dailyMode === "boolean" ? draft.dailyMode : Object.keys(restoredCalData || {}).length > 0);
    setSurcharge(draft.surcharge || { pct: "", label: "" });
    setView("form");
    setToast({ msg: "MPO draft restored.", type: "success" });
  };

  const clearSavedDraft = (draftId = activeDraftId, options = {}) => {
    if (!draftId) {
      syncSavedDrafts([]);
      setActiveDraftId(null);
      if (!options.quiet) setToast({ msg: "All saved MPO drafts cleared.", type: "success" });
      return;
    }
    removeSavedDraft(draftId, options);
  };

  const refreshHistoryModal = async (mpoId, title) => {
    try {
      const events = await fetchAuditEventsForRecord(user.agencyId, "mpo", mpoId);
      setHistoryModal({ mpoId, title, events });
    } catch (error) {
      setToast({ msg: error.message || "Failed to load MPO history.", type: "error" });
    }
  };

  const openMpoHistory = async (mpo) => {
    await refreshHistoryModal(mpo.id, `MPO History — ${mpo.mpoNo || mpo.id}`);
  };

  const openExecutionModal = (mpo) => {
    setExecutionModal({
      mpoId: mpo.id,
      mpoNo: mpo.mpoNo || mpo.id,
      dispatchStatus: mpo.dispatchStatus || "pending",
      dispatchedAt: toIsoInput(mpo.dispatchedAt),
      dispatchContact: mpo.dispatchContact || mpo.vendorName || "",
      dispatchNote: mpo.dispatchNote || "",
      signedMpoUrl: mpo.signedMpoUrl || "",
      invoiceStatus: mpo.invoiceStatus || "pending",
      invoiceNo: mpo.invoiceNo || "",
      invoiceAmount: String(mpo.invoiceAmount ?? mpo.grandTotal ?? ""),
      invoiceReceivedAt: toIsoInput(mpo.invoiceReceivedAt),
      invoiceUrl: mpo.invoiceUrl || "",
      proofStatus: mpo.proofStatus || "pending",
      proofUrl: mpo.proofUrl || "",
      proofReceivedAt: toIsoInput(mpo.proofReceivedAt),
      plannedSpotsExecution: String(mpo.plannedSpotsExecution ?? mpo.totalSpots ?? 0),
      airedSpots: String(mpo.airedSpots ?? 0),
      missedSpots: String(mpo.missedSpots ?? 0),
      makegoodSpots: String(mpo.makegoodSpots ?? 0),
      reconciliationStatus: mpo.reconciliationStatus || "not_started",
      reconciliationNotes: mpo.reconciliationNotes || "",
      reconciledAmount: String(mpo.reconciledAmount ?? mpo.grandTotal ?? 0),
      paymentStatus: mpo.paymentStatus || "unpaid",
      paymentReference: mpo.paymentReference || "",
      paidAt: toIsoInput(mpo.paidAt),
    });
  };

  const uploadExecutionAttachment = async (kind, file) => {
    if (!file) return;
    if (!executionModal?.mpoId) {
      setToast({ msg: "Save or open an existing MPO before uploading attachments.", type: "error" });
      return;
    }
    if (!user?.agencyId) {
      setToast({ msg: "No agency found for this workspace.", type: "error" });
      return;
    }
    const keyMap = {
      signedMpo: "signedMpoUrl",
      invoice: "invoiceUrl",
      proof: "proofUrl",
    };
    const fieldName = keyMap[kind] || "signedMpoUrl";
    setExecutionUploading(state => ({ ...state, [kind]: true }));
    try {
      const url = await uploadMpoAttachmentAndGetUrl({
        agencyId: user.agencyId,
        mpoId: executionModal.mpoId,
        kind,
        file,
      });
      setExecutionModal(modal => ({ ...modal, [fieldName]: url }));
      setToast({ msg: `${kind === "signedMpo" ? "Signed MPO" : kind === "invoice" ? "Invoice" : "Proof"} uploaded. Save execution to persist it.`, type: "success" });
    } catch (e) {
      setToast({ msg: e.message || "Upload failed.", type: "error" });
    } finally {
      setExecutionUploading(state => ({ ...state, [kind]: false }));
    }
  };

  const saveExecutionModal = async () => {
    if (!executionModal?.mpoId) return;
    try {
      const patch = {
        dispatchStatus: executionModal.dispatchStatus,
        dispatchedAt: toIsoOrNull(executionModal.dispatchedAt),
        dispatchedBy: user?.id || null,
        dispatchContact: executionModal.dispatchContact,
        dispatchNote: executionModal.dispatchNote,
        signedMpoUrl: executionModal.signedMpoUrl,
        invoiceStatus: executionModal.invoiceStatus,
        invoiceNo: executionModal.invoiceNo,
        invoiceAmount: executionModal.invoiceAmount,
        invoiceReceivedAt: toIsoOrNull(executionModal.invoiceReceivedAt),
        invoiceUrl: executionModal.invoiceUrl,
        proofStatus: executionModal.proofStatus,
        proofUrl: executionModal.proofUrl,
        proofReceivedAt: toIsoOrNull(executionModal.proofReceivedAt),
        plannedSpotsExecution: executionModal.plannedSpotsExecution,
        airedSpots: executionModal.airedSpots,
        missedSpots: executionModal.missedSpots,
        makegoodSpots: executionModal.makegoodSpots,
        reconciliationStatus: executionModal.reconciliationStatus,
        reconciliationNotes: executionModal.reconciliationNotes,
        reconciledAmount: executionModal.reconciledAmount,
        paymentStatus: executionModal.paymentStatus,
        paymentReference: executionModal.paymentReference,
        paidAt: toIsoOrNull(executionModal.paidAt),
      };
      const updated = await updateMpoExecutionInSupabase(executionModal.mpoId, patch);
      setMpos(ms => ms.map(x => x.id === executionModal.mpoId ? updated : x));
      await createAuditEventInSupabase({
        agencyId: user.agencyId,
        recordType: "mpo",
        recordId: executionModal.mpoId,
        action: "execution_updated",
        actor: user,
        note: `Execution & reconciliation updated for ${executionModal.mpoNo || executionModal.mpoId}.`,
        metadata: {
          mpoNo: executionModal.mpoNo || "",
          dispatchStatus: patch.dispatchStatus,
          invoiceStatus: patch.invoiceStatus,
          proofStatus: patch.proofStatus,
          reconciliationStatus: patch.reconciliationStatus,
          paymentStatus: patch.paymentStatus,
        },
      });
      await notifyExecutionUpdate({ agencyId: user.agencyId, mpo: executionModal, actor: user, patch });
      setExecutionModal(null);
      setToast({ msg: "Execution & reconciliation saved.", type: "success" });
    } catch (error) {
      setToast({ msg: error.message || "Failed to save execution details.", type: "error" });
    }
  };

  const applyMpoStatusChange = async (mpo, nextStatus, note = "") => {
    try {
      const updated = await updateMpoStatusInSupabase(mpo.id, nextStatus);
      setMpos(ms => ms.map(x => x.id === mpo.id ? updated : x));
      await createAuditEventInSupabase({
        agencyId: user.agencyId,
        recordType: "mpo",
        recordId: mpo.id,
        action: "status_changed",
        actor: user,
        note,
        metadata: {
          mpoNo: mpo.mpoNo || "",
          fromStatus: mpo.status || "draft",
          toStatus: nextStatus,
        },
      });
      await notifyMpoWorkflowTransition({ agencyId: user.agencyId, mpo, nextStatus, actor: user, note });
      setStatusModal(null);
      if (historyModal?.mpoId === mpo.id) {
        await refreshHistoryModal(mpo.id, historyModal.title || `MPO History — ${mpo.mpoNo || mpo.id}`);
      }
      setToast({ msg: `MPO moved to ${MPO_STATUS_LABELS[nextStatus] || nextStatus}.`, type: "success" });
    } catch (error) {
      setToast({ msg: error.message || "Failed to update MPO status.", type: "error" });
    }
  };

  const requestMpoStatusChange = (mpo, nextStatus) => {
    const current = String(mpo?.status || "draft").toLowerCase();
    const target = String(nextStatus || current).toLowerCase();
    if (target === current) return;
    const allowedTargets = getAllowedMpoStatusTargets(user, mpo);
    if (!allowedTargets.includes(target)) {
      setToast({ msg: `Your role (${formatRoleLabel(user?.role)}) cannot move this MPO from ${MPO_STATUS_LABELS[current] || current} to ${MPO_STATUS_LABELS[target] || target}.`, type: "error" });
      return;
    }
    const nextOwner = getMpoWorkflowMeta({ ...mpo, status: target });
    if (mpoStatusNeedsNote(target)) {
      setStatusModal({
        mpo,
        nextStatus: target,
        note: "",
        actionLabel: getWorkflowActionLabel(current, target),
        noteLabel: target === "rejected" ? "Request changes note" : "Workflow note",
        helperText: nextOwner?.hint || "",
        nextOwnerLabel: nextOwner?.label || "Next team member",
      });
      return;
    }
    applyMpoStatusChange(mpo, target, "");
  };

  const restoreMPO = async (id) => {
    if (!canManage) return setToast({ msg: readOnlyMessage(user), type: "error" });
    try {
      const restored = await restoreMpoInSupabase(id);
      setMpos(m => m.map(x => x.id === id ? restored : x));
      createAuditEventInSupabase({
        agencyId: user.agencyId,
        recordType: "mpo",
        recordId: id,
        action: "restored",
        actor: user,
        metadata: { mpoNo: restored.mpoNo || "", status: restored.status || "draft" },
      }).catch(error => console.error("Failed to write MPO audit event:", error));
      setToast({ msg: "MPO restored.", type: "success" });
    } catch (e) {
      setToast({ msg: e.message || "Failed to restore MPO.", type: "error" });
    }
  };

  const openNew = () => {
    const blank = blankMPO("", mpos);
    setMpoData(blank); setSpots([]); setStep(1); setEditId(null); setView("form");
    calMonthFallbackRowsRef.current = {};
    setDailyMode(false); setCalRows([blankCalRow()]); setCalData({}); setActiveCalMonth("");
    setActiveDraftId(null);
  };
  const openEdit = (mpo) => {
    if (!canEditMpoContent(user, mpo)) {
      setToast({ msg: `You can only edit MPOs in Draft or Rejected status. Current status: ${MPO_STATUS_LABELS[mpo?.status || "draft"] || (mpo?.status || "draft")}.`, type: "error" });
      return;
    }
    const normalizedSpots = (mpo.spots || []).map(spot => ({
      ...spot,
      bonusSpots: spot?.bonusSpots ?? getSpotBonusCount(spot),
      paidSpots: spot?.paidSpots ?? getPaidSpotCount(spot),
    }));
    const rebuiltCalData = buildCalDataFromSpots(normalizedSpots);
    const rebuiltMonths = Object.keys(rebuiltCalData || {});
    setMpoData({
      campaignId: mpo.campaignId||"",
      vendorId: mpo.vendorId||"",
      mpoNo: mpo.mpoNo||"",
      date: mpo.date||"",
      month: mpo.month||rebuiltMonths[0]||"",
      months: mpo.months?.length ? mpo.months : rebuiltMonths,
      year: mpo.year||"",
      medium: mpo.medium||"",
      signedBy: mpo.signedBy||"",
      signedTitle: mpo.signedTitle||"",
      preparedBy: mpo.preparedBy||user?.name||"",
      preparedContact: mpo.preparedContact||user?.phone||user?.email||"",
      preparedTitle: mpo.preparedTitle||user?.title||"",
      preparedSignature: mpo.preparedSignature||user?.signatureDataUrl||"",
      signedSignature: mpo.signedSignature||"",
      agencyAddress: mpo.agencyAddress||user?.agencyAddress||"",
      transmitMsg: mpo.transmitMsg||"",
      status: mpo.status||"draft",
      discountOverridePct: mpo?.discPct !== undefined && mpo?.discPct !== null ? String((Number(mpo.discPct) || 0) * 100) : ""
    });
    setSurcharge({ pct: mpo.surchPct ? String((mpo.surchPct||0)*100) : "", label: mpo.surchLabel||"" });
    setSpots(normalizedSpots);
    setEditId(mpo.id);
    setStep(2);
    setView("form");
    calMonthFallbackRowsRef.current = rebuiltCalData;
    setCalRows([blankCalRow()]);
    setCalData(rebuiltCalData);
    setActiveCalMonth(rebuiltMonths[0] || "");
    setDailyMode(rebuiltMonths.length > 0);
    setActiveDraftId(null);
  };

  // Add spot rows from calendar schedule — supports single or multi-month
  const addFromCalendarSchedule = (rowsToAdd, monthLabel) => {
    const sourceRows = rowsToAdd || calRows;
    const newSpots = [];
    sourceRows.forEach(row => {
      const calendarDays = expandCountsToDays(row.dayCounts || {});
      if (!row.programme || calendarDays.length === 0) return;
      const rate = vendorRates.find(r => r.id === row.rateId);
      const selectedRate = parseFloat(row.customRate) || parseFloat(rate?.ratePerSpot) || 0;
      const ratePerSpot = row.isComplimentary ? 0 : selectedRate;
      const matFinal = row.material === "__custom__" ? (row.materialCustom || "") : (row.material || "");
      const totalScheduledSpots = calendarDays.length;
      const bonusSpots = row.isComplimentary
        ? totalScheduledSpots
        : Math.max(0, Math.min(Number(row.bonusSpots) || 0, totalScheduledSpots));
      newSpots.push(normalizeBonusAdjustedSpot({
        id: uid(),
        programme: row.programme,
        wd: row.wd || "",
        timeBelt: row.timeBelt,
        material: matFinal,
        duration: row.duration,
        rateId: row.rateId,
        ratePerSpot,
        customRate: row.isComplimentary ? "" : row.customRate,
        calendarDays,
        scheduleMonth: monthLabel || mpoData.month,
        isComplimentary: !!row.isComplimentary,
      }, totalScheduledSpots, bonusSpots));
    });
    if (!newSpots.length) return setToast({ msg: "Fill Programme and add at least one spot on at least one date per row.", type: "error" });
    setSpots(s => [...s, ...newSpots]);
    if (!rowsToAdd) { setDailyMode(false); setCalRows([blankCalRow()]); }
    setToast({ msg: `${newSpots.length} spot row(s) added from calendar.`, type: "success" });
  };

  const addSpot = () => {
    if (!spotForm.programme) return;
    const rate = vendorRates.find(r => r.id === spotForm.rateId);
    const selectedRate = parseFloat(spotForm.customRate) || parseFloat(rate?.ratePerSpot) || 0;
    const ratePerSpot = spotForm.isComplimentary ? 0 : selectedRate;
    const calDays = spotForm.calendarDays?.length
      ? spotForm.calendarDays
      : expandCountsToDays(spotForm.calendarDayCounts || {});
    const spotsCount = calDays.length > 0 ? calDays.length : (parseFloat(spotForm.spots) || 0);
    if (!spotsCount) return setToast({ msg: "Select at least one airing date or enter a spot count.", type: "error" });
    const bonusSpots = spotForm.isComplimentary
      ? spotsCount
      : Math.max(0, Math.min(Number(spotForm.bonusSpots) || 0, spotsCount));
    const newSpot = normalizeBonusAdjustedSpot({ id: editSpotId || uid(), ...spotForm, ratePerSpot, calendarDays: calDays, calendarDayCounts: collapseDaysToCounts(calDays) }, spotsCount, bonusSpots);
    if (editSpotId) { setSpots(s => s.map(x => x.id === editSpotId ? newSpot : x)); setEditSpotId(null); }
    else setSpots(s => [...s, newSpot]);
    setSpotForm(blankSpot); setSpotModal(null);
  };

  const saveMPO = async () => {
    if (!canManage) return setToast({ msg: readOnlyMessage(user), type: "error" });
    const existingEditingMpo = editId ? mpos.find(m => m.id === editId) : null;
    if (existingEditingMpo && !canEditMpoContent(user, existingEditingMpo)) {
      return setToast({ msg: `You can only edit MPOs in Draft or Rejected status. Current status: ${MPO_STATUS_LABELS[existingEditingMpo.status || "draft"] || (existingEditingMpo.status || "draft")}.`, type: "error" });
    }
    if (!mpoData.campaignId) return setToast({ msg: "Please select a campaign.", type: "error" });
    if (!mpoData.vendorId) return setToast({ msg: "Please select a vendor.", type: "error" });
    if (!effectiveSpots.length) return setToast({ msg: "Add at least one spot row before saving this MPO.", type: "error" });
    if (editId && !mpoData.mpoNo) return setToast({ msg: "MPO number is required.", type: "error" });
    const duplicate = editId && mpoData.mpoNo ? mpos.find(m => m.id !== editId && !isArchived(m) && (m.mpoNo || "").trim().toLowerCase() === mpoData.mpoNo.trim().toLowerCase()) : null;
    if (duplicate) return setToast({ msg: "This MPO number already exists.", type: "error" });
    const campaignBudget = parseFloat(campaign?.budget) || 0;
    if (campaignBudget > 0 && grandTotal > campaignBudget) return setToast({ msg: "This MPO total is above the campaign budget. Reduce spots or update the campaign budget first.", type: "error" });

    try {
      const generatedMpoNo = editId ? mpoData.mpoNo : await generateNextMpoNoFromSupabase(campaign?.brand || mpoData.brand || "MPO");
      const existingExec = editId ? (mpos.find(m => m.id === editId) || {}) : {};
      const derivedMonths = deriveMonthsFromSpotRows(effectiveSpots, mpoData.months?.length ? mpoData.months : [mpoData.month]);
      const record = { id: editId || uid(), ...mpoData, preparedSignature: mpoData.preparedSignature || user?.signatureDataUrl || "", signedSignature: mpoData.signedSignature || "", agencyEmail: mpoData.agencyEmail || user?.agencyEmail || "", agencyPhone: mpoData.agencyPhone || user?.agencyPhone || "", mpoNo: generatedMpoNo, vendorName: vendor?.name || "", clientName: client?.name || "", campaignName: campaign?.name || "", brand: campaign?.brand || "", medium: mpoData.medium || campaign?.medium || "", month: derivedMonths[0] || mpoData.month || "", months: derivedMonths, spots: effectiveSpots, totalSpots, totalGross, discPct, discAmt, lessDisc, commPct, commAmt: commAmt, afterComm, surchPct, surchAmt, surchLabel: surcharge.label, netVal, vatPct: VAT_RATE, vatAmt, grandTotal, terms: appSettings?.mpoTerms || DEFAULT_APP_SETTINGS.mpoTerms, roundToWholeNaira: !!appSettings?.roundToWholeNaira,
        dispatchStatus: existingExec.dispatchStatus || "pending",
        dispatchedAt: existingExec.dispatchedAt || null,
        dispatchedBy: existingExec.dispatchedBy || null,
        dispatchContact: existingExec.dispatchContact || "",
        dispatchNote: existingExec.dispatchNote || "",
        signedMpoUrl: existingExec.signedMpoUrl || "",
        invoiceStatus: existingExec.invoiceStatus || "pending",
        invoiceNo: existingExec.invoiceNo || "",
        invoiceAmount: existingExec.invoiceAmount ?? 0,
        invoiceReceivedAt: existingExec.invoiceReceivedAt || null,
        invoiceUrl: existingExec.invoiceUrl || "",
        proofStatus: existingExec.proofStatus || "pending",
        proofUrl: existingExec.proofUrl || "",
        proofReceivedAt: existingExec.proofReceivedAt || null,
        plannedSpotsExecution: existingExec.plannedSpotsExecution ?? totalSpots,
        airedSpots: existingExec.airedSpots ?? 0,
        missedSpots: existingExec.missedSpots ?? 0,
        makegoodSpots: existingExec.makegoodSpots ?? 0,
        reconciliationStatus: existingExec.reconciliationStatus || "not_started",
        reconciliationNotes: existingExec.reconciliationNotes || "",
        reconciledAmount: existingExec.reconciledAmount ?? grandTotal,
        paymentStatus: existingExec.paymentStatus || "unpaid",
        paymentReference: existingExec.paymentReference || "",
        paidAt: existingExec.paidAt || null,
        createdAt: editId ? (mpos.find(m => m.id === editId)?.createdAt || Date.now()) : Date.now(), updatedAt: Date.now() };

      setSpots(effectiveSpots);
      let saved;
      if (editId) {
        saved = { ...(await updateMpoInSupabase(editId, record)), preparedSignature: record.preparedSignature, signedSignature: record.signedSignature };
        setMpos(m => m.map(x => x.id === editId ? saved : x));
        createAuditEventInSupabase({
          agencyId: user.agencyId,
          recordType: "mpo",
          recordId: editId,
          action: "updated",
          actor: user,
          metadata: { mpoNo: saved.mpoNo || generatedMpoNo, status: saved.status || "draft", grandTotal: saved.grandTotal || grandTotal },
        }).catch(error => console.error("Failed to write MPO audit event:", error));
      } else {
        saved = { ...(await createMpoInSupabase(user.agencyId, user.id, record)), preparedSignature: record.preparedSignature, signedSignature: record.signedSignature };
        setMpos(m => [saved, ...m]);
        createAuditEventInSupabase({
          agencyId: user.agencyId,
          recordType: "mpo",
          recordId: saved.id,
          action: "created",
          actor: user,
          metadata: { mpoNo: saved.mpoNo || generatedMpoNo, status: saved.status || "draft", grandTotal: saved.grandTotal || grandTotal },
        }).catch(error => console.error("Failed to write MPO audit event:", error));
      }
      setToast({ msg: editId ? "MPO updated!" : `MPO ${generatedMpoNo} saved!`, type: "success" });
      clearSavedDraft(activeDraftId, { quiet: true });
      setActiveDraftId(null);
      setView("list");
    } catch (e) {
      setToast({ msg: e.message || "Failed to save MPO.", type: "error" });
    }
  };

  const archiveMPO = async (id) => {
    if (!canManage) return setToast({ msg: readOnlyMessage(user), type: "error" });
    try {
      const archived = await archiveMpoInSupabase(id);
      setMpos(m => m.map(x => x.id === id ? archived : x));
      setSelectedMpoIds(ids => ids.filter(item => item !== id));
      createAuditEventInSupabase({
        agencyId: user.agencyId,
        recordType: "mpo",
        recordId: id,
        action: "archived",
        actor: user,
        metadata: { mpoNo: archived.mpoNo || "", status: archived.status || "draft" },
      }).catch(error => console.error("Failed to write MPO audit event:", error));
      setConfirm(null);
      setToast({ msg: "MPO archived.", type: "success" });
    } catch (e) {
      setToast({ msg: e.message || "Failed to archive MPO.", type: "error" });
    }
  };

  const deleteMPO = async (id) => {
    if (!canManage) return setToast({ msg: readOnlyMessage(user), type: "error" });
    try {
      const mpo = mpos.find(item => item.id === id);
      const mod = await import("../services/mpos");
      if (typeof mod.deleteMpoInSupabase !== "function") {
        throw new Error("deleteMpoInSupabase is not available in ../services/mpos yet.");
      }
      await mod.deleteMpoInSupabase(id);
      setMpos(items => items.filter(item => item.id !== id));
      setSelectedMpoIds(ids => ids.filter(item => item !== id));
      createAuditEventInSupabase({
        agencyId: user.agencyId,
        recordType: "mpo",
        recordId: id,
        action: "deleted",
        actor: user,
        metadata: { mpoNo: mpo?.mpoNo || "", status: mpo?.status || "draft" },
      }).catch(error => console.error("Failed to write MPO audit event:", error));
      setConfirm(null);
      setToast({ msg: "MPO deleted.", type: "success" });
    } catch (e) {
      setToast({ msg: e.message || "Failed to delete MPO.", type: "error" });
    }
  };

  const archiveSelectedMpos = async () => {
    if (!selectedVisibleMpoIds.length) return setToast({ msg: "Select at least one MPO first.", type: "error" });
    try {
      for (const id of selectedVisibleMpoIds) {
        const archived = await archiveMpoInSupabase(id);
        setMpos(items => items.map(item => item.id === id ? archived : item));
        createAuditEventInSupabase({
          agencyId: user.agencyId,
          recordType: "mpo",
          recordId: id,
          action: "archived",
          actor: user,
          metadata: { mpoNo: archived.mpoNo || "", status: archived.status || "draft" },
        }).catch(error => console.error("Failed to write MPO audit event:", error));
      }
      setSelectedMpoIds(ids => ids.filter(id => !selectedVisibleMpoIds.includes(id)));
      setConfirm(null);
      setToast({ msg: `${selectedVisibleMpoIds.length} MPO${selectedVisibleMpoIds.length !== 1 ? "s" : ""} archived.`, type: "success" });
    } catch (e) {
      setToast({ msg: e.message || "Failed to archive selected MPOs.", type: "error" });
    }
  };

  const deleteSelectedMpos = async () => {
    if (!selectedVisibleMpoIds.length) return setToast({ msg: "Select at least one MPO first.", type: "error" });
    try {
      const mod = await import("../services/mpos");
      if (typeof mod.deleteMpoInSupabase !== "function") {
        throw new Error("deleteMpoInSupabase is not available in ../services/mpos yet.");
      }
      const selectedItems = mpos.filter(item => selectedVisibleMpoIds.includes(item.id));
      const shouldShowProgress = selectedItems.length > 2;
      if (shouldShowProgress) {
        setDeleteProgress({ current: 0, total: selectedItems.length, label: "Preparing bulk delete..." });
      }
      let completed = 0;
      for (const mpo of selectedItems) {
        if (shouldShowProgress) {
          setDeleteProgress({
            current: completed,
            total: selectedItems.length,
            label: `Deleting ${completed + 1} of ${selectedItems.length}: ${mpo.mpoNo || mpo.id}`,
          });
        }
        await mod.deleteMpoInSupabase(mpo.id);
        completed += 1;
        if (shouldShowProgress) {
          setDeleteProgress({
            current: completed,
            total: selectedItems.length,
            label: completed >= selectedItems.length
              ? `Deleted ${completed} of ${selectedItems.length}`
              : `Deleted ${completed} of ${selectedItems.length}`,
          });
        }
        createAuditEventInSupabase({
          agencyId: user.agencyId,
          recordType: "mpo",
          recordId: mpo.id,
          action: "deleted",
          actor: user,
          metadata: { mpoNo: mpo?.mpoNo || "", status: mpo?.status || "draft" },
        }).catch(error => console.error("Failed to write MPO audit event:", error));
      }
      setMpos(items => items.filter(item => !selectedVisibleMpoIds.includes(item.id)));
      setSelectedMpoIds(ids => ids.filter(id => !selectedVisibleMpoIds.includes(id)));
      setConfirm(null);
      setToast({ msg: `${selectedVisibleMpoIds.length} MPO${selectedVisibleMpoIds.length !== 1 ? "s" : ""} deleted.`, type: "success" });
    } catch (e) {
      setToast({ msg: e.message || "Failed to delete selected MPOs.", type: "error" });
    } finally {
      setDeleteProgress(null);
    }
  };

  const openPreview = async (mpo) => {
    let previewSource = mpo;
    if (mpo?.id) {
      try {
        previewSource = await fetchMappedMpoById(mpo.id);
      } catch (error) {
        console.error("Failed to refresh MPO before preview:", error);
      }
    }
    const safeMpo = sanitizeMPOForExport({ ...previewSource, preparedSignature: previewSource?.preparedSignature || user?.signatureDataUrl || "", signedSignature: previewSource?.signedSignature || "", agencyEmail: previewSource?.agencyEmail || user?.agencyEmail || "", agencyPhone: previewSource?.agencyPhone || user?.agencyPhone || "" });
    const html = buildMPOHTML(safeMpo);
    const pdfBytes = buildMpoPdfBytes(safeMpo);
    const csvHeaders = ["Programme","WD","Time Belt","Material","Duration","Rate/Spot","Spots","Gross Value"];
    const csvRows = (previewSource.spots || []).map(s => [s.programme, s.wd, s.timeBelt, s.material, s.duration+'"', s.ratePerSpot, s.spots, (parseFloat(s.spots||0))*(parseFloat(s.ratePerSpot||0))]);
    csvRows.push([]);
    csvRows.push(["","","","","","Total Spots","",previewSource.totalSpots]);
    csvRows.push(["","","","","","Gross Value","",previewSource.totalGross]);
    csvRows.push(["","","","","","Discount","",-Math.round(previewSource.discAmt||0)]);
    csvRows.push(["","","","","","Net after Disc","",Math.round(previewSource.lessDisc||0)]);
    csvRows.push(["","","","","","Agency Commission","",-Math.round(previewSource.commAmt||0)]);
    csvRows.push(["","","","","","Net Value","",Math.round(previewSource.netVal||0)]);
    csvRows.push(["","","","","",`VAT (${previewSource.vatPct || 7.5}%)`,"",Math.round(previewSource.vatAmt||0)]);
    csvRows.push(["","","","","","TOTAL PAYABLE","",Math.round(previewSource.grandTotal||0)]);
    const csv = buildCSV(csvRows, csvHeaders);
    setPreview({ html, csv, title: `MPO — ${previewSource.mpoNo || "Draft"} | ${previewSource.vendorName} | ${previewSource.month} ${previewSource.year}` });
  };

  const statusColors = MPO_STATUS_COLORS;
  const getMpoClientId = (mpo) => {
    if (mpo?.clientId) return String(mpo.clientId);
    const campaignForMpo = campaigns.find((campaignItem) => String(campaignItem.id || "") === String(mpo?.campaignId || ""));
    return String(campaignForMpo?.clientId || "");
  };
  const visibleMpos = (viewMode === "archived" ? archivedOnly(mpos) : viewMode === "all" ? mpos : activeOnly(mpos)).filter(m => {
    const q = `${m.mpoNo || ""} ${m.vendorName || ""} ${m.clientName || ""} ${m.brand || ""} ${m.campaignName || ""}`.toLowerCase();
    return q.includes(searchTerm.toLowerCase())
      && (statusFilter === "all" || (m.status || "draft") === statusFilter)
      && (campaignFilter === "all" || String(m.campaignId || "") === campaignFilter)
      && (vendorFilter === "all" || String(m.vendorId || "") === vendorFilter)
      && (clientFilter === "all" || getMpoClientId(m) === clientFilter);
  });
  const campaignFilterOptions = [{ value: "all", label: "All Campaigns" }, ...campaigns.map(c => ({ value: c.id, label: c.name }))];
  const vendorFilterOptions = [{ value: "all", label: "All Vendors" }, ...vendors.map(v => ({ value: v.id, label: v.name }))];
  const clientFilterOptions = [{ value: "all", label: "All Clients" }, ...clients.map(c => ({ value: c.id, label: c.name }))];
  const selectedVisibleMpoIds = selectedMpoIds.filter(id => visibleMpos.some(m => m.id === id));
  const hasSelectedMpos = selectedVisibleMpoIds.length > 0;
  const allVisibleSelected = visibleMpos.length > 0 && selectedVisibleMpoIds.length === visibleMpos.length;
  const toggleSelectMpo = (mpoId) => setSelectedMpoIds(ids => ids.includes(mpoId) ? ids.filter(id => id !== mpoId) : [...ids, mpoId]);
  const toggleSelectAllVisibleMpos = () => setSelectedMpoIds(ids => allVisibleSelected ? ids.filter(id => !visibleMpos.some(m => m.id === id)) : Array.from(new Set([...ids, ...visibleMpos.map(m => m.id)])));
  const workflowStats = {
    myQueue: visibleMpos.filter(m => isMpoAwaitingUser(user, m)).length,
    pendingReview: visibleMpos.filter(m => ["submitted", "reviewed"].includes(String(m.status || "draft").toLowerCase())).length,
    readyToSend: visibleMpos.filter(m => String(m.status || "draft").toLowerCase() === "approved").length,
    needsChanges: visibleMpos.filter(m => String(m.status || "draft").toLowerCase() === "rejected").length,
  };
  const myWorkflowQueue = [...visibleMpos]
    .filter(m => isMpoAwaitingUser(user, m))
    .sort((a, b) => (b.updatedAt || b.createdAt || 0) - (a.updatedAt || a.createdAt || 0))
    .slice(0, 6);
  const workflowLaneCounts = MPO_STATUS_OPTIONS.reduce((acc, option) => {
    acc[option.value] = visibleMpos.filter(m => String(m.status || "draft").toLowerCase() === option.value).length;
    return acc;
  }, {});
  const topQuickQueue = myWorkflowQueue.slice(0, 3);
  const visibleMpoCards = [...visibleMpos].sort((a, b) => b.createdAt - a.createdAt);


  if (view === "list") return (
    <div className="fade mpo-page-scroll" style={{ width: "100%" }}>
      {toast && <Toast msg={toast.msg} type={toast.type} onDone={() => setToast(null)} />}
      {deleteProgress && (
        <div
          style={{
            position: "fixed",
            bottom: toast ? 86 : 22,
            right: 22,
            zIndex: 9998,
            minWidth: 280,
            maxWidth: 360,
            background: "var(--bg4)",
            color: "var(--text)",
            padding: "12px 16px",
            borderRadius: 12,
            border: "1px solid var(--border2)",
            boxShadow: "0 8px 28px rgba(0,0,0,.35)",
          }}
        >
          <div style={{ fontSize: 11, fontWeight: 700, color: "var(--accent)", textTransform: "uppercase", letterSpacing: ".06em", marginBottom: 7 }}>
            Bulk Delete
          </div>
          <div style={{ fontSize: 13, fontWeight: 600, marginBottom: 9 }}>
            {deleteProgress.label}
          </div>
          <div style={{ height: 8, borderRadius: 999, overflow: "hidden", background: "rgba(255,255,255,.08)" }}>
            <div
              style={{
                width: `${deleteProgress.total ? Math.round((deleteProgress.current / deleteProgress.total) * 100) : 0}%`,
                height: "100%",
                background: "linear-gradient(90deg,var(--accent),#f7c66a)",
                transition: "width .18s ease",
              }}
            />
          </div>
          <div style={{ marginTop: 8, fontSize: 12, color: "var(--text3)" }}>
            {deleteProgress.current} of {deleteProgress.total} completed
          </div>
        </div>
      )}
      {confirm && <Confirm msg={confirm.msg} onYes={confirm.onYes} onNo={() => setConfirm(null)} />}
      {preview && <PrintPreview html={preview.html} csv={preview.csv} title={preview.title} onClose={() => setPreview(null)} />}
      {mediaPlanImportOpen && (
        <MediaPlanImportModal
          vendors={vendors}
          clients={clients}
          campaigns={campaigns}
          rates={rates}
          user={user}
          appSettings={appSettings}
          setMpos={setMpos}
          setVendors={setVendors}
          onClose={() => setMediaPlanImportOpen(false)}
          onToast={setToast}
          onBulkImportStateChange={onBulkImportStateChange}
          requestMpoRefresh={requestMpoRefresh}
        />
      )}
      {historyModal && (
        <Modal title={historyModal.title || "MPO History"} onClose={() => setHistoryModal(null)} width={760}>
          {historyModal.events?.length ? (
            <div style={{ display: "flex", flexDirection: "column", gap: 10, maxHeight: "70vh", overflowY: "auto" }}>
              {historyModal.events.map(event => (
                <div key={event.id} style={{ border: "1px solid var(--border)", borderRadius: 10, padding: "12px 14px", background: "var(--bg3)" }}>
                  <div style={{ display: "flex", justifyContent: "space-between", gap: 12, flexWrap: "wrap" }}>
                    <div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 14 }}>
                      {(event.action || "updated").replace(/_/g, " ").toUpperCase()}
                    </div>
                    <div style={{ fontSize: 12, color: "var(--text3)" }}>{formatAuditTimestamp(event.createdAt)}</div>
                  </div>
                  <div style={{ fontSize: 12, color: "var(--text2)", marginTop: 4 }}>
                    {event.actorName || "System"} · {formatRoleLabel(event.actorRole || "viewer")}
                  </div>
                  {event.note ? <div style={{ marginTop: 8, fontSize: 13 }}>{event.note}</div> : null}
                  {event.metadata && Object.keys(event.metadata || {}).length ? (
                    <div style={{ marginTop: 10, display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(180px,1fr))", gap: 8 }}>
                      {Object.entries(event.metadata).map(([key, value]) => (
                        <div key={key} style={{ background: "var(--bg2)", border: "1px solid var(--border)", borderRadius: 8, padding: "8px 10px", fontSize: 12 }}>
                          <div style={{ color: "var(--text3)", textTransform: "uppercase", letterSpacing: ".04em", fontSize: 10, fontWeight: 700 }}>{key.replace(/_/g, " ")}</div>
                          <div style={{ marginTop: 4, wordBreak: "break-word" }}>{Array.isArray(value) ? value.join(", ") : String(value ?? "—")}</div>
                        </div>
                      ))}
                    </div>
                  ) : null}
                </div>
              ))}
            </div>
          ) : <Empty icon="🕘" title="No history yet" sub="This MPO has no audit events yet." />}
        </Modal>
      )}
      {executionModal && (
        <Modal title={`Execution & Reconciliation — ${executionModal.mpoNo || "MPO"}`} onClose={() => setExecutionModal(null)} width={920}>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(240px,1fr))", gap: 14 }}>
            <Card style={{ padding: 14 }}>
              <div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, marginBottom: 10 }}>Dispatch & Documents</div>
              <div style={{ display: "grid", gap: 10 }}>
                <Field label="Dispatch Status" value={executionModal.dispatchStatus} onChange={value => setExecutionModal(m => ({ ...m, dispatchStatus: value }))} options={MPO_EXECUTION_STATUS_OPTIONS} />
                <Field label="Dispatched At" type="datetime-local" value={executionModal.dispatchedAt} onChange={value => setExecutionModal(m => ({ ...m, dispatchedAt: value }))} />
                <Field label="Vendor Contact Used" value={executionModal.dispatchContact} onChange={value => setExecutionModal(m => ({ ...m, dispatchContact: value }))} placeholder="Email, phone, or contact name" />
                <AttachmentField label="Signed MPO" url={executionModal.signedMpoUrl} onUrlChange={value => setExecutionModal(m => ({ ...m, signedMpoUrl: value }))} onFileSelected={file => uploadExecutionAttachment("signedMpo", file)} uploading={executionUploading.signedMpo} />
                <Field label="Dispatch Note" rows={3} value={executionModal.dispatchNote} onChange={value => setExecutionModal(m => ({ ...m, dispatchNote: value }))} placeholder="Who sent it, what was shared, any follow-up note..." />
              </div>
            </Card>
            <Card style={{ padding: 14 }}>
              <div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, marginBottom: 10 }}>Invoice & Payment</div>
              <div style={{ display: "grid", gap: 10 }}>
                <Field label="Invoice Status" value={executionModal.invoiceStatus} onChange={value => setExecutionModal(m => ({ ...m, invoiceStatus: value }))} options={MPO_INVOICE_STATUS_OPTIONS} />
                <Field label="Invoice Number" value={executionModal.invoiceNo} onChange={value => setExecutionModal(m => ({ ...m, invoiceNo: value }))} />
                <Field label="Invoice Amount" type="number" value={executionModal.invoiceAmount} onChange={value => setExecutionModal(m => ({ ...m, invoiceAmount: value }))} />
                <Field label="Invoice Received At" type="datetime-local" value={executionModal.invoiceReceivedAt} onChange={value => setExecutionModal(m => ({ ...m, invoiceReceivedAt: value }))} />
                <AttachmentField label="Invoice" url={executionModal.invoiceUrl} onUrlChange={value => setExecutionModal(m => ({ ...m, invoiceUrl: value }))} onFileSelected={file => uploadExecutionAttachment("invoice", file)} uploading={executionUploading.invoice} />
                <Field label="Payment Status" value={executionModal.paymentStatus} onChange={value => setExecutionModal(m => ({ ...m, paymentStatus: value }))} options={MPO_PAYMENT_STATUS_OPTIONS} />
                <Field label="Payment Reference" value={executionModal.paymentReference} onChange={value => setExecutionModal(m => ({ ...m, paymentReference: value }))} />
                <Field label="Paid At" type="datetime-local" value={executionModal.paidAt} onChange={value => setExecutionModal(m => ({ ...m, paidAt: value }))} />
              </div>
            </Card>
            <Card style={{ padding: 14, gridColumn: "1 / -1" }}>
              <div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, marginBottom: 10 }}>Proof of Airing & Reconciliation</div>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(180px,1fr))", gap: 10 }}>
                <Field label="Proof Status" value={executionModal.proofStatus} onChange={value => setExecutionModal(m => ({ ...m, proofStatus: value }))} options={MPO_PROOF_STATUS_OPTIONS} />
                <AttachmentField label="Proof of Airing" url={executionModal.proofUrl} onUrlChange={value => setExecutionModal(m => ({ ...m, proofUrl: value }))} onFileSelected={file => uploadExecutionAttachment("proof", file)} uploading={executionUploading.proof} />
                <Field label="Proof Received At" type="datetime-local" value={executionModal.proofReceivedAt} onChange={value => setExecutionModal(m => ({ ...m, proofReceivedAt: value }))} />
                <Field label="Planned Spots" type="number" value={executionModal.plannedSpotsExecution} onChange={value => setExecutionModal(m => ({ ...m, plannedSpotsExecution: value }))} />
                <Field label="Aired Spots" type="number" value={executionModal.airedSpots} onChange={value => setExecutionModal(m => ({ ...m, airedSpots: value }))} />
                <Field label="Missed Spots" type="number" value={executionModal.missedSpots} onChange={value => setExecutionModal(m => ({ ...m, missedSpots: value }))} />
                <Field label="Make-good Spots" type="number" value={executionModal.makegoodSpots} onChange={value => setExecutionModal(m => ({ ...m, makegoodSpots: value }))} />
                <Field label="Reconciliation Status" value={executionModal.reconciliationStatus} onChange={value => setExecutionModal(m => ({ ...m, reconciliationStatus: value }))} options={MPO_RECON_STATUS_OPTIONS} />
                <Field label="Reconciled Amount" type="number" value={executionModal.reconciledAmount} onChange={value => setExecutionModal(m => ({ ...m, reconciledAmount: value }))} />
              </div>
              <div style={{ marginTop: 10 }}>
                <Field label="Reconciliation Notes" rows={4} value={executionModal.reconciliationNotes} onChange={value => setExecutionModal(m => ({ ...m, reconciliationNotes: value }))} placeholder="Delivered spots, discrepancies, make-goods, invoice notes, approvals..." />
              </div>
              <div style={{ display: "flex", justifyContent: "flex-end", gap: 10, marginTop: 14 }}>
                <Btn variant="ghost" onClick={() => setExecutionModal(null)}>Cancel</Btn>
                <Btn onClick={saveExecutionModal} loading={executionUploading.signedMpo || executionUploading.invoice || executionUploading.proof}>Save Execution</Btn>
              </div>
            </Card>
          </div>
        </Modal>
      )}
      {statusModal && (
        <Modal title={statusModal.actionLabel || `Move MPO to ${MPO_STATUS_LABELS[statusModal.nextStatus] || statusModal.nextStatus}`} onClose={() => setStatusModal(null)} width={560}>
          <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
            <div style={{ fontSize: 13, color: "var(--text2)" }}>
              {statusModal.mpo?.mpoNo || "MPO"} will move from <strong>{MPO_STATUS_LABELS[statusModal.mpo?.status || "draft"] || (statusModal.mpo?.status || "draft")}</strong> to <strong>{MPO_STATUS_LABELS[statusModal.nextStatus] || statusModal.nextStatus}</strong>.
            </div>
            <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
              <Badge color={getMpoWorkflowMeta(statusModal.mpo).color}>{getMpoWorkflowMeta(statusModal.mpo).label}</Badge>
              <Badge color={getMpoWorkflowMeta({ ...statusModal.mpo, status: statusModal.nextStatus }).color}>Next up: {statusModal.nextOwnerLabel || getMpoWorkflowMeta({ ...statusModal.mpo, status: statusModal.nextStatus }).label}</Badge>
            </div>
            <div style={{ fontSize: 12, color: "var(--text2)", background: "var(--bg3)", borderRadius: 10, padding: "10px 12px", border: "1px solid var(--border)" }}>
              {statusModal.helperText || getMpoWorkflowMeta({ ...statusModal.mpo, status: statusModal.nextStatus }).hint}
            </div>
            <Field
              label={statusModal.noteLabel || (statusModal.nextStatus === "rejected" ? "Rejection note" : "Approval / workflow note")}
              rows={4}
              value={statusModal.note || ""}
              onChange={value => setStatusModal(modal => ({ ...modal, note: value }))}
              placeholder={statusModal.nextStatus === "rejected" ? "What should be corrected before this MPO can move forward?" : "Add context for the next team member..."}
              required
            />
            <div style={{ display: "flex", justifyContent: "flex-end", gap: 10 }}>
              <Btn variant="ghost" onClick={() => setStatusModal(null)}>Cancel</Btn>
              <Btn variant={getWorkflowActionVariant(statusModal.nextStatus)} onClick={() => {
                if (mpoStatusNeedsNote(statusModal.nextStatus) && !String(statusModal.note || "").trim()) {
                  setToast({ msg: "A note is required for this workflow action.", type: "error" });
                  return;
                }
                applyMpoStatusChange(statusModal.mpo, statusModal.nextStatus, statusModal.note || "");
              }}>{statusModal.actionLabel || "Confirm Status Change"}</Btn>
            </div>
          </div>
        </Modal>
      )}
      <style>{`
        .mpo-page-scroll {
          width: 100%;
          max-height: calc(100vh - 20px);
          overflow-y: auto;
          overflow-x: hidden;
          padding-right: 8px;
          scrollbar-width: auto;
          scrollbar-color: rgba(160,160,160,.95) rgba(20,24,32,.9);
          scrollbar-gutter: stable;
        }
        .mpo-page-scroll::-webkit-scrollbar {
          width: 12px;
        }
        .mpo-page-scroll::-webkit-scrollbar-track {
          background: rgba(20,24,32,.9);
          border-radius: 999px;
        }
        .mpo-page-scroll::-webkit-scrollbar-thumb {
          background: rgba(190,190,190,.95);
          border-radius: 999px;
          border: 2px solid rgba(20,24,32,.9);
        }
        @media (max-width: 1520px) {
          .mpo-list-card-grid {
            grid-template-columns: 72px minmax(280px,1fr) !important;
          }
          .mpo-list-card-main,
          .mpo-list-card-side,
          .mpo-list-card-actions {
            grid-column: 2 / -1 !important;
          }
          .mpo-list-card-side {
            justify-self: stretch !important;
          }
          .mpo-list-card-summary {
            justify-content: center !important;
          }
        }
        @media (max-width: 980px) {
          .mpo-list-card-grid {
            grid-template-columns: 1fr !important;
          }
          .mpo-list-card-icon,
          .mpo-list-card-main,
          .mpo-list-card-side,
          .mpo-list-card-actions {
            grid-column: auto !important;
          }
          .mpo-list-card-main,
          .mpo-list-card-actions,
          .mpo-list-card-summary {
            text-align: left !important;
            justify-content: flex-start !important;
          }
          .mpo-list-card-main-badges {
            justify-content: flex-start !important;
          }
          .mpo-list-card-summary {
            grid-template-columns: 1fr !important;
          }
        }
      `}</style>
      <div style={{ marginBottom: 18 }}>
        <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", marginBottom: 10, flexWrap: "wrap", gap: 10 }}>
          <div>
            <h1 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 800, fontSize: 24 }}>MPO Generator</h1>
            <p style={{ color: "var(--text2)", marginTop: 3, fontSize: 13 }}>Create, manage & export Media Purchase Orders</p>
          </div>
        </div>
      </div>
      <div style={{ position: "sticky", top: 0, zIndex: 30, marginBottom: 24 }}>
        <div style={{ background: "var(--bg)", padding: "10px 0 12px", borderBottom: "1px solid var(--border)" }}>
          <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center" }}>
            <Field value={searchTerm} onChange={setSearchTerm} placeholder="Search MPO, vendor, client..." />
            <Field value={campaignFilter} onChange={setCampaignFilter} options={campaignFilterOptions} />
            <Field value={vendorFilter} onChange={setVendorFilter} options={vendorFilterOptions} />
            <Field value={clientFilter} onChange={setClientFilter} options={clientFilterOptions} />
            <Field value={statusFilter} onChange={setStatusFilter} options={[{value:"all",label:"All Statuses"}, ...MPO_STATUS_OPTIONS.map(o => ({ value: o.value, label: o.label }))]} />
            <Field value={viewMode} onChange={setViewMode} options={[{value:"active",label:"Active"},{value:"archived",label:"Archived"},{value:"all",label:"All"}]} />
            {canManage && <><Btn variant="ghost" onClick={() => setMediaPlanImportOpen(true)}>Import Media Plan</Btn><Btn icon="+" onClick={openNew}>New MPO</Btn></>}
          </div>
          {canManage && hasSelectedMpos && (
            <Card style={{ marginTop: 12, padding: "14px 18px" }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
                <label style={{ display: "flex", alignItems: "center", gap: 10, fontSize: 13, color: "var(--text2)", fontWeight: 600 }}>
                  <input type="checkbox" checked={allVisibleSelected} onChange={toggleSelectAllVisibleMpos} style={{ width: 16, height: 16, accentColor: "var(--accent)" }} />
                  Select all visible MPOs
                </label>
                <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center" }}>
                  <Badge color={selectedVisibleMpoIds.length ? "accent" : "blue"}>{selectedVisibleMpoIds.length} selected</Badge>
                  <Btn variant="ghost" size="sm" onClick={() => setSelectedMpoIds([])}>Clear Selection</Btn>
                  <Btn variant="danger" size="sm" onClick={() => setConfirm({ msg: `Archive ${selectedVisibleMpoIds.length} selected MPO${selectedVisibleMpoIds.length !== 1 ? "s" : ""}?`, onYes: archiveSelectedMpos })} disabled={Boolean(deleteProgress)}>Archive Selected</Btn>
                  <Btn
                    variant="danger"
                    size="sm"
                    onClick={() => setConfirm({ msg: `Delete ${selectedVisibleMpoIds.length} selected MPO${selectedVisibleMpoIds.length !== 1 ? "s" : ""}? This cannot be undone.`, onYes: deleteSelectedMpos })}
                    disabled={Boolean(deleteProgress)}
                    loading={Boolean(deleteProgress)}
                  >
                    {deleteProgress ? `Deleting ${deleteProgress.current}/${deleteProgress.total}` : "Delete Selected"}
                  </Btn>
                </div>
              </div>
            </Card>
          )}
        </div>
      </div>
      {savedDrafts.length > 0 && (
        <Card style={{ marginBottom: 14, padding: "14px 18px", background: "rgba(59,126,245,.08)", border: "1px solid rgba(59,126,245,.22)" }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 12, flexWrap: "wrap", marginBottom: 12 }}>
            <div>
              <div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 14 }}>Saved MPO drafts</div>
              <div style={{ fontSize: 12, color: "var(--text2)", marginTop: 3 }}>
                Any in-progress MPO you started and left unfinished is saved here so you can continue later.
              </div>
            </div>
            <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center" }}>
              <Badge color="blue">{savedDrafts.length} draft{savedDrafts.length !== 1 ? "s" : ""}</Badge>
              <Btn variant="ghost" size="sm" onClick={() => clearSavedDraft(null)}>Clear All</Btn>
            </div>
          </div>
          <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
            {savedDrafts.map((draft, draftIndex) => (
              <div key={`${draft.id || "draft"}-${draftIndex}`} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 12, flexWrap: "wrap", padding: "10px 12px", borderRadius: 10, background: "var(--bg2)", border: "1px solid var(--border)" }}>
                <div>
                  <div style={{ fontWeight: 700, fontSize: 13 }}>{buildDraftTitle(draft)}</div>
                  <div style={{ fontSize: 12, color: "var(--text2)", marginTop: 3 }}>
                    {buildDraftSubtitle(draft) || "Draft MPO"}{draft.savedAt ? ` · Saved ${new Date(draft.savedAt).toLocaleString()}` : ""}
                  </div>
                </div>
                <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                  <Btn variant="blue" size="sm" onClick={() => resumeSavedDraft(draft.id)}>Continue</Btn>
                  <Btn variant="ghost" size="sm" onClick={() => clearSavedDraft(draft.id)}>Delete Draft</Btn>
                </div>
              </div>
            ))}
          </div>
        </Card>
      )}
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(170px,1fr))", gap: 12, marginBottom: 14 }}>
        <Stat icon="⏳" label="My Queue" value={workflowStats.myQueue} sub="MPOs currently waiting on your role" color="var(--blue)" />
        <Stat icon="🧾" label="Pending Review" value={workflowStats.pendingReview} sub="Submitted and reviewed MPOs in the approval lane" color="var(--purple)" />
        <Stat icon="📤" label="Ready to Send" value={workflowStats.readyToSend} sub="Approved MPOs waiting for dispatch" color="var(--teal)" />
        <Stat icon="🛠" label="Needs Changes" value={workflowStats.needsChanges} sub="Rejected MPOs waiting for revision" color="var(--red)" />
      </div>
      <Card style={{ marginBottom: 14, padding: "16px 18px" }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, flexWrap: "wrap", marginBottom: workflowPanelOpen ? 14 : 0 }}>
          <div>
            <div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 15 }}>Approvals Automation</div>
            <div style={{ fontSize: 12, color: "var(--text2)", marginTop: 4 }}>Role-based queue, waiting-on visibility, and one-click workflow actions.</div>
          </div>
          <Btn variant="ghost" size="sm" onClick={() => setWorkflowPanelOpen(open => !open)}>{workflowPanelOpen ? "Hide" : "Show"} queue</Btn>
        </div>
        {workflowPanelOpen && (
          <div style={{ display: "grid", gridTemplateColumns: "1.3fr .9fr", gap: 14 }}>
            <div style={{ border: "1px solid var(--border)", borderRadius: 12, padding: 14, background: "var(--bg3)" }}>
              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 10 }}>
                <div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 14 }}>My approval / action queue</div>
                <Badge color="blue">{workflowStats.myQueue} item{workflowStats.myQueue !== 1 ? "s" : ""}</Badge>
              </div>
              {topQuickQueue.length === 0 ? (
                <div style={{ fontSize: 12, color: "var(--text2)" }}>Nothing is currently waiting on your role.</div>
              ) : (
                <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                  {topQuickQueue.map((mpo, mpoIndex) => {
                    const workflowMeta = getMpoWorkflowMeta(mpo);
                    const quickActions = getQuickWorkflowActions(user, mpo).slice(0, 2);
                    return (
                      <div key={`${mpo.id || "queue"}-${mpoIndex}`} style={{ border: "1px solid var(--border)", borderRadius: 10, padding: "12px 13px", background: "var(--bg2)" }}>
                        <div style={{ display: "flex", justifyContent: "space-between", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
                          <div style={{ fontWeight: 700, fontSize: 13 }}>{mpo.mpoNo || "MPO"} · {mpo.vendorName || "Vendor"}</div>
                          <Badge color={workflowMeta.color}>Waiting on you</Badge>
                        </div>
                        <div style={{ fontSize: 12, color: "var(--text2)", marginTop: 5 }}>{workflowMeta.hint}</div>
                        <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginTop: 10 }}>
                          {quickActions.map(action => (
                            <Btn key={action.value} size="sm" variant={action.variant} onClick={() => requestMpoStatusChange(mpo, action.value)}>{action.label}</Btn>
                          ))}
                          <Btn size="sm" variant="ghost" onClick={() => openMpoHistory(mpo)}>History</Btn>
                        </div>
                      </div>
                    );
                  })}
                </div>
              )}
            </div>
            <div style={{ border: "1px solid var(--border)", borderRadius: 12, padding: 14, background: "var(--bg3)" }}>
              <div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 14, marginBottom: 10 }}>Workflow lanes</div>
              <div style={{ display: "grid", gap: 8 }}>
                {MPO_STATUS_OPTIONS.map(option => (
                  <div key={option.value} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "9px 10px", borderRadius: 9, background: "var(--bg2)", border: "1px solid var(--border)" }}>
                    <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                      <Badge color={statusColors[option.value] || "accent"}>{option.label}</Badge>
                      <span style={{ fontSize: 12, color: "var(--text2)" }}>{getMpoWorkflowMeta({ status: option.value }).label}</span>
                    </div>
                    <strong style={{ fontFamily: "'Syne',sans-serif" }}>{workflowLaneCounts[option.value] || 0}</strong>
                  </div>
                ))}
              </div>
            </div>
          </div>
        )}
      </Card>
      <div ref={mpoListScrollRef} style={{ width: "100%", overflowX: "auto", overflowY: "hidden", paddingBottom: 2 }}>
        <div style={{ minWidth: 1100, paddingBottom: 2 }}>
      {visibleMpos.length === 0 ? <Card><Empty icon="📄" title="No MPOs yet" sub="Create your first Media Purchase Order" /></Card> :
        <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
          {visibleMpoCards.map((m, cardIndex) => (
            <Card key={`${m.id || "mpo"}-${cardIndex}`} style={{ padding: "12px 16px" }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10, marginBottom: 8, flexWrap: "wrap" }}>
                {canManage ? (
                  <label style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 12, color: "var(--text2)", fontWeight: 600 }}>
                    <input type="checkbox" checked={selectedMpoIds.includes(m.id)} onChange={() => toggleSelectMpo(m.id)} style={{ width: 16, height: 16, accentColor: "var(--accent)" }} />
                    Select MPO
                  </label>
                ) : <div />}
                {m.campaignName ? <Badge color="blue">{m.campaignName}</Badge> : null}
              </div>
              <div
                className="mpo-list-card-grid"
                style={{
                  display: "grid",
                  gridTemplateColumns: "74px minmax(420px,1fr) minmax(290px,360px)",
                  columnGap: 14,
                  rowGap: 12,
                  alignItems: "start",
                }}
              >
                <div
                  className="mpo-list-card-icon"
                  style={{
                    width: 60,
                    height: 60,
                    background: "rgba(240,165,0,.08)",
                    border: "1px solid rgba(240,165,0,.22)",
                    borderRadius: 16,
                    display: "flex",
                    alignItems: "center",
                    justifyContent: "center",
                    fontSize: 26,
                    color: "var(--accent)",
                    gridRow: "1 / span 2",
                    alignSelf: "start",
                    justifySelf: "center",
                  }}
                >
                  📄
                </div>

                <div
                  className="mpo-list-card-main"
                  style={{
                    minWidth: 0,
                    textAlign: "center",
                    paddingTop: 2,
                  }}
                >
                  <div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 800, fontSize: 16, lineHeight: 1.2 }}>
                    {m.mpoNo || "MPO"} <span style={{ color: "var(--text3)", fontSize: 14 }}>— {m.vendorName}</span>
                  </div>
                  <div style={{ fontSize: 13, color: "var(--text2)", marginTop: 8, lineHeight: 1.45 }}>
                    {m.clientName} {m.brand && `· ${m.brand}`} · {(m.months || []).length > 1 ? (m.months || []).join("-") : m.month} {m.year} · {m.totalSpots} spots
                  </div>
                  <div
                    className="mpo-list-card-main-badges"
                    style={{ display: "flex", gap: 8, flexWrap: "wrap", marginTop: 12, justifyContent: "center" }}
                  >
                    <Badge color={statusColors[m.status || "draft"] || "accent"}>{MPO_STATUS_LABELS[m.status || "draft"] || (m.status || "draft")}</Badge>
                    <Badge color={getMpoWorkflowMeta(m).color}>Waiting on: {getMpoWorkflowMeta(m).label}</Badge>
                    {isMpoAwaitingUser(user, m) ? <Badge color="blue">My Queue</Badge> : null}
                    <Badge color={getExecutionHealthColor(m)}>{getExecutionHealthLabel(m)}</Badge>
                    <Badge color="blue">Invoice: {(MPO_INVOICE_STATUS_OPTIONS.find(o => o.value === (m.invoiceStatus || "pending"))?.label || m.invoiceStatus || "pending")}</Badge>
                    <Badge color="purple">Proof: {(MPO_PROOF_STATUS_OPTIONS.find(o => o.value === (m.proofStatus || "pending"))?.label || m.proofStatus || "pending")}</Badge>
                    <Badge color="green">Payment: {(MPO_PAYMENT_STATUS_OPTIONS.find(o => o.value === (m.paymentStatus || "unpaid"))?.label || m.paymentStatus || "unpaid")}</Badge>
                    {isArchived(m) ? <Badge color="red">Archived</Badge> : null}
                  </div>
                  <div style={{ fontSize: 12, color: "var(--text3)", marginTop: 12, lineHeight: 1.45 }}>
                    {getMpoWorkflowMeta(m).hint}
                  </div>
                </div>

                <div
                  className="mpo-list-card-side"
                  style={{
                    minWidth: 0,
                    justifySelf: "end",
                    width: "100%",
                    maxWidth: 360,
                  }}
                >
                  <div
                    className="mpo-list-card-summary"
                    style={{
                      display: "grid",
                      gridTemplateColumns: "minmax(82px,1fr) minmax(112px,1fr) minmax(138px,160px)",
                      gap: 14,
                      alignItems: "center",
                      justifyContent: "center",
                    }}
                  >
                    <div style={{ textAlign: "center" }}>
                      <div style={{ fontSize: 10, color: "var(--text3)", fontWeight: 700, textTransform: "uppercase", letterSpacing: ".04em" }}>Gross</div>
                      <div style={{ fontSize: 14, fontWeight: 700, marginTop: 8, color: "var(--text2)" }}>{fmtN(m.totalGross)}</div>
                    </div>
                    <div style={{ textAlign: "center" }}>
                      <div style={{ fontSize: 10, color: "var(--text3)", fontWeight: 700, textTransform: "uppercase", letterSpacing: ".04em" }}>Net Payable</div>
                      <div style={{ fontSize: 14, fontWeight: 800, marginTop: 8, color: "var(--accent)" }}>{fmtN(m.grandTotal || m.netVal)}</div>
                    </div>
                    <div style={{ minWidth: 0 }}>
                      {canManageStatus ? (
                        <Field
                          value={m.status || "draft"}
                          onChange={v => requestMpoStatusChange(m, v)}
                          options={[
                            { value: m.status || "draft", label: MPO_STATUS_LABELS[m.status || "draft"] || (m.status || "draft") },
                            ...getAllowedMpoStatusTargets(user, m).map(value => ({ value, label: MPO_STATUS_LABELS[value] || value })),
                          ].filter((option, index, arr) => arr.findIndex(item => item.value === option.value) === index)}
                        />
                      ) : (
                        <div style={{ paddingTop: 8, textAlign: "center" }}>
                          <Badge color={statusColors[m.status || "draft"] || "accent"}>{(m.status || "draft").toUpperCase()}</Badge>
                        </div>
                      )}
                    </div>
                  </div>
                </div>

                <div
                  className="mpo-list-card-actions"
                  style={{
                    gridColumn: "2 / -1",
                    display: "flex",
                    gap: 10,
                    flexWrap: "wrap",
                    alignItems: "center",
                    justifyContent: "flex-start",
                  }}
                >
                  {getQuickWorkflowActions(user, m).slice(0, 2).map(action => (
                    <Btn key={action.value} variant={action.variant} size="sm" onClick={() => requestMpoStatusChange(m, action.value)}>{action.label}</Btn>
                  ))}
                  <Btn variant="ghost" size="sm" onClick={() => openEdit(m)} icon="✏️">Edit</Btn>
                  <Btn variant="ghost" size="sm" onClick={() => openExecutionModal(m)} icon="📦">Execution</Btn>
                  <Btn variant="ghost" size="sm" onClick={() => openMpoHistory(m)} icon="🕘">History</Btn>
                  <Btn variant="blue" size="sm" onClick={() => openPreview(m)} icon="⬇">Preview & Export</Btn>
                  {isArchived(m) ? <Btn variant="success" size="sm" onClick={() => restoreMPO(m.id)}>↩</Btn> : <Btn variant="danger" size="sm" onClick={() => setConfirm({ msg: `Archive MPO "${m.mpoNo || m.id}"?`, onYes: () => archiveMPO(m.id) })}>🗄</Btn>}
                  <Btn variant="danger" size="sm" onClick={() => setConfirm({ msg: `Delete MPO "${m.mpoNo || m.id}"? This cannot be undone.`, onYes: () => deleteMPO(m.id) })}>🗑</Btn>
                </div>
              </div>
            </Card>
          ))}
        </div>}
        </div>
      </div>
    </div>
  );

  // FORM VIEW
  const stepLabels = ["Campaign & Vendor", "Spot Schedule", "Costing", "Signatories & Save"];
  return (
    <div className="fade">
      {toast && <Toast msg={toast.msg} type={toast.type} onDone={() => setToast(null)} />}
      {vendorPickerOpen && (
        <Modal title="Search Media House / Vendor" onClose={() => setVendorPickerOpen(false)} width={720}>
          <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
            <Field
              label="Search vendor"
              value={vendorSearch}
              onChange={setVendorSearch}
              placeholder="Type vendor name, type, phone, or email"
            />
            <div style={{ maxHeight: "55vh", overflowY: "auto", border: "1px solid var(--border)", borderRadius: 12, background: "var(--bg3)" }}>
              {filteredVendors.length === 0 ? (
                <div style={{ padding: "18px 16px", fontSize: 13, color: "var(--text2)" }}>
                  No vendor matched “{vendorSearch}”.
                </div>
              ) : (
                filteredVendors.map((item, vendorIndex) => {
                  const isSelected = item.id === mpoData.vendorId;
                  return (
                    <button
                      key={`${item.id || "vendor"}-${vendorIndex}`}
                      type="button"
                      onClick={() => {
                        setMpoData(data => ({ ...data, vendorId: item.id }));
                        setVendorPickerOpen(false);
                      }}
                      style={{
                        width: "100%",
                        textAlign: "left",
                        padding: "13px 15px",
                        border: "none",
                        borderBottom: "1px solid var(--border)",
                        background: isSelected ? "rgba(240,165,0,.12)" : "transparent",
                        cursor: "pointer",
                      }}
                    >
                      <div style={{ display: "flex", justifyContent: "space-between", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
                        <div style={{ fontWeight: 700, fontSize: 13, color: isSelected ? "var(--accent)" : "var(--text)" }}>
                          {isSelected ? "✓ " : ""}{item.name}
                        </div>
                        <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                          {item.type ? <Badge color="blue">{item.type}</Badge> : null}
                          {item.discount ? <Badge color="green">Disc {item.discount}%</Badge> : null}
                          {item.commission ? <Badge color="purple">Comm {item.commission}%</Badge> : null}
                        </div>
                      </div>
                      <div style={{ marginTop: 6, fontSize: 12, color: "var(--text2)", display: "flex", gap: 12, flexWrap: "wrap" }}>
                        {item.contactName ? <span>Contact: {item.contactName}</span> : null}
                        {item.phone ? <span>{item.phone}</span> : null}
                        {item.email ? <span>{item.email}</span> : null}
                      </div>
                    </button>
                  );
                })
              )}
            </div>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
              <div style={{ fontSize: 12, color: "var(--text2)" }}>
                {filteredVendors.length} vendor{filteredVendors.length !== 1 ? "s" : ""} found
              </div>
              <Btn variant="ghost" onClick={() => setVendorPickerOpen(false)}>Close</Btn>
            </div>
          </div>
        </Modal>
      )}
      {spotModal && (
        <Modal title={editSpotId ? "Edit Spot Row" : "Add Spot Row"} onClose={() => { setSpotModal(null); setEditSpotId(null); setSpotForm(blankSpot); }} width={620}>

          {/* Vendor Rate Card Picker */}
          {vendorRates.length > 0 && (
            <div style={{ marginBottom: 18 }}>
              <div style={{ fontSize: 11, fontWeight: 600, color: "var(--text3)", textTransform: "uppercase", letterSpacing: ".06em", marginBottom: 10 }}>
                📺 Select from {vendor?.name || "Vendor"} Rate Cards — click a row to auto-fill
              </div>
              <div style={{ border: "1px solid var(--border2)", borderRadius: 10, overflow: "hidden", maxHeight: 220, overflowY: "auto" }}>
                <table style={{ width: "100%", borderCollapse: "collapse" }}>
                  <thead>
                    <tr style={{ background: "var(--bg3)", position: "sticky", top: 0 }}>
                      {["Programme","Time Belt","Duration","Rate/Spot","Disc %","Net Rate"].map(h => (
                        <th key={h} style={{ padding: "7px 10px", textAlign: "left", fontSize: 10, fontWeight: 600, textTransform: "uppercase", color: "var(--text3)", borderBottom: "1px solid var(--border)", whiteSpace: "nowrap" }}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {vendorRates.map((r, i) => {
                      const selected = spotForm.rateId === r.id;
                      const disc = parseFloat(r.discount) || 0;
                      const net = (parseFloat(r.ratePerSpot) || 0) * (1 - disc / 100);
                      return (
                        <tr key={r.id}
                          onClick={() => {
                            setSpotForm(f => ({
                              ...f,
                              rateId: r.id,
                              programme: r.programme || f.programme,
                              timeBelt: r.timeBelt || f.timeBelt,
                              duration: normalizeDurationValue(r.duration || f.duration, "30"),
                              customRate: "",
                            }));
                          }}
                          style={{ cursor: "pointer", background: selected ? "rgba(240,165,0,.13)" : i % 2 === 0 ? "transparent" : "rgba(255,255,255,.02)", borderBottom: "1px solid var(--border)", transition: "background .12s" }}
                          onMouseEnter={e => { if (!selected) e.currentTarget.style.background = "var(--bg3)"; }}
                          onMouseLeave={e => { if (!selected) e.currentTarget.style.background = i % 2 === 0 ? "transparent" : "rgba(255,255,255,.02)"; }}>
                          <td style={{ padding: "8px 10px", fontWeight: selected ? 700 : 500, fontSize: 12, color: selected ? "var(--accent)" : "var(--text)" }}>
                            {selected && <span style={{ marginRight: 5 }}>✓</span>}{r.programme || "—"}
                          </td>
                          <td style={{ padding: "8px 10px", fontSize: 12, color: "var(--text2)" }}>{r.timeBelt || "—"}</td>
                          <td style={{ padding: "8px 10px", fontSize: 12, color: "var(--text2)" }}>{r.duration ? `${r.duration}"` : "—"}</td>
                          <td style={{ padding: "8px 10px", fontSize: 12, color: "var(--accent)", fontWeight: 600 }}>{fmtN(r.ratePerSpot)}</td>
                          <td style={{ padding: "8px 10px", fontSize: 12, color: disc > 0 ? "var(--green)" : "var(--text3)" }}>{disc > 0 ? `${disc}%` : "—"}</td>
                          <td style={{ padding: "8px 10px", fontSize: 12, fontWeight: 700, color: "var(--green)" }}>{fmtN(net)}</td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
              {spotForm.rateId && (
                <div style={{ marginTop: 8, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                  <span style={{ fontSize: 11, color: "var(--green)" }}>✓ Rate card applied — fields auto-filled below</span>
                  <button onClick={() => setSpotForm(f => ({ ...f, rateId: "", customRate: "" }))} style={{ fontSize: 11, color: "var(--text3)", background: "none", border: "none", cursor: "pointer", textDecoration: "underline" }}>Clear selection</button>
                </div>
              )}
            </div>
          )}

          {vendorRates.length === 0 && (
            <div style={{ marginBottom: 16, background: "rgba(240,165,0,.07)", border: "1px solid rgba(240,165,0,.2)", borderRadius: 9, padding: "10px 14px", fontSize: 12, color: "var(--text2)" }}>
              ⚠️ No rate cards found for this vendor. Go to <strong style={{ color: "var(--accent)" }}>Media Rates</strong> to add them, or fill in the fields manually below.
            </div>
          )}

          {/* Manual / override fields */}
          <div style={{ borderTop: vendorRates.length > 0 ? "1px solid var(--border)" : "none", paddingTop: vendorRates.length > 0 ? 16 : 0 }}>
            <div style={{ fontSize: 11, fontWeight: 600, color: "var(--text3)", textTransform: "uppercase", letterSpacing: ".06em", marginBottom: 12 }}>
              {vendorRates.length > 0 ? "✏️ Review / Override Details" : "✏️ Spot Details"}
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 13 }}>
              <Field label="Programme" value={spotForm.programme} onChange={updS("programme")} placeholder="NTA News, SuperStory…" required />
              <Field label="Fixed Time / Time Belt" value={spotForm.timeBelt} onChange={updS("timeBelt")} placeholder="08:45–09:00" />
              <Field label="Duration (secs)" type="number" value={spotForm.duration} onChange={updS("duration")} placeholder="30" />
              <Field label="Day of Week" value={spotForm.wd} onChange={updS("wd")} options={["Mon","Tue","Wed","Thu","Fri","Sat","Sun","Daily","Weekdays","Weekends"]} />
              <div style={{ gridColumn: "1/-1" }}>
                {campaign?.materialList?.length > 0 ? (
                  <div style={{ display: "flex", flexDirection: "column", gap: 5 }}>
                    <label style={{ fontSize: 11, fontWeight: 600, color: "var(--text3)", textTransform: "uppercase", letterSpacing: ".08em" }}>Material / Spot Name</label>
                    <select value={spotForm.material} onChange={e => updS("material")(e.target.value)}
                      style={{ background: "var(--bg3)", border: "1px solid var(--border2)", borderRadius: 8, padding: "9px 13px", color: spotForm.material ? "var(--text)" : "var(--text3)", fontSize: 13, outline: "none", cursor: "pointer" }}>
                      <option value="">— Select material —</option>
                      {(campaign.materialList||[]).map((m,i) => <option key={i} value={m}>{m}</option>)}
                      <option value="__custom__">Custom…</option>
                    </select>
                    {spotForm.material === "__custom__" && (
                      <input value={spotForm.materialCustom || ""} onChange={e => updS("materialCustom")(e.target.value)} placeholder="Type material name" style={{ background: "var(--bg3)", border: "1px solid var(--border2)", borderRadius: 8, padding: "9px 13px", color: "var(--text)", fontSize: 13, outline: "none" }} />
                    )}
                  </div>
                ) : (
                  <Field label="Material / Spot Name" value={spotForm.material} onChange={updS("material")} placeholder="SM Thematic English 30secs (MP4)" />
                )}
              </div>
              <div style={{ gridColumn: "1/-1" }}>
                <Field label="Rate per Spot (₦)" type="number"
                  value={spotForm.isComplimentary ? "0" : (spotForm.customRate || (vendorRates.find(r => r.id === spotForm.rateId)?.ratePerSpot || ""))}
                  onChange={updS("customRate")}
                  note={spotForm.isComplimentary ? "Complimentary / bonus rows are included at ₦0.00 and do not add to payable cost." : (spotForm.rateId && !spotForm.customRate ? `From rate card: ${fmtN(vendorRates.find(r => r.id === spotForm.rateId)?.ratePerSpot)}` : "Enter rate manually or select from rate card above")}
                  placeholder="0" />
                <div style={{ display: "flex", alignItems: "center", gap: 10, marginTop: 10, padding: "10px 12px", background: "var(--bg3)", border: "1px solid var(--border2)", borderRadius: 8 }}>
                  <input
                    id="single-row-complimentary"
                    type="checkbox"
                    checked={!!spotForm.isComplimentary}
                    onChange={e => setSpotForm(f => {
                      const nextChecked = e.target.checked;
                      const currentSpotCount = (f.calendarDays || []).length || parseFloat(f.spots) || 0;
                      return {
                        ...f,
                        isComplimentary: nextChecked,
                        customRate: nextChecked ? "" : f.customRate,
                        bonusSpots: nextChecked ? String(currentSpotCount || "") : f.bonusSpots,
                      };
                    })}
                    style={{ accentColor: "var(--accent)", width: 16, height: 16 }}
                  />
                  <label htmlFor="single-row-complimentary" style={{ fontSize: 13, color: "var(--text2)", cursor: "pointer" }}>
                    Mark this row as complimentary / bonus spots
                  </label>
                </div>
              </div>
              <div style={{ gridColumn: "1/-1" }}>
                <Field
                  label="Bonus Spots"
                  type="number"
                  value={spotForm.isComplimentary ? (((spotForm.calendarDays||[]).length || spotForm.spots || "") ) : (spotForm.bonusSpots || "")}
                  onChange={updS("bonusSpots")}
                  note={spotForm.isComplimentary ? "Full row is complimentary, so all scheduled spots are treated as bonus spots." : "Bonus spots are deducted from the total scheduled spots on this row for costing."}
                  placeholder="0"
                />
              </div>
            </div>
          </div>

          {/* Calendar day picker */}
          <div style={{ marginTop: 16, borderTop: "1px solid var(--border)", paddingTop: 16 }}>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 10 }}>
              <div style={{ fontSize: 12, fontWeight: 600, color: "var(--text2)", textTransform: "uppercase", letterSpacing: ".06em" }}>
                📅 Select Airing Dates — <span style={{ color: "var(--accent)" }}>{(spotForm.calendarDays||[]).length} day{(spotForm.calendarDays||[]).length !== 1 ? "s" : ""} selected</span>
              </div>
              <div style={{ display: "flex", gap: 6 }}>
                <button onClick={() => { updS("calendarDays")(Array.from({length:31},(_,i)=>i+1)); updS("spots")("31"); }} style={{ padding: "3px 10px", fontSize: 11, background: "var(--bg3)", border: "1px solid var(--border2)", borderRadius: 6, color: "var(--text2)", cursor: "pointer" }}>All</button>
                <button onClick={() => { updS("calendarDays")([]); updS("spots")(""); }} style={{ padding: "3px 10px", fontSize: 11, background: "var(--bg3)", border: "1px solid var(--border2)", borderRadius: 6, color: "var(--text2)", cursor: "pointer" }}>Clear</button>
              </div>
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(11, 1fr)", gap: 4 }}>
              {Array.from({length: 31}, (_, i) => i + 1).map(d => {
                const sel = (spotForm.calendarDays || []).includes(d);
                return (
                  <button key={d} onClick={() => {
                    const cur = spotForm.calendarDays || [];
                    const next = sel ? cur.filter(x => x !== d) : [...cur, d].sort((a,b)=>a-b);
                    updS("calendarDays")(next);
                    updS("spots")(String(next.length));
                    setSpotForm(f => ({
                      ...f,
                      calendarDays: next,
                      spots: String(next.length),
                      bonusSpots: f.isComplimentary ? String(next.length || "") : (next.length > 0 ? String(Math.max(0, Math.min(Number(f.bonusSpots) || 0, next.length))) : ""),
                    }));
                  }}
                  style={{ padding: "6px 2px", border: `1px solid ${sel ? "var(--accent)" : "var(--border2)"}`, borderRadius: 6, background: sel ? "rgba(240,165,0,.18)" : "var(--bg3)", color: sel ? "var(--accent)" : "var(--text3)", fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 12, cursor: "pointer", transition: "all .1s" }}>
                    {d}
                  </button>
                );
              })}
            </div>
          </div>

          <div style={{ marginTop: 14, display: "flex", alignItems: "center", justifyContent: "space-between", background: "var(--bg3)", borderRadius: 9, padding: "10px 14px" }}>
            <span style={{ fontSize: 13, color: "var(--text2)" }}>Number of Spots (auto-counted from dates)</span>
            <span style={{ fontFamily: "'Syne',sans-serif", fontWeight: 800, fontSize: 18, color: "var(--accent)" }}>{(spotForm.calendarDays||[]).length || spotForm.spots || 0}</span>
          </div>

          <div style={{ display: "flex", gap: 10, justifyContent: "flex-end", marginTop: 16 }}>
            <Btn variant="ghost" onClick={() => { setSpotModal(null); setEditSpotId(null); setSpotForm(blankSpot); }}>Cancel</Btn>
            <Btn onClick={addSpot}>{editSpotId ? "Save Changes" : "Add Row"}</Btn>
          </div>
        </Modal>
      )}

      <div
        style={{
          position: "sticky",
          top: 14,
          zIndex: 25,
          marginBottom: 24,
          paddingBottom: 2,
          background: "color-mix(in srgb, var(--bg) 82%, transparent)",
          backdropFilter: "blur(10px)",
        }}
      >
        <div style={{ display: "flex", alignItems: "center", gap: 14, marginBottom: 14 }}>
        <Btn variant="ghost" size="sm" onClick={() => { if (window.confirm("Leave this MPO form? Your latest draft has been autosaved and can be restored later.")) setView("list"); }}>← All MPOs</Btn>
        <h1 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 800, fontSize: 22 }}>{editId ? "Edit MPO" : "New MPO"}</h1>
        {editId && <Badge color="accent">Editing</Badge>}
        </div>

      {/* Step bar */}
      <div
        style={{
          display: "flex",
          gap: 0,
          background: "color-mix(in srgb, var(--bg2) 92%, transparent)",
          border: "1px solid var(--border)",
          borderRadius: 11,
          overflow: "hidden",
          boxShadow: "0 10px 24px rgba(0,0,0,.18)",
        }}
      >
        {stepLabels.map((s, i) => (
          <button key={i} onClick={() => setStep(i + 1)}
            style={{ flex: 1, padding: "11px 6px", border: "none", background: step === i + 1 ? "var(--accent)" : "transparent", color: step === i + 1 ? "#000" : step > i + 1 ? "var(--green)" : "var(--text2)", fontFamily: "'Syne',sans-serif", fontWeight: 600, fontSize: 12, cursor: "pointer", borderRight: i < 3 ? "1px solid var(--border)" : "none", transition: "all .18s" }}>
            <span style={{ display: "block", fontSize: 9, opacity: .75, marginBottom: 1 }}>STEP {i + 1}</span>{s}
          </button>
        ))}
      </div>
      </div>

      {/* Step 1 */}
      {step === 1 && (
        <div className="fade">
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 18, marginBottom: 18 }}>
            <Card>
              <h3 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, marginBottom: 15, fontSize: 14 }}>Campaign Details</h3>
              <div style={{ display: "flex", flexDirection: "column", gap: 13 }}>
                <Field label="Campaign" value={mpoData.campaignId} onChange={upd("campaignId")} options={campaigns.map(c => ({ value: c.id, label: c.name }))} placeholder="Select campaign" />
                {campaign && <div style={{ background: "var(--bg3)", borderRadius: 8, padding: 11, fontSize: 12 }}><div style={{ color: "var(--text2)" }}>Client: <strong style={{ color: "var(--text)" }}>{client?.name}</strong></div><div style={{ color: "var(--text2)", marginTop: 3 }}>Brand: <strong style={{ color: "var(--text)" }}>{campaign.brand || "—"}</strong></div>{campaign.materialList && campaign.materialList.length > 0 && <div style={{ color: "var(--text2)", marginTop: 3 }}>🎬 <strong style={{ color: "var(--teal)" }}>{campaign.materialList.length} material{campaign.materialList.length!==1?"s":""} available</strong></div>}</div>}
                <div>
                  <div style={{ fontSize: 11, fontWeight: 600, color: "var(--text3)", marginBottom: 5, textTransform: "uppercase", letterSpacing: ".06em" }}>MPO Number (auto-generated)</div>
                  <div style={{ background: "var(--bg3)", border: "1px solid var(--border2)", borderRadius: 9, padding: "10px 13px", fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 15, color: "var(--accent)", letterSpacing: ".04em" }}>{mpoData.mpoNo || "—"}</div>
                  <div style={{ fontSize: 10, color: "var(--text3)", marginTop: 4 }}>Format: Brand prefix + sequence + Month/Year</div>
                </div>
                <Field label="MPO Date" type="date" value={mpoData.date} onChange={upd("date")} />
                <Field label="Status" value={mpoData.status} onChange={upd("status")} options={[{value:"draft",label:"Draft"},{value:"sent",label:"Sent"},{value:"approved",label:"Approved"},{value:"rejected",label:"Rejected"}]} />
              </div>
            </Card>
            <Card>
              <h3 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, marginBottom: 15, fontSize: 14 }}>Vendor & Period</h3>
              <div style={{ display: "flex", flexDirection: "column", gap: 13 }}>
                <div style={{ display: "grid", gridTemplateColumns: "1fr auto", gap: 10, alignItems: "end" }}>
                  <Field label="Media House / Vendor" value={mpoData.vendorId} onChange={upd("vendorId")} options={vendors.map(v => ({ value: v.id, label: v.name }))} placeholder="Select vendor" />
                  <Btn variant="ghost" onClick={() => setVendorPickerOpen(true)} icon="🔎">Search Vendor</Btn>
                </div>
                <div style={{ marginTop: -3, fontSize: 11, color: "var(--text3)" }}>
                  Use the search button to quickly find a media house or vendor by name, type, phone, or email.
                </div>
                {vendor && <div style={{ background: "var(--bg3)", borderRadius: 8, padding: 11, fontSize: 12 }}><Badge color="blue">{vendor.type}</Badge><div style={{ marginTop: 6, color: "var(--text2)" }}>Vol Disc: <strong style={{ color: "var(--accent)" }}>{vendor.discount || 0}%</strong> · Comm: <strong style={{ color: "var(--accent)" }}>{vendor.commission || 0}%</strong></div></div>}
                <div style={{ display: "grid", gridTemplateColumns: "3fr 1fr", gap: 11, alignItems: "start" }}>
                  <div>
                    <div style={{ fontSize: 11, fontWeight: 600, color: "var(--text3)", textTransform: "uppercase", letterSpacing: ".08em", marginBottom: 6 }}>Campaign Month(s)</div>
                    <div style={{ display: "flex", flexWrap: "wrap", gap: 5 }}>
                      {months.map(m => {
                        const key = `${m} ${mpoData.year || new Date().getFullYear()}`;
                        const selected = (mpoData.months || []).includes(m) || mpoData.month === m;
                        return (
                          <button key={m} onClick={() => {
                            const cur = mpoData.months?.length ? [...mpoData.months] : (mpoData.month ? [mpoData.month] : []);
                            const next = cur.includes(m) ? cur.filter(x => x !== m) : [...cur, m];
                            const sorted = months.filter(x => next.includes(x));
                            setMpoData(d => ({ ...d, months: sorted, month: sorted[0] || "" }));
                          }}
                          style={{ padding: "4px 11px", borderRadius: 20, fontSize: 11, fontWeight: 700, cursor: "pointer", fontFamily: "'Syne',sans-serif", border: selected ? "2px solid var(--accent)" : "1px solid var(--border2)", background: selected ? "var(--accent)" : "var(--bg3)", color: selected ? "#000" : "var(--text2)", transition: "all .12s" }}>
                            {m.slice(0,3)}
                          </button>
                        );
                      })}
                    </div>
                    {(mpoData.months || []).length > 0 && (
                      <div style={{ marginTop: 6, fontSize: 10, color: "var(--green)" }}>
                        ✓ {(mpoData.months||[]).join(", ")}
                      </div>
                    )}
                  </div>
                  <Field label="Year" value={mpoData.year} onChange={upd("year")} placeholder="2025" />
                </div>
                <Field label="Medium" value={mpoData.medium} onChange={upd("medium")} options={["Television","Radio","Print","Digital","OOH","Multi-Platform"]} />
                <Field label="Transmit Instruction" value={mpoData.transmitMsg} onChange={upd("transmitMsg")} placeholder={`PLEASE TRANSMIT SPOTS ON ${vendor?.name || "VENDOR"} AS SCHEDULED`} rows={2} />
              </div>
            </Card>
          </div>
          <div style={{ display: "flex", justifyContent: "flex-end" }}><Btn onClick={() => setStep(2)}>Next: Spot Schedule →</Btn></div>
        </div>
      )}

      {/* Step 2 */}
      {step === 2 && (
        <div className="fade">
          <Card style={{ marginBottom: 18 }}>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 15, flexWrap: "wrap", gap: 10 }}>
              <div>
                <h3 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 14 }}>Spot Schedule</h3>
                {vendor && vendorRates.length > 0 && (
                  <div style={{ fontSize: 11, color: "var(--green)", marginTop: 3 }}>
                    ✓ {vendorRates.length} rate card{vendorRates.length !== 1 ? "s" : ""} from <strong>{vendor.name}</strong> — select when adding spots
                  </div>
                )}
                {vendor && vendorRates.length === 0 && (
                  <div style={{ fontSize: 11, color: "var(--text3)", marginTop: 3 }}>
                    ⚠️ No rate cards for <strong>{vendor.name}</strong> — add them in Media Rates
                  </div>
                )}
              </div>
              <div style={{ display: "flex", gap: 8 }}>
                <Btn variant={dailyMode ? "primary" : "ghost"} size="sm" onClick={() => { if (!dailyMode && spots.length && !calendarEditHasRows) { const rebuilt = buildCalDataFromSpots(spots); calMonthFallbackRowsRef.current = rebuilt; setCalData(rebuilt); setActiveCalMonth(Object.keys(rebuilt || {})[0] || ""); } setDailyMode(true); }}>📅 Daily Schedule</Btn>
                <Btn variant={!dailyMode ? "secondary" : "ghost"} size="sm" onClick={() => { if (dailyMode && calendarEditHasRows) setSpots(buildSpotsFromCalData(calData)); setDailyMode(false); }}>➕ Single Row</Btn>
              </div>
            </div>

            {/* ── Daily Calendar Schedule — Multi-Month ── */}
            {dailyMode && (() => {
              const selectedMonths = mpoData.months?.length ? mpoData.months : (mpoData.month ? [mpoData.month] : []);
              if (!selectedMonths.length) return (
                <div style={{ background: "rgba(240,165,0,.1)", border: "1px solid rgba(240,165,0,.3)", borderRadius: 8, padding: "13px 16px", fontSize: 12, color: "var(--accent)", marginBottom: 12 }}>
                  ⚠️ Please select at least one <strong>Campaign Month</strong> in Step 1 before using the Daily Calendar.
                </div>
              );
              const curCalMonth = activeCalMonth || selectedMonths[0];
              return (
                <div>
                  {selectedMonths.length > 1 && (
                    <div style={{ display: "flex", gap: 0, marginBottom: 16, background: "var(--bg3)", border: "1px solid var(--border2)", borderRadius: 10, overflow: "hidden" }}>
                      {selectedMonths.map(m => (
                        <button key={m} onClick={() => setActiveCalMonth(m)}
                          style={{ flex: 1, padding: "9px 8px", border: "none", background: curCalMonth === m ? "var(--accent)" : "transparent", color: curCalMonth === m ? "#000" : "var(--text2)", fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 12, cursor: "pointer", borderRight: "1px solid var(--border)", transition: "all .15s" }}>
                          {m.slice(0,3)} <span style={{ fontSize: 10, opacity: .7 }}>{mpoData.year?.slice(-2)}</span>
                          {(calData[m] || []).some(r => totalCountFromDayCounts(r.dayCounts || {}) > 0) && (
                            <span style={{ marginLeft: 4, background: "rgba(34,197,94,.25)", color: "var(--green)", borderRadius: 10, padding: "1px 5px", fontSize: 9 }}>✓</span>
                          )}
                        </button>
                      ))}
                    </div>
                  )}
                  <InlineDailyCalendar
                    month={curCalMonth} year={mpoData.year}
                    calRows={getStableCalRows(curCalMonth)}
                    setCalRows={fn => setStableCalRows(curCalMonth, fn)}
                    vendorRates={vendorRates} fmtN={fmtN}
                    blankCalRow={blankCalRow}
                    campaignMaterials={campaign?.materialList || []}
                    onAdd={() => {
                      const rows = getStableCalRows(curCalMonth);
                      addFromCalendarSchedule(rows, `${curCalMonth} ${mpoData.year}`);
                      const nextCalData = { ...calData, [curCalMonth]: rows };
                      setCalData(nextCalData);
                      calMonthFallbackRowsRef.current = nextCalData;
                      setSpots(buildSpotsFromCalData(nextCalData));
                    }}
                  />
                  {(getStableCalRows(curCalMonth) || []).some(row => row?.programme) && (
                    <div style={{ marginTop: 12, border: "1px solid var(--border)", borderRadius: 10, overflow: "hidden" }}>
                      <div style={{ padding: "10px 14px", background: "var(--bg3)", borderBottom: "1px solid var(--border)", fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 13 }}>
                        Bonus Spots Editor — {curCalMonth}
                      </div>
                      <div style={{ display: "flex", flexDirection: "column" }}>
                        {(getStableCalRows(curCalMonth) || []).filter(row => row?.programme).map((row, index) => {
                          const totalScheduledSpots = totalCountFromDayCounts(row?.dayCounts || {});
                          const bonusValue = row?.isComplimentary ? totalScheduledSpots : Math.max(0, Math.min(Number(row?.bonusSpots) || 0, totalScheduledSpots));
                          const paidValue = Math.max(0, totalScheduledSpots - bonusValue);
                          return (
                            <div key={row.id || index} style={{ display: "grid", gridTemplateColumns: "1.5fr .7fr .7fr .9fr", gap: 10, alignItems: "center", padding: "10px 14px", borderTop: index === 0 ? "none" : "1px solid var(--border)" }}>
                              <div>
                                <div style={{ fontWeight: 700, fontSize: 13 }}>{row.programme || "Untitled Programme"}</div>
                                <div style={{ fontSize: 11, color: "var(--text3)", marginTop: 3 }}>{row.timeBelt || "No time belt"} · {totalScheduledSpots} scheduled spot{totalScheduledSpots !== 1 ? "s" : ""}</div>
                              </div>
                              <div>
                                <div style={{ fontSize: 10, fontWeight: 700, color: "var(--text3)", textTransform: "uppercase", marginBottom: 4 }}>Bonus Spots</div>
                                <input
                                  type="number"
                                  min="0"
                                  max={totalScheduledSpots}
                                  value={row.isComplimentary ? String(totalScheduledSpots) : (row.bonusSpots || "")}
                                  onChange={(e) => updateCalendarRowField(curCalMonth, row.id, { bonusSpots: e.target.value })}
                                  disabled={!!row.isComplimentary}
                                  style={{ width: "100%", background: "var(--bg2)", border: "1px solid var(--border2)", borderRadius: 8, padding: "8px 10px", color: "var(--text)", fontSize: 13, outline: "none", opacity: row.isComplimentary ? 0.65 : 1 }}
                                />
                              </div>
                              <div>
                                <div style={{ fontSize: 10, fontWeight: 700, color: "var(--text3)", textTransform: "uppercase", marginBottom: 4 }}>Paid Spots</div>
                                <div style={{ padding: "8px 10px", borderRadius: 8, background: "var(--bg3)", border: "1px solid var(--border2)", fontFamily: "'Syne',sans-serif", fontWeight: 800, fontSize: 16, color: "var(--green)", textAlign: "center" }}>{paidValue}</div>
                              </div>
                              <div style={{ display: "flex", alignItems: "center", gap: 8, paddingTop: 16 }}>
                                <input
                                  id={`bonus-complimentary-${row.id}`}
                                  type="checkbox"
                                  checked={!!row.isComplimentary}
                                  onChange={(e) => updateCalendarRowField(curCalMonth, row.id, (current, total) => ({
                                    isComplimentary: e.target.checked,
                                    customRate: e.target.checked ? "" : current.customRate,
                                    bonusSpots: e.target.checked ? String(total || "") : current.bonusSpots,
                                  }))}
                                  style={{ accentColor: "var(--accent)", width: 16, height: 16 }}
                                />
                                <label htmlFor={`bonus-complimentary-${row.id}`} style={{ fontSize: 11, color: "var(--text2)", cursor: "pointer", lineHeight: 1.35 }}>
                                  Complimentary row
                                </label>
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    </div>
                  )}
                  {buildSpotsFromCalData(calData).length > 0 && (
                    <div style={{ marginTop: 14, border: "1px solid var(--border)", borderRadius: 10, overflow: "hidden" }}>
                      <div style={{ padding: "10px 14px", background: "var(--bg3)", borderBottom: "1px solid var(--border)", fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 13 }}>
                        Saved Schedule Preview
                      </div>
                      <div style={{ overflowX: "auto" }}>
                        <table style={{ width: "100%", borderCollapse: "collapse" }}>
                          <thead>
                            <tr style={{ background: "var(--bg3)" }}>
                              {["Month","Programme","WD","Time Belt","Material","Scheduled","Bonus","Paid"].map(h => (
                                <th key={h} style={{ padding: "6px 8px", textAlign: "left", fontSize: 8, fontWeight: 600, textTransform: "uppercase", color: "var(--text3)", borderBottom: "1px solid var(--border)", whiteSpace: "nowrap" }}>{h}</th>
                              ))}
                            </tr>
                          </thead>
                          <tbody>
                            {buildSpotsFromCalData(calData).map((row, idx) => (
                              <tr key={`${row.id}-${idx}`} style={{ borderBottom: "1px solid var(--border)" }}>
                                <td style={{ padding: "6px 8px", fontSize: 12 }}>{row.scheduleMonth || "—"}</td>
                                <td style={{ padding: "6px 8px", fontWeight: 600, fontSize: 12 }}>{row.programme}</td>
                                <td style={{ padding: "6px 8px", fontSize: 12 }}>{row.wd || "—"}</td>
                                <td style={{ padding: "6px 8px", fontSize: 12 }}>{row.timeBelt || "—"}</td>
                                <td style={{ padding: "6px 8px", fontSize: 12 }}>{row.material || "—"}</td>
                                <td style={{ padding: "6px 8px", fontWeight: 700, fontSize: 12 }}>{row.spots}</td>
                                <td style={{ padding: "6px 8px", color: "var(--purple)", fontWeight: 700, fontSize: 12 }}>{row.bonusSpots || 0}</td>
                                <td style={{ padding: "6px 8px", color: "var(--green)", fontWeight: 700, fontSize: 12 }}>{row.paidSpots ?? getPaidSpotCount(row)}</td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  )}
                  <div style={{ marginTop: 12, padding: "10px 14px", background: "rgba(59,126,245,.08)", borderRadius: 10, border: "1px solid rgba(59,126,245,.18)", fontSize: 12, color: "var(--text2)" }}>
                    <strong style={{ color: "var(--text)" }}>Edit mode note:</strong> update the schedule directly inside the monthly calendar above. Changes made there are what will be saved for this draft MPO, including bonus-spot changes.
                  </div>
                  {selectedMonths.length > 1 && (
                    <div style={{ marginTop: 10, padding: "10px 14px", background: "var(--bg3)", borderRadius: 9, border: "1px solid var(--border)", fontSize: 11, color: "var(--text2)" }}>
                      💡 Switch months using the tabs above to schedule spots across multiple months. Each month is saved independently.
                    </div>
                  )}
                </div>
              );
            })()}

            {/* Single row modal trigger */}
            {!dailyMode && (
              <div style={{ marginBottom: 14, display: "flex", justifyContent: "flex-end" }}>
                {canManage && <Btn size="sm" icon="+" onClick={() => { setSpotForm(blankSpot); setEditSpotId(null); setSpotModal(true); }}>Add Single Spot Row</Btn>}
              </div>
            )}

            {/* Spots table */}
            {!dailyMode && (spots.length === 0 ? <Empty icon="📋" title="No spots added" sub="Use Daily Schedule or Single Row to add spots" /> :
              <div style={{ overflowX: "auto" }}>
                <table style={{ width: "100%", borderCollapse: "collapse" }}>
                  <thead><tr style={{ background: "var(--bg3)" }}>{["Month","Programme","WD","Time Belt","Material","Dur","Rate/Spot","Spots","Gross",""].map(h => <th key={h} style={{ padding: "6px 8px", textAlign: "left", fontSize: 8, fontWeight: 600, textTransform: "uppercase", color: "var(--text3)", borderBottom: "1px solid var(--border)", whiteSpace: "nowrap" }}>{h}</th>)}</tr></thead>
                  <tbody>
                    {spots.map((r, i) => (
                      <tr key={r.id} style={{ borderBottom: "1px solid var(--border)", background: i % 2 === 0 ? "transparent" : "rgba(255,255,255,.01)" }}>
                        <td style={{ padding: "7px 9px" }}><span style={{ fontSize: 9, background: "rgba(59,126,245,.12)", color: "var(--blue)", padding: "2px 6px", borderRadius: 8, fontWeight: 700, whiteSpace: "nowrap" }}>{r.scheduleMonth || mpoData.month || "—"}</span></td>
                        <td style={{ padding: "7px 9px", fontWeight: 600, fontSize: 12 }}>{r.programme}</td>
                        <td style={{ padding: "7px 9px", fontSize: 12 }}><Badge color="blue">{r.wd}</Badge></td>
                        <td style={{ padding: "7px 9px", color: "var(--text2)", fontSize: 12 }}>{r.timeBelt}</td>
                        <td style={{ padding: "7px 9px", color: "var(--text2)", maxWidth: 120, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", fontSize: 11 }}>{r.material}</td>
                        <td style={{ padding: "7px 9px", fontSize: 12 }}>{r.duration}"</td>
                        <td style={{ padding: "7px 9px", color: "var(--accent)", fontSize: 12 }}>{fmtN(r.ratePerSpot)}</td>
                        <td style={{ padding: "7px 9px", fontWeight: 700, fontSize: 12 }}>{r.spots}</td>
                        <td style={{ padding: "7px 9px", color: "var(--green)", fontWeight: 600, fontSize: 12 }}>{fmtN((parseFloat(r.spots) || 0) * (parseFloat(r.ratePerSpot) || 0))}</td>
                        <td style={{ padding: "7px 9px" }}>
                          <div style={{ display: "flex", gap: 5 }}>
                            {canManage && <Btn variant="ghost" size="sm" onClick={() => {
                              setSpotForm({
                                programme: r.programme,
                                wd: r.wd,
                                timeBelt: r.timeBelt,
                                material: r.material,
                                materialCustom: "",
                                duration: r.duration,
                                rateId: r.rateId || "",
                                customRate: r.isComplimentary ? "" : (r.customRate || r.ratePerSpot || ""),
                                spots: r.spots,
                                bonusSpots: String((r.bonusSpots ?? getSpotBonusCount(r)) || ""),
                                isComplimentary: !!r.isComplimentary,
                                calendarDays: r.calendarDays || [],
                                calendarDayCounts: collapseDaysToCounts(r.calendarDays || []),
                              });
                              setEditSpotId(r.id);
                              setSpotModal(true);
                            }}>✏️</Btn>}
                            <Btn variant="danger" size="sm" onClick={() => setSpots(s => s.filter(x => x.id !== r.id))}>×</Btn>
                          </div>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                  <tfoot><tr style={{ background: "var(--bg3)", borderTop: "2px solid var(--border2)" }}><td colSpan={7} style={{ padding: "7px 9px", fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 12 }}>TOTAL</td><td style={{ padding: "7px 9px", fontWeight: 700, color: "var(--blue)", fontSize: 12 }}>{totalSpots}</td><td style={{ padding: "7px 9px", fontWeight: 700, color: "var(--green)", fontSize: 12 }}>{fmtN(totalGross)}</td><td></td></tr></tfoot>
                </table>
              </div>)}
          </Card>
          <div style={{ display: "flex", justifyContent: "space-between" }}>
            <Btn variant="ghost" onClick={() => setStep(1)}>← Back</Btn>
            <Btn onClick={() => setStep(3)}>Next: Costing →</Btn>
          </div>
        </div>
      )}

      {/* Step 3 */}
      {step === 3 && (
        <div className="fade">
          <div style={{ display: "grid", gridTemplateColumns: "1fr 340px", gap: 18, marginBottom: 18 }}>
            <Card>
              <h3 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, marginBottom: 15, fontSize: 14 }}>Costing Summary</h3>
              <div style={{ background: "rgba(240,165,0,.07)", border: "1px solid rgba(240,165,0,.18)", borderRadius: 8, padding: "8px 13px", fontSize: 12, color: "var(--text2)", marginBottom: 16 }}>
                💡 VAT is applied automatically using your current document settings.
              </div>
              {/* Surcharge input */}
              <div style={{ marginBottom: 14, background: "var(--bg3)", border: "1px solid var(--border2)", borderRadius: 10, padding: "12px 15px" }}>
                <div style={{ fontSize: 11, fontWeight: 700, color: "var(--text3)", textTransform: "uppercase", letterSpacing: ".06em", marginBottom: 10 }}>Surcharge (Optional)</div>
                <div style={{ display: "grid", gridTemplateColumns: "1fr 2fr", gap: 10 }}>
                  <div>
                    <label style={{ fontSize: 11, color: "var(--text3)", fontWeight: 600, display: "block", marginBottom: 4 }}>Surcharge %</label>
                    <input type="number" value={surcharge.pct} onChange={e => setSurcharge(s => ({ ...s, pct: e.target.value }))}
                      placeholder="e.g. 10"
                      style={{ width: "100%", background: "var(--bg2)", border: "1px solid var(--border2)", borderRadius: 8, padding: "8px 12px", color: "var(--text)", fontSize: 13, outline: "none" }} />
                  </div>
                  <div>
                    <label style={{ fontSize: 11, color: "var(--text3)", fontWeight: 600, display: "block", marginBottom: 4 }}>Surcharge Label</label>
                    <input value={surcharge.label} onChange={e => setSurcharge(s => ({ ...s, label: e.target.value }))}
                      placeholder="e.g. Production Surcharge, Agency Fee"
                      style={{ width: "100%", background: "var(--bg2)", border: "1px solid var(--border2)", borderRadius: 8, padding: "8px 12px", color: "var(--text)", fontSize: 13, outline: "none" }} />
                  </div>
                </div>
                {surchPct > 0 && <div style={{ marginTop: 8, fontSize: 11, color: "var(--orange)" }}>⚡ Surcharge adds {fmtN(surchAmt)} to the net cost</div>}
              </div>

              <div style={{ marginTop: 0, border: "1px solid var(--border)", borderRadius: 11, overflow: "hidden" }}>
                {[
                  ["Total Paid Spots", totalSpots, "var(--text)", false],
                  ["Total Gross Value", fmtN(totalGross), "var(--accent)", true],
                  ["Volume Discount", `(${fmtN(discAmt)})`, "var(--red)", false, "discount"],
                  ["Less Discount", fmtN(lessDisc), "var(--text)", false, "text"],
                  ...(commPct > 0 ? [[`Agency Commission (${(commPct * 100).toFixed(0)}%)`, `(${fmtN(commAmt)})`, "var(--red)", false], ["After Commission", fmtN(afterComm), "var(--text)", false]] : []),
                  ...(surchPct > 0 ? [[surcharge.label || `Surcharge (${(surchPct * 100).toFixed(0)}%)`, `+${fmtN(surchAmt)}`, "var(--orange)", false]] : []),
                  ["Net Value", fmtN(netVal), "var(--green)", true],
                  [`VAT (${VAT_RATE}%)`, fmtN(vatAmt), "var(--text)", false]
                ].map(([l, v, c, bold, rowType], i) => (
                  <div key={`${l}-${i}`} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "11px 15px", background: i % 2 === 0 ? "var(--bg3)" : "transparent", borderBottom: "1px solid var(--border)", gap: 12 }}>
                    {rowType === "discount" ? (
                      <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
                        <span style={{ fontSize: 13, color: "var(--text2)" }}>Volume Discount</span>
                        <input
                          type="number"
                          min="0"
                          step="0.01"
                          value={mpoData.discountOverridePct ?? ""}
                          onChange={e => setMpoData(m => ({ ...m, discountOverridePct: e.target.value }))}
                          placeholder={String((vendorDiscPct * 100).toFixed(0))}
                          style={{ width: 92, background: "var(--bg2)", border: "1px solid var(--border2)", borderRadius: 8, padding: "6px 8px", color: "var(--text)", fontSize: 12, outline: "none" }}
                        />
                        <span style={{ fontSize: 12, color: "var(--text3)" }}>%</span>
                        <button
                          type="button"
                          onClick={() => setMpoData(m => ({ ...m, discountOverridePct: vendor ? String(parseFloat(vendor.discount) || 0) : "0" }))}
                          style={{ border: "1px solid var(--border2)", background: "var(--bg2)", borderRadius: 999, padding: "5px 10px", fontSize: 11, fontWeight: 700, color: "var(--text2)", cursor: "pointer" }}
                        >
                          Use Default
                        </button>
                        <span style={{ fontSize: 11, color: "var(--text3)" }}>Applied: {(discPct * 100).toFixed(2).replace(/\.00$/, "")}%</span>
                      </div>
                    ) : (
                      <span style={{ fontSize: 13, color: "var(--text2)" }}>{l}</span>
                    )}
                    <span style={{ fontFamily: "'Syne',sans-serif", fontWeight: bold ? 800 : 600, fontSize: bold ? 17 : 13, color: c }}>{v}</span>
                  </div>
                ))}
                <div style={{ padding: "14px 15px", background: "var(--accent)", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                  <span style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, color: "#000", fontSize: 13 }}>TOTAL AMOUNT PAYABLE</span>
                  <span style={{ fontFamily: "'Syne',sans-serif", fontWeight: 800, fontSize: 20, color: "#000" }}>{fmtN(grandTotal)}</span>
                </div>
              </div>
            </Card>
            <Card>
              <h3 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, marginBottom: 15, fontSize: 14 }}>MPO Summary</h3>
              {[["Client", client?.name || "—"], ["Brand", campaign?.brand || "—"], ["Campaign", campaign?.name || "—"], ["Vendor", vendor?.name || "—"], ["MPO No.", mpoData.mpoNo || "—"], ["Period", `${(mpoData.months||[]).length > 1 ? (mpoData.months||[]).join(", ") : (mpoData.month || "—")} ${mpoData.year}`], ["Total Spots", totalSpots], ["Gross", fmtN(totalGross)], ["Net Payable", fmtN(grandTotal)]].map(([l, v]) => (
                <div key={l} style={{ display: "flex", justifyContent: "space-between", padding: "7px 0", borderBottom: "1px solid var(--border)", fontSize: 12 }}>
                  <span style={{ color: "var(--text3)", fontWeight: 600 }}>{l}</span>
                  <span style={{ fontWeight: 600 }}>{v}</span>
                </div>
              ))}
            </Card>
          </div>
          <div style={{ display: "flex", justifyContent: "space-between" }}>
            <Btn variant="ghost" onClick={() => setStep(2)}>← Back</Btn>
            <Btn onClick={() => setStep(4)}>Next: Signatories →</Btn>
          </div>
        </div>
      )}

      {/* Step 4 */}
      {step === 4 && (
        <div className="fade">
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 18, marginBottom: 18 }}>
            <Card>
              <h3 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, marginBottom: 15, fontSize: 14 }}>Signatories</h3>
              <div style={{ display: "flex", flexDirection: "column", gap: 13 }}>
                <Field label="Signed By (Name)" value={mpoData.signedBy} onChange={upd("signedBy")} placeholder="Nkechi Eluma" />
                <Field label="Signed By (Title)" value={mpoData.signedTitle} onChange={upd("signedTitle")} placeholder="Head Buying and Compliance" />
                {user?.signatureDataUrl && (
                  <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
                    <Btn variant="ghost" size="sm" onClick={() => setMpoData(m => ({ ...m, signedSignature: user.signatureDataUrl }))}>Use My Uploaded Signature</Btn>
                    {mpoData.signedSignature ? <Btn variant="danger" size="sm" onClick={() => setMpoData(m => ({ ...m, signedSignature: "" }))}>Clear Signature</Btn> : null}
                  </div>
                )}
                <div style={{ background: "var(--bg3)", border: "1px solid var(--border2)", borderRadius: 10, padding: "13px 15px" }}>
                  <div style={{ fontSize: 10, fontWeight: 600, color: "var(--text3)", textTransform: "uppercase", letterSpacing: ".06em", marginBottom: 10 }}>Prepared By (from your account)</div>
                  <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                    <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13 }}><span style={{ color: "var(--text3)" }}>Name</span><span style={{ fontWeight: 600 }}>{user?.name || "—"}</span></div>
                    <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13 }}><span style={{ color: "var(--text3)" }}>Title</span><span style={{ fontWeight: 600 }}>{user?.title || "—"}</span></div>
                    <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13 }}><span style={{ color: "var(--text3)" }}>Contact</span><span style={{ fontWeight: 600 }}>{user?.email || "—"}</span></div>
                  </div>
                  {user?.signatureDataUrl && (
                    <div style={{ marginTop: 10 }}>
                      <div style={{ fontSize: 11, color: "var(--text3)", marginBottom: 6 }}>Prepared signature</div>
                      <img src={user.signatureDataUrl} alt="Prepared signature" style={{ maxHeight: 52, maxWidth: 220, objectFit: "contain", background: "#fff", borderRadius: 8, border: "1px solid var(--border)" }} />
                    </div>
                  )}
                  <div style={{ marginTop: 9, fontSize: 11, color: "var(--text3)" }}>To update, edit your profile or re-register with updated details.</div>
                </div>
                <Field label="Agency Address" value={mpoData.agencyAddress} onChange={upd("agencyAddress")} placeholder="5, Craig Street, Ogudu GRA, Lagos" />
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 }}>
                  <Field label="Agency Email" type="email" value={mpoData.agencyEmail || ""} onChange={upd("agencyEmail")} placeholder="hello@agency.com" />
                  <Field label="Agency Phone" type="tel" value={mpoData.agencyPhone || ""} onChange={upd("agencyPhone")} placeholder="+234 800 000 0000" />
                </div>
              </div>
            </Card>
            <Card style={{ background: "linear-gradient(135deg,rgba(240,165,0,.07),rgba(59,126,245,.04))", border: "1px solid rgba(240,165,0,.18)" }}>
              <h3 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, marginBottom: 15, fontSize: 14 }}>Ready to Save</h3>
              <div style={{ display: "flex", flexDirection: "column", gap: 8, fontSize: 13 }}>
                <div>📋 <strong>{campaign?.name || "No campaign selected"}</strong></div>
                <div>🏢 <strong>{vendor?.name || "No vendor selected"}</strong></div>
                <div>👥 {client?.name || "—"}{campaign?.brand && ` · ${campaign.brand}`}</div>
                <div>📅 {mpoData.month} {mpoData.year} · MPO: {mpoData.mpoNo || "—"}</div>
                <div>📋 {totalSpots} spots · {spots.length} rows</div>
                <div style={{ marginTop: 10, padding: "13px 15px", background: "rgba(240,165,0,.1)", borderRadius: 10, border: "1px solid rgba(240,165,0,.2)" }}>
                  <div style={{ fontSize: 11, color: "var(--text3)", marginBottom: 3 }}>TOTAL AMOUNT PAYABLE</div>
                  <div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 800, fontSize: 22, color: "var(--accent)" }}>{fmtN(grandTotal)}</div>
                </div>
                <div style={{ marginTop: 6, fontSize: 11, color: "var(--text2)", lineHeight: 1.5 }}>After saving, you can download this MPO as PDF, Word document, or Excel/CSV from the MPO list.</div>
              </div>
            </Card>
          </div>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
            <Btn variant="ghost" onClick={() => setStep(3)}>← Back</Btn>
            <div style={{ display: "flex", gap: 10 }}>
              <Btn variant="secondary" onClick={() => { if (window.confirm("Leave this MPO form? Your latest draft has been autosaved and can be restored later.")) setView("list"); }}>Cancel</Btn>
              {canManage && <Btn variant="success" icon="✓" onClick={saveMPO}>{editId ? "Save Changes" : "Save MPO"}</Btn>}
            </div>
          </div>
        </div>
      )}
    </div>
  );

}
