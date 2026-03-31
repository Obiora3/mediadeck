import React from "react";
import Field from "../Field";

export default function DailyCalendar({ month, year, calRows, setCalRows, vendorRates, fmtN, blankCalRow, onAdd, campaignMaterials }) {
  const MONTH_NAMES = ["JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"];
  const DAY_LABELS  = ["Sun","Mon","Tue","Wed","Thu","Fri","Sat"];

  const mIdx      = (() => {
    const m = (month||"").toUpperCase();
    // Try 3-letter abbreviation first (JAN, FEB…), then full name (JANUARY…)
    const short = MONTH_NAMES.indexOf(m.slice(0,3));
    return short >= 0 ? short : MONTH_NAMES.findIndex(n => m.startsWith(n));
  })();
  const yr        = parseInt(year) || new Date().getFullYear();
  const validMonth = mIdx >= 0;
  const dim       = validMonth ? new Date(yr, mIdx + 1, 0).getDate() : 31;
  const firstDOW  = validMonth ? new Date(yr, mIdx, 1).getDay() : 0;

  // Build week grid: array of 7-slot rows, null = empty cell
  const cells = [];
  for (let i = 0; i < firstDOW; i++) cells.push(null);
  for (let d = 1; d <= dim; d++) cells.push(d);
  while (cells.length % 7 !== 0) cells.push(null);
  const weeks = [];
  for (let i = 0; i < cells.length; i += 7) weeks.push(cells.slice(i, i + 7));

  const updRow = (id, key, val) =>
    setCalRows(rows => rows.map(r => r.id === id ? { ...r, [key]: val } : r));

  const toggleDate = (rowId, d) =>
    setCalRows(rows => rows.map(r => {
      if (r.id !== rowId) return r;
      const next = new Set(r.selectedDates);
      next.has(d) ? next.delete(d) : next.add(d);
      return { ...r, selectedDates: next };
    }));

  const quickSelect = (rowId, wdNums) =>
    setCalRows(rows => rows.map(r => {
      if (r.id !== rowId) return r;
      const next = new Set();
      if (wdNums === "all") {
        for (let d = 1; d <= dim; d++) next.add(d);
      } else if (wdNums === "clear") {
        /* empty */
      } else {
        for (let d = 1; d <= dim; d++)
          if (validMonth && wdNums.includes(new Date(yr, mIdx, d).getDay())) next.add(d);
      }
      return { ...r, selectedDates: next };
    }));

  const inp = (extra = {}) => ({
    background: "var(--bg2)", border: "1px solid var(--border2)", borderRadius: 6,
    padding: "5px 8px", color: "var(--text)", fontSize: 11, outline: "none",
    width: "100%", ...extra
  });

  if (!validMonth) return (
    <div style={{ background: "rgba(240,165,0,.1)", border: "1px solid rgba(240,165,0,.3)",
        borderRadius: 8, padding: "14px 16px", fontSize: 12, color: "var(--accent)", marginBottom: 14 }}>
      ⚠️ Please set <strong>Month &amp; Year</strong> in Step 1 before using the Daily Calendar.
    </div>
  );

  return (
    <div className="fade" style={{ marginBottom: 12 }}>
      {/* Month header */}
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 14 }}>
        <div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 800, fontSize: 15, letterSpacing: 1 }}>
          📅 {MONTH_NAMES[mIdx]} {yr}
        </div>
        <div style={{ fontSize: 11, color: "var(--text3)" }}>
          Click dates on the calendar to toggle airing days per spot row
        </div>
      </div>

      {calRows.map((row, ri) => {
        const rateObj = vendorRates.find(r => r.id === row.rateId);
        const rateVal = parseFloat(row.customRate) || parseFloat(rateObj?.ratePerSpot) || 0;
        const spotCount = row.selectedDates.size;
        const gross = rateVal * spotCount;

        return (
          <div key={row.id} style={{ border: "1px solid var(--border2)", borderRadius: 10,
              marginBottom: 16, background: "var(--bg2)", overflow: "hidden" }}>

            {/* Row header bar */}
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between",
                background: "var(--bg3)", padding: "10px 14px", borderBottom: "1px solid var(--border)" }}>
              <span style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 12,
                  color: "var(--text2)", textTransform: "uppercase", letterSpacing: 1 }}>
                Spot Row {ri + 1}
              </span>
              <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                {spotCount > 0 && (
                  <span style={{ fontSize: 11, background: "rgba(240,165,0,.18)", color: "var(--accent)",
                      padding: "3px 10px", borderRadius: 12, fontWeight: 700 }}>
                    {spotCount} spot{spotCount !== 1 ? "s" : ""}{rateVal > 0 ? ` · ₦${fmtN(gross)}` : ""}
                  </span>
                )}
                {calRows.length > 1 && (
                  <button onClick={() => setCalRows(r => r.filter(x => x.id !== row.id))}
                    style={{ background: "none", border: "none", cursor: "pointer",
                        color: "var(--text3)", fontSize: 18, lineHeight: 1, padding: "0 2px" }}>×</button>
                )}
              </div>
            </div>

            <div style={{ padding: "12px 14px" }}>
              {/* Programme selector from vendor rate cards */}
              {vendorRates.length > 0 && (
                <div style={{ marginBottom: 12 }}>
                  <div style={{ fontSize: 10, fontWeight: 700, color: "var(--text3)", textTransform: "uppercase", marginBottom: 5, letterSpacing: .5 }}>
                    Select Programme from Rate Card
                  </div>
                  <select
                    value={row.rateId}
                    onChange={e => {
                      const picked = vendorRates.find(r => r.id === e.target.value);
                      if (picked) {
                        setCalRows(rs => rs.map(r => r.id !== row.id ? r : {
                          ...r,
                          rateId:      picked.id,
                          programme:   picked.programme || r.programme,
                          timeBelt:    picked.timeBelt  || r.timeBelt,
                          duration:    picked.duration  || r.duration,
                          customRate:  "",
                        }));
                      } else {
                        updRow(row.id, "rateId", "");
                      }
                    }}
                    style={{ ...inp(), cursor: "pointer", background: row.rateId ? "rgba(240,165,0,.08)" : "var(--bg2)", border: row.rateId ? "1px solid rgba(240,165,0,.4)" : "1px solid var(--border2)" }}>
                    <option value="">— Select programme / rate card —</option>
                    {vendorRates.map(r => (
                      <option key={r.id} value={r.id}>
                        {r.programme}{r.timeBelt ? ` · ${r.timeBelt}` : ""} · {r.duration}s · ₦{fmtN(r.ratePerSpot)}
                      </option>
                    ))}
                  </select>
                  {row.rateId && (
                    <div style={{ marginTop: 4, display: "flex", alignItems: "center", justifyContent: "space-between" }}>
                      <span style={{ fontSize: 10, color: "var(--green)" }}>✓ Rate card applied — fields auto-filled below</span>
                      <button onClick={() => setCalRows(rs => rs.map(r => r.id !== row.id ? r : { ...r, rateId: "", programme: "", timeBelt: "", duration: "30", customRate: "" }))}
                        style={{ fontSize: 10, color: "var(--text3)", background: "none", border: "none", cursor: "pointer", textDecoration: "underline" }}>Clear</button>
                    </div>
                  )}
                </div>
              )}

              {/* Spot detail fields */}
              <div style={{ display: "grid", gridTemplateColumns: "2fr 1.4fr 1.4fr 64px 1.5fr",
                  gap: 8, marginBottom: 12 }}>
                <div>
                  <div style={{ fontSize: 10, fontWeight: 700, color: "var(--text3)",
                      textTransform: "uppercase", marginBottom: 4, letterSpacing: .5 }}>Programme</div>
                  <input value={row.programme} onChange={e => updRow(row.id, "programme", e.target.value)}
                    placeholder="e.g. NTA Network News" style={inp()} />
                </div>
                <div>
                  <div style={{ fontSize: 10, fontWeight: 700, color: "var(--text3)",
                      textTransform: "uppercase", marginBottom: 4, letterSpacing: .5 }}>Time Belt</div>
                  <input value={row.timeBelt} onChange={e => updRow(row.id, "timeBelt", e.target.value)}
                    placeholder="e.g. 9PM–10PM" style={inp()} />
                </div>
                <div>
                  <div style={{ fontSize: 10, fontWeight: 700, color: "var(--text3)",
                      textTransform: "uppercase", marginBottom: 4, letterSpacing: .5 }}>Material</div>
                  {campaignMaterials && campaignMaterials.length > 0 ? (
                    <>
                      <select value={row.material} onChange={e => updRow(row.id, "material", e.target.value)}
                        style={{ ...inp(), cursor: "pointer" }}>
                        <option value="">— Select —</option>
                        {campaignMaterials.map((m, mi) => <option key={mi} value={m}>{m}</option>)}
                        <option value="__custom__">Custom…</option>
                      </select>
                      {row.material === "__custom__" && (
                        <input value={row.materialCustom || ""} onChange={e => updRow(row.id, "materialCustom", e.target.value)}
                          placeholder="Type material" style={{ ...inp(), marginTop: 4 }} />
                      )}
                    </>
                  ) : (
                    <input value={row.material} onChange={e => updRow(row.id, "material", e.target.value)}
                      placeholder="e.g. 30s TVC" style={inp()} />
                  )}
                </div>
                <div>
                  <div style={{ fontSize: 10, fontWeight: 700, color: "var(--text3)",
                      textTransform: "uppercase", marginBottom: 4, letterSpacing: .5 }}>Secs</div>
                  <input type="number" value={row.duration}
                    onChange={e => updRow(row.id, "duration", e.target.value)}
                    placeholder="30" style={inp({ width: 60 })} />
                </div>
                <div>
                  <div style={{ fontSize: 10, fontWeight: 700, color: "var(--text3)",
                      textTransform: "uppercase", marginBottom: 4, letterSpacing: .5 }}>Rate / Spot (₦)</div>
                  {(vendorRates.length === 0 || !row.rateId) ? (
                    <input type="number" value={row.customRate}
                      onChange={e => { updRow(row.id, "customRate", e.target.value); updRow(row.id, "rateId", ""); }}
                      placeholder="Enter rate" style={inp()} />
                  ) : (
                    <div style={{ ...inp(), background: "rgba(240,165,0,.08)", border: "1px solid rgba(240,165,0,.3)", color: "var(--accent)", fontWeight: 700 }}>
                      ₦{fmtN(vendorRates.find(r => r.id === row.rateId)?.ratePerSpot || 0)}
                    </div>
                  )}
                </div>
              </div>

              {/* Quick-select chips */}
              <div style={{ display: "flex", gap: 5, flexWrap: "wrap", marginBottom: 12 }}>
                {[
                  { label: "All",        act: () => quickSelect(row.id, "all") },
                  { label: "Weekdays",   act: () => quickSelect(row.id, [1,2,3,4,5]) },
                  { label: "M·W·F",      act: () => quickSelect(row.id, [1,3,5]) },
                  { label: "T·Th",       act: () => quickSelect(row.id, [2,4]) },
                  { label: "Weekends",   act: () => quickSelect(row.id, [0,6]) },
                  { label: "Mon only",   act: () => quickSelect(row.id, [1]) },
                  { label: "Clear",      act: () => quickSelect(row.id, "clear"), danger: true },
                ].map(({ label, act, danger }) => (
                  <button key={label} onClick={act} style={{
                    padding: "3px 10px", borderRadius: 20, fontSize: 10, fontWeight: 600,
                    cursor: "pointer", border: "1px solid var(--border2)",
                    background: danger ? "rgba(220,50,50,.1)" : "var(--bg3)",
                    color: danger ? "#e05555" : "var(--text2)", transition: "all .1s"
                  }}>{label}</button>
                ))}
              </div>

              {/* ── CALENDAR GRID ── */}
              <div style={{ userSelect: "none" }}>
                {/* Day-of-week headers */}
                <div style={{ display: "grid", gridTemplateColumns: "repeat(7, 1fr)",
                    gap: 3, marginBottom: 4 }}>
                  {DAY_LABELS.map((d, i) => (
                    <div key={d} style={{
                      textAlign: "center", fontSize: 10, fontWeight: 700, letterSpacing: .5,
                      color: i === 0 || i === 6 ? "var(--accent)" : "var(--text3)",
                      padding: "4px 0"
                    }}>{d}</div>
                  ))}
                </div>

                {/* Weeks */}
                {weeks.map((wk, wi) => (
                  <div key={wi} style={{ display: "grid", gridTemplateColumns: "repeat(7, 1fr)", gap: 3, marginBottom: 3 }}>
                    {wk.map((d, di) => {
                      if (!d) return <div key={`e${di}`} />;
                      const isActive = row.selectedDates.has(d);
                      const isWE = di === 0 || di === 6;
                      const dotw = validMonth ? new Date(yr, mIdx, d).getDay() : di;
                      const isWkend = dotw === 0 || dotw === 6;
                      return (
                        <div key={d} onClick={() => toggleDate(row.id, d)}
                          style={{
                            aspectRatio: "1", display: "flex", alignItems: "center",
                            justifyContent: "center", borderRadius: 8, cursor: "pointer",
                            border: isActive ? "2px solid var(--accent)" : "1px solid var(--border2)",
                            background: isActive ? "var(--accent)" : isWkend ? "rgba(240,165,0,.05)" : "var(--bg3)",
                            color: isActive ? "#000" : isWkend ? "var(--accent)" : "var(--text)",
                            fontWeight: isActive ? 800 : isWkend ? 600 : 400,
                            fontSize: 12, transition: "all .1s",
                            boxShadow: isActive ? "0 2px 8px rgba(240,165,0,.35)" : "none"
                          }}>
                          {d}
                        </div>
                      );
                    })}
                  </div>
                ))}
              </div>

            </div>
          </div>
        );
      })}

      {/* Footer actions */}
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginTop: 8 }}>
        <button onClick={() => setCalRows(r => [...r, blankCalRow()])}
          style={{ background: "none", border: "1px dashed var(--border2)", borderRadius: 8,
              padding: "7px 16px", cursor: "pointer", fontSize: 12, color: "var(--text2)",
              display: "flex", alignItems: "center", gap: 6 }}>
          <span style={{ fontSize: 18, lineHeight: 1 }}>+</span> Add Spot Row
        </button>
        <button onClick={onAdd}
          style={{ background: "var(--accent)", color: "#000", border: "none", borderRadius: 8,
              padding: "9px 22px", cursor: "pointer", fontWeight: 700, fontSize: 13,
              fontFamily: "'Syne',sans-serif" }}>
          ✓ Add to Schedule
        </button>
      </div>
    </div>
  );

}