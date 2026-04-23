import { useEffect, useRef, useState } from "react";
import Badge from "../components/Badge";
import Empty from "../components/Empty";
import Toast from "../components/Toast";
import Modal from "../components/Modal";
import Confirm from "../components/Confirm";
import { activeOnly, archivedOnly, isArchived, pctWithin } from "../utils/records";
import { fmtN } from "../utils/formatters";
import { hasPermission, readOnlyMessage, isAdmin, adminOnlyMessage } from "../constants/roles";
import { createRatesInSupabase, updateRateInSupabase, archiveRateInSupabase, restoreRateInSupabase, importRatesInSupabase, deleteRateInSupabase } from "../services/rates";
import { createAuditEventInSupabase } from "../services/notifications";
import { Field, Btn, Card } from "../components/ui/primitives";

const uid = () => Math.random().toString(36).slice(2) + Date.now().toString(36);
const downloadTextFile = (filename, content, mimeType = "text/plain;charset=utf-8") => {
  const blob = new Blob([content], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
};
const downloadRateTemplate = () => {
  const csv = [
    ["Vendor","Media","Programme","Timebelt","Duration","Rate","Discount","Commission","Notes"],
    ["Example FM","Radio","Morning Drive","06:00-09:00","30","150000","10","5","Prime time rate"],
    ["Example TV","Television","News at 9","21:00-21:30","45","450000","0","7.5","Headline bulletin"],
  ].map(row => row.map(v => `"${String(v).replace(/"/g, '""')}"`).join(",")).join("\n");

  downloadTextFile("mediadesk-rate-template.csv", csv, "text/csv;charset=utf-8");
};

/* ── RATES ──────────────────────────────────────────────── */
const blankRateRow = () => ({ _id: uid(), programme: "", timeBelt: "", duration: "30", ratePerSpot: "" });

/* Excel upload helper — uses SheetJS from CDN */
const loadSheetJS = () => new Promise((resolve, reject) => {
  if (window.XLSX) return resolve(window.XLSX);
  const s = document.createElement("script");
  s.src = "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";
  s.onload = () => resolve(window.XLSX);
  s.onerror = reject;
  document.head.appendChild(s);
});

const normaliseExcelRow = (row, vendors) => {
  const get = (...keys) => {
    for (const k of keys) {
      const found = Object.keys(row).find(rk => rk.trim().toLowerCase() === k.toLowerCase());
      if (found !== undefined && row[found] !== undefined && row[found] !== "") return String(row[found]).trim();
    }
    return "";
  };
  const vendorName = get("vendor", "vendor name", "station", "media owner");
  const normalizeVendorName = (value) => String(value ?? "").trim().replace(/\s+/g, " ").toLowerCase();
  const matchedVendor = vendors.find(v => normalizeVendorName(v.name) === normalizeVendorName(vendorName));
  return {
    _vendorName: vendorName,
    vendorId:    matchedVendor?.id || "",
    mediaType:   get("media", "media type", "type", "medium"),
    programme:   get("programme", "program", "slot", "programme / slot", "programme/slot"),
    timeBelt:    get("timebelt", "time belt", "time", "belt", "daypart"),
    duration:    get("duration", "dur", "duration (secs)", "secs") || "30",
    ratePerSpot: get("rate", "rate per spot", "rate/spot", "cost", "amount", "price"),
    discount:    get("discount", "disc", "volume discount", "disc%", "volume discount (%)") || "0",
    commission:  get("commission", "comm", "agency commission", "comm%", "agency commission (%)") || "0",
    vat: "0", notes: get("notes", "note", "remarks"),
    campaignId: "", clientId: "",
  };
};

const normRateText = (value) => String(value ?? "").trim().replace(/\s+/g, " ").toLowerCase();
const normRateTimeBelt = (value) => normRateText(value).replace(/\s*[-–—]\s*/g, "-");
const normRateDuration = (value) => String(value ?? "30").trim() || "30";
const makeRateDuplicateKey = ({ vendorId = "", vendorName = "", mediaType = "", programme = "", timeBelt = "", duration = "30" }) => {
  const vendorPart = vendorId ? `id:${vendorId}` : `name:${normRateText(vendorName)}`;
  return [
    vendorPart,
    normRateText(mediaType),
    normRateText(programme),
    normRateTimeBelt(timeBelt),
    normRateDuration(duration),
  ].join("|");
};


const ExcelImportModal = ({ vendors, existingRates, onImport, onClose }) => {
  const [step, setStep]         = useState("upload");
  const [rows, setRows]         = useState([]);
  const [errors, setErrors]     = useState([]);
  const [fileName, setFileName] = useState("");
  const [loading, setLoading]   = useState(false);
  const [selected, setSelected] = useState([]);
  const fileRef                 = useRef();
  const existingRateKeys = new Set(
    activeOnly(existingRates || []).map((rate) =>
      makeRateDuplicateKey({
        vendorId: rate.vendorId,
        vendorName: vendors.find(v => v.id === rate.vendorId)?.name || "",
        mediaType: rate.mediaType,
        programme: rate.programme,
        timeBelt: rate.timeBelt,
        duration: rate.duration || "30",
      })
    )
  );

  const handleFile = async (file) => {
    if (!file) return;
    setLoading(true); setFileName(file.name);
    try {
      const XLSX = await loadSheetJS();
      const buf  = await file.arrayBuffer();
      const wb   = XLSX.read(buf, { type: "array" });
      const ws   = wb.Sheets[wb.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json(ws, { defval: "" });
      if (!json.length) { setErrors(["The sheet appears to be empty."]); setLoading(false); return; }
      const parsed = json.map((r, i) => ({ ...normaliseExcelRow(r, vendors), _rowIdx: i }));
      const errs = [];
      const seenInFile = new Map();

      parsed.forEach((r, i) => {
        const duplicateKey = makeRateDuplicateKey({
          vendorId: r.vendorId,
          vendorName: r._vendorName,
          mediaType: r.mediaType,
          programme: r.programme,
          timeBelt: r.timeBelt,
          duration: r.duration || "30",
        });
        r._duplicateKey = duplicateKey;

        if (!r._vendorName) errs.push("Row " + (i + 2) + ": Missing Vendor name.");
        if (!r.ratePerSpot || isNaN(parseFloat(r.ratePerSpot))) errs.push("Row " + (i + 2) + ": Invalid or missing Rate.");
        if (!r.vendorId && r._vendorName) errs.push("Row " + (i + 2) + ": Vendor \"" + r._vendorName + "\" not found — vendor will be created automatically during import.");

        if (r.vendorId && r.programme && existingRateKeys.has(duplicateKey)) {
          errs.push("Row " + (i + 2) + ": Duplicate of an existing active rate card.");
        }

        if (r.vendorId && r.programme) {
          if (seenInFile.has(duplicateKey)) {
            errs.push("Row " + (i + 2) + ": Duplicate within this import file (matches row " + (seenInFile.get(duplicateKey) + 2) + ").");
          } else {
            seenInFile.set(duplicateKey, i);
          }
        }
      });

      const blockedRows = new Set(
        errs
          .filter(e => e.includes("Missing Vendor name") || e.includes("Invalid or missing Rate") || e.includes("Duplicate"))
          .map(e => {
            const m = e.match(/^Row\s+(\d+):/);
            return m ? Number(m[1]) - 2 : null;
          })
          .filter(v => v !== null)
      );

      setErrors(errs);
      setRows(parsed);
      setSelected(parsed.map((_, i) => i).filter(i => !blockedRows.has(i)));
      setStep("preview");
    } catch(e) { setErrors(["Failed to parse file. Ensure it is a valid .xlsx or .xls file."]); }
    setLoading(false);
  };

  const toggleRow = i => setSelected(s => s.includes(i) ? s.filter(x => x !== i) : [...s, i]);
  const toggleAll = () => setSelected(selected.length === rows.length ? [] : rows.map((_, i) => i));

  const confirmImport = () => {
    const toImport = rows.filter((_, i) => selected.includes(i)).map(r => ({
      id: uid(), createdAt: Date.now(),
      _vendorName: r._vendorName, vendorId: r.vendorId, mediaType: r.mediaType, programme: r.programme,
      timeBelt: r.timeBelt, duration: r.duration || "30",
      ratePerSpot: r.ratePerSpot, discount: r.discount || "0",
      commission: r.commission || "0", vat: "0",
      notes: r.notes, campaignId: "", clientId: "",
    }));
    onImport(toImport); setStep("done");
  };

  const hardErrors = errors.filter(e => e.includes("Missing Vendor name") || e.includes("Invalid or missing Rate") || e.includes("Duplicate"));
  const warnings   = errors.filter(e => !hardErrors.includes(e));

  return (
    <Modal title="Import Rate Cards from Excel" onClose={onClose} width={860}>
      {step === "upload" && (
        <div>
          <div style={{ background: "var(--bg3)", borderRadius: 12, padding: 16, marginBottom: 20, border: "1px solid var(--border2)" }}>
            <div style={{ fontSize: 12, fontWeight: 700, color: "var(--accent)", fontFamily: "\'Syne\',sans-serif", marginBottom: 10 }}>Required Excel Columns</div>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill,minmax(175px,1fr))", gap: 8 }}>
              {[{col:"Vendor",req:true,note:"Existing vendor will link, new vendor will auto-create"},{col:"Media",req:false,note:"e.g. Television, Radio"},{col:"Type",req:false,note:"Media type / category"},{col:"Programme",req:false,note:"Show or slot name"},{col:"Timebelt",req:false,note:"e.g. 21:00-21:30"},{col:"Duration",req:false,note:"In seconds, e.g. 30"},{col:"Rate",req:true,note:"Rate per spot in N"},{col:"Discounts",req:false,note:"Volume discount %"}].map(({ col, req, note }) => (
                <div key={col} style={{ background: "var(--bg4)", borderRadius: 8, padding: "8px 11px", border: req ? "1px solid rgba(240,165,0,.3)" : "1px solid var(--border)" }}>
                  <div style={{ fontSize: 11, fontWeight: 700, color: req ? "var(--accent)" : "var(--text)", marginBottom: 2 }}>{col}{req && <span style={{ color: "var(--red)", marginLeft: 3 }}>*</span>}</div>
                  <div style={{ fontSize: 10, color: "var(--text3)" }}>{note}</div>
                </div>
              ))}
            </div>
            <div style={{ marginTop: 10, display: "flex", gap: 10, justifyContent: "space-between", alignItems: "center", flexWrap: "wrap" }}><div style={{ fontSize: 11, color: "var(--text3)" }}>Column headers are flexible. First row must be headers.</div><Btn variant="ghost" size="sm" onClick={downloadRateTemplate}>⬇ Download Template</Btn></div>
          </div>
          <div onClick={() => fileRef.current?.click()}
            onDragOver={e => { e.preventDefault(); e.currentTarget.style.borderColor = "var(--accent)"; }}
            onDragLeave={e => { e.currentTarget.style.borderColor = "var(--border2)"; }}
            onDrop={e => { e.preventDefault(); e.currentTarget.style.borderColor = "var(--border2)"; handleFile(e.dataTransfer.files[0]); }}
            style={{ background: "var(--bg3)", border: "2px dashed var(--border2)", borderRadius: 14, padding: "38px 24px", textAlign: "center", cursor: "pointer", transition: "all .2s" }}>
            <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }} onChange={e => handleFile(e.target.files[0])} />
            {loading
              ? <div><div style={{ fontSize: 36, marginBottom: 10 }}>&#x23F3;</div><div style={{ fontFamily: "\'Syne\',sans-serif", fontWeight: 700 }}>Parsing file...</div></div>
              : <><div style={{ fontSize: 48, marginBottom: 10 }}>&#x1F4CA;</div><div style={{ fontFamily: "\'Syne\',sans-serif", fontWeight: 700, fontSize: 16, marginBottom: 6 }}>Drop your Excel file here</div><div style={{ color: "var(--text2)", fontSize: 13, marginBottom: 14 }}>or click to browse</div><Btn variant="secondary" size="sm">Choose File</Btn></>}
          </div>
          {errors.length > 0 && <div style={{ marginTop: 12, background: "rgba(239,68,68,.08)", border: "1px solid rgba(239,68,68,.25)", borderRadius: 10, padding: 12 }}>{errors.map((e, i) => <div key={i} style={{ fontSize: 12, color: "var(--red)" }}>Warning: {e}</div>)}</div>}
        </div>
      )}
      {step === "preview" && (
        <div>
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 12, flexWrap: "wrap", gap: 8 }}>
            <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center" }}>
              <Badge color="blue">{fileName}</Badge>
              <Badge color="green">{rows.length} rows parsed</Badge>
              <Badge color="accent">{selected.length} selected</Badge>
              {warnings.length > 0 && <Badge color="orange">{warnings.length} warning{warnings.length > 1 ? "s" : ""}</Badge>}
              {hardErrors.length > 0 && <Badge color="red">{hardErrors.length} error{hardErrors.length > 1 ? "s" : ""}</Badge>}
            </div>
            <Btn variant="ghost" size="sm" onClick={() => { setStep("upload"); setRows([]); setErrors([]); setSelected([]); }}>Re-upload</Btn>
          </div>
          {warnings.length > 0 && <div style={{ background: "rgba(249,115,22,.08)", border: "1px solid rgba(249,115,22,.25)", borderRadius: 10, padding: "9px 12px", marginBottom: 10 }}>{warnings.map((w, i) => <div key={i} style={{ fontSize: 11, color: "var(--orange)" }}>Warning: {w}</div>)}</div>}
          {hardErrors.length > 0 && <div style={{ background: "rgba(239,68,68,.08)", border: "1px solid rgba(239,68,68,.25)", borderRadius: 10, padding: "9px 12px", marginBottom: 10 }}>{hardErrors.map((e, i) => <div key={i} style={{ fontSize: 11, color: "var(--red)" }}>Error: {e}</div>)}</div>}
          <div style={{ overflowX: "auto", maxHeight: 360, overflowY: "auto", borderRadius: 10, border: "1px solid var(--border)" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 700, fontSize: 12 }}>
              <thead style={{ position: "sticky", top: 0, background: "var(--bg3)", zIndex: 1 }}>
                <tr>
                  <th style={{ padding: "9px 12px", textAlign: "center", borderBottom: "1px solid var(--border)", width: 36 }}>
                    <input type="checkbox" checked={selected.length === rows.length} onChange={toggleAll} style={{ accentColor: "var(--accent)", cursor: "pointer", width: 14, height: 14 }} />
                  </th>
                  {["Vendor","Media Type","Programme","Time Belt","Dur","Rate","Disc%","Comm%","Status"].map(h => (
                    <th key={h} style={{ padding: "9px 12px", textAlign: "left", fontSize: 10, fontWeight: 600, textTransform: "uppercase", color: "var(--text3)", borderBottom: "1px solid var(--border)", whiteSpace: "nowrap" }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {rows.map((r, i) => {
                  const isSelected = selected.includes(i);
                  const hasError   = hardErrors.some(e => e.startsWith("Row " + (i + 2) + ":"));
                  const hasWarn    = warnings.some(w => w.startsWith("Row " + (i + 2) + ":"));
                  return (
                    <tr key={i} style={{ borderBottom: "1px solid var(--border)", opacity: isSelected ? 1 : 0.4 }}>
                      <td style={{ padding: "9px 12px", textAlign: "center" }}><input type="checkbox" checked={isSelected} onChange={() => toggleRow(i)} style={{ accentColor: "var(--accent)", cursor: "pointer", width: 14, height: 14 }} /></td>
                      <td style={{ padding: "9px 12px", fontWeight: 600 }}>
                        {r._vendorName || "—"}
                        {r.vendorId && <div style={{ fontSize: 10, color: "var(--green)" }}>Linked</div>}
                        {!r.vendorId && r._vendorName && <div style={{ fontSize: 10, color: "var(--orange)" }}>Will auto-create</div>}
                      </td>
                      <td style={{ padding: "9px 12px", color: "var(--text2)" }}>{r.mediaType || "—"}</td>
                      <td style={{ padding: "9px 12px", color: "var(--text2)" }}>{r.programme || "—"}</td>
                      <td style={{ padding: "9px 12px", color: "var(--text2)" }}>{r.timeBelt || "—"}</td>
                      <td style={{ padding: "9px 12px", color: "var(--text2)" }}>{r.duration || "30"}</td>
                      <td style={{ padding: "9px 12px", fontWeight: 600, color: r.ratePerSpot ? "var(--accent)" : "var(--red)" }}>{r.ratePerSpot || "MISSING"}</td>
                      <td style={{ padding: "9px 12px", color: "var(--red)" }}>{r.discount || 0}%</td>
                      <td style={{ padding: "9px 12px", color: "var(--red)" }}>{r.commission || 0}%</td>
                      <td style={{ padding: "9px 12px" }}>{hasError ? <Badge color="red">Error</Badge> : hasWarn ? <Badge color="orange">Warning</Badge> : <Badge color="green">Ready</Badge>}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginTop: 16, flexWrap: "wrap", gap: 10 }}>
            <div style={{ fontSize: 12, color: "var(--text2)" }}>{selected.length === 0 ? "No rows selected." : selected.length + " of " + rows.length + " rows will be imported."}</div>
            <div style={{ display: "flex", gap: 10 }}>
              <Btn variant="ghost" onClick={onClose}>Cancel</Btn>
              <Btn onClick={confirmImport} disabled={selected.length === 0} icon="save">Import {selected.length} Rate Card{selected.length !== 1 ? "s" : ""}</Btn>
            </div>
          </div>
        </div>
      )}
      {step === "done" && (
        <div style={{ textAlign: "center", padding: "36px 16px" }}>
          <div style={{ fontSize: 56, marginBottom: 14 }}>&#x2705;</div>
          <div style={{ fontFamily: "\'Syne\',sans-serif", fontWeight: 800, fontSize: 22, marginBottom: 8 }}>Import Complete!</div>
          <div style={{ color: "var(--text2)", fontSize: 14, marginBottom: 24 }}>{selected.length} rate card{selected.length !== 1 ? "s" : ""} from {fileName} saved.</div>
          <Btn onClick={onClose}>Close</Btn>
        </div>
      )}
    </Modal>
  );
};


const RatesPage = ({ rates, setRates, vendors, setVendors, clients, campaigns, user }) => {
  const canManage = hasPermission(user, "manageRates");
  const canDelete = isAdmin(user);
  const [modal, setModal]         = useState(null);
  const [importModal, setImportModal] = useState(false);
  const [search, setSearch]       = useState("");
  const [filterV, setFilterV]     = useState("");
  const [filterChannelType, setFilterChannelType] = useState("");
  const [viewMode, setViewMode]   = useState("active");
  const [toast, setToast]         = useState(null);
  const [confirm, setConfirm]     = useState(null);

  /* shared vendor-level fields */
  const blankHeader = { vendorId: "", mediaType: "", discount: "", commission: "", notes: "" };
  const [hdr, setHdr]   = useState(blankHeader);
  const uh = k => v => setHdr(p => ({ ...p, [k]: v }));

  /* multi-row programme list */
  const [rows, setRows] = useState([blankRateRow()]);
  const updRow = (id, k, v) => setRows(rs => rs.map(r => r._id === id ? { ...r, [k]: v } : r));
  const addRow = () => setRows(rs => [...rs, blankRateRow()]);
  const delRow = id => setRows(rs => rs.length > 1 ? rs.filter(r => r._id !== id) : rs);

  /* editing a single existing rate — keep header + one row */
  const [editId, setEditId] = useState(null);

  /* auto-fill discount/commission from vendor */
  useEffect(() => {
    if (hdr.vendorId) {
      const v = vendors.find(x => x.id === hdr.vendorId);
      if (v) setHdr(p => ({ ...p, discount: p.discount || v.discount || "", commission: p.commission || v.commission || "", mediaType: p.mediaType || v.type || "" }));
    }
  }, [hdr.vendorId]);

  const calcNet = r => {
    const rate = parseFloat(r.ratePerSpot) || 0;
    return rate * (1 - (parseFloat(hdr.discount || r.discount) || 0) / 100) * (1 - (parseFloat(hdr.commission || r.commission) || 0) / 100);
  };
  const calcNetR = (r, disc, comm) => {
    const rate = parseFloat(r.ratePerSpot) || 0;
    return rate * (1 - (parseFloat(disc) || 0) / 100) * (1 - (parseFloat(comm) || 0) / 100);
  };

  const openAdd = () => {
    if (!canManage) return setToast({ msg: readOnlyMessage(user), type: "error" });
    setHdr(blankHeader); setRows([blankRateRow()]); setEditId(null); setModal("add");
  };
  const openEdit = r => {
    setHdr({ vendorId: r.vendorId, mediaType: r.mediaType || "", discount: r.discount || "", commission: r.commission || "", notes: r.notes || "" });
    setRows([{ _id: uid(), programme: r.programme || "", timeBelt: r.timeBelt || "", duration: r.duration || "30", ratePerSpot: r.ratePerSpot || "" }]);
    setEditId(r.id); setModal("edit");
  };

  const save = async () => {
    if (!canManage) return setToast({ msg: readOnlyMessage(user), type: "error" });
    if (!hdr.vendorId) return setToast({ msg: "Please select a vendor.", type: "error" });
    if (!pctWithin(hdr.discount) || !pctWithin(hdr.commission)) return setToast({ msg: "Discount and commission must be between 0 and 100.", type: "error" });
    const validRows = rows.filter(r => r.programme && r.ratePerSpot);
    if (!validRows.length) return setToast({ msg: "Add at least one programme with a rate.", type: "error" });
    if (validRows.some(r => (parseFloat(r.ratePerSpot) || 0) <= 0)) return setToast({ msg: "Every saved rate must be greater than zero.", type: "error" });
    const existingRateKeys = new Set(
      rates
        .filter(existing => existing.id !== editId && !isArchived(existing))
        .map(existing => makeRateDuplicateKey({
          vendorId: existing.vendorId,
          vendorName: vendors.find(v => v.id === existing.vendorId)?.name || "",
          mediaType: existing.mediaType,
          programme: existing.programme,
          timeBelt: existing.timeBelt,
          duration: existing.duration || "30",
        }))
    );

    const seenDraftKeys = new Set();
    const duplicateRow = validRows.find(row => {
      const key = makeRateDuplicateKey({
        vendorId: hdr.vendorId,
        vendorName: vendors.find(v => v.id === hdr.vendorId)?.name || "",
        mediaType: hdr.mediaType,
        programme: row.programme,
        timeBelt: row.timeBelt,
        duration: row.duration || "30",
      });
      if (existingRateKeys.has(key) || seenDraftKeys.has(key)) return true;
      seenDraftKeys.add(key);
      return false;
    });

    if (duplicateRow) {
      return setToast({
        msg: `A matching rate card already exists for ${duplicateRow.programme} (${duplicateRow.timeBelt || "no time belt"}, ${duplicateRow.duration || "30"}s).`,
        type: "error"
      });
    }

    try {
      if (modal === "edit" && editId) {
        const row = validRows[0];
        const updatedRate = await updateRateInSupabase(editId, hdr, row);
        setRates(v => v.map(x => x.id === editId ? updatedRate : x));
        createAuditEventInSupabase({ agencyId: user.agencyId, recordType: "rate", recordId: editId, action: "updated", actor: user, metadata: { vendorId: updatedRate.vendorId || hdr.vendorId || "", programme: updatedRate.programme || row.programme || "" } }).catch(error => console.error("Failed to write audit event:", error));
        setToast({ msg: "Rate updated.", type: "success" });
      } else {
        const createdRates = await createRatesInSupabase(user.agencyId, user.id, hdr, validRows);
        setRates(v => [...createdRates, ...v]);
        createAuditEventInSupabase({ agencyId: user.agencyId, recordType: "rate", recordId: null, action: "created", actor: user, note: `${createdRates.length} rate card${createdRates.length !== 1 ? "s" : ""} added.`, metadata: { count: createdRates.length, vendorId: hdr.vendorId || "" } }).catch(error => console.error("Failed to write audit event:", error));
        setToast({ msg: `${createdRates.length} rate card${createdRates.length !== 1 ? "s" : ""} added.`, type: "success" });
      }
      setModal(null);
    } catch (e) {
      setToast({ msg: e.message || "Failed to save rate.", type: "error" });
    }
  };

  const del = async id => {
    if (!canManage) return setToast({ msg: readOnlyMessage(user), type: "error" });
    if (!canManage) return setToast({ msg: readOnlyMessage(user), type: "error" });
    try {
      const archived = await archiveRateInSupabase(id);
      setRates(v => v.map(x => x.id === id ? archived : x));
      createAuditEventInSupabase({ agencyId: user.agencyId, recordType: "rate", recordId: id, action: "archived", actor: user, metadata: { programme: archived.programme || "", vendorId: archived.vendorId || "" } }).catch(error => console.error("Failed to write audit event:", error));
      setToast({ msg: "Rate card archived.", type: "success" });
      setConfirm(null);
    } catch (e) {
      setToast({ msg: e.message || "Failed to archive rate.", type: "error" });
    }
  };
  const restore = async id => {
    if (!canManage) return setToast({ msg: readOnlyMessage(user), type: "error" });
    if (!canManage) return setToast({ msg: readOnlyMessage(user), type: "error" });
    try {
      const restored = await restoreRateInSupabase(id);
      setRates(v => v.map(x => x.id === id ? restored : x));
      createAuditEventInSupabase({ agencyId: user.agencyId, recordType: "rate", recordId: id, action: "restored", actor: user, metadata: { programme: restored.programme || "", vendorId: restored.vendorId || "" } }).catch(error => console.error("Failed to write audit event:", error));
      setToast({ msg: "Rate card restored.", type: "success" });
    } catch (e) {
      setToast({ msg: e.message || "Failed to restore rate.", type: "error" });
    }
  };
  const hardDelete = async id => {
    if (!canDelete) return setToast({ msg: adminOnlyMessage(user), type: "error" });
    try {
      const target = rates.find(x => x.id === id);
      await deleteRateInSupabase(id);
      setRates(v => v.filter(x => x.id !== id));
      createAuditEventInSupabase({ agencyId: user.agencyId, recordType: "rate", recordId: id, action: "deleted", actor: user, metadata: { programme: target?.programme || "", vendorId: target?.vendorId || "" } }).catch(error => console.error("Failed to write audit event:", error));
      setToast({ msg: "Rate card deleted permanently.", type: "success" });
      setConfirm(null);
    } catch (e) {
      setToast({ msg: e.message || "Failed to delete rate.", type: "error" });
    }
  };

  const handleExcelImport = async (newRates) => {
    try {
      const { insertedRates, duplicateRows, createdVendors = [] } = await importRatesInSupabase(user.agencyId, user.id, newRates);
      setRates(v => [...insertedRates, ...v]);

      if (typeof setVendors === "function" && createdVendors.length) {
        setVendors(prev => {
          const merged = [...createdVendors, ...(prev || [])];
          const seen = new Set();
          return merged.filter(vendor => {
            if (!vendor?.id || seen.has(vendor.id)) return false;
            seen.add(vendor.id);
            return true;
          });
        });
      }

      if (insertedRates.length && duplicateRows.length) {
        setToast({
          msg: `${insertedRates.length} rate card${insertedRates.length !== 1 ? "s" : ""} imported. ${createdVendors.length ? `${createdVendors.length} vendor${createdVendors.length !== 1 ? "s" : ""} auto-created. ` : ""}${duplicateRows.length} duplicate row${duplicateRows.length !== 1 ? "s were" : " was"} skipped.`,
          type: "success"
        });
      } else if (insertedRates.length) {
        setToast({
          msg: `${insertedRates.length} rate card${insertedRates.length !== 1 ? "s" : ""} imported successfully!${createdVendors.length ? ` ${createdVendors.length} vendor${createdVendors.length !== 1 ? "s were" : " was"} auto-created.` : ""}`,
          type: "success"
        });
      } else {
        setToast({ msg: "No new rate cards were imported.", type: "error" });
      }
    } catch (e) {
      setToast({ msg: e.message || "Failed to import rates.", type: "error" });
    }
  };

  const visibleRates = viewMode === "archived" ? archivedOnly(rates) : viewMode === "all" ? rates : activeOnly(rates);
  const channelTypeOptions = Array.from(
    new Set((visibleRates || []).map(rate => String(rate.mediaType || "").trim()).filter(Boolean))
  )
    .sort((a, b) => a.localeCompare(b))
    .map(mediaType => ({ value: mediaType, label: mediaType }));
  const filtered = visibleRates.filter(r => {
    const vn = vendors.find(v => v.id === r.vendorId)?.name || "";
    return `${vn} ${r.programme || ""}`.toLowerCase().includes(search.toLowerCase())
      && (!filterV || r.vendorId === filterV)
      && (!filterChannelType || String(r.mediaType || "") === filterChannelType);
  });

  const inputSt = { background: "var(--bg2)", border: "1px solid var(--border2)", borderRadius: 7, padding: "7px 10px", color: "var(--text)", fontSize: 12, outline: "none", width: "100%" };

  return (
    <div className="fade">
      {toast && <Toast msg={toast.msg} type={toast.type} onDone={() => setToast(null)} />}
      {confirm && <Confirm msg={confirm.msg} onYes={confirm.onYes} onNo={() => setConfirm(null)} />}

      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 24, flexWrap: "wrap", gap: 10 }}>
        <div><h1 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 800, fontSize: 24 }}>Media Rates</h1><p style={{ color: "var(--text2)", marginTop: 3, fontSize: 13 }}>Rate cards across media owners</p></div>
        <div style={{ display: "flex", gap: 10 }}>
          <Btn variant="blue" icon="&#x1F4CA;" onClick={() => setImportModal(true)}>Import from Excel</Btn>
          <Btn icon="+" onClick={openAdd}>Add Rate Card</Btn>
        </div>
      </div>

      <div style={{ display: "flex", gap: 10, marginBottom: 18, flexWrap: "wrap" }}>
        <div style={{ flex: 1, minWidth: 180 }}><Field value={search} onChange={setSearch} placeholder="Search by vendor or programme…" /></div>
        <Field value={filterV} onChange={setFilterV} options={activeOnly(vendors).map(v => ({ value: v.id, label: v.name }))} placeholder="All Vendors" />
        <Field value={filterChannelType} onChange={setFilterChannelType} options={channelTypeOptions} placeholder="All Channel Types" />
        <Field value={viewMode} onChange={setViewMode} options={[{value:"active",label:"Active"},{value:"archived",label:"Archived"},{value:"all",label:"All"}]} />
      </div>

      {filtered.length === 0 ? <Card><Empty icon="💰" title="No rate cards" sub="Add media rates to use in MPO generation" /></Card> :
        <div style={{ overflowX: "auto", overflowY: "auto", maxHeight: "calc(100vh - 260px)", borderRadius: 12, border: "1px solid var(--border)" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 800 }}>
            <thead><tr style={{ background: "var(--bg3)" }}>
              {["Vendor","Programme","Time Belt","Type","Dur","Rate/Spot","Disc%","Comm%","Net Rate",""].map(h => <th key={h} style={{ position: "sticky", top: 0, zIndex: 5, background: "var(--bg3)", padding: "7px 9px", textAlign: "left", fontSize: 9, fontWeight: 600, textTransform: "uppercase", letterSpacing: ".07em", color: "var(--text3)", borderBottom: "1px solid var(--border)", whiteSpace: "nowrap" }}>{h}</th>)}
            </tr></thead>
            <tbody>
              {filtered.map((r, i) => {
                const vn  = vendors.find(v => v.id === r.vendorId)?.name || "—";
                const net = calcNetR(r, r.discount, r.commission);
                return (
                  <tr key={r.id} style={{ borderBottom: "1px solid var(--border)", background: i % 2 === 0 ? "transparent" : "rgba(255,255,255,.01)" }}
                    onMouseEnter={e => e.currentTarget.style.background = "var(--bg3)"}
                    onMouseLeave={e => e.currentTarget.style.background = i % 2 === 0 ? "transparent" : "rgba(255,255,255,.01)"}>
                    <td style={{ padding: "8px 9px", fontWeight: 600, fontSize: 12 }}>{vn}</td>
                    <td style={{ padding: "8px 9px", color: "var(--text2)", fontSize: 12 }}>{r.programme || "—"}</td>
                    <td style={{ padding: "8px 9px", color: "var(--text2)", fontSize: 12 }}>{r.timeBelt || "—"}</td>
                    <td style={{ padding: "8px 9px", fontSize: 12 }}><div style={{ display: "flex", gap: 5, flexWrap: "wrap" }}><Badge color="blue">{r.mediaType || "—"}</Badge>{isArchived(r) && <Badge color="red">Archived</Badge>}</div></td>
                    <td style={{ padding: "8px 9px", color: "var(--text2)", fontSize: 12 }}>{r.duration}"</td>
                    <td style={{ padding: "8px 9px", fontWeight: 600, fontSize: 12, color: "var(--accent)" }}>{fmtN(r.ratePerSpot)}</td>
                    <td style={{ padding: "11px 12px", color: "var(--red)" }}>{r.discount || 0}%</td>
                    <td style={{ padding: "11px 12px", color: "var(--red)" }}>{r.commission || 0}%</td>
                    <td style={{ padding: "8px 9px", fontWeight: 700, color: "var(--green)", fontSize: 12 }}>{fmtN(net)}</td>
                    <td style={{ padding: "11px 12px" }}>
                      <div style={{ display: "flex", gap: 5 }}>
                        {canManage && <Btn variant="ghost" size="sm" onClick={() => openEdit(r)}>✏️</Btn>}
                        {canManage && (isArchived(r) ? <Btn variant="success" size="sm" onClick={() => restore(r.id)}>↩</Btn> : <Btn variant="danger" size="sm" onClick={() => setConfirm({ msg: `Archive this rate card? Existing MPOs will still retain their saved values.`, onYes: () => del(r.id) })}>🗄</Btn>)}
                        {canDelete && <Btn variant="danger" size="sm" onClick={() => setConfirm({ msg: `Delete this rate card permanently? This cannot be undone.`, onYes: () => hardDelete(r.id) })}>🗑</Btn>}
                      </div>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>}

      {/* Excel Import Modal */}
      {importModal && (
        <ExcelImportModal vendors={vendors} existingRates={rates} onImport={handleExcelImport} onClose={() => setImportModal(false)} />
      )}

      {/* Add / Edit Modal */}
      {modal !== null && (
        <Modal title={modal === "edit" ? "Edit Rate Card" : "Add Rate Cards"} onClose={() => setModal(null)} width={780}>

          {/* ── Vendor-level header ── */}
          <div style={{ background: "var(--bg3)", borderRadius: 12, padding: "14px 16px", marginBottom: 18, border: "1px solid var(--border2)" }}>
            <div style={{ fontSize: 11, fontWeight: 700, color: "var(--accent)", textTransform: "uppercase", letterSpacing: ".07em", marginBottom: 12, fontFamily: "'Syne',sans-serif" }}>Vendor Details</div>
            <div style={{ display: "grid", gridTemplateColumns: "2fr 1.5fr 1fr 1fr", gap: 12 }}>
              <Field label="Vendor" value={hdr.vendorId} onChange={uh("vendorId")} options={vendors.map(v => ({ value: v.id, label: v.name }))} placeholder="Select vendor" required note="Disc & Comm auto-filled" />
              <Field label="Media Type" value={hdr.mediaType} onChange={uh("mediaType")} options={["Television","Radio","Print","Digital/Online","Out-of-Home (OOH)","Cinema"]} />
              <Field label="Volume Discount (%)" type="number" value={hdr.discount} onChange={uh("discount")} placeholder="0" note="e.g. 27 for 27%" />
              <Field label="Agency Commission (%)" type="number" value={hdr.commission} onChange={uh("commission")} placeholder="0" note="e.g. 15 for 15%" />
            </div>
            <div style={{ marginTop: 12 }}>
              <Field label="Notes" value={hdr.notes} onChange={uh("notes")} placeholder="Additional notes…" />
            </div>
          </div>

          {/* ── Programme rows ── */}
          <div style={{ marginBottom: 10 }}>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 10 }}>
              <div style={{ fontSize: 11, fontWeight: 700, color: "var(--text2)", textTransform: "uppercase", letterSpacing: ".07em", fontFamily: "'Syne',sans-serif" }}>
                Programmes / Slots — {modal === "edit" ? "1 entry" : `${rows.length} row${rows.length !== 1 ? "s" : ""}`}
              </div>
              {modal !== "edit" && (
                <button onClick={addRow} style={{ display: "flex", alignItems: "center", gap: 6, background: "none", border: "1px dashed var(--border2)", borderRadius: 7, padding: "5px 12px", cursor: "pointer", fontSize: 12, color: "var(--text2)" }}>
                  <span style={{ fontSize: 16 }}>+</span> Add Row
                </button>
              )}
            </div>

            {/* column headers */}
            <div style={{ display: "grid", gridTemplateColumns: "2.5fr 1.5fr 80px 1.5fr 36px", gap: 8, marginBottom: 6, padding: "0 4px" }}>
              {["Programme / Slot","Time Belt","Dur (s)","Rate per Spot (₦)",""].map(h => (
                <div key={h} style={{ fontSize: 10, fontWeight: 600, color: "var(--text3)", textTransform: "uppercase", letterSpacing: ".07em" }}>{h}</div>
              ))}
            </div>

            <div style={{ display: "flex", flexDirection: "column", gap: 8, maxHeight: 340, overflowY: "auto" }}>
              {rows.map((row, ri) => {
                const net = parseFloat(row.ratePerSpot) ? fmtN(parseFloat(row.ratePerSpot) * (1 - (parseFloat(hdr.discount)||0)/100) * (1 - (parseFloat(hdr.commission)||0)/100)) : null;
                return (
                  <div key={row._id} style={{ display: "grid", gridTemplateColumns: "2.5fr 1.5fr 80px 1.5fr 36px", gap: 8, alignItems: "center", background: "var(--bg3)", borderRadius: 9, padding: "10px 12px", border: "1px solid var(--border)" }}>
                    <div>
                      <input value={row.programme} onChange={e => updRow(row._id, "programme", e.target.value)}
                        placeholder="e.g. NTA Network News" style={inputSt}
                        onFocus={e => e.target.style.borderColor="var(--accent)"}
                        onBlur={e => e.target.style.borderColor="var(--border2)"} />
                    </div>
                    <div>
                      <input value={row.timeBelt} onChange={e => updRow(row._id, "timeBelt", e.target.value)}
                        placeholder="e.g. 21:00–21:30" style={inputSt}
                        onFocus={e => e.target.style.borderColor="var(--accent)"}
                        onBlur={e => e.target.style.borderColor="var(--border2)"} />
                    </div>
                    <div>
                      <input type="number" value={row.duration} onChange={e => updRow(row._id, "duration", e.target.value)}
                        placeholder="30" style={inputSt}
                        onFocus={e => e.target.style.borderColor="var(--accent)"}
                        onBlur={e => e.target.style.borderColor="var(--border2)"} />
                    </div>
                    <div>
                      <input type="number" value={row.ratePerSpot} onChange={e => updRow(row._id, "ratePerSpot", e.target.value)}
                        placeholder="0" style={inputSt}
                        onFocus={e => e.target.style.borderColor="var(--accent)"}
                        onBlur={e => e.target.style.borderColor="var(--border2)"} />
                      {net && <div style={{ fontSize: 10, color: "var(--green)", marginTop: 2 }}>Net: {net}</div>}
                    </div>
                    <div style={{ display: "flex", justifyContent: "center" }}>
                      {rows.length > 1 ? (
                        <button onClick={() => delRow(row._id)} style={{ background: "none", border: "none", cursor: "pointer", color: "var(--red)", fontSize: 18, lineHeight: 1, opacity: .7 }}>×</button>
                      ) : (
                        <span style={{ fontSize: 12, color: "var(--text3)", opacity: .4 }}>#{ri+1}</span>
                      )}
                    </div>
                  </div>
                );
              })}
            </div>

            {/* live summary */}
            {rows.filter(r => r.ratePerSpot).length > 0 && modal !== "edit" && (
              <div style={{ marginTop: 12, background: "rgba(240,165,0,.07)", border: "1px solid rgba(240,165,0,.18)", borderRadius: 9, padding: "10px 14px", display: "flex", gap: 20, flexWrap: "wrap", alignItems: "center" }}>
                <span style={{ fontSize: 11, color: "var(--text3)" }}>Summary:</span>
                <span style={{ fontSize: 12, fontWeight: 600, color: "var(--accent)" }}>{rows.filter(r => r.programme && r.ratePerSpot).length} valid row{rows.filter(r => r.programme && r.ratePerSpot).length !== 1 ? "s" : ""} ready to save</span>
                {hdr.discount > 0 && <span style={{ fontSize: 11, color: "var(--text2)" }}>Vol Disc: <strong style={{ color:"var(--red)" }}>{hdr.discount}%</strong></span>}
                {hdr.commission > 0 && <span style={{ fontSize: 11, color: "var(--text2)" }}>Comm: <strong style={{ color:"var(--red)" }}>{hdr.commission}%</strong></span>}
              </div>
            )}
          </div>

          <div style={{ display: "flex", gap: 10, justifyContent: "flex-end", marginTop: 18, paddingTop: 16, borderTop: "1px solid var(--border)" }}>
            <Btn variant="ghost" onClick={() => setModal(null)}>Cancel</Btn>
            <Btn onClick={save} icon="💾">{modal === "edit" ? "Save Changes" : `Add ${rows.filter(r=>r.programme&&r.ratePerSpot).length || ""} Rate Card${rows.filter(r=>r.programme&&r.ratePerSpot).length !== 1 ? "s" : ""}`}</Btn>
          </div>
        </Modal>
      )}
    </div>
  );
};


export default RatesPage;
