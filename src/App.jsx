import React, { useState, useEffect, useRef, useCallback } from "react";
import * as XLSX from "xlsx";
import {
  Download, Upload, Plus, Trash2, X, ChevronDown, GripVertical,
  Image as ImageIcon, FileSpreadsheet, Check, ArrowLeft, ArrowRight,
  Search, Pencil, History, ChevronRight, AlertTriangle, RefreshCw
} from "lucide-react";
import { storageGet, storageSet, storageDelete } from "./lib/storage.js";
import "./styles.css";

/* ============================================================
   CONSTANTS
============================================================ */
const STAGES = ["Lead", "QLead", "Book", "Exam", "Sur"];
const STAGE_LABELS = { Lead: "Lead", QLead: "QLead", Book: "Book", Exam: "Exam", Sur: "Sur" };
const STAGE_COLORS = ["#BFD6FB", "#8FB8F6", "#5A93EF", "#2F6FE0", "#163E9E"];

const DATA_TYPES = [
  { id: "budgetTotal", label: "Budget - Total" },
  { id: "budgetDetail", label: "Budget - Detail" },
  { id: "channelsSummary", label: "Digital Channels - Summary" },
  { id: "channelsBranding", label: "Digital Channels - Branding Ads" },
  { id: "lasikKpi", label: "LASIK KPI Target" },
  { id: "lasikActual", label: "LASIK Actual" },
  { id: "taskList", label: "Task List" },
];

const STATUS_COLORS = {
  "Completed": { bg: "#E4F6EC", fg: "#1E9E63" },
  "In Progress": { bg: "#E7F0FE", fg: "#2D7FF9" },
  "Not Started": { bg: "#EEF0F4", fg: "#7C8595" },
  "Pending Budget": { bg: "#F1E9FB", fg: "#7C4FE0" },
  "In Review": { bg: "#E3F6FA", fg: "#17A2B8" },
};

const uid = () => Math.random().toString(36).slice(2) + Date.now().toString(36);

/* ============================================================
   STORAGE HELPERS  (IndexedDB-backed — persists across sessions,
   see src/lib/storage.js for the underlying implementation)
============================================================ */
async function storeGet(key) {
  try {
    const res = await storageGet(key);
    if (!res) return null;
    return JSON.parse(res.value);
  } catch (e) {
    return null;
  }
}
async function storeSet(key, value) {
  try {
    await storageSet(key, JSON.stringify(value));
    return true;
  } catch (e) {
    console.error("storage set failed", e);
    return false;
  }
}
async function storeDelete(key) {
  try {
    await storageDelete(key);
  } catch (e) { /* ignore */ }
}

/* ============================================================
   SAMPLE DATA (first-run defaults, approximating the reference)
============================================================ */
function sampleReport(id) {
  return {
    id,
    label: id,
    timeRange: "01/08/2026 - 07/08/2026",
    createdAt: Date.now(),
    budgetTotal: {
      columns: ["TT", "BU", "BU Branding Budget 2026", "Chi phí Quý 1", "Ngân sách Quý 2", "Chi phí Quý 2", "Còn lại Q3 + Q4"],
      rows: [
        { TT: 1, BU: "Hospital", "BU Branding Budget 2026": 1000000000, "Chi phí Quý 1": 733032134, "Ngân sách Quý 2": 106967866, "Chi phí Quý 2": 116327336, "Còn lại Q3 + Q4": 150640530 },
        { TT: 2, BU: "Phaco", "BU Branding Budget 2026": 2416800000, "Chi phí Quý 1": 381136009, "Ngân sách Quý 2": 897763991, "Chi phí Quý 2": 404935000, "Còn lại Q3 + Q4": 1630728991 },
        { TT: 3, BU: "Optometry", "BU Branding Budget 2026": 509700000, "Chi phí Quý 1": 58680210, "Ngân sách Quý 2": 223019790, "Chi phí Quý 2": 205822000, "Còn lại Q3 + Q4": 245197790 },
        { TT: 4, BU: "Lasik", "BU Branding Budget 2026": 925000000, "Chi phí Quý 1": 219035019, "Ngân sách Quý 2": 193019790, "Chi phí Quý 2": 119666668, "Còn lại Q3 + Q4": 586298313 },
      ],
    },
    budgetDetail: { columns: ["TT", "BU", "Hạng mục", "Số tiền"], rows: [] },
    channelsSummary: {
      columns: ["Hạng mục", "Tháng 6/2026", "01/07/2026 (01-28/07)", "Tháng 8/2026", "Tháng 9/2026"],
      rows: [
        { "Hạng mục": "Fanpage chính - Followers", "Tháng 6/2026": 49510, "01/07/2026 (01-28/07)": 49169, "Tháng 8/2026": 52000, "Tháng 9/2026": 52000 },
        { "Hạng mục": "Fanpage chính - Engagement", "Tháng 6/2026": 29654, "01/07/2026 (01-28/07)": 19535, "Tháng 8/2026": 2100, "Tháng 9/2026": 2200 },
        { "Hạng mục": "Fanpage PTKX - Followers", "Tháng 6/2026": 3589, "01/07/2026 (01-28/07)": 3598, "Tháng 8/2026": 3900, "Tháng 9/2026": 3900 },
        { "Hạng mục": "Fanpage PTKX - Engagement", "Tháng 6/2026": 10130, "01/07/2026 (01-28/07)": 20274, "Tháng 8/2026": 13000, "Tháng 9/2026": 14000 },
        { "Hạng mục": "Fanpage TTKSCT - Followers", "Tháng 6/2026": 878, "01/07/2026 (01-28/07)": 3229, "Tháng 8/2026": 1500, "Tháng 9/2026": 1800 },
        { "Hạng mục": "Fanpage TTKSCT - Engagement", "Tháng 6/2026": 2309, "01/07/2026 (01-28/07)": 6481, "Tháng 8/2026": 3000, "Tháng 9/2026": 3500 },
        { "Hạng mục": "Google Business - Profile Views", "Tháng 6/2026": 2630, "01/07/2026 (01-28/07)": 2294, "Tháng 8/2026": 12000, "Tháng 9/2026": 13000 },
        { "Hạng mục": "Google Business - Reviews", "Tháng 6/2026": 69, "01/07/2026 (01-28/07)": 80, "Tháng 8/2026": 75, "Tháng 9/2026": 80 },
      ],
    },
    channelsBranding: { columns: ["Hạng mục", "Giá trị"], rows: [] },
    lasik: {
      runrate: 62,
      threshold: 70,
      kpi: { Lead: 200, QLead: 130, Book: 85, Exam: 30, Sur: 27 },
      actual: { Lead: 123, QLead: 80, Book: 40, Exam: 18, Sur: 8 },
      spend: 246888234,
      spendNote: "48%",
      cpa: 12888234,
      cpaNote: "13,000,000",
      creativeImageId: null,
      creativeTitle: "TOP 5 CREATIVE SPEND NHIỀU NHẤT",
    },
    taskList: {
      columns: ["STT", "Hạng mục", "BU", "PIC", "Start date", "Deadline", "Status", "Priority", "Notes"],
      rows: [
        { STT: 1, "Hạng mục": "Tờ rơi lasik", BU: "LASIK", PIC: "Kim", "Start date": "3/8/26", Deadline: "5/8/26", Status: "Completed", Priority: "", Notes: "" },
        { STT: 2, "Hạng mục": "Banner PMAX - Welcome mess T8", BU: "LASIK", PIC: "Kim", "Start date": "1/8/26", Deadline: "31/8/26", Status: "Completed", Priority: "", Notes: "" },
        { STT: 3, "Hạng mục": "POSM lasik tháng 8", BU: "LASIK", PIC: "Kim", "Start date": "1/8/26", Deadline: "31/8/26", Status: "In Progress", Priority: "", Notes: "" },
        { STT: 4, "Hạng mục": "Creative tháng 8", BU: "PHACO", PIC: "Tri", "Start date": "1/8/26", Deadline: "31/8/26", Status: "In Progress", Priority: "", Notes: "17/30" },
        { STT: 5, "Hạng mục": "Vách decal phòng chờ LVC", BU: "HOS", PIC: "Tri", "Start date": "31/7/26", Deadline: "4/8/26", Status: "In Progress", Priority: "", Notes: "Option khác (3/8)" },
        { STT: 6, "Hạng mục": "Creative tháng 8", BU: "PHACO", PIC: "Tri", "Start date": "1/8/26", Deadline: "31/8/26", Status: "Pending Budget", Priority: "", Notes: "Thiết đã được duyệt, chờ ngân sách" },
      ],
    },
    taskImages: { title: "HẠNG MỤC ĐÃ HOÀN THÀNH", images: [] },
    planning: {
      format: "Tất cả",
      bu: "Tất cả",
      items: [{ id: uid(), title: "Tiêu đề", format: "", bu: "PHACO", imageId: null }],
    },
  };
}

/* ============================================================
   SMALL UI PRIMITIVES
============================================================ */
function Card({ children, style }) {
  return <div className="ct-card" style={style}>{children}</div>;
}

function SectionHeader({ title, right }) {
  return (
    <div className="ct-section-header">
      <div className="ct-section-title">{title}</div>
      <div className="ct-section-right">{right}</div>
    </div>
  );
}

function SubTabs({ tabs, active, onChange }) {
  return (
    <div className="ct-subtabs">
      {tabs.map((t) => (
        <button
          key={t}
          className={"ct-subtab" + (active === t ? " active" : "")}
          onClick={() => onChange(t)}
        >
          {t}
        </button>
      ))}
    </div>
  );
}

function EditableText({ value, onChange, placeholder, className, tag }) {
  const Tag = tag || "span";
  return (
    <Tag
      className={"ct-editable " + (className || "")}
      contentEditable
      suppressContentEditableWarning
      onBlur={(e) => onChange(e.currentTarget.textContent)}
      data-placeholder={placeholder || ""}
    >
      {value}
    </Tag>
  );
}

function NumberField({ value, onChange, suffix, width }) {
  return (
    <input
      className="ct-numfield"
      style={width ? { width } : undefined}
      type="number"
      value={value === undefined || value === null ? "" : value}
      onChange={(e) => onChange(e.target.value === "" ? "" : Number(e.target.value))}
    />
  );
}

/* ============================================================
   FORMATTERS
============================================================ */
function fmtNum(v) {
  if (v === null || v === undefined || v === "") return "";
  if (typeof v === "number") return v.toLocaleString("en-US");
  const s = String(v).trim();
  if (/^-?[\d,]+(\.\d+)?$/.test(s)) {
    const n = Number(s.replace(/,/g, ""));
    if (!isNaN(n)) return n.toLocaleString("en-US");
  }
  return v;
}

function isNumericCol(rows, col) {
  return rows.length > 0 && rows.every((r) => {
    const v = r[col];
    return v === "" || v === undefined || v === null || typeof v === "number" ||
      /^-?[\d,]+(\.\d+)?$/.test(String(v).trim());
  }) && rows.some((r) => typeof r[col] === "number");
}

/* ============================================================
   DYNAMIC TABLE  — renders whatever columns/rows exist, no hardcoding
============================================================ */
function DynamicTable({ columns, rows, onCellEdit, editable }) {
  if (!columns || columns.length === 0) {
    return <div className="ct-empty">Chưa có dữ liệu. Hãy Import Excel để bắt đầu.</div>;
  }
  return (
    <div className="ct-table-wrap">
      <table className="ct-table">
        <thead>
          <tr>
            {columns.map((c) => (
              <th key={c}>{c}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {rows.map((row, ri) => (
            <tr key={ri}>
              {columns.map((c) => {
                const raw = row[c];
                const numeric = typeof raw === "number";
                return (
                  <td
                    key={c}
                    className={numeric ? "num" : ""}
                    contentEditable={!!editable}
                    suppressContentEditableWarning
                    onBlur={
                      editable
                        ? (e) => onCellEdit(ri, c, e.currentTarget.textContent)
                        : undefined
                    }
                  >
                    {fmtNum(raw)}
                  </td>
                );
              })}
            </tr>
          ))}
          {rows.length === 0 && (
            <tr><td colSpan={columns.length} className="ct-empty-row">Chưa có dòng dữ liệu.</td></tr>
          )}
        </tbody>
      </table>
    </div>
  );
}

/* ============================================================
   EXCEL PARSING
============================================================ */
function readExcelFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, { type: "array" });
        const sheetName = wb.SheetNames[0];
        const sheet = wb.Sheets[sheetName];
        const raw = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
        if (!raw.length) return resolve({ columns: [], rows: [] });
        const headerRow = raw[0];
        const columns = headerRow.map((h, i) => (h === "" || h === undefined || h === null ? `Cột ${i + 1}` : String(h).trim()));
        const rows = raw.slice(1)
          .filter((r) => r.some((cell) => cell !== "" && cell !== undefined && cell !== null))
          .map((r) => {
            const obj = {};
            columns.forEach((col, i) => {
              let v = r[i];
              if (v === undefined) v = "";
              obj[col] = v;
            });
            return obj;
          });
        resolve({ columns, rows });
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

function readImageFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => resolve(e.target.result);
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

/* fuzzy-match excel stage text to canonical stage key */
function guessStage(text) {
  const s = String(text).toLowerCase();
  if (s.includes("qualified") || s.includes("qlead") || s.includes("q lead")) return "QLead";
  if (s.includes("lead")) return "Lead";
  if (s.includes("book")) return "Book";
  if (s.includes("exam")) return "Exam";
  if (s.includes("surg") || s === "sur" || s.includes("sur")) return "Sur";
  return "";
}

/* ============================================================
   IMPORT MODAL
============================================================ */
function ImportModal({ onClose, onImport }) {
  const [step, setStep] = useState(1);
  const [dataType, setDataType] = useState(DATA_TYPES[0].id);
  const [parsed, setParsed] = useState(null);
  const [fileName, setFileName] = useState("");
  const [error, setError] = useState("");
  const [stageCol, setStageCol] = useState("");
  const [valueCol, setValueCol] = useState("");
  const [stageMap, setStageMap] = useState({});
  const fileInputRef = useRef(null);

  const isLasikType = dataType === "lasikKpi" || dataType === "lasikActual";

  const handleFile = async (file) => {
    setError("");
    setFileName(file.name);
    try {
      const result = await readExcelFile(file);
      setParsed(result);
      if (isLasikType && result.columns.length >= 2) {
        setStageCol(result.columns[0]);
        setValueCol(result.columns[1]);
        const map = {};
        result.rows.forEach((r) => {
          const key = r[result.columns[0]];
          map[key] = guessStage(key);
        });
        setStageMap(map);
      }
      setStep(2);
    } catch (e) {
      setError("Không thể đọc file Excel. Vui lòng kiểm tra định dạng file.");
    }
  };

  const distinctStageValues = parsed && stageCol
    ? Array.from(new Set(parsed.rows.map((r) => r[stageCol])))
    : [];

  const doImport = () => {
    if (isLasikType) {
      const out = { Lead: "", QLead: "", Book: "", Exam: "", Sur: "" };
      parsed.rows.forEach((r) => {
        const raw = r[stageCol];
        const mapped = stageMap[raw];
        if (mapped && STAGES.includes(mapped)) {
          let v = r[valueCol];
          if (typeof v === "string") v = Number(v.replace(/,/g, "")) || 0;
          out[mapped] = v;
        }
      });
      onImport(dataType, out);
    } else {
      onImport(dataType, { columns: parsed.columns, rows: parsed.rows });
    }
    onClose();
  };

  return (
    <div className="ct-modal-overlay" onClick={onClose}>
      <div className="ct-modal" onClick={(e) => e.stopPropagation()}>
        <div className="ct-modal-header">
          <div className="ct-modal-title">IMPORT DATA</div>
          <button className="ct-icon-btn" onClick={onClose}><X size={18} /></button>
        </div>

        {step === 1 && (
          <div className="ct-modal-body">
            <div className="ct-label">Chọn loại dữ liệu</div>
            <div className="ct-radio-list">
              {DATA_TYPES.map((t) => (
                <label key={t.id} className="ct-radio-item">
                  <input
                    type="radio"
                    name="dataType"
                    checked={dataType === t.id}
                    onChange={() => setDataType(t.id)}
                  />
                  <span>{t.label}</span>
                </label>
              ))}
            </div>
            <input
              ref={fileInputRef}
              type="file"
              accept=".xlsx,.xls,.csv"
              style={{ display: "none" }}
              onChange={(e) => e.target.files[0] && handleFile(e.target.files[0])}
            />
            {error && <div className="ct-error"><AlertTriangle size={14}/> {error}</div>}
            <button className="ct-btn ct-btn-primary" onClick={() => fileInputRef.current.click()}>
              <FileSpreadsheet size={16} /> Choose Excel
            </button>
          </div>
        )}

        {step === 2 && parsed && (
          <div className="ct-modal-body">
            <div className="ct-label">File: {fileName} — {parsed.rows.length} dòng, {parsed.columns.length} cột</div>

            {isLasikType && (
              <div className="ct-mapping-box">
                <div className="ct-label">Cấu hình cột</div>
                <div className="ct-mapping-row">
                  <div>
                    <div className="ct-sublabel">Cột tên giai đoạn</div>
                    <select value={stageCol} onChange={(e) => setStageCol(e.target.value)}>
                      {parsed.columns.map((c) => <option key={c} value={c}>{c}</option>)}
                    </select>
                  </div>
                  <div>
                    <div className="ct-sublabel">Cột giá trị</div>
                    <select value={valueCol} onChange={(e) => setValueCol(e.target.value)}>
                      {parsed.columns.map((c) => <option key={c} value={c}>{c}</option>)}
                    </select>
                  </div>
                </div>
                <div className="ct-label" style={{ marginTop: 12 }}>Column Mapping</div>
                {distinctStageValues.map((val) => (
                  <div className="ct-mapping-row" key={String(val)}>
                    <div className="ct-map-source">{String(val)}</div>
                    <ArrowRight size={14} color="#9AA3B2" />
                    <select
                      value={stageMap[val] || ""}
                      onChange={(e) => setStageMap({ ...stageMap, [val]: e.target.value })}
                    >
                      <option value="">— Bỏ qua —</option>
                      {STAGES.map((s) => <option key={s} value={s}>{s}</option>)}
                    </select>
                  </div>
                ))}
              </div>
            )}

            <div className="ct-label" style={{ marginTop: 12 }}>Preview</div>
            <div className="ct-preview-wrap">
              <DynamicTable columns={parsed.columns} rows={parsed.rows.slice(0, 8)} />
            </div>

            <div className="ct-modal-actions">
              <button className="ct-btn" onClick={() => setStep(1)}><ArrowLeft size={14} /> Quay lại</button>
              <button className="ct-btn ct-btn-primary" onClick={doImport}><Check size={14} /> Import</button>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

/* ============================================================
   IMAGE UPLOADER (single slot)
============================================================ */
function SingleImageSlot({ src, onUpload, onDelete, emptyLabel }) {
  const inputRef = useRef(null);
  const handleFile = async (file) => {
    const dataUrl = await readImageFile(file);
    onUpload(dataUrl);
  };
  return (
    <div className="ct-image-slot">
      <input
        ref={inputRef}
        type="file"
        accept="image/png,image/jpeg,image/jpg,image/webp"
        style={{ display: "none" }}
        onChange={(e) => e.target.files[0] && handleFile(e.target.files[0])}
      />
      {src ? (
        <div className="ct-image-filled">
          <img src={src} alt="" style={{ width: "100%", height: "auto", display: "block" }} />
          <div className="ct-image-actions">
            <button className="ct-btn ct-btn-sm" onClick={() => inputRef.current.click()}>Replace</button>
            <button className="ct-btn ct-btn-sm ct-btn-danger" onClick={onDelete}>Delete</button>
          </div>
        </div>
      ) : (
        <button className="ct-image-empty" onClick={() => inputRef.current.click()}>
          <Plus size={26} />
          {emptyLabel && <div className="ct-image-empty-label">{emptyLabel}</div>}
        </button>
      )}
    </div>
  );
}

/* ============================================================
   MULTI IMAGE LIST (reorderable, add/replace/delete)
============================================================ */
function MultiImageList({ images, onAdd, onReplace, onDelete, onReorder }) {
  const addInputRef = useRef(null);
  const replaceInputRef = useRef(null);
  const dragIndex = useRef(null);
  const [replaceId, setReplaceId] = useState(null);

  const handleAdd = async (files) => {
    for (const file of Array.from(files)) {
      const dataUrl = await readImageFile(file);
      onAdd(dataUrl);
    }
  };
  const handleReplace = async (file) => {
    const dataUrl = await readImageFile(file);
    onReplace(replaceId, dataUrl);
  };

  return (
    <div>
      <input
        ref={addInputRef}
        type="file"
        multiple
        accept="image/png,image/jpeg,image/jpg,image/webp"
        style={{ display: "none" }}
        onChange={(e) => { if (e.target.files.length) handleAdd(e.target.files); e.target.value = ""; }}
      />
      <input
        ref={replaceInputRef}
        type="file"
        accept="image/png,image/jpeg,image/jpg,image/webp"
        style={{ display: "none" }}
        onChange={(e) => { if (e.target.files[0]) handleReplace(e.target.files[0]); e.target.value = ""; }}
      />
      <div className="ct-multi-image-grid">
        {images.map((img, idx) => (
          <div
            key={img.id}
            className="ct-multi-image-item"
            draggable
            onDragStart={() => (dragIndex.current = idx)}
            onDragOver={(e) => e.preventDefault()}
            onDrop={() => { onReorder(dragIndex.current, idx); dragIndex.current = null; }}
          >
            <div className="ct-drag-handle"><GripVertical size={14} /></div>
            <img src={img.src} alt="" style={{ width: "100%", height: "auto", display: "block", objectFit: "contain" }} />
            <div className="ct-image-actions">
              <button className="ct-btn ct-btn-sm" onClick={() => { setReplaceId(img.id); replaceInputRef.current.click(); }}>Replace</button>
              <button className="ct-btn ct-btn-sm ct-btn-danger" onClick={() => onDelete(img.id)}>Delete</button>
            </div>
          </div>
        ))}
        <button className="ct-image-empty ct-image-add" onClick={() => addInputRef.current.click()}>
          <Plus size={26} />
          <div className="ct-image-empty-label">ADD IMAGE</div>
        </button>
      </div>
    </div>
  );
}

/* ============================================================
   FUNNEL
============================================================ */
function LasikFunnel({ kpi, actual }) {
  const leadActual = Number(actual.Lead) || 0;
  const maxVal = Math.max(leadActual, 1);
  return (
    <div className="ct-funnel">
      {STAGES.map((stage, i) => {
        const prevStage = STAGES[i - 1];
        const val = Number(actual[stage]) || 0;
        const prevVal = prevStage ? Number(actual[prevStage]) || 0 : 0;
        const pct = i === 0 ? null : (prevVal > 0 ? Math.round((val / prevVal) * 100) : 0);
        const widthPct = Math.max(18, Math.round((val / maxVal) * 100));
        return (
          <div className="ct-funnel-row" key={stage}>
            <div className="ct-funnel-labels">
              <div className="ct-funnel-stage">{STAGE_LABELS[stage]}</div>
              <div className="ct-funnel-kpi">{fmtNum(kpi[stage])}</div>
            </div>
            <div className="ct-funnel-bar-track">
              <div
                className="ct-funnel-bar"
                style={{ width: widthPct + "%", background: STAGE_COLORS[i] }}
              >
                {fmtNum(val)}{pct !== null ? ` (${pct}%)` : ""}
              </div>
            </div>
          </div>
        );
      })}
    </div>
  );
}

/* ============================================================
   NEW REPORT / DUPLICATE MODAL
============================================================ */
function NewReportModal({ existingIds, onCreate, onClose }) {
  const [label, setLabel] = useState("");
  const [timeRange, setTimeRange] = useState("");
  const [conflict, setConflict] = useState(false);

  const submit = () => {
    if (!label.trim()) return;
    if (existingIds.includes(label.trim())) {
      setConflict(true);
      return;
    }
    onCreate(label.trim(), timeRange);
  };

  return (
    <div className="ct-modal-overlay" onClick={onClose}>
      <div className="ct-modal ct-modal-sm" onClick={(e) => e.stopPropagation()}>
        <div className="ct-modal-header">
          <div className="ct-modal-title">NEW REPORT</div>
          <button className="ct-icon-btn" onClick={onClose}><X size={18} /></button>
        </div>
        <div className="ct-modal-body">
          {!conflict ? (
            <>
              <div className="ct-label">Report ID (ví dụ: W33)</div>
              <input className="ct-textfield" value={label} onChange={(e) => setLabel(e.target.value)} placeholder="W33" />
              <div className="ct-label" style={{ marginTop: 12 }}>Time Range</div>
              <input className="ct-textfield" value={timeRange} onChange={(e) => setTimeRange(e.target.value)} placeholder="08/08/2026 - 14/08/2026" />
              <div className="ct-modal-actions">
                <button className="ct-btn" onClick={onClose}>Cancel</button>
                <button className="ct-btn ct-btn-primary" onClick={submit}>Create</button>
              </div>
            </>
          ) : (
            <>
              <div className="ct-error"><AlertTriangle size={14} /> "{label}" already exists.</div>
              <div className="ct-modal-actions" style={{ flexDirection: "column", gap: 8 }}>
                <button className="ct-btn ct-btn-primary" style={{ width: "100%" }} onClick={() => onCreate(label.trim(), timeRange, "update")}>UPDATE EXISTING</button>
                <button className="ct-btn" style={{ width: "100%" }} onClick={() => onCreate(label.trim() + "-v2", timeRange, "new")}>CREATE NEW VERSION</button>
                <button className="ct-btn" style={{ width: "100%" }} onClick={onClose}>CANCEL</button>
              </div>
            </>
          )}
        </div>
      </div>
    </div>
  );
}

/* ============================================================
   MAIN APP
============================================================ */
export default function App() {
  const [loading, setLoading] = useState(true);
  const [reportIndex, setReportIndex] = useState({ ids: [], currentId: null });
  const [report, setReport] = useState(null);
  const [images, setImages] = useState({}); // id -> dataURL
  const [activeTab, setActiveTab] = useState("Digital");
  const [budgetSubTab, setBudgetSubTab] = useState("TOTAL");
  const [channelSubTab, setChannelSubTab] = useState("SUMMARY");
  const [showImport, setShowImport] = useState(false);
  const [showNewReport, setShowNewReport] = useState(false);
  const [showHistory, setShowHistory] = useState(false);
  const [taskFilters, setTaskFilters] = useState({ BU: "Tất cả", PIC: "Tất cả", Status: "Tất cả", Priority: "Tất cả" });
  const [taskSearch, setTaskSearch] = useState("");
  const [saving, setSaving] = useState(false);

  /* ---------- load on mount ---------- */
  useEffect(() => {
    (async () => {
      let idx = await storeGet("report-index");
      if (!idx || !idx.ids || idx.ids.length === 0) {
        const first = sampleReport("W32");
        await storeSet("report:W32", first);
        idx = { ids: ["W32"], currentId: "W32" };
        await storeSet("report-index", idx);
      }
      setReportIndex(idx);
      await loadReport(idx.currentId || idx.ids[idx.ids.length - 1]);
      setLoading(false);
    })();
    // eslint-disable-next-line
  }, []);

  async function loadReport(id) {
    let rep = await storeGet("report:" + id);
    if (!rep) rep = sampleReport(id);
    setReport(rep);
    // collect referenced image ids
    const ids = new Set();
    if (rep.lasik?.creativeImageId) ids.add(rep.lasik.creativeImageId);
    (rep.taskImages?.images || []).forEach((im) => ids.add(im.id));
    (rep.planning?.items || []).forEach((it) => it.imageId && ids.add(it.imageId));
    const imgMap = {};
    for (const iid of ids) {
      const val = await storeGet("image:" + iid);
      if (val) imgMap[iid] = val;
    }
    setImages(imgMap);
  }

  const persist = useCallback(async (nextReport) => {
    setSaving(true);
    await storeSet("report:" + nextReport.id, nextReport);
    setSaving(false);
  }, []);

  const updateReport = useCallback((updater) => {
    setReport((prev) => {
      const next = typeof updater === "function" ? updater(prev) : { ...prev, ...updater };
      persist(next);
      return next;
    });
  }, [persist]);

  async function saveImage(dataUrl) {
    const id = uid();
    await storeSet("image:" + id, dataUrl);
    setImages((prev) => ({ ...prev, [id]: dataUrl }));
    return id;
  }
  async function removeImage(id) {
    await storeDelete("image:" + id);
    setImages((prev) => { const n = { ...prev }; delete n[id]; return n; });
  }

  /* ---------- report switching / creation ---------- */
  async function switchReport(id) {
    const next = { ...reportIndex, currentId: id };
    setReportIndex(next);
    await storeSet("report-index", next);
    await loadReport(id);
    setShowHistory(false);
  }

  async function createReport(label, timeRange, mode) {
    let id = label;
    if (mode === "update") {
      id = label;
    }
    let rep = await storeGet("report:" + id);
    if (!rep || mode !== "update") {
      rep = sampleReport(id);
      rep.timeRange = timeRange || rep.timeRange;
      // fresh report starts empty except structure
      rep.budgetTotal = { columns: report?.budgetTotal?.columns || [], rows: [] };
      rep.budgetDetail = { columns: [], rows: [] };
      rep.channelsSummary = { columns: [], rows: [] };
      rep.channelsBranding = { columns: [], rows: [] };
      rep.lasik = { runrate: "", threshold: "", kpi: { Lead: "", QLead: "", Book: "", Exam: "", Sur: "" }, actual: { Lead: "", QLead: "", Book: "", Exam: "", Sur: "" }, spend: "", spendNote: "", cpa: "", cpaNote: "", creativeImageId: null, creativeTitle: "TOP 5 CREATIVE SPEND NHIỀU NHẤT" };
      rep.taskList = { columns: [], rows: [] };
      rep.taskImages = { title: "Tiêu đề", images: [] };
      rep.planning = { format: "Tất cả", bu: "Tất cả", items: [] };
    }
    await storeSet("report:" + id, rep);
    let ids = reportIndex.ids;
    if (!ids.includes(id)) ids = [...ids, id];
    const nextIdx = { ids, currentId: id };
    setReportIndex(nextIdx);
    await storeSet("report-index", nextIdx);
    await loadReport(id);
    setShowNewReport(false);
  }

  /* ---------- import handler ---------- */
  function handleImport(dataType, payload) {
    updateReport((prev) => {
      const next = { ...prev };
      if (dataType === "budgetTotal") next.budgetTotal = payload;
      else if (dataType === "budgetDetail") next.budgetDetail = payload;
      else if (dataType === "channelsSummary") next.channelsSummary = payload;
      else if (dataType === "channelsBranding") next.channelsBranding = payload;
      else if (dataType === "taskList") next.taskList = payload;
      else if (dataType === "lasikKpi") next.lasik = { ...prev.lasik, kpi: payload };
      else if (dataType === "lasikActual") next.lasik = { ...prev.lasik, actual: payload };
      return next;
    });
  }

  /* ---------- table cell edit (manual entry) ---------- */
  function editCell(section, rowIndex, col, value) {
    updateReport((prev) => {
      const sec = prev[section];
      const rows = sec.rows.map((r, i) => {
        if (i !== rowIndex) return r;
        let v = value;
        if (/^-?[\d,]+(\.\d+)?$/.test(String(value).trim()) && String(value).trim() !== "") {
          v = Number(String(value).replace(/,/g, ""));
        }
        return { ...r, [col]: v };
      });
      return { ...prev, [section]: { ...sec, rows } };
    });
  }

  /* ---------- PDF ---------- */
  function downloadPDF() {
    window.print();
  }

  if (loading || !report) {
    return (
      <div className="ct-root ct-loading">
        <div className="ct-spinner" />
        <div>Đang tải báo cáo...</div>
      </div>
    );
  }

  const taskCols = report.taskList.columns;
  const findCol = (name) => taskCols.find((c) => c.toLowerCase().includes(name.toLowerCase()));
  const buCol = findCol("bu");
  const picCol = findCol("pic");
  const statusCol = findCol("status");
  const priorityCol = findCol("priority");
  const distinctVals = (col) => col ? Array.from(new Set(report.taskList.rows.map((r) => r[col]).filter(Boolean))) : [];

  const filteredTaskRows = report.taskList.rows.filter((r) => {
    if (buCol && taskFilters.BU !== "Tất cả" && r[buCol] !== taskFilters.BU) return false;
    if (picCol && taskFilters.PIC !== "Tất cả" && r[picCol] !== taskFilters.PIC) return false;
    if (statusCol && taskFilters.Status !== "Tất cả" && r[statusCol] !== taskFilters.Status) return false;
    if (priorityCol && taskFilters.Priority !== "Tất cả" && r[priorityCol] !== taskFilters.Priority) return false;
    if (taskSearch.trim()) {
      const hay = Object.values(r).join(" ").toLowerCase();
      if (!hay.includes(taskSearch.trim().toLowerCase())) return false;
    }
    return true;
  });

  const planningFormats = Array.from(new Set(report.planning.items.map((i) => i.format).filter(Boolean)));
  const planningBUs = Array.from(new Set(report.planning.items.map((i) => i.bu).filter(Boolean)));
  const filteredPlanning = report.planning.items.filter((it) => {
    if (report.planning.format !== "Tất cả" && it.format !== report.planning.format) return false;
    if (report.planning.bu !== "Tất cả" && it.bu !== report.planning.bu) return false;
    return true;
  });

  return (
    <div className="ct-root">
      {/* ================= HEADER ================= */}
      <div className="ct-header">
        <div className="ct-brand">
          <div className="ct-logo-mark">
            <svg width="34" height="34" viewBox="0 0 34 34">
              <circle cx="17" cy="17" r="16" fill="none" stroke="#1B4FCE" strokeWidth="2" />
              <circle cx="17" cy="17" r="4.5" fill="#1B4FCE" />
              {[0, 60, 120, 180, 240, 300].map((deg) => (
                <ellipse key={deg} cx="17" cy="7" rx="3.2" ry="6" fill="#1B4FCE" opacity="0.75"
                  transform={`rotate(${deg} 17 17)`} />
              ))}
            </svg>
          </div>
          <div className="ct-brand-text">
            <div className="ct-brand-name">CAO THẮNG</div>
            <div className="ct-brand-sub">MARKETING REPORT</div>
          </div>
        </div>
        <div className="ct-nav">
          {["Digital", "Task list", "Planning"].map((t) => (
            <button
              key={t}
              className={"ct-nav-btn" + (activeTab === t ? " active" : "")}
              onClick={() => setActiveTab(t)}
            >
              {t}
            </button>
          ))}
        </div>
      </div>

      {/* ================= TOOLBAR ================= */}
      <div className="ct-toolbar">
        <div className="ct-toolbar-left">
          <span className="ct-toolbar-label">Timerange:</span>
          <input
            className="ct-textfield ct-timerange"
            value={report.timeRange}
            onChange={(e) => updateReport({ timeRange: e.target.value })}
            placeholder="01/08/2026 - 07/08/2026"
          />
          <div className="ct-history-wrap">
            <button className="ct-btn" onClick={() => setShowHistory((s) => !s)}>
              <History size={14} /> {report.id} <ChevronDown size={14} />
            </button>
            {showHistory && (
              <div className="ct-history-dropdown">
                {reportIndex.ids.slice().reverse().map((id) => (
                  <button key={id} className={"ct-history-item" + (id === report.id ? " active" : "")} onClick={() => switchReport(id)}>
                    {id}{id === report.id && <Check size={13} />}
                  </button>
                ))}
                <button className="ct-history-item ct-history-new" onClick={() => { setShowHistory(false); setShowNewReport(true); }}>
                  <Plus size={13} /> New Report
                </button>
              </div>
            )}
          </div>
          {saving && <span className="ct-saving">Saving…</span>}
        </div>
        <div className="ct-toolbar-right">
          <button className="ct-btn ct-btn-primary" onClick={downloadPDF}><Download size={14} /> Download PDF</button>
          <button className="ct-btn" onClick={() => setShowImport(true)}><Upload size={14} /> Import Data</button>
        </div>
      </div>

      {/* ================= DIGITAL TAB ================= */}
      {activeTab === "Digital" && (
        <div className="ct-page">
          <Card>
            <SectionHeader
              title="CẬP NHẬT NGÂN SÁCH VÀ CHI PHÍ"
              right={<SubTabs tabs={["TOTAL", "DETAIL"]} active={budgetSubTab} onChange={setBudgetSubTab} />}
            />
            {budgetSubTab === "TOTAL" ? (
              <DynamicTable
                columns={report.budgetTotal.columns}
                rows={report.budgetTotal.rows}
                editable
                onCellEdit={(ri, c, v) => editCell("budgetTotal", ri, c, v)}
              />
            ) : (
              <DynamicTable
                columns={report.budgetDetail.columns}
                rows={report.budgetDetail.rows}
                editable
                onCellEdit={(ri, c, v) => editCell("budgetDetail", ri, c, v)}
              />
            )}
          </Card>

          <Card>
            <SectionHeader
              title="THỐNG KÊ CÁC KÊNH DIGITAL"
              right={<SubTabs tabs={["SUMMARY", "BRANDING ADS"]} active={channelSubTab === "SUMMARY" ? "SUMMARY" : "BRANDING ADS"} onChange={(t) => setChannelSubTab(t)} />}
            />
            {channelSubTab === "SUMMARY" ? (
              <DynamicTable
                columns={report.channelsSummary.columns}
                rows={report.channelsSummary.rows}
                editable
                onCellEdit={(ri, c, v) => editCell("channelsSummary", ri, c, v)}
              />
            ) : (
              <DynamicTable
                columns={report.channelsBranding.columns}
                rows={report.channelsBranding.rows}
                editable
                onCellEdit={(ri, c, v) => editCell("channelsBranding", ri, c, v)}
              />
            )}
          </Card>

          <Card>
            <div className="ct-lasik-title">LASIK</div>
            <div className="ct-lasik-runrate">
              THIS MONTH (RUNRATE{" "}
              <input
                className="ct-inline-num"
                type="number"
                value={report.lasik.runrate}
                onChange={(e) => updateReport((p) => ({ ...p, lasik: { ...p.lasik, runrate: e.target.value } }))}
              />
              %)
              <span className="ct-threshold">
                Threshold:
                <input
                  className="ct-inline-num"
                  type="number"
                  value={report.lasik.threshold}
                  onChange={(e) => updateReport((p) => ({ ...p, lasik: { ...p.lasik, threshold: e.target.value } }))}
                />
                %
              </span>
            </div>

            <div className="ct-lasik-grid">
              <div className="ct-lasik-left">
                <LasikFunnel kpi={report.lasik.kpi} actual={report.lasik.actual} />
                <div className="ct-lasik-kpi-edit">
                  <div className="ct-sublabel">Chỉnh KPI Target / Actual thủ công</div>
                  <table className="ct-mini-table">
                    <thead><tr><th></th>{STAGES.map((s) => <th key={s}>{s}</th>)}</tr></thead>
                    <tbody>
                      <tr>
                        <td>KPI</td>
                        {STAGES.map((s) => (
                          <td key={s}>
                            <NumberField
                              value={report.lasik.kpi[s]}
                              onChange={(v) => updateReport((p) => ({ ...p, lasik: { ...p.lasik, kpi: { ...p.lasik.kpi, [s]: v } } }))}
                              width={62}
                            />
                          </td>
                        ))}
                      </tr>
                      <tr>
                        <td>Actual</td>
                        {STAGES.map((s) => (
                          <td key={s}>
                            <NumberField
                              value={report.lasik.actual[s]}
                              onChange={(v) => updateReport((p) => ({ ...p, lasik: { ...p.lasik, actual: { ...p.lasik.actual, [s]: v } } }))}
                              width={62}
                            />
                          </td>
                        ))}
                      </tr>
                    </tbody>
                  </table>
                </div>
                <div className="ct-spend-cards">
                  <div className="ct-spend-card">
                    <div className="ct-spend-label">
                      SPEND (
                      <input
                        className="ct-inline-text"
                        value={report.lasik.spendNote}
                        onChange={(e) => updateReport((p) => ({ ...p, lasik: { ...p.lasik, spendNote: e.target.value } }))}
                      />
                      )
                    </div>
                    <NumberField
                      value={report.lasik.spend}
                      onChange={(v) => updateReport((p) => ({ ...p, lasik: { ...p.lasik, spend: v } }))}
                      width="100%"
                    />
                  </div>
                  <div className="ct-spend-card">
                    <div className="ct-spend-label">
                      CPA (
                      <input
                        className="ct-inline-text"
                        value={report.lasik.cpaNote}
                        onChange={(e) => updateReport((p) => ({ ...p, lasik: { ...p.lasik, cpaNote: e.target.value } }))}
                      />
                      )
                    </div>
                    <NumberField
                      value={report.lasik.cpa}
                      onChange={(v) => updateReport((p) => ({ ...p, lasik: { ...p.lasik, cpa: v } }))}
                      width="100%"
                    />
                  </div>
                </div>
              </div>
              <div className="ct-lasik-right">
                <div className="ct-creative-title">
                  <EditableText
                    value={report.lasik.creativeTitle}
                    onChange={(v) => updateReport((p) => ({ ...p, lasik: { ...p.lasik, creativeTitle: v } }))}
                  />
                </div>
                <SingleImageSlot
                  src={report.lasik.creativeImageId ? images[report.lasik.creativeImageId] : null}
                  emptyLabel="UPLOAD IMAGE"
                  onUpload={async (dataUrl) => {
                    const id = await saveImage(dataUrl);
                    updateReport((p) => ({ ...p, lasik: { ...p.lasik, creativeImageId: id } }));
                  }}
                  onDelete={() => {
                    if (report.lasik.creativeImageId) removeImage(report.lasik.creativeImageId);
                    updateReport((p) => ({ ...p, lasik: { ...p.lasik, creativeImageId: null } }));
                  }}
                />
              </div>
            </div>
          </Card>
        </div>
      )}

      {/* ================= TASK LIST TAB ================= */}
      {activeTab === "Task list" && (
        <div className="ct-page">
          <Card>
            <SectionHeader title="TASK LIST" right={null} />
            <div className="ct-task-controls">
              <div className="ct-search-box">
                <Search size={14} color="#8890A0" />
                <input placeholder="Search task…" value={taskSearch} onChange={(e) => setTaskSearch(e.target.value)} />
              </div>
              {[
                { key: "BU", col: buCol },
                { key: "PIC", col: picCol },
                { key: "Status", col: statusCol },
                { key: "Priority", col: priorityCol },
              ].map(({ key, col }) => col && (
                <select
                  key={key}
                  className="ct-filter-select"
                  value={taskFilters[key]}
                  onChange={(e) => setTaskFilters({ ...taskFilters, [key]: e.target.value })}
                >
                  <option>Tất cả</option>
                  {distinctVals(col).map((v) => <option key={v} value={v}>{v}</option>)}
                </select>
              ))}
            </div>
            <div className="ct-table-wrap">
              <table className="ct-table">
                <thead>
                  <tr>{taskCols.map((c) => <th key={c}>{c}</th>)}</tr>
                </thead>
                <tbody>
                  {filteredTaskRows.map((row, ri) => (
                    <tr key={ri}>
                      {taskCols.map((c) => {
                        const isStatus = c === statusCol;
                        const val = row[c];
                        if (isStatus && val) {
                          const colors = STATUS_COLORS[val] || { bg: "#EEF0F4", fg: "#7C8595" };
                          return (
                            <td key={c}>
                              <span className="ct-badge" style={{ background: colors.bg, color: colors.fg }}>{val}</span>
                            </td>
                          );
                        }
                        return (
                          <td
                            key={c}
                            className={typeof val === "number" ? "num" : ""}
                            contentEditable
                            suppressContentEditableWarning
                            onBlur={(e) => editCell("taskList", report.taskList.rows.indexOf(row), c, e.currentTarget.textContent)}
                          >
                            {fmtNum(val)}
                          </td>
                        );
                      })}
                    </tr>
                  ))}
                  {filteredTaskRows.length === 0 && (
                    <tr><td colSpan={taskCols.length || 1} className="ct-empty-row">Không có task phù hợp.</td></tr>
                  )}
                </tbody>
              </table>
            </div>
          </Card>

          <Card>
            <div className="ct-title-editable">
              <EditableText
                tag="div"
                className="ct-h2"
                value={report.taskImages.title}
                onChange={(v) => updateReport((p) => ({ ...p, taskImages: { ...p.taskImages, title: v } }))}
              />
            </div>
            <MultiImageList
              images={report.taskImages.images.map((im) => ({ id: im.id, src: images[im.id] }))}
              onAdd={async (dataUrl) => {
                const id = await saveImage(dataUrl);
                updateReport((p) => ({ ...p, taskImages: { ...p.taskImages, images: [...p.taskImages.images, { id }] } }));
              }}
              onReplace={async (id, dataUrl) => {
                await storeDelete("image:" + id);
                const newId = await saveImage(dataUrl);
                updateReport((p) => ({
                  ...p,
                  taskImages: { ...p.taskImages, images: p.taskImages.images.map((im) => (im.id === id ? { id: newId } : im)) },
                }));
              }}
              onDelete={(id) => {
                removeImage(id);
                updateReport((p) => ({ ...p, taskImages: { ...p.taskImages, images: p.taskImages.images.filter((im) => im.id !== id) } }));
              }}
              onReorder={(from, to) => {
                updateReport((p) => {
                  const imgs = [...p.taskImages.images];
                  const [moved] = imgs.splice(from, 1);
                  imgs.splice(to, 0, moved);
                  return { ...p, taskImages: { ...p.taskImages, images: imgs } };
                });
              }}
            />
          </Card>
        </div>
      )}

      {/* ================= PLANNING TAB ================= */}
      {activeTab === "Planning" && (
        <div className="ct-page">
          <div className="ct-planning-toolbar">
            <div className="ct-planning-filter">
              <span className="ct-toolbar-label">Format:</span>
              <select value={report.planning.format} onChange={(e) => updateReport((p) => ({ ...p, planning: { ...p.planning, format: e.target.value } }))}>
                <option>Tất cả</option>
                {planningFormats.map((f) => <option key={f} value={f}>{f}</option>)}
              </select>
            </div>
            <div className="ct-planning-filter">
              <span className="ct-toolbar-label">BU:</span>
              <select value={report.planning.bu} onChange={(e) => updateReport((p) => ({ ...p, planning: { ...p.planning, bu: e.target.value } }))}>
                <option>Tất cả</option>
                {planningBUs.map((b) => <option key={b} value={b}>{b}</option>)}
              </select>
            </div>
          </div>

          {filteredPlanning.map((item, idx) => (
            <Card key={item.id}>
              <div className="ct-planning-item-head">
                <EditableText
                  tag="div"
                  className="ct-h2"
                  value={item.title}
                  onChange={(v) => updateReport((p) => ({
                    ...p, planning: { ...p.planning, items: p.planning.items.map((it) => it.id === item.id ? { ...it, title: v } : it) },
                  }))}
                />
                <span className="ct-planning-tag">
                  <input className="ct-tag-input" placeholder="BU" value={item.bu} onChange={(e) => updateReport((p) => ({
                    ...p, planning: { ...p.planning, items: p.planning.items.map((it) => it.id === item.id ? { ...it, bu: e.target.value } : it) },
                  }))} />
                </span>
                <span className="ct-planning-tag">
                  <input className="ct-tag-input" placeholder="Format" value={item.format} onChange={(e) => updateReport((p) => ({
                    ...p, planning: { ...p.planning, items: p.planning.items.map((it) => it.id === item.id ? { ...it, format: e.target.value } : it) },
                  }))} />
                </span>
                <button
                  className="ct-icon-btn ct-danger"
                  onClick={() => {
                    if (item.imageId) removeImage(item.imageId);
                    updateReport((p) => ({ ...p, planning: { ...p.planning, items: p.planning.items.filter((it) => it.id !== item.id) } }));
                  }}
                ><Trash2 size={15} /></button>
              </div>
              <SingleImageSlot
                src={item.imageId ? images[item.imageId] : null}
                emptyLabel="UPLOAD IMAGE"
                onUpload={async (dataUrl) => {
                  const id = await saveImage(dataUrl);
                  updateReport((p) => ({ ...p, planning: { ...p.planning, items: p.planning.items.map((it) => it.id === item.id ? { ...it, imageId: id } : it) } }));
                }}
                onDelete={() => {
                  if (item.imageId) removeImage(item.imageId);
                  updateReport((p) => ({ ...p, planning: { ...p.planning, items: p.planning.items.map((it) => it.id === item.id ? { ...it, imageId: null } : it) } }));
                }}
              />
            </Card>
          ))}

          <button
            className="ct-add-planning-btn"
            onClick={() => updateReport((p) => ({
              ...p, planning: { ...p.planning, items: [...p.planning.items, { id: uid(), title: "Tiêu đề", format: "", bu: "", imageId: null }] },
            }))}
          >
            <Plus size={22} />
          </button>
        </div>
      )}

      {showImport && <ImportModal onClose={() => setShowImport(false)} onImport={handleImport} />}
      {showNewReport && (
        <NewReportModal
          existingIds={reportIndex.ids}
          onClose={() => setShowNewReport(false)}
          onCreate={createReport}
        />
      )}
    </div>
  );
}

