import React, { useState, useCallback } from "react";
import * as XLSX from "xlsx";
import { Upload, FileSpreadsheet, AlertTriangle, X, Plus, RotateCcw, Loader2 } from "lucide-react";
import { ComposedChart, Bar, Line, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer } from "recharts";

const FONT_IMPORT = `@import url('https://fonts.googleapis.com/css2?family=Source+Serif+4:wght@500;600&family=IBM+Plex+Mono:wght@400;500&family=Inter:wght@400;500;600&display=swap');`;

const INK = "#1C2B22";
const PARCHMENT = "#F7F4EC";
const CARD = "#FFFFFF";
const DEED_GREEN = "#3B6E52";
const FLAG_AMBER = "#B8792F";
const FLAG_AMBER_BG = "#FBF0DF";
const SLATE = "#77756A";
const BORDER = "#E4DFD1";

const MONTHS = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];
const MONTH_LABELS = { jan: "Jan", feb: "Feb", mar: "Mar", apr: "Apr", may: "May", jun: "Jun", jul: "Jul", aug: "Aug", sep: "Sep", oct: "Oct", nov: "Nov", dec: "Dec" };

const EXTRA_STYLES = `
@keyframes spin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }
.spin-icon { animation: spin 0.9s linear infinite; }
.flag-cell { position: relative; cursor: help; }
.flag-cell .tip {
  display: none;
  position: absolute;
  top: 100%;
  left: 50%;
  transform: translateX(-50%);
  margin-top: 6px;
  background: ${INK};
  color: #FFFFFF;
  padding: 8px 10px;
  border-radius: 6px;
  font-size: 12px;
  font-family: 'Inter', sans-serif;
  font-weight: 400;
  line-height: 1.45;
  white-space: normal;
  width: max-content;
  max-width: 260px;
  text-align: left;
  z-index: 60;
  box-shadow: 0 6px 16px rgba(0,0,0,0.18);
}
.flag-cell .tip::after {
  content: '';
  position: absolute;
  bottom: 100%;
  left: 50%;
  transform: translateX(-50%);
  border: 5px solid transparent;
  border-bottom-color: ${INK};
}
.flag-cell:hover .tip { display: block; }
.flag-cell .tip div + div { margin-top: 5px; }
`;

function norm(v) {
  return String(v == null ? "" : v).toLowerCase().replace(/\s+/g, " ").trim();
}
function parseNumber(v) {
  if (typeof v === "number") return v;
  if (typeof v === "string") {
    const cleaned = v.replace(/[$,\s]/g, "");
    if (cleaned === "" || cleaned === "-") return NaN;
    const n = parseFloat(cleaned);
    return isNaN(n) ? NaN : n;
  }
  return NaN;
}
function formatMoney(v) {
  if (v === undefined || isNaN(v)) return "\u2014";
  const abs = Math.round(Math.abs(v)).toLocaleString();
  return v < 0 ? `(${abs})` : abs;
}
function firstNonEmpty(row) {
  if (!row) return "";
  for (const c of row) if (norm(c) !== "") return c;
  return "";
}
function findReportSheet(workbook) {
  for (const name of workbook.SheetNames) {
    const grid = XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1, defval: "" });
    for (let r = 0; r < Math.min(grid.length, 15); r++) {
      for (const cell of grid[r]) {
        if (norm(cell) === "monthly status report") return { name, grid };
      }
    }
  }
  const outputSheet = workbook.SheetNames.find((n) => n.toLowerCase().includes("output"));
  const chosen = outputSheet || workbook.SheetNames[0];
  return { name: chosen, grid: XLSX.utils.sheet_to_json(workbook.Sheets[chosen], { header: 1, defval: "" }) };
}
function findMonthHeaderRow(grid) {
  let best = { rowIdx: -1, count: 0, colOfMonth: {} };
  grid.forEach((row, rowIdx) => {
    const colOfMonth = {};
    row.forEach((cell, colIdx) => {
      const c = norm(cell);
      for (const m of MONTHS) {
        if ((c === m || (c.startsWith(m) && c.length <= 9)) && !(m in colOfMonth)) colOfMonth[m] = colIdx;
      }
    });
    const count = Object.keys(colOfMonth).length;
    if (count > best.count) best = { rowIdx, count, colOfMonth };
  });
  let ytdCol = null;
  if (best.rowIdx >= 0) {
    grid[best.rowIdx].forEach((cell, colIdx) => { if (norm(cell) === "ytd" && ytdCol === null) ytdCol = colIdx; });
  }
  return { ...best, ytdCol };
}
function buildLabelIndex(grid) {
  const rows = [];
  grid.forEach((row, rowIdx) => {
    const label = norm(firstNonEmpty(row.slice(0, 4)));
    if (label) rows.push({ rowIdx, label });
  });
  return rows;
}
function findRow(labelRows, matcher) {
  const hit = labelRows.find((r) => matcher(r.label));
  return hit ? hit.rowIdx : null;
}
function findYearedRow(labelRows, regex) {
  const matches = labelRows.map((r) => ({ ...r, m: r.label.match(regex) })).filter((r) => r.m);
  if (matches.length === 0) return null;
  matches.sort((a, b) => parseInt(b.m[1], 10) - parseInt(a.m[1], 10));
  return matches[0].rowIdx;
}
function extractRowValues(grid, rowIdx, colOfMonth, ytdCol) {
  if (rowIdx === null) return { byMonth: {}, ytd: NaN };
  const row = grid[rowIdx];
  const byMonth = {};
  MONTHS.forEach((m) => { if (colOfMonth[m] !== undefined) byMonth[m] = parseNumber(row[colOfMonth[m]]); });
  const ytd = ytdCol !== null ? parseNumber(row[ytdCol]) : NaN;
  return { byMonth, ytd };
}

function processWorkbook(workbook, fileName) {
  const { name: sheetName, grid } = findReportSheet(workbook);
  const statusRowIdx = grid.findIndex((row) => row.some((c) => norm(c) === "monthly status report"));
  const propertyName = statusRowIdx >= 2 ? firstNonEmpty(grid[statusRowIdx - 2]) : "";
  const location = statusRowIdx >= 1 ? firstNonEmpty(grid[statusRowIdx - 1]) : "";
  const reportDateRaw = statusRowIdx >= 0 ? firstNonEmpty(grid[statusRowIdx + 1]) : "";
  const reportDate = reportDateRaw instanceof Date ? reportDateRaw : (reportDateRaw ? new Date(reportDateRaw) : null);

  const labelRows = buildLabelIndex(grid);
  const unitsRowIdx = findRow(labelRows, (l) => l === "number of units");
  const sqftRowIdx = findRow(labelRows, (l) => l === "square footage");
  let units = null, sqft = null, purchaseDate = null, purchasePrice = null;
  if (unitsRowIdx !== null) units = parseNumber(grid[unitsRowIdx].find((c) => typeof c === "number"));
  if (sqftRowIdx !== null) sqft = parseNumber(grid[sqftRowIdx].find((c) => typeof c === "number"));
  grid.slice(0, 10).forEach((row) => {
    row.forEach((cell) => {
      const s = String(cell);
      const pd = s.match(/purchase date:\s*(.+)/i);
      const pp = s.match(/purchase price:\s*(.+)/i);
      if (pd) purchaseDate = pd[1].trim();
      if (pp) purchasePrice = pp[1].trim();
    });
  });

  const { colOfMonth, ytdCol } = findMonthHeaderRow(grid);
  const rowIdx = {
    income: findRow(labelRows, (l) => l === "income"),
    operatingExpenses: findRow(labelRows, (l) => l === "operating expenses"),
    recurringCapex: findRow(labelRows, (l) => l === "recurring capital expenses"),
    ownershipExpenses: findRow(labelRows, (l) => l === "ownership expenses"),
    noi: findRow(labelRows, (l) => l.startsWith("net operating income")),
    nonRecurringCapex: findRow(labelRows, (l) => l === "non-recurring capital expenses"),
    debtService: findRow(labelRows, (l) => l === "debt service"),
    netIncome: findRow(labelRows, (l) => l.startsWith("net income")),
    occupiedPct: findYearedRow(labelRows, /^(\d{4}) occupied %$/),
    moveOuts: findYearedRow(labelRows, /^(\d{4}) move-?outs$/),
    currentVacantUnits: findRow(labelRows, (l) => l === "current vacant units"),
    netExposureVacancy: findRow(labelRows, (l) => l === "net exposure to vacancy"),
  };

  const extracted = {};
  Object.entries(rowIdx).forEach(([key, idx]) => { extracted[key] = extractRowValues(grid, idx, colOfMonth, ytdCol); });

  const monthsPresent = MONTHS.filter((m) => colOfMonth[m] !== undefined);
  const months = monthsPresent.map((m) => {
    const get = (k) => extracted[k].byMonth[m];
    const income = get("income"), opEx = get("operatingExpenses"), recurCapex = get("recurringCapex"), ownershipExp = get("ownershipExpenses");
    const reportedNOI = get("noi"), nonRecurCapex = get("nonRecurringCapex"), debtService = get("debtService"), reportedNetIncome = get("netIncome");
    const occupiedPct = get("occupiedPct"), moveOuts = get("moveOuts"), vacantUnits = get("currentVacantUnits"), netExposure = get("netExposureVacancy");

    // Some templates zero-fill months that haven't happened yet (rather than leaving them blank).
    // Treat a month as "not yet reported" when every core financial figure is exactly zero/blank
    // and occupancy shows no real value either.
    const zeroOrNaN = (v) => v === 0 || isNaN(v);
    const notYetReported = zeroOrNaN(income) && zeroOrNaN(opEx) && zeroOrNaN(reportedNOI) && zeroOrNaN(reportedNetIncome) && isNaN(occupiedPct);

    const flags = [];
    if (!notYetReported) {
      const hasAllNoiInputs = [income, opEx, recurCapex, ownershipExp, reportedNOI].every((v) => !isNaN(v));
      if (hasAllNoiInputs) {
        const computedNOI = income + opEx + recurCapex + ownershipExp;
        if (Math.abs(computedNOI - reportedNOI) > 5) {
          flags.push({ field: "noi", message: `NOI doesn't reconcile: computes to ${formatMoney(computedNOI)}, report shows ${formatMoney(reportedNOI)}` });
        }
      }
      if (!isNaN(reportedNOI) && !isNaN(nonRecurCapex) && !isNaN(debtService) && !isNaN(reportedNetIncome)) {
        const computedNetIncome = reportedNOI + nonRecurCapex + debtService;
        if (Math.abs(computedNetIncome - reportedNetIncome) > 5) {
          flags.push({ field: "netIncome", message: `Net income doesn't reconcile: computes to ${formatMoney(computedNetIncome)}, report shows ${formatMoney(reportedNetIncome)}` });
        }
      }
      if (!isNaN(occupiedPct) && occupiedPct < 0.9) flags.push({ field: "occupiedPct", message: `Occupancy at ${(occupiedPct * 100).toFixed(1)}%, below 90%` });
      if (isNaN(income)) flags.push({ field: "income", message: "No income figure reported for this month" });
    }

    return { key: m, label: MONTH_LABELS[m], income, opEx, recurCapex, ownershipExp, noi: reportedNOI, nonRecurCapex, debtService, netIncome: reportedNetIncome, occupiedPct, moveOuts, vacantUnits, netExposure, flags, notYetReported };
  });

  const reportedMonths = months.filter((mo) => !mo.notYetReported);
  const validIncomes = reportedMonths.map((mo) => mo.income).filter((v) => !isNaN(v) && v > 0).sort((a, b) => a - b);
  const medIncome = validIncomes.length ? validIncomes[Math.floor(validIncomes.length / 2)] : 0;

  const gaps = reportedMonths.map((mo) => (!isNaN(mo.noi) && !isNaN(mo.netIncome)) ? Math.abs(mo.noi - mo.netIncome) : null).filter((v) => v !== null).sort((a, b) => a - b);
  const medAbsGap = gaps.length ? gaps[Math.floor(gaps.length / 2)] : 0;
  const gapThreshold = Math.max(medAbsGap * 2.5, 5000);

  reportedMonths.forEach((mo) => {
    if (!isNaN(mo.income) && medIncome > 0 && (mo.income > medIncome * 2.5 || mo.income < medIncome * 0.4)) {
      mo.flags.push({ field: "income", message: `Income (${formatMoney(mo.income)}) is far from the typical month (median ${formatMoney(medIncome)})` });
    }
    if (!isNaN(mo.netExposure) && units && units > 0) {
      const ratio = mo.netExposure / units;
      if (ratio > 0.05) {
        mo.flags.push({ field: "netExposure", message: `Net exposure to vacancy at ${mo.netExposure} units (${(ratio * 100).toFixed(1)}% of ${units} units)` });
      }
    } else if (!isNaN(mo.netExposure) && mo.netExposure >= 2) {
      mo.flags.push({ field: "netExposure", message: `Net exposure to vacancy at ${mo.netExposure} units` });
    }
    if (!isNaN(mo.noi) && !isNaN(mo.netIncome)) {
      const gap = Math.abs(mo.noi - mo.netIncome);
      if (gap > gapThreshold) {
        mo.flags.push({ field: "netIncome", message: `Gap between NOI (${formatMoney(mo.noi)}) and net income (${formatMoney(mo.netIncome)}) is unusually large for this property \u2014 check capex/debt service this month` });
      }
    }
  });

  const year = reportDate && !isNaN(reportDate) ? reportDate.getFullYear() : parseInt(((fileName || "").match(/(20\d{2})/) || [])[1] || "0", 10) || null;

  const ytdIncome = !isNaN(extracted.income.ytd) ? extracted.income.ytd : reportedMonths.reduce((s, mo) => s + (isNaN(mo.income) ? 0 : mo.income), 0);
  const ytdNetIncome = !isNaN(extracted.netIncome.ytd) ? extracted.netIncome.ytd : reportedMonths.reduce((s, mo) => s + (isNaN(mo.netIncome) ? 0 : mo.netIncome), 0);
  const occValues = reportedMonths.map((mo) => mo.occupiedPct).filter((v) => !isNaN(v));
  const avgOccupancy = occValues.length ? occValues.reduce((a, b) => a + b, 0) / occValues.length : NaN;
  const moveOutValues = reportedMonths.map((mo) => mo.moveOuts).filter((v) => !isNaN(v));
  const totalMoveOuts = moveOutValues.length ? moveOutValues.reduce((a, b) => a + b, 0) : NaN;
  const flaggedMonthCount = reportedMonths.filter((mo) => mo.flags.length > 0).length;

  let period = String(year || "");
  if (reportedMonths.length > 0 && reportedMonths.length < 12) {
    period = `${reportedMonths[0].label}\u2013${reportedMonths[reportedMonths.length - 1].label} ${year || ""}`.trim();
  }

  return {
    sheetName, propertyName, location, reportDate, units, sqft, purchaseDate, purchasePrice, months, year, period,
    ytdIncome, ytdNetIncome, avgOccupancy, totalMoveOuts, flaggedMonthCount, reportedMonthCount: reportedMonths.length,
  };
}

const FIN_ROWS = [
  { key: "income", label: "Income" },
  { key: "opEx", label: "Operating expenses" },
  { key: "recurCapex", label: "Recurring capital expenses" },
  { key: "ownershipExp", label: "Ownership expenses" },
  { key: "noi", label: "Net operating income" },
  { key: "nonRecurCapex", label: "Non-recurring capital expenses" },
  { key: "debtService", label: "Debt service" },
  { key: "netIncome", label: "Net income" },
];

const OCC_ROWS = [
  { key: "occupiedPct", label: "Occupied %", isPct: true },
  { key: "moveOuts", label: "Move-outs" },
  { key: "vacantUnits", label: "Current vacant units" },
  { key: "netExposure", label: "Net exposure to vacancy" },
];

function FlagIndicator({ flags, children }) {
  if (!flags || flags.length === 0) return children;
  return (
    <span className="flag-cell">
      {children}
      <span className="tip">
        {flags.map((f, i) => <div key={i}>{f.message}</div>)}
      </span>
    </span>
  );
}

export default function MonthlyStatusReportComparator() {
  const [reports, setReports] = useState([]);
  const [selectedId, setSelectedId] = useState(null);
  const [dragOver, setDragOver] = useState(false);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);

  const parseFile = (file) => new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: "array", cellDates: true });
        const result = processWorkbook(workbook, file.name);
        if (result.months.length === 0) { reject(`${file.name}: couldn't find a monthly financial section`); return; }
        resolve({ id: `${file.name}-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`, fileName: file.name, ...result });
      } catch (err) {
        reject(`${file.name}: couldn't read this file`);
      }
    };
    reader.onerror = () => reject(`${file.name}: couldn't read this file`);
    reader.readAsArrayBuffer(file);
  });

  const handleFiles = useCallback(async (fileList) => {
    const files = Array.from(fileList || []).filter((f) => /\.xlsx?$/i.test(f.name));
    if (files.length === 0) return;
    setLoading(true);
    setError(null);
    const startedAt = Date.now();
    const results = await Promise.allSettled(files.map(parseFile));
    const elapsed = Date.now() - startedAt;
    const minDuration = 2000;
    if (elapsed < minDuration) await new Promise((res) => setTimeout(res, minDuration - elapsed));
    const succeeded = results.filter((r) => r.status === "fulfilled").map((r) => r.value);
    const failed = results.filter((r) => r.status === "rejected").map((r) => r.reason);
    setReports((prev) => [...prev, ...succeeded]);
    if (succeeded.length > 0) setSelectedId(succeeded[succeeded.length - 1].id);
    if (failed.length > 0) setError(failed.join("; "));
    setLoading(false);
  }, []);

  const removeReport = (id) => {
    setReports((prev) => {
      const next = prev.filter((r) => r.id !== id);
      if (selectedId === id) setSelectedId(next.length ? next[next.length - 1].id : null);
      return next;
    });
  };

  const sortedReports = [...reports].sort((a, b) => (a.propertyName || "").localeCompare(b.propertyName || "") || (a.year || 0) - (b.year || 0));

  const compLabel = (r) => `${r.propertyName || r.fileName}${r.period ? " \u2014 " + r.period : ""}`;

  const comparisonRows = [
    { label: "Location", get: (r) => r.location || "\u2014", isText: true },
    { label: "Reporting period", get: (r) => r.period || "\u2014", isText: true },
    { label: "YTD income", get: (r) => formatMoney(r.ytdIncome) },
    { label: "YTD net income", get: (r) => formatMoney(r.ytdNetIncome) },
    { label: "Avg occupancy", get: (r) => isNaN(r.avgOccupancy) ? "\u2014" : `${(r.avgOccupancy * 100).toFixed(1)}%` },
    { label: "Total move-outs", get: (r) => isNaN(r.totalMoveOuts) ? "\u2014" : Math.round(r.totalMoveOuts).toLocaleString() },
    { label: "Months flagged", get: (r) => `${r.flaggedMonthCount} / ${r.reportedMonthCount}` },
  ];

  const propertyGroups = {};
  reports.forEach((r) => {
    const key = r.propertyName || r.fileName;
    if (!propertyGroups[key]) propertyGroups[key] = [];
    propertyGroups[key].push(r);
  });
  const trendGroups = Object.entries(propertyGroups)
    .filter(([, list]) => list.length >= 2)
    .map(([name, list]) => ({
      name,
      data: [...list].sort((a, b) => (a.year || 0) - (b.year || 0)).map((r) => ({
        year: String(r.year || r.period || r.fileName),
        Income: Math.round(r.ytdIncome || 0),
        "Net income": Math.round(r.ytdNetIncome || 0),
        "Occupancy %": isNaN(r.avgOccupancy) ? null : Math.round(r.avgOccupancy * 1000) / 10,
      })),
    }));

  return (
    <div style={{ fontFamily: "Inter, sans-serif", background: PARCHMENT, minHeight: "100%", padding: "2.5rem 2rem", color: INK }}>
      <style>{FONT_IMPORT}{EXTRA_STYLES}</style>
      <div style={{ maxWidth: 1160, margin: "0 auto" }}>
        <div style={{ marginBottom: "1.75rem" }}>
          <div style={{ fontSize: 12, fontWeight: 600, letterSpacing: "0.12em", color: DEED_GREEN, textTransform: "uppercase", marginBottom: 8 }}>
            Monthly status report comparator
          </div>
          <h1 style={{ fontFamily: "'Source Serif 4', serif", fontWeight: 600, fontSize: 28, margin: 0, lineHeight: 1.2, color: INK }}>
            Normalize and compare property status reports
          </h1>
          <p style={{ color: SLATE, fontSize: 14, marginTop: 10, maxWidth: 640, lineHeight: 1.6 }}>
            Upload one or more monthly status reports (.xls or .xlsx) \u2014 different properties, different years, or both. Each gets checked for NOI/net income reconciliation and unusual months, and you can compare them side by side.
          </p>
        </div>

        {reports.length === 0 ? (
          <div
            onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
            onDragLeave={() => setDragOver(false)}
            onDrop={(e) => { e.preventDefault(); setDragOver(false); handleFiles(e.dataTransfer.files); }}
            style={{ border: `2px dashed ${dragOver ? DEED_GREEN : BORDER}`, background: dragOver ? "#F0F5F1" : CARD, borderRadius: 12, padding: "3.5rem 2rem", textAlign: "center", transition: "border-color 0.15s ease, background 0.15s ease" }}
          >
            {loading ? (
              <>
                <Loader2 size={28} color={DEED_GREEN} className="spin-icon" style={{ marginBottom: 14 }} />
                <div style={{ fontSize: 15, fontWeight: 500, marginBottom: 6 }}>Processing workbook{"\u2026"}</div>
              </>
            ) : (
              <>
                <Upload size={28} color={DEED_GREEN} style={{ marginBottom: 14 }} />
                <div style={{ fontSize: 15, fontWeight: 500, marginBottom: 6 }}>Drop status reports here</div>
              </>
            )}
            <div style={{ fontSize: 13, color: SLATE, marginBottom: 18 }}>.xls or .xlsx \u00b7 select multiple files to compare</div>
            {error && <div style={{ fontSize: 13, color: "#A32D2D", marginBottom: 14 }}>{error}</div>}
            <label style={{ display: "inline-block", background: DEED_GREEN, color: "#FFFFFF", fontSize: 13, fontWeight: 500, padding: "9px 18px", borderRadius: 8, cursor: "pointer" }}>
              Choose files
              <input type="file" accept=".xls,.xlsx" multiple style={{ display: "none" }} onChange={(e) => handleFiles(e.target.files)} />
            </label>
          </div>
        ) : (
          <>
            <div
              onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
              onDragLeave={() => setDragOver(false)}
              onDrop={(e) => { e.preventDefault(); setDragOver(false); handleFiles(e.dataTransfer.files); }}
              style={{ border: `1.5px dashed ${dragOver ? DEED_GREEN : BORDER}`, background: dragOver ? "#F0F5F1" : "transparent", borderRadius: 10, padding: "10px 16px", marginBottom: 20, display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12 }}
            >
              <div style={{ fontSize: 13, color: SLATE, display: "flex", alignItems: "center", gap: 8 }}>
                {loading && <Loader2 size={14} className="spin-icon" color={DEED_GREEN} />}
                {loading ? "Processing workbook(s)\u2026" : "Drop more reports here to add them to the comparison"}
              </div>
              <label style={{ display: "inline-flex", alignItems: "center", gap: 6, background: "transparent", border: `1px solid ${DEED_GREEN}`, color: DEED_GREEN, fontSize: 13, fontWeight: 500, padding: "7px 14px", borderRadius: 8, cursor: "pointer", flexShrink: 0 }}>
                <Plus size={14} /> Add files
                <input type="file" accept=".xls,.xlsx" multiple style={{ display: "none" }} onChange={(e) => handleFiles(e.target.files)} />
              </label>
            </div>
            {error && <div style={{ fontSize: 13, color: "#A32D2D", marginBottom: 16 }}>{error}</div>}

            <div style={{ fontSize: 13, fontWeight: 500, color: SLATE, marginBottom: 10, textTransform: "uppercase", letterSpacing: "0.06em" }}>
              Side-by-side comparison
            </div>
            <div style={{ background: CARD, border: `0.5px solid ${BORDER}`, borderRadius: 12, overflow: "hidden", marginBottom: 28 }}>
              <div style={{ overflowX: "auto" }}>
                <table style={{ borderCollapse: "collapse", fontSize: 13, width: "100%" }}>
                  <thead>
                    <tr style={{ borderBottom: `1px solid ${BORDER}`, background: "#FBFAF6" }}>
                      <th style={{ position: "sticky", left: 0, background: "#FBFAF6", textAlign: "left", padding: "10px 14px", fontWeight: 500, color: SLATE, fontSize: 12, minWidth: 160 }}> </th>
                      {sortedReports.map((r) => (
                        <th key={r.id} style={{ textAlign: "left", padding: "10px 14px", fontWeight: 500, minWidth: 190 }}>
                          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8 }}>
                            <span style={{ fontFamily: "'Source Serif 4', serif", fontWeight: 600, fontSize: 13.5, color: INK }}>
                              {compLabel(r)}
                            </span>
                            <button onClick={() => removeReport(r.id)} title="Remove" style={{ background: "none", border: "none", cursor: "pointer", color: SLATE, padding: 2, flexShrink: 0 }}>
                              <X size={13} />
                            </button>
                          </div>
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {comparisonRows.map((row) => (
                      <tr key={row.label} style={{ borderBottom: `0.5px solid ${BORDER}` }}>
                        <td style={{ position: "sticky", left: 0, background: CARD, padding: "9px 14px", color: SLATE, whiteSpace: "nowrap" }}>{row.label}</td>
                        {sortedReports.map((r) => (
                          <td key={r.id} style={{ padding: "9px 14px", fontFamily: row.isText ? "Inter, sans-serif" : "'IBM Plex Mono', monospace" }}>{row.get(r)}</td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>

            {trendGroups.length > 0 && trendGroups.map((group) => (
              <div key={group.name} style={{ marginBottom: 28 }}>
                <div style={{ fontSize: 13, fontWeight: 500, color: SLATE, marginBottom: 10, textTransform: "uppercase", letterSpacing: "0.06em" }}>
                  {group.name} {"\u2014"} performance over time
                </div>
                <div style={{ background: CARD, border: `0.5px solid ${BORDER}`, borderRadius: 12, padding: "1.5rem 1.25rem 0.5rem" }}>
                  <ResponsiveContainer width="100%" height={280}>
                    <ComposedChart data={group.data} margin={{ top: 4, right: 8, left: 8, bottom: 4 }}>
                      <CartesianGrid stroke={BORDER} strokeDasharray="3 3" vertical={false} />
                      <XAxis dataKey="year" tick={{ fontFamily: "Inter, sans-serif", fontSize: 12, fill: SLATE }} axisLine={{ stroke: BORDER }} tickLine={false} />
                      <YAxis yAxisId="money" tick={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 11, fill: SLATE }} axisLine={false} tickLine={false} tickFormatter={(v) => `$${(v / 1000).toFixed(0)}k`} />
                      <YAxis yAxisId="pct" orientation="right" domain={[0, 100]} tick={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 11, fill: SLATE }} axisLine={false} tickLine={false} tickFormatter={(v) => `${v}%`} />
                      <Tooltip
                        contentStyle={{ background: INK, border: "none", borderRadius: 8, fontFamily: "Inter, sans-serif", fontSize: 12 }}
                        labelStyle={{ color: "#FFFFFF", fontWeight: 600, marginBottom: 4 }}
                        itemStyle={{ color: "#FFFFFF" }}
                        formatter={(value, name) => name === "Occupancy %" ? [`${value}%`, name] : [`$${value.toLocaleString()}`, name]}
                      />
                      <Legend wrapperStyle={{ fontFamily: "Inter, sans-serif", fontSize: 12, color: SLATE }} />
                      <Bar yAxisId="money" dataKey="Income" fill={DEED_GREEN} radius={[4, 4, 0, 0]} barSize={28} />
                      <Bar yAxisId="money" dataKey="Net income" fill="#9BB79E" radius={[4, 4, 0, 0]} barSize={28} />
                      <Line yAxisId="pct" type="monotone" dataKey="Occupancy %" stroke="#5B7B93" strokeWidth={2.5} dot={{ r: 4, fill: "#5B7B93" }} connectNulls />
                    </ComposedChart>
                  </ResponsiveContainer>
                </div>
              </div>
            ))}

            {sortedReports.map((rep) => (
              <div key={rep.id} style={{ marginBottom: 36 }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
                  <div style={{ fontSize: 13, fontWeight: 500, color: SLATE, textTransform: "uppercase", letterSpacing: "0.06em" }}>
                    {compLabel(rep)}
                  </div>
                  {rep.flaggedMonthCount > 0 && (
                    <div style={{ border: `2px solid ${FLAG_AMBER}`, color: FLAG_AMBER, fontFamily: "'IBM Plex Mono', monospace", fontSize: 11, fontWeight: 500, letterSpacing: "0.08em", textTransform: "uppercase", padding: "6px 10px", transform: "rotate(-3deg)", whiteSpace: "nowrap" }}>
                      {rep.flaggedMonthCount} {rep.flaggedMonthCount === 1 ? "month" : "months"} flagged
                    </div>
                  )}
                </div>

                <div style={{ background: CARD, border: `0.5px solid ${BORDER}`, borderRadius: 12, padding: "1.1rem 1.5rem", marginBottom: 20, display: "flex", flexWrap: "wrap", gap: "1.25rem 2.25rem" }}>
                  {[
                    { label: "Units", value: rep.units ?? "\u2014" },
                    { label: "Square footage", value: rep.sqft ? rep.sqft.toLocaleString() : "\u2014" },
                    { label: "Purchase date", value: rep.purchaseDate || "\u2014" },
                    { label: "Purchase price", value: rep.purchasePrice || "\u2014" },
                  ].map((s) => (
                    <div key={s.label}>
                      <div style={{ fontSize: 12, color: SLATE, marginBottom: 4 }}>{s.label}</div>
                      <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 14, fontWeight: 500 }}>{s.value}</div>
                    </div>
                  ))}
                </div>

                <div style={{ background: CARD, border: `0.5px solid ${BORDER}`, borderRadius: 12, overflow: "hidden", marginBottom: 24 }}>
                  <div style={{ overflowX: "auto" }}>
                    <table style={{ borderCollapse: "collapse", fontSize: 13, width: "100%" }}>
                      <thead>
                        <tr style={{ borderBottom: `1px solid ${BORDER}`, background: "#FBFAF6" }}>
                          <th style={{ position: "sticky", left: 0, background: "#FBFAF6", textAlign: "left", padding: "10px 14px", fontWeight: 500, color: SLATE, fontSize: 12, minWidth: 190 }}> </th>
                          {rep.months.map((mo) => (
                            <th key={mo.key} style={{ textAlign: "right", padding: "10px 14px", fontWeight: 500, color: mo.flags.length ? FLAG_AMBER : SLATE, fontSize: 12, whiteSpace: "nowrap" }}>
                              <FlagIndicator flags={mo.flags}>{mo.label}{mo.flags.length ? " \u25CF" : ""}</FlagIndicator>
                            </th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {FIN_ROWS.map((row) => (
                          <tr key={row.key} style={{ borderBottom: `0.5px solid ${BORDER}` }}>
                            <td style={{ position: "sticky", left: 0, background: CARD, padding: "9px 14px", color: row.key === "noi" || row.key === "netIncome" ? INK : SLATE, fontWeight: row.key === "noi" || row.key === "netIncome" ? 500 : 400, whiteSpace: "nowrap" }}>{row.label}</td>
                            {rep.months.map((mo) => {
                              const flag = mo.flags.filter((f) => f.field === row.key);
                              const display = mo.notYetReported ? "\u2014" : formatMoney(mo[row.key]);
                              return (
                                <td key={mo.key} style={{ padding: "9px 14px", textAlign: "right", fontFamily: "'IBM Plex Mono', monospace", background: flag.length ? FLAG_AMBER_BG : "transparent", color: flag.length ? FLAG_AMBER : (mo.notYetReported ? SLATE : INK) }}>
                                  <FlagIndicator flags={flag}>{display}</FlagIndicator>
                                </td>
                              );
                            })}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>

                <div style={{ fontSize: 13, fontWeight: 500, color: SLATE, marginBottom: 10, textTransform: "uppercase", letterSpacing: "0.06em" }}>Occupancy &amp; leasing</div>
                <div style={{ background: CARD, border: `0.5px solid ${BORDER}`, borderRadius: 12, overflow: "hidden", marginBottom: 20 }}>
                  <div style={{ overflowX: "auto" }}>
                    <table style={{ borderCollapse: "collapse", fontSize: 13, width: "100%" }}>
                      <thead>
                        <tr style={{ borderBottom: `1px solid ${BORDER}`, background: "#FBFAF6" }}>
                          <th style={{ position: "sticky", left: 0, background: "#FBFAF6", textAlign: "left", padding: "10px 14px", fontWeight: 500, color: SLATE, fontSize: 12, minWidth: 190 }}> </th>
                          {rep.months.map((mo) => (
                            <th key={mo.key} style={{ textAlign: "right", padding: "10px 14px", fontWeight: 500, color: SLATE, fontSize: 12, whiteSpace: "nowrap" }}>{mo.label}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {OCC_ROWS.map((row) => (
                          <tr key={row.key} style={{ borderBottom: `0.5px solid ${BORDER}` }}>
                            <td style={{ position: "sticky", left: 0, background: CARD, padding: "9px 14px", color: SLATE, whiteSpace: "nowrap" }}>{row.label}</td>
                            {rep.months.map((mo) => {
                              const relevantField = row.key === "occupiedPct" ? "occupiedPct" : "netExposure";
                              const flag = mo.flags.filter((f) => f.field === relevantField);
                              const val = mo[row.key];
                              const display = mo.notYetReported ? "\u2014" : (row.isPct ? (isNaN(val) ? "\u2014" : `${(val * 100).toFixed(1)}%`) : (isNaN(val) ? "\u2014" : Math.round(val).toLocaleString()));
                              return (
                                <td key={mo.key} style={{ padding: "9px 14px", textAlign: "right", fontFamily: "'IBM Plex Mono', monospace", background: flag.length ? FLAG_AMBER_BG : "transparent", color: flag.length ? FLAG_AMBER : (mo.notYetReported ? SLATE : INK) }}>
                                  <FlagIndicator flags={flag}>{display}</FlagIndicator>
                                </td>
                              );
                            })}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>

                <div style={{ fontSize: 12, color: SLATE, display: "flex", alignItems: "center", gap: 6 }}>
                  <FileSpreadsheet size={14} /> {rep.fileName} {"\u00b7"} sheet "{rep.sheetName}"
                </div>
              </div>
            ))}

            <div style={{ marginTop: 32, paddingTop: 20, borderTop: `0.5px solid ${BORDER}`, display: "flex", justifyContent: "center" }}>
              <button
                onClick={() => { setReports([]); setSelectedId(null); setError(null); }}
                style={{ display: "flex", alignItems: "center", gap: 6, background: "transparent", border: `1px solid ${BORDER}`, borderRadius: 8, padding: "9px 16px", fontSize: 13, color: SLATE, cursor: "pointer" }}
              >
                <RotateCcw size={13} /> Clear all &amp; start over
              </button>
            </div>
          </>
        )}
      </div>
    </div>
  );
}
