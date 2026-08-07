"use client";

import React, { useState, useCallback, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import { Upload, FileSpreadsheet, AlertTriangle, X, Plus, RotateCcw, Loader2, Search } from "lucide-react";
import { ComposedChart, Bar, Line, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer } from "recharts";

const FONT_IMPORT = `@import url('https://fonts.googleapis.com/css2?family=Source+Serif+4:wght@500;600&family=IBM+Plex+Mono:wght@400;500&family=Inter:wght@400;500;600&display=swap');`;

const INK = "#EAEEF5";
const PARCHMENT = "#0B1220";
const CARD = "#141C2B";
const HEADER_BG = "#1B2436";
const SURFACE_DARK = "#070B12";
const DEED_GREEN = "#4FAE86";
const FLAG_AMBER = "#E3A452";
const FLAG_AMBER_BG = "#3A2C16";
const SLATE = "#8B96A8";
const BORDER = "#2A3448";
const ERROR_COLOR = "#F0796B";
const GOOD_COLOR = "#4FAE86";
const BAD_COLOR = "#E3826F";
const HOVER_BG = "#1B2536";

const MONTHS = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];
const MONTH_LABELS = { jan: "Jan", feb: "Feb", mar: "Mar", apr: "Apr", may: "May", jun: "Jun", jul: "Jul", aug: "Aug", sep: "Sep", oct: "Oct", nov: "Nov", dec: "Dec" };

const EXTRA_STYLES = `
@keyframes toastFade { 0% { opacity: 0; transform: translate(-50%, 8px); } 10% { opacity: 1; transform: translate(-50%, 0); } 85% { opacity: 1; transform: translate(-50%, 0); } 100% { opacity: 0; transform: translate(-50%, 8px); } }
.toast-msg { animation: toastFade 2.6s ease forwards; }
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
  background: ${SURFACE_DARK};
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
  box-shadow: 0 6px 16px rgba(0,0,0,0.45);
}
.flag-cell .tip::after {
  content: '';
  position: absolute;
  bottom: 100%;
  left: 50%;
  transform: translateX(-50%);
  border: 5px solid transparent;
  border-bottom-color: ${SURFACE_DARK};
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
    let s = v.trim();
    let negative = false;
    if (/^\(.*\)$/.test(s)) { negative = true; s = s.slice(1, -1); }
    const cleaned = s.replace(/[$,\s]/g, "");
    if (cleaned === "" || cleaned === "-") return NaN;
    let n = parseFloat(cleaned);
    if (isNaN(n)) return NaN;
    if (negative) n = -Math.abs(n);
    return n;
  }
  return NaN;
}
// Guards against values like parseNumber("$12.5M") returning 12.5 \u2014 a real dollar amount for a
// property, a loan, or a total cost basis should never be a tiny number. Anything implausibly small
// is treated as unparseable (NaN) rather than silently feeding a nonsense percentage downstream.
function sanitizeDollarAmount(n, minimum) {
  if (isNaN(n)) return NaN;
  return n >= (minimum || 100000) ? n : NaN;
}
function formatMoney(v) {
  if (v === undefined || isNaN(v)) return "\u2014";
  const abs = Math.round(Math.abs(v)).toLocaleString();
  return v < 0 ? `(${abs})` : abs;
}
// Deterministic "what changed" \u2014 computed in JS from the actual month-over-month figures, not
// authored by the model. Only "why" and "action to take" are AI-generated; every number on screen
// traces back to arithmetic, not a language model's guess.
function computeWhatChanged(latest, prior) {
  if (!latest) return ["No reported month available yet."];
  const bullets = [];
  if (prior) {
    if (!isNaN(latest.noi) && !isNaN(prior.noi)) {
      const delta = latest.noi - prior.noi;
      const pct = prior.noi !== 0 ? (delta / Math.abs(prior.noi)) * 100 : null;
      bullets.push(`NOI: ${formatMoney(prior.noi)} \u2192 ${formatMoney(latest.noi)} (${delta >= 0 ? "+" : ""}${formatMoney(delta)}${pct !== null ? `, ${pct >= 0 ? "+" : ""}${pct.toFixed(1)}%` : ""})`);
    }
    if (!isNaN(latest.occupiedPct) && !isNaN(prior.occupiedPct)) {
      const deltaPP = (latest.occupiedPct - prior.occupiedPct) * 100;
      bullets.push(`Occupancy: ${(prior.occupiedPct * 100).toFixed(1)}% \u2192 ${(latest.occupiedPct * 100).toFixed(1)}% (${deltaPP >= 0 ? "+" : ""}${deltaPP.toFixed(1)}pp)`);
    }
    if (!isNaN(latest.opEx) && !isNaN(prior.opEx)) {
      const delta = Math.abs(latest.opEx) - Math.abs(prior.opEx);
      const pct = prior.opEx !== 0 ? (delta / Math.abs(prior.opEx)) * 100 : null;
      bullets.push(`Operating expenses: ${formatMoney(Math.abs(prior.opEx))} \u2192 ${formatMoney(Math.abs(latest.opEx))} (${delta >= 0 ? "+" : ""}${formatMoney(delta)}${pct !== null ? `, ${pct >= 0 ? "+" : ""}${pct.toFixed(1)}%` : ""})`);
    }
  } else {
    bullets.push(`${latest.label}: NOI ${formatMoney(latest.noi)}, occupancy ${isNaN(latest.occupiedPct) ? "\u2014" : (latest.occupiedPct * 100).toFixed(1) + "%"} \u2014 no prior reported month to compare yet.`);
  }
  return bullets.length ? bullets : ["No comparable figures available for this month."];
}
// Insurance: the prompts tell the model never to say "cap rate," but a prompt instruction is not
// a guarantee. Scrub any occurrence from generated text before it ever reaches the screen, since
// this exact distinction (yield on cost vs. a market cap rate we can't actually calculate) is the
// signal this tool is trying to send.
function scrubCapRate(s) {
  return typeof s === "string" ? s.replace(/cap(italization)?\s*rate/gi, "yield on cost") : s;
}
function scrubDeep(value) {
  if (Array.isArray(value)) return value.map(scrubDeep);
  if (typeof value === "string") return scrubCapRate(value);
  if (value && typeof value === "object") {
    const out = {};
    Object.entries(value).forEach(([k, v]) => { out[k] = scrubDeep(v); });
    return out;
  }
  return value;
}

// Simple concurrency limiter so uploading several files doesn't fire a burst of simultaneous AI
// calls (decision support + multi-year overview) all at once, which is a common cause of hitting
// rate limits and every single one failing together.
let activeAICalls = 0;
const aiCallQueue = [];
const AI_CONCURRENCY_LIMIT = 2;
function runQueued(task) {
  return new Promise((resolve, reject) => {
    const run = async () => {
      activeAICalls++;
      try {
        resolve(await task());
      } catch (err) {
        reject(err);
      } finally {
        activeAICalls--;
        if (aiCallQueue.length > 0) aiCallQueue.shift()();
      }
    };
    if (activeAICalls < AI_CONCURRENCY_LIMIT) run();
    else aiCallQueue.push(run);
  });
}

// Calls the Claude API with retry/backoff on transient failures (rate limit, overload, network
// blip). Without this, a single 429/529 during a burst of concurrent calls just fails outright
// with no second attempt, which reads as "wasn't available" for no visible reason.
async function callClaudeText(prompt, maxTokens) {
  return runQueued(async () => {
    let lastErr = null;
    for (let attempt = 0; attempt < 3; attempt++) {
      try {
        // This calls our own Next.js API route (app/api/claude/route.js), not Anthropic directly.
        // The route holds the real API key server-side and proxies the request \u2014 see that file
        // for why this can't just call api.anthropic.com from the browser on a plain Vercel deploy.
        const response = await fetch("/api/claude", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            maxTokens: maxTokens || 1000,
            messages: [{ role: "user", content: prompt }],
          }),
        });
        if (!response.ok) {
          const retryable = response.status === 429 || response.status === 529 || response.status >= 500;
          let detail = "";
          try { detail = JSON.stringify(await response.json()); } catch (e) { /* ignore */ }
          const err = new Error(`API returned ${response.status}${detail ? ": " + detail : ""}`);
          if (retryable && attempt < 2) { lastErr = err; await new Promise((r) => setTimeout(r, 600 * Math.pow(2, attempt))); continue; }
          throw err;
        }
        const data = await response.json();
        const textBlock = (data.content || []).find((b) => b.type === "text");
        if (!textBlock) throw new Error("No text in response");
        return textBlock.text;
      } catch (err) {
        lastErr = err;
        if (attempt < 2) { await new Promise((r) => setTimeout(r, 600 * Math.pow(2, attempt))); continue; }
      }
    }
    throw lastErr || new Error("Request failed");
  });
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
function findYearedRow(labelRows, regex, preferredYear) {
  const matches = labelRows.map((r) => ({ ...r, m: r.label.match(regex) })).filter((r) => r.m);
  if (matches.length === 0) return null;
  if (preferredYear) {
    const exact = matches.find((r) => parseInt(r.m[1], 10) === preferredYear);
    if (exact) return exact.rowIdx;
  }
  matches.sort((a, b) => parseInt(b.m[1], 10) - parseInt(a.m[1], 10));
  return matches[0].rowIdx;
}
function extractRowValues(grid, rowIdx, colOfMonth, ytdCol) {
  if (rowIdx === null) return { byMonth: {}, ytd: NaN, ytdAddr: null };
  const row = grid[rowIdx];
  const byMonth = {};
  MONTHS.forEach((m) => { if (colOfMonth[m] !== undefined) byMonth[m] = parseNumber(row[colOfMonth[m]]); });
  const ytd = ytdCol !== null ? parseNumber(row[ytdCol]) : NaN;
  const ytdAddr = ytdCol !== null ? XLSX.utils.encode_cell({ r: rowIdx, c: ytdCol }) : null;
  return { byMonth, ytd, ytdAddr };
}

function processWorkbook(workbook, fileName) {
  const { name: sheetName, grid } = findReportSheet(workbook);
  const sheetObj = workbook.Sheets[sheetName];
  const statusRowIdx = grid.findIndex((row) => row.some((c) => norm(c) === "monthly status report"));
  const propertyName = statusRowIdx >= 2 ? firstNonEmpty(grid[statusRowIdx - 2]) : "";
  const location = statusRowIdx >= 1 ? firstNonEmpty(grid[statusRowIdx - 1]) : "";
  const reportDateRaw = statusRowIdx >= 0 ? firstNonEmpty(grid[statusRowIdx + 1]) : "";
  const reportDate = reportDateRaw instanceof Date ? reportDateRaw : (reportDateRaw ? new Date(reportDateRaw) : null);
  const yearForMatching = reportDate && !isNaN(reportDate) ? reportDate.getFullYear() : parseInt(((fileName || "").match(/(20\d{2})/) || [])[1] || "0", 10) || null;

  const labelRows = buildLabelIndex(grid);
  const unitsRowIdx = findRow(labelRows, (l) => l === "number of units");
  const sqftRowIdx = findRow(labelRows, (l) => l === "square footage");
  let units = null, sqft = null;
  if (unitsRowIdx !== null) units = parseNumber(grid[unitsRowIdx].find((c) => typeof c === "number"));
  if (sqftRowIdx !== null) sqft = parseNumber(grid[sqftRowIdx].find((c) => typeof c === "number"));
  // Property header fields (purchase price, total cost, existing debt) are laid out inconsistently
  // across properties: sometimes "Label: Value" sits in one cell, sometimes the label is alone and
  // the value is in the next cell over. Search for both patterns rather than assuming one.
  function extractLabeledValue(labelPattern, maxRows) {
    for (let r = 0; r < Math.min(grid.length, maxRows); r++) {
      const row = grid[r];
      for (let c = 0; c < row.length; c++) {
        const s = String(row[c] == null ? "" : row[c]);
        const m = s.match(labelPattern);
        if (m) {
          if (m[1] && m[1].trim()) return { value: m[1].trim(), addr: XLSX.utils.encode_cell({ r, c }) };
          // Label-only cell (value lives in a nearby cell). Search a few cells to the right, but
          // stop if we hit what looks like a *different* field's label (e.g. "Debt:") rather than
          // this field's value \u2014 otherwise an empty field can silently borrow its neighbor's value.
          for (let c2 = c + 1; c2 < row.length && c2 <= c + 4; c2++) {
            const next = String(row[c2] == null ? "" : row[c2]).trim();
            if (next === "") continue;
            if (/^[A-Za-z][\w /]{1,30}:/.test(next)) break;
            return { value: next, addr: XLSX.utils.encode_cell({ r, c: c2 }) };
          }
          return { value: null, addr: null };
        }
      }
    }
    return { value: null, addr: null };
  }

  const purchaseDateRes = extractLabeledValue(/purchase date:\s*(.*)/i, 10);
  const purchasePriceRes = extractLabeledValue(/purchase price:\s*(.*)/i, 10);
  const totalCostRes = extractLabeledValue(/total cost:\s*(.*)/i, 10);
  const debtRes = extractLabeledValue(/(?:existing )?debt:\s*(.*)/i, 10);
  const purchaseDate = purchaseDateRes.value;
  const purchasePrice = purchasePriceRes.value;
  const purchasePriceAddr = purchasePriceRes.addr;
  const totalCost = totalCostRes.value;
  const totalCostAddr = totalCostRes.addr;
  const totalCostNum = totalCost ? sanitizeDollarAmount(parseNumber(totalCost)) : NaN;
  const debtText = debtRes.value;
  const debtAddr = debtRes.addr;
  const loanBalanceMatch = debtText && !/^none$/i.test(debtText.trim()) ? debtText.match(/\$[\d,]+/) : null;
  const loanBalanceNum = loanBalanceMatch ? sanitizeDollarAmount(parseNumber(loanBalanceMatch[0])) : NaN;

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
    netCashFlow: findRow(labelRows, (l) => l === "net cash flow"),
    occupiedPct: findYearedRow(labelRows, /^(\d{4}) occupied %$/, yearForMatching),
    moveOuts: findYearedRow(labelRows, /^(\d{4}) move-?outs$/, yearForMatching),
    currentVacantUnits: findRow(labelRows, (l) => l === "current vacant units"),
    netExposureVacancy: findRow(labelRows, (l) => l === "net exposure to vacancy"),
  };

  const extracted = {};
  Object.entries(rowIdx).forEach(([key, idx]) => { extracted[key] = extractRowValues(grid, idx, colOfMonth, ytdCol); });

  const monthsPresent = MONTHS.filter((m) => colOfMonth[m] !== undefined);
  const dataQualityNotes = [];

  // Sign convention: the reconciliation math (Income + expenses = NOI) only works if expense rows
  // are stored negative. Some templates store them positive instead. Detect via the median of
  // non-zero Operating Expenses values and flip every expense-type row if needed, rather than
  // silently failing every reconciliation check and looking like broken data.
  const opExSampleValues = monthsPresent.map((m) => extracted.operatingExpenses.byMonth[m]).filter((v) => typeof v === "number" && !isNaN(v) && v !== 0);
  const sortedOpEx = [...opExSampleValues].sort((a, b) => a - b);
  const medianOpEx = sortedOpEx.length ? sortedOpEx[Math.floor(sortedOpEx.length / 2)] : 0;
  if (medianOpEx > 0) {
    dataQualityNotes.push("Expense figures in this report appear to be stored as positive numbers (unusual for this template) \u2014 signs were flipped automatically before reconciling. Verify against the source if anything looks off.");
    ["operatingExpenses", "recurringCapex", "ownershipExpenses", "nonRecurringCapex", "debtService"].forEach((field) => {
      Object.keys(extracted[field].byMonth).forEach((m) => {
        const v = extracted[field].byMonth[m];
        if (typeof v === "number" && !isNaN(v)) extracted[field].byMonth[m] = -Math.abs(v);
      });
      if (!isNaN(extracted[field].ytd)) extracted[field].ytd = -Math.abs(extracted[field].ytd);
    });
  }

  // Occupancy scale: some templates store 93 instead of 0.93. Anything above 1.5 is unambiguously
  // a percentage stored as a whole number (150%+ occupancy isn't real), not a rate near 100%.
  let occupancyRescaled = false;
  Object.keys(extracted.occupiedPct.byMonth).forEach((m) => {
    const v = extracted.occupiedPct.byMonth[m];
    if (typeof v === "number" && !isNaN(v) && v > 1.5) { extracted.occupiedPct.byMonth[m] = v / 100; occupancyRescaled = true; }
  });
  if (!isNaN(extracted.occupiedPct.ytd) && extracted.occupiedPct.ytd > 1.5) extracted.occupiedPct.ytd = extracted.occupiedPct.ytd / 100;
  if (occupancyRescaled) dataQualityNotes.push("Occupancy figures were stored as whole numbers (e.g. 93 instead of 0.93) and were rescaled automatically.");

  const cellAddr = (rIdx, cIdx) => (rIdx === null || rIdx === undefined || cIdx === undefined) ? null : XLSX.utils.encode_cell({ r: rIdx, c: cIdx });
  const months = monthsPresent.map((m) => {
    const get = (k) => extracted[k].byMonth[m];
    const income = get("income"), opEx = get("operatingExpenses"), recurCapex = get("recurringCapex"), ownershipExp = get("ownershipExpenses");
    const reportedNOI = get("noi"), nonRecurCapex = get("nonRecurringCapex"), debtService = get("debtService"), reportedNetIncome = get("netIncome");
    const netCashFlow = get("netCashFlow");
    const occupiedPct = get("occupiedPct"), moveOuts = get("moveOuts"), vacantUnits = get("currentVacantUnits"), netExposure = get("netExposureVacancy");
    const colIdx = colOfMonth[m];
    const src = {
      income: cellAddr(rowIdx.income, colIdx), opEx: cellAddr(rowIdx.operatingExpenses, colIdx), recurCapex: cellAddr(rowIdx.recurringCapex, colIdx),
      ownershipExp: cellAddr(rowIdx.ownershipExpenses, colIdx), noi: cellAddr(rowIdx.noi, colIdx), nonRecurCapex: cellAddr(rowIdx.nonRecurringCapex, colIdx),
      debtService: cellAddr(rowIdx.debtService, colIdx), netIncome: cellAddr(rowIdx.netIncome, colIdx), netCashFlow: cellAddr(rowIdx.netCashFlow, colIdx), occupiedPct: cellAddr(rowIdx.occupiedPct, colIdx),
      moveOuts: cellAddr(rowIdx.moveOuts, colIdx), vacantUnits: cellAddr(rowIdx.currentVacantUnits, colIdx), netExposure: cellAddr(rowIdx.netExposureVacancy, colIdx),
    };
    const srcNote = (field) => src[field] ? ` \u2014 source: sheet "${sheetName}", cell ${src[field]}` : "";

    // Some templates zero-fill months that haven't happened yet (rather than leaving them blank).
    // Treat a month as "not yet reported" when every core financial figure is exactly zero/blank
    // and occupancy shows no real value either.
    const zeroOrNaN = (v) => v === 0 || isNaN(v);
    const notYetReported = zeroOrNaN(income) && zeroOrNaN(opEx) && zeroOrNaN(reportedNOI) && zeroOrNaN(reportedNetIncome) && isNaN(occupiedPct);

    const flags = [];
    const estRentPerUnit = (!isNaN(income) && !isNaN(occupiedPct) && occupiedPct > 0 && units > 0) ? income / (units * occupiedPct) : NaN;
    if (!notYetReported) {
      const hasAllNoiInputs = [income, opEx, recurCapex, ownershipExp, reportedNOI].every((v) => !isNaN(v));
      if (hasAllNoiInputs) {
        const computedNOI = income + opEx + recurCapex + ownershipExp;
        if (Math.abs(computedNOI - reportedNOI) > 5) {
          flags.push({ field: "noi", severity: "critical", tier: "data-integrity", dollarImpact: Math.abs(computedNOI - reportedNOI), message: `NOI doesn't reconcile: computes to ${formatMoney(computedNOI)}, report shows ${formatMoney(reportedNOI)}${srcNote("noi")}` });
        }
      }
      if (!isNaN(reportedNOI) && !isNaN(nonRecurCapex) && !isNaN(debtService) && !isNaN(reportedNetIncome)) {
        const computedNetIncome = reportedNOI + nonRecurCapex + debtService;
        if (Math.abs(computedNetIncome - reportedNetIncome) > 5) {
          flags.push({ field: "netIncome", severity: "critical", tier: "data-integrity", dollarImpact: Math.abs(computedNetIncome - reportedNetIncome), message: `Net income doesn't reconcile: computes to ${formatMoney(computedNetIncome)}, report shows ${formatMoney(reportedNetIncome)}${srcNote("netIncome")}` });
        }
      }
      if (!isNaN(occupiedPct) && occupiedPct < 0.9) {
        const vacantUnitEquiv = units > 0 ? units * (1 - occupiedPct) : NaN;
        const impact = (!isNaN(vacantUnitEquiv) && !isNaN(estRentPerUnit)) ? vacantUnitEquiv * estRentPerUnit : null;
        flags.push({ field: "occupiedPct", severity: "review", tier: "business", dollarImpact: impact, message: `Occupancy at ${(occupiedPct * 100).toFixed(1)}%, below 90%${srcNote("occupiedPct")}` });
      }
      if (isNaN(income)) flags.push({ field: "income", severity: "review", tier: "data-integrity", dollarImpact: null, message: `No income figure reported for this month${srcNote("income")}` });
    }

    return { key: m, label: MONTH_LABELS[m], income, opEx, recurCapex, ownershipExp, noi: reportedNOI, nonRecurCapex, debtService, netIncome: reportedNetIncome, netCashFlow, occupiedPct, moveOuts, vacantUnits, netExposure, flags, notYetReported, src, srcNote, estRentPerUnit };
  });

  const reportedMonths = months.filter((mo) => !mo.notYetReported);
  const validIncomes = reportedMonths.map((mo) => mo.income).filter((v) => !isNaN(v) && v > 0).sort((a, b) => a - b);
  const medIncome = validIncomes.length ? validIncomes[Math.floor(validIncomes.length / 2)] : 0;

  const gaps = reportedMonths.map((mo) => (!isNaN(mo.noi) && !isNaN(mo.netIncome)) ? Math.abs(mo.noi - mo.netIncome) : null).filter((v) => v !== null).sort((a, b) => a - b);
  const medAbsGap = gaps.length ? gaps[Math.floor(gaps.length / 2)] : 0;
  const gapThreshold = Math.max(medAbsGap * 2.5, 5000);

  reportedMonths.forEach((mo) => {
    if (!isNaN(mo.income) && medIncome > 0 && (mo.income > medIncome * 2.5 || mo.income < medIncome * 0.4)) {
      mo.flags.push({ field: "income", severity: "review", tier: "data-integrity", dollarImpact: Math.abs(mo.income - medIncome), message: `Income (${formatMoney(mo.income)}) is far from the typical month (median ${formatMoney(medIncome)})${mo.srcNote("income")}` });
    }
    if (!isNaN(mo.netExposure) && units && units > 0) {
      const ratio = mo.netExposure / units;
      if (ratio > 0.05) {
        const impact = !isNaN(mo.estRentPerUnit) ? mo.netExposure * mo.estRentPerUnit : null;
        mo.flags.push({ field: "netExposure", severity: "review", tier: "business", dollarImpact: impact, message: `Net exposure to vacancy at ${mo.netExposure} units (${(ratio * 100).toFixed(1)}% of ${units} units)${mo.srcNote("netExposure")}` });
      }
    } else if (!isNaN(mo.netExposure) && mo.netExposure >= 2) {
      mo.flags.push({ field: "netExposure", severity: "review", tier: "business", dollarImpact: null, message: `Net exposure to vacancy at ${mo.netExposure} units${mo.srcNote("netExposure")}` });
    }
    if (!isNaN(mo.noi) && !isNaN(mo.netIncome)) {
      const gap = Math.abs(mo.noi - mo.netIncome);
      if (gap > gapThreshold) {
        mo.flags.push({ field: "netIncome", severity: "review", tier: "business", dollarImpact: gap, message: `Gap between NOI (${formatMoney(mo.noi)}) and net income (${formatMoney(mo.netIncome)}) is unusually large for this property \u2014 check capex/debt service this month${mo.srcNote("netIncome")}` });
      }
    }
  });

  const year = yearForMatching;

  const ytdIncome = !isNaN(extracted.income.ytd) ? extracted.income.ytd : reportedMonths.reduce((s, mo) => s + (isNaN(mo.income) ? 0 : mo.income), 0);
  const ytdNetIncome = !isNaN(extracted.netIncome.ytd) ? extracted.netIncome.ytd : reportedMonths.reduce((s, mo) => s + (isNaN(mo.netIncome) ? 0 : mo.netIncome), 0);
  const ytdNetCashFlow = !isNaN(extracted.netCashFlow.ytd) ? extracted.netCashFlow.ytd : reportedMonths.reduce((s, mo) => s + (isNaN(mo.netCashFlow) ? 0 : mo.netCashFlow), 0);
  const ytdNOI = reportedMonths.reduce((s, mo) => s + (isNaN(mo.noi) ? 0 : mo.noi), 0);
  const ytdDebtService = reportedMonths.reduce((s, mo) => s + (isNaN(mo.debtService) ? 0 : Math.abs(mo.debtService)), 0);
  const dscr = ytdDebtService > 0 ? ytdNOI / ytdDebtService : NaN;
  const annualizedNOI = reportedMonths.length > 0 ? (ytdNOI / reportedMonths.length) * 12 : NaN;
  const purchasePriceNum = purchasePrice ? sanitizeDollarAmount(parseNumber(purchasePrice)) : NaN;
  const capRate = (!isNaN(annualizedNOI) && !isNaN(purchasePriceNum) && purchasePriceNum > 0) ? annualizedNOI / purchasePriceNum : NaN;
  const yieldOnCost = (!isNaN(annualizedNOI) && !isNaN(totalCostNum) && totalCostNum > 0) ? annualizedNOI / totalCostNum : NaN;
  const debtYield = (!isNaN(annualizedNOI) && !isNaN(loanBalanceNum) && loanBalanceNum > 0) ? annualizedNOI / loanBalanceNum : NaN;
  const occValues = reportedMonths.map((mo) => mo.occupiedPct).filter((v) => !isNaN(v));
  const avgOccupancy = occValues.length ? occValues.reduce((a, b) => a + b, 0) / occValues.length : NaN;
  const moveOutValues = reportedMonths.map((mo) => mo.moveOuts).filter((v) => !isNaN(v));
  const totalMoveOuts = moveOutValues.length ? moveOutValues.reduce((a, b) => a + b, 0) : NaN;
  const flaggedMonthCount = reportedMonths.filter((mo) => mo.flags.length > 0).length;
  const hasFlaggedFinancials = reportedMonths.some((mo) => mo.flags.some((f) => ["noi", "income", "netIncome"].includes(f.field)));
  const allFlags = reportedMonths.flatMap((mo) => mo.flags.map((f) => ({ ...f, monthLabel: mo.label, monthKey: mo.key })));
  // Only sum business-tier exceptions (vacancy exposure, occupancy shortfall, NOI/net-income gap)
  // into a dollar total. Data-integrity flags (a reconciliation gap, an outlier vs. median) are a
  // different kind of number \u2014 a figure that doesn't tie is not "money lost," and adding the two
  // together produces a total that doesn't mean anything. Those still count and show, just not here.
  const businessImpactTotal = allFlags.filter((f) => f.tier === "business").reduce((s, f) => s + (typeof f.dollarImpact === "number" ? f.dollarImpact : 0), 0);
  const criticalCount = allFlags.filter((f) => f.severity === "critical").length;

  // Reconciliation: never present a confident total without checking it against the report's
  // own stated total for the same figure (its YTD/Actual column), not just our own arithmetic.
  // Critical: the two numbers being compared must be derived INDEPENDENTLY. Using a display value
  // that already falls back to the report's own YTD column (as ytdIncome/ytdNetIncome/ytdNetCashFlow
  // do) would compare that value to itself and could never fail.
  const RECON_TOLERANCE = 5;
  const sumMonths = (field) => reportedMonths.reduce((s, mo) => s + (isNaN(mo[field]) ? 0 : mo[field]), 0);
  function buildReconCheck(label, computed, extractedField) {
    const reported = extracted[extractedField].ytd;
    const reportedAddr = extracted[extractedField].ytdAddr;
    if (isNaN(reported) || reportedAddr === null) return { label, computed, reported: NaN, reportedAddr: null, status: "not-checked" };
    const diff = computed - reported;
    return { label, computed, reported, reportedAddr, status: Math.abs(diff) <= RECON_TOLERANCE ? "matches" : "mismatch", diff };
  }
  const reconciliation = [
    buildReconCheck("Income", sumMonths("income"), "income"),
    buildReconCheck("Net operating income", sumMonths("noi"), "noi"),
    buildReconCheck("Net income", sumMonths("netIncome"), "netIncome"),
    buildReconCheck("Net cash flow", sumMonths("netCashFlow"), "netCashFlow"),
  ];
  // "Reconciled" means every checkable figure actually matched \u2014 not merely "nothing failed."
  // A file with no YTD column at all should read "not checked," never "Yes."
  const anyChecked = reconciliation.some((c) => c.status !== "not-checked");
  const anyMismatch = reconciliation.some((c) => c.status === "mismatch");
  const reconciliationStatus = anyMismatch ? "mismatch" : (anyChecked ? "matches" : "not-checked");
  const isReconciled = reconciliationStatus === "matches";

  let period = String(year || "");
  if (reportedMonths.length > 0 && reportedMonths.length < 12) {
    period = `${reportedMonths[0].label}\u2013${reportedMonths[reportedMonths.length - 1].label} ${year || ""}`.trim();
  }

  // Build a high-fidelity preview of the sheet: real merged-cell ranges, real column widths, and
  // each cell's actual Excel-formatted display text (e.g. "(17,822)" for a negative accounting
  // number, not the raw -17822) so the side panel reads like the real workbook. Note: this legacy
  // .xls format doesn't expose per-cell fill/font/border styling through the available parser \u2014
  // that's a genuine format limitation, not something skipped here.
  let lastRow = 0, lastCol = 0;
  grid.forEach((row, r) => {
    row.forEach((cell, c) => {
      if (String(cell).trim() !== "") { if (r > lastRow) lastRow = r; if (c > lastCol) lastCol = c; }
    });
  });
  const previewRowCount = Math.min(lastRow + 1, 150);
  const previewColCount = Math.min(lastCol, 44) + 1;
  const previewGrid = [];
  for (let r = 0; r < previewRowCount; r++) {
    const rowArr = [];
    for (let c = 0; c < previewColCount; c++) {
      const cell = sheetObj[XLSX.utils.encode_cell({ r, c })];
      if (cell) {
        rowArr.push({ v: cell.v, w: cell.w !== undefined ? cell.w : (cell.v !== undefined ? String(cell.v) : ""), t: cell.t, f: cell.f });
      } else {
        rowArr.push({ v: "", w: "", t: undefined, f: undefined });
      }
    }
    previewGrid.push(rowArr);
  }
  const previewMerges = (sheetObj["!merges"] || [])
    .filter((m) => m.s.r < previewRowCount && m.s.c < previewColCount)
    .map((m) => ({ s: { r: m.s.r, c: m.s.c }, e: { r: Math.min(m.e.r, previewRowCount - 1), c: Math.min(m.e.c, previewColCount - 1) } }));
  const previewColWidths = [];
  for (let c = 0; c < previewColCount; c++) {
    const colInfo = (sheetObj["!cols"] || [])[c];
    previewColWidths.push(colInfo && colInfo.wpx ? Math.max(20, Math.round(colInfo.wpx)) : 72);
  }

  return {
    sheetName, propertyName, location, reportDate, units, sqft, purchaseDate, purchasePrice, purchasePriceNum, purchasePriceAddr, months, year, period,
    totalCost, totalCostNum, totalCostAddr, debtText, debtAddr, loanBalanceNum,
    ytdIncome, ytdNetIncome, ytdNetCashFlow, ytdNOI, ytdDebtService, dscr, annualizedNOI, capRate, yieldOnCost, debtYield, hasFlaggedFinancials,
    avgOccupancy, totalMoveOuts, flaggedMonthCount, reportedMonthCount: reportedMonths.length,
    allFlags, businessImpactTotal, criticalCount,
    reconciliation, isReconciled, reconciliationStatus, dataQualityNotes,
    previewGrid, previewMerges, previewColWidths, previewLastCol: previewColCount - 1,
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
  { key: "netCashFlow", label: "Net cash flow" },
];

// Rule-based triage, computed synchronously from hard numbers (flags, DSCR, occupancy) so the
// portfolio view doesn't have to wait on the AI call to tell the owner where to look first.
function computePropertyHealth(latestRep) {
  if (!latestRep) return "healthy";
  const reportedMonths = latestRep.months.filter((mo) => !mo.notYetReported);
  const lastMonth = reportedMonths[reportedMonths.length - 1];
  const priorMonth = reportedMonths[reportedMonths.length - 2];
  const hasCriticalFlag = lastMonth && lastMonth.flags.some((f) => f.severity === "critical");
  const dscr = latestRep.dscr;
  // Use the current month's occupancy, not the yearly average \u2014 a property that ran 95% for
  // ten months and dropped to 81% last month averages ~93% ("healthy") but the current reality is
  // exactly the thing triage should catch. A sharp month-over-month drop counts even if the level
  // alone still looks fine.
  const currentOcc = lastMonth ? lastMonth.occupiedPct : NaN;
  const priorOcc = priorMonth ? priorMonth.occupiedPct : NaN;
  const occDropped = !isNaN(currentOcc) && !isNaN(priorOcc) && (priorOcc - currentOcc) > 0.05;
  if (hasCriticalFlag || (!isNaN(dscr) && dscr < 1.0) || (!isNaN(currentOcc) && currentOcc < 0.80)) return "immediate";
  if (latestRep.flaggedMonthCount > 0 || (!isNaN(dscr) && dscr < 1.2) || (!isNaN(currentOcc) && currentOcc < 0.90) || occDropped) return "review";
  return "healthy";
}
const HEALTH_META = {
  immediate: { label: "Requires immediate action", rank: 0, color: BAD_COLOR },
  review: { label: "Needs review", rank: 1, color: FLAG_AMBER },
  healthy: { label: "Healthy", rank: 2, color: GOOD_COLOR },
};

const PERIOD_BASES = [
  { key: "MTD", label: "MTD" },
  { key: "QTD", label: "QTD" },
  { key: "H1", label: "H1" },
  { key: "H2", label: "H2" },
  { key: "YTD", label: "YTD" },
  { key: "TTM", label: "TTM" },
  { key: "FYE", label: "Annual (FYE)" },
];

// Flattens every reported month across all uploaded years for one property into chronological
// order, so period-basis math (QTD, H1/H2, TTM, FYE) can work the same whether one year or
// several years are loaded, and can correctly span a year boundary for TTM.
function flattenReportedMonths(reportsForProperty) {
  const flat = [];
  const seenAbsIndex = new Set();
  [...reportsForProperty].sort((a, b) => (a.year || 0) - (b.year || 0)).forEach((r) => {
    r.months.filter((mo) => !mo.notYetReported).forEach((mo) => {
      const absIndex = (r.year || 0) * 12 + MONTHS.indexOf(mo.key);
      // Guard against the same calendar month entering twice \u2014 e.g. a full-year report and an
      // overlapping partial-year report for the same property both loaded at once \u2014 which would
      // silently double-count that month in any trailing-window sum (TTM, T-3, T-6).
      if (seenAbsIndex.has(absIndex)) return;
      seenAbsIndex.add(absIndex);
      flat.push({
        year: r.year || 0, monthIdx: MONTHS.indexOf(mo.key), absIndex,
        monthLabel: mo.label, label: mo.label, noi: mo.noi, income: mo.income, opEx: mo.opEx, netIncome: mo.netIncome, netCashFlow: mo.netCashFlow, debtService: mo.debtService,
        occupiedPct: mo.occupiedPct, moveOuts: mo.moveOuts, netExposure: mo.netExposure, hasFlag: mo.flags.length > 0, flags: mo.flags, src: mo.src, rep: r,
      });
    });
  });
  return flat.sort((a, b) => a.absIndex - b.absIndex);
}

function sumField(entries, field) {
  const valid = entries.filter((e) => !isNaN(e[field]));
  return valid.length ? valid.reduce((s, e) => s + e[field], 0) : NaN;
}
function avgField(entries, field) {
  const valid = entries.filter((e) => !isNaN(e[field]));
  return valid.length ? valid.reduce((s, e) => s + e[field], 0) / valid.length : NaN;
}

function computePeriodMetrics(flat, basisKey, purchasePriceRep) {
  if (flat.length === 0) return null;
  const purchasePriceNum = purchasePriceRep ? purchasePriceRep.purchasePriceNum : NaN;
  const latest = flat[flat.length - 1];
  const { year, monthIdx } = latest;
  let entries = [];
  let label = "";
  let annualizeForCapRate = false;

  if (basisKey === "MTD") {
    entries = [latest];
    label = `${MONTH_LABELS[MONTHS[monthIdx]]} ${year}`;
  } else if (basisKey === "QTD") {
    const qStart = Math.floor(monthIdx / 3) * 3;
    entries = flat.filter((e) => e.year === year && e.monthIdx >= qStart && e.monthIdx <= monthIdx);
    label = `Q${Math.floor(monthIdx / 3) + 1} ${year} to date`;
  } else if (basisKey === "H1") {
    entries = flat.filter((e) => e.year === year && e.monthIdx <= 5);
    label = `H1 ${year}${entries.length < 6 ? " (partial)" : ""}`;
  } else if (basisKey === "H2") {
    entries = flat.filter((e) => e.year === year && e.monthIdx >= 6);
    label = entries.length === 0 ? `H2 ${year} (not yet reported)` : `H2 ${year}${entries.length < 6 ? " (partial)" : ""}`;
  } else if (basisKey === "YTD") {
    entries = flat.filter((e) => e.year === year && e.monthIdx <= monthIdx);
    label = `YTD through ${MONTH_LABELS[MONTHS[monthIdx]]} ${year}`;
  } else if (basisKey === "TTM") {
    entries = flat.slice(-12);
    annualizeForCapRate = entries.length === 12;
    label = entries.length === 12 ? "Trailing 12 months" : `Trailing ${entries.length} months`;
  } else if (basisKey === "FYE") {
    entries = flat.filter((e) => e.year === year);
    annualizeForCapRate = entries.length === 12;
    label = entries.length === 12 ? `Full year ${year}` : `${year} (in progress, ${entries.length}/12 months)`;
  }

  const noi = sumField(entries, "noi");
  const netCashFlow = sumField(entries, "netCashFlow");
  const debtService = sumField(entries.filter((e) => !isNaN(e.debtService)), "debtService");
  const dscr = (!isNaN(debtService) && Math.abs(debtService) > 0) ? noi / Math.abs(debtService) : NaN;
  const occupancy = avgField(entries, "occupiedPct");
  const moveOuts = sumField(entries, "moveOuts");
  const hasFlag = entries.some((e) => e.hasFlag);
  let capRate = NaN;
  if (annualizeForCapRate && !isNaN(noi) && purchasePriceNum > 0) capRate = noi / purchasePriceNum;

  return { label, monthCount: entries.length, noi, netCashFlow, dscr, occupancy, moveOuts, hasFlag, capRate, capRateApplicable: annualizeForCapRate, entries, debtService, purchasePriceRep };
}

const OCC_ROWS = [
  { key: "occupiedPct", label: "Occupied %", isPct: true },
  { key: "moveOuts", label: "Move-outs" },
  { key: "vacantUnits", label: "Current vacant units" },
  { key: "netExposure", label: "Net exposure to vacancy" },
];

function makeClickableDot(color, findMode, onDotClick, radius) {
  return (props) => {
    const { cx, cy, payload } = props;
    if (cx == null || cy == null) return null;
    return (
      <circle
        cx={cx}
        cy={cy}
        r={findMode ? radius + 2 : radius}
        fill={color}
        stroke={findMode ? "#FFFFFF" : "none"}
        strokeWidth={findMode ? 1.5 : 0}
        style={{ cursor: findMode ? "zoom-in" : "default" }}
        onClick={(e) => { e.stopPropagation(); if (findMode) onDotClick(payload); }}
      />
    );
  };
}

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

function BulletList({ items }) {
  const list = Array.isArray(items) ? items : [items];
  return (
    <ul style={{ margin: 0, paddingLeft: 16, listStyle: "disc" }}>
      {list.filter(Boolean).map((item, i) => (
        <li key={i} style={{ fontSize: 13.5, lineHeight: 1.5, marginBottom: i < list.length - 1 ? 4 : 0 }}>{item}</li>
      ))}
    </ul>
  );
}

function DecisionSupportCard({ decision, onRetry }) {
  if (!decision) return null;
  const d = decision.data;
  const careColor = !d ? SLATE : d.shouldICare === "Yes" ? BAD_COLOR : d.shouldICare === "Watch" ? FLAG_AMBER : GOOD_COLOR;
  const careLabel = !d ? null : d.shouldICare === "Yes" ? "Should I care: Yes" : d.shouldICare === "Watch" ? "Worth watching" : "On track";
  return (
    <div style={{ background: CARD, border: `0.5px solid ${BORDER}`, borderRadius: 12, padding: "1.25rem 1.5rem", marginBottom: 20 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14, flexWrap: "wrap", gap: 8 }}>
        <div style={{ fontFamily: "'Source Serif 4', serif", fontWeight: 600, fontSize: 15, color: INK }}>Decision support</div>
        {careLabel && (
          <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 11, fontWeight: 600, letterSpacing: "0.06em", textTransform: "uppercase", color: careColor, border: `1px solid ${careColor}`, borderRadius: 20, padding: "3px 10px" }}>
            {careLabel}
          </div>
        )}
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "14px 28px" }}>
        <div>
          <div style={{ fontSize: 11, color: SLATE, textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 4 }}>What changed</div>
          <BulletList items={decision.whatChanged} />
          <div style={{ fontSize: 10, color: SLATE, marginTop: 6 }}>Computed directly from the figures, not AI-generated.</div>
        </div>
        <div>
          <div style={{ fontSize: 11, color: SLATE, textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 4 }}>Why</div>
          {decision.status === "loading" && (
            <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
              <Loader2 size={13} className="spin-icon" color={DEED_GREEN} />
              <span style={{ fontSize: 12, color: SLATE }}>Analyzing{"\u2026"}</span>
            </div>
          )}
          {decision.status === "error" && (
            <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
              <span style={{ fontSize: 12, color: SLATE }} title={decision.errorMessage || undefined}>Not available.</span>
              <button onClick={onRetry} style={{ background: "transparent", border: `1px solid ${BORDER}`, color: INK, borderRadius: 6, padding: "3px 9px", fontSize: 11, cursor: "pointer" }}>Retry</button>
            </div>
          )}
          {decision.status === "done" && <BulletList items={d.whyItChanged} />}
        </div>
        <div style={{ gridColumn: "1 / -1", paddingTop: 4, borderTop: `0.5px solid ${BORDER}` }}>
          <div style={{ fontSize: 11, color: SLATE, textTransform: "uppercase", letterSpacing: "0.05em", margin: "10px 0 4px" }}>What to do</div>
          {decision.status === "done" ? (
            <div style={{ fontWeight: 500 }}><BulletList items={d.actionToTake} /></div>
          ) : (
            <div style={{ fontSize: 13, color: SLATE }}>{decision.status === "loading" ? "\u2014" : "Not available."}</div>
          )}
        </div>
      </div>
      {d && d._backtestWarning && (
        <div style={{ fontSize: 12, color: FLAG_AMBER, marginTop: 12, display: "flex", gap: 6, alignItems: "flex-start" }}>
          <AlertTriangle size={13} style={{ marginTop: 1, flexShrink: 0 }} />
          <span>{d._backtestWarning}</span>
        </div>
      )}
      <div style={{ fontSize: 11, color: SLATE, marginTop: 14 }}>"Why" and "What to do" are AI-generated from the figures above {"\u2014"} useful as a starting read, not a substitute for judgment. "What changed" is not.</div>
    </div>
  );
}

function PropertyOverviewCard({ overview, onRetry }) {
  if (!overview || overview.status === "loading") {
    return (
      <div style={{ background: CARD, border: `0.5px solid ${BORDER}`, borderRadius: 12, padding: "1.1rem 1.5rem", marginBottom: 28, display: "flex", alignItems: "center", gap: 10 }}>
        <Loader2 size={16} className="spin-icon" color={DEED_GREEN} />
        <span style={{ fontSize: 13, color: SLATE }}>Generating multi-year performance report{"\u2026"}</span>
      </div>
    );
  }
  if (overview.status === "error") {
    return (
      <div style={{ background: CARD, border: `0.5px solid ${BORDER}`, borderRadius: 12, padding: "1.1rem 1.5rem", marginBottom: 28, display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, flexWrap: "wrap" }}>
        <div>
          <div style={{ fontSize: 13, color: SLATE }}>Multi-year performance report wasn't available.</div>
          {overview.errorMessage && <div style={{ fontSize: 11, color: SLATE, marginTop: 3, fontFamily: "'IBM Plex Mono', monospace" }}>{overview.errorMessage}</div>}
        </div>
        <button onClick={onRetry} style={{ background: "transparent", border: `1px solid ${BORDER}`, color: INK, borderRadius: 6, padding: "6px 12px", fontSize: 12, cursor: "pointer" }}>Retry</button>
      </div>
    );
  }
  const d = overview.data;
  const verdictColor = d.overallVerdict === "Improving" ? GOOD_COLOR : d.overallVerdict === "Declining" ? BAD_COLOR : FLAG_AMBER;
  return (
    <div style={{ background: CARD, border: `0.5px solid ${BORDER}`, borderRadius: 12, padding: "1.35rem 1.6rem", marginBottom: 28 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16, flexWrap: "wrap", gap: 8 }}>
        <div style={{ fontFamily: "'Source Serif 4', serif", fontWeight: 600, fontSize: 16, color: INK }}>Multi-year performance report</div>
        <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 11, fontWeight: 600, letterSpacing: "0.06em", textTransform: "uppercase", color: verdictColor, border: `1px solid ${verdictColor}`, borderRadius: 20, padding: "3px 10px" }}>
          {d.overallVerdict}
        </div>
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "1fr", gap: 16 }}>
        <div>
          <div style={{ fontSize: 11, color: SLATE, textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 5 }}>Is the property manager doing a good job?</div>
          <BulletList items={d.propertyManagerAssessment} />
        </div>
        <div style={{ paddingTop: 12, borderTop: `0.5px solid ${BORDER}` }}>
          <div style={{ fontSize: 11, color: SLATE, textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 5 }}>Is the asset performing?</div>
          <BulletList items={d.assetPerformanceAssessment} />
        </div>
        <div style={{ paddingTop: 12, borderTop: `0.5px solid ${BORDER}` }}>
          <div style={{ fontSize: 11, color: SLATE, textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 5 }}>Does this change the valuation?</div>
          <BulletList items={d.valuationImpact} />
        </div>
      </div>
      {d._backtestWarning && (
        <div style={{ fontSize: 12, color: FLAG_AMBER, marginTop: 12, display: "flex", gap: 6, alignItems: "flex-start" }}>
          <AlertTriangle size={13} style={{ marginTop: 1, flexShrink: 0 }} />
          <span>{d._backtestWarning}</span>
        </div>
      )}
      <div style={{ fontSize: 11, color: SLATE, marginTop: 16 }}>AI-generated from the standardized year-by-year figures above {"\u2014"} useful as a starting read, not a substitute for judgment.</div>
    </div>
  );
}

export default function MonthlyStatusReportComparator() {
  const [reports, setReports] = useState([]);
  const [selectedId, setSelectedId] = useState(null);
  const [dragOver, setDragOver] = useState(false);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [decisions, setDecisions] = useState({});
  const [propertyOverviews, setPropertyOverviews] = useState({});
  const [activeProperty, setActiveProperty] = useState(null);
  const [findMode, setFindMode] = useState(false);
  const [previewTarget, setPreviewTarget] = useState(null);
  const [showComparisonTable, setShowComparisonTable] = useState(false);
  const [toast, setToast] = useState(null);
  const toastTimerRef = useRef(null);

  const showToast = (message) => {
    if (toastTimerRef.current) clearTimeout(toastTimerRef.current);
    setToast(message);
    toastTimerRef.current = setTimeout(() => setToast(null), 2600);
  };
  const [expandedYearData, setExpandedYearData] = useState({});
  const [periodBasis, setPeriodBasis] = useState("YTD");
  const requestedDecisions = useRef(new Set());
  const requestedOverviews = useRef(new Set());
  const previewCellRef = useRef(null);

  useEffect(() => {
    const onKeyDown = (e) => {
      if (e.key === "Escape") { setFindMode(false); setPreviewTarget(null); }
    };
    window.addEventListener("keydown", onKeyDown);
    return () => window.removeEventListener("keydown", onKeyDown);
  }, []);

  useEffect(() => {
    if (previewTarget && previewCellRef.current) {
      previewCellRef.current.scrollIntoView({ block: "center", inline: "center" });
    }
  }, [previewTarget]);

  const generateDecision = useCallback(async (rep, allReportsForProperty) => {
    const reportedMonths = rep.months.filter((mo) => !mo.notYetReported);
    const latestMonth = reportedMonths[reportedMonths.length - 1] || null;
    // Use the same cross-year flattened list the variance table uses, so a January report compares
    // against December of the prior year instead of reporting "no prior month" just because the
    // prior month lives in a different uploaded file.
    let latest = latestMonth, prior = null;
    if (latestMonth) {
      const flat = flattenReportedMonths(allReportsForProperty && allReportsForProperty.length ? allReportsForProperty : [rep]);
      const latestAbsIndex = (rep.year || 0) * 12 + MONTHS.indexOf(latestMonth.key);
      const idx = flat.findIndex((e) => e.absIndex === latestAbsIndex);
      if (idx >= 0) {
        latest = flat[idx];
        prior = idx > 0 ? flat[idx - 1] : null;
      }
    }
    const whatChanged = computeWhatChanged(latest, prior);
    setDecisions((prev) => ({ ...prev, [rep.id]: { status: "loading", whatChanged } }));
    try {
      const recentFlags = reportedMonths.slice(-3).flatMap((mo) => mo.flags.map((f) => `${mo.label}: ${f.message}`));

      const summary = {
        property: rep.propertyName,
        location: rep.location,
        units: rep.units,
        reportingPeriod: rep.period,
        latestReportedMonth: latest ? { month: latest.label, noi: latest.noi, income: latest.income, operatingExpenses: latest.opEx, netIncome: latest.netIncome, occupiedPct: latest.occupiedPct, netExposureToVacancy: latest.netExposure, moveOuts: latest.moveOuts } : null,
        priorReportedMonth: prior ? { month: prior.label, noi: prior.noi, income: prior.income, operatingExpenses: prior.opEx, netIncome: prior.netIncome, occupiedPct: prior.occupiedPct } : null,
        ytdNOI: rep.ytdNOI, ytdIncome: rep.ytdIncome, ytdNetIncome: rep.ytdNetIncome,
        avgOccupancy: rep.avgOccupancy, dscrYTD: rep.dscr,
        yieldOnCostEstimate: rep.capRate,
        note: "yieldOnCostEstimate is NOI over purchase price, i.e. yield on cost \u2014 NOT a market cap rate. Never call it a cap rate.",
        yieldOnCostAndNOIMayBeDistortedByFlaggedMonth: rep.hasFlaggedFinancials,
        dataQualityFlagsLast3Months: recentFlags,
        whatChangedThisMonth: whatChanged,
      };

      const prompt = `You are a multifamily real estate asset management analyst helping a property OWNER (not a lender or buyer evaluating a deal) quickly understand their own property's monthly status report. The "what changed" numbers have already been computed deterministically in code (see whatChangedThisMonth below) \u2014 do not restate or recompute them, just use them as context for your interpretation. Here is standardized data extracted from the report:

${JSON.stringify(summary, null, 2)}

Respond with ONLY a raw JSON object, no markdown fences, no commentary, with exactly these keys:
{"whyItChanged": ["1-3 short bullet strings on the likely driver behind whatChangedThisMonth, grounded only in the numbers given, no speculation beyond the data, numbers kept"], "shouldICare": "Yes, No, or Watch", "actionToTake": ["1-2 short bullet strings, concrete next steps for the owner, or a single bullet 'No action needed \u2014 performance looks on track' if nothing stands out"]}

Each bullet should read like a terse analyst note, not a sentence in a paragraph \u2014 short, direct, numbers intact, no throat-clearing ("It appears that...", "This suggests..."). Never use the term "cap rate" \u2014 use "yield on cost" for yieldOnCostEstimate.

If yieldOnCostAndNOIMayBeDistortedByFlaggedMonth is true, add a bullet in whyItChanged noting the YTD NOI/yield-on-cost figure includes a flagged month and should be sanity-checked before relying on it.`;

      const responseText = await callClaudeText(prompt, 1000);
      const cleaned = responseText.replace(/```json|```/g, "").trim();
      const parsed = scrubDeep(JSON.parse(cleaned));
      // Backtest: cross-check the AI's "shouldICare" against the hard, rule-based flags we already
      // trust. If the model said "No action needed" on a month that actually has a data-quality or
      // performance flag, surface that mismatch rather than silently trusting the narrative.
      const latestHasFlags = latest && latest.flags.length > 0;
      parsed._backtestWarning = latestHasFlags && parsed.shouldICare === "No"
        ? "This month has a flagged data point, but the analysis said no action needed \u2014 worth a manual check."
        : null;
      setDecisions((prev) => ({ ...prev, [rep.id]: { status: "done", whatChanged, data: parsed } }));
    } catch (err) {
      setDecisions((prev) => ({ ...prev, [rep.id]: { status: "error", whatChanged, errorMessage: err && err.message } }));
    }
  }, []);

  useEffect(() => {
    reports.forEach((rep) => {
      if (!requestedDecisions.current.has(rep.id)) {
        requestedDecisions.current.add(rep.id);
        const siblings = reports.filter((r) => (r.propertyName || r.fileName) === (rep.propertyName || rep.fileName));
        generateDecision(rep, siblings);
      }
    });
  }, [reports, generateDecision]);

  const generatePropertyOverview = useCallback(async (groupKey, propertyName, yearlyReports) => {
    setPropertyOverviews((prev) => ({ ...prev, [groupKey]: { status: "loading" } }));
    try {
      const yearlySummaries = yearlyReports.map((r) => ({
        year: r.year || r.period,
        reportingPeriod: r.period,
        annualizedNOI: r.annualizedNOI,
        avgMonthlyIncome: r.reportedMonthCount ? r.ytdIncome / r.reportedMonthCount : null,
        avgOccupancy: r.avgOccupancy,
        totalMoveOuts: r.totalMoveOuts,
        dscrYTD: r.dscr,
        yieldOnCostEstimate: r.capRate,
        hasFlaggedFinancials: r.hasFlaggedFinancials,
        monthsReported: r.reportedMonthCount,
      }));

      const summary = {
        property: propertyName, units: yearlyReports[0]?.units, yearlySummaries,
        note: "yieldOnCostEstimate is NOI over purchase price, i.e. yield on cost \u2014 NOT a market cap rate. No external market benchmark data is provided; base the assessment only on this property's own figures.",
      };

      const prompt = `You are a multifamily real estate asset management analyst preparing a multi-year performance summary for a property OWNER (long-term hold, family office \u2014 not evaluating a sale). Here is year-by-year standardized data for one property, drawn from multiple uploaded monthly status reports:

${JSON.stringify(summary, null, 2)}

Respond with ONLY a raw JSON object, no markdown fences, no commentary, with exactly these keys:
{"propertyManagerAssessment": ["2-3 short bullet strings on whether the property manager appears to be doing a good job, based on trends in income/rent, occupancy, and turnover (move-outs) across the years given \u2014 grounded only in this data, every number kept"], "assetPerformanceAssessment": ["2-3 short bullet strings on whether the asset is performing well overall, based on the NOI and occupancy trend across years, numbers kept"], "valuationImpact": ["1-2 short bullet strings on whether these trends suggest the property's value is likely improving, stable, or at risk, referencing the yield-on-cost/NOI trend \u2014 note explicitly if a year's figures are flagged and may be unreliable"], "overallVerdict": "Improving, Stable, or Declining"}

Each bullet should read like a terse analyst note, not a sentence in a paragraph \u2014 short, direct, every specific number/dollar figure/percent kept, no throat-clearing. Never use the term "cap rate" \u2014 use "yield on cost".`;

      const responseText = await callClaudeText(prompt, 1000);
      const cleaned = responseText.replace(/```json|```/g, "").trim();
      const parsed = scrubDeep(JSON.parse(cleaned));
      // Backtest: check the AI's verdict against the actual NOI trend across the uploaded years,
      // rather than trusting the label at face value.
      const validNOIs = yearlySummaries.filter((y) => !isNaN(y.annualizedNOI)).map((y) => y.annualizedNOI);
      if (validNOIs.length >= 2) {
        const trendUp = validNOIs[validNOIs.length - 1] > validNOIs[0] * 1.03;
        const trendDown = validNOIs[validNOIs.length - 1] < validNOIs[0] * 0.97;
        const mismatch = (parsed.overallVerdict === "Improving" && trendDown) || (parsed.overallVerdict === "Declining" && trendUp);
        parsed._backtestWarning = mismatch ? `The verdict "${parsed.overallVerdict}" doesn't clearly match the annualized NOI trend from the first to most recent year \u2014 worth a manual look.` : null;
      }
      setPropertyOverviews((prev) => ({ ...prev, [groupKey]: { status: "done", data: parsed } }));
    } catch (err) {
      setPropertyOverviews((prev) => ({ ...prev, [groupKey]: { status: "error", errorMessage: err && err.message } }));
    }
  }, []);

  const parseFile = (file) => new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: "array", cellDates: true, cellStyles: true });
        const result = processWorkbook(workbook, file.name);
        if (result.months.length === 0) { reject(`${file.name}: couldn't find a monthly financial section`); return; }
        resolve({ id: `${file.name}-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`, fileName: file.name, fileURL: URL.createObjectURL(file), ...result });
      } catch (err) {
        reject(`${file.name}: couldn't read this file`);
      }
    };
    reader.onerror = () => reject(`${file.name}: couldn't read this file`);
    reader.readAsArrayBuffer(file);
  });

  const dedupKey = (r) => `${(r.propertyName || r.fileName || "").trim().toLowerCase()}::${r.year || "unknown"}`;

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
    const parsed = results.filter((r) => r.status === "fulfilled").map((r) => r.value);
    const failed = results.filter((r) => r.status === "rejected").map((r) => r.reason);

    const existingKeys = new Set(reports.map(dedupKey));
    const toAdd = [];
    let duplicateCount = 0;
    parsed.forEach((r) => {
      const key = dedupKey(r);
      if (existingKeys.has(key)) {
        duplicateCount++;
        if (r.fileURL) URL.revokeObjectURL(r.fileURL);
      } else {
        existingKeys.add(key);
        toAdd.push(r);
      }
    });

    if (toAdd.length > 0) {
      setReports((prev) => [...prev, ...toAdd]);
      setSelectedId(toAdd[toAdd.length - 1].id);
    }
    if (duplicateCount > 0) {
      showToast(duplicateCount === 1 ? "File already uploaded" : `${duplicateCount} files already uploaded`);
    }
    if (failed.length > 0) setError(failed.join("; "));
    setLoading(false);
  }, [reports]);

  const removeReport = (id) => {
    setReports((prev) => {
      const removed = prev.find((r) => r.id === id);
      if (removed && removed.fileURL) URL.revokeObjectURL(removed.fileURL);
      const next = prev.filter((r) => r.id !== id);
      if (selectedId === id) setSelectedId(next.length ? next[next.length - 1].id : null);
      return next;
    });
    setDecisions((prev) => { const next = { ...prev }; delete next[id]; return next; });
    requestedDecisions.current.delete(id);
  };

  const openCellPreview = (rep, addr) => {
    if (!addr || !rep.previewGrid) return;
    const { r, c } = XLSX.utils.decode_cell(addr);
    setPreviewTarget({ mode: "cell", fileName: rep.fileName, sheetName: rep.sheetName, grid: rep.previewGrid, merges: rep.previewMerges, colWidths: rep.previewColWidths, lastCol: rep.previewLastCol, targetRow: r, targetCol: c, addr });
    setFindMode(false);
  };

  // Opens a calculation breakdown for any derived/aggregate number (period totals, DSCR, cap
  // rate, chart bars) showing exactly which source cells across which uploaded files were summed
  // or divided to produce it, each one clickable through to the raw sheet preview.
  const openCalcPreview = ({ title, formula, rows, note }) => {
    setPreviewTarget({ mode: "calc", title, formula, rows, note });
    setFindMode(false);
  };

  const buildFieldRows = (entries, field, formatFn) => entries.map((e) => ({
    label: `${e.monthLabel} ${e.year}`,
    value: formatFn(e[field]),
    raw: e[field],
    fileName: e.rep.fileName,
    sheetName: e.rep.sheetName,
    addr: e.src ? e.src[field] : null,
    rep: e.rep,
  })).filter((r) => !isNaN(r.raw));

  const propertyGroups = {};
  reports.forEach((r) => {
    const key = r.propertyName || r.fileName;
    if (!propertyGroups[key]) propertyGroups[key] = [];
    propertyGroups[key].push(r);
  });
  const propertyHealth = {};
  const propertyImpact = {};
  const propertyLatest = {};
  Object.entries(propertyGroups).forEach(([key, list]) => {
    const latest = [...list].sort((a, b) => (a.year || 0) - (b.year || 0)).slice(-1)[0];
    propertyHealth[key] = computePropertyHealth(latest);
    propertyImpact[key] = latest.businessImpactTotal || 0;
    propertyLatest[key] = latest;
  });
  // Sort by estimated dollar impact first \u2014 the owner should see which asset needs him first,
  // not scroll to find it alphabetically. Health rank breaks ties among similarly-sized impacts.
  const propertyKeys = Object.keys(propertyGroups).sort((a, b) => {
    const impactDiff = propertyImpact[b] - propertyImpact[a];
    if (Math.abs(impactDiff) > 1) return impactDiff;
    const rankDiff = HEALTH_META[propertyHealth[a]].rank - HEALTH_META[propertyHealth[b]].rank;
    return rankDiff !== 0 ? rankDiff : a.localeCompare(b);
  });
  const healthCounts = { immediate: 0, review: 0, healthy: 0 };
  Object.values(propertyHealth).forEach((h) => { healthCounts[h]++; });
  const portfolioTotalImpact = Object.values(propertyImpact).reduce((s, v) => s + v, 0);

  useEffect(() => {
    if (propertyKeys.length === 0) { setActiveProperty(null); return; }
    if (!activeProperty || !propertyKeys.includes(activeProperty)) setActiveProperty(propertyKeys[0]);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [reports]);

  const activeReports = [...(propertyGroups[activeProperty] || [])].sort((a, b) => (a.year || 0) - (b.year || 0));
  const latestActiveReport = activeReports[activeReports.length - 1] || null;
  const flatForVariance = activeReports.length > 0 ? flattenReportedMonths(activeReports) : [];
  let monthVariance = null;
  if (flatForVariance.length >= 2) {
    const latest = flatForVariance[flatForVariance.length - 1];
    const prior = flatForVariance[flatForVariance.length - 2];
    monthVariance = {
      latestLabel: `${latest.monthLabel} ${latest.year}`,
      priorLabel: `${prior.monthLabel} ${prior.year}`,
      rows: [
        { label: "NOI", latestVal: latest.noi, priorVal: prior.noi, isMoney: true, addr: latest.src ? latest.src.noi : null, rep: latest.rep },
        { label: "Occupancy", latestVal: latest.occupiedPct, priorVal: prior.occupiedPct, isPct: true, addr: latest.src ? latest.src.occupiedPct : null, rep: latest.rep },
        { label: "Net Cash Flow", latestVal: latest.netCashFlow, priorVal: prior.netCashFlow, isMoney: true, addr: latest.src ? latest.src.netCashFlow : null, rep: latest.rep },
      ],
    };
  }
  const monthlyTrendData = latestActiveReport ? latestActiveReport.months.filter((mo) => !mo.notYetReported).map((mo) => ({
    month: mo.label,
    NOI: Math.round(mo.noi || 0),
    "Net Cash Flow": isNaN(mo.netCashFlow) ? null : Math.round(mo.netCashFlow),
    "Occupancy %": isNaN(mo.occupiedPct) ? null : Math.round(mo.occupiedPct * 1000) / 10,
    _addr: mo.src.noi, _cashAddr: mo.src.netCashFlow, _occAddr: mo.src.occupiedPct, _rep: latestActiveReport,
  })) : [];
  const overviewGroupKey = activeProperty ? `${activeProperty}::${activeReports.map((r) => r.id).sort().join(",")}` : null;

  useEffect(() => {
    if (overviewGroupKey && activeReports.length >= 2 && !requestedOverviews.current.has(overviewGroupKey)) {
      requestedOverviews.current.add(overviewGroupKey);
      generatePropertyOverview(overviewGroupKey, activeProperty, activeReports);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [overviewGroupKey]);

  const compLabel = (r) => `${r.propertyName || r.fileName}${r.period ? " \u2014 " + r.period : ""}`;

  const buildReportFieldRows = (r, field, formatFn) => r.months.filter((mo) => !mo.notYetReported).map((mo) => ({
    label: `${mo.label} ${r.year}`, value: formatFn(mo[field]), raw: mo[field],
    fileName: r.fileName, sheetName: r.sheetName, addr: mo.src[field], rep: r,
  })).filter((row) => !isNaN(row.raw));

  const comparisonRows = [
    { label: "Location", get: (r) => r.location || "\u2014", isText: true },
    { label: "Reporting period", get: (r) => r.period || "\u2014", isText: true },
    {
      label: "YTD NOI", get: (r) => formatMoney(r.ytdNOI) + (r.hasFlaggedFinancials ? " *" : ""),
      calc: (r) => {
        const rows = buildReportFieldRows(r, "noi", formatMoney);
        openCalcPreview({ title: `YTD NOI \u2014 ${compLabel(r)}`, formula: `Sum of Net Operating Income across ${rows.length} month(s)`, rows: [...rows, { label: "Total YTD NOI", value: formatMoney(r.ytdNOI), isTotal: true }] });
      },
    },
    { label: "Avg occupancy", get: (r) => isNaN(r.avgOccupancy) ? "\u2014" : `${(r.avgOccupancy * 100).toFixed(1)}%` },
    {
      label: "DSCR (YTD)", get: (r) => isNaN(r.dscr) ? "\u2014" : `${r.dscr.toFixed(2)}x${r.hasFlaggedFinancials ? " *" : ""}`,
      calc: (r) => {
        if (isNaN(r.dscr)) return;
        const noiRows = buildReportFieldRows(r, "noi", formatMoney);
        const debtRows = buildReportFieldRows(r, "debtService", (v) => formatMoney(Math.abs(v)));
        openCalcPreview({
          title: `DSCR (YTD) \u2014 ${compLabel(r)}`,
          formula: "DSCR = Net Operating Income \u00f7 Debt Service",
          rows: [
            { label: "Net operating income", isSection: true }, ...noiRows, { label: "Total NOI", value: formatMoney(r.ytdNOI), isTotal: true },
            { label: "Debt service", isSection: true }, ...debtRows, { label: "Total debt service", value: formatMoney(r.ytdDebtService), isTotal: true },
            { label: "DSCR = NOI \u00f7 Debt service", value: `${r.dscr.toFixed(2)}x`, isTotal: true },
          ],
        });
      },
    },
    {
      label: "Yield on cost (purchase price)", get: (r) => isNaN(r.capRate) ? "\u2014" : `${(r.capRate * 100).toFixed(2)}%${r.hasFlaggedFinancials ? " *" : ""}`,
      calc: (r) => {
        if (isNaN(r.capRate)) return;
        const noiRows = buildReportFieldRows(r, "noi", formatMoney);
        openCalcPreview({
          title: `Yield on cost (purchase price) \u2014 ${compLabel(r)}`,
          formula: "Yield on cost = Annualized NOI \u00f7 Purchase price. This is NOT a market cap rate \u2014 that requires a current valuation, which isn't available here.",
          rows: [
            { label: "Net operating income (annualized from YTD)", isSection: true }, ...noiRows, { label: "Annualized NOI", value: formatMoney(r.annualizedNOI), isTotal: true },
            { label: "Purchase price", isSection: true },
            { label: "Purchase price", value: r.purchasePrice || formatMoney(r.purchasePriceNum), fileName: r.fileName, sheetName: r.sheetName, addr: r.purchasePriceAddr, rep: r },
            { label: "Yield on cost = Annualized NOI \u00f7 Purchase price", value: `${(r.capRate * 100).toFixed(2)}%`, isTotal: true },
          ],
        });
      },
    },
    {
      label: "Yield on cost (total cost)", get: (r) => isNaN(r.yieldOnCost) ? "\u2014 (no total cost on file)" : `${(r.yieldOnCost * 100).toFixed(2)}%${r.hasFlaggedFinancials ? " *" : ""}`,
      calc: (r) => {
        if (isNaN(r.yieldOnCost)) return;
        const noiRows = buildReportFieldRows(r, "noi", formatMoney);
        openCalcPreview({
          title: `Yield on cost (total cost) \u2014 ${compLabel(r)}`,
          formula: "Yield on cost = Annualized NOI \u00f7 Total cost basis (purchase price + capital improvements)",
          rows: [
            { label: "Net operating income (annualized from YTD)", isSection: true }, ...noiRows, { label: "Annualized NOI", value: formatMoney(r.annualizedNOI), isTotal: true },
            { label: "Total cost basis", isSection: true },
            { label: "Total cost", value: r.totalCost || formatMoney(r.totalCostNum), fileName: r.fileName, sheetName: r.sheetName, addr: r.totalCostAddr, rep: r },
            { label: "Yield on cost = Annualized NOI \u00f7 Total cost", value: `${(r.yieldOnCost * 100).toFixed(2)}%`, isTotal: true },
          ],
        });
      },
    },
    {
      label: "Debt yield", get: (r) => isNaN(r.debtYield) ? "\u2014 (no loan balance on file)" : `${(r.debtYield * 100).toFixed(2)}%${r.hasFlaggedFinancials ? " *" : ""}`,
      calc: (r) => {
        if (isNaN(r.debtYield)) return;
        const noiRows = buildReportFieldRows(r, "noi", formatMoney);
        openCalcPreview({
          title: `Debt yield \u2014 ${compLabel(r)}`,
          formula: "Debt yield = Annualized NOI \u00f7 Outstanding loan balance",
          rows: [
            { label: "Net operating income (annualized from YTD)", isSection: true }, ...noiRows, { label: "Annualized NOI", value: formatMoney(r.annualizedNOI), isTotal: true },
            { label: "Loan balance", isSection: true },
            { label: "Existing debt", value: r.debtText || formatMoney(r.loanBalanceNum), fileName: r.fileName, sheetName: r.sheetName, addr: r.debtAddr, rep: r },
            { label: "Debt yield = Annualized NOI \u00f7 Loan balance", value: `${(r.debtYield * 100).toFixed(2)}%`, isTotal: true },
          ],
        });
      },
    },
    { label: "NOI per unit (annualized)", get: (r) => (isNaN(r.annualizedNOI) || !r.units) ? "\u2014" : formatMoney(r.annualizedNOI / r.units), isText: false },
    { label: "NOI per sq ft (annualized)", get: (r) => (isNaN(r.annualizedNOI) || !r.sqft) ? "\u2014" : `$${(r.annualizedNOI / r.sqft).toFixed(2)}`, isText: false },
    { label: "YTD income", get: (r) => formatMoney(r.ytdIncome) },
    { label: "YTD net income", get: (r) => formatMoney(r.ytdNetIncome) },
    { label: "Total move-outs", get: (r) => isNaN(r.totalMoveOuts) ? "\u2014" : Math.round(r.totalMoveOuts).toLocaleString() },
    { label: "Months flagged", get: (r) => `${r.flaggedMonthCount} / ${r.reportedMonthCount}` },
    { label: "Reconciled to source totals", get: (r) => r.reconciliationStatus === "matches" ? "Yes" : r.reconciliationStatus === "mismatch" ? "No \u2014 see warning above" : "Not checked \u2014 no YTD column in source", isText: true },
  ];

  // Trend/TTM data for the active property tab only, when it has 2+ years uploaded.
  let activeTrend = null;
  if (activeReports.length >= 2) {
    const sortedList = activeReports;
    const flatMonths = flattenReportedMonths(sortedList);
    const last12 = flatMonths.slice(-12);
    const ttmNOI = last12.reduce((s, m) => s + (isNaN(m.noi) ? 0 : m.noi), 0);
    const ttmDebtService = last12.reduce((s, m) => s + (isNaN(m.debtService) ? 0 : Math.abs(m.debtService)), 0);
    const ttmDSCR = ttmDebtService > 0 ? ttmNOI / ttmDebtService : NaN;
    const latestPurchasePrice = [...sortedList].reverse().find((r) => !isNaN(r.purchasePriceNum))?.purchasePriceNum;
    const ttmCapRate = (last12.length === 12 && latestPurchasePrice > 0) ? ttmNOI / latestPurchasePrice : NaN;
    activeTrend = {
      isFullTTM: last12.length === 12,
      monthsUsed: last12.length,
      ttmNOI, ttmDSCR, ttmCapRate,
      data: sortedList.map((r) => ({
        year: String(r.year || r.period || r.fileName),
        NOI: Math.round(r.ytdNOI || 0),
        "Net Cash Flow": isNaN(r.ytdNetCashFlow) ? null : Math.round(r.ytdNetCashFlow),
        "Occupancy %": isNaN(r.avgOccupancy) ? null : Math.round(r.avgOccupancy * 1000) / 10,
        _rep: r,
      })),
    };
  }

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
            Upload one or more monthly status reports (.xls or .xlsx) \u2014 different properties, different years, or both. Each gets checked for reconciliation and unusual months. Multiple properties get their own tab, sorted by which needs attention first; multiple years for the same property compare side by side within its tab.
          </p>
        </div>

        {reports.length === 0 ? (
          <div
            onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
            onDragLeave={() => setDragOver(false)}
            onDrop={(e) => { e.preventDefault(); setDragOver(false); handleFiles(e.dataTransfer.files); }}
            style={{ border: `2px dashed ${dragOver ? DEED_GREEN : BORDER}`, background: dragOver ? HOVER_BG : CARD, borderRadius: 12, padding: "3.5rem 2rem", textAlign: "center", transition: "border-color 0.15s ease, background 0.15s ease" }}
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
            {error && <div style={{ fontSize: 13, color: ERROR_COLOR, marginBottom: 14 }}>{error}</div>}
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
              style={{ border: `1.5px dashed ${dragOver ? DEED_GREEN : BORDER}`, background: dragOver ? HOVER_BG : "transparent", borderRadius: 10, padding: "10px 16px", marginBottom: 20, display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12 }}
            >
              <div style={{ fontSize: 13, color: SLATE, display: "flex", alignItems: "center", gap: 8 }}>
                {loading && <Loader2 size={14} className="spin-icon" color={DEED_GREEN} />}
                {loading ? "Processing workbook(s)\u2026" : "Drop more reports here to add them to the comparison"}
              </div>
              <div style={{ display: "flex", gap: 8, flexWrap: "wrap", flexShrink: 0 }}>
                <label style={{ display: "inline-flex", alignItems: "center", gap: 6, background: "transparent", border: `1px solid ${DEED_GREEN}`, color: DEED_GREEN, fontSize: 13, fontWeight: 500, padding: "7px 14px", borderRadius: 8, cursor: "pointer", flexShrink: 0 }}>
                  <Plus size={14} /> Add files
                  <input type="file" accept=".xls,.xlsx" multiple style={{ display: "none" }} onChange={(e) => handleFiles(e.target.files)} />
                </label>
                <button
                  onClick={() => { setFindMode((v) => !v); setPreviewTarget(null); }}
                  style={{ display: "inline-flex", alignItems: "center", gap: 6, background: findMode ? DEED_GREEN : "transparent", border: `1px solid ${DEED_GREEN}`, color: findMode ? "#FFFFFF" : DEED_GREEN, fontSize: 13, fontWeight: 500, padding: "7px 14px", borderRadius: 8, cursor: "pointer", flexShrink: 0 }}
                >
                  <Search size={14} /> {findMode ? "Cancel find cell" : "Find cell"}
                </button>
              </div>
            </div>
            {findMode && (
              <div style={{ fontSize: 13, color: DEED_GREEN, background: HOVER_BG, border: `1px solid ${DEED_GREEN}`, borderRadius: 8, padding: "9px 14px", marginBottom: 16 }}>
                Click any value \u2014 in the tables, KPI cards, or charts \u2014 to see how it was calculated and locate it in the original spreadsheet(s). Press Esc to cancel.
              </div>
            )}
            {error && <div style={{ fontSize: 13, color: ERROR_COLOR, marginBottom: 16 }}>{error}</div>}

            {propertyKeys.length > 0 && (
              <div style={{ display: "flex", flexWrap: "wrap", alignItems: "center", gap: "8px 20px", marginBottom: 16, fontSize: 13 }}>
                <span style={{ fontFamily: "'Source Serif 4', serif", fontWeight: 600, fontSize: 15, color: INK }}>
                  {propertyKeys.length} {propertyKeys.length === 1 ? "Property" : "Properties"}
                </span>
                {healthCounts.immediate > 0 && (
                  <span style={{ display: "flex", alignItems: "center", gap: 6, color: BAD_COLOR }}>
                    <span style={{ width: 8, height: 8, borderRadius: "50%", background: BAD_COLOR, display: "inline-block" }} />
                    {healthCounts.immediate} Require Immediate Action
                  </span>
                )}
                {healthCounts.review > 0 && (
                  <span style={{ display: "flex", alignItems: "center", gap: 6, color: FLAG_AMBER }}>
                    <span style={{ width: 8, height: 8, borderRadius: "50%", background: FLAG_AMBER, display: "inline-block" }} />
                    {healthCounts.review} Need Review
                  </span>
                )}
                {healthCounts.healthy > 0 && (
                  <span style={{ display: "flex", alignItems: "center", gap: 6, color: GOOD_COLOR }}>
                    <span style={{ width: 8, height: 8, borderRadius: "50%", background: GOOD_COLOR, display: "inline-block" }} />
                    {healthCounts.healthy} Healthy
                  </span>
                )}
                {portfolioTotalImpact > 0 && (
                  <span style={{ color: SLATE, fontFamily: "'IBM Plex Mono', monospace", fontSize: 12 }}>
                    {"\u00b7"} {formatMoney(portfolioTotalImpact)} estimated business-impact total (excludes data-integrity flags)
                  </span>
                )}
              </div>
            )}

            {propertyKeys.length > 1 && (
              <div style={{ display: "flex", gap: 4, borderBottom: `1px solid ${BORDER}`, marginBottom: 24, overflowX: "auto" }}>
                {propertyKeys.map((key) => {
                  const isActive = key === activeProperty;
                  const health = HEALTH_META[propertyHealth[key]];
                  const impact = propertyImpact[key];
                  return (
                    <button
                      key={key}
                      onClick={() => setActiveProperty(key)}
                      style={{
                        background: isActive ? HOVER_BG : "none",
                        border: "none",
                        borderBottom: `3px solid ${health.color}`,
                        color: isActive ? INK : SLATE,
                        fontFamily: "'Source Serif 4', serif",
                        fontWeight: 600,
                        fontSize: 14,
                        padding: "10px 18px",
                        cursor: "pointer",
                        whiteSpace: "nowrap",
                        flexShrink: 0,
                        display: "flex",
                        flexDirection: "column",
                        alignItems: "flex-start",
                        gap: 3,
                        borderRadius: "6px 6px 0 0",
                      }}
                    >
                      <span style={{ display: "flex", alignItems: "center", gap: 8 }}>
                        <span style={{ width: 7, height: 7, borderRadius: "50%", background: health.color, flexShrink: 0 }} />
                        {key} <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontWeight: 400, fontSize: 11, color: SLATE }}>({propertyGroups[key].length})</span>
                      </span>
                      {impact > 0 && (
                        <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontWeight: 500, fontSize: 10.5, color: SLATE, paddingLeft: 15 }}>
                          {formatMoney(impact)} impact
                        </span>
                      )}
                    </button>
                  );
                })}
              </div>
            )}

            {activeProperty && (
              <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 20 }}>
                <span style={{ width: 8, height: 8, borderRadius: "50%", background: HEALTH_META[propertyHealth[activeProperty]].color, flexShrink: 0 }} />
                <span style={{ fontSize: 13, fontWeight: 600, color: HEALTH_META[propertyHealth[activeProperty]].color }}>{HEALTH_META[propertyHealth[activeProperty]].label}</span>
              </div>
            )}

            {activeReports.length > 0 && (() => {
              const flat = flattenReportedMonths(activeReports);
              const latestRep = activeReports[activeReports.length - 1];
              const purchasePriceRep = [...activeReports].reverse().find((r) => !isNaN(r.purchasePriceNum)) || null;
              const metrics = computePeriodMetrics(flat, periodBasis, purchasePriceRep);
              if (!metrics) return null;
              return (
                <>
                  <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginBottom: 14 }}>
                    {PERIOD_BASES.map((b) => (
                      <button
                        key={b.key}
                        onClick={() => setPeriodBasis(b.key)}
                        style={{
                          background: periodBasis === b.key ? DEED_GREEN : "transparent",
                          color: periodBasis === b.key ? "#FFFFFF" : SLATE,
                          border: `1px solid ${periodBasis === b.key ? DEED_GREEN : BORDER}`,
                          borderRadius: 20,
                          padding: "5px 13px",
                          fontSize: 12,
                          fontWeight: 600,
                          fontFamily: "'IBM Plex Mono', monospace",
                          cursor: "pointer",
                        }}
                      >
                        {b.label}
                      </button>
                    ))}
                  </div>
                  <div style={{ fontSize: 12, color: SLATE, marginBottom: 12 }}>{metrics.label}</div>
                  <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 12, marginBottom: 28 }}>
                    {[
                      {
                        label: "NOI", value: isNaN(metrics.noi) ? "\u2014" : formatMoney(metrics.noi), flagged: metrics.hasFlag,
                        onClick: () => {
                          const rows = buildFieldRows(metrics.entries, "noi", formatMoney);
                          openCalcPreview({ title: `NOI \u2014 ${metrics.label}`, formula: `Sum of Net Operating Income across ${rows.length} month(s)`, rows: [...rows, { label: "Total NOI", value: formatMoney(metrics.noi), isTotal: true }] });
                        },
                      },
                      {
                        label: "Occupancy (avg)", value: isNaN(metrics.occupancy) ? "\u2014" : `${(metrics.occupancy * 100).toFixed(1)}%`, flagged: false,
                        onClick: () => {
                          const rows = buildFieldRows(metrics.entries, "occupiedPct", (v) => `${(v * 100).toFixed(1)}%`);
                          openCalcPreview({ title: `Occupancy (avg) \u2014 ${metrics.label}`, formula: `Average of ${rows.length} month(s)`, rows: [...rows, { label: "Average occupancy", value: `${(metrics.occupancy * 100).toFixed(1)}%`, isTotal: true }] });
                        },
                      },
                      {
                        label: "Net cash flow", value: isNaN(metrics.netCashFlow) ? "\u2014" : formatMoney(metrics.netCashFlow), flagged: false,
                        onClick: () => {
                          const rows = buildFieldRows(metrics.entries, "netCashFlow", formatMoney);
                          openCalcPreview({ title: `Net cash flow \u2014 ${metrics.label}`, formula: `Sum of Net Cash Flow across ${rows.length} month(s)`, rows: [...rows, { label: "Total net cash flow", value: formatMoney(metrics.netCashFlow), isTotal: true }] });
                        },
                      },
                      {
                        label: "DSCR", value: isNaN(metrics.dscr) ? "\u2014" : `${metrics.dscr.toFixed(2)}x`, flagged: metrics.hasFlag,
                        onClick: () => {
                          const noiRows = buildFieldRows(metrics.entries, "noi", formatMoney);
                          const debtRows = buildFieldRows(metrics.entries, "debtService", (v) => formatMoney(Math.abs(v)));
                          openCalcPreview({
                            title: `DSCR \u2014 ${metrics.label}`,
                            formula: "DSCR = Net Operating Income \u00f7 Debt Service",
                            rows: [
                              { label: "Net operating income", isSection: true },
                              ...noiRows,
                              { label: "Total NOI", value: formatMoney(metrics.noi), isTotal: true },
                              { label: "Debt service", isSection: true },
                              ...debtRows,
                              { label: "Total debt service", value: formatMoney(Math.abs(metrics.debtService)), isTotal: true },
                              { label: "DSCR = NOI \u00f7 Debt service", value: `${metrics.dscr.toFixed(2)}x`, isTotal: true },
                            ],
                          });
                        },
                      },
                      {
                        label: "Yield on cost (purchase price)", value: metrics.capRateApplicable ? (isNaN(metrics.capRate) ? "\u2014" : `${(metrics.capRate * 100).toFixed(2)}%`) : "\u2014 (annual basis only)", flagged: metrics.hasFlag && metrics.capRateApplicable,
                        onClick: () => {
                          if (!metrics.capRateApplicable) return;
                          const noiRows = buildFieldRows(metrics.entries, "noi", formatMoney);
                          const priceRow = metrics.purchasePriceRep ? [{
                            label: "Purchase price", value: metrics.purchasePriceRep.purchasePrice || formatMoney(metrics.purchasePriceRep.purchasePriceNum),
                            fileName: metrics.purchasePriceRep.fileName, sheetName: metrics.purchasePriceRep.sheetName,
                            addr: metrics.purchasePriceRep.purchasePriceAddr, rep: metrics.purchasePriceRep,
                          }] : [];
                          openCalcPreview({
                            title: `Yield on cost (purchase price) \u2014 ${metrics.label}`,
                            formula: "Yield on cost = Net Operating Income \u00f7 Purchase price. Not a market cap rate \u2014 that needs a current valuation, which isn't available here.",
                            rows: [
                              { label: "Net operating income", isSection: true },
                              ...noiRows,
                              { label: "Total NOI", value: formatMoney(metrics.noi), isTotal: true },
                              { label: "Purchase price", isSection: true },
                              ...priceRow,
                              { label: "Yield on cost = NOI \u00f7 Purchase price", value: `${(metrics.capRate * 100).toFixed(2)}%`, isTotal: true },
                            ],
                          });
                        },
                      },
                    ].map((c) => (
                      <div
                        key={c.label}
                        onClick={() => { if (findMode && c.onClick) c.onClick(); }}
                        style={{ background: CARD, border: `0.5px solid ${BORDER}`, borderRadius: 10, padding: "14px 16px", cursor: findMode ? "zoom-in" : "default" }}
                      >
                        <div style={{ fontSize: 12, color: SLATE, marginBottom: 6 }}>{c.label}</div>
                        <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: c.value.length > 10 ? 13 : 20, fontWeight: 600, color: INK }}>{c.value}{c.flagged ? <span style={{ color: FLAG_AMBER, fontSize: 14 }}> *</span> : null}</div>
                      </div>
                    ))}
                  </div>
                </>
              );
            })()}

            {latestActiveReport && latestActiveReport.reconciliationStatus === "mismatch" && (
              <div style={{ background: FLAG_AMBER_BG, border: `1px solid ${FLAG_AMBER}`, borderRadius: 10, padding: "12px 16px", marginBottom: 20, display: "flex", gap: 10, alignItems: "flex-start" }}>
                <AlertTriangle size={16} color={FLAG_AMBER} style={{ marginTop: 1, flexShrink: 0 }} />
                <div>
                  <div style={{ fontSize: 13, fontWeight: 600, color: FLAG_AMBER, marginBottom: 4 }}>This report doesn't tie to its own stated totals</div>
                  <div style={{ fontSize: 12.5, color: INK, lineHeight: 1.5 }}>
                    {latestActiveReport.reconciliation.filter((c) => c.status === "mismatch").map((c) => (
                      <div key={c.label}>
                        {c.label}: months sum to {formatMoney(c.computed)}, but the report's own YTD/Actual column shows {formatMoney(c.reported)} ({formatMoney(Math.abs(c.diff))} off)
                        {findMode && c.reportedAddr && (
                          <button onClick={() => openCellPreview(latestActiveReport, c.reportedAddr)} style={{ marginLeft: 8, background: "none", border: `1px solid ${FLAG_AMBER}`, color: FLAG_AMBER, borderRadius: 4, fontSize: 11, padding: "1px 6px", cursor: "zoom-in" }}>view cell</button>
                        )}
                      </div>
                    ))}
                  </div>
                </div>
              </div>
            )}

            {latestActiveReport && latestActiveReport.reconciliationStatus === "not-checked" && (
              <div style={{ background: HOVER_BG, border: `1px solid ${BORDER}`, borderRadius: 10, padding: "12px 16px", marginBottom: 20, fontSize: 12.5, color: SLATE, lineHeight: 1.5 }}>
                This report has no YTD/Actual column to check against, so totals couldn't be verified against a source-stated total \u2014 only the internal month-to-month math was checked. This is "not checked," not "passed."
              </div>
            )}

            {latestActiveReport && latestActiveReport.dataQualityNotes && latestActiveReport.dataQualityNotes.length > 0 && (
              <div style={{ background: HOVER_BG, border: `1px solid ${BORDER}`, borderRadius: 10, padding: "12px 16px", marginBottom: 20, display: "flex", gap: 10, alignItems: "flex-start" }}>
                <AlertTriangle size={16} color={SLATE} style={{ marginTop: 1, flexShrink: 0 }} />
                <div>
                  <div style={{ fontSize: 13, fontWeight: 600, color: INK, marginBottom: 4 }}>Data was auto-corrected on ingest</div>
                  <div style={{ fontSize: 12.5, color: SLATE, lineHeight: 1.5 }}>
                    {latestActiveReport.dataQualityNotes.map((note, i) => <div key={i}>{note}</div>)}
                  </div>
                </div>
              </div>
            )}

            {monthVariance && (
              <div style={{ marginBottom: 28 }}>
                <div style={{ fontSize: 13, fontWeight: 500, color: SLATE, marginBottom: 10, textTransform: "uppercase", letterSpacing: "0.06em" }}>
                  Variance: {monthVariance.priorLabel} {"\u2192"} {monthVariance.latestLabel}
                </div>
                <div style={{ background: CARD, border: `0.5px solid ${BORDER}`, borderRadius: 12, overflow: "hidden" }}>
                  <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                    <thead>
                      <tr style={{ borderBottom: `1px solid ${BORDER}`, background: HEADER_BG }}>
                        <th style={{ textAlign: "left", padding: "9px 16px", color: SLATE, fontWeight: 500, fontSize: 12 }}></th>
                        <th style={{ textAlign: "right", padding: "9px 16px", color: SLATE, fontWeight: 500, fontSize: 12 }}>{monthVariance.priorLabel}</th>
                        <th style={{ textAlign: "right", padding: "9px 16px", color: SLATE, fontWeight: 500, fontSize: 12 }}>{monthVariance.latestLabel}</th>
                        <th style={{ textAlign: "right", padding: "9px 16px", color: SLATE, fontWeight: 500, fontSize: 12 }}>{"\u0394"}</th>
                        <th style={{ textAlign: "right", padding: "9px 16px", color: SLATE, fontWeight: 500, fontSize: 12 }}>{"\u0394"} %</th>
                      </tr>
                    </thead>
                    <tbody>
                      {monthVariance.rows.map((row) => {
                        const hasBoth = !isNaN(row.latestVal) && !isNaN(row.priorVal);
                        const delta = hasBoth ? row.latestVal - row.priorVal : NaN;
                        const pctDelta = hasBoth && row.priorVal !== 0 ? (delta / Math.abs(row.priorVal)) * 100 : NaN;
                        const deltaColor = !hasBoth ? SLATE : delta > 0 ? GOOD_COLOR : delta < 0 ? BAD_COLOR : SLATE;
                        const fmt = (v) => isNaN(v) ? "\u2014" : row.isPct ? `${(v * 100).toFixed(1)}%` : formatMoney(v);
                        return (
                          <tr
                            key={row.label}
                            onClick={() => { if (findMode && row.addr && row.rep) openCellPreview(row.rep, row.addr); }}
                            style={{ borderBottom: `0.5px solid ${BORDER}`, cursor: findMode && row.addr ? "zoom-in" : "default" }}
                          >
                            <td style={{ padding: "10px 16px", color: SLATE }}>{row.label}</td>
                            <td style={{ padding: "10px 16px", textAlign: "right", fontFamily: "'IBM Plex Mono', monospace" }}>{fmt(row.priorVal)}</td>
                            <td style={{ padding: "10px 16px", textAlign: "right", fontFamily: "'IBM Plex Mono', monospace", fontWeight: 600 }}>{fmt(row.latestVal)}</td>
                            <td style={{ padding: "10px 16px", textAlign: "right", fontFamily: "'IBM Plex Mono', monospace", color: deltaColor }}>
                              {hasBoth ? `${delta > 0 ? "+" : ""}${row.isPct ? (delta * 100).toFixed(1) + "pp" : formatMoney(delta)}` : "\u2014"}
                            </td>
                            <td style={{ padding: "10px 16px", textAlign: "right", fontFamily: "'IBM Plex Mono', monospace", color: deltaColor }}>
                              {!isNaN(pctDelta) ? `${pctDelta > 0 ? "+" : ""}${pctDelta.toFixed(1)}%` : "\u2014"}
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            <div style={{ fontSize: 12, color: SLATE, marginBottom: 10, marginTop: -8 }}>Trend charts below are supporting detail \u2014 the variance table above is the headline.</div>

            {monthlyTrendData.length >= 2 && (
              <div style={{ marginBottom: 28 }}>
                <div style={{ fontSize: 13, fontWeight: 500, color: SLATE, marginBottom: 10, textTransform: "uppercase", letterSpacing: "0.06em" }}>
                  {latestActiveReport.period || latestActiveReport.year} {"\u2014"} monthly trend (detail)
                </div>
                <div style={{ background: CARD, border: `0.5px solid ${BORDER}`, borderRadius: 12, padding: "1.5rem 1.25rem 0.5rem" }}>
                  <ResponsiveContainer width="100%" height={240}>
                    <ComposedChart data={monthlyTrendData} margin={{ top: 4, right: 8, left: 8, bottom: 4 }}>
                      <CartesianGrid stroke={BORDER} strokeDasharray="3 3" vertical={false} />
                      <XAxis dataKey="month" tick={{ fontFamily: "Inter, sans-serif", fontSize: 12, fill: SLATE }} axisLine={{ stroke: BORDER }} tickLine={false} />
                      <YAxis yAxisId="money" tick={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 11, fill: SLATE }} axisLine={false} tickLine={false} tickFormatter={(v) => `$${(v / 1000).toFixed(0)}k`} />
                      <YAxis yAxisId="pct" orientation="right" domain={[0, 100]} tick={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 11, fill: SLATE }} axisLine={false} tickLine={false} tickFormatter={(v) => `${v}%`} />
                      <Tooltip
                        contentStyle={{ background: SURFACE_DARK, border: `1px solid ${BORDER}`, borderRadius: 8, fontFamily: "Inter, sans-serif", fontSize: 12 }}
                        labelStyle={{ color: "#FFFFFF", fontWeight: 600, marginBottom: 4 }}
                        itemStyle={{ color: "#FFFFFF" }}
                        formatter={(value, name) => name === "Occupancy %" ? [`${value}%`, name] : [`$${value.toLocaleString()}`, name]}
                      />
                      <Legend wrapperStyle={{ fontFamily: "Inter, sans-serif", fontSize: 12, color: SLATE }} />
                      <Bar
                        yAxisId="money"
                        dataKey="NOI"
                        fill={DEED_GREEN}
                        radius={[4, 4, 0, 0]}
                        barSize={22}
                        style={{ cursor: findMode ? "zoom-in" : "default" }}
                        onClick={(data) => { if (findMode && data && data._addr) openCellPreview(data._rep, data._addr); }}
                      />
                      <Line
                        yAxisId="money"
                        type="monotone"
                        dataKey="Net Cash Flow"
                        stroke="#C99BE0"
                        strokeWidth={2}
                        dot={makeClickableDot("#C99BE0", findMode, (payload) => { if (payload && payload._cashAddr) openCellPreview(payload._rep, payload._cashAddr); }, 3)}
                        connectNulls
                      />
                      <Line
                        yAxisId="pct"
                        type="monotone"
                        dataKey="Occupancy %"
                        stroke="#7FB3D5"
                        strokeWidth={2.5}
                        dot={makeClickableDot("#7FB3D5", findMode, (payload) => { if (payload && payload._occAddr) openCellPreview(payload._rep, payload._occAddr); }, 3)}
                        connectNulls
                      />
                    </ComposedChart>
                  </ResponsiveContainer>
                </div>
              </div>
            )}

            <button
              onClick={() => setShowComparisonTable((v) => !v)}
              style={{ display: "flex", alignItems: "center", gap: 6, background: "none", border: "none", color: SLATE, fontSize: 13, fontWeight: 500, textTransform: "uppercase", letterSpacing: "0.06em", cursor: "pointer", padding: 0, marginBottom: 10 }}
            >
              {showComparisonTable ? "\u25BE" : "\u25B8"} Full year-by-year comparison table
            </button>
            {showComparisonTable && (
            <>
            <div style={{ background: CARD, border: `0.5px solid ${BORDER}`, borderRadius: 12, overflow: "hidden", marginBottom: 28 }}>
              <div style={{ overflowX: "auto" }}>
                <table style={{ borderCollapse: "collapse", fontSize: 13, width: "100%" }}>
                  <thead>
                    <tr style={{ borderBottom: `1px solid ${BORDER}`, background: HEADER_BG }}>
                      <th style={{ position: "sticky", left: 0, background: HEADER_BG, textAlign: "left", padding: "10px 14px", fontWeight: 500, color: SLATE, fontSize: 12, minWidth: 160 }}> </th>
                      {activeReports.map((r) => (
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
                        {activeReports.map((r) => (
                          <td
                            key={r.id}
                            onClick={() => { if (findMode && row.calc) row.calc(r); }}
                            style={{ padding: "9px 14px", fontFamily: row.isText ? "Inter, sans-serif" : "'IBM Plex Mono', monospace", cursor: findMode && row.calc ? "zoom-in" : "default" }}
                          >
                            {row.get(r)}
                          </td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
            <div style={{ fontSize: 12, color: SLATE, marginBottom: 28, lineHeight: 1.5 }}>
              "Yield on cost" is NOI over cost basis, not a market cap rate \u2014 a true cap rate needs a current valuation, which isn't available here. DSCR uses YTD NOI over YTD debt service. Debt yield uses NOI over the outstanding loan balance where that's stated in the report. IRR isn't shown \u2014 it needs equity cash flow timing this report doesn't include. <span style={{ color: FLAG_AMBER }}>*</span> means the figure is built from a month with a data-quality flag \u2014 check the exceptions in that property's detail table before relying on it.
            </div>
            </>
            )}

            {activeTrend && (
              <div style={{ marginBottom: 28 }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", flexWrap: "wrap", gap: 10, marginBottom: 10 }}>
                  <div style={{ fontSize: 13, fontWeight: 500, color: SLATE, textTransform: "uppercase", letterSpacing: "0.06em" }}>
                    {activeProperty} {"\u2014"} performance over time (detail)
                  </div>
                  <div style={{ fontSize: 12, color: SLATE, fontFamily: "'IBM Plex Mono', monospace" }}>
                    {activeTrend.isFullTTM ? "TTM" : `Trailing ${activeTrend.monthsUsed}mo`} NOI: {formatMoney(activeTrend.ttmNOI)}
                    {!isNaN(activeTrend.ttmDSCR) && <> {"\u00b7"} DSCR {activeTrend.ttmDSCR.toFixed(2)}x</>}
                    {!isNaN(activeTrend.ttmCapRate) && <> {"\u00b7"} Yield on cost {(activeTrend.ttmCapRate * 100).toFixed(2)}%</>}
                  </div>
                </div>
                <div style={{ background: CARD, border: `0.5px solid ${BORDER}`, borderRadius: 12, padding: "1.5rem 1.25rem 0.5rem" }}>
                  <ResponsiveContainer width="100%" height={280}>
                    <ComposedChart data={activeTrend.data} margin={{ top: 4, right: 8, left: 8, bottom: 4 }}>
                      <CartesianGrid stroke={BORDER} strokeDasharray="3 3" vertical={false} />
                      <XAxis dataKey="year" tick={{ fontFamily: "Inter, sans-serif", fontSize: 12, fill: SLATE }} axisLine={{ stroke: BORDER }} tickLine={false} />
                      <YAxis yAxisId="money" tick={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 11, fill: SLATE }} axisLine={false} tickLine={false} tickFormatter={(v) => `$${(v / 1000).toFixed(0)}k`} />
                      <YAxis yAxisId="pct" orientation="right" domain={[0, 100]} tick={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 11, fill: SLATE }} axisLine={false} tickLine={false} tickFormatter={(v) => `${v}%`} />
                      <Tooltip
                        contentStyle={{ background: SURFACE_DARK, border: `1px solid ${BORDER}`, borderRadius: 8, fontFamily: "Inter, sans-serif", fontSize: 12 }}
                        labelStyle={{ color: "#FFFFFF", fontWeight: 600, marginBottom: 4 }}
                        itemStyle={{ color: "#FFFFFF" }}
                        formatter={(value, name) => name === "Occupancy %" ? [`${value}%`, name] : [`$${value.toLocaleString()}`, name]}
                      />
                      <Legend wrapperStyle={{ fontFamily: "Inter, sans-serif", fontSize: 12, color: SLATE }} />
                      <Bar
                        yAxisId="money"
                        dataKey="NOI"
                        fill={DEED_GREEN}
                        radius={[4, 4, 0, 0]}
                        barSize={36}
                        style={{ cursor: findMode ? "zoom-in" : "default" }}
                        onClick={(data) => {
                          if (!findMode || !data || !data._rep) return;
                          const rep = data._rep;
                          const monthRows = rep.months.filter((mo) => !mo.notYetReported).map((mo) => ({
                            label: `${mo.label} ${rep.year}`, value: formatMoney(mo.noi), raw: mo.noi,
                            fileName: rep.fileName, sheetName: rep.sheetName, addr: mo.src.noi, rep,
                          })).filter((r) => !isNaN(r.raw));
                          openCalcPreview({
                            title: `NOI \u2014 ${rep.year || rep.period}`,
                            formula: `Sum of Net Operating Income across ${monthRows.length} reported months`,
                            rows: [...monthRows, { label: "Total NOI", value: formatMoney(data.NOI), isTotal: true }],
                          });
                        }}
                      />
                      <Line
                        yAxisId="money"
                        type="monotone"
                        dataKey="Net Cash Flow"
                        stroke="#C99BE0"
                        strokeWidth={2}
                        dot={makeClickableDot("#C99BE0", findMode, (payload) => {
                          const rep = payload && payload._rep;
                          if (!rep) return;
                          const rows = buildFieldRows(rep.months.filter((mo) => !mo.notYetReported).map((mo) => ({ ...mo, rep, src: mo.src, monthLabel: mo.label, year: rep.year })), "netCashFlow", formatMoney);
                          openCalcPreview({
                            title: `Net Cash Flow \u2014 ${rep.year || rep.period}`,
                            formula: `Sum of Net Cash Flow across ${rows.length} reported months`,
                            rows: [...rows, { label: "Total Net Cash Flow", value: formatMoney(rep.ytdNetCashFlow), isTotal: true }],
                          });
                        }, 4)}
                        connectNulls
                      />
                      <Line
                        yAxisId="pct"
                        type="monotone"
                        dataKey="Occupancy %"
                        stroke="#7FB3D5"
                        strokeWidth={2.5}
                        dot={makeClickableDot("#7FB3D5", findMode, (payload) => {
                          const rep = payload && payload._rep;
                          if (!rep) return;
                          const rows = buildFieldRows(rep.months.filter((mo) => !mo.notYetReported).map((mo) => ({ ...mo, rep, src: mo.src, monthLabel: mo.label, year: rep.year })), "occupiedPct", (v) => `${(v * 100).toFixed(1)}%`);
                          openCalcPreview({
                            title: `Occupancy (avg) \u2014 ${rep.year || rep.period}`,
                            formula: `Average of ${rows.length} reported months`,
                            rows: [...rows, { label: "Average occupancy", value: isNaN(rep.avgOccupancy) ? "\u2014" : `${(rep.avgOccupancy * 100).toFixed(1)}%`, isTotal: true }],
                          });
                        }, 4)}
                        connectNulls
                      />
                    </ComposedChart>
                  </ResponsiveContainer>
                </div>
              </div>
            )}

            {activeReports.length >= 2 && (
              <PropertyOverviewCard overview={propertyOverviews[overviewGroupKey]} onRetry={() => generatePropertyOverview(overviewGroupKey, activeProperty, activeReports)} />
            )}

            {activeReports.map((rep, idx) => (
              <div key={rep.id} style={{ marginBottom: 36 }}>
                {activeReports.length >= 2 && (
                  <div style={{ display: "flex", alignItems: "center", gap: 12, margin: idx === 0 ? "0 0 20px" : "40px 0 20px" }}>
                    <div style={{ flex: 1, height: 1, background: BORDER }} />
                    <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 12, fontWeight: 600, color: DEED_GREEN, letterSpacing: "0.08em", padding: "4px 14px", border: `1px solid ${DEED_GREEN}`, borderRadius: 20, whiteSpace: "nowrap" }}>
                      {rep.year || rep.period}
                    </div>
                    <div style={{ flex: 1, height: 1, background: BORDER }} />
                  </div>
                )}
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

                <DecisionSupportCard decision={decisions[rep.id]} onRetry={() => generateDecision(rep, activeReports)} />

                {rep.flaggedMonthCount > 0 && (
                  <div style={{ marginBottom: 20 }}>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", marginBottom: 10 }}>
                      <div style={{ fontSize: 13, fontWeight: 500, color: SLATE, textTransform: "uppercase", letterSpacing: "0.06em" }}>Exception history</div>
                      {rep.businessImpactTotal > 0 && (
                        <div style={{ fontSize: 12, color: SLATE, fontFamily: "'IBM Plex Mono', monospace" }}>{formatMoney(rep.businessImpactTotal)} estimated business-impact total</div>
                      )}
                    </div>
                    <div style={{ background: CARD, border: `0.5px solid ${BORDER}`, borderRadius: 12, padding: "4px 0" }}>
                      {[...rep.allFlags].sort((a, b) => (b.dollarImpact || 0) - (a.dollarImpact || 0)).map((f, i) => {
                        const tierColor = f.tier === "data-integrity" ? ERROR_COLOR : FLAG_AMBER;
                        const addr = f.field ? (rep.months.find((mo) => mo.key === f.monthKey) || {}).src?.[f.field] : null;
                        return (
                          <div
                            key={i}
                            onClick={() => { if (findMode && addr) openCellPreview(rep, addr); }}
                            style={{ display: "flex", gap: 12, alignItems: "flex-start", justifyContent: "space-between", padding: "10px 18px", borderBottom: `0.5px solid ${BORDER}`, cursor: findMode && addr ? "zoom-in" : "default" }}
                          >
                            <div style={{ display: "flex", gap: 10, alignItems: "flex-start" }}>
                              <span style={{ width: 7, height: 7, borderRadius: "50%", background: tierColor, marginTop: 5, flexShrink: 0 }} />
                              <div>
                                <div style={{ display: "flex", gap: 8, alignItems: "center", marginBottom: 2 }}>
                                  <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 11, color: SLATE }}>{f.monthLabel}</span>
                                  <span style={{ fontSize: 10, fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.04em", color: tierColor, border: `1px solid ${tierColor}`, borderRadius: 10, padding: "1px 7px" }}>
                                    {f.tier === "data-integrity" ? "Data integrity" : "Business"}
                                  </span>
                                </div>
                                <span style={{ fontSize: 13 }}>{f.message.split(" \u2014 source:")[0]}</span>
                              </div>
                            </div>
                            <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 12, fontWeight: 600, color: INK, whiteSpace: "nowrap", flexShrink: 0 }}>
                              {typeof f.dollarImpact === "number" ? `~${formatMoney(f.dollarImpact)}` : ""}
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  </div>
                )}

                <button
                  onClick={() => setExpandedYearData((prev) => ({ ...prev, [rep.id]: !prev[rep.id] }))}
                  style={{ display: "flex", alignItems: "center", gap: 6, background: "none", border: "none", color: SLATE, fontSize: 13, fontWeight: 500, textTransform: "uppercase", letterSpacing: "0.06em", cursor: "pointer", padding: 0, marginBottom: 10 }}
                >
                  {expandedYearData[rep.id] ? "\u25BE" : "\u25B8"} Full standardized data ({rep.reportedMonthCount} months)
                </button>

                {expandedYearData[rep.id] && (
                <>
                <div style={{ fontSize: 13, fontWeight: 500, color: SLATE, marginBottom: 10, textTransform: "uppercase", letterSpacing: "0.06em" }}>Standardized financial data</div>
                <div style={{ background: CARD, border: `0.5px solid ${BORDER}`, borderRadius: 12, overflow: "hidden", marginBottom: 24 }}>
                  <div style={{ overflowX: "auto" }}>
                    <table style={{ borderCollapse: "collapse", fontSize: 13, width: "100%" }}>
                      <thead>
                        <tr style={{ borderBottom: `1px solid ${BORDER}`, background: HEADER_BG }}>
                          <th style={{ position: "sticky", left: 0, background: HEADER_BG, textAlign: "left", padding: "10px 14px", fontWeight: 500, color: SLATE, fontSize: 12, minWidth: 190 }}> </th>
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
                              const addr = mo.src[row.key];
                              const srcTitle = addr ? `Source: sheet "${rep.sheetName}", cell ${addr}` : undefined;
                              return (
                                <td
                                  key={mo.key}
                                  title={flag.length ? undefined : srcTitle}
                                  onClick={() => { if (findMode) openCellPreview(rep, addr); }}
                                  style={{ padding: "9px 14px", textAlign: "right", fontFamily: "'IBM Plex Mono', monospace", background: flag.length ? FLAG_AMBER_BG : "transparent", color: flag.length ? FLAG_AMBER : (mo.notYetReported ? SLATE : INK), cursor: findMode && addr ? "zoom-in" : (flag.length ? "help" : "default") }}
                                >
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
                        <tr style={{ borderBottom: `1px solid ${BORDER}`, background: HEADER_BG }}>
                          <th style={{ position: "sticky", left: 0, background: HEADER_BG, textAlign: "left", padding: "10px 14px", fontWeight: 500, color: SLATE, fontSize: 12, minWidth: 190 }}> </th>
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
                              const addr = mo.src[row.key];
                              const srcTitle = addr ? `Source: sheet "${rep.sheetName}", cell ${addr}` : undefined;
                              return (
                                <td
                                  key={mo.key}
                                  title={flag.length ? undefined : srcTitle}
                                  onClick={() => { if (findMode) openCellPreview(rep, addr); }}
                                  style={{ padding: "9px 14px", textAlign: "right", fontFamily: "'IBM Plex Mono', monospace", background: flag.length ? FLAG_AMBER_BG : "transparent", color: flag.length ? FLAG_AMBER : (mo.notYetReported ? SLATE : INK), cursor: findMode && addr ? "zoom-in" : (flag.length ? "help" : "default") }}
                                >
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
                </>
                )}

                <div style={{ fontSize: 12, color: SLATE, display: "flex", alignItems: "center", gap: 6, flexWrap: "wrap" }}>
                  <FileSpreadsheet size={14} /> {rep.fileName} {"\u00b7"} sheet "{rep.sheetName}"
                  {rep.fileURL && (
                    <>
                      {"\u00b7"}
                      <a href={rep.fileURL} download={rep.fileName} style={{ color: DEED_GREEN, textDecoration: "none", borderBottom: `1px solid ${DEED_GREEN}` }}>
                        Download original file
                      </a>
                    </>
                  )}
                </div>
              </div>
            ))}

            <div style={{ marginTop: 32, paddingTop: 20, borderTop: `0.5px solid ${BORDER}`, display: "flex", justifyContent: "center" }}>
              <button
                onClick={() => { reports.forEach((r) => { if (r.fileURL) URL.revokeObjectURL(r.fileURL); }); setReports([]); setSelectedId(null); setError(null); setDecisions({}); requestedDecisions.current.clear(); setPropertyOverviews({}); requestedOverviews.current.clear(); }}
                style={{ display: "flex", alignItems: "center", gap: 6, background: "transparent", border: `1px solid ${BORDER}`, borderRadius: 8, padding: "9px 16px", fontSize: 13, color: SLATE, cursor: "pointer" }}
              >
                <RotateCcw size={13} /> Clear all &amp; start over
              </button>
            </div>
          </>
        )}
      </div>

      {toast && (
        <div
          className="toast-msg"
          style={{
            position: "fixed",
            bottom: 28,
            left: "50%",
            zIndex: 200,
            background: SURFACE_DARK,
            color: "#FFFFFF",
            border: `1px solid ${BORDER}`,
            borderRadius: 20,
            padding: "9px 18px",
            fontSize: 13,
            fontFamily: "Inter, sans-serif",
            boxShadow: "0 8px 24px rgba(0,0,0,0.5)",
            whiteSpace: "nowrap",
          }}
        >
          {toast}
        </div>
      )}

      {reports.length > 0 && !previewTarget && (
        <button
          onClick={() => { setFindMode((v) => !v); setPreviewTarget(null); }}
          title={findMode ? "Cancel find cell (Esc)" : "Find cell"}
          style={{
            position: "fixed",
            right: 0,
            top: "50%",
            transform: "translateY(-50%)",
            zIndex: 95,
            background: findMode ? BAD_COLOR : DEED_GREEN,
            color: "#FFFFFF",
            border: "none",
            borderRadius: "12px 0 0 12px",
            padding: "20px 12px",
            cursor: "pointer",
            display: "flex",
            flexDirection: "column",
            alignItems: "center",
            gap: 10,
            boxShadow: `-6px 0 20px rgba(0,0,0,0.5), 0 0 0 1px ${findMode ? BAD_COLOR : DEED_GREEN}`,
          }}
        >
          <Search size={20} strokeWidth={2.75} />
          <span style={{ writingMode: "vertical-rl", fontSize: 13, fontWeight: 700, letterSpacing: "0.1em", fontFamily: "'IBM Plex Mono', monospace", textTransform: "uppercase" }}>
            {findMode ? "Cancel" : "Find cell"}
          </span>
        </button>
      )}

      {previewTarget && (
        <>
          <div onClick={() => setPreviewTarget(null)} style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.55)", zIndex: 90 }} />
          <div style={{ position: "fixed", top: 0, right: 0, bottom: 0, width: "min(600px, 92vw)", background: CARD, borderLeft: `1px solid ${BORDER}`, zIndex: 100, display: "flex", flexDirection: "column", boxShadow: "-8px 0 24px rgba(0,0,0,0.5)" }}>
            <div style={{ padding: "16px 18px", borderBottom: `1px solid ${BORDER}`, display: "flex", justifyContent: "space-between", alignItems: "center", flexShrink: 0 }}>
              <div>
                {previewTarget.mode === "calc" ? (
                  <>
                    <div style={{ fontFamily: "'Source Serif 4', serif", fontWeight: 600, fontSize: 14, color: INK }}>{previewTarget.title}</div>
                    <div style={{ fontSize: 12, color: SLATE, marginTop: 2 }}>{previewTarget.formula}</div>
                  </>
                ) : (
                  <>
                    <div style={{ fontFamily: "'Source Serif 4', serif", fontWeight: 600, fontSize: 14, color: INK }}>{previewTarget.fileName}</div>
                    <div style={{ fontSize: 12, color: SLATE, marginTop: 2 }}>Sheet "{previewTarget.sheetName}" {"\u00b7"} cell {previewTarget.addr}</div>
                  </>
                )}
              </div>
              <div style={{ display: "flex", alignItems: "center", gap: 10, flexShrink: 0 }}>
                {previewTarget.mode === "cell" && previewTarget.returnTo && (
                  <button onClick={() => setPreviewTarget(previewTarget.returnTo)} style={{ background: "none", border: `1px solid ${BORDER}`, color: SLATE, cursor: "pointer", padding: "5px 10px", borderRadius: 6, fontSize: 12 }}>
                    {"\u2039"} Back to breakdown
                  </button>
                )}
                <button onClick={() => setPreviewTarget(null)} style={{ background: "none", border: "none", color: SLATE, cursor: "pointer", padding: 4 }}>
                  <X size={18} />
                </button>
              </div>
            </div>

            {previewTarget.mode === "cell" && (
              <div style={{ display: "flex", alignItems: "center", background: "#FFFFFF", borderBottom: "1px solid #D4D4D4", flexShrink: 0 }}>
                <div style={{ width: 70, padding: "4px 8px", borderRight: "1px solid #D4D4D4", fontFamily: "Calibri, 'Segoe UI', Arial, sans-serif", fontSize: 12.5, color: "#000000" }}>
                  {previewTarget.addr}
                </div>
                <div style={{ width: 30, textAlign: "center", borderRight: "1px solid #D4D4D4", fontFamily: "Georgia, serif", fontStyle: "italic", fontSize: 13, color: "#616161" }}>fx</div>
                <div style={{ flex: 1, padding: "4px 10px", fontFamily: "Calibri, 'Segoe UI', Arial, sans-serif", fontSize: 12.5, color: "#000000", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>
                  {(() => {
                    const cell = previewTarget.targetRow !== undefined && previewTarget.grid[previewTarget.targetRow] ? previewTarget.grid[previewTarget.targetRow][previewTarget.targetCol] : null;
                    if (!cell) return "";
                    if (cell.f) return `=${cell.f}`;
                    return cell.v === undefined ? "" : String(cell.v);
                  })()}
                </div>
              </div>
            )}

            {previewTarget.mode === "calc" ? (
              <div style={{ overflow: "auto", flex: 1, padding: "8px 4px" }}>
                {previewTarget.note && (
                  <div style={{ fontSize: 12, color: SLATE, padding: "8px 14px", lineHeight: 1.5 }}>{previewTarget.note}</div>
                )}
                {previewTarget.rows.map((row, i) => (
                  row.isSection ? (
                    <div key={i} style={{ padding: "12px 18px 6px", fontSize: 11, fontWeight: 600, color: DEED_GREEN, textTransform: "uppercase", letterSpacing: "0.06em" }}>
                      {row.label}
                    </div>
                  ) : (
                  <div
                    key={i}
                    onClick={() => {
                      if (row.addr && row.rep && row.rep.previewGrid) {
                        const { r, c } = XLSX.utils.decode_cell(row.addr);
                        setPreviewTarget({ mode: "cell", fileName: row.fileName, sheetName: row.sheetName, grid: row.rep.previewGrid, merges: row.rep.previewMerges, colWidths: row.rep.previewColWidths, lastCol: row.rep.previewLastCol, targetRow: r, targetCol: c, addr: row.addr, returnTo: previewTarget });
                      }
                    }}
                    style={{
                      display: "flex", justifyContent: "space-between", alignItems: "center", gap: 12,
                      padding: "11px 18px", borderBottom: `0.5px solid ${BORDER}`,
                      cursor: row.addr ? "zoom-in" : "default",
                      fontWeight: row.isTotal ? 600 : 400,
                      background: row.isTotal ? HEADER_BG : "transparent",
                    }}
                  >
                    <div>
                      <div style={{ fontSize: 13, color: row.isTotal ? INK : undefined }}>{row.label}</div>
                      {row.fileName && (
                        <div style={{ fontSize: 11, color: SLATE, marginTop: 2 }}>{row.fileName} {"\u00b7"} {row.sheetName}{row.addr ? `, cell ${row.addr}` : ""}</div>
                      )}
                    </div>
                    <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 13, fontWeight: row.isTotal ? 600 : 500, color: INK, whiteSpace: "nowrap" }}>{row.value}</div>
                  </div>
                  )
                ))}
              </div>
            ) : (
              <div style={{ overflow: "auto", flex: 1, padding: 0, background: "#FFFFFF" }}>
                {(() => {
                  const merges = previewTarget.merges || [];
                  const mergeAnchor = new Map();
                  const mergeCovered = new Set();
                  merges.forEach((m) => {
                    const rowSpan = m.e.r - m.s.r + 1;
                    const colSpan = m.e.c - m.s.c + 1;
                    if (rowSpan <= 1 && colSpan <= 1) return;
                    mergeAnchor.set(`${m.s.r},${m.s.c}`, { rowSpan, colSpan });
                    for (let rr = m.s.r; rr <= m.e.r; rr++) {
                      for (let cc = m.s.c; cc <= m.e.c; cc++) {
                        if (rr === m.s.r && cc === m.s.c) continue;
                        mergeCovered.add(`${rr},${cc}`);
                      }
                    }
                  });
                  const widths = previewTarget.colWidths || [];
                  return (
                    <table style={{ borderCollapse: "collapse", fontSize: 12, fontFamily: "Calibri, 'Segoe UI', Arial, sans-serif", color: "#000000", tableLayout: "fixed" }}>
                      <colgroup>
                        <col style={{ width: 34 }} />
                        {Array.from({ length: previewTarget.lastCol + 1 }, (_, c) => (
                          <col key={c} style={{ width: widths[c] || 72 }} />
                        ))}
                      </colgroup>
                      <thead>
                        <tr>
                          <th style={{ position: "sticky", top: 0, left: 0, background: "#F3F2F1", zIndex: 3, padding: "3px 6px", border: "1px solid #D4D4D4", fontWeight: 400, color: "#616161" }}></th>
                          {Array.from({ length: previewTarget.lastCol + 1 }, (_, c) => (
                            <th key={c} style={{ position: "sticky", top: 0, background: "#F3F2F1", color: "#616161", padding: "3px 6px", border: "1px solid #D4D4D4", fontWeight: 400, zIndex: 2, textAlign: "center", overflow: "hidden" }}>
                              {XLSX.utils.encode_col(c)}
                            </th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {previewTarget.grid.map((row, r) => (
                          <tr key={r}>
                            <td style={{ position: "sticky", left: 0, background: "#F3F2F1", color: "#616161", textAlign: "center", padding: "3px 6px", border: "1px solid #D4D4D4" }}>{r + 1}</td>
                            {Array.from({ length: previewTarget.lastCol + 1 }, (_, c) => {
                              if (mergeCovered.has(`${r},${c}`)) return null;
                              const span = mergeAnchor.get(`${r},${c}`);
                              const isTarget = r === previewTarget.targetRow && c === previewTarget.targetCol;
                              const cell = row[c];
                              const display = cell && cell.w !== undefined ? cell.w : "";
                              const isNumeric = cell && cell.t === "n";
                              return (
                                <td
                                  key={c}
                                  ref={isTarget ? previewCellRef : null}
                                  colSpan={span ? span.colSpan : 1}
                                  rowSpan={span ? span.rowSpan : 1}
                                  style={{
                                    padding: "3px 6px",
                                    border: isTarget ? "2px solid #9C6500" : "1px solid #D4D4D4",
                                    background: isTarget ? "#FFEB9C" : "#FFFFFF",
                                    color: isTarget ? "#9C6500" : "#000000",
                                    whiteSpace: "nowrap",
                                    overflow: "hidden",
                                    textOverflow: "ellipsis",
                                    textAlign: isNumeric ? "right" : "left",
                                    fontWeight: isTarget ? 600 : 400,
                                  }}
                                >
                                  {display}
                                </td>
                              );
                            })}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  );
                })()}
              </div>
            )}
          </div>
        </>
      )}
    </div>
  );
}
