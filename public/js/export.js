/**
 * export.js
 * ─────────────────────────────────────────────────────────────────────────────
 * Standalone HTML report generator.
 * Packages the current client's processed data into a self-contained file
 * that clients can open, print, and share without the main dashboard app.
 *
 * Contents:
 *   1. exportReport()      — collects data from the live dashboard, builds
 *                            and downloads the standalone HTML file
 *   2. buildExportHTML()   — constructs the full standalone HTML document
 *   3. Export print dialog — the settings modal inside the exported file
 *
 * ── Architecture notes ───────────────────────────────────────────────────────
 *
 *  NO DUPLICATE PARSING
 *  The previous version of this file re-parsed the Excel data from scratch,
 *  duplicating every function from parser.js and charts.js. That has been
 *  removed entirely. exportReport() now reads the already-processed data
 *  from the live dashboard's DOM and client state — the same data the user
 *  is already looking at. This guarantees the export is always in sync with
 *  what is shown on screen.
 *
 *  DATA PACKAGE
 *  All processed data is serialised into a JSON payload embedded in the
 *  exported HTML as a <script type="application/json"> tag. The exported
 *  file's own inline script reads this payload on load and rebuilds the
 *  charts using the same Chart.js library (loaded from CDN).
 *
 *  SUMMARY SANITISATION
 *  The executive summary HTML is passed through sanitiseSummaryHTML()
 *  (from ui.js) before being embedded. This strips any <script> tags or
 *  event-handler attributes that could have been typed into the editor.
 *
 *  PRINT PIPELINE
 *  The exported file has its own print dialog and print pipeline, matching
 *  the main dashboard's print.js behaviour. This is intentional — the
 *  exported file is standalone and cannot depend on the main app's JS.
 * ─────────────────────────────────────────────────────────────────────────────
 */


/* ═══════════════════════════════════════════════════════════════════════════
 * 1. exportReport()
 * ═══════════════════════════════════════════════════════════════════════════ */

/**
 * exportReport(clientName)
 * Entry point called by the "Export & Send" button in the dashboard header.
 *
 * Reads processed data from the current client state (already parsed by
 * parser.js and rendered by charts.js), packages it into a JSON payload,
 * and downloads a standalone HTML report file.
 *
 * @param {string} clientName — used as the filename prefix
 */
function exportReport(clientName) {
  const i       = activeClient;
  const client  = clients[i];
  const sheets  = client.sheets;
  const names   = Object.keys(sheets);

  /* ── Re-derive sheet references (same logic as renderClient) ── */
  const scoringSheet = names.find(n => /scor/i.test(n) && !/prev/i.test(n)) || names[0];
  const utilSheet    = names.find(n => /util/i.test(n))
                    || names.find(n => n !== scoringSheet)
                    || names[0];
  const prevSheet    = names.find(n => /prev/i.test(n) || /last/i.test(n) || /prior/i.test(n));

  const scoringRows = sheets[scoringSheet] || [];
  const utilRows    = sheets[utilSheet]    || [];
  const prevRows    = prevSheet ? sheets[prevSheet] : [];

  /* ── Re-use parser.js functions (no duplication) ── */
  const sCols = detectCols(scoringRows);
  const uCols = detectCols(utilRows);
  const pCols = prevRows.length ? detectCols(prevRows) : {};

  const vehicleMap = buildVehicleMap(scoringRows, sCols);

  /* Day columns and info */
  const uKeys   = Object.keys(utilRows[0] || {});
  const dayCols = detectDayCols(uKeys);
  const dayInfo = dayCols.map(k => ({ key: k, ...parseDayInfo(k) }));

  /* Enrich vehicle map with utilisation data */
  enrichFromUtilSheet(vehicleMap, utilRows, uCols, dayInfo);

  /* Daily totals */
  const dailyTotals      = buildDailyTotals(utilRows, dayInfo);
  const vehicleDailyData = buildVehicleDailyData(utilRows, uCols, dayInfo);

  /* Previous month map */
  let prevMap = {};
  if (prevRows.length) {
    const result = buildPrevMap(prevRows, pCols);
    prevMap      = result.prevMap;
    reconcilePrevMap(vehicleMap, prevMap, result.normIndex);
  }

  const vehicles = Object.values(vehicleMap).filter(v => v.name && v.name !== 'Unknown');

  /* ── Save any in-progress summary edit before exporting ── */
  const liveSummaryEl = document.getElementById(`execBody_${activeClient}`);
  if (liveSummaryEl?.contentEditable === 'true') {
    /* Auto-save the draft before packaging */
    if (client) {
      client.customSummaryHTML = sanitiseSummaryHTML(liveSummaryEl.innerHTML);
    }
    liveSummaryEl.contentEditable = 'false';
    liveSummaryEl.classList.remove('is-editing');
  }

  /* ── Read current summary and notes from the live dashboard ── */
  const exportSummaryEl = document.getElementById(`execBody_${activeClient}`);
  const exportNotesEl   = document.getElementById(`execNotes_${activeClient}`);
  const summaryHTML     = exportSummaryEl
    ? sanitiseSummaryHTML(exportSummaryEl.innerHTML)
    : '';
  const notes = exportNotesEl ? exportNotesEl.value.trim() : '';

  /* ── Build the data payload ── */
  const payload = {
    clientName,
    reportMonth  : client?.month || '',
    generatedOn  : new Date().toLocaleDateString('en-GB', {
      day: 'numeric', month: 'long', year: 'numeric',
    }),

    /* Processed vehicle data */
    vehicles,
    prevMap,
    hasPrev      : Object.keys(prevMap).length > 0,

    /* Daily chart data */
    dailyTotals,
    dayLabels    : dayInfo.map(d => d.label),
    vehicleDailyData,

    /* Day metadata — allows exported file to reconstruct date filtering */
    dayInfoDates : dayInfo.map(d =>
      d.date && d.date.getFullYear() > 2000
        ? { year: d.date.getFullYear(), month: d.date.getMonth() + 1, day: d.date.getDate() }
        : null
    ),
    dayNums      : dayInfo.map(d => d.dayNum || null),
    dayKeys      : dayInfo.map(d => d.key),
    hasRealDates : dayInfo.some(d => d.date && d.date.getFullYear() > 2000),

    /* Raw util rows — sanitised: strip undefined/Date so JSON.stringify never throws */
    utilRows : JSON.parse(JSON.stringify(utilRows, (_, v) =>
      v === undefined   ? null
      : v instanceof Date ? v.toISOString()
      : v
    )),
    uCols,

    /* Violation definitions — from constants.js */
    violations   : VIOLATIONS,
    violColors   : VIOL_COLORS,

    /* Executive summary */
    summaryHTML,
    summaryIsCustom: !!(client?.customSummaryHTML),
    notes,
  };

  /* ── Serialise payload — catch any unexpected non-serialisable values ── */
  let payloadJson;
  try {
    payloadJson = JSON.stringify(payload)
      /* Escape </script> sequences so the JSON data island cannot break the
         HTML parser. Use \u003c (Unicode) not <\/ (backslash) — the backslash
         form is NOT valid JSON and causes JSON.parse to throw when the
         exported file is opened in a browser.                                */
      .replace(/</g, '\\u003c')
      .replace(/>/g, '\\u003e')
      .replace(/&/g, '\\u0026');
  } catch (err) {
    showToast('error', 'Export failed',
      `Could not serialise report data: ${err.message}`, 0);
    return;
  }

  /* ── Generate and download the standalone HTML file ── */
  const html = buildExportHTML(payloadJson);
  const blob = new Blob([html], { type: 'text/html' });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');

  a.href            = url;
  a.download        = `${clientName}_Fleet_Report_${new Date().toISOString().slice(0, 10)}.html`;
  a.style.display   = 'none';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);

  /* Release object URL after a short delay to allow the download to start */
  setTimeout(() => URL.revokeObjectURL(url), 1000);

  showToast('success', 'Report exported',
    `${clientName}_Fleet_Report downloaded successfully.`, 3000);
}


/* ═══════════════════════════════════════════════════════════════════════════
 * 2. buildExportHTML()
 * Constructs the full standalone HTML document as a string.
 *
 * The document is intentionally self-contained:
 *   - Chart.js loaded from CDN (cdnjs)
 *   - All CSS inlined in <style>
 *   - All data embedded in <script type="application/json">
 *   - All JS inlined in <script>
 *   - No external file references except the CDN
 * ═══════════════════════════════════════════════════════════════════════════ */

function buildExportHTML(payloadJson) {
  return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Fleet Report</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.js"><\/script>
<style>
/* ── Design tokens ── */
:root {
  --bg      : #0f1117; --surface : #1a1d27; --surface2: #22263a;
  --border  : rgba(255,255,255,0.08); --border2: rgba(255,255,255,0.14);
  --text    : #f0f3fb; --text2   : #e3e9f8; --text3   : #cfd8ee;
  --accent  : #4f8ef7; --red     : #e05353; --amber   : #e09545;
  --green   : #3db87a; --teal    : #2ec4b6;
  --radius  : 10px;    --radius-lg: 14px;
}
*{box-sizing:border-box;margin:0;padding:0}
body{background:var(--bg);color:var(--text);font-family:'Segoe UI',system-ui,sans-serif;font-size:15px;min-height:100vh}
::-webkit-scrollbar{width:6px;height:6px}
::-webkit-scrollbar-track{background:var(--surface2);border-radius:10px}
::-webkit-scrollbar-thumb{background:var(--border2);border-radius:10px}
::-webkit-scrollbar-thumb:hover{background:var(--accent)}
*{scrollbar-width:thin;scrollbar-color:var(--border2) var(--surface2)}

/* ── Topbar ── */
.topbar{background:var(--surface);border-bottom:1px solid var(--border);padding:0 1.5rem;height:56px;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:100}
.topbar-brand{font-size:13px;font-weight:600;letter-spacing:.08em;color:var(--text2);text-transform:uppercase}
.topbar-brand span{color:var(--accent)}
.topbar-badge{font-size:11px;background:rgba(61,184,122,.15);color:var(--green);padding:3px 10px;border-radius:20px;border:1px solid rgba(61,184,122,.3)}
.topbar-right{display:flex;gap:8px}
/* ── Layout ── */
.content{padding:1.5rem;max-width:1400px;margin:0 auto}
.report-header{margin-bottom:1.5rem}
.report-header h2{font-size:18px;font-weight:600}
.report-header p{color:var(--text2);font-size:12px;margin-top:4px}

/* ── KPI grid ── */
.kpi-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(148px,1fr));gap:10px;margin-bottom:1.5rem}
.kpi{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);padding:1rem}
.kpi-label{font-size:12px;color:var(--text2);text-transform:uppercase;letter-spacing:.06em;margin-bottom:6px}
.kpi-value{font-size:24px;font-weight:600;color:var(--text)}
.kpi-sub{font-size:12px;color:var(--text2);margin-top:4px}
.kpi-change{font-size:11px;margin-top:4px;font-weight:500}
.kpi-up{color:var(--red)}.kpi-down{color:var(--green)}.kpi-same{color:var(--text3)}

/* ── Cards & charts ── */
.grid-2{display:grid;grid-template-columns:1fr 1fr;gap:1rem;margin-bottom:1rem}
.card{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius-lg);padding:1.25rem;overflow:clip}
.card-title{font-size:13px;font-weight:600;color:var(--text2);text-transform:uppercase;letter-spacing:.06em;margin-bottom:1rem}
.chart-wrap{position:relative;width:100%}
.section-title{font-size:11px;color:var(--text3);text-transform:uppercase;letter-spacing:.08em;margin-bottom:10px;padding-bottom:6px;border-bottom:1px solid var(--border)}

/* ── Tables ── */
.tbl-wrap{overflow-x:auto}
table{width:100%;border-collapse:collapse;font-size:12px}
th{text-align:left;padding:8px 10px;color:var(--text3);font-weight:500;font-size:11px;text-transform:uppercase;letter-spacing:.05em;border-bottom:1px solid var(--border)}
td{padding:8px 10px;border-bottom:1px solid var(--border);color:var(--text)}
tr:last-child td{border-bottom:none}
tr:hover td{background:var(--surface2)}
.badge{display:inline-block;padding:2px 8px;border-radius:4px;font-size:11px;font-weight:500}
.b-red{background:rgba(224,83,83,.15);color:var(--red)}
.b-amber{background:rgba(224,149,69,.15);color:var(--amber)}
.b-green{background:rgba(61,184,122,.15);color:var(--green)}
.arr-up{color:var(--red)}.arr-down{color:var(--green)}

/* ── Legend ── */
.legend-row{display:flex;flex-wrap:wrap;gap:12px;margin-bottom:10px}
.leg{display:flex;align-items:center;gap:5px;font-size:11px;color:var(--text2)}
.leg-dot{width:10px;height:10px;border-radius:2px;flex-shrink:0}

/* ── Violation cards ── */
.viol-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(260px,1fr));gap:8px;margin-bottom:1.5rem}
.viol-card{background:var(--surface2);border-radius:var(--radius);padding:12px 14px;border-left:3px solid var(--border2)}
.viol-name{font-size:12px;font-weight:600;color:var(--text);margin-bottom:4px}
.viol-desc{font-size:11px;color:var(--text2);line-height:1.5}
.viol-high{border-left-color:var(--red)}
.viol-med{border-left-color:var(--amber)}
.viol-low{border-left-color:var(--green)}

/* ── Executive summary ── */
.exec-summary{background:var(--surface);border:1px solid var(--border2);border-radius:var(--radius-lg);padding:1.5rem;margin-bottom:1.5rem}
.exec-summary-title{font-size:11px;font-weight:600;color:var(--text3);text-transform:uppercase;letter-spacing:.08em;margin-bottom:1rem}
.exec-summary-body{font-size:14px;color:var(--text2);line-height:1.85}
.exec-summary-body p{margin-bottom:.65rem;padding-bottom:.65rem;border-bottom:1px solid rgba(255,255,255,.05)}
.exec-summary-body p:last-child{margin-bottom:0;padding-bottom:0;border-bottom:none}
.exec-summary-body strong{color:var(--text)}
.exec-summary-body .good{color:var(--green);font-weight:600}
.exec-summary-body .bad{color:var(--red);font-weight:600}
.exec-summary-body .neutral{color:var(--amber);font-weight:600}
.exec-notes-wrap{margin-top:1rem;padding-top:1rem;border-top:1px solid var(--border)}
.exec-notes-label{font-size:11px;color:var(--text3);text-transform:uppercase;letter-spacing:.06em;margin-bottom:6px}
.exec-notes-readonly{font-size:13px;color:var(--text2);line-height:1.6;white-space:pre-wrap}

/* ── Filter bar ── */
.global-filter-bar{background:var(--surface);border:1px solid var(--border2);border-radius:var(--radius-lg);padding:12px 1.25rem;margin-bottom:1rem;display:flex;align-items:center;gap:12px;flex-wrap:wrap}
.global-filter-label{font-size:11px;color:var(--text3);text-transform:uppercase;letter-spacing:.06em;white-space:nowrap}
.global-filter-toggle{display:flex;border:1px solid var(--border2);border-radius:var(--radius);overflow:hidden;flex-shrink:0}
.global-filter-toggle button{padding:6px 16px;font-size:12px;border:none;cursor:pointer;transition:all .15s;white-space:nowrap}
.compare-tags{display:flex;flex-wrap:wrap;align-items:center;gap:6px;flex:1}
.compare-tag{display:flex;align-items:center;gap:4px;padding:3px 8px;border-radius:4px;font-size:11px;font-weight:500;color:#fff;white-space:nowrap}
.compare-tag-x{background:none;border:none;color:rgba(255,255,255,0.7);cursor:pointer;font-size:13px;line-height:1;padding:0;margin-left:2px}
.compare-tag-x:hover{color:#fff}
.compare-add{background:none;border:1px dashed var(--border2);color:var(--text3);border-radius:4px;padding:3px 10px;font-size:11px;cursor:pointer;white-space:nowrap}
.compare-add:hover{border-color:var(--accent);color:var(--accent)}
.compare-add:disabled{opacity:.4;cursor:not-allowed}
.compare-limit{font-size:10px;color:var(--text3);white-space:nowrap}
.date-range-wrap{display:flex;align-items:center;gap:8px;flex-wrap:wrap}
.date-range-label{font-size:11px;color:var(--text3);white-space:nowrap}
.date-input{background:rgba(255,255,255,.04);color:var(--text);border:1px solid var(--border2);border-radius:var(--radius);padding:6px 11px;font-size:13px;outline:none;cursor:pointer;color-scheme:dark}
.date-input::-webkit-calendar-picker-indicator{filter:invert(1);opacity:.85;cursor:pointer}
.date-range-btn{background:var(--accent);color:#fff;border:none;border-radius:var(--radius);padding:5px 14px;font-size:12px;cursor:pointer}
.date-range-reset{background:none;color:var(--text3);border:1px solid var(--border2);border-radius:var(--radius);padding:5px 12px;font-size:12px;cursor:pointer}

/* ── Toast ── */
#toastContainer{position:fixed;bottom:1.5rem;right:1.5rem;z-index:99999;display:flex;flex-direction:column;gap:8px;pointer-events:none}
.toast{display:flex;align-items:flex-start;gap:10px;padding:12px 16px;border-radius:var(--radius);font-size:13px;min-width:280px;max-width:380px;pointer-events:all;border:1px solid;animation:toastIn .25s ease}
@keyframes toastIn{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:translateY(0)}}
@keyframes toastOut{from{opacity:1}to{opacity:0;transform:translateY(8px)}}
.t-icon{font-size:14px;flex-shrink:0;margin-top:1px}.t-body{flex:1}
.t-title{font-weight:600;margin-bottom:2px}.t-msg{font-size:12px;color:var(--text2);line-height:1.5}
.t-close{background:none;border:none;color:var(--text3);cursor:pointer;font-size:16px;padding:0;flex-shrink:0}
.t-error{background:#1e1215;border-color:rgba(224,83,83,.4)}.t-error .t-title{color:var(--red)}
.t-warn{background:#1a170e;border-color:rgba(224,149,69,.4)}.t-warn .t-title{color:var(--amber)}
.t-success{background:#0e1a13;border-color:rgba(61,184,122,.4)}.t-success .t-title{color:var(--green)}
.t-info{background:#0e1320;border-color:rgba(79,142,247,.4)}.t-info .t-title{color:var(--accent)}

/* ── Responsive ── */
@media(max-width:900px){.grid-2{grid-template-columns:1fr}}
@media(max-width:600px){
  .content{padding:.75rem}
  .kpi-grid{grid-template-columns:1fr 1fr;gap:8px}
  .kpi{padding:.75rem}.kpi-value{font-size:20px}
  .grid-2{grid-template-columns:1fr}
  .viol-grid{grid-template-columns:1fr}
  table{min-width:500px}
  #toastContainer{left:1rem;right:1rem;bottom:1rem}
  .toast{min-width:unset;max-width:100%}
}
@media(max-width:380px){.kpi-grid{grid-template-columns:1fr}.kpi-value{font-size:18px}}
@media(min-width:1400px){
  .content{padding:2rem 3rem}
  .kpi-grid{grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:14px}
  .kpi-value{font-size:28px}
  .grid-2{gap:1.5rem}
  .card{padding:1.5rem}
  .viol-grid{grid-template-columns:repeat(auto-fill,minmax(300px,1fr));gap:12px}
}

</style>
</head>
<body>
<div id="toastContainer"></div>

<!-- Topbar -->
<div class="topbar">
  <div style="display:flex;align-items:center;gap:12px">
    <div class="topbar-brand">Data <span>Sphere</span></div>
    <span class="topbar-badge">Read-only report</span>
  </div>
</div>

<!-- Main content area — populated by JS below -->
<div class="content" id="main"></div>

<!-- Data payload — read by the inline script below -->
<script id="__payload__" type="application/json">` + payloadJson + `</script>

<script>
/* ════════════════════════════════════════════════════════════════════════════
 * Exported report script
 * Reads the embedded JSON payload and renders the full dashboard.
 * Self-contained — does not depend on any external JS files.
 * ════════════════════════════════════════════════════════════════════════════ */

/* ── Parse payload ── */
/* All top-level initialisation is wrapped in try/catch so that any parse
   or data error surfaces as a visible message instead of a blank page.   */
var D, VIOLATIONS, VIOL_COLORS, vehicles, prevMap, hasPrev,
    dailyTotals, dayLabels, vehicleDailyData, utilRows, uCols, dayKeys,
    dayInfo, totalDist, avgScore, vehiclesWithIdle, avgIdle,
    vehiclesWithTrips, flaggedCount, activeViolations, prevVehicles,
    prevAvgScore, prevTotalDist, violTotals, flaggedVehicles;

try {
  D           = JSON.parse(document.getElementById('__payload__').textContent);
  if (!D || typeof D !== 'object') throw new Error('Payload is not a valid object');

  VIOLATIONS      = D.violations  || [];
  VIOL_COLORS     = D.violColors  || [];
  vehicles        = D.vehicles    || [];
  prevMap         = D.prevMap     || {};
  hasPrev         = D.hasPrev     || false;
  dailyTotals     = D.dailyTotals || {};
  dayLabels       = D.dayLabels   || [];
  vehicleDailyData= D.vehicleDailyData || {};
  utilRows        = D.utilRows    || [];
  uCols           = D.uCols       || {};
  dayKeys         = D.dayKeys     || dayLabels || [];
} catch (initErr) {
  (function() {
    var main = document.getElementById('main');
    if (main) {
      main.innerHTML =
        '<div style="margin:2rem;padding:1.5rem;background:rgba(224,83,83,.1);border:1px solid rgba(224,83,83,.3);border-radius:10px;color:#e05353;font-size:13px;line-height:1.6">' +
        '<strong>Report failed to load — data error.</strong><br>' +
        'The embedded report data could not be read. This can happen if the file was corrupted during download or if the export was interrupted.<br><br>' +
        '<span style="font-family:monospace;font-size:11px;color:#f87171">' + String(initErr) + '</span>' +
        '</div>';
    }
  })();
  throw initErr; /* halt — nothing further can run without valid data */
}

/* Score band thresholds — mirrors constants.js */
const SCORE_BANDS = {
  safe     : { max: 20, label: 'Low Risk',      color: '#3db87a' },
  moderate : { max: 40, label: 'Moderate Risk',  color: '#e09545' },
  high     : { min: 41, label: 'High Risk',      color: '#e05353' },
};
const COMPARE_COLORS = ['#4f8ef7','#3db87a','#e09545','#a855f7','#2ec4b6'];

function getScoreBand(score) {
  const s = Number(score) || 0;
  if (s <= SCORE_BANDS.safe.max)     return SCORE_BANDS.safe;
  if (s <= SCORE_BANDS.moderate.max) return SCORE_BANDS.moderate;
  return SCORE_BANDS.high;
}
function isAtRisk(score) { return (Number(score) || 0) > SCORE_BANDS.safe.max; }

/* ── Reconstruct dayInfo ── */
dayInfo = (dayLabels || []).map(function(lbl, i) {
  var di     = D.dayInfoDates && D.dayInfoDates[i];
  var date   = di ? new Date(di.year, di.month - 1, di.day) : null;
  var dayNum = D.dayNums && D.dayNums[i] != null ? D.dayNums[i] : null;
  if (dayNum === null) {
    var s = String(lbl).trim();
    var m = s.match(/^[SMTWFsmtwf]-(\\d+)$/);
    if (m) dayNum = parseInt(m[1]);
    else {
      var sl = s.split('/');
      if (sl.length >= 2) dayNum = parseInt(sl[0]);
    }
  }
  return { label: lbl, date: date, dayNum: dayNum, key: dayKeys[i] || lbl };
});

/* ── Computed KPIs ── */
var eActiveDateFilter = null;
totalDist       = vehicles.reduce(function(s,v){return s+(v.totalDist||0);},0);
avgScore        = vehicles.length ? vehicles.reduce(function(s,v){return s+(v.score||0);},0)/vehicles.length : 0;
vehiclesWithIdle  = vehicles.filter(function(v){return v.daysIdle!==undefined;});
avgIdle         = vehiclesWithIdle.length ? vehiclesWithIdle.reduce(function(s,v){return s+(v.daysIdle||0);},0)/vehiclesWithIdle.length : 0;
vehiclesWithTrips = vehicles.filter(function(v){return (v.daysActive||0)>0;}).length;
flaggedCount    = vehicles.filter(function(v){return isAtRisk(v.score);}).length;
activeViolations  = VIOLATIONS.filter(function(v){return vehicles.some(function(vh){return (vh[v.key]||0)>0;});});
prevVehicles    = Object.values(prevMap);
prevAvgScore    = prevVehicles.length ? prevVehicles.reduce(function(s,v){return s+(v.score||0);},0)/prevVehicles.length : null;
prevTotalDist   = prevVehicles.reduce(function(s,v){return s+(v.totalDist||0);},0);
violTotals      = activeViolations.map(function(v){return vehicles.reduce(function(s,vh){return s+(vh[v.key]||0);},0);});
flaggedVehicles = vehicles.filter(function(v){return isAtRisk(v.score);}).sort(function(a,b){return (b.score||0)-(a.score||0);});

/* ── Selected vehicle & compare state ── */
var eRiskSel=''; var eWdweSel=''; var eUtilSel=''; var ePrevSel=''; var eDistSel='';
var eCompareSelected=[]; var eFilterMode='fleet';
const COMPARE_MAX=5;
function getCompareColor(idx){return COMPARE_COLORS[idx%COMPARE_COLORS.length];}

/* ── Utility: set element text ── */
function setEl(id,text){var el=document.getElementById(id);if(el)el.textContent=text;}
function setKpi(n,label,value,sub){setEl('eKL'+n,label);setEl('eKV'+n,value);setEl('eKS'+n,sub);}

/* ── changeTag() ── */
function changeTag(curr,prev,lowerIsBetter){
  if(prev===null||prev===undefined||prev===0)return'';
  var diff=curr-prev;var rawPct=Math.abs(diff)/Math.abs(prev)*100;
  if(rawPct<1)return'<span class="kpi-change kpi-same">→ no change</span>';
  var improved=lowerIsBetter?diff<0:diff>0;
  var cls=improved?'kpi-down':'kpi-up';var arrow=diff>0?'▲':'▼';
  if(rawPct>100){var absDiff=Math.round(Math.abs(diff));var sign=diff>0?'+':'-';return'<span class="kpi-change '+cls+'">'+arrow+' '+sign+absDiff+' vs last month</span>';}
  return'<span class="kpi-change '+cls+'">'+arrow+' '+Math.round(rawPct)+'% vs last month</span>';
}

/* ── Toast ── */
function showToast(type,title,msg,duration){
  if(duration===undefined)duration=6000;
  var icons={error:'✕',warn:'⚠',success:'✓',info:'i'};
  var box=document.createElement('div');
  box.className='toast t-'+type;
  box.innerHTML='<span class="t-icon">'+(icons[type]||'i')+'</span><div class="t-body"><div class="t-title">'+escHTML(title)+'</div>'+(msg?'<div class="t-msg">'+escHTML(msg)+'</div>':'')+'</div><button class="t-close" onclick="this.parentNode.remove()">×</button>';
  document.getElementById('toastContainer').appendChild(box);
  if(duration>0)setTimeout(function(){box.style.animation='toastOut .25s ease forwards';setTimeout(function(){box.remove();},250);},duration);
}
function escHTML(s){return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;');}

/* ══════════════════════════════════════════════════════════════════════════
 * BUILD DASHBOARD HTML
 * ══════════════════════════════════════════════════════════════════════════ */
function buildDashboard(){
  var hasDailyData=Object.keys(dailyTotals).length>0;
  var main=document.getElementById('main');
  main.innerHTML=
    '<div class="report-header">'+
      '<h2>'+escHTML(D.clientName)+' — Fleet Performance Report</h2>'+
      '<p>'+vehicles.length+' vehicles &nbsp;·&nbsp; Generated '+escHTML(D.generatedOn)+(D.reportMonth?' &nbsp;·&nbsp; '+escHTML(D.reportMonth):'')+'</p>'+
    '</div>'+

    '<div class="global-filter-bar">'+
      '<span class="global-filter-label">Filter</span>'+
      '<div class="global-filter-toggle">'+
        '<button id="eBtnFleet"   onclick="eSetMode(\\'fleet\\')"   style="background:var(--accent);color:#fff">Whole fleet</button>'+
        '<button id="eBtnVehicle" onclick="eSetMode(\\'vehicle\\')" style="background:transparent;color:var(--text2)">Single vehicle</button>'+
        '<button id="eBtnCompare" onclick="eSetMode(\\'compare\\')" style="background:transparent;color:var(--text2)">Compare</button>'+
      '</div>'+
      '<div id="eComparePanel" style="display:none;align-items:center;gap:8px;flex-wrap:wrap;flex:1">'+
        '<div class="compare-tags" id="eCompareTags"></div>'+
        '<div style="position:relative">'+
          '<button class="compare-add" id="eCompareAddBtn" onclick="eToggleCompareDrop()">+ Add vehicle</button>'+
          '<span class="compare-limit" id="eCompareLimit">0/5</span>'+
          '<div id="eCompareDropPanel" style="display:none;position:absolute;top:calc(100% + 4px);left:0;min-width:220px;background:var(--surface);border:1px solid var(--border2);border-radius:var(--radius);z-index:9999;box-shadow:0 8px 24px rgba(0,0,0,.5)">'+
            '<div style="padding:6px"><input id="eCompareSearch" type="text" placeholder="Search plate..." oninput="eBuildCompareDrop(this.value)" style="width:100%;background:var(--surface2);color:var(--text);border:1px solid var(--border2);border-radius:6px;padding:6px 8px;font-size:12px;outline:none"/></div>'+
            '<div id="eCompareDropList" style="max-height:220px;overflow-y:auto;padding-bottom:4px"></div>'+
          '</div>'+
        '</div>'+
      '</div>'+
      '<div id="eVDrop" style="display:none;position:relative;width:220px">'+
        '<div onclick="eToggleVDrop()" style="display:flex;align-items:center;justify-content:space-between;background:var(--surface2);border:1px solid var(--border2);border-radius:var(--radius);padding:6px 10px;font-size:12px;cursor:pointer;user-select:none">'+
          '<span id="eVDropLbl" style="color:var(--text3)">Select vehicle...</span>'+
          '<span style="color:var(--text3);font-size:10px;margin-left:6px">&#9662;</span>'+
        '</div>'+
        '<div id="eVDropPanel" style="display:none;position:absolute;top:calc(100% + 4px);left:0;right:0;background:var(--surface);border:1px solid var(--border2);border-radius:var(--radius);z-index:9999;box-shadow:0 8px 24px rgba(0,0,0,.5)">'+
          '<div style="padding:6px"><input id="eVDropSearch" type="text" placeholder="Search plate..." oninput="eBuildVDrop(this.value)" style="width:100%;background:var(--surface2);color:var(--text);border:1px solid var(--border2);border-radius:6px;padding:6px 8px;font-size:12px;outline:none"/></div>'+
          '<div id="eVDropList" style="max-height:220px;overflow-y:auto;padding-bottom:4px"></div>'+
        '</div>'+
        '<select id="eVehicleSelect" style="display:none"></select>'+
      '</div>'+
      '<div style="width:1px;background:var(--border2);align-self:stretch;margin:0 4px"></div>'+
      '<div class="date-range-wrap">'+
        '<span class="date-range-label">From</span>'+
        '<input type="date" class="date-input" id="eDateFrom"/>'+
        '<span class="date-range-label">To</span>'+
        '<input type="date" class="date-input" id="eDateTo"/>'+
        '<button class="date-range-btn" onclick="eApplyDateRange()">Apply</button>'+
        '<button class="date-range-reset" id="eDateResetBtn" onclick="eResetAllFilters()" style="display:none">Reset all</button>'+
      '</div>'+
    '</div>'+

    '<div class="exec-summary">'+
      '<div class="exec-summary-title">Executive Summary</div>'+
      '<div class="exec-summary-body" id="eExecBody"></div>'+
      (D.notes?'<div class="exec-notes-wrap"><div class="exec-notes-label">Additional notes</div><div class="exec-notes-readonly">'+escHTML(D.notes)+'</div></div>':'')+
    '</div>'+

    '<div class="kpi-grid">'+
      '<div class="kpi"><div class="kpi-label" id="eKL0">Total vehicles</div><div class="kpi-value" id="eKV0">'+vehicles.length+'</div><div class="kpi-sub" id="eKS0">in the fleet</div></div>'+
      '<div class="kpi"><div class="kpi-label" id="eKL1">Total distance</div><div class="kpi-value" id="eKV1">'+Math.round(totalDist).toLocaleString()+'</div><div class="kpi-sub" id="eKS1">km fleet total</div><span id="eKC1">'+changeTag(totalDist,prevTotalDist,false)+'</span></div>'+
      '<div class="kpi"><div class="kpi-label" id="eKL2">Fleet avg score</div><div class="kpi-value" id="eKV2">'+Math.round(avgScore).toLocaleString()+'</div><div class="kpi-sub" id="eKS2">lower is safer</div>'+(hasPrev&&prevAvgScore!==null?changeTag(avgScore,prevAvgScore,true):'')+'</div>'+
      '<div class="kpi"><div class="kpi-label" id="eKL3">Avg idle days</div><div class="kpi-value" id="eKV3">'+avgIdle.toFixed(1)+'</div><div class="kpi-sub" id="eKS3">per vehicle</div></div>'+
      '<div class="kpi"><div class="kpi-label" id="eKL4">Vehicles active</div><div class="kpi-value" id="eKV4">'+vehiclesWithTrips+'</div><div class="kpi-sub" id="eKS4">recorded trips this month</div></div>'+
      '<div class="kpi"><div class="kpi-label" id="eKL5">At-risk vehicles</div><div class="kpi-value" id="eKV5" style="color:var(--red)">'+flaggedCount+'</div><div class="kpi-sub" id="eKS5">Moderate or High risk (score 21+)</div></div>'+
      '<div class="kpi"><div class="kpi-label" id="eKL6">Violation types</div><div class="kpi-value" id="eKV6">'+activeViolations.length+'</div><div class="kpi-sub" id="eKS6">detected in data</div></div>'+
    '</div>'+

    (hasDailyData?'<div class="card" style="margin-bottom:1rem"><div style="margin-bottom:1rem"><div class="card-title" id="eDCTitle">Daily fleet distance (km)</div><div style="font-size:11px;color:var(--text3)" id="eDCSub">Total km covered by all vehicles each day</div></div><div class="chart-wrap" style="height:180px"><canvas id="eDailyChart"></canvas></div></div>':'')+

    '<div class="grid-2" style="margin-bottom:1rem">'+
      '<div class="card"><div class="card-title" id="eRiskTitle">All vehicles — advanced score ranking</div><div id="eRiskScroll" style="overflow-y:auto;max-height:320px"><div id="eRiskWrap" style="position:relative;height:320px"><canvas id="eRiskChart"></canvas></div></div></div>'+
      '<div class="card"><div class="card-title" id="eViolTitle">Violation breakdown — fleet total</div><div class="legend-row" id="eViolLeg"></div><div class="chart-wrap" style="height:280px"><canvas id="eViolChart"></canvas></div></div>'+
    '</div>'+

    '<div class="grid-2" style="margin-bottom:1rem">'+
      '<div class="card"><div class="card-title" id="eWdweTitle">Weekday vs weekend distance</div><div class="legend-row"><span class="leg"><span class="leg-dot" style="background:#4f8ef7"></span>Weekday</span><span class="leg"><span class="leg-dot" style="background:#2ec4b6"></span>Weekend</span></div><div id="eWdweScroll" style="overflow-y:auto;max-height:320px"><div id="eWdweWrap" style="position:relative;height:320px"><canvas id="eWdweChart"></canvas></div></div></div>'+
      '<div class="card"><div class="card-title" id="eUtilTitle">Most idle vehicles — days active vs idle</div><div class="legend-row"><span class="leg"><span class="leg-dot" style="background:#4f8ef7"></span>Active</span><span class="leg"><span class="leg-dot" style="background:#555b72"></span>Idle</span></div><div id="eUtilScroll" style="overflow-y:auto;max-height:320px"><div id="eUtilWrap" style="position:relative;height:320px"><canvas id="eUtilChart"></canvas></div></div></div>'+
    '</div>'+

    (hasPrev?'<div class="card" style="margin-bottom:1rem"><div class="card-title" id="ePrevTitle">Month-on-month Advance score — current vs previous</div><div class="legend-row"><span class="leg"><span class="leg-dot" style="background:#3b6edc"></span>Current</span><span class="leg"><span class="leg-dot" style="background:#6ea8ff"></span>Previous</span></div><div id="ePrevScroll" style="overflow-y:auto;max-height:400px"><div id="ePrevWrap" style="position:relative;height:400px"><canvas id="ePrevChart"></canvas></div></div></div>':'')+

    (hasPrev?'<div class="card print-section-distcomp" style="margin-bottom:1rem"><div class="card-title" id="eDCCompTitle">Distance comparison — current vs previous month</div><div class="legend-row"><span class="leg"><span class="leg-dot" style="background:#3db87a"></span>Increased</span><span class="leg"><span class="leg-dot" style="background:#e05353"></span>Decreased</span><span class="leg"><span class="leg-dot" style="background:rgba(79,142,247,0.45)"></span>Previous month</span></div><div id="eDCCompScroll" style="overflow-y:auto;max-height:400px"><div id="eDCCompWrap" style="position:relative;height:400px"><canvas id="eDCCompChart"></canvas></div></div></div>':'')+

    '<div class="card print-section-flagged" style="margin-bottom:1.5rem">'+
      '<div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:10px;flex-wrap:wrap;gap:8px">'+
        '<div class="card-title" id="eFlaggedTitle" style="margin-bottom:0">Vehicles Scoring</div>'+
        '<div style="display:flex;gap:6px">'+
          '<button id="eBtnFlaggedOnly" onclick="eSetTableMode(\\'flagged\\')" style="padding:4px 10px;font-size:11px;border-radius:20px;border:1px solid var(--accent);background:var(--accent);color:#fff;cursor:pointer;font-weight:500">Flagged only</button>'+
          '<button id="eBtnAllVehicles" onclick="eSetTableMode(\\'all\\')" style="padding:4px 10px;font-size:11px;border-radius:20px;border:1px solid var(--border2);background:transparent;color:var(--text2);cursor:pointer">All vehicles</button>'+
        '</div>'+
      '</div>'+
      '<div class="tbl-wrap" style="max-height:420px;overflow-x:auto;overflow-y:auto;width:100%">'+
        '<table style="min-width:'+Math.max(900,500+activeViolations.length*90)+'px;white-space:nowrap"><thead><tr><th>Vehicle</th><th>Total km</th>'+(hasPrev?'<th>Prev km</th><th>Dist Change</th>':'')+'<th>Advanced Score</th>'+
        (hasPrev?'<th>Prev Score</th><th>Score Change</th>':'')+
        '<th>Days Active</th><th>Days Idle</th><th>Weekday km</th><th>Weekend km</th>'+
        activeViolations.map(function(v){return'<th>'+escHTML(v.short)+'</th>';}).join('')+
        (function(){var ek=[];vehicles.forEach(function(v){Object.keys(v._extra||{}).forEach(function(k){if(ek.indexOf(k)===-1)ek.push(k);});});return ek.map(function(k){return'<th>'+escHTML(k)+'</th>';}).join('');})() +
        '<th>Risk</th></tr></thead><tbody id="eAlertBody"></tbody></table>'+
      '</div>'+
    '</div>'+

    '<div class="card print-section-ranked" style="margin-bottom:1rem">'+
      '<div class="card-title">Top Ranked Vehicles Per 100 KM</div>'+
      '<div style="font-size:12px;color:var(--text3);margin-bottom:12px">Ranked by violations per 100 km — lower is safer.</div>'+
      '<div class="tbl-wrap"><table><thead><tr><th>Rank</th><th>Vehicle</th><th>Total km</th><th>Violations</th><th>Per 100 km</th><th>Advance Score</th><th>Rating</th></tr></thead><tbody id="eBestPerfBody"></tbody></table></div>'+
    '</div>'+

    '<div class="section-title" style="margin-bottom:12px">Violation reference guide</div>'+
    '<div class="viol-grid" id="eViolGuide"></div>'+

    '<div class="card" style="margin-bottom:1rem">'+
      '<div class="card-title">Advanced Score Grading Guide</div>'+
      '<div style="font-size:12px;color:var(--text3);margin-bottom:14px">Lower score = safer driver. Use this guide to interpret scores.</div>'+
      '<div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:12px">'+
        '<div style="background:rgba(61,184,122,.08);border:1px solid rgba(61,184,122,.25);border-radius:8px;padding:14px"><div style="font-size:18px;font-weight:700;color:var(--green);margin-bottom:4px">1 – 20</div><div style="font-size:13px;font-weight:600;color:var(--text);margin-bottom:6px">✓ Safe Driving</div><div style="font-size:11px;color:var(--text3);line-height:1.5">Excellent behaviour. Vehicle is within acceptable safety parameters.</div></div>'+
        '<div style="background:rgba(255,183,77,.08);border:1px solid rgba(255,183,77,.25);border-radius:8px;padding:14px"><div style="font-size:18px;font-weight:700;color:#f7c04f;margin-bottom:4px">21 – 40</div><div style="font-size:13px;font-weight:600;color:var(--text);margin-bottom:6px">⚠ Needs Attention</div><div style="font-size:11px;color:var(--text3);line-height:1.5">Moderate risk. Driver coaching and monitoring recommended.</div></div>'+
        '<div style="background:rgba(224,83,83,.08);border:1px solid rgba(224,83,83,.25);border-radius:8px;padding:14px"><div style="font-size:18px;font-weight:700;color:var(--red);margin-bottom:4px">41+</div><div style="font-size:13px;font-weight:600;color:var(--text);margin-bottom:6px">✗ High Risk</div><div style="font-size:11px;color:var(--text3);line-height:1.5">Serious violations detected. Immediate intervention required.</div></div>'+
      '</div>'+
    '</div>';

  buildCharts();
  buildVehicleDropdown();
  buildViolationGuide();
  setDateDefaults();
  renderVehicleTable();
  renderBestPerfTable();
  eGenerateSummary();
}

/* ══════════════════════════════════════════════════════════════════════════
 * CHART INSTANCES
 * ══════════════════════════════════════════════════════════════════════════ */
if(typeof Chart!=='undefined')Chart.defaults.color='#8b90a7';
var eDailyChart=null,eRiskChart=null,eViolChart=null,eWdweChart=null,eUtilChart=null,ePrevChart=null,eDCCompChart=null;

function tickColor(label,selGetter){
  if(label===selGetter())return'#3db87a';
  var ci=eCompareSelected.indexOf(label);
  if(ci>-1)return getCompareColor(ci);
  return'#8b90a7';
}
function tickFont(label,selGetter){
  if(label===selGetter()||eCompareSelected.indexOf(label)>-1)return{size:10,weight:'bold'};
  return{size:10};
}

function buildCharts(){
  var hasDailyData=Object.keys(dailyTotals).length>0;
  var allRisk=[...vehicles].sort(function(a,b){return(b.score||0)-(a.score||0);});
  var allWdwe=[...vehicles].sort(function(a,b){return(b.totalDist||0)-(a.totalDist||0);});
  var allUtil=[...vehicles].filter(function(v){return v.daysIdle!==undefined;}).sort(function(a,b){return(b.daysIdle||0)-(a.daysIdle||0);});

  var riskH=Math.max(320,allRisk.length*32);
  var wdweH=Math.max(320,allWdwe.length*32);
  var utilH=Math.max(320,allUtil.length*32);

  document.getElementById('eRiskWrap').style.height=riskH+'px';
  document.getElementById('eWdweWrap').style.height=wdweH+'px';
  document.getElementById('eUtilWrap').style.height=utilH+'px';

  if(hasDailyData){
    eDailyChart=new Chart(document.getElementById('eDailyChart'),{
      type:'line',data:{labels:dayLabels,datasets:[{label:'Fleet total',data:Object.values(dailyTotals),borderColor:'#4f8ef7',backgroundColor:'rgba(79,142,247,0.08)',fill:true,tension:0.3,pointRadius:3,borderWidth:2,pointHoverRadius:5,pointBackgroundColor:'#4f8ef7'}]},
      options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:{callbacks:{label:function(ctx){return Math.round(ctx.raw).toLocaleString()+' km';}}}},scales:{x:{ticks:{font:{size:10},maxRotation:45,color:'#555b72'},grid:{color:'rgba(255,255,255,0.04)'},border:{display:false}},y:{ticks:{callback:function(v){return v.toLocaleString()+' km';},color:'#555b72'},grid:{color:'rgba(255,255,255,0.04)'},border:{display:false}}}}
    });
  }

  eRiskChart=new Chart(document.getElementById('eRiskChart'),{
    type:'bar',data:{labels:allRisk.map(function(v){return v.name;}),datasets:[{data:allRisk.map(function(v){return v.score||0;}),backgroundColor:allRisk.map(function(){return'#e05353';}),borderRadius:3}]},
    options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{x:{ticks:{callback:function(v){return v.toLocaleString();},color:'#555b72'},grid:{color:'rgba(255,255,255,0.04)'},border:{display:false}},y:{ticks:{font:function(ctx){return tickFont(ctx.tick.label,function(){return eRiskSel;});},color:function(ctx){return tickColor(ctx.tick.label,function(){return eRiskSel;});}}}}}
  });

  document.getElementById('eViolLeg').innerHTML=activeViolations.map(function(v,idx){return'<span class="leg"><span class="leg-dot" style="background:'+VIOL_COLORS[idx%VIOL_COLORS.length]+'"></span>'+escHTML(v.short)+'</span>';}).join('');
  eViolChart=new Chart(document.getElementById('eViolChart'),{
    type:'doughnut',data:{labels:activeViolations.map(function(v){return v.short;}),datasets:[{data:violTotals,backgroundColor:VIOL_COLORS.slice(0,activeViolations.length),borderWidth:0}]},
    options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}}}
  });

  eWdweChart=new Chart(document.getElementById('eWdweChart'),{
    type:'bar',data:{labels:allWdwe.map(function(v){return v.name;}),datasets:[{label:'Weekday',data:allWdwe.map(function(v){return v.weekdayDist||0;}),backgroundColor:'#4f8ef7',borderRadius:2},{label:'Weekend',data:allWdwe.map(function(v){return v.weekendDist||0;}),backgroundColor:'#2ec4b6',borderRadius:2}]},
    options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{x:{stacked:true,ticks:{callback:function(v){return v.toLocaleString();},color:'#555b72'},grid:{color:'rgba(255,255,255,0.04)'},border:{display:false}},y:{stacked:true,ticks:{font:function(ctx){return tickFont(ctx.tick.label,function(){return eWdweSel;});},color:function(ctx){return tickColor(ctx.tick.label,function(){return eWdweSel;});}}}}}
  });

  eUtilChart=new Chart(document.getElementById('eUtilChart'),{
    type:'bar',data:{labels:allUtil.map(function(v){return v.name;}),datasets:[{label:'Active',data:allUtil.map(function(v){return v.daysActive||0;}),backgroundColor:'#4f8ef7',borderRadius:2},{label:'Idle',data:allUtil.map(function(v){return v.daysIdle||0;}),backgroundColor:'#555b72',borderRadius:2}]},
    options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{x:{stacked:true,ticks:{color:'#555b72'},grid:{color:'rgba(255,255,255,0.04)'},border:{display:false}},y:{stacked:true,ticks:{font:function(ctx){return tickFont(ctx.tick.label,function(){return eUtilSel;});},color:function(ctx){return tickColor(ctx.tick.label,function(){return eUtilSel;});}}}}}
  });

  if(hasPrev){
    var allPrev=vehicles.filter(function(v){return prevMap[v.name];}).sort(function(a,b){return(b.score||0)-(a.score||0);});
    var prevH=Math.max(400,allPrev.length*36);
    document.getElementById('ePrevWrap').style.height=prevH+'px';
    ePrevChart=new Chart(document.getElementById('ePrevChart'),{
      type:'bar',data:{labels:allPrev.map(function(v){return v.name;}),datasets:[{label:'Current',data:allPrev.map(function(v){return v.score||0;}),backgroundColor:allPrev.map(function(){return'#3b6edc';}),borderRadius:3},{label:'Previous',data:allPrev.map(function(v){return prevMap[v.name]?prevMap[v.name].score||0:0;}),backgroundColor:'#6ea8ff',borderRadius:3}]},
      options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{x:{ticks:{callback:function(v){return v.toLocaleString();},color:'#555b72'},grid:{color:'rgba(255,255,255,0.04)'},border:{display:false}},y:{ticks:{font:function(ctx){return tickFont(ctx.tick.label,function(){return ePrevSel;});},color:function(ctx){return tickColor(ctx.tick.label,function(){return ePrevSel;});}}}}}
    });

    var allNames=new Set([...vehicles.map(function(v){return v.name;}),...Object.keys(prevMap)]);
    var allDC=[...allNames].map(function(name){var curr=vehicles.find(function(v){return v.name===name;});return{name:name,currDist:curr?curr.totalDist||0:0,prevDist:prevMap[name]?prevMap[name].totalDist||0:0};}).sort(function(a,b){return b.currDist-a.currDist;});
    var dcH=Math.max(400,allDC.length*34);
    document.getElementById('eDCCompWrap').style.height=dcH+'px';
    eDCCompChart=new Chart(document.getElementById('eDCCompChart'),{
      type:'bar',data:{labels:allDC.map(function(v){return v.name;}),datasets:[{label:'Current',data:allDC.map(function(v){return v.currDist;}),backgroundColor:allDC.map(function(v){return v.currDist>=v.prevDist?'#3db87a':'#e05353';}),borderRadius:3},{label:'Previous',data:allDC.map(function(v){return v.prevDist;}),backgroundColor:allDC.map(function(v){return v.currDist>=v.prevDist?'#e05353':'#3db87a';}),borderRadius:3}]},
      options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:{callbacks:{label:function(ctx){return ctx.dataset.label+': '+Math.round(ctx.raw).toLocaleString()+' km';}}}},scales:{x:{ticks:{callback:function(v){return v.toLocaleString()+' km';},color:'#555b72'},grid:{color:'rgba(255,255,255,0.04)'},border:{display:false}},y:{ticks:{font:function(ctx){return tickFont(ctx.tick.label,function(){return eDistSel;});},color:function(ctx){return tickColor(ctx.tick.label,function(){return eDistSel;});}},border:{display:false}}}}
    });
  }
}

function eRestoreChartsToFleet(){
  var allRisk=[...vehicles].sort(function(a,b){return(b.score||0)-(a.score||0);});
  var allWdwe=[...vehicles].sort(function(a,b){return(b.totalDist||0)-(a.totalDist||0);});
  var allUtil=[...vehicles].filter(function(v){return v.daysIdle!==undefined;}).sort(function(a,b){return(b.daysIdle||0)-(a.daysIdle||0);});
  var riskH=Math.max(320,allRisk.length*32);
  var wdweH=Math.max(320,allWdwe.length*32);
  var utilH=Math.max(320,allUtil.length*32);
  document.getElementById('eRiskWrap').style.height=riskH+'px';
  document.getElementById('eRiskScroll').style.maxHeight='320px';
  document.getElementById('eWdweWrap').style.height=wdweH+'px';
  document.getElementById('eWdweScroll').style.maxHeight='320px';
  document.getElementById('eUtilWrap').style.height=utilH+'px';
  document.getElementById('eUtilScroll').style.maxHeight='320px';
  if(eDailyChart){
    eDailyChart.data.labels=dayLabels;
    eDailyChart.data.datasets=[{label:'Fleet total',data:Object.values(dailyTotals),borderColor:'#4f8ef7',backgroundColor:'rgba(79,142,247,0.08)',fill:true,tension:0.3,pointRadius:3,borderWidth:2,pointHoverRadius:5,pointBackgroundColor:'#4f8ef7'}];
    eDailyChart.update();
  }
  setEl('eDCTitle','Daily fleet distance (km)');setEl('eDCSub','Total km covered by all vehicles each day');
  if(eRiskChart){
    eRiskChart.data.labels=allRisk.map(function(v){return v.name;});
    eRiskChart.data.datasets[0].data=allRisk.map(function(v){return v.score||0;});
    eRiskChart.data.datasets[0].backgroundColor=allRisk.map(function(){return'#e05353';});
    eRiskChart.update();eRiskChart.resize();
  }
  setEl('eRiskTitle','All vehicles — advanced score ranking');
  if(eViolChart){eViolChart.data.datasets[0].data=violTotals;eViolChart.update();}
  setEl('eViolTitle','Violation breakdown — fleet total');
  if(eWdweChart){
    eWdweChart.data.labels=allWdwe.map(function(v){return v.name;});
    eWdweChart.data.datasets[0].data=allWdwe.map(function(v){return v.weekdayDist||0;});
    delete eWdweChart.data.datasets[0].barPercentage;delete eWdweChart.data.datasets[0].categoryPercentage;
    eWdweChart.data.datasets[1].data=allWdwe.map(function(v){return v.weekendDist||0;});
    delete eWdweChart.data.datasets[1].barPercentage;delete eWdweChart.data.datasets[1].categoryPercentage;
    eWdweChart.update();eWdweChart.resize();
  }
  setEl('eWdweTitle','Weekday vs weekend distance');
  if(eUtilChart){
    eUtilChart.data.labels=allUtil.map(function(v){return v.name;});
    eUtilChart.data.datasets[0].data=allUtil.map(function(v){return v.daysActive||0;});
    delete eUtilChart.data.datasets[0].barPercentage;delete eUtilChart.data.datasets[0].categoryPercentage;
    eUtilChart.data.datasets[1].data=allUtil.map(function(v){return v.daysIdle||0;});
    delete eUtilChart.data.datasets[1].barPercentage;delete eUtilChart.data.datasets[1].categoryPercentage;
    eUtilChart.update();eUtilChart.resize();
  }
  setEl('eUtilTitle','Most idle vehicles — days active vs idle');
  if(hasPrev&&ePrevChart){
    var allPrev=vehicles.filter(function(v){return prevMap[v.name];}).sort(function(a,b){return(b.score||0)-(a.score||0);});
    var prevH=Math.max(400,allPrev.length*36);
    document.getElementById('ePrevWrap').style.height=prevH+'px';
    document.getElementById('ePrevScroll').style.maxHeight='400px';
    ePrevChart.data.labels=allPrev.map(function(v){return v.name;});
    ePrevChart.data.datasets[0].data=allPrev.map(function(v){return v.score||0;});
    ePrevChart.data.datasets[0].backgroundColor=allPrev.map(function(){return'#3b6edc';});
    delete ePrevChart.data.datasets[0].barPercentage;delete ePrevChart.data.datasets[0].categoryPercentage;
    ePrevChart.data.datasets[1].data=allPrev.map(function(v){return prevMap[v.name]?prevMap[v.name].score||0:0;});
    ePrevChart.data.datasets[1].backgroundColor='#6ea8ff';
    delete ePrevChart.data.datasets[1].barPercentage;delete ePrevChart.data.datasets[1].categoryPercentage;
    ePrevChart.update();ePrevChart.resize();
    setEl('ePrevTitle','Month-on-month Advance score — current vs previous');
  }
  if(hasPrev&&eDCCompChart){
    var allNames=new Set([...vehicles.map(function(v){return v.name;}),...Object.keys(prevMap)]);
    var allDC=[...allNames].map(function(name){var curr=vehicles.find(function(v){return v.name===name;});return{name:name,currDist:curr?curr.totalDist||0:0,prevDist:prevMap[name]?prevMap[name].totalDist||0:0};}).sort(function(a,b){return b.currDist-a.currDist;});
    var dcH=Math.max(400,allDC.length*34);
    document.getElementById('eDCCompWrap').style.height=dcH+'px';
    document.getElementById('eDCCompScroll').style.maxHeight='400px';
    eDCCompChart.data.labels=allDC.map(function(v){return v.name;});
    eDCCompChart.data.datasets[0].data=allDC.map(function(v){return v.currDist;});
    eDCCompChart.data.datasets[0].backgroundColor=allDC.map(function(v){return v.currDist>=v.prevDist?'#3db87a':'#e05353';});
    delete eDCCompChart.data.datasets[0].barPercentage;delete eDCCompChart.data.datasets[0].categoryPercentage;
    eDCCompChart.data.datasets[1].data=allDC.map(function(v){return v.prevDist;});
    eDCCompChart.data.datasets[1].backgroundColor=allDC.map(function(v){return v.currDist>=v.prevDist?'#e05353':'#3db87a';});
    delete eDCCompChart.data.datasets[1].barPercentage;delete eDCCompChart.data.datasets[1].categoryPercentage;
    eDCCompChart.update();eDCCompChart.resize();
    setEl('eDCCompTitle','Distance comparison — current vs previous month');
  }
}

function eClearVehicleSelectionUI(){
  var sel=document.getElementById('eVehicleSelect');
  if(sel)sel.value='';
  var lbl=document.getElementById('eVDropLbl');
  if(lbl){lbl.textContent='Select vehicle...';lbl.style.color='var(--text3)';}
  var vSearch=document.getElementById('eVDropSearch');
  if(vSearch)vSearch.value='';
  var cSearch=document.getElementById('eCompareSearch');
  if(cSearch)cSearch.value='';
  var vPanel=document.getElementById('eVDropPanel');
  if(vPanel)vPanel.style.display='none';
  var cPanel=document.getElementById('eCompareDropPanel');
  if(cPanel)cPanel.style.display='none';
}

function eUpdateResetBtnVisibility(){
  var btn=document.getElementById('eDateResetBtn');
  if(!btn)return;
  var active=eActiveDateFilter!==null||eCompareSelected.length>0||!!eRiskSel||eFilterMode!=='fleet';
  btn.style.display=active?'':'none';
}

/* ══════════════════════════════════════════════════════════════════════════
 * EXECUTIVE SUMMARY (regenerates on fleet / vehicle / compare filters)
 * ══════════════════════════════════════════════════════════════════════════ */
function eGenerateSummary(selVehicles,label,forceRefresh){
  var bodyEl=document.getElementById('eExecBody');
  if(!bodyEl)return;
  var isFiltered=selVehicles&&selVehicles.length<vehicles.length;
  if(!forceRefresh&&D.summaryIsCustom&&!isFiltered){
    bodyEl.innerHTML=D.summaryHTML;
    return;
  }
  var vList=selVehicles||vehicles;
  var fleetScore=Math.round(vList.reduce(function(s,v){return s+(v.score||0);},0)/Math.max(1,vList.length));
  var totalD=vList.reduce(function(s,v){return s+(v.totalDist||0);},0);
  var lines=[];
  if(isFiltered){
    lines.push('<strong>'+escHTML(label||vList.map(function(v){return v.name;}).join(', '))+'</strong> — '+(vList.length===1?'individual vehicle analysis':'comparison of '+vList.length+' vehicles')+'.');
  }else{
    var band=getScoreBand(fleetScore);
    lines.push('Fleet of <strong>'+vList.length+' vehicles</strong> · Status: <span style="color:'+band.color+'">'+band.label+'</span> · Average Advanced Score: <strong>'+fleetScore+'</strong>.');
  }
  if(hasPrev&&prevVehicles.length&&!isFiltered){
    var prevAvg=prevVehicles.reduce(function(s,v){return s+(v.score||0);},0)/prevVehicles.length;
    var diff=fleetScore-Math.round(prevAvg);
    var rawPct=Math.abs(diff)/Math.max(1,prevAvg)*100;
    if(rawPct>=1){
      var dir=diff<0?'<span class="good">improved</span>':'<span class="bad">worsened</span>';
      var pctStr=rawPct>100?String(Math.abs(diff))+' points':String(Math.round(rawPct))+'%';
      lines.push('<strong>Score Trend:</strong> Fleet average score '+dir+' by <strong>'+(diff<0?'▼':'▲')+' '+pctStr+'</strong> vs last month ('+Math.round(prevAvg)+' → '+fleetScore+'). '+(diff<0?'Driving behaviour is heading in the right direction.':'Immediate fleet-wide coaching is recommended.'));
    }else{
      lines.push('<strong>Score Trend:</strong> Fleet average score remained <span class="neutral">stable</span> at <strong>'+fleetScore+'</strong>.');
    }
  }else if(isFiltered){
    vList.forEach(function(v){
      var pv=prevMap[v.name];
      var vDiff=pv?(v.score||0)-(pv.score||0):null;
      if(vDiff!==null&&Math.abs(vDiff)>0){
        var vDir=vDiff<0?'<span class="good">improved</span>':'<span class="bad">worsened</span>';
        lines.push(escHTML(v.name)+': score <strong>'+(v.score||0)+'</strong> — '+vDir+' by <strong>'+(vDiff<0?'▼':'▲')+' '+Math.abs(vDiff)+'</strong> vs last month.');
      }else{
        lines.push(escHTML(v.name)+': score <strong>'+(v.score||0)+'</strong>.');
      }
    });
  }
  var vWithActive=vList.filter(function(v){return(v.daysActive||0)>0;});
  var avgActiveDays=vWithActive.length?Math.round(vWithActive.reduce(function(s,v){return s+(v.daysActive||0);},0)/vWithActive.length):0;
  var idleVehicles=vList.filter(function(v){return(v.daysIdle||0)>=5;});
  if(!isFiltered){
    lines.push('<strong>Distance & Utilisation:</strong> Fleet covered <strong>'+Math.round(totalD).toLocaleString()+' km</strong> this period. Average active days: <strong>'+avgActiveDays+' days</strong>. '+(idleVehicles.length>0?'<span class="bad">'+idleVehicles.length+' vehicle'+(idleVehicles.length!==1?'s':'')+'</span> recorded 5+ idle weekdays — review deployment schedules.':'All vehicles maintained adequate utilisation.'));
  }else{
    lines.push('<strong>Distance:</strong> '+vList.map(function(v){return escHTML(v.name)+': <strong>'+Math.round(v.totalDist||0).toLocaleString()+' km</strong>';}).join(' · ')+'.');
  }
  var highRisk=[...vList].filter(function(v){return(v.score||0)>SCORE_BANDS.high.min-1;}).sort(function(a,b){return(b.score||0)-(a.score||0);}).slice(0,5);
  var moderate=[...vList].filter(function(v){return(v.score||0)>SCORE_BANDS.safe.max&&(v.score||0)<=SCORE_BANDS.moderate.max;}).sort(function(a,b){return(b.score||0)-(a.score||0);}).slice(0,3);
  if(highRisk.length){
    lines.push('<strong>High Risk Vehicles (Score '+SCORE_BANDS.high.min+'+):</strong> '+highRisk.map(function(v){return'<strong>'+escHTML(v.name)+'</strong> ('+v.score+')';}).join(', ')+'. '+(highRisk.length===1?'This vehicle requires':'These vehicles require')+' immediate intervention.');
  }
  if(moderate.length&&vList.length>3){
    lines.push('<strong>Needs Attention (Score '+(SCORE_BANDS.safe.max+1)+'–'+SCORE_BANDS.moderate.max+'):</strong> '+moderate.map(function(v){return'<strong>'+escHTML(v.name)+'</strong> ('+v.score+')';}).join(', ')+'. Targeted coaching recommended.');
  }
  var zeroes=[...vList].filter(function(v){return(v.score||0)===0&&(v.totalDist||0)>0;}).slice(0,3);
  var lowScorers=[...vList].filter(function(v){return(v.score||0)>0&&(v.score||0)<=SCORE_BANDS.safe.max;}).sort(function(a,b){return(a.score||0)-(b.score||0);}).slice(0,3);
  if(zeroes.length){
    lines.push('<strong>Top Performers (Score 0):</strong> '+zeroes.map(function(v){return'<strong>'+escHTML(v.name)+'</strong>';}).join(', ')+' recorded zero violations — excellent driving behaviour.');
  }else if(lowScorers.length&&vList.length>3){
    lines.push('<strong>Top Performers:</strong> '+lowScorers.map(function(v){return'<strong>'+escHTML(v.name)+'</strong> ('+v.score+')';}).join(', ')+' — safe driving, within acceptable parameters.');
  }
  var violCounts=activeViolations.map(function(v){return Object.assign({},v,{count:vList.reduce(function(s,vh){return s+(vh[v.key]||0);},0)});}).filter(function(v){return v.count>0;}).sort(function(a,b){return b.count-a.count;});
  if(violCounts.length){
    var topViol=violCounts.slice(0,3).map(function(v,idx){return(idx===0?'<strong>'+v.short+' ('+v.count.toLocaleString()+')</strong>':v.short+' ('+v.count.toLocaleString()+')');}).join(', ');
    var highViols=violCounts.filter(function(v){return v.risk==='high';});
    lines.push('<strong>Violation Summary:</strong> Top violations — '+topViol+'. '+(highViols.length?'High-severity events: <span class="bad">'+highViols.map(function(v){return v.short;}).join(', ')+'</span>. These carry the highest risk.':''));
  }
  var actions=[];
  if(highRisk.length)actions.push('Schedule driver interviews for '+highRisk.length+' high-risk vehicle'+(highRisk.length!==1?'s':''));
  if(idleVehicles.length&&!isFiltered)actions.push('Review deployment for '+idleVehicles.length+' frequently idle vehicle'+(idleVehicles.length!==1?'s':''));
  violCounts.forEach(function(v){
    if(v.key&&v.key.toLowerCase().indexOf('brake')>-1&&v.count>0)actions.push('Run defensive driving refresher focused on braking distances');
    if(v.key&&v.key.toLowerCase().indexOf('speed')>-1&&v.count>0)actions.push('Enforce speed policy — consider geofenced speed alerts');
  });
  if(actions.length){
    var uniqueActions=actions.filter(function(a,i){return actions.indexOf(a)===i;});
    lines.push('<strong>Recommended Actions:</strong> '+uniqueActions.map(function(a,idx){return(idx+1)+'. '+a;}).join(' &nbsp;·&nbsp; ')+'.');
  }
  bodyEl.innerHTML=lines.map(function(l){return'<p>'+l+'</p>';}).join('');
}

/* ══════════════════════════════════════════════════════════════════════════
 * MODE SWITCHING (fleet / vehicle / compare)
 * ══════════════════════════════════════════════════════════════════════════ */
function eSetMode(mode){
  eFilterMode=mode;
  var btnF=document.getElementById('eBtnFleet');
  var btnV=document.getElementById('eBtnVehicle');
  var btnC=document.getElementById('eBtnCompare');
  [btnF,btnV,btnC].forEach(function(b){if(b){b.style.background='transparent';b.style.color='var(--text2)';}});
  var active=mode==='fleet'?btnF:mode==='vehicle'?btnV:btnC;
  if(active){active.style.background='var(--accent)';active.style.color='#fff';}
  document.getElementById('eVDrop').style.display=mode==='vehicle'?'block':'none';
  document.getElementById('eComparePanel').style.display=mode==='compare'?'flex':'none';
  if(mode==='fleet'){
    eCompareSelected=[];
    eRiskSel='';eWdweSel='';eUtilSel='';ePrevSel='';eDistSel='';
    eClearVehicleSelectionUI();
    eRenderCompareTags();
    eRestoreChartsToFleet();
    setKpi(0,'Total vehicles',vehicles.length,'in the fleet');
    setKpi(1,'Total distance',Math.round(totalDist).toLocaleString(),'km fleet total');
    setKpi(2,'Fleet avg score',Math.round(avgScore).toLocaleString(),'lower is safer');
    setKpi(3,'Avg idle days',avgIdle.toFixed(1),'per vehicle');
    setKpi(4,'Vehicles active',vehiclesWithTrips,'recorded trips this month');
    setKpi(5,'At-risk vehicles',flaggedCount,'Moderate or High risk (score 21+)');
    var kv0=document.getElementById('eKV0');if(kv0){kv0.style.fontSize='';kv0.style.color='';}
    var kv5=document.getElementById('eKV5');if(kv5){kv5.style.fontSize='';kv5.style.color='var(--red)';}
    setKpi(6,'Violation types',activeViolations.length,'detected in data');
    var chg1=document.getElementById('eKC1');if(chg1)chg1.style.display='';
    /* Preserve eSelectionContext: switching to fleet/all and back to flagged
       should still show the previously selected vehicle(s), not reset to generic flagged list */
    eSetTableMode('flagged');renderBestPerfTable();
    eGenerateSummary();
    eUpdateResetBtnVisibility();
  }
  if(mode==='compare'){
    eRenderCompareTags();
    if(eCompareSelected.length>0)eUpdateCompare();
    eUpdateResetBtnVisibility();
  }
  if(mode==='vehicle'){
    eUpdateResetBtnVisibility();
  }
}

/* ── Vehicle dropdown ── */
function buildVehicleDropdown(){
  var sel=document.getElementById('eVehicleSelect');
  if(!sel)return;
  Object.keys(vehicleDailyData).sort().forEach(function(name){
    var opt=document.createElement('option');opt.value=name;opt.textContent=name;sel.appendChild(opt);
  });
  eBuildVDrop('');
}
function eBuildVDrop(q){
  var list=document.getElementById('eVDropList');if(!list)return;
  list.innerHTML='';
  var query=(q||'').trim().toLowerCase();
  Object.keys(vehicleDailyData).sort().forEach(function(name){
    if(query&&name.toLowerCase().indexOf(query)===-1)return;
    var d=document.createElement('div');d.textContent=name;
    d.style.cssText='padding:7px 12px;font-size:12px;cursor:pointer;color:var(--text)';
    d.onmouseenter=function(){d.style.background='var(--surface2)';};d.onmouseleave=function(){d.style.background='';};
    d.onclick=function(){
      document.getElementById('eVehicleSelect').value=name;
      var lbl=document.getElementById('eVDropLbl');if(lbl){lbl.textContent=name;lbl.style.color='var(--text)';}
      document.getElementById('eVDropPanel').style.display='none';
      document.getElementById('eVDropSearch').value='';eBuildVDrop('');
      eUpdateVehicle(name);
    };
    list.appendChild(d);
  });
  if(!list.children.length)list.innerHTML='<div style="padding:8px 12px;font-size:12px;color:var(--text3)">No vehicles found</div>';
}
function eToggleVDrop(){
  var panel=document.getElementById('eVDropPanel');var open=panel.style.display!=='none';
  panel.style.display=open?'none':'block';
  if(!open){document.getElementById('eVDropSearch').focus();eBuildVDrop('');}
}
document.addEventListener('click',function(e){
  var wrap=document.getElementById('eVDrop');if(wrap&&!wrap.contains(e.target)){var p=document.getElementById('eVDropPanel');if(p)p.style.display='none';}
  var cpanel=document.getElementById('eCompareDropPanel');var cbtn=document.getElementById('eCompareAddBtn');
  if(cpanel&&!cpanel.contains(e.target)&&cbtn&&!cbtn.contains(e.target))cpanel.style.display='none';
});

function eUpdateVehicle(name){
  if(!name)return;
  var veh=vehicles.find(function(v){return v.name===name;})||{};
  var data=vehicleDailyData[name]||[];
  var active=data.filter(function(v){return v>0;}).length;
  var idle=data.filter(function(v){return v===0;}).length;
  var maxKm=data.length?Math.max.apply(null,data):0;
  var maxDay=dayLabels[data.indexOf(maxKm)]||'—';
  var avgKm=active>0?data.reduce(function(s,v){return s+v;},0)/active:0;
  if(eDailyChart){eDailyChart.data.datasets[0].data=data;eDailyChart.data.datasets[0].label=name;eDailyChart.data.datasets[0].borderColor='#3db87a';eDailyChart.data.datasets[0].backgroundColor='rgba(61,184,122,0.08)';eDailyChart.data.datasets[0].pointBackgroundColor='#3db87a';eDailyChart.update();}
  setEl('eDCTitle','Daily distance — '+name);setEl('eDCSub','km covered by this vehicle each day');
  var band=getScoreBand(veh.score||0);
  setKpi(0,'Selected vehicle',name,'of '+vehicles.length+' total');document.getElementById('eKV0').style.fontSize='13px';document.getElementById('eKV0').style.color='var(--accent)';
  setKpi(1,'Total distance',Math.round(veh.totalDist||0).toLocaleString(),'km this month');
  setKpi(2,'Advanced score',(veh.score||0).toLocaleString(),'lower is safer');
  setKpi(3,'Idle days',idle,'days no movement');
  setKpi(4,'Active days',active,'days with trips');
  setKpi(5,'Best day',Math.round(maxKm).toLocaleString()+' km','on '+maxDay);document.getElementById('eKV5').style.color='var(--green)';
  setKpi(6,'Avg per active day',Math.round(avgKm).toLocaleString()+' km','when moving');
  var chg1=document.getElementById('eKC1');if(chg1)chg1.style.display='none';
  eRiskSel=name;
  if(eRiskChart){eRiskChart.data.labels=[name];eRiskChart.data.datasets[0].data=[veh.score||0];eRiskChart.data.datasets[0].backgroundColor=[band.color];var sH=Math.max(80,72);document.getElementById('eRiskWrap').style.height=sH+'px';document.getElementById('eRiskScroll').style.maxHeight=sH+'px';eRiskChart.update();eRiskChart.resize();}
  setEl('eRiskTitle','Advanced score — '+name);
  if(eViolChart){eViolChart.data.datasets[0].data=activeViolations.map(function(v){return veh[v.key]||0;});eViolChart.update();}
  setEl('eViolTitle','Violation breakdown — '+name);
  eWdweSel=name;
  if(eWdweChart){eWdweChart.data.labels=[name];eWdweChart.data.datasets[0].data=[veh.weekdayDist||0];eWdweChart.data.datasets[0].barPercentage=0.35;eWdweChart.data.datasets[0].categoryPercentage=0.5;eWdweChart.data.datasets[1].data=[veh.weekendDist||0];eWdweChart.data.datasets[1].barPercentage=0.35;eWdweChart.data.datasets[1].categoryPercentage=0.5;var wH=Math.max(80,100);document.getElementById('eWdweWrap').style.height=wH+'px';document.getElementById('eWdweScroll').style.maxHeight=wH+'px';eWdweChart.update();}
  setEl('eWdweTitle','Weekday vs weekend — '+name);
  eUtilSel=name;
  if(eUtilChart){eUtilChart.data.labels=[name];eUtilChart.data.datasets[0].data=[veh.daysActive||0];eUtilChart.data.datasets[0].barPercentage=0.35;eUtilChart.data.datasets[0].categoryPercentage=0.5;eUtilChart.data.datasets[1].data=[veh.daysIdle||0];eUtilChart.data.datasets[1].barPercentage=0.35;eUtilChart.data.datasets[1].categoryPercentage=0.5;var uH=Math.max(80,100);document.getElementById('eUtilWrap').style.height=uH+'px';document.getElementById('eUtilScroll').style.maxHeight=uH+'px';eUtilChart.update();}
  setEl('eUtilTitle','Active vs idle — '+name);
  ePrevSel=name;
  if(ePrevChart&&hasPrev){var prevScore=prevMap[name]?prevMap[name].score||0:0;var pH=Math.max(80,90);document.getElementById('ePrevWrap').style.height=pH+'px';document.getElementById('ePrevScroll').style.maxHeight=pH+'px';ePrevChart.data.labels=[name];ePrevChart.data.datasets[0].data=[veh.score||0];ePrevChart.data.datasets[0].backgroundColor=['#4f8ef7'];ePrevChart.data.datasets[0].barPercentage=0.35;ePrevChart.data.datasets[0].categoryPercentage=0.5;ePrevChart.data.datasets[1].data=[prevScore];ePrevChart.data.datasets[1].backgroundColor=['rgba(79,142,247,0.45)'];ePrevChart.data.datasets[1].barPercentage=0.35;ePrevChart.data.datasets[1].categoryPercentage=0.5;ePrevChart.update();setEl('ePrevTitle','Month-on-month — '+name);}
  eDistSel=name;
  if(eDCCompChart&&hasPrev){var prevDist=prevMap[name]?prevMap[name].totalDist||0:0;var currDist=veh.totalDist||0;eDCCompChart.data.labels=[name];eDCCompChart.data.datasets[0].data=[currDist];eDCCompChart.data.datasets[0].backgroundColor=[currDist>=prevDist?'#3db87a':'#e05353'];eDCCompChart.data.datasets[0].barPercentage=0.35;eDCCompChart.data.datasets[0].categoryPercentage=0.5;eDCCompChart.data.datasets[1].data=[prevDist];eDCCompChart.data.datasets[1].backgroundColor=[currDist>=prevDist?'#e05353':'#3db87a'];eDCCompChart.data.datasets[1].barPercentage=0.35;eDCCompChart.data.datasets[1].categoryPercentage=0.5;eDCCompChart.update();document.getElementById('eDCCompWrap').style.height='120px';document.getElementById('eDCCompScroll').style.maxHeight='120px';setEl('eDCCompTitle','Distance comparison — '+name);}
  eSelectionLabel='Details — '+name;renderVehicleTable([veh]);renderBestPerfTable([veh]);setEl('eFlaggedTitle',eSelectionLabel);
  eGenerateSummary([veh],name);
  eUpdateResetBtnVisibility();
}

/* ── Compare mode ── */
function eRenderCompareTags(){
  var tagsEl=document.getElementById('eCompareTags');var limitEl=document.getElementById('eCompareLimit');var addBtn=document.getElementById('eCompareAddBtn');
  if(!tagsEl)return;tagsEl.innerHTML='';
  eCompareSelected.forEach(function(name,idx){
    var col=getCompareColor(idx);var tag=document.createElement('span');
    tag.className='compare-tag';tag.style.background=col;
    tag.innerHTML=escHTML(name)+' <button class="compare-tag-x">×</button>';
    tag.querySelector('.compare-tag-x').addEventListener('click',function(e){
      e.stopPropagation();eCompareSelected=eCompareSelected.filter(function(n){return n!==name;});
      eRenderCompareTags();if(eCompareSelected.length>0)eUpdateCompare();
    });
    tagsEl.appendChild(tag);
  });
  if(limitEl)limitEl.textContent=eCompareSelected.length+'/'+COMPARE_MAX;
  if(addBtn)addBtn.disabled=eCompareSelected.length>=COMPARE_MAX;
}
function eToggleCompareDrop(){
  var panel=document.getElementById('eCompareDropPanel');var open=panel.style.display==='block';
  panel.style.display=open?'none':'block';
  if(!open){document.getElementById('eCompareSearch').value='';eBuildCompareDrop('');setTimeout(function(){document.getElementById('eCompareSearch').focus();},50);}
}
function eBuildCompareDrop(q){
  var list=document.getElementById('eCompareDropList');if(!list)return;list.innerHTML='';
  var query=(q||'').trim().toLowerCase();
  Object.keys(vehicleDailyData).sort().forEach(function(name){
    if(query&&name.toLowerCase().indexOf(query)===-1)return;
    if(eCompareSelected.indexOf(name)>-1)return;
    var d=document.createElement('div');d.textContent=name;
    d.style.cssText='padding:7px 12px;font-size:12px;cursor:pointer;color:var(--text)';
    d.onmouseenter=function(){d.style.background='var(--surface2)';};d.onmouseleave=function(){d.style.background='';};
    d.onclick=function(){
      if(eCompareSelected.length>=COMPARE_MAX){showToast('warn','Limit reached','Maximum 5 vehicles.',3000);return;}
      eCompareSelected.push(name);eRenderCompareTags();
      document.getElementById('eCompareDropPanel').style.display='none';eUpdateCompare();
    };
    list.appendChild(d);
  });
  if(!list.children.length){var e=document.createElement('div');e.textContent='No vehicles found';e.style.cssText='padding:10px 12px;font-size:12px;color:var(--text3)';list.appendChild(e);}
}
function eUpdateCompare(){
  if(!eCompareSelected.length)return;
  var datasets=eCompareSelected.map(function(name,idx){
    var col=getCompareColor(idx);return{label:name,data:vehicleDailyData[name]||[],borderColor:col,backgroundColor:'rgba(0,0,0,0)',fill:false,tension:0.3,pointRadius:3,borderWidth:2,pointHoverRadius:5,pointBackgroundColor:col};
  });
  if(eDailyChart){eDailyChart.data.labels=dayLabels;eDailyChart.data.datasets=datasets;eDailyChart.update();}
  var titleNames=eCompareSelected.length<=3?eCompareSelected.join(', '):eCompareSelected.slice(0,2).join(', ')+' + '+(eCompareSelected.length-2)+' more';
  setEl('eDCTitle','Comparing — '+titleNames);setEl('eDCSub',eCompareSelected.length+' vehicles');
  var selVeh=vehicles.filter(function(v){return eCompareSelected.indexOf(v.name)>-1;});
  var selDist=selVeh.reduce(function(s,v){return s+(v.totalDist||0);},0);
  var selAvg=selVeh.length?selVeh.reduce(function(s,v){return s+(v.score||0);},0)/selVeh.length:0;
  setKpi(0,'Comparing',eCompareSelected.length+' vehicles','selected for comparison');document.getElementById('eKV0').style.fontSize='16px';document.getElementById('eKV0').style.color='var(--accent)';
  setKpi(1,'Combined distance',Math.round(selDist).toLocaleString(),'km this month');
  setKpi(2,'Avg score',Math.round(selAvg).toLocaleString(),'lower is safer');
  var chg1=document.getElementById('eKC1');if(chg1)chg1.style.display='none';
  var cmpRisk=eCompareSelected.map(function(nm,idx){var v=vehicles.find(function(v){return v.name===nm;})||{name:nm,score:0};return{name:nm,score:v.score||0,color:getCompareColor(idx)};});
  if(eRiskChart){eRiskChart.data.labels=cmpRisk.map(function(v){return v.name;});eRiskChart.data.datasets[0].data=cmpRisk.map(function(v){return v.score;});eRiskChart.data.datasets[0].backgroundColor=cmpRisk.map(function(v){return v.color;});var cH=Math.max(120,cmpRisk.length*32+40);document.getElementById('eRiskWrap').style.height=cH+'px';document.getElementById('eRiskScroll').style.maxHeight=cH+'px';eRiskChart.update();eRiskChart.resize();}
  if(eViolChart){eViolChart.data.datasets[0].data=activeViolations.map(function(v){return eCompareSelected.reduce(function(s,nm){var vh=vehicles.find(function(x){return x.name===nm;})||{};return s+(vh[v.key]||0);},0);});eViolChart.update();}
  setEl('eViolTitle','Violations — '+eCompareSelected.length+' vehicles (combined)');
  eSelectionLabel='Comparison — '+eCompareSelected.length+' vehicles';renderVehicleTable(selVeh);renderBestPerfTable(selVeh);setEl('eFlaggedTitle',eSelectionLabel);
  eGenerateSummary(selVeh,'Comparison');
  eUpdateResetBtnVisibility();
}

/* ══════════════════════════════════════════════════════════════════════════
 * DATE FILTER
 * ══════════════════════════════════════════════════════════════════════════ */
function setDateDefaults(){
  var pad2=function(n){return String(n).padStart(2,'0');};
  var cm=D.reportMonth||'';
  if(cm){var parts=cm.split('-').map(Number);var yr=parts[0];var mo=parts[1];var days=new Date(yr,mo,0).getDate();document.getElementById('eDateFrom').value=yr+'-'+pad2(mo)+'-01';document.getElementById('eDateTo').value=yr+'-'+pad2(mo)+'-'+pad2(days);}
  else if(D.hasRealDates&&dayInfo.length){var first=dayInfo[0].date;var last=dayInfo[dayInfo.length-1].date;if(first)document.getElementById('eDateFrom').value=first.getFullYear()+'-'+pad2(first.getMonth()+1)+'-'+pad2(first.getDate());if(last)document.getElementById('eDateTo').value=last.getFullYear()+'-'+pad2(last.getMonth()+1)+'-'+pad2(last.getDate());}
}
function eApplyDateRange(){
  var fromVal=document.getElementById('eDateFrom').value;var toVal=document.getElementById('eDateTo').value;
  if(!fromVal||!toVal)return;
  var hasReal=D.hasRealDates;var filtered;
  if(hasReal){var fp=fromVal.split('-').map(Number);var tp=toVal.split('-').map(Number);var from=new Date(fp[0],fp[1]-1,fp[2]);var to=new Date(tp[0],tp[1]-1,tp[2]);if(from>to){showToast('warn','Invalid range','Start date must be before end date.',4000);return;}filtered=dayInfo.filter(function(d){if(!d.date)return false;var nd=new Date(d.date.getFullYear(),d.date.getMonth(),d.date.getDate());return nd>=from&&nd<=to;});}
  else{var fd=parseInt(fromVal.split('-')[2])||1;var td=parseInt(toVal.split('-')[2])||31;if(fd>td){showToast('warn','Invalid range','Start day must be before end day.',4000);return;}filtered=dayInfo.filter(function(d){return d.dayNum&&d.dayNum>=fd&&d.dayNum<=td;});}
  if(!filtered.length){showToast('warn','No data','Select a range that exists in this report.',5000);eUpdateResetBtnVisibility();return;}
  eActiveDateFilter=filtered;
  var filteredLabels=filtered.map(function(d){return d.label;});
  var filteredTotals=filtered.map(function(d){return utilRows.reduce(function(s,r){return s+Number(r[d.key]||0);},0);});
  if(eDailyChart){eDailyChart.data.labels=filteredLabels;eDailyChart.data.datasets[0].data=filteredTotals;eDailyChart.data.datasets[0].borderColor='#4f8ef7';eDailyChart.data.datasets[0].backgroundColor='rgba(79,142,247,0.08)';eDailyChart.update();}
  setEl('eDCTitle','Daily distance — '+filteredLabels[0]+' to '+filteredLabels[filteredLabels.length-1]);
  setEl('eDCSub',filtered.length+' days selected · '+Math.round(filteredTotals.reduce(function(s,v){return s+v;},0)).toLocaleString()+' km');
  eUpdateResetBtnVisibility();
  showToast('success','Date range applied','Showing '+filtered.length+' day'+(filtered.length!==1?'s':'')+' from '+filteredLabels[0]+' to '+filteredLabels[filteredLabels.length-1],3000);
}
function eResetAllFilters(){
  eActiveDateFilter=null;
  eCompareSelected=[];
  eRiskSel='';eWdweSel='';eUtilSel='';ePrevSel='';eDistSel='';
  setDateDefaults();
  eClearVehicleSelectionUI();
  eClearTableSelection(); /* Explicit reset clears selection — unlike fleet-mode switch which preserves it */
  eSetTableMode('flagged');
  eSetMode('fleet');
  eUpdateResetBtnVisibility();
  showToast('info','Filters reset','All filters cleared — showing full fleet report.',2500);
}

/* ══════════════════════════════════════════════════════════════════════════
 * TABLES
 * ══════════════════════════════════════════════════════════════════════════ */
/* eTableMode        — 'flagged' or 'all' (the toggle buttons)
   eTableOverride    — vehicles set by a selection (single vehicle / compare)
   eSelectionContext — persists the selection so toggling All/Flagged and back restores it */
var eTableMode='flagged'; var eTableOverride=null; var eSelectionContext=null; var eSelectionLabel='';
function eSetTableMode(mode){
  eTableMode=mode;
  /* Do NOT clear eSelectionContext — toggling All/Flagged should not lose the selection */
  var btnF=document.getElementById('eBtnFlaggedOnly');var btnA=document.getElementById('eBtnAllVehicles');
  if(btnF){btnF.style.background=mode==='flagged'?'var(--accent)':'transparent';btnF.style.color=mode==='flagged'?'#fff':'var(--text2)';btnF.style.borderColor=mode==='flagged'?'var(--accent)':'var(--border2)';}
  if(btnA){btnA.style.background=mode==='all'?'var(--accent)':'transparent';btnA.style.color=mode==='all'?'#fff':'var(--text2)';btnA.style.borderColor=mode==='all'?'var(--accent)':'var(--border2)';}
  eRenderVehicleTableBody();
}
function renderVehicleTable(overrideVehicles){
  if(overrideVehicles){
    eTableOverride=overrideVehicles;
    eSelectionContext=overrideVehicles; /* remember selection so All/Flagged toggle can restore it */
  }
  eRenderVehicleTableBody();
}
function eClearTableSelection(){
  /* Called only when filter mode truly resets (back to whole fleet) */
  eTableOverride=null;
  eSelectionContext=null;
  eSelectionLabel='';
}
function eRenderVehicleTableBody(){
  var tbody=document.getElementById('eAlertBody');if(!tbody)return;tbody.innerHTML='';
  var tableVehicles;
  if(eTableMode==='all'){
    /* Show all vehicles but still highlight that a selection is active */
    tableVehicles=[...vehicles].sort(function(a,b){return(b.score||0)-(a.score||0);});
    setEl('eFlaggedTitle','All vehicles ('+vehicles.length+')');
  }else if(eSelectionContext){
    /* Flagged mode with an active selection — show the selection, not generic flagged list */
    tableVehicles=eSelectionContext;
    setEl('eFlaggedTitle',eSelectionLabel||'Selected vehicles');
  }else{
    tableVehicles=flaggedVehicles;
    setEl('eFlaggedTitle','Vehicles Scoring');
  }
  tableVehicles.forEach(function(v){
    var score=v.score||0;var band=getScoreBand(score);
    var badgeCls=score>40?'b-red':score>20?'b-amber':'b-green';
    var prev=prevMap[v.name];var changeCell='';var distCell='';
    if(hasPrev){
      var ps=prev?prev.score:null;var psc=ps!==null?ps.toLocaleString():'—';
      var chg='<span style="color:var(--text3)">—</span>';
      if(ps!==null&&ps>0){var d=score-ps;var rawP=Math.abs(d)/ps*100;var pct=Math.round(rawP);if(rawP>100){var sign=d>0?'+':'-';chg=d>0?'<span class="arr-up">▲ '+sign+Math.round(Math.abs(d))+' worse</span>':'<span class="arr-down">▼ '+sign+Math.round(Math.abs(d))+' better</span>';}else if(rawP>1)chg=d>0?'<span class="arr-up">▲ '+pct+'% worse</span>':'<span class="arr-down">▼ '+pct+'% better</span>';else chg='<span style="color:var(--text3)">→</span>';}
      changeCell='<td>'+psc+'</td><td>'+chg+'</td>';
      if(prev){var pd=prev.totalDist||0;var cd=v.totalDist||0;if(pd>0){var dd=cd-pd;var rawD=Math.abs(dd)/pd*100;var dpct=Math.round(rawD);var dc;if(rawD>100)dc=dd>0?'<span style="color:var(--green)">▲ +'+Math.round(Math.abs(dd)).toLocaleString()+' km</span>':'<span style="color:var(--red)">▼ -'+Math.round(Math.abs(dd)).toLocaleString()+' km</span>';else if(rawD>1)dc=dd>0?'<span style="color:var(--green)">▲ '+dpct+'%</span>':'<span style="color:var(--red)">▼ '+dpct+'%</span>';else dc='<span style="color:var(--text3)">→</span>';distCell='<td>'+Math.round(pd).toLocaleString()+'</td><td>'+dc+'</td>';}}
      if(!distCell)distCell='<td>—</td><td><span style="color:var(--text3)">—</span></td>';
    }
    var utilCells='<td>'+(v.daysActive!==undefined?v.daysActive:'—')+'</td><td>'+(v.daysIdle!==undefined?v.daysIdle:'—')+'</td><td>'+(v.weekdayDist!==undefined?Math.round(v.weekdayDist).toLocaleString():'—')+'</td><td>'+(v.weekendDist!==undefined?Math.round(v.weekendDist).toLocaleString():'—')+'</td>';
    var tr=document.createElement('tr');
    var extraKeys=[];vehicles.forEach(function(vh){Object.keys(vh._extra||{}).forEach(function(k){if(extraKeys.indexOf(k)===-1)extraKeys.push(k);});});
    var extraCells=extraKeys.map(function(k){return'<td>'+escHTML(String(v._extra&&v._extra[k]!==undefined?v._extra[k]:'—'))+'</td>';}).join('');
    tr.innerHTML='<td>'+escHTML(v.name)+'</td><td>'+Math.round(v.totalDist||0).toLocaleString()+'</td>'+distCell+'<td>'+score.toLocaleString()+'</td>'+changeCell+utilCells+activeViolations.map(function(vd){return'<td>'+(v[vd.key]||0).toLocaleString()+'</td>';}).join('')+extraCells+'<td><span class="badge '+badgeCls+'">'+band.label+'</span></td>';
    tbody.appendChild(tr);
  });
}
function renderBestPerfTable(overrideVehicles){
  var tbody=document.getElementById('eBestPerfBody');if(!tbody)return;tbody.innerHTML='';
  var src=overrideVehicles&&overrideVehicles.length?overrideVehicles:vehicles;
  var perfs=[...src].filter(function(v){return(v.totalDist||0)>0;}).map(function(v){
    var totalViol=VIOLATIONS.reduce(function(s,vd){return s+(v[vd.key]||0);},0);
    var per100=totalViol/(v.totalDist/100);
    var rating=per100===0?'Excellent':per100<1?'Good':per100<3?'Fair':'Poor';
    return Object.assign({},v,{totalViol:totalViol,per100:per100,rating:rating});
  }).sort(function(a,b){return a.per100-b.per100;}).slice(0,15);
  if(!perfs.length){tbody.innerHTML='<tr><td colspan="7" style="color:var(--text3)">No ranked vehicles for this selection.</td></tr>';return;}
  perfs.forEach(function(v,idx){
    var rCls=v.rating==='Excellent'||v.rating==='Good'?'b-green':v.rating==='Fair'?'b-amber':'b-red';
    var tr=document.createElement('tr');
    tr.innerHTML='<td>'+(idx+1)+'</td><td>'+escHTML(v.name)+'</td><td>'+Math.round(v.totalDist).toLocaleString()+'</td><td>'+v.totalViol.toLocaleString()+'</td><td>'+v.per100.toFixed(2)+'</td><td>'+(v.score||0).toLocaleString()+'</td><td><span class="badge '+rCls+'">'+v.rating+'</span></td>';
    tbody.appendChild(tr);
  });
}

/* ── Violation guide ── */
function buildViolationGuide(){
  var guide=document.getElementById('eViolGuide');if(!guide)return;
  VIOLATIONS.forEach(function(v){
    var riskCls=v.risk==='high'?'viol-high':v.risk==='med'?'viol-med':'viol-low';
    var badgeCls=v.risk==='high'?'b-red':v.risk==='med'?'b-amber':'b-green';
    var count=vehicles.reduce(function(s,vh){return s+(vh[v.key]||0);},0);
    guide.innerHTML+='<div class="viol-card '+riskCls+'"><span class="badge '+badgeCls+'" style="margin-bottom:6px;display:inline-block">'+(v.risk==='high'?'High risk':v.risk==='med'?'Medium risk':'Low risk')+'</span><div class="viol-name">'+escHTML(v.short)+'</div><div class="viol-desc">'+escHTML(v.desc)+'</div>'+(count>0?'<div style="margin-top:8px;font-size:11px;color:var(--accent)">Fleet total: '+count.toLocaleString()+' events</div>':'')+'</div>';
  });
}

/* ── Initialise ── */
function _safeInit() {
  if (typeof Chart === 'undefined') {
    var main = document.getElementById('main');
    if (main) {
      main.innerHTML =
        '<div style="margin:2rem;padding:1.5rem;background:rgba(224,83,83,.1);border:1px solid rgba(224,83,83,.3);border-radius:10px;color:#e05353;font-size:13px;line-height:1.6">' +
        '<strong>Chart library failed to load.</strong><br>' +
        'This report requires an internet connection to load Chart.js from CDN. ' +
        'Please open this file while connected to the internet, or contact your report sender.' +
        '</div>';
    }
    return;
  }
  try {
    buildDashboard();
  } catch(err) {
    var main = document.getElementById('main');
    if (main) {
      main.innerHTML =
        '<div style="margin:2rem;padding:1.5rem;background:rgba(224,83,83,.1);border:1px solid rgba(224,83,83,.3);border-radius:10px;color:#e05353;font-size:13px;line-height:1.6">' +
        '<strong>Report failed to render.</strong><br>' +
        '<span style="font-family:monospace;font-size:11px;color:#f87171">' + String(err) + '</span>' +
        '</div>';
    }
  }
}
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', _safeInit);
} else {
  _safeInit();
}
<\/script>
</body>
</html>`;
}
