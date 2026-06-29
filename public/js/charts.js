/**
 * charts.js
 * ─────────────────────────────────────────────────────────────────────────────
 * Main dashboard rendering engine.
 * Called once per client tab after Excel files are parsed.
 *
 * Contents:
 *   1. renderClient(i)          — orchestrator: coordinates all sections below
 *   2. buildDashboardHTML()     — generates the static HTML shell
 *   3. buildCharts()            — creates all 7 Chart.js instances
 *   4. buildFilterBar()         — wires up fleet/vehicle/compare toggle
 *   5. buildCompareMode()       — compare mode state, tags, dropdown
 *   6. buildVehicleMode()       — single-vehicle dropdown and chart updates
 *   7. buildDateFilter()        — date range apply/reset logic
 *   8. buildTables()            — flagged vehicles + best performers tables
 *   9. buildExecSummary()       — auto-generated executive summary
 *  10. buildViolationGuide()    — violation reference cards
 *
 * ── Architecture notes ───────────────────────────────────────────────────────
 *
 *  SCOPE PATTERN
 *  All per-client state (chart instances, selected vehicle, compare list,
 *  date filter) lives inside renderClient()'s closure. This means multiple
 *  client tabs are fully isolated — one tab's state can never affect another's.
 *
 *  EVENT DELEGATION vs window[] GLOBALS
 *  Previously, functions like switchDistView, toggleDrop, applyDateRange were
 *  registered as window['fnName_0'], window['fnName_1'] etc. via HTML onclick
 *  attributes, leaving stale entries in the global scope on every render.
 *
 *  They are now attached directly to DOM elements using addEventListener()
 *  after the HTML is injected. The window[] registrations are gone.
 *
 *  VARIABLE HOISTING FOR CHART CALLBACKS
 *  Chart.js tick/tooltip callbacks execute asynchronously. Any variable they
 *  reference must be declared BEFORE the `new Chart(...)` call — not inside
 *  a conditional block above it. All such variables are declared at the top
 *  of buildCharts() for this reason.
 *
 *  SCORE BANDS
 *  All risk classification uses getScoreBand() and SCORE_BANDS from
 *  constants.js. The threshold is never hardcoded in this file.
 * ─────────────────────────────────────────────────────────────────────────────
 */


/* ═══════════════════════════════════════════════════════════════════════════
 * 1. renderClient(i)  —  main orchestrator
 * ═══════════════════════════════════════════════════════════════════════════ */

/**
 * renderClient(i)
 * Entry point called by switchClient() when the user selects a client tab.
 * Parses the client's sheet data, builds all DOM, then wires up all
 * interactive behaviour.
 *
 * @param {number} i — index into the global `clients` array
 */
function renderClient(i) {
  const client     = clients[i];
  const sheets     = client.sheets;
  const sheetNames = Object.keys(sheets);

  /* ── Sheet selection ── */
  const scoringSheet = sheetNames.find(n => /scor/i.test(n) && !/prev/i.test(n)) || sheetNames[0];
  const utilSheet    = sheetNames.find(n => /util/i.test(n))
                    || sheetNames.find(n => n !== scoringSheet)
                    || sheetNames[0];
  const prevSheet    = sheetNames.find(n => /prev/i.test(n) || /last/i.test(n) || /prior/i.test(n));

  const scoringRows = sheets[scoringSheet] || [];
  const utilRows    = sheets[utilSheet]    || [];
  const prevRows    = prevSheet ? sheets[prevSheet] : [];

  /* ── Column detection ── */
  const sCols = detectCols(scoringRows);
  const uCols = detectCols(utilRows);
  const pCols = prevRows.length ? detectCols(prevRows) : {};

  /* ── Vehicle data ── */
  const vehicleMap = buildVehicleMap(scoringRows, sCols);

  /* ── Day columns — detect and parse from util sheet ── */
  const uKeys  = Object.keys(utilRows[0] || {});
  const dayCols = detectDayCols(uKeys);
  const dayInfo = dayCols.map(k => ({ key: k, ...parseDayInfo(k) }));

  /* ── Enrich vehicle map with utilisation data ── */
  enrichFromUtilSheet(vehicleMap, utilRows, uCols, dayInfo);

  /* ── Daily fleet totals (line chart) ── */
  const dailyTotals = buildDailyTotals(utilRows, dayInfo);

  /* ── Per-vehicle daily data (vehicle mode + compare mode) ── */
  const vehicleDailyData = buildVehicleDailyData(utilRows, uCols, dayInfo);

  /* ── Previous month map ── */
  let prevMap   = {};
  let normIndex = {};
  if (prevRows.length) {
    const result = buildPrevMap(prevRows, pCols);
    prevMap      = result.prevMap;
    normIndex    = result.normIndex;
    reconcilePrevMap(vehicleMap, prevMap, normIndex);
  }
  const hasPrev    = Object.keys(prevMap).length > 0;
  const prevVehicles = Object.values(prevMap);

  /* ── Vehicle list ── */
  const vehicles = Object.values(vehicleMap).filter(v => v.name && v.name !== 'Unknown');

  /* ── Render nothing if no vehicles found ── */
  if (vehicles.length === 0) {
    document.getElementById('dashContent').innerHTML =
      `<div style="background:#1e1215;border:1px solid rgba(224,83,83,.3);
                   border-radius:14px;padding:1.5rem;margin:1rem">
         <div style="color:var(--red);font-weight:600;font-size:14px;margin-bottom:6px">
           No vehicles found
         </div>
         <div style="color:var(--text2);font-size:13px;line-height:1.6">
           Could not extract any vehicle records from this file.
           Check that your Excel has a column named "Grouping", "Row Labels", or "Vehicle".
         </div>
         <div style="color:var(--text3);font-size:11px;margin-top:10px;font-family:monospace;
                     background:rgba(0,0,0,.3);padding:8px;border-radius:6px">
           Sheets found: ${sheetNames.join(', ')}<br>
           Scoring sheet used: ${scoringSheet}
         </div>
       </div>`;
    return;
  }

  /* ── Fleet KPIs ── */
  const totalDist        = vehicles.reduce((s, v) => s + (v.totalDist || 0), 0);
  const avgScore         = vehicles.length
                         ? vehicles.reduce((s, v) => s + (v.score || 0), 0) / vehicles.length
                         : 0;
  const vehiclesWithIdle = vehicles.filter(v => v.daysIdle !== undefined);
  const avgIdle          = vehiclesWithIdle.length
                         ? vehiclesWithIdle.reduce((s, v) => s + (v.daysIdle || 0), 0) / vehiclesWithIdle.length
                         : 0;
  const vehiclesWithTrips = vehicles.filter(v => (v.daysActive || 0) > 0).length;

  /*
   * flaggedCount — uses SCORE_BANDS.safe.max (20) from constants.js.
   * Label updated to match actual band names: "Moderate or High risk (score 21+)"
   */
  const flaggedCount     = vehicles.filter(v => isAtRisk(v.score)).length;
  const activeViolations = VIOLATIONS.filter(v => vehicles.some(vh => (vh[v.key] || 0) > 0));

  /* Previous month KPIs */
  const prevAvgScore  = prevVehicles.length
                      ? prevVehicles.reduce((s, v) => s + (v.score || 0), 0) / prevVehicles.length
                      : null;
  const prevTotalDist = prevVehicles.reduce((s, v) => s + (v.totalDist || 0), 0);

  /* ── Derived chart datasets ── */
  const top15dist        = [...vehicles].sort((a, b) => (b.totalDist || 0) - (a.totalDist || 0)).slice(0, 15);
  const violTotals       = activeViolations.map(v => vehicles.reduce((s, vh) => s + (vh[v.key] || 0), 0));
  const flaggedVehicles  = vehicles.filter(v => isAtRisk(v.score))
                                   .sort((a, b) => (b.score || 0) - (a.score || 0));
  const dayLabels        = dayInfo.map(d => d.label);

  /* ── Build HTML shell ── */
  buildDashboardHTML(i, client, vehicles, hasPrev, activeViolations,
    dailyTotals, totalDist, avgScore, avgIdle, vehiclesWithTrips,
    flaggedCount, prevTotalDist, prevAvgScore);

  /* ── Data quality warnings ── */
  const allWarns = (client.warnings || []).concat(validateVehicleData(vehicles));
  if (allWarns.length) {
    const banner = document.createElement('div');
    banner.className = 'warn-banner';
    banner.innerHTML =
      `<span style="font-size:16px;flex-shrink:0">&#9888;</span>` +
      `<div><b>Data warnings for ${escapeHTML(client.name)}</b>` +
      `<ul>${allWarns.map(w => `<li>${escapeHTML(w)}</li>`).join('')}</ul></div>`;
    const dc = document.getElementById('dashContent');
    dc.insertBefore(banner, dc.firstChild);
  }

  /* ── Chart.js global defaults ── */
  Chart.defaults.color = '#8b90a7';

  /* ── Per-client mutable state ─────────────────────────────────────────────
   * These are declared HERE, before any chart is created, so that
   * Chart.js tick callbacks can safely close over them.
   * (Placing them inside conditionals above the chart constructors would
   * cause "variable not defined" errors inside tick/tooltip callbacks.) */
  let compareSelected  = [];   /* array of vehicle names in compare mode  */
  let riskSelected     = '';   /* highlighted vehicle name in risk chart   */
  let wdweSelected     = '';   /* highlighted vehicle in wdwe chart        */
  let utilSelected     = '';   /* highlighted vehicle in util chart        */
  let prevSelected     = '';   /* highlighted vehicle in prev chart        */
  let distCompSel      = '';   /* highlighted vehicle in dist comp chart   */
  let activeDateFilter = null; /* filtered dayInfo subset, or null         */

  const COMPARE_MAX    = 5;

  /** Returns the colour for vehicle at compare index idx */
  function getCompareColor(idx) {
    return COMPARE_COLORS_MAIN[idx % COMPARE_COLORS_MAIN.length];
  }

  /* ── Build charts ── */
  const charts = buildCharts(i, vehicles, vehicleDailyData, dayLabels,
    dailyTotals, dayInfo, activeViolations, violTotals,
    hasPrev, prevMap, prevVehicles,
    /* tick callback state — passed by reference via closure */
    () => compareSelected, () => riskSelected, () => wdweSelected,
    () => utilSelected,    () => prevSelected,  () => distCompSel,
    getCompareColor);

  const {
    dailyChartInst, riskChartInst, violChartInst,
    wdweChartInst,  utilChartInst, prevChartInst,
    distCompInst,   allRiskVehicles, allWdweVehicles,
    allUtilVehicles, riskBarHeight,  riskFullH,
  } = charts;

  /* ── Wire up date filter ── */
  const dateApi = buildDateFilter(
    i, dayInfo, dayLabels, dailyTotals, vehicleDailyData,
    vehicles, utilRows, client,
    avgScore, avgIdle, vehiclesWithTrips, flaggedCount, activeViolations,
    () => activeDateFilter,
    f  => { activeDateFilter = f; },
    () => compareSelected,
    dailyChartInst,
    renderVehicleTable, renderBestPerfTable,
    updateCompareCharts
  );

  /* ── Wire up filter bar (fleet / vehicle / compare toggle) ── */
  buildFilterBar(
    i, vehicles, vehicleDailyData, dayLabels, dailyTotals,
    hasPrev, prevMap, prevVehicles,
    activeViolations, violTotals, flaggedVehicles,
    avgScore, avgIdle, vehiclesWithTrips, flaggedCount,
    totalDist, prevTotalDist, prevAvgScore,
    dailyChartInst, riskChartInst, violChartInst,
    wdweChartInst,  utilChartInst, prevChartInst, distCompInst,
    allRiskVehicles, allWdweVehicles, allUtilVehicles,
    riskBarHeight, riskFullH, hasPrev,
    () => compareSelected, s => { compareSelected = s; },
    () => activeDateFilter,
    renderVehicleTable, renderBestPerfTable, generateSummary,
    updateCompareCharts, updateVehicleChart,
    COMPARE_MAX, getCompareColor
  );

  /* ── Wire up vehicle mode dropdown ── */
  buildVehicleMode(
    i, vehicles, vehicleDailyData, dayLabels,
    hasPrev, prevMap, activeViolations,
    dailyChartInst, riskChartInst, violChartInst,
    wdweChartInst,  utilChartInst, prevChartInst, distCompInst,
    riskBarHeight,
    s => { riskSelected = s; },
    s => { wdweSelected = s; },
    s => { utilSelected = s; },
    s => { prevSelected = s; },
    s => { distCompSel  = s; },
    () => activeDateFilter,
    renderVehicleTable, renderBestPerfTable, generateSummary
  );

  /* ── Build tables ── */
  function renderVehicleTable(overrideVehicles, forceAll) {
    if (overrideVehicles) vehicleTableContext = overrideVehicles;
    _renderVehicleTable(
      i, vehicles, flaggedVehicles, hasPrev, prevMap,
      activeViolations,
      overrideVehicles !== undefined ? overrideVehicles : vehicleTableContext,
      forceAll
    );
  }

  function renderBestPerfTable(overrideVehicles) {
    _renderBestPerfTable(i, vehicles, overrideVehicles);
  }

  /* ── Vehicle table mode toggle ── */
  let vehicleTableMode    = 'flagged';
  let vehicleTableContext = null;

  document.getElementById(`btnFlaggedOnly_${i}`)
    ?.addEventListener('click', () => setVehicleTableMode('flagged'));
  document.getElementById(`btnAllVehicles_${i}`)
    ?.addEventListener('click', () => setVehicleTableMode('all'));

  function setVehicleTableMode(mode) {
    vehicleTableMode    = mode;
    vehicleTableContext = null;
    const btnF = document.getElementById(`btnFlaggedOnly_${i}`);
    const btnA = document.getElementById(`btnAllVehicles_${i}`);
    if (btnF) {
      btnF.style.background  = mode === 'flagged' ? 'var(--accent)' : 'transparent';
      btnF.style.color       = mode === 'flagged' ? '#fff' : 'var(--text2)';
      btnF.style.borderColor = mode === 'flagged' ? 'var(--accent)' : 'var(--border2)';
    }
    if (btnA) {
      btnA.style.background  = mode === 'all' ? 'var(--accent)' : 'transparent';
      btnA.style.color       = mode === 'all' ? '#fff' : 'var(--text2)';
      btnA.style.borderColor = mode === 'all' ? 'var(--accent)' : 'var(--border2)';
    }
    renderVehicleTable(null, mode === 'all');
  }

  /* ── Compare mode chart updates ── */
  function updateCompareCharts() {
    _updateCompareCharts(
      i, vehicles, vehicleDailyData, dayLabels, dailyTotals,
      hasPrev, prevMap, activeViolations,
      () => compareSelected,
      () => activeDateFilter,
      dailyChartInst, riskChartInst, violChartInst,
      wdweChartInst,  utilChartInst, prevChartInst, distCompInst,
      riskBarHeight, getCompareColor,
      renderVehicleTable, renderBestPerfTable, generateSummary,
      avgScore, avgIdle, flaggedCount, activeViolations
    );
  }

  /* ── Single vehicle chart update ── */
  function updateVehicleChart() {
    _updateVehicleChart(
      i, vehicles, vehicleDailyData, dayLabels,
      hasPrev, prevMap, activeViolations,
      dailyChartInst, riskChartInst, violChartInst,
      wdweChartInst,  utilChartInst, prevChartInst, distCompInst,
      riskBarHeight,
      s => { riskSelected = s; },
      s => { wdweSelected = s; },
      s => { utilSelected = s; },
      s => { prevSelected = s; },
      s => { distCompSel  = s; },
      () => activeDateFilter,
      renderVehicleTable, renderBestPerfTable, generateSummary
    );
  }

  /* ── Executive summary ── */
  function generateSummary(selVehicles, label, forceRefresh) {
    _generateSummary(i, client, vehicles, prevMap, prevVehicles,
      activeViolations, hasPrev, avgIdle,
      selVehicles, label, forceRefresh);
  }

  /* ── Wire up summary edit buttons ── */
  buildExecSummaryControls(i, client, generateSummary);

  /* ── Build violation guide ── */
  buildViolationGuide(i, vehicles, prevVehicles, hasPrev, activeViolations);

  /* ── Auto-apply date filter if a month was chosen on upload ── */
  if (client.month) {
    dateApi.setInitialFiltered(true);
    dateApi.apply(true /* silent */);
  }

  /* ── Initial renders ── */
  generateSummary();
  renderVehicleTable();
  renderBestPerfTable();
}


/* ═══════════════════════════════════════════════════════════════════════════
 * 2. buildDashboardHTML()
 * Generates the static HTML shell for the dashboard.
 * All interactive elements are wired up with addEventListener() after
 * injection — no inline onclick attributes.
 * ═══════════════════════════════════════════════════════════════════════════ */

function buildDashboardHTML(
  i, client, vehicles, hasPrev, activeViolations,
  dailyTotals, totalDist, avgScore, avgIdle, vehiclesWithTrips,
  flaggedCount, prevTotalDist, prevAvgScore
) {
  const hasDailyData = Object.keys(dailyTotals).length > 0;

  const content = document.getElementById('dashContent');
  content.innerHTML = `
    <!-- Report header -->
    <div class="report-header">
      <div class="report-header-left">
        <h2>${escapeHTML(client.name)} — Fleet Performance Report</h2>
        <p>${vehicles.length} vehicles &nbsp;·&nbsp;
           ${new Date().toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric' })}</p>
      </div>
      <div class="report-header-right">
        <button class="btn-download btn-export" id="exportBtn_${i}">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4"/>
            <polyline points="17,8 12,3 7,8"/>
            <line x1="12" y1="3" x2="12" y2="15"/>
          </svg>
          Export &amp; Send
        </button>
      </div>
    </div>

    <!-- Global filter bar -->
    <div class="global-filter-bar" id="globalFilterBar_${i}">
      <span class="global-filter-label">Filter</span>
      <div class="global-filter-toggle">
        <button id="btnFleet_${i}"   style="background:var(--accent);color:#fff">Whole fleet</button>
        <button id="btnVehicle_${i}" style="background:transparent;color:var(--text2)">Single vehicle</button>
        <button id="btnCompare_${i}" style="background:transparent;color:var(--text2)">Compare</button>
      </div>

      <!-- Compare panel (hidden until compare mode activated) -->
      <div id="comparePanel_${i}" style="display:none;align-items:center;gap:8px;flex-wrap:wrap;flex:1">
        <div class="compare-tags" id="compareTags_${i}"></div>
        <div style="position:relative">
          <button class="compare-add" id="compareAddBtn_${i}">+ Add vehicle</button>
          <span  class="compare-limit" id="compareLimit_${i}">0/5</span>
          <div id="compareDropPanel_${i}" style="
            display:none;position:absolute;top:calc(100% + 4px);left:0;
            min-width:220px;background:var(--surface);border:1px solid var(--border2);
            border-radius:var(--radius);z-index:9999;box-shadow:0 8px 24px rgba(0,0,0,.5)">
            <div style="padding:6px">
              <input id="compareSearch_${i}" type="text" placeholder="Search plate..."
                autocomplete="off" style="
                  width:100%;background:var(--surface2);color:var(--text);
                  border:1px solid var(--border2);border-radius:6px;
                  padding:6px 8px;font-size:12px;outline:none"/>
            </div>
            <div id="compareDropList_${i}" style="max-height:220px;overflow-y:auto;padding-bottom:4px"></div>
          </div>
        </div>
      </div>

      <!-- Single-vehicle dropdown (hidden until vehicle mode activated) -->
      <div id="vDrop_${i}" style="display:none;position:relative;width:220px">
        <div id="vDropTrigger_${i}" style="
          display:flex;align-items:center;justify-content:space-between;
          background:var(--surface2);border:1px solid var(--border2);
          border-radius:var(--radius);padding:6px 10px;font-size:12px;
          cursor:pointer;user-select:none">
          <span id="vDropLbl_${i}" style="color:var(--text3)">Select vehicle...</span>
          <span style="color:var(--text3);font-size:10px;margin-left:6px">&#9662;</span>
        </div>
        <div id="vDropPanel_${i}" style="
          display:none;position:absolute;top:calc(100% + 4px);left:0;right:0;
          background:var(--surface);border:1px solid var(--border2);
          border-radius:var(--radius);z-index:9999;box-shadow:0 8px 24px rgba(0,0,0,.5)">
          <div style="padding:6px">
            <input id="vDropSearch_${i}" type="text" placeholder="Search plate..."
              autocomplete="off" style="
                width:100%;background:var(--surface2);color:var(--text);
                border:1px solid var(--border2);border-radius:6px;
                padding:6px 8px;font-size:12px;outline:none"/>
          </div>
          <div id="vDropList_${i}" style="max-height:220px;overflow-y:auto;padding-bottom:4px"></div>
        </div>
        <select id="vehicleSelect_${i}" style="display:none"></select>
      </div>

      <div style="width:1px;background:var(--border2);align-self:stretch;margin:0 4px"></div>

      <!-- Date range picker -->
      <div class="date-range-wrap" id="dateRangeBar_${i}">
        <span class="date-range-label">From</span>
        <input type="date" class="date-input" id="dateFrom_${i}"/>
        <span class="date-range-label">To</span>
        <input type="date" class="date-input" id="dateTo_${i}"/>
        <button class="date-range-btn"   id="dateApplyBtn_${i}">Apply</button>
        <button class="date-range-reset" id="dateResetBtn_${i}" style="display:none">Reset</button>
      </div>
    </div>

    <!-- Executive summary -->
    <div class="exec-summary" id="execSummary_${i}">
      <div style="display:flex;align-items:center;justify-content:space-between;
                  gap:10px;flex-wrap:wrap;margin-bottom:10px">
        <div class="exec-summary-title" style="margin-bottom:0">Executive Summary</div>
        <div style="display:flex;gap:6px">
          <button class="exec-edit-btn"            id="execEditBtn_${i}">Edit</button>
          <button class="exec-edit-btn"            id="execSaveBtn_${i}"   style="display:none">Save</button>
          <button class="exec-edit-btn exec-edit-btn-muted" id="execCancelBtn_${i}" style="display:none">Cancel</button>
          <button class="exec-edit-btn exec-edit-btn-muted" id="execAutoBtn_${i}"
                  title="Regenerate summary from current data">Auto</button>
        </div>
      </div>
      <div class="exec-summary-body" id="execBody_${i}">Generating summary...</div>
      <div class="exec-notes-wrap">
        <div class="exec-notes-label">Additional notes</div>
        <textarea class="exec-notes" id="execNotes_${i}"
          placeholder="Add your own observations, context or recommendations here before exporting..."></textarea>
      </div>
    </div>

    <!-- KPI cards -->
    <div class="kpi-grid" id="kpiGrid_${i}">
      <div class="kpi">
        <div class="kpi-label" id="kpiLbl0_${i}">Total vehicles</div>
        <div class="kpi-value" id="kpiVal0_${i}">${vehicles.length}</div>
        <div class="kpi-sub"   id="kpiSub0_${i}">in the fleet</div>
      </div>
      <div class="kpi">
        <div class="kpi-label" id="kpiLbl1_${i}">Total distance</div>
        <div class="kpi-value" id="kpiVal1_${i}">${Math.round(totalDist).toLocaleString()}</div>
        <div class="kpi-sub"   id="kpiSub1_${i}">km fleet total</div>
        <span id="kpiChange1_${i}">${hasPrev ? changeTag(totalDist, prevTotalDist, false) : ''}</span>
      </div>
      <div class="kpi">
        <div class="kpi-label" id="kpiLbl2_${i}">Fleet avg score</div>
        <div class="kpi-value" id="kpiVal2_${i}">${Math.round(avgScore).toLocaleString()}</div>
        <div class="kpi-sub"   id="kpiSub2_${i}">lower is safer</div>
        ${hasPrev && prevAvgScore !== null ? changeTag(avgScore, prevAvgScore, true) : ''}
      </div>
      <div class="kpi">
        <div class="kpi-label" id="kpiLbl3_${i}">Avg idle days</div>
        <div class="kpi-value" id="kpiVal3_${i}">${avgIdle.toFixed(1)}</div>
        <div class="kpi-sub"   id="kpiSub3_${i}">per vehicle (weekdays only)</div>
      </div>
      <div class="kpi">
        <div class="kpi-label" id="kpiLbl4_${i}">Vehicles active</div>
        <div class="kpi-value" id="kpiVal4_${i}">${vehiclesWithTrips}</div>
        <div class="kpi-sub"   id="kpiSub4_${i}">recorded trips this month</div>
      </div>
      <div class="kpi">
        <div class="kpi-label" id="kpiLbl5_${i}">At-risk vehicles</div>
        <div class="kpi-value" id="kpiVal5_${i}" style="color:var(--red)">${flaggedCount}</div>
        <div class="kpi-sub"   id="kpiSub5_${i}">Moderate or High risk (score ${SCORE_BANDS.safe.max + 1}+)</div>
      </div>
      <div class="kpi">
        <div class="kpi-label" id="kpiLbl6_${i}">Violation types</div>
        <div class="kpi-value" id="kpiVal6_${i}">${activeViolations.length}</div>
        <div class="kpi-sub"   id="kpiSub6_${i}">detected in data</div>
      </div>
    </div>

    <!-- Daily distance chart (only shown when daily data is available) -->
    ${hasDailyData ? `
    <div class="card" style="margin-bottom:1rem">
      <div style="display:flex;align-items:center;justify-content:space-between;
                  flex-wrap:wrap;gap:10px;margin-bottom:1rem">
        <div>
          <div class="card-title" style="margin-bottom:2px" id="distChartTitle_${i}">
            Daily fleet distance (km)
          </div>
          <div style="font-size:11px;color:var(--text3)" id="distChartSub_${i}">
            Total km covered by all vehicles each day
          </div>
        </div>
      </div>
      <div class="chart-wrap" style="height:180px"><canvas id="dailyChart_${i}"></canvas></div>
      <div id="vehicleStatGrid_${i}"
           style="display:none;grid-template-columns:repeat(auto-fit,minmax(130px,1fr));
                  gap:8px;margin-top:12px"></div>
    </div>` : ''}

    <!-- Score ranking + violation donut -->
    <div class="grid-2" style="margin-bottom:1rem">
      <div class="card">
        <div class="card-title" id="riskTitle_${i}">All vehicles — advanced score ranking</div>
        <div id="riskScroll_${i}" style="overflow-y:auto;max-height:320px">
          <div id="riskWrap_${i}" style="position:relative;height:320px">
            <canvas id="riskChart_${i}"></canvas>
          </div>
        </div>
      </div>
      <div class="card">
        <div class="card-title" id="violTitle_${i}">Violation breakdown — fleet total</div>
        <div class="legend-row" id="violLeg_${i}"></div>
        <div class="chart-wrap" style="height:280px"><canvas id="violChart_${i}"></canvas></div>
      </div>
    </div>

    <!-- Weekday/weekend + utilisation -->
    <div class="grid-2" style="margin-bottom:1rem">
      <div class="card">
        <div class="card-title" id="wdweTitle_${i}">Weekday vs weekend distance</div>
        <div class="legend-row">
          <span class="leg"><span class="leg-dot" style="background:#4f8ef7"></span>Weekday</span>
          <span class="leg"><span class="leg-dot" style="background:#2ec4b6"></span>Weekend</span>
        </div>
        <div id="wdweScroll_${i}" style="overflow-y:auto;max-height:320px">
          <div id="wdweWrap_${i}" style="position:relative;height:320px">
            <canvas id="wdweChart_${i}"></canvas>
          </div>
        </div>
      </div>
      <div class="card">
        <div class="card-title" id="utilTitle_${i}">Most idle vehicles — days active vs idle</div>
        <div class="legend-row">
          <span class="leg"><span class="leg-dot" style="background:#4f8ef7"></span>Active</span>
          <span class="leg"><span class="leg-dot" style="background:#555b72"></span>Idle</span>
        </div>
        <div id="utilScroll_${i}" style="overflow-y:auto;max-height:320px">
          <div id="utilWrap_${i}" style="position:relative;height:320px">
            <canvas id="utilChart_${i}"></canvas>
          </div>
        </div>
      </div>
    </div>

    <!-- Month-on-month score chart (only shown when prev data available) -->
    ${hasPrev ? `
    <div class="card" style="margin-bottom:1rem">
      <div class="card-title" id="prevTitle_${i}">
        Month-on-month Advance score — current vs previous
      </div>
      <div class="legend-row">
        <span class="leg"><span class="leg-dot" style="background:#3b6edc"></span>Current month</span>
        <span class="leg"><span class="leg-dot" style="background:#6ea8ff"></span>Previous month</span>
      </div>
      <div id="prevScroll_${i}" style="overflow-y:auto;max-height:400px">
        <div id="prevWrap_${i}" style="position:relative;height:400px">
          <canvas id="prevChart_${i}"></canvas>
        </div>
      </div>
    </div>` : ''}

    <!-- Flagged vehicles table -->
    <div class="card print-section-flagged" style="margin-bottom:1.5rem">
      <div style="display:flex;align-items:center;justify-content:space-between;
                  margin-bottom:10px;flex-wrap:wrap;gap:8px">
        <div class="card-title" id="flaggedTitle_${i}" style="margin-bottom:0">Vehicles Scoring</div>
        <div style="display:flex;gap:6px">
          <button id="btnFlaggedOnly_${i}" style="
            padding:4px 10px;font-size:11px;border-radius:20px;
            border:1px solid var(--accent);background:var(--accent);
            color:#fff;cursor:pointer;font-weight:500">Flagged only</button>
          <button id="btnAllVehicles_${i}" style="
            padding:4px 10px;font-size:11px;border-radius:20px;
            border:1px solid var(--border2);background:transparent;
            color:var(--text2);cursor:pointer">All vehicles</button>
        </div>
      </div>
      <div class="tbl-wrap" style="max-height:420px;overflow-x:auto;overflow-y:auto;width:100%">
        <table style="min-width:max(900px, calc(500px + ${activeViolations.length * 90}px));white-space:nowrap">
          <thead><tr>
            <th>Vehicle</th><th>Total km</th>
            ${hasPrev ? '<th>Prev km</th><th>Dist Change</th>' : ''}
            <th>Advanced Score</th>
            ${hasPrev ? '<th>Prev Score</th><th>Score Change</th>' : ''}
            <th>Days Active</th><th>Days Idle</th>
            <th>Weekday km</th><th>Weekend km</th>
            ${activeViolations.map(v => `<th>${v.short}</th>`).join('')}
            <th>Risk</th>
          </tr></thead>
          <tbody id="alertBody_${i}"></tbody>
        </table>
      </div>
    </div>

    <!-- Top ranked by violations per 100km -->
    <div class="card print-section-ranked" style="margin-bottom:1rem">
      <div class="card-title" id="bestPerfTitle_${i}">Top Ranked Vehicles Per 100 KM</div>
      <div style="font-size:12px;color:var(--text3);margin-bottom:12px">
        Ranked by violations per 100 km — a normalised safety metric that accounts
        for distance driven. Lower is safer.
      </div>
      <div class="tbl-wrap">
        <table>
          <thead><tr>
            <th>Rank</th><th>Vehicle</th><th>Total km</th>
            <th>Violations</th><th>Per 100 km</th>
            <th>Advance Score</th><th>Rating</th>
          </tr></thead>
          <tbody id="bestPerfBody_${i}"></tbody>
        </table>
      </div>
    </div>

    <!-- Violation reference guide -->
    <div class="section-title" style="margin-bottom:12px">Violation reference guide</div>
    <div class="viol-grid" id="violGuide_${i}"></div>

    <!-- Score grading guide -->
    <div class="card" style="margin-bottom:1rem">
      <div class="card-title">Advanced Score Grading Guide</div>
      <div style="font-size:12px;color:var(--text3);margin-bottom:14px">
        The advanced score is calculated from all violation types weighted by severity.
        Use this guide to interpret fleet and individual vehicle scores.
      </div>
      <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:12px">
        <div style="background:rgba(61,184,122,.08);border:1px solid rgba(61,184,122,.25);
                    border-radius:8px;padding:14px">
          <div style="font-size:18px;font-weight:700;color:var(--green);margin-bottom:4px">
            1 – ${SCORE_BANDS.safe.max}
          </div>
          <div style="font-size:13px;font-weight:600;color:var(--text);margin-bottom:6px">
            ✓ Safe Driving
          </div>
          <div style="font-size:11px;color:var(--text3);line-height:1.5">
            Excellent driving behaviour. Minimal violations detected.
            Vehicle is operating within acceptable safety parameters.
          </div>
        </div>
        <div style="background:rgba(255,183,77,.08);border:1px solid rgba(255,183,77,.25);
                    border-radius:8px;padding:14px">
          <div style="font-size:18px;font-weight:700;color:#f7c04f;margin-bottom:4px">
            ${SCORE_BANDS.moderate.min} – ${SCORE_BANDS.moderate.max}
          </div>
          <div style="font-size:13px;font-weight:600;color:var(--text);margin-bottom:6px">
            ⚠ Needs Attention
          </div>
          <div style="font-size:11px;color:var(--text3);line-height:1.5">
            Moderate risk level. Some recurring violations observed.
            Driver coaching and monitoring recommended.
          </div>
        </div>
        <div style="background:rgba(224,83,83,.08);border:1px solid rgba(224,83,83,.25);
                    border-radius:8px;padding:14px">
          <div style="font-size:18px;font-weight:700;color:var(--red);margin-bottom:4px">
            ${SCORE_BANDS.high.min}+
          </div>
          <div style="font-size:13px;font-weight:600;color:var(--text);margin-bottom:6px">
            ✗ High Risk Vehicle
          </div>
          <div style="font-size:11px;color:var(--text3);line-height:1.5">
            Serious violation patterns detected. Immediate intervention required.
            Consider vehicle grounding pending review.
          </div>
        </div>
      </div>
    </div>
  `;

  /* Wire up export button now that HTML is in the DOM */
  document.getElementById(`exportBtn_${i}`)
    ?.addEventListener('click', () => exportReport(client.name));
}


/* ═══════════════════════════════════════════════════════════════════════════
 * 3. buildCharts()
 * Creates all Chart.js instances for a client.
 * Returns an object containing every chart instance and derived data arrays.
 * ═══════════════════════════════════════════════════════════════════════════ */

function buildCharts(
  i, vehicles, vehicleDailyData, dayLabels, dailyTotals, dayInfo,
  activeViolations, violTotals, hasPrev, prevMap, prevVehicles,
  getCompareSelected, getRiskSel, getWdweSel,
  getUtilSel, getPrevSel, getDistCompSel,
  getCompareColor
) {
  /*
   * HOISTING NOTE: All variables referenced by Chart.js tick/tooltip callbacks
   * must be declared here — before any `new Chart(...)` call — regardless of
   * whether they are used. Declaring them inside conditionals would make them
   * inaccessible when the callback fires asynchronously.
   */
  const allRiskVehicles  = [...vehicles].sort((a, b) => (b.score || 0) - (a.score || 0));
  const allWdweVehicles  = [...vehicles].sort((a, b) => (b.totalDist || 0) - (a.totalDist || 0));
  const allUtilVehicles  = [...vehicles]
    .filter(v => v.daysIdle !== undefined)
    .sort((a, b) => (b.daysIdle || 0) - (a.daysIdle || 0));

  const riskBarHeight = 32;
  const riskFullH     = Math.max(320, allRiskVehicles.length * riskBarHeight);

  /* ── Tick colour helper ── uses getters to read latest closure state */
  function tickColor(label, defaultColor, selGetter) {
    if (label === selGetter())       return '#3db87a';
    const ci = getCompareSelected().indexOf(label);
    if (ci > -1)                     return getCompareColor(ci);
    return defaultColor;
  }

  function tickFont(label, selGetter) {
    if (label === selGetter())                       return { size: 10, weight: 'bold' };
    if (getCompareSelected().indexOf(label) > -1)   return { size: 10, weight: 'bold' };
    return { size: 10 };
  }

  /* ── Daily line chart ── */
  let dailyChartInst = null;
  if (Object.keys(dailyTotals).length > 0) {
    dailyChartInst = new Chart(document.getElementById(`dailyChart_${i}`), {
      type: 'line',
      data: {
        labels  : dayLabels,
        datasets: [{
          label           : 'Fleet total',
          data            : Object.values(dailyTotals),
          borderColor     : '#4f8ef7',
          backgroundColor : 'rgba(79,142,247,0.08)',
          fill            : true,
          tension         : 0.3,
          pointRadius     : 3,
          borderWidth     : 2,
          pointHoverRadius    : 5,
          pointBackgroundColor: '#4f8ef7',
        }],
      },
      options: {
        responsive         : true,
        maintainAspectRatio: false,
        plugins: {
          legend : { display: false },
          tooltip: { callbacks: { label: ctx => `${Math.round(ctx.raw).toLocaleString()} km` } },
        },
        scales: {
          x: {
            ticks : { font: { size: 10 }, maxRotation: 45, color: '#555b72' },
            grid  : { color: 'rgba(255,255,255,0.04)' },
            border: { display: false },
          },
          y: {
            ticks : { callback: v => v.toLocaleString() + ' km', color: '#555b72' },
            grid  : { color: 'rgba(255,255,255,0.04)' },
            border: { display: false },
          },
        },
      },
    });
  }

  /* ── Risk (score ranking) bar chart ── */
  document.getElementById(`riskWrap_${i}`).style.height = `${riskFullH}px`;

  const riskChartInst = new Chart(document.getElementById(`riskChart_${i}`), {
    type: 'bar',
    data: {
      labels  : allRiskVehicles.map(v => v.name),
      datasets: [{
        data            : allRiskVehicles.map(v => v.score || 0),
        backgroundColor : allRiskVehicles.map(() => SCORE_BANDS.high.color),
        borderRadius    : 3,
      }],
    },
    options: {
      indexAxis          : 'y',
      responsive         : true,
      maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: {
        x: {
          ticks : { callback: v => v.toLocaleString(), color: '#555b72' },
          grid  : { color: 'rgba(255,255,255,0.04)' },
          border: { display: false },
        },
        y: {
          ticks: {
            font : ctx => tickFont(ctx.tick.label, getRiskSel),
            color: ctx => tickColor(ctx.tick.label, '#8b90a7', getRiskSel),
          },
        },
      },
    },
  });

  /* ── Violation donut ── */
  document.getElementById(`violLeg_${i}`).innerHTML = activeViolations
    .map((v, idx) =>
      `<span class="leg">
         <span class="leg-dot" style="background:${VIOL_COLORS[idx % VIOL_COLORS.length]}"></span>
         ${escapeHTML(v.short)}
       </span>`
    ).join('');

  const violChartInst = new Chart(document.getElementById(`violChart_${i}`), {
    type: 'doughnut',
    data: {
      labels  : activeViolations.map(v => v.short),
      datasets: [{
        data           : violTotals,
        backgroundColor: VIOL_COLORS.slice(0, activeViolations.length),
        borderWidth    : 0,
      }],
    },
    options: {
      responsive         : true,
      maintainAspectRatio: false,
      plugins            : { legend: { display: false } },
    },
  });

  /* ── Weekday/weekend stacked bar ── */
  const wdweFullH = Math.max(320, allWdweVehicles.length * 32);
  document.getElementById(`wdweWrap_${i}`).style.height = `${wdweFullH}px`;

  const wdweChartInst = new Chart(document.getElementById(`wdweChart_${i}`), {
    type: 'bar',
    data: {
      labels  : allWdweVehicles.map(v => v.name),
      datasets: [
        {
          label          : 'Weekday',
          data           : allWdweVehicles.map(v => v.weekdayDist || 0),
          backgroundColor: '#4f8ef7',
          borderRadius   : 2,
        },
        {
          label          : 'Weekend',
          data           : allWdweVehicles.map(v => v.weekendDist || 0),
          backgroundColor: '#2ec4b6',
          borderRadius   : 2,
        },
      ],
    },
    options: {
      indexAxis          : 'y',
      responsive         : true,
      maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: {
        x: {
          stacked: true,
          ticks  : { callback: v => v.toLocaleString(), color: '#555b72' },
          grid   : { color: 'rgba(255,255,255,0.04)' },
          border : { display: false },
        },
        y: {
          stacked: true,
          ticks  : {
            font : ctx => tickFont(ctx.tick.label, getWdweSel),
            color: ctx => tickColor(ctx.tick.label, '#8b90a7', getWdweSel),
          },
        },
      },
    },
  });

  /* ── Utilisation stacked bar ── */
  const utilFullH = Math.max(320, allUtilVehicles.length * 32);
  document.getElementById(`utilWrap_${i}`).style.height = `${utilFullH}px`;

  const utilChartInst = new Chart(document.getElementById(`utilChart_${i}`), {
    type: 'bar',
    data: {
      labels  : allUtilVehicles.map(v => v.name),
      datasets: [
        {
          label          : 'Active',
          data           : allUtilVehicles.map(v => v.daysActive || 0),
          backgroundColor: '#4f8ef7',
          borderRadius   : 2,
        },
        {
          label          : 'Idle',
          data           : allUtilVehicles.map(v => v.daysIdle || 0),
          backgroundColor: '#555b72',
          borderRadius   : 2,
        },
      ],
    },
    options: {
      indexAxis          : 'y',
      responsive         : true,
      maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: {
        x: {
          stacked: true,
          ticks  : { color: '#555b72' },
          grid   : { color: 'rgba(255,255,255,0.04)' },
          border : { display: false },
        },
        y: {
          stacked: true,
          ticks  : {
            font : ctx => tickFont(ctx.tick.label, getUtilSel),
            color: ctx => tickColor(ctx.tick.label, '#8b90a7', getUtilSel),
          },
        },
      },
    },
  });

  /* ── Month-on-month comparison bar ── */
  let prevChartInst = null;
  if (hasPrev) {
    const allPrevVehicles = vehicles
      .filter(v => prevMap[v.name])
      .sort((a, b) => (b.score || 0) - (a.score || 0));
    const prevFullH = Math.max(400, allPrevVehicles.length * 36);
    document.getElementById(`prevWrap_${i}`).style.height = `${prevFullH}px`;

    prevChartInst = new Chart(document.getElementById(`prevChart_${i}`), {
      type: 'bar',
      data: {
        labels  : allPrevVehicles.map(v => v.name),
        datasets: [
          {
            label          : 'Current',
            data           : allPrevVehicles.map(v => v.score || 0),
            backgroundColor: allPrevVehicles.map(() => '#3b6edc'),
            borderRadius   : 3,
          },
          {
            label          : 'Previous',
            data           : allPrevVehicles.map(v => prevMap[v.name]?.score || 0),
            backgroundColor: '#6ea8ff',
            borderRadius   : 3,
          },
        ],
      },
      options: {
        indexAxis          : 'y',
        responsive         : true,
        maintainAspectRatio: false,
        plugins: { legend: { display: false } },
        scales: {
          x: {
            ticks : { callback: v => v.toLocaleString(), color: '#555b72' },
            grid  : { color: 'rgba(255,255,255,0.04)' },
            border: { display: false },
          },
          y: {
            ticks: {
              font : ctx => tickFont(ctx.tick.label, getPrevSel),
              color: ctx => tickColor(ctx.tick.label, '#8b90a7', getPrevSel),
            },
          },
        },
      },
    });
  }

  /* ── Distance comparison bar (injected dynamically) ── */
  let distCompInst = null;
  if (hasPrev) {
    /* Inject the distance comparison card just before the flagged vehicles card */
    const dcCard = document.createElement('div');
    dcCard.innerHTML = `
      <div class="card print-section-distcomp" style="margin-bottom:1rem">
        <div class="card-title" id="distCompTitle_${i}">
          Distance comparison — current vs previous month
        </div>
        <div class="legend-row">
          <span class="leg"><span class="leg-dot" style="background:#3db87a"></span>Increased</span>
          <span class="leg"><span class="leg-dot" style="background:#e05353"></span>Decreased</span>
          <span class="leg"><span class="leg-dot" style="background:rgba(79,142,247,0.45)"></span>Previous month</span>
        </div>
        <div id="distCompScroll_${i}" style="overflow-y:auto;max-height:400px">
          <div id="distCompWrap_${i}" style="position:relative;height:400px">
            <canvas id="distCompChart_${i}"></canvas>
          </div>
        </div>
      </div>`;

    const flaggedCard = document.getElementById(`flaggedTitle_${i}`)?.closest('.card');
    const dashContent = document.getElementById('dashContent');
    if (flaggedCard && dashContent) dashContent.insertBefore(dcCard, flaggedCard);
    else if (dashContent) dashContent.appendChild(dcCard);

    const allNames    = new Set([...vehicles.map(v => v.name), ...Object.keys(prevMap)]);
    const allDistComp = [...allNames].map(name => {
      const curr = vehicles.find(v => v.name === name);
      return {
        name,
        currDist: curr ? curr.totalDist || 0 : 0,
        prevDist: prevMap[name] ? prevMap[name].totalDist || 0 : 0,
      };
    }).sort((a, b) => b.currDist - a.currDist);

    const dcWrap = document.getElementById(`distCompWrap_${i}`);
    if (dcWrap) dcWrap.style.height = `${Math.max(400, allDistComp.length * 34)}px`;

    const dcCanvas = document.getElementById(`distCompChart_${i}`);
    if (dcCanvas) {
      /* Destroy any stale instance before creating a new one */
      const existing = Chart.getChart(dcCanvas);
      if (existing) existing.destroy();

      distCompInst = new Chart(dcCanvas, {
        type: 'bar',
        data: {
          labels  : allDistComp.map(v => v.name),
          datasets: [
            {
              label          : 'Current',
              data           : allDistComp.map(v => v.currDist),
              backgroundColor: allDistComp.map(v => v.currDist >= v.prevDist ? '#3db87a' : '#e05353'),
              borderRadius   : 3,
            },
            {
              label          : 'Previous',
              data           : allDistComp.map(v => v.prevDist),
              backgroundColor: allDistComp.map(v => v.currDist >= v.prevDist ? '#e05353' : '#3db87a'),
              borderRadius   : 3,
            },
          ],
        },
        options: {
          indexAxis          : 'y',
          responsive         : true,
          maintainAspectRatio: false,
          plugins: {
            legend : { display: false },
            tooltip: {
              callbacks: {
                label: ctx => `${ctx.dataset.label}: ${Math.round(ctx.raw).toLocaleString()} km`,
              },
            },
          },
          scales: {
            x: {
              ticks : { callback: v => v.toLocaleString() + ' km', color: '#555b72' },
              grid  : { color: 'rgba(255,255,255,0.04)' },
              border: { display: false },
            },
            y: {
              ticks: {
                font : ctx => tickFont(ctx.tick.label, getDistCompSel),
                color: ctx => tickColor(ctx.tick.label, '#8b90a7', getDistCompSel),
              },
              border: { display: false },
            },
          },
        },
      });
    }
  }

  return {
    dailyChartInst, riskChartInst, violChartInst,
    wdweChartInst,  utilChartInst, prevChartInst, distCompInst,
    allRiskVehicles, allWdweVehicles, allUtilVehicles,
    riskBarHeight, riskFullH,
  };
}


/* ═══════════════════════════════════════════════════════════════════════════
 * 4. changeTag()
 * Returns an HTML <span> showing month-on-month change for KPI cards.
 * ═══════════════════════════════════════════════════════════════════════════ */

/**
 * changeTag(curr, prev, lowerIsBetter)
 * Generates a coloured percentage or absolute change indicator.
 *
 * When the change exceeds 100%, shows an absolute difference instead of
 * a percentage (e.g. "▲ +20 vs last month" rather than "▲ 900%").
 *
 * @param {number}  curr           — current month value
 * @param {number}  prev           — previous month value
 * @param {boolean} lowerIsBetter  — true for scores; false for distance
 * @returns {string} HTML string
 */
function changeTag(curr, prev, lowerIsBetter = true) {
  if (prev === null || prev === undefined || prev === 0) return '';

  const diff   = curr - prev;
  const rawPct = Math.abs(diff) / Math.abs(prev) * 100;

  if (rawPct < 1) return `<span class="kpi-change kpi-same">→ no change</span>`;

  const improved = lowerIsBetter ? diff < 0 : diff > 0;
  const cls      = improved ? 'kpi-down' : 'kpi-up';
  const arrow    = diff > 0 ? '▲' : '▼';

  /* Beyond 100% change: show absolute difference to avoid misleading "900%" */
  if (rawPct > 100) {
    const absDiff = Math.round(Math.abs(diff));
    const sign    = diff > 0 ? '+' : '-';
    return `<span class="kpi-change ${cls}">${arrow} ${sign}${absDiff} vs last month</span>`;
  }

  const pct = Math.round(rawPct);
  return `<span class="kpi-change ${cls}">${arrow} ${pct}% vs last month</span>`;
}


/* ═══════════════════════════════════════════════════════════════════════════
 * 5. buildFilterBar()
 * Wires up the fleet / vehicle / compare toggle buttons.
 * Restores full-fleet state when switching back to fleet mode.
 * ═══════════════════════════════════════════════════════════════════════════ */

function buildFilterBar(
  i, vehicles, vehicleDailyData, dayLabels, dailyTotals,
  hasPrev, prevMap, prevVehicles,
  activeViolations, violTotals, flaggedVehicles,
  avgScore, avgIdle, vehiclesWithTrips, flaggedCount,
  totalDist, prevTotalDist, prevAvgScore,
  dailyChartInst, riskChartInst, violChartInst,
  wdweChartInst,  utilChartInst, prevChartInst, distCompInst,
  allRiskVehicles, allWdweVehicles, allUtilVehicles,
  riskBarHeight, riskFullH, _hasPrev,
  getCompareSelected, setCompareSelected,
  getActiveDateFilter,
  renderVehicleTable, renderBestPerfTable, generateSummary,
  updateCompareCharts, updateVehicleChart,
  COMPARE_MAX, getCompareColor
) {
  const btnFleet   = document.getElementById(`btnFleet_${i}`);
  const btnVehicle = document.getElementById(`btnVehicle_${i}`);
  const btnCompare = document.getElementById(`btnCompare_${i}`);

  function setActiveBtn(mode) {
    [btnFleet, btnVehicle, btnCompare].forEach(b => {
      if (!b) return;
      b.style.background = 'transparent';
      b.style.color      = 'var(--text2)';
    });
    const active = mode === 'fleet' ? btnFleet : mode === 'vehicle' ? btnVehicle : btnCompare;
    if (active) { active.style.background = 'var(--accent)'; active.style.color = '#fff'; }
  }

  /* ── Fleet mode ── */
  btnFleet?.addEventListener('click', () => {
    setActiveBtn('fleet');
    setCompareSelected([]);

    document.getElementById(`vDrop_${i}`).style.display        = 'none';
    document.getElementById(`comparePanel_${i}`).style.display = 'none';

    /* Restore daily chart */
    if (dailyChartInst) {
      dailyChartInst.data.labels      = dayLabels;
      dailyChartInst.data.datasets    = [{
        label           : 'Fleet total',
        data            : Object.values(dailyTotals),
        borderColor     : '#4f8ef7',
        backgroundColor : 'rgba(79,142,247,0.08)',
        fill: true, tension: 0.3, pointRadius: 3, borderWidth: 2,
        pointHoverRadius: 5, pointBackgroundColor: '#4f8ef7',
      }];
      dailyChartInst.update();
    }
    setEl(`distChartTitle_${i}`, 'Daily fleet distance (km)');
    setEl(`distChartSub_${i}`,   'Total km covered by all vehicles each day');

    /* Restore KPI cards */
    restoreFleetKpis(i, vehicles, totalDist, avgScore, avgIdle,
      vehiclesWithTrips, flaggedCount, activeViolations,
      hasPrev, prevTotalDist, prevAvgScore);

    /* Restore all charts to full fleet */
    restoreAllChartsToFleet(
      i, vehicles, hasPrev, prevMap, activeViolations, violTotals,
      riskChartInst, violChartInst, wdweChartInst, utilChartInst,
      prevChartInst, distCompInst,
      allRiskVehicles, allWdweVehicles, allUtilVehicles, riskFullH
    );

    vehicleTableContext = null;
    generateSummary();
    renderVehicleTable(null, vehicleTableMode === 'all');
    renderBestPerfTable();
  });

  /* ── Vehicle mode ── */
  btnVehicle?.addEventListener('click', () => {
    setActiveBtn('vehicle');
    document.getElementById(`vDrop_${i}`).style.display        = 'block';
    document.getElementById(`comparePanel_${i}`).style.display = 'none';
    updateVehicleChart();
  });

  /* ── Compare mode ── */
  btnCompare?.addEventListener('click', () => {
    setActiveBtn('compare');
    document.getElementById(`vDrop_${i}`).style.display        = 'none';
    document.getElementById(`comparePanel_${i}`).style.display = 'flex';

    /* Reset to fleet chart until vehicles are added */
    if (dailyChartInst) {
      dailyChartInst.data.labels   = dayLabels;
      dailyChartInst.data.datasets = [{
        label: 'Fleet total', data: Object.values(dailyTotals),
        borderColor: '#4f8ef7', backgroundColor: 'rgba(79,142,247,0.08)',
        fill: true, tension: 0.3, pointRadius: 3, borderWidth: 2,
        pointHoverRadius: 5, pointBackgroundColor: '#4f8ef7',
      }];
      dailyChartInst.update();
    }
    setEl(`distChartTitle_${i}`, 'Compare vehicles — select up to 5');
    setEl(`distChartSub_${i}`,   'Add vehicles using the + button above');

    buildCompareMode(
      i, vehicles, vehicleDailyData, dayLabels, dailyTotals,
      hasPrev, prevMap, activeViolations,
      getCompareSelected, setCompareSelected,
      getActiveDateFilter,
      dailyChartInst, riskChartInst, violChartInst,
      wdweChartInst,  utilChartInst, prevChartInst, distCompInst,
      riskBarHeight, COMPARE_MAX, getCompareColor,
      renderVehicleTable, renderBestPerfTable, generateSummary,
      updateCompareCharts
    );

    if (getCompareSelected().length > 0) updateCompareCharts();
  });
}


/* ═══════════════════════════════════════════════════════════════════════════
 * 6. buildCompareMode()
 * Sets up the compare mode tag list and vehicle search dropdown.
 * ═══════════════════════════════════════════════════════════════════════════ */

function buildCompareMode(
  i, vehicles, vehicleDailyData, dayLabels, dailyTotals,
  hasPrev, prevMap, activeViolations,
  getCompareSelected, setCompareSelected,
  getActiveDateFilter,
  dailyChartInst, riskChartInst, violChartInst,
  wdweChartInst,  utilChartInst, prevChartInst, distCompInst,
  riskBarHeight, COMPARE_MAX, getCompareColor,
  renderVehicleTable, renderBestPerfTable, generateSummary,
  updateCompareCharts
) {
  /** Re-renders the compare tags and updates the Add button state */
  function renderCompareTags() {
    const tagsEl  = document.getElementById(`compareTags_${i}`);
    const limitEl = document.getElementById(`compareLimit_${i}`);
    const addBtn  = document.getElementById(`compareAddBtn_${i}`);
    if (!tagsEl) return;

    tagsEl.innerHTML = '';
    getCompareSelected().forEach((name, idx) => {
      const col = getCompareColor(idx);
      const tag = document.createElement('span');
      tag.className        = 'compare-tag';
      tag.style.background = col;
      tag.innerHTML = `${escapeHTML(name)} <button class="compare-tag-x" title="Remove">×</button>`;
      tag.querySelector('.compare-tag-x').addEventListener('click', e => {
        e.stopPropagation();
        setCompareSelected(getCompareSelected().filter(n => n !== name));
        renderCompareTags();
        if (getCompareSelected().length > 0) {
          updateCompareCharts();
        } else {
          /* Last vehicle removed — revert daily chart to fleet total */
          if (dailyChartInst) {
            dailyChartInst.data.labels   = dayLabels;
            dailyChartInst.data.datasets = [{
              label: 'Fleet total', data: Object.values(dailyTotals),
              borderColor: '#4f8ef7', backgroundColor: 'rgba(79,142,247,0.08)',
              fill: true, tension: 0.3, pointRadius: 3, borderWidth: 2,
              pointHoverRadius: 5, pointBackgroundColor: '#4f8ef7',
            }];
            dailyChartInst.update();
          }
          setEl(`distChartTitle_${i}`, 'Daily fleet distance (km)');
          setEl(`distChartSub_${i}`,   'Select vehicles to compare');
          renderBestPerfTable();
        }
      });
      tagsEl.appendChild(tag);
    });

    if (limitEl) limitEl.textContent = `${getCompareSelected().length}/${COMPARE_MAX}`;
    if (addBtn)  addBtn.disabled = getCompareSelected().length >= COMPARE_MAX;
  }

  /** Builds the compare vehicle search dropdown list */
  function buildDropList(query) {
    const list = document.getElementById(`compareDropList_${i}`);
    if (!list) return;
    list.innerHTML = '';

    const q = (query || '').trim().toLowerCase();
    Object.keys(vehicleDailyData).sort().forEach(name => {
      if (q && !name.toLowerCase().includes(q)) return;
      if (getCompareSelected().includes(name))   return; /* already selected */

      const d = document.createElement('div');
      d.textContent    = name;
      d.style.cssText  = 'padding:7px 12px;font-size:12px;cursor:pointer;color:var(--text)';
      d.onmouseenter   = () => { d.style.background = 'var(--surface2)'; };
      d.onmouseleave   = () => { d.style.background = ''; };
      d.addEventListener('click', () => {
        if (getCompareSelected().length >= COMPARE_MAX) {
          showToast('warn', 'Limit reached', 'Maximum 5 vehicles can be compared at once.', 3000);
          return;
        }
        setCompareSelected([...getCompareSelected(), name]);
        renderCompareTags();
        document.getElementById(`compareDropPanel_${i}`).style.display = 'none';
        updateCompareCharts();
        /* Force tick callbacks to re-evaluate with updated compare list */
        riskChartInst?.update('none');
        wdweChartInst?.update('none');
        utilChartInst?.update('none');
        prevChartInst?.update('none');
        distCompInst?.update('none');
      });
      list.appendChild(d);
    });

    if (!list.children.length) {
      const empty = document.createElement('div');
      empty.textContent   = 'No vehicles found';
      empty.style.cssText = 'padding:10px 12px;font-size:12px;color:var(--text3)';
      list.appendChild(empty);
    }
  }

  /* Wire up the Add button */
  const addBtn = document.getElementById(`compareAddBtn_${i}`);
  addBtn?.addEventListener('click', () => {
    const panel = document.getElementById(`compareDropPanel_${i}`);
    if (!panel) return;
    const isOpen = panel.style.display === 'block';
    panel.style.display = isOpen ? 'none' : 'block';
    if (!isOpen) {
      const search = document.getElementById(`compareSearch_${i}`);
      if (search) { search.value = ''; }
      buildDropList('');
      setTimeout(() => document.getElementById(`compareSearch_${i}`)?.focus(), 50);
    }
  });

  /* Wire up the search input */
  document.getElementById(`compareSearch_${i}`)
    ?.addEventListener('input', e => buildDropList(e.target.value));

  /* Close dropdown on outside click */
  document.addEventListener('click', e => {
    const panel = document.getElementById(`compareDropPanel_${i}`);
    const btn   = document.getElementById(`compareAddBtn_${i}`);
    if (panel && !panel.contains(e.target) && btn && !btn.contains(e.target)) {
      panel.style.display = 'none';
    }
  });

  renderCompareTags();
  buildDropList('');
}


/* ═══════════════════════════════════════════════════════════════════════════
 * 7. _updateCompareCharts()
 * Updates all charts and KPI cards when compare mode is active.
 * ═══════════════════════════════════════════════════════════════════════════ */

function _updateCompareCharts(
  i, vehicles, vehicleDailyData, dayLabels, dailyTotals,
  hasPrev, prevMap, activeViolations,
  getCompareSelected, getActiveDateFilter,
  dailyChartInst, riskChartInst, violChartInst,
  wdweChartInst,  utilChartInst, prevChartInst, distCompInst,
  riskBarHeight, getCompareColor,
  renderVehicleTable, renderBestPerfTable, generateSummary,
  avgScore, avgIdle, flaggedCount, _activeViolations
) {
  const sel       = getCompareSelected();
  if (!sel.length) return;

  const hasFilter = !!(getActiveDateFilter()?.length);
  const filter    = getActiveDateFilter();
  const labels    = hasFilter ? filter.map(d => d.label) : dayLabels;

  /* ── Daily chart — one line per selected vehicle ── */
  const datasets = sel.map((name, idx) => {
    const col      = getCompareColor(idx);
    const fullData = vehicleDailyData[name] || [];
    const data     = hasFilter
      ? filter.map(d => {
          const di = dayLabels.indexOf(d.label);
          return di > -1 ? fullData[di] || 0 : 0;
        })
      : fullData;
    return {
      label: name, data,
      borderColor     : col,
      backgroundColor : col.replace(')', ',0.08)').replace('rgb', 'rgba'),
      fill: false, tension: 0.3, pointRadius: 3, borderWidth: 2,
      pointHoverRadius: 5, pointBackgroundColor: col,
    };
  });

  if (dailyChartInst) {
    dailyChartInst.data.labels   = labels;
    dailyChartInst.data.datasets = datasets;
    dailyChartInst.update();
  }

  const titleNames = sel.length <= 3
    ? sel.join(', ')
    : sel.slice(0, 2).join(', ') + ' + ' + (sel.length - 2) + ' more';
  setEl(`distChartTitle_${i}`, `Comparing — ${titleNames}`);
  setEl(`distChartSub_${i}`,   `${sel.length} vehicles${hasFilter ? ' · ' + labels.length + ' days' : ''}`);

  /* ── KPI cards for selected vehicles ── */
  const selVehicles  = vehicles.filter(v => sel.includes(v.name));
  const selTotalDist = hasFilter
    ? selVehicles.reduce((sum, v) => {
        const vd = vehicleDailyData[v.name] || [];
        return sum + filter.reduce((s, d) => {
          const di = dayLabels.indexOf(d.label);
          return s + (di > -1 ? vd[di] || 0 : 0);
        }, 0);
      }, 0)
    : selVehicles.reduce((s, v) => s + (v.totalDist || 0), 0);

  const selAvgScore = selVehicles.length
    ? selVehicles.reduce((s, v) => s + (v.score || 0), 0) / selVehicles.length
    : 0;

  /* Update KPI labels and values */
  setKpi(i, 0, 'Comparing', sel.length + ' vehicles', 'selected for comparison');
  document.getElementById(`kpiVal0_${i}`).style.fontSize = '16px';
  document.getElementById(`kpiVal0_${i}`).style.color    = 'var(--accent)';
  setKpi(i, 1,
    hasFilter ? 'Combined distance (filtered)' : 'Combined distance',
    Math.round(selTotalDist).toLocaleString(),
    hasFilter ? `km · ${labels.length} days` : 'km this month'
  );
  setKpi(i, 2, 'Avg score', Math.round(selAvgScore).toLocaleString(), 'lower is safer');
  const chg1 = document.getElementById(`kpiChange1_${i}`);
  if (chg1) chg1.style.display = 'none';

  /* ── Risk chart — collapse to selected vehicles ── */
  const cmpRisk = sel.map((nm, idx) => {
    const v = vehicles.find(v => v.name === nm) || { name: nm, score: 0 };
    return { name: nm, score: v.score || 0, color: getCompareColor(idx) };
  });
  if (riskChartInst) {
    riskChartInst.data.labels                     = cmpRisk.map(v => v.name);
    riskChartInst.data.datasets[0].data           = cmpRisk.map(v => v.score);
    riskChartInst.data.datasets[0].backgroundColor= cmpRisk.map(v => v.color);
    const h = Math.max(120, cmpRisk.length * riskBarHeight + 40);
    setWrapHeight(i, 'risk', h, h);
    riskChartInst.update();
    riskChartInst.resize();
  }
  setEl(`riskTitle_${i}`, `Score comparison — ${sel.length} vehicles`);

  /* ── Violation donut ── */
  if (violChartInst) {
    violChartInst.data.datasets[0].data = activeViolations.map(v =>
      sel.reduce((s, nm) => {
        const vh = vehicles.find(x => x.name === nm) || {};
        return s + (vh[v.key] || 0);
      }, 0)
    );
    violChartInst.update();
  }
  setEl(`violTitle_${i}`, `Violations — ${sel.length} vehicles (combined)`);

  /* ── Weekday/weekend and utilisation charts ── */
  collapseBarChart(wdweChartInst, i, 'wdwe',
    sel.map(nm => ({ name: nm, v: vehicles.find(v => v.name === nm) || {} }))
       .map(({ name, v }) => ({ name, weekday: v.weekdayDist || 0, weekend: v.weekendDist || 0 })),
    ['weekday', 'weekend'], 60, `Weekday vs weekend — ${sel.length} vehicles`,
    `wdweTitle_${i}`);

  collapseBarChart(utilChartInst, i, 'util',
    sel.map(nm => ({ name: nm, v: vehicles.find(v => v.name === nm) || {} }))
       .map(({ name, v }) => ({ name, active: v.daysActive || 0, idle: v.daysIdle || 0 })),
    ['active', 'idle'], 60, `Active vs idle — ${sel.length} vehicles`,
    `utilTitle_${i}`);

  /* ── Month-on-month chart ── */
  if (prevChartInst && hasPrev) {
    const selPrev = sel.map(name => {
      const curr = vehicles.find(v => v.name === name) || {};
      return { name, score: curr.score || 0, prevScore: prevMap[name]?.score || 0 };
    });
    const h = Math.max(120, selPrev.length * 50);
    setWrapHeight(i, 'prev', h, h);
    prevChartInst.data.labels                         = selPrev.map(v => v.name);
    prevChartInst.data.datasets[0].data               = selPrev.map(v => v.score);
    prevChartInst.data.datasets[0].backgroundColor    = selPrev.map(() => '#4f8ef7');
    prevChartInst.data.datasets[0].barPercentage      = 0.35;
    prevChartInst.data.datasets[0].categoryPercentage = 0.5;
    prevChartInst.data.datasets[1].data               = selPrev.map(v => v.prevScore);
    prevChartInst.data.datasets[1].backgroundColor    = selPrev.map(() => 'rgba(79,142,247,0.45)');
    prevChartInst.data.datasets[1].barPercentage      = 0.35;
    prevChartInst.data.datasets[1].categoryPercentage = 0.5;
    prevChartInst.update();
    setEl(`prevTitle_${i}`, `Month-on-month — ${sel.length} vehicles`);
  }

  /* ── Distance comparison chart ── */
  if (distCompInst && hasPrev) {
    const selDC = sel.map(nm => {
      const curr = vehicles.find(v => v.name === nm) || {};
      return { name: nm, currDist: curr.totalDist || 0, prevDist: prevMap[nm]?.totalDist || 0 };
    });
    const h = Math.max(120, selDC.length * 70);
    setWrapHeight(i, 'distComp', h, h);
    distCompInst.data.labels                          = selDC.map(v => v.name);
    distCompInst.data.datasets[0].data                = selDC.map(v => v.currDist);
    distCompInst.data.datasets[0].backgroundColor     = selDC.map(v => v.currDist >= v.prevDist ? '#3db87a' : '#e05353');
    distCompInst.data.datasets[0].barPercentage       = 0.35;
    distCompInst.data.datasets[0].categoryPercentage  = 0.5;
    distCompInst.data.datasets[1].data                = selDC.map(v => v.prevDist);
    distCompInst.data.datasets[1].backgroundColor     = selDC.map(v => v.currDist >= v.prevDist ? '#e05353' : '#3db87a');
    distCompInst.data.datasets[1].barPercentage       = 0.35;
    distCompInst.data.datasets[1].categoryPercentage  = 0.5;
    distCompInst.update();
    setEl(`distCompTitle_${i}`, `Distance comparison — ${sel.length} vehicles`);
  }

  /* Update summary, flagged table, and best performers */
  const cmpVehicles = vehicles.filter(v => sel.includes(v.name));
  setEl(`flaggedTitle_${i}`, `Comparison — ${sel.length} vehicles`);
  renderVehicleTable(cmpVehicles);
  renderBestPerfTable(cmpVehicles);
  generateSummary(cmpVehicles, 'Comparison');
}


/* ═══════════════════════════════════════════════════════════════════════════
 * 8. buildVehicleMode() + _updateVehicleChart()
 * Wires up the single-vehicle searchable dropdown.
 * ═══════════════════════════════════════════════════════════════════════════ */

function buildVehicleMode(
  i, vehicles, vehicleDailyData, dayLabels,
  hasPrev, prevMap, activeViolations,
  dailyChartInst, riskChartInst, violChartInst,
  wdweChartInst,  utilChartInst, prevChartInst, distCompInst,
  riskBarHeight,
  setRiskSel, setWdweSel, setUtilSel, setPrevSel, setDistCompSel,
  getActiveDateFilter,
  renderVehicleTable, renderBestPerfTable, generateSummary
) {
  /* Populate the hidden <select> with all vehicle names */
  const sel = document.getElementById(`vehicleSelect_${i}`);
  if (!sel) return;

  const blank = document.createElement('option');
  blank.value = ''; blank.textContent = 'Select vehicle...'; blank.disabled = true;
  sel.appendChild(blank);

  Object.keys(vehicleDailyData).sort().forEach(name => {
    const opt = document.createElement('option');
    opt.value = name; opt.textContent = name;
    sel.appendChild(opt);
  });
  sel.value = '';

  /* Build the visible custom dropdown list */
  function buildDropList(query) {
    const list = document.getElementById(`vDropList_${i}`);
    if (!list) return;
    list.innerHTML = '';

    const q = (query || '').trim().toLowerCase();
    Array.from(sel.options).forEach(opt => {
      if (!opt.value) return; /* skip the blank placeholder */
      if (q && !opt.value.toLowerCase().includes(q)) return;

      const d = document.createElement('div');
      d.textContent    = opt.value;
      d.style.cssText  = 'padding:7px 12px;font-size:12px;cursor:pointer;color:var(--text)';
      d.onmouseenter   = () => { d.style.background = 'var(--surface2)'; };
      d.onmouseleave   = () => { d.style.background = ''; };
      d.addEventListener('click', () => {
        sel.value = opt.value;
        const lbl = document.getElementById(`vDropLbl_${i}`);
        if (lbl) { lbl.textContent = opt.value; lbl.style.color = 'var(--text)'; }
        document.getElementById(`vDropPanel_${i}`).style.display = 'none';
        document.getElementById(`vDropSearch_${i}`).value        = '';
        buildDropList('');
        updateVehicleChart();
      });
      list.appendChild(d);
    });

    if (!list.children.length) {
      list.innerHTML = '<div style="padding:8px 12px;font-size:12px;color:var(--text3)">No vehicles found</div>';
    }
  }

  /* Toggle dropdown open/closed */
  document.getElementById(`vDropTrigger_${i}`)?.addEventListener('click', () => {
    const panel = document.getElementById(`vDropPanel_${i}`);
    const open  = panel.style.display !== 'none';
    panel.style.display = open ? 'none' : 'block';
    if (!open) {
      document.getElementById(`vDropSearch_${i}`)?.focus();
      buildDropList('');
    }
  });

  /* Search input filters the list */
  document.getElementById(`vDropSearch_${i}`)
    ?.addEventListener('input', e => buildDropList(e.target.value));

  /* Close on outside click */
  document.addEventListener('click', e => {
    const wrap = document.getElementById(`vDrop_${i}`);
    if (wrap && !wrap.contains(e.target)) {
      document.getElementById(`vDropPanel_${i}`).style.display = 'none';
    }
  });

  /* <select> change event (backup for programmatic changes) */
  sel.addEventListener('change', () => updateVehicleChart());

  buildDropList('');

  function updateVehicleChart() {
    _updateVehicleChart(
      i, vehicles, vehicleDailyData, dayLabels,
      hasPrev, prevMap, activeViolations,
      dailyChartInst, riskChartInst, violChartInst,
      wdweChartInst,  utilChartInst, prevChartInst, distCompInst,
      riskBarHeight,
      setRiskSel, setWdweSel, setUtilSel, setPrevSel, setDistCompSel,
      getActiveDateFilter,
      renderVehicleTable, renderBestPerfTable, generateSummary
    );
  }
}

function _updateVehicleChart(
  i, vehicles, vehicleDailyData, dayLabels,
  hasPrev, prevMap, activeViolations,
  dailyChartInst, riskChartInst, violChartInst,
  wdweChartInst,  utilChartInst, prevChartInst, distCompInst,
  riskBarHeight,
  setRiskSel, setWdweSel, setUtilSel, setPrevSel, setDistCompSel,
  getActiveDateFilter,
  renderVehicleTable, renderBestPerfTable, generateSummary
) {
  const sel  = document.getElementById(`vehicleSelect_${i}`);
  const name = sel?.value;
  if (!name) return;

  const veh      = vehicles.find(v => v.name === name) || {};
  const data     = vehicleDailyData[name] || [];
  const vScore   = veh.score || 0;
  const isFilter = !!(getActiveDateFilter()?.length);
  const filter   = getActiveDateFilter();

  /* Daily distance stats */
  const totalKm  = data.reduce((s, v) => s + v, 0);
  const active   = data.filter(v => v > 0).length;
  const idle     = data.filter(v => v === 0).length;
  const maxKm    = data.length ? Math.max(...data) : 0;
  const maxDay   = dayLabels[data.indexOf(maxKm)] || '—';
  const avgKm    = active > 0 ? totalKm / active : 0;

  /* ── Daily chart ── */
  if (dailyChartInst) {
    dailyChartInst.data.datasets[0].data               = data;
    dailyChartInst.data.datasets[0].label              = name;
    dailyChartInst.data.datasets[0].borderColor        = '#3db87a';
    dailyChartInst.data.datasets[0].backgroundColor    = 'rgba(61,184,122,0.08)';
    dailyChartInst.data.datasets[0].pointBackgroundColor = '#3db87a';
    dailyChartInst.update();
  }
  setEl(`distChartTitle_${i}`, `Daily distance — ${name}`);
  setEl(`distChartSub_${i}`,   'km covered by this vehicle each day');

  /* ── KPI cards ── */
  const filteredData = isFilter
    ? filter.map(d => { const di = dayLabels.indexOf(d.label); return di > -1 ? data[di] || 0 : 0; })
    : null;
  const dispDist  = filteredData ? Math.round(filteredData.reduce((s, v) => s + v, 0)) : Math.round(veh.totalDist || 0);
  const dispIdle  = filteredData ? filteredData.filter(v => v === 0).length : idle;

  setKpi(i, 0, 'Selected vehicle', name, `of ${vehicles.length} total`);
  document.getElementById(`kpiVal0_${i}`).style.fontSize = '13px';
  document.getElementById(`kpiVal0_${i}`).style.color    = 'var(--accent)';
  setKpi(i, 1,
    isFilter ? 'Distance (filtered)' : 'Total distance',
    dispDist.toLocaleString(),
    isFilter ? `km · ${filter.length} days` : 'km this month'
  );
  setKpi(i, 2, 'Advanced score', vScore.toLocaleString(), 'lower is safer');
  setKpi(i, 3,
    isFilter ? 'Idle days (filtered)' : 'Idle days',
    dispIdle,
    isFilter ? `of ${filter.length} days` : 'days no movement'
  );
  setKpi(i, 4, 'Active days',               active,                            'days with trips');
  setKpi(i, 5, 'Best day',                  Math.round(maxKm).toLocaleString() + ' km', `on ${maxDay}`);
  setKpi(i, 6, 'Avg per active day',        Math.round(avgKm).toLocaleString() + ' km', 'when moving');
  document.getElementById(`kpiVal5_${i}`).style.color = 'var(--green)';
  const chg1 = document.getElementById(`kpiChange1_${i}`);
  if (chg1) chg1.style.display = 'none';

  /* ── Violation donut ── */
  if (violChartInst) {
    violChartInst.data.datasets[0].data = activeViolations.map(v => veh[v.key] || 0);
    violChartInst.update();
  }
  setEl(`violTitle_${i}`, `Violation breakdown — ${name}`);

  /* ── Risk chart — single bar ── */
  setRiskSel(name);
  const band    = getScoreBand(vScore);
  const singleH = Math.max(80, riskBarHeight + 40);
  if (riskChartInst) {
    riskChartInst.data.labels                     = [name];
    riskChartInst.data.datasets[0].data           = [vScore];
    riskChartInst.data.datasets[0].backgroundColor= [band.color];
    setWrapHeight(i, 'risk', singleH, singleH);
    riskChartInst.update();
    riskChartInst.resize();
  }
  setEl(`riskTitle_${i}`, `Advanced score — ${name}`);

  /* ── Weekday/weekend — single bar ── */
  setWdweSel(name);
  const singleWdH = Math.max(80, 60 + 40);
  setWrapHeight(i, 'wdwe', singleWdH, singleWdH);
  if (wdweChartInst) {
    wdweChartInst.data.labels                          = [name];
    wdweChartInst.data.datasets[0].data                = [veh.weekdayDist || 0];
    wdweChartInst.data.datasets[0].backgroundColor     = ['#4f8ef7'];
    wdweChartInst.data.datasets[0].barPercentage       = 0.35;
    wdweChartInst.data.datasets[0].categoryPercentage  = 0.5;
    wdweChartInst.data.datasets[1].data                = [veh.weekendDist || 0];
    wdweChartInst.data.datasets[1].backgroundColor     = ['#2ec4b6'];
    wdweChartInst.data.datasets[1].barPercentage       = 0.35;
    wdweChartInst.data.datasets[1].categoryPercentage  = 0.5;
    wdweChartInst.update();
  }
  setEl(`wdweTitle_${i}`, `Weekday vs weekend — ${name}`);

  /* ── Utilisation — single bar ── */
  setUtilSel(name);
  const singleUtH = Math.max(80, 60 + 40);
  setWrapHeight(i, 'util', singleUtH, singleUtH);
  if (utilChartInst) {
    utilChartInst.data.labels                          = [name];
    utilChartInst.data.datasets[0].data                = [veh.daysActive || 0];
    utilChartInst.data.datasets[0].backgroundColor     = ['#4f8ef7'];
    utilChartInst.data.datasets[0].barPercentage       = 0.35;
    utilChartInst.data.datasets[0].categoryPercentage  = 0.5;
    utilChartInst.data.datasets[1].data                = [veh.daysIdle || 0];
    utilChartInst.data.datasets[1].backgroundColor     = ['#555b72'];
    utilChartInst.data.datasets[1].barPercentage       = 0.35;
    utilChartInst.data.datasets[1].categoryPercentage  = 0.5;
    utilChartInst.update();
  }
  setEl(`utilTitle_${i}`, `Active vs idle — ${name}`);

  /* ── Month-on-month — single bar ── */
  setPrevSel(name);
  if (prevChartInst && hasPrev) {
    const prevScore  = prevMap[name]?.score || 0;
    const singlePH   = Math.max(80, 50 + 40);
    setWrapHeight(i, 'prev', singlePH, singlePH);
    prevChartInst.data.labels                          = [name];
    prevChartInst.data.datasets[0].data                = [veh.score || 0];
    prevChartInst.data.datasets[0].backgroundColor     = ['#4f8ef7'];
    prevChartInst.data.datasets[0].barPercentage       = 0.35;
    prevChartInst.data.datasets[0].categoryPercentage  = 0.5;
    prevChartInst.data.datasets[1].data                = [prevScore];
    prevChartInst.data.datasets[1].backgroundColor     = ['rgba(79,142,247,0.45)'];
    prevChartInst.data.datasets[1].barPercentage       = 0.35;
    prevChartInst.data.datasets[1].categoryPercentage  = 0.5;
    prevChartInst.update();
    setEl(`prevTitle_${i}`, `Month-on-month — ${name}`);
  }

  /* ── Distance comparison — single bar ── */
  setDistCompSel(name);
  if (distCompInst && hasPrev) {
    const prev     = prevMap[name]?.totalDist || 0;
    const curr     = veh.totalDist || 0;
    setWrapHeight(i, 'distComp', 120, 120);
    distCompInst.data.labels                          = [name];
    distCompInst.data.datasets[0].data                = [curr];
    distCompInst.data.datasets[0].backgroundColor     = [curr >= prev ? '#3db87a' : '#e05353'];
    distCompInst.data.datasets[0].barPercentage       = 0.35;
    distCompInst.data.datasets[0].categoryPercentage  = 0.5;
    distCompInst.data.datasets[1].data                = [prev];
    distCompInst.data.datasets[1].backgroundColor     = [curr >= prev ? '#e05353' : '#3db87a'];
    distCompInst.data.datasets[1].barPercentage       = 0.35;
    distCompInst.data.datasets[1].categoryPercentage  = 0.5;
    distCompInst.update();
    setEl(`distCompTitle_${i}`, `Distance comparison — ${name}`);
  }

  /* ── Summary and tables ── */
  generateSummary([veh], name);
  renderVehicleTable([veh]);
  renderBestPerfTable([veh]);
  setEl(`flaggedTitle_${i}`, `Details — ${name}`);
}


/* ═══════════════════════════════════════════════════════════════════════════
 * 9. buildDateFilter()
 * Sets date input defaults and wires up Apply / Reset buttons.
 * ═══════════════════════════════════════════════════════════════════════════ */

function buildDateFilter(
  i, dayInfo, dayLabels, dailyTotals, vehicleDailyData,
  vehicles, utilRows, client,
  avgScore, avgIdle, vehiclesWithTrips, flaggedCount, activeViolations,
  getActiveDateFilter, setActiveDateFilter,
  getCompareSelected,
  dailyChartInst,
  renderVehicleTable, renderBestPerfTable, updateCompareCharts
) {
  let initialFrom      = '';
  let initialTo        = '';
  let initialFiltered  = false;

  const pad2 = n => String(n).padStart(2, '0');
  const fmt  = d => `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;

  /* ── Set default From/To values ── */
  const datesWithDate = dayInfo.filter(d => d.date);
  const cm            = client?.month || '';

  if (cm) {
    const [yr, mo] = cm.split('-').map(Number);
    const days     = new Date(yr, mo, 0).getDate();
    document.getElementById(`dateFrom_${i}`).value = `${yr}-${pad2(mo)}-01`;
    document.getElementById(`dateTo_${i}`).value   = `${yr}-${pad2(mo)}-${pad2(days)}`;
  } else if (datesWithDate.length) {
    const hasReal = datesWithDate.some(d => d.date.getFullYear() > 2000);
    if (hasReal) {
      document.getElementById(`dateFrom_${i}`).value = fmt(datesWithDate[0].date);
      document.getElementById(`dateTo_${i}`).value   = fmt(datesWithDate[datesWithDate.length - 1].date);
    }
  }

  initialFrom = document.getElementById(`dateFrom_${i}`).value;
  initialTo   = document.getElementById(`dateTo_${i}`).value;

  /* ── Apply button ── */
  function apply(silent) {
    const fromVal = document.getElementById(`dateFrom_${i}`).value;
    const toVal   = document.getElementById(`dateTo_${i}`).value;
    if (!fromVal || !toVal) return;

    const reportMonth = client?.month || null;

    /* Validate that the selected range is within the report month */
    if (reportMonth) {
      const fromKey = fromVal.slice(0, 7);
      const toKey   = toVal.slice(0, 7);
      if (fromKey !== reportMonth || toKey !== reportMonth) {
        setActiveDateFilter(null);
        if (dailyChartInst) {
          dailyChartInst.data.labels            = [];
          dailyChartInst.data.datasets[0].data  = [];
          dailyChartInst.update();
        }
        setEl(`distChartTitle_${i}`, 'No data available for selected range');
        setEl(`distChartSub_${i}`,   `This report contains data for ${reportMonth}.`);
        document.getElementById(`dateResetBtn_${i}`).style.display = '';
        if (!silent) showToast('warn', 'No data available',
          `This report contains data for ${reportMonth}.`, 5000);
        return;
      }
    }

    /* Filter dayInfo to the selected range */
    const hasReal = dayInfo.some(d => d.date && d.date.getFullYear() > 2000);
    let filtered;

    if (hasReal) {
      const [fy, fm, fd] = fromVal.split('-').map(Number);
      const [ty, tm, td] = toVal.split('-').map(Number);
      const from = new Date(fy, fm - 1, fd);
      const to   = new Date(ty, tm - 1, td);
      if (from > to) {
        showToast('warn', 'Invalid range', 'Start date must be before end date.', 4000);
        return;
      }
      filtered = dayInfo.filter(d => {
        if (!d.date) return false;
        const nd = new Date(d.date.getFullYear(), d.date.getMonth(), d.date.getDate());
        return nd >= from && nd <= to;
      });
    } else {
      const fromDay = parseInt(fromVal.split('-')[2]) || 1;
      const toDay   = parseInt(toVal.split('-')[2])   || 31;
      if (fromDay > toDay) {
        showToast('warn', 'Invalid range', 'Start day must be before end day.', 4000);
        return;
      }
      filtered = dayInfo.filter(d => d.dayNum && d.dayNum >= fromDay && d.dayNum <= toDay);
    }

    if (!filtered.length) {
      setActiveDateFilter(null);
      if (!silent) showToast('warn', 'No data available',
        'Select a date range that exists in this report.', 5000);
      document.getElementById(`dateResetBtn_${i}`).style.display = '';
      return;
    }

    setActiveDateFilter(filtered);
    const filteredLabels = filtered.map(d => d.label);

    /* If compare mode is active, rebuild compare charts with filtered dates */
    if (getCompareSelected().length > 0) {
      updateCompareCharts();
      document.getElementById(`dateResetBtn_${i}`).style.display = '';
      if (!silent) showToast('success', 'Date range applied',
        `Showing ${filtered.length} day${filtered.length !== 1 ? 's' : ''} from ${filteredLabels[0]} to ${filteredLabels[filteredLabels.length - 1]}`, 3000);
      return;
    }

    /* Fleet mode — update daily chart with filtered totals */
    const filteredTotals = filtered.map(d =>
      utilRows.reduce((s, r) => s + Number(r[d.key] || 0), 0)
    );
    if (dailyChartInst) {
      dailyChartInst.data.labels            = filteredLabels;
      dailyChartInst.data.datasets[0].data  = filteredTotals;
      dailyChartInst.data.datasets[0].borderColor     = '#4f8ef7';
      dailyChartInst.data.datasets[0].backgroundColor = 'rgba(79,142,247,0.08)';
      dailyChartInst.data.datasets[0].pointBackgroundColor = '#4f8ef7';
      dailyChartInst.update();
    }
    setEl(`distChartTitle_${i}`,
      `Daily distance — ${filteredLabels[0]} to ${filteredLabels[filteredLabels.length - 1]}`);
    setEl(`distChartSub_${i}`,
      `${filtered.length} days selected · ${Math.round(filteredTotals.reduce((s, v) => s + v, 0)).toLocaleString()} km`);

    document.getElementById(`dateResetBtn_${i}`).style.display = '';
    const chg1 = document.getElementById(`kpiChange1_${i}`);
    if (chg1) chg1.style.display = 'none';

    if (!silent) showToast('success', 'Date range applied',
      `Showing ${filtered.length} day${filtered.length !== 1 ? 's' : ''} from ${filteredLabels[0]} to ${filteredLabels[filteredLabels.length - 1]}`, 3000);
  }

  /* ── Reset button ── */
  function reset() {
    const fromEl = document.getElementById(`dateFrom_${i}`);
    const toEl   = document.getElementById(`dateTo_${i}`);
    if (fromEl && initialFrom) fromEl.value = initialFrom;
    if (toEl   && initialTo)   toEl.value   = initialTo;
    setActiveDateFilter(null);

    if (dailyChartInst) {
      dailyChartInst.data.labels            = dayLabels;
      dailyChartInst.data.datasets[0].data  = Object.values(dailyTotals);
      dailyChartInst.data.datasets[0].borderColor     = '#4f8ef7';
      dailyChartInst.data.datasets[0].backgroundColor = 'rgba(79,142,247,0.08)';
      dailyChartInst.data.datasets[0].pointBackgroundColor = '#4f8ef7';
      dailyChartInst.update();
    }
    setEl(`distChartTitle_${i}`, 'Daily fleet distance (km)');
    setEl(`distChartSub_${i}`,   'Total km covered by all vehicles each day');
    document.getElementById(`dateResetBtn_${i}`).style.display = 'none';
    const chg1 = document.getElementById(`kpiChange1_${i}`);
    if (chg1) chg1.style.display = '';
  }

  document.getElementById(`dateApplyBtn_${i}`)?.addEventListener('click', () => apply(false));
  document.getElementById(`dateResetBtn_${i}`)?.addEventListener('click', () => reset());

  return {
    apply,
    reset,
    setInitialFiltered: v => { initialFiltered = v; },
  };
}


/* ═══════════════════════════════════════════════════════════════════════════
 * 10. buildTables() — flagged vehicles + best performers
 * ═══════════════════════════════════════════════════════════════════════════ */

/**
 * _renderVehicleTable()
 * Renders the flagged / all-vehicles table.
 * Uses getScoreBand() from constants.js for consistent risk labelling.
 */
function _renderVehicleTable(
  i, vehicles, flaggedVehicles, hasPrev, prevMap,
  activeViolations, overrideVehicles, forceAll
) {
  const tbody = document.getElementById(`alertBody_${i}`);
  if (!tbody) return;
  tbody.innerHTML = '';

  let tableVehicles;
  if (forceAll) {
    tableVehicles = [...vehicles].sort((a, b) => (b.score || 0) - (a.score || 0));
    setEl(`flaggedTitle_${i}`, `All vehicles (${vehicles.length})`);
  } else if (overrideVehicles) {
    tableVehicles = overrideVehicles;
  } else {
    tableVehicles = flaggedVehicles;
    setEl(`flaggedTitle_${i}`, 'Vehicles Scoring');
  }

  tableVehicles.forEach(v => {
    const score = v.score || 0;
    const band  = getScoreBand(score);
    /* Map band label to badge class */
    const badgeCls = score > SCORE_BANDS.high.min - 1 ? 'b-red'
                   : score > SCORE_BANDS.safe.max     ? 'b-amber'
                   : 'b-green';

    const prev = prevMap[v.name];
    let changeCell     = '';
    let distChangeCell = '';

    if (hasPrev) {
      /* Score change cell */
      const ps         = prev?.score ?? null;
      const prevScoreStr = ps !== null ? ps.toLocaleString() : '—';
      let chgStr       = '<span style="color:var(--text3)">—</span>';
      if (ps !== null && ps > 0) {
        const d    = score - ps;
        const rawP = Math.abs(d) / ps * 100;
        const pct  = Math.round(rawP);
        if      (rawP > 100) chgStr = d > 0
          ? `<span class="arr-up">▲ +${Math.round(Math.abs(d))} worse</span>`
          : `<span class="arr-down">▼ -${Math.round(Math.abs(d))} better</span>`;
        else if (rawP > 1)   chgStr = d > 0
          ? `<span class="arr-up">▲ ${pct}% worse</span>`
          : `<span class="arr-down">▼ ${pct}% better</span>`;
        else                 chgStr = '<span style="color:var(--text3)">→</span>';
      }
      changeCell = `<td>${prevScoreStr}</td><td>${chgStr}</td>`;

      /* Distance change cell */
      if (prev) {
        const pd = prev.totalDist || 0;
        const cd = v.totalDist   || 0;
        if (pd > 0) {
          const dd   = cd - pd;
          const rawD = Math.abs(dd) / pd * 100;
          const dpct = Math.round(rawD);
          let distChg;
          if      (rawD > 100) distChg = dd > 0
            ? `<span style="color:var(--green)">▲ +${Math.round(Math.abs(dd)).toLocaleString()} km</span>`
            : `<span style="color:var(--red)">▼ -${Math.round(Math.abs(dd)).toLocaleString()} km</span>`;
          else if (rawD > 1)   distChg = dd > 0
            ? `<span style="color:var(--green)">▲ ${dpct}%</span>`
            : `<span style="color:var(--red)">▼ ${dpct}%</span>`;
          else                 distChg = '<span style="color:var(--text3)">→</span>';
          distChangeCell = `<td>${Math.round(pd).toLocaleString()}</td><td>${distChg}</td>`;
        }
      }
      if (!distChangeCell) {
        distChangeCell = '<td>—</td><td><span style="color:var(--text3)">—</span></td>';
      }
    }

    const utilCells =
      `<td>${(v.daysActive  !== undefined ? v.daysActive  : '—')}</td>` +
      `<td>${(v.daysIdle    !== undefined ? v.daysIdle    : '—')}</td>` +
      `<td>${(v.weekdayDist !== undefined ? Math.round(v.weekdayDist).toLocaleString() : '—')}</td>` +
      `<td>${(v.weekendDist !== undefined ? Math.round(v.weekendDist).toLocaleString() : '—')}</td>`;

    const tr = document.createElement('tr');
    tr.innerHTML =
      `<td>${escapeHTML(v.name)}</td>` +
      `<td>${Math.round(v.totalDist || 0).toLocaleString()}</td>` +
      distChangeCell +
      `<td>${score.toLocaleString()}</td>` +
      changeCell +
      utilCells +
      activeViolations.map(vd => `<td>${(v[vd.key] || 0).toLocaleString()}</td>`).join('') +
      `<td><span class="badge ${badgeCls}">${band.label}</span></td>`;
    tbody.appendChild(tr);
  });
}

/**
 * _renderBestPerfTable()
 * Renders the top-performers table ranked by violations per 100 km.
 * Vehicles with 0 km are excluded to avoid division-by-zero.
 */
function _renderBestPerfTable(i, vehicles, overrideVehicles) {
  const tbody = document.getElementById(`bestPerfBody_${i}`);
  if (!tbody) return;
  tbody.innerHTML = '';

  const src   = overrideVehicles?.length ? overrideVehicles : vehicles;
  const perfs = [...src]
    .filter(v => (v.totalDist || 0) > 0)
    .map(v => {
      const totalViol = VIOLATIONS.reduce((s, vd) => s + (v[vd.key] || 0), 0);
      const per100    = totalViol / (v.totalDist / 100);
      const rating    = per100 === 0 ? 'Excellent' : per100 < 1 ? 'Good' : per100 < 3 ? 'Fair' : 'Poor';
      return { ...v, totalViol, per100, rating };
    })
    .sort((a, b) => a.per100 - b.per100)
    .slice(0, 15);

  if (!perfs.length) {
    tbody.innerHTML = '<tr><td colspan="7" style="color:var(--text3)">No ranked vehicles for this selection.</td></tr>';
    return;
  }

  perfs.forEach((v, idx) => {
    const rCls = v.rating === 'Excellent' || v.rating === 'Good' ? 'b-green'
               : v.rating === 'Fair' ? 'b-amber' : 'b-red';
    const tr   = document.createElement('tr');
    tr.innerHTML =
      `<td>${idx + 1}</td>` +
      `<td>${escapeHTML(v.name)}</td>` +
      `<td>${Math.round(v.totalDist).toLocaleString()}</td>` +
      `<td>${v.totalViol.toLocaleString()}</td>` +
      `<td>${v.per100.toFixed(2)}</td>` +
      `<td>${(v.score || 0).toLocaleString()}</td>` +
      `<td><span class="badge ${rCls}">${v.rating}</span></td>`;
    tbody.appendChild(tr);
  });
}


/* ═══════════════════════════════════════════════════════════════════════════
 * 11. buildExecSummaryControls() + _generateSummary()
 * ═══════════════════════════════════════════════════════════════════════════ */

function buildExecSummaryControls(i, client, generateSummary) {
  /* Edit button — makes summary contenteditable */
  document.getElementById(`execEditBtn_${i}`)?.addEventListener('click', () => {
    const body   = document.getElementById(`execBody_${i}`);
    const edit   = document.getElementById(`execEditBtn_${i}`);
    const save   = document.getElementById(`execSaveBtn_${i}`);
    const cancel = document.getElementById(`execCancelBtn_${i}`);
    if (!body) return;

    if (client) client.draftSummaryHTML = body.innerHTML;
    body.contentEditable = 'true';
    body.classList.add('is-editing');
    edit.style.display   = 'none';
    save.style.display   = '';
    cancel.style.display = '';
    body.focus();
  });

  /* Save button — persists the edited HTML */
  document.getElementById(`execSaveBtn_${i}`)?.addEventListener('click', () => {
    const body = document.getElementById(`execBody_${i}`);
    if (!body) return;
    if (client) client.customSummaryHTML = sanitiseSummaryHTML(body.innerHTML);
    body.contentEditable = 'false';
    body.classList.remove('is-editing');
    document.getElementById(`execEditBtn_${i}`).style.display   = '';
    document.getElementById(`execSaveBtn_${i}`).style.display   = 'none';
    document.getElementById(`execCancelBtn_${i}`).style.display = 'none';
    showToast('success', 'Summary saved', 'Your edits were saved.', 2500);
  });

  /* Cancel button — restores the pre-edit draft */
  document.getElementById(`execCancelBtn_${i}`)?.addEventListener('click', () => {
    const body = document.getElementById(`execBody_${i}`);
    if (!body) return;
    if (client?.draftSummaryHTML !== undefined) body.innerHTML = client.draftSummaryHTML;
    body.contentEditable = 'false';
    body.classList.remove('is-editing');
    document.getElementById(`execEditBtn_${i}`).style.display   = '';
    document.getElementById(`execSaveBtn_${i}`).style.display   = 'none';
    document.getElementById(`execCancelBtn_${i}`).style.display = 'none';
  });

  /* Auto button — regenerates the summary from current data */
  document.getElementById(`execAutoBtn_${i}`)?.addEventListener('click', () => {
    if (client) client.customSummaryHTML = '';
    generateSummary(null, null, true);
    showToast('info', 'Summary reset', 'Auto-generated summary restored.', 2500);
  });
}

/**
 * _generateSummary()
 * Auto-generates a structured executive summary from fleet data.
 * Sections: Overview, Score trend, Distance & utilisation,
 *           High-risk vehicles, Top performers, Violation breakdown,
 *           Recommended actions.
 */
function _generateSummary(
  i, client, vehicles, prevMap, prevVehicles,
  activeViolations, hasPrev, avgIdleFleet,
  selVehicles, label, forceRefresh
) {
  const bodyEl = document.getElementById(`execBody_${i}`);
  if (!bodyEl) return;

  /* If the user has a custom saved summary, show that unless forced to refresh */
  if (!forceRefresh && client?.customSummaryHTML) {
    bodyEl.innerHTML = client.customSummaryHTML;
    return;
  }

  const vList      = selVehicles || vehicles;
  const isFiltered = selVehicles && selVehicles.length < vehicles.length;
  const fleetScore = Math.round(vList.reduce((s, v) => s + (v.score || 0), 0) / Math.max(1, vList.length));
  const totalD     = vList.reduce((s, v) => s + (v.totalDist || 0), 0);
  const lines      = [];

  /* ── Overview ── */
  if (isFiltered) {
    lines.push(
      `<strong>${escapeHTML(label || vList.map(v => v.name).join(', '))}</strong> — ` +
      `${vList.length === 1 ? 'individual vehicle analysis' : 'comparison of ' + vList.length + ' vehicles'}.`
    );
  } else {
    const band        = getScoreBand(fleetScore);
    const healthLabel = `<span style="color:${band.color}">${band.label}</span>`;
    lines.push(
      `Fleet of <strong>${vList.length} vehicles</strong> · Status: ${healthLabel} · ` +
      `Average Advanced Score: <strong>${fleetScore}</strong>.`
    );
  }

  /* ── Score trend vs previous month ── */
  if (hasPrev && prevVehicles.length && !isFiltered) {
    const prevAvg = prevVehicles.reduce((s, v) => s + (v.score || 0), 0) / prevVehicles.length;
    const diff    = fleetScore - Math.round(prevAvg);
    const rawPct  = Math.abs(diff) / Math.max(1, prevAvg) * 100;
    if (rawPct >= 1) {
      const dir    = diff < 0
        ? '<span class="good">improved</span>'
        : '<span class="bad">worsened</span>';
      const pctStr = rawPct > 100
        ? `${Math.abs(diff)} points`
        : `${Math.round(rawPct)}%`;
      lines.push(
        `<strong>Score Trend:</strong> Fleet average score ${dir} by ` +
        `<strong>${diff < 0 ? '▼' : '▲'} ${pctStr}</strong> vs last month ` +
        `(${Math.round(prevAvg)} → ${fleetScore}). ` +
        `${diff < 0 ? 'Driving behaviour is heading in the right direction.' : 'Immediate fleet-wide coaching is recommended.'}`
      );
    } else {
      lines.push(
        `<strong>Score Trend:</strong> Fleet average score remained ` +
        `<span class="neutral">stable</span> at <strong>${fleetScore}</strong>.`
      );
    }
  } else if (isFiltered) {
    vList.forEach(v => {
      const pv   = prevMap[v.name];
      const diff = pv ? (v.score || 0) - (pv.score || 0) : null;
      if (diff !== null && Math.abs(diff) > 0) {
        const dir = diff < 0 ? '<span class="good">improved</span>' : '<span class="bad">worsened</span>';
        lines.push(`${escapeHTML(v.name)}: score <strong>${v.score || 0}</strong> — ${dir} by <strong>${diff < 0 ? '▼' : '▲'} ${Math.abs(diff)}</strong> vs last month.`);
      } else {
        lines.push(`${escapeHTML(v.name)}: score <strong>${v.score || 0}</strong>.`);
      }
    });
  }

  /* ── Distance & utilisation ── */
  const vWithActive   = vList.filter(v => (v.daysActive || 0) > 0);
  const avgActiveDays = vWithActive.length
    ? Math.round(vWithActive.reduce((s, v) => s + (v.daysActive || 0), 0) / vWithActive.length)
    : 0;
  const idleVehicles  = vList.filter(v => (v.daysIdle || 0) >= 5);

  if (!isFiltered) {
    lines.push(
      `<strong>Distance & Utilisation:</strong> Fleet covered ` +
      `<strong>${Math.round(totalD).toLocaleString()} km</strong> this period. ` +
      `Average active days: <strong>${avgActiveDays} days</strong>. ` +
      (idleVehicles.length > 0
        ? `<span class="bad">${idleVehicles.length} vehicle${idleVehicles.length !== 1 ? 's' : ''}</span> recorded 5+ idle weekdays — review deployment schedules.`
        : 'All vehicles maintained adequate utilisation.')
    );
  } else {
    lines.push(
      `<strong>Distance:</strong> ` +
      vList.map(v => `${escapeHTML(v.name)}: <strong>${Math.round(v.totalDist || 0).toLocaleString()} km</strong>`).join(' · ') + '.'
    );
  }

  /* ── High-risk vehicles ── */
  const highRisk   = [...vList].filter(v => (v.score || 0) > SCORE_BANDS.high.min - 1)
                               .sort((a, b) => (b.score || 0) - (a.score || 0)).slice(0, 5);
  const moderate   = [...vList].filter(v => (v.score || 0) > SCORE_BANDS.safe.max && (v.score || 0) <= SCORE_BANDS.moderate.max)
                               .sort((a, b) => (b.score || 0) - (a.score || 0)).slice(0, 3);

  if (highRisk.length) {
    lines.push(
      `<strong>High Risk Vehicles (Score ${SCORE_BANDS.high.min}+):</strong> ` +
      highRisk.map(v => `<strong>${escapeHTML(v.name)}</strong> (${v.score})`).join(', ') +
      `. ${highRisk.length === 1 ? 'This vehicle requires' : 'These vehicles require'} immediate intervention.`
    );
  }
  if (moderate.length && vList.length > 3) {
    lines.push(
      `<strong>Needs Attention (Score ${SCORE_BANDS.moderate.min}–${SCORE_BANDS.moderate.max}):</strong> ` +
      moderate.map(v => `<strong>${escapeHTML(v.name)}</strong> (${v.score})`).join(', ') +
      `. Targeted coaching recommended.`
    );
  }

  /* ── Top performers ── */
  const zeroes     = [...vList].filter(v => (v.score || 0) === 0 && (v.totalDist || 0) > 0).slice(0, 3);
  const lowScorers = [...vList].filter(v => (v.score || 0) > 0 && (v.score || 0) <= SCORE_BANDS.safe.max)
                               .sort((a, b) => (a.score || 0) - (b.score || 0)).slice(0, 3);

  if (zeroes.length) {
    lines.push(
      `<strong>Top Performers (Score 0):</strong> ` +
      zeroes.map(v => `<strong>${escapeHTML(v.name)}</strong>`).join(', ') +
      ` recorded zero violations — excellent driving behaviour.`
    );
  } else if (lowScorers.length && vList.length > 3) {
    lines.push(
      `<strong>Top Performers:</strong> ` +
      lowScorers.map(v => `<strong>${escapeHTML(v.name)}</strong> (${v.score})`).join(', ') +
      ` — safe driving, within acceptable parameters.`
    );
  }

  /* ── Violation breakdown ── */
  const violCounts = activeViolations
    .map(v => ({ ...v, count: vList.reduce((s, vh) => s + (vh[v.key] || 0), 0) }))
    .filter(v => v.count > 0)
    .sort((a, b) => b.count - a.count);

  if (violCounts.length) {
    const topViol     = violCounts.slice(0, 3)
      .map((v, idx) => idx === 0 ? `<strong>${v.short} (${v.count.toLocaleString()})</strong>` : `${v.short} (${v.count.toLocaleString()})`)
      .join(', ');
    const highViols   = violCounts.filter(v => v.risk === 'high');
    lines.push(
      `<strong>Violation Summary:</strong> Top violations — ${topViol}. ` +
      (highViols.length
        ? `High-severity events: <span class="bad">${highViols.map(v => v.short).join(', ')}</span>. These carry the highest risk.`
        : '')
    );
  }

  /* ── Recommended actions ── */
  const actions = [];
  if (highRisk.length)
    actions.push(`Schedule driver interviews for ${highRisk.length} high-risk vehicle${highRisk.length !== 1 ? 's' : ''}`);
  if (idleVehicles.length && !isFiltered)
    actions.push(`Review deployment for ${idleVehicles.length} frequently idle vehicle${idleVehicles.length !== 1 ? 's' : ''}`);
  if (violCounts.find(v => v.key?.toLowerCase().includes('brake') && v.count > 0))
    actions.push('Run defensive driving refresher focused on braking distances');
  if (violCounts.find(v => v.key?.toLowerCase().includes('speed') && v.count > 0))
    actions.push('Enforce speed policy — consider geofenced speed alerts');

  if (actions.length) {
    lines.push(
      `<strong>Recommended Actions:</strong> ` +
      actions.map((a, idx) => `${idx + 1}. ${a}`).join(' &nbsp;·&nbsp; ') + '.'
    );
  }

  const html = lines.map(l => `<p>${l}</p>`).join('');
  bodyEl.innerHTML   = html;
  if (client) client.autoSummaryHTML = html;
}


/* ═══════════════════════════════════════════════════════════════════════════
 * 12. buildViolationGuide()
 * Renders the violation reference cards at the bottom of the dashboard.
 * ═══════════════════════════════════════════════════════════════════════════ */

function buildViolationGuide(i, vehicles, prevVehicles, hasPrev, activeViolations) {
  const guide = document.getElementById(`violGuide_${i}`);
  if (!guide) return;

  VIOLATIONS.forEach(v => {
    const riskCls  = v.risk === 'high' ? 'viol-high' : v.risk === 'med' ? 'viol-med' : 'viol-low';
    const badgeCls = v.risk === 'high' ? 'b-red'     : v.risk === 'med' ? 'b-amber'  : 'b-green';
    const count    = vehicles.reduce((s, vh) => s + (vh[v.key] || 0), 0);

    let trendStr = '';
    if (hasPrev && prevVehicles.length) {
      const prevCount = prevVehicles.reduce((s, vh) => s + (vh[v.key] || 0), 0);
      if (prevCount > 0) {
        const d    = count - prevCount;
        const rawP = Math.abs(d) / prevCount * 100;
        const pct  = Math.round(rawP);
        if (rawP > 100) {
          const sign = d > 0 ? '+' : '-';
          trendStr = d > 0
            ? `<span style="color:var(--red);font-size:11px">▲ ${sign}${Math.round(Math.abs(d))} vs last month</span>`
            : `<span style="color:var(--green);font-size:11px">▼ ${sign}${Math.round(Math.abs(d))} vs last month</span>`;
        } else if (rawP > 1) {
          trendStr = d > 0
            ? `<span style="color:var(--red);font-size:11px">▲ ${pct}% vs last month</span>`
            : `<span style="color:var(--green);font-size:11px">▼ ${pct}% vs last month</span>`;
        }
      }
    }

    guide.innerHTML +=
      `<div class="viol-card ${riskCls}">` +
        `<span class="badge ${badgeCls}" style="margin-bottom:6px;display:inline-block">` +
          `${v.risk === 'high' ? 'High risk' : v.risk === 'med' ? 'Medium risk' : 'Low risk'}` +
        `</span>` +
        `<div class="viol-name">${escapeHTML(v.short)}</div>` +
        `<div class="viol-desc">${escapeHTML(v.desc)}</div>` +
        (count > 0
          ? `<div style="margin-top:8px;font-size:11px;color:var(--accent)">Fleet total: ${count.toLocaleString()} events</div>`
          : '') +
        (trendStr ? `<div style="margin-top:2px">${trendStr}</div>` : '') +
      `</div>`;
  });
}


/* ═══════════════════════════════════════════════════════════════════════════
 * UTILITY HELPERS
 * Small shared helpers used throughout this file.
 * ═══════════════════════════════════════════════════════════════════════════ */

/**
 * setEl(id, text)
 * Sets the textContent of a DOM element by ID. Safe no-op if not found.
 */
function setEl(id, text) {
  const el = document.getElementById(id);
  if (el) el.textContent = text;
}

/**
 * setKpi(i, n, label, value, sub)
 * Updates a KPI card's label, value, and sub-text in one call.
 */
function setKpi(i, n, label, value, sub) {
  setEl(`kpiLbl${n}_${i}`, label);
  setEl(`kpiVal${n}_${i}`, value);
  setEl(`kpiSub${n}_${i}`, sub);
}

/**
 * setWrapHeight(i, name, height, maxHeight)
 * Sets the pixel height of a chart's wrap and scroll container.
 * Naming convention: wrapEl = `${name}Wrap_${i}`, scrollEl = `${name}Scroll_${i}`
 */
function setWrapHeight(i, name, height, maxHeight) {
  const wrap   = document.getElementById(`${name}Wrap_${i}`);
  const scroll = document.getElementById(`${name}Scroll_${i}`);
  if (wrap)   wrap.style.height      = `${height}px`;
  if (scroll) scroll.style.maxHeight = `${maxHeight}px`;
}

/**
 * collapseBarChart(chartInst, i, name, data, fields, barH, title, titleId)
 * Collapses a stacked horizontal bar chart to a subset of vehicles.
 * Used in both compare mode and single-vehicle mode.
 *
 * @param {Chart}    chartInst — Chart.js instance
 * @param {number}   i         — client index
 * @param {string}   name      — wrap/scroll ID prefix (e.g. 'wdwe', 'util')
 * @param {Object[]} data      — array of { name, [field1], [field2] }
 * @param {string[]} fields    — two property names matching the two datasets
 * @param {number}   barH      — px per bar row
 * @param {string}   title     — new card title text
 * @param {string}   titleId   — element ID for the title
 */
function collapseBarChart(chartInst, i, name, data, fields, barH, title, titleId) {
  if (!chartInst) return;
  const h = Math.max(120, data.length * barH);
  setWrapHeight(i, name, h, h);
  chartInst.data.labels                          = data.map(v => v.name);
  chartInst.data.datasets[0].data                = data.map(v => v[fields[0]] || 0);
  chartInst.data.datasets[0].barPercentage       = 0.35;
  chartInst.data.datasets[0].categoryPercentage  = 0.5;
  chartInst.data.datasets[1].data                = data.map(v => v[fields[1]] || 0);
  chartInst.data.datasets[1].barPercentage       = 0.35;
  chartInst.data.datasets[1].categoryPercentage  = 0.5;
  chartInst.update();
  setEl(titleId, title);
}

/**
 * restoreFleetKpis()
 * Restores all 7 KPI cards to their full-fleet values.
 * Called when switching back to fleet mode.
 */
function restoreFleetKpis(
  i, vehicles, totalDist, avgScore, avgIdle,
  vehiclesWithTrips, flaggedCount, activeViolations,
  hasPrev, prevTotalDist, prevAvgScore
) {
  setKpi(i, 0, 'Total vehicles',  vehicles.length,                        'in the fleet');
  setKpi(i, 1, 'Total distance',  Math.round(totalDist).toLocaleString(), 'km fleet total');
  setKpi(i, 2, 'Fleet avg score', Math.round(avgScore).toLocaleString(),  'lower is safer');
  setKpi(i, 3, 'Avg idle days',   avgIdle.toFixed(1),                     'per vehicle (weekdays only)');
  setKpi(i, 4, 'Vehicles active', vehiclesWithTrips,                      'recorded trips this month');
  setKpi(i, 5, 'At-risk vehicles',flaggedCount,                           `Moderate or High risk (score ${SCORE_BANDS.safe.max + 1}+)`);
  setKpi(i, 6, 'Violation types', activeViolations.length,                'detected in data');

  /* Reset KPI value overrides from vehicle/compare modes */
  [0, 5].forEach(n => {
    const val = document.getElementById(`kpiVal${n}_${i}`);
    if (val) { val.style.fontSize = ''; val.style.color = ''; }
  });
  document.getElementById(`kpiVal5_${i}`).style.color = 'var(--red)';

  /* Restore the MoM change tag */
  const chg1 = document.getElementById(`kpiChange1_${i}`);
  if (chg1) chg1.style.display = '';
}

/**
 * restoreAllChartsToFleet()
 * Restores all bar charts to their full-fleet datasets after
 * a vehicle/compare mode selection is cleared.
 */
function restoreAllChartsToFleet(
  i, vehicles, hasPrev, prevMap, activeViolations, violTotals,
  riskChartInst, violChartInst, wdweChartInst, utilChartInst,
  prevChartInst, distCompInst,
  allRiskVehicles, allWdweVehicles, allUtilVehicles, riskFullH
) {
  /* Risk chart */
  if (riskChartInst) {
    riskChartInst.data.labels                     = allRiskVehicles.map(v => v.name);
    riskChartInst.data.datasets[0].data           = allRiskVehicles.map(v => v.score || 0);
    riskChartInst.data.datasets[0].backgroundColor= allRiskVehicles.map(() => SCORE_BANDS.high.color);
    setWrapHeight(i, 'risk', riskFullH, 320);
    riskChartInst.update();
    riskChartInst.resize();
  }
  setEl(`riskTitle_${i}`, 'All vehicles — advanced score ranking');

  /* Violation donut */
  if (violChartInst) {
    violChartInst.data.datasets[0].data = violTotals;
    violChartInst.update();
  }
  setEl(`violTitle_${i}`, 'Violation breakdown — fleet total');

  /* Weekday/weekend */
  if (wdweChartInst) {
    const h = Math.max(320, allWdweVehicles.length * 32);
    setWrapHeight(i, 'wdwe', h, 320);
    wdweChartInst.data.labels                     = allWdweVehicles.map(v => v.name);
    wdweChartInst.data.datasets[0].data           = allWdweVehicles.map(v => v.weekdayDist || 0);
    wdweChartInst.data.datasets[0].backgroundColor= allWdweVehicles.map(() => '#4f8ef7');
    wdweChartInst.data.datasets[1].data           = allWdweVehicles.map(v => v.weekendDist || 0);
    wdweChartInst.data.datasets[1].backgroundColor= allWdweVehicles.map(() => '#2ec4b6');
    wdweChartInst.update();
  }
  setEl(`wdweTitle_${i}`, 'Weekday vs weekend distance');

  /* Utilisation */
  if (utilChartInst) {
    const h = Math.max(320, allUtilVehicles.length * 32);
    setWrapHeight(i, 'util', h, 320);
    utilChartInst.data.labels                     = allUtilVehicles.map(v => v.name);
    utilChartInst.data.datasets[0].data           = allUtilVehicles.map(v => v.daysActive || 0);
    utilChartInst.data.datasets[0].backgroundColor= allUtilVehicles.map(() => '#4f8ef7');
    utilChartInst.data.datasets[1].data           = allUtilVehicles.map(v => v.daysIdle || 0);
    utilChartInst.data.datasets[1].backgroundColor= allUtilVehicles.map(() => '#555b72');
    utilChartInst.update();
  }
  setEl(`utilTitle_${i}`, 'Most idle vehicles — days active vs idle');

  /* Month-on-month */
  if (prevChartInst && hasPrev) {
    const allPrev = vehicles.filter(v => prevMap[v.name]).sort((a, b) => (b.score || 0) - (a.score || 0));
    const h       = Math.max(400, allPrev.length * 36);
    setWrapHeight(i, 'prev', h, 400);
    prevChartInst.data.labels                     = allPrev.map(v => v.name);
    prevChartInst.data.datasets[0].data           = allPrev.map(v => v.score || 0);
    prevChartInst.data.datasets[0].backgroundColor= allPrev.map(() => '#3b6edc');
    prevChartInst.data.datasets[1].data           = allPrev.map(v => prevMap[v.name]?.score || 0);
    prevChartInst.update();
  }
  if (hasPrev) setEl(`prevTitle_${i}`, 'Month-on-month Advance score — current vs previous');

  /* Distance comparison */
  if (distCompInst && hasPrev) {
    const allNames = new Set([...vehicles.map(v => v.name), ...Object.keys(prevMap)]);
    const allDC    = [...allNames].map(name => {
      const curr = vehicles.find(v => v.name === name);
      return {
        name,
        currDist: curr?.totalDist || 0,
        prevDist: prevMap[name]?.totalDist || 0,
      };
    }).sort((a, b) => b.currDist - a.currDist);
    const h = Math.max(400, allDC.length * 34);
    setWrapHeight(i, 'distComp', h, 400);
    distCompInst.data.labels                      = allDC.map(v => v.name);
    distCompInst.data.datasets[0].data            = allDC.map(v => v.currDist);
    distCompInst.data.datasets[0].backgroundColor = allDC.map(v => v.currDist >= v.prevDist ? '#3db87a' : '#e05353');
    distCompInst.data.datasets[1].data            = allDC.map(v => v.prevDist);
    distCompInst.data.datasets[1].backgroundColor = allDC.map(v => v.currDist >= v.prevDist ? '#e05353' : '#3db87a');
    distCompInst.update();
    setEl(`distCompTitle_${i}`, 'Distance comparison — current vs previous month');
  }
}
