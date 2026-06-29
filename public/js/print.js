/**
 * print.js
 * ─────────────────────────────────────────────────────────────────────────────
 * Print and PDF export pipeline — settings dialog, chart pagination,
 * canvas height management, and post-print restoration.
 *
 * Contents:
 *   1. State & constants        — orientation, colour mode, guard flags
 *   2. Dialog open / close      — show/hide the print settings modal
 *   3. Settings selectors       — orientation and colour mode pickers
 *   4. Helpers                  — paint gate, preparation overlay
 *   5. Confirm & print          — the main async print pipeline
 *   6. Post-print restoration   — undo all DOM mutations after printing
 *   7. Public aliases           — downloadPDF(), printReport()
 *
 * ── Key engineering decisions ────────────────────────────────────────────────
 *
 *  A. DOUBLE-CLICK GUARD (isPrinting flag)
 *     Blocks re-entry while a print job is being prepared. Without this, every
 *     extra click spawns another full set of temporary Chart instances on top
 *     of the first, causing a GPU memory spike and duplicate chart pages.
 *
 *  B. SCOPED CHART INSTANCES
 *     Chart.instances holds every Chart created in the session, including charts
 *     from previously viewed clients. We filter to only canvases inside the
 *     active client's #dashContent subtree so stale instances from other clients
 *     are never paginated.
 *
 *  C. TEMP CHART CAP (MAX_TEMP_CHARTS)
 *     A 300-vehicle fleet across 5 bar charts (CHUNK=15) would create ~100
 *     temporary Chart instances synchronously, spiking CPU and GPU memory.
 *     We cap at MAX_TEMP_CHARTS and warn the user if the fleet is too large
 *     to fully paginate within the safe limit.
 *
 *  D. ASYNC RENDER GATE (waitForPaint)
 *     setTimeout(120) was a blind guess. On slow machines or large fleets the
 *     browser had not finished painting all canvases before window.print()
 *     fired, so charts appeared blank in the print preview. We now wait for
 *     two requestAnimationFrame ticks (one layout pass, one paint pass) per
 *     batch of charts, guaranteeing every canvas is drawn before the print
 *     dialog opens.
 *
 *  E. RISK CHART HEIGHT STRATEGY
 *     The print stylesheet applies height:auto to all canvases. For the risk
 *     chart (capped to top 10 vehicles) this would cause the canvas to expand
 *     to full-fleet size before the cap takes effect. We set an explicit inline
 *     px height directly on the canvas element — inline styles win over
 *     stylesheet rules — and re-enforce it in a beforeprint listener as a
 *     safety net.
 * ─────────────────────────────────────────────────────────────────────────────
 */


/* ═══════════════════════════════════════════════════════════════════════════
 * 1. STATE & CONSTANTS
 * ═══════════════════════════════════════════════════════════════════════════ */

let currentOrientation = 'portrait';
let currentColourMode  = 'colour';
let pendingPrintAction = null;

/**
 * isPrinting — re-entry guard (decision A).
 * Set to true when confirmPrintDialog() starts; released in restore().
 */
let isPrinting = false;

/**
 * MAX_TEMP_CHARTS — hard cap on temporary paginated chart instances (decision C).
 * Keeps memory usage predictable on large fleets.
 */
const MAX_TEMP_CHARTS = 40;


/* ═══════════════════════════════════════════════════════════════════════════
 * 2. DIALOG OPEN / CLOSE
 * ═══════════════════════════════════════════════════════════════════════════ */

/**
 * openPrintDialog(type, clientName)
 * Opens the print settings modal and configures its labels for either
 * print or PDF export mode.
 *
 * @param {string} type       — 'print' | 'pdf'
 * @param {string} clientName — used for PDF filename (currently reserved)
 */
function openPrintDialog(type, clientName) {
  /* Block if a print job is already in progress */
  if (isPrinting) {
    showToast('info', 'Print in progress', 'Please wait for the current print job to finish.', 3000);
    return;
  }

  pendingPrintAction = type || 'print';
  const isPDF = type === 'pdf';

  /* Update dialog labels depending on print vs PDF mode */
  const setText = (id, val) => {
    const el = document.getElementById(id);
    if (el) el.textContent = val;
  };

  setText('printDialogTitle',
    isPDF ? 'Export PDF' : 'Print / Save PDF'
  );
  setText('printSidebarLabel',
    isPDF ? 'Export PDF' : 'Print / Save PDF'
  );
  setText('printDialogMainTitle',
    isPDF ? 'Export Settings' : 'Print Settings'
  );
  setText('printDialogSub',
    isPDF
      ? 'Configure your document before exporting to PDF'
      : 'Configure your document before printing or saving as PDF'
  );
  setText('printDialogConfirmLabel',
    isPDF ? 'Download PDF' : 'Print / Save PDF'
  );

  /* Reset to defaults each time the dialog opens */
  selectOrient('portrait');
  selectColourMode('colour');
  const cms = document.getElementById('colourModeSelect');
  if (cms) cms.value = 'colour';

  document.getElementById('printDialogOverlay').classList.add('show');
}

/**
 * closePrintDialog()
 * Hides the print settings modal without printing.
 */
function closePrintDialog() {
  document.getElementById('printDialogOverlay').classList.remove('show');
  pendingPrintAction = null;
}

/* Close on overlay background click or Escape key */
document.addEventListener('click', e => {
  if (e.target.id === 'printDialogOverlay') closePrintDialog();
});
document.addEventListener('keydown', e => {
  if (e.key === 'Escape' && !isPrinting) closePrintDialog();
});


/* ═══════════════════════════════════════════════════════════════════════════
 * 3. SETTINGS SELECTORS
 * ═══════════════════════════════════════════════════════════════════════════ */

/**
 * selectOrient(o)
 * Updates the orientation state and the live preview thumbnail in the dialog.
 *
 * @param {string} o — 'portrait' | 'landscape'
 */
function selectOrient(o) {
  currentOrientation = o;

  /* Sync the select element if present */
  const sel = document.getElementById('orientSelect');
  if (sel) sel.value = o;

  const page = document.getElementById('previewPage');
  const dims = document.getElementById('previewDims');
  const sub  = document.getElementById('orientSub');

  if (o === 'landscape') {
    if (page) { page.style.width = '96px'; page.style.height = '66px'; }
    if (dims) dims.innerHTML  = 'A4 Landscape<br>29.7 x 21 cm';
    if (sub)  sub.textContent = 'Landscape — A4 horizontal (29.7 x 21 cm)';
  } else {
    if (page) { page.style.width = '66px'; page.style.height = '88px'; }
    if (dims) dims.innerHTML  = 'A4 Portrait<br>21 x 29.7 cm';
    if (sub)  sub.textContent = 'Portrait — A4 vertical (21 x 29.7 cm)';
  }
}

/**
 * selectColourMode(m)
 * Switches between full-colour and black-and-white print output.
 * The preview thumbnail updates instantly so the user can see the effect.
 *
 * @param {string} m — 'colour' | 'bw'
 */
function selectColourMode(m) {
  currentColourMode = m;

  const sub   = document.getElementById('colourModeSub');
  const lines = document.getElementById('previewPage')
                       ?.querySelector('.print-preview-lines');

  if (sub) {
    sub.textContent = m === 'bw'
      ? 'Black & white — grayscale output, ink-friendly'
      : 'Full colour — badges and charts in colour';
  }

  /* Grayscale the preview lines to mimic B&W output */
  if (lines) lines.style.filter = m === 'bw' ? 'grayscale(1)' : 'none';
}


/* ═══════════════════════════════════════════════════════════════════════════
 * 4. HELPERS
 * ═══════════════════════════════════════════════════════════════════════════ */

/**
 * waitForPaint()
 * Yields to the browser for one full layout + paint cycle.
 * Uses two nested requestAnimationFrame calls:
 *   - First rAF: browser processes pending layout (reflow)
 *   - Second rAF: browser completes the paint pass
 *
 * This is the async render gate (decision D). Called after each batch of
 * temporary charts is created to ensure canvases are drawn before the next
 * batch starts or the print dialog opens.
 *
 * @returns {Promise<void>}
 */
function waitForPaint() {
  return new Promise(resolve =>
    requestAnimationFrame(() => requestAnimationFrame(resolve))
  );
}

/**
 * setPrintOverlay(visible, msg)
 * Shows or hides a full-screen "Preparing…" overlay during the async print
 * pipeline. Created lazily on first call and reused on subsequent calls.
 *
 * @param {boolean} visible — true to show, false to hide
 * @param {string}  msg     — status message shown to the user
 */
function setPrintOverlay(visible, msg) {
  let overlay = document.getElementById('printPrepOverlay');

  /* Create the overlay element once and cache it in the DOM */
  if (!overlay) {
    overlay    = document.createElement('div');
    overlay.id = 'printPrepOverlay';
    overlay.style.cssText = [
      'position:fixed', 'inset:0', 'background:rgba(0,0,0,.72)',
      'z-index:9998', 'display:flex', 'flex-direction:column',
      'align-items:center', 'justify-content:center', 'gap:14px',
    ].join(';');
    overlay.innerHTML = `
      <div style="
        width:36px; height:36px;
        border:3px solid rgba(255,255,255,.15);
        border-top-color:#4f8ef7;
        border-radius:50%;
        animation:spin .7s linear infinite
      "></div>
      <div id="printPrepMsg" style="
        color:#e8eaf0;
        font-size:13px;
        font-family:'Segoe UI',system-ui,sans-serif
      "></div>`;
    document.body.appendChild(overlay);
  }

  const msgEl = overlay.querySelector('#printPrepMsg');
  if (msgEl) msgEl.textContent = msg || '';

  overlay.style.display = visible ? 'flex' : 'none';
}


/* ═══════════════════════════════════════════════════════════════════════════
 * 5. CONFIRM & PRINT  (main async pipeline)
 * ═══════════════════════════════════════════════════════════════════════════ */

/**
 * confirmPrintDialog()
 * The main print pipeline. Called when the user clicks "Print / Save PDF"
 * in the settings dialog.
 *
 * Pipeline steps:
 *   1. Apply @page orientation via injected <style>
 *   2. Apply colour mode class to <body>
 *   3. Unlock scrollable chart containers so all content is visible
 *   4. Paginate horizontal bar charts into printable chunks
 *   5. Cap the risk chart to top N vehicles
 *   6. Optionally truncate table rows for very large fleets
 *   7. Wait for all canvases to paint (decision D)
 *   8. Open window.print()
 *   9. Restore all DOM mutations after the user closes the print dialog
 */
async function confirmPrintDialog() {
  /* Decision A: block re-entry */
  if (isPrinting) return;
  isPrinting = true;

  const orient  = currentOrientation;
  const colMode = currentColourMode;

  closePrintDialog();
  setPrintOverlay(true, 'Preparing your document…');
  await waitForPaint();

  /* ── Step 1: inject @page orientation rule ── */
  let styleEl = document.getElementById('orientStyle');
  if (!styleEl) {
    styleEl    = document.createElement('style');
    styleEl.id = 'orientStyle';
    document.head.appendChild(styleEl);
  }
  styleEl.textContent = `@media print { @page { size: A4 ${orient}; margin: ${
    orient === 'landscape' ? '10mm 14mm' : '14mm 12mm'
  }; } }`;

  /* ── Step 2: apply colour mode ── */
  document.body.classList.toggle('print-bw', colMode === 'bw');

  /* ── Step 3: unlock scroll containers ── */
  /* Save current styles so we can restore them after printing */
  const containers      = document.querySelectorAll('[id*="Scroll_"],[id*="Wrap_"]');
  const savedContainers = [];

  containers.forEach(el => {
    savedContainers.push({
      el,
      overflow : el.style.overflow,
      maxH     : el.style.maxHeight,
      height   : el.style.height,
    });

    el.style.overflow  = 'visible';
    el.style.maxHeight = 'none';

    /*
     * Decision E: do NOT set height:auto on riskWrap elements.
     * The risk chart cap (step 4) sets an explicit pixel height after slicing
     * to top N bars. Setting auto here first causes the canvas to expand to
     * full-fleet height before the cap can clamp it, overflowing the card.
     */
    const isRiskWrap = el.id && /riskWrap_/i.test(el.id);
    if (el.id && el.id.includes('Wrap_') && !isRiskWrap) {
      el.style.height = 'auto';
    }
  });

  /* ── Step 4 & 5: chart pagination and risk chart cap ── */

  /* Saved original chart data so we can restore it after printing */
  const chartSaved = [];
  /* Temporary Chart.js instances created for pagination pages */
  const tempCharts = [];
  /* Cards whose original chart is hidden while paginated versions display */
  const hiddenCards = [];

  /* Container for paginated print pages — appended to <body>, hidden on screen */
  const printPages    = document.createElement('div');
  printPages.id       = 'fleetPrintPages';
  printPages.style.cssText = 'display:none;font-family:Segoe UI,system-ui,sans-serif';
  document.body.appendChild(printPages);

  /*
   * Decision B: only look at charts inside the active client's DOM subtree.
   * Chart.instances includes every chart ever created in the session.
   * Filtering by dashContent prevents stale charts from other client tabs
   * from being paginated.
   */
  const dashContent = document.getElementById('dashContent');

  try {
    /* Collect only horizontal bar charts belonging to this client */
    const allCharts = Object.values(Chart.instances).filter(chart => {
      if (!chart?.canvas || !chart?.config)      return false;
      if (chart.config.type !== 'bar')           return false;
      if (chart.options?.indexAxis !== 'y')      return false;
      if (!chart.data.labels?.length)            return false;
      return dashContent ? dashContent.contains(chart.canvas) : true;
    });

    /* Determine fleet size to pick appropriate chunk and bar height values */
    const maxBars      = allCharts.reduce(
      (m, c) => Math.max(m, c.data?.labels?.length || 0), 0
    );
    const isLargeFleet = maxBars >= 80;
    const isSmallFleet = maxBars > 0 && maxBars <= 20;

    /*
     * CHUNK: how many vehicles fit per page.
     * BAR_H: pixel height per bar row (used to size the canvas).
     * topBars: how many bars to show in the risk chart.
     * rowLimit: maximum table rows for extremely large fleets.
     */
    const CHUNK = orient === 'landscape'
      ? (isLargeFleet ? 24 : isSmallFleet ? 14 : 20)
      : (isLargeFleet ? 18 : isSmallFleet ? 12 : 15);

    const BAR_H     = isLargeFleet ? 24 : isSmallFleet ? 30 : 27;
    const topBars   = isLargeFleet ? 8 : 10;
    const rowLimit  = maxBars >= 140
      ? (orient === 'landscape' ? 90 : 70)
      : Number.POSITIVE_INFINITY;

    /* Decision C: cap is relaxed slightly for large fleets since CHUNK is smaller */
    const tempChartCap = isLargeFleet ? 60 : MAX_TEMP_CHARTS;

    /* Running count of temporary charts created (decision C) */
    let tempCount = 0;

    for (const chart of allCharts) {
      const canvas = chart.canvas;
      const labels = chart.data.labels;
      const isRisk = /risk/i.test(canvas.id || '');

      /* Walk up the DOM to find the containing .card element */
      let card = canvas.parentElement;
      while (card && !card.classList.contains('card')) card = card.parentElement;

      const sectionTitle = card
        ?.querySelector('.card-title')
        ?.textContent.trim().toUpperCase() || '';

      /* ── Risk chart: cap to top N, no pagination ── */
      if (isRisk) {
        if (labels.length > topBars) {
          /* Save original data for restoration */
          const entry = {
            chart,
            labels  : labels.slice(),
            datasets: chart.data.datasets.map(ds => ({
              data : (ds.data || []).slice(),
              bg   : Array.isArray(ds.backgroundColor)
                      ? ds.backgroundColor.slice()
                      : ds.backgroundColor,
            })),
          };

          /* Find the corresponding wrap and scroll elements by ID convention */
          const wrapEl = document.getElementById(
            canvas.id
              .replace(/Chart_(\d+)/, 'Wrap_$1')
              .replace(/Chart$/, 'Wrap')
          );
          if (wrapEl) { entry.wrapEl = wrapEl; entry.wrapH = wrapEl.style.height; }
          chartSaved.push(entry);

          /* Slice data to top N vehicles */
          chart.data.labels = labels.slice(0, topBars);
          chart.data.datasets.forEach(ds => {
            ds.data = (ds.data || []).slice(0, topBars);
            if (Array.isArray(ds.backgroundColor)) {
              ds.backgroundColor = ds.backgroundColor.slice(0, topBars);
            }
          });

          /* Calculate the exact pixel height for the top-N block */
          const top10H = topBars * BAR_H + 42;   /* bars + axis padding */

          /*
           * Decision E: set explicit heights on wrap, scroll, and canvas.
           * Inline styles win over the print stylesheet's height:auto rule,
           * which is what prevents the chart from expanding to full-fleet size.
           */
          if (wrapEl) {
            wrapEl.style.height    = `${top10H}px`;
            wrapEl.style.maxHeight = `${top10H}px`;
            wrapEl.style.overflow  = 'hidden';
          }

          const scrollEl = wrapEl?.parentElement;
          if (scrollEl && /Scroll_/i.test(scrollEl.id || '')) {
            scrollEl.style.height    = `${top10H}px`;
            scrollEl.style.maxHeight = `${top10H}px`;
            scrollEl.style.overflow  = 'hidden';
          }

          /* Canvas itself must also have an explicit height (decision E) */
          canvas.style.height    = `${top10H}px`;
          canvas.style.maxHeight = `${top10H}px`;

          chart.update('none');
          chart.resize(canvas.offsetWidth || 600, top10H);
        }
        continue;   /* risk chart handled — skip to next chart */
      }

      /* ── Other bar charts: paginate into chunks ── */
      const pagesNeeded = Math.ceil(labels.length / CHUNK);

      /* Decision C: check cap before spawning temp charts for this chart */
      if (tempCount + pagesNeeded > tempChartCap) {
        showToast(
          'warn',
          'Large fleet — print truncated',
          `This chart has ${labels.length} vehicles. Only the first ` +
          `${CHUNK * (tempChartCap - tempCount)} fit within the safe print limit. ` +
          `Use Export & Send for a full report.`,
          0
        );
        /* Hide the original card even though we're truncating */
        if (card) {
          hiddenCards.push({ el: card, disp: card.style.display });
          card.style.display = 'none';
        }
        continue;
      }

      /* Hide the original scrollable card — the paginated pages replace it */
      if (card) {
        hiddenCards.push({ el: card, disp: card.style.display });
        card.style.display = 'none';
      }

      const isStacked = !!(chart.options.scales?.x?.stacked);

      /* Create one page div per chunk */
      for (let p = 0; p < pagesNeeded; p++) {
        const sliceStart  = p * CHUNK;
        const sliceEnd    = Math.min((p + 1) * CHUNK, labels.length);
        const sliceLabels = labels.slice(sliceStart, sliceEnd);

        /* Page wrapper — force page break after each chunk except the last */
        const pageDiv = document.createElement('div');
        pageDiv.style.cssText =
          'page-break-inside:avoid;break-inside:avoid;padding:0' +
          (p < pagesNeeded - 1 ? ';page-break-after:always;break-after:page' : '');

        /* Section heading shows the chart title and page count */
        const hdr = document.createElement('div');
        hdr.style.cssText =
          'font-size:10px;font-weight:700;color:#374151;text-transform:uppercase;' +
          'letter-spacing:.06em;margin-bottom:10px;padding-bottom:7px;' +
          'border-bottom:2px solid #e5e7eb';
        hdr.textContent = sectionTitle +
          (pagesNeeded > 1 ? ` — Page ${p + 1} of ${pagesNeeded}` : '');
        pageDiv.appendChild(hdr);

        /* Canvas wrapper — exact pixel height prevents float/overflow issues */
        const wrapDiv = document.createElement('div');
        wrapDiv.style.cssText =
          `position:relative;height:${Math.max(150, sliceLabels.length * BAR_H + 16)}px;width:100%`;

        const cv = document.createElement('canvas');
        wrapDiv.appendChild(cv);
        pageDiv.appendChild(wrapDiv);
        printPages.appendChild(pageDiv);

        /* Create a clean Chart.js instance for this page slice */
        tempCharts.push(new Chart(cv, {
          type: 'bar',
          data: {
            labels: sliceLabels,
            datasets: chart.data.datasets.map(ds => ({
              label              : ds.label || '',
              data               : (ds.data || []).slice(sliceStart, sliceEnd),
              backgroundColor    : Array.isArray(ds.backgroundColor)
                                   ? ds.backgroundColor.slice(sliceStart, sliceEnd)
                                   : ds.backgroundColor,
              borderColor        : ds.borderColor,
              borderWidth        : ds.borderWidth        || 1,
              borderRadius       : ds.borderRadius       || 0,
              barPercentage      : ds.barPercentage,
              categoryPercentage : ds.categoryPercentage,
            })),
          },
          options: {
            indexAxis           : 'y',
            responsive          : true,
            maintainAspectRatio : false,
            animation           : false,   /* synchronous draw — no animation overhead in print */
            plugins: {
              legend: { display: false },
            },
            scales: {
              x: {
                stacked : isStacked,
                ticks   : {
                  color    : '#6b7280',
                  font     : { size: 9 },
                  callback : v => typeof v === 'number' ? v.toLocaleString() : v,
                },
                grid   : { color: 'rgba(0,0,0,0.06)' },
                border : { display: false },
              },
              y: {
                stacked : isStacked,
                ticks   : { color: '#374151', font: { size: 9 } },
                border  : { display: false },
              },
            },
          },
        }));

        tempCount++;
      }

      /* Decision D: yield to the browser after each chart so it can paint
         before spinning up the next one — prevents a single blocking burst */
      setPrintOverlay(
        true,
        `Rendering charts… (${tempCount} of ${Math.min(tempCount + pagesNeeded, tempChartCap)})`
      );
      await waitForPaint();
    }

    /* ── Step 6: optionally truncate table rows for very large fleets ── */
    const savedRows = [];
    document.querySelectorAll('table tbody').forEach(tbody => {
      Array.from(tbody.querySelectorAll('tr')).forEach((tr, idx) => {
        if (idx >= rowLimit) {
          savedRows.push({ tr, disp: tr.style.display });
          tr.style.display = 'none';
        }
      });
    });

    /* ── Step 7: final paint gate — all canvases must be drawn ── */
    setPrintOverlay(true, 'Opening print dialog…');
    await waitForPaint();
    setPrintOverlay(false);

    /* ── Step 8: beforeprint safety net ──
     * The browser's print layout engine runs after window.print() is called,
     * which can re-trigger CSS height:auto on our explicitly-sized canvases.
     * This listener re-enforces the inline px heights as a safety net.
     */
    const _beforePrint = () => {
      chartSaved.forEach(cs => {
        if (!cs.chart?.canvas) return;

        const cv     = cs.chart.canvas;
        const top10H = topBars * BAR_H + 42;

        cv.style.setProperty('height',     `${top10H}px`, 'important');
        cv.style.setProperty('max-height', `${top10H}px`, 'important');

        if (cs.wrapEl) {
          cs.wrapEl.style.setProperty('height',     `${top10H}px`, 'important');
          cs.wrapEl.style.setProperty('max-height', `${top10H}px`, 'important');
          cs.wrapEl.style.setProperty('overflow',   'hidden',      'important');
        }

        const sc = cs.wrapEl?.parentElement;
        if (sc && /Scroll_/i.test(sc.id || '')) {
          sc.style.setProperty('height',     `${top10H}px`, 'important');
          sc.style.setProperty('max-height', `${top10H}px`, 'important');
          sc.style.setProperty('overflow',   'hidden',      'important');
        }
      });
    };

    window.addEventListener('beforeprint', _beforePrint);

    /* ── Step 8: open native print dialog ── */
    window.print();

    /* ── Step 9: restore everything after the print dialog closes ──
     * window.print() is synchronous — it blocks until the user closes the
     * dialog. We wait 2.5s to give the browser time to complete the print
     * job before tearing down the temporary DOM structures. */
    window.removeEventListener('beforeprint', _beforePrint);
    setTimeout(restore, 2500);

    /* ── Restoration function ── */
    function restore() {
      isPrinting = false;                              /* decision A: release guard */
      document.body.classList.remove('print-bw');

      /* Restore scroll container styles */
      savedContainers.forEach(s => {
        s.el.style.overflow  = s.overflow || '';
        s.el.style.maxHeight = s.maxH     || '';
        if (s.el.id?.includes('Wrap_')) s.el.style.height = s.height || '';
      });

      /* Restore chart data and dimensions */
      chartSaved.forEach(cs => {
        cs.chart.data.labels = cs.labels;
        cs.chart.data.datasets.forEach((ds, i) => {
          ds.data            = cs.datasets[i].data;
          ds.backgroundColor = cs.datasets[i].bg;
        });

        if (cs.wrapEl) {
          cs.wrapEl.style.height    = cs.wrapH || '';
          cs.wrapEl.style.maxHeight = '';
          cs.wrapEl.style.overflow  = '';

          /* Restore the scroll container */
          const sc = cs.wrapEl.parentElement;
          if (sc && /Scroll_/i.test(sc.id || '')) {
            sc.style.height    = '';
            sc.style.maxHeight = '320px';
            sc.style.overflow  = 'auto';
          }

          /* Remove inline dimensions set during print */
          if (cs.chart.canvas) {
            cs.chart.canvas.style.height    = '';
            cs.chart.canvas.style.maxHeight = '';
          }
        }

        cs.chart.update('none');
        cs.chart.resize();
      });

      /* Destroy temporary pagination charts */
      tempCharts.forEach(c => { try { c.destroy(); } catch (e) { /* already destroyed */ } });

      /* Remove the paginated print pages container */
      document.getElementById('fleetPrintPages')?.remove();

      /* Un-hide original chart cards */
      hiddenCards.forEach(h => { h.el.style.display = h.disp || ''; });

      /* Un-hide truncated table rows */
      savedRows.forEach(s => { s.tr.style.display = s.disp || ''; });

      setPrintOverlay(false);
    }

  } catch (err) {
    /* If anything goes wrong, always release the guard and hide the overlay
       so the user is not left staring at a frozen "Preparing…" screen */
    console.warn('Print pipeline error:', err);
    isPrinting = false;
    setPrintOverlay(false);
    showToast('error', 'Print failed', err.message || 'An unexpected error occurred.', 0);
  }
}


/* ═══════════════════════════════════════════════════════════════════════════
 * 6. POST-PRINT RESTORATION
 * (restore() is defined inline inside confirmPrintDialog above
 *  so it closes over the saved state arrays — this is intentional)
 * ═══════════════════════════════════════════════════════════════════════════ */


/* ═══════════════════════════════════════════════════════════════════════════
 * 7. PUBLIC ALIASES
 * ═══════════════════════════════════════════════════════════════════════════ */

/**
 * downloadPDF(clientName)
 * Opens the print dialog in PDF export mode.
 * Called from the "Download PDF" button in the report header.
 *
 * @param {string} clientName — used for the title label in the dialog
 */
function downloadPDF(clientName) {
  openPrintDialog('print', clientName);
}

/**
 * printReport()
 * Opens the print dialog in standard print mode.
 * Called from the "Print" button in the topbar.
 */
function printReport() {
  openPrintDialog('print', '');
}
