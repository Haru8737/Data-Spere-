/**
 * dashboard.js
 * ─────────────────────────────────────────────────────────────────────────────
 * Application entry point.
 * Owns the top-level state arrays and the shared DashState object.
 * All other logic lives in dashboard.bundle.js.
 *
 * Load order in index.html:
 *   1. constants.js          ← SCORE_BANDS, VIOLATIONS, getScoreBand(), isAtRisk()
 *   2. utils.js              ← setEl(), changeTag(), fmt(), fmtPct(), etc.
 *   3. dashboard.js          ← this file (globals + upload zone wiring)
 *   4. dashboard.bundle.js   ← all rendering, parsing, chart logic
 * ─────────────────────────────────────────────────────────────────────────────
 */


/* ── Top-level app state ─────────────────────────────────────────────────── */

/** All loaded client objects. Each entry: { name, sheets, month, warnings } */
let clients = [];

/** Index of the currently visible client tab */
let activeClient = 0;

/** Files queued on the upload screen, waiting for "Go" */
let pendingFiles = [];


/* ── Shared Dashboard State ──────────────────────────────────────────────────
 *
 * Single source of truth for vehicle table UI state.
 * Modules read and write through DashState — no hidden closure state.
 *
 * Rules:
 *   - Never re-declare these as separate `let` variables inside functions
 *   - Always call DashState.reset() when switching clients
 *   - Add new cross-module state here, not inside closures
 *
 * Phase 1A complete:  vehicleTable state centralised here
 * Phase 1B pending:   compareSelected, activeDateFilter, riskSelected etc.
 *                     still use closure getter/setters — migrates in Phase 3
 *                     when the getter/setter pattern is refactored out
 * ─────────────────────────────────────────────────────────────────────────── */

/* ── Upload zone wiring ──────────────────────────────────────────────────────
 *
 * Wires the file input and drag-and-drop zone to handleFiles().
 * Must run after the DOM is ready; handleFiles() is defined in the bundle
 * which loads after this file, so calling it from DOMContentLoaded is safe.
 * ─────────────────────────────────────────────────────────────────────────── */

/* ── runDashboard() ──────────────────────────────────────────────────────────
 *
 * Called by the "Go" button on the upload screen.
 * Parses all pending Excel files, populates clients[], then renders the
 * dashboard. This is the correct entry point — NOT buildDashboard(), which
 * is the standalone export-report renderer.
 * ─────────────────────────────────────────────────────────────────────────── */

async function runDashboard() {
  if (pendingFiles.length === 0) return;

  const goBtn = document.getElementById('goBtn');
  goBtn.disabled   = true;
  goBtn.textContent = 'Loading…';

  try {
    clients      = [];
    activeClient = 0;

    for (const p of pendingFiles) {
      const sheets = await parseExcel(p.file);
      clients.push({
        name    : p.client || p.file.name.replace(/\.[^.]+$/, ''),
        sheets,
        month   : p.month  || '',
        warnings: [],
      });
    }

    /* Switch screens */
    document.getElementById('uploadScreen').style.display = 'none';
    document.getElementById('dashboard').style.display    = 'block';

    renderTabs();
    renderClient(0);

  } catch (err) {
    showToast('error', 'Failed to load file', err.message, 0);
    goBtn.disabled    = false;
    goBtn.textContent = 'Go';
  }
}

document.addEventListener('DOMContentLoaded', function () {
  const fileInput = document.getElementById('fileInput');
  const dropZone  = document.getElementById('dropZone');

  if (fileInput) {
    fileInput.addEventListener('change', function () {
      handleFiles(Array.from(this.files));
      /* Reset so the same file can be re-selected after removal */
      this.value = '';
    });
  }

  if (dropZone) {
    dropZone.addEventListener('dragover', function (e) {
      e.preventDefault();
      dropZone.classList.add('over');
    });

    dropZone.addEventListener('dragleave', function () {
      dropZone.classList.remove('over');
    });

    dropZone.addEventListener('drop', function (e) {
      e.preventDefault();
      dropZone.classList.remove('over');
      handleFiles(Array.from(e.dataTransfer.files));
    });
  }
});


const DashState = {

  /* ── Vehicle scoring table ── */
  vehicleTableMode     : 'flagged',  /* 'flagged' | 'all'             */
  vehicleTableContext  : null,       /* current render override        */
  vehicleSelectionCtx  : null,       /* sticky — survives mode toggle  */
  vehicleSelectionLabel: '',         /* title to restore on Flagged    */

  /* ── Reset ────────────────────────────────────────────────────────────────
   * Call when switching clients or loading a new file.
   * Resets all UI state to defaults without touching clients/activeClient.
   * ──────────────────────────────────────────────────────────────────────── */
  reset() {
    this.vehicleTableMode      = 'flagged';
    this.vehicleTableContext   = null;
    this.vehicleSelectionCtx   = null;
    this.vehicleSelectionLabel = '';
  },
};
