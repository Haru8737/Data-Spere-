/**
 * ui.js
 * ─────────────────────────────────────────────────────────────────────────────
 * UI helper functions used across the entire dashboard.
 * Nothing in here touches vehicle data or charts directly.
 *
 * Contents:
 *   1. Security helpers     — safe text escaping to prevent XSS
 *   2. Toast notifications  — success / warn / error / info messages
 *   3. Loading overlay      — full-screen spinner shown during file parsing
 *   4. Data validation      — sheet-level and vehicle-level sanity checks
 *   5. Client tabs          — tab bar rendering and switching
 *   6. Screen navigation    — moving between upload screen and dashboard
 * ─────────────────────────────────────────────────────────────────────────────
 */


/* ═══════════════════════════════════════════════════════════════════════════
 * 1. SECURITY HELPERS
 * ═══════════════════════════════════════════════════════════════════════════ */

/**
 * escapeHTML(str)
 * Escapes a string so it is safe to inject into innerHTML.
 * Always use this when inserting user-supplied text (client names, notes,
 * summary content) into the DOM via innerHTML.
 *
 * Without this, a client name containing <script> tags or event handlers
 * could execute arbitrary JavaScript in the exported report.
 *
 * @param  {*}      str — any value (will be coerced to string)
 * @returns {string}    — HTML-safe string
 */
function escapeHTML(str) {
  return String(str ?? '')
    .replace(/&/g,  '&amp;')
    .replace(/</g,  '&lt;')
    .replace(/>/g,  '&gt;')
    .replace(/"/g,  '&quot;')
    .replace(/'/g,  '&#39;');
}

/**
 * sanitiseSummaryHTML(html)
 * Strips dangerous tags and attributes from the executive summary HTML
 * before it is injected into the exported report.
 *
 * Allows safe formatting tags (b, i, em, strong, br, p, ul, li, span)
 * but removes script tags, event handler attributes (onclick, onerror…),
 * and javascript: href values.
 *
 * @param  {string} html — raw innerHTML from the summary editor
 * @returns {string}     — safe HTML
 */
function sanitiseSummaryHTML(html) {
  /* Use a detached div so the browser parses the HTML for us,
     then walk the tree removing anything dangerous */
  const div = document.createElement('div');
  div.innerHTML = html;

  const ALLOWED_TAGS = new Set([
    'b', 'i', 'em', 'strong', 'br', 'p',
    'ul', 'ol', 'li', 'span', 'div',
  ]);

  function clean(node) {
    const toRemove = [];
    node.childNodes.forEach(child => {
      if (child.nodeType === Node.ELEMENT_NODE) {
        const tag = child.tagName.toLowerCase();
        if (!ALLOWED_TAGS.has(tag)) {
          /* Replace disallowed element with its text content */
          toRemove.push({ node: child, replacement: document.createTextNode(child.textContent) });
          return;
        }
        /* Strip all event-handler attributes and javascript: hrefs */
        Array.from(child.attributes).forEach(attr => {
          if (
            attr.name.startsWith('on') ||
            (attr.name === 'href' && attr.value.trim().toLowerCase().startsWith('javascript:'))
          ) {
            child.removeAttribute(attr.name);
          }
        });
        clean(child);
      }
    });
    toRemove.forEach(({ node: n, replacement }) => n.parentNode.replaceChild(replacement, n));
  }

  clean(div);
  return div.innerHTML;
}


/* ═══════════════════════════════════════════════════════════════════════════
 * 2. TOAST NOTIFICATIONS
 * ═══════════════════════════════════════════════════════════════════════════ */

/**
 * showToast(type, title, msg, duration)
 * Displays a non-blocking notification in the bottom-right corner.
 *
 * @param {string} type     — 'success' | 'warn' | 'error' | 'info'
 * @param {string} title    — bold heading line
 * @param {string} msg      — optional detail text (pass '' to omit)
 * @param {number} duration — ms before auto-dismiss (0 = stays until closed)
 */
function showToast(type, title, msg, duration) {
  if (duration === undefined) duration = 6000;

  const icons = {
    error  : '✕',
    warn   : '⚠',
    success: '✓',
    info   : 'i',
  };

  const box       = document.createElement('div');
  box.className   = `toast t-${type}`;

  /* Use escapeHTML on all user-supplied text to prevent XSS via
     error messages that might contain raw Excel cell content */
  box.innerHTML =
    `<span class="t-icon" aria-hidden="true">${icons[type] ?? 'i'}</span>` +
    `<div class="t-body">` +
      `<div class="t-title">${escapeHTML(title)}</div>` +
      (msg ? `<div class="t-msg">${escapeHTML(msg)}</div>` : '') +
    `</div>` +
    `<button class="t-close" aria-label="Dismiss notification" onclick="this.parentNode.remove()">×</button>`;

  document.getElementById('toastContainer').appendChild(box);

  /* Auto-dismiss after duration ms, with a fade-out animation */
  if (duration > 0) {
    setTimeout(() => {
      box.style.animation = 'toastOut .25s ease forwards';
      setTimeout(() => box.remove(), 250);
    }, duration);
  }
}


/* ═══════════════════════════════════════════════════════════════════════════
 * 3. LOADING OVERLAY
 * ═══════════════════════════════════════════════════════════════════════════ */

/**
 * showLoading(msg, sub)
 * Shows a full-screen loading overlay with a message and optional sub-text.
 * Used during Excel parsing which can take a few seconds on large files.
 *
 * @param {string} msg — primary loading message e.g. 'Reading file...'
 * @param {string} sub — secondary line e.g. the filename being processed
 */
function showLoading(msg, sub) {
  document.getElementById('loadMsg').textContent = msg || 'Processing...';
  document.getElementById('loadSub').textContent = sub || '';
  document.getElementById('loadOverlay').classList.add('show');
}

/**
 * hideLoading()
 * Hides the loading overlay after parsing completes.
 */
function hideLoading() {
  document.getElementById('loadOverlay').classList.remove('show');
}


/* ═══════════════════════════════════════════════════════════════════════════
 * 4. DATA VALIDATION
 * ═══════════════════════════════════════════════════════════════════════════ */

/**
 * validateSheets(sheets, fileName)
 * Checks that an uploaded workbook has the expected sheet structure.
 * Throws if the file is completely unusable; returns warnings for
 * recoverable issues that the user should know about.
 *
 * @param  {Object} sheets   — result of parseExcel()
 * @param  {string} fileName — used in error/warning messages
 * @returns {string[]}       — array of warning strings (may be empty)
 * @throws {Error}           — if the file cannot produce a dashboard at all
 */
function validateSheets(sheets, fileName) {
  const warns = [];
  const names = Object.keys(sheets);

  /* Fatal: empty workbook */
  if (names.length === 0) {
    throw new Error(
      `No sheets found in "${fileName}". The file may be empty or corrupted.`
    );
  }

  /* Warn: expected sheet names not found */
  const hasScoring = names.some(n => /scor/i.test(n) && !/prev/i.test(n));
  const hasUtil    = names.some(n => /util/i.test(n));

  if (!hasScoring) {
    warns.push(
      'No "Scoring" sheet found — using first sheet instead. ' +
      'Column detection may be limited.'
    );
  }
  if (!hasUtil) {
    warns.push(
      'No "Utilization" sheet found — daily distance chart and ' +
      'idle stats will not be available.'
    );
  }

  /* Fatal: scoring sheet is completely empty */
  const sName = names.find(n => /scor/i.test(n) && !/prev/i.test(n)) || names[0];
  const sRows = sheets[sName] || [];

  if (sRows.length === 0) {
    throw new Error(
      `The scoring sheet in "${fileName}" is empty. Please check your Excel file.`
    );
  }

  /* Warn: critical columns missing (uses parser.js detectCriticalColWarnings) */
  const sCols      = detectCols(sRows);
  const colWarns   = detectCriticalColWarnings(sCols, fileName);
  colWarns.forEach(w => warns.push(w));

  return warns;
}

/**
 * validateVehicleData(vehicles)
 * Sanity-checks the processed vehicle array for values that suggest
 * data errors in the Excel file. Returns warnings shown as toasts.
 *
 * Checks:
 *   - Negative scores (impossible — score is a non-negative penalty)
 *   - Implausibly large distances (> 50,000 km in one month)
 *   - More than half the fleet showing 0 km (likely a parsing issue)
 *   - Active or idle days exceeding 31 (impossible for a monthly report)
 *
 * @param  {Object[]} vehicles — array of vehicle objects from buildVehicleMap()
 * @returns {string[]}         — array of warning strings (may be empty)
 */
function validateVehicleData(vehicles) {
  const warns   = [];
  const neg     = vehicles.filter(v => (v.score     || 0) < 0);
  const huge    = vehicles.filter(v => (v.totalDist || 0) > 50000);
  const zero    = vehicles.filter(v => (v.totalDist || 0) === 0);
  const badDays = vehicles.filter(v =>
    (v.daysIdle   || 0) > 31 ||
    (v.daysActive || 0) > 31
  );

  if (neg.length) {
    warns.push(
      `${neg.length} vehicle(s) have a negative score — possible data error.`
    );
  }
  if (huge.length) {
    warns.push(
      `${huge.length} vehicle(s) show over 50,000 km in one month — please verify.`
    );
  }
  if (zero.length > vehicles.length * 0.5) {
    warns.push(
      `${zero.length} of ${vehicles.length} vehicles show 0 km distance. ` +
      `Over half the fleet — possible data issue.`
    );
  }
  if (badDays.length) {
    warns.push(
      `${badDays.length} vehicle(s) show more than 31 active or idle days — ` +
      `impossible for a monthly report.`
    );
  }

  return warns;
}


/* ═══════════════════════════════════════════════════════════════════════════
 * 5. CLIENT TABS
 * ═══════════════════════════════════════════════════════════════════════════ */

/**
 * renderTabs()
 * Re-renders the client tab bar to reflect the current clients array
 * and marks the active tab. Called after loading files and after switching.
 */
function renderTabs() {
  document.getElementById('clientTabs').innerHTML = clients
    .map((c, i) =>
      `<div class="ctab ${i === activeClient ? 'active' : ''}"
            onclick="switchClient(${i})"
            title="${escapeHTML(c.name)}">
         ${escapeHTML(c.name)}
       </div>`
    )
    .join('');
}

/**
 * switchClient(i)
 * Switches the dashboard to show a different client's data.
 * Destroys all existing Chart.js instances first to prevent canvas conflicts,
 * then re-renders the full dashboard for the selected client.
 *
 * @param {number} i — index into the clients array
 */
function switchClient(i) {
  /* Destroy all charts before re-rendering to avoid "canvas already in use" errors */
  destroyAllCharts();
  activeClient = i;
  DashState.reset();
  renderTabs();
  try {
    renderClient(i);
  } catch (e) {
    showToast('error', 'Could not load client', e.message, 0);
  }
}

/**
 * destroyAllCharts()
 * Safely destroys every active Chart.js instance.
 * Wrapped in try/catch per instance so one bad chart can't block the rest.
 */
function destroyAllCharts() {
  Object.values(Chart.instances).forEach(ch => {
    try { ch.destroy(); } catch (e) { /* ignore already-destroyed charts */ }
  });
}


/* ═══════════════════════════════════════════════════════════════════════════
 * 6. SCREEN NAVIGATION
 * ═══════════════════════════════════════════════════════════════════════════ */

/**
 * backToUpload()
 * Returns the user to the upload screen.
 * Shows a confirmation dialog first since all loaded data will be lost —
 * the user would have to re-upload their files to get back to the dashboard.
 *
 * Clears pendingFiles so the upload screen starts fresh.
 */
function backToUpload() {
  /* Confirm before discarding all loaded data */
  const confirmed = window.confirm(
    'Return to the upload screen?\n\n' +
    'All loaded client data will be cleared. ' +
    'You will need to re-upload your files.'
  );
  if (!confirmed) return;

  destroyAllCharts();
  document.getElementById('uploadScreen').style.display = 'flex';
  document.getElementById('dashboard').style.display   = 'none';

  /* Reset file list so the upload screen starts clean */
  pendingFiles = [];
  renderFileList();
}
