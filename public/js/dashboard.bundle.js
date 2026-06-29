/**
 * parser.js
 * ─────────────────────────────────────────────────────────────────────────────
 * All data-reading logic: file uploads, Excel parsing, vehicle extraction,
 * column detection, and day-column parsing.
 *
 * This is the SINGLE source of truth for all parsing functions.
 * export.js must NOT duplicate any of these — it calls them directly.
 *
 * Contents:
 *   1. File upload handling  — drag-drop, file list, month guessing
 *   2. Excel parsing         — XLSX.js wrapper
 *   3. Vehicle name          — extraction and normalisation
 *   4. Column detection      — maps friendly keys to actual Excel headers
 *   5. Vehicle map builder   — builds the per-vehicle data object
 *   6. Day-column parsing    — detects and parses date columns in util sheet
 *   7. Util sheet enrichment — merges daily data into the vehicle map
 * ─────────────────────────────────────────────────────────────────────────────
 */


/* ═══════════════════════════════════════════════════════════════════════════
 * 1. FILE UPLOAD HANDLING
 * ═══════════════════════════════════════════════════════════════════════════ */

/**
 * guessMonthFromFilename(filename)
 * Tries to extract a YYYY-MM string from a filename so the month picker
 * on the upload screen is pre-filled automatically.
 *
 * Handles patterns like:
 *   ClientName_Feb_2026.xlsx  →  2026-02
 *   2026_02_report.xlsx       →  2026-02
 *   02_2026.xlsx              →  2026-02
 *
 * @param  {string} filename
 * @returns {string} YYYY-MM string, or '' if no match found
 */
function guessMonthFromFilename(filename) {
  const months = {
    jan: '01', feb: '02', mar: '03', apr: '04',
    may: '05', jun: '06', jul: '07', aug: '08',
    sep: '09', oct: '10', nov: '11', dec: '12',
  };
  const s = filename.toLowerCase();

  /* Pattern: word_Feb_2026 or Feb2026 or February_2026 */
  const abbr = s.match(/\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*[_\-\s]*(\d{4})\b/);
  if (abbr) return `${abbr[2]}-${months[abbr[1]]}`;

  /* Pattern: 2026_02 or 2026-02 */
  const yymm = s.match(/\b(20\d{2})[_\-](0[1-9]|1[0-2])\b/);
  if (yymm) return `${yymm[1]}-${yymm[2]}`;

  /* Pattern: 02_2026 */
  const mmyy = s.match(/\b(0[1-9]|1[0-2])[_\-](20\d{2})\b/);
  if (mmyy) return `${mmyy[2]}-${mmyy[1]}`;

  return '';
}

/**
 * handleFiles(files)
 * Called when the user selects or drops .xlsx / .xls files onto the upload
 * zone. Adds each valid file to the pendingFiles array and refreshes the list.
 *
 * @param {File[]} files — FileList or Array of File objects
 */
function handleFiles(files) {
  files.forEach(f => {
    /* Silently skip non-Excel files dropped alongside valid ones */
    if (!f.name.match(/\.(xlsx|xls)$/i)) return;

    const client = f.name.replace(/\.[^.]+$/, '');   /* strip extension for default client name */
    const month  = guessMonthFromFilename(f.name);
    pendingFiles.push({ file: f, client, month });
  });
  renderFileList();
}

/**
 * renderFileList()
 * Renders the list of pending files on the upload screen.
 * Each row shows the filename, an editable client name, a month picker,
 * and a remove button.
 */
function renderFileList() {
  document.getElementById('fileList').innerHTML = pendingFiles.map((p, i) => `
    <div class="file-item">
      <div class="file-item-name">${p.file.name}</div>
      <input
        class="file-item-client-input"
        type="text"
        value="${p.client}"
        placeholder="Client name"
        oninput="pendingFiles[${i}].client = this.value.trim() || pendingFiles[${i}].file.name.replace(/\\.[^.]+$/,'')"
      />
      <input
        class="file-item-month-input"
        type="month"
        value="${p.month}"
        title="Month for this file's date filter"
        onchange="pendingFiles[${i}].month = this.value"
      />
      <button class="file-item-remove" onclick="removeFile(${i})">×</button>
    </div>`).join('');

  /* Disable the Go button until at least one file is queued */
  document.getElementById('goBtn').disabled = pendingFiles.length === 0;
}

/**
 * removeFile(i)
 * Removes a single file from the pending list and re-renders.
 *
 * @param {number} i — index in pendingFiles
 */
function removeFile(i) {
  pendingFiles.splice(i, 1);
  renderFileList();
}


/* ═══════════════════════════════════════════════════════════════════════════
 * 2. EXCEL PARSING
 * ═══════════════════════════════════════════════════════════════════════════ */

/**
 * parseExcel(file)
 * Reads an Excel file using the XLSX.js library and returns all sheets as
 * an object keyed by sheet name, each containing an array of row objects.
 *
 * defval:0 ensures missing cells return 0 instead of undefined, which
 * prevents NaN values propagating through the violation totals.
 *
 * @param  {File}    file — a .xlsx or .xls File object
 * @returns {Promise<Object>} — { SheetName: [ {col: val, ...}, ... ], ... }
 */
async function parseExcel(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const wb     = XLSX.read(e.target.result, { type: 'array' });
        const sheets = {};
        wb.SheetNames.forEach(name => {
          sheets[name] = XLSX.utils.sheet_to_json(wb.Sheets[name], { defval: 0 });
        });
        resolve(sheets);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = () => reject(new Error('File could not be read.'));
    reader.readAsArrayBuffer(file);
  });
}


/* ═══════════════════════════════════════════════════════════════════════════
 * 3. VEHICLE NAME EXTRACTION AND NORMALISATION
 * ═══════════════════════════════════════════════════════════════════════════ */

/**
 * extractVehicleName(raw)
 * Cleans a raw vehicle name from the Excel row into a consistent plate string.
 *
 * Handles cases like:
 *   "Alliad - KDV 233Q"  →  "KDV 233Q"
 *   "KDV 233Q"           →  "KDV 233Q"
 *   "CompanyName - ABC"  →  "ABC"
 *
 * @param  {*}      raw — raw cell value from Excel
 * @returns {string}    — cleaned vehicle name, or 'Unknown'
 */
function extractVehicleName(raw) {
  if (!raw) return 'Unknown';
  const s = String(raw).trim();

  /* Strip client prefix before a Kenyan plate number (e.g. "Alliad - KDV 233Q") */
  const plateMatch = s.match(/([A-Z]{2,3}\s+\d{3}[A-Z].*)/i);
  if (plateMatch) return plateMatch[1].trim();

  /* Fallback: strip anything before and including a dash */
  return s.replace(/^[^-]+-\s*/, '').trim() || s;
}

/**
 * normaliseName(n)
 * Strips spaces, hyphens and underscores from a name for fuzzy comparison.
 * Used when matching vehicles across sheets where spacing may differ.
 *
 * e.g. "KDK 708G" → "kdk708g"  matches  "KDK708G" → "kdk708g"
 *
 * @param  {string} n
 * @returns {string} — lowercase, stripped
 */
function normaliseName(n) {
  return String(n).toLowerCase().replace(/[\s\-_]/g, '');
}


/* ═══════════════════════════════════════════════════════════════════════════
 * 4. COLUMN DETECTION
 * ═══════════════════════════════════════════════════════════════════════════ */

/**
 * detectCols(rows)
 * Inspects the first row of a sheet to map friendly internal keys to the
 * actual column headers used in the Excel file.
 *
 * Returns an object like:
 *   {
 *     vehicle   : 'Row Labels',
 *     totalDist : 'Total Distance (KM)',
 *     score     : 'Advanced Score',
 *     daysActive: 'Days With Trips',
 *     daysIdle  : 'Days Without Trips',
 *     weekdayDist: 'Weekday Distance (km)',
 *     weekendDist: 'Weekend Distance (km)',
 *     'Over Speeding': 'Over Speeding',   ← one key per violation
 *     ...
 *   }
 *
 * Missing keys are simply absent — callers use fallback lookups.
 *
 * @param  {Object[]} rows — sheet rows from parseExcel()
 * @returns {Object}       — column key map
 */
function detectCols(rows) {
  if (!rows || rows.length === 0) return {};

  const keys  = Object.keys(rows[0]);
  const found = {};

  keys.forEach(k => {
    const kl = k.toLowerCase().trim();

    /* Vehicle identifier column */
    if (
      kl.includes('grouping')      ||
      kl === 'vehicle'             ||
      kl.includes('vehicle name')  ||
      kl === 'row labels'          ||
      kl.includes('row label')
    ) found.vehicle = k;

    /* Distance column */
    if (
      kl.includes('total distance') ||
      kl === 'distance (km)'        ||
      kl === 'distance(km)'
    ) found.totalDist = k;

    /* Score column */
    if (
      (kl.includes('advance') && kl.includes('score')) ||
      kl === 'score'
    ) found.score = k;

    /* Utilisation columns */
    if (kl.includes('days with trips'))  found.daysActive  = k;
    if (kl.includes('days without'))     found.daysIdle    = k;
    if (kl.includes('weekday dist'))     found.weekdayDist = k;
    if (kl.includes('weekend dist'))     found.weekendDist = k;
  });

  /* Map each violation key to its matching Excel column (first 10 chars match) */
  VIOLATIONS.forEach(v => {
    const match = keys.find(k =>
      k.toLowerCase().trim().includes(v.key.toLowerCase().trim().slice(0, 10))
    );
    if (match) found[v.key] = match;
  });

  return found;
}

/**
 * detectCriticalColWarnings(cols, filename)
 * Returns an array of human-readable warning strings for any critical columns
 * that detectCols() could not find. These are shown as toast notifications
 * so the user knows why some data may be missing.
 *
 * @param  {Object} cols     — result of detectCols()
 * @param  {string} filename — used in the warning message
 * @returns {string[]}       — array of warning strings (may be empty)
 */
function detectCriticalColWarnings(cols, filename) {
  const warns = [];
  const f     = filename ? `"${filename}"` : 'the file';

  if (!cols.vehicle)   warns.push(`${f}: Could not find a vehicle name column. Check column is named "Row Labels", "Vehicle", or "Grouping".`);
  if (!cols.score)     warns.push(`${f}: Could not find a score column. Check column is named "Advanced Score" or "Advance Score".`);
  if (!cols.totalDist) warns.push(`${f}: Could not find a distance column. Check column is named "Total Distance (km)" or similar.`);

  return warns;
}


/* ═══════════════════════════════════════════════════════════════════════════
 * 5. VEHICLE MAP BUILDER
 * ═══════════════════════════════════════════════════════════════════════════ */

/**
 * buildVehicleMap(rows, cols)
 * Converts raw Excel scoring rows into a keyed map of vehicle objects.
 *
 * Each vehicle object contains:
 *   { name, totalDist, score, daysActive, daysIdle,
 *     weekdayDist, weekendDist, [violationKey]: count, ... }
 *
 * Days are capped at 31 to guard against grand-total rows or data errors.
 * active + idle is also capped at 31 (one calendar month maximum).
 *
 * @param  {Object[]} rows — scoring sheet rows
 * @param  {Object}   cols — result of detectCols()
 * @returns {Object}       — { vehicleName: vehicleObject, ... }
 */
function buildVehicleMap(rows, cols) {
  const map = {};

  rows.forEach(row => {
    /* Extract and validate vehicle name */
    const name = extractVehicleName(
      row[cols.vehicle] || row['Grouping'] || row['Row Labels'] || row['Vehicle'] || ''
    );
    if (!name || name === 'Unknown') return;

    map[name] = map[name] || { name };

    /* Distance */
    const dist =
      row[cols.totalDist]           ||
      row['Total Distance (km)']    ||
      row['Total Distance (KM)']    ||
      row['Distance (KM)']          ||
      row['Distance (km)']          || 0;
    map[name].totalDist = Number(dist);

    /* Score */
    const sc =
      row[cols.score]       ||
      row['Advanced Score'] ||
      row['Advance Score']  || 0;
    map[name].score = Number(sc);

    /* Active days — capped at 31 */
    if (cols.daysActive || row['Days With Trips'] !== undefined) {
      const rawActive       = Number(row[cols.daysActive] || row['Days With Trips'] || 0);
      map[name].daysActive  = Math.min(rawActive, 31);
    }

    /* Idle days — capped so active + idle never exceeds 31 */
    if (cols.daysIdle || row['Days Without Trips'] !== undefined) {
      const rawIdle      = Number(row[cols.daysIdle] || row['Days Without Trips'] || 0);
      const activeVal    = map[name].daysActive || 0;
      map[name].daysIdle = Math.min(rawIdle, 31 - activeVal);
    }

    /* Weekday / weekend distance */
    if (cols.weekdayDist || row['Weekday Distance (km)'] !== undefined) {
      map[name].weekdayDist = Number(
        row[cols.weekdayDist]         ||
        row['Weekday Distance (km)']  ||
        row['Weekday Distance (KM)']  || 0
      );
    }
    if (cols.weekendDist || row['Weekend Distance (km)'] !== undefined) {
      map[name].weekendDist = Number(
        row[cols.weekendDist]         ||
        row['Weekend Distance (km)']  ||
        row['Weekend Distance (KM)']  || 0
      );
    }

    /* Violation counts — one property per violation type */
    VIOLATIONS.forEach(v => {
      map[name][v.key] = Number(row[cols[v.key]] || row[v.key] || row[v.short] || 0);
    });

    /* ── Violation anomaly detection ── */
    /* Flags vehicles where one violation type is disproportionately high
     * vs all others — signature of a sensor loop fault, not bad driving  */
    map[name]._warnings = map[name]._warnings || [];
    const violCounts = VIOLATIONS.map(v => map[name][v.key] || 0);
    const totalViol  = violCounts.reduce((s, c) => s + c, 0);
    const meanViol   = totalViol / (violCounts.length || 1);

    VIOLATIONS.forEach((v, i) => {
      const count = violCounts[i];
      if (count > 10000) {
        map[name]._warnings.push({
          type   : 'HARD_CAP',
          field  : v.key,
          count,
          message: `${map[name].name} — ${v.key}: ${count.toLocaleString()} events — exceeds physical possibility. Possible sensor loop or data corruption.`,
        });
      } else if (meanViol > 0 && count > meanViol * 15) {
        map[name]._warnings.push({
          type   : 'DISPROPORTIONATE',
          field  : v.key,
          count,
          ratio  : Math.round(count / meanViol),
          message: `${map[name].name} — ${v.key}: ${count.toLocaleString()} events (${Math.round(count / meanViol)}× avg) — likely sensor fault, not driver behaviour.`,
        });
      }
    });

  });  // ← end rows.forEach

  return map;
}

/**
 * buildPrevMap(prevRows, pCols)
 * Builds the previous-month vehicle map used for month-on-month comparisons.
 * Identical structure to buildVehicleMap but only extracts score, distance,
 * and violation counts (no days or utilisation data needed for MoM).
 *
 * Also builds a normalised name index so slight spelling differences between
 * months (e.g. "KDK 708G" vs "KDK708G") still match correctly.
 *
 * @param  {Object[]} prevRows — previous month scoring rows
 * @param  {Object}   pCols    — detectCols() result for the prev sheet
 * @returns {Object}           — { vehicleName: { score, totalDist, violations } }
 */
function buildPrevMap(prevRows, pCols) {
  const prevMap     = {};
  const normIndex   = {};   /* normalised name → original name */

  prevRows.forEach(row => {
    const name = extractVehicleName(
      row[pCols.vehicle] || row['Row Labels'] || row['Grouping'] || row['Vehicle'] || ''
    );
    if (!name || name === 'Unknown') return;

    prevMap[name] = {
      score: Number(
        row[pCols.score] || row['Advanced Score'] || row['Advance Score'] || 0
      ),
      totalDist: Number(
        row[pCols.totalDist]          ||
        row['Total Distance (km)']    ||
        row['Total Distance (KM)']    ||
        row['Distance (KM)']          ||
        row['Distance (km)']          || 0
      ),
    };

    VIOLATIONS.forEach(v => {
      prevMap[name][v.key] = Number(row[pCols[v.key]] || row[v.key] || row[v.short] || 0);
    });


    /* Store normalised key for fuzzy lookup */
    normIndex[normaliseName(name)] = name;
  });

  return { prevMap, normIndex };
}

/**
 * reconcilePrevMap(vehicleMap, prevMap, normIndex)
 * Tries to match current-month vehicles to previous-month vehicles whose
 * names differ only in spacing/capitalisation.
 *
 * Mutates prevMap in-place by adding aliases for any matched vehicles.
 * This means callers can always do prevMap[currentName] after reconciling.
 *
 * @param {Object} vehicleMap — current month vehicle map
 * @param {Object} prevMap    — previous month vehicle map
 * @param {Object} normIndex  — normalised name → prev name index
 */
function reconcilePrevMap(vehicleMap, prevMap, normIndex) {
  Object.keys(vehicleMap).forEach(vname => {
    /* Already matched — nothing to do */
    if (prevMap[vname]) return;

    /* Try normalised lookup */
    const normKey = normaliseName(vname);
    if (normIndex[normKey]) {
      /* Create an alias so current name resolves correctly */
      prevMap[vname] = prevMap[normIndex[normKey]];
    }
  });
}


/* ═══════════════════════════════════════════════════════════════════════════
 * 6. DAY-COLUMN PARSING (UTILISATION SHEET)
 * ═══════════════════════════════════════════════════════════════════════════ */

/**
 * detectDayCols(uKeys)
 * Filters a list of sheet column keys down to only those that look like
 * date or day columns in the utilisation sheet.
 *
 * Recognised formats:
 *   S-1, M-15           (day-of-week letter + day number)
 *   1/3, 15/3/2026      (slash-delimited dates)
 *   1-3, 03-15-2026     (dash-delimited dates)
 *   15 Mar 2026         (natural language dates)
 *   Day 1, Day 15       (generic day labels)
 *   2026-03-15          (ISO dates)
 *   45000               (Excel serial date numbers)
 *
 * @param  {string[]} uKeys — all column keys from the utilisation sheet
 * @returns {string[]}      — only the date/day columns
 */
function detectDayCols(uKeys) {
  return uKeys.filter(k => {
    const t = String(k).trim();
    return (
      /^[SMTFW]-\d+$/.test(t)                                    /* S-1, M-15        */
      || /^\d{1,2}\/\d{1,2}(\/\d{2,4})?$/.test(t)               /* 1/3, 15/3/2026   */
      || /^\d{1,2}-\d{1,2}(-\d{2,4})?$/.test(t)                 /* 1-3, 03-15-2026  */
      || /^\d{1,2}\s+\w+\s*\d{0,4}$/.test(t)                    /* 15 Mar 2026      */
      || /^\w+\s+\d{1,2}(,?\s*\d{4})?$/.test(t)                 /* Mar 15, 2026     */
      || /^Day\s*\d+$/i.test(t)                                  /* Day 1            */
      || /^\d{4}-\d{2}-\d{2}/.test(t)                           /* 2026-03-15 (ISO) */
      || (/^\d{5}$/.test(t) && Number(t) > 40000 && Number(t) < 60000) /* Excel serial */
    );
  });
}

/**
 * excelSerialToLocalDate(serial)
 * Converts an Excel serial date number to a local-timezone Date object.
 *
 * ── TIMEZONE FIX ────────────────────────────────────────────────────────────
 * The naive approach  new Date((serial - 25569) * 86400 * 1000)  constructs
 * a UTC midnight Date. When you then call .getDate() / .getDay(), JavaScript
 * returns the value in the LOCAL timezone. In negative-offset zones (Americas)
 * UTC midnight becomes the previous evening locally, shifting the day back by 1.
 *
 * The fix: extract Y/M/D from the UTC interpretation and then construct a
 * LOCAL date from those parts, so .getDate() always returns the intended day.
 * ────────────────────────────────────────────────────────────────────────────
 *
 * @param  {number} serial — Excel date serial (e.g. 45000)
 * @returns {Date}         — local-timezone Date at midnight
 */
function excelSerialToLocalDate(serial) {
  const utc = new Date(Math.round((serial - 25569) * 86400 * 1000));
  /* Re-construct from UTC parts as a local date to avoid timezone shift */
  return new Date(utc.getUTCFullYear(), utc.getUTCMonth(), utc.getUTCDate());
}

/**
 * parseDayInfo(k)
 * Parses a single day-column key into a structured object used for
 * chart labels, weekend detection, and date filtering.
 *
 * This is the SINGLE authoritative implementation.
 * The copy that previously lived inside export.js has been removed.
 *
 * Returns:
 *   {
 *     label    : string   — display label e.g. "15/3"
 *     isWeekend: boolean  — true for Saturday and Sunday
 *     date     : Date     — local-timezone Date object (or null if unknown)
 *     dayNum   : number   — day-of-month integer (or null)
 *   }
 *
 * @param  {string|number} k — raw column key from the utilisation sheet
 * @returns {Object}
 */
function parseDayInfo(k) {
  const s  = String(k).trim();
  let   dt = null;

  /* ── ISO date: 2026-03-15T... ── */
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) {
    const parts = s.slice(0, 10).split('-');
    dt = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));

  /* ── Excel serial date ── */
  } else if (/^\d{5}$/.test(s) && Number(s) > 40000) {
    dt = excelSerialToLocalDate(Number(s));

  } else {
    /* ── Slash/dash delimited: 1/3/2026, 03-15-2026, 15-03-2026 ── */
    const dmyOrMdy = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
    if (dmyOrMdy) {
      let a = Number(dmyOrMdy[1]);
      let b = Number(dmyOrMdy[2]);
      let y = Number(dmyOrMdy[3]);
      if (y < 100) y += 2000;

      let day, month;
      if      (a > 12) { day = a; month = b; }     /* a must be day    */
      else if (b > 12) { day = b; month = a; }     /* b must be day    */
      else             { day = a; month = b; }     /* ambiguous: D/M   */

      dt = new Date(y, month - 1, day);

    /* ── Natural language: "Mar 15 2026" or "15 Mar 2026" ── */
    } else if (
      /^\w+\s+\d{1,2}(,?\s*\d{4})?$/.test(s) ||
      /^\d{1,2}\s+\w+\s*\d{0,4}$/.test(s)
    ) {
      const parsed = new Date(s);
      if (!isNaN(parsed)) {
        /* Use UTC parts → local date to avoid timezone shift */
        dt = new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate());
      }
    }
  }

  /* ── Successful date parse ── */
  if (dt && !isNaN(dt)) {
    const dow = dt.getDay();   /* 0 = Sunday, 6 = Saturday */
    return {
      label    : `${dt.getDate()}/${dt.getMonth() + 1}`,
      isWeekend: dow === 0 || dow === 6,
      date     : dt,
      dayNum   : dt.getDate(),
    };
  }

  /* ── Day-of-week letter format: S-1, M-15 ── */
  if (/^[SMTFW]-\d+$/.test(s)) {
    const letter = s[0];
    const dayNum = parseInt(s.split('-')[1]) || 1;
    return {
      label    : s,
      isWeekend: letter === 'S',
      date     : new Date(2000, 0, dayNum),   /* placeholder year */
      dayNum,
    };
  }

  /* ── Fallback: just extract a number if present ── */
  const numMatch = s.match(/\d+/);
  const dayN     = numMatch ? parseInt(numMatch[0]) : null;
  return {
    label    : s,
    isWeekend: false,
    date     : dayN ? new Date(2000, 0, dayN) : null,
    dayNum   : dayN,
  };
}


/* ═══════════════════════════════════════════════════════════════════════════
 * 7. UTILISATION SHEET ENRICHMENT
 * ═══════════════════════════════════════════════════════════════════════════ */

/**
 * enrichFromUtilSheet(vehicleMap, utilRows, uCols, dayInfo)
 * Merges daily utilisation data from the util sheet into the vehicleMap
 * that was built from the scoring sheet.
 *
 * Calculates (or reads) for each vehicle:
 *   - daysActive  — number of days the vehicle made at least one trip
 *   - daysIdle    — weekdays with no trips
 *   - weekdayDist — total km on Monday–Friday
 *   - weekendDist — total km on Saturday–Sunday
 *   - totalDist   — fallback if not already set from scoring sheet
 *
 * Strategy for daysActive / daysIdle when no explicit column exists:
 *   1. If we have real date info, use only weekdays as the idle basis
 *      (a vehicle not running on Saturday is not "idle").
 *   2. If column headers are generic labels (Day 1, Day 2…), use all
 *      day columns as the basis since we cannot identify weekends.
 *   3. Active = columns where km > 0. Idle = basis − active.
 *      Never allow idle > (total columns − active).
 *
 * @param {Object}   vehicleMap — mutated in-place
 * @param {Object[]} utilRows   — utilisation sheet rows
 * @param {Object}   uCols      — detectCols() result for util sheet
 * @param {Object[]} dayInfo    — parseDayInfo() results for each day column
 */
function enrichFromUtilSheet(vehicleMap, utilRows, uCols, dayInfo) {
  const hasRealDates  = dayInfo.some(d => d.date && d.date.getFullYear() > 2000);
  const weekdayInfo   = dayInfo.filter(d => !d.isWeekend);

  utilRows.forEach(row => {
    const name = extractVehicleName(
      row[uCols.vehicle] || row['Row Labels'] || row['Grouping'] || row['Vehicle'] || ''
    );
    if (!name || name === 'Unknown') return;

    vehicleMap[name] = vehicleMap[name] || { name };
    const u = vehicleMap[name];

    /* ── Days active / idle ── */
    const hasExplicitDays =
      uCols.daysActive || row['Days With Trips'] !== undefined;

    if (hasExplicitDays) {
      /* Use the sheet's own pre-calculated columns when available */
      u.daysActive = Number(row[uCols.daysActive] || row['Days With Trips']  || 0);
      u.daysIdle   = Number(row[uCols.daysIdle]   || row['Days Without Trips'] || 0);
    } else {
      /* Calculate from daily columns */
      const basisDays     = hasRealDates && weekdayInfo.length > 0 ? weekdayInfo : dayInfo;
      const activeDays    = basisDays.filter(d => Number(row[d.key] || 0) > 0).length;
      const totalActive   = dayInfo.filter(d => Number(row[d.key] || 0) > 0).length;
      const maxIdle       = dayInfo.length - totalActive;

      /* Take the larger of weekday-basis and all-day counts for active */
      u.daysActive = Math.max(activeDays, totalActive);
      /* Idle must not exceed what is physically possible */
      u.daysIdle   = Math.min(basisDays.length - activeDays, maxIdle);
    }

    /* ── Weekday / weekend distance ── */
    const hasExplicitWknd = !!(uCols.weekdayDist && uCols.weekendDist);

    if (hasExplicitWknd) {
      u.weekdayDist = Number(row[uCols.weekdayDist] || 0);
      u.weekendDist = Number(row[uCols.weekendDist] || 0);
    } else {
      /* Sum from individual day columns using weekend flag */
      u.weekdayDist = dayInfo
        .filter(d => !d.isWeekend)
        .reduce((sum, d) => sum + Number(row[d.key] || 0), 0);
      u.weekendDist = dayInfo
        .filter(d => d.isWeekend)
        .reduce((sum, d) => sum + Number(row[d.key] || 0), 0);
    }

    /* ── Total distance fallback ── */
    if (!u.totalDist || u.totalDist === 0) {
      u.totalDist = Number(
        row['Total Distance (KM)'] ||
        row['Total Distance (km)'] ||
        row[uCols.totalDist]       || 0
      );
    }
  });
}

/**
 * buildDailyTotals(utilRows, dayInfo)
 * Calculates the fleet-wide total distance for each day column.
 * Used to power the daily distance line chart.
 *
 * @param  {Object[]} utilRows — utilisation sheet rows
 * @param  {Object[]} dayInfo  — parseDayInfo() results
 * @returns {Object}           — { "15/3": 4500, "16/3": 3200, ... }
 */
function buildDailyTotals(utilRows, dayInfo) {
  const totals = {};
  dayInfo.forEach(d => {
    totals[d.label] = utilRows.reduce(
      (sum, row) => sum + Number(row[d.key] || 0), 0
    );
  });
  return totals;
}

/**
 * buildVehicleDailyData(utilRows, uCols, dayInfo)
 * Builds a per-vehicle array of daily km values.
 * Used for the single-vehicle daily distance chart and the exported report.
 *
 * @param  {Object[]} utilRows — utilisation sheet rows
 * @param  {Object}   uCols    — detectCols() result
 * @param  {Object[]} dayInfo  — parseDayInfo() results
 * @returns {Object}           — { vehicleName: [day1km, day2km, ...], ... }
 */
function buildVehicleDailyData(utilRows, uCols, dayInfo) {
  const result = {};
  utilRows.forEach(row => {
    const name = extractVehicleName(
      row[uCols.vehicle] || row['Row Labels'] || row['Grouping'] || row['Vehicle'] || ''
    );
    if (!name || name === 'Unknown') return;
    result[name] = dayInfo.map(d => Math.round(Number(row[d.key] || 0) * 100) / 100);
  });
  return result;
}
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

  /* ── Per-vehicle anomaly warnings from buildVehicleMap() ── */
  vehicles.forEach(v => {
    (v._warnings || []).forEach(w => warns.push(w.message));
  });

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
        <div id="riskScroll_${i}" style="overflow-y:auto;overflow-x:auto;max-height:320px">
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
        <div id="wdweScroll_${i}" style="overflow-y:auto;overflow-x:auto;max-height:320px">
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
        <div id="utilScroll_${i}" style="overflow-y:auto;overflow-x:auto;max-height:320px">
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
      <div id="prevScroll_${i}" style="overflow-y:auto;overflow-x:auto;max-height:400px">
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
        <div id="distCompScroll_${i}" style="overflow-y:auto;overflow-x:auto;max-height:400px">
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
.warn-banner{display:flex;align-items:flex-start;gap:12px;background:rgba(224,149,69,.08);border:1px solid rgba(224,149,69,.35);border-radius:var(--radius-lg);padding:14px 18px;margin-bottom:1.25rem;font-size:13px;color:var(--text2);line-height:1.6}.warn-banner b{color:var(--text)}.warn-banner ul{margin:6px 0 0 1.1rem;padding:0}.warn-banner li{margin-bottom:3px}
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

    /* ── Data anomaly warning banner — matches main dashboard warn-banner style ── */
    (vehicles.some(function(v){return(v._warnings||[]).length>0;}) ?
      '<div class="warn-banner">'+
        '<span style="font-size:16px;flex-shrink:0">&#9888;</span>'+
        '<div><b>Data warnings for '+escHTML(D.clientName)+'</b>'+
          '<ul>'+
            vehicles.filter(function(v){return(v._warnings||[]).length>0;}).map(function(v){
              return v._warnings.map(function(w){
                return '<li>'+escHTML(v.name)+' — '+escHTML(w.message)+'</li>';
              }).join('');
            }).join('')+
          '</ul>'+
        '</div>'+
      '</div>'
    : '')+

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
      '<div class="card"><div class="card-title" id="eRiskTitle">All vehicles — advanced score ranking</div><div id="eRiskScroll" style="overflow-y:auto;overflow-x:auto;max-height:320px"><div id="eRiskWrap" style="position:relative;height:320px"><canvas id="eRiskChart"></canvas></div></div></div>'+
      '<div class="card"><div class="card-title" id="eViolTitle">Violation breakdown — fleet total</div><div class="legend-row" id="eViolLeg"></div><div class="chart-wrap" style="height:280px"><canvas id="eViolChart"></canvas></div></div>'+
    '</div>'+

    '<div class="grid-2" style="margin-bottom:1rem">'+
      '<div class="card"><div class="card-title" id="eWdweTitle">Weekday vs weekend distance</div><div class="legend-row"><span class="leg"><span class="leg-dot" style="background:#4f8ef7"></span>Weekday</span><span class="leg"><span class="leg-dot" style="background:#2ec4b6"></span>Weekend</span></div><div id="eWdweScroll" style="overflow-y:auto;overflow-x:auto;max-height:320px"><div id="eWdweWrap" style="position:relative;height:320px"><canvas id="eWdweChart"></canvas></div></div></div>'+
      '<div class="card"><div class="card-title" id="eUtilTitle">Most idle vehicles — days active vs idle</div><div class="legend-row"><span class="leg"><span class="leg-dot" style="background:#4f8ef7"></span>Active</span><span class="leg"><span class="leg-dot" style="background:#555b72"></span>Idle</span></div><div id="eUtilScroll" style="overflow-y:auto;overflow-x:auto;max-height:320px"><div id="eUtilWrap" style="position:relative;height:320px"><canvas id="eUtilChart"></canvas></div></div></div>'+
    '</div>'+

    (hasPrev?'<div class="card" style="margin-bottom:1rem"><div class="card-title" id="ePrevTitle">Month-on-month Advance score — current vs previous</div><div class="legend-row"><span class="leg"><span class="leg-dot" style="background:#3b6edc"></span>Current</span><span class="leg"><span class="leg-dot" style="background:#6ea8ff"></span>Previous</span></div><div id="ePrevScroll" style="overflow-y:auto;overflow-x:auto;max-height:400px"><div id="ePrevWrap" style="position:relative;height:400px"><canvas id="ePrevChart"></canvas></div></div></div>':'')+

    (hasPrev?'<div class="card print-section-distcomp" style="margin-bottom:1rem"><div class="card-title" id="eDCCompTitle">Distance comparison — current vs previous month</div><div class="legend-row"><span class="leg"><span class="leg-dot" style="background:#3db87a"></span>Increased</span><span class="leg"><span class="leg-dot" style="background:#e05353"></span>Decreased</span><span class="leg"><span class="leg-dot" style="background:rgba(79,142,247,0.45)"></span>Previous month</span></div><div id="eDCCompScroll" style="overflow-y:auto;overflow-x:auto;max-height:400px"><div id="eDCCompWrap" style="position:relative;height:400px"><canvas id="eDCCompChart"></canvas></div></div></div>':'')+

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
    options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{x:{ticks:{callback:function(v){return v.toLocaleString();},color:'#555b72'},grid:{color:'rgba(255,255,255,0.04)'},border:{display:false}},y:{ticks:{font:function(ctx){return tickFont(ctx.tick.label,function(){return eRiskSel;});},color:function(ctx){return tickColor(ctx.tick.label,function(){return eRiskSel;});}},afterFit:function(axis){axis.width=Math.max(axis.width,90);}}}}
  });

  document.getElementById('eViolLeg').innerHTML=activeViolations.map(function(v,idx){return'<span class="leg"><span class="leg-dot" style="background:'+VIOL_COLORS[idx%VIOL_COLORS.length]+'"></span>'+escHTML(v.short)+'</span>';}).join('');
  eViolChart=new Chart(document.getElementById('eViolChart'),{
    type:'doughnut',data:{labels:activeViolations.map(function(v){return v.short;}),datasets:[{data:violTotals,backgroundColor:VIOL_COLORS.slice(0,activeViolations.length),borderWidth:0}]},
    options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}}}
  });

  eWdweChart=new Chart(document.getElementById('eWdweChart'),{
    type:'bar',data:{labels:allWdwe.map(function(v){return v.name;}),datasets:[{label:'Weekday',data:allWdwe.map(function(v){return v.weekdayDist||0;}),backgroundColor:'#4f8ef7',borderRadius:2},{label:'Weekend',data:allWdwe.map(function(v){return v.weekendDist||0;}),backgroundColor:'#2ec4b6',borderRadius:2}]},
    options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{x:{stacked:true,ticks:{callback:function(v){return v.toLocaleString();},color:'#555b72'},grid:{color:'rgba(255,255,255,0.04)'},border:{display:false}},y:{stacked:true,ticks:{font:function(ctx){return tickFont(ctx.tick.label,function(){return eWdweSel;});},color:function(ctx){return tickColor(ctx.tick.label,function(){return eWdweSel;});}},afterFit:function(axis){axis.width=Math.max(axis.width,90);}}}}
  });

  eUtilChart=new Chart(document.getElementById('eUtilChart'),{
    type:'bar',data:{labels:allUtil.map(function(v){return v.name;}),datasets:[{label:'Active',data:allUtil.map(function(v){return v.daysActive||0;}),backgroundColor:'#4f8ef7',borderRadius:2},{label:'Idle',data:allUtil.map(function(v){return v.daysIdle||0;}),backgroundColor:'#555b72',borderRadius:2}]},
    options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{x:{stacked:true,ticks:{color:'#555b72'},grid:{color:'rgba(255,255,255,0.04)'},border:{display:false}},y:{stacked:true,ticks:{font:function(ctx){return tickFont(ctx.tick.label,function(){return eUtilSel;});},color:function(ctx){return tickColor(ctx.tick.label,function(){return eUtilSel;});}},afterFit:function(axis){axis.width=Math.max(axis.width,90);}}}}
  });

  if(hasPrev){
    var allPrev=vehicles.filter(function(v){return prevMap[v.name];}).sort(function(a,b){return(b.score||0)-(a.score||0);});
    var prevH=Math.max(400,allPrev.length*36);
    document.getElementById('ePrevWrap').style.height=prevH+'px';
    ePrevChart=new Chart(document.getElementById('ePrevChart'),{
      type:'bar',data:{labels:allPrev.map(function(v){return v.name;}),datasets:[{label:'Current',data:allPrev.map(function(v){return v.score||0;}),backgroundColor:allPrev.map(function(){return'#3b6edc';}),borderRadius:3},{label:'Previous',data:allPrev.map(function(v){return prevMap[v.name]?prevMap[v.name].score||0:0;}),backgroundColor:'#6ea8ff',borderRadius:3}]},
      options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{x:{ticks:{callback:function(v){return v.toLocaleString();},color:'#555b72'},grid:{color:'rgba(255,255,255,0.04)'},border:{display:false}},y:{ticks:{font:function(ctx){return tickFont(ctx.tick.label,function(){return ePrevSel;});},color:function(ctx){return tickColor(ctx.tick.label,function(){return ePrevSel;});}},afterFit:function(axis){axis.width=Math.max(axis.width,90);}}}}
    });

    var allNames=new Set([...vehicles.map(function(v){return v.name;}),...Object.keys(prevMap)]);
    var allDC=[...allNames].map(function(name){var curr=vehicles.find(function(v){return v.name===name;});return{name:name,currDist:curr?curr.totalDist||0:0,prevDist:prevMap[name]?prevMap[name].totalDist||0:0};}).sort(function(a,b){return b.currDist-a.currDist;});
    var dcH=Math.max(400,allDC.length*34);
    document.getElementById('eDCCompWrap').style.height=dcH+'px';
    eDCCompChart=new Chart(document.getElementById('eDCCompChart'),{
      type:'bar',data:{labels:allDC.map(function(v){return v.name;}),datasets:[{label:'Current',data:allDC.map(function(v){return v.currDist;}),backgroundColor:allDC.map(function(v){return v.currDist>=v.prevDist?'#3db87a':'#e05353';}),borderRadius:3},{label:'Previous',data:allDC.map(function(v){return v.prevDist;}),backgroundColor:allDC.map(function(v){return v.currDist>=v.prevDist?'#e05353':'#3db87a';}),borderRadius:3}]},
      options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:{callbacks:{label:function(ctx){return ctx.dataset.label+': '+Math.round(ctx.raw).toLocaleString()+' km';}}}},scales:{x:{ticks:{callback:function(v){return v.toLocaleString()+' km';},color:'#555b72'},grid:{color:'rgba(255,255,255,0.04)'},border:{display:false}},y:{ticks:{font:function(ctx){return tickFont(ctx.tick.label,function(){return eDistSel;});},color:function(ctx){return tickColor(ctx.tick.label,function(){return eDistSel;});}},border:{display:false},afterFit:function(axis){axis.width=Math.max(axis.width,90);}}}}
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