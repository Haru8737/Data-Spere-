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
