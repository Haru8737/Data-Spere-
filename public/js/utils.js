/**
 * utils.js
 * ─────────────────────────────────────────────────────────────────────────────
 * Shared utility helpers used across dashboard.bundle.js and export.js.
 * All functions here are pure — no DOM side effects, no global state.
 *
 * Contents:
 *   1. DOM helpers      — setEl, setKpi, setWrapHeight
 *   2. Number helpers   — fmt, fmtPct, safeNum, pct
 *   3. Date helpers     — formatDate, pad2
 *   4. String helpers   — escapeHTML (alias of escHTML in export context)
 *   5. Score helpers    — changeTag (shared KPI change indicator)
 * ─────────────────────────────────────────────────────────────────────────────
 */


/* ═══════════════════════════════════════════════════════════════════════════
 * 1. DOM HELPERS
 * ═══════════════════════════════════════════════════════════════════════════ */

/**
 * setEl(id, text)
 * Sets the textContent of a DOM element by ID.
 * Safe no-op if the element does not exist.
 *
 * @param {string} id   — element ID
 * @param {*}      text — value (coerced to string)
 */
function setEl(id, text) {
  const el = document.getElementById(id);
  if (el) el.textContent = text;
}

/**
 * setKpi(i, n, label, value, sub)
 * Updates a KPI card's label, value, and sub-text in one call.
 * Follows the dashboard's KPI element ID convention: kpiLbl{n}_{i}
 *
 * @param {number} i     — client index
 * @param {number} n     — KPI card index (0–6)
 * @param {string} label — top label text
 * @param {*}      value — main value (displayed large)
 * @param {string} sub   — sub-label below the value
 */
function setKpi(i, n, label, value, sub) {
  setEl(`kpiLbl${n}_${i}`, label);
  setEl(`kpiVal${n}_${i}`, value);
  setEl(`kpiSub${n}_${i}`, sub);
}

/**
 * setWrapHeight(i, name, height, maxHeight)
 * Sets the pixel height of a chart's wrap and scroll container.
 *
 * Naming convention used throughout the dashboard:
 *   wrapEl   = `${name}Wrap_${i}`
 *   scrollEl = `${name}Scroll_${i}`
 *
 * @param {number} i         — client index
 * @param {string} name      — chart name prefix e.g. 'risk', 'wdwe', 'util'
 * @param {number} height    — px height for the inner wrap div
 * @param {number} maxHeight — px maxHeight for the outer scroll div
 */
function setWrapHeight(i, name, height, maxHeight) {
  const wrap   = document.getElementById(`${name}Wrap_${i}`);
  const scroll = document.getElementById(`${name}Scroll_${i}`);
  if (wrap)   wrap.style.height      = `${height}px`;
  if (scroll) scroll.style.maxHeight = `${maxHeight}px`;
}


/* ═══════════════════════════════════════════════════════════════════════════
 * 2. NUMBER HELPERS
 * ═══════════════════════════════════════════════════════════════════════════ */

/**
 * fmt(n)
 * Formats a number with locale-appropriate thousands separators.
 * Returns '0' for null/undefined/NaN.
 *
 * @param  {*} n
 * @returns {string}
 */
function fmt(n) {
  return (Number(n) || 0).toLocaleString();
}

/**
 * fmtPct(n)
 * Formats a number as a rounded percentage string e.g. '42%'.
 *
 * @param  {number} n
 * @returns {string}
 */
function fmtPct(n) {
  return `${Math.round(Number(n) || 0)}%`;
}

/**
 * safeNum(val, fallback)
 * Coerces a value to a number, returning fallback if the result is NaN.
 * Guards against undefined/null/empty string producing NaN in calculations.
 *
 * @param  {*}      val
 * @param  {number} fallback — default 0
 * @returns {number}
 */
function safeNum(val, fallback = 0) {
  const n = Number(val);
  return isNaN(n) ? fallback : n;
}

/**
 * pct(a, b)
 * Returns (a / b) * 100 rounded to the nearest integer.
 * Returns 0 safely when b is 0 or falsy.
 *
 * @param  {number} a — numerator
 * @param  {number} b — denominator
 * @returns {number}
 */
function pct(a, b) {
  return b ? Math.round((a / b) * 100) : 0;
}


/* ═══════════════════════════════════════════════════════════════════════════
 * 3. DATE HELPERS
 * ═══════════════════════════════════════════════════════════════════════════ */

/**
 * pad2(n)
 * Zero-pads a number to at least 2 digits.
 * e.g. pad2(3) → '03', pad2(12) → '12'
 *
 * @param  {number} n
 * @returns {string}
 */
function pad2(n) {
  return String(n).padStart(2, '0');
}

/**
 * formatDate(date)
 * Formats a Date object as a YYYY-MM-DD string using local timezone parts.
 * Avoids UTC-offset bugs that occur with toISOString().
 *
 * @param  {Date} date
 * @returns {string} e.g. '2026-03-15'
 */
function formatDate(date) {
  return `${date.getFullYear()}-${pad2(date.getMonth() + 1)}-${pad2(date.getDate())}`;
}


/* ═══════════════════════════════════════════════════════════════════════════
 * 4. CHANGE TAG
 * Shared between dashboard.bundle.js KPI cards and export.js KPI cards.
 * Previously duplicated in both files — now a single source of truth.
 * ═══════════════════════════════════════════════════════════════════════════ */

/**
 * changeTag(curr, prev, lowerIsBetter)
 * Returns an HTML <span> showing a month-on-month change indicator.
 *
 * When the change exceeds 100%, shows an absolute difference instead of
 * a percentage to avoid misleading figures like "▲ 900%".
 *
 * @param  {number}  curr          — current month value
 * @param  {number}  prev          — previous month value
 * @param  {boolean} lowerIsBetter — true for scores, false for distance
 * @returns {string} HTML string (safe to inject via innerHTML)
 */
function changeTag(curr, prev, lowerIsBetter = true) {
  if (prev === null || prev === undefined || prev === 0) return '';

  const diff   = curr - prev;
  const rawPct = Math.abs(diff) / Math.abs(prev) * 100;

  if (rawPct < 1) return `<span class="kpi-change kpi-same">→ no change</span>`;

  const improved = lowerIsBetter ? diff < 0 : diff > 0;
  const cls      = improved ? 'kpi-down' : 'kpi-up';
  const arrow    = diff > 0 ? '▲' : '▼';

  if (rawPct > 100) {
    const absDiff = Math.round(Math.abs(diff));
    const sign    = diff > 0 ? '+' : '-';
    return `<span class="kpi-change ${cls}">${arrow} ${sign}${absDiff} vs last month</span>`;
  }

  return `<span class="kpi-change ${cls}">${arrow} ${Math.round(rawPct)}% vs last month</span>`;
}
