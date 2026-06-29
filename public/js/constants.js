/**
 * constants.js
 * ─────────────────────────────────────────────────────────────────────────────
 * Single source of truth for all fixed values used across the dashboard.
 * Every other file reads from here — nothing is hardcoded elsewhere.
 *
 * Contents:
 *   1. SCORE_BANDS   — risk classification thresholds and display properties
 *   2. VIOLATIONS    — violation definitions, Excel column keys, and risk levels
 *   3. VIOL_COLORS   — chart colour palette for violation breakdowns
 *   4. COMPARE_COLORS — colour palette for the vehicle compare mode
 * ─────────────────────────────────────────────────────────────────────────────
 */


/* ── 1. SCORE BANDS ──────────────────────────────────────────────────────────
 *
 * Defines the three risk tiers used everywhere in the dashboard:
 *   - KPI "at-risk" counter
 *   - Score ranking bar colours
 *   - Flagged vehicles table badges
 *   - Progress heatmap cells
 *
 * Lower score = safer driver (score is a penalty, not a rating).
 * A vehicle is considered "at-risk" when its score exceeds SCORE_BANDS.safe.max.
 * ─────────────────────────────────────────────────────────────────────────── */

const SCORE_BANDS = {
  safe: {
    label     : 'Low Risk',
    min       : 0,
    max       : 20,
    color     : '#3db87a',   /* green  */
    colorLight: '#d1f5e3',   /* green tint for backgrounds */
    textColor : '#1a6640',   /* dark green for text on light backgrounds */
  },
  moderate: {
    label     : 'Moderate Risk',
    min       : 21,
    max       : 40,
    color     : '#e09545',   /* amber  */
    colorLight: '#fdf0dc',
    textColor : '#7a4a10',
  },
  high: {
    label     : 'High Risk',
    min       : 41,
    max       : Infinity,
    color     : '#e05353',   /* red    */
    colorLight: '#fde8e8',
    textColor : '#7a1a1a',
  },
};

/**
 * getScoreBand(score)
 * Returns the SCORE_BANDS entry that the given score falls into.
 * Use this instead of repeating if/else chains across files.
 *
 * @param  {number} score  — vehicle advanced score
 * @returns {object}       — one of SCORE_BANDS.safe / .moderate / .high
 */
function getScoreBand(score) {
  const s = Number(score) || 0;
  if (s <= SCORE_BANDS.safe.max)     return SCORE_BANDS.safe;
  if (s <= SCORE_BANDS.moderate.max) return SCORE_BANDS.moderate;
  return SCORE_BANDS.high;
}

/**
 * isAtRisk(score)
 * Returns true when a vehicle's score exceeds the safe band (score > 20).
 * Used by the KPI counter and the flagged vehicles table.
 *
 * @param  {number} score
 * @returns {boolean}
 */
function isAtRisk(score) {
  return (Number(score) || 0) > SCORE_BANDS.safe.max;
}


/* ── 2. VIOLATIONS ───────────────────────────────────────────────────────────
 *
 * Each entry describes one violation type tracked in the Excel data.
 *
 * Fields:
 *   key   — exact column header used in the Excel scoring sheet
 *           (used by detectCols() and buildVehicleMap() in parser.js)
 *   short — abbreviated label used in charts and table headers
 *   risk  — severity tier: 'high' | 'med' | 'low'
 *   desc  — plain-English explanation shown in the violation reference guide
 * ─────────────────────────────────────────────────────────────────────────── */

const VIOLATIONS = [
  {
    key  : 'Diagnostic: Fault no Engine RPM',
    short: 'Engine RPM Fault',
    risk : 'high',
    desc : 'Causes vehicles to miss critical events such as freewheeling and over-revving.',
  },
  {
    key  : 'Free Wheeling',
    short: 'Free Wheeling',
    risk : 'high',
    desc : 'Likely to cause gearbox damage and engine problems. Increased chance of accident if wrong gear is engaged.',
  },
  {
    key  : 'Possible impact',
    short: 'Possible Impact',
    risk : 'high',
    desc : 'Monitors impact severity. May lead to drive shaft breaking.',
  },
  {
    key  : 'Harsh Acceleration',
    short: 'Harsh Acceleration',
    risk : 'med',
    desc : 'Reduces tire life and increases fuel consumption.',
  },
  {
    key  : 'Harsh Braking',
    short: 'Harsh Braking',
    risk : 'high',
    desc : 'Damages brake pads, drums and suspension. May lead to tire burst.',
  },
  {
    key  : 'Idle - excessive',
    short: 'Excessive Idle',
    risk : 'low',
    desc : 'Results in higher fuel consumption.',
  },
  {
    key  : 'Night Driving',
    short: 'Night Driving',
    risk : 'med',
    desc : 'Increases risk of theft and accidents due to poor visibility.',
  },
  {
    key  : 'Over Revving',
    short: 'Over Revving',
    risk : 'med',
    desc : 'Causes increased wear of engine parts and high fuel consumption.',
  },
  {
    key  : 'Over Speeding',
    short: 'Over Speeding',
    risk : 'high',
    desc : 'Results in high fuel consumption and a high risk of accidents.',
  },
  {
    key  : 'Harsh Cornering',
    short: 'Harsh Cornering',
    risk : 'med',
    desc : 'Increases tire wear and risk of loss of vehicle control.',
  },
  {
    key  : '3-Axis - Possible Accident (In Trip)',
    short: 'Possible Accident',
    risk : 'high',
    desc : 'Monitors impact severity in-trip. May lead to drive shaft breakage.',
  },
  {
    key  : 'Over speeding in location',
    short: 'Speed in Zone',
    risk : 'high',
    desc : 'Results in high fuel consumption and high risk of accidents in specific zones.',
  },
];


/* ── 3. VIOLATION COLOURS ────────────────────────────────────────────────────
 *
 * One colour per violation, assigned by index position.
 * Matches the order of the VIOLATIONS array above — VIOL_COLORS[0] is used
 * for VIOLATIONS[0] (Engine RPM Fault), and so on.
 *
 * Used by: violation donut chart, violation breakdown table.
 * ─────────────────────────────────────────────────────────────────────────── */

const VIOL_COLORS = [
  '#e05353',  /* Engine RPM Fault    — red        */
  '#4f8ef7',  /* Free Wheeling       — blue        */
  '#e09545',  /* Possible Impact     — amber       */
  '#3db87a',  /* Harsh Acceleration  — green       */
  '#d45b9f',  /* Harsh Braking       — pink        */
  '#2ec4b6',  /* Excessive Idle      — teal        */
  '#a855f7',  /* Night Driving       — purple      */
  '#f97316',  /* Over Revving        — orange      */
  '#eab308',  /* Over Speeding       — yellow      */
  '#06b6d4',  /* Harsh Cornering     — cyan        */
  '#84cc16',  /* Possible Accident   — lime        */
  '#f43f5e',  /* Speed in Zone       — rose        */
];


/* ── 4. COMPARE MODE COLOURS ─────────────────────────────────────────────────
 *
 * Used when the user selects multiple vehicles in Compare mode.
 * Up to 5 vehicles can be compared; each gets a distinct colour from this list.
 * Colours are preserved per vehicle for the duration of the session so the
 * same vehicle always appears in the same colour across all charts.
 * ─────────────────────────────────────────────────────────────────────────── */

const COMPARE_COLORS_MAIN = [
  '#4f8ef7',  /* blue   */
  '#3db87a',  /* green  */
  '#e09545',  /* amber  */
  '#a855f7',  /* purple */
  '#2ec4b6',  /* teal   */
];
