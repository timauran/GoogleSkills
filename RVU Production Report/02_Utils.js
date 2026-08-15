// ============================================================
// 02_Utils — shared helpers
// Spreadsheet access, date canonicalization, shifts, colors, logging.
// ============================================================

// ── Spreadsheet access ──────────────────────────────────────

/**
 * Opens the data spreadsheet using the ID stored in Script Properties.
 * Throws a clear error if createAndSetupSpreadsheet() has not been run.
 */
function getDataSS() {
  const id = PropertiesService.getScriptProperties().getProperty(CONFIG.PROP_SPREADSHEET_ID);
  if (!id) {
    throw new Error(
      'Data spreadsheet not configured. ' +
      'Run createAndSetupSpreadsheet() from the Apps Script editor first.'
    );
  }
  return SpreadsheetApp.openById(id);
}

/** Get an existing sheet by name, or create it if missing. */
function getOrCreateSheet(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

// ── Column resolution ───────────────────────────────────────

/**
 * Maps the four logical fields onto the column indices of a header row,
 * matching CONFIG.COL_ALIASES case-insensitively after trimming.
 * Shared by the emailed CSV (03_Ingest) and imported sheets (09_ImportSheet)
 * so a renamed report column only has to be fixed in one place.
 *
 * @return {{col:{date:number,provider:number,shift:number,rvu:number},
 *           missing:string[]}}  missing lists required fields not found;
 *           `shift` is -1 when absent (ingest proceeds with a blank shift).
 */
function resolveColumns(headers) {
  const clean   = headers.map(h => String(h).trim().toLowerCase());
  const col     = {};
  const missing = [];

  Object.keys(CONFIG.COL_ALIASES).forEach(field => {
    let idx = -1;
    for (const alias of CONFIG.COL_ALIASES[field]) {
      idx = clean.indexOf(alias.toLowerCase());
      if (idx !== -1) break;
    }
    col[field] = idx;
    if (idx === -1 && field !== 'shift') missing.push(field);   // shift is optional
  });

  return { col, missing };
}

/** Reads a cell by resolved column index; '' when the column is absent. */
function cellAt(row, idx) {
  return idx === -1 || row[idx] === undefined || row[idx] === null ? '' : String(row[idx]).trim();
}

// ── Shift classification ────────────────────────────────────

/** True if the shift is "not scheduled" (case-insensitive, trimmed). */
function isNS(shift) {
  return String(shift).trim().toLowerCase() === CONFIG.NOT_SCHEDULED_KEY;
}

/**
 * True if a shift name contains "weekend". Weekend shifts have no aveRVU
 * and are excluded from mismatch alerts, thresholds, and Raw Data.
 */
function isWeekendShift(shift) {
  return String(shift).trim().toLowerCase().includes('weekend');
}

/**
 * True if a shift is exempt from green/orange colour-coding in the daily
 * email — Not Scheduled, or any name matching CONFIG.NO_HIGHLIGHT_SHIFTS
 * (e.g. "PM Imaging"). Exempt rotations still show their RVU and %.
 */
function isNoHighlightShift(shift) {
  const s = String(shift).trim().toLowerCase();
  if (!s || isNS(s)) return true;
  return CONFIG.NO_HIGHLIGHT_SHIFTS.some(x => s.includes(String(x).toLowerCase()));
}

// ── Provider names ──────────────────────────────────────────

/**
 * Comparator sorting provider names by last name, A→Z. Names arrive as
 * "Last, First"; anything without a comma falls back to its final word,
 * so "Timothy Auran" and "Auran, Timothy" both sort under A.
 */
function byLastName(a, b) {
  const last = (name) => {
    const s = String(name).trim();
    const i = s.indexOf(',');
    const l = i !== -1 ? s.slice(0, i) : s.split(/\s+/).pop();
    return l.trim().toLowerCase();
  };
  const la = last(a), lb = last(b);
  if (la !== lb) return la < lb ? -1 : 1;
  return String(a).toLowerCase() < String(b).toLowerCase() ? -1 : 1;   // tie → full name
}

// ── Dates ───────────────────────────────────────────────────
// Canonical internal date format is ISO "yyyy-MM-dd" (date only).
// ISO sorts correctly as plain text and is unambiguous for downstream
// scripts. US "M/D/YYYY" is used only for display (headers, email).

/**
 * Normalize any date value to canonical ISO "yyyy-MM-dd".
 * Accepts Date objects and strings ("4/15/2026 6:00 AM", "2026-04-15",
 * "04/15/2026", etc.). Returns '' if the value cannot be parsed.
 *
 * Common formats are parsed by hand to avoid the JS/Apps Script
 * timezone pitfalls of new Date() (e.g. "2026-04-15" parsing as UTC
 * midnight and rolling back a day in Pacific time).
 */
function toISO(val) {
  if (val === null || val === undefined || val === '') return '';

  if (val instanceof Date) {
    if (isNaN(val.getTime())) return '';
    return Utilities.formatDate(val, CONFIG.TIMEZONE, 'yyyy-MM-dd');
  }

  const s = String(val).trim();
  if (!s) return '';

  // Already ISO (optionally with a time component) → take the date part
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return `${m[1]}-${m[2]}-${m[3]}`;

  // US M/D/YYYY (optionally with a time component)
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) {
    return `${m[3]}-${m[1].padStart(2, '0')}-${m[2].padStart(2, '0')}`;
  }

  // Fallback: let Date parse it, then format in the project timezone
  const d = new Date(s);
  if (!isNaN(d.getTime())) return Utilities.formatDate(d, CONFIG.TIMEZONE, 'yyyy-MM-dd');

  return '';
}

/** ISO "yyyy-MM-dd" → US "M/D/YYYY" for display. Passes through non-ISO input. */
function isoToUS(iso) {
  const m = String(iso).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return m ? `${Number(m[2])}/${Number(m[3])}/${m[1]}` : String(iso);
}

/** ISO "yyyy-MM-dd" → a local-midnight Date (safe for weekday/holiday math). */
function isoToLocalDate(iso) {
  const m = String(iso).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : null;
}

/** Sort ISO date strings newest-first (most-recent column on the left). */
function sortISOsDesc(arr) {
  return arr.slice().sort((a, b) => (a < b ? 1 : a > b ? -1 : 0));
}

/** Prior calendar day in Pacific time as ISO "yyyy-MM-dd". */
function priorDayISO() {
  const todayIso = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd');
  const d = isoToLocalDate(todayIso);
  d.setDate(d.getDate() - 1);
  const mo = String(d.getMonth() + 1).padStart(2, '0');
  const da = String(d.getDate()).padStart(2, '0');
  return `${d.getFullYear()}-${mo}-${da}`;
}

/** True if an ISO date falls on a weekend or observed US federal holiday. */
function isoIsWeekendOrHoliday(iso) {
  const d = isoToLocalDate(iso);
  if (!d) return false;
  const dow = d.getDay();               // 0=Sun, 6=Sat
  if (dow === 0 || dow === 6) return true;
  return isUSFederalHoliday(d);
}

/** True if local-midnight Date d is an observed US federal holiday. */
function isUSFederalHoliday(d) {
  const yr = d.getFullYear();

  const nthWD = (mo, wd, n) => {         // nth weekday of a month
    const first = new Date(yr, mo - 1, 1);
    const day = ((wd - first.getDay()) + 7) % 7 + 1 + (n - 1) * 7;
    return new Date(yr, mo - 1, day);
  };
  const lastWD = (mo, wd) => {           // last weekday of a month
    const last = new Date(yr, mo, 0);
    return new Date(yr, mo - 1, last.getDate() - ((last.getDay() - wd + 7) % 7));
  };
  const obs = (date) => {                // observed: Sat→Fri, Sun→Mon
    const dw = date.getDay();
    if (dw === 6) return new Date(date.getFullYear(), date.getMonth(), date.getDate() - 1);
    if (dw === 0) return new Date(date.getFullYear(), date.getMonth(), date.getDate() + 1);
    return date;
  };
  const eq = (holiday) => {
    const o = obs(holiday);
    return d.getFullYear() === o.getFullYear()
        && d.getMonth()    === o.getMonth()
        && d.getDate()     === o.getDate();
  };

  return (
    eq(new Date(yr, 0,  1))  ||   // New Year's Day
    eq(nthWD(1, 1, 3))       ||   // MLK Day        (3rd Mon Jan)
    eq(nthWD(2, 1, 3))       ||   // Presidents Day (3rd Mon Feb)
    eq(lastWD(5, 1))         ||   // Memorial Day   (last Mon May)
    eq(new Date(yr, 5, 19))  ||   // Juneteenth
    eq(new Date(yr, 6,  4))  ||   // Independence Day
    eq(nthWD(9, 1, 1))       ||   // Labor Day      (1st Mon Sep)
    eq(nthWD(10, 1, 2))      ||   // Columbus Day   (2nd Mon Oct)
    eq(new Date(yr, 10, 11)) ||   // Veterans Day
    eq(nthWD(11, 4, 4))      ||   // Thanksgiving   (4th Thu Nov)
    eq(new Date(yr, 11, 25))      // Christmas
  );
}

/** Wrapper for Utilities.formatDate with the project timezone default. */
function fmtDate(d, pattern) {
  return Utilities.formatDate(d, CONFIG.TIMEZONE, pattern || 'yyyy-MM-dd HH:mm z');
}

// ── Numbers & colors ────────────────────────────────────────

/** Round to 2 decimal places, returned as a number. */
function roundRVU(n) {
  return Math.round(n * 100) / 100;
}

/**
 * Red→yellow→green hex color for a value within [min, max].
 * min → red, midpoint → yellow, max → green. Used for heat-mapping.
 */
function rvuHeatColor(value, min, max) {
  if (max <= min) return '#ffffff';
  const t = (value - min) / (max - min);   // 0 = lowest, 1 = highest

  let r, g, b;
  if (t <= 0.5) {                            // red → yellow
    const u = t * 2;
    r = Math.round(234 + (251 - 234) * u);
    g = Math.round(67  + (188 - 67)  * u);
    b = Math.round(53  + (4   - 53)  * u);
  } else {                                   // yellow → green
    const u = (t - 0.5) * 2;
    r = Math.round(251 + (52  - 251) * u);
    g = Math.round(188 + (168 - 188) * u);
    b = Math.round(4   + (83  - 4)   * u);
  }
  const hex = (x) => x.toString(16).padStart(2, '0');
  return '#' + hex(r) + hex(g) + hex(b);
}

/**
 * Pivot column widths: label column fixed, all date columns fixed
 * so they never collapse regardless of zoom or content length.
 */
function setPivotColumnWidths(sheet, nCols) {
  try {
    sheet.setColumnWidth(1, CONFIG.LABEL_COL_WIDTH);
    for (let c = 2; c <= nCols; c++) sheet.setColumnWidth(c, CONFIG.DATE_COL_WIDTH);
  } catch (e) {
    Logger.log(`setPivotColumnWidths error: ${e}`);
  }
}

// ── Admin alerts & logging ──────────────────────────────────

/** A fresh, empty admin-alerts accumulator (populated during a run). */
function blankAlerts() {
  return {
    missingEmail:    false,   // no matching report email found
    noCSV:           false,   // email found but no CSV attachment
    aveRVUMissing:   false,   // aveRVU sheet not found
    shiftMismatches: [],      // shift names not in the aveRVU sheet
    newProviders:    [],      // providers seen for the first time
    changes:         []       // WorkRVU changes on existing rows
  };
}

/**
 * Appends log lines to the Run Log sheet, newest first, capped at 2000 rows.
 * Both columns are forced to plain text so lines starting with "=" are never
 * interpreted as formulas.
 */
function logToSheet(ss, logLines) {
  try {
    const sheet = getOrCreateSheet(ss, CONFIG.SHEETS.RUN_LOG);

    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, 2)
        .setValues([['Timestamp', 'Message']])
        .setFontWeight('bold')
        .setBackground(CONFIG.COLOR.HEADER_DARK)
        .setFontColor('white');
      sheet.setFrozenRows(1);
      sheet.setColumnWidth(1, 175);
      sheet.setColumnWidth(2, 650);
      sheet.getRange('A:B').setNumberFormat('@');
    }

    if (!logLines || !logLines.length) return;

    const ts   = fmtDate(new Date(), 'yyyy-MM-dd HH:mm:ss');
    const rows = logLines.map(msg => [ts, msg]);

    sheet.getRange('A:B').setNumberFormat('@');
    sheet.insertRowsBefore(2, rows.length);       // newest at the top
    sheet.getRange(2, 1, rows.length, 2).setValues(rows);

    const maxRows = 2000;
    const total   = sheet.getLastRow() - 1;        // exclude header
    if (total > maxRows) sheet.deleteRows(maxRows + 2, total - maxRows);

  } catch (e) {
    Logger.log(`logToSheet error: ${e}`);
  }
}
