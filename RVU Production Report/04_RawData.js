// ============================================================
// 04_RawData — canonical Raw Data engine
// ------------------------------------------------------------
// The Raw Data sheet is the single source of truth for daily RVU
// data. It is kept canonical on every run:
//   • Date stored as ISO "yyyy-MM-dd" text (no time-of-day)
//   • Exactly one row per Date + Provider + Shift (case-insensitive)
//   • Duplicates merged (highest WorkRVU wins)
// This self-healing normalize+dedupe is what fixes the historical
// "mixed date formats / uncaught duplicate" problem, and makes the
// sheet safe for other scripts to read (see getDailyRVU / getAllRVU).
// ============================================================

/** Case-insensitive composite key for a canonical record. */
function recordKey(rec) {
  return rec.iso + '|||'
       + String(rec.provider).trim().toLowerCase() + '|||'
       + String(rec.shift).trim().toLowerCase();
}

/** Ensure the Raw Data header row exists and column A is plain text. */
function ensureRawHeaders(sheet) {
  const headers = CONFIG.RAW_HEADERS;
  const first   = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const ok      = headers.every((h, i) => String(first[i]).trim() === h);
  if (!ok) {
    sheet.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setFontWeight('bold')
      .setBackground(CONFIG.COLOR.HEADER_DARK)
      .setFontColor('white');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 110);
  }
  sheet.getRange(1, 1, 1, headers.length).setNumberFormat('@');
}

/**
 * Reads the Raw Data sheet into canonical form, deduping as it goes.
 * Returns { records: Map<key, {iso,provider,shift,rvu}>,
 *           notes:   Map<key, [note,note,note,note]> }.
 * On duplicate keys the highest-RVU row wins; notes are merged.
 */
function readRawRecords(sheet) {
  const records = new Map();
  const notes   = new Map();
  const last    = sheet.getLastRow();
  if (last < 2) return { records, notes };

  const n     = last - 1;
  const vals  = sheet.getRange(2, 1, n, 4).getValues();
  const cellN = sheet.getRange(2, 1, n, 4).getNotes();

  for (let i = 0; i < n; i++) {
    const iso      = toISO(vals[i][0]);
    const provider = String(vals[i][1]).trim();
    const shift    = String(vals[i][2]).trim();
    const rvu      = parseFloat(vals[i][3]) || 0;
    if (!iso || !provider) continue;

    const key = recordKey({ iso, provider, shift });
    const ex  = records.get(key);
    if (!ex || rvu > ex.rvu) records.set(key, { iso, provider, shift, rvu });

    // Merge notes — keep the first non-empty note per column
    const cur = notes.get(key) || ['', '', '', ''];
    for (let c = 0; c < 4; c++) if (cellN[i][c] && !cur[c]) cur[c] = cellN[i][c];
    notes.set(key, cur);
  }
  return { records, notes };
}

/**
 * Rewrites the Raw Data body from canonical records, sorted newest date
 * first then provider A→Z. Applies merged notes and highlights the
 * WorkRVU cell of any key in changedKeys (yellow).
 */
function writeRawRecords(sheet, records, notes, changedKeys, log) {
  ensureRawHeaders(sheet);

  // Clear the existing data body (keep the header row)
  const last = sheet.getLastRow();
  if (last > 1) {
    sheet.getRange(2, 1, last - 1, 4).clearContent().clearNote().setBackground(null);
  }

  const keys = [...records.keys()].sort((a, b) => {
    const ra = records.get(a), rb = records.get(b);
    if (ra.iso !== rb.iso) return ra.iso < rb.iso ? 1 : -1;                 // date desc
    return ra.provider.toLowerCase() < rb.provider.toLowerCase() ? -1 : 1;  // provider asc
  });

  if (!keys.length) { if (log) log.push('Raw Data: no rows to write'); return; }

  const values = [], noteRows = [], changedRows = [];
  keys.forEach((k, i) => {
    const r = records.get(k);
    values.push([r.iso, r.provider, r.shift, roundRVU(r.rvu)]);
    noteRows.push(notes.get(k) || ['', '', '', '']);
    if (changedKeys && changedKeys.has(k)) changedRows.push(i + 2);   // 1-based sheet row
  });

  const rng = sheet.getRange(2, 1, values.length, 4);
  rng.setValues(values);
  rng.setNotes(noteRows);
  sheet.getRange(2, 1, values.length, 1).setNumberFormat('@');          // dates as text
  sheet.getRange(2, 4, values.length, 1).setNumberFormat('#,##0.00');   // WorkRVU

  // Highlight only the WorkRVU cells that changed this run (body was cleared above).
  changedRows.forEach(row => sheet.getRange(row, 4).setBackground(CONFIG.COLOR.CHANGE_CELL));

  if (log) log.push(`Raw Data: wrote ${values.length} canonical rows`);
}

/**
 * Merges the incoming CSV batch into Raw Data and rewrites it canonically.
 * Detects new providers and WorkRVU changes (recorded in adminAlerts).
 * Since the key is Date+Provider+Shift, the only field that can change on
 * an existing row is WorkRVU — date-format churn no longer looks like a change.
 */
function updateRawData(ss, incomingRows, log, adminAlerts) {
  const sheet = getOrCreateSheet(ss, CONFIG.SHEETS.RAW_DATA);
  ensureRawHeaders(sheet);

  const { records, notes } = readRawRecords(sheet);
  const existingProviders  = new Set([...records.values()].map(r => r.provider.toLowerCase()));
  log.push(`Raw Data: ${records.size} existing canonical rows`);

  // Pre-aggregate the incoming batch (dedupe, highest RVU per key)
  const incoming = new Map();
  incomingRows.forEach(r => {
    const iso      = toISO(r.date);
    const provider = String(r.provider).trim();
    const shift    = String(r.shift).trim();
    const rvu      = parseFloat(r.rvu) || 0;
    if (!iso || !provider) return;
    const key = recordKey({ iso, provider, shift });
    const ex  = incoming.get(key);
    if (!ex || rvu > ex.rvu) incoming.set(key, { iso, provider, shift, rvu });
  });

  const now          = fmtDate(new Date(), 'yyyy-MM-dd HH:mm z');
  const changedKeys  = new Set();
  const newProviders = new Set();
  let added = 0, updated = 0;

  incoming.forEach((rec, key) => {
    const prev = records.get(key);
    if (!prev) {
      records.set(key, rec);
      added++;
      if (!existingProviders.has(rec.provider.toLowerCase())) newProviders.add(rec.provider);
    } else if (roundRVU(prev.rvu) !== roundRVU(rec.rvu)) {
      const noteRow = notes.get(key) || ['', '', '', ''];
      const entry   = `${now}: was "${roundRVU(prev.rvu)}"`;
      noteRow[3]    = noteRow[3] ? noteRow[3] + '\n' + entry : entry;
      notes.set(key, noteRow);
      adminAlerts.changes.push({
        provider: rec.provider, date: rec.iso, shift: rec.shift,
        col: 'WorkRVU', oldVal: roundRVU(prev.rvu), newVal: roundRVU(rec.rvu)
      });
      records.set(key, rec);
      changedKeys.add(key);
      updated++;
    }
  });

  writeRawRecords(sheet, records, notes, changedKeys, log);

  if (newProviders.size) {
    adminAlerts.newProviders = [...newProviders].sort();
    log.push(`New providers: ${adminAlerts.newProviders.join(', ')}`);
  }
  log.push(`Raw Data: +${added} new, ${updated} RVU update(s), ${records.size} total after dedupe/normalize`);
}

/** Menu action: normalize dates and merge duplicate rows in place. */
function cleanRawData() {
  const ss    = getDataSS();
  const log   = [];
  const sheet = getOrCreateSheet(ss, CONFIG.SHEETS.RAW_DATA);
  ensureRawHeaders(sheet);

  const before = Math.max(0, sheet.getLastRow() - 1);
  const { records, notes } = readRawRecords(sheet);
  writeRawRecords(sheet, records, notes, new Set(), log);
  const removed = before - records.size;

  log.push(`Clean Raw Data: ${before} → ${records.size} rows (${removed} merged/removed)`);
  logToSheet(ss, log);
  SpreadsheetApp.getUi().alert(
    `Raw Data cleaned.\n\n` +
    `Before: ${before} rows\nAfter:  ${records.size} rows\n` +
    `Merged/removed: ${removed} duplicate row(s)\n\n` +
    `All dates normalized to yyyy-MM-dd.`
  );
}

/**
 * Reads canonical Raw Data as objects { date, provider, shift, rvu }
 * (date is ISO "yyyy-MM-dd"), filtered to included providers and deduped.
 * `log` is optional.
 */
function readAllRawData(ss, log) {
  const sheet = ss.getSheetByName(CONFIG.SHEETS.RAW_DATA);
  if (!sheet || sheet.getLastRow() < 2) {
    if (log) log.push('readAllRawData: sheet empty or missing');
    return [];
  }

  const n        = sheet.getLastRow() - 1;
  const vals     = sheet.getRange(2, 1, n, 4).getValues();
  const included = getIncludedProviders(ss, log);

  const map = new Map();
  vals.forEach(row => {
    const iso      = toISO(row[0]);
    const provider = String(row[1]).trim();
    const shift    = String(row[2]).trim();
    const rvu      = parseFloat(row[3]) || 0;
    if (!iso || !provider) return;
    if (included && !included.has(provider)) return;
    const key = recordKey({ iso, provider, shift });
    const ex  = map.get(key);
    if (!ex || rvu > ex.rvu) map.set(key, { date: iso, provider, shift, rvu });
  });

  const rows = [...map.values()];
  if (log) log.push(`readAllRawData: ${rows.length} rows after provider filter and dedup`);
  return rows;
}

// ── PUBLIC API — for use by other Apps Script projects ──────

/**
 * Returns the daily RVU rows for a single calendar day.
 * @param {string|Date} dateInput  'yyyy-MM-dd', 'M/D/YYYY', or a Date.
 * @return {{date:string, provider:string, shift:string, rvu:number}[]}
 *         ISO-dated rows (empty array if the date is unparseable or has no data).
 */
function getDailyRVU(dateInput) {
  const iso = toISO(dateInput);
  if (!iso) return [];
  return readAllRawData(getDataSS(), null).filter(r => r.date === iso);
}

/** Returns all canonical RVU rows (ISO-dated, provider-filtered, deduped). */
function getAllRVU() {
  return readAllRawData(getDataSS(), null);
}

// ============================================================
// PROVIDER LIST SHEET
// One row per unique provider; unchecked providers are excluded from
// all pivots, the email, and the monthly summary.
// Columns: Provider | Include (checkbox) | Notes
// ============================================================

function updateProviderList(ss, dataRows, log) {
  const sheet      = getOrCreateSheet(ss, CONFIG.SHEETS.PROVIDER_LIST);
  const HEADERS    = ['Provider', 'Include', 'Notes'];
  const DATA_START = 2;

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length)
      .setValues([HEADERS])
      .setFontWeight('bold')
      .setBackground(CONFIG.COLOR.HEADER_DARK)
      .setFontColor('white');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 200);
    sheet.setColumnWidth(2, 70);
    sheet.setColumnWidth(3, 300);
    log.push('Provider List: sheet initialised');
  }

  const lastRow  = sheet.getLastRow();
  const existing = {};
  if (lastRow >= DATA_START) {
    sheet.getRange(DATA_START, 1, lastRow - 1, 1).getValues().forEach((r, i) => {
      const name = r[0].toString().trim();
      if (name) existing[name] = DATA_START + i;
    });
  }

  const seen  = new Set(dataRows.map(r => r.provider).filter(Boolean));
  const toAdd = [...seen].filter(p => !existing[p]).sort();

  if (toAdd.length) {
    const insertRow = sheet.getLastRow() + 1;
    sheet.getRange(insertRow, 1, toAdd.length, 3).setValues(toAdd.map(p => [p, true, '']));
    sheet.getRange(insertRow, 2, toAdd.length, 1)
      .setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
    log.push(`Provider List: added ${toAdd.length} new provider(s): ${toAdd.join(', ')}`);
  } else {
    log.push('Provider List: no new providers');
  }

  // Idempotently ensure the whole Include column has checkbox validation
  const finalRow = sheet.getLastRow();
  if (finalRow >= DATA_START) {
    sheet.getRange(DATA_START, 2, finalRow - 1, 1)
      .setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
  }
}

/**
 * Returns a Set of provider names with Include = TRUE, or null (meaning
 * "include everyone") if the sheet is missing or empty. `log` is optional.
 */
function getIncludedProviders(ss, log) {
  const sheet = ss.getSheetByName(CONFIG.SHEETS.PROVIDER_LIST);
  if (!sheet || sheet.getLastRow() < 2) {
    if (log) log.push('Provider List: sheet missing or empty — including all providers');
    return null;
  }
  const vals = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  const included = new Set(vals.filter(r => r[0] && r[1] === true).map(r => r[0].toString().trim()));
  if (log) log.push(`Provider List: ${included.size} provider(s) included in reporting`);
  return included;
}
