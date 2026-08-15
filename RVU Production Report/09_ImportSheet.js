// ============================================================
// 09_ImportSheet — ingest an imported sheet into Raw Data
// ------------------------------------------------------------
// For backfills and re-imports: instead of waiting for the daily
// "[SECURE] RVU Comp Report" email, drop a CSV/XLSX into the data
// spreadsheet (File → Import → Insert new sheet), rename that tab to
// CONFIG.IMPORT_SHEET ("data"), and run this.
//
// The imported tab is read with the same rules as the emailed CSV
// (header-driven, weekend shifts skipped), then merged through
// updateRawData(), so the existing canonical behaviour applies:
//   • key = Date + Provider + Shift (case-insensitive)
//   • rows not already present are added
//   • rows whose WorkRVU differs are REPLACED — the prior value is
//     kept as a cell note and the cell is highlighted yellow
//   • rows that match exactly are left alone
// The imported tab itself is never modified.
// ============================================================

/** Rows scanned from the top of an imported sheet while hunting for the header row. */
const IMPORT_HEADER_SCAN_ROWS = 10;

/** True if `name` is one of this project's own sheets (never an import target). */
function isSystemSheet(name) {
  return Object.keys(CONFIG.SHEETS).some(k => CONFIG.SHEETS[k] === name);
}

/** Sheet names in `ss` that could plausibly be an imported data tab. */
function listImportCandidates(ss) {
  return ss.getSheets().map(s => s.getName()).filter(n => !isSystemSheet(n));
}

/**
 * Locates the header row in an imported sheet and maps our four logical
 * fields onto its column indices.
 * Returns { headerRow, col:{date,provider,shift,rvu}, headers } or null.
 */
function findImportHeaders(values, log) {
  const limit = Math.min(IMPORT_HEADER_SCAN_ROWS, values.length);

  for (let r = 0; r < limit; r++) {
    const headers = values[r].map(h => String(h).trim());
    const { col, missing } = resolveColumns(headers);      // 02_Utils — same aliases as the CSV

    if (!missing.length) {
      if (col.shift === -1) log.push('Import: no Shift column found — shift will be blank');
      log.push(`Import: header row ${r + 1} → [${headers.filter(String).join(' | ')}]`);
      return { headerRow: r, col, headers };
    }
  }

  log.push('ERROR: Import sheet has no recognizable header row '
         + `(need date, provider and WorkRVU columns in the first ${limit} rows). `
         + 'Add the new header name to CONFIG.COL_ALIASES.');
  return null;
}

/**
 * Reads an imported sheet into parseCSV-shaped rows: { date, provider, shift, rvu }.
 * Dates are left raw here; updateRawData canonicalizes them to ISO.
 */
function parseImportSheet(sheet, log) {
  const last = sheet.getLastRow();
  if (last < 2) {
    log.push(`ERROR: Import sheet "${sheet.getName()}" is empty.`);
    return { dataRows: [] };
  }

  const values = sheet.getRange(1, 1, last, sheet.getLastColumn()).getValues();
  const found  = findImportHeaders(values, log);
  if (!found) return { dataRows: [] };

  const { headerRow, col } = found;
  const dataRows = [];
  let skipped = 0;

  for (let i = headerRow + 1; i < values.length; i++) {
    const r = values[i];
    if (!r || r.every(c => String(c).trim() === '')) { skipped++; continue; }

    // Sheets often imports the date column as a real Date — pass it through
    // untouched so toISO() formats it in the project timezone.
    const raw      = col.date === -1 ? '' : r[col.date];
    const date     = raw instanceof Date ? raw : cellAt(r, col.date);
    const provider = cellAt(r, col.provider);
    const shift    = cellAt(r, col.shift);
    const rvu      = parseFloat(r[col.rvu]) || 0;

    if (!date || !provider)   { skipped++; continue; }
    if (!toISO(date))         { log.push(`  Row ${i + 1} skipped: unparseable date "${date}"`); skipped++; continue; }
    if (isWeekendShift(shift)) { skipped++; continue; }

    dataRows.push({ date, provider, shift, rvu });
  }

  log.push(`Import parsed: ${dataRows.length} valid rows, ${skipped} skipped`);
  return { dataRows };
}

/**
 * Core routine — ingest a named sheet into Raw Data and rebuild everything
 * downstream. Callable headlessly (e.g. from another script or a trigger):
 *   ingestImportedSheetByName('data')
 * Returns a summary object; throws nothing on data problems, only reports.
 */
function ingestImportedSheetByName(sheetName) {
  const ss     = getDataSS();
  const log    = [];
  const alerts = blankAlerts();
  log.push(`=== Import ingest start: "${sheetName}" ===`);

  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    log.push(`ERROR: Sheet "${sheetName}" not found.`);
    logToSheet(ss, log);
    return { ok: false, message: `Sheet "${sheetName}" not found.` };
  }
  if (isSystemSheet(sheetName)) {
    log.push(`ERROR: "${sheetName}" is a system sheet — refusing to ingest it.`);
    logToSheet(ss, log);
    return { ok: false, message: `"${sheetName}" is one of the report's own sheets, not an import.` };
  }

  const { dataRows } = parseImportSheet(sheet, log);
  if (!dataRows.length) {
    logToSheet(ss, log);
    return { ok: false, message: `No usable rows found in "${sheetName}". See the Run Log for details.` };
  }

  const r = ingestDataRows(ss, dataRows, log, alerts);   // 07_Engine — shared with the email scan

  const summary = {
    ok:       true,
    sheet:    sheetName,
    parsed:   r.parsed,
    added:    r.added,
    replaced: r.replaced,
    total:    r.total,
    newProviders:    alerts.newProviders,
    shiftMismatches: alerts.shiftMismatches
  };

  log.push(`Import ingest done: +${summary.added} new, ${summary.replaced} replaced, ${summary.total} total rows`);
  logToSheet(ss, log);
  return summary;
}

/**
 * Menu action: ingest the "data" tab and report what changed.
 * If that tab is missing, falls back to asking which sheet to use.
 */
function ingestImportedSheet() {
  const ui   = SpreadsheetApp.getUi();
  const ss   = getDataSS();
  let   name = CONFIG.IMPORT_SHEET;

  if (!ss.getSheetByName(name)) {
    const candidates = listImportCandidates(ss);
    if (!candidates.length) {
      ui.alert(
        `No "${CONFIG.IMPORT_SHEET}" sheet found`,
        'Add the data first: File → Import → Upload your RVU file → '
        + `"Insert new sheet(s)", then rename that tab to "${CONFIG.IMPORT_SHEET}". `
        + 'Then run this again.',
        ui.ButtonSet.OK
      );
      return;
    }

    const response = ui.prompt(
      'Ingest Imported Sheet',
      `No "${CONFIG.IMPORT_SHEET}" tab found.\n\nAvailable: ${candidates.join(', ')}\n\n`
      + `Sheet name (blank = "${candidates[0]}"):`,
      ui.ButtonSet.OK_CANCEL
    );
    if (response.getSelectedButton() !== ui.Button.OK) return;
    name = response.getResponseText().trim() || candidates[0];
  }

  const result = ingestImportedSheetByName(name);
  if (!result.ok) { ui.alert('Import failed', result.message, ui.ButtonSet.OK); return; }

  let msg = `Ingested "${result.sheet}".\n\n`
          + `Rows read:        ${result.parsed}\n`
          + `New rows added:   ${result.added}\n`
          + `Rows replaced:    ${result.replaced}  (old WorkRVU kept as a cell note, cell highlighted)\n`
          + `Raw Data total:   ${result.total}\n\n`
          + 'Pivot sheets rebuilt.';
  if (result.newProviders.length)    msg += `\n\nNew provider(s): ${result.newProviders.join(', ')}`;
  if (result.shiftMismatches.length) msg += `\n\nShift(s) not in aveRVU: ${result.shiftMismatches.join(', ')}`;
  msg += `\n\nThe "${result.sheet}" tab was not modified — delete it when you are done.`;

  ui.alert('Import complete', msg, ui.ButtonSet.OK);
}
