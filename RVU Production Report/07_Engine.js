// ============================================================
// 07_Engine — the nightly pipeline
// ============================================================

/**
 * Main entry point (runs from the daily trigger).
 * Finds the report email, updates Raw Data, rebuilds all pivots, and
 * drafts/sends the daily email. All alerts are consolidated into a
 * single admin email at the end.
 */
function checkRVUEmail() {
  const log      = [];
  const runStart = new Date();
  const alerts   = blankAlerts();
  log.push(`=== RVU Run Start: ${fmtDate(runStart, 'yyyy-MM-dd HH:mm:ss z')} ===`);

  let ss;
  try {
    ss = getDataSS();
    log.push(`Opened data spreadsheet: "${ss.getName()}" (${ss.getId()})`);
  } catch (e) {
    Logger.log(`FATAL: Cannot open data spreadsheet — ${e}`);
    GmailApp.sendEmail(CONFIG.MY_EMAIL, 'RVU Script FATAL: Cannot open data spreadsheet',
      `${e}\n\nHave you run createAndSetupSpreadsheet() yet?`);
    return;
  }

  try {
    // 1 — find email & extract CSV
    const { content, thread } = findAndExtractCSV(log, alerts);
    if (!content) { sendAdminSummaryEmail(alerts, log); logToSheet(ss, log); return; }

    // 2 — parse CSV
    const { dataRows } = parseCSV(content, log);
    if (!dataRows.length) {
      log.push('ERROR: No valid data rows — aborting.');
      sendAdminSummaryEmail(alerts, log); logToSheet(ss, log); return;
    }

    // 3 — load aveRVU reference & 4 — check shift mismatches
    const aveRVUMap = loadAveRVU(log, alerts);
    checkShiftMismatches(dataRows, aveRVUMap, log, alerts);

    // 5 — update Raw Data (canonical) & Provider List
    updateRawData(ss, dataRows, log, alerts);
    updateProviderList(ss, dataRows, log);

    // 6 — read canonical data & 7 — rebuild pivots
    const allData = readAllRawData(ss, log);
    updateTotalRVUSheet(ss, allData, log);
    updateNotScheduledSheet(ss, allData, log);
    updateThresholdSheet(ss, allData, aveRVUMap, log);
    updateMonthlySummary(ss, allData, log);

    // 8 — daily email & 9 — admin summary
    createDailyEmailDraft(ss, allData, aveRVUMap, log);
    sendAdminSummaryEmail(alerts, log);

    // 10 — trash the source report email
    if (thread) {
      try { thread.moveToTrash(); log.push('Source report email moved to trash.'); }
      catch (e) { log.push(`WARNING: Could not trash source email: ${e}`); }
    }

    log.push(`=== Run complete in ${((new Date() - runStart) / 1000).toFixed(1)}s ===`);

  } catch (e) {
    log.push(`FATAL ERROR: ${e}`);
    log.push(`Stack: ${e.stack}`);
    GmailApp.sendEmail(CONFIG.MY_EMAIL, 'RVU Script FATAL ERROR', `${e}\n\n${e.stack}`);
  }

  logToSheet(ss, log);
}

// ============================================================
// Shared ingest tail — used by the manual email scan (below) and by
// the imported-sheet ingest (09_ImportSheet). Merges parsed rows into
// Raw Data, refreshes the Provider List, and rebuilds every sheet that
// reads Raw Data. Does NOT send any email.
// Returns { parsed, added, replaced, total }.
// ============================================================
function ingestDataRows(ss, dataRows, log, alerts) {
  const rawSheet = getOrCreateSheet(ss, CONFIG.SHEETS.RAW_DATA);
  const before   = readRawRecords(rawSheet).records.size;

  const aveRVUMap = loadAveRVU(log, alerts);
  checkShiftMismatches(dataRows, aveRVUMap, log, alerts);

  updateRawData(ss, dataRows, log, alerts);      // adds new, replaces differing WorkRVU
  updateProviderList(ss, dataRows, log);

  const after = readRawRecords(rawSheet).records.size;

  const allData = readAllRawData(ss, log);
  updateTotalRVUSheet(ss, allData, log);
  updateNotScheduledSheet(ss, allData, log);
  updateThresholdSheet(ss, allData, aveRVUMap, log);
  updateMonthlySummary(ss, allData, log);

  return {
    parsed:   dataRows.length,
    added:    after - before,
    replaced: alerts.changes.length,
    total:    after
  };
}

/**
 * Menu action: scan Gmail for the report email and ingest it now.
 * Unlike Run Now, this stops after Raw Data + pivots — no daily email,
 * no admin summary, and the source thread is labelled but not trashed,
 * so it stays available to re-read manually if needed.
 */
function scanEmailAndIngest() {
  const ui     = SpreadsheetApp.getUi();
  const ss     = getDataSS();
  const log    = [];
  const alerts = blankAlerts();
  log.push('=== Manual email scan & ingest ===');

  const { content } = findAndExtractCSV(log, alerts);
  if (!content) {
    logToSheet(ss, log);
    ui.alert('Nothing to ingest',
      alerts.noCSV
        ? 'The report email was found but had no CSV attachment.'
        : 'No unprocessed report email found in the last 2 days.\n\n'
          + `(Threads already labelled "${CONFIG.PROCESSED_LABEL}" are skipped. `
          + 'Remove that label to re-ingest one.)',
      ui.ButtonSet.OK);
    return;
  }

  const { dataRows } = parseCSV(content, log);
  if (!dataRows.length) {
    logToSheet(ss, log);
    ui.alert('Nothing to ingest', 'The CSV had no valid data rows. See the Run Log.', ui.ButtonSet.OK);
    return;
  }

  const r = ingestDataRows(ss, dataRows, log, alerts);
  log.push(`Manual email ingest: +${r.added} new, ${r.replaced} replaced, ${r.total} total rows`);
  logToSheet(ss, log);

  let msg = `Ingested from the report email.\n\n`
          + `Rows read:      ${r.parsed}\n`
          + `New rows added: ${r.added}\n`
          + `Rows replaced:  ${r.replaced}  (old WorkRVU kept as a cell note, cell highlighted)\n`
          + `Raw Data total: ${r.total}\n\n`
          + 'Pivot sheets rebuilt. No email was sent.';
  if (alerts.newProviders.length)    msg += `\n\nNew provider(s): ${alerts.newProviders.join(', ')}`;
  if (alerts.shiftMismatches.length) msg += `\n\nShift(s) not in aveRVU: ${alerts.shiftMismatches.join(', ')}`;

  ui.alert('Email ingest complete', msg, ui.ButtonSet.OK);
}
