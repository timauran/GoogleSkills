// ============================================================
// 07_Engine — the nightly pipeline
// ============================================================

/**
 * Main entry point (runs from the daily trigger).
 * Finds the report email, updates Raw Data, rebuilds all pivots, and
 * drafts/sends the daily email. All alerts are consolidated into a
 * single admin email at the end.
 */
/**
 * Trigger handler. Runs under the central Email Automation Hub, which records
 * successes and errors (never empty scans), alerts if the daily run errors or stops
 * happening, and can pause it from the hub's Registry sheet. Falls back to a plain
 * call if the EmailHub library is not attached.
 */
function checkRVUEmail() {
  if (typeof EmailHub === 'undefined') return checkRVUEmailInner_();
  return EmailHub.runJob('rvu-daily-report', checkRVUEmailInner_, {
    project: 'RVU Production Report', handler: 'checkRVUEmail', intervalMin: 1440, phi: false
  });
}

/** Returns {processed, summary} on success, {errors, errorMessages} on failure, 0 when no new email. */
function checkRVUEmailInner_() {
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
    return { errors: 1, errorMessages: ['Cannot open data spreadsheet: ' + e] };
  }

  try {
    // 1 — find email & extract CSV
    const { content, thread } = findAndExtractCSV(log, alerts);
    if (!content) { sendAdminSummaryEmail(alerts, log); logToSheet(ss, log); return 0; }

    // 2 — parse CSV
    const { dataRows } = parseCSV(content, log);
    if (!dataRows.length) {
      log.push('ERROR: No valid data rows — aborting.');
      sendAdminSummaryEmail(alerts, log); logToSheet(ss, log);
      return { errors: 1, errorMessages: ['Report email had no valid data rows'] };
    }

    // 3 — load aveRVU reference & 4 — check shift mismatches
    const aveRVUMap = loadAveRVU(log, alerts);
    checkShiftMismatches(dataRows, aveRVUMap, log, alerts);

    // 5 — update Raw Data (canonical) & Provider List
    updateRawData(ss, dataRows, log, alerts);
    updateProviderList(ss, dataRows, log);

    // 6 — read canonical data & 7 — rebuild pivots
    ensureHolidaysSheet(ss, log);                 // keeps this + next year seeded
    const allData = readAllRawData(ss, log);
    updateTotalRVUSheet(ss, allData, log);
    updateNotScheduledSheet(ss, allData, log);
    updateThresholdSheet(ss, allData, aveRVUMap, log);
    updateMonthlySummary(ss, allData, log);

    // 8 — daily email (workdays only, never the same day twice) & 9 — admin summary
    const day = chooseReportDay(dataRows, log);
    if (day) {
      if (createDailyEmailDraft(ss, allData, aveRVUMap, log, day)) markReported(day, log);
    }
    sendAdminSummaryEmail(alerts, log);

    // 10 — trash the source report email
    if (thread) {
      try { thread.moveToTrash(); log.push('Source report email moved to trash.'); }
      catch (e) { log.push(`WARNING: Could not trash source email: ${e}`); }
    }

    log.push(`=== Run complete in ${((new Date() - runStart) / 1000).toFixed(1)}s ===`);
    logToSheet(ss, log);
    return { processed: dataRows.length, summary: `Ingested ${dataRows.length} RVU rows` };

  } catch (e) {
    log.push(`FATAL ERROR: ${e}`);
    log.push(`Stack: ${e.stack}`);
    GmailApp.sendEmail(CONFIG.MY_EMAIL, 'RVU Script FATAL ERROR', `${e}\n\n${e.stack}`);
    logToSheet(ss, log);
    return { errors: 1, errorMessages: [String(e)] };
  }
}

// ============================================================
// Report-day selection for the automatic daily email.
// ------------------------------------------------------------
// The day comes from the DATA, not the clock: the latest workday present
// in this report that has not already been emailed. So:
//   • a report containing only Saturday/Sunday/holiday data → no email
//   • Monday's report (Sun data) → no email; Friday went out on Saturday
//   • a report re-sending Friday alongside the weekend → Friday not re-sent
//   • today is never reported: rows are dated by SignDateTime, so reads
//     signed after midnight on last night's call carry today's date even
//     though the day has barely started
// ============================================================

/** Returns the ISO day to email for this batch, or null (reason logged). */
function chooseReportDay(dataRows, log) {
  const today = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd');
  const all   = [...new Set(dataRows.map(r => toISO(r.date)).filter(Boolean))].sort();
  const isos  = all.filter(i => i < today);
  if (isos.length < all.length) log.push(`Daily email: ignoring today/future date(s) ${all.filter(i => i >= today).map(isoToUS).join(', ')} — day not complete`);
  if (!isos.length) { log.push('Daily email: no completed dated rows — skipped'); return null; }

  const skipped = CONFIG.SKIP_NONWORKDAY_REPORTS ? isos.filter(isoIsWeekendOrHoliday) : [];
  const workdays = isos.filter(i => !skipped.includes(i));
  if (skipped.length) {
    log.push(`Daily email: not reporting weekend/holiday date(s) ${skipped.map(isoToUS).join(', ')}`);
  }

  const last   = getLastReported_();
  const fresh  = workdays.filter(i => i > last);
  if (!fresh.length) {
    log.push(workdays.length
      ? `Daily email: ${workdays.map(isoToUS).join(', ')} already reported (last: ${isoToUS(last)}) — skipped`
      : 'Daily email: batch has only weekend/holiday data — skipped');
    return null;
  }
  const day = fresh[fresh.length - 1];
  log.push(`Daily email: reporting ${isoToUS(day)}`);
  return day;
}

/**
 * Last emailed ISO date ('' if none). One-time repair: before today-dates
 * were excluded, the 11/3 and 11/4 runs emailed same-day stubs and pushed
 * the marker to 2026-11-04, which would block the real 11/3 and 11/4
 * emails. Roll it back to 11/2 (the last day correctly emailed) once.
 */
function getLastReported_() {
  const props = PropertiesService.getScriptProperties();
  let last = props.getProperty(CONFIG.PROP_LAST_REPORTED) || '';
  if (!props.getProperty('RVU_FIX_SAMEDAY_20261104')) {
    if (last === '2026-11-03' || last === '2026-11-04') {
      last = '2026-11-02';
      props.setProperty(CONFIG.PROP_LAST_REPORTED, last);
    }
    props.setProperty('RVU_FIX_SAMEDAY_20261104', 'done');
  }
  return last;
}

/** Records that `iso` has been emailed (only ever moves forward). */
function markReported(iso, log) {
  const props = PropertiesService.getScriptProperties();
  const last  = getLastReported_();
  if (iso > last) props.setProperty(CONFIG.PROP_LAST_REPORTED, iso);
  if (log) log.push(`Daily email: last reported date = ${isoToUS(iso > last ? iso : last)}`);
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
