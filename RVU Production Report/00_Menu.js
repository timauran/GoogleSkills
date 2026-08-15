// ============================================================
// 00_Menu — spreadsheet menu + thin action wrappers
// ============================================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('⚙️ RVU Report')
    .addItem('▶ Run Now (full pipeline)', 'runNow')
    .addSeparator()
    .addSubMenu(ui.createMenu('📥 Ingest Data')
      .addItem('Scan Email & Ingest Now', 'scanEmailAndIngest')
      .addItem('Ingest "data" Sheet',     'ingestImportedSheet'))
    .addSeparator()
    .addSubMenu(ui.createMenu('📊 Rebuild Sheets')
      .addItem('Rebuild All Pivot Sheets', 'rebuildAllPivots')
      .addItem('Rebuild Total RVU Only',   'rebuildTotalRVU')
      .addItem('Rebuild Not Scheduled',    'rebuildNotScheduled')
      .addItem('Rebuild Productivity',     'rebuildThreshold')
      .addItem('Rebuild Monthly Summary',  'rebuildMonthly'))
    .addSeparator()
    .addSubMenu(ui.createMenu('📧 Email')
      .addItem('Send / Draft Today\'s Email Now', 'sendEmailNow')
      .addItem('Draft Email for a Prior Date…',   'draftEmailForDate')
      .addItem('Toggle Send ↔ Draft Mode',        'toggleSendMode'))
    .addSeparator()
    .addSubMenu(ui.createMenu('🔧 Setup & Admin')
      .addItem('Initial Setup (create all sheets)', 'createAndSetupSpreadsheet')
      .addItem('Install Daily Trigger',             'createDailyTrigger')
      .addItem('Clean / Dedupe Raw Data',           'cleanRawData')
      .addItem('Show Spreadsheet Info',             'showSpreadsheetInfo'))
    .addToUi();
}

// ── Rebuild shims — run a pivot sub-step without the full pipeline ──

function rebuildAllPivots() {
  const ss = getDataSS(), log = [];
  const allData   = readAllRawData(ss, log);
  const aveRVUMap = loadAveRVU(log, blankAlerts());
  updateTotalRVUSheet(ss, allData, log);
  updateNotScheduledSheet(ss, allData, log);
  updateThresholdSheet(ss, allData, aveRVUMap, log);
  updateMonthlySummary(ss, allData, log);
  logToSheet(ss, log);
  SpreadsheetApp.getUi().alert('All pivot sheets rebuilt.');
}

function rebuildTotalRVU() {
  const ss = getDataSS(), log = [];
  updateTotalRVUSheet(ss, readAllRawData(ss, log), log);
  logToSheet(ss, log);
  SpreadsheetApp.getUi().alert('Total RVU sheet rebuilt.');
}

function rebuildNotScheduled() {
  const ss = getDataSS(), log = [];
  updateNotScheduledSheet(ss, readAllRawData(ss, log), log);
  logToSheet(ss, log);
  SpreadsheetApp.getUi().alert('Not Scheduled sheet rebuilt.');
}

function rebuildThreshold() {
  const ss = getDataSS(), log = [];
  const aveRVUMap = loadAveRVU(log, blankAlerts());
  updateThresholdSheet(ss, readAllRawData(ss, log), aveRVUMap, log);
  logToSheet(ss, log);
  SpreadsheetApp.getUi().alert('Productivity sheet rebuilt.');
}

function rebuildMonthly() {
  const ss = getDataSS(), log = [];
  updateMonthlySummary(ss, readAllRawData(ss, log), log);
  logToSheet(ss, log);
  SpreadsheetApp.getUi().alert('Monthly Summary rebuilt.');
}

// ── Email shims ─────────────────────────────────────────────

function sendEmailNow() {
  const ss = getDataSS(), log = [];
  const allData   = readAllRawData(ss, log);
  const aveRVUMap = loadAveRVU(log, blankAlerts());
  createDailyEmailDraft(ss, allData, aveRVUMap, log);
  logToSheet(ss, log);
  const cfg = readEmailConfig(ss, log);
  SpreadsheetApp.getUi().alert(cfg.sendMode ? 'Email sent.' : 'Draft created in Gmail.');
}

function draftEmailForDate() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Draft Email for Prior Date',
    'Enter the date to report on (M/D/YYYY — e.g. 4/15/2026):',
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  const input = response.getResponseText().trim();
  if (!input) { ui.alert('No date entered.'); return; }

  const iso = toISO(input);
  if (!iso) {
    ui.alert(`Could not parse date: "${input}"\nPlease use M/D/YYYY, e.g. 4/15/2026.`);
    return;
  }

  const ss = getDataSS(), log = [];
  const allData   = readAllRawData(ss, log);
  const aveRVUMap = loadAveRVU(log, blankAlerts());

  const dayData = allData.filter(r => r.date === iso);
  if (!dayData.length) {
    ui.alert(`No data found for ${isoToUS(iso)}.\nCheck that Raw Data contains records for this date.`);
    logToSheet(ss, log);
    return;
  }

  // Always a DRAFT for historical dates, regardless of Send Mode.
  createDailyEmailDraft(ss, allData, aveRVUMap, log, iso, /* forceDraft= */ true);
  logToSheet(ss, log);
  ui.alert(`Draft created for ${isoToUS(iso)} (${dayData.length} records).\nCheck Gmail Drafts.`);
}

function toggleSendMode() {
  const ss    = getDataSS();
  const cfgSh = ss.getSheetByName(CONFIG.SHEETS.CONFIG_SHEET);
  if (!cfgSh) { SpreadsheetApp.getUi().alert('Config sheet not found. Run Initial Setup first.'); return; }

  const vals = cfgSh.getDataRange().getValues();
  for (let i = 1; i < vals.length; i++) {
    if (vals[i][0] && vals[i][0].toString().trim().toLowerCase() === 'send mode') {
      const next = !(vals[i][1] === true);
      cfgSh.getRange(i + 1, 2).setValue(next);
      SpreadsheetApp.getUi().alert(`Send Mode is now: ${next ? 'SEND (live)' : 'DRAFT'}`);
      return;
    }
  }
  SpreadsheetApp.getUi().alert('Send Mode row not found in Config sheet.');
}
