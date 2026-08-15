// ============================================================
// 08_Setup — one-time setup, triggers, manual runners
// ============================================================

/**
 * SETUP STEP 1 — run once.
 * Creates the data spreadsheet, initialises every working sheet, and
 * stores the spreadsheet ID in Script Properties.
 */
function createAndSetupSpreadsheet() {
  const S = CONFIG.SHEETS;
  const SHEET_NAMES = [
    S.RAW_DATA, S.TOTAL_RVU, S.NOT_SCHEDULED, S.THRESHOLD, S.MONTHLY_SUMMARY,
    S.PROVIDER_LIST, S.EMAIL_LIST, S.CONFIG_SHEET, S.RUN_LOG
  ];

  const ss  = SpreadsheetApp.create('RVU Production Report');
  const id  = ss.getId();
  const url = ss.getUrl();

  ss.getActiveSheet().setName(SHEET_NAMES[0]);
  SHEET_NAMES.slice(1).forEach(name => ss.insertSheet(name));

  // Raw Data — canonical schema
  const raw = ss.getSheetByName(S.RAW_DATA);
  raw.getRange(1, 1, 1, CONFIG.RAW_HEADERS.length)
    .setValues([CONFIG.RAW_HEADERS])
    .setFontWeight('bold').setBackground(CONFIG.COLOR.HEADER_DARK).setFontColor('white');
  raw.getRange('A:A').setNumberFormat('@');   // ISO dates as text
  raw.setFrozenRows(1);
  raw.setColumnWidth(1, 110);

  // Run Log
  const logSheet = ss.getSheetByName(S.RUN_LOG);
  logSheet.getRange(1, 1, 1, 2)
    .setValues([['Timestamp', 'Message']])
    .setFontWeight('bold').setBackground(CONFIG.COLOR.HEADER_DARK).setFontColor('white');
  logSheet.setFrozenRows(1);
  logSheet.setColumnWidth(1, 175);
  logSheet.setColumnWidth(2, 650);

  // Pivot placeholders
  [S.TOTAL_RVU, S.NOT_SCHEDULED, S.THRESHOLD].forEach(name => {
    const sh = ss.getSheetByName(name);
    sh.getRange(1, 1).setValue('Rad \\ Date')
      .setFontWeight('bold')
      .setBackground(name === S.NOT_SCHEDULED ? CONFIG.COLOR.HEADER_NS
                   : name === S.THRESHOLD     ? CONFIG.COLOR.HEADER_PROD
                   :                            CONFIG.COLOR.HEADER_DARK)
      .setFontColor('white');
    sh.setFrozenRows(1);
    sh.setFrozenColumns(1);
    sh.setColumnWidth(1, CONFIG.LABEL_COL_WIDTH);
  });

  // Monthly Summary
  const ms = ss.getSheetByName(S.MONTHLY_SUMMARY);
  ms.getRange(1, 1, 1, 5)
    .setValues([['Month', 'Provider', 'Total RVU', 'Scheduled RVU', 'Not Scheduled RVU']])
    .setFontWeight('bold').setBackground(CONFIG.COLOR.HEADER_DARK).setFontColor('white');
  ms.setFrozenRows(1);

  // Email List
  const el = ss.getSheetByName(S.EMAIL_LIST);
  el.getRange(1, 1, 1, 3)
    .setValues([['Name', 'Email', 'Role (To / CC / BCC)']])
    .setFontWeight('bold').setBackground(CONFIG.COLOR.HEADER_DARK).setFontColor('white');
  el.setFrozenRows(1);
  el.setColumnWidth(1, 180);
  el.setColumnWidth(2, 240);
  el.setColumnWidth(3, 160);
  el.getRange(2, 1, 1, 3).setValues([['Dr. Example', 'example@ra-slo.com', 'To']])
    .setFontColor('#999999').setFontStyle('italic');

  // Config
  const cfg = ss.getSheetByName(S.CONFIG_SHEET);
  cfg.getRange(1, 1, 1, 3)
    .setValues([['Setting', 'Value', 'Description']])
    .setFontWeight('bold').setBackground(CONFIG.COLOR.HEADER_DARK).setFontColor('white');
  cfg.setFrozenRows(1);
  cfg.setColumnWidth(1, 160);
  cfg.setColumnWidth(2, 100);
  cfg.setColumnWidth(3, 380);
  cfg.getRange(2, 1, 1, 3).setValues([['Send Mode', false, 'TRUE = send email automatically  |  FALSE = save as draft']]);
  cfg.getRange(2, 2).setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
  cfg.getRange(2, 1).setFontWeight('bold');

  PropertiesService.getScriptProperties().setProperty(CONFIG.PROP_SPREADSHEET_ID, id);

  Logger.log('='.repeat(60));
  Logger.log('✅ RVU Production Report spreadsheet created.');
  Logger.log(`   ID  : ${id}`);
  Logger.log(`   URL : ${url}`);
  Logger.log(`   Sheets: ${SHEET_NAMES.join(', ')}`);
  Logger.log(`   ID stored in Script Properties as "${CONFIG.PROP_SPREADSHEET_ID}"`);
  Logger.log('Next: run createDailyTrigger(), then runNow() to test.');
  Logger.log('='.repeat(60));
}

/**
 * SETUP STEP 2 — run once.
 * Installs a daily trigger for checkRVUEmail() at ~3:30 AM Pacific.
 */
function createDailyTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'checkRVUEmail')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('checkRVUEmail')
    .timeBased().atHour(3).nearMinute(30).everyDays(1).create();

  Logger.log('✅ Trigger created: checkRVUEmail daily at ~3:30 AM (America/Los_Angeles)');
}

/** Manual test runner — process immediately. */
function runNow() {
  checkRVUEmail();
}

/** Prints the data spreadsheet ID/URL/sheets to the execution log. */
function showSpreadsheetInfo() {
  const id = PropertiesService.getScriptProperties().getProperty(CONFIG.PROP_SPREADSHEET_ID);
  if (!id) { Logger.log('⚠️  No data spreadsheet set. Run createAndSetupSpreadsheet() first.'); return; }
  try {
    const ss = SpreadsheetApp.openById(id);
    Logger.log(`Data spreadsheet : "${ss.getName()}"`);
    Logger.log(`ID               : ${id}`);
    Logger.log(`URL              : ${ss.getUrl()}`);
    Logger.log(`Sheets           : ${ss.getSheets().map(s => s.getName()).join(', ')}`);
  } catch (e) {
    Logger.log(`❌ Cannot open spreadsheet ID "${id}": ${e}`);
  }
}
