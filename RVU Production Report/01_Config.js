// ============================================================
// RVU PRODUCTION REPORT — Google Apps Script
// Version 2.0  |  tauran@ra-slo.com
// ============================================================
// Processes the daily "[SECURE] RVU Comp Report" email, maintains a
// canonical Raw Data sheet, rebuilds pivot/summary sheets, and drafts
// or sends the daily production email.
//
// FILE MAP
//   01_Config     CONFIG constants (this file)
//   02_Utils      shared helpers: dates, colors, logging, spreadsheet access
//   03_Ingest     Gmail fetch, CSV parse, aveRVU load, shift-mismatch check
//   04_RawData    canonical Raw Data engine (+ Provider List, public accessors)
//   05_Pivots     Total RVU / Not Scheduled / Productivity / Monthly Summary
//   06_Email      daily production email + admin summary email
//   07_Engine     checkRVUEmail() — the nightly pipeline
//   08_Setup      one-time setup, triggers, manual runners
//   09_ImportSheet ingest a manually imported sheet into Raw Data
//   00_Menu       spreadsheet menu + thin action wrappers
//
// FIRST-TIME SETUP (run in this order from the Apps Script editor):
//   1. Project Settings → Time zone: America/Los_Angeles
//   2. createAndSetupSpreadsheet()   creates the data spreadsheet + sheets
//   3. createDailyTrigger()          installs the 3:30 AM cron
//   4. runNow()                      process immediately as a test
//
// RAW DATA CONTRACT (for other scripts):
//   Sheet "Raw Data" columns: Date | Provider | Shift | WorkRVU
//   Date is always canonical ISO text "yyyy-MM-dd" (no time-of-day).
//   One row per unique Date + Provider + Shift (deduped every run).
//   Other scripts should read it via getDailyRVU() / getAllRVU() (see 04_RawData).
// ============================================================

const CONFIG = {
  // ── aveRVU reference spreadsheet (read-only) ─────────────────────
  // Sheet "aveRVU" header: Shift | AveRVU | Count | Start Date | End Date.
  // Only col A (Shift) and col B (AveRVU) are read; the rest are ignored.
  AVE_RVU_SPREADSHEET_ID: '1rThkXpgdC9Vo1tYD3sWK30akVhNj8DueX2ZEZoV7Gcs',

  // Data spreadsheet ID is stored in Script Properties under this key
  // by createAndSetupSpreadsheet(). Not hard-coded here.
  PROP_SPREADSHEET_ID: 'RVU_SPREADSHEET_ID',

  MY_EMAIL:        'tauran@ra-slo.com',
  // Matched as a Gmail subject: prefix — the report's subject line has
  // changed wording after this point before, so keep it short.
  SUBJECT_PATTERN: '[SECURE] RVU Comp Report',
  TIMEZONE:        'America/Los_Angeles',

  // Gmail label applied to threads after successful processing so the
  // same report is never picked up twice. Auto-created on first run.
  PROCESSED_LABEL: 'RVU-Processed',

  SHEETS: {
    RAW_DATA:        'Raw Data',
    TOTAL_RVU:       'Total RVU',
    NOT_SCHEDULED:   'Not Scheduled',
    THRESHOLD:       'Productivity',
    AVE_RVU:         'aveRVU',
    RUN_LOG:         'Run Log',
    MONTHLY_SUMMARY: 'Monthly Summary',
    PROVIDER_LIST:   'Provider List',
    EMAIL_LIST:      'Email List',
    CONFIG_SHEET:    'Config'
  },

  // Column-name aliases, matched case-insensitively after trimming, first
  // match wins. Used for BOTH the emailed CSV and a manually imported
  // sheet (see resolveColumns in 02_Utils). The report has renamed its
  // date column more than once — add new spellings to the front of the
  // list rather than replacing the old ones, so historical files still
  // ingest. Extra columns (e.g. ShiftAvg_90Day) are ignored: averages
  // come from the aveRVU sheet, not the report.
  COL_ALIASES: {
    date:     ['SignDateTime', 'ShiftStartDateTime2', 'ShiftStartDateTime', 'ShiftStartDate', 'Date'],
    provider: ['RenderingProvider1', 'RenderingProvider', 'Provider'],
    shift:    ['Shift'],
    rvu:      ['WorkRVU', 'Work RVU', 'RVU']
  },

  // Canonical Raw Data sheet schema. Column order is fixed; downstream
  // scripts rely on it. Date is stored as ISO "yyyy-MM-dd" text.
  RAW_HEADERS: ['Date', 'Provider', 'Shift', 'WorkRVU'],

  // Tab name a manually imported RVU file is renamed to before ingesting
  // (File → Import → Insert new sheet, then rename it to this).
  // Deliberately NOT in SHEETS — it is data in transit, not a report sheet.
  IMPORT_SHEET: 'data',

  NOT_SCHEDULED_KEY: 'not scheduled',   // compared case-insensitively after trim
  THRESHOLD_GREEN:   120,               // >= this % → green
  THRESHOLD_ORANGE:  75,                // <  this % → orange

  // Rotations whose RVU is never colour-coded in the daily email: their
  // volume is not comparable to the shift average, so a green/orange call
  // would be misleading. Matched as a case-insensitive substring of the
  // shift name. Not Scheduled lines are exempt regardless of this list.
  NO_HIGHLIGHT_SHIFTS: ['PM Imaging'],

  // Pivot sheet column widths (px)
  DATE_COL_WIDTH:  72,
  LABEL_COL_WIDTH: 160,

  COLOR: {
    HEADER_DARK:  '#1a5276',
    HEADER_MED:   '#2e86c1',
    HEADER_NS:    '#b45f06',
    HEADER_PROD:  '#274e13',
    ROW_ALT:      '#eaf3fb',
    ROW_ALT_NS:   '#fef9f5',
    ROW_ALT_PD:   '#f0f4f0',
    SUBTOTAL:     '#d6e4f0',
    SUBTOTAL_NS:  '#fce5cd',
    SUBTOTAL_PD:  '#d9ead3',
    GRAND_TOTAL:  '#c8d8e8',
    WEEKEND_GRAY: '#e8eaed',   // non-work day column data rows
    WEEKEND_HDR:  '#9aa0a6',   // non-work day column header row
    CHANGE_CELL:  '#FFFF00',
    GREEN_PCT:    '#b7e1cd',
    ORANGE_PCT:   '#f9cb9c'
  }
};
