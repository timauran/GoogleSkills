// ============================================================
// 03_Ingest — Gmail fetch, CSV parse, aveRVU load, mismatch check
// ============================================================

/**
 * Finds the newest unprocessed report email, extracts its CSV attachment,
 * and labels the thread as processed. Returns { content, thread }.
 * content is null if nothing usable was found (adminAlerts is updated).
 */
function findAndExtractCSV(log, adminAlerts) {
  const query = `subject:"${CONFIG.SUBJECT_PATTERN}" newer_than:2d -label:${CONFIG.PROCESSED_LABEL}`;
  log.push(`Gmail search: ${query}`);

  let threads;
  try {
    threads = GmailApp.search(query, 0, 10);
  } catch (e) {
    log.push(`Gmail search error: ${e}`);
    return { content: null, thread: null };
  }

  log.push(`Threads found (unprocessed): ${threads.length}`);

  if (threads.length === 0) {
    const processed = GmailApp.search(
      `subject:"${CONFIG.SUBJECT_PATTERN}" newer_than:2d label:${CONFIG.PROCESSED_LABEL}`, 0, 1);
    if (processed.length > 0) {
      log.push('INFO: Report email exists but was already processed — skipping.');
    } else {
      log.push('WARNING: No matching email found.');
      adminAlerts.missingEmail = true;
    }
    return { content: null, thread: null };
  }

  threads.sort((a, b) => b.getLastMessageDate() - a.getLastMessageDate());
  const thread = threads[0];
  const msgs   = thread.getMessages();
  const msg    = msgs[msgs.length - 1];

  log.push(`Email: "${msg.getSubject()}" | From: ${msg.getFrom()} | Date: ${msg.getDate()}`);

  const attachments = msg.getAttachments();
  log.push(`Attachments: ${attachments.length}`);

  for (const att of attachments) {
    const name = att.getName();
    const type = att.getContentType().toLowerCase();
    log.push(`  → "${name}" (${type})`);

    if (name.toLowerCase().endsWith('.csv') || type.includes('csv') || type.includes('text/plain')) {
      const content = att.getDataAsString();
      log.push(`  CSV extracted: ${content.length} chars, ~${content.split('\n').length} lines`);

      // Mark thread processed so it is never picked up twice
      try {
        let label = GmailApp.getUserLabelByName(CONFIG.PROCESSED_LABEL)
                 || GmailApp.createLabel(CONFIG.PROCESSED_LABEL);
        thread.addLabel(label);
        log.push(`  Thread labelled "${CONFIG.PROCESSED_LABEL}"`);
      } catch (e) {
        log.push(`  WARNING: Could not apply processed label: ${e}`);
      }
      return { content, thread };
    }
  }

  log.push('ERROR: No CSV attachment found in email.');
  adminAlerts.noCSV = true;
  return { content: null, thread: null };
}

/**
 * Parses the CSV into rows: { date, provider, shift, rvu }.
 * date is the raw datetime string; it is canonicalized to ISO later in
 * updateRawData. Weekend shifts (no aveRVU) are skipped.
 */
function parseCSV(csvContent, log) {
  // The attachment is normally comma-delimited, but has arrived tab-delimited
  // (.csv extension, TSV content) — pick the delimiter from the header line.
  const firstLine = String(csvContent).split(/\r?\n/)[0] || '';
  const isTSV     = firstLine.indexOf('\t') !== -1 && firstLine.indexOf(',') === -1;

  let rows;
  try {
    rows = isTSV ? Utilities.parseCsv(csvContent, '\t') : Utilities.parseCsv(csvContent);
  } catch (e) {
    log.push(`ERROR: Could not parse attachment: ${e}`);
    return { dataRows: [] };
  }
  if (!rows || rows.length === 0) {
    log.push('ERROR: Utilities.parseCsv returned empty.');
    return { dataRows: [] };
  }
  if (isTSV) log.push('Attachment is tab-delimited.');

  const headers = rows[0].map(h => h.trim());
  log.push(`CSV headers: [${headers.join(' | ')}]`);

  const { col, missing } = resolveColumns(headers);
  if (missing.length) {
    log.push(`ERROR: Could not find a column for: ${missing.join(', ')}. `
           + 'Add the new header name to CONFIG.COL_ALIASES.');
    return { dataRows: [] };
  }
  if (col.shift === -1) log.push('No Shift column found — shift will be blank.');

  const dataRows = [];
  let skipped = 0;

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r || r.every(c => !c.trim())) { skipped++; continue; }

    const date     = cellAt(r, col.date);
    const provider = cellAt(r, col.provider);
    const shift    = cellAt(r, col.shift);
    const rvu      = parseFloat(r[col.rvu]) || 0;

    if (!date || !provider) { log.push(`  Row ${i + 1} skipped: empty date or provider`); skipped++; continue; }
    if (isWeekendShift(shift)) { skipped++; continue; }

    dataRows.push({ date, provider, shift, rvu });
  }

  log.push(`Parsed: ${dataRows.length} valid rows, ${skipped} skipped`);
  return { dataRows };
}

/**
 * Loads the aveRVU reference from the RVUComp3 spreadsheet.
 * Returns a map: lowercase shift name → { name, ave }.
 */
function loadAveRVU(log, adminAlerts) {
  let aveSS;
  try {
    aveSS = SpreadsheetApp.openById(CONFIG.AVE_RVU_SPREADSHEET_ID);
  } catch (e) {
    log.push(`ERROR: Cannot open aveRVU spreadsheet (RVUComp3): ${e}`);
    adminAlerts.aveRVUMissing = true;
    return {};
  }

  const sheet = aveSS.getSheetByName(CONFIG.SHEETS.AVE_RVU);
  if (!sheet) {
    log.push(`ERROR: Sheet "${CONFIG.SHEETS.AVE_RVU}" not found in RVUComp3!`);
    adminAlerts.aveRVUMissing = true;
    return {};
  }

  const data = sheet.getDataRange().getValues();
  const map = {};
  for (let i = 1; i < data.length; i++) {
    const name = data[i][0] ? data[i][0].toString().trim() : '';
    const ave  = parseFloat(data[i][1]) || 0;
    if (name) map[name.toLowerCase()] = { name, ave };
  }

  log.push(`aveRVU: ${Object.keys(map).length} shifts loaded`);
  return map;
}

/** Flags shift names present in the CSV but missing from the aveRVU sheet. */
function checkShiftMismatches(dataRows, aveRVUMap, log, adminAlerts) {
  const unmatched = new Set();
  dataRows.forEach(r => {
    const key = r.shift.trim().toLowerCase();
    if (key !== CONFIG.NOT_SCHEDULED_KEY && !isWeekendShift(r.shift) && !aveRVUMap[key]) {
      unmatched.add(r.shift);
    }
  });

  if (unmatched.size) {
    adminAlerts.shiftMismatches = [...unmatched].sort();
    log.push(`Shift mismatch(es): ${adminAlerts.shiftMismatches.join(', ')}`);
  } else {
    log.push('Shift name check: all shifts matched (or were Not Scheduled).');
  }
}
