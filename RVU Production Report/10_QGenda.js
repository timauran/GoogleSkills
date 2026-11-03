// ============================================================
// 10_QGenda — vacation lookup from the QGenda iCal schedule
// ============================================================
// The RVU report only knows what was read, not who was scheduled off, so
// "worked on vacation" comes from QGenda. Same feed AutoAssign uses:
// all-day events whose SUMMARY is "VACATION 2 [SELPH]" (rotation, then the
// rad's LAST NAME in brackets; QGenda notes may trail the bracket).
//
// The feed URL contains a secret key and this repo is public, so it lives
// in Script Property CONFIG.PROP_QGENDA_URL — set it from the menu
// (Setup & Admin ▸ Set QGenda iCal URL…). No URL → vacation flags are
// simply skipped; the daily email never fails because of QGenda.
// ============================================================

let _qgendaEvents = null;   // per-execution cache of parsed vacation events

/** Parsed VACATION events: [{ start, end (exclusive ISO), staff, rotation }]. */
function loadQGendaVacations_(log) {
  if (_qgendaEvents) return _qgendaEvents;
  _qgendaEvents = [];

  const url = PropertiesService.getScriptProperties().getProperty(CONFIG.PROP_QGENDA_URL);
  if (!url) { if (log) log.push('QGenda: no iCal URL set — vacation check skipped'); return _qgendaEvents; }

  let ics;
  try {
    const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) {
      if (log) log.push(`QGenda: iCal HTTP ${resp.getResponseCode()} — vacation check skipped`);
      return _qgendaEvents;
    }
    ics = resp.getContentText();
  } catch (e) {
    if (log) log.push(`QGenda: fetch failed (${e}) — vacation check skipped`);
    return _qgendaEvents;
  }

  // RFC 5545: unfold continuation lines, then walk VEVENT blocks
  const lines = ics.replace(/\r?\n[ \t]/g, '').split(/\r?\n/);
  const icsDate = l => { const m = l.substring(l.indexOf(':') + 1).match(/^(\d{4})(\d{2})(\d{2})/); return m ? `${m[1]}-${m[2]}-${m[3]}` : null; };
  let cur = null;
  lines.forEach(line => {
    if (line === 'BEGIN:VEVENT') { cur = {}; return; }
    if (line === 'END:VEVENT') {
      if (cur && cur.summary && cur.start) {
        const m = cur.summary.match(/^(.*?)\s*\[([^\]]+)\]/);   // rotation [STAFF] trailing-notes
        if (m && isVacationShift(m[1])) {
          _qgendaEvents.push({ start: cur.start, end: cur.end || cur.start, rotation: m[1].trim(), staff: m[2].trim() });
        }
      }
      cur = null;
      return;
    }
    if (!cur) return;
    if (line.indexOf('SUMMARY') === 0)      cur.summary = line.substring(line.indexOf(':') + 1).replace(/\\([,;\\])/g, '$1').trim();
    else if (line.indexOf('DTSTART') === 0) cur.start = icsDate(line);
    else if (line.indexOf('DTEND') === 0)   cur.end   = icsDate(line);
  });

  if (log) log.push(`QGenda: ${_qgendaEvents.length} vacation event(s) loaded`);
  return _qgendaEvents;
}

/** Letters-only uppercase key, dropping NP/PA suffixes: "Vergara NP" → "VERGARA". */
function qgendaNameKey_(s) {
  return String(s).toUpperCase().replace(/\b(NP|PA|MD|DO)\b/g, '').replace(/[^A-Z]/g, '');
}

/**
 * Of the given providers ("Last, First"), which are on a QGenda VACATION
 * shift on `iso`? Matched on last name. Returns Map<provider, rotation>.
 */
function getVacationProviders(iso, providers, log) {
  const onDay = loadQGendaVacations_(log).filter(e => iso === e.start || (e.start < iso && iso < e.end));   // DTEND exclusive
  const byLast = new Map(onDay.map(e => [qgendaNameKey_(e.staff), e.rotation]));
  const out = new Map();
  providers.forEach(p => {
    const rot = byLast.get(qgendaNameKey_(String(p).split(',')[0]));
    if (rot) out.set(p, rot);
  });
  return out;
}

/** Menu action: store the QGenda iCal URL and test it (also triggers the UrlFetch consent). */
function setQGendaUrl() {
  const ui    = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const has   = !!props.getProperty(CONFIG.PROP_QGENDA_URL);
  const r = ui.prompt('QGenda iCal URL',
    (has ? 'A URL is already set. Paste a new one to replace it, or Cancel.\n\n' : '')
    + 'Paste the QGenda iCal URL (https://app.qgenda.com/iCal?key=…):',
    ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;

  const url = r.getResponseText().trim();
  if (!/^https:\/\/app\.qgenda\.com\/ical\?key=/i.test(url)) { ui.alert('That does not look like a QGenda iCal URL — nothing saved.'); return; }
  props.setProperty(CONFIG.PROP_QGENDA_URL, url);

  const log = [];
  _qgendaEvents = null;
  loadQGendaVacations_(log);
  ui.alert('QGenda', 'Saved.\n\n' + log.join('\n'), ui.ButtonSet.OK);
}
