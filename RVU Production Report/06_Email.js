// ============================================================
// 06_Email — daily production email + admin summary email
// ============================================================

/**
 * Reads send mode and recipients from the spreadsheet.
 *   Config sheet:     "Send Mode" row → TRUE (send) / FALSE (draft)
 *   Email List sheet: Name | Email | Role (To/CC/BCC), blank role = To
 * MY_EMAIL is always BCC'd. Returns { to, cc, bcc, sendMode }.
 */
function readEmailConfig(ss, log) {
  let sendMode = false;
  const cfgSheet = ss.getSheetByName(CONFIG.SHEETS.CONFIG_SHEET);
  if (cfgSheet && cfgSheet.getLastRow() >= 1) {
    cfgSheet.getDataRange().getValues().forEach(row => {
      const key = row[0] ? row[0].toString().trim().toLowerCase() : '';
      if (key === 'send mode') {
        sendMode = row[1] === true || String(row[1]).trim().toLowerCase() === 'true';
      }
    });
  } else {
    log.push('Email Config: Config sheet missing — defaulting to DRAFT mode');
  }
  log.push(`Email Config: send mode = ${sendMode ? 'SEND' : 'DRAFT'}`);

  const toList = [], ccList = [], bccList = [];
  const listSheet = ss.getSheetByName(CONFIG.SHEETS.EMAIL_LIST);
  if (listSheet && listSheet.getLastRow() >= 2) {
    listSheet.getRange(2, 1, listSheet.getLastRow() - 1, 3).getValues().forEach(row => {
      const email = row[1] ? row[1].toString().trim() : '';
      if (!email || !email.includes('@')) return;
      const role = row[2] ? row[2].toString().trim().toLowerCase() : 'to';
      if      (role === 'cc')  ccList.push(email);
      else if (role === 'bcc') bccList.push(email);
      else                     toList.push(email);
    });
  } else {
    log.push('Email Config: Email List sheet missing or empty — sending only to MY_EMAIL');
  }

  if (![...toList, ...ccList, ...bccList].includes(CONFIG.MY_EMAIL)) bccList.push(CONFIG.MY_EMAIL);

  const to  = toList.length ? toList.join(',') : CONFIG.MY_EMAIL;
  const cc  = ccList.join(',');
  const bcc = bccList.join(',');
  log.push(`Email Config: To=${to}, CC=${cc || '(none)'}, BCC=${bcc || '(none)'}`);
  return { to, cc: cc || null, bcc: bcc || null, sendMode };
}

/**
 * Builds and sends/drafts the daily production email for one day.
 * @param targetIso  optional ISO "yyyy-MM-dd"; defaults to the prior day.
 * @param forceDraft optional; always create a draft (used for historical dates).
 */
function createDailyEmailDraft(ss, allData, aveRVUMap, log, targetIso, forceDraft) {
  const dayIso  = targetIso || priorDayISO();
  const dayDisp = isoToUS(dayIso);
  log.push(`Email draft: targeting ${dayIso}${targetIso ? ' (manual)' : ' (prior day)'}`);

  const dayData = allData.filter(r => r.date === dayIso);
  if (!dayData.length) { log.push(`No data for ${dayIso} — skipping email draft`); return; }

  // Per-provider summary
  const provMap = {};
  dayData.forEach(r => {
    if (!provMap[r.provider]) provMap[r.provider] = { scheduledShifts: [], totalRVU: 0, nsRVU: 0 };
    provMap[r.provider].totalRVU += r.rvu;
    if (isNS(r.shift)) {
      provMap[r.provider].nsRVU += r.rvu;
    } else {
      const entry = aveRVUMap[r.shift.trim().toLowerCase()];
      const ave   = (entry && entry.ave > 0) ? entry.ave : 0;
      const pct   = ave > 0 ? (r.rvu / ave) * 100 : null;
      provMap[r.provider].scheduledShifts.push({ shift: r.shift, rvu: r.rvu, pct, ave });
    }
  });

  // Section 1: every rad with any RVU that day, by last name.
  // Not Scheduled gets its own line inside each rad's block, so a rad with
  // only NS RVU still belongs here and the grand total covers the day.
  const provs = Object.keys(provMap)
    .filter(p => provMap[p].scheduledShifts.length > 0 || provMap[p].nsRVU > 0)
    .sort(byLastName);

  // Section 2: any rad with NS RVU, same alphabetical order
  const nsProvs = Object.keys(provMap)
    .filter(p => provMap[p].nsRVU > 0)
    .sort(byLastName);

  // Rotations listed highest-RVU first within each rad's block
  provs.forEach(p => provMap[p].scheduledShifts.sort((a, b) => b.rvu - a.rvu));

  const html = buildDraftHTML(dayDisp, provs, provMap, nsProvs);

  const cfg     = readEmailConfig(ss, log);
  const options = { htmlBody: html };
  if (cfg.cc)  options.cc  = cfg.cc;
  if (cfg.bcc) options.bcc = cfg.bcc;
  const subject = `RVU Production — ${dayDisp}`;

  if (!forceDraft && cfg.sendMode) {
    GmailApp.sendEmail(cfg.to, subject, '', options);
    log.push(`Email SENT: to="${cfg.to}", ${provs.length} providers, ${nsProvs.length} with Not Scheduled`);
  } else {
    GmailApp.createDraft(cfg.to, subject, '', options);
    log.push(`Email DRAFT created: to="${cfg.to}", ${provs.length} providers, ${nsProvs.length} with Not Scheduled`);
  }
}

/**
 * Builds the daily production email HTML. `date` is a display string.
 *
 * Section 1 gives every radiologist one row per rotation worked that day,
 * plus a separate "Not Scheduled" row where applicable. The Radiologist and
 * Total RVU cells span that whole block, so Total RVU is the sum of every
 * row shown for the rad — scheduled rotations and Not Scheduled alike.
 * The RVU Obtained cell is shaded against the rotation's average:
 * green ≥120%, plain 75–120%, orange <75% (CONFIG.THRESHOLD_*).
 */
function buildDraftHTML(date, provs, provMap, nsProvs) {
  const tableStyle = 'border-collapse:collapse;margin:0 0 4px;';
  const thL = 'padding:3px 10px 3px 6px;text-align:left;font-size:10px;white-space:nowrap;color:#fff;';
  const thR = 'padding:3px 10px 3px 6px;text-align:right;font-size:10px;white-space:nowrap;color:#fff;';
  const tdL = 'padding:2px 10px 2px 6px;vertical-align:top;font-size:10px;white-space:nowrap;';
  const tdLw= 'padding:2px 10px 2px 6px;vertical-align:top;font-size:10px;';
  const tdR = 'padding:2px 10px 2px 6px;vertical-align:top;text-align:right;font-family:monospace;font-size:10px;white-space:nowrap;';
  const tdRb= tdR + 'font-weight:bold;';
  const NCOLS = 6;
  const RULE  = 'border-top:1px solid #c9d9e8;';   // separates one rad's block from the next
  const DASH  = '<span style="color:#bbb;">—</span>';

  const sectionHeader = (title, bg, cols) =>
    `<tr><td colspan="${cols || NCOLS}" style="background:${bg};color:#fff;font-weight:bold;font-size:11px;padding:3px 8px;white-space:nowrap;">${title}</td></tr>`;

  /** Background/foreground for an RVU-obtained cell, by % of rotation average. */
  const pctStyle = (pct) => {
    if (pct === null) return '';                                   // no average → no judgement
    if (pct >= CONFIG.THRESHOLD_GREEN)  return `background:${CONFIG.COLOR.GREEN_PCT};color:#1e7e34;font-weight:bold;`;
    if (pct <  CONFIG.THRESHOLD_ORANGE) return `background:${CONFIG.COLOR.ORANGE_PCT};color:#b05000;font-weight:bold;`;
    return '';                                                     // regular
  };

  // Section 1 — one row per rotation, plus a Not Scheduled row, per radiologist
  const provRows = provs.map((p, i) => {
    const d      = provMap[p];
    const bg     = i % 2 === 0 ? '#ffffff' : '#eaf3fb';
    const nRows  = d.scheduledShifts.length + (d.nsRVU > 0 ? 1 : 0);
    if (!nRows) return '';

    // Cells that span the rad's whole block
    const nameCell  = `<td rowspan="${nRows}" style="${tdL}${RULE}font-weight:bold;vertical-align:middle;">${p}</td>`;
    const totalCell = `<td rowspan="${nRows}" style="${tdRb}${RULE}color:#1a5276;vertical-align:middle;">${d.totalRVU.toFixed(1)}</td>`;

    const lines = d.scheduledShifts.map(s => {
      const pct = s.ave > 0 ? (s.rvu / s.ave) * 100 : null;
      return {
        rotation: s.shift,
        ave:      s.ave > 0 ? s.ave.toFixed(1) : DASH,
        obtained: s.rvu.toFixed(1),
        pct:      pct === null ? DASH : `${pct.toFixed(0)}%`,
        // "PM Imaging" and friends show their numbers but are never judged
        style:    isNoHighlightShift(s.shift) ? '' : pctStyle(pct)
      };
    });

    if (d.nsRVU > 0) {
      lines.push({
        rotation: '<em style="color:#b45f06;">Not Scheduled</em>',
        ave:      DASH,
        obtained: d.nsRVU.toFixed(1),
        pct:      DASH,
        style:    ''      // no average to compare against — never colour-coded
      });
    }

    return lines.map((ln, k) => {
      const rule = k === 0 ? RULE : '';           // rule only on the block's first row
      return `<tr style="background:${bg};">
        ${k === 0 ? nameCell : ''}
        <td style="${tdLw}${rule}">${ln.rotation}</td>
        <td style="${tdR}${rule}color:#666;">${ln.ave}</td>
        <td style="${tdR}${rule}${ln.style}">${ln.obtained}</td>
        <td style="${tdR}${rule}${ln.style}">${ln.pct}</td>
        ${k === 0 ? totalCell : ''}
      </tr>`;
    }).join('');
  }).join('');

  const grandTotal = provs.reduce((s, p) => s + provMap[p].totalRVU, 0);
  const grandTotalRow = `<tr style="background:${CONFIG.COLOR.GRAND_TOTAL};">
    <td style="${tdL}font-weight:bold;color:#1a5276;border-top:2px solid #1a5276;" colspan="5">GRAND TOTAL</td>
    <td style="${tdRb}color:#1a5276;border-top:2px solid #1a5276;">${grandTotal.toFixed(1)}</td>
  </tr>`;

  // Section 2 — Not Scheduled RVU
  const nsRows = nsProvs.length === 0
    ? `<tr><td colspan="2" style="${tdL}color:#999;"><em>None</em></td></tr>`
    : nsProvs.map((p, i) => {
        const bg = i % 2 === 0 ? '#fff8f4' : '#fce5cd';
        return `<tr style="background:${bg};">
          <td style="${tdL}font-weight:bold;">${p}</td>
          <td style="${tdR}">${provMap[p].nsRVU.toFixed(1)}</td>
        </tr>`;
      }).join('');

  return `<!DOCTYPE html><html><head><meta charset="UTF-8"></head>
<body style="font-family:Arial,Helvetica,sans-serif;font-size:11px;color:#222;margin:0;padding:8px;">
<p style="margin:0;line-height:1.6;">&nbsp;</p>
<p style="margin:0;line-height:1.6;">&nbsp;</p>
<div style="background:#d6e4f0;border-left:4px solid #1a5276;padding:5px 10px;margin-bottom:6px;">
  <strong style="font-size:13px;color:#1a5276;">RVU Daily Production — ${date}</strong>
</div>

<table style="${tableStyle}">
  <thead>
    ${sectionHeader('Total Daily RVU Production', '#1a5276', NCOLS)}
    <tr style="background:#2e86c1;">
      <th style="${thL}">Radiologist</th>
      <th style="${thL}">Rotation</th>
      <th style="${thR}">Avg RVU</th>
      <th style="${thR}">RVU Obtained</th>
      <th style="${thR}">% of Avg</th>
      <th style="${thR}">Total RVU</th>
    </tr>
  </thead>
  <tbody>${provRows}${grandTotalRow}</tbody>
</table>

<p style="margin:2px 0 0;font-size:9px;color:#777;">
  % of rotation average:
  <span style="background:${CONFIG.COLOR.GREEN_PCT};color:#1e7e34;padding:0 4px;">&ge;${CONFIG.THRESHOLD_GREEN}%</span>
  <span style="padding:0 4px;">${CONFIG.THRESHOLD_ORANGE}&ndash;${CONFIG.THRESHOLD_GREEN}%</span>
  <span style="background:${CONFIG.COLOR.ORANGE_PCT};color:#b05000;padding:0 4px;">&lt;${CONFIG.THRESHOLD_ORANGE}%</span>
  &nbsp;&middot;&nbsp; ${CONFIG.NO_HIGHLIGHT_SHIFTS.join(', ')} and Not Scheduled are not colour-coded.
</p>

<p style="margin:0;line-height:1.6;">&nbsp;</p>

<table style="${tableStyle}">
  <thead>
    ${sectionHeader('Not Scheduled RVU', '#b45f06', 2)}
    <tr style="background:#cc7a00;">
      <th style="${thL}">Radiologist</th>
      <th style="${thR}">Not Scheduled RVU</th>
    </tr>
  </thead>
  <tbody>${nsRows}</tbody>
</table>
</body></html>`;
}

/**
 * Sends a single consolidated admin email covering missing report,
 * no CSV, missing aveRVU, shift mismatches, new providers, and WorkRVU
 * changes. Suppressed when there is nothing to report. Admin-only.
 */
function sendAdminSummaryEmail(adminAlerts, log) {
  const { missingEmail, noCSV, aveRVUMissing, shiftMismatches, newProviders, changes } = adminAlerts;
  const hasAlerts = missingEmail || noCSV || aveRVUMissing
                 || shiftMismatches.length || newProviders.length || changes.length;
  if (!hasAlerts) { log.push('Admin summary: no alerts — email suppressed.'); return; }

  const ts = fmtDate(new Date(), 'yyyy-MM-dd HH:mm z');
  const H2 = (title, bg) =>
    `<tr><td colspan="6" style="background:${bg};color:#fff;font-weight:bold;font-size:11px;padding:4px 10px;border-top:3px solid #fff;">${title}</td></tr>`;
  const banner = (bg, text) =>
    `<tr style="background:${bg};"><td colspan="6" style="padding:5px 10px;font-size:10px;">${text}</td></tr>`;

  let rows = '';

  if (missingEmail) {
    rows += H2('⚠️ No Matching Email Found', '#c0392b');
    rows += banner('#fdf2f2', `No email matching <code>${CONFIG.SUBJECT_PATTERN}</code> was found in the last 48 hours.<br>The report may not have run, or the subject format may have changed.`);
  }
  if (noCSV) {
    rows += H2('⚠️ Email Found — No CSV Attachment', '#c0392b');
    rows += banner('#fdf2f2', 'A matching email was found but contained no CSV attachment. Processing was stopped.');
  }
  if (aveRVUMissing) {
    rows += H2('⚠️ aveRVU Sheet Not Found', '#c0392b');
    rows += banner('#fdf2f2', 'The <strong>aveRVU</strong> sheet was not found. Threshold calculations were skipped.');
  }
  if (shiftMismatches.length) {
    rows += H2(`⚠️ Shift Name Mismatch (${shiftMismatches.length})`, '#e67e22');
    rows += banner('#fef5ec', 'These shift names appear in the CSV but have no entry in the <strong>aveRVU</strong> sheet. Threshold calculations are incomplete for them.');
    shiftMismatches.forEach((s, i) => {
      rows += `<tr style="background:${i % 2 === 0 ? '#fff8f0' : '#fef5ec'};"><td colspan="6" style="padding:3px 10px;font-size:10px;font-family:monospace;">"${s}"</td></tr>`;
    });
  }
  if (newProviders.length) {
    rows += H2(`ℹ️ New Provider(s) Detected (${newProviders.length})`, '#1a5276');
    rows += banner('#eaf3fb', 'The following provider(s) appear for the first time. Please verify names.');
    newProviders.forEach((p, i) => {
      rows += `<tr style="background:${i % 2 === 0 ? '#f0f7ff' : '#eaf3fb'};"><td colspan="6" style="padding:3px 10px;font-size:10px;font-weight:bold;">${p}</td></tr>`;
    });
  }

  const MAX_CHANGES = 100;
  if (changes.length) {
    const shown    = changes.slice(0, MAX_CHANGES);
    const overflow = changes.length - shown.length;
    rows += H2(`📝 Raw Data Changes (${changes.length} cell${changes.length > 1 ? 's' : ''})`, '#274e13');
    rows += `<tr style="background:#f0f4f0;">${
      ['Provider', 'Date', 'Shift', 'Col', 'Old', 'New']
        .map(h => `<th style="padding:2px 8px;font-size:9px;text-align:left;color:#274e13;white-space:nowrap;">${h}</th>`).join('')
    }</tr>`;
    shown.forEach((c, i) => {
      const bg = i % 2 === 0 ? '#f8fbf8' : '#eaf3ea';
      rows += `<tr style="background:${bg};">
        <td style="padding:1px 8px;font-size:9px;">${c.provider}</td>
        <td style="padding:1px 8px;font-size:9px;white-space:nowrap;">${isoToUS(c.date)}</td>
        <td style="padding:1px 8px;font-size:9px;">${c.shift}</td>
        <td style="padding:1px 8px;font-size:9px;font-family:monospace;">${c.col}</td>
        <td style="padding:1px 8px;font-size:9px;font-family:monospace;color:#c0392b;">${String(c.oldVal).substring(0, 40)}</td>
        <td style="padding:1px 8px;font-size:9px;font-family:monospace;color:#1e8449;">${String(c.newVal).substring(0, 40)}</td>
      </tr>`;
    });
    if (overflow > 0) {
      rows += `<tr style="background:#e8f4e8;"><td colspan="6" style="padding:3px 8px;font-size:9px;color:#274e13;font-style:italic;">… and ${overflow} more change${overflow > 1 ? 's' : ''} — see Run Log sheet.</td></tr>`;
    }
  }

  const flags = [];
  if (missingEmail || noCSV)  flags.push('NO REPORT');
  if (aveRVUMissing)          flags.push('aveRVU MISSING');
  if (shiftMismatches.length) flags.push(`${shiftMismatches.length} SHIFT MISMATCH`);
  if (newProviders.length)    flags.push(`${newProviders.length} NEW PROVIDER`);
  if (changes.length)         flags.push(`${changes.length} CHANGE`);
  const subject = `RVU Admin: ${flags.join(' · ')} — ${ts}`;

  const html = `<!DOCTYPE html><html><head><meta charset="UTF-8"></head>
<body style="font-family:Arial,Helvetica,sans-serif;font-size:11px;color:#222;margin:0;padding:8px;">
<div style="background:#1a5276;color:#fff;padding:5px 10px;margin-bottom:6px;">
  <strong style="font-size:12px;">RVU Script Admin Summary</strong>
  <span style="font-size:10px;float:right;opacity:.8;">${ts}</span>
</div>
<table style="border-collapse:collapse;width:100%;">${rows}</table>
</body></html>`;

  // Gmail hard-limits HTML bodies; fall back to plain text if oversized.
  const BODY_LIMIT = 150 * 1024;
  const htmlBytes  = Utilities.newBlob(html).getBytes().length;

  if (htmlBytes > BODY_LIMIT) {
    log.push(`WARNING: HTML body ${htmlBytes} bytes exceeds limit — sending plain-text fallback.`);
    const plain = [
      `RVU Script Admin Summary — ${ts}`,
      '='.repeat(50),
      missingEmail  ? '⚠ No matching report email found.'    : null,
      noCSV         ? '⚠ Email found but no CSV attachment.' : null,
      aveRVUMissing ? '⚠ aveRVU sheet not found.'            : null,
      shiftMismatches.length ? `⚠ Shift mismatches (${shiftMismatches.length}):\n` + shiftMismatches.map(s => `  "${s}"`).join('\n') : null,
      newProviders.length    ? `ℹ New providers (${newProviders.length}):\n` + newProviders.join('\n') : null,
      changes.length ? `${changes.length} data change(s) — see Run Log sheet.\n`
        + changes.slice(0, 50).map(c => `  ${c.provider} | ${isoToUS(c.date)} | ${c.shift} | ${c.col}: "${c.oldVal}" → "${c.newVal}"`).join('\n')
        + (changes.length > 50 ? `\n  … and ${changes.length - 50} more.` : '') : null
    ].filter(Boolean).join('\n\n');
    GmailApp.sendEmail(CONFIG.MY_EMAIL, subject, plain);
  } else {
    GmailApp.sendEmail(CONFIG.MY_EMAIL, subject, '', { htmlBody: html });
  }

  log.push(`Admin summary email sent (${htmlBytes} bytes) — ${flags.join(', ')}`);
}
