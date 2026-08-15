// ============================================================
// 05_Pivots — Total RVU, Not Scheduled, Productivity, Monthly Summary
// ------------------------------------------------------------
// All three date-grid pivots share the same layout (frozen header + a
// totals row + one row per provider, with weekend shading and fixed
// column widths). renderPivot() builds that grid once; each pivot just
// supplies its dates, providers, per-cell values, and colors.
// ============================================================

/**
 * Renders a provider × date pivot grid onto a sheet.
 * opts:
 *   headerColor            header row background
 *   totalsLabel, totalsBg  row-1 label + background
 *   dates      ISO date strings (already sorted, newest first)
 *   providers  provider names (already sorted)
 *   totalsFn(iso)          → value for the totals row (number or string)
 *   cellFn(provider, iso)  → { v, bg?, note? }  (v '' when no data)
 *   heat                   apply RVU heat-map to numeric provider cells
 *   numberFormat           number format for value cells (omit for text)
 *   altColor               zebra color for odd provider rows
 *   protect                Set of cell colors weekend shading must not override
 *   emptyMsg               shown when there is nothing to render
 */
function renderPivot(sheet, opts) {
  sheet.clearContents();
  sheet.clearNotes();
  sheet.clearFormats();

  const { dates, providers } = opts;
  if (!dates.length || !providers.length) {
    sheet.getRange(1, 1).setValue(opts.emptyMsg || 'No data.');
    return;
  }

  const nCols   = dates.length + 1;
  const nRows   = providers.length + 2;
  const altBase = opts.altColor || CONFIG.COLOR.ROW_ALT;

  const values = [], bgs = [], notes = [];

  // Row 0 — header
  values.push(['Rad \\ Date', ...dates.map(isoToUS)]);
  bgs.push(Array(nCols).fill(opts.headerColor));
  notes.push(Array(nCols).fill(''));

  // Row 1 — totals
  const totalsRow = [opts.totalsLabel];
  dates.forEach(d => totalsRow.push(opts.totalsFn(d)));
  values.push(totalsRow);
  bgs.push(Array(nCols).fill(opts.totalsBg));
  notes.push(Array(nCols).fill(''));

  // Rows 2+ — providers
  providers.forEach((p, pIdx) => {
    const base = pIdx % 2 === 1 ? altBase : '#ffffff';
    const vRow = [p], bRow = [base], nRow = [''];
    dates.forEach(d => {
      const c = opts.cellFn(p, d) || {};
      vRow.push(c.v === undefined ? '' : c.v);
      bRow.push(c.bg || base);
      nRow.push(c.note || '');
    });
    values.push(vRow); bgs.push(bRow); notes.push(nRow);
  });

  // Heat-map (numeric provider cells only, per column)
  if (opts.heat) {
    for (let col = 1; col < nCols; col++) {
      const nums = [];
      for (let r = 2; r < nRows; r++) if (typeof values[r][col] === 'number') nums.push(values[r][col]);
      if (nums.length < 2) continue;
      const mn = Math.min(...nums), mx = Math.max(...nums);
      for (let r = 2; r < nRows; r++) {
        if (typeof values[r][col] === 'number') bgs[r][col] = rvuHeatColor(values[r][col], mn, mx);
      }
    }
  }

  // Weekend / holiday shading (skips protected colors, e.g. threshold green/orange)
  dates.forEach((d, dIdx) => {
    if (!isoIsWeekendOrHoliday(d)) return;
    const col = dIdx + 1;
    bgs[0][col] = CONFIG.COLOR.WEEKEND_HDR;
    for (let r = 1; r < nRows; r++) {
      if (!opts.protect || !opts.protect.has(bgs[r][col])) bgs[r][col] = CONFIG.COLOR.WEEKEND_GRAY;
    }
  });

  // Single batch write
  const range = sheet.getRange(1, 1, nRows, nCols);
  range.setValues(values);
  range.setBackgrounds(bgs);
  range.setNotes(notes);

  sheet.getRange(1, 1, 1, nCols).setFontWeight('bold').setFontColor('white');
  sheet.getRange(2, 1, 1, nCols).setFontWeight('bold');
  if (opts.numberFormat) {
    sheet.getRange(2, 2, nRows - 1, dates.length)
      .setNumberFormat(opts.numberFormat)
      .setHorizontalAlignment('right');
  }
  sheet.setFrozenRows(2);
  sheet.setFrozenColumns(1);
  setPivotColumnWidths(sheet, nCols);
}

// ── Total RVU ───────────────────────────────────────────────

function updateTotalRVUSheet(ss, allData, log) {
  const map = {}, dateSet = new Set();
  allData.forEach(r => {
    dateSet.add(r.date);
    if (!map[r.provider])         map[r.provider] = {};
    if (!map[r.provider][r.date]) map[r.provider][r.date] = { total: 0, shifts: [] };
    map[r.provider][r.date].total += r.rvu;
    map[r.provider][r.date].shifts.push({ shift: r.shift, rvu: r.rvu });
  });

  const dates     = sortISOsDesc([...dateSet]);
  const providers = Object.keys(map).sort();

  renderPivot(getOrCreateSheet(ss, CONFIG.SHEETS.TOTAL_RVU), {
    headerColor: CONFIG.COLOR.HEADER_DARK,
    totalsLabel: 'TOTAL', totalsBg: CONFIG.COLOR.SUBTOTAL,
    dates, providers, heat: true, numberFormat: '#,##0.0',
    emptyMsg: 'No data.',
    totalsFn: d => roundRVU(providers.reduce((s, p) =>
      s + ((map[p][d] && map[p][d].total) || 0), 0)),
    cellFn: (p, d) => {
      const c = map[p][d];
      if (!c) return { v: '' };
      return { v: roundRVU(c.total), note: c.shifts.map(s => `${s.shift}: ${s.rvu.toFixed(2)}`).join('\n') };
    }
  });

  log.push(`Total RVU: ${providers.length} providers × ${dates.length} dates`);
}

// ── Not Scheduled ───────────────────────────────────────────

function updateNotScheduledSheet(ss, allData, log) {
  const map = {}, dateSet = new Set();
  allData.filter(r => isNS(r.shift)).forEach(r => {
    dateSet.add(r.date);
    if (!map[r.provider]) map[r.provider] = {};
    map[r.provider][r.date] = (map[r.provider][r.date] || 0) + r.rvu;
  });

  // Only include rads who also worked a scheduled (non-NS) shift.
  const scheduled = new Set(allData.filter(r => !isNS(r.shift)).map(r => r.provider));
  const dates     = sortISOsDesc([...dateSet]);
  const providers = Object.keys(map).filter(p => scheduled.has(p)).sort();

  renderPivot(getOrCreateSheet(ss, CONFIG.SHEETS.NOT_SCHEDULED), {
    headerColor: CONFIG.COLOR.HEADER_NS,
    totalsLabel: 'TOTAL', totalsBg: CONFIG.COLOR.SUBTOTAL_NS,
    dates, providers, heat: true, numberFormat: '#,##0.0',
    altColor: CONFIG.COLOR.ROW_ALT_NS,
    emptyMsg: 'No "Not Scheduled" data for scheduled providers.',
    totalsFn: d => {
      const t = providers.reduce((s, p) => s + (map[p][d] || 0), 0);
      return t > 0 ? roundRVU(t) : '';
    },
    cellFn: (p, d) => ({ v: map[p][d] ? roundRVU(map[p][d]) : '' })
  });

  log.push(`Not Scheduled: ${providers.length} providers × ${dates.length} dates`);
}

// ── Productivity (threshold) ────────────────────────────────

function updateThresholdSheet(ss, allData, aveRVUMap, log) {
  const map = {}, dateSet = new Set();
  allData
    .filter(r => !isNS(r.shift) && !!aveRVUMap[r.shift.trim().toLowerCase()])
    .forEach(r => {
      const ave = aveRVUMap[r.shift.trim().toLowerCase()].ave;
      dateSet.add(r.date);
      if (!map[r.provider])         map[r.provider] = {};
      if (!map[r.provider][r.date]) map[r.provider][r.date] = { actual: 0, expected: 0, shifts: [] };
      map[r.provider][r.date].actual   += r.rvu;
      map[r.provider][r.date].expected += ave;
      map[r.provider][r.date].shifts.push({ name: r.shift, rvu: r.rvu, ave });
    });

  const dates     = sortISOsDesc([...dateSet]);
  const providers = Object.keys(map).sort();
  const protect   = new Set([CONFIG.COLOR.GREEN_PCT, CONFIG.COLOR.ORANGE_PCT]);

  renderPivot(getOrCreateSheet(ss, CONFIG.SHEETS.THRESHOLD), {
    headerColor: CONFIG.COLOR.HEADER_PROD,
    totalsLabel: 'AVG THRESHOLD', totalsBg: CONFIG.COLOR.SUBTOTAL_PD,
    dates, providers, heat: false, altColor: CONFIG.COLOR.ROW_ALT_PD, protect,
    emptyMsg: 'No scheduled shift data.',
    totalsFn: d => {
      const pcts = providers.map(p => {
        const c = map[p] && map[p][d];
        return (c && c.expected > 0) ? (c.actual / c.expected) * 100 : null;
      }).filter(v => v !== null);
      return pcts.length ? (pcts.reduce((a, b) => a + b, 0) / pcts.length).toFixed(1) + '%' : '';
    },
    cellFn: (p, d) => {
      const c = map[p] && map[p][d];
      if (!c) return { v: '' };
      if (c.expected === 0) {
        return {
          v: c.actual.toFixed(2) + ' (no avg)',
          note: c.shifts.map(s => `${s.name}: ${s.rvu.toFixed(1)} RVU (avg: n/a)`).join('\n')
        };
      }
      const pct  = (c.actual / c.expected) * 100;
      const note = c.shifts.map(s => {
        const sp = s.ave > 0 ? (s.rvu / s.ave * 100).toFixed(0) + '%' : 'no avg';
        return `${s.name}: ${s.rvu.toFixed(1)} RVU  (avg: ${s.ave.toFixed(1)}, ${sp})`;
      }).join('\n');
      const bg = pct >= CONFIG.THRESHOLD_GREEN  ? CONFIG.COLOR.GREEN_PCT
               : pct <  CONFIG.THRESHOLD_ORANGE ? CONFIG.COLOR.ORANGE_PCT
               : undefined;
      return { v: pct.toFixed(1) + '%', bg, note };
    }
  });

  log.push(`Threshold: ${providers.length} providers × ${dates.length} dates`);
}

// ── Monthly Summary (flat table) ────────────────────────────

function updateMonthlySummary(ss, allData, log) {
  const sheet = getOrCreateSheet(ss, CONFIG.SHEETS.MONTHLY_SUMMARY);
  sheet.clearContents();
  sheet.clearFormats();

  // month key = ISO year-month (r.date is already ISO "yyyy-MM-dd")
  const monthMap = {};
  allData.forEach(r => {
    const month = r.date.substring(0, 7);
    if (!monthMap[month]) monthMap[month] = {};
    if (!monthMap[month][r.provider]) monthMap[month][r.provider] = { total: 0, scheduled: 0, ns: 0 };
    monthMap[month][r.provider].total += r.rvu;
    if (isNS(r.shift)) monthMap[month][r.provider].ns        += r.rvu;
    else               monthMap[month][r.provider].scheduled += r.rvu;
  });

  const headers = ['Month', 'Provider', 'Total RVU', 'Scheduled RVU', 'Not Scheduled RVU'];
  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground(CONFIG.COLOR.HEADER_DARK)
    .setFontColor('white');

  // Only providers with scheduled RVU in the month; highest total first.
  const rows = [], bgs = [];
  Object.keys(monthMap).sort().forEach(month => {
    Object.keys(monthMap[month])
      .filter(p => monthMap[month][p].scheduled > 0)
      .sort((a, b) => monthMap[month][b].total - monthMap[month][a].total)
      .forEach((p, i) => {
        const d = monthMap[month][p];
        rows.push([month, p, roundRVU(d.total), roundRVU(d.scheduled), roundRVU(d.ns)]);
        bgs.push(Array(5).fill(i % 2 === 1 ? CONFIG.COLOR.ROW_ALT : '#ffffff'));
      });
  });

  if (rows.length) {
    const range = sheet.getRange(2, 1, rows.length, 5);
    range.setValues(rows);
    range.setBackgrounds(bgs);
  }

  sheet.setFrozenRows(1);
  try { sheet.autoResizeColumns(1, headers.length); } catch (e) {}
  log.push(`Monthly Summary: ${Object.keys(monthMap).length} months, ${rows.length} detail rows`);
}
