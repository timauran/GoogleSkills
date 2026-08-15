# RVU Production Report

Google Apps Script that processes the daily **[SECURE] RVU Comp Report** email,
maintains a canonical **Raw Data** sheet, rebuilds pivot/summary sheets, and
drafts or sends the daily production email.

## Deployment / clasp

This project is managed with [clasp](https://github.com/google/clasp). The IDs
below are what any environment (desktop **or** cloud) needs to push and pull.

| Item | Value |
| --- | --- |
| Apps Script ID (`.clasp.json`) | `1ae6afecQUBBfDHlHOaXTN_V_eDucMzLoseQ6_IRE67EPHQEWmJaHrSY5` |
| aveRVU reference spreadsheet | `1rThkXpgdC9Vo1tYD3sWK30akVhNj8DueX2ZEZoV7Gcs` (sheet `aveRVU`) |
| Data spreadsheet ID | stored in **Script Properties** as `RVU_SPREADSHEET_ID` (set by `createAndSetupSpreadsheet()`) |

`.clasp.json` is committed on purpose so the cloud environment can push without
re-cloning. To work on this project:

```bash
cd "RVU Production Report"
clasp pull      # get the latest from Apps Script
# …edit…
clasp push      # send changes back
```

> The data spreadsheet ID is **not** in source — it lives in Script Properties,
> so the same code drives whichever spreadsheet the deployed script is bound to.

## File map

| File | Responsibility |
| --- | --- |
| `01_Config.js` | `CONFIG` constants (IDs, sheet names, thresholds, colors) |
| `02_Utils.js` | shared helpers: dates, column resolution, name sort, logging |
| `03_Ingest.js` | Gmail fetch, CSV/TSV parse, aveRVU load, shift-mismatch check |
| `04_RawData.js` | canonical Raw Data engine + Provider List + public accessors |
| `05_Pivots.js` | Total RVU / Not Scheduled / Productivity / Monthly Summary |
| `06_Email.js` | daily production email + admin summary email |
| `07_Engine.js` | `checkRVUEmail()` nightly pipeline + manual email-scan ingest |
| `08_Setup.js` | one-time setup, triggers, manual runners |
| `09_ImportSheet.js` | ingest a manually imported `data` sheet into Raw Data |
| `00_Menu.js` | spreadsheet menu + thin action wrappers |

## Menu (⚙️ RVU Report)

- **Run Now (full pipeline)** — find email → Raw Data → pivots → email.
- **Ingest Data ▸ Scan Email & Ingest Now** — pull the report email into Raw
  Data + pivots only (no email sent, source thread not trashed).
- **Ingest Data ▸ Ingest "data" Sheet** — ingest a manually imported tab named
  `data` (File → Import → Insert new sheet, rename it to `data`).
- Rebuild / Email / Setup submenus as before.

## Ingest behavior

Raw Data is keyed on **Date + Provider + Shift** (case-insensitive). New rows
are added; a row whose **WorkRVU** differs is replaced, the old value kept as a
cell note and the cell highlighted yellow; identical rows are left alone.

Column names are matched via `CONFIG.COL_ALIASES`, so a renamed report header
(e.g. `SignDateTime` vs. `ShiftStartDateTime2`) is a one-line config fix.

## Daily email layout

One row per rotation per radiologist (sorted by last name), plus a separate
**Not Scheduled** line. Columns: Radiologist · Rotation · Avg RVU · RVU Obtained
· % of Avg · Total RVU. **RVU Obtained** and **% of Avg** are shaded green
(≥120%), plain (75–120%), or orange (<75%) — except rotations in
`CONFIG.NO_HIGHLIGHT_SHIFTS` (e.g. `PM Imaging`) and Not Scheduled lines, which
are never color-coded. Total RVU spans all of a radiologist's lines.
