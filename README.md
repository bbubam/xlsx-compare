# XLSX Compare

A lightweight, zero-install browser tool for comparing two `.xlsx` or `.csv` files side by side — with row-level diffing, inline change highlighting, and an interactive grid.

## Features

- **Row-level diff** — auto-detects the unique key column(s) per sheet to correctly match rows across files, no manual config needed
- **Inline cell diff** — modified cells show old value (red block, strikethrough) and new value (green block) stacked in a single cell
- **Smart row matching** — falls back to positional (index-based) matching when no unique key column is found, so changed rows show as modified instead of deleted + added
- **New/deleted column detection** — new columns show plain green, deleted columns show plain red; both always visible even with Hide Unchanged on
- **Date formatting** — Excel date serial numbers are rendered as readable dates using the cell's own format string
- **CSV support** — drop or select `.csv` files alongside `.xlsx`
- **Ignore columns** — Alt+click a column header to exclude it from diff comparison entirely; ignored columns show a greyed badge and are hidden when Hide Unchanged is on
- **3-tab view**
  - **Summary** — cards per sheet showing added / deleted / modified row counts, click to jump to grid
  - **Sheet Grid** — full data grid with green (added), red (deleted), yellow (modified) highlights
  - **All Changes** — flat filterable list of every change with old and new values; stays in sync with ignored columns
- **Interactive grid**
  - `Click` col header — lock left · `Shift+click` — lock right · `Alt+click` — ignore col
  - `Click` row # — lock row · `Drag` edges — resize
  - Locked columns always visible; ignored columns hidden when Hide Unchanged is on (unless locked)
  - Lock priority: Locked > Ignored > Hide Unchanged
- **No install, no server** — open `index.html` directly in any browser

## Usage

1. Open `index.html` in a browser
2. Drop or select **File A** (original) and **File B** (modified) — `.xlsx` or `.csv`
3. Click **Compare**

## Files

| File | Purpose |
|------|---------|
| `index.html` | UI layout and styles |
| `script.js` | All comparison logic and rendering |
| `compare.js` | Node CLI tool for headless comparison |

## Changelog

### v1.2.2
- CSV missing column detection: LCS-based column mapping for headerless CSVs
- Fixed tie-breaking to correctly identify last matching deleted column position
- Fixed summary counter showing 0 when only deleted columns differ
- Headerless CSVs no longer show data values as column headers

### v1.2.1
- Right-locked column lock icon now appears after the header text (on the right)

### v1.2.0
- Alt+click column header to ignore columns (excluded from diff, hidden with Hide Unchanged)
- Ignored columns sync to All Changes tab in real time
- New/deleted column cells correctly show as added/deleted (not modified) in All Changes tab
- Modified row count only increments for real cell changes, not new-column additions
- Lock icon (🔒) on column headers shows pin direction (left vs right)
- Improved hint bar with styled keyboard shortcuts
- Column name no longer cut off in All Changes tab
- Summary row counts fixed to reflect row-level changes only

### v1.1.0
- CSV file support (`.csv` accepted alongside `.xlsx`)

### v1.0.0
- Initial release: row diff, inline cell diff, smart key detection, positional fallback, date formatting, 3-tab view, column/row locking, resize, Hide Unchanged

## Roadmap

### v1.3 — Export Diff
"Export" button in the grid toolbar. Downloads an `.xlsx` file with the diff results — added rows in green, deleted in red, modified with old/new values side by side.

### v1.4 — Merge View
Side-by-side two-panel table (File A left, File B right), scrolled in sync. Modified cells highlighted in both panels.

### v1.5 — Share / Permalink
"Share" button that serializes the current diff into a compressed base64 URL hash. Anyone opening that URL sees the same diff without re-uploading files.

## Dependencies

- [SheetJS](https://sheetjs.com/) — loaded from CDN, no install needed
