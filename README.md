# XLSX Compare

A lightweight, zero-install browser tool for comparing two `.xlsx` files side by side — with row-level diffing, inline change highlighting, and an interactive grid.

## Features

- **Row-level diff** — auto-detects the unique key column(s) per sheet to correctly match rows across files, no manual config needed
- **Inline cell diff** — modified cells show old value (red block, strikethrough) and new value (green block) stacked in a single cell
- **Smart row matching** — falls back to positional (index-based) matching when no unique key column is found, so changed rows show as modified instead of deleted + added
- **New/deleted column detection** — new columns show plain green, deleted columns show plain red; both always visible even with Hide Unchanged on
- **Date formatting** — Excel date serial numbers are rendered as readable dates using the cell's own format string
- **3-tab view**
  - **Summary** — cards per sheet showing added / deleted / modified row counts, click to jump to grid
  - **Sheet Grid** — full data grid with green (added), red (deleted), yellow (modified) highlights
  - **All Changes** — flat filterable list of every change with old and new values
- **Interactive grid**
  - Lock columns left or right (sticky scroll) — click header to pin left, Shift+click to pin right
  - Lock rows — click row number to pin
  - Resize columns and rows by dragging edges
  - Hide unchanged rows/columns toggle
- **No install, no server** — open `index.html` directly in any browser

## Usage

1. Open `index.html` in a browser
2. Drop or select **File A** (original) and **File B** (modified)
3. Click **Compare**

## Files

| File | Purpose |
|------|---------|
| `index.html` | UI layout and styles |
| `script.js` | All comparison logic and rendering |

## Roadmap

### v1.1 — CSV Support
Accept `.csv` files alongside `.xlsx`. Parse with SheetJS (`XLSX.read(text, {type:'string'})`). Dropzones update to accept `.xlsx,.csv`. Single-sheet result since CSV has no sheet concept.

### v1.2 — Ignore Columns
Click a column header to toggle it as "ignored" — excluded from diff comparison entirely. Ignored columns still show in the grid but are never highlighted as changed. Stored as a per-sheet `ignoredCols` Set. Small "ignored" badge on the header.

### v1.3 — Export Diff
"Export" button in the grid toolbar. Downloads an `.xlsx` file with the diff results — added rows in green, deleted in red, modified with old/new values side by side. Uses SheetJS write API (already loaded).

### v1.4 — Merge View
New toggle in the grid toolbar: "Merge View" switches from the current inline old → new single-row layout to a side-by-side two-panel table (File A left, File B right), scrolled in sync. Modified cells highlighted in both panels.

### v1.5 — Share / Permalink
"Share" button that serializes the current diff result into a compressed base64 string and puts it in `window.location.hash`. Anyone opening that URL sees the same diff without re-uploading files. Uses `CompressionStream` API (built into modern browsers, no CDN).

## Dependencies

- [SheetJS](https://sheetjs.com/) — loaded from CDN, no install needed
