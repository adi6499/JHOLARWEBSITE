# Sheets Dashboard

Offline dashboard + "what's wrong" checker for Excel / Google Sheets downloads.

- **Read-only.** The app never edits, uploads or deletes anything in Google Drive.
- **Stays on this PC.** Files are read inside the app; saved copies live only in this PC's browser storage.

## Start (localhost)
Double-click **Start Dashboard.bat** — it opens http://localhost:5180 and fetches the sheets listed in `sources.json` automatically (refreshes every 5 min, ↻ to refresh now).
(First time only: run `npm install` in this folder.)

## Where the data comes from
- **data/ folder** — every Excel file placed here is shown. To update, replace the file (same name) and press Refresh.
- **sources.json** — Google Drive files shared "Anyone with the link → Viewer". Copy `sources.example.json` to `sources.json` and put the file IDs in. It is git-ignored on purpose: a Drive ID is enough to download a link-shared file.

## Use (old notes)
1. In Google Sheets: File → Download → Microsoft Excel (.xlsx).
2. Drop the file on the left. Dropping a file with the same name replaces the old copy.
3. Pick a tab → **Dashboard**, **What's wrong** (row numbers match the sheet), **Data**.
4. Optional: paste a Drive link (only works if the file is shared "Anyone with the link → Viewer"); ↻ re-downloads it.

## Screens
- **Upload** — drag & drop, choose files, choose folder; Replace / Remove per file.
- **Overview** — data health, critical issues, warnings, spreadsheet count; "Needs attention" list; next-step fix list; one tile per tab.
- **Per tab** — Summary (KPIs + charts, or Target vs Achieved), Problems (Critical / Warnings / Suggestions, each with what · where · why · what to do), Rows (Cards ⇄ Table, shared search + filters, "View in table" locates the row).
- **Global search** (press /) — tabs and rows across every file.

## Files
- `server.js` — local server, 127.0.0.1 only
- `public/analyze.js` — header detection, column types, all data checks (pure, testable in Node)
- `public/app.js` — screens and charts
- `public/vendor/` — SheetJS, Chart.js and the Lucide icons used (icons.js), built by `npm install` / `node scripts/copy-vendor.js` (offline)

## Next step
Wrap as a Windows app (.exe) with Electron.
