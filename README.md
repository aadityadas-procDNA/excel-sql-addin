# SQL Query: an Excel add-in

Run SQL (SQLite dialect) against Excel tables, named ranges, whole sheets and any cell range in the open workbook. Results can be previewed, inserted on a new sheet or at the selected cell, and optionally kept as a **refreshable table** that re-runs its query on demand.

## Install (current user, no admin)

```powershell
.\scripts\install.ps1 -AutoStart
```

- Creates a localhost HTTPS certificate and trusts it for your user (Windows asks you to confirm).
- Registers `manifest.xml` with Excel desktop.
- `-AutoStart` starts the local server at logon. Without it, run `npm start` before using the add-in.

Restart Excel, then go to **Data > SQL Query**.

Uninstall with `.\scripts\uninstall.ps1`.

## Referencing data

| Source        | SQL                                   |
|---------------|---------------------------------------|
| Excel table   | `SELECT * FROM Orders`                |
| Named range   | `SELECT * FROM Targets`               |
| Whole sheet   | `SELECT * FROM [Sheet 1]`             |
| Any range     | `SELECT * FROM [Sheet1!A1:D200]`      |

- The first row of a range is used as the header row. Quote names that contain spaces: `[Unit Price]`.
- Date-formatted cells become ISO text (`2026-03-14`), so `strftime`, `date()` and text comparisons work.
- Only the objects a query mentions are loaded, and they are read fresh on every run.

## Example (sample.xlsx)

```sql
SELECT c.Region, t.Target,
       ROUND(SUM(o.Qty * o.[Unit Price]), 2) AS Revenue,
       ROUND(SUM(o.Qty * o.[Unit Price]) * 100.0 / t.Target, 1) AS PctOfTarget
FROM Orders o
JOIN Customers c ON c.CustomerId = o.CustomerId
JOIN Targets t   ON t.Region = c.Region
WHERE o.OrderDate >= '2026-01-01'
GROUP BY c.Region, t.Target
ORDER BY Revenue DESC;
```

## Lookup columns (XLOOKUP replacement)

**Lookup** tab: choose the table to add columns to, where to look values up, the key column(s) and the columns to bring over. The values are written once as plain values, with no formulas recalculating, so this stays fast on large tables.

- **Ignore case** matches like XLOOKUP. **Loose** also ignores extra spaces and matches numbers stored as text (`"123"` = `123`), the usual causes of `#N/A`.
- Composite keys (several key columns) are supported. You can choose whether duplicates use the first or last match, and what text to show when nothing is found.
- The status line reports matched, not-found and duplicate-key counts.
- Number formats (dates, currency) are copied from the source column.

## Join builder

**Join** tab: pick a starting table, then add joins (LEFT / INNER / FULL). Keys are guessed from shared column names, and you choose the columns. It generates SQL that you can run directly or open in the editor. Use the **Ignore case** and **ignore spaces** options to match like XLOOKUP.

## Auto-refresh

**Manage** tab: turn on auto-refresh to update refreshable query tables and lookup columns whenever their source cells change. It runs in a shared runtime, so it keeps working with the pane closed, and the add-in loads automatically with that workbook. Changes made by the add-in itself are ignored, so it can't loop.

Limitations: only cell edits trigger it. Values that change because a formula recalculated do not trigger it; use **Refresh everything** for those.

## Tests

`npm test` runs the SQL and lookup logic tests, plus UI tests. The UI tests drive the real task pane in jsdom against an in-memory fake of the Excel API.

## Layout

- `public/sqlcore.js`: SQL and data conversion logic (tested by `npm test`)
- `public/taskpane.*`: task pane UI and Excel I/O
- `public/lookup.js`, `public/joinbuilder.js`, `public/autorefresh.js`: the features above
- `server.js`: static HTTPS server on `https://localhost:3021` (loopback only)
- `manifest.xml`: add-in manifest (ribbon button on the Data tab)

Saved queries and query-table definitions are stored inside the workbook, so they travel with the file.
