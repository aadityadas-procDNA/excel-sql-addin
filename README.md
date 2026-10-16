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

## Layout

- `public/sqlcore.js`: SQL and data conversion logic (tested by `npm test`)
- `public/taskpane.*`: task pane UI and Excel I/O
- `server.js`: static HTTPS server on `https://localhost:3021` (loopback only)
- `manifest.xml`: add-in manifest (ribbon button on the Data tab)

Saved queries and query-table definitions are stored inside the workbook, so they travel with the file.
