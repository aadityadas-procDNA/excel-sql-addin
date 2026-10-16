const assert = require("assert");
const initSqlJs = require("sql.js");
const S = require("../public/sqlcore.js");

const tests = [];
const test = (name, fn) => tests.push({ name, fn });

test("extractIdentifiers skips strings and comments", () => {
  const ids = S.extractIdentifiers(
    "SELECT a.[Unit Price], \"Qty\" FROM Sales a -- Hidden\nJOIN [My Sheet!A1:C9] b ON 1 /* Other */ WHERE x = 'Regions'''"
  );
  assert.ok(ids.includes("Sales"));
  assert.ok(ids.includes("Unit Price"));
  assert.ok(ids.includes("Qty"));
  assert.ok(ids.includes("My Sheet!A1:C9"));
  assert.ok(!ids.includes("Hidden"));
  assert.ok(!ids.includes("Other"));
  assert.ok(!ids.includes("Regions"));
});

test("resolveSources prefers tables, then names, then sheets", () => {
  const objs = [
    { kind: "sheet", name: "Sales" },
    { kind: "table", name: "Sales" },
    { kind: "name", name: "Rates" },
    { kind: "sheet", name: "Data 2024" },
  ];
  const src = S.resolveSources(["sales", "SALES", "Rates", "Data 2024", "Sheet1!A1:B2", "col"], objs);
  assert.deepStrictEqual(
    src.map((s) => [s.kind, s.ref]),
    [["table", "Sales"], ["name", "Rates"], ["sheet", "Data 2024"], ["range", "Sheet1!A1:B2"]]
  );
});

test("splitAddress handles quoted sheet names", () => {
  assert.deepStrictEqual(S.splitAddress("'Bob''s Data'!A1:C3"), { sheet: "Bob's Data", address: "A1:C3" });
  assert.deepStrictEqual(S.splitAddress("Sheet1!A:D"), { sheet: "Sheet1", address: "A:D" });
});

test("quoteRef", () => {
  assert.strictEqual(S.quoteRef("Sales"), "Sales");
  assert.strictEqual(S.quoteRef("Unit Price"), "[Unit Price]");
  assert.strictEqual(S.quoteRef("Order"), "[Order]");
});

test("normalizeHeaders", () => {
  assert.deepStrictEqual(S.normalizeHeaders(["id", "", "ID", "name", null]), ["id", "Column2", "ID_2", "name", "Column5"]);
});

test("date formats and serial conversion", () => {
  assert.ok(S.isDateFormat("m/d/yyyy"));
  assert.ok(S.isDateFormat("[$-409]dd-mmm-yy;@"));
  assert.ok(S.isDateFormat("yyyy-mm-dd hh:mm"));
  assert.ok(!S.isDateFormat("General"));
  assert.ok(!S.isDateFormat("#,##0.00"));
  assert.ok(!S.isDateFormat('0.00" days"'));
  assert.ok(!S.isDateFormat("[h]:mm:ss"));
  assert.strictEqual(S.serialToIso(45292), "2024-01-01");
  assert.strictEqual(S.serialToIso(45292.5), "2024-01-01 12:00:00");
  assert.strictEqual(S.serialToIso(0.75), "18:00:00");
});

test("prepareTable infers types and converts values", () => {
  const t = S.prepareTable(
    ["Date", "Qty", "Price", "Name", "Flag", "Mixed", ""],
    [
      [45292, 2, 1.5, "a", true, 1, ""],
      ["", "", "", "", "", "", ""],
      [45293, 3, 2, "b", false, "x", ""],
    ],
    [["m/d/yyyy", "General", "0.00", "@", "General", "General", "General"]]
  );
  assert.deepStrictEqual(t.columns.map((c) => c.type), ["DATE", "INTEGER", "REAL", "TEXT", "BOOLEAN", "ANY", "ANY"]);
  assert.strictEqual(t.rows.length, 2);
  assert.deepStrictEqual(t.rows[0], ["2024-01-01", 2, 1.5, "a", 1, 1, null]);
});

test("end to end: load tables, join, aggregate, output", async () => {
  const SQL = await initSqlJs();
  const db = new SQL.Database();
  S.loadTable(db, "Sales", S.prepareTable(
    ["Region", "Date", "Amount"],
    [["N", 45292, 10], ["S", 45293, 20], ["N", 45324, 5]],
    [["General", "d-mmm-yy", "General"]]
  ));
  S.loadTable(db, "My Sheet!A1:B3", S.prepareTable(["Region", "Manager"], [["N", "Ann"], ["S", "=Bob"]], []));
  const res = S.runQuery(db, `
    WITH m AS (SELECT * FROM [my sheet!a1:b3])
    SELECT s.Region, m.Manager, SUM(Amount) AS Total, strftime('%m', MIN(Date)) AS FirstMonth,
           RANK() OVER (ORDER BY SUM(Amount) DESC) AS rnk
    FROM Sales s JOIN m ON m.Region = s.Region
    GROUP BY s.Region ORDER BY s.Region;`);
  assert.deepStrictEqual(res.columns, ["Region", "Manager", "Total", "FirstMonth", "rnk"]);
  assert.deepStrictEqual(res.values, [["N", "Ann", 15, "01", 2], ["S", "=Bob", 20, "01", 1]]);
  const out = S.toExcelValues(res);
  assert.strictEqual(out.length, 3);
  assert.strictEqual(out[2][1], "'=Bob");
  assert.strictEqual(S.toExcelValues({ columns: ["a"], values: [] }, true).length, 2);
  db.close();
});

(async () => {
  let failed = 0;
  for (const t of tests) {
    try {
      await t.fn();
      console.log("ok   " + t.name);
    } catch (e) {
      failed++;
      console.log("FAIL " + t.name + "\n     " + e.message);
    }
  }
  console.log(`\n${tests.length - failed}/${tests.length} passed`);
  process.exit(failed ? 1 : 0);
})();
