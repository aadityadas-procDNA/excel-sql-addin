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

test("lookup modes match like XLOOKUP, loose and exact", () => {
  const src = [["C01", "Acme", 10], ["c02 ", "Globex", 20], [123, "Num", 30], ["C01", "Acme dup", 40], ["", "blank", 0]];
  const target = [["c01"], ["C02"], ["123"], [" 123 "], ["zzz"], [""]];

  const nocase = S.buildLookupIndex(src, [0], "nocase", "first");
  assert.strictEqual(nocase.duplicates, 1);
  let r = S.applyLookup(target, nocase.index, [1, 2], "nocase", "#N/A");
  assert.deepStrictEqual(r.columns[0], ["Acme", "#N/A", "#N/A", "#N/A", "#N/A", ""]);
  assert.strictEqual(r.matched, 1);
  assert.strictEqual(r.missing, 4);

  const loose = S.buildLookupIndex(src, [0], "loose", "last");
  r = S.applyLookup(target, loose.index, [1, 2], "loose", "");
  assert.deepStrictEqual(r.columns[0], ["Acme dup", "Globex", "Num", "Num", "", ""]);
  assert.deepStrictEqual(r.columns[1], [40, 20, 30, 30, "", ""]);

  const exact = S.buildLookupIndex(src, [0], "exact", "first");
  r = S.applyLookup([["C01"], ["c01"]], exact.index, [1], "exact", "none");
  assert.deepStrictEqual(r.columns[0], ["Acme", "none"]);
});

test("lookup with composite keys and formula-like text", () => {
  const src = [["N", 2026, "=bad"], ["N", 2025, "ok"]];
  const { index } = S.buildLookupIndex(src, [0, 1], "nocase", "first");
  const r = S.applyLookup([["n", 2026], ["N", 2025], ["N", 2024]], index, [2], "nocase", "");
  assert.deepStrictEqual(r.columns[0], ["'=bad", "ok", ""]);
});

test("parseAddress, intersects, contains", () => {
  assert.deepStrictEqual(S.parseAddress("B2:D9"), [{ r1: 1, r2: 8, c1: 1, c2: 3 }]);
  assert.deepStrictEqual(S.parseAddress("Sheet1!$A$1"), [{ r1: 0, r2: 0, c1: 0, c2: 0 }]);
  assert.strictEqual(S.parseAddress("C:E")[0].r2, 1048575);
  assert.strictEqual(S.parseAddress("3:5")[0].c2, 16383);
  assert.strictEqual(S.parseAddress("A1,AA10:AB12")[1].c1, 26);
  const t = S.boundsOf(0, 0, 61, 6);
  assert.ok(S.intersects(t, S.parseAddress("F61")[0]));
  assert.ok(!S.intersects(t, S.parseAddress("G1")[0]));
  assert.ok(S.contains(t, S.parseAddress("B2:C3")[0]));
  assert.ok(!S.contains(t, S.parseAddress("A:A")[0]));
});

test("buildJoinSql and guessJoinKey", async () => {
  assert.deepStrictEqual(S.guessJoinKey(["OrderId", "CustomerId", "Name"], ["customerid", "Name"]), { left: "CustomerId", right: "customerid" });
  assert.strictEqual(S.guessJoinKey(["a"], ["b"]), null);
  const sql = S.buildJoinSql({
    base: { name: "Orders", alias: "a" },
    joins: [{ type: "LEFT", name: "My Customers", alias: "b", on: [{ leftAlias: "a", leftColumn: "CustomerId", rightColumn: "CustomerId" }] }],
    columns: [{ alias: "a", column: "OrderId" }, { alias: "b", column: "Name", as: "Name (My Customers)" }],
    nocase: true,
    trim: true,
  });
  assert.ok(sql.includes("LEFT JOIN [My Customers] AS b"));
  assert.ok(sql.includes("TRIM(a.CustomerId) = TRIM(b.CustomerId) COLLATE NOCASE"));

  const SQL = await initSqlJs();
  const db = new SQL.Database();
  S.loadTable(db, "Orders", S.prepareTable(["OrderId", "CustomerId"], [[1, "c01"], [2, "C02 "], [3, "X"]], []));
  S.loadTable(db, "My Customers", S.prepareTable(["CustomerId", "Name"], [["C01", "Acme"], ["C02", "Globex"]], []));
  const res = S.runQuery(db, sql);
  assert.deepStrictEqual(res.columns, ["OrderId", "Name (My Customers)"]);
  assert.deepStrictEqual(res.values, [[1, "Acme"], [2, "Globex"], [3, null]]);
  db.close();
});

// ---- reference mode

const gridReader = (cell) => {
  let calls = 0;
  const read = async (list) => {
    calls++;
    return list.map((b) =>
      Array.from({ length: b.r2 - b.r1 + 1 }, (_, i) => Array.from({ length: b.c2 - b.c1 + 1 }, (_, j) => cell(b.r1 + i, b.c1 + j)))
    );
  };
  return { read, calls: () => calls };
};

test("currentRegion grows like Ctrl+A, including diagonals, and stops at gaps", async () => {
  const rows = [
    "......ZZ",
    ".AB.....",
    ".CD.....",
    "..E.....",
    "...Y....",
    "........",
    "A.......",
  ];
  const { read } = gridReader((r, c) => (rows[r] && rows[r][c] !== "." && rows[r][c]) || "");
  const used = { r1: 0, c1: 0, r2: 6, c2: 7 };
  assert.deepStrictEqual(await S.currentRegion({ r1: 2, c1: 2, r2: 2, c2: 2 }, used, read), { r1: 1, c1: 1, r2: 4, c2: 3 });
  assert.deepStrictEqual(await S.currentRegion({ r1: 0, c1: 7, r2: 0, c2: 7 }, used, read), { r1: 0, c1: 6, r2: 0, c2: 7 });
  assert.strictEqual(await S.currentRegion({ r1: 5, c1: 5, r2: 5, c2: 5 }, used, read), null, "empty cell");
  assert.strictEqual(await S.currentRegion({ r1: 9, c1: 0, r2: 9, c2: 0 }, used, read), null, "outside used range");
});

test("currentRegion on a large block needs few reads", async () => {
  const g = gridReader((r, c) => (r < 5000 && c < 3 ? 1 : ""));
  const region = await S.currentRegion({ r1: 2500, c1: 1, r2: 2500, c2: 1 }, { r1: 0, c1: 0, r2: 5000, c2: 2 }, g.read);
  assert.deepStrictEqual(region, { r1: 0, c1: 0, r2: 4999, c2: 2 });
  assert.ok(g.calls() < 15, "reads: " + g.calls());
});

test("parseNameFormula", () => {
  assert.deepStrictEqual(S.parseNameFormula("='My Sheet'!$F$1:$G$4"), { sheet: "My Sheet", bounds: { r1: 0, c1: 5, r2: 3, c2: 6 } });
  assert.deepStrictEqual(S.parseNameFormula("=Data!$A:$C").bounds, { r1: 0, c1: 0, r2: 1048575, c2: 2 });
  assert.strictEqual(S.parseNameFormula("=OFFSET(Data!$A$1,0,0,10,2)"), null);
  assert.strictEqual(S.parseNameFormula("=Data!$A$1,Data!$C$1"), null);
  assert.strictEqual(S.parseNameFormula("=0.2"), null);
  assert.strictEqual(S.toA1({ r1: 0, c1: 0, r2: 8, c2: 27 }), "A1:AB9");
});

test("sqlLiteral", () => {
  assert.strictEqual(S.sqlLiteral(12.5, "General"), "12.5");
  assert.strictEqual(S.sqlLiteral(true), "1");
  assert.strictEqual(S.sqlLiteral("O'Neil", "@"), "'O''Neil'");
  assert.strictEqual(S.sqlLiteral(46023.5, "yyyy-mm-dd hh:mm"), "'2026-01-01 12:00:00'");
  assert.strictEqual(S.sqlLiteral("", "General"), null);
});

test("reference choice: table, named range, sheet, range, columns, value", () => {
  const B = (a) => S.parseAddress(a)[0];
  const base = {
    sheet: "Data 1",
    used: B("A1:J20"),
    tables: [{ name: "Orders", bounds: B("A1:C10") }],
    tableNames: ["orders", "customers"],
    names: [
      { name: "Rates", sheet: "Data 1", bounds: B("E1:F5") },
      { name: "Customers", sheet: "Data 1", bounds: B("H1:H3") }, // shadowed by a table
    ],
  };
  const choose = (sel, region, headers = ["Id", "", "Id"], cell) => {
    const info = { ...base, sel: B(sel), region };
    const pick = S.pickReferenceBlock(info);
    return pick.needsRegion ? pick : S.referenceCandidates(info, pick, headers, ["orders", "customers", "rates"], cell);
  };
  const texts = (r) => r.candidates.map((c) => `${c.kind}:${c.text}`);

  assert.deepStrictEqual(texts(choose("B5")), ["table:Orders", "column:Column2"]);
  assert.deepStrictEqual(texts(choose("B5", undefined, undefined, { value: "C'02", format: "General" })), ["table:Orders", "column:Column2", "value:'C''02'"]);
  assert.deepStrictEqual(texts(choose("A5", undefined, undefined, { value: 46023, format: "dd-mmm-yy" })), ["table:Orders", "column:Id", "value:'2026-01-01'"]);
  assert.deepStrictEqual(texts(choose("A2:B4")), ["table:Orders", "columns:Id, Column2"]);
  assert.deepStrictEqual(texts(choose("A1:B4")), ["table:Orders", "range:[Data 1!A1:B4]", "columns:Id, Column2"]);
  assert.deepStrictEqual(texts(choose("A1:C10")), ["table:Orders"]);
  assert.deepStrictEqual(texts(choose("C1")), ["table:Orders", "column:Id_2"]);
  assert.deepStrictEqual(texts(choose("F2", undefined, ["Rate", "Value"])), ["name:Rates", "range:[Data 1!E1:F5]", "column:Value"]);
  assert.deepStrictEqual(texts(choose("E1:F5", undefined, ["Rate", "Value"])), ["name:Rates", "range:[Data 1!E1:F5]"]);
  assert.deepStrictEqual(choose("H2"), { needsRegion: true }, "named range shadowed by a table name is skipped");
  assert.deepStrictEqual(texts(choose("H2", B("H1:J20"))), ["range:[Data 1!H1:J20]", "column:Id"]);
  assert.deepStrictEqual(texts(choose("A1:Z99")), ["sheet:[Data 1]", "range:[Data 1!A1:J20]"]);
  assert.deepStrictEqual(texts(choose("H1", B("H1:I4"), ["Region", "Target"])), ["range:[Data 1!H1:I4]", "column:Region"]);
  assert.deepStrictEqual(choose("Z50", null).candidates, []);
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
