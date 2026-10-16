// Drives the real task pane scripts in jsdom against the fake Excel workbook.
const fs = require("fs");
const path = require("path");
const vm = require("vm");
const assert = require("assert");
const { JSDOM } = require("jsdom");
const initSqlJs = require("sql.js");
const { Workbook, makeOffice } = require("./fake-excel");

const PUB = path.join(__dirname, "..", "public");
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

function buildWorkbook() {
  const wb = new Workbook();
  const orders = wb.addSheet("Orders");
  orders.write("A1", [
    ["OrderId", "OrderDate", "CustomerId", "Qty"],
    [1, 46023, "C01", 2],
    [2, 46024, "c02", 5],
    [3, 46025, " C03 ", 1],
    [4, 46026, "C99", 7],
    [5, 46027, 123, 1],
  ]);
  orders.getRange("B2:B6").numberFormat = "yyyy-mm-dd";
  wb.addTable(orders, "Orders", "A1:D6");
  const cust = wb.addSheet("Customers");
  cust.write("A1", [
    ["CustomerId", "Name", "Region", "Since"],
    ["C01", "Acme", "North", 45292],
    ["C02", "Globex", "South", 45300],
    ["C03", "Initech", "North", 45310],
    ["123", "Numeric Co", "West", 45320],
    ["C01", "Acme Dup", "East", 45330],
  ]);
  cust.getRange("D2:D6").numberFormat = "dd-mmm-yy";
  wb.addTable(cust, "Customers", "A1:D6");
  cust.write("F1", [["Region", "Target"], ["North", 10], ["South", 5], ["West", 3]]);
  wb.addName("Targets", cust, "F1:G4");
  return wb;
}

async function boot(wb) {
  const html = fs.readFileSync(path.join(PUB, "taskpane.html"), "utf8").replace(/<script[^>]*\ssrc=[^>]*><\/script>/g, "");
  const dom = new JSDOM(html, { url: "https://localhost:3021/taskpane.html", runScripts: "outside-only", pretendToBeVisual: true });
  const w = dom.window;
  const { Office, Excel } = makeOffice(wb);
  Object.assign(w, { Office, Excel, OfficeExtension: { Error: class extends Error {} }, initSqlJs: () => initSqlJs() });
  const ctx = dom.getInternalVMContext();
  for (const f of ["sqlcore.js", "taskpane.js", "joinbuilder.js", "lookup.js", "autorefresh.js"]) {
    new vm.Script(fs.readFileSync(path.join(PUB, f), "utf8"), { filename: f }).runInContext(ctx);
  }
  const g = (expr) => new vm.Script(expr).runInContext(ctx);
  await Office.boot();
  const doc = w.document;
  const $ = (id) => doc.getElementById(id);
  return {
    w, doc, g, Office, $,
    status: () => $("status").textContent,
    select(id, value) {
      $(id).value = value;
      assert.strictEqual($(id).value, value, `option ${value} missing in #${id}`);
      $(id).dispatchEvent(new w.Event("change"));
    },
    click(id) { $(id).click(); },
    async idle(ms = 30) {
      await sleep(ms);
      for (let i = 0; i < 400; i++) {
        if (g("busyCount") === 0) return;
        await sleep(25);
      }
      throw new Error("timed out waiting for the pane to finish");
    },
  };
}

(async () => {
  const wb = buildWorkbook();
  const ui = await boot(wb);
  await ui.idle();
  const step = (name) => console.log("ok   " + name);

  // ---- objects tree
  assert.strictEqual(ui.doc.querySelectorAll("#objects > li").length, 5, ui.status());
  step("workbook objects listed (2 tables, 1 named range, 2 sheets)");

  // ---- lookup columns
  ui.select("lkTarget", "Orders");
  ui.select("lkSource", "Customers");
  const keySelects = ui.doc.querySelectorAll("#lkKeys select");
  assert.deepStrictEqual([keySelects[0].value, keySelects[1].value], ["CustomerId", "CustomerId"]);
  ui.click("lkAll");
  assert.strictEqual(ui.doc.querySelectorAll("#lkColumns input:checked").length, 3);
  ui.select("lkMode", "loose");
  ui.$("lkNotFound").value = "N/A";
  ui.click("lkApply");
  await ui.idle();
  const orders = wb.tableValues("Orders");
  assert.deepStrictEqual(orders[0], ["OrderId", "OrderDate", "CustomerId", "Qty", "Name", "Region", "Since"], ui.status());
  assert.deepStrictEqual(orders.slice(1).map((r) => r[4]), ["Acme", "Globex", "Initech", "N/A", "Numeric Co"]);
  assert.strictEqual(orders[1][6], 45292);
  assert.strictEqual(wb.tables.getItem("Orders").getDataBodyRange().getCell(0, 6).numberFormat[0][0], "dd-mmm-yy");
  assert.match(ui.status(), /4 of 5 rows matched, 1 not found/);
  assert.match(ui.status(), /1 duplicate key/);
  assert.strictEqual(ui.Office.context.document.settings.get("sqlLookups").length, 1);
  step("lookup columns: loose match, first-of-duplicates, not-found text, date format copied");

  // ---- join builder
  ui.select("joinBase", "Orders");
  ui.click("joinAdd");
  const sql = ui.$("joinPreview").textContent;
  assert.match(sql, /LEFT JOIN Customers AS b\n\s+ON a\.CustomerId = b\.CustomerId COLLATE NOCASE/, sql);
  assert.match(sql, /b\.Name AS \[Name \(Customers\)\]/, sql);
  ui.click("joinRun");
  await ui.idle();
  assert.strictEqual(ui.g("lastResult.values.length"), 6, ui.status());
  step("join builder: guessed key, NOCASE, clash aliasing, runs (6 rows incl. duplicate customer)");

  ui.click("toSheet");
  await ui.idle();
  const linked = ui.Office.context.document.settings.get("sqlQueryTables");
  assert.deepStrictEqual(Object.keys(linked), ["SQL_Result_1"], ui.status());
  assert.strictEqual(wb.tableValues("SQL_Result_1").length, 7);
  step("result inserted as refreshable table SQL_Result_1");

  // ---- auto-refresh
  let lookupRuns = 0;
  let queryRuns = 0;
  ui.w.__count = (k) => (k === "l" ? lookupRuns++ : queryRuns++);
  ui.g(`{ const L = refreshLookup, Q = refreshLinked;
          refreshLookup = (...a) => { __count("l"); return L(...a); };
          refreshLinked = (...a) => { __count("q"); return Q(...a); }; }`);

  ui.$("autoRefresh").checked = true;
  ui.$("autoRefresh").dispatchEvent(new ui.w.Event("change"));
  await ui.idle();
  assert.strictEqual(ui.Office.startup, "Load");
  assert.ok(ui.g("auto.deps.length") >= 4, "deps: " + ui.g("JSON.stringify(auto.deps)"));
  step("auto-refresh on: add-in set to load with the workbook, dependencies mapped");

  // A user edit in the lookup source refreshes the lookup and the query table.
  const cust = wb.worksheets.getItem("Customers");
  cust.write("B2", [["Acme Renamed"]]);
  await sleep(1600);
  await ui.idle();
  assert.strictEqual(wb.tableValues("Orders")[1][4], "Acme Renamed", ui.status());
  assert.ok(wb.tableValues("SQL_Result_1").some((r) => r.includes("Acme Renamed")), ui.status());
  assert.match(ui.status(), /Auto-refreshed/);
  step(`edit in Customers -> lookup + query table auto-refreshed (lookup ×${lookupRuns}, query ×${queryRuns})`);

  // The add-in's own writes must not trigger more refreshes.
  const before = lookupRuns + queryRuns;
  await sleep(2500);
  await ui.idle();
  assert.strictEqual(lookupRuns + queryRuns, before, "refresh loop detected");
  step("no refresh loop from the add-in's own writes");

  // Changing a key in the target table re-runs the lookup.
  wb.worksheets.getItem("Orders").write("C5", [["C02"]]);
  await sleep(1600);
  await ui.idle();
  assert.strictEqual(wb.tableValues("Orders")[4][4], "Globex", ui.status());
  step("key change in Orders -> lookup re-run");

  // Edits elsewhere are ignored.
  const runs = lookupRuns + queryRuns;
  wb.worksheets.getItem("Customers").write("K20", [["unrelated"]]);
  await sleep(1600);
  await ui.idle();
  assert.strictEqual(lookupRuns + queryRuns, runs);
  step("unrelated edit ignored");

  ui.$("autoRefresh").checked = false;
  ui.$("autoRefresh").dispatchEvent(new ui.w.Event("change"));
  await ui.idle();
  assert.strictEqual(wb.handlers.length, 0);
  assert.strictEqual(ui.Office.startup, "None");
  step("auto-refresh off: event handler removed");

  // ---- manual refresh-all still works
  ui.click("refreshAll");
  await ui.idle();
  assert.match(ui.status(), /Refreshed 2 items/, ui.status());
  step("refresh everything");

  console.log("\nUI tests passed");
  process.exit(0);
})().catch((e) => {
  console.error("FAIL", e);
  process.exit(1);
});
