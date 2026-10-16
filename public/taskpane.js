/* global Office, Excel, OfficeExtension, initSqlJs, SqlCore,
   initJoinBuilder, renderJoinBuilder, initLookup, renderLookupForm, renderLookups, refreshLookup,
   initAutoRefresh, markSelfWrite, rebuildDeps, initRefMode */
"use strict";

const SETTING_SAVED = "sqlSavedQueries"; // [{ name, sql }]
const SETTING_LINKED = "sqlQueryTables"; // { tableName: { sql, refreshed } }
const PREVIEW_ROWS = 200;
const CELLS_PER_CHUNK = 100000; // keeps each Excel round trip well under payload limits
const FORMAT_SAMPLE_ROWS = 25;

let SQL = null;
let lastResult = null; // { columns, values, sql }
let objectCache = []; // described workbook objects: { kind, name, columns, rows, where }

const $ = (id) => document.getElementById(id);

Office.onReady(async (info) => {
  if (info.host !== Office.HostType.Excel) {
    setStatus("This add-in only runs in Excel.", "error");
    return;
  }
  bindUi();
  initJoinBuilder();
  initLookup();
  initRefMode();
  try {
    $("sql").value = localStorage.getItem("sqlDraft") || "";
  } catch (e) { /* storage unavailable */ }
  try {
    SQL = await initSqlJs({ locateFile: (f) => "vendor/" + f });
  } catch (e) {
    setStatus("Could not load the SQLite engine: " + e.message, "error");
    return;
  }
  renderManage();
  await busy(refreshObjects);
  await busy(initAutoRefresh);
});

// ---------------------------------------------------------------- UI wiring

function bindUi() {
  document.querySelectorAll(".tab").forEach((b) => (b.onclick = () => showTab(b.dataset.tab)));

  const ta = $("sql");
  ta.addEventListener("keydown", (e) => {
    if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) {
      e.preventDefault();
      busy(runFromEditor);
    } else if (e.key === "Tab" && !e.shiftKey) {
      e.preventDefault();
      ta.setRangeText("  ", ta.selectionStart, ta.selectionEnd, "end");
    }
  });
  ta.addEventListener("input", () => {
    try { localStorage.setItem("sqlDraft", ta.value); } catch (e) { /* ignore */ }
  });

  $("run").onclick = () => busy(runFromEditor);
  $("clearBtn").onclick = () => { ta.value = ""; ta.dispatchEvent(new Event("input")); ta.focus(); };
  $("toSheet").onclick = () => busy(() => insertResult("sheet"));
  $("toCell").onclick = () => busy(() => insertResult("cell"));
  $("refreshAll").onclick = () => busy(refreshEverything);
  $("refreshObjects").onclick = (e) => { e.preventDefault(); e.stopPropagation(); busy(refreshObjects); };
  $("filter").oninput = renderObjects;

  $("saveBtn").onclick = () => {
    $("saveRow").hidden = false;
    $("saveName").focus();
  };
  $("saveCancel").onclick = () => { $("saveRow").hidden = true; };
  $("saveOk").onclick = () => busy(saveCurrentQuery);
  $("saveName").addEventListener("keydown", (e) => {
    if (e.key === "Enter") busy(saveCurrentQuery);
    if (e.key === "Escape") $("saveRow").hidden = true;
  });
}

function showTab(name) {
  document.querySelectorAll(".tab").forEach((b) => b.classList.toggle("active", b.dataset.tab === name));
  document.querySelectorAll("[data-panel]").forEach((p) => (p.hidden = p.dataset.panel !== name));
}

function setStatus(msg, kind) {
  const node = $("status");
  node.textContent = msg || "";
  node.className = "status" + (kind ? " " + kind : "");
}

let busyCount = 0;
async function busy(fn) {
  busyCount++;
  document.querySelectorAll("button").forEach((b) => (b.disabled = true));
  try {
    await fn();
  } catch (e) {
    setStatus(errorText(e), "error");
    console.error(e);
  } finally {
    if (--busyCount === 0) document.querySelectorAll("button").forEach((b) => (b.disabled = false));
  }
}

function errorText(e) {
  if (typeof OfficeExtension !== "undefined" && e instanceof OfficeExtension.Error) {
    return "Excel: " + e.message + (e.debugInfo && e.debugInfo.errorLocation ? ` (${e.debugInfo.errorLocation})` : "");
  }
  return e && e.message ? e.message : String(e);
}

// Inserts at the cursor (or over the selected text), adding a space where it would otherwise run into a word.
function insertAtCursor(text) {
  const ta = $("sql");
  const start = ta.selectionStart;
  const end = ta.selectionEnd;
  if (start > 0 && /[\w\]"]/.test(ta.value[start - 1])) text = " " + text;
  if (end < ta.value.length && /[\w["]/.test(ta.value[end])) text += " ";
  ta.focus();
  ta.setRangeText(text, start, end, "end");
  ta.dispatchEvent(new Event("input"));
}

function loadIntoEditor(sql) {
  showTab("query");
  $("sql").value = sql;
  $("sql").dispatchEvent(new Event("input"));
  $("sql").focus();
}

function el(tag, props, ...children) {
  const node = document.createElement(tag);
  Object.assign(node, props || {});
  for (const c of children) if (c !== null && c !== undefined) node.append(c);
  return node;
}

function fillSelect(select, options, value) {
  select.replaceChildren(...options.map((o) => el("option", { value: o.value, textContent: o.label })));
  if (options.some((o) => o.value === value)) select.value = value;
  return select.value;
}

// ---------------------------------------------------------------- workbook objects

async function listObjects(ctx) {
  const wb = ctx.workbook;
  const tables = wb.tables.load("items/name");
  const names = wb.names.load("items/name,items/type,items/visible");
  const sheets = wb.worksheets.load("items/name");
  await ctx.sync();
  return [
    ...tables.items.map((t) => ({ kind: "table", name: t.name })),
    ...names.items.filter((n) => n.type === "Range" && n.visible).map((n) => ({ kind: "name", name: n.name })),
    ...sheets.items.map((s) => ({ kind: "sheet", name: s.name })),
  ];
}

// Finds a workbook object by name (table > named range > sheet), or a "Sheet!A1:B2" range.
function findObject(name, objs) {
  return SqlCore.resolveSources([name], objs)[0] || null;
}

// Range for a source, clipped to the sheet's used area so whole-column references stay small.
function sourceRange(ctx, src) {
  const wb = ctx.workbook;
  switch (src.kind) {
    case "table":
      return wb.tables.getItem(src.ref).getRange();
    case "sheet":
      return wb.worksheets.getItem(src.ref).getUsedRangeOrNullObject(true);
    case "name": {
      const r = wb.names.getItem(src.ref).getRange();
      return r.getIntersectionOrNullObject(r.worksheet.getUsedRange(true));
    }
    case "range": {
      const { sheet, address } = SqlCore.splitAddress(src.ref);
      const ws = wb.worksheets.getItem(sheet);
      return ws.getRange(address).getIntersectionOrNullObject(ws.getUsedRange(true));
    }
  }
  throw new Error("Unknown source kind " + src.kind);
}

// Reads a block of cells in chunks. With formatRows > 0 also returns number formats of the first rows.
async function readBlock(ctx, ws, rowIndex, columnIndex, rowCount, columnCount, formatRows) {
  const step = Math.max(1, Math.floor(CELLS_PER_CHUNK / columnCount));
  const values = [];
  let formats = [];
  for (let r = 0; r < rowCount; r += step) {
    const part = ws.getRangeByIndexes(rowIndex + r, columnIndex, Math.min(step, rowCount - r), columnCount).load("values");
    let fmt = null;
    if (r === 0 && formatRows > 0) {
      fmt = ws.getRangeByIndexes(rowIndex, columnIndex, Math.min(formatRows, rowCount), columnCount).load("numberFormat");
    }
    await ctx.sync();
    for (const row of part.values) values.push(row);
    if (fmt) formats = fmt.numberFormat;
  }
  return { values, formats };
}

// Returns header row, data rows and number formats of the first data rows.
async function readSource(ctx, src) {
  const range = sourceRange(ctx, src);
  range.load("rowIndex,columnIndex,rowCount,columnCount");
  await ctx.sync();
  if (range.isNullObject) return { headers: [], rows: [], formats: [] };
  const { rowIndex, columnIndex, rowCount, columnCount } = range;
  const ws = range.worksheet;
  const head = await readBlock(ctx, ws, rowIndex, columnIndex, 1, columnCount, 0);
  if (rowCount < 2) return { headers: head.values[0], rows: [], formats: [] };
  const body = await readBlock(ctx, ws, rowIndex + 1, columnIndex, rowCount - 1, columnCount, FORMAT_SAMPLE_ROWS);
  return { headers: head.values[0], rows: body.values, formats: body.formats };
}

async function refreshObjects() {
  objectCache = await Excel.run(async (ctx) => {
    const objs = await listObjects(ctx);
    const wb = ctx.workbook;
    const pending = objs.map((o) => {
      if (o.kind === "table") {
        const t = wb.tables.getItem(o.name);
        const head = t.getHeaderRowRange().load("values");
        const body = t.getDataBodyRange().load("rowCount");
        const ws = t.worksheet.load("name");
        return () => ({ ...o, headers: head.values[0], rows: body.rowCount, where: ws.name });
      }
      const range = sourceRange(ctx, { ...o, ref: o.name });
      range.load("address,rowCount");
      return () => (range.isNullObject ? { ...o, empty: true } : { ...o, range, rows: range.rowCount - 1, where: range.address });
    });
    await ctx.sync();
    const described = pending.map((f) => f());

    // Second pass: header rows of named ranges and sheets.
    const heads = described.map((o) => (o.range ? o.range.getRow(0).load("values") : null));
    await ctx.sync();
    return described.map((o, i) => {
      const headers = heads[i] ? heads[i].values[0] : o.headers || [];
      const { range, ...rest } = o;
      return { ...rest, columns: SqlCore.normalizeHeaders(headers) };
    });
  });
  renderObjects();
  renderJoinBuilder();
  renderLookupForm();
}

function columnsOf(name) {
  const o = objectCache.find((x) => x.name === name);
  return o ? o.columns : [];
}

function renderObjects() {
  const list = $("objects");
  const q = $("filter").value.trim().toLowerCase();
  list.replaceChildren();
  const kindLabel = { table: "T", name: "N", sheet: "S" };
  const kindTitle = { table: "Excel table", name: "Named range", sheet: "Worksheet (used range)" };

  const shown = objectCache.filter(
    (o) => !q || o.name.toLowerCase().includes(q) || o.columns.some((c) => c.toLowerCase().includes(q))
  );
  if (!shown.length) {
    list.append(el("li", { className: "muted small", textContent: objectCache.length ? "No matches." : "No tables, named ranges or sheets with data." }));
    return;
  }
  for (const o of shown) {
    const ref = SqlCore.quoteRef(o.name);
    const cols = el("ul", { className: "cols", hidden: !q });
    for (const c of o.columns) {
      cols.append(el("li", { title: "Insert column name", onclick: () => insertAtCursor(SqlCore.quoteRef(c)) }, el("span", { textContent: c })));
    }
    const caret = el("span", { className: "caret", textContent: q ? "▾" : "▸" });
    caret.onclick = () => { cols.hidden = !cols.hidden; caret.textContent = cols.hidden ? "▸" : "▾"; };
    const meta = o.empty ? "empty" : `${Math.max(o.rows, 0).toLocaleString()} rows · ${o.columns.length} cols`;
    const row = el(
      "div", { className: "obj" },
      caret,
      el("span", { className: "kind " + o.kind, textContent: kindLabel[o.kind], title: kindTitle[o.kind] }),
      el("span", { className: "name", textContent: o.name, title: `Insert ${ref}${o.where ? "\n" + o.where : ""}`, onclick: () => insertAtCursor(ref) }),
      el("span", { className: "meta", textContent: meta }),
      el("button", {
        className: "icon", textContent: "▶", title: "Preview first 100 rows",
        onclick: () => {
          loadIntoEditor(`SELECT *\nFROM ${ref}\nLIMIT 100;`);
          busy(runFromEditor);
        },
      })
    );
    list.append(el("li", null, row, cols));
  }
}

// ---------------------------------------------------------------- running queries

async function executeSql(sql) {
  const loaded = [];
  const db = new SQL.Database();
  try {
    await Excel.run(async (ctx) => {
      const objs = await listObjects(ctx);
      const sources = SqlCore.resolveSources(SqlCore.extractIdentifiers(sql), objs);
      for (const src of sources) {
        let data;
        try {
          data = await readSource(ctx, src);
        } catch (e) {
          throw new Error(`Could not read ${src.name}: ${errorText(e)}`);
        }
        const table = SqlCore.prepareTable(data.headers, data.rows, data.formats);
        SqlCore.loadTable(db, src.name, table);
        loaded.push(`${src.name} (${table.rows.length.toLocaleString()})`);
      }
    });
    try {
      return { result: SqlCore.runQuery(db, sql), loaded };
    } catch (e) {
      let msg = "SQL error: " + e.message;
      if (/no such table/i.test(e.message)) {
        msg += "\nUse an Excel table name, a named range, a sheet name ([Sheet 1]) or a range ([Sheet1!A1:D100]).";
      }
      throw new Error(msg);
    }
  } finally {
    db.close();
  }
}

async function runFromEditor() {
  if (!SQL) throw new Error("SQLite engine not loaded.");
  const ta = $("sql");
  const selected = ta.value.substring(ta.selectionStart, ta.selectionEnd);
  const sql = (selected.trim() ? selected : ta.value).trim();
  if (!sql) return setStatus("Type a query first.");

  setStatus("Running…");
  const t0 = performance.now();
  const { result, loaded } = await executeSql(sql);
  const ms = Math.round(performance.now() - t0);

  lastResult = { ...result, sql };
  renderGrid(result);
  const from = loaded.length ? `\nLoaded: ${loaded.join(", ")}` : "";
  if (!result.columns.length) {
    setStatus(`Done in ${ms} ms. No result set returned.${from}`, "ok");
  } else {
    setStatus(`${result.values.length.toLocaleString()} rows × ${result.columns.length} cols · ${ms} ms${from}`, "ok");
  }
}

function renderGrid(result) {
  const section = $("resultSection");
  const grid = $("grid");
  grid.replaceChildren();
  section.hidden = !result.columns.length;
  if (!result.columns.length) return;

  grid.append(el("thead", null, el("tr", null, ...result.columns.map((c) => el("th", { textContent: c, title: c })))));
  const body = el("tbody");
  for (const row of result.values.slice(0, PREVIEW_ROWS)) {
    body.append(
      el("tr", null, ...row.map((v) => {
        if (v === null) return el("td", { className: "null", textContent: "NULL" });
        const text = v instanceof Uint8Array ? `[BLOB ${v.length} bytes]` : String(v);
        return el("td", { className: typeof v === "number" ? "num" : "", textContent: text, title: text });
      }))
    );
  }
  grid.append(body);
  $("gridNote").textContent =
    result.values.length > PREVIEW_ROWS ? `Showing first ${PREVIEW_ROWS} of ${result.values.length.toLocaleString()} rows. Insert to get all rows.` : "";
}

// ---------------------------------------------------------------- writing results

// Writes a 2D block in chunks and records it as our own change (so auto-refresh ignores it).
// formats: optional per-column number formats (null entries are left alone).
async function writeValues(ctx, ws, row0, col0, data, formats) {
  const cols = data[0].length;
  ws.load("name");
  await ctx.sync();
  const step = Math.max(1, Math.floor(CELLS_PER_CHUNK / cols));
  for (let r = 0; r < data.length; r += step) {
    const chunk = data.slice(r, r + step);
    markSelfWrite(ws.name, SqlCore.boundsOf(row0 + r, col0, chunk.length, cols));
    const range = ws.getRangeByIndexes(row0 + r, col0, chunk.length, cols);
    if (formats && formats.some((f) => f)) {
      // Format first, so text such as "001" stays text in "@" columns.
      formats.forEach((f, c) => {
        if (f) ws.getRangeByIndexes(row0 + r, col0 + c, chunk.length, 1).numberFormat = chunk.map(() => [f]);
      });
    }
    range.values = chunk;
    await ctx.sync();
  }
}

async function assertEmpty(ctx, ranges, hint) {
  const used = ranges.map((r) => r.getUsedRangeOrNullObject(true).load("address"));
  await ctx.sync();
  const hit = used.find((u) => !u.isNullObject);
  if (hit) throw new Error(`Stopped: ${hit.address} already contains data. ${hint || "Select an empty area or use “New sheet”."}`);
}

function uniqueName(existing, base, alwaysNumber = true) {
  const taken = new Set(existing.map((s) => s.toLowerCase()));
  if (!alwaysNumber && !taken.has(base.toLowerCase())) return base;
  for (let i = alwaysNumber ? 1 : 2; ; i++) {
    const name = alwaysNumber ? base + i : `${base} ${i}`;
    if (!taken.has(name.toLowerCase())) return name;
  }
}

async function insertResult(mode) {
  if (!lastResult || !lastResult.columns.length) throw new Error("Run a query that returns rows first.");
  const asTable = $("asTable").checked;
  const data = SqlCore.toExcelValues(lastResult, asTable);

  const tableName = await Excel.run(async (ctx) => {
    const wb = ctx.workbook;
    const tables = wb.tables.load("items/name");
    const sheets = wb.worksheets.load("items/name");
    let ws, row0 = 0, col0 = 0;
    if (mode === "sheet") {
      await ctx.sync();
      ws = wb.worksheets.add(uniqueName(sheets.items.map((s) => s.name), "Query "));
    } else {
      const cell = wb.getActiveCell().load("rowIndex,columnIndex");
      ws = cell.worksheet;
      await ctx.sync();
      row0 = cell.rowIndex;
      col0 = cell.columnIndex;
      await assertEmpty(ctx, [ws.getRangeByIndexes(row0, col0, data.length, data[0].length)]);
    }
    await writeValues(ctx, ws, row0, col0, data);
    const full = ws.getRangeByIndexes(row0, col0, data.length, data[0].length);

    let name = null;
    if (asTable) {
      const t = ws.tables.add(full, true);
      name = uniqueName(tables.items.map((x) => x.name), "SQL_Result_");
      t.name = name;
    }
    if (data.length <= 5000) full.format.autofitColumns();
    ws.activate();
    full.getCell(0, 0).select();
    await ctx.sync();
    return name;
  });

  if (tableName) {
    const linked = getSetting(SETTING_LINKED, {});
    linked[tableName] = { sql: lastResult.sql, refreshed: new Date().toISOString() };
    await setSetting(SETTING_LINKED, linked);
    renderLinked();
    setStatus(`Inserted ${lastResult.values.length.toLocaleString()} rows as table ${tableName}.`, "ok");
  } else {
    setStatus(`Inserted ${lastResult.values.length.toLocaleString()} rows.`, "ok");
  }
  await refreshObjects();
  await rebuildDeps();
}

// ---------------------------------------------------------------- refreshable query tables

async function refreshLinked(name) {
  const linked = getSetting(SETTING_LINKED, {});
  const entry = linked[name];
  if (!entry) throw new Error(`No query stored for ${name}.`);
  const { result } = await executeSql(entry.sql);
  if (!result.columns.length) throw new Error(`The query for ${name} returned no result set.`);
  const data = SqlCore.toExcelValues(result, true);
  const rows = data.length;
  const cols = data[0].length;

  await Excel.run(async (ctx) => {
    const t = ctx.workbook.tables.getItemOrNullObject(name).load("name");
    await ctx.sync();
    if (t.isNullObject) throw new Error(`Table ${name} no longer exists (renamed or deleted).`);
    const old = t.getRange().load("rowIndex,columnIndex,rowCount,columnCount");
    const ws = t.worksheet.load("name");
    await ctx.sync();

    // Cells the table will grow into must be empty.
    const growth = [];
    if (rows > old.rowCount) growth.push(ws.getRangeByIndexes(old.rowIndex + old.rowCount, old.columnIndex, rows - old.rowCount, Math.max(cols, old.columnCount)));
    if (cols > old.columnCount) growth.push(ws.getRangeByIndexes(old.rowIndex, old.columnIndex + old.columnCount, Math.min(rows, old.rowCount), cols - old.columnCount));
    if (growth.length) await assertEmpty(ctx, growth, `Table ${name} needs more room to refresh.`);

    markSelfWrite(ws.name, SqlCore.boundsOf(old.rowIndex, old.columnIndex, old.rowCount, old.columnCount));
    const target = ws.getRangeByIndexes(old.rowIndex, old.columnIndex, rows, cols);
    if (Office.context.requirements.isSetSupported("ExcelApi", "1.13")) {
      // Keep the table (and its style); swap its contents and size.
      t.getRange().clear(Excel.ClearApplyTo.contents);
      await writeValues(ctx, ws, old.rowIndex, old.columnIndex, data);
      t.resize(target);
    } else {
      t.getRange().clear(Excel.ClearApplyTo.contents);
      t.delete();
      await writeValues(ctx, ws, old.rowIndex, old.columnIndex, data);
      ws.tables.add(target, true).name = name;
    }
    await ctx.sync();
  });

  linked[name].refreshed = new Date().toISOString();
  await setSetting(SETTING_LINKED, linked);
  return result.values.length;
}

// Refreshes every query table and lookup. Order: query tables first (lookups may read them).
async function refreshEverything() {
  const jobs = [
    ...Object.keys(getSetting(SETTING_LINKED, {})).map((name) => ({ label: name, run: () => refreshLinked(name) })),
    ...getSetting(SETTING_LOOKUPS, []).map((l) => ({ label: `${l.target} ← ${l.source}`, run: () => refreshLookup(l.id) })),
  ];
  if (!jobs.length) return setStatus("Nothing to refresh yet. Insert a refreshable table or add lookup columns.");
  const failures = [];
  for (const [i, job] of jobs.entries()) {
    setStatus(`Refreshing ${job.label} (${i + 1}/${jobs.length})…`);
    try {
      await job.run();
    } catch (e) {
      failures.push(`${job.label}: ${errorText(e)}`);
    }
  }
  renderManage();
  await refreshObjects();
  await rebuildDeps();
  const done = jobs.length - failures.length;
  if (failures.length) setStatus(`Refreshed ${done} of ${jobs.length}.\n${failures.join("\n")}`, "error");
  else setStatus(`Refreshed ${done} item${done === 1 ? "" : "s"}.`, "ok");
}

function renderManage() {
  renderLinked();
  renderLookups();
  renderSaved();
}

function renderLinked() {
  const linked = getSetting(SETTING_LINKED, {});
  const names = Object.keys(linked);
  const list = $("linked");
  $("linkedCount").textContent = names.length ? `(${names.length})` : "";
  list.replaceChildren();
  if (!names.length) {
    list.append(el("li", { className: "muted small", textContent: "None yet. Run a query and insert it with “Refreshable table” checked." }));
    return;
  }
  for (const name of names) {
    const entry = linked[name];
    const when = entry.refreshed ? new Date(entry.refreshed).toLocaleString() : "";
    list.append(
      el(
        "li", null,
        el(
          "div", { className: "title" },
          el("span", { textContent: name, title: when ? "Last refreshed " + when : "" }),
          el("button", {
            textContent: "Refresh",
            onclick: () => busy(async () => {
              setStatus(`Refreshing ${name}…`);
              const n = await refreshLinked(name);
              renderLinked();
              setStatus(`Refreshed ${name}: ${n.toLocaleString()} rows.`, "ok");
            }),
          }),
          el("button", { className: "icon", textContent: "✎", title: "Open query in editor", onclick: () => loadIntoEditor(entry.sql) }),
          el("button", { className: "icon", textContent: "⌖", title: "Go to table", onclick: () => busy(() => selectTable(name)) }),
          el("button", {
            className: "icon", textContent: "✕", title: "Stop tracking (table stays)",
            onclick: () => busy(async () => {
              const all = getSetting(SETTING_LINKED, {});
              delete all[name];
              await setSetting(SETTING_LINKED, all);
              renderLinked();
              await rebuildDeps();
            }),
          })
        ),
        el("pre", { textContent: entry.sql })
      )
    );
  }
}

async function selectTable(name) {
  await Excel.run(async (ctx) => {
    const t = ctx.workbook.tables.getItemOrNullObject(name).load("name");
    await ctx.sync();
    if (t.isNullObject) throw new Error(`Table ${name} no longer exists.`);
    t.worksheet.activate();
    t.getRange().select();
    await ctx.sync();
  });
}

// ---------------------------------------------------------------- saved queries

async function saveCurrentQuery() {
  const name = $("saveName").value.trim();
  const sql = $("sql").value.trim();
  if (!name) return setStatus("Enter a name for the query.", "error");
  if (!sql) return setStatus("The editor is empty.", "error");
  const saved = getSetting(SETTING_SAVED, []).filter((q) => q.name.toLowerCase() !== name.toLowerCase());
  saved.push({ name, sql });
  saved.sort((a, b) => a.name.localeCompare(b.name));
  await setSetting(SETTING_SAVED, saved);
  $("saveRow").hidden = true;
  $("saveName").value = "";
  renderSaved();
  setStatus(`Saved “${name}” in this workbook (kept when the workbook is saved).`, "ok");
}

function renderSaved() {
  const saved = getSetting(SETTING_SAVED, []);
  const list = $("saved");
  $("savedCount").textContent = saved.length ? `(${saved.length})` : "";
  list.replaceChildren();
  if (!saved.length) {
    list.append(el("li", { className: "muted small", textContent: "None yet. Use Save above the editor." }));
    return;
  }
  for (const q of saved) {
    list.append(
      el(
        "li", null,
        el(
          "div", { className: "title" },
          el("span", { textContent: q.name }),
          el("button", { textContent: "Open", onclick: () => { loadIntoEditor(q.sql); $("saveName").value = q.name; } }),
          el("button", { className: "icon", textContent: "▶", title: "Run", onclick: () => { loadIntoEditor(q.sql); busy(runFromEditor); } }),
          el("button", {
            className: "icon", textContent: "✕", title: "Delete",
            onclick: () => busy(async () => {
              await setSetting(SETTING_SAVED, getSetting(SETTING_SAVED, []).filter((x) => x.name !== q.name));
              renderSaved();
            }),
          })
        ),
        el("pre", { textContent: q.sql })
      )
    );
  }
}

// ---------------------------------------------------------------- document settings (stored in the workbook)

function getSetting(key, fallback) {
  const v = Office.context.document.settings.get(key);
  return v === null || v === undefined ? fallback : v;
}

function setSetting(key, value) {
  Office.context.document.settings.set(key, value);
  return new Promise((resolve, reject) =>
    Office.context.document.settings.saveAsync((r) =>
      r.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(r.error)
    )
  );
}
