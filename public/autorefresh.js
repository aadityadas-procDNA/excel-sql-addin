/* global Office, Excel, SqlCore, $, busy, busyCount, setStatus, errorText, listObjects, findObject, columnsOf,
   refreshLinked, refreshLookup, renderManage, getSetting, setSetting, SETTING_LINKED, SETTING_LOOKUPS */
"use strict";

// Auto-refresh: watches cell changes and refreshes query tables / lookups whose sources changed.
// Runs in the shared runtime, so it keeps working with the task pane closed.
const SETTING_AUTO = "sqlAutoRefresh";
const SELF_WRITE_MS = 4000; // our own writes are ignored for this long
const DEBOUNCE_MS = 1200;
const MAX_CASCADE = 5;

const auto = {
  on: false,
  handler: null,
  deps: [], // [{ key: "q:<table>" | "l:<id>", sheet, bounds, table? }]
  sheetNames: new Map(), // worksheet id -> name
  selfWrites: [], // [{ sheet, bounds, until }]
  pending: new Set(),
  timer: null,
};

function markSelfWrite(sheet, bounds) {
  auto.selfWrites.push({ sheet, bounds, until: Date.now() + SELF_WRITE_MS });
}

async function initAutoRefresh() {
  const box = $("autoRefresh");
  if (!Office.context.requirements.isSetSupported("ExcelApi", "1.9")) {
    box.disabled = true;
    $("autoNote").textContent = "Auto-refresh needs a newer version of Excel (ExcelApi 1.9).";
    return;
  }
  box.checked = getSetting(SETTING_AUTO, false);
  box.onchange = () => busy(() => setAutoRefresh(box.checked));
  if (box.checked) await startAuto();
}

async function setAutoRefresh(on) {
  await setSetting(SETTING_AUTO, on);
  if (on) await startAuto();
  else await stopAuto();
  // Load the add-in (hidden) whenever this workbook opens, so auto-refresh survives restarts.
  if (Office.addin && Office.addin.setStartupBehavior) {
    try {
      await Office.addin.setStartupBehavior(on ? Office.StartupBehavior.load : Office.StartupBehavior.none);
    } catch (e) {
      console.warn("setStartupBehavior failed", e);
    }
  }
  setStatus(
    on ? `Auto-refresh on: watching ${new Set(auto.deps.map((d) => d.key)).size} item(s). Save the workbook to keep this setting.` : "Auto-refresh off.",
    "ok"
  );
}

async function startAuto() {
  auto.on = true;
  if (!auto.handler) {
    await Excel.run(async (ctx) => {
      auto.handler = ctx.workbook.worksheets.onChanged.add(onSheetChanged);
      await ctx.sync();
    });
  }
  await rebuildDeps();
}

async function stopAuto() {
  auto.on = false;
  auto.deps = [];
  auto.pending.clear();
  clearTimeout(auto.timer);
  if (auto.handler) {
    const h = auto.handler;
    auto.handler = null;
    await Excel.run(h.context, async (ctx) => {
      h.remove();
      await ctx.sync();
    });
  }
}

// Recomputes which cells each refreshable item depends on.
async function rebuildDeps() {
  if (!auto.on) return;
  const linked = getSetting(SETTING_LINKED, {});
  const lookups = getSetting(SETTING_LOOKUPS, []);
  try {
    auto.deps = await Excel.run(async (ctx) => {
      const wb = ctx.workbook;
      const objs = await listObjects(ctx);
      const sheets = wb.worksheets.load("items/id,items/name");

      const wanted = [];
      for (const [name, e] of Object.entries(linked)) {
        for (const src of SqlCore.resolveSources(SqlCore.extractIdentifiers(e.sql), objs)) wanted.push({ key: "q:" + name, src });
      }
      for (const l of lookups) {
        const src = findObject(l.source, objs);
        if (src) wanted.push({ key: "l:" + l.id, src });
        for (const k of l.keys) wanted.push({ key: "l:" + l.id, src: { kind: "column", table: l.target, column: k.target } });
      }

      const hasSheet = (s) => objs.some((o) => o.kind === "sheet" && o.name.toLowerCase() === s.toLowerCase());
      const hasTable = (t) => objs.some((o) => o.kind === "table" && o.name.toLowerCase() === t.toLowerCase());
      const items = [];
      for (const w of wanted) {
        const s = w.src;
        if (s.kind === "sheet") {
          items.push({ key: w.key, sheet: s.ref, bounds: SqlCore.boundsOf(0, 0, 1048576, 16384) });
          continue;
        }
        let r;
        let padRows = 0;
        let padCols = 0;
        if (s.kind === "table") {
          r = wb.tables.getItem(s.ref).getRange();
          padRows = 1; // typing just below a table extends it
          padCols = 1;
        } else if (s.kind === "name") {
          r = wb.names.getItem(s.ref).getRange();
        } else if (s.kind === "range") {
          const { sheet, address } = SqlCore.splitAddress(s.ref);
          if (!hasSheet(sheet)) continue;
          r = wb.worksheets.getItem(sheet).getRange(address);
        } else if (s.kind === "column") {
          if (!hasTable(s.table) || !columnsOf(s.table).some((c) => c.toLowerCase() === s.column.toLowerCase())) continue;
          r = wb.tables.getItem(s.table).columns.getItem(s.column).getRange();
          padRows = 1;
        }
        r.load("rowIndex,columnIndex,rowCount,columnCount");
        items.push({ key: w.key, r, ws: r.worksheet.load("name"), padRows, padCols, table: s.kind === "table" ? s.ref : null });
      }
      await ctx.sync();
      auto.sheetNames = new Map(sheets.items.map((s) => [s.id, s.name]));
      return items.map((it) =>
        it.r
          ? {
              key: it.key,
              sheet: it.ws.name,
              table: it.table,
              bounds: SqlCore.boundsOf(it.r.rowIndex, it.r.columnIndex, it.r.rowCount + it.padRows, it.r.columnCount + it.padCols),
            }
          : it
      );
    });
  } catch (e) {
    console.warn("Auto-refresh: could not map sources", e);
    setStatus("Auto-refresh: could not map some sources: " + errorText(e), "error");
  }
}

async function onSheetChanged(event) {
  if (!auto.on) return;
  let sheet = auto.sheetNames.get(event.worksheetId);
  if (!sheet) {
    await rebuildDeps(); // new or renamed sheet
    sheet = auto.sheetNames.get(event.worksheetId);
    if (!sheet) return;
  }
  let areas;
  try {
    areas = SqlCore.parseAddress(event.address);
  } catch (e) {
    return;
  }
  const now = Date.now();
  auto.selfWrites = auto.selfWrites.filter((w) => w.until > now);
  const ours = areas.every((a) => auto.selfWrites.some((w) => w.sheet === sheet && SqlCore.contains(w.bounds, a)));
  if (ours) return;

  for (const d of auto.deps) {
    if (d.sheet === sheet && areas.some((a) => SqlCore.intersects(d.bounds, a))) auto.pending.add(d.key);
  }
  if (auto.pending.size) schedule(DEBOUNCE_MS);
}

function schedule(ms) {
  clearTimeout(auto.timer);
  auto.timer = setTimeout(flush, ms);
}

async function flush() {
  if (!auto.on || !auto.pending.size) return;
  if (busyCount > 0) return schedule(800); // let a running user action finish first

  await busy(async () => {
    const done = [];
    const failed = [];
    const seen = new Set();
    for (let round = 0; auto.pending.size && round < MAX_CASCADE; round++) {
      // Query tables first: lookups may read them.
      const keys = [...auto.pending].sort((a, b) => (a[0] === b[0] ? 0 : a[0] === "q" ? -1 : 1));
      auto.pending.clear();
      for (const key of keys) {
        const id = key.slice(2);
        let label = id;
        try {
          if (key[0] === "q") {
            await refreshLinked(id);
            cascadeFrom(id, key);
          } else {
            const { def } = await refreshLookup(id);
            label = `${def.target} ← ${def.source}`;
            cascadeFrom(def.target, key);
          }
          if (!seen.has(key)) done.push(label);
          seen.add(key);
        } catch (e) {
          failed.push(`${label}: ${errorText(e)}`);
        }
      }
    }
    renderManage();
    await rebuildDeps();
    const time = new Date().toLocaleTimeString();
    if (failed.length) setStatus(`Auto-refresh (${time}) failed:\n${failed.join("\n")}`, "error");
    else if (done.length) setStatus(`Auto-refreshed ${done.join(", ")} at ${time}.`, "ok");
  });
}

// A refreshed table may feed other items; queue those too (our own writes don't raise events we act on).
function cascadeFrom(tableName, sourceKey) {
  for (const d of auto.deps) {
    if (d.key !== sourceKey && d.table && d.table.toLowerCase() === tableName.toLowerCase()) auto.pending.add(d.key);
  }
}
