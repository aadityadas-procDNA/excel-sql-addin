/* global Excel, SqlCore, $, el, setStatus, errorText, insertAtCursor */
"use strict";

// Reference mode: select cells in the sheet, see which SQL reference they correspond to, insert it into the editor.
const REF_DEBOUNCE_MS = 250;
const REF_MAX_NAMES = 500;
const REF_KIND_LABEL = { table: "table", name: "named range", sheet: "whole sheet", range: "range", column: "column" };

const ref = { handler: null, timer: null, seq: 0 };

function initRefMode() {
  const box = $("refMode");
  box.onchange = () => setRefMode(box.checked).catch((e) => setStatus(errorText(e), "error"));
  let saved = false;
  try { saved = localStorage.getItem("sqlRefMode") === "1"; } catch (e) { /* storage unavailable */ }
  if (saved) {
    box.checked = true;
    box.onchange();
  }
}

async function setRefMode(on) {
  try { localStorage.setItem("sqlRefMode", on ? "1" : "0"); } catch (e) { /* ignore */ }
  $("refBar").hidden = !on;
  if (on && !ref.handler) {
    await Excel.run(async (ctx) => {
      ref.handler = ctx.workbook.onSelectionChanged.add(onRefSelection);
      await ctx.sync();
    });
    await updateRefPreview();
  } else if (!on && ref.handler) {
    clearTimeout(ref.timer);
    const h = ref.handler;
    ref.handler = null;
    await Excel.run(h.context, async (ctx) => {
      h.remove();
      await ctx.sync();
    });
  }
}

function onRefSelection() {
  clearTimeout(ref.timer);
  ref.timer = setTimeout(updateRefPreview, REF_DEBOUNCE_MS);
}

async function updateRefPreview() {
  const seq = ++ref.seq;
  let info;
  try {
    info = await readSelectionRef();
  } catch (e) {
    const msg = errorText(e);
    info = { message: /multiple|areas/i.test(msg) ? "Select a single block of cells." : msg };
  }
  if (seq !== ref.seq || !ref.handler) return; // a newer selection is being read, or the mode was turned off
  renderRefBar(info);
}

// Reads the selection and the tables, named ranges and used range around it, then picks the references.
async function readSelectionRef() {
  return Excel.run(async (ctx) => {
    const wb = ctx.workbook;
    const sel = wb.getSelectedRange().load("address,rowIndex,columnIndex,rowCount,columnCount");
    const ws = sel.worksheet.load("name");
    const used = ws.getUsedRangeOrNullObject(true).load("rowIndex,columnIndex,rowCount,columnCount");
    const sheetTables = ws.tables.load("items/name");
    const allTables = wb.tables.load("items/name");
    const names = wb.names.load("items/name,items/type,items/visible,items/formula");
    await ctx.sync();

    const boundsOf = (r) => SqlCore.boundsOf(r.rowIndex, r.columnIndex, r.rowCount, r.columnCount);
    const tableRanges = sheetTables.items.map((t) => ({ name: t.name, r: t.getRange().load("rowIndex,columnIndex,rowCount,columnCount") }));
    await ctx.sync();

    const rangeNames = names.items.filter((n) => n.type === "Range" && n.visible).slice(0, REF_MAX_NAMES);
    const info = {
      sheet: ws.name,
      sel: boundsOf(sel),
      used: used.isNullObject ? null : boundsOf(used),
      tables: tableRanges.map((t) => ({ name: t.name, bounds: boundsOf(t.r) })),
      tableNames: allTables.items.map((t) => t.name.toLowerCase()),
      names: rangeNames.map((n) => ({ name: n.name, ...SqlCore.parseNameFormula(n.formula) })).filter((n) => n.bounds),
    };

    let pick = SqlCore.pickReferenceBlock(info);
    if (pick.needsRegion) {
      info.region = await SqlCore.currentRegion(info.sel, info.used, async (list) => {
        const ranges = list.map((b) => ws.getRangeByIndexes(b.r1, b.c1, b.r2 - b.r1 + 1, b.c2 - b.c1 + 1).load("values"));
        await ctx.sync();
        return ranges.map((r) => r.values);
      });
      pick = SqlCore.pickReferenceBlock(info);
    }

    let headers = [];
    if (pick.block) {
      const b = pick.block;
      const head = ws.getRangeByIndexes(b.r1, b.c1, 1, b.c2 - b.c1 + 1).load("values");
      await ctx.sync();
      headers = head.values[0];
    }
    const taken = info.tableNames.concat(rangeNames.map((n) => n.name.toLowerCase()));
    return { address: sel.address, ...SqlCore.referenceCandidates(info, pick, headers, taken) };
  });
}

function renderRefBar(info) {
  const text = $("refInfo");
  const buttons = $("refButtons");
  buttons.replaceChildren();
  if (info.message) {
    text.textContent = info.message;
    return;
  }
  if (!info.candidates.length) {
    text.textContent = `${info.address}: no data here. Select a table, a cell inside a block of data, or a range.`;
    return;
  }

  const best = info.candidates.find((c) => c.kind !== "column");
  const shown = info.columns.slice(0, 6).join(", ") + (info.columns.length > 6 ? `, … (${info.columns.length})` : "");
  text.replaceChildren(
    el("b", { textContent: best.label }),
    ` (${REF_KIND_LABEL[best.kind]}) · ${Math.max(info.rows, 0).toLocaleString()} rows · ${shown}`
  );
  info.candidates.forEach((c, i) => {
    buttons.append(
      el("button", {
        className: i === 0 ? "primary" : "",
        textContent: c.kind === "column" ? `Insert column ${c.text}` : `Insert ${c.text}`,
        title: `Insert the ${REF_KIND_LABEL[c.kind]} reference at the editor cursor`,
        onclick: () => insertAtCursor(c.text),
      })
    );
  });
}
