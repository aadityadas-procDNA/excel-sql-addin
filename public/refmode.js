/* global Office, Excel, SqlCore, $, el, busyCount, setStatus, errorText, insertAtCursor */
"use strict";

// Reference mode: select cells in the sheet, see which SQL reference they correspond to, insert it into the editor.
const REF_DEBOUNCE_MS = 250;
const REF_MAX_NAMES = 500;
const REF_KIND_LABEL = {
  table: "table", name: "named range", sheet: "whole sheet", range: "range", column: "column", columns: "columns", value: "value",
};

const REF_POLL_MS = 1000;

// Selection events don't reach the pane in every Excel build, so a light poll of the selection address backs them up.
const ref = { on: false, timer: null, poll: null, seq: 0, lastAddress: null, checking: false };

function initRefMode() {
  const box = $("refMode");
  box.onchange = () => setRefMode(box.checked).catch((e) => setStatus(errorText(e), "error"));
  $("refNoHeader").onchange = () => updateRefPreview();
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
  if (on === ref.on) return;
  ref.on = on;
  const doc = Office.context.document;
  const type = Office.EventType.DocumentSelectionChanged;
  if (on) {
    doc.addHandlerAsync(type, onRefSelection, (r) => {
      if (r.status !== Office.AsyncResultStatus.Succeeded) console.warn("Selection event unavailable; polling only", r.error);
    });
    ref.poll = setInterval(pollSelection, REF_POLL_MS);
    await updateRefPreview();
  } else {
    clearTimeout(ref.timer);
    clearInterval(ref.poll);
    ref.lastAddress = null;
    doc.removeHandlerAsync(type, { handler: onRefSelection }, () => {});
  }
}

function onRefSelection() {
  clearTimeout(ref.timer);
  ref.timer = setTimeout(updateRefPreview, REF_DEBOUNCE_MS);
}

// Re-reads the preview when the selection moved without an event arriving.
async function pollSelection() {
  if (!ref.on || ref.checking || busyCount > 0 || document.hidden) return;
  ref.checking = true;
  try {
    const address = await Excel.run(async (ctx) => {
      const sel = ctx.workbook.getSelectedRange().load("address");
      await ctx.sync();
      return sel.address;
    }).catch(() => "(multiple areas)");
    if (ref.on && address !== ref.lastAddress) await updateRefPreview();
  } finally {
    ref.checking = false;
  }
}

async function updateRefPreview() {
  const seq = ++ref.seq;
  let info;
  try {
    info = await readSelectionRef();
  } catch (e) {
    const msg = errorText(e);
    info = { address: "(multiple areas)", message: /multiple|areas/i.test(msg) ? "Select a single block of cells." : msg };
  }
  if (seq !== ref.seq || !ref.on) return; // a newer selection is being read, or the mode was turned off
  ref.lastAddress = info.address;
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
      noHeader: $("refNoHeader").checked,
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
    let cell = null;
    if (pick.block) {
      const b = pick.block;
      const head = ws.getRangeByIndexes(b.r1, b.c1, 1, b.c2 - b.c1 + 1).load("values");
      const one = sel.rowCount === 1 && sel.columnCount === 1
        ? ws.getRangeByIndexes(sel.rowIndex, sel.columnIndex, 1, 1).load("values,numberFormat")
        : null;
      await ctx.sync();
      headers = head.values[0];
      if (one) cell = { value: one.values[0][0], format: one.numberFormat[0][0] };
    }
    const taken = info.tableNames.concat(rangeNames.map((n) => n.name.toLowerCase()));
    return { address: sel.address, ...SqlCore.referenceCandidates(info, pick, headers, taken, cell) };
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
  const selected = info.address.slice(info.address.lastIndexOf("!") + 1);
  text.replaceChildren(
    `${selected} → in `,
    el("b", { textContent: best.label }),
    ` (${REF_KIND_LABEL[best.kind]}) · ${Math.max(info.rows, 0).toLocaleString()} rows · ${shown}`,
    best.kind === "table" && $("refNoHeader").checked ? " · tables always have a header row, so “No header row” is ignored" : ""
  );
  const prefix = { column: "column ", columns: "columns ", value: "value " };
  info.candidates.forEach((c, i) => {
    buttons.append(
      el("button", {
        className: i === 0 ? "primary" : "",
        textContent: `Insert ${prefix[c.kind] || ""}${c.text}`,
        title: `Insert the ${REF_KIND_LABEL[c.kind]} reference at the editor cursor`,
        onclick: () => insertAtCursor(c.text),
      })
    );
  });
}
