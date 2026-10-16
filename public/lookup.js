/* global Excel, SqlCore, $, el, busy, setStatus, fillSelect, uniqueName, columnsOf, objectCache, listObjects, findObject,
   readSource, readBlock, writeValues, assertEmpty, markSelfWrite, refreshObjects, rebuildDeps, getSetting, setSetting, selectTable */
"use strict";

// Lookup columns: add columns to an Excel table by matching keys in another table/range (XLOOKUP replacement).
// Definitions are stored in the workbook so they can be refreshed: [{ id, target, source, keys, columns, mode, dup, notFound }]
const SETTING_LOOKUPS = "sqlLookups";

const lk = { target: "", source: "", keys: [{ target: "", source: "" }], picked: new Set() };

function initLookup() {
  $("lkTarget").onchange = () => { lk.target = $("lkTarget").value; guessLookupKeys(); renderLookupForm(); };
  $("lkSource").onchange = () => { lk.source = $("lkSource").value; lk.picked.clear(); guessLookupKeys(); renderLookupForm(); };
  $("lkAddKey").onclick = () => { lk.keys.push({ target: "", source: "" }); renderLookupForm(); };
  $("lkAll").onclick = (e) => { e.preventDefault(); lookupValueColumns().forEach((c) => lk.picked.add(c)); renderLookupForm(); };
  $("lkNone").onclick = (e) => { e.preventDefault(); lk.picked.clear(); renderLookupForm(); };
  $("lkApply").onclick = () => busy(applyLookupFromForm);
}

function guessLookupKeys() {
  if (!lk.target || !lk.source) return;
  const k = SqlCore.guessJoinKey(columnsOf(lk.target), columnsOf(lk.source));
  lk.keys = [k ? { target: k.left, source: k.right } : { target: "", source: "" }];
}

// Source columns that can be brought over (everything except the key columns).
function lookupValueColumns() {
  const keys = new Set(lk.keys.map((k) => k.source.toLowerCase()));
  return columnsOf(lk.source).filter((c) => !keys.has(c.toLowerCase()));
}

function renderLookupForm() {
  const tables = objectCache.filter((o) => o.kind === "table");
  lk.target = fillSelect($("lkTarget"), [{ value: "", label: "Choose a table…" }, ...tables.map((t) => ({ value: t.name, label: t.name }))], lk.target);
  const kindNote = { table: "", name: " (named range)", sheet: " (sheet)" };
  const sources = objectCache.filter((o) => o.name !== lk.target);
  lk.source = fillSelect($("lkSource"), [{ value: "", label: "Choose…" }, ...sources.map((o) => ({ value: o.name, label: o.name + kindNote[o.kind] }))], lk.source);

  const colOptions = (cols) => [{ value: "", label: "column…" }, ...cols.map((c) => ({ value: c, label: c }))];
  const keyBox = $("lkKeys");
  keyBox.replaceChildren();
  lk.keys.forEach((k, i) => {
    const ts = el("select", { title: `Column in ${lk.target || "the table"}` });
    k.target = fillSelect(ts, colOptions(columnsOf(lk.target)), k.target);
    ts.onchange = () => { k.target = ts.value; };
    const ss = el("select", { title: `Column in ${lk.source || "the lookup source"}` });
    k.source = fillSelect(ss, colOptions(columnsOf(lk.source)), k.source);
    ss.onchange = () => { k.source = ss.value; renderLookupForm(); };
    const rm = el("button", {
      className: "icon", textContent: "✕", title: "Remove key", hidden: lk.keys.length === 1,
      onclick: () => { lk.keys.splice(i, 1); renderLookupForm(); },
    });
    keyBox.append(el("div", { className: "keyrow" }, ts, el("span", { className: "eq", textContent: "=" }), ss, rm));
  });

  const box = $("lkColumns");
  box.replaceChildren();
  const cols = lookupValueColumns();
  if (!cols.length) box.append(el("div", { className: "empty", textContent: lk.source ? "No other columns." : "Choose where to look up." }));
  for (const c of cols) {
    const cb = el("input", { type: "checkbox", checked: lk.picked.has(c) });
    cb.onchange = () => (cb.checked ? lk.picked.add(c) : lk.picked.delete(c));
    box.append(el("label", null, cb, c));
  }
}

async function applyLookupFromForm() {
  if (!lk.target) throw new Error("Choose the table to add columns to.");
  if (!lk.source) throw new Error("Choose where to look up values.");
  if (lk.keys.some((k) => !k.target || !k.source)) throw new Error("Pick a column on both sides of every key.");
  const picked = lookupValueColumns().filter((c) => lk.picked.has(c));
  if (!picked.length) throw new Error("Tick at least one column to bring over.");

  const taken = [...columnsOf(lk.target)];
  const columns = picked.map((c) => {
    const clash = taken.some((t) => t.toLowerCase() === c.toLowerCase());
    const dest = clash ? uniqueName(taken, `${c} (${lk.source})`, false) : c;
    taken.push(dest);
    return { source: c, dest };
  });
  const def = {
    id: "lk" + Date.now().toString(36),
    target: lk.target,
    source: lk.source,
    keys: lk.keys.map((k) => ({ ...k })),
    columns,
    mode: $("lkMode").value,
    dup: $("lkDup").value,
    notFound: $("lkNotFound").value,
  };

  setStatus("Looking up…");
  const stats = await runLookup(def);
  if ($("lkKeep").checked) {
    const all = getSetting(SETTING_LOOKUPS, []);
    all.push({ ...def, refreshed: new Date().toISOString() });
    await setSetting(SETTING_LOOKUPS, all);
    renderLookups();
  }
  lk.picked.clear();
  await refreshObjects();
  await rebuildDeps();
  setStatus(lookupSummary("Added", def, stats), stats.missing ? "" : "ok");
}

function lookupSummary(verb, def, s) {
  const n = def.columns.length;
  let msg = `${verb} ${n} column${n === 1 ? "" : "s"} in ${def.target}: ${s.matched.toLocaleString()} of ${s.rows.toLocaleString()} rows matched`;
  if (s.missing) msg += `, ${s.missing.toLocaleString()} not found`;
  msg += ".";
  if (s.duplicates) msg += ` ${def.source} has ${s.duplicates.toLocaleString()} duplicate key${s.duplicates === 1 ? "" : "s"}; used the ${def.dup} match.`;
  if (s.missing && def.mode !== "loose") msg += "\nTip: “Loose” matching also ignores extra spaces and numbers stored as text.";
  return msg;
}

// Computes the lookup and writes the values into the target table (adding columns as needed).
async function runLookup(def) {
  return Excel.run(async (ctx) => {
    const objs = await listObjects(ctx);
    if (!objs.some((o) => o.kind === "table" && o.name.toLowerCase() === def.target.toLowerCase())) {
      throw new Error(`Table ${def.target} no longer exists.`);
    }
    const src = findObject(def.source, objs);
    if (!src) throw new Error(`${def.source} no longer exists.`);

    const sdata = await readSource(ctx, src);
    const sHeaders = SqlCore.normalizeHeaders(sdata.headers);
    const sIdx = (name) => {
      const i = sHeaders.findIndex((h) => h.toLowerCase() === name.toLowerCase());
      if (i < 0) throw new Error(`Column ${name} not found in ${def.source}.`);
      return i;
    };
    const keyIdx = def.keys.map((k) => sIdx(k.source));
    const valIdx = def.columns.map((c) => sIdx(c.source));
    const formats = valIdx.map((c) => {
      const f = SqlCore.sampleFormat(sdata.rows, sdata.formats, c);
      return f && f !== "General" ? f : null;
    });

    const t = ctx.workbook.tables.getItem(def.target);
    const ws = t.worksheet.load("name");
    const header = t.getHeaderRowRange().load("values,rowIndex,columnIndex,columnCount");
    const body = t.getDataBodyRange().load("rowIndex,columnIndex,rowCount");
    await ctx.sync();
    let tHeaders = header.values[0].map(String);
    const tIdx = (name) => {
      const i = tHeaders.findIndex((h) => h.toLowerCase() === name.toLowerCase());
      if (i < 0) throw new Error(`Column ${name} not found in ${def.target}.`);
      return i;
    };
    const keyCols = def.keys.map((k) => tIdx(k.target));

    // Read only the target's key columns.
    const keyData = [];
    for (const c of keyCols) keyData.push((await readBlock(ctx, ws, body.rowIndex, body.columnIndex + c, body.rowCount, 1, 0)).values);
    const targetKeys = keyData[0].map((_, r) => keyData.map((col) => col[r][0]));

    const { index, duplicates } = SqlCore.buildLookupIndex(sdata.rows, keyIdx, def.mode, def.dup);
    const res = SqlCore.applyLookup(targetKeys, index, valIdx, def.mode, def.notFound);

    // Add destination columns that don't exist yet (first run, or a user deleted one).
    const missing = def.columns.filter((c) => !tHeaders.some((h) => h.toLowerCase() === c.dest.toLowerCase()));
    if (missing.length) {
      const right = SqlCore.boundsOf(header.rowIndex, header.columnIndex + header.columnCount, body.rowCount + 1, missing.length);
      await assertEmpty(
        ctx,
        [ws.getRangeByIndexes(right.r1, right.c1, body.rowCount + 1, missing.length)],
        `New columns go to the right of ${def.target}; clear those cells or move the table.`
      );
      markSelfWrite(ws.name, right);
      for (const c of missing) t.columns.add(null, null, c.dest);
      const fresh = t.getHeaderRowRange().load("values");
      await ctx.sync();
      tHeaders = fresh.values[0].map(String);
    }

    for (const [j, c] of def.columns.entries()) {
      const ci = tIdx(c.dest);
      if (keyCols.includes(ci)) throw new Error(`Refusing to overwrite key column ${c.dest}.`);
      await writeValues(ctx, ws, body.rowIndex, body.columnIndex + ci, res.columns[j].map((v) => [v]), [formats[j]]);
    }
    return { rows: targetKeys.length, matched: res.matched, missing: res.missing, duplicates };
  });
}

async function refreshLookup(id) {
  const all = getSetting(SETTING_LOOKUPS, []);
  const def = all.find((l) => l.id === id);
  if (!def) throw new Error("Lookup definition not found.");
  const stats = await runLookup(def);
  def.refreshed = new Date().toISOString();
  await setSetting(SETTING_LOOKUPS, all);
  return { def, stats };
}

const MODE_LABELS = { nocase: "ignore case", loose: "loose", exact: "exact" };

function renderLookups() {
  const all = getSetting(SETTING_LOOKUPS, []);
  const list = $("lookups");
  $("lookupCount").textContent = all.length ? `(${all.length})` : "";
  list.replaceChildren();
  if (!all.length) {
    list.append(el("li", { className: "muted small", textContent: "None yet. Use the Lookup tab with “Refreshable” checked." }));
    return;
  }
  for (const def of all) {
    const keys = def.keys.map((k) => (k.target === k.source ? k.target : `${k.target} = ${k.source}`)).join(" + ");
    const when = def.refreshed ? "Last refreshed " + new Date(def.refreshed).toLocaleString() : "";
    list.append(
      el(
        "li", null,
        el(
          "div", { className: "title" },
          el("span", { textContent: `${def.target} ← ${def.source}`, title: when }),
          el("button", {
            textContent: "Refresh",
            onclick: () => busy(async () => {
              setStatus(`Refreshing lookup into ${def.target}…`);
              const { stats } = await refreshLookup(def.id);
              renderLookups();
              setStatus(lookupSummary("Refreshed", def, stats), stats.missing ? "" : "ok");
            }),
          }),
          el("button", { className: "icon", textContent: "⌖", title: "Go to table", onclick: () => busy(() => selectTable(def.target)) }),
          el("button", {
            className: "icon", textContent: "✕", title: "Stop tracking (columns stay)",
            onclick: () => busy(async () => {
              await setSetting(SETTING_LOOKUPS, getSetting(SETTING_LOOKUPS, []).filter((l) => l.id !== def.id));
              renderLookups();
              await rebuildDeps();
            }),
          })
        ),
        el("pre", { textContent: `${def.columns.map((c) => c.dest).join(", ")}\nkey: ${keys} · ${MODE_LABELS[def.mode] || def.mode}` })
      )
    );
  }
}
