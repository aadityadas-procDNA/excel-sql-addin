/* global SqlCore, $, el, busy, setStatus, fillSelect, uniqueName, columnsOf, objectCache, loadIntoEditor, runFromEditor */
"use strict";

// Visual join builder. items[0] is the starting table; each later item is joined to any earlier one.
// item: { table, type, keys: [{ left: item, leftCol, rightCol }], picked: Set(column) }
const jb = { items: [] };

const JOIN_TYPES = [
  { value: "LEFT", label: "LEFT (keep all rows so far)" },
  { value: "INNER", label: "INNER (only matching rows)" },
  { value: "FULL", label: "FULL (all rows from both)" },
];

const aliasOf = (i) => String.fromCharCode(97 + i); // a, b, c...

function initJoinBuilder() {
  $("joinBase").onchange = () => {
    const t = $("joinBase").value;
    jb.items = [];
    if (t) jb.items.push(makeItem(t, 0));
    renderJoinBuilder();
  };
  $("joinAdd").onclick = () => {
    if (!jb.items.length) return setStatus("Choose the table to start with.", "error");
    const used = new Set(jb.items.map((i) => i.table));
    const next = objectCache.find((o) => !used.has(o.name));
    jb.items.push(makeItem(next ? next.name : "", jb.items.length));
    renderJoinBuilder();
  };
  $("joinNocase").onchange = renderJoinPreview;
  $("joinTrim").onchange = renderJoinPreview;
  $("joinRun").onclick = () => busy(async () => {
    loadIntoEditor(currentJoinSql());
    await runFromEditor();
  });
  $("joinEdit").onclick = () => busy(async () => loadIntoEditor(currentJoinSql()));
}

function makeItem(table, position) {
  const item = { table, type: "LEFT", keys: [], picked: new Set() };
  setItemTable(item, table, position);
  return item;
}

// Picks default keys (shared column names with an earlier table) and default columns.
function setItemTable(item, table, position) {
  item.table = table;
  item.keys = [];
  const cols = columnsOf(table);
  if (position > 0) {
    for (let i = 0; i < position && !item.keys.length; i++) {
      const k = SqlCore.guessJoinKey(columnsOf(jb.items[i].table), cols);
      if (k) item.keys.push({ left: jb.items[i], leftCol: k.left, rightCol: k.right });
    }
    if (!item.keys.length) item.keys.push({ left: jb.items[0], leftCol: "", rightCol: "" });
  }
  const keyCols = new Set(item.keys.map((k) => k.rightCol.toLowerCase()));
  item.picked = new Set(cols.filter((c) => !keyCols.has(c.toLowerCase())));
}

function renderJoinBuilder() {
  const opts = objectCache.map((o) => ({ value: o.name, label: o.name }));
  const base = jb.items[0];
  const chosen = fillSelect($("joinBase"), [{ value: "", label: "Choose a table…" }, ...opts], base ? base.table : "");
  if (!chosen) jb.items = [];

  const list = $("joinList");
  list.replaceChildren();
  jb.items.slice(1).forEach((item, k) => {
    const pos = k + 1;
    const type = el("select", { title: "Join type" });
    item.type = fillSelect(type, JOIN_TYPES, item.type);
    type.onchange = () => { item.type = type.value; renderJoinPreview(); };

    const tbl = el("select", { title: "Table to join" });
    fillSelect(tbl, [{ value: "", label: "table…" }, ...opts], item.table);
    tbl.onchange = () => { setItemTable(item, tbl.value, pos); renderJoinBuilder(); };

    const remove = el("button", {
      className: "icon", textContent: "✕", title: "Remove this join",
      onclick: () => {
        jb.items.splice(pos, 1);
        // Keys pointing at the removed table fall back to the first table.
        for (const it of jb.items) for (const key of it.keys) if (!jb.items.includes(key.left)) Object.assign(key, { left: jb.items[0], leftCol: "" });
        renderJoinBuilder();
      },
    });

    // Left side of each key: any column of any earlier table.
    const leftOptions = [{ value: "", label: "column…" }];
    jb.items.slice(0, pos).forEach((it, i) => columnsOf(it.table).forEach((c) => leftOptions.push({ value: `${i}|${c}`, label: `${aliasOf(i)}.${c}` })));
    const rightOptions = [{ value: "", label: "column…" }, ...columnsOf(item.table).map((c) => ({ value: c, label: `${aliasOf(pos)}.${c}` }))];

    const on = el("div", { className: "on" });
    item.keys.forEach((key, ki) => {
      const ls = el("select");
      const lv = fillSelect(ls, leftOptions, key.leftCol ? `${jb.items.indexOf(key.left)}|${key.leftCol}` : "");
      if (!lv) key.leftCol = "";
      ls.onchange = () => {
        const [i, ...rest] = ls.value.split("|");
        Object.assign(key, ls.value ? { left: jb.items[Number(i)], leftCol: rest.join("|") } : { leftCol: "" });
        renderJoinPreview();
      };
      const rs = el("select");
      key.rightCol = fillSelect(rs, rightOptions, key.rightCol);
      rs.onchange = () => { key.rightCol = rs.value; renderJoinPreview(); };
      const rmKey = el("button", {
        className: "icon", textContent: "✕", title: "Remove key", hidden: item.keys.length === 1,
        onclick: () => { item.keys.splice(ki, 1); renderJoinBuilder(); },
      });
      on.append(el("div", { className: "keyrow" }, el("span", { className: "muted small", textContent: ki ? "and" : "on" }), ls, el("span", { className: "eq", textContent: "=" }), rs, rmKey));
    });
    const addKey = el("a", {
      href: "#", className: "small", textContent: "+ key column",
      onclick: (e) => { e.preventDefault(); item.keys.push({ left: jb.items[0], leftCol: "", rightCol: "" }); renderJoinBuilder(); },
    });

    list.append(el("div", { className: "join" }, el("div", { className: "head" }, type, tbl, el("span", { className: "alias", textContent: aliasOf(pos) }), remove), on, addKey));
  });

  renderJoinColumns();
  renderJoinPreview();
}

function renderJoinColumns() {
  const box = $("joinColumns");
  box.replaceChildren();
  if (!jb.items.length) {
    box.append(el("div", { className: "empty", textContent: "Choose a table to start with." }));
    return;
  }
  jb.items.forEach((item, i) => {
    if (!item.table) return;
    const cols = columnsOf(item.table);
    const setAll = (on) => (e) => {
      e.preventDefault();
      item.picked = new Set(on ? cols : []);
      renderJoinColumns();
      renderJoinPreview();
    };
    box.append(
      el("div", { className: "group" },
        el("span", { textContent: `${aliasOf(i)} · ${item.table}` }),
        el("a", { href: "#", textContent: "all", onclick: setAll(true) }),
        el("a", { href: "#", textContent: "none", onclick: setAll(false) }))
    );
    for (const c of cols) {
      const cb = el("input", { type: "checkbox", checked: item.picked.has(c) });
      cb.onchange = () => { cb.checked ? item.picked.add(c) : item.picked.delete(c); renderJoinPreview(); };
      box.append(el("label", null, cb, c));
    }
  });
}

function joinSpec() {
  if (!jb.items.length) throw new Error("Choose a table to start with.");
  const used = [];
  const columns = [];
  jb.items.forEach((item, i) => {
    for (const c of columnsOf(item.table)) {
      if (!item.picked.has(c)) continue;
      const clash = used.some((u) => u.toLowerCase() === c.toLowerCase());
      const as = clash ? uniqueName(used, `${c} (${item.table})`, false) : c;
      used.push(as);
      columns.push({ alias: aliasOf(i), column: c, as });
    }
  });
  const joins = jb.items.slice(1).map((item, k) => {
    if (!item.table) throw new Error(`Join ${k + 1}: choose a table.`);
    const on = item.keys
      .filter((key) => key.leftCol && key.rightCol && jb.items.includes(key.left))
      .map((key) => ({ leftAlias: aliasOf(jb.items.indexOf(key.left)), leftColumn: key.leftCol, rightColumn: key.rightCol }));
    if (!on.length) throw new Error(`Join ${k + 1} (${item.table}): pick the key columns to match on.`);
    return { type: item.type, name: item.table, alias: aliasOf(k + 1), on };
  });
  return {
    base: { name: jb.items[0].table, alias: "a" },
    joins,
    columns,
    nocase: $("joinNocase").checked,
    trim: $("joinTrim").checked,
  };
}

function currentJoinSql() {
  return SqlCore.buildJoinSql(joinSpec());
}

function renderJoinPreview() {
  const pre = $("joinPreview");
  try {
    pre.textContent = currentJoinSql();
  } catch (e) {
    pre.textContent = e.message;
  }
}
