// Minimal in-memory stand-in for the Office.js Excel API, covering what the task pane uses.
// Reads are live (no batching); writes apply immediately and raise worksheet onChanged events like Excel does.
const SqlCore = require("../public/sqlcore.js");

const colName = (c) => {
  let s = "";
  for (c++; c > 0; c = Math.floor((c - 1) / 26)) s = String.fromCharCode(65 + ((c - 1) % 26)) + s;
  return s;
};
const isEmpty = (v) => v === "" || v === null || v === undefined;

class NullObject {
  constructor() { this.isNullObject = true; }
  load() { return this; }
}

class Range {
  constructor(ws, r, c, rows, cols) { Object.assign(this, { ws, r, c, rows, cols, isNullObject: false }); }
  load() { return this; }
  get rowIndex() { return this.r; }
  get columnIndex() { return this.c; }
  get rowCount() { return this.rows; }
  get columnCount() { return this.cols; }
  get address() {
    const a = colName(this.c) + (this.r + 1);
    const b = colName(this.c + this.cols - 1) + (this.r + this.rows);
    return `${this.ws.name}!${a === b ? a : a + ":" + b}`;
  }
  get localAddress() { return this.address.split("!")[1]; }
  get worksheet() { return this.ws; }
  grid(fn) { return Array.from({ length: this.rows }, (_, i) => Array.from({ length: this.cols }, (_, j) => fn(this.r + i, this.c + j))); }
  get values() { return this.grid((r, c) => this.ws.get(r, c)); }
  set values(v) {
    const scalar = !Array.isArray(v);
    this.grid((r, c) => this.ws.set(r, c, scalar ? v : v[r - this.r][c - this.c]));
    this.ws.changed(this);
  }
  get numberFormat() { return this.grid((r, c) => this.ws.formats.get(r + "," + c) || "General"); }
  set numberFormat(v) { this.grid((r, c) => this.ws.formats.set(r + "," + c, Array.isArray(v) ? v[r - this.r][c - this.c] : v)); }
  get format() { return { autofitColumns() {} }; }
  getRow(i) { return new Range(this.ws, this.r + i, this.c, 1, this.cols); }
  getCell(i, j) { return new Range(this.ws, this.r + i, this.c + j, 1, 1); }
  select() {}
  clear() {
    this.grid((r, c) => this.ws.set(r, c, ""));
    this.ws.changed(this);
  }
  bounds() { return SqlCore.boundsOf(this.r, this.c, this.rows, this.cols); }
  getIntersectionOrNullObject(other) {
    const a = this.bounds();
    const b = other.bounds();
    if (!SqlCore.intersects(a, b)) return new NullObject();
    const r1 = Math.max(a.r1, b.r1), c1 = Math.max(a.c1, b.c1), r2 = Math.min(a.r2, b.r2), c2 = Math.min(a.c2, b.c2);
    return new Range(this.ws, r1, c1, r2 - r1 + 1, c2 - c1 + 1);
  }
  getUsedRangeOrNullObject() {
    let b = null;
    this.grid((r, c) => {
      if (isEmpty(this.ws.get(r, c))) return;
      b = b ? { r1: Math.min(b.r1, r), c1: Math.min(b.c1, c), r2: Math.max(b.r2, r), c2: Math.max(b.c2, c) } : { r1: r, c1: c, r2: r, c2: c };
    });
    return b ? new Range(this.ws, b.r1, b.c1, b.r2 - b.r1 + 1, b.c2 - b.c1 + 1) : new NullObject();
  }
}

class Table {
  constructor(wb, ws, name, r, c, rows, cols) { Object.assign(this, { wb, ws, name, r, c, rows, cols }); }
  load() { return this; }
  get worksheet() { return this.ws; }
  getRange() { return new Range(this.ws, this.r, this.c, this.rows, this.cols); }
  getHeaderRowRange() { return new Range(this.ws, this.r, this.c, 1, this.cols); }
  getDataBodyRange() { return new Range(this.ws, this.r + 1, this.c, Math.max(1, this.rows - 1), this.cols); }
  get columns() {
    const t = this;
    const item = (name) => {
      const i = t.getHeaderRowRange().values[0].findIndex((h) => String(h).toLowerCase() === name.toLowerCase());
      return i < 0 ? null : { name, load() { return this; }, getRange: () => new Range(t.ws, t.r, t.c + i, t.rows, 1) };
    };
    return {
      load() { return this; },
      add(index, values, name) {
        t.cols++;
        t.ws.set(t.r, t.c + t.cols - 1, name || "Column" + t.cols);
        t.ws.changed(new Range(t.ws, t.r, t.c + t.cols - 1, t.rows, 1));
      },
      getItem(name) {
        const it = item(name);
        if (!it) throw new Error("ItemNotFound: column " + name);
        return it;
      },
    };
  }
  resize(range) { Object.assign(this, { r: range.r, c: range.c, rows: range.rows, cols: range.cols }); }
  delete() { this.wb.tableList.splice(this.wb.tableList.indexOf(this), 1); }
}

class Worksheet {
  constructor(wb, name, id) {
    Object.assign(this, { wb, name, id, cells: new Map(), formats: new Map() });
  }
  load() { return this; }
  get(r, c) { return this.cells.has(r + "," + c) ? this.cells.get(r + "," + c) : ""; }
  set(r, c, v) {
    if (typeof v === "string" && v.startsWith("'")) v = v.slice(1);
    if (isEmpty(v)) this.cells.delete(r + "," + c);
    else this.cells.set(r + "," + c, v);
  }
  changed(range) { this.wb.emit(this, range.localAddress); }
  usedBounds() {
    let b = null;
    for (const key of this.cells.keys()) {
      const [r, c] = key.split(",").map(Number);
      b = b ? { r1: Math.min(b.r1, r), c1: Math.min(b.c1, c), r2: Math.max(b.r2, r), c2: Math.max(b.c2, c) } : { r1: r, c1: c, r2: r, c2: c };
    }
    return b;
  }
  getRangeByIndexes(r, c, rows, cols) { return new Range(this, r, c, rows, cols); }
  getRange(address) {
    const [b] = SqlCore.parseAddress(address);
    return new Range(this, b.r1, b.c1, b.r2 - b.r1 + 1, b.c2 - b.c1 + 1);
  }
  getUsedRange() {
    const b = this.usedBounds();
    return b ? new Range(this, b.r1, b.c1, b.r2 - b.r1 + 1, b.c2 - b.c1 + 1) : new Range(this, 0, 0, 1, 1);
  }
  getUsedRangeOrNullObject() { return this.usedBounds() ? this.getUsedRange() : new NullObject(); }
  activate() {}
  get tables() {
    return {
      items: this.wb.tableList.filter((t) => t.ws === this),
      load() { return this; },
      add: (range, hasHeaders) => {
        const t = new Table(this.wb, this, "Table" + (this.wb.tableList.length + 1), range.r, range.c, range.rows, range.cols);
        this.wb.tableList.push(t);
        return t;
      },
    };
  }

  // Test helpers
  write(address, rows) {
    const r = this.getRange(address);
    new Range(this, r.r, r.c, rows.length, rows[0].length).values = rows;
  }
}

class Workbook {
  constructor() {
    this.sheetList = [];
    this.tableList = [];
    this.nameList = [];
    this.handlers = [];
    this.selectionHandlers = [];
    this.events = [];
    this.active = null;
    this.selection = null;
  }
  // Test helper: the user selects cells; raises onSelectionChanged like Excel.
  select(ws, address) {
    this.active = ws;
    this.selection = address === null ? null : ws.getRange(address);
    for (const h of this.selectionHandlers) setTimeout(() => h({}), 0);
  }
  getSelectedRange() {
    if (!this.selection) throw new Error("InvalidSelection: The operation is not supported on multiple areas.");
    return this.selection;
  }
  get onSelectionChanged() {
    const wb = this;
    return {
      add(handler) {
        wb.selectionHandlers.push(handler);
        return { context: {}, remove: () => wb.selectionHandlers.splice(wb.selectionHandlers.indexOf(handler), 1) };
      },
    };
  }
  addSheet(name) {
    const ws = new Worksheet(this, name, "{sheet-" + (this.sheetList.length + 1) + "}");
    this.sheetList.push(ws);
    this.active = this.active || ws;
    return ws;
  }
  addTable(ws, name, address) {
    const r = ws.getRange(address);
    const t = new Table(this, ws, name, r.r, r.c, r.rows, r.cols);
    this.tableList.push(t);
    return t;
  }
  addName(name, ws, address) {
    const formula = `='${ws.name}'!` + address.replace(/([A-Z]+)(\d+)/g, "$$$1$$$2");
    this.nameList.push({ name, type: "Range", visible: true, formula, getRange: () => ws.getRange(address) });
  }
  emit(ws, address) {
    const event = { worksheetId: ws.id, address, changeType: "RangeEdited", source: "Local" };
    this.events.push(event);
    for (const h of this.handlers) setTimeout(() => h(event), 0);
  }
  collection(list, what) {
    const find = (name) => list.find((x) => x.name.toLowerCase() === String(name).toLowerCase());
    return {
      items: list,
      load() { return this; },
      getItem(name) {
        const x = find(name);
        if (!x) throw new Error(`ItemNotFound: ${what} ${name}`);
        return x;
      },
      getItemOrNullObject(name) { return find(name) || new NullObject(); },
    };
  }
  get tables() { return this.collection(this.tableList, "table"); }
  get names() { return this.collection(this.nameList, "name"); }
  get worksheets() {
    const wb = this;
    return Object.assign(this.collection(this.sheetList, "sheet"), {
      add: (name) => wb.addSheet(name),
      onChanged: {
        add(handler) {
          wb.handlers.push(handler);
          return { context: {}, remove: () => wb.handlers.splice(wb.handlers.indexOf(handler), 1) };
        },
      },
    });
  }
  getActiveCell() { return new Range(this.active, 0, 0, 1, 1); }
  tableValues(name) { return this.tables.getItem(name).getRange().values; }
}

function makeOffice(wb) {
  const settings = new Map();
  let ready = null;
  const ctx = { workbook: wb, sync: async () => {} };
  const Office = {
    HostType: { Excel: "Excel" },
    AsyncResultStatus: { Succeeded: "succeeded" },
    EventType: { DocumentSelectionChanged: "documentSelectionChanged" },
    StartupBehavior: { load: "Load", none: "None" },
    startup: null,
    onReady(cb) { ready = cb; },
    boot: () => ready({ host: "Excel" }),
    addin: { async setStartupBehavior(b) { Office.startup = b; } },
    context: {
      requirements: { isSetSupported: () => true },
      document: {
        addHandlerAsync(type, handler, cb) {
          if (type === "documentSelectionChanged") wb.selectionHandlers.push(handler);
          if (cb) cb({ status: "succeeded" });
        },
        removeHandlerAsync(type, options, cb) {
          const i = wb.selectionHandlers.indexOf(options.handler);
          if (i >= 0) wb.selectionHandlers.splice(i, 1);
          if (cb) cb({ status: "succeeded" });
        },
        settings: {
          get: (k) => (settings.has(k) ? JSON.parse(settings.get(k)) : null),
          set: (k, v) => settings.set(k, JSON.stringify(v)),
          saveAsync: (cb) => cb({ status: "succeeded" }),
        },
      },
    },
  };
  const Excel = {
    ClearApplyTo: { contents: "Contents" },
    run: async (a, b) => (typeof a === "function" ? a(ctx) : b(ctx)),
  };
  return { Office, Excel };
}

module.exports = { Workbook, makeOffice };
