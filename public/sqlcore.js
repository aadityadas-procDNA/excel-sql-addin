/*
 * Pure SQL helpers shared by the task pane and the Node tests.
 * Converts Excel range values into SQLite tables and query results back into Excel values.
 */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.SqlCore = factory();
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  const KEYWORDS = new Set(
    ("select from where group by order having limit offset join inner left right full outer cross on as and or not " +
      "null is in like glob between case when then else end union intersect except all distinct table index values " +
      "insert update delete create drop alter with into set exists natural using default primary key check")
      .split(" ")
  );

  // Collects every identifier-like token in the SQL, skipping string literals and comments.
  function extractIdentifiers(sql) {
    const ids = [];
    const n = sql.length;
    let i = 0;
    while (i < n) {
      const c = sql[i];
      if (c === "'") {
        i++;
        while (i < n && !(sql[i] === "'" && sql[i + 1] !== "'")) i += sql[i] === "'" ? 2 : 1;
        i++;
      } else if (c === "-" && sql[i + 1] === "-") {
        while (i < n && sql[i] !== "\n") i++;
      } else if (c === "/" && sql[i + 1] === "*") {
        const end = sql.indexOf("*/", i + 2);
        i = end < 0 ? n : end + 2;
      } else if (c === '"' || c === "`") {
        let j = i + 1;
        let s = "";
        while (j < n) {
          if (sql[j] === c) {
            if (sql[j + 1] === c) { s += c; j += 2; continue; }
            break;
          }
          s += sql[j++];
        }
        ids.push(s);
        i = j + 1;
      } else if (c === "[") {
        const end = sql.indexOf("]", i + 1);
        const stop = end < 0 ? n : end;
        ids.push(sql.slice(i + 1, stop));
        i = stop + 1;
      } else if (/[A-Za-z_À-￿]/.test(c)) {
        let j = i + 1;
        while (j < n && /[A-Za-z0-9_$À-￿]/.test(sql[j])) j++;
        ids.push(sql.slice(i, j));
        i = j;
      } else {
        i++;
      }
    }
    return ids;
  }

  // Maps identifiers used in the SQL to workbook objects. Priority: table > named range > sheet.
  // Identifiers containing "!" are treated as range addresses, e.g. [Sheet1!A1:D50].
  function resolveSources(ids, objects) {
    const priority = { table: 0, name: 1, sheet: 2 };
    const byLower = new Map();
    [...objects]
      .sort((a, b) => priority[a.kind] - priority[b.kind])
      .forEach((o) => { if (!byLower.has(o.name.toLowerCase())) byLower.set(o.name.toLowerCase(), o); });

    const sources = new Map();
    for (const id of ids) {
      const key = id.toLowerCase();
      if (!id || sources.has(key)) continue;
      if (id.includes("!")) sources.set(key, { kind: "range", name: id, ref: id });
      else if (byLower.has(key)) sources.set(key, { kind: byLower.get(key).kind, name: id, ref: byLower.get(key).name });
    }
    return [...sources.values()];
  }

  // "'My Sheet'!A1:B5" -> { sheet: "My Sheet", address: "A1:B5" }
  function splitAddress(ref) {
    const at = ref.lastIndexOf("!");
    let sheet = ref.slice(0, at).trim();
    if (sheet.startsWith("'") && sheet.endsWith("'")) sheet = sheet.slice(1, -1).replace(/''/g, "'");
    return { sheet, address: ref.slice(at + 1).trim() };
  }

  function quoteRef(name) {
    if (/^[A-Za-z_][A-Za-z0-9_]*$/.test(name) && !KEYWORDS.has(name.toLowerCase())) return name;
    if (!name.includes("]")) return "[" + name + "]";
    return quoteIdent(name);
  }

  function quoteIdent(name) {
    return '"' + String(name).replace(/"/g, '""') + '"';
  }

  // Blank headers become ColumnN; duplicates get a _2, _3... suffix.
  function normalizeHeaders(row) {
    const used = new Set();
    return row.map((h, i) => {
      const base = h === null || h === undefined || String(h).trim() === "" ? "Column" + (i + 1) : String(h).trim();
      let name = base;
      for (let k = 2; used.has(name.toLowerCase()); k++) name = base + "_" + k;
      used.add(name.toLowerCase());
      return name;
    });
  }

  function isDateFormat(fmt) {
    if (typeof fmt !== "string") return false;
    if (/\[(h+|m+|s+)\]/i.test(fmt)) return false; // elapsed-time formats stay numeric
    const s = fmt.replace(/"[^"]*"/g, "").replace(/\\./g, "").replace(/\[[^\]]*\]/g, "").trim();
    if (s === "" || /^(general|@)$/i.test(s)) return false;
    return /[dmyhs]/i.test(s);
  }

  // Excel serial date -> "YYYY-MM-DD", "YYYY-MM-DD HH:MM:SS" or "HH:MM:SS" (time only).
  function serialToIso(v) {
    const totalSec = Math.round(v * 86400);
    const d = new Date((totalSec - 25569 * 86400) * 1000);
    if (isNaN(d.getTime())) return v;
    const iso = d.toISOString();
    if (v >= 0 && v < 1) return iso.slice(11, 19);
    return totalSec % 86400 === 0 ? iso.slice(0, 10) : iso.slice(0, 19).replace("T", " ");
  }

  const isBlank = (v) => v === "" || v === null || v === undefined;

  function mergeKind(a, b) {
    if (a === null || a === b) return b;
    if ((a === "int" && b === "real") || (a === "real" && b === "int")) return "real";
    return "mixed";
  }

  const SQL_TYPES = { int: "INTEGER", real: "REAL", text: "TEXT", bool: "INTEGER", date: "TEXT", mixed: "", empty: "" };
  const TYPE_LABELS = { int: "INTEGER", real: "REAL", text: "TEXT", bool: "BOOLEAN", date: "DATE", mixed: "ANY", empty: "ANY" };

  // headerRow: array of header values; rows: 2D data values;
  // formats: 2D number formats for the first few data rows (used to detect date columns).
  function prepareTable(headerRow, rows, formats) {
    const names = normalizeHeaders(headerRow || []);
    const data = rows.filter((r) => r.some((v) => !isBlank(v)));

    const columns = names.map((name, c) => {
      let fmt = null;
      for (let r = 0; formats && r < formats.length && r < rows.length; r++) {
        if (typeof rows[r][c] === "number") { fmt = formats[r][c]; break; }
      }
      let kind = null;
      for (const row of data) {
        const v = row[c];
        if (isBlank(v)) continue;
        const k = typeof v === "number" ? (Number.isInteger(v) ? "int" : "real") : typeof v === "boolean" ? "bool" : "text";
        kind = mergeKind(kind, k);
        if (kind === "mixed") break;
      }
      const isDate = isDateFormat(fmt);
      if (isDate && (kind === "int" || kind === "real")) kind = "date";
      kind = kind || "empty";
      return { name, kind, isDate, type: TYPE_LABELS[kind], sqlType: SQL_TYPES[kind] };
    });

    const converted = data.map((row) =>
      columns.map((col, c) => {
        const v = row[c];
        if (isBlank(v)) return null;
        if (typeof v === "boolean") return v ? 1 : 0;
        if (typeof v === "number" && col.isDate) return serialToIso(v);
        return v;
      })
    );
    return { columns, rows: converted };
  }

  function loadTable(db, name, table) {
    if (!table.columns.length) {
      db.run(`CREATE TABLE ${quoteIdent(name)} (${quoteIdent("Column1")})`);
      return;
    }
    const cols = table.columns.map((c) => (quoteIdent(c.name) + " " + c.sqlType).trim()).join(", ");
    db.run(`CREATE TABLE ${quoteIdent(name)} (${cols})`);
    const stmt = db.prepare(`INSERT INTO ${quoteIdent(name)} VALUES (${table.columns.map(() => "?").join(", ")})`);
    db.run("BEGIN");
    try {
      for (const row of table.rows) stmt.run(row);
      db.run("COMMIT");
    } catch (e) {
      db.run("ROLLBACK");
      throw e;
    } finally {
      stmt.free();
    }
  }

  // Runs one or more statements; returns the last result set.
  function runQuery(db, sql) {
    const results = db.exec(sql);
    return results.length ? results[results.length - 1] : { columns: [], values: [] };
  }

  function toCell(v) {
    if (v === null || v === undefined) return "";
    if (v instanceof Uint8Array) return `[BLOB ${v.length} bytes]`;
    if (typeof v === "string") {
      const s = v.length > 32767 ? v.slice(0, 32767) : v;
      return s.startsWith("=") ? "'" + s : s; // never let result text become a formula
    }
    return v;
  }

  // Header row + data rows. ensureRow adds an empty data row (Excel tables need one).
  function toExcelValues(result, ensureRow) {
    const out = [result.columns.map(toCell)];
    for (const row of result.values) out.push(row.map(toCell));
    if (ensureRow && out.length === 1) out.push(result.columns.map(() => ""));
    return out;
  }

  // ---------------------------------------------------------------- lookups (XLOOKUP replacement)

  const NUMERIC_TEXT = /^[-+]?(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?$/;

  // mode: "exact" (case-sensitive), "nocase" (like XLOOKUP), "loose" (nocase + trims spaces + numbers stored as text match numbers)
  function keyPart(v, mode) {
    if (isBlank(v)) return null;
    if (typeof v === "number") return "n" + v;
    if (typeof v === "boolean") return "b" + v;
    let s = String(v);
    if (mode === "loose") {
      s = s.trim().replace(/\s+/g, " ");
      if (s === "") return null;
      if (NUMERIC_TEXT.test(s)) return "n" + Number(s);
    }
    if (mode !== "exact") s = s.toLowerCase();
    return "s" + s;
  }

  function lookupKey(values, mode) {
    const parts = values.map((v) => keyPart(v, mode));
    return parts.includes(null) ? null : parts.join("\u0001");
  }

  // Index source rows by key. dup: "first" | "last". Returns { index: Map(key -> row), duplicates }.
  function buildLookupIndex(rows, keyIdx, mode, dup) {
    const index = new Map();
    let duplicates = 0;
    for (const row of rows) {
      const key = lookupKey(keyIdx.map((i) => row[i]), mode);
      if (key === null) continue;
      if (index.has(key)) {
        duplicates++;
        if (dup !== "last") continue;
      }
      index.set(key, row);
    }
    return { index, duplicates };
  }

  // targetKeys: rows of key values. Returns one output column (array) per valIdx, plus match stats.
  // Blank keys give blank output; non-blank keys with no match give notFound.
  function applyLookup(targetKeys, index, valIdx, mode, notFound) {
    const cols = valIdx.map(() => new Array(targetKeys.length));
    let matched = 0;
    let missing = 0;
    targetKeys.forEach((keys, r) => {
      const key = lookupKey(keys, mode);
      const row = key === null ? undefined : index.get(key);
      if (row) matched++;
      else if (key !== null) missing++;
      valIdx.forEach((c, j) => {
        cols[j][r] = row ? toCell(row[c]) : key === null ? "" : notFound;
      });
    });
    return { columns: cols, matched, missing };
  }

  // Number format of the first non-blank value in column c (formats aligned with rows).
  function sampleFormat(rows, formats, c) {
    for (let r = 0; formats && r < formats.length && r < rows.length; r++) {
      if (!isBlank(rows[r][c])) return formats[r][c];
    }
    return null;
  }

  // ---------------------------------------------------------------- A1 addresses (for change tracking)

  const MAX_ROW = 1048575;
  const MAX_COL = 16383;

  function colToIndex(letters) {
    let n = 0;
    for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n - 1;
  }

  function parseCellRef(ref) {
    const m = /^([A-Z]*)(\d*)$/i.exec(ref.trim());
    if (!m || (m[1] === "" && m[2] === "")) throw new Error("Bad cell reference: " + ref);
    const hasC = m[1] !== "";
    const hasR = m[2] !== "";
    return {
      r1: hasR ? Number(m[2]) - 1 : 0,
      r2: hasR ? Number(m[2]) - 1 : MAX_ROW,
      c1: hasC ? colToIndex(m[1]) : 0,
      c2: hasC ? colToIndex(m[1]) : MAX_COL,
    };
  }

  // "A1", "B2:D9", "A:C", "3:5", "Sheet1!A1:B2", comma-separated areas -> [{ r1, c1, r2, c2 }] (0-based, inclusive)
  function parseAddress(address) {
    return address.split(",").map((area) => {
      const a = area.slice(area.lastIndexOf("!") + 1).replace(/\$/g, "");
      const [p1, p2 = p1] = a.split(":");
      const x = parseCellRef(p1);
      const y = parseCellRef(p2);
      return { r1: Math.min(x.r1, y.r1), r2: Math.max(x.r2, y.r2), c1: Math.min(x.c1, y.c1), c2: Math.max(x.c2, y.c2) };
    });
  }

  function boundsOf(rowIndex, columnIndex, rowCount, columnCount) {
    return { r1: rowIndex, c1: columnIndex, r2: rowIndex + rowCount - 1, c2: columnIndex + columnCount - 1 };
  }

  const intersects = (a, b) => a.r1 <= b.r2 && b.r1 <= a.r2 && a.c1 <= b.c2 && b.c1 <= a.c2;
  const contains = (outer, inner) => outer.r1 <= inner.r1 && inner.r2 <= outer.r2 && outer.c1 <= inner.c1 && inner.c2 <= outer.c2;
  const sameBounds = (a, b) => a.r1 === b.r1 && a.c1 === b.c1 && a.r2 === b.r2 && a.c2 === b.c2;

  function intersection(a, b) {
    if (!intersects(a, b)) return null;
    return { r1: Math.max(a.r1, b.r1), c1: Math.max(a.c1, b.c1), r2: Math.min(a.r2, b.r2), c2: Math.min(a.c2, b.c2) };
  }

  function indexToCol(c) {
    let s = "";
    for (c++; c > 0; c = Math.floor((c - 1) / 26)) s = String.fromCharCode(65 + ((c - 1) % 26)) + s;
    return s;
  }

  // { r1: 0, c1: 0, r2: 8, c2: 3 } -> "A1:D9"
  function toA1(b) {
    const a = indexToCol(b.c1) + (b.r1 + 1);
    const z = indexToCol(b.c2) + (b.r2 + 1);
    return a === z ? a : a + ":" + z;
  }

  // ---------------------------------------------------------------- reference mode

  // "='My Sheet'!$A$1:$C$9" -> { sheet, bounds }. Formulas, constants and multi-area names give null.
  function parseNameFormula(formula) {
    if (typeof formula !== "string") return null;
    const f = formula.replace(/^=/, "").trim();
    if (!f.includes("!") || /[(),]/.test(f.replace(/'[^']*'/g, ""))) return null;
    const { sheet, address } = splitAddress(f);
    try {
      return { sheet, bounds: parseAddress(address)[0] };
    } catch (e) {
      return null;
    }
  }

  const filled = (v) => v !== "" && v !== null && v !== undefined;

  // Expands a non-empty cell to its block of data, like Ctrl+A (Excel's "current region"): keeps adding
  // adjacent rows and columns (diagonals included) that contain a value. Stays inside limits (the used range).
  // read(boundsList) resolves to one 2D value array per bounds; probes grow geometrically to keep round trips few.
  async function currentRegion(start, limits, read) {
    if (!limits || !contains(limits, start)) return null;
    if (!filled((await read([start]))[0][0][0])) return null;
    const b = { ...start };
    let rowStep = 200;
    let colStep = 20;
    for (let i = 0; i < 5000; i++) {
      const c1 = Math.max(b.c1 - 1, limits.c1), c2 = Math.min(b.c2 + 1, limits.c2);
      const r1 = Math.max(b.r1 - 1, limits.r1), r2 = Math.min(b.r2 + 1, limits.r2);
      const probes = [];
      if (b.r2 < limits.r2) probes.push({ side: "down", b: { r1: b.r2 + 1, r2: Math.min(b.r2 + rowStep, limits.r2), c1, c2 } });
      if (b.r1 > limits.r1) probes.push({ side: "up", b: { r1: Math.max(b.r1 - rowStep, limits.r1), r2: b.r1 - 1, c1, c2 } });
      if (b.c2 < limits.c2) probes.push({ side: "right", b: { r1, r2, c1: b.c2 + 1, c2: Math.min(b.c2 + colStep, limits.c2) } });
      if (b.c1 > limits.c1) probes.push({ side: "left", b: { r1, r2, c1: Math.max(b.c1 - colStep, limits.c1), c2: b.c1 - 1 } });
      if (!probes.length) break;

      const values = await read(probes.map((p) => p.b));
      let grew = false;
      probes.forEach((p, k) => {
        const v = values[k];
        const rows = v.length;
        const cols = v[0].length;
        const rowHas = (r) => v[r].some(filled);
        const colHas = (c) => v.some((row) => filled(row[c]));
        let n = 0;
        if (p.side === "down") { while (n < rows && rowHas(n)) n++; b.r2 += n; }
        if (p.side === "up") { while (n < rows && rowHas(rows - 1 - n)) n++; b.r1 -= n; }
        if (p.side === "right") { while (n < cols && colHas(n)) n++; b.c2 += n; }
        if (p.side === "left") { while (n < cols && colHas(cols - 1 - n)) n++; b.c1 -= n; }
        if (n) grew = true;
        const full = p.side === "down" || p.side === "up" ? n === rows : n === cols;
        if (full && (p.side === "down" || p.side === "up")) rowStep = Math.min(rowStep * 2, Math.max(200, Math.floor(200000 / (c2 - c1 + 1))));
        if (full && (p.side === "left" || p.side === "right")) colStep = Math.min(colStep * 2, Math.max(20, Math.floor(200000 / (r2 - r1 + 1))));
      });
      if (!grew) break;
    }
    return b;
  }

  // Decides which block of cells a selection refers to.
  // info: { sheet, sel, used, tables: [{ name, bounds }] (this sheet), tableNames: all, lower-cased,
  //         names: [{ name, sheet, bounds }], region? }
  // Returns { block, table?, name? }, or { needsRegion: true } when a single loose cell must be expanded first.
  function pickReferenceBlock(info) {
    const { sel, used } = info;
    const single = sel.r1 === sel.r2 && sel.c1 === sel.c2;
    const table = info.tables.find((t) => contains(t.bounds, sel));
    if (table) return { block: table.bounds, table };

    // A named range sharing a table's name resolves to the table, so it can't be offered.
    const tableNames = new Set((info.tableNames || []).concat(info.tables.map((t) => t.name.toLowerCase())));
    const name = info.names.find(
      (n) =>
        n.sheet.toLowerCase() === info.sheet.toLowerCase() &&
        !tableNames.has(n.name.toLowerCase()) &&
        (single ? contains(n.bounds, sel) : sameBounds(n.bounds, sel))
    );
    if (!used) return { block: null };
    if (name) return { block: intersection(name.bounds, used), name };
    if (!single) return { block: intersection(sel, used) };
    if (info.region === undefined) return { needsRegion: true };
    return { block: info.region };
  }

  // SQL literal for a cell value, matching how the cell is loaded (dates become ISO text).
  function sqlLiteral(value, format) {
    if (!filled(value)) return null;
    if (typeof value === "boolean") return value ? "1" : "0";
    if (typeof value === "number") return isDateFormat(format) ? "'" + serialToIso(value) + "'" : String(value);
    return "'" + String(value).replace(/'/g, "''") + "'";
  }

  // Lists the SQL references for a picked block, best first. headers: values of the block's first row.
  // taken: lower-cased names of all tables and named ranges (a sheet with such a name can't be referenced by name).
  // cell: { value, format } of a single selected cell, offered as a literal for WHERE clauses.
  function referenceCandidates(info, pick, headers, taken, cell) {
    const block = pick.block;
    if (!block) return { block: null, candidates: [], columns: [] };
    const columns = normalizeHeaders(headers || []);
    const { sel } = info;
    const single = sel.r1 === sel.r2 && sel.c1 === sel.c2;
    const out = [];
    if (pick.table) {
      out.push({ kind: "table", label: pick.table.name, text: quoteRef(pick.table.name) });
      // Part of a table that starts at its header row is still a valid source with column names.
      if (!single && sel.r1 === block.r1 && !sameBounds(sel, block)) {
        const address = `${info.sheet}!${toA1(sel)}`;
        out.push({ kind: "range", label: address, text: `[${address}]` });
      }
    } else {
      if (pick.name) out.push({ kind: "name", label: pick.name.name, text: quoteRef(pick.name.name) });
      if (info.used && sameBounds(block, info.used) && !(taken || []).includes(info.sheet.toLowerCase())) {
        out.push({ kind: "sheet", label: info.sheet, text: quoteRef(info.sheet) });
      }
      const address = `${info.sheet}!${toA1(block)}`;
      out.push({ kind: "range", label: address, text: `[${address}]` });
    }
    // Names of the selected columns (all of them only when it's a single column: otherwise that's just "*").
    const c1 = Math.max(sel.c1, block.c1);
    const c2 = Math.min(sel.c2, block.c2);
    const picked = columns.slice(c1 - block.c1, c2 - block.c1 + 1);
    if (block.r2 > block.r1 && picked.length && (picked.length === 1 || picked.length < columns.length)) {
      out.push({ kind: picked.length === 1 ? "column" : "columns", label: picked.join(", "), text: picked.map(quoteRef).join(", ") });
    }
    if (single && sel.r1 > block.r1 && cell) {
      const literal = sqlLiteral(cell.value, cell.format);
      if (literal !== null) out.push({ kind: "value", label: literal, text: literal });
    }
    return { block, candidates: out, columns, rows: block.r2 - block.r1 };
  }

  // ---------------------------------------------------------------- join builder

  // spec: { base: { name, alias }, joins: [{ type, name, alias, on: [{ leftAlias, leftColumn, rightColumn }] }],
  //         columns: [{ alias, column, as }], nocase, trim }
  function buildJoinSql(spec) {
    const key = (alias, col) => {
      const ref = alias + "." + quoteRef(col);
      return spec.trim ? `TRIM(${ref})` : ref;
    };
    const select = spec.columns.length
      ? spec.columns
          .map((c) => `  ${c.alias}.${quoteRef(c.column)}` + (c.as && c.as !== c.column ? ` AS ${quoteRef(c.as)}` : ""))
          .join(",\n")
      : "  *";
    const lines = ["SELECT", select, `FROM ${quoteRef(spec.base.name)} AS ${spec.base.alias}`];
    for (const j of spec.joins) {
      const on = j.on
        .map((k) => `${key(k.leftAlias, k.leftColumn)} = ${key(j.alias, k.rightColumn)}${spec.nocase ? " COLLATE NOCASE" : ""}`)
        .join("\n   AND ");
      lines.push(`${j.type} JOIN ${quoteRef(j.name)} AS ${j.alias}`, `    ON ${on}`);
    }
    return lines.join("\n") + ";";
  }

  // Picks the most likely join key between two column lists: shared names, preferring *id/key/code.
  function guessJoinKey(leftColumns, rightColumns) {
    const right = new Map(rightColumns.map((c) => [c.toLowerCase(), c]));
    const shared = leftColumns.filter((c) => right.has(c.toLowerCase()));
    if (!shared.length) return null;
    const best = shared.find((c) => /(id|key|code|no|number)$/i.test(c)) || shared[0];
    return { left: best, right: right.get(best.toLowerCase()) };
  }

  return {
    buildLookupIndex,
    applyLookup,
    lookupKey,
    sampleFormat,
    parseAddress,
    boundsOf,
    intersects,
    contains,
    toA1,
    parseNameFormula,
    currentRegion,
    pickReferenceBlock,
    referenceCandidates,
    sqlLiteral,
    buildJoinSql,
    guessJoinKey,
    extractIdentifiers,
    resolveSources,
    splitAddress,
    quoteRef,
    quoteIdent,
    normalizeHeaders,
    isDateFormat,
    serialToIso,
    prepareTable,
    loadTable,
    runQuery,
    toExcelValues,
  };
});
