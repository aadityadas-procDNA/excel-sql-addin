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

  return {
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
