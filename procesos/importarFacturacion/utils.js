const fs = require("fs");
const path = require("path");

function _str(value) {
  if (value === null || value === undefined) return "";
  return String(value).trim();
}

function _toInt(value) {
  if (value === null || value === undefined || value === "") return null;
  if (typeof value === "number" && Number.isFinite(value)) {
    return Math.trunc(value);
  }
  const s = String(value).trim();
  if (s === "") return null;
  const n = Number(s);
  if (!Number.isFinite(n)) return null;
  return Math.trunc(n);
}

function _toFloat(value) {
  if (value === null || value === undefined || value === "") return null;
  if (typeof value === "number") return Number.isFinite(value) ? value : null;
  const s = String(value).trim();
  if (s === "") return null;
  const n = Number(s);
  return Number.isFinite(n) ? n : null;
}

// xlsx-populate devuelve fechas como número serial de Excel cuando la celda es
// tipo fecha; convertirlo a Date. Si viene ya como Date u objeto, respetamos.
function excelSerialToDate(serial) {
  // Epoch Excel: 1899-12-30 (compensa bug 1900 leap year)
  const epoch = Date.UTC(1899, 11, 30);
  const ms = Math.round(serial * 86400 * 1000);
  return new Date(epoch + ms);
}

function _toDate(value, fallback) {
  if (value instanceof Date && !isNaN(value)) return value;
  if (typeof value === "number" && Number.isFinite(value)) {
    return excelSerialToDate(value);
  }
  if (typeof value === "string" && value.trim()) {
    const s = value.trim();
    // ISO YYYY-MM-DD
    let m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
    // DD/MM/YYYY o DD-MM-YYYY
    m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
    if (m) return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
  }
  return fallback;
}

function pad5(n) {
  const s = String(n);
  return s.length >= 5 ? s : "0".repeat(5 - s.length) + s;
}

function isoDate(d) {
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}`;
}

function stampYYYYMMDDHHmm(d) {
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  const hh = String(d.getHours()).padStart(2, "0");
  const mi = String(d.getMinutes()).padStart(2, "0");
  return `${yyyy}${mm}${dd}_${hh}${mi}`;
}

function ensureDir(dir) {
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
}

// xlsx-populate: `usedRange().value()` devuelve un array 2D RELATIVO a la
// esquina superior izquierda del rango usado. Si el rango no empieza en A1,
// los índices de columna quedan desplazados. Esta función devuelve las filas
// alineadas por índice absoluto de columna (cells[0] siempre es col A).
function readAbsoluteRows(sheet) {
  const usedRange = sheet.usedRange();
  if (!usedRange) return { rows: [], startRow: 1 };
  const values = usedRange.value();
  const startRow = usedRange.startCell().rowNumber();
  const startCol = usedRange.startCell().columnNumber();
  const padLeft = startCol - 1;
  const rows = values.map((row, idx) => {
    const cells = padLeft > 0 ? new Array(padLeft).fill(undefined).concat(row || []) : row || [];
    return { rowIndex: startRow + idx, cells };
  });
  return { rows, startRow };
}

function csvEscape(value) {
  if (value === null || value === undefined) return "";
  let s = typeof value === "string" ? value : String(value);
  if (s.includes(",") || s.includes('"') || s.includes("\n") || s.includes("\r")) {
    s = '"' + s.replace(/"/g, '""') + '"';
  }
  return s;
}

function writeCsv(filePath, headers, rows) {
  const lines = [];
  lines.push(headers.map(csvEscape).join(","));
  for (const row of rows) {
    lines.push(row.map(csvEscape).join(","));
  }
  fs.writeFileSync(filePath, lines.join("\r\n") + "\r\n", { encoding: "utf8" });
}

module.exports = {
  _str,
  _toInt,
  _toFloat,
  _toDate,
  excelSerialToDate,
  pad5,
  isoDate,
  stampYYYYMMDDHHmm,
  ensureDir,
  readAbsoluteRows,
  writeCsv,
};
