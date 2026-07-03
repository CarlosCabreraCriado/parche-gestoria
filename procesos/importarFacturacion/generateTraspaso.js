const fs = require("fs");
const path = require("path");
const XlsxPopulate = require("xlsx-populate");

const DEFAULT_TEMPLATE = "M:\\A3\\A3GESW\\PLANTILLA DE TRASPASO DE DATOS A A3GES.XLSX";
const DATA_START_ROW = 3;

const ISO_DATE = /^\d{4}-\d{2}-\d{2}$/;
const INT = /^-?(?:0|[1-9]\d*)$/;
const MONEY = /^-?\d+\.\d{1,2}$/;

// Coerce a value read from CSV (siempre string) al tipo correcto para A3GES.
// Reglas críticas validadas en PoC empresa 26:
//   - Fechas ISO YYYY-MM-DD -> Date (A3 rechaza strings)
//   - "890.00" -> float (locale ES interpreta punto como miles si es string)
//   - "26" -> int; "02290" queda como string (preservar leading zero)
//   - "0.909", "3.010" -> string literal (regex MONEY NO matchea 3+ decimales)
function coerce(value) {
  if (typeof value !== "string") return value;
  if (ISO_DATE.test(value)) {
    const [y, m, d] = value.split("-").map(Number);
    return new Date(y, m - 1, d);
  }
  if (INT.test(value)) return parseInt(value, 10);
  if (MONEY.test(value)) return parseFloat(value);
  return value;
}

// Normaliza texto para matching de encabezados (quita acentos, minúsculas,
// solo alfanumérico).
function normalize(value) {
  if (value === null || value === undefined) return "";
  let s = String(value).trim().toLowerCase();
  s = s.normalize("NFD").replace(/[̀-ͯ]/g, "");
  s = s.replace(/\(\*\)/g, "");
  s = s.replace(/[^0-9a-z]/g, "");
  return s;
}

// Lee CSV a array de objetos {header -> value}. Maneja BOM, quoted fields, y CRLF.
function loadCsv(filePath) {
  const raw = fs.readFileSync(filePath, "utf8").replace(/^﻿/, "");
  const rows = parseCsv(raw);
  if (rows.length === 0) return [];
  const headers = rows[0];
  return rows.slice(1).map((row) => {
    const obj = {};
    for (let i = 0; i < headers.length; i++) obj[headers[i]] = row[i] ?? "";
    return obj;
  });
}

function parseCsv(text) {
  const rows = [];
  let cur = [];
  let field = "";
  let inQuotes = false;
  let i = 0;
  const push = () => {
    cur.push(field);
    field = "";
  };
  const endRow = () => {
    push();
    rows.push(cur);
    cur = [];
  };
  while (i < text.length) {
    const ch = text[i];
    if (inQuotes) {
      if (ch === '"') {
        if (text[i + 1] === '"') {
          field += '"';
          i += 2;
          continue;
        }
        inQuotes = false;
        i++;
        continue;
      }
      field += ch;
      i++;
      continue;
    }
    if (ch === '"') { inQuotes = true; i++; continue; }
    if (ch === ",") { push(); i++; continue; }
    if (ch === "\r") { i++; continue; }
    if (ch === "\n") { endRow(); i++; continue; }
    field += ch;
    i++;
  }
  if (field !== "" || cur.length > 0) endRow();
  // La última "fila" puede ser vacía si el fichero acaba en \n
  if (rows.length && rows[rows.length - 1].length === 1 && rows[rows.length - 1][0] === "") {
    rows.pop();
  }
  return rows;
}

// Construye el mapa de columnas (encabezados fila 1 + fila 2) → índice de
// columna (1-based). Devuelve dos mapas:
//   qualified: nombreNormalizado(grupo+field) -> col
//   bare: nombreNormalizado(field) -> col (solo si el field es único en la hoja)
function columnMap(sheet) {
  const row1 = sheet.row(1);
  const row2 = sheet.row(2);
  // Determinar número máximo de columnas en las filas 1 y 2
  let maxCol = 0;
  const usedRange = sheet.usedRange();
  if (usedRange) maxCol = usedRange.endCell().columnNumber();

  let lastGroup = "";
  const qualified = new Map();
  const bareFirst = new Map();
  const bareCount = new Map();
  for (let c = 1; c <= maxCol; c++) {
    const gRaw = row1.cell(c).value();
    const fRaw = row2.cell(c).value();
    const g = gRaw !== undefined && gRaw !== null ? String(gRaw).trim() : "";
    if (g && !g.startsWith("***")) lastGroup = g;
    if (fRaw === undefined || fRaw === null || String(fRaw).trim() === "") continue;
    const field = String(fRaw).trim();
    const nf = normalize(field);
    const nq = lastGroup ? normalize(`${lastGroup}${field}`) : nf;
    if (!qualified.has(nq)) qualified.set(nq, c);
    bareCount.set(nf, (bareCount.get(nf) || 0) + 1);
    if (!bareFirst.has(nf)) bareFirst.set(nf, c);
  }
  const bare = new Map();
  for (const [k, v] of bareFirst) {
    if (bareCount.get(k) === 1) bare.set(k, v);
  }
  return { qualified, bare };
}

function appendRows(sheet, rows) {
  const { qualified, bare } = columnMap(sheet);
  let nextRow = DATA_START_ROW;
  let written = 0;
  const ambiguous = new Set();
  for (const row of rows) {
    const allEmpty = Object.values(row).every((v) => v === null || v === undefined || v === "");
    if (allEmpty) continue;
    for (const [rawKey, value] of Object.entries(row)) {
      if (value === null || value === undefined || value === "") continue;
      const nk = normalize(rawKey);
      const col = qualified.get(nk) ?? bare.get(nk);
      if (col === undefined) {
        if (nk && !qualified.has(nk) && !bare.has(nk)) ambiguous.add(rawKey);
        continue;
      }
      const coerced = coerce(value);
      const cell = sheet.cell(nextRow, col);
      cell.value(coerced);
      if (coerced instanceof Date) {
        // Formato fecha corta español; A3 lee la celda como date real, el
        // formato es cosmético.
        cell.style("numberFormat", "dd/mm/yyyy");
      }
    }
    nextRow++;
    written++;
  }
  return { written, ambiguous: Array.from(ambiguous).sort() };
}

// Copia la plantilla A3 a `output`, inyecta cada CSV encontrado en `inputDir`
// cuyo nombre coincida con una hoja de la plantilla, y guarda.
async function run(inputDir, output, template = DEFAULT_TEMPLATE, { verbose = true } = {}) {
  if (!fs.existsSync(template)) {
    throw new Error(`Plantilla no encontrada: ${template}`);
  }
  if (!fs.existsSync(inputDir) || !fs.statSync(inputDir).isDirectory()) {
    throw new Error(`Directorio de entrada no existe: ${inputDir}`);
  }

  const outputDir = path.dirname(output);
  if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir, { recursive: true });
  fs.copyFileSync(template, output);

  const workbook = await XlsxPopulate.fromFileAsync(output);
  const sheetNames = workbook.sheets().map((s) => s.name());

  let total = 0;
  const sheets = [];
  for (const sheetName of sheetNames) {
    const csvPath = path.join(inputDir, `${sheetName}.csv`);
    if (!fs.existsSync(csvPath)) continue;
    const rows = loadCsv(csvPath);
    const { written, ambiguous } = appendRows(workbook.sheet(sheetName), rows);
    sheets.push({
      sheet: sheetName,
      rows: written,
      csv: path.basename(csvPath),
      unknown_cols: ambiguous,
    });
    if (verbose) {
      const cols = ambiguous.length ? `   columnas ignoradas: ${JSON.stringify(ambiguous)}` : "";
      console.log(`  ${sheetName.padEnd(35)}  ${String(written).padStart(5)} filas  (${path.basename(csvPath)})${cols}`);
    }
    total += written;
  }

  await workbook.toFileAsync(output);

  if (verbose) {
    console.log();
    console.log(`Salida: ${output}`);
    console.log(`CSVs usados: ${sheets.length}   Filas totales anexadas: ${total}`);
    if (sheets.length === 0) {
      console.log("AVISO: ningún CSV coincidió con un nombre de hoja de la plantilla.");
      console.log("Hojas disponibles:");
      for (const sn of sheetNames) console.log(`  - ${sn}`);
    }
  }

  return { output, sheets, total_rows: total };
}

module.exports = { run, coerce, normalize, DEFAULT_TEMPLATE };
