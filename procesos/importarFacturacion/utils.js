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

// Precio puntual escrito a mano en la columna IMPORTE de una fila (nóminas y
// trámites). Devuelve { valor:número } si hay precio, { valor:null } si la celda
// está vacía, o { error:texto } si hay algo escrito que no es un número. Un
// IMPORTE ilegible NO cae a la tarifa de catálogo: facturaría un importe
// distinto del que el usuario quiso teclear y nadie lo notaría.
function leerImporte(raw) {
  if (raw === null || raw === undefined) return { valor: null };
  if (typeof raw === "number") {
    return Number.isFinite(raw) ? { valor: raw } : { error: String(raw) };
  }
  const original = _str(raw);
  if (original === "") return { valor: null };
  let s = original.replace(/[€\s]/g, "");
  // Con coma se asume formato español: el punto es separador de millares.
  if (s.includes(",")) s = s.replace(/\./g, "").replace(",", ".");
  const n = Number(s);
  return Number.isFinite(n) ? { valor: n } : { error: original };
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

// La fecha de facturación la elige el usuario en el formulario (datepicker de
// Material). Llega como Date, pero al cruzar el IPC de Electron se serializa a
// ISO en UTC: se reconstruye desde los componentes LOCALES para que no se
// desplace un día. Devuelve null si no hay valor o no es interpretable.
function fechaDesdeFormulario(value) {
  if (value === null || value === undefined || value === "") return null;
  if (typeof value === "string") {
    const m = value.trim().match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  }
  const d = value instanceof Date ? value : new Date(value);
  if (isNaN(d)) return null;
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

function fechaCorta(d) {
  const dd = String(d.getDate()).padStart(2, "0");
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  return `${dd}/${mm}/${d.getFullYear()}`;
}

// Añade la fecha del archivo de origen al final de la descripción. Recorta la
// BASE dejando hueco al sufijo: concatenar y recortar después se come la fecha
// justo en las descripciones largas, que es cuando más falta hace.
function conFecha(base, fecha, max = 250) {
  const texto = String(base ?? "");
  if (!fecha) return texto.slice(0, max);
  const sufijo = ` - ${fechaCorta(fecha)}`;
  return texto.slice(0, Math.max(0, max - sufijo.length)) + sufijo;
}

// Límites reales de A3 en la línea de "concepto pendiente" (medidos 2026-07-24,
// ver procesos/importarFacturacion/prueba-limite-a3/): la Descripción es un campo
// de ancho fijo de 50 caracteres —lo que pase de ahí lo descarta A3— y la
// Descripción Ampliada admite del orden de 500.
const LIMITE_DESC = 50;
const LIMITE_DESC_AMPLIADA = 500;

// Recorta `texto` a `max` caracteres sin partir palabras: corta en el último
// espacio que quepa y limpia la puntuación de cola, para que la línea no acabe a
// mitad de palabra. Si la primera palabra ya excede `max` (no hay ningún
// espacio), corta a pelo para no devolver una cadena vacía.
function recortarPorPalabra(texto, max) {
  const s = String(texto ?? "").trim();
  if (s.length <= max) return s;
  const trozo = s.slice(0, max);
  const corte = trozo.lastIndexOf(" ");
  const base = corte > 0 ? trozo.slice(0, corte) : trozo;
  return base.replace(/[\s\-–—.,;:]+$/, "").trim() || trozo.trim();
}

// Reparte un concepto y su detalle entre Descripción (50) y Descripción Ampliada
// (500). Como A3 solo guarda 50 caracteres de la Descripción, ahí va el concepto
// recortado por palabra; si no cupo entero, su texto ÍNTEGRO se conserva al
// inicio de la Ampliada (así no se pierde nada) seguido del resto del detalle
// (nombre, observación...). Normaliza espacios y saltos de línea a un solo
// espacio para que los textos multilínea del catálogo salgan en una sola línea.
// Devuelve { descripcion, ampliada }.
function repartirDescripcion(concepto, extras = []) {
  const norm = (x) => String(x ?? "").replace(/\s+/g, " ").trim();
  const texto = norm(concepto);
  const descripcion = recortarPorPalabra(texto, LIMITE_DESC);
  const partes = [];
  if (texto.length > LIMITE_DESC) partes.push(texto);
  for (const e of extras) {
    const s = norm(e);
    if (s) partes.push(s);
  }
  const ampliada = recortarPorPalabra(partes.join(" · "), LIMITE_DESC_AMPLIADA);
  return { descripcion, ampliada };
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

// Normaliza el texto de una cabecera para poder emparejarla sin depender de
// mayúsculas, acentos, espacios ni signos: "CONCEPTO FACT" -> "conceptofact".
function normalizeHeader(value) {
  if (value === null || value === undefined) return "";
  let s = String(value).trim().toLowerCase();
  s = s.normalize("NFD").replace(/[̀-ͯ]/g, "");
  s = s.replace(/[^0-9a-z]/g, "");
  return s;
}

// Dado el array de celdas de una fila de cabecera y un diccionario
// { campoLogico: [sinonimoNormalizado, ...] }, devuelve:
//   - cols:    { campoLogico -> índice de columna (0-based) } para los encontrados
//   - present: Map(cabeceraNormalizada -> índice) de todas las cabeceras
// Se usa el primer sinónimo presente y la primera aparición de cada cabecera.
//
// El emparejamiento tiene dos niveles. El NIVEL 1 (exacto) es la vía principal:
// determinista y sin falsos positivos. El NIVEL 2 (por prefijo, opt-in con
// `fuzzy`) solo actúa sobre lo que el nivel 1 no resolvió, para tolerar variantes
// del cliente —plural, sufijos— sin listar cada una. Su regla de oro: ante la
// menor ambigüedad NO asigna. En una importación de facturación es preferible
// dejar un campo sin resolver (dato de menos, error claro) a capturar la columna
// equivocada en silencio y facturar mal.
function resolveHeaderColumns(headerCells, synonyms, options = {}) {
  const { fuzzy = false, minRoot = 4 } = options;
  const present = new Map();
  (headerCells || []).forEach((cell, idx) => {
    const nk = normalizeHeader(cell);
    if (nk && !present.has(nk)) present.set(nk, idx);
  });

  const cols = {};
  const claimed = new Set(); // índices de columna ya asignados

  // Nivel 1 — coincidencia EXACTA del nombre normalizado con un sinónimo.
  for (const [field, names] of Object.entries(synonyms)) {
    for (const n of names) {
      if (present.has(n)) {
        cols[field] = present.get(n);
        claimed.add(present.get(n));
        break;
      }
    }
  }

  if (!fuzzy) return { cols, present };

  // Nivel 2 — por prefijo, solo para campos aún sin resolver y columnas aún sin
  // dueño. Una cabecera es candidata de un campo si EMPIEZA por alguno de sus
  // sinónimos (raíz de longitud >= minRoot, para que raíces cortas no capturen de
  // más): así "observacion" capta "OBSERVACIONES" o "OBSERVACION 2" sin enumerar
  // cada variante. Se asigna solo con correspondencia ÚNICA EN AMBOS SENTIDOS:
  // una sola columna candidata para el campo Y esa columna no candidata de ningún
  // otro campo pendiente. Cualquier ambigüedad se deja sin asignar.
  const unresolved = Object.keys(synonyms).filter((f) => cols[f] === undefined);
  const avail = [...present.entries()].filter(([, idx]) => !claimed.has(idx));

  const candByField = new Map(); // field -> [idx, ...]
  const fieldsByIdx = new Map(); // idx -> Set(field) que la reclaman
  for (const field of unresolved) {
    const roots = synonyms[field].filter((n) => n.length >= minRoot);
    const cand = [];
    for (const [key, idx] of avail) {
      if (roots.some((r) => key.startsWith(r))) {
        cand.push(idx);
        if (!fieldsByIdx.has(idx)) fieldsByIdx.set(idx, new Set());
        fieldsByIdx.get(idx).add(field);
      }
    }
    candByField.set(field, cand);
  }

  for (const field of unresolved) {
    const cand = candByField.get(field);
    if (cand.length !== 1) continue; // 0 = ninguna; >1 = ambigua para el campo
    const idx = cand[0];
    if (fieldsByIdx.get(idx).size !== 1) continue; // otra columna la disputa
    cols[field] = idx;
    claimed.add(idx);
  }

  return { cols, present };
}

// Nº máx. de filas iniciales a escanear en cada hoja buscando la cabecera.
const HEADER_SCAN_ROWS = 30;

// Localiza, dentro de un libro, la primera hoja cuya fila de cabecera satisface
// `isHeader(cols)`, resolviendo las columnas por nombre según `synonyms`. Permite
// que cada importador funcione con distintos formatos/hojas del cliente sin
// depender de posiciones fijas ni del orden de las hojas.
// Opciones:
//   - scanRows: nº de filas iniciales a inspeccionar por hoja.
//   - mergeUp: cabecera en dos filas. Combina cada fila candidata con la de
//     arriba usando el texto de la fila SUPERIOR si existe y, si no, el de la
//     inferior (p.ej. notificaciones: nombres A3 en la fila 1 y nombres A–E en
//     la fila 2). El `headerRow` devuelto es la fila inferior; los datos
//     empiezan en headerRow + 1.
//   - fuzzy: activa el nivel 2 (por prefijo) de resolveHeaderColumns para
//     tolerar variantes de nombre del cliente (plural, sufijos). Ver esa función.
// Devuelve { sheet, sheetName, headerRow, cols } o null si no la encuentra.
function locateHeaderTable(workbook, synonyms, isHeader, options = {}) {
  const { scanRows = HEADER_SCAN_ROWS, mergeUp = false, fuzzy = false } = options;
  for (const sheet of workbook.sheets()) {
    const { rows } = readAbsoluteRows(sheet);
    if (!rows.length) continue;
    const limit = rows[0].rowIndex + scanRows;
    for (let i = 0; i < rows.length; i++) {
      const { rowIndex, cells } = rows[i];
      if (rowIndex > limit) break;
      let header = cells;
      if (mergeUp && i > 0) {
        const above = rows[i - 1].cells || [];
        const width = Math.max(cells.length, above.length);
        header = [];
        for (let c = 0; c < width; c++) {
          const upper = above[c];
          header[c] =
            upper !== null && upper !== undefined && String(upper).trim() !== ""
              ? upper
              : cells[c];
        }
      }
      const { cols } = resolveHeaderColumns(header, synonyms, { fuzzy });
      if (isHeader(cols)) {
        return { sheet, sheetName: sheet.name(), headerRow: rowIndex, cols };
      }
    }
  }
  return null;
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
  _toDate,
  leerImporte,
  excelSerialToDate,
  fechaDesdeFormulario,
  fechaCorta,
  conFecha,
  LIMITE_DESC,
  LIMITE_DESC_AMPLIADA,
  recortarPorPalabra,
  repartirDescripcion,
  pad5,
  isoDate,
  stampYYYYMMDDHHmm,
  ensureDir,
  readAbsoluteRows,
  normalizeHeader,
  resolveHeaderColumns,
  locateHeaderTable,
  writeCsv,
};
