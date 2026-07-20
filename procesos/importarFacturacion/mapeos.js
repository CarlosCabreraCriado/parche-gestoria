const path = require("path");
const XlsxPopulate = require("xlsx-populate");
const { _str, _toInt, readAbsoluteRows, resolveHeaderColumns } = require("./utils");

const SHEET_CLIENTES_EXPTES = "ClientesXExptes";
const SHEET_CONCEPTOS_FACTURABLES = "ConceptosFacturables";
const SHEET_EMPRESAS_NO_FACTURABLES = "EmpresasNoFacturables";

const REDIRECT_NADA = "NADA";

// Las columnas se resuelven POR NOMBRE, nunca por posición. Leerlas por índice
// fijo hizo que añadir "Frecuencia" en la columna C de ConceptosFacturables
// —desplazando "Importe" a la D— dejara el catálogo entero sin precios: la
// corrida salía con 0 conceptos y las 945 filas en incidencias, sin un solo
// error que lo delatara. El primer sinónimo de cada campo es el nombre actual
// de la columna; los demás permiten renombrarlas sin tocar código.
const HEADER_SYNONYMS = {
  [SHEET_CLIENTES_EXPTES]: {
    cliente: ["expte", "cliente", "codigocliente"],
    expediente: ["exptefact", "expediente", "expedientefact"],
  },
  [SHEET_CONCEPTOS_FACTURABLES]: {
    codigo: ["codigo", "cod", "codconcepto"],
    descripcion: ["descripcion", "concepto", "nombre"],
    frecuencia: ["frecuencia", "periodicidad"],
    importe: ["importe", "precio", "tarifa"],
  },
  [SHEET_EMPRESAS_NO_FACTURABLES]: {
    origen: ["empresasenlasquenosefacturanada", "origen", "nofacturar"],
    destino: ["empresasenlasquefacturarlostramites", "destino", "facturaren"],
  },
};

// Campos sin los que la hoja no se puede interpretar. `frecuencia` y
// `descripcion` quedan fuera a propósito: son opcionales y su ausencia degrada
// (sin validación / sin nombre), no corrompe.
const HEADER_REQUIRED = {
  [SHEET_CLIENTES_EXPTES]: ["cliente", "expediente"],
  [SHEET_CONCEPTOS_FACTURABLES]: ["codigo", "importe"],
  [SHEET_EMPRESAS_NO_FACTURABLES]: ["origen", "destino"],
};

const HEADER_SCAN_ROWS = 5;

function readSheetRows(workbook, sheetName) {
  const sheet = workbook.sheet(sheetName);
  if (!sheet) {
    throw new Error(
      `Hoja '${sheetName}' no encontrada en el archivo de mapeos. Hojas requeridas: ${SHEET_CLIENTES_EXPTES}, ${SHEET_CONCEPTOS_FACTURABLES}, ${SHEET_EMPRESAS_NO_FACTURABLES}.`
    );
  }
  return readAbsoluteRows(sheet).rows;
}

// Devuelve { rows, headerRow, cols }. Si falta una columna obligatoria aborta
// nombrando la hoja y lo que se esperaba: fallar claro es mejor que leer de la
// columna de al lado y facturar mal.
function readSheetTable(workbook, sheetName) {
  const rows = readSheetRows(workbook, sheetName);
  const synonyms = HEADER_SYNONYMS[sheetName];
  const required = HEADER_REQUIRED[sheetName];

  for (const { rowIndex, cells } of rows) {
    if (rowIndex > HEADER_SCAN_ROWS) break;
    const { cols } = resolveHeaderColumns(cells, synonyms);
    if (required.every((f) => cols[f] !== undefined)) {
      return { rows, headerRow: rowIndex, cols };
    }
  }

  const esperadas = required.map((f) => `'${synonyms[f][0]}'`).join(", ");
  throw new Error(
    `Hoja '${sheetName}' del archivo de mapeos: no se encuentra la fila de cabecera ` +
      `(se requieren las columnas ${esperadas} en las primeras ${HEADER_SCAN_ROWS} filas). ` +
      `Revisa que no se hayan renombrado ni borrado.`
  );
}

class ExpteShortLookup {
  constructor() {
    this._byCliente = new Map();
    this.warnings = [];
  }

  static fromWorkbook(workbook) {
    const obj = new ExpteShortLookup();
    const { rows, headerRow, cols } = readSheetTable(workbook, SHEET_CLIENTES_EXPTES);
    for (const { rowIndex, cells } of rows) {
      if (rowIndex <= headerRow) continue;
      const raw = cells ? cells[cols.cliente] : undefined;
      if (raw === undefined || raw === null) continue;
      const key = _toInt(raw);
      if (key === null) {
        obj.warnings.push(
          `Fila ${rowIndex}: código cliente no numérico '${raw}' — ignorado`
        );
        continue;
      }
      const val = _str(cells[cols.expediente]);
      if (!val) {
        obj.warnings.push(
          `Fila ${rowIndex}: cliente ${key} sin expediente — ignorado`
        );
        continue;
      }
      const prev = obj._byCliente.get(key);
      if (prev !== undefined && prev !== val) {
        obj.warnings.push(
          `Fila ${rowIndex}: cliente ${key} ya mapeado a '${prev}', sobrescrito con '${val}'`
        );
      }
      obj._byCliente.set(key, val);
    }
    return obj;
  }

  resolve(clienteCorto) {
    const k = _toInt(clienteCorto);
    if (k === null) return null;
    return this._byCliente.get(k) ?? null;
  }

  size() {
    return this._byCliente.size;
  }
}

class TarifaCatalog {
  constructor() {
    this._prices = new Map();
    this._names = new Map();
    this._escalado = new Set();
    this._sinPrecio = new Set();
    this._frecuencias = new Map();
    this.warnings = [];
  }

  static fromWorkbook(workbook) {
    const obj = new TarifaCatalog();
    const { rows, headerRow, cols } = readSheetTable(workbook, SHEET_CONCEPTOS_FACTURABLES);
    const seen = new Map();
    for (const { rowIndex, cells } of rows) {
      if (rowIndex <= headerRow) continue;
      if (!cells || cells[cols.codigo] === undefined || cells[cols.codigo] === null) continue;
      const code = _str(cells[cols.codigo]);
      if (!code) continue;
      if (seen.has(code)) {
        obj.warnings.push(
          `Fila ${rowIndex}: código '${code}' duplicado (previo en fila ${seen.get(code)}) — usado el último`
        );
      }
      seen.set(code, rowIndex);
      obj._names.set(code, cols.descripcion !== undefined ? _str(cells[cols.descripcion]) : "");

      // Frecuencias declaradas para el concepto. Admite varias separadas por
      // "/" ("MENSUAL/TRIMESTRAL"): el 111 lo presentan mensual las grandes
      // empresas y trimestral el resto, así que un concepto puede tener más de
      // una legítimamente. Solo se usa para avisar de discrepancias, nunca para
      // decidir si una fila se factura — eso lo fija la FRECUENCIA de la fila.
      if (cols.frecuencia !== undefined) {
        const frecs = _str(cells[cols.frecuencia])
          .toUpperCase()
          .split(/[\/,;]/)
          .map((f) => f.trim())
          .filter(Boolean);
        if (frecs.length) obj._frecuencias.set(code, new Set(frecs));
        else obj._frecuencias.delete(code);
      }

      const priceRaw = cells[cols.importe] ?? null;
      obj._prices.delete(code);
      obj._escalado.delete(code);
      obj._sinPrecio.delete(code);
      if (priceRaw === null || priceRaw === undefined) {
        obj._sinPrecio.add(code);
      } else if (typeof priceRaw === "string") {
        const txt = priceRaw.trim();
        if (txt === "") {
          obj._sinPrecio.add(code);
        } else if (txt.toUpperCase() === "ESCALADO") {
          obj._escalado.add(code);
        } else {
          const n = Number(txt.replace(",", "."));
          if (!Number.isFinite(n)) {
            obj.warnings.push(
              `Fila ${rowIndex}: código '${code}' con precio no numérico '${priceRaw}' — tratado como sin precio`
            );
            obj._sinPrecio.add(code);
          } else {
            obj._prices.set(code, n);
          }
        }
      } else if (typeof priceRaw === "number" && Number.isFinite(priceRaw)) {
        obj._prices.set(code, priceRaw);
      } else {
        obj._sinPrecio.add(code);
      }
    }
    return obj;
  }

  resolve(codigo) {
    const key = String(codigo ?? "").trim();
    return this._prices.get(key) ?? null;
  }

  // Descripción del catálogo ("Modelo 111 -"). Cadena vacía si no está: quien la
  // use decide el texto alternativo.
  describe(codigo) {
    const key = String(codigo ?? "").trim();
    return this._names.get(key) ?? "";
  }

  missReason(codigo) {
    const key = String(codigo ?? "").trim();
    if (this._prices.has(key)) return "ok";
    if (this._escalado.has(key)) return "escalado";
    if (this._sinPrecio.has(key)) return "sin_precio";
    return "no_en_catalogo";
  }

  known(codigo) {
    const key = String(codigo ?? "").trim();
    return this._prices.has(key) || this._escalado.has(key) || this._sinPrecio.has(key);
  }

  // Frecuencias declaradas en el catálogo, o null si el concepto no la declara
  // (hoy la inmensa mayoría). Informativa: solo alimenta el aviso de
  // discrepancia contra la FRECUENCIA de la fila, que es la que manda.
  frecuencias(codigo) {
    const key = String(codigo ?? "").trim();
    return this._frecuencias.get(key) ?? null;
  }

  size() {
    return this._prices.size;
  }
}

class ClienteRedirect {
  constructor() {
    this._byOrigen = new Map();
    this.warnings = [];
  }

  static fromWorkbook(workbook) {
    const obj = new ClienteRedirect();
    const { rows, headerRow, cols } = readSheetTable(workbook, SHEET_EMPRESAS_NO_FACTURABLES);
    for (const { rowIndex, cells } of rows) {
      if (rowIndex <= headerRow) continue;
      if (!cells) continue;
      const origenRaw = cells[cols.origen];
      const origen = _toInt(origenRaw);
      if (origen === null) {
        if (origenRaw !== null && origenRaw !== undefined && origenRaw !== "") {
          obj.warnings.push(
            `Fila ${rowIndex}: origen no numérico '${origenRaw}' — ignorado`
          );
        }
        continue;
      }
      const destinoRaw = cells[cols.destino];
      let destino;
      if (typeof destinoRaw === "string" && destinoRaw.trim().toUpperCase() === REDIRECT_NADA) {
        destino = REDIRECT_NADA;
      } else if (destinoRaw === null || destinoRaw === undefined || destinoRaw === "") {
        obj.warnings.push(`Fila ${rowIndex}: origen ${origen} sin destino — ignorado`);
        continue;
      } else {
        const destInt = _toInt(destinoRaw);
        if (destInt === null) {
          obj.warnings.push(
            `Fila ${rowIndex}: destino '${destinoRaw}' no interpretable — ignorado`
          );
          continue;
        }
        destino = destInt;
      }
      const prev = obj._byOrigen.get(origen);
      if (prev !== undefined && prev !== destino) {
        obj.warnings.push(
          `Fila ${rowIndex}: origen ${origen} ya mapeado a ${JSON.stringify(prev)}, sobrescrito con ${JSON.stringify(destino)}`
        );
      }
      obj._byOrigen.set(origen, destino);
    }
    return obj;
  }

  // int destino si hay redirección, 'NADA' si no facturable, null si no aplica.
  resolve(clienteOrigen) {
    const k = _toInt(clienteOrigen);
    if (k === null) return null;
    return this._byOrigen.get(k) ?? null;
  }

  size() {
    return this._byOrigen.size;
  }
}

class Mapeos {
  constructor(exptes, tarifas, redirect) {
    this.exptes = exptes;
    this.tarifas = tarifas;
    this.redirect = redirect;
  }

  static async fromFile(filePath) {
    const workbook = await XlsxPopulate.fromFileAsync(path.normalize(filePath));
    const exptes = ExpteShortLookup.fromWorkbook(workbook);
    const tarifas = TarifaCatalog.fromWorkbook(workbook);
    const redirect = ClienteRedirect.fromWorkbook(workbook);
    return new Mapeos(exptes, tarifas, redirect);
  }

  allWarnings() {
    return [
      ...this.exptes.warnings.map((w) => `[exptes] ${w}`),
      ...this.tarifas.warnings.map((w) => `[tarifas] ${w}`),
      ...this.redirect.warnings.map((w) => `[redirect] ${w}`),
    ];
  }

  summary() {
    return {
      exptes_cargados: this.exptes.size(),
      tarifas_cargadas: this.tarifas.size(),
      tarifas_escaladas: this.tarifas._escalado.size,
      tarifas_sin_precio: this.tarifas._sinPrecio.size,
      tarifas_con_frecuencia: this.tarifas._frecuencias.size,
      redirecciones: this.redirect.size(),
      warnings: this.allWarnings().length,
    };
  }
}

module.exports = {
  ExpteShortLookup,
  TarifaCatalog,
  ClienteRedirect,
  Mapeos,
  REDIRECT_NADA,
  SHEET_CLIENTES_EXPTES,
  SHEET_CONCEPTOS_FACTURABLES,
  SHEET_EMPRESAS_NO_FACTURABLES,
};
