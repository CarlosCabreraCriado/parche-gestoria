const path = require("path");
const XlsxPopulate = require("xlsx-populate");
const { _str, _toInt, readAbsoluteRows } = require("./utils");

const SHEET_CLIENTES_EXPTES = "ClientesXExptes";
const SHEET_CONCEPTOS_FACTURABLES = "ConceptosFacturables";
const SHEET_EMPRESAS_NO_FACTURABLES = "EmpresasNoFacturables";

const REDIRECT_NADA = "NADA";

function readSheetRows(workbook, sheetName) {
  const sheet = workbook.sheet(sheetName);
  if (!sheet) {
    throw new Error(
      `Hoja '${sheetName}' no encontrada en el archivo de mapeos. Hojas requeridas: ${SHEET_CLIENTES_EXPTES}, ${SHEET_CONCEPTOS_FACTURABLES}, ${SHEET_EMPRESAS_NO_FACTURABLES}.`
    );
  }
  return readAbsoluteRows(sheet).rows;
}

class ExpteShortLookup {
  constructor() {
    this._byCliente = new Map();
    this.warnings = [];
  }

  static fromWorkbook(workbook) {
    const obj = new ExpteShortLookup();
    const rows = readSheetRows(workbook, SHEET_CLIENTES_EXPTES);
    for (const { rowIndex, cells } of rows) {
      if (rowIndex < 2) continue; // saltar cabecera
      if (!cells || cells[0] === undefined || cells[0] === null) continue;
      const key = _toInt(cells[0]);
      if (key === null) {
        obj.warnings.push(
          `Fila ${rowIndex}: código cliente no numérico '${cells[0]}' — ignorado`
        );
        continue;
      }
      const val = _str(cells[1]);
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
    this.warnings = [];
  }

  static fromWorkbook(workbook) {
    const obj = new TarifaCatalog();
    const rows = readSheetRows(workbook, SHEET_CONCEPTOS_FACTURABLES);
    const seen = new Map();
    for (const { rowIndex, cells } of rows) {
      if (rowIndex < 2) continue;
      if (!cells || cells[0] === undefined || cells[0] === null) continue;
      const code = _str(cells[0]);
      if (!code) continue;
      if (seen.has(code)) {
        obj.warnings.push(
          `Fila ${rowIndex}: código '${code}' duplicado (previo en fila ${seen.get(code)}) — usado el último`
        );
      }
      seen.set(code, rowIndex);
      obj._names.set(code, _str(cells[1]));
      const priceRaw = cells.length > 2 ? cells[2] : null;
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
    const rows = readSheetRows(workbook, SHEET_EMPRESAS_NO_FACTURABLES);
    for (const { rowIndex, cells } of rows) {
      if (rowIndex < 2) continue;
      if (!cells || cells.length < 2) continue;
      const origen = _toInt(cells[0]);
      if (origen === null) {
        if (cells[0] !== null && cells[0] !== undefined && cells[0] !== "") {
          obj.warnings.push(
            `Fila ${rowIndex}: origen no numérico '${cells[0]}' — ignorado`
          );
        }
        continue;
      }
      const destinoRaw = cells[1];
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
