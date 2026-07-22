const path = require("path");
const XlsxPopulate = require("xlsx-populate");
const {
  _str,
  _toInt,
  normalizeHeader,
  readAbsoluteRows,
  resolveHeaderColumns,
} = require("./utils");

const SHEET_CLIENTES_EXPTES = "ClientesXExptes";
const SHEET_CONCEPTOS_FACTURABLES = "ConceptosFacturables";
const SHEET_EMPRESAS_NO_FACTURABLES = "EmpresasNoFacturables";

// Escalas de precio por tramos. Viven en su propia hoja y no en
// ConceptosFacturables porque son otra entidad: una tabla rango→precio, no un
// precio. Esa hoja es además la lista con la que la gestoría cotiza al cliente,
// y trae DOS escalas sobre LOS MISMOS tramos (el modelo 190 y su certificado de
// retenciones), una por columna. Mantenerlas juntas es justo lo que impide que
// los tramos deriven entre ambas: si el año que viene "6 A 25" pasa a "6 A 30",
// se toca una celda y vale para las dos.
const SHEET_ESCALAS = "Modelo 190";

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

// "0 A 2" → [0,2] · "250 en adelante" → [250,∞) · "7" → [7,7]. Devuelve null si
// la fila no describe un tramo (títulos, notas, filas vacías de la hoja).
function parseTramo(texto) {
  const s = _str(texto).replace(/\s+/g, " ").trim();
  if (!s) return null;
  let m = s.match(/^(\d+)\s*(?:A|-|hasta)\s*(\d+)$/i);
  if (m) return { desde: Number(m[1]), hasta: Number(m[2]) };
  m = s.match(/^(\d+)\s*(?:en adelante|o m[aá]s|y m[aá]s|\+)$/i);
  if (m) return { desde: Number(m[1]), hasta: Infinity };
  m = s.match(/^(\d+)$/);
  if (m) return { desde: Number(m[1]), hasta: Number(m[1]) };
  return null;
}

function precioDeCelda(raw) {
  if (raw === null || raw === undefined) return null;
  if (typeof raw === "number") return Number.isFinite(raw) ? raw : null;
  const s = _str(raw).replace(/[€\s]/g, "");
  if (!s) return null;
  const n = Number(s.includes(",") ? s.replace(/\./g, "").replace(",", ".") : s);
  return Number.isFinite(n) ? n : null;
}

// Tabla tramo→precio de la hoja `Modelo 190`. Una escala por columna de precio;
// la clave es el prefijo normalizado de su cabecera ("MODELO 190", "CERTIFICADOS"),
// no la cabecera entera: así el año del título puede pasar a 2027 sin romper nada.
//
// La hoja es OPCIONAL: los otros importadores (nóminas, trámites,
// notificaciones) comparten este archivo de mapeos y no usan escalas, así que
// su ausencia no puede impedir cargarlo. Quien necesite una escala y no la
// encuentre falla en su sitio, con el nombre de la que buscaba.
class EscalaCatalog {
  constructor() {
    this._escalas = new Map(); // clave normalizada -> { nombre, tramos: [...] }
    this.warnings = [];
  }

  static fromWorkbook(workbook) {
    const obj = new EscalaCatalog();
    const sheet = workbook.sheet(SHEET_ESCALAS);
    if (!sheet) return obj;

    const rows = readAbsoluteRows(sheet).rows;

    // La cabecera es la última fila con contenido antes del primer tramo. Se
    // localiza por contenido y no por número de fila porque la hoja lleva un
    // título encima ("MODELO 190-CERTIFICADOS RETENCION") que podría ganar o
    // perder líneas.
    const primerTramo = rows.find((r) => parseTramo((r.cells || [])[0]) !== null);
    if (!primerTramo) {
      obj.warnings.push(
        `Hoja '${SHEET_ESCALAS}': no se encuentra ninguna fila de tramos ("0 A 2", "250 en adelante"…) — no se carga ninguna escala.`
      );
      return obj;
    }
    const filaCabecera = [...rows]
      .reverse()
      .find((r) => r.rowIndex < primerTramo.rowIndex && (r.cells || []).some((c) => _str(c) !== ""));
    if (!filaCabecera) {
      obj.warnings.push(
        `Hoja '${SHEET_ESCALAS}': los tramos empiezan en la fila ${primerTramo.rowIndex} y no hay ninguna fila de cabecera encima que nombre las columnas de precio.`
      );
      return obj;
    }

    // Toda columna de la cabecera salvo la primera (que son los tramos) es una
    // escala. Así añadir una tercera columna de precio a la hoja no exige tocar
    // código: basta con que quien la use la nombre.
    const headerRow = filaCabecera.rowIndex;
    const cabecera = filaCabecera.cells || [];
    const columnas = [];
    cabecera.forEach((celda, idx) => {
      if (idx === 0) return;
      const nombre = _str(celda);
      const clave = normalizeHeader(celda);
      if (!clave) return;
      columnas.push({ idx, nombre, clave });
      obj._escalas.set(clave, { nombre, tramos: [] });
    });

    for (const { rowIndex, cells } of rows) {
      if (rowIndex <= headerRow) continue;
      const tramo = parseTramo((cells || [])[0]);
      if (tramo === null) continue;
      const etiqueta = _str((cells || [])[0]).replace(/\s+/g, " ");
      for (const col of columnas) {
        const precio = precioDeCelda((cells || [])[col.idx]);
        if (precio === null) {
          obj.warnings.push(
            `Hoja '${SHEET_ESCALAS}' fila ${rowIndex}: el tramo '${etiqueta}' no tiene precio en la columna '${col.nombre}' — ese tramo queda sin cubrir.`
          );
          continue;
        }
        obj._escalas.get(col.clave).tramos.push({ ...tramo, etiqueta, precio, fila: rowIndex });
      }
    }

    for (const escala of obj._escalas.values()) obj._validar(escala);
    return obj;
  }

  // Solapes, huecos y falta de tramo abierto al final. Son avisos y no errores:
  // `resolve` es determinista igualmente (gana el primer tramo que casa, en el
  // orden de la hoja), así que una corrida no se bloquea por esto. Pero cada uno
  // significa que alguien cobraría de más o de menos, así que tienen que verse.
  _validar(escala) {
    const t = escala.tramos;
    if (!t.length) {
      this.warnings.push(`Escala '${escala.nombre}': sin ningún tramo con precio.`);
      return;
    }
    for (let i = 0; i < t.length; i++) {
      if (t[i].desde > t[i].hasta) {
        this.warnings.push(
          `Escala '${escala.nombre}', fila ${t[i].fila}: el tramo '${t[i].etiqueta}' está al revés (${t[i].desde} > ${t[i].hasta}).`
        );
      }
      if (i === 0) continue;
      const prev = t[i - 1];
      if (t[i].desde <= prev.hasta) {
        this.warnings.push(
          `Escala '${escala.nombre}': los tramos '${prev.etiqueta}' (${prev.precio}) y '${t[i].etiqueta}' (${t[i].precio}) se solapan ` +
            `en ${t[i].desde}${prev.hasta === Infinity ? " en adelante" : `–${Math.min(prev.hasta, t[i].hasta)}`} — ` +
            `se aplica el primero (${prev.precio}). Corrígelo en la hoja '${SHEET_ESCALAS}'.`
        );
      } else if (t[i].desde > prev.hasta + 1) {
        this.warnings.push(
          `Escala '${escala.nombre}': hueco entre '${prev.etiqueta}' y '${t[i].etiqueta}' — ` +
            `los valores ${prev.hasta + 1}–${t[i].desde - 1} no tienen precio.`
        );
      }
    }
    const ultimo = t[t.length - 1];
    if (ultimo.hasta !== Infinity) {
      this.warnings.push(
        `Escala '${escala.nombre}': el último tramo ('${ultimo.etiqueta}') está cerrado en ${ultimo.hasta} — ` +
          `por encima de ese valor no hay precio. Usa "N en adelante".`
      );
    }
    const primero = t[0];
    if (primero.desde > 0) {
      this.warnings.push(
        `Escala '${escala.nombre}': el primer tramo empieza en ${primero.desde} — ` +
          `por debajo de ese valor no hay precio.`
      );
    }
  }

  _get(escala) {
    return this._escalas.get(normalizeHeader(escala)) ?? this._porPrefijo(escala);
  }

  // La cabecera real lleva el año ("MODELO 190 2026"), así que se busca por
  // prefijo: quien pide "MODELO 190" sigue encontrándola en 2027.
  _porPrefijo(escala) {
    const clave = normalizeHeader(escala);
    if (!clave) return undefined;
    for (const [k, v] of this._escalas) {
      if (k.startsWith(clave)) return v;
    }
    return undefined;
  }

  has(escala) {
    return this._get(escala) !== undefined;
  }

  // { precio, tramo } o null si la escala no existe o la cantidad no cae en
  // ningún tramo. Gana el primer tramo que casa, en el orden de la hoja.
  resolve(escala, cantidad) {
    const e = this._get(escala);
    if (!e || !Number.isFinite(cantidad)) return null;
    const t = e.tramos.find((x) => cantidad >= x.desde && cantidad <= x.hasta);
    return t ? { precio: t.precio, tramo: t.etiqueta } : null;
  }

  missReason(escala, cantidad) {
    const e = this._get(escala);
    if (!e) {
      const disponibles = [...this._escalas.values()].map((x) => `'${x.nombre}'`).join(", ");
      return (
        `no existe la escala '${escala}' en la hoja '${SHEET_ESCALAS}' del archivo de mapeos` +
        (disponibles ? ` (hay: ${disponibles})` : " (la hoja no existe o está vacía)")
      );
    }
    if (!Number.isFinite(cantidad)) return `cantidad '${cantidad}' no numérica`;
    return `${cantidad} no cae en ningún tramo de '${e.nombre}'`;
  }

  nombres() {
    return [...this._escalas.values()].map((e) => e.nombre);
  }

  size() {
    return this._escalas.size;
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
  constructor(exptes, tarifas, redirect, escalas) {
    this.exptes = exptes;
    this.tarifas = tarifas;
    this.redirect = redirect;
    this.escalas = escalas;
  }

  static async fromFile(filePath) {
    const workbook = await XlsxPopulate.fromFileAsync(path.normalize(filePath));
    const exptes = ExpteShortLookup.fromWorkbook(workbook);
    const tarifas = TarifaCatalog.fromWorkbook(workbook);
    const redirect = ClienteRedirect.fromWorkbook(workbook);
    const escalas = EscalaCatalog.fromWorkbook(workbook);
    return new Mapeos(exptes, tarifas, redirect, escalas);
  }

  allWarnings() {
    return [
      ...this.exptes.warnings.map((w) => `[exptes] ${w}`),
      ...this.tarifas.warnings.map((w) => `[tarifas] ${w}`),
      ...this.redirect.warnings.map((w) => `[redirect] ${w}`),
      ...this.escalas.warnings.map((w) => `[escalas] ${w}`),
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
      escalas_cargadas: this.escalas.size(),
      warnings: this.allWarnings().length,
    };
  }
}

module.exports = {
  ExpteShortLookup,
  TarifaCatalog,
  ClienteRedirect,
  EscalaCatalog,
  Mapeos,
  parseTramo,
  REDIRECT_NADA,
  SHEET_CLIENTES_EXPTES,
  SHEET_CONCEPTOS_FACTURABLES,
  SHEET_EMPRESAS_NO_FACTURABLES,
  SHEET_ESCALAS,
};
