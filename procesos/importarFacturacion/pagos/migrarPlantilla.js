// Migración: 4PAGOS<año>.xls → plantilla normalizada A3 (.xlsx).
//
// Genera un libro nuevo donde cada hoja de modelo fiscal tiene:
//   - Zona A (cols A–F): bloque estándar idéntico en todas las hojas
//     CONCEPTO FACT | EXPTE | NIF | EMPRESA | FACTURAR | FRECUENCIA
//   - Col G: separador
//   - Zona B (col H+): columnas originales no consumidas (incl. P1–P4,
//     OBSERVACIONES y F.BAJA), copiadas tal cual
// CONCEPTO FACT es el concepto facturable de la fila (código de la hoja
// `ConceptosFacturables` de mapeos_facturacion.xlsx: 0.016, 0.012…). Se deriva
// del modelo fiscal vía `MODELO_CONCEPTO` y es lo que fija el importe a facturar,
// no el importe que hubiera en el periodo. Por eso P1–P4 (y OBSERVACIONES) dejan
// de ser campos que lee el importador y vuelven a la zona libre como dato crudo
// original. Si un modelo no está mapeado, CONCEPTO FACT cae al propio nº de
// modelo (señal visible de que falta mapearlo).
// OBSERVACIONES / P1–P4 se copian tal cual del original a la zona libre.
// FRECUENCIA (TRIMESTRAL/MENSUAL/ANUAL/OTRA) hace explícita la periodicidad de
// cada fila. Por defecto la hereda la hoja o la sección (igual que FACTURAR),
// pero en la hoja 111 algunas empresas grandes declaran mensualmente y hoy esa
// excepción vivía escondida como texto "MENSUAL" en la columna F.BAJA — ver
// `frecuenciaDesdeColumna` en su spec, que la promueve a este campo explícito
// sin tocar F.BAJA (que se sigue copiando tal cual a Zona B).
// FACTURAR es la única fuente de verdad sobre si una fila se factura: no hay
// ningún otro campo (p. ej. F.BAJA) que el importador cruce o use de fallback.
// Las hojas sin estructura tabular (alcohol, 184, EXP) se copian literalmente.
// Las secciones internas del original (BAJAS, EXENTOS, MOROSOS, sub-bloques de
// otro modelo…) se localizan por el TEXTO de su título (`match` en el spec).
// El invariante que lo sustenta: una fila de sección nunca lleva expte numérico.
// Si un título no aparece, o aparece más de una vez, el script falla: el archivo
// del cliente cambia cada trimestre y un ancla que se desplaza sin avisar se
// come filas de datos (las convierte en separador) y deja de facturarlas.
// Quedan specs con anclas legacy por número de fila exacto (`fila`), calibradas
// contra 4PAGOS2026 (1).xls; se validan igual contra ese invariante y hay que
// migrarlas a `match` antes de procesar esas hojas con otro 4PAGOS.
//
// Uso: node migrarPlantilla.js [input.xls] [output.xlsx] [--hojas 111,130]

const path = require("path");
const fs = require("fs");
const XLSX = require("xlsx");
const XlsxPopulate = require("xlsx-populate");

const INPUTS_DIR = path.join(__dirname, "..", "inputs", "pagos");
const DEFAULT_INPUT = path.join(INPUTS_DIR, "4PAGOS2026_2T.xls");
const DEFAULT_OUTPUT = path.join(INPUTS_DIR, "4PAGOS2026_2T - PLANTILLA A3 v2.xlsx");

// v2 añadió la columna IMPORTE (precio puntual por fila). El importador acepta
// v1 y v2; ver SENTINEL_VERSION en pagos.js.
const SENTINEL = "A3PAGOS v2";
const ZONA_A = ["CONCEPTO FACT", "EXPTE", "NIF", "EMPRESA", "FACTURAR", "FRECUENCIA", "IMPORTE"];
const FRECUENCIAS = ["TRIMESTRAL", "MENSUAL", "ANUAL", "OTRA"];
const COL_IMPORTE = 7; // dentro de Zona A; se deja vacía (opt-in del usuario)
const COL_SEP = ZONA_A.length + 1; // separador tras Zona A
const ZONA_B_START = COL_SEP + 1; // primera columna de Zona B

// Modelo fiscal → código de concepto facturable (hoja `ConceptosFacturables` de
// mapeos_facturacion.xlsx). Todos son "Modelo N" en esa tabla; el importe se
// resuelve luego cruzando el código, no desde aquí.
const MODELO_CONCEPTO = {
  "130": "0.016", "131": "0.017", "111": "0.012", "115": "0.013", "046": "0.115",
  "420": "0.014", "421": "0.015", "417": "0.025", "412": "0.096",
  "202": "0.018", "210": "0.019",
  "303": "3.186", "349": "3.248", "369": "0.121", "340": "3.183",
  "123": "3.185", "193": "3.184",
};
function conceptoDe(modelo) {
  return MODELO_CONCEPTO[modelo] || modelo; // fallback: nº de modelo si no está mapeado
}

// Colores
const C_HEADER_A = "305496"; // azul oscuro
const C_HEADER_B = "808080"; // gris
const C_SECCION = "FCE4D6"; // naranja claro
const C_NOTA = "F2F2F2"; // gris claro
const C_SEP = "404040"; // separador

// ---------------------------------------------------------------- helpers

function toInt(v) {
  if (v === null || v === undefined || v === "") return null;
  const n = Number(String(v).trim());
  return Number.isFinite(n) ? Math.trunc(n) : null;
}
function str(v) {
  if (v === null || v === undefined) return "";
  return String(v).trim();
}
function isDateSerial(v) {
  return typeof v === "number" && Number.isFinite(v) && v >= 25000 && v <= 60000;
}
const DATE_HEADER_RX = /alta|baja|fecha|constit|oblig|vto|^o$/i;

function colLetter(n0) {
  // índice 0-based → letra Excel
  let s = "";
  let n = n0 + 1;
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

// ---------------------------------------------------------------- specs
//
// cols: índices 0-based del archivo ORIGINAL.
// secciones: anclas. La fila ancla nunca es dato. Dos formas:
//   { match: /regex/i, titulo, facturar, modelo?, frecuencia?, skip? }
//       localiza la sección por el texto de su título — resiste que el archivo
//       gane o pierda filas entre trimestres. Preferir siempre esta.
//   { fila, titulo, ... }
//       legacy: número de fila exacto (1-based) de 4PAGOS2026 (1).xls. Solo vale
//       para ese archivo; migrar a `match` al procesar la hoja con otro 4PAGOS.
//   skip=true: cabecera repetida, se descarta sin emitir separador.
// El estado inicial (antes de la primera ancla) es `inicial`.

const SPECS = [
  {
    hoja: "130",
    headerRow: 1,
    labelRows: [1],
    dataFrom: 2,
    cols: { expte: 2, nif: 5, empresa: 6, p: [11, 12, 13, 14], obs: 16 },
    inicial: { modelo: "130", facturar: "SI", frecuencia: "TRIMESTRAL", titulo: "MODELO 130 TRIMESTRAL" },
    secciones: [
      { fila: 78, titulo: "BAJAS", facturar: "NO" },
      { fila: 82, titulo: "EXENTOS", facturar: "NO" },
      { fila: 112, titulo: "MOROSOS O SE FUERON — COMPROBAR", facturar: "REVISAR" },
    ],
  },
  {
    hoja: "420",
    headerRow: 1,
    labelRows: [1],
    dataFrom: 2,
    cols: { expte: 2, nif: 5, empresa: 6, p: [11, 12, 13, 14], obs: 16 },
    inicial: { modelo: "417", facturar: "REVISAR", frecuencia: "MENSUAL", titulo: "MODELO 417 GRAN EMPRESA SII (mensual)" },
    secciones: [
      { fila: 8, titulo: "MODELO 412 MENSUAL", modelo: "412", facturar: "REVISAR", frecuencia: "MENSUAL" },
      { fila: 11, titulo: "MODELO 420 TRIMESTRAL", modelo: "420", facturar: "SI", frecuencia: "TRIMESTRAL" },
      { fila: 169, titulo: "BAJAS", facturar: "NO" },
      { fila: 185, titulo: "REGIMEN PEQUEÑO EMPRESARIO — REPEP (exento trimestral)", facturar: "NO" },
      { fila: 241, titulo: "MINORISTAS (exentos)", facturar: "NO" },
      { fila: 261, titulo: "EPIGRAFES EXENTOS ESPECIALES", facturar: "NO" },
      { fila: 284, titulo: "SE ENCARGAN ELLOS", facturar: "NO" },
      { fila: 287, titulo: "MOROSOS — COMPROBAR", facturar: "REVISAR" },
      { fila: 290, titulo: "NULAS O SE FUERON O LO TRAMITAN ELLOS", facturar: "NO" },
    ],
  },
  {
    hoja: "303-349-369-340",
    headerRow: 2,
    labelRows: [2, 1],
    dataFrom: 3,
    cols: { expte: 2, nif: 5, empresa: 6, p: [11, 12, 13, 14], obs: 16 },
    inicial: { modelo: "303", facturar: "SI", frecuencia: "TRIMESTRAL", titulo: "MODELO 303 TRIMESTRAL" },
    secciones: [
      { fila: 8, titulo: "BAJAS", facturar: "NO" },
      { fila: 16, titulo: "MODELO 303 MENSUAL", modelo: "303", facturar: "REVISAR", frecuencia: "MENSUAL" },
      { fila: 17, skip: true },
      { fila: 21, titulo: "MODELO 349 INTERCOMUNITARIO", modelo: "349", facturar: "SI", frecuencia: "TRIMESTRAL" },
      { fila: 22, skip: true },
      { fila: 26, titulo: "BAJAS (349)", facturar: "NO" },
      { fila: 31, titulo: "MODELO 369", modelo: "369", facturar: "SI", frecuencia: "TRIMESTRAL" },
      { fila: 32, skip: true },
      { fila: 38, titulo: "MODELO 340 REDEME (mensual)", modelo: "340", facturar: "REVISAR", frecuencia: "MENSUAL" },
      { fila: 39, skip: true },
    ],
  },
  {
    hoja: "123-193",
    headerRow: 2,
    labelRows: [2, 1],
    dataFrom: 3,
    cols: { expte: 2, nif: 5, empresa: 6, p: [11, 12, 13, 14], obs: 16 },
    inicial: { modelo: "123", facturar: "SI", frecuencia: "TRIMESTRAL", titulo: "MODELO 123 TRIMESTRAL" },
    secciones: [{ fila: 29, titulo: "SE FUERON ASESORIA", facturar: "NO" }],
  },
  {
    hoja: "202",
    headerRow: 2,
    labelRows: [2, 3, 1],
    dataFrom: 4,
    cols: { expte: 2, nif: 5, empresa: 6, p: [11, 12, 13], obs: null },
    inicial: { modelo: "202", facturar: "SI", frecuencia: "TRIMESTRAL", titulo: "MODELO 202 PAGOS FRACCIONADOS IS (P1=abr, P2=oct, P3=dic)" },
    secciones: [
      { fila: 91, titulo: "NO TIENE OBLIGACION MOD202", facturar: "NO" },
      { fila: 153, titulo: "LO HACEN ELLOS O SE FUERON", facturar: "NO" },
      { fila: 176, titulo: "SOCIEDADES DISUELTAS", facturar: "NO" },
    ],
  },
  {
    hoja: "210",
    headerRow: 2,
    labelRows: [2, 1],
    dataFrom: 3,
    cols: { expte: 2, nif: 5, empresa: 6, p: [8, 9, 10, 11], obs: 12 },
    inicial: { modelo: "210", facturar: "SI", frecuencia: "TRIMESTRAL", titulo: "MODELO 210 NO RESIDENTES" },
    secciones: [],
  },
  {
    hoja: "111",
    headerRow: 2,
    labelRows: [2, 1],
    dataFrom: 3,
    cols: { expte: 1, nif: 4, empresa: 5, p: [12, 13, 14, 15], obs: 17 },
    drop: [18, 19, 20, 21, 22, 23, 24, 25, 26, 27], // bloque A3 manual (ejemplo hecho a mano)
    // Algunas empresas grandes declaran el 111 mensualmente. Hoy esa excepción
    // vive como texto "MENSUAL" en la columna BAJA (col. original 8) en vez de
    // en un campo propio — la promovemos a FRECUENCIA sin tocar F.BAJA.
    frecuenciaDesdeColumna: { col: 8, valores: { MENSUAL: "MENSUAL" } },
    inicial: { modelo: "111", facturar: "SI", frecuencia: "TRIMESTRAL", titulo: "MODELO 111 TRIMESTRAL" },
    secciones: [
      { match: /BAJA\s+ASESORIA/i, titulo: "BAJA ASESORIA 2026", facturar: "NO" },
      { match: /EMPRESAS\s+REALIZAN\s+ELLOS/i, titulo: "EMPRESAS REALIZAN ELLOS MODELOS 111 - 190", facturar: "NO" },
    ],
  },
  {
    hoja: "131",
    headerRow: 2,
    labelRows: [2, 1],
    dataFrom: 3,
    cols: { expte: 3, nif: 6, empresa: 7, p: [13, 16, 20, 24], obs: 29 },
    labelOverrides: { 0: "(marca NO)" },
    inicial: { modelo: "131", facturar: "SI", frecuencia: "TRIMESTRAL", titulo: "MODELO 131 TRIMESTRAL" },
    secciones: [
      { fila: 13, titulo: "BAJAS", facturar: "NO" },
      { fila: 15, titulo: "EXENTOS DE PAGOS TRIMESTRALES — VER AMORTIZACION PARA RENTA + 347 + 415", facturar: "NO" },
      { fila: 18, titulo: "NO TRAMITAR — SE ENCARGAN ELLOS O PENDIENTE", facturar: "NO" },
      { fila: 20, titulo: "MOROSOS", facturar: "REVISAR" },
      { fila: 21, titulo: "SE FUERON DE LA ASESORIA O NULOS — NO SE TRAMITAN", facturar: "NO" },
    ],
  },
  {
    hoja: "421",
    headerRow: 2,
    labelRows: [2, 1],
    dataFrom: 3,
    cols: { expte: 2, nif: 5, empresa: 6, p: [12, 13, 14, 15], obs: 18 },
    inicial: { modelo: "421", facturar: "SI", frecuencia: "TRIMESTRAL", titulo: "MODELO 421 TRIMESTRAL" },
    secciones: [
      { fila: 9, titulo: "BAJA ACTIVIDAD O CAMBIO A CONTABILIDAD", facturar: "NO" },
      { fila: 12, titulo: "REGIMEN PEQUEÑOS EMPRESARIOS Y PROFESIONALES (exentos)", facturar: "NO" },
      { fila: 14, titulo: "MINORISTAS (exentos)", facturar: "NO" },
      { fila: 18, titulo: "REGIMEN ESPECIAL GANADERIA", facturar: "REVISAR" },
      { fila: 21, titulo: "MOROSOS — ¿TRAMITAR?", facturar: "REVISAR" },
      { fila: 23, titulo: "NO TRAMITAR / SE FUERON / ELLOS / CONTABILIDAD", facturar: "NO" },
    ],
  },
  {
    hoja: "115",
    headerRow: 4,
    labelRows: [4, 1],
    dataFrom: 2,
    cols: { expte: 2, nif: 5, empresa: 6, p: [12, 13, 14, 15], obs: null },
    inicial: { modelo: "046", facturar: "SI", frecuencia: "TRIMESTRAL", titulo: "MODELO 046 TRIMESTRAL" },
    secciones: [
      { fila: 2, skip: true }, // título original "MODELO 046 - TRIMESTRAL"
      { fila: 4, titulo: "MODELO 115 Y 180 TRIMESTRAL", modelo: "115", facturar: "SI" }, // fila de cabecera
      { fila: 80, titulo: "BAJAS 2026", facturar: "NO" },
      { fila: 82, skip: true }, // cabecera repetida
      { fila: 87, titulo: "EMPRESAS EXENTAS O QUE NO APLICAN RETENCION", facturar: "NO" },
      { fila: 95, titulo: "SE ENCARGAN ELLOS / SE FUERON", facturar: "NO" },
      { fila: 100, titulo: "MOROSOS — ¿TRAMITAR?", facturar: "REVISAR" },
    ],
  },
];

// Hojas que se copian literalmente (estructura no tabular o de referencia).
const VERBATIM = [
  { hoja: "alcohol", motivo: "leyenda de modelos IIEE y plazos, sin filas de clientes", dateCols: [8, 9] },
  { hoja: "184", motivo: "estructura por comuneros (anual informativa), se decidirá en otra fase", dateCols: [9, 10] },
  { hoja: "EXP", motivo: "referencia expte → expte facturación (equivale al mapeo ClientesXExptes)", dateCols: [] },
];

// ---------------------------------------------------------------- lectura y transformación

function readSheet(wb, name) {
  // sheet_to_json(undefined) devuelve [] en vez de lanzar: sin esta guarda, una
  // hoja que el original ya no trae se convierte en una hoja vacía sin avisar.
  if (!wb.Sheets[name]) throw new Error(`Hoja '${name}' no encontrada en el original`);
  return XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: null });
}

function filaVacia(cells) {
  return !cells.some((c) => c !== null && String(c).trim() !== "");
}

// ¿La fila describe un cliente o es una nota suelta? Solo para filas ya
// conocidas como no vacías: con `cells` vacío devuelve true (los undefined
// cuentan en nonNull). Da falsos positivos con las cabeceras repetidas (su NIF
// trae el literal "N.I.F." / "CIF."), que es lo que parchean los `skip:true`;
// por eso decidir si una fila es sección se hace con `esFilaSeccionCandidata`.
function esDato(spec, cells) {
  if (toInt(cells[spec.cols.expte]) !== null) return true;
  const p = spec.cols.p.map((i) => (i !== null && i !== undefined ? cells[i] : null));
  const mapped = [cells[spec.cols.nif], cells[spec.cols.empresa], ...p];
  const nonNull = mapped.filter((c) => c !== null && String(c).trim() !== "").length;
  return str(cells[spec.cols.nif]) !== "" || nonNull >= 3;
}

// Invariante del formato: una fila de sección nunca lleva expte numérico.
function esFilaSeccionCandidata(spec, cells) {
  return !filaVacia(cells) && toInt(cells[spec.cols.expte]) === null;
}

function textoFila(cells, dropped) {
  return cells
    .map((c, j) => (c !== null && !dropped.has(j) ? str(c) : ""))
    .filter(Boolean)
    .join(" ");
}

// Localiza cada sección del spec en el archivo concreto y devuelve la lista
// resuelta [{...seccion, fila, textoMatch}]. Falla antes de transformar nada:
// un ancla mal resuelta se come una fila de datos en silencio y esa fila deja de
// facturarse, así que aquí todo lo dudoso es error, no warning.
function resolveSecciones(spec, rows, warnings) {
  const dropped = new Set(spec.drop || []);
  const resueltas = [];

  for (const s of spec.secciones || []) {
    if (s.fila !== undefined) {
      // Fuera de rango no lanzaría por sí solo (rows[...] → undefined) y el bucle
      // de transformSheet nunca alcanzaría la fila: el separador no se emitiría y
      // sus clientes heredarían el FACTURAR de la sección anterior, en silencio.
      if (s.fila < 1 || s.fila > rows.length) {
        throw new Error(
          `[${spec.hoja}] el ancla legacy de la fila ${s.fila} ("${s.titulo}") queda fuera del archivo ` +
            `(${rows.length} filas): el archivo ha cambiado respecto al que calibró el spec. ` +
            `Migra esta sección a { match: /.../ }.`
        );
      }
      const cells = rows[s.fila - 1] || [];
      if (toInt(cells[spec.cols.expte]) !== null) {
        throw new Error(
          `[${spec.hoja}] el ancla legacy de la fila ${s.fila} ("${s.titulo}") cae sobre una fila con EXPTE ` +
            `${str(cells[spec.cols.expte])}: el archivo ha cambiado respecto al que calibró el spec. ` +
            `Migra esta sección a { match: /.../ }.`
        );
      }
      resueltas.push({ ...s, fila: s.fila, textoMatch: textoFila(cells, dropped) });
      continue;
    }

    const hits = [];
    for (let fila = spec.dataFrom; fila <= rows.length; fila++) {
      if (fila === spec.headerRow) continue;
      const cells = rows[fila - 1] || [];
      if (!esFilaSeccionCandidata(spec, cells)) continue;
      const texto = textoFila(cells, dropped);
      if (s.match.test(texto)) hits.push({ fila, texto });
    }
    if (hits.length === 0) {
      throw new Error(`[${spec.hoja}] sección "${s.titulo}" no encontrada (patrón ${s.match}).`);
    }
    if (hits.length > 1) {
      throw new Error(
        `[${spec.hoja}] sección "${s.titulo}" ambigua (patrón ${s.match}): ` +
          `filas ${hits.map((h) => h.fila).join(", ")}. Afina el patrón.`
      );
    }
    resueltas.push({ ...s, fila: hits[0].fila, textoMatch: hits[0].texto });
  }

  // Dos secciones en la misma fila colapsarían en el Map de anclas y una
  // desaparecería sin ruido: cada `match` puede tener 1 hit y aun así chocar.
  const porFila = new Map();
  for (const s of resueltas) {
    if (porFila.has(s.fila)) {
      throw new Error(
        `[${spec.hoja}] las secciones "${porFila.get(s.fila).titulo}" y "${s.titulo}" resuelven ambas a la fila ${s.fila}.`
      );
    }
    porFila.set(s.fila, s);
  }

  // Si el orden se altera, cada patrón sigue teniendo 1 hit pero las secciones
  // quedan cruzadas: mismo separador, título equivocado.
  for (let i = 1; i < resueltas.length; i++) {
    if (resueltas[i].fila <= resueltas[i - 1].fila) {
      warnings.push(
        `[${spec.hoja}] las secciones no salen en el orden del spec: "${resueltas[i - 1].titulo}" (fila ` +
          `${resueltas[i - 1].fila}) va antes que "${resueltas[i].titulo}" (fila ${resueltas[i].fila}). ` +
          `Revisa que cada patrón matchee el título que le toca.`
      );
    }
  }

  return resueltas;
}

function buildLabels(spec, rows, maxCol) {
  const labels = [];
  for (let c = 0; c <= maxCol; c++) {
    if (spec.labelOverrides && spec.labelOverrides[c] !== undefined) {
      labels[c] = spec.labelOverrides[c];
      continue;
    }
    const parts = [];
    for (const lr of spec.labelRows) {
      const v = (rows[lr - 1] || [])[c];
      const s = str(v);
      if (s && !parts.includes(s)) parts.push(s);
    }
    labels[c] = parts.join(" ") || `(col ${colLetter(c)})`;
  }
  return labels;
}

function transformSheet(spec, rows, informe, warnings) {
  const dropped = new Set(spec.drop || []);
  const skipRows = new Set(spec.skipRows || []);
  const secciones = resolveSecciones(spec, rows, warnings);
  const anchors = new Map(secciones.map((s) => [s.fila, s]));

  let maxCol = 0;
  rows.forEach((r) => (r || []).forEach((c, j) => {
    if (c !== null && !dropped.has(j) && j > maxCol) maxCol = j;
  }));

  // Qué columnas de `drop` traían datos de verdad en este archivo: el informe no
  // debe presumir de descartar un bloque que el original ya no trae.
  const dropConDatos = [...dropped]
    .filter((j) => rows.some((r) => (r || [])[j] !== null && (r || [])[j] !== undefined && String((r || [])[j]).trim() !== ""))
    .sort((a, b) => a - b);

  const labels = buildLabels(spec, rows, maxCol);
  // La Zona A ya no consume P1..P4 ni OBSERVACIONES: se quedan en su posición
  // original dentro de la zona libre (dato crudo que el importador no lee), pues
  // la facturación sale del CONCEPTO FACT y no del importe del periodo.
  const consumed = new Set([spec.cols.expte, spec.cols.nif, spec.cols.empresa]);
  const zonaBCols = [];
  for (let c = 0; c <= maxCol; c++) {
    if (!consumed.has(c) && !dropped.has(c)) zonaBCols.push(c);
  }

  const out = []; // { tipo: 'seccion'|'dato', ... }
  const notas = [];
  let cur = { ...spec.inicial };
  const stats = { SI: 0, NO: 0, REVISAR: 0, sinExpte: 0, notas: 0, frecuenciaOverrides: 0 };
  const porPeriodo = [0, 0, 0, 0]; // filas SI con P rellena

  for (let fila = spec.dataFrom; fila <= rows.length; fila++) {
    const anchor = anchors.get(fila);
    if (anchor) {
      if (!anchor.skip) {
        cur = {
          modelo: anchor.modelo || cur.modelo,
          facturar: anchor.facturar,
          frecuencia: anchor.frecuencia || cur.frecuencia,
          titulo: anchor.titulo,
        };
        out.push({ tipo: "seccion", titulo: cur.titulo });
      }
      continue;
    }
    if (fila === spec.headerRow || skipRows.has(fila)) continue;

    const cells = rows[fila - 1] || [];
    if (filaVacia(cells)) continue;

    const expte = toInt(cells[spec.cols.expte]);
    const nif = str(cells[spec.cols.nif]);
    const p = spec.cols.p.map((i) => (i !== null && i !== undefined ? cells[i] : null));
    while (p.length < 4) p.push(null);

    // fila sin expte: ¿cliente sin código o nota suelta?
    if (!esDato(spec, cells)) {
      const texto = cells
        .map((c, j) => (c !== null && !dropped.has(j) ? str(c) : ""))
        .filter(Boolean)
        .join(" · ");
      notas.push({ filaOrigen: fila, texto });
      stats.notas++;
      continue;
    }

    if (expte === null) {
      stats.sinExpte++;
      informe.sinExpte.push({
        hoja: spec.hoja, filaOrigen: fila,
        empresa: str(cells[spec.cols.empresa]) || nif, seccion: cur.titulo,
      });
    }
    stats[cur.facturar]++;

    p.forEach((v, i) => {
      if (v === null || v === undefined) return;
      const s = String(v).trim();
      if (s === "") return;
      if (cur.facturar === "SI") porPeriodo[i]++;
      if (typeof v !== "number" && s.toUpperCase() !== "X") {
        informe.pNoConformes.push({
          hoja: spec.hoja, filaOrigen: fila, periodo: `P${i + 1}`, valor: s,
          empresa: str(cells[spec.cols.empresa]), facturar: cur.facturar, seccion: cur.titulo,
        });
      }
    });

    let frecuencia = cur.frecuencia;
    if (spec.frecuenciaDesdeColumna) {
      const { col, valores } = spec.frecuenciaDesdeColumna;
      const raw = str(cells[col]).toUpperCase();
      if (valores[raw]) {
        frecuencia = valores[raw];
        stats.frecuenciaOverrides++;
      }
    }

    out.push({
      tipo: "dato",
      filaOrigen: fila,
      modelo: cur.modelo,
      concepto: conceptoDe(cur.modelo),
      expte,
      nif,
      empresa: str(cells[spec.cols.empresa]),
      facturar: cur.facturar,
      frecuencia,
      // P1..P4 y OBSERVACIONES ya no van en Zona A; salen en zonaB con su valor crudo.
      zonaB: zonaBCols.map((c) => cells[c] ?? null),
    });
  }

  return { out, notas, labels, zonaBCols, stats, porPeriodo, maxCol, secciones, dropConDatos };
}

// ---------------------------------------------------------------- escritura

// Tercer elemento `true` = línea de título (negrita). Evita indexar filas a
// mano: añadir/quitar una línea ya no desalinea qué queda en negrita.
function writeLeeme(sheet, fechaGen, inputName, specsSel, verbatimSel) {
  const hojasModelo = specsSel.map((s) => s.hoja).join(", ") || "ninguna";
  const lines = [
    ["PLANTILLA 4PAGOS → IMPORTACIÓN A3GES", "", true],
    [`Generada el ${fechaGen} a partir de "${inputName}".`, ""],
    [`Hojas de modelo que cubre esta plantilla: ${hojasModelo}.`, ""],
    ["", ""],
    ["CÓMO FUNCIONA", "", true],
    ["Cada hoja de modelo tiene dos zonas:", ""],
    ["  · ZONA A (columnas A a G): bloque estándar que lee el importador. NO insertar, borrar ni renombrar columnas aquí.", ""],
    ["  · ZONA B (columna I en adelante): zona libre de cada modelo (incluye P1–P4, OBSERVACIONES, F.BAJA y el resto de columnas originales). El importador no la lee; se puede modificar libremente.", ""],
    ["", ""],
    ["EL MODELO EN UNA LÍNEA", "CONCEPTO FACT = qué · FACTURAR = si sí o no · FRECUENCIA = cuándo · EXPTE = a quién · IMPORTE = cuánto (opcional).", true],
    ["", ""],
    ["COLUMNAS DE LA ZONA A", "", true],
    ["  CONCEPTO FACT", "Concepto facturable de la fila (código de la hoja ConceptosFacturables: 0.016, 0.012…). Es lo que fija el importe a facturar al cruzarlo con mapeos_facturacion.xlsx. Se deriva del modelo fiscal (130, 111…)."],
    ["  EXPTE", "Código de cliente (número corto). No decide si la fila se factura (eso es FACTURAR), pero una fila SI sin EXPTE no se puede facturar: se migra vacía y sale en incidencias hasta que se le asigne."],
    ["  NIF", "Informativo / control de calidad."],
    ["  EMPRESA", "Informativo."],
    ["  FACTURAR", "Única fuente de verdad sobre la decisión de facturar la fila: SI = se factura · NO = nunca · REVISAR = no se factura y sale en incidencias para decidir a mano. El importador no cruza ningún otro campo para decidirlo: ni P1–P4, ni la fecha de baja de la Zona B."],
    ["  FRECUENCIA", "TRIMESTRAL (por defecto) · MENSUAL · ANUAL · OTRA. Periodicidad real de la fila; puede venir heredada de la hoja/sección o corregida a mano por excepción (p. ej. una empresa grande que declara el 111 mensualmente). Cada ejecución factura las filas cuya frecuencia toca en el periodo elegido: en un cierre de trimestre entran tanto las TRIMESTRAL como las MENSUAL; en un mes intermedio, solo las MENSUAL."],
    ["  IMPORTE", "Precio puntual de ESTA fila. Vacío (lo normal) = se factura la tarifa del catálogo. Con un número = se factura ese importe, ignorando el catálogo (sirve para cobrar más o menos a un cliente concreto, o para dar precio a un concepto ESCALADO). OJO: la plantilla se reutiliza entre trimestres — un importe que se deja escrito se vuelve a facturar. Cada ejecución lista todos los importes puntuales usados en precios_manuales.csv para poder revisarlos."],
    ["", ""],
    ["P1..P4 y OBSERVACIONES (ahora en la ZONA LIBRE)", "", true],
    ["  Se copian tal cual del archivo original y el importador NO las lee: son dato informativo. Periodos trimestrales P1=1T…P4=4T (modelo 202: P1=abril, P2=octubre, P3=diciembre). La facturación la fija CONCEPTO FACT; el importe del periodo no interviene.", ""],
    ["", ""],
    ["FILAS NARANJAS", "Separadores de sección heredados del original. El importador las ignora (no tienen EXPTE)."],
    ["FILAS GRISES AL FINAL", "Notas y leyendas del archivo original, conservadas para no perder información."],
  ];
  if (verbatimSel.length) {
    lines.push([
      `HOJAS ${verbatimSel.map((v) => v.hoja).join(" / ")}`,
      "Copiadas tal cual del original; el importador no las procesa por ahora.",
    ]);
  }
  lines.push(["", ""]);
  lines.push(["El importador solo procesa hojas cuya celda A1 contenga: " + SENTINEL, ""]);
  lines.forEach(([label, detail, isTitle], i) => {
    const cell = sheet.cell(i + 1, 1).value(label);
    if (detail) sheet.cell(i + 1, 2).value(detail);
    if (isTitle) cell.style({ bold: true });
  });
  sheet.cell(1, 1).style({ fontSize: 14 });
  sheet.column("A").width(30);
  sheet.column("B").width(110);
}

function writeModelSheet(sheet, spec, t, fechaGen) {
  const nZonaB = t.zonaBCols.length;

  // Fila 1: centinela + título
  sheet.cell(1, 1).value(SENTINEL).style({ bold: true, fontColor: "808080" });
  sheet.cell(1, 4).value(`Hoja "${spec.hoja}" — plantilla A3 generada ${fechaGen}`).style({ bold: true, fontSize: 12 });
  if (nZonaB > 0) {
    sheet.cell(1, ZONA_B_START).value("ZONA LIBRE (el importador no lee estas columnas) →").style({ bold: true, fontColor: "808080" });
  }

  // Fila 2: cabeceras
  ZONA_A.forEach((h, i) => {
    sheet.cell(2, i + 1).value(h).style({ bold: true, fontColor: "FFFFFF", fill: C_HEADER_A });
  });
  sheet.cell(2, COL_SEP).value("").style({ fill: C_SEP });
  t.zonaBCols.forEach((c, i) => {
    sheet.cell(2, ZONA_B_START + i).value(t.labels[c]).style({ bold: true, fontColor: "FFFFFF", fill: C_HEADER_B });
  });

  // Datos
  let r = 3;
  for (const row of t.out) {
    if (row.tipo === "seccion") {
      for (let c = 1; c <= ZONA_A.length; c++) sheet.cell(r, c).style({ fill: C_SECCION, bold: true });
      sheet.cell(r, 4).value(`— ${row.titulo} —`);
      r++;
      continue;
    }
    sheet.cell(r, 1).value(row.concepto);
    if (row.expte !== null) sheet.cell(r, 2).value(row.expte);
    if (row.nif) sheet.cell(r, 3).value(row.nif);
    if (row.empresa) sheet.cell(r, 4).value(row.empresa);
    sheet.cell(r, 5).value(row.facturar);
    sheet.cell(r, 6).value(row.frecuencia);
    sheet.cell(r, COL_SEP).style({ fill: C_SEP });
    row.zonaB.forEach((v, i) => {
      if (v === null || v === undefined || String(v).trim() === "") return;
      const cell = sheet.cell(r, ZONA_B_START + i).value(v);
      if (isDateSerial(v) && DATE_HEADER_RX.test(t.labels[t.zonaBCols[i]])) {
        cell.style("numberFormat", "dd/mm/yyyy");
      }
    });
    r++;
  }

  // Desplegables en FACTURAR y FRECUENCIA (todos los atributos explícitos:
  // xlsx-populate serializa literalmente lo que recibe y "undefined" corrompe el XML).
  if (r > 3) {
    sheet.dataValidation(`E3:E${r - 1}`, {
      type: "list", allowBlank: true, showInputMessage: false, prompt: "", promptTitle: "",
      showErrorMessage: true, error: "Valores permitidos: SI, NO, REVISAR", errorTitle: "FACTURAR",
      operator: "between", formula1: '"SI,NO,REVISAR"', formula2: "",
    });
    sheet.dataValidation(`F3:F${r - 1}`, {
      type: "list", allowBlank: true, showInputMessage: false, prompt: "", promptTitle: "",
      showErrorMessage: true, error: `Valores permitidos: ${FRECUENCIAS.join(", ")}`, errorTitle: "FRECUENCIA",
      operator: "between", formula1: `"${FRECUENCIAS.join(",")}"`, formula2: "",
    });
  }

  // Notas del original
  if (t.notas.length) {
    r++;
    sheet.cell(r, 4).value("— NOTAS Y LEYENDAS DEL ARCHIVO ORIGINAL (el importador las ignora) —").style({ bold: true, fill: C_NOTA });
    r++;
    for (const n of t.notas) {
      sheet.cell(r, 4)
        .value(`${n.texto}  [fila ${n.filaOrigen} del original]`)
        .style({ italic: true, fontColor: "595959", fill: C_NOTA });
      r++;
    }
  }

  // Presentación
  sheet.freezePanes(0, 2);
  // G = IMPORTE (Zona A), H = separador. El resto de Zona B se dimensiona abajo.
  const widths = { A: 13, B: 8, C: 13, D: 42, E: 11, F: 12, G: 11, H: 2 };
  Object.entries(widths).forEach(([col, w]) => sheet.column(col).width(w));
  // Formato moneda solo de display: el importador lee el valor numérico, no el
  // texto. Deja vacías las celdas hasta que el usuario escriba un precio.
  sheet.column(colLetter(COL_IMPORTE)).style("numberFormat", "#,##0.00 €");
  for (let i = 0; i < nZonaB; i++) {
    sheet.column(colLetter(ZONA_B_START - 1 + i)).width(14);
  }
}

function writeVerbatimSheet(sheet, rows, dateCols) {
  const dc = new Set(dateCols || []);
  rows.forEach((row, i) => {
    (row || []).forEach((v, j) => {
      if (v === null || v === undefined) return;
      const cell = sheet.cell(i + 1, j + 1).value(v);
      if (dc.has(j) && isDateSerial(v)) cell.style("numberFormat", "dd/mm/yyyy");
    });
  });
}

// ---------------------------------------------------------------- informe

function writeInforme(informePath, ctx) {
  const { inputName, outputName, fechaGen, resumen, informe, verbatims,
    omitidasVacias, sinSpecConDatos, noSeleccionadas, warnings } = ctx;
  const L = [];
  L.push(`# Informe de migración — 4PAGOS → plantilla A3`);
  L.push("");
  L.push(`- **Origen:** \`${inputName}\``);
  L.push(`- **Generado:** \`${outputName}\` (${fechaGen})`);
  L.push(`- **Hojas migradas:** ${[...resumen.map((r) => r.hoja), ...verbatims.map((v) => v.hoja)].join(", ") || "ninguna"}`);
  L.push("");
  if (warnings.length) {
    L.push(`## Avisos`);
    L.push("");
    for (const w of warnings) L.push(`- ${w}`);
    L.push("");
  }
  L.push(`## Resumen por hoja`);
  L.push("");
  L.push(`| Hoja | Filas SI | Filas NO | Filas REVISAR | Sin expte | Notas conservadas | P1 | P2 | P3 | P4 |`);
  L.push(`|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|`);
  for (const r of resumen) {
    L.push(`| ${r.hoja} | ${r.stats.SI} | ${r.stats.NO} | ${r.stats.REVISAR} | ${r.stats.sinExpte} | ${r.stats.notas} | ${r.porPeriodo.join(" | ")} |`);
  }
  L.push("");
  L.push(`P1–P4 es un **diagnóstico del archivo original**: cuenta las filas con FACTURAR=SI que traían el periodo relleno. No es una regla de facturación — el importador no lee esas columnas (ver LEEME).`);
  L.push("");

  L.push(`## Columna FRECUENCIA`);
  L.push("");
  L.push(`| Hoja | Frecuencia por defecto | Excepciones por fila detectadas |`);
  L.push(`|---|---|---:|`);
  for (const r of resumen) {
    L.push(`| ${r.hoja} | ${r.spec.inicial.frecuencia} | ${r.stats.frecuenciaOverrides} |`);
  }
  L.push("");
  const r111 = resumen.find((r) => r.hoja === "111");
  if (r111) {
    L.push(`En la hoja 111, ${r111.stats.frecuenciaOverrides} empresas tenían "MENSUAL" escrito en la columna BAJA del original — se promovió a FRECUENCIA=MENSUAL y F.BAJA se dejó intacta en Zona B.`);
    L.push("");
  }

  const sinObs = resumen.filter((r) => r.spec.cols.obs === null || r.spec.cols.obs === undefined).map((r) => r.hoja);
  L.push(`## Columna OBSERVACIONES`);
  L.push("");
  L.push(`Unifica bajo un único nombre la columna de notas de cada hoja (OBS/OBSERV/OBSERVACIONES según el modelo).`);
  L.push(
    sinObs.length
      ? `Hojas sin columna de notas en el original (OBSERVACIONES queda vacía): ${sinObs.join(", ")}.`
      : `Todas las hojas traían columna de notas en el original.`
  );
  L.push("");

  L.push(`## Secciones aplicadas (decisión FACTURAR heredada del bloque del original)`);
  L.push("");
  L.push(`Cada sección se localiza por el texto de su título en el archivo de origen. Se listan la fila donde se resolvió y el texto crudo que hizo match: es el rastro para comprobar que el bloque es el que se esperaba.`);
  L.push("");
  L.push(`| Hoja | Sección | → FACTURAR | Fila en el origen | Texto encontrado |`);
  L.push(`|---|---|---|---:|---|`);
  for (const r of resumen) {
    L.push(`| ${r.hoja} | ${r.spec.inicial.titulo} _(inicial)_ | ${r.spec.inicial.facturar} | ${r.spec.dataFrom} | — |`);
    for (const s of r.secciones.filter((s) => !s.skip)) {
      L.push(`| ${r.hoja} | ${s.titulo} | ${s.facturar} | ${s.fila} | \`${s.textoMatch.slice(0, 60)}\` |`);
    }
  }
  L.push("");

  L.push(`## Celdas de periodo con texto (revisar FACTURAR a mano)`);
  L.push("");
  L.push(`Diagnóstico del archivo original: el importador no lee P1–P4, así que estas celdas no bloquean nada por sí solas. Importan porque contradicen la fila: una empresa con FACTURAR=SI cuyo periodo pone "NO" o "??" es candidata a poner en REVISAR o NO. La decisión es manual — el script no la toma.`);
  L.push("");
  const pRelevantes = informe.pNoConformes.filter((p) => p.facturar !== "NO");
  const pEnNo = informe.pNoConformes.filter((p) => p.facturar === "NO");
  if (pRelevantes.length === 0) {
    L.push(`Ninguna en filas SI/REVISAR.`);
  } else {
    L.push(`En filas **SI/REVISAR**: ${pRelevantes.length}`);
    L.push("");
    L.push(`| Hoja | Fila original | Periodo | Valor | Empresa | FACTURAR |`);
    L.push(`|---|---|---|---|---|---|`);
    for (const p of pRelevantes.slice(0, 200)) {
      L.push(`| ${p.hoja} | ${p.filaOrigen} | ${p.periodo} | ${p.valor} | ${p.empresa} | ${p.facturar} |`);
    }
    if (pRelevantes.length > 200) L.push(`| … | | | (${pRelevantes.length - 200} más) | | |`);
  }
  L.push("");
  if (pEnNo.length) {
    const porValor = new Map();
    for (const p of pEnNo) {
      const k = `${p.hoja} · "${p.valor}"`;
      porValor.set(k, (porValor.get(k) || 0) + 1);
    }
    L.push(`En filas **NO** (no afectan a la facturación, solo informativo): ${pEnNo.length} celdas. Desglose:`);
    for (const [k, n] of [...porValor.entries()].sort((a, b) => b[1] - a[1])) L.push(`- ${k}: ${n}`);
    L.push("");
  }

  L.push(`## Filas de cliente sin EXPTE (migradas con EXPTE vacío — no facturables hasta asignarlo)`);
  L.push("");
  if (informe.sinExpte.length === 0) L.push(`Ninguna.`);
  else {
    L.push(`| Hoja | Fila original | Empresa/NIF | Sección |`);
    L.push(`|---|---|---|---|`);
    for (const s of informe.sinExpte) L.push(`| ${s.hoja} | ${s.filaOrigen} | ${s.empresa} | ${s.seccion} |`);
  }
  L.push("");

  L.push(`## Hojas copiadas tal cual (sin Zona A)`);
  L.push("");
  L.push(verbatims.length ? verbatims.map((v) => `- **${v.hoja}** — ${v.motivo}`).join("\n") : "Ninguna.");
  L.push("");
  L.push(`## Hojas del original no migradas`);
  L.push("");
  L.push(`- **No seleccionadas** (tienen spec, excluidas con \`--hojas\`): ${noSeleccionadas.join(", ") || "ninguna"}`);
  L.push(`- **Vacías en el original**: ${omitidasVacias.join(", ") || "ninguna"}`);
  L.push(`- **Con datos y sin spec** (se pierden si hacen falta): ${sinSpecConDatos.join(", ") || "ninguna"}`);
  L.push("");
  L.push(`## Columnas descartadas`);
  L.push("");
  const conDrop = resumen.filter((r) => r.dropConDatos.length);
  if (conDrop.length === 0) {
    L.push(`Ninguna: las columnas marcadas para descartar en los specs no traían datos en este archivo.`);
  } else {
    for (const r of conDrop) {
      L.push(`- Hoja **${r.hoja}**, columnas ${r.dropConDatos.map(colLetter).join(", ")} del original: bloque construido a mano (ejemplo del output A3). Se descarta porque es dato derivado que ahora genera el proceso.`);
    }
  }
  L.push("");

  fs.writeFileSync(informePath, L.join("\n"), "utf8");
}

// ---------------------------------------------------------------- main

// Separa los flags de los posicionales: leer process.argv[2] a pelo hace que
// `--hojas 111` acabe como ruta de entrada.
function parseArgs(argv) {
  const pos = [];
  let hojas = null;
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === "--hojas") {
      hojas = argv[++i];
      if (hojas === undefined) throw new Error("--hojas necesita un valor: --hojas 111,130");
    } else if (a.startsWith("--hojas=")) hojas = a.slice("--hojas=".length);
    else if (a.startsWith("--")) throw new Error(`Opción desconocida: ${a}`);
    else pos.push(a);
  }
  const sel = hojas === null ? null : hojas.split(",").map((h) => h.trim()).filter(Boolean);
  if (sel && sel.length === 0) throw new Error("--hojas no puede ir vacío");
  if (sel) {
    const conocidas = new Set([...SPECS.map((s) => s.hoja), ...VERBATIM.map((v) => v.hoja)]);
    const desconocidas = sel.filter((h) => !conocidas.has(h));
    if (desconocidas.length) {
      throw new Error(
        `--hojas: sin spec para ${desconocidas.join(", ")}. Disponibles: ${[...conocidas].join(", ")}`
      );
    }
  }
  return { input: pos[0] || DEFAULT_INPUT, output: pos[1] || DEFAULT_OUTPUT, hojas: sel };
}

async function main() {
  const { input, output, hojas } = parseArgs(process.argv.slice(2));
  const informePath = output.replace(/\.xlsx$/i, "") + " - informe migracion.md";
  const fechaGen = new Date().toISOString().slice(0, 10);

  const specsSel = hojas ? SPECS.filter((s) => hojas.includes(s.hoja)) : SPECS;
  const verbatimSel = hojas ? VERBATIM.filter((v) => hojas.includes(v.hoja)) : VERBATIM;

  console.log(`Leyendo ${input}`);
  const wb = XLSX.readFile(input);

  const informe = { pNoConformes: [], sinExpte: [] };
  const warnings = [];
  const resumen = [];
  const transformadas = new Map();

  for (const spec of specsSel) {
    const rows = readSheet(wb, spec.hoja);
    const t = transformSheet(spec, rows, informe, warnings);
    transformadas.set(spec.hoja, t);
    resumen.push({ hoja: spec.hoja, stats: t.stats, porPeriodo: t.porPeriodo, spec, secciones: t.secciones, dropConDatos: t.dropConDatos });
    console.log(
      `  ${spec.hoja.padEnd(18)} SI=${String(t.stats.SI).padStart(4)}  NO=${String(t.stats.NO).padStart(4)}  ` +
      `REVISAR=${String(t.stats.REVISAR).padStart(3)}  sinExpte=${t.stats.sinExpte}  notas=${t.stats.notas}` +
      (t.stats.frecuenciaOverrides ? `  frecOverrides=${t.stats.frecuenciaOverrides}` : "")
    );
  }

  // Tres cosas distintas que antes se agrupaban como "omitidas (vacías)":
  const generadas = new Set([...specsSel.map((s) => s.hoja), ...verbatimSel.map((v) => v.hoja)]);
  const conSpec = new Set([...SPECS.map((s) => s.hoja), ...VERBATIM.map((v) => v.hoja)]);
  const noSeleccionadas = [...conSpec].filter((n) => !generadas.has(n));
  const sinSpec = wb.SheetNames.filter((n) => !conSpec.has(n));
  const hojaVacia = (n) => readSheet(wb, n).every((r) => !r || filaVacia(r));
  const omitidasVacias = sinSpec.filter(hojaVacia);
  const sinSpecConDatos = sinSpec.filter((n) => !hojaVacia(n));
  for (const n of sinSpecConDatos) {
    warnings.push(`La hoja '${n}' del original tiene datos y no está en SPECS ni en VERBATIM: no se migra.`);
  }

  console.log(`Escribiendo ${output}`);
  const outWb = await XlsxPopulate.fromBlankAsync();
  outWb.sheet(0).name("LEEME");
  writeLeeme(outWb.sheet(0), fechaGen, path.basename(input), specsSel, verbatimSel);

  for (const spec of specsSel) {
    const sheet = outWb.addSheet(spec.hoja);
    writeModelSheet(sheet, spec, transformadas.get(spec.hoja), fechaGen);
  }
  for (const v of verbatimSel) {
    const sheet = outWb.addSheet(v.hoja);
    writeVerbatimSheet(sheet, readSheet(wb, v.hoja), v.dateCols);
  }

  await outWb.toFileAsync(output);
  writeInforme(informePath, {
    inputName: path.basename(input), outputName: path.basename(output),
    fechaGen, resumen, informe, verbatims: verbatimSel,
    omitidasVacias, sinSpecConDatos, noSeleccionadas, warnings,
  });

  console.log(`Informe: ${informePath}`);
  if (noSeleccionadas.length) console.log(`Hojas no seleccionadas (--hojas): ${noSeleccionadas.join(", ")}`);
  console.log(`Hojas omitidas (vacías): ${omitidasVacias.join(", ") || "ninguna"}`);
  console.log(
    `P con texto a normalizar: ${informe.pNoConformes.length} · filas sin expte: ${informe.sinExpte.length}`
  );
  if (warnings.length) {
    console.log(`\nAvisos (${warnings.length}):`);
    for (const w of warnings) console.log(`  · ${w}`);
  }
}

// Solo corre como CLI: `actualizarLeeme.js` requiere este módulo para reutilizar
// `writeLeeme` (el texto del LEEME vive aquí y en un solo sitio), y sin esta
// guarda el simple require lanzaría la migración completa.
if (require.main === module) {
  main().catch((err) => {
    console.error(err);
    process.exit(1);
  });
}

module.exports = { writeLeeme, SPECS, VERBATIM, SENTINEL };
