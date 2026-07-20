const path = require("path");
const XlsxPopulate = require("xlsx-populate");
const { REDIRECT_NADA } = require("./mapeos");
const {
  _str,
  _toInt,
  pad5,
  isoDate,
  ensureDir,
  readAbsoluteRows,
  resolveHeaderColumns,
  writeCsv,
} = require("./utils");

const EMPRESA_FACTURADORA = 14;
const TIPO_IVA = 3;
const UNIDADES = 1;

// Centinela de A1 que marca las hojas que este importador lee. Lo escribe
// `pagos/migrarPlantilla.js` y lo documenta el LEEME de la plantilla: las hojas
// que no lo llevan (LEEME, hojas copiadas tal cual) se ignoran sin ruido.
const SENTINEL = "A3PAGOS v1";

// Zona A de la plantilla (cols A–F): el bloque estándar que lee el importador. Se
// localiza por el texto de la cabecera, como el resto de importadores. La Zona B
// (col H+) es zona libre y NO se lee: P1–P4, OBSERVACIONES y F.BAJA viven ahí como
// dato crudo. Quién se factura lo fija FACTURAR, qué y cuánto CONCEPTO FACT,
// cuándo FRECUENCIA y a quién EXPTE.
const HEADER_SYNONYMS = {
  concepto: ["conceptofact", "concepto"],
  expte: ["expte", "expt"],
  nif: ["nif"],
  empresa: ["empresa", "razonsocial"],
  facturar: ["facturar"],
  frecuencia: ["frecuencia"],
};

const HEADER_SCAN_ROWS = 10;

// Qué frecuencias entran en cada tipo de corrida: manda el periodo de la
// ejecución, no la fila. En un cierre de trimestre entran las mensuales igual que
// las trimestrales, todas con Unidades=1 (no se acumulan tres meses); en un mes
// intermedio solo las mensuales.
const FACTURA_EN = {
  TRIMESTRAL: new Set(["TRIMESTRAL", "MENSUAL"]),
  MENSUAL: new Set(["MENSUAL"]),
};

// Las cuatro que admite el desplegable de la columna FRECUENCIA de la plantilla.
const FRECUENCIAS_CONOCIDAS = new Set(["TRIMESTRAL", "MENSUAL", "ANUAL", "OTRA"]);

// Conocidas que ninguna corrida factura. Salen como incidencia en vez de saltarse
// en silencio: si no, una fila SI marcada ANUAL no se facturaría nunca y nada lo
// delataría.
const SIN_CORRIDA = new Set(["ANUAL", "OTRA"]);

const FACTURAR_VALIDOS = new Set(["SI", "NO", "REVISAR"]);

const PERIODO_RX = /^(\d{4})-(?:([1-4])T|(0[1-9]|1[0-2]))$/;

// "2026-2T" -> cierre trimestral; "2026-05" -> mes intermedio. El formato es el
// que discrimina la cadencia de la corrida.
function parsePeriodo(raw) {
  const s = _str(raw).toUpperCase();
  const m = s.match(PERIODO_RX);
  if (!m) {
    throw new Error(
      `Periodo '${_str(raw)}' no válido. Formatos admitidos: '2026-2T' (cierre trimestral) ` +
        `o '2026-05' (mes intermedio).`
    );
  }
  const anio = Number(m[1]);
  if (m[2]) {
    const trimestre = Number(m[2]);
    return {
      tipo: "TRIMESTRAL",
      anio,
      trimestre,
      // Fecha de devengo = último día del periodo, para que la línea caiga dentro
      // del trimestre que factura aunque el proceso se lance más tarde.
      fecha: finDeMes(anio, trimestre * 3),
      etiqueta: `${trimestre}T ${anio}`,
    };
  }
  const mes = Number(m[3]);
  return {
    tipo: "MENSUAL",
    anio,
    mes,
    fecha: finDeMes(anio, mes),
    etiqueta: `${String(mes).padStart(2, "0")}/${anio}`,
  };
}

// Día 0 del mes siguiente = último día de `mes` (1-based).
function finDeMes(anio, mes) {
  return new Date(anio, mes, 0);
}

function tieneCentinela(rows) {
  const fila1 = rows.find((r) => r.rowIndex === 1);
  if (!fila1) return false;
  return _str(fila1.cells[0]).toUpperCase().includes(SENTINEL.toUpperCase());
}

function isZonaAHeader(cols) {
  return (
    cols.concepto !== undefined &&
    cols.expte !== undefined &&
    cols.facturar !== undefined &&
    cols.frecuencia !== undefined
  );
}

// Una hoja con centinela cuya Zona A no se puede leer es una plantilla rota:
// saltarla en silencio dejaría de facturar a todos sus clientes, así que es error.
function locateZonaA(sheetName, rows) {
  for (const { rowIndex, cells } of rows) {
    if (rowIndex > HEADER_SCAN_ROWS) break;
    const { cols } = resolveHeaderColumns(cells, HEADER_SYNONYMS);
    if (isZonaAHeader(cols)) return { headerRow: rowIndex, cols };
  }
  throw new Error(
    `Hoja '${sheetName}': lleva el centinela ${SENTINEL} pero no se encuentra la cabecera de la Zona A ` +
      `(se requieren CONCEPTO FACT, EXPTE, FACTURAR y FRECUENCIA en las primeras ${HEADER_SCAN_ROWS} filas).`
  );
}

// El catálogo describe el concepto ("Modelo 111 -"); la etiqueta del periodo la
// pone la corrida. Se limpia la puntuación de cola con la que vienen varios
// nombres del catálogo para no acabar en "Modelo 111 - - 2T 2026".
function buildDescripcion(nombreConcepto, concepto, periodo) {
  const base = _str(nombreConcepto).replace(/[\s\-–—.,;:]+$/, "") || `Concepto ${concepto}`;
  return `${base} - ${periodo.etiqueta}`.slice(0, 250);
}

function buildDescAmpliada(empresa, nif) {
  return [empresa, nif].filter(Boolean).join(", ").slice(0, 500);
}

async function transform(inputPath, mapeos, outputDir, options = {}) {
  ensureDir(outputDir);
  const periodo = parsePeriodo(options.periodo);
  const facturables = FACTURA_EN[periodo.tipo];

  // El 4PAGOS original (.xls) y la plantilla (.xlsx) viven en la misma carpeta con
  // nombres casi idénticos, así que elegir el de origen es el error fácil. Sin esta
  // guarda el fallo sale como un críptico "Cannot read properties of null" de
  // xlsx-populate, que solo abre .xlsx.
  if (/\.xls$/i.test(inputPath)) {
    throw new Error(
      `'${path.basename(inputPath)}' es un .xls: este importador consume la PLANTILLA A3 de pagos (.xlsx), ` +
        `no el 4PAGOS original del cliente. La plantilla se genera antes con ` +
        `procesos/importarFacturacion/pagos/migrarPlantilla.js.`
    );
  }

  const workbook = await XlsxPopulate.fromFileAsync(path.normalize(inputPath));

  const conceptos = [];
  const incidencias = [];
  const warningsQc = [];
  const hojas = [];
  // EXPTE+concepto ya emitido en esta corrida -> dónde salió. La plantilla se
  // corrige a mano y la misma empresa repetida con el mismo modelo se facturaría
  // dos veces. La clave es el EXPTE de origen y NO el cliente que paga: en un
  // grupo, varias empresas redirigen al mismo pagador y cada una presenta su
  // propio modelo, así que le tocan varias líneas iguales y es correcto (las
  // distingue la Descripción Ampliada, que lleva la razón social de cada una).
  const vistos = new Map();

  for (const sheet of workbook.sheets()) {
    const { rows } = readAbsoluteRows(sheet);
    if (!rows.length) continue;
    if (!tieneCentinela(rows)) continue;

    const nombreHoja = sheet.name();
    const { headerRow, cols } = locateZonaA(nombreHoja, rows);
    const get = (cells, field) =>
      cols[field] !== undefined ? cells[cols[field]] : undefined;

    const stats = { conceptos: 0, si: 0, no: 0, revisar: 0, fuera_de_periodo: 0, incidencias: 0 };

    for (const { rowIndex: filaIdx, cells } of rows) {
      if (filaIdx <= headerRow) continue;

      const facturar = _str(get(cells, "facturar")).toUpperCase();
      // Las filas de sección (naranjas) y las notas del original no tienen
      // FACTURAR: es lo que las distingue de una fila de cliente.
      if (!facturar) continue;

      const concepto = _str(get(cells, "concepto"));
      const expte = _toInt(get(cells, "expte"));
      const nif = _str(get(cells, "nif"));
      const empresa = _str(get(cells, "empresa"));
      const frecuencia = _str(get(cells, "frecuencia")).toUpperCase();

      const addInc = (motivo) => {
        incidencias.push({
          hoja: nombreHoja,
          fila_origen: filaIdx,
          motivo,
          concepto,
          expte: expte !== null ? String(expte) : "",
          nif,
          empresa,
          facturar,
          frecuencia,
        });
        stats.incidencias++;
      };

      if (!FACTURAR_VALIDOS.has(facturar)) {
        addInc(`FACTURAR='${facturar}' no reconocido (valores válidos: SI, NO, REVISAR)`);
        continue;
      }
      if (facturar === "NO") {
        stats.no++;
        continue;
      }
      if (facturar === "REVISAR") {
        stats.revisar++;
        addInc("Marcada REVISAR en la plantilla — decidir a mano si se factura");
        continue;
      }
      stats.si++;

      if (!frecuencia) {
        addInc("Sin FRECUENCIA: no se sabe si le toca este periodo");
        continue;
      }
      if (!FRECUENCIAS_CONOCIDAS.has(frecuencia)) {
        addInc(
          `FRECUENCIA='${frecuencia}' no reconocida (${[...FRECUENCIAS_CONOCIDAS].join(", ")})`
        );
        continue;
      }
      if (SIN_CORRIDA.has(frecuencia)) {
        addInc(
          `FRECUENCIA=${frecuencia}: ninguna corrida (trimestral ni mensual) la factura — facturar a mano`
        );
        continue;
      }
      if (!facturables.has(frecuencia)) {
        // Trimestral en una corrida mensual: no le toca, y es lo esperado.
        stats.fuera_de_periodo++;
        continue;
      }

      if (expte === null) {
        addInc("Fila FACTURAR=SI sin EXPTE: no se puede facturar hasta asignarle código de cliente");
        continue;
      }
      if (!concepto) {
        addInc("Sin CONCEPTO FACT: no se sabe qué facturar");
        continue;
      }

      // 1. Redirección
      const redirectTarget = mapeos.redirect.resolve(expte);
      if (redirectTarget === REDIRECT_NADA) {
        addInc(`Cliente ${expte} marcado como no facturable (NADA)`);
        continue;
      }
      const clienteEfectivo = typeof redirectTarget === "number" ? redirectTarget : expte;

      // 2. Expediente
      const codigoExpediente = mapeos.exptes.resolve(clienteEfectivo);
      if (!codigoExpediente) {
        if (typeof redirectTarget === "number") {
          addInc(
            `Cliente destino ${clienteEfectivo} (redirect de ${expte}) sin expediente en mapeo`
          );
        } else {
          addInc(`Cliente ${expte} sin expediente formato B en mapeo`);
        }
        continue;
      }

      // 3. Tarifa
      const tarifa = mapeos.tarifas.resolve(concepto);
      if (tarifa === null) {
        const motivo = mapeos.tarifas.missReason(concepto);
        addInc(`Tarifa concepto '${concepto}' no resoluble: ${motivo}`);
        continue;
      }

      // 4. QC duplicados
      const clave = `${expte}|${concepto}`;
      const previo = vistos.get(clave);
      if (previo) {
        warningsQc.push(
          `Hoja ${nombreHoja} fila ${filaIdx}: el EXPTE ${expte} (${empresa}) ya lleva el concepto ` +
            `${concepto} en la hoja ${previo.hoja} fila ${previo.fila} — se factura dos veces este periodo`
        );
      } else {
        vistos.set(clave, { hoja: nombreHoja, fila: filaIdx });
      }

      conceptos.push({
        empresa: EMPRESA_FACTURADORA,
        codigo_cliente: pad5(clienteEfectivo),
        codigo_concepto: concepto,
        fecha: periodo.fecha,
        descripcion: buildDescripcion(mapeos.tarifas.describe(concepto), concepto, periodo),
        tipo_iva: TIPO_IVA,
        unidades: UNIDADES,
        importe_gastos: "",
        importe_honorarios: Math.round(tarifa * 100) / 100,
        codigo_expediente: codigoExpediente,
        descripcion_ampliada: buildDescAmpliada(empresa, nif),
      });
      stats.conceptos++;
    }

    hojas.push({ hoja: nombreHoja, fila_cabecera: headerRow, ...stats });
  }

  if (hojas.length === 0) {
    throw new Error(
      `No se encontró ninguna hoja de pagos en '${path.basename(inputPath)}'. ` +
        `El importador solo procesa hojas cuya celda A1 contenga '${SENTINEL}' ` +
        `(las genera procesos/importarFacturacion/pagos/migrarPlantilla.js).`
    );
  }

  writeConceptos(path.join(outputDir, "Conceptos Pendientes Facturar.csv"), conceptos);
  writeIncidencias(path.join(outputDir, "incidencias.csv"), incidencias);
  writeWarnings(path.join(outputDir, "warnings_qc.csv"), warningsQc);

  const total = conceptos.reduce((a, c) => a + c.importe_honorarios * c.unidades, 0);

  return {
    input: inputPath,
    output_dir: outputDir,
    periodo: {
      valor: _str(options.periodo).toUpperCase(),
      tipo: periodo.tipo,
      etiqueta: periodo.etiqueta,
      fecha_linea: isoDate(periodo.fecha),
      frecuencias_facturadas: [...facturables],
    },
    hojas,
    conceptos: conceptos.length,
    incidencias: incidencias.length,
    warnings_qc: warningsQc.length,
    importe_total: Math.round(total * 100) / 100,
  };
}

function writeConceptos(filePath, rows) {
  const hdr = [
    "Empresa Facturadora",
    "Código Cliente",
    "Cód. Concepto Facturable",
    "Fecha",
    "Descripción",
    "Tipo de IVA",
    "Unidades",
    "Importe Gastos",
    "Importe Honorarios",
    "Código Expediente",
    "Descripción Ampliada",
  ];
  const data = rows.map((r) => [
    r.empresa,
    r.codigo_cliente,
    r.codigo_concepto,
    isoDate(r.fecha),
    r.descripcion,
    r.tipo_iva,
    r.unidades,
    r.importe_gastos,
    r.importe_honorarios.toFixed(2),
    r.codigo_expediente,
    r.descripcion_ampliada,
  ]);
  writeCsv(filePath, hdr, data);
}

function writeIncidencias(filePath, rows) {
  const hdr = [
    "hoja",
    "fila_origen",
    "motivo",
    "CONCEPTO FACT",
    "EXPTE",
    "NIF",
    "EMPRESA",
    "FACTURAR",
    "FRECUENCIA",
  ];
  const data = rows.map((r) => [
    r.hoja,
    r.fila_origen,
    r.motivo,
    r.concepto,
    r.expte,
    r.nif,
    r.empresa,
    r.facturar,
    r.frecuencia,
  ]);
  writeCsv(filePath, hdr, data);
}

function writeWarnings(filePath, rows) {
  writeCsv(filePath, ["mensaje"], rows.map((m) => [m]));
}

module.exports = { transform, parsePeriodo, SENTINEL };
