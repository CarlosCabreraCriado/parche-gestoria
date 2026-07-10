const path = require("path");
const XlsxPopulate = require("xlsx-populate");
const { REDIRECT_NADA } = require("./mapeos");
const {
  _str,
  _toInt,
  _toFloat,
  _toDate,
  pad5,
  isoDate,
  ensureDir,
  readAbsoluteRows,
  locateHeaderTable,
  writeCsv,
} = require("./utils");

const EMPRESA_FACTURADORA = 14;
const TIPO_IVA = 3;

// El cliente maneja dos formatos de archivo (p.ej. "VARIOS" en .xlsx y
// "Facturar" en .xlsm) con la MISMA información pero distinto orden de columnas
// y distinta hoja. En lugar de posiciones fijas, cada campo se localiza por el
// texto de su cabecera (normalizado). Añadir un sinónimo aquí basta para
// soportar una variante nueva.
const HEADER_SYNONYMS = {
  expt: ["expt"],
  empresa: ["empresa", "razonsocial"],
  nombre_trab: ["nombretrabajador", "nombretrab", "trabajador"],
  fecha: ["fecha"],
  tipo_tramite: ["tipotramite"],
  concepto: ["conceptofact", "concepto"],
  importe: ["importe"],
};

// Cabecera válida de trámites: EXPT + IMPORTE + (TIPO TRAMITE o CONCEPTO). Así se
// ignoran hojas auxiliares (p.ej. "Datos", "Conceptos" del .xlsm).
function isTramitesHeader(cols) {
  return (
    cols.expt !== undefined &&
    cols.importe !== undefined &&
    (cols.tipo_tramite !== undefined || cols.concepto !== undefined)
  );
}

function buildDescripcion(tipoTramite) {
  const base = (tipoTramite || "").replace(/[ \-.]+$/, "").trim() || "Trámite laboral";
  return base.slice(0, 250);
}

async function transform(inputPath, mapeos, outputDir) {
  ensureDir(outputDir);

  const workbook = await XlsxPopulate.fromFileAsync(path.normalize(inputPath));
  const table = locateHeaderTable(workbook, HEADER_SYNONYMS, isTramitesHeader);
  if (!table) {
    throw new Error(
      `No se encontró la tabla de trámites en '${path.basename(inputPath)}'. ` +
        `Se requiere una hoja con cabeceras EXPT, IMPORTE y TIPO TRAMITE/CONCEPTO FACT.`
    );
  }
  const { sheet, cols, headerRow } = table;
  const dataStartRow = headerRow + 1;
  const { rows: absRows } = readAbsoluteRows(sheet);

  const get = (row, field) =>
    cols[field] !== undefined ? row[cols[field]] : undefined;

  const conceptos = [];
  const incidencias = [];
  const warningsQc = [];

  for (const { rowIndex: filaIdx, cells } of absRows) {
    if (filaIdx < dataStartRow) continue;
    const row = cells;
    if (!row.length || row.every((c) => c === null || c === undefined)) continue;

    const empresaRaw = _str(get(row, "empresa"));
    const importeOrigen = _toFloat(get(row, "importe"));
    const exptCorto = _toInt(get(row, "expt"));
    const conceptoRaw = _str(get(row, "concepto"));
    const tipoTramite = _str(get(row, "tipo_tramite"));
    const nombreTrab = _str(get(row, "nombre_trab"));
    const fecha = _toDate(get(row, "fecha"), null);

    // Fila totalmente ruido
    if (!empresaRaw && importeOrigen === null && !conceptoRaw && !tipoTramite) continue;

    const addInc = (motivo) => {
      incidencias.push({
        fila_origen: filaIdx,
        motivo,
        empresa: empresaRaw,
        expt: exptCorto !== null ? String(exptCorto) : "",
        concepto: conceptoRaw,
        tipo_tramite: tipoTramite,
        importe: importeOrigen,
        nombre_trabajador: nombreTrab,
      });
    };

    if (exptCorto === null) {
      if (empresaRaw || importeOrigen || tipoTramite) {
        addInc("Sin código cliente en columna EXPT");
      }
      continue;
    }
    if (importeOrigen === null || importeOrigen <= 0) {
      if (tipoTramite || conceptoRaw) addInc("Sin IMPORTE");
      continue;
    }
    if (!conceptoRaw) {
      addInc("Sin CONCEPTO FACT");
      continue;
    }
    if (!fecha) {
      addInc("Sin FECHA válida en el archivo");
      continue;
    }

    // 1. Redirección
    const redirectTarget = mapeos.redirect.resolve(exptCorto);
    if (redirectTarget === REDIRECT_NADA) {
      addInc(`Cliente ${exptCorto} marcado como no facturable (NADA)`);
      continue;
    }
    const clienteEfectivo =
      typeof redirectTarget === "number" ? redirectTarget : exptCorto;

    // 2. Expediente
    const codigoExpediente = mapeos.exptes.resolve(clienteEfectivo);
    if (!codigoExpediente) {
      if (typeof redirectTarget === "number") {
        addInc(
          `Cliente destino ${clienteEfectivo} (redirect de ${exptCorto}) sin expediente en mapeo`
        );
      } else {
        addInc(`Cliente ${exptCorto} sin expediente formato B en mapeo`);
      }
      continue;
    }

    // 3. Tarifa
    const tarifa = mapeos.tarifas.resolve(conceptoRaw);
    if (tarifa === null) {
      const motivo = mapeos.tarifas.missReason(conceptoRaw);
      addInc(`Tarifa concepto '${conceptoRaw}' no resoluble: ${motivo}`);
      continue;
    }

    // 4. QC importe origen vs tarifa (unidades=1)
    if (Math.abs(importeOrigen - tarifa) > Math.max(0.01, tarifa * 0.01)) {
      warningsQc.push(
        `Fila ${filaIdx}: IMPORTE origen ${importeOrigen.toFixed(2)}€ != tarifa ${tarifa.toFixed(2)}€ (concepto ${conceptoRaw})`
      );
    }

    conceptos.push({
      empresa: EMPRESA_FACTURADORA,
      codigo_cliente: pad5(clienteEfectivo),
      codigo_concepto: conceptoRaw,
      fecha,
      descripcion: buildDescripcion(tipoTramite),
      tipo_iva: TIPO_IVA,
      unidades: 1,
      importe_gastos: "",
      importe_honorarios: Math.round(tarifa * 100) / 100,
      codigo_expediente: codigoExpediente,
      descripcion_ampliada: nombreTrab,
    });
  }

  writeConceptos(path.join(outputDir, "Conceptos Pendientes Facturar.csv"), conceptos);
  writeIncidencias(path.join(outputDir, "incidencias.csv"), incidencias);
  writeWarnings(path.join(outputDir, "warnings_qc.csv"), warningsQc);

  const total = conceptos.reduce((a, c) => a + c.importe_honorarios, 0);

  return {
    input: inputPath,
    output_dir: outputDir,
    hoja: table.sheetName,
    fila_cabecera: headerRow,
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
    "fila_origen",
    "motivo",
    "EMPRESA",
    "EXPT",
    "CONCEPTO",
    "TIPO TRAMITE",
    "IMPORTE",
    "NOMBRE TRABAJADOR",
  ];
  const data = rows.map((r) => [
    r.fila_origen,
    r.motivo,
    r.empresa,
    r.expt,
    r.concepto,
    r.tipo_tramite,
    r.importe,
    r.nombre_trabajador,
  ]);
  writeCsv(filePath, hdr, data);
}

function writeWarnings(filePath, rows) {
  writeCsv(filePath, ["mensaje"], rows.map((m) => [m]));
}

module.exports = { transform };
