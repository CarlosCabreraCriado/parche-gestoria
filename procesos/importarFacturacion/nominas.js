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
  writeCsv,
} = require("./utils");

const EMPRESA_FACTURADORA = 14;
const CODIGO_CONCEPTO_DEFAULT = "0.010";
const TIPO_IVA = 3;
const DATA_START_ROW = 8;

const COLS = {
  EXPT_FACT: 0,
  EXPT: 1,
  EMPRESA: 2,
  DNI_TRAB: 3,
  NOMBRE_TRAB: 4,
  FECHA: 8,
  OBSERVACION: 9,
  TIPO_TRAMITE: 11,
  CONCEPTO_FACT: 12,
  IMPORTE: 13,
};

function buildDescripcion(tipoTramite, observacion) {
  const base = (tipoTramite || "").replace(/[ -]+$/, "").trim() || "Nóminas";
  const obsInt = _toInt(observacion);
  if (obsInt && obsInt > 0) {
    const plural = obsInt !== 1 ? "s" : "";
    return `${base} - ${obsInt} nómina${plural}`.slice(0, 250);
  }
  const obs = _str(observacion);
  if (obs) return `${base} - ${obs}`.slice(0, 250);
  return base.slice(0, 250);
}

async function transform(inputPath, mapeos, outputDir, fechaDefault) {
  const fecha0 = fechaDefault || new Date();
  ensureDir(outputDir);

  const workbook = await XlsxPopulate.fromFileAsync(path.normalize(inputPath));
  const sheet = workbook.sheets()[0];
  const { rows: absRows } = readAbsoluteRows(sheet);

  const conceptos = [];
  const incidencias = [];
  const warningsQc = [];

  for (const { rowIndex: filaIdx, cells } of absRows) {
    if (filaIdx < DATA_START_ROW) continue;
    const row = cells;
    if (!row.length || row.every((c) => c === null || c === undefined)) continue;

    const empresaRaw = _str(row[COLS.EMPRESA]);
    const importeOrigen = _toFloat(row[COLS.IMPORTE]);
    const exptFactInput = _str(row[COLS.EXPT_FACT]);
    const exptCorto = _toInt(row[COLS.EXPT]);
    const observacion = row[COLS.OBSERVACION];
    const tipoTramite = _str(row[COLS.TIPO_TRAMITE]);
    const conceptoRaw = _str(row[COLS.CONCEPTO_FACT]) || CODIGO_CONCEPTO_DEFAULT;
    const nombreTrab = _str(row[COLS.NOMBRE_TRAB]);
    const fecha = _toDate(row[COLS.FECHA], fecha0);

    if (empresaRaw === "" && importeOrigen === null && exptCorto === null) continue;

    const addInc = (motivo) => {
      incidencias.push({
        fila_origen: filaIdx,
        motivo,
        empresa: empresaRaw,
        expt_fact: exptFactInput,
        expt: exptCorto !== null ? String(exptCorto) : "",
        concepto: conceptoRaw,
        importe: importeOrigen,
        nombre_trabajador: nombreTrab,
      });
    };

    if (exptCorto === null) {
      addInc("Sin código cliente en columna EXPT");
      continue;
    }
    if (importeOrigen === null || importeOrigen <= 0) {
      addInc("Sin IMPORTE");
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

    // 2. Expediente formato B
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

    // 3. QC: expt_fact input vs mapeo
    if (
      exptFactInput &&
      exptFactInput !== codigoExpediente &&
      redirectTarget === null
    ) {
      warningsQc.push(
        `Fila ${filaIdx}: EXPT FACT del input '${exptFactInput}' difiere del mapeo '${codigoExpediente}' — se usa el mapeo`
      );
    }

    // 4. Tarifa
    const tarifa = mapeos.tarifas.resolve(conceptoRaw);
    if (tarifa === null) {
      const motivo = mapeos.tarifas.missReason(conceptoRaw);
      addInc(`Tarifa concepto '${conceptoRaw}' no resoluble: ${motivo}`);
      continue;
    }

    const unidades = _toInt(observacion) || 1;

    // 5. QC importe origen vs tarifa × unidades
    const esperado = Math.round(tarifa * unidades * 100) / 100;
    if (Math.abs(importeOrigen - esperado) > Math.max(0.01, esperado * 0.01)) {
      warningsQc.push(
        `Fila ${filaIdx}: IMPORTE origen ${importeOrigen.toFixed(2)}€ != tarifa×uds ${esperado.toFixed(2)}€ (concepto ${conceptoRaw}, ${unidades} uds × ${tarifa.toFixed(2)}€)`
      );
    }

    conceptos.push({
      empresa: EMPRESA_FACTURADORA,
      codigo_cliente: pad5(clienteEfectivo),
      codigo_concepto: conceptoRaw,
      fecha,
      descripcion: buildDescripcion(tipoTramite, observacion),
      tipo_iva: TIPO_IVA,
      unidades,
      importe_gastos: "",
      importe_honorarios: Math.round(tarifa * 100) / 100,
      codigo_expediente: codigoExpediente,
      descripcion_ampliada: nombreTrab,
    });
  }

  writeConceptos(path.join(outputDir, "Conceptos Pendientes Facturar.csv"), conceptos);
  writeIncidencias(path.join(outputDir, "incidencias.csv"), incidencias);
  writeWarnings(path.join(outputDir, "warnings_qc.csv"), warningsQc);

  const totalUnit = conceptos.reduce((a, c) => a + c.importe_honorarios, 0);
  const totalEfectivo = conceptos.reduce(
    (a, c) => a + c.importe_honorarios * c.unidades,
    0
  );

  return {
    input: inputPath,
    output_dir: outputDir,
    conceptos: conceptos.length,
    incidencias: incidencias.length,
    warnings_qc: warningsQc.length,
    importe_total_unitario: Math.round(totalUnit * 100) / 100,
    importe_total_efectivo: Math.round(totalEfectivo * 100) / 100,
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
    "EXPT FACT",
    "EXPT",
    "CONCEPTO",
    "IMPORTE",
    "NOMBRE TRABAJADOR",
  ];
  const data = rows.map((r) => [
    r.fila_origen,
    r.motivo,
    r.empresa,
    r.expt_fact,
    r.expt,
    r.concepto,
    r.importe,
    r.nombre_trabajador,
  ]);
  writeCsv(filePath, hdr, data);
}

function writeWarnings(filePath, rows) {
  writeCsv(filePath, ["mensaje"], rows.map((m) => [m]));
}

module.exports = { transform };
