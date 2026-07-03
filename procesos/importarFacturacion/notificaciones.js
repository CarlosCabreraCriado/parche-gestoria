const path = require("path");
const XlsxPopulate = require("xlsx-populate");
const { REDIRECT_NADA } = require("./mapeos");
const {
  _str,
  _toInt,
  _toDate,
  pad5,
  isoDate,
  ensureDir,
  readAbsoluteRows,
  writeCsv,
} = require("./utils");

const EMPRESA_FACTURADORA = 14;
const CODIGO_CONCEPTO_DEFAULT = "3.010";
const TIPO_IVA = 3;
const DATA_START_ROW = 3;

const COLS = {
  EXPT_CORTO: 0,
  CLIENTE: 1,
  EMISOR: 2,
  F_LECTURA: 3,
  ASUNTO: 4,
  EMP_FACT: 5,
  COD_CLIENTE: 6,
  COD_CONCEPTO: 7,
  FECHA: 8,
  DESCRIPCION: 9,
  TIPO_IVA: 10,
  UNIDADES: 11,
  IMPORTE_RAW: 12,
  COD_EXPEDIENTE: 13,
  DESC_AMPLIADA: 14,
};

function normalizarConcepto(raw) {
  const s = _str(raw);
  if (!s) return CODIGO_CONCEPTO_DEFAULT;
  if (/^\d+$/.test(s) && s.length >= 4) {
    return `${s.slice(0, -3)}.${s.slice(-3)}`;
  }
  return s;
}

function buildDescripcion(asunto, fecha) {
  const fechaStr = `${String(fecha.getDate()).padStart(2, "0")}/${String(
    fecha.getMonth() + 1
  ).padStart(2, "0")}/${fecha.getFullYear()}`;
  let base = "Aviso Notificacion";
  const a = (asunto || "").trim();
  if (a) base = `${base} - ${a}`;
  return `${base} - ${fechaStr}`.slice(0, 250);
}

function buildDescAmpliada(descAmpliadaRaw, asunto, emisor) {
  if (descAmpliadaRaw) return descAmpliadaRaw.slice(0, 500);
  const parts = ["Aviso Notificacion"];
  if (asunto) parts.push(asunto);
  if (emisor) parts.push(emisor);
  return parts.join(", ").slice(0, 500);
}

async function transform(inputPath, mapeos, outputDir, fechaDefault) {
  const fecha0 = fechaDefault || new Date();
  ensureDir(outputDir);

  const workbook = await XlsxPopulate.fromFileAsync(path.normalize(inputPath));
  const sheetNames = workbook.sheets().map((s) => s.name());
  const sheetName = sheetNames.includes("Hoja1") ? "Hoja1" : sheetNames[0];
  const sheet = workbook.sheet(sheetName);
  const { rows: absRows } = readAbsoluteRows(sheet);

  const conceptos = [];
  const incidencias = [];
  const warningsQc = [];

  for (const { rowIndex: filaIdx, cells } of absRows) {
    if (filaIdx < DATA_START_ROW) continue;
    const row = cells;
    if (!row.length || row.every((c) => c === null || c === undefined)) continue;

    const exptCortoStr = _str(row[COLS.EXPT_CORTO]);
    const clienteRazon = _str(row[COLS.CLIENTE]);
    const emisor = _str(row[COLS.EMISOR]);
    const asunto = _str(row[COLS.ASUNTO]);
    const codClienteRaw = _str(row[COLS.COD_CLIENTE]);
    const codExpedienteInput = _str(row[COLS.COD_EXPEDIENTE]);
    const descAmpliadaRaw = _str(row[COLS.DESC_AMPLIADA]);
    const concepto = normalizarConcepto(row[COLS.COD_CONCEPTO]);

    if (!(exptCortoStr || clienteRazon || codExpedienteInput || codClienteRaw)) continue;

    const clienteInput = _toInt(codClienteRaw) ?? _toInt(exptCortoStr);

    const addInc = (motivo) => {
      incidencias.push({
        fila_origen: filaIdx,
        motivo,
        cliente: clienteRazon,
        expt_corto: exptCortoStr,
        cod_cliente: codClienteRaw,
        concepto,
        cod_expediente: codExpedienteInput,
        asunto,
      });
    };

    if (clienteInput === null) {
      addInc("Sin código cliente numérico (columnas EXPT/COD_CLIENTE)");
      continue;
    }

    // 1. Redirección
    const redirectTarget = mapeos.redirect.resolve(clienteInput);
    if (redirectTarget === REDIRECT_NADA) {
      addInc(`Cliente ${clienteInput} marcado como no facturable (NADA)`);
      continue;
    }
    const clienteEfectivo =
      typeof redirectTarget === "number" ? redirectTarget : clienteInput;

    // 2. Expediente
    const codigoExpediente = mapeos.exptes.resolve(clienteEfectivo);
    if (!codigoExpediente) {
      if (typeof redirectTarget === "number") {
        addInc(
          `Cliente destino ${clienteEfectivo} (redirect de ${clienteInput}) sin expediente en mapeo`
        );
      } else {
        addInc(`Cliente ${clienteInput} sin expediente formato B en mapeo`);
      }
      continue;
    }

    // 3. QC vs expte del archivo
    if (
      codExpedienteInput &&
      codExpedienteInput !== codigoExpediente &&
      redirectTarget === null
    ) {
      warningsQc.push(
        `Fila ${filaIdx}: COD EXPEDIENTE del input '${codExpedienteInput}' difiere del mapeo '${codigoExpediente}' — se usa el mapeo`
      );
    }

    // 4. Tarifa
    const tarifa = mapeos.tarifas.resolve(concepto);
    if (tarifa === null) {
      const motivo = mapeos.tarifas.missReason(concepto);
      addInc(`Tarifa concepto '${concepto}' no resoluble: ${motivo}`);
      continue;
    }

    const unidades = _toInt(row[COLS.UNIDADES]) || 1;
    const fechaLinea = _toDate(row[COLS.FECHA], fecha0);

    conceptos.push({
      empresa: EMPRESA_FACTURADORA,
      codigo_cliente: pad5(clienteEfectivo),
      codigo_concepto: concepto,
      fecha: fechaLinea,
      descripcion: buildDescripcion(asunto, fechaLinea),
      tipo_iva: TIPO_IVA,
      unidades,
      importe_gastos: "",
      importe_honorarios: Math.round(tarifa * 100) / 100,
      codigo_expediente: codigoExpediente,
      descripcion_ampliada: buildDescAmpliada(descAmpliadaRaw, asunto, emisor),
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
    "CLIENTE",
    "EXPT CORTO",
    "COD CLIENTE",
    "CONCEPTO",
    "COD EXPEDIENTE",
    "ASUNTO",
  ];
  const data = rows.map((r) => [
    r.fila_origen,
    r.motivo,
    r.cliente,
    r.expt_corto,
    r.cod_cliente,
    r.concepto,
    r.cod_expediente,
    r.asunto,
  ]);
  writeCsv(filePath, hdr, data);
}

function writeWarnings(filePath, rows) {
  writeCsv(filePath, ["mensaje"], rows.map((m) => [m]));
}

module.exports = { transform };
