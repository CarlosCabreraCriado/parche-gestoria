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
  locateHeaderTable,
  writeCsv,
} = require("./utils");

const EMPRESA_FACTURADORA = 14;
const CODIGO_CONCEPTO_DEFAULT = "3.010";
const TIPO_IVA = 3;

// Cabecera en dos filas: los nombres destino A3 (Código Cliente, Cód. Concepto
// Facturable, Fecha, Unidades…) están en la fila superior y los de las primeras
// columnas (Expediente, Cliente, Emisor, Asunto) en la inferior. Se localiza por
// nombre combinando ambas filas (mergeUp), sin depender de posiciones fijas.
const HEADER_SYNONYMS = {
  expediente: ["expediente", "exptcorto", "expt"],
  cliente: ["cliente", "razonsocial"],
  emisor: ["emisor"],
  f_lectura: ["flectura"],
  asunto: ["asunto"],
  emp_fact: ["empresafacturadora"],
  cod_cliente: ["codigocliente", "codcliente"],
  cod_concepto: ["codconceptofacturable", "codigoconceptofacturable", "codconcepto"],
  fecha: ["fecha"],
  descripcion: ["descripcion"],
  tipo_iva: ["tipodeiva", "tipoiva"],
  unidades: ["unidades"],
  importe_gastos: ["importegastos", "importe"],
  cod_expediente: ["codigoexpediente", "codexpediente"],
  desc_ampliada: ["descripcionampliada"],
};

// Cabecera válida: exige un campo que solo aparece en la fila inferior
// (cliente/expediente) y otro de la superior (cod_concepto) para no confundir la
// fila 1 sola con la cabecera y desplazar los datos.
function isNotificacionesHeader(cols) {
  return (
    (cols.cliente !== undefined || cols.expediente !== undefined) &&
    cols.cod_concepto !== undefined &&
    cols.fecha !== undefined
  );
}

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

async function transform(inputPath, mapeos, outputDir) {
  ensureDir(outputDir);

  const workbook = await XlsxPopulate.fromFileAsync(path.normalize(inputPath));
  const table = locateHeaderTable(workbook, HEADER_SYNONYMS, isNotificacionesHeader, {
    mergeUp: true,
  });
  if (!table) {
    throw new Error(
      `No se encontró la tabla de notificaciones en '${path.basename(inputPath)}'. ` +
        `Se requiere una hoja con cabeceras Cliente/Expediente, Cód. Concepto Facturable y Fecha.`
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

    const exptCortoStr = _str(get(row, "expediente"));
    const clienteRazon = _str(get(row, "cliente"));
    const emisor = _str(get(row, "emisor"));
    const asunto = _str(get(row, "asunto"));
    const codClienteRaw = _str(get(row, "cod_cliente"));
    const codExpedienteInput = _str(get(row, "cod_expediente"));
    const descAmpliadaRaw = _str(get(row, "desc_ampliada"));
    const concepto = normalizarConcepto(get(row, "cod_concepto"));

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

    const unidades = _toInt(get(row, "unidades")) || 1;
    const fechaLinea = _toDate(get(row, "fecha"), null);
    if (!fechaLinea) {
      addInc("Sin FECHA válida en el archivo");
      continue;
    }

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
    hoja: table.sheetName,
    fila_cabecera: headerRow,
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
