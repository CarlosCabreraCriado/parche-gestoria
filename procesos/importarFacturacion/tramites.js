const path = require("path");
const XlsxPopulate = require("xlsx-populate");
const { REDIRECT_NADA } = require("./mapeos");
const {
  _str,
  _toInt,
  _toDate,
  leerImporte,
  fechaCorta,
  repartirDescripcion,
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
  // Variantes reales del cliente: singular, plural y abreviatura. Además, con
  // `fuzzy` (ver locateHeaderTable) "observacion" capta por prefijo cualquier
  // "OBSERVACION…" que no esté listada; "obs" se pone explícito porque es más
  // corta que la raíz mínima y el nivel por prefijo no la alcanzaría.
  observacion: ["observacion", "observaciones", "obs"],
  tipo_tramite: ["tipotramite"],
  concepto: ["conceptofact", "concepto"],
  // Precio puntual de la fila. Su VALOR es opcional: vacío = tarifa del
  // catálogo, con número = manda ese importe. La columna sí se exige en la
  // cabecera porque es parte de la seña de identidad de la hoja de trámites.
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

// A3 solo guarda 50 caracteres en la Descripción, así que ahí va SOLO el tipo de
// trámite (recortado por palabra) y todo lo demás se conserva en la Descripción
// Ampliada: el texto íntegro del trámite si no cupo en 50, el nombre del
// trabajador, la OBSERVACION —detalle concreto del trámite ("B Vol - Pte Ss +
// Cert") que el TIPO TRAMITE no recoge— y la fecha de la fila.
// La fecha ya no va en la Descripción: A3 trae su propia columna Fecha en la
// línea y ese hueco se necesita para el concepto. La OBSERVACION no se repite si
// dice lo mismo que el tipo de trámite.
function buildLinea(tipoTramite, observacion, nombreTrab, fecha) {
  const concepto =
    _str(tipoTramite).replace(/[ \-.]+$/, "").trim() || "Trámite laboral";
  const obs = _str(observacion).replace(/[ \-.]+$/, "").trim();
  const extras = [nombreTrab];
  if (obs && obs.toLowerCase() !== concepto.toLowerCase()) extras.push(obs);
  if (fecha) extras.push(fechaCorta(fecha));
  return repartirDescripcion(concepto, extras);
}

async function transform(inputPath, mapeos, outputDir, options = {}) {
  ensureDir(outputDir);
  // Sin ella todas las líneas saldrían sin Fecha y el fallo no aparecería hasta
  // escribir el CSV, como un críptico error de getFullYear.
  const fechaFactura = options.fechaFactura;
  if (!fechaFactura) {
    throw new Error(
      "Falta la fecha de facturación: la elige el usuario en el formulario y es obligatoria."
    );
  }

  const workbook = await XlsxPopulate.fromFileAsync(path.normalize(inputPath));
  const table = locateHeaderTable(workbook, HEADER_SYNONYMS, isTramitesHeader, {
    fuzzy: true,
  });
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
  let importesPuntuales = 0;

  for (const { rowIndex: filaIdx, cells } of absRows) {
    if (filaIdx < dataStartRow) continue;
    const row = cells;
    if (!row.length || row.every((c) => c === null || c === undefined)) continue;

    const empresaRaw = _str(get(row, "empresa"));
    // IMPORTE es un precio puntual opcional, no un dato obligatorio: lo normal
    // es que venga vacío y se facture la tarifa del catálogo.
    const importePuntual = leerImporte(get(row, "importe"));
    const tieneImporte =
      importePuntual.valor !== null || importePuntual.error !== undefined;
    const importeOrigen = importePuntual.error ?? importePuntual.valor;
    const exptCorto = _toInt(get(row, "expt"));
    const conceptoRaw = _str(get(row, "concepto"));
    const tipoTramite = _str(get(row, "tipo_tramite"));
    const observacion = get(row, "observacion");
    const nombreTrab = _str(get(row, "nombre_trab"));
    const fecha = _toDate(get(row, "fecha"), null);

    // Fila totalmente ruido
    if (!empresaRaw && !tieneImporte && !conceptoRaw && !tipoTramite) continue;

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
      if (empresaRaw || tieneImporte || tipoTramite) {
        addInc("Sin código cliente en columna EXPT");
      }
      continue;
    }
    // Sin concepto no hay nada que facturar. Las filas que además vienen sin
    // trámite ni importe son relleno de la hoja y se saltan sin incidencia.
    if (!conceptoRaw) {
      if (tipoTramite || tieneImporte) addInc("Sin CONCEPTO FACT");
      continue;
    }
    // Un IMPORTE ilegible no cae a la tarifa: facturaría algo distinto de lo
    // que el usuario quiso teclear y nadie lo notaría.
    if (importePuntual.error !== undefined) {
      addInc(
        `IMPORTE '${importePuntual.error}' no es un número válido — corrige la celda o déjala vacía`
      );
      continue;
    }
    if (importePuntual.valor !== null && importePuntual.valor <= 0) {
      addInc(`IMPORTE ${importePuntual.valor} no válido: debe ser mayor que 0`);
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

    // 3. Precio: manda el IMPORTE de la fila; si viene vacío, la tarifa del
    // catálogo. Con IMPORTE la fila se factura aunque el concepto no tenga
    // tarifa (ESCALADO o sin precio); sin él, no hay de dónde sacar el importe.
    const tarifaCatalogo = mapeos.tarifas.resolve(conceptoRaw);
    let importeAplicado;
    if (importePuntual.valor !== null) {
      importeAplicado = importePuntual.valor;
      importesPuntuales++;
      // Se avisa solo de la discrepancia: es el caso que merece revisión.
      if (
        tarifaCatalogo !== null &&
        Math.abs(importeAplicado - tarifaCatalogo) > Math.max(0.01, tarifaCatalogo * 0.01)
      ) {
        warningsQc.push(
          `Fila ${filaIdx}: IMPORTE puntual ${importeAplicado.toFixed(2)}€ != tarifa catálogo ${tarifaCatalogo.toFixed(2)}€ (concepto ${conceptoRaw}) — se factura el puntual`
        );
      }
    } else {
      if (tarifaCatalogo === null) {
        const motivo = mapeos.tarifas.missReason(conceptoRaw);
        addInc(
          `Tarifa concepto '${conceptoRaw}' no resoluble: ${motivo} — rellena IMPORTE en la fila para facturarla`
        );
        continue;
      }
      importeAplicado = tarifaCatalogo;
    }

    // La fila se factura igual sin fecha; solo se pierde ese dato en la
    // descripción ampliada. Se avisa aquí y no antes para no reportar filas que
    // después se descartan por otro motivo.
    if (!fecha) {
      warningsQc.push(
        `Fila ${filaIdx}: sin FECHA válida en el archivo — se factura igual, la ampliada sale sin la fecha del trámite`
      );
    }

    const linea = buildLinea(tipoTramite, observacion, nombreTrab, fecha);
    conceptos.push({
      empresa: EMPRESA_FACTURADORA,
      codigo_cliente: pad5(clienteEfectivo),
      codigo_concepto: conceptoRaw,
      fecha: fechaFactura,
      descripcion: linea.descripcion,
      tipo_iva: TIPO_IVA,
      unidades: 1,
      importe_gastos: "",
      importe_honorarios: Math.round(importeAplicado * 100) / 100,
      codigo_expediente: codigoExpediente,
      descripcion_ampliada: linea.ampliada,
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
    importes_puntuales: importesPuntuales,
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
