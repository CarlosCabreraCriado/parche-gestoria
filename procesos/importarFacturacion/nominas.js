const path = require("path");
const XlsxPopulate = require("xlsx-populate");
const { REDIRECT_NADA } = require("./mapeos");
const {
  _str,
  _toInt,
  _toDate,
  leerImporte,
  conFecha,
  pad5,
  isoDate,
  ensureDir,
  readAbsoluteRows,
  locateHeaderTable,
  writeCsv,
} = require("./utils");

const EMPRESA_FACTURADORA = 14;
const CODIGO_CONCEPTO_DEFAULT = "0.010";
const TIPO_IVA = 3;

// Igual que en trámites: el archivo ya no trae la columna EXPT FACT y las
// columnas pueden variar de posición. Cada campo se localiza por el texto de su
// cabecera (normalizado). OBSERVACION lleva el nº de nóminas (unidades).
const HEADER_SYNONYMS = {
  expt: ["expt"],
  empresa: ["empresa", "razonsocial"],
  nombre_trab: ["nombretrabajador", "nombretrab", "trabajador"],
  fecha: ["fecha"],
  observacion: ["observacion"],
  tipo_tramite: ["tipotramite"],
  concepto: ["conceptofact", "concepto"],
  // Precio puntual POR UNIDAD (por nómina, no total de la fila). Su valor es
  // opcional: vacío = tarifa del catálogo. La columna sí se exige en la
  // cabecera porque es parte de la seña de identidad de la hoja de nóminas.
  importe: ["importe"],
};

// Cabecera válida de nóminas: EXPT + IMPORTE + CONCEPTO/TIPO TRAMITE.
function isNominasHeader(cols) {
  return (
    cols.expt !== undefined &&
    cols.importe !== undefined &&
    (cols.concepto !== undefined || cols.tipo_tramite !== undefined)
  );
}

// La fecha es la de la fila del archivo del cliente: ya no decide cuándo se
// factura (eso lo fija el formulario), solo documenta a qué día corresponde el
// trabajo. Si la fila no la trae, la descripción sale sin ella.
function buildDescripcion(tipoTramite, observacion, fecha) {
  const base = (tipoTramite || "").replace(/[ -]+$/, "").trim() || "Nóminas";
  const obsInt = _toInt(observacion);
  let texto = base;
  if (obsInt && obsInt > 0) {
    const plural = obsInt !== 1 ? "s" : "";
    texto = `${base} - ${obsInt} nómina${plural}`;
  } else {
    const obs = _str(observacion);
    if (obs) texto = `${base} - ${obs}`;
  }
  return conFecha(texto, fecha);
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
  const table = locateHeaderTable(workbook, HEADER_SYNONYMS, isNominasHeader);
  if (!table) {
    throw new Error(
      `No se encontró la tabla de nóminas en '${path.basename(inputPath)}'. ` +
        `Se requiere una hoja con cabeceras EXPT, IMPORTE y CONCEPTO FACT/TIPO TRAMITE.`
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
    const observacion = get(row, "observacion");
    const tipoTramite = _str(get(row, "tipo_tramite"));
    const conceptoCelda = _str(get(row, "concepto"));
    const conceptoRaw = conceptoCelda || CODIGO_CONCEPTO_DEFAULT;
    const nombreTrab = _str(get(row, "nombre_trab"));
    const fecha = _toDate(get(row, "fecha"), null);

    if (empresaRaw === "" && !tieneImporte && exptCorto === null) continue;

    const addInc = (motivo) => {
      incidencias.push({
        fila_origen: filaIdx,
        motivo,
        empresa: empresaRaw,
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
    // El concepto por defecto solo se aplica a filas que sean trabajo real. Una
    // fila con cliente pero sin concepto, sin trámite y sin importe es relleno
    // de la hoja: antes la descartaba el IMPORTE obligatorio y ahora, sin esa
    // guarda, se facturaría sola al concepto por defecto.
    if (!conceptoCelda && !tipoTramite && !tieneImporte) continue;
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

    const unidades = _toInt(observacion) || 1;

    // 3. Precio unitario: manda el IMPORTE de la fila; si viene vacío, la tarifa
    // del catálogo. El IMPORTE es por nómina, no el total de la fila: la línea
    // sale con Unidades = nº de nóminas y A3 multiplica. Con IMPORTE la fila se
    // factura aunque el concepto no tenga tarifa (ESCALADO o sin precio); sin
    // él, no hay de dónde sacar el importe.
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
          `Fila ${filaIdx}: IMPORTE puntual ${importeAplicado.toFixed(2)}€/ud != tarifa catálogo ${tarifaCatalogo.toFixed(2)}€/ud (concepto ${conceptoRaw}, ${unidades} uds) — se factura el puntual`
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

    // La fila se factura igual sin fecha; solo se pierde el dato en la
    // descripción. Se avisa aquí y no antes para no reportar filas que después
    // se descartan por otro motivo.
    if (!fecha) {
      warningsQc.push(
        `Fila ${filaIdx}: sin FECHA válida en el archivo — se factura igual, la descripción sale sin fecha`
      );
    }

    conceptos.push({
      empresa: EMPRESA_FACTURADORA,
      codigo_cliente: pad5(clienteEfectivo),
      codigo_concepto: conceptoRaw,
      fecha: fechaFactura,
      descripcion: buildDescripcion(tipoTramite, observacion, fecha),
      tipo_iva: TIPO_IVA,
      unidades,
      importe_gastos: "",
      importe_honorarios: Math.round(importeAplicado * 100) / 100,
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
    hoja: table.sheetName,
    fila_cabecera: headerRow,
    conceptos: conceptos.length,
    incidencias: incidencias.length,
    warnings_qc: warningsQc.length,
    importes_puntuales: importesPuntuales,
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
    "EXPT",
    "CONCEPTO",
    "IMPORTE",
    "NOMBRE TRABAJADOR",
  ];
  const data = rows.map((r) => [
    r.fila_origen,
    r.motivo,
    r.empresa,
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
