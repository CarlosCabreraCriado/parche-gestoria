const path = require("path");
const XlsxPopulate = require("xlsx-populate");
const { REDIRECT_NADA } = require("./mapeos");
const {
  _str,
  _toInt,
  _toDate,
  conFecha,
  recortarPorPalabra,
  LIMITE_DESC,
  LIMITE_DESC_AMPLIADA,
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

// El cliente manda dos formatos de listado y ambos deben funcionar:
//   - v1 ("FACT NOTIFICACIONES"): cabecera en DOS filas, con los nombres destino
//     A3 (Código Cliente, Cód. Concepto Facturable, Fecha, Unidades…) en la
//     superior y los del listado (Expediente, Cliente, Emisor, Asunto) en la
//     inferior. Se combinan con mergeUp.
//   - v2 ("Notificaciones v2"): export en crudo con UNA fila de cabecera y solo
//     las cinco columnas del listado (Expediente, Cliente, Emisor, F. Lectura,
//     Asunto). Todo lo demás (concepto, importe, expediente A3) lo pone este
//     proceso desde los mapeos y las constantes de arriba.
// En ambos casos se localiza por nombre, sin depender de posiciones fijas.
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

// Cabecera válida: quién es el cliente (cliente/expediente) MÁS algún dato
// propio del listado de notificaciones (emisor, asunto o F. Lectura). No se
// exigen las columnas A3 (cod_concepto, fecha) porque el formato v2 no las trae;
// pedirlas dejaba el proceso sin encontrar la tabla. Exigir la parte del listado
// evita además confundir la fila 1 sola de la v1 —que solo tiene nombres A3— con
// la cabecera y desplazar los datos.
function isNotificacionesHeader(cols) {
  const identificaCliente =
    cols.cliente !== undefined || cols.expediente !== undefined;
  const esListadoNotificaciones =
    cols.asunto !== undefined ||
    cols.emisor !== undefined ||
    cols.f_lectura !== undefined;
  return identificaCliente && esListadoNotificaciones;
}

function normalizarConcepto(raw) {
  const s = _str(raw);
  if (!s) return CODIGO_CONCEPTO_DEFAULT;
  if (/^\d+$/.test(s) && s.length >= 4) {
    return `${s.slice(0, -3)}.${s.slice(-3)}`;
  }
  return s;
}

// La descripción que ve el cliente en la factura es fija: prefijo + fecha. El
// asunto ya no entra aquí porque llega del portal en crudo —códigos ilegibles
// del tipo "REGIMENES SEG. SOCIAL-NOT.DEUDOR.DIL.LEV.EMB.", interrogantes de
// una exportación mal codificada y textos de 230 caracteres que se truncaban a
// mitad de palabra—; su sitio es la descripción ampliada.
// La fecha es la de la fila del archivo del cliente (F. LECTURA): no decide
// cuándo se factura (eso lo fija el formulario), solo documenta a qué día
// corresponde el aviso. Si la fila no la trae, la descripción sale sin ella.
function buildDescripcion(fecha) {
  // Fija y corta ("Aviso Notificación - dd/mm/aaaa", ~31); el recorte a 50 solo
  // blinda el límite de A3, igual que en el resto de importadores.
  return recortarPorPalabra(conFecha("Aviso Notificación", fecha), LIMITE_DESC);
}

// Aquí va el detalle. Importa más que antes: como la descripción corta ya no
// lleva asunto, dos avisos del mismo cliente el mismo día salen con líneas
// idénticas en A3 y esta columna es lo único que permite distinguirlos. Se
// devuelve entera y la recorta quien llama, que así puede avisar del recorte.
function buildDescAmpliada(descAmpliadaRaw, asunto, emisor) {
  if (descAmpliadaRaw) return descAmpliadaRaw;
  return [asunto, emisor].filter(Boolean).join(", ");
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
  const table = locateHeaderTable(workbook, HEADER_SYNONYMS, isNotificacionesHeader, {
    mergeUp: true,
    fuzzy: true,
  });
  if (!table) {
    throw new Error(
      `No se encontró la tabla de notificaciones en '${path.basename(inputPath)}'. ` +
        `Se requiere una hoja con cabeceras Cliente o Expediente y, al menos, ` +
        `Asunto, Emisor o F. Lectura.`
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
  // Para detectar avisos repetidos. Se agrupa por el dato de ORIGEN (cliente +
  // asunto + día), no por la descripción ya construida: como esta se quedó en
  // prefijo + fecha, compararla marcaría cualquier par del mismo día y el aviso
  // sería ruido. Sobre el origen señala justo lo que hay que mirar, que el
  // listado del portal traiga dos veces la misma notificación.
  const posiblesRepetidos = new Map();

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
    // La v1 duplica la fecha en una columna FECHA propia de A3; la v2 solo trae
    // F. LECTURA. Son el mismo dato, así que vale cualquiera de las dos.
    const fechaLinea = _toDate(get(row, "fecha") ?? get(row, "f_lectura"), null);

    // La fila se factura igual sin fecha; solo se pierde el dato en la
    // descripción. Se avisa aquí y no antes para no reportar filas que después
    // se descartan por otro motivo.
    if (!fechaLinea) {
      warningsQc.push(
        `Fila ${filaIdx}: sin FECHA / F. LECTURA válida en el archivo — se factura igual, la descripción sale sin fecha`
      );
    }

    const ampliada = buildDescAmpliada(descAmpliadaRaw, asunto, emisor);
    if (ampliada.length > LIMITE_DESC_AMPLIADA) {
      warningsQc.push(
        `Fila ${filaIdx}: descripción ampliada de ${ampliada.length} caracteres — se recorta a ${LIMITE_DESC_AMPLIADA}`
      );
    }

    const clienteFinal = pad5(clienteEfectivo);
    const diaAviso = fechaLinea ? isoDate(fechaLinea) : "sin fecha";
    const claveRepetido = `${clienteFinal}||${asunto.toLowerCase()}||${diaAviso}`;
    const grupo = posiblesRepetidos.get(claveRepetido);
    if (grupo) grupo.filas.push(filaIdx);
    else {
      posiblesRepetidos.set(claveRepetido, {
        filas: [filaIdx],
        cliente: clienteFinal,
        asunto,
        dia: diaAviso,
      });
    }

    conceptos.push({
      empresa: EMPRESA_FACTURADORA,
      codigo_cliente: clienteFinal,
      codigo_concepto: concepto,
      fecha: fechaFactura,
      descripcion: buildDescripcion(fechaLinea),
      tipo_iva: TIPO_IVA,
      unidades,
      importe_gastos: "",
      importe_honorarios: Math.round(tarifa * 100) / 100,
      codigo_expediente: codigoExpediente,
      descripcion_ampliada: ampliada.slice(0, LIMITE_DESC_AMPLIADA),
    });
  }

  // Se avisa al final y no dentro del bucle porque hasta no recorrerlo entero no
  // se sabe cuántas veces aparece cada aviso. No es un descarte: las líneas se
  // facturan igual, solo se marcan para que alguien las mire antes de emitir.
  for (const g of posiblesRepetidos.values()) {
    if (g.filas.length < 2) continue;
    const asuntoCorto =
      g.asunto.length > 80 ? `${g.asunto.slice(0, 80)}…` : g.asunto;
    warningsQc.push(
      `Cliente ${g.cliente}: ${g.filas.length} líneas con el mismo asunto '${asuntoCorto}' el ${g.dia} ` +
        `(filas ${g.filas.join(", ")}) — se facturan todas; revisar si son avisos distintos o el listado repite la notificación`
    );
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
