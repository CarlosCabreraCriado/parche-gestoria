const fs = require("fs");
const path = require("path");
const { simpleParser } = require("mailparser");
//const cheerio = require("cheerio");
const nodemailer = require("nodemailer");

//  Utilidades 
function excelSerialToDate(serial) {
  if (serial === undefined || serial === null || serial === "") return null;
  const ms = Math.round((serial - 25569) * 86400 * 1000);
  if (Number.isNaN(ms)) return null;
  return new Date(ms);
}
function formatDateFromExcel(serial) {
  const d = excelSerialToDate(serial);
  if (!d) return "";
  const dd = String(d.getDate()).padStart(2, "0");
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const yyyy = d.getFullYear();
  return `${dd}/${mm}/${yyyy}`;
}
function safe(v, fallback = "") {
  return v === undefined || v === null ? fallback : String(v);
}
function safeFilename(str, max = 120) {
  const s = safe(str, "archivo")
    .replace(/[<>:"/\\|?*\x00-\x1F]/g, "_")
    .slice(0, max);
  return s || "archivo";
}
function escapeHtml(str) {
  return safe(str, "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function normalizeContingencia(raw) {
  return safe(raw, "").replace(/^\d+=/, "");
}

function buildFilasITHtml(records, tipoDoc) {
  return records
    .map((r) => {
      const nombre = escapeHtml(r.nombre || "");
      const contingencia = escapeHtml(normalizeContingencia(r.contingencia));
      const fechaBaja = escapeHtml(formatDateFromExcel(r.fechaBajaIt) || "");
      const fechaAlta = escapeHtml(formatDateFromExcel(r.fechaFinIt) || "");

      let fechaConfirmacion = "";
      let proximaRevision = "";

      if (Array.isArray(r.partesConfirmacion) && r.partesConfirmacion.length > 0) {
        fechaConfirmacion =
          formatDateFromExcel(r.partesConfirmacion[0].fechaDelParteDeConfirmacion) || "";
        proximaRevision =
          formatDateFromExcel(r.partesConfirmacion[0].fechaSiguienteRevisionMedica) || "";
      } else if (tipoDoc === "BAJAS") {
        proximaRevision = formatDateFromExcel(r.fechaProximaRevisionParteBaja) || "";
      }

      const tipoProceso = escapeHtml(calcularTipoProceso(r));

      let observaciones = "";
      if (tipoDoc === "ALTAS") {
        observaciones = `PARTE DE ALTA MÉDICA.<br/>${tipoProceso}`;
      } else if (tipoDoc === "BAJAS") {
        observaciones = `PARTE DE BAJA MÉDICA.<br/>${tipoProceso}`;
      } else if (tipoDoc === "CONFIRMACION") {
        let num = 0;
        if (Array.isArray(r.partesConfirmacion) && r.partesConfirmacion.length > 0) {
          num = r.partesConfirmacion[0].numeroDeParteDeConfirmacion || 0;
        }
        observaciones = `PARTE DE CONFIRMACIÓN Nº${escapeHtml(String(num))}<br/>${tipoProceso}`;
      }

      // Mantengo estilos compatibles con Outlook y coherentes con tu cabecera
      return `
<tr>
  <td style="border:1px solid #000;padding:4px;font-size:12px;line-height:14px;overflow-wrap:anywhere;">${nombre}</td>
  <td style="border:1px solid #000;padding:4px;font-size:12px;line-height:14px;overflow-wrap:anywhere;">${contingencia}</td>
  <td style="border:1px solid #000;padding:4px;font-size:12px;line-height:14px;overflow-wrap:anywhere;">${fechaBaja}</td>
  <td style="border:1px solid #000;padding:4px;font-size:12px;line-height:14px;overflow-wrap:anywhere;">${escapeHtml(fechaConfirmacion)}</td>
  <td style="border:1px solid #000;padding:4px;font-size:12px;line-height:14px;overflow-wrap:anywhere;">${escapeHtml(proximaRevision)}</td>
  <td style="border:1px solid #000;padding:4px;font-size:12px;line-height:14px;overflow-wrap:anywhere;">${fechaAlta}</td>
  <td style="border:1px solid #000;padding:4px;font-size:12px;line-height:14px;overflow-wrap:anywhere;">${observaciones}</td>
</tr>`.trim();
    })
    .join("\n");
}

//  Config 
const TEMPLATE_EML_ALTA = path.join(__dirname, "fie", "alta.eml"); // tu plantilla
const TEMPLATE_EML_BAJA = path.join(__dirname, "fie", "baja.eml"); // tu plantilla
const TEMPLATE_HTML_CONFIRMACION = path.join(__dirname, "fie", "confirmacion.html");


// Puedes fijar un FROM/TO/CC por defecto si quieres sobreescribir los de la plantilla
const DEFAULT_FROM = null;
const DEFAULT_TO = null;
const DEFAULT_CC = null;

// Construye asunto nuevo a partir del record
function buildSubject(record, tipoDoc) {
  var fechaRecep = formatDateFromExcel(record.fechaRecepcion);
  switch (tipoDoc) {
    case "BAJAS":
      fechaRecep = formatDateFromExcel(record.fechaBajaIt) || fechaRecep;
      return `${safe(record.expte)} - COMUNICACIÓN IT - ${safe(record.nombre)} - PB ${fechaRecep}`;
    case "ALTAS":
      fechaRecep = formatDateFromExcel(record.fechaFinIt) || fechaRecep;
      return `${safe(record.expte)} - COMUNICACIÓN IT - ${safe(record.nombre)} - PA ${fechaRecep}`;
    case "CONFIRMACION":
      var parteConfirmacion = 0;
      if (Array.isArray(record.partesConfirmacion)) {
        fechaRecep =
          formatDateFromExcel(
            record.partesConfirmacion[0].fechaDelParteDeConfirmacion,
          ) || fechaRecep;
        parteConfirmacion =
          record.partesConfirmacion[0].numeroDeParteDeConfirmacion || 0;
      }
      return `${safe(record.expte)} - COMUNICACIÓN IT - ${safe(record.nombre)} - PC${parteConfirmacion} ${fechaRecep}`;
  }
  return `null`;
}

// util fechas para "fecha actual" 
function formatTodayDDMMYYYY() {
  const d = new Date();
  const dd = String(d.getDate()).padStart(2, "0");
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const yyyy = d.getFullYear();
  return `${dd}/${mm}/${yyyy}`;
}

function formatTodayDDMMYYYY_noSlash() {
  const d = new Date();
  const dd = String(d.getDate()).padStart(2, "0");
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const yyyy = d.getFullYear();
  return `${dd}${mm}${yyyy}`;
}

// calcula tipo de proceso (igual que en el individual) 
function calcularTipoProceso(record) {
  var tipoProceso = "";
  const tipoProcStr = safe(record.tipoDeProceso, "");
  const primeraLetra = (tipoProcStr && tipoProcStr[0]) || null;

  switch (Number(primeraLetra)) {
    case 1:
      tipoProceso = "PROCESO MUY CORTO";
      break;
    case 2:
      tipoProceso = "PROCESO CORTO";
      break;
    case 3:
      tipoProceso = "PROCESO INTERMEDIO";
      break;
    case 4:
      tipoProceso = "PROCESO LARGO";
      break;
    default:
      tipoProceso = "";
      break;
  }
  return tipoProceso;
}

// reemplaza placeholders en una fila <tr> con un record concreto 
function personalizarFilaHtml(filaHtml, record, tipoDoc) {
  let row = filaHtml;

  row = row.replaceAll(/{{\s*nombre\s*}}/g, record.nombre || "");
  row = row.replaceAll(
    /{{\s*fechaAlta\s*}}/g,
    formatDateFromExcel(record.fechaFinIt) || "",
  );
  row = row.replaceAll(
    /{{\s*fechaBaja\s*}}/g,
    formatDateFromExcel(record.fechaBajaIt) || "",
  );
  row = row.replaceAll(/{{\s*contingencia\s*}}/g, record.contingencia || "");

  // Próxima revisión: depende de BAJAS o CONFIRMACION
  if (tipoDoc === "BAJAS") {
    row = row.replaceAll(
      /{{\s*proximaRevision\s*}}/g,
      formatDateFromExcel(record.fechaProximaRevisionParteBaja) || "",
    );
  }

  let numeroParteConfirmacion = 0;
  if (
    record.partesConfirmacion &&
    Array.isArray(record.partesConfirmacion) &&
    record.partesConfirmacion.length > 0
  ) {
    row = row.replaceAll(
      /{{\s*proximaRevision\s*}}/g,
      formatDateFromExcel(
        record.partesConfirmacion[0].fechaSiguienteRevisionMedica,
      ) || "",
    );
    row = row.replaceAll(
      /{{\s*fechaConfirmacion\s*}}/g,
      formatDateFromExcel(
        record.partesConfirmacion[0].fechaDelParteDeConfirmacion,
      ) || "",
    );
    numeroParteConfirmacion =
      record.partesConfirmacion[0].numeroDeParteDeConfirmacion || 0;
  }

  const tipoProceso = calcularTipoProceso(record);

  // Observaciones (misma lógica que el individual)
  switch (tipoDoc) {
    case "ALTAS":
      row = row.replaceAll(
        /{{\s*observaciones\s*}}/g,
        "PARTE DE ALTA MÉDICA. <br/>" + tipoProceso,
      );
      break;
    case "BAJAS":
      row = row.replaceAll(
        /{{\s*observaciones\s*}}/g,
        "PARTE DE BAJA MÉDICA. <br/>" + tipoProceso,
      );
      break;
    case "CONFIRMACION":
      row = row.replaceAll(
        /{{\s*observaciones\s*}}/g,
        "PARTE DE CONFIRMACIÓN Nº" +
          numeroParteConfirmacion +
          "<br/>" +
          tipoProceso,
      );
      break;
  }

  return row;
}

// genera un texto plano agrupado 
function buildGroupedText(records, tipoDoc, expteEmpresa) {
  const hoy = formatTodayDDMMYYYY();

  const lines = records.map((r) => {
    const trabajador = safe(r.nombre);
    const contingencia = safe(r.contingencia);
    const fBaja = formatDateFromExcel(r.fechaBajaIt) || "";
    const fAlta = formatDateFromExcel(r.fechaFinIt) || "";

    let parteConf = "";
    let proxRev = "";

    if (
      r.partesConfirmacion &&
      Array.isArray(r.partesConfirmacion) &&
      r.partesConfirmacion.length > 0
    ) {
      parteConf = safe(r.partesConfirmacion[0].numeroDeParteDeConfirmacion || "");
      proxRev = formatDateFromExcel(
        r.partesConfirmacion[0].fechaSiguienteRevisionMedica,
      ) || "";
    } else if (tipoDoc === "BAJAS") {
      proxRev = formatDateFromExcel(r.fechaProximaRevisionParteBaja) || "";
    }

    return `- ${trabajador} | ${contingencia} | Baja: ${fBaja} | Conf: ${parteConf} | PróxRev: ${proxRev} | Alta: ${fAlta}`;
  });

  return `Buenos días,

Indicarles que hemos recibido información telemática de su empresa (${safe(expteEmpresa)}) correspondiente a procesos de IT (${tipoDoc}) que se detallan:

${lines.join("\n")}

Atentamente,
Susasesores.com
`;
}

// genera .eml agrupado (1 correo por empresa y tipología) 
async function generarEmailFieAgrupadoDesdePlantilla(records, tipoDoc, OUTPUT_DIR, overrideAddresses = {}) {
  if (!Array.isArray(records) || records.length === 0) {
    throw new Error("generarEmailFieAgrupadoDesdePlantilla: records vacío");
  }

  const expteEmpresa = records[0].expedienteEmpresa || "";
  const subject = `${safe(expteEmpresa)} - COMUNICACIÓN IT - (${tipoDoc}) - ${formatTodayDDMMYYYY()}`;

  let html = "";
  let parsed = null;

  if (tipoDoc === "CONFIRMACION") {
    // ✅ plantilla HTML pura
    html = fs.readFileSync(TEMPLATE_HTML_CONFIRMACION, "utf-8");
  } else {
    // ✅ ALTAS/BAJAS siguen con EML
    const TEMPLATE_EML = tipoDoc === "ALTAS" ? TEMPLATE_EML_ALTA : TEMPLATE_EML_BAJA;
    const raw = fs.readFileSync(TEMPLATE_EML);
    parsed = await simpleParser(raw);
    html = parsed.html || "";
  }

  // ✅ Inyección FILAS_IT (prioridad)
  if (/{{\s*FILAS_IT\s*}}/i.test(html)) {
    const filas = buildFilasITHtml(records, tipoDoc);
    html = html.replace(/{{\s*FILAS_IT\s*}}/gi, filas);
  } else {
    // fallback antiguo (por si alguna plantilla vieja)
    const trRegex = /<tr\b[^>]*>[\s\S]*?<\/tr>/gi;
    const allTr = html.match(trRegex) || [];
    const rowTemplate = allTr.find((tr) => /{{\s*nombre\s*}}/i.test(tr));
    if (rowTemplate) {
      const rows = records.map((r) => personalizarFilaHtml(rowTemplate, r, tipoDoc));
      html = html.replace(rowTemplate, rows.join("\n"));
    } else {
      console.warn("[emails-fie] No se encontró {{FILAS_IT}} ni una fila <tr> con {{nombre}} en la plantilla.");
    }
  }

  const text = buildGroupedText(records, tipoDoc, expteEmpresa);

  const rawNew = await saveAsNewEml(parsed, subject, html, text, {
    to: overrideAddresses.to ?? [],
  });

  const base = `${safeFilename(expteEmpresa)}_${tipoDoc}_${formatTodayDDMMYYYY_noSlash()}`;
  const outPath = path.join(OUTPUT_DIR, `${base}.eml`);
  fs.writeFileSync(outPath, rawNew, "utf-8");

  console.log("Guardando en:", outPath);
  return outPath;
}



// Modifica el HTML de la plantilla para meter los datos del cliente
/*
function personalizeHtml(html, record, tipoDoc) {
  const trabajador = safe(record.nombre);
  const contingencia = safe(record.contingencia).replace(/^\d+=/, "");
  const fBaja = formatDateFromExcel(record.fechaBajaIt) || "";
  const fAlta = formatDateFromExcel(record.fechaFinIt) || "";
  const parteConf = ""; // si lo tienes en tus datos, colócalo aquí
  const proximaRev = ""; // ídem
  const observ = tipoDoc === "ALTAS" ? "PARTE DE ALTA MÉDICA." : "";

  const $ = cheerio.load(html, { decodeEntities: false });

  // 1) Actualiza el título si lo hay (opcional)
  // $('title').text('Nuevo asunto o título'); // opcional

  // 2) Localiza la fila de datos. Estrategia: buscar encabezados y moverte a su fila.
  // Aquí asumimos que hay UNA fila de datos. Si hay varias, duplica la fila como plantilla.
  const headers = $("table th")
    .map((_, th) => $(th).text().trim().toLowerCase())
    .get();

  // Encuentra la primera fila de <tbody>
  const row = $("table tbody tr").first();
  if (row.length) {
    const cells = row.find("td");

    // Mapear según el orden esperado de columnas
    // [Trabajador/a, Contingencia, F. Baja Médica, Parte de Confirmación, Próxima Revisión, F. Alta Médica, Observaciones]
    const setCell = (index, value) => {
      if (cells.eq(index).length) cells.eq(index).html(value || "&nbsp;");
    };

    // Si prefieres localizar por nombre del th en lugar de índice:
    // const idxTrab = headers.indexOf('trabajador/a');
    // ... y luego setCell(idxTrab, trabajador);

    setCell(0, trabajador);
    setCell(1, contingencia);
    setCell(2, fBaja);
    setCell(3, parteConf);
    setCell(4, proximaRev);
    setCell(5, fAlta);
    setCell(6, observ);
  }

  return $.html(); // devuelve HTML completo
}
*/

// Crea un cuerpo de texto plano a partir de los datos (o toma el de la plantilla y lo adapta)
function personalizeText(originalText, record, tipoDoc) {
  // Si el .eml trae texto plano, puedes hacer reemplazos.
  // Si prefieres generar uno nuevo, ignora originalText y crea desde cero:
  const trabajador = safe(record.nombre);
  const contingencia = safe(record.contingencia).replace(/^\d+=/, "");
  const fBaja = formatDateFromExcel(record.fechaBajaIt) || "";
  const fAlta = formatDateFromExcel(record.fechaFinIt) || "";
  const observ = tipoDoc === "ALTAS" ? "PARTE DE ALTA MÉDICA." : "";

  return `Buenos días,

Indicarles que hemos recibido información telemática de su empresa correspondiente a procesos de IT que se detallan:

Trabajador/a: ${trabajador}
Contingencia: ${contingencia}
F. Baja Médica: ${fBaja}
Parte de Confirmación:
Próxima Revisión:
F. Alta Médica: ${fAlta}
Observaciones: ${observ}

Atentamente,
Susasesores.com
`;
}

// Recompila y guarda como .eml nuevo (conservando inline images y adjuntos)
async function saveAsNewEml(parsed, newSubject, newHtml, newText, override = {}) {
  const transport = nodemailer.createTransport({
    streamTransport: true,
    buffer: true,
    newline: "windows",
  });

  const safeParsed = parsed || {};

  const attachments = (safeParsed.attachments || []).map((att) => ({
    filename: att.filename || "adjunto",
    content: att.content,
    contentType: att.contentType,
    cid: att.contentId || undefined,
  }));

  // Cabeceras originales (si venimos de un .eml plantilla)
  const originalHeaders = {};
  for (const h of safeParsed.headerLines || []) {
    const key = (h.key || "").toString();
    if (!key) continue;
    const val = safeParsed.headers?.get(key) ?? h.line.replace(/^[^:]+:\s*/, "");
    originalHeaders[key] = val;
  }

  const mergedHeaders = {
    ...originalHeaders,
    ...(override.headers || {}),
    "X-Unsent": "1", // ✅ Outlook lo abre como borrador
  };

  delete mergedHeaders["Message-ID"];
  delete mergedHeaders["Date"];

  const message = {
    from: override.from || safeParsed.from?.value?.[0] || undefined,
    to: override.to || [],
    cc: override.cc || [],
    bcc: override.bcc || [],
    subject: newSubject || safeParsed.subject || "",
    date: new Date(),
    headers: mergedHeaders,
    text: newText,
    html: newHtml,
    attachments,
  };

  const info = await transport.sendMail(message);
  return info.message.toString("utf-8");
}



// Función principal: abre plantilla, personaliza y guarda
async function generarEmailFieDesdePlantilla(
  record,
  tipoDoc,
  OUTPUT_DIR,
  overrideAddresses = {},
) {
  let TEMPLATE_EML = null;
  switch (tipoDoc) {
    case "BAJAS":
      TEMPLATE_EML = TEMPLATE_EML_BAJA;
      break;
    case "ALTAS":
      TEMPLATE_EML = TEMPLATE_EML_ALTA;
      break;
    case "CONFIRMACION":
      TEMPLATE_EML = TEMPLATE_HTML_CONFIRMACION;
      break;
    default:
      throw new Error(`Tipo de documento desconocido: ${tipoDoc}`);
  }

  const raw = fs.readFileSync(TEMPLATE_EML);

  const parsed = await simpleParser(raw);

  //console.log(parsed);

  const subject = buildSubject(record, tipoDoc);
  //const html = personalizeHtml(parsed.html || "", record, tipoDoc);
  var html = parsed.html || "";

  html = html.replaceAll(/{{\s*nombre\s*}}/g, record.nombre || "");
  html = html.replaceAll(
    /{{\s*fechaAlta\s*}}/g,
    formatDateFromExcel(record.fechaFinIt) || "",
  );

  html = html.replaceAll(
    /{{\s*fechaBaja\s*}}/g,
    formatDateFromExcel(record.fechaBajaIt) || "",
  );
  html = html.replaceAll(/{{\s*contingencia\s*}}/g, record.contingencia || "");

  if (tipoDoc === "BAJAS") {
    html = html.replaceAll(
      /{{\s*proximaRevision\s*}}/g,
      formatDateFromExcel(record.fechaProximaRevisionParteBaja) || "",
    );
  }

  var numeroParteConfirmacion = 0;
  if (
    record.partesConfirmacion &&
    Array.isArray(record.partesConfirmacion) &&
    record.partesConfirmacion.length > 0
  ) {
    html = html.replaceAll(
      /{{\s*proximaRevision\s*}}/g,
      formatDateFromExcel(
        record.partesConfirmacion[0].fechaSiguienteRevisionMedica,
      ) || "",
    );
    html = html.replaceAll(
      /{{\s*fechaConfirmacion\s*}}/g,
      formatDateFromExcel(
        record.partesConfirmacion[0].fechaDelParteDeConfirmacion,
      ) || "",
    );
    numeroParteConfirmacion =
      record.partesConfirmacion[0].numeroDeParteDeConfirmacion || 0;
  }

  //Tipo de proceso:
  var tipoProceso = "";
  const tipoProcStr = safe(record.tipoDeProceso, ""); // safe ya existe arriba
  const primeraLetra = (tipoProcStr && tipoProcStr[0]) || null;
  
  switch (Number(primeraLetra)) {
    case 1:
      tipoProceso = "PROCESO MUY CORTO";
      break;
    case 2:
      tipoProceso = "PROCESO CORTO";
      break;
    case 3:
      tipoProceso = "PROCESO INTERMEDIO";
      break;
    case 4:
      tipoProceso = "PROCESO LARGO";
      break;
    default:
      tipoProceso = ""; // o algún texto por defecto si quieres
      break;
  }


  //Procesar observaciones:
  switch (tipoDoc) {
    case "ALTAS":
      html = html.replaceAll(
        /{{\s*observaciones\s*}}/g,
        "PARTE DE ALTA MÉDICA. <br/>" + tipoProceso,
      );
      break;
    case "BAJAS":
      html = html.replaceAll(
        /{{\s*observaciones\s*}}/g,
        "PARTE DE BAJA MÉDICA. <br/>" + tipoProceso,
      );
      break;
    case "CONFIRMACION":
      html = html.replaceAll(
        /{{\s*observaciones\s*}}/g,
        "PARTE DE CONFIRMACIÓN Nº" +
          numeroParteConfirmacion +
          "<br/>" +
          tipoProceso,
      );
      break;
  }

  const text = personalizeText(parsed.text || "", record, tipoDoc);
  const rawNew = await saveAsNewEml(parsed, subject, html, text, {
    to: overrideAddresses.to ?? [],
  });

  var base = "";

  switch (tipoDoc) {
    case "BAJAS":
      base = `${safeFilename(record.expte || "")}_${tipoDoc}_${safeFilename(record.dni)}_${safeFilename(formatDateFromExcel(record.fechaBajaIt) || "")}`;
      break;
    case "ALTAS":
      base = `${safeFilename(record.expte || "")}_${tipoDoc}_${safeFilename(record.dni)}_${safeFilename(formatDateFromExcel(record.fechaFinIt) || "")}`;
      break;
    case "CONFIRMACION":
      base = `${safeFilename(record.expte || "")}_${tipoDoc}_${safeFilename(record.dni)}_${safeFilename(formatDateFromExcel(record.partesConfirmacion[0].fechaDelParteDeConfirmacion) || "")}`;
      break;
  }

  const outPath = path.join(OUTPUT_DIR, `${base}.eml`);
  fs.writeFileSync(outPath, rawNew, "utf-8");
  console.log("Guardando en:", outPath);
  return outPath;
}

module.exports = generarEmailFieDesdePlantilla;

module.exports.generarEmailFieAgrupadoDesdePlantilla = generarEmailFieAgrupadoDesdePlantilla;


