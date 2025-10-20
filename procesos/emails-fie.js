const fs = require("fs");
const path = require("path");
const { simpleParser } = require("mailparser");
//const cheerio = require("cheerio");
const nodemailer = require("nodemailer");

// ===== Utilidades =====
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

// ===== Config =====
const TEMPLATE_EML_ALTA = path.join(__dirname, "fie", "alta.eml"); // tu plantilla
const TEMPLATE_EML_BAJA = path.join(__dirname, "fie", "baja.eml"); // tu plantilla
const TEMPLATE_EML_CONFIRMACION = path.join(
  __dirname,
  "fie",
  "confirmacion.eml",
); // tu plantilla

// Puedes fijar un FROM/TO/CC por defecto si quieres sobreescribir los de la plantilla
const DEFAULT_FROM = null;
const DEFAULT_TO = null;
const DEFAULT_CC = null;

// Construye asunto nuevo a partir del record
function buildSubject(record, tipoDoc) {
  const fechaRecep = formatDateFromExcel(record.fechaRecepcion);
  switch (tipoDoc) {
    case "BAJAS":
      return `${safe(record.expte)} - COMUNICACIÓN IT - ${safe(record.nombre)} - PB ${fechaRecep}`;
    case "ALTAS":
      return `${safe(record.expte)} - COMUNICACIÓN IT - ${safe(record.nombre)} - PA ${fechaRecep}`;
    case "CONFIRMACION":
      var parteConfirmacion = 0;
      if (Array.isArray(record.parteConfirmacion)) {
        parteConfirmacion =
          record.parteConfirmacion[0].numeroDeParteDeConfirmacion || 0;
      }
      return `${safe(record.expte)} - COMUNICACIÓN IT - ${safe(record.nombre)} - PC${parteConfirmacion} ${fechaRecep}`;
  }
  return `null`;
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
async function saveAsNewEml(
  parsed,
  newSubject,
  newHtml,
  newText,
  override = {},
) {
  const transport = nodemailer.createTransport({
    streamTransport: true,
    buffer: true,
    newline: "windows",
  });

  // Reconstruye attachments preservando cid para inline images
  const attachments = (parsed.attachments || []).map((att) => {
    // Mailparser nos da content (Buffer), contentId (cid), filename, contentType, headers, etc.
    const item = {
      filename: att.filename || "adjunto",
      content: att.content, // Buffer
      contentType: att.contentType,
      cid: att.contentId || undefined, // para inline
    };
    return item;
  });

  // Convierte headerLines de mailparser a objeto clave:valor
  const originalHeaders = {};
  for (const h of parsed.headerLines || []) {
    // h.key, h.line -> "Key: value"; mejor usar parsed.headers.get(h.key) pero esto vale
    const key = (h.key || "").toString();
    if (!key) continue;
    // Si hay claves repetidas, mailparser suele agrupar; nos quedamos con el valor "bonito"
    const val = parsed.headers?.get(key) ?? h.line.replace(/^[^:]+:\s*/, "");
    originalHeaders[key] = val;
  }

  // Forzamos borrador:
  const mergedHeaders = {
    ...originalHeaders,
  };

  delete mergedHeaders["Message-ID"];
  delete mergedHeaders["Date"];

  const message = {
    from: override.from || parsed.from?.value?.[0] || undefined,
    to: override.to || [],
    cc: override.cc || [],
    bcc: override.bcc || [],
    subject: newSubject || parsed.subject || "",
    date: parsed.date || new Date(),
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
      TEMPLATE_EML = TEMPLATE_EML_CONFIRMACION;
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
  switch (Number(record.tipoDeProceso.slice(1))) {
    case 1:
      tipoProceso = "PROCESO MUY CORTO";
      break;
    case 2:
      tipoProceso = "PROCESO CORTO";
      break;
    case 3:
      tipoProceso = "PROCESO MEDIO";
      break;
    case 4:
      tipoProceso = "PROCESO LARGO";
      break;
  }

  //Procesar observaciones:
  switch (tipoDoc) {
    case "ALTAS":
      html = html.replaceAll(
        /{{\s*observaciones\s*}}/g,
        "PARTE DE ALTA MÉDICA. ",
      );
      break;
    case "BAJAS":
      html = html.replaceAll(
        /{{\s*observaciones\s*}}/g,
        "PARTE DE BAJA MÉDICA. " + tipoProceso,
      );
      break;
    case "CONFIRMACION":
      html = html.replaceAll(
        /{{\s*observaciones\s*}}/g,
        "PARTE DE CONFIRMACIÓN Nº" + numeroParteConfirmacion,
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
      base = `${tipoDoc}_${safeFilename(record.dni)}_${safeFilename(formatDateFromExcel(record.fechaBajaIt) || "")}`;
      break;
    case "ALTAS":
      base = `${tipoDoc}_${safeFilename(record.dni)}_${safeFilename(formatDateFromExcel(record.fechaBajaIt) || "")}`;
      break;
    case "CONFIRMACION":
      base = `${tipoDoc}_${safeFilename(record.dni)}_${safeFilename(formatDateFromExcel(record.fechaBajaIt) || "")}`;
      break;
  }

  const outPath = path.join(OUTPUT_DIR, `${base}.eml`);
  fs.writeFileSync(outPath, rawNew, "utf-8");
  console.log("Guardando en:", outPath);
  return outPath;
}

module.exports = generarEmailFieDesdePlantilla;
