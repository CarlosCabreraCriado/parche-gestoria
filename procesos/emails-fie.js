const fs = require("fs");
const path = require("path");
const { simpleParser } = require("mailparser");
const nodemailer = require("nodemailer");

// =======================
// Utilidades
// =======================
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

// =======================
// Config de plantillas (TODO en HTML)
// =======================
const TEMPLATE_HTML_ALTA = path.join(__dirname, "fie", "alta.html");
const TEMPLATE_HTML_BAJA = path.join(__dirname, "fie", "baja.html");
const TEMPLATE_HTML_CONFIRMACION = path.join(__dirname, "fie", "confirmacion.html");

// Puedes fijar un FROM/TO/CC por defecto si quieres sobreescribir los de la plantilla
// (En HTML no hay headers, así que esto viene bien para que Outlook no “cojee”)
const DEFAULT_FROM = null; // ejemplo: 'Susasesores <ro@susasesores.com>'
const DEFAULT_TO = null;
const DEFAULT_CC = null;

// =======================
// Helpers fecha "hoy"
// =======================
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

// =======================
// Lógica de negocio
// =======================
function calcularTipoProceso(record) {
  let tipoProceso = "";
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

// Construye asunto nuevo a partir del record (individual)
function buildSubject(record, tipoDoc) {
  let fechaRecep = formatDateFromExcel(record.fechaRecepcion);

  switch (tipoDoc) {
    case "BAJAS":
      fechaRecep = formatDateFromExcel(record.fechaBajaIt) || fechaRecep;
      return `${safe(record.expte)} - COMUNICACIÓN IT - ${safe(record.nombre)} - PB ${fechaRecep}`;
    case "ALTAS":
      fechaRecep = formatDateFromExcel(record.fechaFinIt) || fechaRecep;
      return `${safe(record.expte)} - COMUNICACIÓN IT - ${safe(record.nombre)} - PA ${fechaRecep}`;
    case "CONFIRMACION": {
      let parteConfirmacion = 0;
      if (Array.isArray(record.partesConfirmacion) && record.partesConfirmacion.length > 0) {
        fechaRecep =
          formatDateFromExcel(record.partesConfirmacion[0].fechaDelParteDeConfirmacion) ||
          fechaRecep;
        parteConfirmacion = record.partesConfirmacion[0].numeroDeParteDeConfirmacion || 0;
      }
      return `${safe(record.expte)} - COMUNICACIÓN IT - ${safe(record.nombre)} - PC${parteConfirmacion} ${fechaRecep}`;
    }
  }

  return `null`;
}

// =======================
// HTML filas tabla dinámica (agrupado)
// =======================
function buildFilasITHtml(records, tipoDoc) {
  const border = "#E7E3EA";
  const tdBase =
    `border:1px solid ${border}; padding:8px 6px; font-size:12px; line-height:16px; overflow-wrap:anywhere; background:#FFFFFF;`;

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
      if (tipoDoc === "ALTAS") observaciones = `PARTE DE ALTA MÉDICA.<br/>${tipoProceso}`;
      if (tipoDoc === "BAJAS") observaciones = `PARTE DE BAJA MÉDICA.<br/>${tipoProceso}`;
      if (tipoDoc === "CONFIRMACION") {
        let num = 0;
        if (Array.isArray(r.partesConfirmacion) && r.partesConfirmacion.length > 0) {
          num = r.partesConfirmacion[0].numeroDeParteDeConfirmacion || 0;
        }
        observaciones = `PARTE DE CONFIRMACIÓN Nº${escapeHtml(String(num))}<br/>${tipoProceso}`;
      }

      return `
<tr>
  <td bgcolor="#FFFFFF" style="${tdBase}">${nombre}</td>
  <td bgcolor="#FFFFFF" style="${tdBase}">${contingencia}</td>
  <td bgcolor="#FFFFFF" style="${tdBase} text-align:center; white-space:nowrap;">${fechaBaja}</td>
  <td bgcolor="#FFFFFF" style="${tdBase} text-align:center; white-space:nowrap;">${escapeHtml(fechaConfirmacion)}</td>
  <td bgcolor="#FFFFFF" style="${tdBase} text-align:center; white-space:nowrap;">${escapeHtml(proximaRevision)}</td>
  <td bgcolor="#FFFFFF" style="${tdBase} text-align:center; white-space:nowrap;">${fechaAlta}</td>
  <td bgcolor="#FFFFFF" style="${tdBase}">${observaciones}</td>
</tr>`.trim();
    })
    .join("\n");
}

// fallback (si pillas una plantilla vieja sin {{FILAS_IT}})
function personalizarFilaHtml(filaHtml, record, tipoDoc) {
  let row = filaHtml;

  row = row.replaceAll(/{{\s*nombre\s*}}/g, record.nombre || "");
  row = row.replaceAll(/{{\s*fechaAlta\s*}}/g, formatDateFromExcel(record.fechaFinIt) || "");
  row = row.replaceAll(/{{\s*fechaBaja\s*}}/g, formatDateFromExcel(record.fechaBajaIt) || "");
  row = row.replaceAll(/{{\s*contingencia\s*}}/g, record.contingencia || "");

  if (tipoDoc === "BAJAS") {
    row = row.replaceAll(
      /{{\s*proximaRevision\s*}}/g,
      formatDateFromExcel(record.fechaProximaRevisionParteBaja) || "",
    );
  }

  let numeroParteConfirmacion = 0;
  if (Array.isArray(record.partesConfirmacion) && record.partesConfirmacion.length > 0) {
    row = row.replaceAll(
      /{{\s*proximaRevision\s*}}/g,
      formatDateFromExcel(record.partesConfirmacion[0].fechaSiguienteRevisionMedica) || "",
    );
    row = row.replaceAll(
      /{{\s*fechaConfirmacion\s*}}/g,
      formatDateFromExcel(record.partesConfirmacion[0].fechaDelParteDeConfirmacion) || "",
    );
    numeroParteConfirmacion = record.partesConfirmacion[0].numeroDeParteDeConfirmacion || 0;
  }

  const tipoProceso = calcularTipoProceso(record);

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
        "PARTE DE CONFIRMACIÓN Nº" + numeroParteConfirmacion + "<br/>" + tipoProceso,
      );
      break;
  }

  return row;
}

// =======================
// Texto plano (agrupado)
// =======================
function buildGroupedText(records, tipoDoc, expteEmpresa) {
  const lines = records.map((r) => {
    const trabajador = safe(r.nombre);
    const contingencia = safe(r.contingencia);
    const fBaja = formatDateFromExcel(r.fechaBajaIt) || "";
    const fAlta = formatDateFromExcel(r.fechaFinIt) || "";

    let parteConf = "";
    let proxRev = "";

    if (Array.isArray(r.partesConfirmacion) && r.partesConfirmacion.length > 0) {
      parteConf = safe(r.partesConfirmacion[0].numeroDeParteDeConfirmacion || "");
      proxRev =
        formatDateFromExcel(r.partesConfirmacion[0].fechaSiguienteRevisionMedica) || "";
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

// =======================
// Plantilla por tipo (HTML)
// =======================
function getTemplatePathByTipo(tipoDoc) {
  switch (tipoDoc) {
    case "ALTAS":
      return TEMPLATE_HTML_ALTA;
    case "BAJAS":
      return TEMPLATE_HTML_BAJA;
    case "CONFIRMACION":
      return TEMPLATE_HTML_CONFIRMACION;
    default:
      throw new Error(`Tipo de documento desconocido: ${tipoDoc}`);
  }
}

// =======================
// saveAsNewEml (genera .eml listo para Outlook)
// =======================
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
    "X-Unsent": "1",
  };

  delete mergedHeaders["Message-ID"];
  delete mergedHeaders["Date"];

  const message = {
    from: override.from || DEFAULT_FROM || safeParsed.from?.value?.[0] || undefined,
    to: override.to || DEFAULT_TO || [],
    cc: override.cc || DEFAULT_CC || [],
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

// =======================
// AGRUPADO (1 correo por empresa y tipología)
// =======================
async function generarEmailFieAgrupadoDesdePlantilla(
  records,
  tipoDoc,
  OUTPUT_DIR,
  overrideAddresses = {},
) {
  if (!Array.isArray(records) || records.length === 0) {
    throw new Error("generarEmailFieAgrupadoDesdePlantilla: records vacío");
  }

  const expteEmpresa = records[0].expedienteEmpresa || "";
  const subject = `${safe(expteEmpresa)} - COMUNICACIÓN IT - (${tipoDoc}) - ${formatTodayDDMMYYYY()}`;

  // ✅ En agrupado: SIEMPRE HTML (alta/baja/confirmacion)
  const templatePath = getTemplatePathByTipo(tipoDoc);
  let html = fs.readFileSync(templatePath, "utf-8");
  const parsed = null; // no viene de EML

  // ✅ Inyección FILAS_IT (prioridad)
  if (/{{\s*FILAS_IT\s*}}/i.test(html)) {
    const filas = buildFilasITHtml(records, tipoDoc);
    html = html.replace(/{{\s*FILAS_IT\s*}}/gi, filas);
  } else {
    // fallback viejo
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
    cc: overrideAddresses.cc ?? [],
    bcc: overrideAddresses.bcc ?? [],
    from: overrideAddresses.from ?? null,
  });

  const base = `${safeFilename(expteEmpresa)}_${tipoDoc}_${formatTodayDDMMYYYY_noSlash()}`;
  const outPath = path.join(OUTPUT_DIR, `${base}.eml`);
  fs.writeFileSync(outPath, rawNew, "utf-8");

  console.log("Guardando en:", outPath);
  return outPath;
}

// =======================
// INDIVIDUAL (1 correo por registro)
// =======================
function personalizeText(originalText, record, tipoDoc) {
  const trabajador = safe(record.nombre);
  const contingencia = safe(record.contingencia).replace(/^\d+=/, "");
  const fBaja = formatDateFromExcel(record.fechaBajaIt) || "";
  const fAlta = formatDateFromExcel(record.fechaFinIt) || "";

  const tipoProceso = calcularTipoProceso(record);

  let observ = "";
  if (tipoDoc === "ALTAS") observ = "PARTE DE ALTA MÉDICA. " + tipoProceso;
  if (tipoDoc === "BAJAS") observ = "PARTE DE BAJA MÉDICA. " + tipoProceso;

  if (tipoDoc === "CONFIRMACION") {
    const num =
      Array.isArray(record.partesConfirmacion) && record.partesConfirmacion.length > 0
        ? record.partesConfirmacion[0].numeroDeParteDeConfirmacion || 0
        : 0;
    observ = `PARTE DE CONFIRMACIÓN Nº${num}. ${tipoProceso}`;
  }

  return `Buenos días,

Indicarles que hemos recibido información telemática de su empresa correspondiente a procesos de IT que se detallan:

Trabajador/a: ${trabajador}
Contingencia: ${contingencia}
F. Baja Médica: ${fBaja}
F. Alta Médica: ${fAlta}
Observaciones: ${observ}

Atentamente,
Susasesores.com
`;
}

async function generarEmailFieDesdePlantilla(record, tipoDoc, OUTPUT_DIR, overrideAddresses = {}) {
  const templatePath = getTemplatePathByTipo(tipoDoc);

  // ✅ HTML puro
  let html = fs.readFileSync(templatePath, "utf-8");
  const parsed = null;

  // ✅ Sustituciones individuales
  html = html.replaceAll(/{{\s*nombre\s*}}/g, escapeHtml(record.nombre || ""));
  html = html.replaceAll(/{{\s*fechaAlta\s*}}/g, escapeHtml(formatDateFromExcel(record.fechaFinIt) || ""));
  html = html.replaceAll(/{{\s*fechaBaja\s*}}/g, escapeHtml(formatDateFromExcel(record.fechaBajaIt) || ""));
  html = html.replaceAll(/{{\s*contingencia\s*}}/g, escapeHtml(normalizeContingencia(record.contingencia || "")));

  // Próxima revisión (BAJAS) / Confirmación (CONFIRMACION)
  if (tipoDoc === "BAJAS") {
    html = html.replaceAll(
      /{{\s*proximaRevision\s*}}/g,
      escapeHtml(formatDateFromExcel(record.fechaProximaRevisionParteBaja) || ""),
    );
  }

  let numeroParteConfirmacion = 0;
  if (Array.isArray(record.partesConfirmacion) && record.partesConfirmacion.length > 0) {
    html = html.replaceAll(
      /{{\s*proximaRevision\s*}}/g,
      escapeHtml(formatDateFromExcel(record.partesConfirmacion[0].fechaSiguienteRevisionMedica) || ""),
    );
    html = html.replaceAll(
      /{{\s*fechaConfirmacion\s*}}/g,
      escapeHtml(formatDateFromExcel(record.partesConfirmacion[0].fechaDelParteDeConfirmacion) || ""),
    );
    numeroParteConfirmacion = record.partesConfirmacion[0].numeroDeParteDeConfirmacion || 0;
  } else {
    // por si en la plantilla existen esos placeholders
    html = html.replaceAll(/{{\s*fechaConfirmacion\s*}}/g, "");
    html = html.replaceAll(/{{\s*proximaRevision\s*}}/g, "");
  }

  const tipoProceso = calcularTipoProceso(record);

  // Observaciones
  switch (tipoDoc) {
    case "ALTAS":
      html = html.replaceAll(/{{\s*observaciones\s*}}/g, `PARTE DE ALTA MÉDICA. <br/>${escapeHtml(tipoProceso)}`);
      break;
    case "BAJAS":
      html = html.replaceAll(/{{\s*observaciones\s*}}/g, `PARTE DE BAJA MÉDICA. <br/>${escapeHtml(tipoProceso)}`);
      break;
    case "CONFIRMACION":
      html = html.replaceAll(
        /{{\s*observaciones\s*}}/g,
        `PARTE DE CONFIRMACIÓN Nº${escapeHtml(String(numeroParteConfirmacion))}<br/>${escapeHtml(tipoProceso)}`,
      );
      break;
  }

  const subject = buildSubject(record, tipoDoc);
  const text = personalizeText("", record, tipoDoc);

  const rawNew = await saveAsNewEml(parsed, subject, html, text, {
    to: overrideAddresses.to ?? [],
    cc: overrideAddresses.cc ?? [],
    bcc: overrideAddresses.bcc ?? [],
    from: overrideAddresses.from ?? null,
  });

  let base = "";
  switch (tipoDoc) {
    case "BAJAS":
      base = `${safeFilename(record.expte || "")}_${tipoDoc}_${safeFilename(record.dni)}_${safeFilename(formatDateFromExcel(record.fechaBajaIt) || "")}`;
      break;
    case "ALTAS":
      base = `${safeFilename(record.expte || "")}_${tipoDoc}_${safeFilename(record.dni)}_${safeFilename(formatDateFromExcel(record.fechaFinIt) || "")}`;
      break;
    case "CONFIRMACION": {
      const f = Array.isArray(record.partesConfirmacion) && record.partesConfirmacion.length > 0
        ? formatDateFromExcel(record.partesConfirmacion[0].fechaDelParteDeConfirmacion) || ""
        : "";
      base = `${safeFilename(record.expte || "")}_${tipoDoc}_${safeFilename(record.dni)}_${safeFilename(f)}`;
      break;
    }
  }

  const outPath = path.join(OUTPUT_DIR, `${base}.eml`);
  fs.writeFileSync(outPath, rawNew, "utf-8");
  console.log("Guardando en:", outPath);
  return outPath;
}

// =======================
// Exports
// =======================
module.exports = generarEmailFieDesdePlantilla;
module.exports.generarEmailFieAgrupadoDesdePlantilla = generarEmailFieAgrupadoDesdePlantilla;
