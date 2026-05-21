const fs = require("fs");
const path = require("path");
const nodemailer = require("nodemailer");

// =======================
// Utilidades
// =======================
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
// Tipos de certificado (orden y etiqueta)
// =======================
const TIPOS_CERTIFICADO = [
  { key: "nombreArchivoSS",    label: "SS" },
  { key: "nombreArchivoTrib",  label: "AEAT" },
  { key: "nombreArchivoATC",   label: "ATC" },
  { key: "nombreArchivoITA",   label: "ITA" },
  { key: "nombreArchivoArt42", label: "ART42" },
];

// =======================
// Construye filas HTML de la tabla de certificados
// =======================
function buildFilasCertificadosHtml(adjuntosInfo) {
  const border = "#E7E3EA";
  const tdBase = `border:1px solid ${border}; padding:8px 6px; font-size:12px; line-height:16px; background:#FFFFFF;`;

  return adjuntosInfo
    .map(
      ({ label, filename }) => `
<tr>
  <td bgcolor="#FFFFFF" style="${tdBase} font-weight:bold; white-space:nowrap;">${escapeHtml(label)}</td>
  <td bgcolor="#FFFFFF" style="${tdBase}">${escapeHtml(filename)}</td>
</tr>`.trim()
    )
    .join("\n");
}

// =======================
// Texto plano del correo
// =======================
function buildTexto(grupo, adjuntosInfo) {
  const codigo = safe(grupo[0].codigo);
  const empresa = safe(grupo[0].empresa);
  const lineas = adjuntosInfo.map(({ label, filename }) => `- ${label}: ${filename}`);

  return `Buenos días,

Indicarles que adjuntamos la documentación solicitada en este mensaje. Cualquier documentación adicional que necesiten por favor, hacédnoslo saber y la remitiremos a la mayor brevedad.

${empresa} (expediente ${codigo}):
${lineas.join("\n")}

Quedamos a su disposición para cualquier consulta o gestión adicional que pueda necesitar.

Atentamente,
Susasesores.com
`;
}

// =======================
// Crear archivo .EML con nodemailer
// =======================
async function saveAsNewEml(newSubject, newHtml, newText, toAddresses, adjuntos) {
  const transport = nodemailer.createTransport({
    streamTransport: true,
    buffer: true,
    newline: "windows",
  });

  const logoPath = path.join(__dirname, "..", "..", "src", "assets", "correos", "logo_susasesores.png");

  const attachments = [
    ...adjuntos,
    {
      filename: "logo_susasesores.png",
      content: fs.readFileSync(logoPath),
      contentType: "image/png",
      cid: "logo_susasesores",
    },
  ];

  const message = {
    to: toAddresses,
    cc: [],
    bcc: [],
    subject: newSubject,
    date: new Date(),
    headers: { "X-Unsent": "1" },
    text: newText,
    html: newHtml,
    attachments,
  };

  const info = await transport.sendMail(message);
  return info.message.toString("utf-8");
}

// =======================
// Función principal exportada
// =======================
async function generarEmailCertificados(grupo, carpetaRaiz, correos, carpetaCorreos) {
  if (!Array.isArray(grupo) || grupo.length === 0) {
    throw new Error("generarEmailCertificados: grupo vacío");
  }

  const codigo = safe(grupo[0].codigo);
  const empresa = safe(grupo[0].empresa);
  const subject = `${codigo} - CERTIFICADOS DE ESTAR AL CORRIENTE - ${empresa} - ${formatTodayDDMMYYYY()}`;

  // Recopilar adjuntos que existen en disco
  const adjuntosInfo = [];
  const adjuntos = [];
  for (const cliente of grupo) {
    for (const { key, label } of TIPOS_CERTIFICADO) {
      if (cliente[key]) {
        const ruta = path.join(carpetaRaiz, cliente[key]);
        if (fs.existsSync(ruta)) {
          adjuntosInfo.push({ label, filename: cliente[key] });
          adjuntos.push({ filename: cliente[key], content: fs.readFileSync(ruta) });
        }
      }
    }
  }

  if (adjuntosInfo.length === 0) {
    console.log(`[EMAIL] Sin certificados descargados para expediente ${codigo}, omitiendo correo.`);
    return null;
  }

  // Cargar plantilla HTML e inyectar filas
  const templatePath = path.join(__dirname, "certificado.html");
  let html = fs.readFileSync(templatePath, "utf-8");
  const filas = buildFilasCertificadosHtml(adjuntosInfo);
  html = html.replace(/{{\s*FILAS_CERTIFICADOS\s*}}/gi, filas);

  const text = buildTexto(grupo, adjuntosInfo);

  const rawEml = await saveAsNewEml(subject, html, text, correos, adjuntos);

  const base = `${safeFilename(codigo)}_${safeFilename(empresa)}_${formatTodayDDMMYYYY_noSlash()}`;
  const outPath = path.join(carpetaCorreos, `${base}.eml`);
  fs.writeFileSync(outPath, rawEml, "utf-8");
  console.log(`[EMAIL] Borrador generado: ${outPath}`);
  return outPath;
}

module.exports = { generarEmailCertificados };
