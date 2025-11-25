const fs = require("fs");
const fsp = fs.promises;
const path = require("path");
const PDFDocument = require("pdfkit");

// ===== Config =====
//const OUTPUT_DIR = path.join(__dirname, 'out');
const FONT_PATH = path.join(__dirname, "Roboto-Regular.ttf"); // Cambia si usas otra fuente
//if (!fs.existsSync(OUTPUT_DIR)) fs.mkdirSync(OUTPUT_DIR, { recursive: true });

// ===== Utilidades =====

// Convierte número de fecha Excel -> Date (s/fusos horarios raros)
function excelSerialToDate(serial) {
  if (serial === undefined || serial === null || serial === "") return null;
  // Excel base 1900 (con bug del 1900-02-29). 25569 = 1970-01-01
  const ms = Math.round((serial - 25569) * 86400 * 1000);
  if (Number.isNaN(ms)) return null;
  return new Date(ms);
}

function formatDateFromExcel(serial) {
  const d = excelSerialToDate(serial);
  if (!d) return "-";
  const dd = String(d.getDate()).padStart(2, "0");
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const yyyy = d.getFullYear();
  return `${dd}/${mm}/${yyyy}`;
}

// Asegura string y limpia control chars
function safeStr(v) {
  if (v === undefined || v === null) return "-";
  return String(v).replace(/\s+/g, " ").trim();
}

// Genera nombre de archivo seguro
function safeFilename(str, max = 80) {
  const cleaned = safeStr(str).replace(/[<>:"/\\|?*\x00-\x1F]/g, "_");
  return cleaned.slice(0, max) || "archivo";
}

// Dibuja una pareja (Etiqueta: valor)
function drawRow(doc, label, value, opts = {}) {
  const {
    labelWidth = 180, // ancho de la columna de etiqueta
    valueWidth = 360, // ancho de la columna de valor
    gap = 10, // separación entre ":" y el valor
    lineGap = 2, // espacio vertical entre filas
    labelAlign = "left",
    startX, // opcional: X de inicio fija (si no, margen izq. de la página)
  } = opts;

  // X fija para todas las filas (no dependas de doc.x)
  const baseX = startX !== undefined ? startX : doc.page.margins.left;
  const y = doc.y;

  const labelText = String(label ?? "-");
  const valueText = String(value ?? "-");

  // Medimos alturas para avanzar la Y uniformemente
  const labelMeasure = { width: labelWidth, align: labelAlign };
  const valueMeasure = { width: valueWidth };

  const labelHeight = doc.heightOfString(labelText, labelMeasure);
  const valueHeight = doc.heightOfString(valueText, valueMeasure);
  const rowHeight = Math.max(labelHeight, valueHeight);

  // Dibuja etiqueta siempre en la misma X
  doc.fontSize(9).fillColor("#333").text(labelText, baseX, y, {
    width: labelWidth,
    align: labelAlign,
    lineBreak: false,
    ellipsis: true,
  });

  // Dos puntos pegados al final de la columna de etiqueta
  doc.fillColor("#333").text(":", baseX + labelWidth, y, { lineBreak: false });

  // Valor empieza SIEMPRE en la misma vertical
  doc
    .fillColor("#000")
    .text(valueText, baseX + labelWidth + gap, y, { width: valueWidth });

  // Avanza a la siguiente fila, reseteando X para no “arrastrar” desplazamientos
  doc.x = baseX;
  doc.y = y + rowHeight + lineGap;
}

// Encabezado por tipo
function drawHeader(doc, tipo, opts = {}) {
  const { repeatOnlyLogo = false } = opts;
  const LOGO_PATH = path.join(__dirname, "../src/assets/Icono.png");
  const startY = doc.y;
  const left = doc.page.margins.left;
  const right = doc.page.width - doc.page.margins.right;

  // Logo (si existe)
  const logoHeight = 36; // ajusta a tu gusto (px)
  const logoTop = 40; // margen superior visual
  let headerBottomY = logoTop + logoHeight;

  if (fs.existsSync(LOGO_PATH)) {
    // Dibuja el logo con altura fija, manteniendo proporción
    doc.image(LOGO_PATH, left, logoTop, { height: logoHeight });
  }

  // Título (solo en la primera página)
  if (!repeatOnlyLogo) {
    doc
      .fontSize(18)
      .fillColor("#111")
      .text(`Parte de ${tipo}`, left + 140, logoTop + 6, { align: "left" }); // mueve si necesitas más espacio
    headerBottomY = Math.max(headerBottomY, logoTop + 24);
  }

  // Línea separadora
  doc
    .moveTo(left, headerBottomY + 10)
    .lineTo(right, headerBottomY + 10)
    .strokeColor("#999")
    .stroke();

  // Sitúa el cursor debajo del encabezado
  doc.moveTo(left, headerBottomY + 20);
  doc.y = headerBottomY + 24;
  doc.x = left;
}

function addFooterWithPageNumbers(doc) {
  const range = doc.bufferedPageRange(); // { start, count }
  for (let i = 0; i < range.count; i++) {
    doc.switchToPage(range.start + i);
    const footerText = `Página ${i + 1} de ${range.count}`;
    const y = doc.page.height - doc.page.margins.bottom + 10;
    doc.fontSize(8).fillColor("#666");
    doc.text(footerText, doc.page.margins.left, y, {
      width: doc.page.width - doc.page.margins.left - doc.page.margins.right,
      align: "center",
    });
  }
}

// Sección con título fino
function sectionTitle(doc, title) {
  doc.moveDown(0.5);
  doc.fontSize(12).fillColor("#222").text(title);
  doc.moveDown(0.2);
  doc
    .moveTo(doc.x, doc.y)
    .lineTo(doc.page.width - doc.page.margins.right, doc.y)
    .strokeColor("#ddd")
    .stroke();
  doc.moveDown(0.3);
}

// Mapa de campos por tipo
const FIELD_MAPS = [
  ["Clave autorización", "claveAutorizacion"],
  ["Fecha recepción", (r) => formatDateFromExcel(r.fechaRecepcion)],
  ["CCC", "ccc"],
  ["Empresa", "empresa"],
  ["Expediente", "expte"],
  ["NAF", "naf"],
  ["NIF", "nif"],
  ["Nombre", "nombre"],
  ["CNAE", "cnae"],
  ["Detalle CNAE", "detalleCnae"],
  [
    "Inicio relación laboral",
    (r) => formatDateFromExcel(r.fechaInicioRelacionLaboral),
  ],
  [
    "Extinción relación laboral",
    (r) => formatDateFromExcel(r.fechaExtincionRelacionLaboral),
  ],
  ["Fecha baja IT", (r) => formatDateFromExcel(r.fechaBajaIt)],
  ["Contingencia", "contingencia"],
  ["Entidad responsable", "entidadResponsable"],
  ["Recaída", "indicadorDeRecaida"],
  ["Fecha proceso inicial", (r) => formatDateFromExcel(r.fechaProcesoInicial)],
  [
    "Fecha proceso anterior",
    (r) => formatDateFromExcel(r.fechaProcesoAnterior),
  ],
  ["Días acumulados", "diasAcumulados"],
  [
    "IT inexistente (fecha)",
    (r) => formatDateFromExcel(r.fechaProcesoItInexistente),
  ],
  ["IT inexistente (causa)", "causaItProcesoInexistente"],
  ["Carencia", "indicadorCarencia"],
  ["Tipo de proceso", "tipoDeProceso"],
  ["Duración estimada (días)", "duracionEstimada"],
  [
    "Fin pago delegado (fecha)",
    (r) => formatDateFromExcel(r.fechaFinPagoDelegado),
  ],
  ["Fin pago delegado (causa)", "causaFinPagoDelegado"],
  ["Fin IT (fecha)", (r) => formatDateFromExcel(r.fechaFinIt)],
  ["Fin IT (causa)", "causaFinIt"],
  ["Parte de baja anulado", "parteDeBajaAnulado"],
  ["Parte de alta anulado", "parteDeAltaAnulado"],
  ["Modalidad de pago", "modalidadDePago"],
  ["Situaciones especiales IT", "situacionesEspecialesDeIt"],
  [
    "Peculiaridades pago/cotización",
    "procesosConPeculiaridadesEnPagoYCotizacion",
  ],
  ["IT internacional", "indicadorDeItInternacional"],
];

// Genera un PDF para un registro concreto
async function generatePDF(record, tipo, OUTPUT_DIR) {
  //Validacion de campos obligatorios:
  if (!record.expte) {
    console.error(
      "Error generando PDF: El campo 'expte' no puede estar vacío.",
    );
    return null;
  }

  if (!record.dni) {
    console.error("Error generando PDF: El campo 'dni' no puede estar vacío.");
    return null;
  }

  const doc = new PDFDocument({
    size: "A4",
    margin: 50,
    info: {
      Title: `Parte de ${tipo} - ${safeStr(record.dni)}`,
      Author: "Gestoría",
    },
  });

  // Archivo destino
  var baseName = "";
  switch (tipo) {
    case "BAJAS":
      baseName = `FIE B${safeFilename(formatDateFromExcel(record.fechaBajaIt))}.pdf`;
      break;
    case "ALTAS":
      baseName = `FIE A${safeFilename(formatDateFromExcel(record.fechaFinIt))}.pdf`;
      break;
    case "CONFIRMACIONES":
      baseName = `PC${safeFilename(record.partesConfirmacion[0].numeroDeParteDeConfirmacion)} ${safeFilename(formatDateFromExcel(record.partesConfirmacion[0].fechaDelParteDeConfirmacion))}.pdf`;
      break;
  }

  const filepath = path.join(
    OUTPUT_DIR,
    String(record.expte),
    String(record.dni),
    baseName,
  );

  const outDir = path.dirname(filepath);

  try {
    // Crea todas las carpetas necesarias si no existen
    await fsp.mkdir(outDir, { recursive: true });

    const stream = fs.createWriteStream(filepath);
    doc.pipe(stream);

    // Fuente Unicode para acentos, ñ, etc.
    if (fs.existsSync(FONT_PATH)) {
      doc.font(FONT_PATH);
    }

    // Cabecera
    drawHeader(doc, tipo);

    // Sección datos de empresa/cliente
    sectionTitle(doc, "Datos del cliente");

    const datosBasicos = [
      ["Nombre", "nombre"],
      ["NIF", "nif"],
      ["NAF", "naf"],

      ["Empresa", "empresa"],
      ["Expediente", "expte"],
      ["CCC", "ccc"],
    ];
    datosBasicos.forEach(([label, key]) =>
      drawRow(doc, label, safeStr(record[key])),
    );

    // Sección proceso IT
    sectionTitle(doc, "Datos del proceso");
    const map = FIELD_MAPS;
    map.forEach(([label, keyOrFn]) => {
      const value =
        typeof keyOrFn === "function"
          ? keyOrFn(record)
          : safeStr(record[keyOrFn]);
      drawRow(doc, label, safeStr(value));
      // Si se acerca al final de página, crear una nueva
      if (doc.y > doc.page.height - doc.page.margins.bottom - 80) {
        //addFooterWithPageNumbers(doc);
        //doc.addPage();
      }
    });

    //Añade pagina si hay parte o datos adicionales:
    if (
      record.datosAdicionales ||
      (tipo === "CONFIRMACIONES" &&
        Array.isArray(record.partesConfirmacion) &&
        record.partesConfirmacion.length > 0)
    ) {
      doc.addPage();
    }

    // Sección datos adicionales
    if (record.datosAdicionales) {
      sectionTitle(doc, "Datos Adicionales");
      const datosAdicionalesMap = [
        ["Fecha AT/EP", (p) => formatDateFromExcel(p["fechaAt/Ep"])],
        ["Tipo Accidente", "tipoAccidente"],
        ["Tipo de asistencia", "tipoDeAsistencia"],
        [
          "Fecha siguiente revisión médica Parte de Baja",
          (p) => formatDateFromExcel(p.fechaSiguienteRevisionMedicaParteDeBaja),
        ],
        [
          "Fecha Agotamiento 545 días de IT	",
          (p) => formatDateFromExcel(p.fechaAgotamiento545DiasDeIt),
        ],
        [
          "Fecha de envío PB, PC o PA al INSS",
          (p) => formatDateFromExcel(p["fechaDeEnvioPb,PcOPaAlInss"]),
        ],

        ["Base reguladora INSS", "baseReguladoraInss"],
        [
          "Código de resolución IT emitida por el INSS",
          "codigoDeResolucionItEmitidaPorElInss",
        ],
        [
          "Fecha de resolución IT emitida por el INSS",
          (p) => formatDateFromExcel(p.fechaDeResolucionItEmitidaPorElInss),
        ],
        [
          "Fecha notificación denegación IP al trabajador",
          (p) =>
            formatDateFromExcel(p.fechaNotificacionDenegacionIpAlTrabajador),
        ],
      ];

      datosAdicionalesMap.forEach(([label, keyOrFn]) => {
        const value =
          typeof keyOrFn === "function"
            ? keyOrFn(record.datosAdicionales)
            : safeStr(record.datosAdicionales[keyOrFn]);
        drawRow(doc, label, safeStr(value));
        // Si se acerca al final de página, crear una nueva
        if (doc.y > doc.page.height - doc.page.margins.bottom - 80) {
          //addFooterWithPageNumbers(doc);
          //doc.addPage();
        }
      });
      doc.moveDown(1.2);
    }

    if (
      tipo === "CONFIRMACIONES" &&
      Array.isArray(record.partesConfirmacion) &&
      record.partesConfirmacion.length > 0
    ) {
      record.partesConfirmacion.forEach((parte, index) => {
        sectionTitle(
          doc,
          `Parte de Confirmación ${parte.numeroDeParteDeConfirmacion}`,
        );

        const parteMap = [
          ["Número de parte", "numeroDeParteDeConfirmacion"],
          [
            "Fecha del parte de confirmación",
            (p) => formatDateFromExcel(p.fechaDelParteDeConfirmacion),
          ],
          ["Nombre", "nombre"],
          ["Empresa", "empresa"],
          ["NIF", "nif"],
          [
            "Fecha de inicio de relación laboral",
            (p) => formatDateFromExcel(p.fechaInicioRelacionLaboral),
          ],
          [
            "Fecha de extinción de relación laboral",
            (p) => formatDateFromExcel(p.fechaExtincionRelacionLaboral),
          ],
          [
            "Fecha de cumplimiento 365 días",
            (p) => formatDateFromExcel(p.fechaDeCumplimientoDe365Dias),
          ],
          [
            "Fecha siguiente revisión",
            (p) => formatDateFromExcel(p.fechaSiguienteRevisionMedica),
          ],
        ];
        parteMap.forEach(([label, keyOrFn]) => {
          const value =
            typeof keyOrFn === "function"
              ? keyOrFn(parte)
              : safeStr(parte[keyOrFn]);

          drawRow(doc, label, safeStr(value));
        });
      });
    }

    // Pie
    doc.moveDown(1.2);
    doc
      .fontSize(8)
      .fillColor("#555")
      .text(`Generado: ${new Date().toLocaleString()}`);

    //addFooterWithPageNumbers(doc);

    doc.end();

    return new Promise((resolve) => {
      stream.on("finish", () => resolve(filepath));
    });
  } catch (err) {
    console.error("No se pudo guardar el PDF:", err);
    return null;
  }
}

module.exports = generatePDF;
