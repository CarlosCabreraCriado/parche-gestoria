/**
 * Proceso: Bases y recibos al cobro autónomos
 *
 * Flujo basado en:
 * - Consulta de bases y cuotas ingresadas (año actual) + Imprimir (nueva pestaña) :contentReference[oaicite:5]{index=5}
 * - Consulta de recibos emitidos régimen de autónomos + primer recibo + imprimir/guardar :contentReference[oaicite:6]{index=6}
 */

const path = require("path");
const fs = require("fs");
const XlsxPopulate = require("xlsx-populate");
const puppeteer = require("puppeteer");

const {
  waitForPopup,
  descargarPdfRawViaFetchCDP,
  descargarPdfConReintento,
} = require("../utils/pdfNuevaPestanaCdp"); // ajusta ruta según tu repo

class BasesYRecibosAutonomos {
  constructor() {
    this.DEFAULT_ANIO_ECONOMICO = "2025";
    this.MAX_REINTENTOS_POR_REGISTRO = 2;
  }

  // ---------------------------
  // Utils
  // ---------------------------
  async esperar(ms) {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }

  ensureDir(dir) {
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
  }

  safeFileName(name) {
    return String(name || "")
      .trim()
      .replace(/[<>:"/\\|?*\x00-\x1F]/g, " ")
      .replace(/\s+/g, " ")
      .trim();
  }

  dniToFolder(dni) {
    const digits = String(dni || "").replace(/\D/g, "");
    return digits || "DNI_DESCONOCIDO";
  }

  getMesAnioString(fecha = new Date()) {
    const mm = String(fecha.getMonth() + 1).padStart(2, "0");
    const yyyy = String(fecha.getFullYear());
    return `${mm}-${yyyy}`; // ✅ mes actual (como pediste)
  }

  nowStamp() {
    const d = new Date();
    const yyyy = d.getFullYear();
    const mm = String(d.getMonth() + 1).padStart(2, "0");
    const dd = String(d.getDate()).padStart(2, "0");
    const hh = String(d.getHours()).padStart(2, "0");
    const mi = String(d.getMinutes()).padStart(2, "0");
    const ss = String(d.getSeconds()).padStart(2, "0");
    return `${yyyy}${mm}${dd}_${hh}${mi}${ss}`;
  }

  async withTimeout(promise, ms, errorMsg) {
    let t;
    const timeout = new Promise((_, reject) => {
      t = setTimeout(() => reject(new Error(errorMsg || `Timeout ${ms}ms`)), ms);
    });
    try {
      return await Promise.race([promise, timeout]);
    } finally {
      clearTimeout(t);
    }
  }

  // ---------------------------
  // 1) Lectura Excel (incluye columnas ocultas)
  // ---------------------------
  async leerExcel(pathExcel) {
    const workbook = await XlsxPopulate.fromFileAsync(pathExcel);
    const sheet = workbook.sheet(0);

    const usedRange = sheet.usedRange();
    const maxRow = usedRange.endCell().rowNumber();
    const maxCol = usedRange.endCell().columnNumber();

    // Buscar fila cabecera por "EXPTE." en columna A
    let headerRow = null;
    for (let r = 1; r <= Math.min(maxRow, 50); r++) {
      const val = sheet.cell(r, 1).value();
      if (String(val || "").trim().toUpperCase() === "EXPTE.") {
        headerRow = r;
        break;
      }
    }
    if (!headerRow) throw new Error("No se encontró la fila de cabeceras (buscando 'EXPTE.' en columna A).");

    // Map cabeceras por columna (ocultas incluidas)
    const headersByCol = {};
    for (let c = 1; c <= maxCol; c++) headersByCol[c] = sheet.cell(headerRow, c).value();

    // Forzar nombres G/H -> NAF1/NAF2
    headersByCol[7] = "NAF1";
    headersByCol[8] = "NAF2";

    // Leer registros
    const registros = [];
    let emptyAStreak = 0;

    for (let r = headerRow + 1; r <= maxRow; r++) {
      const expte = sheet.cell(r, 1).value();

      if (expte === null || expte === undefined || String(expte).trim() === "") {
        emptyAStreak++;
        if (emptyAStreak >= 5) break;
        continue;
      }
      emptyAStreak = 0;

      const rowObj = {};
      for (let c = 1; c <= maxCol; c++) {
        const key = headersByCol[c] || `COL_${c}`;
        rowObj[key] = sheet.cell(r, c).value();
      }

      // Normalización de claves más usadas
      rowObj.DNI = rowObj["DNI"] ?? "";
      rowObj.ADMINISTRADOR = rowObj["ADMINISTRADOR "] ?? rowObj["ADMINISTRADOR"] ?? "";
      rowObj.NAF1 = String(rowObj["NAF1"] ?? "").trim();
      rowObj.NAF2 = String(rowObj["NAF2"] ?? "").trim();

      registros.push({ row: r, data: rowObj });
    }

    return { headerRow, maxCol, headersByCol, registros };
  }

  validarRegistro(reg) {
    const errores = [];
    const dni = String(reg.DNI || "").trim();
    const admin = String(reg.ADMINISTRADOR || "").trim();
    const naf1 = String(reg.NAF1 || "").trim();
    const naf2 = String(reg.NAF2 || "").trim();

    if (!dni) errores.push("DNI vacío");
    if (!admin) errores.push("ADMINISTRADOR vacío");
    if (!/^\d{2}$/.test(naf1)) errores.push("NAF1 inválido (2 dígitos numéricos)");
    if (!/^\d{10}$/.test(naf2)) errores.push("NAF2 inválido (10 dígitos numéricos)");

    return errores;
  }

  // ---------------------------
  // Helpers Puppeteer (frames)
  // ---------------------------
  async clickLinkByTextInAnyFrame(page, text, timeoutMs = 15000) {
    const end = Date.now() + timeoutMs;
    while (Date.now() < end) {
      const frames = page.frames();
      for (const frame of frames) {
        const found = await frame.evaluate((t) => {
          const norm = (s) => (s || "").replace(/\s+/g, " ").trim();
          const links = Array.from(document.querySelectorAll("a"));
          const link = links.find((a) => norm(a.textContent).includes(t));
          if (link) {
            link.scrollIntoView({ block: "center" });
            link.click();
            return true;
          }
          return false;
        }, text);
        if (found) return true;
      }
      await this.esperar(250);
    }
    throw new Error(`No se encontró link con texto: "${text}"`);
  }

  async waitForSelectorInAnyFrame(page, selector, timeoutMs = 15000) {
    const end = Date.now() + timeoutMs;
    while (Date.now() < end) {
      for (const frame of page.frames()) {
        const h = await frame.$(selector);
        if (h) return { frame, handle: h };
      }
      await this.esperar(250);
    }
    throw new Error(`No se encontró selector en ningún frame: ${selector}`);
  }

  async typeInAnyFrame(page, selector, value, timeoutMs = 15000) {
    const { frame } = await this.waitForSelectorInAnyFrame(page, selector, timeoutMs);
    await frame.focus(selector);
    await frame.click(selector, { clickCount: 3 });
    await frame.keyboard.press("Backspace");
    await frame.type(selector, String(value), { delay: 20 });
  }

  async clickInAnyFrame(page, selector, timeoutMs = 15000) {
    const { handle } = await this.waitForSelectorInAnyFrame(page, selector, timeoutMs);
    await handle.click();
  }

  async readTextIfExistsInAnyFrame(page, selector) {
    for (const frame of page.frames()) {
      const txt = await frame
        .$eval(selector, (el) => {
          const style = window.getComputedStyle(el);
          const visible = style && style.display !== "none" && style.visibility !== "hidden";
          return visible ? (el.textContent || "").trim() : "";
        })
        .catch(() => "");
      if (txt) return txt;
    }
    return "";
  }

  async savePageAsPDF(page, filePath) {
    await page.emulateMediaType("screen");
    await page.pdf({
      path: filePath,
      format: "A4",
      printBackground: true,
      margin: { top: "10mm", right: "10mm", bottom: "10mm", left: "10mm" },
    });
  }

  // ---------------------------
  // Parte 1: Bases/cuotas + guardar PDF en nueva pestaña con CDP/Fetch
  // ---------------------------
  async ejecutarParte1(page, browser, { naf1, naf2, anioEconomico, outPdfPath }) {
    await page.goto("https://w2.seg-social.es/fs/indexframes.html", { waitUntil: "domcontentloaded" });

    // Certificado manual
    await this.esperar(1500);

    await this.clickLinkByTextInAnyFrame(page, "Cotización RETA");
    await this.clickLinkByTextInAnyFrame(page, "Consulta de bases y cuotas ingresadas");

    await this.typeInAnyFrame(page, "#SDFWPROVNAF", naf1);
    await this.typeInAnyFrame(page, "#SDFWRESTONAF", naf2);
    await this.typeInAnyFrame(page, "#SDFWAOMAPA", anioEconomico);

    await this.clickInAnyFrame(page, "#Sub2207101004_35"); // Continuar

    // Error DIL -> saltar registro (según tu criterio)
    await this.esperar(900);
    const dil = await this.readTextIfExistsInAnyFrame(page, "#DIL");
    if (dil) throw new Error(`Error DIL en Parte 1: ${dil}`);

    // Imprimir abre nueva pestaña -> descargar PDF real via CDP/Fetch
    await descargarPdfConReintento({
      label: "CUOTAS_Y_BASES",
      reintentos: 2,
      openPdfFn: async () => {
        await this.clickInAnyFrame(page, "#Sub2204801005_67"); // Imprimir
      },
      getPopupFn: async () => {
        return await waitForPopup(browser, page, 45000);
      },
      downloadFn: async (popup) => {
        // Muy robusto: intercepta "*" y valida %PDF.
        // Si luego quieres afinar, podemos restringir patterns cuando veamos la URL real.
        await descargarPdfRawViaFetchCDP(popup, outPdfPath, 90000, {
          fetchPatterns: [{ urlPattern: "*", requestStage: "Response" }],
          forceReload: true,
        });
      },
      closePopupFn: async (popup) => {
        try { await popup.close(); } catch (_) {}
      },
    });
  }

  // ---------------------------
  // Parte 2: Recibos (autorizado fijo 316077 + primer recibo)
  // ---------------------------
  async ejecutarParte2(page, browser, { naf1, naf2, outPdfPath }) {
    await page.goto("https://w2.seg-social.es/fs/indexframes.html", { waitUntil: "domcontentloaded" });
    await this.esperar(1200);

    await this.clickLinkByTextInAnyFrame(page, "Cotización RETA");
    await this.clickLinkByTextInAnyFrame(page, "Consulta de recibos emitidos régimen de autónomos");

    await page.waitForTimeout(1500);

    // Capturar dialogs (popups nativos)
    let lastDialogMsg = "";
    page.removeAllListeners("dialog");
    page.on("dialog", async (dialog) => {
      lastDialogMsg = dialog.message();
      await dialog.dismiss().catch(() => {});
    });

    // ✅ Autorizado fijo 316077
    const clickedById = await page
      .evaluate(() => {
        const el = document.querySelector("#enlace_316077");
        if (el) {
          el.scrollIntoView({ block: "center" });
          el.click();
          return true;
        }
        return false;
      })
      .catch(() => false);

    if (!clickedById) {
      const clickedByText = await page
        .evaluate(() => {
          const links = Array.from(document.querySelectorAll("a"));
          const t = links.find((a) => (a.textContent || "").trim() === "316077");
          if (t) {
            t.scrollIntoView({ block: "center" });
            t.click();
            return true;
          }
          return false;
        })
        .catch(() => false);

      if (!clickedByText) throw new Error("No se encontró el enlace del autorizado 316077.");
    }

    await page.waitForSelector("#seleccion_1", { timeout: 20000 });
    await page.select("#seleccion_1", "0521");
    await page.select("#seleccion_3", "07");

    // Inputs NAF
    await page.focus("#idTexto1");
    await page.click("#idTexto1", { clickCount: 3 });
    await page.keyboard.press("Backspace");
    await page.type("#idTexto1", naf1, { delay: 20 });

    await page.focus("#idTexto2");
    await page.click("#idTexto2", { clickCount: 3 });
    await page.keyboard.press("Backspace");
    await page.type("#idTexto2", naf2, { delay: 20 });

    await page.click("#botConRegIde");

    if (lastDialogMsg) throw new Error(`Popup/diálogo en Parte 2: ${lastDialogMsg}`);

    // Aviso importante
    await page.waitForSelector("#cheAviImport", { timeout: 20000 });
    await page.click("#cheAviImport").catch(() => {});
    await page.click("#botContAviso");

    // ✅ Primer recibo (primer enlace detalle)
    await page.waitForSelector("a.enlaceFuncDetalle", { timeout: 20000 });
    await page.click("a.enlaceFuncDetalle");

    await page.waitForTimeout(1500);

    // Guardar PDF: preferimos botón imprimir si abre pestaña, si no PDF del detalle
    let printed = false;

    const printSelectorCandidates = [
      "button[title*='Imprimir']",
      "a[title*='Imprimir']",
      "a[href*='IMPRIMIR']",
      "button[onclick*='print']",
      "a[onclick*='print']",
    ];

    for (const sel of printSelectorCandidates) {
      const exists = await page.$(sel);
      if (!exists) continue;

      try {
        const printPage = await this.waitForNewPageFromAction(browser, async () => {
          await page.click(sel);
        }, 8000);

        await printPage.bringToFront();
        await printPage.waitForTimeout(1200);
        await this.savePageAsPDF(printPage, outPdfPath);
        await printPage.close();
        printed = true;
      } catch (_) {
        // No abrió pestaña: fallback PDF del detalle
      }
      break;
    }

    if (!printed) {
      await this.savePageAsPDF(page, outPdfPath);
    }
  }

  async waitForNewPageFromAction(browser, actionFn, timeoutMs = 15000) {
    const p = new Promise((resolve, reject) => {
      const t = setTimeout(() => reject(new Error("No se abrió nueva pestaña a tiempo")), timeoutMs);
      browser.once("targetcreated", async (target) => {
        clearTimeout(t);
        try {
          const newPage = await target.page();
          resolve(newPage);
        } catch (e) {
          reject(e);
        }
      });
    });

    await actionFn();
    return await p;
  }

  // ---------------------------
  // MAIN
  // ---------------------------
  async run(argumentos) {
    const chromeExe = argumentos.googleChrome;
    const excelPath = argumentos.excelClientes;
    const outDir = argumentos.rutaSalida;

    this.ensureDir(outDir);

    const logPath = path.join(outDir, `LOG_BASES_RECIBOS_${this.nowStamp()}.txt`);
    const logLines = [];
    const log = (msg) => {
      const line = `[${new Date().toISOString()}] ${msg}`;
      logLines.push(line);
      console.log(line);
      fs.writeFileSync(logPath, logLines.join("\n"), "utf-8");
    };

    log("INICIO proceso: Bases y recibos al cobro autónomos");

    // 1) Leer Excel
    const excelInfo = await this.leerExcel(excelPath);
    log(`Excel leído. headerRow=${excelInfo.headerRow}. Registros detectados=${excelInfo.registros.length}`);

    // 2) Validar
    const validos = [];
    const invalidos = [];

    for (const r of excelInfo.registros) {
      const errs = this.validarRegistro(r.data);
      if (errs.length) invalidos.push({ ...r, errores: errs });
      else validos.push(r);
    }

    log(`Validación: válidos=${validos.length} | inválidos=${invalidos.length}`);
    for (const inv of invalidos) {
      log(`SKIP fila ${inv.row} DNI=${inv.data.DNI || ""} -> ${inv.errores.join(" | ")}`);
    }

    // 3) Browser
    const browser = await puppeteer.launch({
      headless: false,
      executablePath: chromeExe,
      defaultViewport: null,
      args: ["--start-maximized", "--disable-notifications", "--no-sandbox", "--disable-dev-shm-usage"],
    });

    const page = await browser.newPage();
    page.setDefaultTimeout(25000);

    let ok = 0, ko = 0;
    const startAll = Date.now();

    // 4) Procesar registros
    for (let i = 0; i < validos.length; i++) {
      const rec = validos[i].data;
      const idx = `${i + 1}/${validos.length}`;

      const dniFolder = this.dniToFolder(rec.DNI);
      const carpetaRegistro = path.join(outDir, dniFolder);
      this.ensureDir(carpetaRegistro);

      const mesAnio = this.getMesAnioString(new Date()); // ✅ mes actual
      const adminSafe = this.safeFileName(rec.ADMINISTRADOR);

      const pdf1 = path.join(
        carpetaRegistro,
        this.safeFileName(`CUOTAS Y BASES INGRESADAS ${adminSafe} ${mesAnio}`) + ".pdf"
      );
      const pdf2 = path.join(
        carpetaRegistro,
        this.safeFileName(`RECIBOS AL COBRO ${adminSafe} ${mesAnio}`) + ".pdf"
      );

      log(`Procesando ${idx} | fila=${validos[i].row} | DNI=${rec.DNI} | NAF=${rec.NAF1}-${rec.NAF2}`);

      let success = false;
      let lastErr = "";

      for (let intento = 1; intento <= this.MAX_REINTENTOS_POR_REGISTRO; intento++) {
        try {
          log(`  Intento ${intento}/${this.MAX_REINTENTOS_POR_REGISTRO}`);

          // ✅ Parte 1 (si falla -> se salta completo; no entra Parte 2)
          await this.withTimeout(
            this.ejecutarParte1(page, browser, {
              naf1: rec.NAF1,
              naf2: rec.NAF2,
              anioEconomico: this.DEFAULT_ANIO_ECONOMICO,
              outPdfPath: pdf1,
            }),
            100000,
            "Timeout Parte 1 (bases/cuotas)"
          );

          // ✅ Parte 2 solo si Parte 1 OK
          await this.withTimeout(
            this.ejecutarParte2(page, browser, {
              naf1: rec.NAF1,
              naf2: rec.NAF2,
              outPdfPath: pdf2,
            }),
            110000,
            "Timeout Parte 2 (recibos)"
          );

          log(`  OK -> PDFs guardados:\n    - ${pdf1}\n    - ${pdf2}`);
          success = true;
          break;
        } catch (e) {
          lastErr = e.message || String(e);
          log(`  ERROR intento ${intento}: ${lastErr}`);

          // Limpieza para evitar estados raros
          try { await page.goto("about:blank"); } catch (_) {}
          await this.esperar(1200);
        }
      }

      if (success) ok++;
      else {
        ko++;
        log(`  FAIL definitivo | DNI=${rec.DNI} | Motivo: ${lastErr}`);
      }
    }

    const totalMs = Date.now() - startAll;
    log("FIN proceso");
    log(`RESUMEN -> OK=${ok} | KO=${ko} | Invalidos=${invalidos.length} | Tiempo=${Math.round(totalMs / 1000)}s`);

    await browser.close();
    return true;
  }
}

module.exports = BasesYRecibosAutonomos;
