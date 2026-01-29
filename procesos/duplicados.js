const path = require("path");
const fs = require("fs");
const XlsxPopulate = require("xlsx-populate");
const { registrarEjecucion } = require("../metricas");
const puppeteer = require("puppeteer");

/**
 * Procesos de Duplicados (TA2 / SS)
 */
class ProcesosDuplicados {
  constructor(pathToDbFolder, nombreProyecto, proyectoDB) {
    this.pathToDbFolder = pathToDbFolder;
    this.nombreProyecto = nombreProyecto;
    this.proyectoDB = proyectoDB;
  }

  async esperar(ms) {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }

  getCurrentDateString() {
    const d = new Date();
    const yyyy = d.getFullYear();
    const mm = String(d.getMonth() + 1).padStart(2, "0");
    const dd = String(d.getDate()).padStart(2, "0");
    return `${yyyy}-${mm}-${dd}`;
  }

  // =========================
  // Helpers / normalizadores
  // =========================
  _stripDiacritics(str) {
    return String(str ?? "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "");
  }

  _normHeader(str) {
    return this._stripDiacritics(str)
      .toLowerCase()
      .trim()
      .replace(/\s+/g, " ")
      .replace(/[^\w\s]/g, "");
  }

  _safeFileName(str) {
    return String(str ?? "")
      .trim()
      .replace(/[\\\/:*?"<>|]/g, " ")
      .replace(/\s+/g, " ")
      .trim();
  }

  _dniNorm(dni) {
    return String(dni ?? "")
      .toUpperCase()
      .replace(/\s+/g, "")
      .trim();
  }

  _digitsOnly(val) {
    return String(val ?? "").replace(/\D/g, "");
  }

  _padLeftDigits(val, len) {
    const s = this._digitsOnly(val);
    return s.padStart(len, "0");
  }

  async ensureDir(dir) {
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
  }

  // =========================
  // Excel: lectura formato NUEVO
  // =========================
  async leerExcelDuplicados(pathExcel) {
    const SPEC = {
      exp: { col: "Emp->Código_de_la_Empresa" },
      empresa: { col: "Emp->Nombre_de_la_Empresa" },
      prov_ccc: { col: "Cent->2_primeras_cifras_Segsoc" },
      ccc: {
        parts: [
          "Cent->7_siguientes_cifras_SegSoc",
          "Cent->2_últimas_cifras_SegSoc",
        ],
      },
      trabajador: { col: "Trab->Apellidos_y_Nombre_del_Trabajador" },
      dni: { col: "Trab->DNI_del_Trabajador" },
      prov_naf: { col: "Trab->2_primeras_cifras_Segsoc" },
      naf: {
        parts: [
          "Trab->8_siguientes_cifras_SegSoc",
          "Trab->2_últimas_cifras_SegSoc",
        ],
      },
    };

    const tryReadWithXlsxPopulate = async () => {
      const wb = await XlsxPopulate.fromFileAsync(path.normalize(pathExcel));
      const sh = wb.sheet(0);
      const used = sh.usedRange();
      const numRows = used ? used._numRows : 0;
      const numCols = used ? used._numColumns : 0;
      const getCell = (r, c) => sh.cell(r, c).value();
      return { numRows, numCols, getCell };
    };

    const tryReadWithSheetJS = async () => {
      let XLSX = null;
      try {
        XLSX = require("xlsx");
      } catch (_) {
        XLSX = null;
      }
      if (!XLSX) {
        throw new Error(
          "No se pudo leer el Excel. Si es .xls, convierte a .xlsx o instala la dependencia 'xlsx' (SheetJS).",
        );
      }

      const wb = XLSX.readFile(path.normalize(pathExcel), { cellDates: true });
      const sheetName = wb.SheetNames[0];
      const ws = wb.Sheets[sheetName];
      const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true });

      const numRows = aoa.length;
      const numCols = aoa.reduce((m, row) => Math.max(m, row.length), 0);
      const getCell = (r, c) =>
        (aoa[r - 1] && aoa[r - 1][c - 1]) ?? undefined;
      return { numRows, numCols, getCell };
    };

    let reader = null;
    try {
      reader = await tryReadWithXlsxPopulate();
    } catch (e) {
      reader = await tryReadWithSheetJS();
    }

    const { numRows, numCols, getCell } = reader;

    const maxScan = Math.min(25, numRows);
    let headerRow = null;

    const findCol = (rowHeaders, headerName) => {
      const target = this._normHeader(headerName);
      const found = rowHeaders.find((h) => h.norm === target);
      return found ? found.c : null;
    };

    const headerMap = {};

    for (let r = 1; r <= maxScan; r++) {
      const rowHeaders = [];
      for (let c = 1; c <= numCols; c++) {
        const v = getCell(r, c);
        if (v === null || v === undefined || v === "") continue;
        rowHeaders.push({ c, norm: this._normHeader(v) });
      }
      if (!rowHeaders.length) continue;

      const dniCol = findCol(rowHeaders, SPEC.dni.col);
      const expCol = findCol(rowHeaders, SPEC.exp.col);

      const provNafCol = findCol(rowHeaders, SPEC.prov_naf.col);
      const provCccCol = findCol(rowHeaders, SPEC.prov_ccc.col);

      const nafP1 = findCol(rowHeaders, SPEC.naf.parts[0]);
      const nafP2 = findCol(rowHeaders, SPEC.naf.parts[1]);

      const cccP1 = findCol(rowHeaders, SPEC.ccc.parts[0]);
      const cccP2 = findCol(rowHeaders, SPEC.ccc.parts[1]);

      if (
        dniCol &&
        expCol &&
        provNafCol &&
        provCccCol &&
        nafP1 &&
        nafP2 &&
        cccP1 &&
        cccP2
      ) {
        headerRow = r;

        headerMap.exp = expCol;
        headerMap.empresa = findCol(rowHeaders, SPEC.empresa.col);
        headerMap.trabajador = findCol(rowHeaders, SPEC.trabajador.col);
        headerMap.dni = dniCol;

        headerMap.prov_naf = provNafCol;
        headerMap.prov_ccc = provCccCol;

        headerMap.naf = { p1: nafP1, p2: nafP2 };
        headerMap.ccc = { p1: cccP1, p2: cccP2 };
        break;
      }
    }

    if (!headerRow || !headerMap.dni || !headerMap.naf || !headerMap.ccc) {
      throw new Error(
        "No se encontró una fila de cabecera válida con el formato nuevo. Revisa las columnas Trab->DNI..., Trab->8_siguientes..., Trab->2_últimas..., Cent->7_siguientes..., Cent->2_últimas..., etc.",
      );
    }

    const rows = [];
    for (let r = headerRow + 1; r <= numRows; r++) {
      const dni = getCell(r, headerMap.dni);
      if (dni === null || dni === undefined || String(dni).trim() === "") break;

      const reg = {
        exp: getCell(r, headerMap.exp),
        empresa: getCell(r, headerMap.empresa),
        provCCC: getCell(r, headerMap.prov_ccc),

        ccc7: getCell(r, headerMap.ccc.p1),
        ccc2: getCell(r, headerMap.ccc.p2),

        trabajador: getCell(r, headerMap.trabajador),
        dni: getCell(r, headerMap.dni),

        provNAF: getCell(r, headerMap.prov_naf),
        naf8: getCell(r, headerMap.naf.p1),
        naf2: getCell(r, headerMap.naf.p2),

        _row: r,
      };

      rows.push(this.normalizarRegistro(reg));
    }

    return { rows, headerRow, headerMap };
  }

  normalizarRegistro(r) {
    const dni = this._dniNorm(r.dni);

    const provCCC = this._padLeftDigits(r.provCCC, 2);
    const provNAF = this._padLeftDigits(r.provNAF, 2);

    const ccc7 = this._padLeftDigits(r.ccc7, 7);
    const ccc2 = this._padLeftDigits(r.ccc2, 2);
    const ccc = `${ccc7}${ccc2}`;

    const naf8 = this._padLeftDigits(r.naf8, 8);
    const naf2 = this._padLeftDigits(r.naf2, 2);
    const naf = `${naf8}${naf2}`;

    return {
      exp: String(r.exp ?? "").trim(),
      empresa: String(r.empresa ?? "").trim(),
      regimen: "",
      provCCC,
      ccc,
      trabajador: String(r.trabajador ?? "").trim(),
      dni,
      provNAF,
      naf,
      _row: r._row,
    };
  }

  validarRegistro(r) {
    const errores = [];

    if (!r.dni) errores.push("DNI vacío");
    if (!r.trabajador) errores.push("TRABAJADOR/A vacío");
    if (!r.regimen) errores.push("REGIMEN vacío (input manual)");
    if (!r.ccc) errores.push("CCC vacío");
    if (!r.naf) errores.push("NAF vacío");

    if (r.provCCC && !/^\d{2}$/.test(r.provCCC))
      errores.push("PROV CCC no parece 2 dígitos");
    if (r.provNAF && !/^\d{2}$/.test(r.provNAF))
      errores.push("PROV NAF no parece 2 dígitos");

    if (r.ccc && !/^\d{9}$/.test(r.ccc))
      errores.push("CCC no parece 9 dígitos (7+2)");
    if (r.naf && !/^\d{10}$/.test(r.naf))
      errores.push("NAF no parece 10 dígitos (8+2)");

    if (r.regimen && !/^\d{4}$/.test(r.regimen))
      errores.push("REGIMEN no parece 4 dígitos (ej: 0111)");

    return errores;
  }

  deduplicarPorDNI(rows) {
    const seen = new Set();
    const kept = [];
    const skipped = [];

    for (const r of rows) {
      if (!r.dni) {
        kept.push(r);
        continue;
      }
      if (seen.has(r.dni)) {
        skipped.push({
          dni: r.dni,
          row: r._row,
          reason: "SKIP: DNI duplicado (se procesa la primera aparición)",
        });
        continue;
      }
      seen.add(r.dni);
      kept.push(r);
    }
    return { kept, skipped };
  }

  // =========================
  // Web helpers
  // =========================
  async findFrameWithSelector(page, selector, timeoutMs = 25000, pollMs = 400) {
    const start = Date.now();
    while (Date.now() - start < timeoutMs) {
      for (const fr of page.frames()) {
        try {
          const el = await fr.$(selector);
          if (el) return fr;
        } catch (_) {}
      }
      await this.esperar(pollMs);
    }
    return null;
  }

  async clickLinkInFrames(page, { hrefIncludes, textIncludes }, timeoutMs = 25000) {
    const norm = (s) =>
      String(s ?? "")
        .trim()
        .toLowerCase()
        .replace(/\s+/g, " ");
    const targetText = textIncludes ? norm(textIncludes) : null;

    const start = Date.now();
    while (Date.now() - start < timeoutMs) {
      for (const fr of page.frames()) {
        try {
          if (hrefIncludes) {
            const a = await fr.$(`a[href*="${hrefIncludes}"]`);
            if (a) {
              await a.evaluate((el) => (el.target = "_self"));
              await a.click({ delay: 40 });
              return true;
            }
          }

          if (targetText) {
            const ok = await fr.evaluate((t) => {
              const norm2 = (s) =>
                (s || "").trim().toLowerCase().replace(/\s+/g, " ");
              const a = Array.from(document.querySelectorAll("a")).find((x) =>
                norm2(x.textContent).includes(t),
              );
              if (a) {
                a.target = "_self";
                a.click();
                return true;
              }
              return false;
            }, targetText);
            if (ok) return true;
          }
        } catch (_) {}
      }
      await this.esperar(400);
    }
    return false;
  }

  async safeScreenshot(page, fullPathPng) {
    try {
      await page.screenshot({ path: fullPathPng, fullPage: true });
      return true;
    } catch (e) {
      console.warn("[DUPLICADOS] No se pudo guardar screenshot:", e?.message || e);
      return false;
    }
  }

  async waitForLabelInAnyFrame(page, textIncludes, timeoutMs = 70000, pollMs = 500) {
    const target = String(textIncludes ?? "").trim();
    const start = Date.now();

    while (Date.now() - start < timeoutMs) {
      for (const fr of page.frames()) {
        try {
          const found = await fr.evaluate((t) => {
            const labels = Array.from(document.querySelectorAll("label"));
            return labels.some((l) => (l.textContent || "").includes(t));
          }, target);

          if (found) return fr;
        } catch (_) {}
      }

      await this.esperar(pollMs);
    }

    return null;
  }

  async detectPossibleErrorInFrames(page) {
    const patterns = [
      "error",
      "no existe",
      "no se ha encontrado",
      "datos incorrectos",
      "debe introducir",
      "no se pudo",
      "se ha producido",
    ];
    for (const fr of page.frames()) {
      try {
        const hit = await fr.evaluate((pats) => {
          const txt = (document.body ? document.body.innerText : "") || "";
          const low = txt.toLowerCase();
          return pats.find((p) => low.includes(p));
        }, patterns);

        if (hit)
          return `Posible error detectado en pantalla (contiene '${hit}')`;
      } catch (_) {}
    }
    return null;
  }

  async clickContinuarRobusta({ page, frameForm, timeoutMs = 30000 }) {
    const primarySel = "#Sub2207601004";

    const tryClickBySelector = async (sel) => {
      const el = await frameForm.$(sel);
      if (!el) return false;

      try {
        await el.evaluate((node) =>
          node.scrollIntoView({ block: "center", inline: "center" }),
        );
      } catch (_) {}

      try {
        await el.click({ delay: 30 });
        return true;
      } catch (_) {}

      try {
        const ok = await frameForm.evaluate((s) => {
          const btn = document.querySelector(s);
          if (!btn) return false;
          btn.scrollIntoView({ block: "center", inline: "center" });
          btn.click();
          return true;
        }, sel);
        return !!ok;
      } catch (_) {}

      return false;
    };

    const tryClickByText = async (text) => {
      try {
        const ok = await frameForm.evaluate((t) => {
          const norm = (s) =>
            (s || "").replace(/\s+/g, " ").trim().toLowerCase();
          const target = norm(t);

          const candidates = [
            ...Array.from(document.querySelectorAll("button")),
            ...Array.from(
              document.querySelectorAll(
                'input[type="button"], input[type="submit"]',
              ),
            ),
            ...Array.from(document.querySelectorAll("a")),
          ];

          const el = candidates.find((x) => {
            const label = x.tagName === "INPUT" ? x.value : x.textContent;
            return norm(label).includes(target);
          });

          if (!el) return false;
          el.scrollIntoView({ block: "center", inline: "center" });
          el.click();
          return true;
        }, text);
        return !!ok;
      } catch (_) {
        return false;
      }
    };

    const start = Date.now();
    while (Date.now() - start < timeoutMs) {
      try {
        await frameForm.waitForSelector(primarySel, { timeout: 5000 });
      } catch (_) {}

      if (await tryClickBySelector(primarySel)) return true;
      if (await tryClickByText("Continuar")) return true;

      await this.esperar(600);
    }

    return false;
  }

  async waitForPopup(browser, openerPage, timeoutMs = 45000) {
    const target = await browser
      .waitForTarget(
        (t) => {
          try {
            return (
              t.type() === "page" &&
              t.opener() &&
              t.opener() === openerPage.target()
            );
          } catch (_) {
            return false;
          }
        },
        { timeout: timeoutMs },
      )
      .catch(() => null);

    if (!target) return null;
    return target.page();
  }

  // =========================
  // ✅ OPTIMO REAL: Descarga PDF capturando la RESPUESTA de red (CDP)
  // =========================
  /**
   * Descarga el PDF desde la pestaña popup capturando el body de la respuesta
   * desde la propia sesión de Chrome (CDP).
   *
   * Esto evita el 403 de Axios (certificado/sesión) y no depende del visor/shadow DOM.
   */
  async descargarPdfDesdeRespuestaCDP(popupPage, outputPath, timeoutMs = 90000) {
    await popupPage.bringToFront().catch(() => {});
    console.log("[DUPLICADOS] Popup URL:", popupPage.url());

    const client = await popupPage.target().createCDPSession();
    await client.send("Network.enable");
    await client.send("Network.setCacheDisabled", { cacheDisabled: true }).catch(() => {});

    const timer = new Promise((_, rej) =>
      setTimeout(() => rej(new Error("Timeout esperando la respuesta del PDF en el popup.")), timeoutMs),
    );

    const pdfPromise = new Promise((resolve, reject) => {
      const onResponse = async (params) => {
        try {
          const url = String(params?.response?.url || "");
          const status = Number(params?.response?.status || 0);
          const mime = String(params?.response?.mimeType || "").toLowerCase();
          const headers = params?.response?.headers || {};

          const ct = String(headers["content-type"] || headers["Content-Type"] || "").toLowerCase();
          const cd = String(headers["content-disposition"] || headers["Content-Disposition"] || "").toLowerCase();

          const looksPdf =
            url.includes("/ImprPDF/") ||
            url.includes("InSeNaCoder") ||
            url.toLowerCase().endsWith(".pdf") ||
            mime.includes("pdf") ||
            ct.includes("pdf") ||
            cd.includes("pdf") ||
            cd.includes("attachment");

          if (!looksPdf) return;

          if (!(status >= 200 && status < 300)) {
            console.log("[DUPLICADOS][CDP] response NO OK:", status, url);
            return;
          }

          console.log("[DUPLICADOS][CDP] response OK:", status, url);
          console.log("[DUPLICADOS][CDP] mime:", mime);
          console.log("[DUPLICADOS][CDP] content-type:", ct);
          console.log("[DUPLICADOS][CDP] content-disposition:", cd);

          client.off("Network.responseReceived", onResponse);

          const { body, base64Encoded } = await client.send("Network.getResponseBody", {
            requestId: params.requestId,
          });

          resolve({ body, base64Encoded });
        } catch (e) {
          reject(e);
        }
      };

      client.on("Network.responseReceived", onResponse);
    });

    // ✅ Clave: recargamos DESPUÉS de enganchar listeners para no llegar tarde
    await popupPage.reload({ waitUntil: "domcontentloaded" }).catch(() => {});
    await popupPage.waitForSelector("body", { timeout: 20000 }).catch(() => {});
    await popupPage.waitForFunction(() => document.readyState !== "loading", { timeout: 30000 }).catch(() => {});
    await this.esperar(800);

    const { body, base64Encoded } = await Promise.race([pdfPromise, timer]);

    const buffer = base64Encoded ? Buffer.from(body, "base64") : Buffer.from(body, "utf8");
    fs.writeFileSync(outputPath, buffer);

    try {
      await client.detach();
    } catch (_) {}
  }

  // =========================
  // PROCESO: DUPLICADOS TA2
  // =========================
  async duplicadosTa2(argumentos) {
    console.log("[DUPLICADOS] Iniciando proceso DUPLICADOS TA2");

    const nombreProceso = "DUPLICADOS TA2";
    let registrosProcesados = 0;

    return new Promise(async (resolve) => {
      let browser = null;

      try {
        const chromeExePath = argumentos?.formularioControl?.[0];
        const pathExcel = argumentos?.formularioControl?.[1];
        const regimenManual = argumentos?.formularioControl?.[2];
        const pathSalidaBase = argumentos?.formularioControl?.[3];

        const regimen4 = this._padLeftDigits(regimenManual || "0111", 4);

        if (!chromeExePath || !fs.existsSync(chromeExePath)) {
          console.error("[DUPLICADOS] Ruta a chrome.exe no válida.");
          return resolve(false);
        }
        if (!pathExcel || typeof pathExcel !== "string" || !fs.existsSync(pathExcel)) {
          console.error("[DUPLICADOS] Ruta a Excel no válida.");
          return resolve(false);
        }
        if (!pathSalidaBase || typeof pathSalidaBase !== "string" || !pathSalidaBase.trim()) {
          console.error("[DUPLICADOS] Ruta de salida no válida.");
          return resolve(false);
        }
        if (!/^\d{4}$/.test(regimen4)) {
          console.error("[DUPLICADOS] Régimen inválido. Debe ser 4 dígitos (ej: 0111).");
          return resolve(false);
        }

        const rootOut = path.join(
          path.normalize(pathSalidaBase),
          `Duplicados TA2 (${this.getCurrentDateString()})`,
        );
        const dirPdf = path.join(rootOut, "PDF");
        const dirPng = path.join(rootOut, "CAPTURAS");
        const dirLogs = path.join(rootOut, "LOGS");

        await this.ensureDir(rootOut);
        await this.ensureDir(dirPdf);
        await this.ensureDir(dirPng);
        await this.ensureDir(dirLogs);

        const resumenPath = path.join(dirLogs, "resumen.json");
        let resumen = { ok: [], error: [], skipped: [] };
        if (fs.existsSync(resumenPath)) {
          try {
            resumen = JSON.parse(fs.readFileSync(resumenPath, "utf8"));
          } catch (_) {}
        }
        const okSet = new Set((resumen.ok || []).map((x) => this._dniNorm(x)));

        console.log("[DUPLICADOS] Leyendo Excel:", path.normalize(pathExcel));
        const { rows } = await this.leerExcelDuplicados(pathExcel);

        for (const r of rows) r.regimen = regimen4;

        const logsPorDni = new Map();
        const validRows = [];

        for (const r of rows) {
          const errs = this.validarRegistro(r);
          if (errs.length) {
            logsPorDni.set(
              r.dni || `ROW_${r._row}`,
              `ERROR: ${errs.join(" | ")}`,
            );
            continue;
          }
          validRows.push(r);
        }

        const { kept, skipped } = this.deduplicarPorDNI(validRows);
        for (const s of skipped) {
          logsPorDni.set(s.dni, s.reason);
          resumen.skipped.push(s);
        }

        console.log(
          `[DUPLICADOS] Registros leídos: ${rows.length}. Válidos: ${validRows.length}. Tras dedupe: ${kept.length}. Skips: ${skipped.length}.`,
        );

        if (!kept.length) {
          console.warn("[DUPLICADOS] No hay registros válidos para procesar.");
          fs.writeFileSync(resumenPath, JSON.stringify(resumen, null, 2), "utf8");
          return resolve(false);
        }

        const urlFS = "https://w2.seg-social.es/fs/indexframes.html";

        browser = await puppeteer.launch({
          headless: false,
          defaultViewport: null,
          executablePath: chromeExePath,
          args: [
            "--start-maximized",
            "--no-sandbox",
            "--disable-setuid-sandbox",
            "--disable-features=IsolateOrigins,site-per-process",
            "--disable-popup-blocking",
          ],
        });

        const opened = await browser.pages();
        const page = opened.length ? opened[0] : await browser.newPage();

        page.on("dialog", async (dialog) => {
          try {
            await dialog.accept();
          } catch (_) {}
        });

        await page.goto(urlFS, { waitUntil: "domcontentloaded" });
        console.log("[DUPLICADOS] FS abierto. Selecciona el certificado si aparece.");

        const openAFIOnlineReal = async () => {
          const ok = await this.clickLinkInFrames(page, {
            hrefIncludes: "menuAFI-REMESAS.html",
            textIncludes: "Inscripción y Afiliación Online Real",
          });
          if (!ok) throw new Error("No se pudo clicar 'Inscripción y Afiliación Online Real'");
          await this.esperar(1200);
        };

        const openATR65Duplicados = async () => {
          const ok = await this.clickLinkInFrames(page, {
            hrefIncludes: "TRANSACCION=ATR65",
            textIncludes: "Duplicados de documentos trabajador",
          });
          if (!ok) throw new Error("No se pudo clicar 'Duplicados de documentos trabajador' (ATR65)");
          await this.esperar(1200);
        };

        const flushResumen = () => {
          try {
            fs.writeFileSync(resumenPath, JSON.stringify(resumen, null, 2), "utf8");
          } catch (e) {
            console.warn("[DUPLICADOS] No se pudo escribir resumen.json:", e?.message || e);
          }
        };

        const fillInput = async (frame, selector, value, { timeout = 30000 } = {}) => {
          await frame.waitForSelector(selector, { timeout });
          const el = await frame.$(selector);
          if (!el) throw new Error(`No se encontró el input ${selector}`);

          await el.click({ clickCount: 3 });
          await page.keyboard.down("Control");
          await page.keyboard.press("A");
          await page.keyboard.up("Control");
          await page.keyboard.press("Backspace");
          await page.keyboard.type(String(value ?? ""), { delay: 20 });
        };

        const procesarRegistro = async (r, idx) => {
          const dni = this._dniNorm(r.dni);
          const trabajador = this._safeFileName(r.trabajador);

          const pngPath = path.join(dirPng, `Cuadro TA2 SS ${trabajador}.png`);
          const pdfPath = path.join(dirPdf, `TA2 A010112 ${trabajador}.pdf`);

          if (okSet.has(dni)) {
            logsPorDni.set(dni, "SKIP: ya estaba OK (modo resume)");
            return;
          }

          console.log(
            `[DUPLICADOS] Procesando ${idx + 1}/${kept.length} | DNI: ${dni} | TRABAJADOR: ${r.trabajador}`,
          );

          await page.goto(urlFS, { waitUntil: "domcontentloaded" });
          await this.esperar(800);

          await openAFIOnlineReal();
          await openATR65Duplicados();

          const frameForm = await this.findFrameWithSelector(page, "#SDFTESNAF", 30000);
          if (!frameForm) throw new Error("No se encontró el formulario ATR65 (selector #SDFTESNAF)");

          const provNAF = this._padLeftDigits(r.provNAF, 2);
          const naf10 = this._padLeftDigits(r.naf, 10);
          const provCCC = this._padLeftDigits(r.provCCC, 2);
          const ccc9 = this._padLeftDigits(r.ccc, 9);

          await fillInput(frameForm, "#SDFTESNAF", provNAF);
          await fillInput(frameForm, "#SDFNAF", naf10);
          await fillInput(frameForm, "#SDFREGCTA_NH", regimen4);
          await fillInput(frameForm, "#SDFTESCTA", provCCC);
          await fillInput(frameForm, "#SDFCUENTA", ccc9);

          await frameForm.waitForSelector("#ListaTipoImpresion", { timeout: 30000 });
          await frameForm.select("#ListaTipoImpresion", "OnLine");

          const clickedContinuar = await this.clickContinuarRobusta({
            page,
            frameForm,
            timeoutMs: 35000,
          });
          if (!clickedContinuar) {
            await this.safeScreenshot(page, pngPath);
            throw new Error("No se pudo pulsar 'Continuar' (click robusto falló).");
          }

          const frameTabla = await this.waitForLabelInAnyFrame(page, "ALTA (SIT.ACTUAL)", 90000);
          if (!frameTabla) {
            const maybeError = await this.detectPossibleErrorInFrames(page);
            await this.safeScreenshot(page, pngPath);
            throw new Error(
              `No se encontró el listado que contiene 'ALTA (SIT.ACTUAL)' tras continuar.${maybeError ? " " + maybeError : ""}`,
            );
          }

          await this.safeScreenshot(page, pngPath);

          // ✅ PREPARAR CAPTURA POPUP ANTES DEL DOBLE CLICK (evita carreras)
          const popupPromise = this.waitForPopup(browser, page, 45000);

          const didOpen = await frameTabla.evaluate(() => {
            const label = Array.from(document.querySelectorAll("label")).find(
              (l) => (l.textContent || "").includes("ALTA (SIT.ACTUAL)"),
            );
            if (!label) return false;

            label.scrollIntoView({ block: "center", inline: "center" });

            label.dispatchEvent(
              new MouseEvent("dblclick", {
                bubbles: true,
                cancelable: true,
                view: window,
              }),
            );

            label.click();
            label.click();

            return true;
          });

          if (!didOpen) {
            throw new Error("No se encontró 'ALTA (SIT.ACTUAL)' para abrir el PDF");
          }

          const popupPage = await popupPromise;
          if (!popupPage) {
            throw new Error("Se esperaba una nueva pestaña con el PDF, pero no se abrió.");
          }

          // ✅ FLUJO CORRECTO: descargar el PDF desde la sesión del propio Chrome (CDP response)
          await this.descargarPdfDesdeRespuestaCDP(popupPage, pdfPath, 90000);

          logsPorDni.set(dni, `OK: PDF descargado (CDP response) -> ${path.basename(pdfPath)}`);
          resumen.ok.push(dni);
          okSet.add(dni);

          try {
            await popupPage.close();
          } catch (_) {}
        };

        let okCount = 0;
        let errCount = 0;

        for (let i = 0; i < kept.length; i++) {
          registrosProcesados++;
          const r = kept[i];
          const dni = this._dniNorm(r.dni);

          try {
            await procesarRegistro(r, i);
            if ((logsPorDni.get(dni) || "").startsWith("OK")) okCount++;
          } catch (e) {
            errCount++;
            const msg = `ERROR: ${e?.message || e}`;
            logsPorDni.set(dni, msg);
            resumen.error.push({ dni, error: msg });
            console.warn("[DUPLICADOS]", msg);

            const trabajador = this._safeFileName(r.trabajador);
            await this.safeScreenshot(page, path.join(dirPng, `ERROR ${trabajador}.png`)).catch(() => {});
          }

          if ((i + 1) % 5 === 0) flushResumen();
        }

        flushResumen();

        console.log(
          `[DUPLICADOS] Terminado. OK: ${okCount} | ERROR: ${errCount} | Procesados: ${registrosProcesados}`,
        );

        registrarEjecucion({ nombreProceso, registrosProcesados });

        try {
          if (browser) await browser.close();
        } catch (_) {}

        return resolve(true);
      } catch (err) {
        console.error("[DUPLICADOS] Error general:", err?.message || err);
        try {
          if (globalThis?.mainProcess?.mostrarError) {
            await globalThis.mainProcess.mostrarError(
              "No se ha podido completar el proceso",
              "Se ha producido un error interno ejecutando DUPLICADOS TA2.",
            );
          }
        } catch (_) {}
        try {
          if (browser) await browser.close();
        } catch (_) {}
        return resolve(false);
      }
    });
  }

  async dUPLICADOSTA2(argumentos) {
    console.warn("[DUPLICADOS] Alias dUPLICADOSTA2() llamado. Usa duplicadosTa2().");
    return this.duplicadosTa2(argumentos);
  }
}

module.exports = ProcesosDuplicados;
