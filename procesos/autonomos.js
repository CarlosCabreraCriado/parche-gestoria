const path = require("path");
const fs = require("fs");
const XlsxPopulate = require("xlsx-populate");
const puppeteer = require("puppeteer");
const {
  waitForPopup,
  descargarPdfRawViaFetchCDP,
  descargarPdfConReintento,
  waitForPrintPreviewPopup,
  descargarPdfDesdePrintPreview,
} = require("./utils/pdfNuevaPestanaCdp");

class ProcesosBasesRecibosAutonomos {
  constructor(pathToDbFolder, nombreProyecto, proyectoDB) {
    this.pathToDbFolder = pathToDbFolder;
    this.nombreProyecto = nombreProyecto;
    this.proyectoDB = proyectoDB;
  }

  // -------------------------
  // Utilidades
  // -------------------------
  async esperar(ms) {
    return new Promise((r) => setTimeout(r, ms));
  }

  getCurrentDateString() {
    const d = new Date();
    const yyyy = d.getFullYear();
    const mm = String(d.getMonth() + 1).padStart(2, "0");
    const dd = String(d.getDate()).padStart(2, "0");
    return `${yyyy}-${mm}-${dd}`;
  }

  _injectBaseTag(html, baseHref) {
    // Inserta <base> dentro de <head> para que /GestionDomiciliacionCuenta/... resuelva bien
    if (!html) return html;
    if (/<base\s/i.test(html)) return html;

    if (/<head[^>]*>/i.test(html)) {
      return html.replace(
        /<head[^>]*>/i,
        (m) => `${m}\n<base href="${baseHref}">`,
      );
    }
    // fallback raro si no hay <head>
    return `<base href="${baseHref}">\n${html}`;
  }

  _stripHeavyScripts(html) {
    // Quita scripts (gtm, analytics, prosa.js) para que setContent no se quede “esperando”
    // Mantiene los <link rel="stylesheet"...> (que es lo que nos interesa)
    if (!html) return html;
    return html.replace(
      /<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi,
      "",
    );
  }

  _safeFileName(str) {
    return String(str ?? "")
      .trim()
      .replace(/[\\\/\:*?"<>|]/g, " ")
      .replace(/\s+/g, " ")
      .trim();
  }

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

  _digitsOnly(val) {
    return String(val ?? "").replace(/\D/g, "");
  }

  _padLeftDigitsOrEmpty(val, len) {
    const s = this._digitsOnly(val);
    if (!s) return "";
    return s.padStart(len, "0");
  }

  async ensureDir(dir) {
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
  }

  _excelColName(n) {
    // 1 -> A, 2 -> B ...
    let s = "";
    while (n > 0) {
      const m = (n - 1) % 26;
      s = String.fromCharCode(65 + m) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }

  // -------------------------
  // Excel: lectura incluyendo columnas ocultas (XlsxPopulate las devuelve igual)
  // -------------------------
  async leerExcelInput(pathExcel) {
    const wb = await XlsxPopulate.fromFileAsync(path.normalize(pathExcel));
    const sh = wb.sheet(0);
    const used = sh.usedRange();

    const numRows = used ? used._numRows : 0;
    const numCols = used ? used._numColumns : 0;

    const getCell = (r, c) => sh.cell(r, c).value();

    const maxScan = Math.min(30, numRows);
    let headerRow = null;
    let headerMap = null;

    const findInRow = (rowHeaders, name) => {
      const target = this._normHeader(name);
      const found = rowHeaders.find((h) => h.norm === target);
      return found ? found.c : null;
    };

    for (let r = 1; r <= maxScan; r++) {
      const rowHeaders = [];
      for (let c = 1; c <= numCols; c++) {
        const v = getCell(r, c);
        if (v === null || v === undefined || v === "") continue;
        rowHeaders.push({ c, raw: v, norm: this._normHeader(v) });
      }
      if (!rowHeaders.length) continue;

      const colExpte = findInRow(rowHeaders, "EXPTE.");
      const colAdmin = findInRow(rowHeaders, "ADMINISTRADOR");

      if (colExpte && colAdmin) {
        headerRow = r;
        headerMap = {
          EXPTE: colExpte,
          EMPRESA: findInRow(rowHeaders, "EMPRESA"),
          NAF: findInRow(rowHeaders, "NAF"),
          CLAVE: findInRow(rowHeaders, "CLAVE"),
          FALTA_BAJA: findInRow(rowHeaders, "F.ALTA/BAJA"),
          ADMIN: colAdmin,
          BASE: findInRow(rowHeaders, "BASE"),
          TOTAL: findInRow(rowHeaders, "TOTAL"),
          PREV_ANO: findInRow(rowHeaders, "PREV AÑO"),
        };
        break;
      }
    }

    if (!headerRow || !headerMap?.ADMIN) {
      throw new Error(
        "No se encontró la fila de cabecera. Necesito al menos 'EXPTE.' y 'ADMINISTRADOR'.",
      );
    }

    // 2) Forzar columnas G/H para NAF1/NAF2
    // Excel: A=1, B=2 ... G=7, H=8
    const colNAF1 = 7;
    const colNAF2 = 8;

    // 3) Lista de cabeceras completa (debug)
    const headers = [];
    for (let c = 1; c <= numCols; c++) {
      const v = getCell(headerRow, c);
      headers.push({
        col: c,
        excelCol: this._excelColName(c),
        header: v ?? "",
        norm: this._normHeader(v ?? ""),
      });
    }

    // 4) Parsear registros: procesar filas con NAF1/NAF2/ADMIN
    // Criterio de fin: fila vacía real (sin NAF1, NAF2 y ADMIN)
    const rows = [];
    for (let r = headerRow + 1; r <= numRows; r++) {
      const naf1Cell = getCell(r, colNAF1);
      const naf2Cell = getCell(r, colNAF2);
      const adminCell = headerMap.ADMIN ? getCell(r, headerMap.ADMIN) : "";

      const isEmptyRow =
        (naf1Cell === null ||
          naf1Cell === undefined ||
          String(naf1Cell).trim() === "") &&
        (naf2Cell === null ||
          naf2Cell === undefined ||
          String(naf2Cell).trim() === "") &&
        (adminCell === null ||
          adminCell === undefined ||
          String(adminCell).trim() === "");

      if (isEmptyRow) break;

      const rec = {
        EXPTE: headerMap.EXPTE ? getCell(r, headerMap.EXPTE) : "",
        EMPRESA: headerMap.EMPRESA ? getCell(r, headerMap.EMPRESA) : "",
        NAF: headerMap.NAF ? getCell(r, headerMap.NAF) : "",
        CLAVE: headerMap.CLAVE ? getCell(r, headerMap.CLAVE) : "",
        FALTA_BAJA: headerMap.FALTA_BAJA
          ? getCell(r, headerMap.FALTA_BAJA)
          : "",
        ADMIN: adminCell,
        BASE: headerMap.BASE ? getCell(r, headerMap.BASE) : "",
        TOTAL: headerMap.TOTAL ? getCell(r, headerMap.TOTAL) : "",
        PREV_ANO: headerMap.PREV_ANO ? getCell(r, headerMap.PREV_ANO) : "",
        NAF1: naf1Cell,
        NAF2: naf2Cell,
        _row: r,
      };

      rows.push(this.normalizarRegistro(rec));
    }

    return { headers, headerRow, rows };
  }

  normalizarRegistro(r) {
    const naf1 = this._padLeftDigitsOrEmpty(r.NAF1, 2);
    const naf2 = this._padLeftDigitsOrEmpty(r.NAF2, 10);

    return {
      expte: String(r.EXPTE ?? "").trim(),
      empresa: String(r.EMPRESA ?? "").trim(),
      nafRaw: String(r.NAF ?? "").trim(),
      clave: String(r.CLAVE ?? "").trim(),
      fechaAltaBaja: String(r.FALTA_BAJA ?? "").trim(),
      administrador: String(r.ADMIN ?? "").trim(),
      base: String(r.BASE ?? "").trim(),
      total: String(r.TOTAL ?? "").trim(),
      prevAno: String(r.PREV_ANO ?? "").trim(),
      naf1,
      naf2,
      _row: r._row,
    };
  }

  validarRegistro(r) {
    const missing = [];
    const invalid = [];

    const req = (val, msg) => {
      if (val === null || val === undefined || String(val).trim() === "")
        missing.push(msg);
    };

    req(r.administrador, "ADMINISTRADOR vacío");
    req(r.naf1, "NAF1 vacío (columna G)");
    req(r.naf2, "NAF2 vacío (columna H)");

    if (r.naf1 && !/^\d{2}$/.test(r.naf1)) invalid.push("NAF1 no es 2 dígitos");
    if (r.naf2 && !/^\d{10}$/.test(r.naf2))
      invalid.push("NAF2 no es 10 dígitos");

    return { missing, invalid };
  }

  // -------------------------
  // Web helpers (frames + popups)
  // -------------------------
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

  async clickLinkInFrames(
    page,
    { hrefIncludes, textIncludes },
    timeoutMs = 25000,
  ) {
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

  async waitForPrintPreviewAny(browser, openerPage, timeoutMs = 45000) {
    const start = Date.now();

    // A) Caso 1: misma pestaña navega a chrome://print
    const waitSameTab = (async () => {
      while (Date.now() - start < timeoutMs) {
        try {
          const u = openerPage.url();
          if (u && u.startsWith("chrome://print")) return openerPage;
        } catch (_) {}
        await this.esperar(200);
      }
      return null;
    })();

    // B) Caso 2: se abre como popup/target nuevo chrome://print
    const waitNewTarget = browser
      .waitForTarget(
        (t) =>
          t.type() === "page" && (t.url() || "").startsWith("chrome://print"),
        { timeout: timeoutMs },
      )
      .then((t) => t.page().catch(() => null))
      .catch(() => null);

    const printPage = await Promise.race([waitSameTab, waitNewTarget]);

    if (!printPage)
      throw new Error("Timeout esperando vista de impresión (chrome://print).");

    try {
      await printPage.bringToFront();
    } catch (_) {}

    return printPage;
  }

  _isPdfBuffer(buf) {
    if (!buf || !Buffer.isBuffer(buf) || buf.length < 5) return false;
    return buf.subarray(0, 5).toString("utf8") === "%PDF-";
  }

  _isPdfDownloadRetryableError(err) {
    const msg = String(err?.message || err).toLowerCase();
    return (
      msg.includes("no es un pdf") ||
      msg.includes("timeout") ||
      msg.includes("pdf")
    );
  }

  /**
   * Descarga un PDF real capturando la respuesta con CDP Fetch.
   * Muy robusto en visores que devuelven HTML si no se captura el request.
   */
  async descargarPdfRawViaFetchCDP(popupPage, outputPath, timeoutMs = 90000) {
    await popupPage.bringToFront().catch(() => {});
    const client = await popupPage.target().createCDPSession();

    let done = false;
    let onPaused = null;

    try {
      await client.send("Network.enable").catch(() => {});
      await client
        .send("Network.setCacheDisabled", { cacheDisabled: true })
        .catch(() => {});

      await client
        .send("Fetch.enable", {
          patterns: [{ urlPattern: "*", requestStage: "Response" }],
        })
        .catch(() => {});

      const timer = new Promise((_, rej) =>
        setTimeout(
          () => rej(new Error("Timeout esperando el PDF (Fetch CDP).")),
          timeoutMs,
        ),
      );

      const pdfPromise = new Promise((resolve, reject) => {
        onPaused = async (ev) => {
          if (done) {
            try {
              await client
                .send("Fetch.continueRequest", { requestId: ev.requestId })
                .catch(() => {});
            } catch (_) {}
            return;
          }

          try {
            const reqId = ev.requestId;
            const status = ev.responseStatusCode || 0;
            const headers = {};
            for (const h of ev.responseHeaders || []) {
              headers[String(h.name || "").toLowerCase()] = String(
                h.value || "",
              );
            }
            const ctype = (headers["content-type"] || "").toLowerCase();

            const maybePdf =
              status >= 200 &&
              status < 300 &&
              (ctype.includes("application/pdf") ||
                ctype.includes("octet-stream") ||
                ctype.includes("binary") ||
                true);

            if (!maybePdf) {
              await client
                .send("Fetch.continueRequest", { requestId: reqId })
                .catch(() => {});
              return;
            }

            const bodyResp = await client.send("Fetch.getResponseBody", {
              requestId: reqId,
            });
            const buf = bodyResp?.base64Encoded
              ? Buffer.from(bodyResp.body || "", "base64")
              : Buffer.from(bodyResp.body || "", "utf8");

            await client
              .send("Fetch.continueRequest", { requestId: reqId })
              .catch(() => {});

            if (!this._isPdfBuffer(buf)) {
              // No era un PDF real -> seguir escuchando
              return;
            }

            done = true;
            fs.writeFileSync(outputPath, buf);
            resolve(true);
          } catch (e) {
            reject(e);
          }
        };

        client.on("Fetch.requestPaused", onPaused);
      });

      // Forzar reload para disparar el request del PDF
      await popupPage.reload({ waitUntil: "domcontentloaded" }).catch(() => {});
      await popupPage
        .waitForSelector("body", { timeout: 20000 })
        .catch(() => {});
      await this.esperar(800);

      await Promise.race([pdfPromise, timer]);
      return true;
    } finally {
      try {
        if (onPaused) client.removeListener("Fetch.requestPaused", onPaused);
      } catch (_) {}
      try {
        await client.send("Fetch.disable").catch(() => {});
      } catch (_) {}
      try {
        await client.detach();
      } catch (_) {}
    }
  }

  /**
   * Descarga PDF con 1 reintento.
   */
  async descargarPdfConReintento({ openPopupFn, outputPath, label }) {
    let lastErr = null;

    for (let intento = 1; intento <= 2; intento++) {
      let popup = null;
      try {
        popup = await openPopupFn();
        if (!popup) throw new Error("No se abrió el popup/pestaña del PDF.");

        await this.descargarPdfRawViaFetchCDP(popup, outputPath, 90000);

        try {
          await popup.close();
        } catch (_) {}
        return { ok: true, intento };
      } catch (e) {
        lastErr = e;

        try {
          if (popup) await popup.close();
        } catch (_) {}

        if (!this._isPdfDownloadRetryableError(e) || intento === 2) break;

        console.warn(
          `[${label}] Fallo de descarga (intento ${intento}). Reintentando 1 vez...`,
          e?.message || e,
        );
        await this.esperar(1200);
      }
    }

    throw lastErr;
  }

  // -------------------------
  // Selección “recibo más reciente”
  // -------------------------
  async seleccionarLinkReciboMasRecienteEnContexto(ctx /* Page o Frame */) {
    // ctx puede ser: pageB (Page) o frListado (Frame)
    const $all = (sel) => (ctx.$$ ? ctx.$$(sel) : []);
    const $ = (sel) => (ctx.$ ? ctx.$(sel) : null);

    // ✅ Tu HTML real usa paramRecibo=0 y SPM.ACC.AC_VER_RECIBO
    const selectorFuerte =
      'a[href*="SPM.ACC.AC_VER_RECIBO=AC_VER_RECIBO"], a.enlaceFuncDetalle[href*="AC_VER_RECIBO"], a.pr_enlaceLocal[href*="AC_VER_RECIBO"]';

    let links = await $all(selectorFuerte);

    // Fallback más laxo
    if (!links || links.length === 0) {
      links = await $all('a[href*="AC_VER_RECIBO"], a[href*="paramRecibo="]');
    }

    if (!links || links.length === 0) {
      throw new Error(
        "No se encontró ningún enlace 'Ver detalle del recibo' (AC_VER_RECIBO).",
      );
    }

    // Si hay varios, podemos elegir el que tenga mayor paramRecibo (normalmente 0 es el último/primero según orden)
    const scored = [];
    for (const a of links) {
      const href = await a.evaluate((el) => el.getAttribute("href") || "");
      const m = href.match(/paramRecibo=(\d+)/);
      const n = m ? Number(m[1]) : 0;

      // score: cuanto más pequeño, más “reciente” en tu ejemplo (0)
      scored.push({ a, score: n, href });
    }

    scored.sort((x, y) => x.score - y.score);

    return {
      handle: scored[0].a,
      strategy: `paramRecibo_min(${scored[0].score})`,
    };
  }

  async findAnyFrameWithAnySelector(
    page,
    selectors,
    timeoutMs = 60000,
    pollMs = 400,
  ) {
    const start = Date.now();
    while (Date.now() - start < timeoutMs) {
      for (const fr of page.frames()) {
        for (const sel of selectors) {
          try {
            const el = await fr.$(sel);
            if (el) return { frame: fr, selector: sel };
          } catch (_) {}
        }
      }
      await this.esperar(pollMs);
    }
    return null;
  }

  /**
   * Tras clicar "Consulta de recibos emitidos", la UI puede:
   * - cargar dentro de un frame de FS, o
   * - abrir una nueva pestaña/ventana.
   * Esta función devuelve { pageB, frameB } donde están los selectores PROSA.
   */
  async getProsaContext(browser, seedPage, timeoutMs = 60000) {
    const start = Date.now();

    const looksLikeProsaFrame = async (fr) => {
      try {
        return await fr.evaluate(() => {
          // Huellas PROSA / Seguridad Social (según tus HTML)
          const hasTicket =
            !!document.querySelector('input[name="ARQ.SPM.TICKET"]') ||
            !!document.querySelector("#ARQ_SPM_TICKET");
          const hasFrontend = !!document.querySelector("#FRONTEND");
          const hasProsaJs = Array.from(document.scripts || []).some((s) =>
            (s.src || "").includes("prosa.min.js"),
          );
          const hasProsaInput = hasTicket || hasFrontend;
          const hasHeader =
            (document.title || "")
              .toLowerCase()
              .includes("consulta de recibos") ||
            (document.body?.innerText || "")
              .toLowerCase()
              .includes("oficina virtual");

          return (
            (hasProsaJs && (hasProsaInput || hasHeader)) ||
            (hasProsaInput && hasHeader)
          );
        });
      } catch {
        return false;
      }
    };

    const findInPageFrames = async (pg) => {
      for (const fr of pg.frames()) {
        if (await looksLikeProsaFrame(fr)) return { pageB: pg, frameB: fr };
      }
      // A veces PROSA es el main frame
      if (await looksLikeProsaFrame(pg.mainFrame()))
        return { pageB: pg, frameB: pg.mainFrame() };
      return null;
    };

    while (Date.now() - start < timeoutMs) {
      // 1) Intentar en la seedPage (frames)
      const inSeed = await findInPageFrames(seedPage);
      if (inSeed) return inSeed;

      // 2) Intentar en todas las pestañas
      const pages = await browser.pages();
      for (const pg of pages) {
        const url = (pg.url() || "").toLowerCase();
        const urlHint =
          url.includes("prosainternet") ||
          url.includes("gestiondomiciliacioncuenta") ||
          url.includes("xv24e003");
        if (!urlHint && pg !== seedPage) {
          // aunque no haya hint, puede ser que el url sea "about:blank" pero el frame tenga el contenido
        }

        const found = await findInPageFrames(pg);
        if (found) return found;
      }

      await this.esperar(350);
    }

    throw new Error(
      "No se pudo localizar el contexto PROSA (ni en frames ni en pestaña nueva).",
    );
  }

  // -------------------------
  // PROCESO PRINCIPAL
  // -------------------------
  async basesYRecibosAutonomos(argumentos) {
    console.log("[BASES/RECIBOS] Iniciando proceso Bases y Recibos Autónomos");

    const nombreProceso = "Bases y recibos al cobro autónomos";
    let registrosProcesados = 0;

    return new Promise(async (resolve) => {
      let browser = null;

      try {
        const chromeExePath = argumentos?.formularioControl?.[0];
        const pathExcel = argumentos?.formularioControl?.[1];
        const ejercicioEconomicoRaw = argumentos?.formularioControl?.[2];
        const pathSalidaBase = argumentos?.formularioControl?.[3];

        if (!chromeExePath || !fs.existsSync(chromeExePath)) {
          console.error("[BASES/RECIBOS][INPUT] Ruta a chrome.exe no válida.");
          return resolve(false);
        }
        if (!pathExcel || !fs.existsSync(pathExcel)) {
          console.error("[BASES/RECIBOS][INPUT] Ruta a Excel no válida.");
          return resolve(false);
        }
        if (!pathSalidaBase || !String(pathSalidaBase).trim()) {
          console.error("[BASES/RECIBOS][INPUT] Ruta de salida no válida.");
          return resolve(false);
        }
        const ejercicioEconomico = String(ejercicioEconomicoRaw ?? "").trim();
        if (!/^\d{4}$/.test(ejercicioEconomico)) {
          console.error(
            "[BASES/RECIBOS][INPUT] Ejercicio económico inválido. Debe ser AAAA (ej: 2025).",
          );
          return resolve(false);
        }

        // Mes/Año actuales para nombre
        const nowGlobal = new Date();
        const mesActual = String(nowGlobal.getMonth() + 1).padStart(2, "0");
        const anioActual = String(nowGlobal.getFullYear());
        const tagMesAno = `${mesActual}-${anioActual}`;

        const rootOut = path.join(
          path.normalize(pathSalidaBase),
          `BasesYRecibosAutonomos (${this.getCurrentDateString()})`,
        );
        await this.ensureDir(rootOut);

        const dirA = path.join(rootOut, "CUOTAS Y BASES INGRESADAS");
        const dirB = path.join(rootOut, "RECIBOS AL COBRO");

        await this.ensureDir(dirA);
        await this.ensureDir(dirB);

        const logPath = path.join(rootOut, "log.tsv"); // tabla tipo Excel
        const logRows = []; // array de filas (objetos)

        // Columnas fijas de la tabla (ordenadas)
        const LOG_COLS = [
          "fila_excel",
          "administrador",
          "naf",
          "estado_a",
          "intento_a",
          "pdf_a",
          "dil_a",
          "estado_b",
          "b_registros",
          "pdfs_b",
          "error_a",
          "error_b",
        ];

        const esc = (v) =>
          String(v ?? "")
            .replace(/\r?\n/g, " ")
            .trim();

        const flushLog = () => {
          const header = LOG_COLS.join("\t");
          const lines = logRows.map((row) =>
            LOG_COLS.map((c) => esc(row[c])).join("\t"),
          );
          fs.writeFileSync(logPath, [header, ...lines].join("\n"), "utf8");
        };

        console.log(
          "[BASES/RECIBOS][INPUT] Leyendo Excel:",
          path.normalize(pathExcel),
        );
        const { headerRow, rows } = await this.leerExcelInput(pathExcel);

        console.log(
          "[BASES/RECIBOS][INPUT] Cabecera detectada en fila:",
          headerRow,
        );
        console.log("[BASES/RECIBOS][INPUT] Registros leídos:", rows.length);

        // Validación
        const toProcess = [];
        for (const r of rows) {
          const admin = this._safeFileName(r.administrador || "SIN_ADMIN");
          const logKey = this._safeFileName(
            `FILA_${r._row} | ${admin} | ${r.naf1}-${r.naf2}`,
          );

          const { missing, invalid } = this.validarRegistro(r);

          if (missing.length || invalid.length) {
            logRows.push({
              fila_excel: r._row,
              administrador: admin,
              naf: `${r.naf1}-${r.naf2}`,
              estado_a: "SKIP",
              estado_b: "SKIP",
              error_a: missing.length
                ? `SKIP_FALTA_DATOS: ${missing.join(" | ")}`
                : "",
              error_b: invalid.length
                ? `SKIP_FORMATO_INVALIDO: ${invalid.join(" | ")}`
                : "",
              intento_a: "",
              pdf_a: "",
              dil_a: "",
              b_registros: "",
              pdfs_b: "",
            });
            continue;
          }

          toProcess.push({ ...r, _logKey: logKey });
        }

        if (!toProcess.length) {
          console.warn("[BASES/RECIBOS] No hay registros válidos.");
          flushLog();
          return resolve(false);
        }

        const urlFS = "https://w2.seg-social.es/fs/indexframes.html";

        browser = await puppeteer.launch({
          headless: false,
          defaultViewport: null,
          executablePath: chromeExePath,
          protocolTimeout: 120000,
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

        // Aceptar diálogos JS (por si acaso)
        page.on("dialog", async (dialog) => {
          try {
            await dialog.accept();
          } catch (_) {}
        });

        // Helpers navegación FS
        const openCotizacionRETA = async () => {
          const ok = await this.clickLinkInFrames(
            page,
            {
              hrefIncludes: "menuSLD-RETA.html",
              textIncludes: "Cotización RETA",
            },
            30000,
          );
          if (!ok) throw new Error("No se pudo clicar 'Cotización RETA'.");
          await this.esperar(800);
        };

        const openConsultaBasesCuotas = async () => {
          const ok = await this.clickLinkInFrames(
            page,
            {
              hrefIncludes: "TRANSACCION=RCRS3",
              textIncludes: "Consulta de bases y cuotas ingresadas",
            },
            30000,
          );
          if (!ok)
            throw new Error(
              "No se pudo clicar 'Consulta de bases y cuotas ingresadas'.",
            );
          await this.esperar(900);
        };

        const openConsultaRecibosEmitidos = async () => {
          const ok = await this.clickLinkInFrames(
            page,
            {
              hrefIncludes: "XV24E003",
              textIncludes: "Consulta de recibos emitidos",
            },
            30000,
          );
          if (!ok)
            throw new Error(
              "No se pudo clicar 'Consulta de recibos emitidos régimen de autónomos'.",
            );
          await this.esperar(900);
        };

        const fillInFrame = async (frame, selector, value, timeout = 30000) => {
          await frame.waitForSelector(selector, { timeout });
          const el = await frame.$(selector);
          if (!el) throw new Error(`No se encontró el input ${selector}`);
          await el.click({ clickCount: 3 });
          await page.keyboard.down("Control");
          await page.keyboard.press("A");
          await page.keyboard.up("Control");
          await page.keyboard.press("Backspace");
          await page.keyboard.type(String(value ?? ""), { delay: 15 });
        };

        const readDIL = async () => {
          for (const fr of page.frames()) {
            try {
              const txt = await fr.evaluate(() => {
                const el = document.querySelector("#DIL");
                return el ? (el.textContent || "").trim() : "";
              });
              if (txt) return txt;
            } catch (_) {}
          }
          return "";
        };

        // -------------------------
        // LOOP por registro
        // -------------------------
        let okA = 0,
          errA = 0,
          okB = 0,
          errB = 0;

        for (let i = 0; i < toProcess.length; i++) {
          registrosProcesados++;
          const r = toProcess[i];

          const admin = this._safeFileName(r.administrador || "SIN_ADMIN");

          const naf = `${r.naf1}-${r.naf2}`;

          const rowLog = {
            fila_excel: r._row,
            administrador: admin,
            naf,
            estado_a: "",
            intento_a: "",
            pdf_a: "",
            dil_a: "",
            estado_b: "",
            b_registros: "",
            pdfs_b: "",
            error_a: "",
            error_b: "",
          };

          logRows.push(rowLog);

          const pdfA = path.join(
            dirA,
            this._safeFileName(`${admin} - ${ejercicioEconomico}.pdf`),
          );

          console.log(
            `[BASES/RECIBOS] ${i + 1}/${toProcess.length} | NAF: ${r.naf1}-${r.naf2}`,
          );

          // ==========
          // PARTE A
          // ==========
          try {
            await page.goto(urlFS, { waitUntil: "domcontentloaded" });
            console.log(
              "[BASES/RECIBOS] FS abierto. Selecciona certificado si aparece.",
            );

            await openCotizacionRETA();
            await openConsultaBasesCuotas();

            const frameForm = await this.findFrameWithSelector(
              page,
              "#SDFWPROVNAF",
              30000,
            );
            if (!frameForm)
              throw new Error(
                "No se encontró el formulario de bases/cuotas (#SDFWPROVNAF).",
              );

            await fillInFrame(frameForm, "#SDFWPROVNAF", r.naf1);
            await fillInFrame(frameForm, "#SDFWRESTONAF", r.naf2);
            await fillInFrame(frameForm, "#SDFWAOMAPA", ejercicioEconomico);

            // Continuar
            await frameForm
              .waitForSelector("#Sub2207101004_35", { timeout: 25000 })
              .catch(() => {});
            const btnCont = await frameForm.$("#Sub2207101004_35");
            if (!btnCont)
              throw new Error(
                "No se encontró el botón Continuar (bases/cuotas).",
              );

            await btnCont.click({ delay: 40 });
            await this.esperar(900);

            // DIL de error
            const dil = await readDIL();
            if (dil) {
              errA++;
              rowLog.estado_a = "DIL";
              rowLog.dil_a = dil;
              rowLog.error_a = "";
              console.warn("[BASES/RECIBOS][A] DIL:", dil);
            } else {
              // Imprimir -> nueva pestaña
              const openPopupFn = async () => {
                let frBtn = null;
                for (const fr of page.frames()) {
                  try {
                    const b = await fr.$("#Sub2204801005_67");
                    if (b) {
                      frBtn = fr;
                      break;
                    }
                  } catch (_) {}
                }
                if (!frBtn)
                  throw new Error(
                    "No se encontró el botón Imprimir (parte A).",
                  );

                const popupPromise = this.waitForPopup(browser, page, 30000);
                const b = await frBtn.$("#Sub2204801005_67");
                await b.click({ delay: 40 });

                const popup = await popupPromise;
                return popup;
              };

              const rtaA = await this.descargarPdfConReintento({
                openPopupFn,
                outputPath: pdfA,
                label: "PARTE_A_PDF",
              });

              okA++;
              rowLog.estado_a = "OK";
              rowLog.intento_a = rtaA?.intento || 1;
              rowLog.pdf_a = path.basename(pdfA);
            }
          } catch (e) {
            errA++;
            rowLog.estado_a = "ERROR";
            rowLog.error_a = String(e?.message || e);
            console.warn("[BASES/RECIBOS][A] ERROR_A:", rowLog.error_a);
          }

          // ==========
          // PARTE B
          // ==========
          const self = this;

          async function clickAnywhere(page, selector, timeoutMs = 60000) {
            const start = Date.now();

            while (Date.now() - start < timeoutMs) {
              for (const f of page.frames()) {
                try {
                  const el = await f.$(selector);
                  if (!el) continue;

                  // intenta scroll (no siempre hace falta, pero ayuda)
                  try {
                    await f.$eval(selector, (e) =>
                      e.scrollIntoView({ block: "center", inline: "center" }),
                    );
                  } catch (_) {}

                  // click normal + fallback
                  try {
                    await f.click(selector, { delay: 50 });
                  } catch (e) {
                    console.warn(
                      `[clickAnywhere] click normal falló en ${f.name()} | ${f.url()} -> ${e.message}`,
                    );
                    await f.$eval(selector, (e) => e.click());
                  }

                  return;
                } catch (_) {
                  // sigue con otros frames
                }
              }

              await self.esperar(300);
            }

            throw new Error(`No se encontró ${selector} en ningún frame`);
          }
          async function getFrameTabla2(page, timeoutMs = 60000) {
            const fr = await self.findFrameWithSelector(
              page,
              "#TABLA_2",
              timeoutMs,
            );
            return fr;
          }

          async function extraerFilasTabla2(fr) {
            return await fr.evaluate(() => {
              const table = document.querySelector("#TABLA_2");
              if (!table) return [];

              const rows = Array.from(table.querySelectorAll("tbody tr"));
              const out = [];

              for (const tr of rows) {
                const tds = Array.from(tr.querySelectorAll("td"));
                const concepto = (tds[0]?.textContent || "")
                  .trim()
                  .replace(/\s+/g, " ");
                const periodo = (tds[1]?.textContent || "")
                  .trim()
                  .replace(/\s+/g, " ");

                const a = tr.querySelector(
                  "a.enlaceFuncDetalle, a[href*='AC_VER_RECIBO']",
                );
                const href = a ? a.getAttribute("href") || "" : "";

                let paramRecibo = null;
                const m = href.match(/paramRecibo=(\d+)/);
                if (m) paramRecibo = Number(m[1]);

                if (concepto && periodo && href)
                  out.push({ concepto, periodo, href, paramRecibo });
              }

              return out;
            });
          }

          function buildPdfBName(concepto, admin, periodo) {
            const c = self._safeFileName(concepto || "RECIBO");
            const a = self._safeFileName(admin || "SIN_ADMIN");
            const p = self._safeFileName(periodo || "SIN_PERIODO");
            return `${c} - ${a} - ${p}.pdf`;
          }

          let pdfsGeneradosB = 0;

          try {
            await page.goto(urlFS, { waitUntil: "domcontentloaded" });

            await openCotizacionRETA();
            await openConsultaRecibosEmitidos();

            // 1) Si estamos en pantalla de autorizados -> clicar 316077
            console.log("Esperando click 316077...");
            await clickAnywhere(page, "#enlace_316077");
            console.log("click detalle realizado");
            await this.esperar(2000);

            // 2) Ya en formulario: obtener el frame que contiene #seleccion_1 (puede haber cambiado)
            console.log("Esperando formulario...");
            const frForm = await this.findFrameWithSelector(
              page,
              "#seleccion_1",
              60000,
            );

            if (!frForm) {
              throw new Error(
                "No se encontró el frame del formulario (#seleccion_1) tras clicar 316077",
              );
            }

            await frForm.select("#seleccion_1", "0521");
            await frForm.select("#seleccion_3", "07");

            await frForm.waitForSelector("#idTexto1", { timeout: 60000 });
            await frForm.click("#idTexto1", { clickCount: 3 });
            await frForm.type("#idTexto1", r.naf1, { delay: 10 });

            await frForm.waitForSelector("#idTexto2", { timeout: 60000 });
            await frForm.click("#idTexto2", { clickCount: 3 });
            await frForm.type("#idTexto2", r.naf2, { delay: 10 });

            await frForm.waitForSelector("#botConRegIde", { timeout: 60000 });
            await frForm.click("#botConRegIde", { delay: 40 });
            await this.esperar(1000);
            console.log("Formulario completo");

            // ✅ Tick
            const frAviso = await this.findFrameWithSelector(
              page,
              "#cheAviImport",
              60000,
            );
            await frAviso.waitForSelector("#cheAviImport", {
              timeout: 20000,
            });
            const isChecked = await frAviso
              .$eval("#cheAviImport", (el) => el.checked)
              .catch(() => false);
            if (!isChecked) await frAviso.click("#cheAviImport", { delay: 30 });

            // ✅ Continuar
            await frAviso.waitForSelector("#botContAviso", {
              timeout: 20000,
            });
            await frAviso.click("#botContAviso", { delay: 40 });
            await this.esperar(1000);

            // -------------
            // LISTA recibos: leer TABLA_2 + log + loop por cada recibo
            // -------------
            console.log("[BASES/RECIBOS][B] Esperando tabla TABLA_2...");
            const frTabla2 = await getFrameTabla2(page, 60000);
            if (!frTabla2)
              throw new Error("No se encontró #TABLA_2 (listado de recibos).");

            const items = await extraerFilasTabla2(frTabla2);

            console.log(
              `[BASES/RECIBOS][B] Registros encontrados en TABLA_2: ${items.length}`,
            );
            rowLog.b_registros = items.length;

            if (!items.length)
              throw new Error("TABLA_2 no contiene registros de recibos.");

            // Loop: un PDF por fila de TABLA_2
            for (let j = 0; j < items.length; j++) {
              const it = items[j];

              console.log(
                `[BASES/RECIBOS][B] (${j + 1}/${items.length}) Click detalle -> ${it.concepto} | ${it.periodo} | paramRecibo=${it.paramRecibo}`,
              );

              // Click del enlace concreto (si hay paramRecibo, lo usamos)
              if (it.paramRecibo !== null && it.paramRecibo !== undefined) {
                await clickAnywhere(
                  page,
                  `a[href*="AC_VER_RECIBO"][href*="paramRecibo=${it.paramRecibo}"]`,
                  60000,
                );
              } else {
                await clickAnywhere(page, "a.enlaceFuncDetalle", 60000);
              }

              // -------------
              // Generar PDF del detalle (1 por recibo)
              // -------------

              const frDetalle = await this.findFrameWithSelector(
                page,
                "#TABLA_5",
                60000,
              );
              if (!frDetalle)
                throw new Error("No encontré #TABLA_5 tras abrir el detalle.");

              let htmlDetalle = await frDetalle.content();
              if (!htmlDetalle || htmlDetalle.length < 500)
                throw new Error(
                  "El HTML del detalle está vacío o es demasiado corto.",
                );

              const baseHref = "https://w2.seg-social.es";
              htmlDetalle = this._stripHeavyScripts(htmlDetalle);
              htmlDetalle = this._injectBaseTag(htmlDetalle, baseHref);

              // Nombre PDF dinámico: Concepto + Admin + Periodo
              const pdfName = buildPdfBName(it.concepto, admin, it.periodo);
              const pdfPath = path.join(dirB, pdfName);

              const pdfPage = await browser.newPage();
              await pdfPage.setViewport({
                width: 1280,
                height: 720,
                deviceScaleFactor: 1,
              });
              await pdfPage.setContent(htmlDetalle, { waitUntil: "load" });
              await pdfPage.waitForSelector("#TABLA_5", { timeout: 60000 });

              await pdfPage.emulateMediaType("print");
              await pdfPage.pdf({
                path: pdfPath,
                format: "A4",
                printBackground: true,
                preferCSSPageSize: true,
                margin: {
                  top: "10mm",
                  right: "10mm",
                  bottom: "10mm",
                  left: "10mm",
                },
              });

              await pdfPage.close().catch(() => {});

              console.log("[BASES/RECIBOS][B] PDF guardado:", pdfPath);

              pdfsGeneradosB++;
              rowLog.pdfs_b = pdfsGeneradosB;

              // Volver al listado SOLO si quedan más recibos por procesar
              if (j < items.length - 1) {
                console.log(
                  `[BASES/RECIBOS][B] Volviendo al listado para siguiente recibo...`,
                );

                const volverOk = await (async () => {
                  // Intento 1: back rápido
                  try {
                    await this.esperar(500);
                    await page.goBack({
                      waitUntil: "domcontentloaded",
                      timeout: 15000,
                    });
                    return true;
                  } catch (_) {}

                  // Intento 2: back sin esperar carga (a veces PROSA nunca dispara load)
                  try {
                    await this.esperar(500);
                    await page
                      .goBack({ waitUntil: "networkidle2", timeout: 15000 })
                      .catch(() =>
                        page.goBack({
                          waitUntil: "domcontentloaded",
                          timeout: 15000,
                        }),
                      );
                    return true;
                  } catch (_) {}

                  // Intento 3: fallback seguro -> re-ejecutar navegación hasta llegar al listado otra vez
                  try {
                    console.warn(
                      "[BASES/RECIBOS][B] goBack falló. Re-navegando al listado...",
                    );
                    await page.goto(urlFS, { waitUntil: "domcontentloaded" });
                    await openCotizacionRETA();
                    await openConsultaRecibosEmitidos();

                    // autorizado 316077
                    await clickAnywhere(page, "#enlace_316077");
                    await this.esperar(1200);

                    // NOTA: aquí ya estamos en formulario, pero NO hace falta rellenarlo otra vez si PROSA mantiene contexto.
                    // Sin embargo, si en tu caso te obliga, descomenta rellenado y consulta igual que al inicio:
                    //
                    // const frFormAgain = await this.findFrameWithSelector(page, "#seleccion_1", 60000);
                    // await frFormAgain.select("#seleccion_1", "0521");
                    // await frFormAgain.select("#seleccion_3", "07");
                    // await frFormAgain.click("#idTexto1", { clickCount: 3 }); await frFormAgain.type("#idTexto1", r.naf1, { delay: 10 });
                    // await frFormAgain.click("#idTexto2", { clickCount: 3 }); await frFormAgain.type("#idTexto2", r.naf2, { delay: 10 });
                    // await frFormAgain.click("#botConRegIde", { delay: 40 });
                    // await this.esperar(900);
                    // const frAvisoAgain = await this.findFrameWithSelector(page, "#cheAviImport", 60000);
                    // const isChecked2 = await frAvisoAgain.$eval("#cheAviImport", el => el.checked).catch(()=>false);
                    // if (!isChecked2) await frAvisoAgain.click("#cheAviImport", { delay: 30 });
                    // await frAvisoAgain.click("#botContAviso", { delay: 40 });
                    // await this.esperar(900);

                    return true;
                  } catch (e) {
                    console.warn(
                      "[BASES/RECIBOS][B] Fallback re-navegación falló:",
                      e?.message || e,
                    );
                    return false;
                  }
                })();

                if (!volverOk) {
                  // En vez de petar el proceso, registramos y seguimos (ya tenemos PDFs generados hasta aquí)
                  console.warn(
                    "[BASES/RECIBOS][B] No pude volver al listado, pero continúo (fin loop).",
                  );
                  break;
                }

                // Asegurar que el listado está otra vez (mejor: esperar selector)
                const frTabla2Again = await getFrameTabla2(page, 20000).catch(
                  () => null,
                );
                if (!frTabla2Again) {
                  console.warn(
                    "[BASES/RECIBOS][B] No aparece #TABLA_2 tras volver. Continúo (fin loop).",
                  );
                  break;
                }
              }
            }
          } catch (e) {
            const msg = `ERROR_B: ${e?.message || e}`;
            rowLog.error_b = String(e?.message || e);
            console.warn("[BASES/RECIBOS][B]", msg);
          }
          if (pdfsGeneradosB > 0) {
            okB++;
            rowLog.estado_b = "OK";
          } else {
            errB++;
            rowLog.estado_b = "ERROR";
            rowLog.error_b = rowLog.error_b || "No se generó ningún PDF";
          }
          if ((i + 1) % 5 === 0) flushLog();
        }

        flushLog();

        console.log("[BASES/RECIBOS] Terminado.");
        console.log(
          `[BASES/RECIBOS] OK_A=${okA} ERR_A=${errA} | OK_B=${okB} ERR_B=${errB}`,
        );
        console.log(`[BASES/RECIBOS] Procesados: ${registrosProcesados}`);

        // Si tu repo lo tiene:
        // registrarEjecucion({ nombreProceso, registrosProcesados });

        try {
          if (browser) await browser.close();
        } catch (_) {}

        return resolve(true);
      } catch (err) {
        console.error("[BASES/RECIBOS] Error general:", err?.message || err);
        try {
          if (browser) await browser.close();
        } catch (_) {}
        return resolve(false);
      }
    });
  }

  // Alias cómodo
  async ["basesYRecibosAlCobroAutonomos"](argumentos) {
    return this.basesYRecibosAutonomos(argumentos);
  }
}

module.exports = ProcesosBasesRecibosAutonomos;
