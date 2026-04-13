const path = require("path");
const fs = require("fs");
const XlsxPopulate = require("xlsx-populate");
const puppeteer = require("puppeteer");
const { registrarEjecucion } = require("../metricas");

/**
 * Bases y recibos al cobro autónomos
 * - Parte A: Bases y cuotas ingresadas (PDF por NAF y año)
 * - Parte B: Recibos al cobro (PDF por cada fila del listado TABLA_2)
 *
 * Objetivo: simple, mantenible, sin inventos.
 */
class ProcesosBasesRecibosAutonomos {
  constructor(pathToDbFolder, nombreProyecto, proyectoDB) {
    this.pathToDbFolder = pathToDbFolder;
    this.nombreProyecto = nombreProyecto;
    this.proyectoDB = proyectoDB;

    // Prefijo único para logs en consola
    this.TAG = "[BASES/RECIBOS]";
  }

  // ==========================================================
  // Utils básicas
  // ==========================================================
  esperar(ms) {
    return new Promise((r) => setTimeout(r, ms));
  }

  getCurrentDateString() {
    const d = new Date();
    const yyyy = d.getFullYear();
    const mm = String(d.getMonth() + 1).padStart(2, "0");
    const dd = String(d.getDate()).padStart(2, "0");
    return `${yyyy}-${mm}-${dd}`;
  }

  ensureDir(dir) {
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
  }

  log(msg, ...rest) {
    console.log(`${this.TAG} ${msg}`, ...rest);
  }

  logWarn(msg, ...rest) {
    console.warn(`${this.TAG} ${msg}`, ...rest);
  }

  logErr(msg, ...rest) {
    console.error(`${this.TAG} ${msg}`, ...rest);
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

  _excelColName(n) {
    let s = "";
    while (n > 0) {
      const m = (n - 1) % 26;
      s = String.fromCharCode(65 + m) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }

  _injectBaseTag(html, baseHref) {
    if (!html) return html;
    if (/<base\s/i.test(html)) return html;

    if (/<head[^>]*>/i.test(html)) {
      return html.replace(
        /<head[^>]*>/i,
        (m) => `${m}\n<base href="${baseHref}">`,
      );
    }
    return `<base href="${baseHref}">\n${html}`;
  }

  _stripHeavyScripts(html) {
    if (!html) return html;
    return html.replace(
      /<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi,
      "",
    );
  }

  // ==========================================================
  // Logger CSV
  // ==========================================================
  createCsvLogger(rootOut) {
    const logPath = path.join(rootOut, "log.csv");
    const logRows = [];

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

    // Escapado CSV: comillas dobles y campo entre comillas
    const esc = (v) => {
      const s = String(v ?? "")
        .replace(/\r?\n/g, " ")
        .trim();
      return `"${s.replace(/"/g, '""')}"`;
    };

    const flush = () => {
      const header = LOG_COLS.map(esc).join(",");
      const lines = logRows.map((row) =>
        LOG_COLS.map((c) => esc(row[c])).join(","),
      );
      fs.writeFileSync(logPath, [header, ...lines].join("\n"), "utf8");
    };

    return { logPath, logRows, flush };
  }

  // ==========================================================
  // Excel
  // ==========================================================
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
        "No se encontró la cabecera. Necesito 'EXPTE.' y 'ADMINISTRADOR'.",
      );
    }

    // Columnas fijas: G/H -> NAF1/NAF2
    const colNAF1 = 7;
    const colNAF2 = 8;

    // Debug headers (por si necesitas)
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

    const rows = [];
    for (let r = headerRow + 1; r <= numRows; r++) {
      const naf1Cell = getCell(r, colNAF1);
      const naf2Cell = getCell(r, colNAF2);
      const adminCell = headerMap.ADMIN ? getCell(r, headerMap.ADMIN) : "";

      const empty =
        (naf1Cell == null || String(naf1Cell).trim() === "") &&
        (naf2Cell == null || String(naf2Cell).trim() === "") &&
        (adminCell == null || String(adminCell).trim() === "");

      if (empty) break;

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
      administrador: String(r.ADMIN ?? "").trim(),
      naf1,
      naf2,
      _row: r._row,
      // el resto lo dejamos por si lo usas en futuro
      expte: String(r.EXPTE ?? "").trim(),
      empresa: String(r.EMPRESA ?? "").trim(),
      nafRaw: String(r.NAF ?? "").trim(),
      clave: String(r.CLAVE ?? "").trim(),
      fechaAltaBaja: String(r.FALTA_BAJA ?? "").trim(),
      base: String(r.BASE ?? "").trim(),
      total: String(r.TOTAL ?? "").trim(),
      prevAno: String(r.PREV_ANO ?? "").trim(),
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

  // ==========================================================
  // Frames / clicks (unificado)
  // ==========================================================
  async findFrameWithSelector(page, selector, timeoutMs = 25000, pollMs = 350) {
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
      await this.esperar(350);
    }

    return false;
  }

  /**
   * Click simple en cualquier frame. Es tu clickAnywhere, solo renombrado.
   */
  async clickInFrames(page, selector, timeoutMs = 60000) {
    const start = Date.now();

    while (Date.now() - start < timeoutMs) {
      for (const fr of page.frames()) {
        try {
          const el = await fr.$(selector);
          if (!el) continue;

          try {
            await fr.$eval(selector, (e) =>
              e.scrollIntoView({ block: "center", inline: "center" }),
            );
          } catch (_) {}

          try {
            await fr.click(selector, { delay: 50 });
          } catch (_) {
            await fr.$eval(selector, (e) => e.click());
          }

          return true;
        } catch (_) {}
      }

      await this.esperar(300);
    }

    throw new Error(`No se encontró ${selector} en ningún frame`);
  }

  async typeInFrame(page, frame, selector, value, timeout = 30000) {
    await frame.waitForSelector(selector, { timeout });
    const el = await frame.$(selector);
    if (!el) throw new Error(`No se encontró el input ${selector}`);

    await el.click({ clickCount: 3 });

    await page.keyboard.down("Control");
    await page.keyboard.press("A");
    await page.keyboard.up("Control");
    await page.keyboard.press("Backspace");

    await page.keyboard.type(String(value ?? ""), { delay: 15 });
  }

  async readDIL(page) {
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
  }

  // ==========================================================
  // PDF A (popup + CDP Fetch) - igual que el que ya funciona
  // ==========================================================
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
    return target.page().catch(() => null);
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

            const okStatus = status >= 200 && status < 300;
            if (!okStatus) {
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

            if (!this._isPdfBuffer(buf)) return;

            done = true;
            fs.writeFileSync(outputPath, buf);
            resolve(true);
          } catch (e) {
            reject(e);
          }
        };

        client.on("Fetch.requestPaused", onPaused);
      });

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

        const retryable = this._isPdfDownloadRetryableError(e);
        if (!retryable || intento === 2) break;

        this.logWarn(
          `[${label}] Fallo al descargar (intento ${intento}). Reintentando 1 vez...`,
          e?.message || e,
        );
        await this.esperar(1200);
      }
    }

    throw lastErr;
  }

  // ==========================================================
  // Navegación FS
  // ==========================================================
  async openCotizacionRETA(page) {
    const ok = await this.clickLinkInFrames(
      page,
      { hrefIncludes: "menuSLD-RETA.html", textIncludes: "Cotización RETA" },
      30000,
    );
    if (!ok) throw new Error("No se pudo clicar 'Cotización RETA'.");
    await this.esperar(800);
  }

  async openConsultaBasesCuotas(page) {
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
  }

  async openConsultaRecibosEmitidos(page) {
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
  }

  // ==========================================================
  // Parte B helpers
  // ==========================================================
  async getFrameTabla2(page, timeoutMs = 60000) {
    return this.findFrameWithSelector(page, "#TABLA_2", timeoutMs);
  }

  async extraerFilasTabla2(fr) {
    return fr.evaluate(() => {
      const table = document.querySelector("#TABLA_2");
      if (!table) return [];

      const rows = Array.from(table.querySelectorAll("tbody tr"));
      const out = [];

      for (const tr of rows) {
        const tds = Array.from(tr.querySelectorAll("td"));
        const concepto = (tds[0]?.textContent || "")
          .trim()
          .replace(/\s+/g, " ");
        const periodo = (tds[1]?.textContent || "").trim().replace(/\s+/g, " ");

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

  buildPdfBName(concepto, admin, periodo) {
    const c = this._safeFileName(concepto || "RECIBO");
    const a = this._safeFileName(admin || "SIN_ADMIN");
    const p = this._safeFileName(periodo || "SIN_PERIODO");
    return `${c} - ${a} - ${p}.pdf`;
  }

  async volverAlListadoB(page, urlFS) {
    try {
      await this.esperar(500);
      await page.goBack({ waitUntil: "domcontentloaded", timeout: 15000 });
      return true;
    } catch (_) {}

    try {
      await this.esperar(500);
      await page.goBack({ waitUntil: "networkidle2", timeout: 15000 });
      return true;
    } catch (_) {}

    try {
      this.logWarn("[B] goBack falló. Re-navegando al listado...");
      await page.goto(urlFS, { waitUntil: "domcontentloaded" });
      await this.openCotizacionRETA(page);
      await this.openConsultaRecibosEmitidos(page);
      await this.clickInFrames(page, "#enlace_316077");
      await this.esperar(1200);
      return true;
    } catch (e) {
      this.logWarn("[B] Fallback re-navegación falló:", e?.message || e);
      return false;
    }
  }

  async renderDetalleFrameToPdf(browser, frDetalle, pdfPath) {
    let htmlDetalle = await frDetalle.content();
    if (!htmlDetalle || htmlDetalle.length < 500) {
      throw new Error("El HTML del detalle está vacío o es demasiado corto.");
    }

    htmlDetalle = this._stripHeavyScripts(htmlDetalle);
    htmlDetalle = this._injectBaseTag(htmlDetalle, "https://w2.seg-social.es");

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
      margin: { top: "10mm", right: "10mm", bottom: "10mm", left: "10mm" },
    });

    await pdfPage.close().catch(() => {});
  }

  // ==========================================================
  // Inputs
  // ==========================================================
  validarInputs(argumentos) {
    const chromeExePath = argumentos?.formularioControl?.[0];
    const pathExcel = argumentos?.formularioControl?.[1];
    const ejercicioEconomicoRaw = argumentos?.formularioControl?.[2];
    const pathSalidaBase = argumentos?.formularioControl?.[3];

    if (!chromeExePath || !fs.existsSync(chromeExePath)) {
      throw new Error("Ruta a chrome.exe no válida.");
    }
    if (!pathExcel || !fs.existsSync(pathExcel)) {
      throw new Error("Ruta a Excel no válida.");
    }
    if (!pathSalidaBase || !String(pathSalidaBase).trim()) {
      throw new Error("Ruta de salida no válida.");
    }

    const ejercicioEconomico = String(ejercicioEconomicoRaw ?? "").trim();
    if (!/^\d{4}$/.test(ejercicioEconomico)) {
      throw new Error(
        "Ejercicio económico inválido. Debe ser AAAA (ej: 2025).",
      );
    }

    return { chromeExePath, pathExcel, ejercicioEconomico, pathSalidaBase };
  }

  // ==========================================================
  // Parte A (extraída)
  // ==========================================================
  async runParteA({
    browser,
    page,
    urlFS,
    ejercicioEconomico,
    r,
    pdfAPath,
    rowLog,
  }) {
    try {
      await page.goto(urlFS, { waitUntil: "domcontentloaded" });
      this.log("[A] FS abierto. Selecciona certificado si aparece.");

      await this.openCotizacionRETA(page);
      await this.openConsultaBasesCuotas(page);

      const frForm = await this.findFrameWithSelector(
        page,
        "#SDFWPROVNAF",
        30000,
      );
      if (!frForm)
        throw new Error("No se encontró el formulario (#SDFWPROVNAF).");

      await this.typeInFrame(page, frForm, "#SDFWPROVNAF", r.naf1);
      await this.typeInFrame(page, frForm, "#SDFWRESTONAF", r.naf2);
      await this.typeInFrame(page, frForm, "#SDFWAOMAPA", ejercicioEconomico);

      const btnSelector = "#Sub2207101004_35";
      await frForm
        .waitForSelector(btnSelector, { timeout: 25000 })
        .catch(() => {});
      const btnCont = await frForm.$(btnSelector);
      if (!btnCont) throw new Error("No se encontró el botón Continuar.");

      await btnCont.click({ delay: 40 });
      await this.esperar(900);

      const dil = await this.readDIL(page);
      if (dil) {
        rowLog.estado_a = "DIL";
        rowLog.dil_a = dil;
        this.logWarn("[A] DIL:", dil);
        return { ok: false, dil: true, intento: "" };
      }

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
        if (!frBtn) throw new Error("No se encontró el botón Imprimir (A).");

        const popupPromise = this.waitForPopup(browser, page, 30000);
        const b = await frBtn.$("#Sub2204801005_67");
        await b.click({ delay: 40 });

        return popupPromise;
      };

      const rtaA = await this.descargarPdfConReintento({
        openPopupFn,
        outputPath: pdfAPath,
        label: "PDF_A",
      });

      rowLog.estado_a = "OK";
      rowLog.intento_a = rtaA?.intento || 1;
      rowLog.pdf_a = path.basename(pdfAPath);

      return { ok: true, dil: false, intento: rtaA?.intento || 1 };
    } catch (e) {
      rowLog.estado_a = "ERROR";
      rowLog.error_a = String(e?.message || e);
      this.logWarn("[A] ERROR:", rowLog.error_a);
      return { ok: false, dil: false, intento: "" };
    }
  }

  // ==========================================================
  // Parte B (extraída)
  // ==========================================================
  async runParteB({ browser, page, urlFS, r, admin, dirB, rowLog }) {
    let pdfsGeneradosB = 0;

    try {
      await page.goto(urlFS, { waitUntil: "domcontentloaded" });

      await this.openCotizacionRETA(page);
      await this.openConsultaRecibosEmitidos(page);

      this.log("[B] Seleccionando autorizado 316077...");
      await this.clickInFrames(page, "#enlace_316077");
      await this.esperar(2000);

      this.log("[B] Esperando formulario...");
      const frForm = await this.findFrameWithSelector(
        page,
        "#seleccion_1",
        60000,
      );
      if (!frForm)
        throw new Error("No se encontró el formulario (#seleccion_1).");

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

      // aviso + tick
      const frAviso = await this.findFrameWithSelector(
        page,
        "#cheAviImport",
        60000,
      );
      if (!frAviso) throw new Error("No se encontró el aviso (#cheAviImport).");

      const isChecked = await frAviso
        .$eval("#cheAviImport", (el) => el.checked)
        .catch(() => false);
      if (!isChecked) await frAviso.click("#cheAviImport", { delay: 30 });

      await frAviso.waitForSelector("#botContAviso", { timeout: 20000 });
      await frAviso.click("#botContAviso", { delay: 40 });
      await this.esperar(1000);

      // listado
      this.log("[B] Esperando listado (TABLA_2)...");
      const frTabla2 = await this.getFrameTabla2(page, 60000);
      if (!frTabla2) throw new Error("No se encontró TABLA_2.");

      const items = await this.extraerFilasTabla2(frTabla2);
      rowLog.b_registros = items.length;

      this.log(`[B] Recibos encontrados: ${items.length}`);
      if (!items.length) throw new Error("TABLA_2 no contiene recibos.");

      for (let j = 0; j < items.length; j++) {
        const it = items[j];

        this.log(
          `[B] (${j + 1}/${items.length}) Abriendo detalle: ${it.concepto} | ${it.periodo}`,
        );

        if (it.paramRecibo !== null && it.paramRecibo !== undefined) {
          await this.clickInFrames(
            page,
            `a[href*="AC_VER_RECIBO"][href*="paramRecibo=${it.paramRecibo}"]`,
            60000,
          );
        } else {
          await this.clickInFrames(page, "a.enlaceFuncDetalle", 60000);
        }

        const frDetalle = await this.findFrameWithSelector(
          page,
          "#TABLA_5",
          60000,
        );
        if (!frDetalle)
          throw new Error("No se encontró TABLA_5 en el detalle.");

        const pdfName = this.buildPdfBName(it.concepto, admin, it.periodo);
        const pdfPath = path.join(dirB, pdfName);

        await this.renderDetalleFrameToPdf(browser, frDetalle, pdfPath);

        pdfsGeneradosB++;
        rowLog.pdfs_b = pdfsGeneradosB;

        // volver si quedan más
        if (j < items.length - 1) {
          const volverOk = await this.volverAlListadoB(page, urlFS);
          if (!volverOk) break;

          const frAgain = await this.getFrameTabla2(page, 20000).catch(
            () => null,
          );
          if (!frAgain) break;
        }
      }

      rowLog.estado_b = pdfsGeneradosB > 0 ? "OK" : "ERROR";
      if (pdfsGeneradosB === 0) rowLog.error_b = "No se generó ningún PDF";

      return { pdfs: pdfsGeneradosB, ok: pdfsGeneradosB > 0 };
    } catch (e) {
      rowLog.estado_b = "ERROR";
      rowLog.error_b = String(e?.message || e);
      this.logWarn("[B] ERROR:", rowLog.error_b);
      return { pdfs: pdfsGeneradosB, ok: false };
    }
  }

  // ==========================================================
  // PROCESO PRINCIPAL
  // ==========================================================
  async basesYRecibosAutonomos(argumentos) {
    this.log("Inicio del proceso");
    const nombreProceso = "Bases y Recibos Autónomos";

    let browser = null;
    let registrosProcesados = 0;

    try {
      const { chromeExePath, pathExcel, ejercicioEconomico, pathSalidaBase } =
        this.validarInputs(argumentos);

      const rootOut = path.join(
        path.normalize(pathSalidaBase),
        `BasesYRecibosAutonomos (${this.getCurrentDateString()})`,
      );
      this.ensureDir(rootOut);

      const dirA = path.join(rootOut, "CUOTAS Y BASES INGRESADAS");
      const dirB = path.join(rootOut, "RECIBOS AL COBRO");
      this.ensureDir(dirA);
      this.ensureDir(dirB);

      const logger = this.createCsvLogger(rootOut);

      this.log(`Leyendo Excel: ${path.normalize(pathExcel)}`);
      const { headerRow, rows } = await this.leerExcelInput(pathExcel);
      this.log(`Cabecera detectada en fila: ${headerRow}`);
      this.log(`Registros leídos: ${rows.length}`);

      // Filtrar/validar
      const toProcess = [];
      for (const r of rows) {
        const admin = this._safeFileName(r.administrador || "SIN_ADMIN");
        const { missing, invalid } = this.validarRegistro(r);

        if (missing.length || invalid.length) {
          logger.logRows.push({
            fila_excel: r._row,
            administrador: admin,
            naf: `${r.naf1}-${r.naf2}`,
            estado_a: "SKIP",
            estado_b: "SKIP",
            error_a: missing.length ? `SKIP: ${missing.join(" | ")}` : "",
            error_b: invalid.length ? `SKIP: ${invalid.join(" | ")}` : "",
            intento_a: "",
            pdf_a: "",
            dil_a: "",
            b_registros: "",
            pdfs_b: "",
          });
          continue;
        }

        toProcess.push(r);
      }

      if (!toProcess.length) {
        this.logWarn("No hay registros válidos. Fin.");
        logger.flush();
        return false;
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

      const pages = await browser.pages();
      const page = pages.length ? pages[0] : await browser.newPage();

      page.on("dialog", async (dialog) => {
        try {
          await dialog.accept();
        } catch (_) {}
      });

      let okA = 0,
        errA = 0,
        okB = 0,
        errB = 0;

      for (let i = 0; i < toProcess.length; i++) {
        registrosProcesados++;
        const r = toProcess[i];

        const admin = this._safeFileName(r.administrador || "SIN_ADMIN");
        const naf = `${r.naf1}-${r.naf2}`;

        this.log(`Registro ${i + 1}/${toProcess.length} | NAF ${naf}`);

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
        logger.logRows.push(rowLog);

        // ---- Parte A
        const pdfAPath = path.join(
          dirA,
          this._safeFileName(`${admin} - ${ejercicioEconomico}.pdf`),
        );
        const resA = await this.runParteA({
          browser,
          page,
          urlFS,
          ejercicioEconomico,
          r,
          pdfAPath,
          rowLog,
        });
        if (resA.ok) okA++;
        else errA++;

        // ---- Parte B
        const resB = await this.runParteB({
          browser,
          page,
          urlFS,
          r,
          admin,
          dirB,
          rowLog,
        });
        if (resB.ok) okB++;
        else errB++;

        if ((i + 1) % 5 === 0) logger.flush();
      }

      logger.flush();

      this.log("Proceso terminado");
      this.log(`Resumen: A OK=${okA} ERR=${errA} | B OK=${okB} ERR=${errB}`);
      this.log(`Procesados: ${registrosProcesados}`);
      this.log(`Log: ${logger.logPath}`);

      try {
        if (browser) await browser.close();
      } catch (_) {}

      await registrarEjecucion({ nombreProceso, registrosProcesados });
      return true;
    } catch (err) {
      this.logErr("Error general:", err?.message || err);
      try {
        if (browser) await browser.close();
      } catch (_) {}
      return false;
    }
  }

  // Alias cómodo
  async ["basesYRecibosAlCobroAutonomos"](argumentos) {
    return this.basesYRecibosAutonomos(argumentos);
  }
}

module.exports = ProcesosBasesRecibosAutonomos;
