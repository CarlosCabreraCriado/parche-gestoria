const path = require("path");
const fs = require("fs");
const XlsxPopulate = require("xlsx-populate");
const puppeteer = require("puppeteer");

/**
 * Procesos Bases y Recibos Autónomos
 *
 * Inputs (argumentos.formularioControl):
 *  [0] chromeExePath  (ruta chrome.exe)
 *  [1] pathExcelInput (fichero input)
 *  [2] pathSalidaBase (carpeta destino)
 *
 * Notas:
 * - Mes y año del nombre: MES y AÑO actual (sistema).
 * - Recibos emitidos: descargar el primero que salga en la lista.
 * - Robusto: timeouts razonables, 1 reintento en descarga PDF A (popup),
 *           PDF B se genera con printToPDF (sin chrome://print).
 * - Si falla un registro se loggea y se continúa.
 */
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

  _dniNorm(dni) {
    return String(dni ?? "")
      .toUpperCase()
      .replace(/\s+/g, "")
      .trim();
  }

  _dniFolderKey(dni) {
    const s = this._dniNorm(dni);
    if (!s) return "SIN_DNI";
    // Quitar última letra si existe (DNI/NIE)
    if (/[A-Z]$/.test(s)) return s.slice(0, -1);
    return s;
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

    // 1) Encontrar fila cabecera buscando “EXPTE.”, “DNI” y “ADMINISTRADOR”
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
      const colDni = findInRow(rowHeaders, "DNI");
      const colAdmin = findInRow(rowHeaders, "ADMINISTRADOR");

      if (colExpte && colDni && colAdmin) {
        headerRow = r;
        headerMap = {
          EXPTE: colExpte,
          EMPRESA: findInRow(rowHeaders, "EMPRESA"),
          NAF: findInRow(rowHeaders, "NAF"),
          DNI: colDni,
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

    if (!headerRow || !headerMap?.DNI) {
      throw new Error(
        "No se encontró la fila de cabecera. Necesito al menos 'EXPTE.', 'DNI' y 'ADMINISTRADOR'.",
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

    // 4) Parsear registros hasta primer DNI vacío
    const rows = [];
    for (let r = headerRow + 1; r <= numRows; r++) {
      const dni = getCell(r, headerMap.DNI);
      if (dni === null || dni === undefined || String(dni).trim() === "") break;

      const rec = {
        EXPTE: headerMap.EXPTE ? getCell(r, headerMap.EXPTE) : "",
        EMPRESA: headerMap.EMPRESA ? getCell(r, headerMap.EMPRESA) : "",
        NAF: headerMap.NAF ? getCell(r, headerMap.NAF) : "",
        DNI: dni,
        CLAVE: headerMap.CLAVE ? getCell(r, headerMap.CLAVE) : "",
        FALTA_BAJA: headerMap.FALTA_BAJA
          ? getCell(r, headerMap.FALTA_BAJA)
          : "",
        ADMIN: headerMap.ADMIN ? getCell(r, headerMap.ADMIN) : "",
        BASE: headerMap.BASE ? getCell(r, headerMap.BASE) : "",
        TOTAL: headerMap.TOTAL ? getCell(r, headerMap.TOTAL) : "",
        PREV_ANO: headerMap.PREV_ANO ? getCell(r, headerMap.PREV_ANO) : "",

        // NAF1/NAF2 “forzados” desde G/H
        NAF1: getCell(r, colNAF1),
        NAF2: getCell(r, colNAF2),

        _row: r,
      };

      rows.push(this.normalizarRegistro(rec));
    }

    return { headers, headerRow, rows };
  }

  normalizarRegistro(r) {
    const dni = this._dniNorm(r.DNI);

    const naf1 = this._padLeftDigitsOrEmpty(r.NAF1, 2);
    const naf2 = this._padLeftDigitsOrEmpty(r.NAF2, 10);

    return {
      expte: String(r.EXPTE ?? "").trim(),
      empresa: String(r.EMPRESA ?? "").trim(),
      nafRaw: String(r.NAF ?? "").trim(),
      dni,
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

    req(r.dni, "DNI vacío");
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

  /**
   * Click ultra-robusto dentro de un FRAME usando DOM click (evita "not clickable").
   * - Espera selector
   * - Scroll into view
   * - Intenta click() DOM
   */
  async safeDomClick(frame, selector, { timeout = 60000, label = "" } = {}) {
    await frame.waitForSelector(selector, { timeout });

    const ok = await frame.evaluate((sel) => {
      const el = document.querySelector(sel);
      if (!el) return { ok: false, reason: "not_found" };

      // Scroll
      try {
        el.scrollIntoView({ block: "center", inline: "center" });
      } catch (_) {}

      // Por si es <a> con target blank
      try {
        el.target = "_self";
      } catch (_) {}

      // Click DOM
      try {
        el.click();
        return { ok: true };
      } catch (e) {
        return { ok: false, reason: String(e) };
      }
    }, selector);

    if (!ok?.ok) {
      throw new Error(
        `${label ? label + ": " : ""}safeDomClick falló en selector "${selector}" (reason=${ok?.reason || "unknown"})`,
      );
    }
  }

  /**
   * Para inputs: set value “a lo bruto” (evita teclas/overlay raros)
   */
  async safeSetInput(
    frame,
    selector,
    value,
    { timeout = 60000, label = "" } = {},
  ) {
    await frame.waitForSelector(selector, { timeout });

    const ok = await frame.evaluate(
      (sel, val) => {
        const el = document.querySelector(sel);
        if (!el) return { ok: false, reason: "not_found" };

        try {
          el.scrollIntoView({ block: "center", inline: "center" });
        } catch (_) {}

        // set value + events
        el.focus();
        el.value = String(val ?? "");
        el.dispatchEvent(new Event("input", { bubbles: true }));
        el.dispatchEvent(new Event("change", { bubbles: true }));
        return { ok: true };
      },
      selector,
      value,
    );

    if (!ok?.ok) {
      throw new Error(
        `${label ? label + ": " : ""}safeSetInput falló en selector "${selector}" (reason=${ok?.reason || "unknown"})`,
      );
    }
  }

  /**
   * Click al primer recibo del listado (DOM click, sin ElementHandle)
   */
  async clickPrimerReciboEnListado(frListado) {
    const ok = await frListado.evaluate(() => {
      const pick =
        document.querySelector("#enlace_0") ||
        document.querySelector('a[href*="AC_VER_RECIBO"]') ||
        document.querySelector("a.enlaceFuncDetalle") ||
        document.querySelector("a.pr_enlaceLocal");

      if (!pick) return false;

      try {
        pick.target = "_self";
      } catch (_) {}
      try {
        pick.scrollIntoView({ block: "center", inline: "center" });
      } catch (_) {}
      pick.click();
      return true;
    });

    if (!ok)
      throw new Error(
        "No se encontró enlace al detalle del recibo en el listado.",
      );
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
   * Útil cuando la web abre un popup/visor que descarga "de verdad" un PDF.
   */
  async descargarPdfRawViaFetchCDP(
    popupPage,
    outputPath,
    timeoutMs = 90000,
    { label = "PDF", triggerFn = null, doReload = false } = {},
  ) {
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
            } catch {}
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
                ctype.includes("binary"));

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
              // No era PDF real
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

      // 1) disparador opcional
      if (typeof triggerFn === "function") {
        await triggerFn().catch((e) => {
          throw new Error(`triggerFn fallo: ${e?.message || e}`);
        });
      }

      // 2) reload opcional (normalmente NO hace falta)
      if (doReload) {
        await popupPage
          .reload({ waitUntil: "domcontentloaded" })
          .catch(() => {});
        await popupPage
          .waitForSelector("body", { timeout: 20000 })
          .catch(() => {});
        await this.esperar(800);
      }

      await Promise.race([pdfPromise, timer]);
      return true;
    } finally {
      try {
        if (onPaused) client.removeListener("Fetch.requestPaused", onPaused);
      } catch {}
      try {
        await client.send("Fetch.disable").catch(() => {});
      } catch {}
      try {
        await client.detach();
      } catch {}
    }
  }

  async descargarPdfConReintento({
    openPopupFn,
    outputPath,
    label,
    triggerFn,
  }) {
    let lastErr = null;

    for (let intento = 1; intento <= 2; intento++) {
      let popup = null;
      try {
        popup = await openPopupFn();
        if (!popup) throw new Error("No se abrió el popup/pestaña del PDF.");

        await this.descargarPdfRawViaFetchCDP(popup, outputPath, 90000, {
          label,
          triggerFn: triggerFn ? () => triggerFn(popup) : null,
          doReload: false,
        });

        try {
          await popup.close();
        } catch (_) {}
        return true;
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

  /**
   * ✅ Genera PDF de la página actual SIN abrir chrome://print
   * (DevTools Protocol -> Page.printToPDF)
   */
  async guardarPdfConPrintToPDF(page, outputPath, { timeoutMs = 60000 } = {}) {
    await page.bringToFront().catch(() => {});
    await page.waitForSelector("body", { timeout: 20000 }).catch(() => {});
    await this.esperar(250);

    const client = await page.target().createCDPSession();
    try {
      await client.send("Page.enable").catch(() => {});
      // Para que respete estilos de pantalla (el HTML tiene estilos de impresión en media="print")
      await client
        .send("Emulation.setEmulatedMedia", { media: "screen" })
        .catch(() => {});

      const timer = new Promise((_, rej) =>
        setTimeout(
          () => rej(new Error("Timeout en Page.printToPDF")),
          timeoutMs,
        ),
      );

      const pdfPromise = client.send("Page.printToPDF", {
        printBackground: true,
        preferCSSPageSize: true,
        marginTop: 0.4,
        marginBottom: 0.4,
        marginLeft: 0.4,
        marginRight: 0.4,
      });

      const pdf = await Promise.race([pdfPromise, timer]);
      const buf = Buffer.from(pdf.data, "base64");
      fs.writeFileSync(outputPath, buf);
      return true;
    } finally {
      try {
        await client.detach();
      } catch (_) {}
    }
  }

  // -------------------------
  // Selección “primer recibo que salga” (SIN ElementHandle)
  // -------------------------
  async clickPrimerReciboEnListado(frListado) {
    const ok = await frListado.evaluate(() => {
      const pick =
        document.querySelector("#enlace_0") ||
        document.querySelector('a[href*="AC_VER_RECIBO"]') ||
        document.querySelector("a.enlaceFuncDetalle") ||
        document.querySelector("a.pr_enlaceLocal") ||
        document.querySelector("a");

      if (!pick) return false;

      try {
        pick.target = "_self";
      } catch (_) {}
      try {
        pick.scrollIntoView({ block: "center", inline: "center" });
      } catch (_) {}

      // Click DOM (robusto aunque esté “no clickable” para puppeteer)
      pick.click();
      return true;
    });

    if (!ok) {
      throw new Error(
        "No se encontró enlace clicable al detalle del recibo en el listado.",
      );
    }
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
   * Esta función devuelve { pageB, frameB } donde está PROSA.
   */
  async getProsaContext(browser, seedPage, timeoutMs = 60000) {
    const start = Date.now();

    const looksLikeProsaFrame = async (fr) => {
      try {
        return await fr.evaluate(() => {
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
              .includes("detalle adeudo");

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
      if (await looksLikeProsaFrame(pg.mainFrame()))
        return { pageB: pg, frameB: pg.mainFrame() };
      return null;
    };

    while (Date.now() - start < timeoutMs) {
      const inSeed = await findInPageFrames(seedPage);
      if (inSeed) return inSeed;

      const pages = await browser.pages();
      for (const pg of pages) {
        const found = await findInPageFrames(pg);
        if (found) return found;
      }

      await this.esperar(350);
    }

    throw new Error(
      "No se pudo localizar el contexto PROSA (frames o pestaña nueva).",
    );
  }

  async findFirstInFrames(page, selectors, timeoutMs = 30000, pollMs = 300) {
    const start = Date.now();
    while (Date.now() - start < timeoutMs) {
      for (const fr of page.frames()) {
        for (const sel of selectors) {
          try {
            const el = await fr.$(sel);
            if (el) return { frame: fr, selector: sel, el };
          } catch (_) {}
        }
      }
      await this.esperar(pollMs);
    }
    return null;
  }

  /**
   * Guarda el PDF de lo que se ve en la pestaña actual (incluye frames).
   * Mucho más estable que clicar imprimir/popup.
   */
  async guardarPdfPantallaActual(page, outputPath) {
    await page.bringToFront().catch(() => {});
    await page.emulateMediaType("print").catch(() => {});
    await this.esperar(400);

    await page.pdf({
      path: outputPath,
      format: "A4",
      printBackground: true,
      preferCSSPageSize: true,
      margin: { top: "10mm", right: "10mm", bottom: "10mm", left: "10mm" },
    });
  }

  /**
   * Espera a que la pantalla de resultados de Bases/Cuotas esté "cargada".
   * (Pon varios selectores para aguantar cambios del portal)
   */
  async waitResultadosBasesCuotas(page) {
    const hit = await this.findFirstInFrames(
      page,
      [
        // tablas/listados típicos
        "table",
        ".pr_tablaResponsive",
        "[id^='TABLA_']",
        // a veces aparece el botón imprimir pero no lo necesitamos
        "#Sub2204801005_67",
        "button[aria-label*='Imprimir']",
        "#botonImprimir",
      ],
      40000,
    );

    if (!hit)
      throw new Error(
        "No se detectó la pantalla de resultados (no hay tabla/listado).",
      );
  }

  // -------------------------
  // PROCESO PRINCIPAL
  // -------------------------
  async basesYRecibosAutonomos(argumentos) {
    console.log("[BASES/RECIBOS] Iniciando proceso Bases y Recibos Autónomos");

    let registrosProcesados = 0;

    return new Promise(async (resolve) => {
      let browser = null;

      try {
        const chromeExePath = argumentos?.formularioControl?.[0];
        const pathExcel = argumentos?.formularioControl?.[1];
        const pathSalidaBase = argumentos?.formularioControl?.[2];

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

        const logPath = path.join(rootOut, "log.txt");
        const logMap = new Map(); // key: dniKey, value: texto

        const flushLog = () => {
          const lines = Array.from(logMap.entries()).map(
            ([k, v]) => `${k} -> ${v}`,
          );
          fs.writeFileSync(logPath, lines.join("\n"), "utf8");
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
          const dniKey = this._dniFolderKey(r.dni);
          const { missing, invalid } = this.validarRegistro(r);

          if (missing.length) {
            logMap.set(
              dniKey,
              `SKIP_FALTA_DATOS: ${missing.join(" | ")} (fila ${r._row})`,
            );
            continue;
          }
          if (invalid.length) {
            logMap.set(
              dniKey,
              `SKIP_FORMATO_INVALIDO: ${invalid.join(" | ")} (fila ${r._row})`,
            );
            continue;
          }

          toProcess.push(r);
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

          const dniKey = this._dniFolderKey(r.dni);
          const admin = this._safeFileName(r.administrador || "SIN_ADMIN");

          const dirDni = path.join(rootOut, this._safeFileName(dniKey));
          await this.ensureDir(dirDni);

          const pdfA = path.join(
            dirDni,
            this._safeFileName(
              `CUOTAS Y BASES INGRESADAS - ${admin} - ${tagMesAno}.pdf`,
            ),
          );
          const pdfB = path.join(
            dirDni,
            this._safeFileName(
              `RECIBOS AL COBRO - ${admin} - ${tagMesAno}.pdf`,
            ),
          );

          console.log(
            `[BASES/RECIBOS] ${i + 1}/${toProcess.length} | DNI: ${r.dni} | NAF: ${r.naf1}-${r.naf2}`,
          );

          // ==========
          // PARTE A (CORREGIDA - sin popup, sin Fetch)
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
            if (!frameForm) {
              throw new Error(
                "No se encontró el formulario de bases/cuotas (#SDFWPROVNAF).",
              );
            }

            await fillInFrame(frameForm, "#SDFWPROVNAF", r.naf1);
            await fillInFrame(frameForm, "#SDFWRESTONAF", r.naf2);
            await fillInFrame(frameForm, "#SDFWAOMAPA", "2025");

            // Continuar
            await frameForm.waitForSelector("#Sub2207101004_35", {
              timeout: 25000,
            });
            const btnCont = await frameForm.$("#Sub2207101004_35");
            if (!btnCont)
              throw new Error(
                "No se encontró el botón Continuar (bases/cuotas).",
              );

            await btnCont.click({ delay: 40 });
            await this.esperar(1200);

            // DIL de error
            const dil = await readDIL();
            if (dil) {
              errA++;
              logMap.set(dniKey, `PARTE_A_ERROR_DIL: ${dil}`);
              console.warn("[BASES/RECIBOS][A] DIL:", dil);
            } else {
              // ✅ Esperar a pantalla de resultados y guardar PDF directo
              await this.waitResultadosBasesCuotas(page);

              // ✅ PDF directo (incluye frames)
              await this.guardarPdfPantallaActual(page, pdfA);

              okA++;
              logMap.set(
                dniKey,
                `OK_A: PDF A guardado -> ${path.basename(pdfA)}`,
              );
            }
          } catch (e) {
            errA++;
            const msg = `ERROR_A: ${e?.message || e}`;
            logMap.set(dniKey, msg);
            console.warn("[BASES/RECIBOS][A]", msg);
          }

          // ========== PARTE B ==========
          try {
            await page.goto(urlFS, { waitUntil: "domcontentloaded" });

            await openCotizacionRETA();
            await openConsultaRecibosEmitidos();

            // localizar PROSA
            const ctx1 = await this.getProsaContext(browser, page, 60000);
            const { pageB } = ctx1;

            if (pageB !== page) {
              try {
                await pageB.bringToFront();
              } catch (_) {}
            }

            // 1) Pantalla autorizados -> clicar 316077 si aparece
            const ctxAuth = await this.getProsaContext(browser, pageB, 60000);
            const frameMaybeAuth = ctxAuth.frameB;

            const existe316 = await frameMaybeAuth
              .$("#enlace_316077")
              .then(Boolean)
              .catch(() => false);
            if (existe316) {
              await this.safeDomClick(frameMaybeAuth, "#enlace_316077", {
                label: "AUTH 316077",
              });

              const next = await this.findAnyFrameWithAnySelector(
                pageB,
                ["#seleccion_1"],
                60000,
              );
              if (!next)
                throw new Error("Tras clicar 316077 no aparece #seleccion_1.");
            }

            // 2) formulario con #seleccion_1
            const frm = await this.findAnyFrameWithAnySelector(
              pageB,
              ["#seleccion_1"],
              60000,
            );
            if (!frm)
              throw new Error(
                "No se encontró #seleccion_1 (formulario de régimen).",
              );

            const frForm = frm.frame;

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

            await this.esperar(900);

            // 3) AVISO IMPORTANTE (si aparece)
            const foundAviso = await this.findAnyFrameWithAnySelector(
              pageB,
              ["#cheAviImport", "#botContAviso"],
              8000,
            );

            if (foundAviso) {
              const frAviso = foundAviso.frame;
              await frAviso.waitForSelector("#cheAviImport", {
                timeout: 20000,
              });
              const isChecked = await frAviso
                .$eval("#cheAviImport", (el) => el.checked)
                .catch(() => false);

              if (!isChecked)
                await frAviso.click("#cheAviImport", { delay: 30 });

              await frAviso.waitForSelector("#botContAviso", {
                timeout: 20000,
              });
              await frAviso.click("#botContAviso", { delay: 40 });

              await this.esperar(900);
            }

            // 4) LISTADO recibos (frame donde esté el enlace)
            const listado = await this.findAnyFrameWithAnySelector(
              pageB,
              [
                "#TABLA_2",
                'a[href*="AC_VER_RECIBO"]',
                "#enlace_0",
                "a.enlaceFuncDetalle",
                "a.pr_enlaceLocal",
              ],
              60000,
            );

            if (!listado)
              throw new Error(
                "No se encontró la lista de recibos (TABLA_2 / AC_VER_RECIBO).",
              );

            const frListado = listado.frame;

            // 5) Click al primer recibo (DENTRO del frame) + esperar a que aparezca el detalle
            // OJO: muchas veces NO hay navegación “real”, solo cambia el contenido.
            const waitDetalle = pageB
              .waitForSelector("#TABLA_3", { timeout: 45000 })
              .catch(() => null);

            const waitNav = pageB
              .waitForNavigation({
                waitUntil: "domcontentloaded",
                timeout: 15000,
              })
              .catch(() => null);

            // ✅ Click robusto (DOM click) en el frame del listado
            await this.clickPrimerReciboEnListado(frListado);

            // Espera a que pase “algo”: navegación o que aparezca la tabla del detalle
            await Promise.race([waitDetalle, waitNav]);

            // Asegura que estamos en detalle (tu HTML real tiene #TABLA_3)
            await pageB.waitForSelector("#TABLA_3", { timeout: 45000 });
            await this.esperar(400);

            // 6) Ya estamos en DETALLE ADEUDO (HTML normal). Esperar algo propio del detalle
            // (con tu HTML: id="TABLA_3" existe en el detalle)
            await pageB
              .waitForSelector("#TABLA_3", { timeout: 45000 })
              .catch(async () => {
                // fallback: si por lo que sea no aparece TABLA_3, al menos espera body
                await pageB
                  .waitForSelector("body", { timeout: 20000 })
                  .catch(() => {});
              });

            // ✅ 7) Generar PDF sin imprimir (sin chrome://print)
            await this.guardarPdfConPrintToPDF(pageB, pdfB, {
              timeoutMs: 60000,
            });

            okB++;
            const prev = logMap.get(dniKey) || "";
            logMap.set(
              dniKey,
              `${prev}${prev ? " | " : ""}OK_B: PDF B guardado -> ${path.basename(pdfB)}`,
            );
          } catch (e) {
            errB++;
            const prev = logMap.get(dniKey) || "";
            const msg = `ERROR_B: ${e?.message || e}`;
            logMap.set(dniKey, prev ? `${prev} | ${msg}` : msg);
            console.warn("[BASES/RECIBOS][B]", msg);
          }

          if ((i + 1) % 5 === 0) flushLog();
        }

        flushLog();

        console.log("[BASES/RECIBOS] Terminado.");
        console.log(
          `[BASES/RECIBOS] OK_A=${okA} ERR_A=${errA} | OK_B=${okB} ERR_B=${errB}`,
        );
        console.log(`[BASES/RECIBOS] Procesados: ${registrosProcesados}`);

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
