const path = require("path");
const fs = require("fs");
const XlsxPopulate = require("xlsx-populate");
const { registrarEjecucion } = require("../metricas");
const puppeteer = require("puppeteer");

// Activa esto solo si quieres ver console.log del navegador (frames FS)
// const DEBUG_FS_CONSOLE = true;
const DEBUG_FS_CONSOLE = false;

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

  /**
   * ✅ Si no hay dígitos, devuelve "" (no "00"/"000...").
   * Esto permite detectar campos faltantes en validación.
   */
  _padLeftDigitsOrEmpty(val, len) {
    const s = this._digitsOnly(val);
    if (!s) return "";
    return s.padStart(len, "0");
  }

  async ensureDir(dir) {
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
  }

  // =========================
  // ✅ Validar que el buffer ES un PDF real
  // =========================
  _isPdfBuffer(buf) {
    if (!buf || !Buffer.isBuffer(buf) || buf.length < 5) return false;
    return buf.subarray(0, 5).toString("utf8") === "%PDF-";
  }

  // =========================
  // ✅ Carpetas por cliente (DNI)
  // =========================
  async ensureClientDirs(rootOut, dni) {
    const dniFolder = this._safeFileName(this._dniNorm(dni) || "SIN_DNI");
    const dirCliente = path.join(rootOut, dniFolder);
    const dirClientePdf = path.join(dirCliente, "PDF");
    const dirClientePng = path.join(dirCliente, "CAPTURAS");

    await this.ensureDir(dirCliente);
    await this.ensureDir(dirClientePdf);
    await this.ensureDir(dirClientePng);

    return { dirCliente, dirClientePdf, dirClientePng, dniFolder };
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

    // ✅ Si falta algún campo, se queda vacío para que validarRegistro lo detecte.
    const provCCC = this._padLeftDigitsOrEmpty(r.provCCC, 2);
    const provNAF = this._padLeftDigitsOrEmpty(r.provNAF, 2);

    const ccc7 = this._padLeftDigitsOrEmpty(r.ccc7, 7);
    const ccc2 = this._padLeftDigitsOrEmpty(r.ccc2, 2);
    const ccc = ccc7 && ccc2 ? `${ccc7}${ccc2}` : "";

    const naf8 = this._padLeftDigitsOrEmpty(r.naf8, 8);
    const naf2 = this._padLeftDigitsOrEmpty(r.naf2, 2);
    const naf = naf8 && naf2 ? `${naf8}${naf2}` : "";

    return {
      exp: String(r.exp ?? "").trim(),
      empresa: String(r.empresa ?? "").trim(),
      regimen: "",

      // ✅ obligatorios para completar el proceso
      provCCC,
      ccc,

      trabajador: String(r.trabajador ?? "").trim(),
      dni,

      provNAF,
      naf,

      _row: r._row,
    };
  }

  /**
   * ✅ validación separada (faltan datos vs formato inválido)
   */
  validarRegistro(r) {
    const missing = [];
    const invalid = [];

    const req = (val, msg) => {
      if (val === null || val === undefined || String(val).trim() === "")
        missing.push(msg);
    };

    // Obligatorios (para completar el proceso completo)
    req(r.dni, "DNI vacío");
    req(r.trabajador, "TRABAJADOR/A vacío");
    req(r.regimen, "REGIMEN vacío (input manual)");
    req(r.provNAF, "PROV NAF vacío");
    req(r.naf, "NAF vacío");
    req(r.provCCC, "PROV CCC vacío");
    req(r.ccc, "CCC vacío");

    // Formatos
    if (r.provCCC && !/^\d{2}$/.test(r.provCCC))
      invalid.push("PROV CCC no parece 2 dígitos");
    if (r.provNAF && !/^\d{2}$/.test(r.provNAF))
      invalid.push("PROV NAF no parece 2 dígitos");
    if (r.ccc && !/^\d{9}$/.test(r.ccc))
      invalid.push("CCC no parece 9 dígitos (7+2)");
    if (r.naf && !/^\d{10}$/.test(r.naf))
      invalid.push("NAF no parece 10 dígitos (8+2)");
    if (r.regimen && !/^\d{4}$/.test(r.regimen))
      invalid.push("REGIMEN no parece 4 dígitos (ej: 0111)");

    return { missing, invalid };
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
          reason: "SKIP_DUPLICADO: DNI duplicado (se procesa la primera aparición)",
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
      console.warn("[DUPLICADOS][CAPTURA] No se pudo guardar screenshot:", e?.message || e);
      return false;
    }
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

        if (hit) return `Posible error detectado en pantalla (contiene '${hit}')`;
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
  // ✅ NUEVO: Reintento para errores de "detached frame"
  // =========================
  async withDetachedFrameRetry(fn, { retries = 2, label = "acción" } = {}) {
    let lastErr = null;
    for (let i = 0; i <= retries; i++) {
      try {
        return await fn();
      } catch (e) {
        lastErr = e;
        const msg = String(e?.message || e);
        if (!msg.toLowerCase().includes("detached frame")) throw e;
        console.warn(
          `[DUPLICADOS][RETRY] ${label}: detached frame (intento ${i + 1}/${retries + 1})`,
        );
        await this.esperar(700);
      }
    }
    throw lastErr;
  }

  // =========================
  // ✅ DIL (errores/avisos tras Continuar)
  // =========================
  async readDILInAnyFrame(page) {
    for (const fr of page.frames()) {
      try {
        const txt = await fr.evaluate(() => {
          const el = document.querySelector("#DIL");
          return el ? (el.textContent || "").trim() : "";
        });
        if (txt) return { frame: fr, text: txt };
      } catch (_) {}
    }
    return { frame: null, text: "" };
  }

  async waitForDILAfterContinuar(page, prevText, timeoutMs = 15000, pollMs = 300) {
    const start = Date.now();
    while (Date.now() - start < timeoutMs) {
      const { text } = await this.readDILInAnyFrame(page);
      if (text && (!prevText || text !== prevText)) return text;
      await this.esperar(pollMs);
    }
    const { text } = await this.readDILInAnyFrame(page);
    return text || "";
  }

  interpretarDIL(text) {
    const t = String(text || "").trim();
    if (!t) return null;

    if (t.includes("3083*") || t.toUpperCase().includes("INTRODUZCA LOS DATOS"))
      return null;

    if (/^\d{4}\*/.test(t)) return t;

    const low = t.toLowerCase();
    if (
      low.includes("incorrect") ||
      low.includes("error") ||
      low.includes("no se ha encontrado") ||
      low.includes("no existe") ||
      low.includes("datos incorrectos")
    ) {
      return t;
    }

    return null;
  }

  // =========================
  // ✅ Detectar el listado por cabeceras (más robusto)
  // - NO nos basamos en "ALTA (SIT.ACTUAL)" para localizar el frame.
  // - Localizamos el frame que contiene la tabla y además nos aseguramos
  //   de que existe una tabla con varias filas de datos.
  // =========================
  async findListadoTAFrame(page, timeoutMs = 60000, pollMs = 400) {
    const start = Date.now();
    while (Date.now() - start < timeoutMs) {
      for (const fr of page.frames()) {
        try {
          const ok = await fr.evaluate(() => {
            const bodyTxt = document.body?.innerText || "";
            if (!bodyTxt.includes("Documento TA") || !bodyTxt.includes("Fecha Real")) return false;

            const tables = Array.from(document.querySelectorAll("table"));
            const candidates = tables
              .map((t) => {
                const rows = Array.from(t.querySelectorAll("tr"));
                const dataRows = rows.filter((tr) => tr.querySelectorAll("td").length >= 2);
                return {
                  rowsCount: rows.length,
                  dataCount: dataRows.length,
                  txt: t.innerText || "",
                };
              })
              .filter((x) => x.txt.includes("Documento TA") && x.txt.includes("Fecha Real"))
              .sort((a, b) => b.dataCount - a.dataCount);

            return candidates.length && candidates[0].dataCount >= 1;
          });

          if (ok) return fr;
        } catch (_) {}
      }
      await this.esperar(pollMs);
    }
    return null;
  }

  // =========================
  // ✅ Seleccionar la fila "ALTA" con Fecha Real más actual (robusto)
  //    - Aplica la lógica del código antiguo: dblclick + click + click
  //    - Prioriza label/a dentro de la celda si existe
  //    - Ignora filas sin fecha parseable / vacías
  // =========================
async seleccionarAltaMasReciente(frameListado) {
  return await frameListado.evaluate(() => {
    const norm = (s) => String(s || "").replace(/\s+/g, " ").trim();

    // Acepta: 27/01/2026 | 27-01-2026 | 27.01.2026 | 27 01 2026
    const parseFecha = (s) => {
      const txt = norm(s);
      const m = txt.match(/(\d{2})\s*[\/\-. ]\s*(\d{2})\s*[\/\-. ]\s*(\d{4})/);
      if (!m) return null;
      const dd = Number(m[1]);
      const mm = Number(m[2]);
      const yy = Number(m[3]);
      const d = new Date(yy, mm - 1, dd);
      return isNaN(d.getTime()) ? null : d;
    };

    // ✅ Click “modo antiguo”: dblclick + click + click
    const fireClicksOldStyle = (el) => {
      if (!el) return false;
      try { el.scrollIntoView({ block: "center", inline: "center" }); } catch (_) {}

      try {
        el.dispatchEvent(new MouseEvent("dblclick", { bubbles: true, cancelable: true, view: window }));
      } catch (_) {}

      try { el.click(); } catch (_) {}
      try { el.click(); } catch (_) {}

      return true;
    };

    // 1) Seleccionamos la tabla más probable del listado
    const tables = Array.from(document.querySelectorAll("table"))
      .map((t) => {
        const txt = t.innerText || "";
        const rows = Array.from(t.querySelectorAll("tr"));
        const dataRows = rows.filter((tr) => tr.querySelectorAll("td").length >= 2);
        return { t, txt, dataCount: dataRows.length };
      })
      .filter((x) => x.txt.includes("Documento TA") && x.txt.includes("Fecha Real"))
      .sort((a, b) => b.dataCount - a.dataCount);

    const target = tables[0]?.t || null;
    if (!target) {
      return { ok: false, reason: "No se encontró la tabla del listado (Documento TA / Fecha Real)." };
    }

    const rows = Array.from(target.querySelectorAll("tr"));

    // 2) Detectar índices reales de columnas por la cabecera
    let docIdx = 0;
    let fechaIdx = 1;

    const headerRow = rows.find((tr) => {
      const cells = Array.from(tr.querySelectorAll("th,td"));
      const txts = cells.map((c) => norm(c.innerText || ""));
      return txts.some((x) => x.includes("Documento TA")) && txts.some((x) => x.includes("Fecha Real"));
    });

    if (headerRow) {
      const cells = Array.from(headerRow.querySelectorAll("th,td"));
      const txts = cells.map((c) => norm(c.innerText || ""));

      const foundDoc = txts.findIndex((x) => x.includes("Documento TA"));
      const foundFecha = txts.findIndex((x) => x.includes("Fecha Real"));

      if (foundDoc >= 0) docIdx = foundDoc;
      if (foundFecha >= 0) fechaIdx = foundFecha;
    }

    // 3) Buscar filas cuyo Documento EMPIECE por "ALTA" y elegir la fecha más actual
    const candidates = [];

    for (const tr of rows) {
      const tds = Array.from(tr.querySelectorAll("td"));
      if (tds.length < 2) continue;

      const docCell = tds[docIdx] || tds[0];
      const fechaCell = tds[fechaIdx] || tds[1];

      const docTxt = norm(docCell?.innerText || "");
      const fechaTxt = norm(fechaCell?.innerText || "");

      // ✅ "empiece por ALTA" (más estricto que includes)
      if (!/^ALTA\b/i.test(docTxt)) continue;

      const fecha = parseFecha(fechaTxt);
      if (!fecha) continue;

      // ✅ Prioridad: label > a > celda
      const label = docCell?.querySelector("label");
      const anchor = docCell?.querySelector("a");
      const clickable = label || anchor || docCell;

      candidates.push({
        doc: docTxt,
        fechaTxt,
        fechaMs: fecha.getTime(),
        clickable,
      });
    }

    if (!candidates.length) {
      return {
        ok: false,
        reason: "No se encontraron filas cuyo Documento empiece por 'ALTA' con Fecha Real parseable.",
      };
    }

    candidates.sort((a, b) => b.fechaMs - a.fechaMs);
    const best = candidates[0];

    const did = fireClicksOldStyle(best.clickable);
    if (!did) {
      return { ok: false, reason: "No se pudo clicar el elemento (modo antiguo) en la fila ALTA seleccionada." };
    }

    return { ok: true, doc: best.doc, fecha: best.fechaTxt };
  });
}


  // =========================
  // ✅ Descargar PDF RAW interceptando la respuesta (Fetch CDP)
  // =========================
  async descargarPdfRawViaFetchCDP(popupPage, outputPath, timeoutMs = 90000) {
    await popupPage.bringToFront().catch(() => {});
    const client = await popupPage.target().createCDPSession();

    const urlMatch = (url) => {
      const u = String(url || "");
      return (
        u.includes("/ImprPDF/") ||
        u.includes("InSeNaCoder") ||
        u.toLowerCase().endsWith(".pdf")
      );
    };

    await client
      .send("Fetch.enable", {
        patterns: [
          { urlPattern: "*w2.seg-social.es/ImprPDF/*", requestStage: "Response" },
          { urlPattern: "*/ImprPDF/*", requestStage: "Response" },
        ],
      })
      .catch(() => {});

    await client.send("Network.enable").catch(() => {});
    await client
      .send("Network.setCacheDisabled", { cacheDisabled: true })
      .catch(() => {});

    const timer = new Promise((_, rej) =>
      setTimeout(() => rej(new Error("Timeout esperando el PDF (Fetch CDP).")), timeoutMs),
    );

    let done = false;

    const pdfPromise = new Promise((resolve, reject) => {
      const onPaused = async (ev) => {
        if (done) {
          try {
            await client.send("Fetch.continueRequest", { requestId: ev.requestId }).catch(() => {});
          } catch (_) {}
          return;
        }

        try {
          const reqId = ev.requestId;
          const url = ev.request?.url || "";
          const status = ev.responseStatusCode || 0;

          if (!urlMatch(url) || !(status >= 200 && status < 300)) {
            await client.send("Fetch.continueRequest", { requestId: reqId }).catch(() => {});
            return;
          }

          const bodyResp = await client.send("Fetch.getResponseBody", { requestId: reqId });
          const body = bodyResp?.body || "";
          const base64Encoded = !!bodyResp?.base64Encoded;

          await client.send("Fetch.continueRequest", { requestId: reqId }).catch(() => {});

          const buf = base64Encoded ? Buffer.from(body, "base64") : Buffer.from(body, "utf8");

          if (!this._isPdfBuffer(buf)) {
            return reject(
              new Error(
                "El contenido capturado NO es un PDF (no empieza por %PDF-). Probablemente es el wrapper del visor.",
              ),
            );
          }

          done = true;
          try {
            client.removeListener("Fetch.requestPaused", onPaused);
          } catch (_) {}

          resolve({ buf, url, status });
        } catch (e) {
          try {
            client.removeListener("Fetch.requestPaused", onPaused);
          } catch (_) {}
          reject(e);
        }
      };

      client.on("Fetch.requestPaused", onPaused);
    });

    await popupPage.reload({ waitUntil: "domcontentloaded" }).catch(() => {});
    await popupPage.waitForSelector("body", { timeout: 20000 }).catch(() => {});
    await this.esperar(800);

    const { buf, url, status } = await Promise.race([pdfPromise, timer]);

    fs.writeFileSync(outputPath, buf);
    console.log("[DUPLICADOS][PDF] PDF RAW guardado:", status, url, "->", outputPath);

    try {
      await client.send("Fetch.disable");
    } catch (_) {}
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
          console.error("[DUPLICADOS][INPUT] Ruta a chrome.exe no válida.");
          return resolve(false);
        }
        if (!pathExcel || typeof pathExcel !== "string" || !fs.existsSync(pathExcel)) {
          console.error("[DUPLICADOS][INPUT] Ruta a Excel no válida.");
          return resolve(false);
        }
        if (!pathSalidaBase || typeof pathSalidaBase !== "string" || !pathSalidaBase.trim()) {
          console.error("[DUPLICADOS][INPUT] Ruta de salida no válida.");
          return resolve(false);
        }
        if (!/^\d{4}$/.test(regimen4)) {
          console.error("[DUPLICADOS][INPUT] Régimen inválido. Debe ser 4 dígitos (ej: 0111).");
          return resolve(false);
        }

        const rootOut = path.join(
          path.normalize(pathSalidaBase),
          `Duplicados TA2 (${this.getCurrentDateString()})`,
        );

        const dirLogs = path.join(rootOut, "LOGS");
        await this.ensureDir(rootOut);
        await this.ensureDir(dirLogs);

        const resumenPath = path.join(dirLogs, "resumen.json");
        const detallePath = path.join(dirLogs, "detalle.json");
        const detalleTxtPath = path.join(dirLogs, "detalle.log");

        let resumen = {
          ok: [],
          error: [],
          skipped: [],
          stats: {},
          generated_at: new Date().toISOString(),
        };
        if (fs.existsSync(resumenPath)) {
          try {
            resumen = JSON.parse(fs.readFileSync(resumenPath, "utf8"));
          } catch (_) {}
        }
        const okSet = new Set((resumen.ok || []).map((x) => this._dniNorm(x)));

        const flushLogs = (logsPorDni, stats) => {
          try {
            resumen.stats = {
              ...(resumen.stats || {}),
              ...stats,
              generated_at: new Date().toISOString(),
            };
            fs.writeFileSync(resumenPath, JSON.stringify(resumen, null, 2), "utf8");

            const detalle = Array.from(logsPorDni.entries()).map(([k, v]) => ({ key: k, msg: v }));
            fs.writeFileSync(detallePath, JSON.stringify({ detalle }, null, 2), "utf8");

            const lines = detalle.map((x) => `${x.key} -> ${x.msg}`);
            fs.writeFileSync(detalleTxtPath, lines.join("\n"), "utf8");
          } catch (e) {
            console.warn("[DUPLICADOS][LOGS] No se pudo escribir logs:", e?.message || e);
          }
        };

        console.log("[DUPLICADOS][INPUT] Leyendo Excel:", path.normalize(pathExcel));
        const { rows } = await this.leerExcelDuplicados(pathExcel);

        for (const r of rows) r.regimen = regimen4;

        const logsPorDni = new Map();
        const toProcess = [];
        const skippedMissing = [];
        const skippedInvalid = [];

        for (const r of rows) {
          const { missing, invalid } = this.validarRegistro(r);

          if (missing.length) {
            const key = this._dniNorm(r.dni) || `ROW_${r._row}`;
            const reason = `SKIP_FALTA_DATOS: ${missing.join(" | ")}`;
            logsPorDni.set(key, reason);
            skippedMissing.push({ dni: this._dniNorm(r.dni), row: r._row, reason });
            resumen.skipped.push({ dni: this._dniNorm(r.dni), row: r._row, reason });
            continue;
          }

          if (invalid.length) {
            const key = this._dniNorm(r.dni) || `ROW_${r._row}`;
            const reason = `SKIP_FORMATO_INVALIDO: ${invalid.join(" | ")}`;
            logsPorDni.set(key, reason);
            skippedInvalid.push({ dni: this._dniNorm(r.dni), row: r._row, reason });
            resumen.skipped.push({ dni: this._dniNorm(r.dni), row: r._row, reason });
            continue;
          }

          toProcess.push(r);
        }

        const { kept, skipped } = this.deduplicarPorDNI(toProcess);
        for (const s of skipped) {
          logsPorDni.set(s.dni, s.reason);
          resumen.skipped.push(s);
        }

        const stats = {
          total_read: rows.length,
          total_to_process_pre_dedupe: toProcess.length,
          total_skip_missing: skippedMissing.length,
          total_skip_invalid: skippedInvalid.length,
          total_skip_duplicate: skipped.length,
          total_to_process: kept.length,
        };

        console.log(`[DUPLICADOS][INPUT] Leídos: ${stats.total_read}`);
        console.log(`[DUPLICADOS][INPUT] A procesar (pre-dedupe): ${stats.total_to_process_pre_dedupe}`);
        console.log(`[DUPLICADOS][INPUT] Skips falta datos: ${stats.total_skip_missing}`);
        console.log(`[DUPLICADOS][INPUT] Skips formato inválido: ${stats.total_skip_invalid}`);
        console.log(`[DUPLICADOS][INPUT] Skips duplicado: ${stats.total_skip_duplicate}`);
        console.log(`[DUPLICADOS][INPUT] A procesar (final): ${stats.total_to_process}`);

        if (!kept.length) {
          console.warn("[DUPLICADOS][INPUT] No hay registros válidos para procesar.");
          flushLogs(logsPorDni, stats);
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

        // ✅ SOLO una vez (evita duplicados)
        if (DEBUG_FS_CONSOLE) {
          page.on("console", (msg) => {
            try {
              console.log("[FS-CONSOLE]", msg.text());
            } catch (_) {}
          });
        }

        page.on("dialog", async (dialog) => {
          try {
            await dialog.accept();
          } catch (_) {}
        });

        await page.goto(urlFS, { waitUntil: "domcontentloaded" });
        console.log("[DUPLICADOS][FS] FS abierto. Selecciona el certificado si aparece.");

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

          const { dirClientePdf, dirClientePng } = await this.ensureClientDirs(rootOut, dni);

          const pngPath = path.join(dirClientePng, `Cuadro TA2 SS ${trabajador}.png`);
          const pdfPath = path.join(dirClientePdf, `TA2 A010112 ${trabajador}.pdf`);

          if (okSet.has(dni)) {
            logsPorDni.set(dni, "SKIP_OK_PREVIO: ya estaba OK (modo resume)");
            return;
          }

          console.log(
            `[DUPLICADOS][PROC] ${idx + 1}/${kept.length} | DNI: ${dni} | TRABAJADOR: ${r.trabajador}`,
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

          await this.withDetachedFrameRetry(() => fillInput(frameForm, "#SDFTESNAF", provNAF), {
            label: "fill PROV NAF",
          });
          await this.withDetachedFrameRetry(() => fillInput(frameForm, "#SDFNAF", naf10), {
            label: "fill NAF",
          });
          await this.withDetachedFrameRetry(() => fillInput(frameForm, "#SDFREGCTA_NH", regimen4), {
            label: "fill REGIMEN",
          });
          await this.withDetachedFrameRetry(() => fillInput(frameForm, "#SDFTESCTA", provCCC), {
            label: "fill PROV CCC",
          });
          await this.withDetachedFrameRetry(() => fillInput(frameForm, "#SDFCUENTA", ccc9), {
            label: "fill CCC",
          });

          await frameForm.waitForSelector("#ListaTipoImpresion", { timeout: 30000 });
          await frameForm.select("#ListaTipoImpresion", "OnLine");

          const { text: dilBefore } = await this.readDILInAnyFrame(page);

          const clickedContinuar = await this.clickContinuarRobusta({
            page,
            frameForm,
            timeoutMs: 35000,
          });
          if (!clickedContinuar) {
            throw new Error("No se pudo pulsar 'Continuar' (click robusto falló).");
          }

          const dilAfter = await this.waitForDILAfterContinuar(page, dilBefore, 15000);
          if (dilAfter) console.log(`[DUPLICADOS][DIL] ${dni} -> ${dilAfter}`);

          const dilError = this.interpretarDIL(dilAfter);
          if (dilError) {
            throw new Error(`Validación FS (DIL): ${dilError}`);
          }

          // ✅ Localizar listado
          const frameListado = await this.findListadoTAFrame(page, 90000);
          if (!frameListado) {
            const maybeError = await this.detectPossibleErrorInFrames(page);
            throw new Error(
              `No se detectó el listado TA tras Continuar.${maybeError ? " " + maybeError : ""}`,
            );
          }

          // ✅ REQUISITO: CAPTURA LO PRIMERO (tabla cargada) ANTES de abrir PDF
          await this.safeScreenshot(page, pngPath);

          // ✅ Preparar popupPromise ANTES de hacer doble click
          const popupPromise = this.waitForPopup(browser, page, 45000);

          // ✅ Seleccionar ALTA más reciente (con reintento detached)
          const sel = await this.withDetachedFrameRetry(
            () => this.seleccionarAltaMasReciente(frameListado),
            { label: "seleccionar ALTA más reciente" },
          );

          if (!sel.ok) {
            // ✅ Si no existe ALTA (solo BAJA/CAMBIO): error, log y siguiente registro
            throw new Error(`No se ha encontrado ALTA en el listado: ${sel.reason}`);
          }

          console.log(`[DUPLICADOS][SELECCION] ${dni} -> ${sel.doc} | Fecha Real: ${sel.fecha}`);

          // Esperar popup
          let popupPage = await popupPromise;

          // ✅ Fallback: a veces FS no abre la pestaña a la primera (timing/handler)
          if (!popupPage) {
            console.warn(`[DUPLICADOS][POPUP] No abrió a la primera. Reintentando click ALTA...`);

            // Prepara otra espera de popup y repite la interacción
            const popupPromise2 = this.waitForPopup(browser, page, 25000);

            const sel2 = await this.withDetachedFrameRetry(
              () => this.seleccionarAltaMasReciente(frameListado),
              { label: "reintento seleccionar ALTA" },
            );

            if (!sel2.ok) {
              throw new Error(`Reintento: no se ha encontrado ALTA en el listado: ${sel2.reason}`);
            }

            popupPage = await popupPromise2;
          }

          if (!popupPage) {
            throw new Error(
              "Se esperaba una nueva pestaña con el PDF, pero no se abrió (tras reintento).",
            );
          }

          await this.descargarPdfRawViaFetchCDP(popupPage, pdfPath, 90000);

          logsPorDni.set(dni, `OK: PDF guardado -> ${path.basename(pdfPath)}`);
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
            logsPorDni.set(dni || `ROW_${r._row}`, msg);
            resumen.error.push({ dni, row: r._row, error: msg });
            console.warn("[DUPLICADOS]", msg);
          }

          if ((i + 1) % 5 === 0) flushLogs(logsPorDni, stats);
        }

        flushLogs(logsPorDni, stats);

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
