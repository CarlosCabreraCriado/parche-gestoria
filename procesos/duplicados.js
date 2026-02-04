const path = require("path");
const fs = require("fs");
const XlsxPopulate = require("xlsx-populate");
const { registrarEjecucion } = require("../metricas");
const puppeteer = require("puppeteer");

// Activa esto solo si quieres ver console.log del navegador (frames FS)
// const DEBUG_FS_CONSOLE = true;
const DEBUG_FS_CONSOLE = false;

/**
 * Procesos de Duplicados (TA2 / IDC)
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

  // ✅ Espera simple (sin console.time)
  async esperarLog(ms, _label) {
    await this.esperar(ms);
  }

  // ✅ Wrapper sin console.time (mantiene la firma para no romper llamadas)
  async timeAsync(_label, fn) {
    return await fn();
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

  _dniFolderKey(dni) {
    const s = this._dniNorm(dni);
    if (!s) return "";
    if (/[A-Z]$/.test(s)) return s.slice(0, -1);
    return s;
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

  // Convierte:
  // - "27/01/2026" | "27-01-2026" | "27.01.2026" -> "A270126"
  // - "01 01 2012" -> "A010112"  ✅ (IDC y a veces TA2 viene con espacios)
  _fechaRealToA(fechaRealTxt) {
    const s = String(fechaRealTxt || "").trim();

    // dd/mm/yyyy o dd-mm-yyyy o dd.mm.yyyy (y también con espacios)
    let m = s.match(/(\d{2})\s*[\/\-. ]\s*(\d{2})\s*[\/\-. ]\s*(\d{4})/);

    // dd mm yyyy (espacios)
    if (!m) m = s.match(/(\d{2})\s+(\d{2})\s+(\d{4})/);

    if (!m) return "A010112"; // fallback conservador
    const dd = m[1];
    const mm = m[2];
    const yy = m[3].slice(-2);
    return `A${dd}${mm}${yy}`;
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
  // Carpetas por cliente (DNI)
  // =========================
  async ensureClientDirFlat(rootOut, dni) {
    const dniKey = this._dniFolderKey(dni);
    const dniFolder = this._safeFileName(dniKey || "SIN_DNI");
  
    const dirDni = path.join(rootOut, dniFolder);
    await this.ensureDir(dirDni);
  
    return { dirDni, dniFolder };
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
    const missing = [];
    const invalid = [];

    const req = (val, msg) => {
      if (val === null || val === undefined || String(val).trim() === "")
        missing.push(msg);
    };

    req(r.dni, "DNI vacío");
    req(r.trabajador, "TRABAJADOR/A vacío");
    req(r.regimen, "REGIMEN vacío (input manual)");
    req(r.provNAF, "PROV NAF vacío");
    req(r.naf, "NAF vacío");
    req(r.provCCC, "PROV CCC vacío");
    req(r.ccc, "CCC vacío");

    if (r.provCCC && !/^\d{2}$/.test(r.provCCC)) invalid.push("PROV CCC no parece 2 dígitos");
    if (r.provNAF && !/^\d{2}$/.test(r.provNAF)) invalid.push("PROV NAF no parece 2 dígitos");
    if (r.ccc && !/^\d{9}$/.test(r.ccc)) invalid.push("CCC no parece 9 dígitos (7+2)");
    if (r.naf && !/^\d{10}$/.test(r.naf)) invalid.push("NAF no parece 10 dígitos (8+2)");
    if (r.regimen && !/^\d{4}$/.test(r.regimen)) invalid.push("REGIMEN no parece 4 dígitos (ej: 0111)");

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
              const norm2 = (s) => (s || "").trim().toLowerCase().replace(/\s+/g, " ");
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
          const norm = (s) => (s || "").replace(/\s+/g, " ").trim().toLowerCase();
          const target = norm(t);

          const candidates = [
            ...Array.from(document.querySelectorAll("button")),
            ...Array.from(
              document.querySelectorAll('input[type="button"], input[type="submit"]'),
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
            return t.type() === "page" && t.opener() && t.opener() === openerPage.target();
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
  // ✅ Reintento detached frame
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
  // ✅ DIL
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

    if (t.includes("3083*") || t.toUpperCase().includes("INTRODUZCA LOS DATOS")) return null;

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
  // ✅ TA2: detectar listado por cabeceras
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
                return { dataCount: dataRows.length, txt: t.innerText || "" };
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

  // ✅ TA2: seleccionar ALTA más reciente (clic en label/a/celda de Documento)
  async seleccionarAltaMasReciente(frameListado) {
    return await frameListado.evaluate(() => {
      const norm = (s) => String(s || "").replace(/\s+/g, " ").trim();

      const parseFecha = (s) => {
        const txt = norm(s);
        const m = txt.match(/(\d{2})\s*[\/\-. ]\s*(\d{2})\s*[\/\-. ]\s*(\d{4})/);
        if (!m) return null;
        const d = new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
        return isNaN(d.getTime()) ? null : d;
      };

      const fireClicksOldStyle = (el) => {
        if (!el) return false;
        try { el.scrollIntoView({ block: "center", inline: "center" }); } catch (_) {}
        try { el.dispatchEvent(new MouseEvent("dblclick", { bubbles: true, cancelable: true, view: window })); } catch (_) {}
        try { el.click(); } catch (_) {}
        try { el.click(); } catch (_) {}
        return true;
      };

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
      if (!target) return { ok: false, reason: "No se encontró la tabla del listado (Documento TA / Fecha Real)." };

      const rows = Array.from(target.querySelectorAll("tr"));

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

      const candidates = [];

      for (const tr of rows) {
        const tds = Array.from(tr.querySelectorAll("td"));
        if (tds.length < 2) continue;

        const docCell = tds[docIdx] || tds[0];
        const fechaCell = tds[fechaIdx] || tds[1];

        const docTxt = norm(docCell?.innerText || "");
        const fechaTxt = norm(fechaCell?.innerText || "");

        if (!/^ALTA\b/i.test(docTxt)) continue;

        const fecha = parseFecha(fechaTxt);
        if (!fecha) continue;

        const label = docCell?.querySelector("label");
        const anchor = docCell?.querySelector("a");
        const clickable = label || anchor || docCell;

        candidates.push({ doc: docTxt, fechaTxt, fechaMs: fecha.getTime(), clickable });
      }

      if (!candidates.length) {
        return { ok: false, reason: "No se encontraron filas 'ALTA' con Fecha Real parseable." };
      }

      candidates.sort((a, b) => b.fechaMs - a.fechaMs);
      const best = candidates[0];

      const did = fireClicksOldStyle(best.clickable);
      if (!did) return { ok: false, reason: "No se pudo clicar el elemento en la fila ALTA seleccionada." };

      return { ok: true, doc: best.doc, fecha: best.fechaTxt };
    });
  }

  // =========================
  // ✅ IDC: detectar listado por cabeceras
  // =========================
  async findListadoIDCFrame(page, timeoutMs = 60000, pollMs = 400) {
    const start = Date.now();
    while (Date.now() - start < timeoutMs) {
      for (const fr of page.frames()) {
        try {
          const ok = await fr.evaluate(() => {
            const bodyTxt = document.body?.innerText || "";
            if (!bodyTxt.includes("F. R. Alta")) return false;

            const tables = Array.from(document.querySelectorAll("table"));
            const candidates = tables
              .map((t) => ({
                t,
                txt: t.innerText || "",
                dataRows: Array.from(t.querySelectorAll("tr")).filter((tr) => tr.querySelectorAll("td").length >= 2).length,
              }))
              .filter((x) => x.txt.includes("F. R. Alta") && x.txt.includes("F. R. Baja"))
              .sort((a, b) => b.dataRows - a.dataRows);

            return candidates.length && candidates[0].dataRows >= 1;
          });

          if (ok) return fr;
        } catch (_) {}
      }
      await this.esperar(pollMs);
    }
    return null;
  }

  // ✅ IDC: seleccionar F. R. Alta más reciente (dblclick en LABEL de la FECHA)
  async seleccionarAltaIDCMasReciente(frameListado) {
    return await frameListado.evaluate(() => {
      const norm = (s) => String(s || "").replace(/\s+/g, " ").trim();

      const parseFechaIDC = (s) => {
        const txt = norm(s);
        const m = txt.match(/(\d{2})\s+(\d{2})\s+(\d{4})/);
        if (!m) return null;
        const d = new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
        return isNaN(d.getTime()) ? null : d;
      };

      const fireClicksOldStyle = (el) => {
        if (!el) return false;
        try { el.scrollIntoView({ block: "center", inline: "center" }); } catch (_) {}
        try { el.dispatchEvent(new MouseEvent("dblclick", { bubbles: true, cancelable: true, view: window })); } catch (_) {}
        try { el.click(); } catch (_) {}
        try { el.click(); } catch (_) {}
        return true;
      };

      const tables = Array.from(document.querySelectorAll("table"))
        .map((t) => {
          const txt = t.innerText || "";
          const rows = Array.from(t.querySelectorAll("tr"));
          const dataRows = rows.filter((tr) => tr.querySelectorAll("td").length >= 2);
          return { t, txt, dataCount: dataRows.length };
        })
        .filter((x) => x.txt.includes("F. R. Alta") && x.txt.includes("F. R. Baja"))
        .sort((a, b) => b.dataCount - a.dataCount);

      const target = tables[0]?.t || null;
      if (!target) return { ok: false, reason: "No se encontró la tabla IDC (F. R. Alta / F. R. Baja)." };

      const rows = Array.from(target.querySelectorAll("tr"));

      let altaIdx = 0;
      const headerRow = rows.find((tr) => {
        const cells = Array.from(tr.querySelectorAll("th,td"));
        const txts = cells.map((c) => norm(c.innerText || ""));
        return txts.some((x) => x.includes("F. R. Alta"));
      });

      if (headerRow) {
        const cells = Array.from(headerRow.querySelectorAll("th,td"));
        const txts = cells.map((c) => norm(c.innerText || ""));
        const idx = txts.findIndex((x) => x.includes("F. R. Alta"));
        if (idx >= 0) altaIdx = idx;
      }

      const candidates = [];

      for (const tr of rows) {
        const tds = Array.from(tr.querySelectorAll("td"));
        if (!tds.length) continue;

        const altaCell = tds[altaIdx] || tds[0];

        // ✅ Lo que me confirmas: dblclick en esa columna, pero en el registro de la FECHA (label)
        const label = altaCell?.querySelector("label");
        const altaTxt = norm(label?.innerText || altaCell?.innerText || "");
        if (!altaTxt) continue;

        const fecha = parseFechaIDC(altaTxt);
        if (!fecha) continue;

        const clickable = label || altaCell;

        candidates.push({ altaTxt, fechaMs: fecha.getTime(), clickable });
      }

      if (!candidates.length) {
        return { ok: false, reason: "No hay filas con 'F. R. Alta' parseable (dd mm yyyy)." };
      }

      candidates.sort((a, b) => b.fechaMs - a.fechaMs);
      const best = candidates[0];

      const did = fireClicksOldStyle(best.clickable);
      if (!did) return { ok: false, reason: "No se pudo hacer dblclick/click en el label de la fecha (F. R. Alta)." };

      return { ok: true, fecha: best.altaTxt };
    });
  }


  // =========================
  // ✅ Descargar PDF RAW (Fetch CDP) - ROBUSTO (cleanup SIEMPRE)
  // =========================
  async descargarPdfRawViaFetchCDP(popupPage, outputPath, timeoutMs = 90000) {
    await popupPage.bringToFront().catch(() => {});
    const client = await popupPage.target().createCDPSession();

    const urlMatch = (url) => {
      const u = String(url || "");
      return u.includes("/ImprPDF/") || u.includes("InSeNaCoder") || u.toLowerCase().endsWith(".pdf");
    };

    let done = false;
    let onPaused = null;

    try {
      await client.send("Network.enable").catch(() => {});
      await client.send("Network.setCacheDisabled", { cacheDisabled: true }).catch(() => {});

      await client
        .send("Fetch.enable", {
          patterns: [
            { urlPattern: "*w2.seg-social.es/ImprPDF/*", requestStage: "Response" },
            { urlPattern: "*/ImprPDF/*", requestStage: "Response" },
          ],
        })
        .catch(() => {});

      const timer = new Promise((_, rej) =>
        setTimeout(() => rej(new Error("Timeout esperando el PDF (Fetch CDP).")), timeoutMs),
      );

      const pdfPromise = new Promise((resolve, reject) => {
        onPaused = async (ev) => {
          // Si ya resolvimos, no bloquees nada.
          if (done) {
            try { await client.send("Fetch.continueRequest", { requestId: ev.requestId }).catch(() => {}); } catch (_) {}
            return;
          }

          try {
            const reqId = ev.requestId;
            const url = ev.request?.url || "";
            const status = ev.responseStatusCode || 0;

            // No es el PDF -> continúa.
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
            resolve({ buf, url, status });
          } catch (e) {
            reject(e);
          }
        };

        client.on("Fetch.requestPaused", onPaused);
      });

      // A veces el visor carga “HTML wrapper” primero; el reload ayuda.
      await popupPage.reload({ waitUntil: "domcontentloaded" }).catch(() => {});
      await popupPage.waitForSelector("body", { timeout: 20000 }).catch(() => {});
      await this.esperarLog(800, "descargarPdfRawViaFetchCDP post_body");

      const { buf, url, status } = await Promise.race([pdfPromise, timer]);

      fs.writeFileSync(outputPath, buf);

      console.log(
        "[DUPLICADOS][PDF] PDF RAW guardado:",
        status,
        url,
        "->",
        outputPath,
        `| bytes=${buf?.length || 0}`,
      );

      return true;
    } finally {
      // ✅ Cleanup SIEMPRE, haya ok o error (evita “se queda colgado”)
      try {
        if (onPaused) client.removeListener("Fetch.requestPaused", onPaused);
      } catch (_) {}

      try { await client.send("Fetch.disable").catch(() => {}); } catch (_) {}
      try { await client.detach(); } catch (_) {}
    }
  }

  _isPdfDownloadRetryableError(err) {
    const msg = String(err?.message || err).toLowerCase();
    return (
      msg.includes("no es un pdf") ||
      msg.includes("wrapper del visor") ||
      msg.includes("timeout esperando el pdf")
    );
  }
  
  async descargarPdfCon1Reintento({ popupPage, outputPath, timeoutMs = 90000, label = "PDF" }) {
    try {
      await this.descargarPdfRawViaFetchCDP(popupPage, outputPath, timeoutMs);
      return;
    } catch (e) {
      // Si no es un fallo típico de “visor/tiempos”, no reintentes.
      if (!this._isPdfDownloadRetryableError(e)) throw e;
    
      console.warn(`[DUPLICADOS][${label}] Fallo descarga (1er intento). Reintentando 1 vez...`, e?.message || e);
    
      // Cierra el popup “malo” antes de reintentar
      try { await popupPage.close(); } catch (_) {}
    
      await this.esperarLog(1200, `retry_${label}_wait`);
    
      // El reintento NO crea popup nuevo aquí: lo haremos desde el caller (TA2/IDC),
      // porque necesitas volver a hacer click/dblclick para abrir de nuevo el PDF.
      throw e; // el caller detectará y hará la reapertura + 2º intento
    }
  }
  
  

  // =========================
  // PROCESO ÚNICO: TA2 + IDC por empleado
  // =========================
  async duplicadosTa2(argumentos) {
    console.log("[DUPLICADOS] Iniciando proceso DUPLICADOS TA2 + IDC (por empleado)");

    const nombreProceso = "DUPLICADOS TA2+IDC";
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

        const rootOut = path.join(path.normalize(pathSalidaBase), `Duplicados (${this.getCurrentDateString()})`);
        const dirLogs = path.join(rootOut, "LOGS");
        await this.ensureDir(rootOut);
        await this.ensureDir(dirLogs);

        // ✅ logs separados para resume independiente
        const detalleTA2Path = path.join(dirLogs, "detalle_ta2.log");
        const detalleIDCPath = path.join(dirLogs, "detalle_idc.log");

        const loadOkSet = (filePath) => {
          const set = new Set();
          if (!fs.existsSync(filePath)) return set;
          try {
            const prev = fs.readFileSync(filePath, "utf8");
            prev.split(/\r?\n/).forEach((line) => {
              const m = line.match(/^(.+?)\s*->\s*OK:/i);
              if (m) set.add(this._dniNorm(m[1]));
            });
          } catch (_) {}
          return set;
        };

        const okSetTA2 = loadOkSet(detalleTA2Path);
        const okSetIDC = loadOkSet(detalleIDCPath);

        const flushLogs = (filePath, logsMap) => {
          try {
            const lines = Array.from(logsMap.entries()).map(([k, v]) => `${k} -> ${v}`);
            fs.writeFileSync(filePath, lines.join("\n"), "utf8");
          } catch (e) {
            console.warn("[DUPLICADOS][LOGS] No se pudo escribir log:", filePath, e?.message || e);
          }
        };

        console.log("[DUPLICADOS][INPUT] Leyendo Excel:", path.normalize(pathExcel));
        const { rows } = await this.timeAsync("excel_leerExcelDuplicados", () => this.leerExcelDuplicados(pathExcel));
        for (const r of rows) r.regimen = regimen4;

        const logsTA2 = new Map();
        const logsIDC = new Map();

        const toProcess = [];
        const skippedMissing = [];
        const skippedInvalid = [];

        for (const r of rows) {
          const { missing, invalid } = this.validarRegistro(r);

          if (missing.length) {
            const key = this._dniNorm(r.dni) || `ROW_${r._row}`;
            const reason = `SKIP_FALTA_DATOS: ${missing.join(" | ")}`;
            logsTA2.set(key, reason);
            logsIDC.set(key, reason);
            skippedMissing.push({ dni: this._dniNorm(r.dni), row: r._row, reason });
            continue;
          }

          if (invalid.length) {
            const key = this._dniNorm(r.dni) || `ROW_${r._row}`;
            const reason = `SKIP_FORMATO_INVALIDO: ${invalid.join(" | ")}`;
            logsTA2.set(key, reason);
            logsIDC.set(key, reason);
            skippedInvalid.push({ dni: this._dniNorm(r.dni), row: r._row, reason });
            continue;
          }

          toProcess.push(r);
        }

        const { kept, skipped } = this.deduplicarPorDNI(toProcess);
        for (const s of skipped) {
          logsTA2.set(s.dni, s.reason);
          logsIDC.set(s.dni, s.reason);
        }

        console.log(`[DUPLICADOS][INPUT] Leídos: ${rows.length}`);
        console.log(`[DUPLICADOS][INPUT] A procesar (final): ${kept.length}`);

        if (!kept.length) {
          console.warn("[DUPLICADOS][INPUT] No hay registros válidos para procesar.");
          flushLogs(detalleTA2Path, logsTA2);
          flushLogs(detalleIDCPath, logsIDC);
          return resolve(false);
        }

        const urlFS = "https://w2.seg-social.es/fs/indexframes.html";

        browser = await this.timeAsync("puppeteer_launch", () =>
          puppeteer.launch({
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
          }),
        );

        const opened = await browser.pages();
        const page = opened.length ? opened[0] : await browser.newPage();

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
          await this.esperarLog(1200, "openAFIOnlineReal");
        };

        const openATR65Duplicados = async () => {
          const ok = await this.clickLinkInFrames(page, {
            hrefIncludes: "TRANSACCION=ATR65",
            textIncludes: "Duplicados de documentos trabajador",
          });
          if (!ok) throw new Error("No se pudo clicar 'Duplicados de documentos trabajador' (ATR65)");
          await this.esperarLog(1200, "openATR65Duplicados");
        };

        const openATR37IDC = async () => {
          const ok = await this.clickLinkInFrames(page, {
            hrefIncludes: "TRANSACCION=ATR37",
            textIncludes: "Informe datos de cotización-Trab.Cuenta Ajena",
          });
          if (!ok) throw new Error("No se pudo clicar ATR37 (IDC)");
          await this.esperarLog(1200, "openATR37IDC");
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

        // -------------------------
        // ✅ Ejecutar TA2 para 1 empleado
        // -------------------------
        const procesarTA2 = async (r, idx, total) => {
          const dni = this._dniNorm(r.dni);
          const trabajador = this._safeFileName(r.trabajador);

          if (okSetTA2.has(dni)) {
            logsTA2.set(dni, "SKIP_OK_PREVIO: TA2 ya estaba OK (resume)");
            return;
          }

          const { dirDni } = await this.ensureClientDirFlat(rootOut, dni);
          const pngPath = path.join(dirDni, `Cuadro TA2 SS ${trabajador}.png`);

          console.log(`[DUPLICADOS][TA2] ${idx + 1}/${total} | DNI: ${dni} | TRABAJADOR: ${r.trabajador}`);

          await page.goto(urlFS, { waitUntil: "domcontentloaded" });
          await this.esperarLog(800);

          await openAFIOnlineReal();
          await openATR65Duplicados();

          const frameForm = await this.findFrameWithSelector(page, "#SDFTESNAF", 30000);
          if (!frameForm) throw new Error("No se encontró el formulario ATR65 (selector #SDFTESNAF)");

          const provNAF = this._padLeftDigits(r.provNAF, 2);
          const naf10 = this._padLeftDigits(r.naf, 10);
          const provCCC = this._padLeftDigits(r.provCCC, 2);
          const ccc9 = this._padLeftDigits(r.ccc, 9);

          await this.withDetachedFrameRetry(() => fillInput(frameForm, "#SDFTESNAF", provNAF), { label: "TA2 fill PROV NAF" });
          await this.withDetachedFrameRetry(() => fillInput(frameForm, "#SDFNAF", naf10), { label: "TA2 fill NAF" });
          await this.withDetachedFrameRetry(() => fillInput(frameForm, "#SDFREGCTA_NH", regimen4), { label: "TA2 fill REGIMEN" });
          await this.withDetachedFrameRetry(() => fillInput(frameForm, "#SDFTESCTA", provCCC), { label: "TA2 fill PROV CCC" });
          await this.withDetachedFrameRetry(() => fillInput(frameForm, "#SDFCUENTA", ccc9), { label: "TA2 fill CCC" });

          await frameForm.waitForSelector("#ListaTipoImpresion", { timeout: 30000 });
          await frameForm.select("#ListaTipoImpresion", "OnLine");

          const { text: dilBefore } = await this.readDILInAnyFrame(page);

          const clicked = await this.clickContinuarRobusta({ page, frameForm, timeoutMs: 35000 });
          if (!clicked) throw new Error("TA2: no se pudo pulsar 'Continuar'.");

          const listadoPromise = this.findListadoTAFrame(page, 20000, 250);
          const dilPromise = this.waitForDILAfterContinuar(page, dilBefore, 6000);

          const winner = await Promise.race([
            listadoPromise.then((fr) => ({ type: "listado", fr })),
            dilPromise.then((text) => ({ type: "dil", text })),
          ]);

          let frameListado = null;
          if (winner.type === "listado") {
            frameListado = winner.fr;
          } else {
            const dilAfter = winner.text || "";
            const dilError = this.interpretarDIL(dilAfter);
            if (dilError) throw new Error(`TA2: Validación FS (DIL): ${dilError}`);
            frameListado = await this.findListadoTAFrame(page, 15000, 250);
          }

          if (!frameListado) {
            const maybeError = await this.detectPossibleErrorInFrames(page);
            throw new Error(`TA2: no se detectó listado tras Continuar.${maybeError ? " " + maybeError : ""}`);
          }

          await this.safeScreenshot(page, pngPath);

          const popupPromise = this.waitForPopup(browser, page, 45000);

          const sel = await this.withDetachedFrameRetry(() => this.seleccionarAltaMasReciente(frameListado), {
            label: "TA2 seleccionar ALTA",
          });

          if (!sel.ok) throw new Error(`TA2: no se ha encontrado ALTA: ${sel.reason}`);

          const aFecha = this._fechaRealToA(sel.fecha);
          const pdfPath = path.join(dirDni, `TA2 ${aFecha} ${trabajador}.pdf`);

          console.log(`[DUPLICADOS][TA2][SELECCION] ${dni} -> ${sel.doc} | Fecha Real: ${sel.fecha}`);

          let popupPage = await popupPromise;
          if (!popupPage) {
            console.warn(`[DUPLICADOS][TA2][POPUP] No abrió a la primera. Reintentando click...`);
            const popupPromise2 = this.waitForPopup(browser, page, 25000);
            const sel2 = await this.withDetachedFrameRetry(() => this.seleccionarAltaMasReciente(frameListado), {
              label: "TA2 reintento seleccionar ALTA",
            });
            if (!sel2.ok) throw new Error(`TA2 reintento: no se ha encontrado ALTA: ${sel2.reason}`);
            popupPage = await popupPromise2;
          }

          if (!popupPage) throw new Error("TA2: no se abrió la pestaña del PDF tras reintento.");

          // 1er intento
          try {
            await this.descargarPdfRawViaFetchCDP(popupPage, pdfPath, 90000);
          } catch (e1) {
            if (!this._isPdfDownloadRetryableError(e1)) throw e1;
          
            console.warn(`[DUPLICADOS][TA2][PDF] Reintento por fallo de visor/PDF...`, e1?.message || e1);
          
            // Cierra popup malo
            try { await popupPage.close(); } catch (_) {}
          
            // Reabrir popup (click de nuevo sobre la fila)
            const popupPromiseR = this.waitForPopup(browser, page, 25000);
          
            const selR = await this.withDetachedFrameRetry(() => this.seleccionarAltaMasReciente(frameListado), {
              label: "TA2 reintento seleccionar ALTA (para PDF)",
            });
            if (!selR.ok) throw new Error(`TA2 reintento: no se ha encontrado ALTA: ${selR.reason}`);
          
            const popupPageR = await popupPromiseR;
            if (!popupPageR) throw new Error("TA2 reintento: no se abrió la pestaña del PDF.");
          
            // 2º intento (si falla aquí, se captura arriba y se pasa al siguiente registro)
            try {
              await this.descargarPdfRawViaFetchCDP(popupPageR, pdfPath, 90000);
            } finally {
              try { await popupPageR.close(); } catch (_) {}
            }
          }

          logsTA2.set(dni, `OK: PDF guardado -> ${path.basename(pdfPath)}`);
          okSetTA2.add(dni);

          try { await popupPage.close(); } catch (_) {}
        };

        // -------------------------
        // ✅ Ejecutar IDC para 1 empleado
        // -------------------------
        const procesarIDC = async (r, idx, total) => {
          const dni = this._dniNorm(r.dni);
          const trabajador = this._safeFileName(r.trabajador);

          if (okSetIDC.has(dni)) {
            logsIDC.set(dni, "SKIP_OK_PREVIO: IDC ya estaba OK (resume)");
            return;
          }

          const { dirDni } = await this.ensureClientDirFlat(rootOut, dni);

          console.log(`[DUPLICADOS][IDC] ${idx + 1}/${total} | DNI: ${dni} | TRABAJADOR: ${r.trabajador}`);

          await page.goto(urlFS, { waitUntil: "domcontentloaded" });
          await this.esperarLog(800);

          await openAFIOnlineReal();
          await openATR37IDC();

          const frameForm = await this.findFrameWithSelector(page, "#SDFTESNAF", 30000);
          if (!frameForm) throw new Error("IDC: no se encontró el formulario ATR37 (selector #SDFTESNAF)");

          const provNAF = this._padLeftDigits(r.provNAF, 2);
          const naf10 = this._padLeftDigits(r.naf, 10);
          const provCCC = this._padLeftDigits(r.provCCC, 2);
          const ccc9 = this._padLeftDigits(r.ccc, 9);

          await this.withDetachedFrameRetry(() => fillInput(frameForm, "#SDFTESNAF", provNAF), { label: "IDC fill PROV NAF" });
          await this.withDetachedFrameRetry(() => fillInput(frameForm, "#SDFNAF", naf10), { label: "IDC fill NAF" });
          await this.withDetachedFrameRetry(() => fillInput(frameForm, "#SDFREGCTA", regimen4), { label: "IDC fill REGIMEN" });
          await this.withDetachedFrameRetry(() => fillInput(frameForm, "#SDFTESCTA", provCCC), { label: "IDC fill PROV CCC" });
          await this.withDetachedFrameRetry(() => fillInput(frameForm, "#SDFCUENTA", ccc9), { label: "IDC fill CCC" });

          await frameForm.waitForSelector("#ListaTipoImpresion", { timeout: 30000 });
          await frameForm.select("#ListaTipoImpresion", "OnLine");

          const { text: dilBefore } = await this.readDILInAnyFrame(page);

          const clicked = await this.clickContinuarRobusta({ page, frameForm, timeoutMs: 35000 });
          if (!clicked) throw new Error("IDC: no se pudo pulsar 'Continuar'.");

          const dilAfter = await this.waitForDILAfterContinuar(page, dilBefore, 1000);
          const dilError = this.interpretarDIL(dilAfter);
          if (dilError) throw new Error(`IDC: Validación FS (DIL): ${dilError}`);

          const frameListado = await this.findListadoIDCFrame(page, 30000, 250);
          if (!frameListado) throw new Error("IDC: no se detectó el listado tras Continuar.");

          const popupPromise = this.waitForPopup(browser, page, 45000);

          const sel = await this.withDetachedFrameRetry(() => this.seleccionarAltaIDCMasReciente(frameListado), {
            label: "IDC seleccionar F. R. Alta",
          });
          if (!sel.ok) throw new Error(`IDC: no se pudo seleccionar fila: ${sel.reason}`);

          const aFecha = this._fechaRealToA(sel.fecha); // "01 01 2012" -> A010112
          const pdfPath = path.join(dirDni, `IDC ${aFecha} ${trabajador}.pdf`);

          console.log(`[DUPLICADOS][IDC][SELECCION] ${dni} -> F. R. Alta: ${sel.fecha}`);

          let popupPage = await popupPromise;
          if (!popupPage) {
            console.warn(`[DUPLICADOS][IDC][POPUP] No abrió a la primera. Reintentando doble click...`);
            const popupPromise2 = this.waitForPopup(browser, page, 25000);
            const sel2 = await this.withDetachedFrameRetry(() => this.seleccionarAltaIDCMasReciente(frameListado), {
              label: "IDC reintento seleccionar fila",
            });
            if (!sel2.ok) throw new Error(`IDC reintento: no se pudo seleccionar fila: ${sel2.reason}`);
            popupPage = await popupPromise2;
          }

          if (!popupPage) throw new Error("IDC: no se abrió la pestaña del PDF tras reintento.");

          // 1er intento
    try {
      await this.descargarPdfRawViaFetchCDP(popupPage, pdfPath, 90000);
    } catch (e1) {
      if (!this._isPdfDownloadRetryableError(e1)) throw e1;
    
      console.warn(`[DUPLICADOS][IDC][PDF] Reintento por fallo de visor/PDF...`, e1?.message || e1);
    
      try { await popupPage.close(); } catch (_) {}
    
      // Reabrir popup (doble click de nuevo)
      const popupPromiseR = this.waitForPopup(browser, page, 25000);
    
      const selR = await this.withDetachedFrameRetry(() => this.seleccionarAltaIDCMasReciente(frameListado), {
        label: "IDC reintento seleccionar F. R. Alta (para PDF)",
      });
      if (!selR.ok) throw new Error(`IDC reintento: no se pudo seleccionar fila: ${selR.reason}`);
    
      const popupPageR = await popupPromiseR;
      if (!popupPageR) throw new Error("IDC reintento: no se abrió la pestaña del PDF.");
    
      try {
        await this.descargarPdfRawViaFetchCDP(popupPageR, pdfPath, 90000);
      } finally {
        try { await popupPageR.close(); } catch (_) {}
      }
    }
    

          logsIDC.set(dni, `OK: PDF guardado -> ${path.basename(pdfPath)}`);
          okSetIDC.add(dni);

          try { await popupPage.close(); } catch (_) {}
        };

        // -------------------------
        // ✅ LOOP principal: por empleado -> TA2 y luego IDC
        // -------------------------
        let okTA2 = 0, errTA2 = 0;
        let okIDC = 0, errIDC = 0;

        for (let i = 0; i < kept.length; i++) {
          registrosProcesados++;
          const r = kept[i];
          const dni = this._dniNorm(r.dni);

          // 1) TA2
          try {
            await procesarTA2(r, i, kept.length);
            if ((logsTA2.get(dni) || "").startsWith("OK")) okTA2++;
          } catch (e) {
            errTA2++;
            const key = dni || `ROW_${r._row}`;
            const msg = `ERROR: ${e?.message || e}`;
            logsTA2.set(key, msg);
            console.warn("[DUPLICADOS][TA2]", msg);
          }

          // 2) IDC (siempre después)
          try {
            await procesarIDC(r, i, kept.length);
            if ((logsIDC.get(dni) || "").startsWith("OK")) okIDC++;
          } catch (e) {
            errIDC++;
            const key = dni || `ROW_${r._row}`;
            const msg = `ERROR: ${e?.message || e}`;
            logsIDC.set(key, msg);
            console.warn("[DUPLICADOS][IDC]", msg);
          }

          if ((i + 1) % 5 === 0) {
            flushLogs(detalleTA2Path, logsTA2);
            flushLogs(detalleIDCPath, logsIDC);
          }
        }

        flushLogs(detalleTA2Path, logsTA2);
        flushLogs(detalleIDCPath, logsIDC);

        console.log(`[DUPLICADOS] Terminado.`);
        console.log(`[DUPLICADOS][TA2] OK: ${okTA2} | ERROR: ${errTA2}`);
        console.log(`[DUPLICADOS][IDC] OK: ${okIDC} | ERROR: ${errIDC}`);
        console.log(`[DUPLICADOS] Procesados: ${registrosProcesados}`);

        registrarEjecucion({ nombreProceso, registrosProcesados });

        try { if (browser) await browser.close(); } catch (_) {}

        return resolve(true);
      } catch (err) {
        console.error("[DUPLICADOS] Error general:", err?.message || err);
        try {
          if (globalThis?.mainProcess?.mostrarError) {
            await globalThis.mainProcess.mostrarError(
              "No se ha podido completar el proceso",
              "Se ha producido un error interno ejecutando DUPLICADOS TA2+IDC.",
            );
          }
        } catch (_) {}
        try { if (browser) await browser.close(); } catch (_) {}

        return resolve(false);
      }
    });
  }

  // ✅ Alias que AppService ya llama: ahora hace TA2+IDC
  async ["dUPLICADOSTA2+IDC"](argumentos) {
    console.warn("[DUPLICADOS] Alias dUPLICADOSTA2+IDC() llamado. Ejecuta TA2+IDC (por empleado).");
    return this.duplicadosTa2(argumentos);
  }
}

module.exports = ProcesosDuplicados;
