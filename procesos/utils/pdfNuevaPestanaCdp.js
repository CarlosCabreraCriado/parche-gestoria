/**
 * Guardar PDF en nueva pestaña usando CDP + Fetch (Response interception)
 * Robusto para portales que abren visor o HTML y el PDF real va por request interna.
 */

const fs = require("fs");

function sleep(ms) {
  return new Promise((r) => setTimeout(r, ms));
}

/**
 * Espera una nueva pestaña (popup) originada tras una acción.
 * - parentPage: page desde la que se dispara la nueva pestaña
 */
async function waitForPopup(browser, parentPage, timeoutMs = 30000) {
  const parentTargetId = parentPage.target()._targetId; // interno, pero práctico
  const start = Date.now();

  while (Date.now() - start < timeoutMs) {
    const targets = browser.targets();
    // Buscamos un target "page" distinto al parent
    const pageTargets = targets.filter(
      (t) => t.type() === "page" && t._targetId !== parentTargetId,
    );

    // A veces se crean varias páginas; nos quedamos con la más reciente
    if (pageTargets.length) {
      const t = pageTargets[pageTargets.length - 1];
      const p = await t.page().catch(() => null);
      if (p) return p;
    }

    await sleep(250);
  }

  throw new Error("Timeout esperando nueva pestaña/popup del PDF.");
}

/**
 * Descarga un PDF real desde una pestaña usando CDP Fetch (requestStage: Response)
 * - popupPage: la pestaña nueva
 * - outPath: ruta de salida del PDF
 * - timeoutMs: timeout duro
 * - options:
 *    - fetchPatterns: [{urlPattern, requestStage}]
 *    - forceReload: recarga la pestaña para forzar la request del PDF
 */
async function descargarPdfRawViaFetchCDP(
  popupPage,
  outPath,
  timeoutMs = 90000,
  options = {},
) {
  const {
    fetchPatterns = [{ urlPattern: "*", requestStage: "Response" }],
    forceReload = true,
  } = options;

  const client = await popupPage.target().createCDPSession();

  let done = false;
  let lastError = "";
  let timer;

  try {
    // Activar Fetch
    await client.send("Fetch.enable", { patterns: fetchPatterns });

    // Promesa que resuelve cuando detecta y guarda el PDF real
    const pdfPromise = new Promise((resolve, reject) => {
      timer = setTimeout(() => {
        reject(
          new Error(
            "Timeout descargando PDF (CDP/Fetch). " + (lastError || ""),
          ),
        );
      }, timeoutMs);

      client.on("Fetch.requestPaused", async (ev) => {
        if (done) {
          // Continuar para no bloquear
          try {
            await client.send("Fetch.continueRequest", {
              requestId: ev.requestId,
            });
          } catch (_) {}
          return;
        }

        try {
          // Solo nos interesa cuando hay Response disponible
          // (en algunos casos puede venir sin responseStatus)
          const status = ev.responseStatusCode ?? ev.responseStatus; // ✅ compat
          const hasResponse = typeof status === "number";

          if (!hasResponse) {
            await client.send("Fetch.continueRequest", {
              requestId: ev.requestId,
            });
            return;
          }

          // Heurística rápida por headers (content-type)
          const headers = {};
          for (const h of ev.responseHeaders || []) {
            headers[String(h.name || "").toLowerCase()] = String(h.value || "");
          }
          const ctype = (headers["content-type"] || "").toLowerCase();

          // Intentamos leer body SI parece pdf o si no sabemos (a veces viene mal content-type)
          const shouldTryBody =
            ctype.includes("application/pdf") ||
            ctype.includes("octet-stream") ||
            ctype.includes("binary") ||
            true; // mantenemos robustez, porque algunos portales devuelven html con redirección

          if (!shouldTryBody) {
            await client.send("Fetch.continueRequest", {
              requestId: ev.requestId,
            });
            return;
          }

          const bodyResp = await client.send("Fetch.getResponseBody", {
            requestId: ev.requestId,
          });
          const buf = bodyResp.base64Encoded
            ? Buffer.from(bodyResp.body, "base64")
            : Buffer.from(bodyResp.body, "utf8");

          // Validación PDF real: cabecera %PDF
          const head = buf.slice(0, 4).toString("utf8");
          const isPdf = head === "%PDF";

          if (isPdf) {
            fs.writeFileSync(outPath, buf);
            done = true;
            clearTimeout(timer);

            // IMPORTANTÍSIMO: continuar request para no “colgar” la sesión
            try {
              await client.send("Fetch.continueRequest", {
                requestId: ev.requestId,
              });
            } catch (_) {}

            resolve(true);
            return;
          }

          // No era PDF -> continuar
          await client.send("Fetch.continueRequest", {
            requestId: ev.requestId,
          });
        } catch (e) {
          lastError = e.message || String(e);
          // Continuar para no bloquear
          try {
            await client.send("Fetch.continueRequest", {
              requestId: ev.requestId,
            });
          } catch (_) {}
        }
      });
    });

    // Forzar que el portal haga la request del PDF
    if (forceReload) {
      try {
        await popupPage.bringToFront();
        // Espera mínima para que cargue algo
        await popupPage.waitForTimeout(800);
        await popupPage
          .reload({ waitUntil: "domcontentloaded" })
          .catch(() => {});
      } catch (_) {}
    }

    await pdfPromise;
    return true;
  } finally {
    clearTimeout(timer);
    try {
      await client.send("Fetch.disable");
    } catch (_) {}
    try {
      await client.detach();
    } catch (_) {}
  }
}

/**
 * Wrapper con reintento:
 * - openPdfFn: dispara la acción (click imprimir)
 * - getPopupFn: obtiene la pestaña nueva
 * - downloadFn: descarga el pdf
 * - closePopupFn: cierre seguro
 */
async function descargarPdfConReintento({
  label = "PDF",
  openPdfFn,
  getPopupFn,
  downloadFn,
  closePopupFn,
  reintentos = 2,
  esperaEntre = 1200,
}) {
  let lastErr = "";

  for (let i = 1; i <= reintentos; i++) {
    try {
      await openPdfFn();
      const popup = await getPopupFn();
      await downloadFn(popup);
      if (closePopupFn) await closePopupFn(popup);
      return true;
    } catch (e) {
      lastErr = e.message || String(e);
      // Intento de limpieza: cerrar popup si existe
      try {
        // closePopupFn puede fallar si popup no existe
      } catch (_) {}
      if (i < reintentos) await sleep(esperaEntre);
    }
  }

  throw new Error(
    `[${label}] No se pudo descargar PDF tras ${reintentos} intentos. Último error: ${lastErr}`,
  );
}

async function waitForPrintPreviewPopup(
  browser,
  openerPage,
  timeoutMs = 30000,
) {
  const start = Date.now();

  while (Date.now() - start < timeoutMs) {
    const targets = browser.targets();

    // Busca cualquier page con chrome://print (no solo opener exacto)
    const t = targets.find(
      (x) =>
        x.type() === "page" && (x.url() || "").startsWith("chrome://print"),
    );

    if (t) {
      const p = await t.page().catch(() => null);
      if (p) {
        await p.bringToFront().catch(() => {});
        return p;
      }
    }

    await new Promise((r) => setTimeout(r, 250));
  }

  throw new Error("Timeout esperando popup chrome://print.");
}

async function descargarPdfDesdePrintPreview(popupPage, outputPath, opts = {}) {
  const { timeoutMs = 20000, urlPattern = "*print.pdf*", debug = false } = opts;

  const fs = require("fs");

  const client = await popupPage.target().createCDPSession();
  let saved = false;

  await client.send("Fetch.enable", {
    patterns: [{ urlPattern, requestStage: "Response" }],
  });

  client.on("Fetch.requestPaused", async (ev) => {
    try {
      const status = ev.responseStatusCode ?? ev.responseStatus;
      const hasResponse = typeof status === "number";
      const url = ev.request?.url || "";

      if (!hasResponse) {
        await client.send("Fetch.continueRequest", { requestId: ev.requestId });
        return;
      }

      const headers = (ev.responseHeaders || []).reduce((acc, h) => {
        acc[String(h.name || "").toLowerCase()] = String(h.value || "");
        return acc;
      }, {});
      const ct = headers["content-type"] || "";

      const parecePdf =
        url.includes("print.pdf") ||
        ct.includes("application/pdf") ||
        ct.includes("octet-stream");

      if (!parecePdf) {
        await client.send("Fetch.continueRequest", { requestId: ev.requestId });
        return;
      }

      const body = await client.send("Fetch.getResponseBody", {
        requestId: ev.requestId,
      });
      const buff = body.base64Encoded
        ? Buffer.from(body.body, "base64")
        : Buffer.from(body.body, "utf8");

      if (!buff.slice(0, 5).toString("utf8").startsWith("%PDF-")) {
        if (debug)
          console.warn("[PRINT] Capturado pero no parece PDF. Se ignora.");
        await client.send("Fetch.continueRequest", { requestId: ev.requestId });
        return;
      }

      fs.writeFileSync(outputPath, buff);
      saved = true;

      await client.send("Fetch.continueRequest", { requestId: ev.requestId });
    } catch (e) {
      try {
        await client.send("Fetch.continueRequest", { requestId: ev.requestId });
      } catch {}
    }
  });

  // Fuerza a que el print preview cargue el recurso print.pdf
  await popupPage.reload({ waitUntil: "domcontentloaded" });

  const start = Date.now();
  while (!saved && Date.now() - start < timeoutMs) {
    await new Promise((r) => setTimeout(r, 200));
  }

  await client.send("Fetch.disable");
  try {
    await client.detach();
  } catch (_) {}

  if (!saved)
    throw new Error(
      "No se pudo capturar print.pdf desde chrome://print (timeout).",
    );
}



module.exports = {
  waitForPopup,
  descargarPdfRawViaFetchCDP,
  descargarPdfConReintento,
  waitForPrintPreviewPopup,
  descargarPdfDesdePrintPreview,
};
