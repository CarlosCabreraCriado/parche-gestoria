const path = require("path");
const fs = require("fs");
const XlsxPopulate = require("xlsx-populate");
const { DateTime } = require("luxon");

const { registrarEjecucion, agruparPorEmpresa } = require("../metricas");
const puppeteer = require("puppeteer");

class ProcesosCertificados {
  constructor(pathToDbFolder, nombreProyecto, proyectoDB) {
    this.pathToDbFolder = pathToDbFolder;
    this.nombreProyecto = nombreProyecto;
    this.proyectoDB = proyectoDB;
  }

  async esperar(tiempo) {
    return new Promise((resolve) => {
      setTimeout(resolve, tiempo);
    });
  }

  async certificadosDeEstarAlCorriente(argumentos) {
    return new Promise((resolve) => {
      console.log("Certificados unificados — iniciando");
      const nombreProceso = "Certificados Unificados";
      let registrosProcesados = 0;

      const chromiumExecutablePath = path.normalize(
        argumentos.formularioControl[0],
      );
      const pathArchivoEtiquetas = argumentos.formularioControl[1];
      const pathBase = argumentos.formularioControl[2];
      const modoManual = !!argumentos.formularioControl[3];
      const codigosEmpresaInput = argumentos.formularioControl[4];

      let runSS = !!argumentos.formularioControl[5];
      let runTrib = !!argumentos.formularioControl[6];
      let runATC = !!argumentos.formularioControl[7];
      let runITA = !!argumentos.formularioControl[8];
      let runArt42 = !!argumentos.formularioControl[9];

      const empresaAutRegimen = String(argumentos.formularioControl[10] || "");
      const empresaAutTesoreria = String(
        argumentos.formularioControl[11] || "",
      );
      const empresaAutCuenta = String(argumentos.formularioControl[12] || "");

      console.log(
        `[MODO] ${modoManual ? "Manual (form-driven)" : "Automático (Excel-driven)"}`,
      );

      if (modoManual && !runSS && !runTrib && !runATC && !runITA && !runArt42) {
        console.log(
          "No se ha seleccionado ningún certificado. Nada que hacer.",
        );
        return resolve(false);
      }

      const parsearCodigos = (input) =>
        new Set(
          String(input || "")
            .split(/[,;\-\s]+/)
            .map((t) => t.replace(/\D/g, ""))
            .filter((t) => t !== "")
            .map((t) => t.padStart(4, "0")),
        );
      const codigosEmpresaObjetivo = parsearCodigos(codigosEmpresaInput);
      if (codigosEmpresaObjetivo.size === 0) {
        console.log(
          "No se ha especificado ningún código de empresa. Se procesarán todos los expedientes.",
        );
      }

      const fechaEjecucion = DateTime.now()
        .setZone("Europe/Madrid")
        .toFormat("dd-MM-yyyy");
      const carpetaFecha = `Certificados de estar al corriente (${fechaEjecucion})`;

      const carpetaRaiz = path.join(path.normalize(pathBase), carpetaFecha);

      try {
        XlsxPopulate.fromFileAsync(path.normalize(pathArchivoEtiquetas))
          .then(async (workbook) => {
            console.log("Excel cargado (certificados unificados)");
            const archivo = workbook;
            const hoja = archivo.sheet("BASE DE DATOS (NO TOCAR)");
            if (!hoja) {
              console.warn(
                "[CERT] Hoja 'BASE DE DATOS (NO TOCAR)' no encontrada en el Excel.",
              );
              return resolve(false);
            }
            const columnas = hoja.usedRange()._numColumns;
            const filas = hoja.usedRange()._numRows;


            const cabeceras = [];
            for (let i = 1; i <= columnas; i++) {
              cabeceras.push(hoja.cell(1, i).value());
            }
            console.log("Cabeceras: " + cabeceras);

            // Mapa dinámico: nombre de cabecera → índice de columna (1-based)
            const colIdx = {};
            cabeceras.forEach((h, i) => {
              if (h !== undefined && h !== null) {
                colIdx[String(h).trim()] = i + 1;
              }
            });

            // Añadir columnas LOG al final, de forma dinámica
            const addLogCol = (nombre) => {
              const nextCol = Object.keys(colIdx).length + 1;
              hoja.cell(1, nextCol).value(nombre);
              colIdx[nombre] = nextCol;
            };
            addLogCol("LOG SS");
            addLogCol("LOG TRIB");
            addLogCol("LOG ATC");
            addLogCol("LOG ITA");
            addLogCol("LOG ART42");
            console.log("Índices de columnas dinámicos:", colIdx);

            let clientes = [];
            for (let i = 2; i <= filas; i++) {
              const objetoCliente = {
                filaExcel: i,
                errores: [],
                flagDupeNIF: false,
                flagSS: false,
                flagAEAT: false,
                flagATC: false,
                flagITA: false,
              };
              for (let j = 1; j <= columnas; j++) {
                const cellVal = hoja.cell(i, j).value();
                if (cellVal !== undefined) {
                  switch (cabeceras[j - 1]) {
                    case "Código Cuenta Cotización (CCC)":
                      objetoCliente.ccc = cellVal;
                      const c = String(cellVal);
                      objetoCliente.ccc1 = c.substring(0, 4);
                      objetoCliente.ccc2 = c.substring(4, 6);
                      objetoCliente.ccc3 = c.substring(6);
                      break;
                    case "EMPRESA":
                      objetoCliente.empresa = cellVal;
                      break;
                    case "Expediente":
                      objetoCliente.codigo = cellVal;
                      break;
                    case "NIF":
                      objetoCliente.nif = cellVal;
                      break;
                    case "SS":
                      objetoCliente.flagSS =
                        String(cellVal || "")
                          .trim()
                          .toLowerCase() === "x";
                      break;
                    case "AEAT":
                      objetoCliente.flagAEAT =
                        String(cellVal || "")
                          .trim()
                          .toLowerCase() === "x";
                      break;
                    case "ATC":
                      objetoCliente.flagATC =
                        String(cellVal || "")
                          .trim()
                          .toLowerCase() === "x";
                      break;
                    case "ITA":
                      objetoCliente.flagITA =
                        String(cellVal || "")
                          .trim()
                          .toLowerCase() === "x";
                      break;
                  }
                }
              }

              const codigoNormalizado = String(objetoCliente.codigo || "")
                .replace(/\D/g, "")
                .padStart(4, "0");

              const debeProcesarseManual =
                codigoNormalizado !== "" &&
                (codigosEmpresaObjetivo.size === 0 ||
                  codigosEmpresaObjetivo.has(codigoNormalizado));

              const tieneAlgunFlag =
                objetoCliente.flagSS ||
                objetoCliente.flagAEAT ||
                objetoCliente.flagATC ||
                objetoCliente.flagITA;

              const debeProcesarse = modoManual
                ? debeProcesarseManual
                : codigoNormalizado !== "" && tieneAlgunFlag;

              if (
                debeProcesarse &&
                objetoCliente.ccc !== "" &&
                objetoCliente.ccc !== null &&
                objetoCliente.ccc !== undefined
              ) {
                const fechaHoy = DateTime.now()
                  .setZone("Europe/Madrid")
                  .toFormat("ddMMyy");
                objetoCliente.nombreArchivoSS =
                  objetoCliente.codigo +
                  " CERTIFICADO ESTAR AL CORRIENTE SS " +
                  objetoCliente.empresa +
                  " " +
                  fechaHoy +
                  ".pdf";
                objetoCliente.nombreArchivoTrib =
                  objetoCliente.codigo +
                  " CERTIFICADO ESTAR AL CORRIENTE AEAT " +
                  objetoCliente.empresa +
                  " " +
                  fechaHoy +
                  ".pdf";
                objetoCliente.nombreArchivoATC =
                  objetoCliente.codigo +
                  " CERTIFICADO ESTAR AL CORRIENTE ATC " +
                  objetoCliente.empresa +
                  " " +
                  fechaHoy +
                  ".pdf";
                objetoCliente.nombreArchivoITA = `${objetoCliente.codigo} CERTIFICADO ESTAR AL CORRIENTE ITA ${objetoCliente.empresa} ${fechaHoy}.pdf`;
                objetoCliente.nombreArchivoArt42 = `${objetoCliente.codigo} CERTIFICADO ESTAR AL CORRIENTE ART42 ${objetoCliente.empresa} ${fechaHoy}.png`;
                clientes.push(Object.assign({}, objetoCliente));
              }
            }

            if (!modoManual) {
              runSS = clientes.some((c) => c.flagSS);
              runTrib = clientes.some((c) => c.flagAEAT);
              runATC = clientes.some((c) => c.flagATC);
              runITA = clientes.some((c) => c.flagITA);
              runArt42 = false;

              if (!runSS && !runTrib && !runATC && !runITA) {
                console.log(
                  "No hay empresas con certificados marcados en el Excel. Nada que hacer.",
                );
                return resolve(false);
              }
              console.log(
                `[AUTO] Procesos requeridos: SS=${runSS}, TRIB=${runTrib}, ATC=${runATC}, ITA=${runITA}`,
              );
            }

            const paths = {};
            if (runSS)
              paths.ss = { excel: carpetaRaiz, resultados: carpetaRaiz };
            if (runTrib)
              paths.trib = { excel: carpetaRaiz, resultados: carpetaRaiz };
            if (runATC)
              paths.atc = { excel: carpetaRaiz, resultados: carpetaRaiz };
            if (runITA)
              paths.ita = { excel: carpetaRaiz, resultados: carpetaRaiz };
            if (runArt42)
              paths.art42 = { excel: carpetaRaiz, resultados: carpetaRaiz };
            if (!fs.existsSync(carpetaRaiz)) {
              fs.mkdirSync(carpetaRaiz, { recursive: true });
              console.log(`Carpeta creada: ${carpetaRaiz}`);
            } else {
              console.log(`La carpeta ya existe: ${carpetaRaiz}`);
            }

            const downloadPathInicial = carpetaRaiz;

            if (runTrib || runATC) {
              const vistos = new Set();
              clientes = clientes.map((obj) => {
                const nifKey = String(obj.nif || "").trim();
                if (nifKey && vistos.has(nifKey)) {
                  return { ...obj, flagDupeNIF: true };
                }
                if (nifKey) vistos.add(nifKey);
                return obj;
              });
            }

            console.log(`Clientes a procesar: ${clientes.length}`);
            console.log("Clientes: ");
            console.log(clientes);

            let browser;
            try {
              browser = await puppeteer.launch({
                executablePath: chromiumExecutablePath,
                headless: false,
              });
              console.log("[CERT] Navegador iniciado:", chromiumExecutablePath);
            } catch (e) {
              console.warn("[CERT] Error lanzando Chromium:", e?.message || e);
              return resolve(false);
            }

            const prepararPagina = async (pageObj) => {
              pageObj.on("dialog", async (dialog) => {
                const tipo = dialog.type();
                if (tipo === "beforeunload") {
                  try {
                    await dialog.accept();
                  } catch (_) {}
                }
              });
              await pageObj._client().send("Page.setDownloadBehavior", {
                behavior: "allow",
                downloadPath: downloadPathInicial,
              });
              await pageObj.setViewport({ width: 1080, height: 1024 });
              pageObj.setDefaultTimeout(60000);
            };

            let page = await browser.newPage();
            await prepararPagina(page);

            try {
              await this._preinicializarCertificados({
                browser,
                page,
                runSS,
                runTrib,
                runATC,
                runArt42,
              });
            } catch (e) {
              console.warn(
                "[CERT INIT] Error en pre-inicialización:",
                e?.message || e,
              );
              try {
                await browser.close();
              } catch (_) {}
              return resolve(false);
            }

            for (let i = 0; i < clientes.length; i++) {
              registrosProcesados += 1;

              if (i % 10 === 0 && i > 0) {
                console.log("[CERT] Reciclando página en iteración", i);
                try {
                  await page.close();
                } catch (_) {}
                page = await browser.newPage();
                await prepararPagina(page);
              }

              console.log("Procesando cliente: " + i);
              console.log(clientes[i]);

              const clientRunSS = modoManual ? runSS : clientes[i].flagSS;
              const clientRunTrib = modoManual ? runTrib : clientes[i].flagAEAT;
              const clientRunATC = modoManual ? runATC : clientes[i].flagATC;
              const clientRunITA = modoManual ? runITA : clientes[i].flagITA;
              const clientRunArt42 = modoManual && runArt42;

              if (
                clientes[i].ccc === "" ||
                clientes[i].ccc === null ||
                clientes[i].ccc === undefined
              ) {
                clientes[i].errores = ["Campo CCC no definidos."];

                if (clientRunSS)
                  hoja
                    .cell(clientes[i].filaExcel, colIdx["LOG SS"])
                    .value("ERROR: Campo CCC no definido.");
                if (clientRunTrib)
                  hoja
                    .cell(clientes[i].filaExcel, colIdx["LOG TRIB"])
                    .value("ERROR: Campo CCC no definido.");
                if (clientRunATC)
                  hoja
                    .cell(clientes[i].filaExcel, colIdx["LOG ATC"])
                    .value("ERROR: Campo CCC no definido.");
                if (clientRunITA)
                  hoja
                    .cell(clientes[i].filaExcel, colIdx["LOG ITA"])
                    .value("ERROR: Campo CCC no definido.");
                if (clientRunArt42)
                  hoja
                    .cell(clientes[i].filaExcel, colIdx["LOG ART42"])
                    .value("ERROR: Campo CCC no definido.");
                continue;
              }

              if (clientRunSS) {
                try {
                  await this._procesarCertificadoSS({
                    browser,
                    page,
                    cliente: clientes[i],
                    paths: paths.ss,
                    hoja,
                    colIdx,
                  });
                } catch (e) {
                  const msg = String(e?.message || e);
                  console.warn("[CERT SS] Error:", msg);
                  hoja.cell(clientes[i].filaExcel, colIdx["LOG SS"]).value("ERROR: " + msg);
                  try {
                    await page.goto("about:blank");
                  } catch (_) {}
                }
              }

              if (clientRunTrib) {
                try {
                  await this._procesarCertificadoAEAT({
                    browser,
                    page,
                    cliente: clientes[i],
                    paths: paths.trib,
                    hoja,
                    colIdx,
                  });
                } catch (e) {
                  const msg = String(e?.message || e);
                  console.warn("[CERT TRIB] Error:", msg);
                  hoja.cell(clientes[i].filaExcel, colIdx["LOG TRIB"]).value("ERROR: " + msg);
                  try {
                    await page.goto("about:blank");
                  } catch (_) {}
                }
              }

              if (clientRunATC) {
                try {
                  await this._procesarCertificadoATC({
                    browser,
                    page,
                    cliente: clientes[i],
                    paths: paths.atc,
                    hoja,
                    colIdx,
                  });
                } catch (e) {
                  const msg = String(e?.message || e);
                  console.warn("[CERT ATC] Error:", msg);
                  hoja.cell(clientes[i].filaExcel, colIdx["LOG ATC"]).value("ERROR: " + msg);
                  try {
                    await page.goto("about:blank");
                  } catch (_) {}
                }
              }

              if (clientRunITA) {
                try {
                  await this._procesarInformeITA({
                    browser,
                    page,
                    cliente: clientes[i],
                    paths: paths.ita,
                    hoja,
                    colIdx,
                  });
                } catch (e) {
                  const msg = String(e?.message || e);
                  console.warn("[ITA] Error:", msg);
                  hoja.cell(clientes[i].filaExcel, colIdx["LOG ITA"]).value("ERROR: " + msg);
                  try {
                    await page.goto("about:blank");
                  } catch (_) {}
                }
              }

              if (clientRunArt42) {
                try {
                  await this._procesarCertificadoArt42({
                    browser,
                    page,
                    cliente: clientes[i],
                    paths: paths.art42,
                    hoja,
                    colIdx,
                    empresaAutRegimen,
                    empresaAutTesoreria,
                    empresaAutCuenta,
                  });
                } catch (e) {
                  const msg = String(e?.message || e);
                  console.warn("[ART42] Error:", msg);
                  hoja.cell(clientes[i].filaExcel, colIdx["LOG ART42"]).value("ERROR: " + msg);
                  try {
                    await page.goto("about:blank");
                  } catch (_) {}
                }
              }

              console.log("Nuevo cliente");
              await this.esperar(1000);
            }

            try {
              await browser.close();
            } catch (_) {}

            const excelOutBase = runSS
              ? paths.ss.excel
              : runTrib
                ? paths.trib.excel
                : runATC
                  ? paths.atc.excel
                  : runITA
                    ? paths.ita.excel
                    : paths.art42.excel;
            console.log("Escribiendo archivo...");
            console.log("Path: " + path.normalize(excelOutBase));
            try {
              await archivo.toFileAsync(
                path.normalize(
                  path.join(excelOutBase, "Certificados-Procesado.xlsx"),
                ),
              );
              console.log("XLSX escrito correctamente");
            } catch (err) {
              console.log("Error escribiendo XLSX:", err?.message || err);
              return resolve(false);
            }

            try {
              registrarEjecucion({
                nombreProceso,
                registrosProcesados: registrosProcesados,
                empresas: agruparPorEmpresa(clientes),
              });
            } catch (_) {}
            console.log("Fin del procesamiento (certificados unificados)");
            resolve(true);
          })
          .then(() => {})
          .catch((err) => {
            console.log("ERROR: ", err?.message || err);
            resolve(false);
          });
      } catch (err) {
        console.log(
          "Se ha producido un error interno cargando los archivos:",
          err?.message || err,
        );
        resolve(false);
      }
    }).catch((err) => {
      console.log("Se ha producido un error interno: ", err?.message || err);
      return false;
    });
  }

  async _preinicializarCertificados({
    browser,
    page,
    runSS,
    runTrib,
    runATC,
    runArt42,
  }) {
    console.log(
      "[CERT INIT] Iniciando pre-selección de certificados digitales...",
    );

    if (runSS) {
      console.log("[CERT INIT] SS — navegando para seleccionar certificado...");
      for (let intento = 1; intento <= 2; intento++) {
        try {
          await page.goto(
            "https://w2.seg-social.es/ProsaInternet/OnlineAccess?ARQ.SPM.ACTION=LOGIN&ARQ.SPM.APPTYPE=SERVICE&ARQ.IDAPP=XV21F001",
            { waitUntil: "networkidle0" },
          );
          break;
        } catch (e) {
          if (intento === 2) throw e;
          await this.esperar(1500);
        }
      }
      console.log("[CERT INIT] SS listo.");
    }

    if (runTrib) {
      console.log(
        "[CERT INIT] TRIB — navegando para seleccionar certificado...",
      );
      for (let intento = 1; intento <= 2; intento++) {
        try {
          await page.goto(
            "https://www1.agenciatributaria.gob.es/wlpl/EMCE-JDIT/ECOTInternetCiudadanosServlet",
            { waitUntil: "networkidle0" },
          );
          break;
        } catch (e) {
          if (intento === 2) throw e;
          await this.esperar(1500);
        }
      }
      try {
        const botonModal = await page.waitForSelector(
          'button[data-dismiss="modal"]',
          { timeout: 1000 },
        );
        if (botonModal) await botonModal.click();
      } catch (_) {}
      console.log("[CERT INIT] TRIB listo.");
    }

    if (runATC) {
      console.log(
        "[CERT INIT] ATC — navegando para seleccionar certificado...",
      );
      for (let intento = 1; intento <= 2; intento++) {
        try {
          await page.goto(
            "https://sede.gobiernodecanarias.org/tributos/ov/seguro/certificados/individual/listado.jsp",
            { waitUntil: "networkidle0" },
          );
          break;
        } catch (e) {
          if (intento === 2) throw e;
          await this.esperar(1500);
        }
      }
      await this.esperar(1000);

      try {
        await page.waitForSelector(
          'img[alt="img_dig1"], img[src*="certificadoDigital"]',
          { timeout: 3000 },
        );
        await page.evaluate(() => {
          const img =
            document.querySelector('img[alt="img_dig1"]') ||
            document.querySelector('img[src*="certificadoDigital"]');
          if (img?.parentElement?.tagName === "A") img.parentElement.click();
        });
        await page
          .waitForNavigation({ waitUntil: "networkidle0", timeout: 10000 })
          .catch(() => {});
        await this.esperar(1000);
      } catch (_) {}

      if (page.url().includes("/publico/validacion/")) {
        try {
          const botonEntrar = await page.waitForSelector(
            'input[id="btnValidar"]',
            { timeout: 5000 },
          );
          if (botonEntrar) await botonEntrar.click();
        } catch (_) {}

        try {
          await page.waitForFunction(
            () => !window.location.href.includes("/publico/validacion/"),
            { timeout: 120000 },
          );
        } catch (_) {
          throw new Error(
            "Tiempo de autenticación ATC agotado en la fase de inicialización.",
          );
        }
        await this.esperar(2000);
      }
      console.log("[CERT INIT] ATC listo.");
    }

    if (runArt42) {
      console.log(
        "[CERT INIT] ART42 — navegando para seleccionar certificado digital...",
      );
      for (let intento = 1; intento <= 2; intento++) {
        try {
          await page.goto("https://w2.seg-social.es/fs/indexframes.html", {
            waitUntil: "networkidle0",
          });
          break;
        } catch (e) {
          if (intento === 2) throw e;
          await this.esperar(1500);
        }
      }
      console.log("[CERT INIT] ART42 listo.");
    }

    console.log(
      "[CERT INIT] Todos los certificados pre-seleccionados. Iniciando procesamiento de clientes...",
    );
  }

  async _descargarPDF({
    browser,
    botonClick,
    rutaArchivo,
    etiqueta,
    timeoutMs = 15000,
  }) {
    let resuelto = false;
    let timeoutId = null;

    const resultado = await new Promise((resolve) => {
      const finalizar = (valor) => {
        if (resuelto) return;
        resuelto = true;
        clearTimeout(timeoutId);
        browser.off("targetcreated", onTargetCreated);
        resolve(valor);
      };

      const onTargetCreated = async (target) => {
        if (resuelto) return;
        try {
          const newPage = await target.page();
          if (!newPage) return;
          newPage.on("response", async (response) => {
            if (resuelto) return;
            const contentType = response.headers()["content-type"] || "";
            if (
              response.url().startsWith("chrome-extension://") &&
              contentType.includes("application/pdf")
            ) {
              console.log(`PDF detectado (${etiqueta}):`, response.url());
              const pdfBuffer = await response.buffer();
              fs.writeFileSync(rutaArchivo, pdfBuffer);
              console.log(`PDF ${etiqueta} descargado en:`, rutaArchivo);
              finalizar(newPage);
            }
          });
        } catch (_) {}
      };

      browser.on("targetcreated", onTargetCreated);
      timeoutId = setTimeout(() => finalizar(false), timeoutMs);

      botonClick();
    });

    return resultado;
  }

  async _procesarInformeITA({ browser, page, cliente, paths, hoja, colIdx }) {
    console.log(
      `[ITA] Iniciando para cliente: ${cliente.codigo} - ${cliente.ccc}`,
    );
    const filePath = path.join(paths.resultados, cliente.nombreArchivoITA);

    await page.goto(
      "https://w2.seg-social.es/Xhtml?JacadaApplicationName=SGIRED&TRANSACCION=ATR64&E=I&AP=AFIR",
      { waitUntil: "networkidle0" },
    );
    await this.esperar(1000);

    await page.locator('input[name="txt_SDFREG62_ayuda"]').wait();
    await page.type('input[name="txt_SDFREG62_ayuda"]', String(cliente.ccc1));
    await page.locator('input[name="txt_SDFTESO62"]').wait();
    await page.type('input[name="txt_SDFTESO62"]', String(cliente.ccc2));
    await page.locator('input[name="txt_SDFNUM62"]').wait();
    await page.type('input[name="txt_SDFNUM62"]', String(cliente.ccc3));

    await this.esperar(1000);
    await page.select('select[name="cbo_ListaTipoImpresion"]', "OnLine");
    await this.esperar(1000);

    let tabITA;
    try {
      tabITA = await new Promise(async (resolvePromise) => {
        let resuelto = false;
        let timeoutId = null;

        const finalizar = (resultado) => {
          if (resuelto) return;
          resuelto = true;
          if (timeoutId) clearTimeout(timeoutId);
          browser.off("targetcreated", onTargetCreated);
          resolvePromise(resultado);
        };

        const onTargetCreated = async (target) => {
          if (resuelto) return;
          try {
            const newPage = await target.page();
            if (!newPage) return;
            newPage.on("response", async (response) => {
              if (resuelto) return;
              if (
                !response.url().endsWith(".js") &&
                !response.url().endsWith(".css") &&
                response.url().startsWith("chrome-extension://")
              ) {
                console.log("[ITA] PDF detectado:", response.url());
                const pdfBuffer = await response.buffer();
                fs.writeFileSync(filePath, pdfBuffer);
                console.log("[ITA] PDF descargado en:", filePath);
                finalizar(newPage);
              }
            });
          } catch (_) {}
        };

        browser.on("targetcreated", onTargetCreated);
        timeoutId = setTimeout(() => finalizar(false), 15000);

        await page.locator('input[name="btn_Sub2207601004"]').wait();
        await page.locator('input[name="btn_Sub2207601004"]').click();
      });
    } catch (e) {
      console.log("[ITA] Error en descarga de PDF:", e);
    }

    if (!tabITA && fs.existsSync(filePath)) tabITA = true;

    let descargaOk = !!tabITA || fs.existsSync(filePath);

    if (!descargaOk) {
      let mensajeError = "ERROR: No se ha podido descargar el informe.";
      try {
        const textoDIL = await page.$eval("#DIL", (el) =>
          el.textContent.trim(),
        );
        if (textoDIL) mensajeError = "ERROR: " + textoDIL;
      } catch (_) {
        try {
          const textoBody = await page.$eval("body", (el) =>
            el.innerText.trim().slice(0, 200),
          );
          if (textoBody) mensajeError = "ERROR (página): " + textoBody;
        } catch (_2) {}
      }
      hoja.cell(cliente.filaExcel, colIdx["LOG ITA"]).value(mensajeError);
      console.warn("[ITA] Error en descarga:", mensajeError);
    } else {
      hoja.cell(cliente.filaExcel, colIdx["LOG ITA"]).value("OK");
      if (tabITA && typeof tabITA.close === "function") {
        try {
          await tabITA.close();
        } catch (_) {}
      }
    }
    await this.esperar(1000);
  }

  async _procesarCertificadoSS({ browser, page, cliente, paths, hoja, colIdx }) {
    console.log(
      `[CERT SS] Iniciando para cliente: ${cliente.codigo} - ${cliente.empresa}`,
    );
    const ccc = String(cliente.ccc);

    for (let intento = 1; intento <= 2; intento++) {
      try {
        await page.goto(
          "https://w2.seg-social.es/ProsaInternet/OnlineAccess?ARQ.SPM.ACTION=LOGIN&ARQ.SPM.APPTYPE=SERVICE&ARQ.IDAPP=XV21F001",
          { waitUntil: "networkidle0" },
        );
        break;
      } catch (e) {
        console.warn(
          `[CERT SS] Fallo navegación (intento ${intento}):`,
          e?.message || e,
        );
        if (intento === 2) throw e;
        await this.esperar(1500);
      }
    }

    try {
      await page.locator('a[id="enlace_316077"]').click();
    } catch (e) {
      throw new Error(`[SS-Paso enlace ARED] ${e.message}`);
    }
    try {
      await page.locator('button[name="SPM.ACC.AC_BUSCAR_OAR"]').click();
    } catch (e) {
      throw new Error(`[SS-Paso botón buscar inicial] ${e.message}`);
    }

    try {
      await page.waitForSelector(`input[title="Buscar por CCC o NAF"]`, {
        timeout: 60000,
      });
    } catch (e) {
      throw new Error(`[SS-Paso radio CCC/NAF] ${e.message}`);
    }
    const radio = await page.$(`input[title="Buscar por CCC o NAF"]`);
    if (radio) await radio.click();

    try {
      await page.waitForSelector('input[name="criteriosBusquedaCccNaf"]', {
        timeout: 60000,
      });
    } catch (e) {
      throw new Error(`[SS-Paso campo CCC] ${e.message}`);
    }
    await page.type('input[name="criteriosBusquedaCccNaf"]', ccc);
    await this.esperar(1000);
    try {
      await page.locator('button[name="SPM.ACC.AC_BUSCAR_OAR"]').click();
    } catch (e) {
      throw new Error(`[SS-Paso botón buscar CCC] ${e.message}`);
    }

    const selectorResultado = 'a[id="enlace_' + String(Number(ccc)) + '"]';
    const enlaceResultado = await page
      .waitForSelector(selectorResultado, { timeout: 10000 })
      .catch(() => null);
    if (!enlaceResultado) {
      throw new Error("CCC no encontrado en el sistema ARED: " + ccc);
    }
    await enlaceResultado.click();

    await page.locator('button[name="SPM.ACC.CONTINUAR"]').click();
    await page.locator('button[name="SPM.ACC.IMPRIMIR"]').click();
    await page.waitForNavigation({ waitUntil: "load" });

    const enlaces = await page.$$("a");
    let enlaceEncontrado = null;
    for (const enlace of enlaces) {
      const texto = await page.evaluate((el) => el.innerText, enlace);
      if (texto.includes("Certificado genérico")) {
        enlaceEncontrado = enlace;
        break;
      }
    }
    if (!enlaceEncontrado) {
      throw new Error("No se encontró el enlace 'Certificado genérico'.");
    }

    const rutaSS = path.join(paths.resultados, cliente.nombreArchivoSS);
    let nuevaPagina = await this._descargarPDF({
      browser,
      botonClick: () => enlaceEncontrado.click(),
      rutaArchivo: rutaSS,
      etiqueta: "SS",
      timeoutMs: 15000,
    });

    if (!nuevaPagina) {
      console.log("[CERT SS] Reintentando descarga...");
      await this.esperar(3000);
      nuevaPagina = await this._descargarPDF({
        browser,
        botonClick: () => enlaceEncontrado.click(),
        rutaArchivo: rutaSS,
        etiqueta: "SS",
        timeoutMs: 15000,
      });
    }

    await this.esperar(1000);

    if (!nuevaPagina) {
      console.log("[CERT SS] ERROR EN DESCARGA");
      hoja
        .cell(cliente.filaExcel, 13)
        .value("ERROR: No se ha podido descargar el certificado.");
    } else {
      hoja.cell(cliente.filaExcel, colIdx["LOG SS"]).value("OK, certificado descargado.");
      try {
        await nuevaPagina.close();
      } catch (_) {}
    }
  }

  async _procesarCertificadoAEAT({ browser, page, cliente, paths, hoja, colIdx }) {
    if (cliente.flagDupeNIF) {
      hoja
        .cell(cliente.filaExcel, colIdx["LOG TRIB"])
        .value("WARNING: Solicitud evitada por duplicidad en NIF.");
      return;
    }

    console.log(
      `[CERT TRIB] Iniciando para cliente: ${cliente.codigo} - ${cliente.empresa}`,
    );

    for (let intento = 1; intento <= 2; intento++) {
      try {
        await page.goto(
          "https://www1.agenciatributaria.gob.es/wlpl/EMCE-JDIT/ECOTInternetCiudadanosServlet",
          { waitUntil: "networkidle0" },
        );
        break;
      } catch (e) {
        console.warn(
          `[CERT TRIB] Fallo navegación (intento ${intento}):`,
          e?.message || e,
        );
        if (intento === 2) throw e;
        await this.esperar(1500);
      }
    }

    try {
      const botonModal = await page.waitForSelector(
        'button[data-dismiss="modal"]',
        { timeout: 1000 },
      );
      if (botonModal) {
        await botonModal.click();
      }
    } catch (_) {}

    await page.locator(`input[id="fTipoRepresentacion1"]`).wait();
    const radio1 = await page.$(`input[id="fTipoRepresentacion1"]`);
    if (radio1) await radio1.click();

    await page.locator('input[name="fNifT"]').wait();
    await page.type('input[name="fNifT"]', String(cliente.nif));
    await this.esperar(500);

    await page.locator('input[name="fNombreT"]').wait();
    await page.type('input[name="fNombreT"]', String(cliente.empresa));
    await this.esperar(500);

    await page.locator(`input[id="fTipoCertificado4"]`).wait();
    const radio2 = await page.$(`input[id="fTipoCertificado4"]`);
    if (radio2) await radio2.click();

    await page.locator('input[id="validarSolicitud"]').click();
    await page.waitForNavigation({ waitUntil: "load" });

    await page.locator('input[value="Firmar Enviar"]').wait();

    let firmaOk;
    try {
      [firmaOk] = await Promise.all([
        new Promise((resolvePromise) => {
          const onTargetCreated = async (target) => {
            const newPage = await target.page();
            await this.esperar(1000);
            await newPage.locator('input[id="Conforme"]').wait();
            await newPage.locator('input[id="Conforme"]').click();
            await this.esperar(500);
            await newPage.locator('input[name="Firmar"]').wait();
            await newPage.locator('input[name="Firmar"]').click();
            try {
              await newPage.close();
            } catch (_) {}
            resolvePromise(true);
          };
          browser.once("targetcreated", onTargetCreated);
          setTimeout(() => {
            browser.off("targetcreated", onTargetCreated);
            resolvePromise(false);
          }, 10000);
        }),
        page.locator('input[value="Firmar Enviar"]').click(),
      ]);
    } catch (e) {
      console.log("[CERT TRIB] Error firma: ", e?.message || e);
    }

    await this.esperar(1000);
    console.log("[CERT TRIB] Descargando...");

    await page.locator('input[id="descarga"]').wait();

    const rutaTrib = path.join(paths.resultados, cliente.nombreArchivoTrib);
    let nuevaPagina = await this._descargarPDF({
      browser,
      botonClick: () => page.locator('input[id="descarga"]').click(),
      rutaArchivo: rutaTrib,
      etiqueta: "TRIB",
      timeoutMs: 15000,
    });

    if (!nuevaPagina) {
      console.log("[CERT TRIB] Reintentando descarga...");
      await this.esperar(3000);
      nuevaPagina = await this._descargarPDF({
        browser,
        botonClick: () => page.locator('input[id="descarga"]').click(),
        rutaArchivo: rutaTrib,
        etiqueta: "TRIB",
        timeoutMs: 15000,
      });
    }

    if (!nuevaPagina) {
      console.log("[CERT TRIB] ERROR EN DESCARGA");
      hoja
        .cell(cliente.filaExcel, colIdx["LOG TRIB"])
        .value("ERROR: No se ha podido generar el resguardo de la solicitud.");
    } else {
      hoja
        .cell(cliente.filaExcel, colIdx["LOG TRIB"])
        .value("OK, resguardo de solicitud descargado.");
      try {
        await nuevaPagina.close();
      } catch (_) {}
    }
  }

  async _procesarCertificadoATC({ browser, page, cliente, paths, hoja, colIdx }) {
    if (cliente.flagDupeNIF) {
      hoja
        .cell(cliente.filaExcel, colIdx["LOG ATC"])
        .value("WARNING: Solicitud evitada por duplicidad en NIF.");
      return;
    }

    console.log(
      `[CERT ATC] Iniciando para cliente: ${cliente.codigo} - ${cliente.empresa}`,
    );

    for (let intento = 1; intento <= 2; intento++) {
      try {
        await page.goto(
          "https://sede.gobiernodecanarias.org/tributos/ov/seguro/certificados/individual/listado.jsp",
          { waitUntil: "networkidle0" },
        );
        break;
      } catch (e) {
        console.warn(
          `[CERT ATC] Fallo navegación (intento ${intento}):`,
          e?.message || e,
        );
        if (intento === 2) throw e;
        await this.esperar(1500);
      }
    }
    await this.esperar(1000);

    // PASO 1: Página selectora de login (cert vs clave)
    // Detecta por la imagen del certificado digital, no por URL
    try {
      await page.waitForSelector(
        'img[alt="img_dig1"], img[src*="certificadoDigital"]',
        { timeout: 3000 },
      );
      await page.evaluate(() => {
        const img =
          document.querySelector('img[alt="img_dig1"]') ||
          document.querySelector('img[src*="certificadoDigital"]');
        if (img?.parentElement?.tagName === "A") img.parentElement.click();
      });
      await page
        .waitForNavigation({ waitUntil: "networkidle0", timeout: 10000 })
        .catch(() => {});
      await this.esperar(1000);
    } catch (_) {
      // Ya autenticado o página no encontrada, se continúa
    }

    // PASO 2: valida.jsp — clicar "Entrar" y esperar selección de certificado
    if (page.url().includes("/publico/validacion/")) {
      try {
        const botonEntrar = await page.waitForSelector(
          'input[id="btnValidar"]',
          { timeout: 5000 },
        );
        if (botonEntrar) await botonEntrar.click();
      } catch (_) {}

      try {
        await page.waitForFunction(
          () => !window.location.href.includes("/publico/validacion/"),
          { timeout: 120000 },
        );
      } catch (_) {
        throw new Error(
          "Tiempo de autenticación ATC agotado. Seleccione el certificado cuando se le pida.",
        );
      }
      await this.esperar(2000);
    }

    try {
      const botonSolicitar = await page.waitForSelector(
        'input[id="btnSolicitar"]',
        { timeout: 60000 },
      );
      if (botonSolicitar) {
        await botonSolicitar.click();
      }
    } catch (_) {
      throw new Error("No se localizó el botón Solicitar inicial (ATC).");
    }

    await page.locator(`select[name="tiposCertificado"]`).wait();
    await this.esperar(500);
    await page.select('select[name="tiposCertificado"]', "AS");

    await page.locator(`input[id="id_tipo_terceros"]`).wait();
    const radio = await page.$(`input[id="id_tipo_terceros"]`);
    if (radio) await radio.click();

    await this.esperar(1000);

    await page.locator('input[id="idNifTitular"]').wait();
    await page.type('input[id="idNifTitular"]', String(cliente.nif));
    await this.esperar(500);

    await page.locator('input[id="idNombreTitular"]').wait();
    await page.type('input[id="idNombreTitular"]', String(cliente.empresa));
    await this.esperar(500);

    await page.locator('input[id="btnSolicitar"]').wait();
    await page.locator('input[id="btnSolicitar"]').click();

    if ((await page.evaluate(() => document.readyState)) !== "complete") {
      await page.waitForNavigation({ waitUntil: "load" });
    }

    console.log("[CERT ATC] Solicitud realizada");

    if ((await page.evaluate(() => document.readyState)) !== "complete") {
      await page.waitForNavigation({ waitUntil: "load" });
    }

    console.log("[CERT ATC] Descargando...");
    try {
      await page.waitForSelector('input[id="btnDescargar"]', {
        timeout: 40000,
      });
    } catch (_) {
      hoja
        .cell(cliente.filaExcel, colIdx["LOG ATC"])
        .value("ERROR: No se ha podido generar la solicitud.");
      return;
    }

    await this.esperar(1000);

    const rutaATC = path.join(paths.resultados, cliente.nombreArchivoATC);
    let nuevaPagina = await this._descargarPDF({
      browser,
      botonClick: () => page.locator('input[id="btnDescargar"]').click(),
      rutaArchivo: rutaATC,
      etiqueta: "ATC",
      timeoutMs: 20000,
    });

    if (!nuevaPagina) {
      console.log("[CERT ATC] Reintentando descarga...");
      await this.esperar(3000);
      nuevaPagina = await this._descargarPDF({
        browser,
        botonClick: () => page.locator('input[id="btnDescargar"]').click(),
        rutaArchivo: rutaATC,
        etiqueta: "ATC",
        timeoutMs: 20000,
      });
    }

    if (!nuevaPagina) {
      console.log("[CERT ATC] ERROR ABRIENDO DESCARGA");
      hoja
        .cell(cliente.filaExcel, colIdx["LOG ATC"])
        .value("ERROR: No se ha podido generar el resguardo de la solicitud.");
    } else {
      hoja
        .cell(cliente.filaExcel, colIdx["LOG ATC"])
        .value("OK, resguardo de solicitud descargado.");
      try {
        await nuevaPagina.close();
      } catch (_) {}
    }
  }

  async _procesarCertificadoArt42({
    browser,
    page,
    cliente,
    paths,
    hoja,
    colIdx,
    empresaAutRegimen,
    empresaAutTesoreria,
    empresaAutCuenta,
  }) {
    console.log(
      `[ART42] Iniciando para cliente: ${cliente.codigo} - ${cliente.empresa}`,
    );

    for (let intento = 1; intento <= 2; intento++) {
      try {
        await page.goto("https://w2.seg-social.es/fs/indexframes.html", {
          waitUntil: "networkidle0",
        });
        break;
      } catch (e) {
        console.warn(
          `[ART42] Fallo navegación (intento ${intento}):`,
          e?.message || e,
        );
        if (intento === 2) throw e;
        await this.esperar(1500);
      }
    }
    await this.esperar(1000);

    const getFrame = () => page.mainFrame().childFrames()[0];

    let frame = getFrame();
    if (!frame) throw new Error("[ART42] No se encontró el iframe del menú.");

    await frame.waitForSelector("a", { timeout: 10000 });
    const clickedGestion = await frame.evaluate(() => {
      const link = Array.from(document.querySelectorAll("a")).find((a) =>
        a.textContent.includes("Gestión de Deuda"),
      );
      if (link) {
        link.click();
        return true;
      }
      return false;
    });
    if (!clickedGestion) {
      throw new Error('[ART42] Enlace "Gestión de Deuda" no encontrado.');
    }
    await this.esperar(1000);

    frame = getFrame();
    if (!frame) {
      throw new Error(
        "[ART42] No se encontró el iframe tras expandir Gestión de Deuda.",
      );
    }
    await frame.waitForSelector("a", { timeout: 10000 });
    const clickedArt42 = await frame.evaluate(() => {
      const link = Array.from(document.querySelectorAll("a")).find((a) =>
        a.textContent.includes("Autorización Certificado Art.42"),
      );
      if (link) {
        link.click();
        return true;
      }
      return false;
    });
    if (!clickedArt42) {
      throw new Error(
        '[ART42] Enlace "Autorización Certificado Art.42 Est.Trab." no encontrado.',
      );
    }

    await page
      .waitForNavigation({ waitUntil: "networkidle0", timeout: 20000 })
      .catch(() => {});
    await this.esperar(500);

    frame = getFrame();
    if (!frame)
      throw new Error("[ART42] No se encontró el frame del formulario.");

    try {
      await frame.waitForSelector("#SDFREGIMEN", { timeout: 15000 });
    } catch (e) {
      throw new Error(`[ART42] #SDFREGIMEN no apareció: ${e.message}`);
    }

    await frame.type("#SDFREGIMEN", String(cliente.ccc1));
    await frame.type("#SDFPROVINCIA", String(cliente.ccc2));
    await frame.type("#SDFNISS", String(cliente.ccc3));

    try {
      await frame.select("#SDFOPCION", "Alta");
    } catch (e) {
      throw new Error(
        `[ART42] Error seleccionando Alta en #SDFOPCION: ${e.message}`,
      );
    }

    await Promise.all([
      page
        .waitForNavigation({ waitUntil: "networkidle0", timeout: 20000 })
        .catch(() => {}),
      frame.click("#Sub2207001004_35"),
    ]);
    await this.esperar(500);

    frame = getFrame();
    if (!frame)
      throw new Error("[ART42] No se encontró el frame tras Continuar 1.");

    try {
      await frame.waitForSelector("#SDFREGKCGK", { timeout: 15000 });
    } catch (e) {
      throw new Error(`[ART42] #SDFREGKCGK no apareció: ${e.message}`);
    }

    await frame.type("#SDFREGKCGK", empresaAutRegimen);
    await frame.type("#SDFTESCCGK", empresaAutTesoreria);
    await frame.type("#SDFCCONCGK9", empresaAutCuenta);

    const ahora = DateTime.now().setZone("Europe/Madrid");
    const hasta = ahora.plus({ years: 1 });

    await frame.type("#SDFDIADESDE", ahora.toFormat("dd"));
    await frame.type("#SDFMESDESDE", ahora.toFormat("MM"));
    await frame.type("#SDFAODESDE", ahora.toFormat("yyyy"));
    await frame.type("#SDFDIAHASTA", hasta.toFormat("dd"));
    await frame.type("#SDFMESHASTA", hasta.toFormat("MM"));
    await frame.type("#SDFAOHASTA", hasta.toFormat("yyyy"));

    await Promise.all([
      page
        .waitForNavigation({ waitUntil: "networkidle0", timeout: 20000 })
        .catch(() => {}),
      frame.click("#Sub2207001004_75"),
    ]);
    await this.esperar(500);

    frame = getFrame();
    if (!frame)
      throw new Error("[ART42] No se encontró el frame tras Continuar 2.");

    try {
      await frame.waitForSelector("#Sub2204701006_74", { timeout: 15000 });
    } catch (e) {
      throw new Error(`[ART42] Botón Confirmar no apareció: ${e.message}`);
    }

    await Promise.all([
      page
        .waitForNavigation({ waitUntil: "networkidle0", timeout: 30000 })
        .catch(() => {}),
      frame.click("#Sub2204701006_74"),
    ]);
    await this.esperar(1000);

    const rutaScreenshot = path.join(
      paths.resultados,
      cliente.nombreArchivoArt42,
    );
    try {
      await page.screenshot({ path: rutaScreenshot, fullPage: false });
      console.log(`[ART42] Screenshot guardado: ${rutaScreenshot}`);
    } catch (e) {
      throw new Error(`[ART42] Error guardando screenshot: ${e.message}`);
    }

    hoja.cell(cliente.filaExcel, colIdx["LOG ART42"]).value("OK, autorización generada.");
    await this.esperar(1000);
  }
}

module.exports = ProcesosCertificados;
