const path = require("path");
const fs = require("fs");
const XlsxPopulate = require("xlsx-populate");
const { DateTime } = require("luxon");
const { execSync } = require("child_process");
const os = require("os");

const { registrarEjecucion, agruparPorEmpresa } = require("../metricas");
const puppeteer = require("puppeteer");

class ProcesosCertificados {
  constructor(pathToDbFolder, nombreProyecto, proyectoDB) {
    this.pathToDbFolder = pathToDbFolder;
    this.nombreProyecto = nombreProyecto;
    this.proyectoDB = proyectoDB;

    // Selectores CSS por portal
    // IMPORTANTE: Estos selectores son frágiles y pueden cambiar con actualizaciones del portal.
    // Se han parametrizado aquí para facilitar futuros mantenimientos.
    this.SELECTORS = {
      SS: {
        // Portal ARED (Seguridad Social)
        enlaceAred: 'a[id="enlace_316077"]', // ID de sección ARED en portal SS
        btnBuscarOAR: 'button[name="SPM.ACC.AC_BUSCAR_OAR"]',
        radioCCC: 'input[title="Buscar por CCC o NAF"]',
        campoCCC: 'input[name="criteriosBusquedaCccNaf"]',
        btnBuscarCCC: 'button[name="SPM.ACC.AC_BUSCAR_OAR"]',
        enlaceResultado: (ccc) => `a[id="enlace_${ccc}"]`, // Dinámico por CCC
        btnContinuar: 'button[name="SPM.ACC.CONTINUAR"]',
        btnImprimir: 'button[name="SPM.ACC.IMPRIMIR"]',
        linkCertGenerico: 'a', // Buscado por texto "Certificado genérico"
      },
      AEAT: {
        // Portal Agencia Tributaria
        radioBuscadorTipo: 'input[id="fTipoRepresentacion0"]', // Tipo representación
        radioCertificadoTipo: 'input[id="fTipoCertificado4"]', // Tipo certificado
        btnValidarSolicitud: 'input[id="validarSolicitud"]',
        btnFirmarEnviar: 'input[value="Firmar Enviar"]',
        btnConforme: 'input[id="Conforme"]', // En popup de firma
        btnFirmar: 'input[name="Firmar"]', // En popup de firma
        btnDescarga: 'input[id="descarga"]',
      },
      ATC: {
        // Sede de Canarias (Autoridad Tributaria Canaria)
        imgCertDigital: 'img[alt="img_dig1"], img[src*="certificadoDigital"]',
        btnValidar: 'input[id="btnValidar"]',
        btnSolicitarInicial: 'input[id="btnSolicitar"]',
        selectTipoCertificado: 'select[name="tiposCertificado"]',
        radioTipoTerceros: 'input[id="id_tipo_terceros"]',
        campNifTitular: 'input[id="idNifTitular"]',
        campNombreTitular: 'input[id="idNombreTitular"]',
        btnSolicitar: 'input[id="btnSolicitar"]',
        btnDescargar: 'input[id="btnDescargar"]',
      },
      ITA: {
        // SS Informe de Trabajadores Activos (SGIRED)
        campRegimen: 'input[name="txt_SDFREG62_ayuda"]',
        campTesoreria: 'input[name="txt_SDFTESO62"]',
        campNumero: 'input[name="txt_SDFNUM62"]',
        selectTipoImpresion: 'select[name="cbo_ListaTipoImpresion"]',
        btnSubmit: 'input[name="btn_Sub2207601004"]',
      },
      ART42: {
        // SS Autorización Certificado Art.42
        campRegimen: '#SDFREGIMEN',
        campProvincia: '#SDFPROVINCIA',
        campNISS: '#SDFNISS',
        selectOpcion: '#SDFOPCION',
        btnContinuar1: '#Sub2207001004_35',
        campRegKemsoCGK: '#SDFREGKCGK',
        campTesoreriaCGK: '#SDFTESCCGK',
        campCuentaCGK: '#SDFCCONCGK9',
        btnContinuar2: '#Sub2207001004_75',
        btnConfirmar: '#Sub2204701006_74',
      },
    };

    // Nombre de la hoja de datos en Excel
    this.NOMBRE_HOJA_DATOS = "BASE DE DATOS (NO TOCAR)";
  }

  async esperar(tiempo) {
    return new Promise((resolve) => {
      setTimeout(resolve, tiempo);
    });
  }

  _validarInputs(argumentos) {
    if (!argumentos || !argumentos.formularioControl || !Array.isArray(argumentos.formularioControl)) {
      throw new Error("Argumentos inválidos: esperado argumentos.formularioControl como array");
    }

    const fc = argumentos.formularioControl;
    const expectedIndices = [
      { idx: 0, name: "chromiumExecutablePath", type: "string" },
      { idx: 1, name: "pathArchivoEtiquetas", type: "string" },
      { idx: 2, name: "pathBase", type: "string" },
      { idx: 3, name: "modoManual", type: "boolean-or-truthy" },
      { idx: 4, name: "codigosEmpresaInput", type: "string-or-empty" },
      { idx: 5, name: "runSS", type: "boolean-or-truthy" },
      { idx: 6, name: "runAEAT", type: "boolean-or-truthy" },
      { idx: 7, name: "runATC", type: "boolean-or-truthy" },
      { idx: 8, name: "runITA", type: "boolean-or-truthy" },
      { idx: 9, name: "runArt42", type: "boolean-or-truthy" },
      { idx: 10, name: "empresaAutRegimen", type: "string-or-empty" },
      { idx: 11, name: "empresaAutTesoreria", type: "string-or-empty" },
      { idx: 12, name: "empresaAutCuenta", type: "string-or-empty" },
    ];

    for (const { idx, name } of expectedIndices) {
      if (fc.length <= idx || fc[idx] === undefined) {
        throw new Error(`Argumento faltante en índice ${idx} (${name})`);
      }
    }

    // Validar que las rutas existen
    if (!fs.existsSync(fc[0])) {
      throw new Error(`Ejecutable Chromium no encontrado: ${fc[0]}`);
    }
    if (!fs.existsSync(fc[1])) {
      throw new Error(`Archivo Excel no encontrado: ${fc[1]}`);
    }
    if (!fs.existsSync(fc[2])) {
      throw new Error(`Carpeta base no encontrada: ${fc[2]}`);
    }

    return {
      chromiumExecutablePath: fc[0],
      pathArchivoEtiquetas: fc[1],
      pathBase: fc[2],
      modoManual: !!fc[3],
      codigosEmpresaInput: fc[4],
      runSS: !!fc[5],
      runAEAT: !!fc[6],
      runATC: !!fc[7],
      runITA: !!fc[8],
      runArt42: !!fc[9],
      empresaAutRegimen: fc[10],
      empresaAutTesoreria: fc[11],
      empresaAutCuenta: fc[12],
    };
  }

  async _ejecutarConReintentos(fn, descripcion, page, maxReintentos = 2) {
    let ultimoError;
    for (let intento = 1; intento <= maxReintentos; intento++) {
      try {
        await fn();
        return;
      } catch (e) {
        ultimoError = e;
        if (intento < maxReintentos) {
          console.warn(`[${descripcion}] Error (intento ${intento}/${maxReintentos}): ${e?.message || e}. Reintentando...`);
          try {
            await page.goto("about:blank");
          } catch (_) {}
          await this.esperar(2000);
        }
      }
    }
    throw ultimoError;
  }

  async _esperarSelector(page, selector, timeoutTotal = 60000, reintentos = 3) {
    let ultimoError;
    const timeoutPorIntento = Math.ceil(timeoutTotal / reintentos);
    for (let i = 1; i <= reintentos; i++) {
      try {
        await page.waitForSelector(selector, { timeout: timeoutPorIntento });
        return;
      } catch (e) {
        ultimoError = e;
        if (i < reintentos) await this.esperar(1000);
      }
    }
    throw ultimoError;
  }

  async _navegarConReintentos(page, url, maxReintentos = 2) {
    let ultimoError;
    for (let intento = 1; intento <= maxReintentos; intento++) {
      try {
        await page.goto(url, { waitUntil: "networkidle0" });
        return;
      } catch (e) {
        ultimoError = e;
        if (intento < maxReintentos) {
          await this.esperar(1500);
        }
      }
    }
    throw ultimoError;
  }

  async _descargaPDFConReintento(pdfOptions) {
    let nuevaPagina = await this._descargarPDF(pdfOptions);
    if (!nuevaPagina) {
      console.log(`[${pdfOptions.etiqueta}] Reintentando descarga...`);
      await this.esperar(3000);
      nuevaPagina = await this._descargarPDF(pdfOptions);
    }
    return nuevaPagina;
  }

  async _procesarLoginATC(page) {
    try {
      await page.waitForSelector(
        this.SELECTORS.ATC.imgCertDigital,
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
          this.SELECTORS.ATC.btnValidar,
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
        throw new Error("Timeout esperando autenticación ATC (seleccionar certificado cuando se pida)");
      }
      await this.esperar(2000);
    }
  }

  _obtenerCNcertificado(nif) {
    const scriptPath = path.join(os.tmpdir(), `cert_lookup_${Date.now()}.ps1`);
    const nifSafe = (nif || "").replace(/'/g, "''");
    const script = `[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$nif = '${nifSafe}'
$today = Get-Date
$store = New-Object System.Security.Cryptography.X509Certificates.X509Store('My', [System.Security.Cryptography.X509Certificates.StoreLocation]::CurrentUser)
$store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadOnly)
$cert = $store.Certificates |
  Where-Object { $_.Subject -match [regex]::Escape($nif) -and $_.NotAfter -gt $today } |
  Sort-Object NotAfter -Descending |
  Select-Object -First 1
$store.Close()
if ($cert) {
  $subjectCN = ($cert.Subject -split ',') |
    Where-Object { $_.Trim().StartsWith('CN=') } |
    Select-Object -First 1
  $subjectCN = $subjectCN.Trim().Substring(3)
  $issuerCN = ($cert.Issuer -split ',') |
    Where-Object { $_.Trim().StartsWith('CN=') } |
    Select-Object -First 1
  $issuerCN = $issuerCN.Trim().Substring(3)
  Write-Output "$subjectCN|||$issuerCN"
} else {
  Write-Output "NOT_FOUND"
}`;
    fs.writeFileSync(scriptPath, '﻿' + script, "utf8");
    try {
      const out = execSync(
        `powershell -NoProfile -ExecutionPolicy Bypass -File "${scriptPath}"`,
        { encoding: "utf8", timeout: 15000 }
      ).trim();
      if (!out || out === "NOT_FOUND") return null;
      const [subjectCN, issuerCN] = out.split("|||");
      return { subjectCN: subjectCN.trim(), issuerCN: issuerCN.trim() };
    } finally {
      try { fs.unlinkSync(scriptPath); } catch (_) {}
    }
  }

  _setAutoSelectPolicy({ subjectCN }) {
    const scriptPath = path.join(os.tmpdir(), `cert_policy_${Date.now()}.ps1`);
    const filter = {};
    if (subjectCN) filter.SUBJECT = { CN: subjectCN };
    const policy = JSON.stringify({ pattern: "https://[*.]agenciatributaria.gob.es", filter });
    const safePolicy = policy.replace(/'/g, "''");
    const script = [
      `New-Item -Path 'HKCU:\\Software\\Policies\\Google\\Chrome\\AutoSelectCertificateForUrls' -Force | Out-Null`,
      `Set-ItemProperty -Path 'HKCU:\\Software\\Policies\\Google\\Chrome\\AutoSelectCertificateForUrls' -Name '1' -Value '${safePolicy}'`,
    ].join("\r\n");
    // BOM UTF-8 (﻿): PowerShell 5.x lee archivos sin BOM como ANSI, corrompiendo
    // los caracteres acentuados de los CN. Con BOM los lee como UTF-8 y escribe los
    // caracteres Unicode reales al registro, que Chrome compara directamente.
    fs.writeFileSync(scriptPath, '﻿' + script, "utf8");
    try {
      execSync(`powershell -NoProfile -ExecutionPolicy Bypass -File "${scriptPath}"`, { encoding: "utf8", timeout: 30000 });
      console.log(`[POLICY] AutoSelect policy set: ${policy}`);
    } finally {
      try { fs.unlinkSync(scriptPath); } catch (_) {}
    }
  }

  _limpiarAutoSelectPolicy() {
    const scriptPath = path.join(os.tmpdir(), `cert_policy_clean_${Date.now()}.ps1`);
    const script = `Remove-Item -Path 'HKCU:\\Software\\Policies\\Google\\Chrome\\AutoSelectCertificateForUrls' -Force -Recurse -ErrorAction SilentlyContinue`;
    fs.writeFileSync(scriptPath, script, "utf8");
    try {
      execSync(`powershell -NoProfile -ExecutionPolicy Bypass -File "${scriptPath}"`, { encoding: "utf8", timeout: 10000 });
    } catch (_) {} finally {
      try { fs.unlinkSync(scriptPath); } catch (_) {}
    }
  }

  async certificadosSSITAATC(argumentos) {
    return this._ejecutarCertificados(argumentos, {
      habilitarSS: true, habilitarAEAT: false, habilitarATC: true, habilitarITA: true, habilitarArt42: false,
      nombreProceso: 'Certificados SS ITA ATC'
    });
  }

  async certificadoAEAT(argumentos) {
    return this._ejecutarCertificados(argumentos, {
      habilitarSS: false, habilitarAEAT: true, habilitarATC: false, habilitarITA: false, habilitarArt42: false,
      nombreProceso: 'Certificado AEAT'
    });
  }

  async certificadoArt42(argumentos) {
    return this._ejecutarCertificados(argumentos, {
      habilitarSS: false, habilitarAEAT: false, habilitarATC: false, habilitarITA: false, habilitarArt42: true,
      nombreProceso: 'Certificado Art 42'
    });
  }

  async certificadosDeEstarAlCorriente(argumentos) {
    return this._ejecutarCertificados(argumentos, {
      habilitarSS: true, habilitarAEAT: true, habilitarATC: true, habilitarITA: true, habilitarArt42: false,
      nombreProceso: 'Certificados Unificados'
    });
  }

  async _ejecutarCertificados(argumentos, config) {
    return new Promise((resolve) => {
      console.log(`${config.nombreProceso} — iniciando`);
      const nombreProceso = config.nombreProceso;
      let registrosProcesados = 0;

      let validacion;
      try {
        validacion = this._validarInputs(argumentos);
      } catch (e) {
        console.error(`[CERT] Error en validación de entrada: ${e.message}`);
        return resolve(false);
      }

      const chromiumExecutablePath = path.normalize(validacion.chromiumExecutablePath);
      const pathArchivoEtiquetas = validacion.pathArchivoEtiquetas;
      const pathBase = validacion.pathBase;
      const modoManual = validacion.modoManual;
      const codigosEmpresaInput = validacion.codigosEmpresaInput;

      let runSS = config.habilitarSS && validacion.runSS;
      let runTrib = config.habilitarAEAT && validacion.runAEAT;
      let runATC = config.habilitarATC && validacion.runATC;
      let runITA = config.habilitarITA && validacion.runITA;
      let runArt42 = config.habilitarArt42 && validacion.runArt42;

      const empresaAutRegimen = String(validacion.empresaAutRegimen || "");
      const empresaAutTesoreria = String(validacion.empresaAutTesoreria || "");
      const empresaAutCuenta = String(validacion.empresaAutCuenta || "");

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
            const hoja = archivo.sheet(this.NOMBRE_HOJA_DATOS);
            if (!hoja) {
              console.warn(
                `[CERT] Hoja '${this.NOMBRE_HOJA_DATOS}' no encontrada en el Excel.`,
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
            let nextColIdx = 1;
            cabeceras.forEach((h, i) => {
              if (h !== undefined && h !== null) {
                colIdx[String(h).trim()] = i + 1;
                nextColIdx = Math.max(nextColIdx, i + 2);
              }
            });

            // Añadir columnas LOG al final, de forma dinámica
            const addLogCol = (nombre) => {
              hoja.cell(1, nextColIdx).value(nombre);
              colIdx[nombre] = nextColIdx;
              nextColIdx++;
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

              const tieneAlgunFlagHabilitado =
                (config.habilitarSS && objetoCliente.flagSS) ||
                (config.habilitarAEAT && objetoCliente.flagAEAT) ||
                (config.habilitarATC && objetoCliente.flagATC) ||
                (config.habilitarITA && objetoCliente.flagITA);

              const debeProcesarse = modoManual
                ? debeProcesarseManual
                : codigoNormalizado !== "" && tieneAlgunFlagHabilitado;

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
              runSS = config.habilitarSS && clientes.some((c) => c.flagSS);
              runTrib = config.habilitarAEAT && clientes.some((c) => c.flagAEAT);
              runATC = config.habilitarATC && clientes.some((c) => c.flagATC);
              runITA = config.habilitarITA && clientes.some((c) => c.flagITA);
              runArt42 = false;

              if (!runSS && !runTrib && !runATC && !runITA && !runArt42) {
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
                try {
                  await dialog.accept();
                } catch (_) {}
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
              this._limpiarAutoSelectPolicy();
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

              const clientRunSS = modoManual ? runSS : (config.habilitarSS && clientes[i].flagSS);
              const clientRunTrib = modoManual ? runTrib : (config.habilitarAEAT && clientes[i].flagAEAT);
              const clientRunATC = modoManual ? runATC : (config.habilitarATC && clientes[i].flagATC);
              const clientRunITA = modoManual ? runITA : (config.habilitarITA && clientes[i].flagITA);
              const clientRunArt42 = modoManual && runArt42;

              if (clientRunSS) {
                await this._ejecutarConReintentos(
                  () => this._procesarCertificadoSS({
                    browser,
                    page,
                    cliente: clientes[i],
                    paths: paths.ss,
                    hoja,
                    colIdx,
                  }),
                  "CERT SS",
                  page
                ).catch((e) => {
                  hoja
                    .cell(clientes[i].filaExcel, colIdx["LOG SS"])
                    .value("ERROR: " + (e?.message || e));
                });
              }

              if (clientRunTrib) {
                await this._ejecutarConReintentos(
                  () => this._procesarCertificadoAEAT({
                    browser,
                    page,
                    cliente: clientes[i],
                    paths: paths.trib,
                    hoja,
                    colIdx,
                    executablePath: chromiumExecutablePath,
                  }),
                  "CERT TRIB",
                  page
                ).catch((e) => {
                  hoja
                    .cell(clientes[i].filaExcel, colIdx["LOG TRIB"])
                    .value("ERROR: " + (e?.message || e));
                });
              }

              if (clientRunATC) {
                await this._ejecutarConReintentos(
                  () => this._procesarCertificadoATC({
                    browser,
                    page,
                    cliente: clientes[i],
                    paths: paths.atc,
                    hoja,
                    colIdx,
                  }),
                  "CERT ATC",
                  page
                ).catch((e) => {
                  hoja
                    .cell(clientes[i].filaExcel, colIdx["LOG ATC"])
                    .value("ERROR: " + (e?.message || e));
                });
              }

              if (clientRunITA) {
                await this._ejecutarConReintentos(
                  () => this._procesarInformeITA({
                    browser,
                    page,
                    cliente: clientes[i],
                    paths: paths.ita,
                    hoja,
                    colIdx,
                  }),
                  "[ITA]",
                  page
                ).catch((e) => {
                  hoja
                    .cell(clientes[i].filaExcel, colIdx["LOG ITA"])
                    .value("ERROR: " + (e?.message || e));
                });
              }

              if (clientRunArt42) {
                await this._ejecutarConReintentos(
                  () => this._procesarCertificadoArt42({
                    browser,
                    page,
                    cliente: clientes[i],
                    paths: paths.art42,
                    hoja,
                    colIdx,
                    empresaAutRegimen,
                    empresaAutTesoreria,
                    empresaAutCuenta,
                  }),
                  "[ART42]",
                  page
                ).catch((e) => {
                  hoja
                    .cell(clientes[i].filaExcel, colIdx["LOG ART42"])
                    .value("ERROR: " + (e?.message || e));
                });
              }

              console.log("Nuevo cliente");
              await this.esperar(1000);
            }

            this._limpiarAutoSelectPolicy();

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
      await this._navegarConReintentos(page, "https://w2.seg-social.es/ProsaInternet/OnlineAccess?ARQ.SPM.ACTION=LOGIN&ARQ.SPM.APPTYPE=SERVICE&ARQ.IDAPP=XV21F001");
      console.log("[CERT INIT] SS listo.");
    }

    if (runTrib) {
      console.log(
        "[CERT INIT] TRIB — Los certificados se seleccionarán automáticamente por empresa",
      );
    }

    if (runATC) {
      console.log(
        "[CERT INIT] ATC — navegando para seleccionar certificado...",
      );
      await this._navegarConReintentos(page, "https://sede.gobiernodecanarias.org/tributos/ov/seguro/certificados/individual/listado.jsp");
      await this.esperar(1000);

      try {
        await this._procesarLoginATC(page);
      } catch (e) {
        console.warn("[CERT INIT ATC] Error en login:", e?.message || e);
      }
      console.log("[CERT INIT] ATC listo.");
    }

    if (runArt42) {
      console.log(
        "[CERT INIT] ART42 — navegando para seleccionar certificado digital...",
      );
      await this._navegarConReintentos(page, "https://w2.seg-social.es/fs/indexframes.html");
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
    isPDFResponse = null,
  }) {
    let resuelto = false;
    let timeoutId = null;

    if (!isPDFResponse) {
      isPDFResponse = (response) => {
        const contentType = response.headers()["content-type"] || "";
        return (
          response.url().startsWith("chrome-extension://") &&
          contentType.includes("application/pdf")
        );
      };
    }

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
            if (isPDFResponse(response)) {
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

    await this._navegarConReintentos(page, "https://w2.seg-social.es/Xhtml?JacadaApplicationName=SGIRED&TRANSACCION=ATR64&E=I&AP=AFIR");
    await this.esperar(1000);

    await page.locator(this.SELECTORS.ITA.campRegimen).wait();
    await page.type(this.SELECTORS.ITA.campRegimen, String(cliente.ccc1));
    await page.locator(this.SELECTORS.ITA.campTesoreria).wait();
    await page.type(this.SELECTORS.ITA.campTesoreria, String(cliente.ccc2));
    await page.locator(this.SELECTORS.ITA.campNumero).wait();
    await page.type(this.SELECTORS.ITA.campNumero, String(cliente.ccc3));

    await this.esperar(1000);
    await page.select(this.SELECTORS.ITA.selectTipoImpresion, "OnLine");
    await this.esperar(1000);

    const itaPDFFilter = (response) => {
      return (
        !response.url().endsWith(".js") &&
        !response.url().endsWith(".css") &&
        response.url().startsWith("chrome-extension://")
      );
    };

    const tabITA = await this._descargarPDF({
      browser,
      botonClick: async () => {
        await page.locator(this.SELECTORS.ITA.btnSubmit).wait();
        await page.locator(this.SELECTORS.ITA.btnSubmit).click();
      },
      rutaArchivo: filePath,
      etiqueta: "ITA",
      timeoutMs: 15000,
      isPDFResponse: itaPDFFilter,
    });

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
      if (tabITA) {
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

    await this._navegarConReintentos(page, "https://w2.seg-social.es/ProsaInternet/OnlineAccess?ARQ.SPM.ACTION=LOGIN&ARQ.SPM.APPTYPE=SERVICE&ARQ.IDAPP=XV21F001");

    await this.esperar(2000);

    try {
      const botonModal = await page.waitForSelector('button[data-dismiss="modal"]', { timeout: 2000 });
      if (botonModal) {
        await botonModal.click();
        await this.esperar(500);
      }
    } catch (_) {}

    try {
      await page.locator(this.SELECTORS.SS.enlaceAred).click();
    } catch (e) {
      throw new Error(`[SS-Paso enlace ARED] ${e.message}`);
    }
    try {
      await page.locator(this.SELECTORS.SS.btnBuscarOAR).click();
    } catch (e) {
      throw new Error(`[SS-Paso botón buscar inicial] ${e.message}`);
    }

    try {
      await this._esperarSelector(page, this.SELECTORS.SS.radioCCC, 60000, 3);
    } catch (e) {
      throw new Error(`[SS-Paso radio CCC/NAF] ${e.message}`);
    }
    const radio = await page.$(this.SELECTORS.SS.radioCCC);
    if (radio) await radio.click();

    try {
      await this._esperarSelector(page, this.SELECTORS.SS.campoCCC, 60000, 3);
    } catch (e) {
      throw new Error(`[SS-Paso campo CCC] ${e.message}`);
    }
    await page.type(this.SELECTORS.SS.campoCCC, ccc);
    await this.esperar(1000);
    try {
      await page.locator(this.SELECTORS.SS.btnBuscarCCC).click();
    } catch (e) {
      throw new Error(`[SS-Paso botón buscar CCC] ${e.message}`);
    }

    const selectorResultado = this.SELECTORS.SS.enlaceResultado(ccc);
    const enlaceResultado = await page
      .waitForSelector(selectorResultado, { timeout: 10000 })
      .catch(() => null);
    if (!enlaceResultado) {
      throw new Error("CCC no encontrado en el sistema ARED: " + ccc);
    }
    await enlaceResultado.click();

    await page.locator(this.SELECTORS.SS.btnContinuar).click();
    await page.locator(this.SELECTORS.SS.btnImprimir).click();
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
    const nuevaPagina = await this._descargaPDFConReintento({
      browser,
      botonClick: () => enlaceEncontrado.click(),
      rutaArchivo: rutaSS,
      etiqueta: "SS",
      timeoutMs: 15000,
    });

    await this.esperar(1000);

    if (!nuevaPagina) {
      console.log("[CERT SS] ERROR EN DESCARGA");
      hoja
        .cell(cliente.filaExcel, colIdx["LOG SS"])
        .value("ERROR: No se ha podido descargar el certificado.");
    } else {
      hoja.cell(cliente.filaExcel, colIdx["LOG SS"]).value("OK, certificado descargado.");
      try {
        await nuevaPagina.close();
      } catch (_) {}
    }
  }

  async _procesarCertificadoAEAT({
    browser,
    page,
    cliente,
    paths,
    hoja,
    colIdx,
    executablePath = null,
  }) {
    if (cliente.flagDupeNIF) {
      hoja
        .cell(cliente.filaExcel, colIdx["LOG TRIB"])
        .value("WARNING: Solicitud evitada por duplicidad en NIF.");
      return;
    }

    console.log(
      `[CERT TRIB] Iniciando para cliente: ${cliente.codigo} - ${cliente.empresa}`,
    );

    // Buscar certificado directamente en el almacén Windows por NIF del cliente
    const certInfo = this._obtenerCNcertificado(cliente.nif);
    if (!certInfo) {
      hoja
        .cell(cliente.filaExcel, colIdx["LOG TRIB"])
        .value(`ERROR: No se encontró certificado en almacén Windows para NIF ${cliente.nif}`);
      return;
    }

    console.log(`[CERT TRIB] Certificado encontrado: CN="${certInfo.subjectCN}", ISSUER="${certInfo.issuerCN}"`);

    // Configurar auto-selección Chrome por CN real del cert
    this._setAutoSelectPolicy(certInfo);
    const certBrowser = await puppeteer.launch({ executablePath, headless: false });
    const aeatPage = await certBrowser.newPage();
    await aeatPage.setViewport({ width: 1080, height: 1024 });
    aeatPage.setDefaultTimeout(60000);
    aeatPage.on("dialog", async (dialog) => {
      try { await dialog.accept(); } catch (_) {}
    });
    const activeBrowser = certBrowser;

    try {
      await this._navegarConReintentos(aeatPage, "https://www1.agenciatributaria.gob.es/wlpl/EMCE-JDIT/ECOTInternetCiudadanosServlet");

      try {
        const botonModal = await aeatPage.waitForSelector(
          'button[data-dismiss="modal"]',
          { timeout: 1000 },
        );
        if (botonModal) {
          await botonModal.click();
        }
      } catch (_) {}

      await aeatPage.locator(this.SELECTORS.AEAT.radioBuscadorTipo).wait();
      const radio1 = await aeatPage.$(this.SELECTORS.AEAT.radioBuscadorTipo);
      if (radio1) await radio1.click();
      await this.esperar(500);

      await aeatPage.locator(this.SELECTORS.AEAT.radioCertificadoTipo).wait();
      const radio2 = await aeatPage.$(this.SELECTORS.AEAT.radioCertificadoTipo);
      if (radio2) await radio2.click();

      await aeatPage.locator(this.SELECTORS.AEAT.btnValidarSolicitud).click();
      await aeatPage.waitForNavigation({ waitUntil: "load" });

      await aeatPage.locator(this.SELECTORS.AEAT.btnFirmarEnviar).wait();

      let firmaOk;
      try {
        [firmaOk] = await Promise.all([
          new Promise((resolvePromise) => {
            const onTargetCreated = async (target) => {
              const newPage = await target.page();
              await this.esperar(1000);
              await newPage.locator(this.SELECTORS.AEAT.btnConforme).wait();
              await newPage.locator(this.SELECTORS.AEAT.btnConforme).click();
              await this.esperar(500);
              await newPage.locator(this.SELECTORS.AEAT.btnFirmar).wait();
              await newPage.locator(this.SELECTORS.AEAT.btnFirmar).click();
              try {
                await newPage.close();
              } catch (_) {}
              resolvePromise(true);
            };
            activeBrowser.once("targetcreated", onTargetCreated);
            setTimeout(() => {
              activeBrowser.off("targetcreated", onTargetCreated);
              resolvePromise(false);
            }, 10000);
          }),
          aeatPage.locator('input[value="Firmar Enviar"]').click(),
        ]);
      } catch (e) {
        console.log("[CERT TRIB] Error firma: ", e?.message || e);
      }

      await this.esperar(1000);
      console.log("[CERT TRIB] Descargando...");

      await aeatPage.locator(this.SELECTORS.AEAT.btnDescarga).wait();

      const rutaTrib = path.join(paths.resultados, cliente.nombreArchivoTrib);
      const nuevaPagina = await this._descargaPDFConReintento({
        browser: activeBrowser,
        botonClick: () => aeatPage.locator(this.SELECTORS.AEAT.btnDescarga).click(),
        rutaArchivo: rutaTrib,
        etiqueta: "TRIB",
        timeoutMs: 15000,
      });

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
    } finally {
      if (certBrowser) await certBrowser.close();
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

    await this._navegarConReintentos(page, "https://sede.gobiernodecanarias.org/tributos/ov/seguro/certificados/individual/listado.jsp");
    await this.esperar(1000);

    await this._procesarLoginATC(page);

    try {
      const botonSolicitar = await page.waitForSelector(
        this.SELECTORS.ATC.btnSolicitarInicial,
        { timeout: 60000 },
      );
      if (botonSolicitar) {
        await botonSolicitar.click();
      }
    } catch (_) {
      throw new Error("No se localizó el botón Solicitar inicial (ATC).");
    }

    await page.locator(this.SELECTORS.ATC.selectTipoCertificado).wait();
    await this.esperar(500);
    await page.select(this.SELECTORS.ATC.selectTipoCertificado, "AS");

    await page.locator(this.SELECTORS.ATC.radioTipoTerceros).wait();
    const radio = await page.$(this.SELECTORS.ATC.radioTipoTerceros);
    if (radio) await radio.click();

    await this.esperar(1000);

    await page.locator(this.SELECTORS.ATC.campNifTitular).wait();
    await page.type(this.SELECTORS.ATC.campNifTitular, String(cliente.nif));
    await this.esperar(500);

    await page.locator(this.SELECTORS.ATC.campNombreTitular).wait();
    await page.type(this.SELECTORS.ATC.campNombreTitular, String(cliente.empresa));
    await this.esperar(500);

    await page.locator(this.SELECTORS.ATC.btnSolicitar).wait();
    await page.locator(this.SELECTORS.ATC.btnSolicitar).click();

    if ((await page.evaluate(() => document.readyState)) !== "complete") {
      await page.waitForNavigation({ waitUntil: "load" });
    }

    console.log("[CERT ATC] Solicitud realizada y cargado");
    console.log("[CERT ATC] Descargando...");
    try {
      await page.waitForSelector(this.SELECTORS.ATC.btnDescargar, {
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
    const nuevaPagina = await this._descargaPDFConReintento({
      browser,
      botonClick: () => page.locator(this.SELECTORS.ATC.btnDescargar).click(),
      rutaArchivo: rutaATC,
      etiqueta: "ATC",
      timeoutMs: 20000,
    });

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

    await this._navegarConReintentos(page, "https://w2.seg-social.es/fs/indexframes.html");
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

    await frame.type(this.SELECTORS.ART42.campRegimen, String(cliente.ccc1));
    await frame.type(this.SELECTORS.ART42.campProvincia, String(cliente.ccc2));
    await frame.type(this.SELECTORS.ART42.campNISS, String(cliente.ccc3));

    try {
      await frame.select(this.SELECTORS.ART42.selectOpcion, "Alta");
    } catch (e) {
      throw new Error(
        `[ART42] Error seleccionando Alta en ${this.SELECTORS.ART42.selectOpcion}: ${e.message}`,
      );
    }

    await Promise.all([
      page
        .waitForNavigation({ waitUntil: "networkidle0", timeout: 20000 })
        .catch(() => {}),
      frame.click(this.SELECTORS.ART42.btnContinuar1),
    ]);
    await this.esperar(500);

    frame = getFrame();
    if (!frame)
      throw new Error("[ART42] No se encontró el frame tras Continuar 1.");

    try {
      await frame.waitForSelector(this.SELECTORS.ART42.campRegKemsoCGK, { timeout: 15000 });
    } catch (e) {
      throw new Error(`[ART42] ${this.SELECTORS.ART42.campRegKemsoCGK} no apareció: ${e.message}`);
    }

    await frame.type(this.SELECTORS.ART42.campRegKemsoCGK, empresaAutRegimen);
    await frame.type(this.SELECTORS.ART42.campTesoreriaCGK, empresaAutTesoreria);
    await frame.type(this.SELECTORS.ART42.campCuentaCGK, empresaAutCuenta);

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
      frame.click(this.SELECTORS.ART42.btnContinuar2),
    ]);
    await this.esperar(500);

    frame = getFrame();
    if (!frame)
      throw new Error("[ART42] No se encontró el frame tras Continuar 2.");

    try {
      await frame.waitForSelector(this.SELECTORS.ART42.btnConfirmar, { timeout: 15000 });
    } catch (e) {
      throw new Error(`[ART42] Botón Confirmar no apareció: ${e.message}`);
    }

    await Promise.all([
      page
        .waitForNavigation({ waitUntil: "networkidle0", timeout: 30000 })
        .catch(() => {}),
      frame.click(this.SELECTORS.ART42.btnConfirmar),
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
