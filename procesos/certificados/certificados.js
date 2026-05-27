const path = require("path");
const fs = require("fs");
const XlsxPopulate = require("xlsx-populate");
const { DateTime } = require("luxon");
const { execSync } = require("child_process");
const os = require("os");

const { registrarEjecucion, agruparPorEmpresa } = require("../../metricas");
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

  _obtenerCNcertificado(nif) {
    const scriptPath = path.join(os.tmpdir(), `cert_lookup_${Date.now()}.ps1`);
    const nifSafe = (nif || "").replace(/'/g, "''");
    const script = `[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$nif = '${nifSafe}'
$today = Get-Date
$cert = $null
$locations = @(
  [System.Security.Cryptography.X509Certificates.StoreLocation]::CurrentUser,
  [System.Security.Cryptography.X509Certificates.StoreLocation]::LocalMachine
)
foreach ($loc in $locations) {
  $store = New-Object System.Security.Cryptography.X509Certificates.X509Store('My', $loc)
  try {
    $store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadOnly)
    $found = $store.Certificates |
      Where-Object { $_.Subject -match [regex]::Escape($nif) -and $_.NotAfter -gt $today } |
      Sort-Object NotAfter -Descending |
      Select-Object -First 1
    if ($found) { $cert = $found; break }
  } finally {
    $store.Close()
  }
}
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
      `$ErrorActionPreference = 'Stop'`,
      `$kp = 'HKCU:\\Software\\Policies\\Google\\Chrome\\AutoSelectCertificateForUrls'`,
      `if (-not (Test-Path $kp)) { New-Item -Path $kp -Force | Out-Null }`,
      `Set-ItemProperty -Path $kp -Name '1' -Value '${safePolicy}'`,
    ].join("\r\n");
    // BOM UTF-8 (﻿): PowerShell 5.x lee archivos sin BOM como ANSI, corrompiendo
    // los caracteres acentuados de los CN. Con BOM los lee como UTF-8 y escribe los
    // caracteres Unicode reales al registro, que Chrome compara directamente.
    fs.writeFileSync(scriptPath, '﻿' + script, "utf8");
    try {
      execSync(`powershell -NoProfile -ExecutionPolicy Bypass -File "${scriptPath}"`, { encoding: "utf8", timeout: 30000 });
      console.log(`[POLICY] AutoSelect policy set: ${policy}`);
      return true;
    } catch (e) {
      console.warn(`[POLICY] Escritura normal fallida, intentando con elevación UAC...`);
      try {
        // Lanza el mismo script con privilegios de administrador (muestra diálogo UAC al usuario).
        // -Wait hace que este proceso espere a que el elevado termine antes de continuar.
        execSync(
          `powershell -NoProfile -ExecutionPolicy Bypass -Command "Start-Process powershell -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File \\"${scriptPath}\\"' -Verb RunAs -Wait"`,
          { encoding: "utf8", timeout: 60000 }
        );
        console.log(`[POLICY] AutoSelect policy set (elevado): ${policy}`);
        return true;
      } catch (e2) {
        console.warn(`[POLICY] No se pudo escribir la política ni con elevación (el usuario deberá seleccionar el certificado manualmente): ${e2?.message || e2}`);
        return false;
      }
    } finally {
      try { fs.unlinkSync(scriptPath); } catch (_) {}
    }
  }

  _limpiarAutoSelectPolicy() {
    const scriptPath = path.join(os.tmpdir(), `cert_policy_clean_${Date.now()}.ps1`);
    // Solo borra el valor "1" dentro de la clave, no la clave en sí.
    // Así la clave persiste entre ejecuciones y no hace falta crearla (con admin) cada vez.
    const script = `Remove-ItemProperty -Path 'HKCU:\\Software\\Policies\\Google\\Chrome\\AutoSelectCertificateForUrls' -Name '1' -ErrorAction SilentlyContinue`;
    fs.writeFileSync(scriptPath, script, "utf8");
    try {
      execSync(`powershell -NoProfile -ExecutionPolicy Bypass -File "${scriptPath}"`, { encoding: "utf8", timeout: 10000 });
    } catch (_) {} finally {
      try { fs.unlinkSync(scriptPath); } catch (_) {}
    }
  }

  // Garantiza que el usuario actual puede escribir en la clave del registro sin elevar.
  // Si no puede, hace UNA sola elevación UAC que crea la clave y concede FullControl.
  // Llamar una vez al inicio (pre-inicialización TRIB) para que el bucle de clientes
  // nunca necesite elevar.
  _garantizarPermisoRegistroCerts() {
    const kp = 'HKCU:\\Software\\Policies\\Google\\Chrome\\AutoSelectCertificateForUrls';
    const testScript = [
      `$ErrorActionPreference = 'Stop'`,
      `$kp = '${kp}'`,
      `if (-not (Test-Path $kp)) { New-Item -Path $kp -Force | Out-Null }`,
      `Set-ItemProperty -Path $kp -Name '__test__' -Value '1'`,
      `Remove-ItemProperty -Path $kp -Name '__test__' -ErrorAction SilentlyContinue`,
    ].join("\r\n");
    const testPath = path.join(os.tmpdir(), `cert_regtest_${Date.now()}.ps1`);
    fs.writeFileSync(testPath, '﻿' + testScript, "utf8");
    try {
      execSync(`powershell -NoProfile -ExecutionPolicy Bypass -File "${testPath}"`, { encoding: "utf8", timeout: 15000 });
      console.log("[POLICY] Permisos de registro OK (sin elevación).");
      return true;
    } catch (_) { /* continúa con elevación */ }
    finally { try { fs.unlinkSync(testPath); } catch (_) {} }

    console.log("[POLICY] Sin permisos directos — elevando UNA vez para configurar clave del registro...");
    const elevScript = [
      `$ErrorActionPreference = 'Stop'`,
      `$kp = '${kp}'`,
      `if (-not (Test-Path $kp)) { New-Item -Path $kp -Force | Out-Null }`,
      `$acl = Get-Acl $kp`,
      `$rule = New-Object System.Security.AccessControl.RegistryAccessRule(`,
      `  [System.Security.Principal.WindowsIdentity]::GetCurrent().Name,`,
      `  'FullControl', 'Allow'`,
      `)`,
      `$acl.SetAccessRule($rule)`,
      `Set-Acl -Path $kp -AclObject $acl`,
    ].join("\r\n");
    const elevPath = path.join(os.tmpdir(), `cert_reginit_${Date.now()}.ps1`);
    fs.writeFileSync(elevPath, '﻿' + elevScript, "utf8");
    try {
      execSync(
        `powershell -NoProfile -ExecutionPolicy Bypass -Command "Start-Process powershell -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File \\"${elevPath}\\"' -Verb RunAs -Wait"`,
        { encoding: "utf8", timeout: 60000 }
      );
      console.log("[POLICY] Clave de registro inicializada con permisos correctos.");
      return true;
    } catch (e) {
      console.warn("[POLICY] No se pudo inicializar la clave del registro:", e?.message || e);
      return false;
    } finally {
      try { fs.unlinkSync(elevPath); } catch (_) {}
    }
  }

  async certificadosSSITAATC(argumentos) {
    return this._ejecutarCertificados(argumentos, {
      habilitarSS: true, habilitarAEAT: false, habilitarATC: true, habilitarITA: true, habilitarArt42: false,
      nombreProceso: 'Certificados SS ITA ATC'
    });
  }

  async certificadoAEAT(argumentos) {
    // El formulario standalone tiene: [0]=chrome, [1]=excel, [2]=outDir, [3]=modoManual, [4]=codigosEmpresa, [5]=certTributario
    // _ejecutarCertificados espera runTrib en [6], no en [5]
    const fc = argumentos.formularioControl;
    const remapped = [
      fc[0],  // [0] chrome
      fc[1],  // [1] excel
      fc[2],  // [2] outDir
      fc[3],  // [3] modoManual
      fc[4],  // [4] codigosEmpresa
      false,  // [5] runSS (habilitarSS=false, no aplica)
      fc[5],  // [6] certTributario → runTrib
      false,  // [7] runATC
      false,  // [8] runITA
      false,  // [9] runArt42
    ];
    return this._ejecutarCertificados(
      { ...argumentos, formularioControl: remapped },
      { habilitarSS: false, habilitarAEAT: true, habilitarATC: false, habilitarITA: false, habilitarArt42: false, nombreProceso: 'Certificado AEAT' }
    );
  }

  async certificadoArt42(argumentos) {
    // El formulario standalone tiene: [0]=chrome, [1]=excel, [2]=outDir, [3]=codigoEmpresa, [4]=regimen, [5]=tesoreria, [6]=cuenta
    // _ejecutarCertificados espera:   [0]=chrome, [1]=excel, [2]=outDir, [3]=modoManual, [4]=codigosEmpresa, [9]=certArt42, [10]=regimen, [11]=tesoreria, [12]=cuenta
    const fc = argumentos.formularioControl;
    const remapped = [
      fc[0],  // [0] chrome
      fc[1],  // [1] excel
      fc[2],  // [2] outDir
      true,   // [3] modoManual (siempre manual en proceso standalone)
      fc[3],  // [4] codigosEmpresa (filtro opcional)
      false,  // [5] runSS
      false,  // [6] runTrib
      false,  // [7] runATC
      false,  // [8] runITA
      true,   // [9] certArt42 (activa el proceso)
      fc[4],  // [10] art42EmpresaRegimen
      fc[5],  // [11] art42EmpresaTesoreria
      fc[6],  // [12] art42EmpresaCuenta
    ];
    return this._ejecutarCertificados(
      { ...argumentos, formularioControl: remapped },
      { habilitarSS: false, habilitarAEAT: false, habilitarATC: false, habilitarITA: false, habilitarArt42: true, nombreProceso: 'Certificado Art 42' }
    );
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

      const chromiumExecutablePath = path.normalize(
        argumentos.formularioControl[0],
      );
      const pathArchivoEtiquetas = argumentos.formularioControl[1];
      const pathBase = argumentos.formularioControl[2];
      const modoManual = !!argumentos.formularioControl[3];
      const codigosEmpresaInput = argumentos.formularioControl[4];

      let runSS = config.habilitarSS && !!argumentos.formularioControl[5];
      let runTrib = config.habilitarAEAT && !!argumentos.formularioControl[6];
      let runATC = config.habilitarATC && !!argumentos.formularioControl[7];
      let runITA = config.habilitarITA && !!argumentos.formularioControl[8];
      let runArt42 = config.habilitarArt42 && !!argumentos.formularioControl[9];

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
            addLogCol("LOG ATC");
            addLogCol("LOG ITA");
            addLogCol("LOG AEAT");
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
                    case "EMAIL":
                      objetoCliente.email = String(cellVal || "").trim();
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
                  `${objetoCliente.codigo} CERT CORRIENTE SS ${objetoCliente.empresa} ${fechaHoy}.pdf`;
                objetoCliente.nombreArchivoTrib =
                  `${objetoCliente.codigo} CERT CORRIENTE AEAT ${objetoCliente.empresa} ${fechaHoy}.pdf`;
                objetoCliente.nombreArchivoATC =
                  `${objetoCliente.codigo} CERT CORRIENTE ATC ${objetoCliente.empresa} ${fechaHoy}.pdf`;
                objetoCliente.nombreArchivoITA =
                  `${objetoCliente.codigo} Informe ITA ${objetoCliente.ccc} ${objetoCliente.empresa} ${fechaHoy}.pdf`;
                objetoCliente.nombreArchivoArt42 =
                  `${objetoCliente.codigo} CERT CORRIENTE ART42 ${objetoCliente.ccc} ${objetoCliente.empresa} ${fechaHoy}.png`;
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

            // Si ya existe un output previo del mismo día, restaurar sus logs
            const rutaOutputExistente = path.join(carpetaRaiz, "Certificados-Procesado.xlsx");
            if (fs.existsSync(rutaOutputExistente)) {
              try {
                const wbPrevio = await XlsxPopulate.fromFileAsync(rutaOutputExistente);
                const hojaPrev = wbPrevio.sheet("BASE DE DATOS (NO TOCAR)");
                if (hojaPrev) {
                  const colsPrev = hojaPrev.usedRange()._numColumns;
                  const filasPrev = hojaPrev.usedRange()._numRows;
                  const cabPrev = [];
                  for (let i = 1; i <= colsPrev; i++) cabPrev.push(hojaPrev.cell(1, i).value());
                  const idxPrev = {};
                  cabPrev.forEach((h, i) => { if (h != null) idxPrev[String(h).trim()] = i + 1; });

                  const logCols = ["LOG SS", "LOG ATC", "LOG ITA", "LOG AEAT", "LOG ART42"];
                  const expColPrev = idxPrev["Expediente"];
                  const cccColPrev = idxPrev["Código Cuenta Cotización (CCC)"];

                  if (expColPrev && cccColPrev) {
                    const logMap = {};
                    for (let i = 2; i <= filasPrev; i++) {
                      const cod = String(hojaPrev.cell(i, expColPrev).value() || "").replace(/\D/g, "").padStart(4, "0");
                      const ccc = String(hojaPrev.cell(i, cccColPrev).value() || "").trim();
                      if (!cod || !ccc) continue;
                      logMap[`${cod}_${ccc}`] = {};
                      for (const logCol of logCols) {
                        if (idxPrev[logCol]) {
                          const v = hojaPrev.cell(i, idxPrev[logCol]).value();
                          if (v != null) logMap[`${cod}_${ccc}`][logCol] = v;
                        }
                      }
                    }

                    const expCol = colIdx["Expediente"];
                    const cccCol = colIdx["Código Cuenta Cotización (CCC)"];
                    if (expCol && cccCol) {
                      for (let i = 2; i <= filas; i++) {
                        const cod = String(hoja.cell(i, expCol).value() || "").replace(/\D/g, "").padStart(4, "0");
                        const ccc = String(hoja.cell(i, cccCol).value() || "").trim();
                        const prev = logMap[`${cod}_${ccc}`];
                        if (!prev) continue;
                        for (const logCol of logCols) {
                          if (prev[logCol] != null && colIdx[logCol]) {
                            hoja.cell(i, colIdx[logCol]).value(prev[logCol]);
                          }
                        }
                      }
                    }
                  }
                }
                console.log("[LOG] Logs previos restaurados desde archivo existente.");
              } catch (e) {
                console.warn("[LOG] No se pudo leer el archivo de output previo:", e?.message || e);
              }
            }

            const downloadPathInicial = carpetaRaiz;

            if (runSS || runTrib || runATC || runArt42) {
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
                    .cell(clientes[i].filaExcel, colIdx["LOG AEAT"])
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
                if (clientes[i].flagDupeNIF) {
                  hoja
                    .cell(clientes[i].filaExcel, colIdx["LOG SS"])
                    .value(`SKIP: ya procesado para esta empresa (NIF: ${clientes[i].nif}).`);
                } else {
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
              }

              if (clientRunTrib) {
                if (clientes[i].flagDupeNIF) {
                  hoja
                    .cell(clientes[i].filaExcel, colIdx["LOG AEAT"])
                    .value(`SKIP: ya procesado para esta empresa (NIF: ${clientes[i].nif}).`);
                } else {
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
                      .cell(clientes[i].filaExcel, colIdx["LOG AEAT"])
                      .value("ERROR: " + (e?.message || e));
                  });
                }
              }

              if (clientRunATC) {
                if (clientes[i].flagDupeNIF) {
                  hoja
                    .cell(clientes[i].filaExcel, colIdx["LOG ATC"])
                    .value(`SKIP: ya procesado para esta empresa (NIF: ${clientes[i].nif}).`);
                } else {
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
                if (clientes[i].flagDupeNIF) {
                  hoja
                    .cell(clientes[i].filaExcel, colIdx["LOG ART42"])
                    .value(`SKIP: ya procesado para esta empresa (NIF: ${clientes[i].nif}).`);
                } else {
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
              }

              console.log("Nuevo cliente");
              await this.esperar(1000);
            }

            this._limpiarAutoSelectPolicy();

            try {
              await browser.close();
            } catch (_) {}

            // Generar borradores de correo (.eml) agrupados por expediente
            try {
              const { generarEmailCertificados } = require("./emails");
              const carpetaCorreos = path.join(carpetaRaiz, "Correos");
              if (!fs.existsSync(carpetaCorreos)) fs.mkdirSync(carpetaCorreos, { recursive: true });

              const gruposPorExpediente = {};
              for (const cliente of clientes) {
                const key = String(cliente.codigo || "").trim();
                if (!key) continue;
                if (!gruposPorExpediente[key]) gruposPorExpediente[key] = [];
                gruposPorExpediente[key].push(cliente);
              }

              for (const codigo of Object.keys(gruposPorExpediente)) {
                const grupo = gruposPorExpediente[codigo];
                const emailRaw = (grupo.find((c) => c.email) || {}).email || "";
                const correos = emailRaw.split(/[;,]/).map((e) => e.trim()).filter(Boolean);
                if (correos.length === 0) continue;
                try {
                  await generarEmailCertificados(grupo, carpetaRaiz, correos, carpetaCorreos);
                } catch (e) {
                  console.warn(`[EMAIL] Error generando .eml para expediente ${codigo}:`, e?.message || e);
                }
              }
            } catch (e) {
              console.warn("[EMAIL] Error en generación de borradores:", e?.message || e);
            }

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
                empresas: agruparPorEmpresa(clientes, ["codigo"], ["empresa"]),
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
      this._garantizarPermisoRegistroCerts();
      console.log(
        "[CERT INIT] TRIB — Los certificados se seleccionarán automáticamente por empresa",
      );
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
              try {
                const pdfBuffer = await response.buffer();
                fs.writeFileSync(rutaArchivo, pdfBuffer);
                console.log(`PDF ${etiqueta} descargado en:`, rutaArchivo);
                finalizar(newPage);
              } catch (err) {
                console.warn(`PDF ${etiqueta} - error al leer buffer:`, err?.message);
                finalizar(false);
              }
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

    for (let intento = 1; intento <= 2; intento++) {
      try {
        await page.goto(
          "https://w2.seg-social.es/Xhtml?JacadaApplicationName=SGIRED&TRANSACCION=ATR64&E=I&AP=AFIR",
          { waitUntil: "networkidle0" },
        );
        break;
      } catch (e) {
        console.warn(
          `[ITA] Fallo navegación (intento ${intento}):`,
          e?.message || e,
        );
        if (intento === 2) throw e;
        await this.esperar(1500);
      }
    }
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

    await this.esperar(2000);

    try {
      const botonModal = await page.waitForSelector('button[data-dismiss="modal"]', { timeout: 2000 });
      if (botonModal) {
        await botonModal.click();
        await this.esperar(500);
      }
    } catch (_) {}

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
      await this._esperarSelector(page, `input[title="Buscar por CCC o NAF"]`, 60000, 3);
    } catch (e) {
      throw new Error(`[SS-Paso radio CCC/NAF] ${e.message}`);
    }
    const radio = await page.$(`input[title="Buscar por CCC o NAF"]`);
    if (radio) await radio.click();

    try {
      await this._esperarSelector(page, 'input[name="criteriosBusquedaCccNaf"]', 60000, 3);
    } catch (e) {
      throw new Error(`[SS-Paso campo CCC] ${e.message}`);
    }
    await page.type('input[name="criteriosBusquedaCccNaf"]', ccc);
    await this.esperar(1000);
    try {
      await Promise.all([
        page.waitForNavigation({ waitUntil: "networkidle0", timeout: 30000 }).catch(() => {}),
        page.locator('button[name="SPM.ACC.AC_BUSCAR_OAR"]').click(),
      ]);
    } catch (e) {
      throw new Error(`[SS-Paso botón buscar CCC] ${e.message}`);
    }

    const selectorResultado = 'a[id="enlace_' + String(Number(ccc)) + '"]';
    const enlaceResultado = await page
      .waitForSelector(selectorResultado, { timeout: 30000 })
      .catch(() => null);
    if (!enlaceResultado) {
      try {
        const screenshotPath = path.join(paths.resultados, `SS_debug_${ccc}_${Date.now()}.png`);
        await page.screenshot({ path: screenshotPath, fullPage: true });
        console.warn(`[CERT SS] Screenshot de diagnóstico guardado en: ${screenshotPath}`);
      } catch (_) {}
      throw new Error("CCC no encontrado en el sistema ARED: " + ccc);
    }
    await Promise.all([
      page.waitForNavigation({ waitUntil: "networkidle0", timeout: 30000 }).catch(() => {}),
      enlaceResultado.click(),
    ]);

    await page.locator('button[name="SPM.ACC.CONTINUAR"]').wait();
    await page.locator('button[name="SPM.ACC.CONTINUAR"]').click();
    await page.locator('button[name="SPM.ACC.IMPRIMIR"]').wait();
    await Promise.all([
      page.waitForNavigation({ waitUntil: "load" }),
      page.locator('button[name="SPM.ACC.IMPRIMIR"]').click(),
    ]);

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
        .cell(cliente.filaExcel, colIdx["LOG AEAT"])
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
        .cell(cliente.filaExcel, colIdx["LOG AEAT"])
        .value(`ERROR: No se encontró certificado en almacén Windows para NIF ${cliente.nif}`);
      return;
    }

    console.log(`[CERT TRIB] Certificado encontrado: CN="${certInfo.subjectCN}", ISSUER="${certInfo.issuerCN}"`);

    // Configurar auto-selección Chrome por CN real del cert
    const policyOk = this._setAutoSelectPolicy(certInfo);
    if (!policyOk) {
      hoja
        .cell(cliente.filaExcel, colIdx["LOG AEAT"])
        .value(`WARNING: Selecciona manualmente el certificado "${certInfo.subjectCN}" en el diálogo de Chrome.`);
    }
    const certBrowser = await puppeteer.launch({ executablePath, headless: false });
    const aeatPage = await certBrowser.newPage();
    await aeatPage.setViewport({ width: 1080, height: 1024 });
    aeatPage.setDefaultTimeout(60000);
    aeatPage.on("dialog", async (dialog) => {
      try { await dialog.accept(); } catch (_) {}
    });
    const activeBrowser = certBrowser;

    try {
      for (let intento = 1; intento <= 2; intento++) {
        try {
          await aeatPage.goto(
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
        const botonModal = await aeatPage.waitForSelector(
          'button[data-dismiss="modal"]',
          { timeout: 1000 },
        );
        if (botonModal) {
          await botonModal.click();
        }
      } catch (_) {}

      await aeatPage.locator(`input[id="fTipoRepresentacion0"]`).wait();
      const radio1 = await aeatPage.$(`input[id="fTipoRepresentacion0"]`);
      if (radio1) await radio1.click();
      await this.esperar(500);

      await aeatPage.locator(`input[id="fTipoCertificado4"]`).wait();
      const radio2 = await aeatPage.$(`input[id="fTipoCertificado4"]`);
      if (radio2) await radio2.click();

      await aeatPage.locator('input[id="validarSolicitud"]').click();
      await aeatPage.waitForNavigation({ waitUntil: "load" });

      await aeatPage.locator('input[value="Firmar Enviar"]').wait();

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

      await aeatPage.locator('input[id="descarga"]').wait();

      const rutaTrib = path.join(paths.resultados, cliente.nombreArchivoTrib);
      let nuevaPagina = await this._descargarPDF({
        browser: activeBrowser,
        botonClick: () => aeatPage.$('input[id="descarga"]').then(el => el?.click()),
        rutaArchivo: rutaTrib,
        etiqueta: "TRIB",
        timeoutMs: 15000,
      });

      if (!nuevaPagina) {
        console.log("[CERT TRIB] Reintentando descarga...");
        await this.esperar(3000);
        nuevaPagina = await this._descargarPDF({
          browser: activeBrowser,
          botonClick: () => aeatPage.$('input[id="descarga"]').then(el => el?.click()),
          rutaArchivo: rutaTrib,
          etiqueta: "TRIB",
          timeoutMs: 15000,
        });
      }

      if (!nuevaPagina) {
        console.log("[CERT TRIB] ERROR EN DESCARGA");
        hoja
          .cell(cliente.filaExcel, colIdx["LOG AEAT"])
          .value("ERROR: No se ha podido generar el resguardo de la solicitud.");
      } else {
        hoja
          .cell(cliente.filaExcel, colIdx["LOG AEAT"])
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
    await this.esperar(200);

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

    frame = getFrame();
    if (!frame) {
      throw new Error(
        "[ART42] No se encontró el iframe tras expandir Gestión de Deuda.",
      );
    }
    await frame.waitForFunction(
      () => Array.from(document.querySelectorAll("a")).some((a) =>
        a.textContent.includes("Autorización Certificado Art.42"),
      ),
      { timeout: 10000 },
    );
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

    frame = getFrame();
    if (!frame)
      throw new Error("[ART42] No se encontró el frame del formulario.");

    try {
      await frame.waitForSelector("#SDFREGIMEN", { timeout: 15000 });
    } catch (e) {
      throw new Error(`[ART42] #SDFREGIMEN no apareció: ${e.message}`);
    }

    await frame.evaluate((ccc1, ccc2, ccc3) => {
      const set = (sel, val) => { const el = document.querySelector(sel); if (el) el.value = val; };
      set('#SDFREGIMEN', ccc1);
      set('#SDFPROVINCIA', ccc2);
      set('#SDFNISS', ccc3);
      set('#SDFOPCION', 'Alta');
    }, String(cliente.ccc1), String(cliente.ccc2), String(cliente.ccc3));

    await frame.click("#Sub2207001004_35");

    frame = getFrame();
    if (!frame)
      throw new Error("[ART42] No se encontró el frame tras Continuar 1.");

    try {
      await frame.waitForSelector("#SDFREGKCGK", { timeout: 15000 });
    } catch (e) {
      throw new Error(`[ART42] #SDFREGKCGK no apareció: ${e.message}`);
    }

    const ahora = DateTime.now().setZone("Europe/Madrid");
    const hasta = ahora.plus({ years: 1 });

    await frame.evaluate((reg, tes, cta, diaD, mesD, anyoD, diaH, mesH, anyoH) => {
      const set = (sel, val) => { const el = document.querySelector(sel); if (el) el.value = val; };
      set('#SDFREGKCGK', reg);
      set('#SDFTESCCGK', tes);
      set('#SDFCCONCGK9', cta);
      set('#SDFDIADESDE', diaD);
      set('#SDFMESDESDE', mesD);
      set('#SDFAODESDE', anyoD);
      set('#SDFDIAHASTA', diaH);
      set('#SDFMESHASTA', mesH);
      set('#SDFAOHASTA', anyoH);
    },
      empresaAutRegimen, empresaAutTesoreria, empresaAutCuenta,
      ahora.toFormat("dd"), ahora.toFormat("MM"), ahora.toFormat("yyyy"),
      hasta.toFormat("dd"), hasta.toFormat("MM"), hasta.toFormat("yyyy"),
    );

    await frame.click("#Sub2207001004_75");

    frame = getFrame();
    if (!frame)
      throw new Error("[ART42] No se encontró el frame tras Continuar 2.");

    try {
      await frame.waitForSelector("#Sub2204701006_74", { timeout: 15000 });
    } catch (e) {
      throw new Error(`[ART42] Botón Confirmar no apareció: ${e.message}`);
    }

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

    await frame.click("#Sub2204701006_74");

    hoja.cell(cliente.filaExcel, colIdx["LOG ART42"]).value("OK, autorización generada.");
  }
}

module.exports = ProcesosCertificados;
