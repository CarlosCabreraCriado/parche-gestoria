const path = require("path");
const fs = require("fs");
const fsExtra = require("fs-extra");
const pdf = require("pdf-parse");
const pdfLib = require("pdf-lib");
const readline = require("readline");
const axios = require("axios");
const moment = require("moment");
const XlsxPopulate = require("xlsx-populate");
const Datastore = require("nedb");
const _ = require("lodash");
const { DateTime } = require("luxon");

const { ipcRenderer } = require("electron");
const puppeteer = require("puppeteer");

class ProcesosAsesoria {
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

  async extractInfoFromPdf(filePath) {
    /*
    const dataBuffer = await fsExtra.readFile(filePath);
    const data = await pdf(dataBuffer);

    const cccMatch = data.text.match(
      /Código de Cuenta de Cotización:\s*([\d\s]+)/,
    );

    console.log(cccMatch);
    if (!cccMatch) return null;
    const raw = cccMatch[1].replace(/\s+/g, "");
    const ccc = raw.slice(-11); // Últimos 11 dígitos

    console.log("CCC: ", ccc);
    const hasVacaciones = /Calificador de Liquidación:\s*L13/i.test(data.text);
        */

    const dataBuffer = await fsExtra.readFile(filePath);
    const data = await pdf(dataBuffer);
    const lines = data.text.split("\n").map((line) => line.trim());

    let ccc = null;
    let hasVacaciones = false;
    for (let i = 0; i < lines.length; i++) {
      if (lines[i].startsWith("Código de Cuenta de Cotización:")) {
        const nextLine = lines[i + 4]?.replace(/\s+/g, "") || "";
        const lineaVacaciones = lines[i + 5] || "";

        if (lineaVacaciones.includes("L13")) {
          hasVacaciones = true;
        }

        if (/^\d{4}\d{11}$/.test(nextLine)) {
          ccc = nextLine.slice(-11); // Tomamos los últimos 11 dígitos
        }
        break;
      }
    }

    console.log("CCC: ", ccc);
    console.log("Vacaciones: ", hasVacaciones);
    if (!ccc) return null;

    return { filePath, ccc, hasVacaciones };
  }

  async formatearRecibosDeLiquidacion(argumentos) {
    return new Promise((resolve) => {
      console.log("Excel de datos de clientes...");
      console.log(argumentos.formularioControl[0]);

      var archivoLiquidacion = {};
      var empresas = [];
      var autonomos = [];

      var pathArchivoLiquidacion = argumentos.formularioControl[0];
      var pathSalidaExcel = path.join(
        path.normalize(argumentos.formularioControl[1]),
        "Liquidacion-Procesado",
      );

      var pathPDFs = path.normalize(argumentos.formularioControl[1]);
      var pathGuardarPDFs = path.join(
        path.normalize(argumentos.formularioControl[1]),
        "Liquidacion-Procesado",
        "Resultados",
      );

      // Verificar si la carpeta "Resultados" existe y crearla si no
      if (!fs.existsSync(pathGuardarPDFs)) {
        fs.mkdirSync(pathGuardarPDFs, { recursive: true });
        console.log(`Carpeta creada: ${pathGuardarPDFs}`);
      } else {
        console.log(`La carpeta ya existe: ${pathGuardarPDFs}`);
      }

      try {
        XlsxPopulate.fromFileAsync(path.normalize(pathArchivoLiquidacion))
          .then(async (workbook) => {
            console.log("Archivo Cargado: Liquidacion");
            archivoLiquidacion = workbook;

            //Procesamiento de HOJA DATOS:
            const columnasEmpresas = archivoLiquidacion
              .sheet("DATOS")
              .usedRange()._numColumns;
            const filasEmpresas = archivoLiquidacion
              .sheet("DATOS")
              .usedRange()._numRows;

            var objetoEmpresa = {};
            var cabecerasEmpresas = [];
            for (var i = 1; i <= columnasEmpresas; i++) {
              cabecerasEmpresas.push(
                archivoLiquidacion.sheet("DATOS").cell(1, i).value(),
              );
            }

            console.log("Cabeceras EMPRESAS: " + cabecerasEmpresas);
            for (var i = 2; i <= filasEmpresas; i++) {
              objetoEmpresa = {};
              for (var j = 1; j <= columnasEmpresas; j++) {
                if (
                  archivoLiquidacion.sheet("DATOS").cell(i, j).value() !==
                  undefined
                ) {
                  switch (cabecerasEmpresas[j - 1]) {
                    case "CÓDIGO":
                      objetoEmpresa["codigo"] = archivoLiquidacion
                        .sheet("DATOS")
                        .cell(i, j)
                        .value();
                      break;
                    case "EMPRESA":
                      objetoEmpresa["empresa"] = archivoLiquidacion
                        .sheet("DATOS")
                        .cell(i, j)
                        .value();
                      break;
                    case "CCC":
                      objetoEmpresa["ccc"] = archivoLiquidacion
                        .sheet("DATOS")
                        .cell(i, j)
                        .value();
                      break;
                  }
                }
              }
              objetoEmpresa["errores"] = [];
              empresas.push(Object.assign({}, objetoEmpresa));
            }

            //Procesamiento de HOJA AUTONOMOS:
            const columnasAutonomos = archivoLiquidacion
              .sheet("AUTONOMOS")
              .usedRange()._numColumns;
            const filasAutonomos = archivoLiquidacion
              .sheet("DATOS")
              .usedRange()._numRows;

            var objetoAutonomos = {};
            var cabecerasAutonomos = [];
            for (var i = 1; i <= columnasAutonomos; i++) {
              cabecerasAutonomos.push(
                archivoLiquidacion.sheet("AUTONOMOS").cell(1, i).value(),
              );
            }

            console.log("Cabeceras AUTONOMOS: " + cabecerasAutonomos);
            for (var i = 2; i <= filasAutonomos; i++) {
              objetoAutonomos = {};
              for (var j = 1; j <= columnasAutonomos; j++) {
                if (
                  archivoLiquidacion.sheet("AUTONOMOS").cell(i, j).value() !==
                  undefined
                ) {
                  switch (cabecerasAutonomos[j - 1]) {
                    case "CÓDIGO":
                      objetoAutonomos["codigo"] = archivoLiquidacion
                        .sheet("AUTONOMOS")
                        .cell(i, j)
                        .value();
                      break;
                    case "AUTONOMOS":
                      objetoAutonomos["autonomo"] = archivoLiquidacion
                        .sheet("AUTONOMOS")
                        .cell(i, j)
                        .value();
                      break;
                    case "DNI":
                      objetoAutonomos["dni"] = archivoLiquidacion
                        .sheet("AUTONOMOS")
                        .cell(i, j)
                        .value();
                      break;
                    case "CCC":
                      objetoAutonomos["ccc"] = archivoLiquidacion
                        .sheet("AUTONOMOS")
                        .cell(i, j)
                        .value();
                      break;
                  }
                }
              }
              objetoAutonomos["errores"] = [];
              autonomos.push(Object.assign({}, objetoAutonomos));
            }

            console.log("Empresas: ", empresas);
            console.log("Autonomos: ", autonomos);

            //PROCESAMIENTO DE PDFs:

            console.log("PROCESANDO PDFs: ", pathPDFs);

            const files = await fsExtra.readdir(pathPDFs);

            console.log("PDFs: ", files);
            const pdfFiles = files.filter((file) =>
              file.toLowerCase().endsWith(".pdf"),
            );

            console.log("PDFs encontrados: ", pdfFiles);

            const groupedByCCC = {};

            for (const file of pdfFiles) {
              const fullPath = path.join(pathPDFs, file);
              const info = await this.extractInfoFromPdf(fullPath);

              if (!info) continue;

              const codigoEmpresa = obtenerCodigoEmpresaPorCCC(
                info.ccc,
                empresas,
                autonomos,
              );
              info["codigo"] = codigoEmpresa;

              if (!groupedByCCC[info["codigo"]])
                groupedByCCC[info["codigo"]] = [];
              groupedByCCC[info["codigo"]].push(info);
            }

            console.log("Agrupados: ", groupedByCCC);

            for (const codigo in groupedByCCC) {
              const files = groupedByCCC[codigo];
              const outputPdf = await pdfLib.PDFDocument.create();

              let hasVacaciones = false;

              // Separar primero los archivos
              const sinVacaciones = files.filter((f) => !f.hasVacaciones);
              const conVacaciones = files.filter((f) => f.hasVacaciones);

              // Primero añadimos los que NO tienen vacaciones
              for (const { filePath } of sinVacaciones) {
                const pdfBytes = await fsExtra.readFile(filePath);
                const pdfDoc = await pdfLib.PDFDocument.load(pdfBytes);
                const copiedPages = await outputPdf.copyPages(
                  pdfDoc,
                  pdfDoc.getPageIndices(),
                );
                copiedPages.forEach((page) => outputPdf.addPage(page));
              }

              // Luego añadimos los que SÍ tienen vacaciones
              for (const { filePath } of conVacaciones) {
                const pdfBytes = await fsExtra.readFile(filePath);
                const pdfDoc = await pdfLib.PDFDocument.load(pdfBytes);
                const copiedPages = await outputPdf.copyPages(
                  pdfDoc,
                  pdfDoc.getPageIndices(),
                );
                copiedPages.forEach((page) => outputPdf.addPage(page));
                hasVacaciones = true;
              }

              //Buscar codigo de empresa o autonomo:
              let año = new Date().getFullYear();
              const nombreMesAnterior = obtenerNombreMesAnterior();

              const pageCount = outputPdf.getPageCount();
              if (pageCount > 1) {
                hasVacaciones = false;
              }

              const fileName = `${codigo} TC1 ${nombreMesAnterior} ${año}${hasVacaciones ? " Vacaciones" : ""}.pdf`;
              const outputPath = path.join(pathGuardarPDFs, fileName);
              const finalPdfBytes = await outputPdf.save();
              await fsExtra.writeFile(outputPath, finalPdfBytes);
              console.log(`Archivo generado: ${outputPath}`);
            }

            function obtenerNombreMesAnterior() {
              const meses = [
                "enero",
                "febrero",
                "marzo",
                "abril",
                "mayo",
                "junio",
                "julio",
                "agosto",
                "septiembre",
                "octubre",
                "noviembre",
                "diciembre",
              ];

              const ahora = new Date();
              const mesAnterior =
                ahora.getMonth() === 0 ? 11 : ahora.getMonth() - 1;

              return meses[mesAnterior];
            }

            function obtenerCodigoEmpresaPorCCC(ccc, empresas, autonomos) {
              const empresa = empresas.find((e) => e.ccc === Number(ccc));
              if (empresa) return empresa.codigo;

              const autonomo = autonomos.find((a) => a.ccc === Number(ccc));
              if (autonomo) return autonomo.codigo;

              return null;
            }

            resolve(true);
          })
          .then(() => {})
          .catch((err) => {
            console.log("ERROR");

            throw err;
          });
      } catch (err) {
        var tituloError = "No se ha podido cargar el archivo";
        var mensajeError =
          "Se ha producido un error interno cargando los archivos.";
        mainProcess.mostrarError(tituloError, mensajeError).then((result) => {
          resolve(false);
        });
      }
    }).catch((err) => {
      console.log("Se ha producido un error interno: ");
      console.log(err);
      var tituloError = "No se ha podido cargar el archivo";
      var mensajeError =
        "Se ha producido un error interno cargando los archivos.";
      mainProcess.mostrarError(tituloError, mensajeError).then((result) => {
        resolve(false);
      });
    });
  }

  async iRPF2024(argumentos) {
    return new Promise((resolve) => {
      console.log("Calculo de IRPF...");
      console.log(argumentos.formularioControl[1]);
      console.log("Ruta Google...");
      console.log(argumentos.formularioControl[0]);

      var archivoIRPF = {};
      var clientes = [];
      var pathArchivoIRPF = argumentos.formularioControl[1];
      var pathSalidaExcel = path.join(
        path.normalize(argumentos.formularioControl[2]),
        "IRPF-Procesado",
      );
      var pathSalida = path.join(
        path.normalize(argumentos.formularioControl[2]),
        "IRPF-Procesado",
        "Resultados",
      );

      // Verificar si la carpeta "Resultados" existe y crearla si no
      if (!fs.existsSync(pathSalida)) {
        fs.mkdirSync(pathSalida, { recursive: true });
        console.log(`Carpeta creada: ${pathSalida}`);
      } else {
        console.log(`La carpeta ya existe: ${pathSalida}`);
      }

      try {
        XlsxPopulate.fromFileAsync(path.normalize(pathArchivoIRPF))
          .then(async (workbook) => {
            console.log("Archivo Cargado: IRPF");
            archivoIRPF = workbook;
            var columnas = archivoIRPF.sheet(0).usedRange()._numColumns;

            var filas = archivoIRPF.sheet(0).usedRange()._numRows;

            var objetoCliente = {};

            var cabeceras = [];
            for (var i = 1; i <= columnas; i++) {
              cabeceras.push(archivoIRPF.sheet(0).cell(2, i).value());
            }

            console.log("Cabeceras: " + cabeceras);

            for (var i = 3; i <= filas; i++) {
              objetoCliente = {};
              for (var j = 1; j <= columnas; j++) {
                if (archivoIRPF.sheet(0).cell(i, j).value() !== undefined) {
                  switch (cabeceras[j - 1]) {
                    case "Emp->Código_de_la_Empresa":
                      objetoCliente["cod_empresa"] = archivoIRPF
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;
                    case "Emp->Nombre_de_la_Empresa":
                      objetoCliente["nombre_empresa"] = archivoIRPF
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;
                    case "Trab->Código_del_Trabajador":
                      objetoCliente["cod_trabajador"] = archivoIRPF
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;
                    case "Trab->DNI_del_Trabajador":
                      objetoCliente["dni_trabajador"] = archivoIRPF
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;
                    case "Trab->Apellidos_y_Nombre_del_Trabajador":
                      objetoCliente["nombre_trabajador"] = archivoIRPF
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;
                    case "Trab->Número_de_hijos":
                      objetoCliente["num_hijos"] = archivoIRPF
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;
                    case "Trab->Porcentaje_retención":
                      objetoCliente["porcentaje_retencion"] = archivoIRPF
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;
                    case "Trab->Tipo_de_retención":
                      objetoCliente["tipo_retencion"] = archivoIRPF
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;
                    case "Trab->Ingresos_anuales":
                      objetoCliente["ingresos_anuales"] = archivoIRPF
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;
                    case "Trab->IRPF_Grado_Discapacidad":
                      objetoCliente["grado_discapacidad"] = archivoIRPF
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;
                    case "Trab->Tipo_Contrato_(3_posiciones)":
                      objetoCliente["tipo_contrato"] = archivoIRPF
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;
                    case "Trab->Edad_Trabajador":
                      objetoCliente["edad_trabajador"] = archivoIRPF
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;
                    case "Trab->Fecha_Nacimiento_(AAAA/MM/DD)":
                      objetoCliente["fecha_nacimiento"] = archivoIRPF
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;
                    case "Trab->Situación_Familiar":
                      objetoCliente["situacion_familiar"] = archivoIRPF
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;
                    case "Trab->DNI_Conyuge":
                      objetoCliente["dni_conyuge"] = archivoIRPF
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;
                    case "Trab->Año_Nacimiento_Hijo_01":
                      objetoCliente["anio_nacimiento_hijo_01"] = archivoIRPF
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;
                    case "Trab->Año_Nacimiento_Hijo_02":
                      objetoCliente["anio_nacimiento_hijo_02"] = archivoIRPF
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;
                    case "Trab->Año_Nacimiento_Hijo_03":
                      objetoCliente["anio_nacimiento_hijo_03"] = archivoIRPF
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;
                    case "Trab->Indicador_Adquisición_Vivienda":
                      objetoCliente["adquisicion_vivienda"] = archivoIRPF
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;
                    case "Trab->Pensiones_Compensatorias_Cónyuge":
                      objetoCliente["pension_conyuge"] = archivoIRPF
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;
                    case "Trab->Anualidades_en_Favor_de_los_Hijos":
                      objetoCliente["anualidades_hijos"] = archivoIRPF
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;
                    case "Trab->Sumatorio_015_de_conceptos_de_paga":
                      objetoCliente["sumatorio_015"] = archivoIRPF
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;
                    case "Trab->Sumatorio_016_de_conceptos_de_paga":
                      objetoCliente["sumatorio_016"] = archivoIRPF
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;
                    case "Trab->Sumatorio_017_de_conceptos_de_paga":
                      objetoCliente["sumatorio_017"] = archivoIRPF
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;
                  }
                }
              }

              objetoCliente["errores"] = [];

              if (
                objetoCliente.dni_trabajador !== "" &&
                objetoCliente.dni_trabajador !== null &&
                objetoCliente.dni_trabajador !== undefined
              ) {
                clientes.push(Object.assign({}, objetoCliente));
              }
            }

            var chromiumExecutablePath = path.normalize(
              argumentos.formularioControl[0],
            );

            //Inicio de procesamiento:
            const browser = await puppeteer.launch({
              executablePath: chromiumExecutablePath,
              headless: false,
            });

            var page = await browser.newPage();

            // Configurar el comportamiento de descarga
            await page._client().send("Page.setDownloadBehavior", {
              behavior: "allow",
              downloadPath: pathSalida,
            });

            await page.setViewport({ width: 1080, height: 1024 });

            var hoy = new Date();
            for (var i = 0; i < clientes.length; i++) {
              //Recargar cada 10 clientes:
              if (i % 10 == 0 && i > 0) {
                //await browser.close();
                await page.close();
                page = await browser.newPage();

                // Configurar el comportamiento de descarga
                await page._client().send("Page.setDownloadBehavior", {
                  behavior: "allow",
                  downloadPath: pathSalida,
                });
                await page.setViewport({ width: 1080, height: 1024 });
              }

              console.log("Procesando cliente: " + i);
              console.log(clientes[i]);

              if (
                clientes[i].dni_trabajador == "" ||
                clientes[i].dni_trabajador == null ||
                clientes[i].dni_trabajador == undefined
              ) {
                clientes[i]["errores"] = ["DNI del trabajador no definido."];
                continue;
              }

              await page.goto(
                "https://prewww2.aeat.es/wlpl/PRET-R200/R242/index.zul",
                { waitUntil: "networkidle0" },
              );

              //Procesado:

              //********
              // DNI
              //********
              await page.locator('input[title="NIF del perceptor"]').wait();
              await page.type(
                'input[title="NIF del perceptor"]',
                String(clientes[i].dni_trabajador),
              );

              //********
              // AÑO DE NACIMIENTO
              //********
              var anioNacimiento = clientes[i].fecha_nacimiento.slice(-4);
              await page.locator('input[title="Año de nacimiento"]').wait();
              await page.type(
                'input[title="Año de nacimiento"]',
                anioNacimiento,
              );

              //********
              //Seleccion de discapacidad:
              //********
              var spanSelector = 'span[title="Sin discapacidad"]';

              if (
                clientes[i].grado_discapacidad == "" ||
                clientes[i].grado_discapacidad == null ||
                clientes[i].grado_discapacidad == undefined
              ) {
                spanSelector = 'span[title="Sin discapacidad"]';
              } else if (clientes[i].grado_discapacidad >= 65) {
                spanSelector = 'span[title="Superior o igual al 65%"]';
              } else if (clientes[i].grado_discapacidad >= 33) {
                spanSelector =
                  'span[title="Superior o igual al 33% e inferior al 65%"]';
              }

              await page.locator(`${spanSelector} input[type="radio"]`).wait();
              var radioButton = await page.$(
                `${spanSelector} input[type="radio"]`,
              );

              if (radioButton) {
                await radioButton.click(); // Hacer clic en el radio button
                console.log("Radio button seleccionado.");
              } else {
                console.log("No se encontró el radio button.");
              }

              //********
              //Seleccion situacion familiar:
              //********
              var spanSelector = "";

              switch (clientes[i].situacion_familiar) {
                case "Soltero,divorciado,v":
                  spanSelector = `span[title='Situación 1: Soltero/a, viudo/a, divorciado/a o separado/a legalmente, con hijos solteros menores de 18 años o incapacitados judicialmente que convivan exclusivamente con el perceptor, sin convivir también con el otro progenitor, siempre que proceda consignar al menos un hijo o descendiente en el apartado "Ascendientes y  Descendientes"']`;
                  break;
                case "Conyuge a Cargo":
                  spanSelector =
                    'span[title="Situación 2: Perceptor casado y no separado legalmente cuyo cónyuge no obtenga rentas superiores a 1.500 euros anuales, excluidas las exentas."]';
                  break;
                case "Sin conyuge a Cargo":
                  spanSelector =
                    'span[title="Situación 3: Perceptor cuya situación familiar es distinta de las dos anteriores (v. gr.: solteros sin hijos; casados cuyo cónyuge obtiene rentas superiores a 1.500 euros anuales, excluidas las exentas, etc.).También se marcará esta casilla cuando el perceptor no desee manifestar su situación familiar"]';
                  break;
              }

              await page.locator(`${spanSelector} input[type="radio"]`).wait();
              await page.locator(`${spanSelector}`).click();

              //Si hay conyuge a cargo pone su DNI:
              if (clientes[i].situacion_familiar == "Conyuge a Cargo") {
                await page.locator('input[title="NIF del cónyuge"]').wait();
                await page.type(
                  'input[title="NIF del cónyuge"]',
                  clientes[i].dni_conyuge,
                );
              }

              //********************
              // TIPO CONTRATO:
              //********************
              spanSelector = 'span[title="General"]';

              if (clientes[i].tipo_contrato >= 300) {
                spanSelector =
                  'span[title="Duración inferior al año o relación laboral especial de las personas artistas que desarrollan actividades escénicas, audiovisuales y musicales, y de quienes realizan actividades técnicas o auxiliares necesarias para el desarrollo de dicha actividad (excepto relaciones esporádicas: peonadas y jornales diarios)."]';
              }

              await page.locator(`${spanSelector} input[type="radio"]`).wait();
              await page.locator(`${spanSelector}`).click();

              // ******************
              // DATOS ASCENDIENTES / DESCENDIENTES:
              // ******************

              if (
                clientes[i].anio_nacimiento_hijo_01 ||
                clientes[i].anio_nacimiento_hijo_02 ||
                clientes[i].anio_nacimiento_hijo_03
              ) {
                await page
                  .locator("span ::-p-text('Ascendientes y descendientes')")
                  .wait();
                await page
                  .locator("span ::-p-text('Ascendientes y descendientes')")
                  .click();

                //Hijo 01:
                if (
                  clientes[i].anio_nacimiento_hijo_01 &&
                  hoy.getFullYear() -
                    Number(clientes[i].anio_nacimiento_hijo_01) <
                    25
                ) {
                  await page.locator(".z-icon-user-plus").wait();
                  await page.locator(".z-icon-user-plus").click();

                  await page
                    .locator('[role="dialog"] input[title="Año de nacimiento"]')
                    .wait();
                  await page.type(
                    '[role="dialog"] input[title="Año de nacimiento"]',
                    String(clientes[i].anio_nacimiento_hijo_01),
                  );

                  await page.locator("button ::-p-text(' Aceptar')").wait();
                  await page.locator("button ::-p-text(' Aceptar')").click();
                  await page.waitForSelector('[role="dialog"]', {
                    hidden: true,
                  });
                }

                //Hijo 02:
                if (
                  clientes[i].anio_nacimiento_hijo_02 &&
                  hoy.getFullYear() -
                    Number(clientes[i].anio_nacimiento_hijo_02) <
                    25
                ) {
                  await page.locator(".z-icon-user-plus").wait();
                  await page.locator(".z-icon-user-plus").click();

                  await page
                    .locator('[role="dialog"] input[title="Año de nacimiento"]')
                    .wait();
                  await page.type(
                    '[role="dialog"] input[title="Año de nacimiento"]',
                    String(clientes[i].anio_nacimiento_hijo_02),
                  );

                  await page.locator("button ::-p-text(' Aceptar')").click();
                  await page.waitForSelector('[role="dialog"]', {
                    hidden: true,
                  });
                }

                //Hijo 03:
                if (
                  clientes[i].anio_nacimiento_hijo_03 &&
                  hoy.getFullYear() -
                    Number(clientes[i].anio_nacimiento_hijo_03) <
                    25
                ) {
                  await page.locator(".z-icon-user-plus").wait();
                  await page.locator(".z-icon-user-plus").click();

                  await page
                    .locator('[role="dialog"] input[title="Año de nacimiento"]')
                    .wait();
                  await page.type(
                    '[role="dialog"] input[title="Año de nacimiento"]',
                    String(clientes[i].anio_nacimiento_hijo_03),
                  );

                  await page.locator("button ::-p-text(' Aceptar')").click();
                  await page.waitForSelector('[role="dialog"]', {
                    hidden: true,
                  });
                }
              } //Fin ascentientes y descendientes.

              // ******************
              // DATOS ECONOMICOS:
              // ******************
              await page.locator("span ::-p-text('Datos económicos')").wait();
              await page.locator("span ::-p-text('Datos económicos')").click();
              await page
                .locator(
                  'input[title="Retribuciones totales (dinerarias y en especie)."]',
                )
                .wait();
              await page
                .locator(
                  'input[title="Gastos deducibles (Art. 19.2, letras a, b y c de la LIRPF: Seguridad Social, Mutualidades de funcionarios, derechos pasivos, colegios de huérfanos o instituciones similares)"]',
                )
                .wait();
              await page
                .locator(
                  'input[title="Pensión compensatoria a favor del cónyuge. Importe fijado judicialmente"]',
                )
                .wait();
              await page
                .locator(
                  'input[title="Anualidades por alimentos en favor de los hijos. Importe fijado judicialmente"]',
                )
                .wait();
              await page
                .locator(
                  'span[title="El perceptor ha comunicado en el modelo 145 que está efectuando pagos por préstamos destinados a la adquisición o rehabilitación de su vivienda habitual por los que va a tener derecho a deducción por inversión en vivienda habitual en el IRPF y que la suma de los rendimientos íntegros del trabajo procedentes de todos sus pagadores es inferior a 33.007,20 euros anuales."]',
                )
                .wait();

              if (clientes[i].sumatorio_015) {
                await page.type(
                  'input[title="Retribuciones totales (dinerarias y en especie)."]',
                  String(clientes[i].sumatorio_015),
                );
              }

              if (clientes[i].sumatorio_017) {
                await page.type(
                  'input[title="Gastos deducibles (Art. 19.2, letras a, b y c de la LIRPF: Seguridad Social, Mutualidades de funcionarios, derechos pasivos, colegios de huérfanos o instituciones similares)"]',
                  String(clientes[i].sumatorio_017),
                );
              }

              if (clientes[i].pension_conyuge) {
                await page.type(
                  'input[title="Pensión compensatoria a favor del cónyuge. Importe fijado judicialmente"]',
                  String(clientes[i].pension_conyuge),
                );
              }

              if (clientes[i].anualidades_hijos) {
                clientes[i].anualidades_hijos =
                  parseFloat(clientes[i].anualidades_hijos) / 12;

                await page.type(
                  'input[title="Anualidades por alimentos en favor de los hijos. Importe fijado judicialmente"]',
                  String(clientes[i].anualidades_hijos),
                );
              }

              if (clientes[i].adquisicion_vivienda == "Destina (ant.2010)") {
                if (clientes[i].sumatorio_015 < 33007.2) {
                  await page
                    .locator(
                      'span[title="El perceptor ha comunicado en el modelo 145 que está efectuando pagos por préstamos destinados a la adquisición o rehabilitación de su vivienda habitual por los que va a tener derecho a deducción por inversión en vivienda habitual en el IRPF y que la suma de los rendimientos íntegros del trabajo procedentes de todos sus pagadores es inferior a 33.007,20 euros anuales."]',
                    )
                    .click();
                } else {
                  clientes[i]["errores"].push(
                    "WARN: Ingresos superiores a 33.007,20 euros anuales. Omitiendo deducción por vivienda habitual.",
                  );
                }
              }

              // ******************
              // RESULTADOS:
              // ******************

              if (!clientes[i].sumatorio_017) {
                clientes[i]["errores"].push(
                  "ERROR: Faltan datos de sumatorio_017",
                );
                await page.reload();
                continue;
              }
              await page.locator("span ::-p-text('Resultados')").wait();
              await page.locator("span ::-p-text('Resultados')").click();

              await this.esperar(2000);

              const found = await page.evaluate(() => {
                const div = document.querySelector("div");
                return div && div.textContent.includes("Relación de errores");
              });

              if (found) {
                console.log("ERROR EN EL PROCESAMIENTO", i);

                await this.esperar(2000);

                var errores = await page.$$eval(".z-label", (spans) =>
                  spans.map((span) => span.textContent.trim()),
                );

                clientes[i]["errores"].push(...errores);

                console.log("ERRORES", errores);

                await page.reload();
                continue;
              }
              if (
                hoy.getFullYear() -
                  Number(clientes[i].anio_nacimiento_hijo_01) >=
                25
              ) {
                clientes[i]["errores"].push("WARNING: Hijo 1 mayor de 25 años");
              }
              if (
                hoy.getFullYear() -
                  Number(clientes[i].anio_nacimiento_hijo_02) >=
                25
              ) {
                clientes[i]["errores"].push("WARNING: Hijo 2 mayor de 25 años");
              }
              if (
                hoy.getFullYear() -
                  Number(clientes[i].anio_nacimiento_hijo_03) >=
                25
              ) {
                clientes[i]["errores"].push("WARNING: Hijo 3 mayor de 25 años");
              }

              if (clientes[i].num_hijos > 3) {
                clientes[i]["errores"].push(
                  "ERROR: Faltan datos de descendencia (más de 3 hijos)",
                );
              }

              //********************
              // DESCARGA:
              //********************
              await page.locator("button ::-p-text(' Generar PDF')").wait();
              await page.locator("button ::-p-text(' Generar PDF')").click();

              await page.waitForSelector(".resultado");
              var resultados = await page.$$eval(".resultado", (spans) =>
                spans.map((span) => span.textContent.trim()),
              );

              clientes[i]["retencion_aplicable"] = parseFloat(
                resultados[0].replace(/\./g, "").replace(",", "."),
              );
              clientes[i]["resultado"] = parseFloat(
                resultados[1].replace(/\./g, "").replace(",", "."),
              );

              console.log("RESULTADO IRPF", resultados, clientes[i]);

              await this.esperar(2000);
              //await page.reload();
            } // FIN FOR CLIENTES

            //Cerrar navedador
            await browser.close();

            //Procesado de los resultados en XLSX:
            archivoIRPF
              .sheet(0)
              .cell(2, columnas + 1)
              .value("Retención Aplicable");
            archivoIRPF
              .sheet(0)
              .cell(2, columnas + 2)
              .value("Resultado IRPF");
            archivoIRPF
              .sheet(0)
              .cell(2, columnas + 3)
              .value("DIFF");
            archivoIRPF
              .sheet(0)
              .cell(2, columnas + 4)
              .value("Errores");

            var diff = 0;
            for (var i = 0; i < clientes.length; i++) {
              diff =
                (clientes[i].resultado || 0) - (clientes[i].sumatorio_016 || 0);

              archivoIRPF
                .sheet(0)
                .cell(i + 3, columnas + 1)
                .value(clientes[i].retencion_aplicable || 0);
              archivoIRPF
                .sheet(0)
                .cell(i + 3, columnas + 2)
                .value(clientes[i].resultado || 0);
              archivoIRPF
                .sheet(0)
                .cell(i + 3, columnas + 3)
                .value(diff);
              if (
                clientes[i].errores !== undefined &&
                clientes[i].errores !== null &&
                Array.isArray(clientes[i].errores) &&
                clientes[i].errores.length > 0
              ) {
                archivoIRPF
                  .sheet(0)
                  .cell(i + 3, columnas + 4)
                  .value(clientes[i].errores.join(" // "));
              } else {
                if (diff == 0) {
                  archivoIRPF
                    .sheet(0)
                    .cell(i + 3, columnas + 4)
                    .value("OK");
                }
              }
            }

            //ESCRITURA XLSX:
            console.log("Escribiendo archivo...");
            console.log("Path: " + path.normalize(pathSalidaExcel));

            archivoIRPF
              .toFileAsync(
                path.normalize(
                  path.join(pathSalidaExcel, "IRPF-Procesado.xlsx"),
                ),
              )
              .then(() => {
                console.log("Fin del procesamiento");
                //console.log(archivoIRPF)

                resolve(true);
              })
              .catch((err) => {
                console.log("Se ha producido un error interno: ");
                console.log(err);
                var tituloError =
                  "Se ha producido un error escribiendo el archivo: " +
                  path.normalize(pathSalidaExcel);
                resolve(false);
              });

            resolve(true);
          })
          .then(() => {})
          .catch((err) => {
            console.log("ERROR");

            throw err;
          });
      } catch (err) {
        var tituloError = "No se ha podido cargar el archivo";
        var mensajeError =
          "Se ha producido un error interno cargando los archivos.";
        mainProcess.mostrarError(tituloError, mensajeError).then((result) => {
          resolve(false);
        });
      }
    }).catch((err) => {
      console.log("Se ha producido un error interno: ");
      console.log(err);
      var tituloError = "No se ha podido cargar el archivo";
      var mensajeError =
        "Se ha producido un error interno cargando los archivos.";
      mainProcess.mostrarError(tituloError, mensajeError).then((result) => {
        resolve(false);
      });
    });
  }

  async iRPF2025(argumentos) {
    return new Promise((resolve) => {
      console.log("Calculo de IRPF...");
      console.log(argumentos.formularioControl[1]);
      console.log("Ruta Google...");
      console.log(argumentos.formularioControl[0]);

      var archivoIRPF = {};
      var clientes = [];
      var pathArchivoIRPF = argumentos.formularioControl[1];
      var pathSalidaExcel = path.join(
        path.normalize(argumentos.formularioControl[2]),
        "IRPF-Procesado",
      );
      var pathSalida = path.join(
        path.normalize(argumentos.formularioControl[2]),
        "IRPF-Procesado",
        "Resultados",
      );

      // Verificar si la carpeta "Resultados" existe y crearla si no
      if (!fs.existsSync(pathSalida)) {
        fs.mkdirSync(pathSalida, { recursive: true });
        console.log(`Carpeta creada: ${pathSalida}`);
      } else {
        console.log(`La carpeta ya existe: ${pathSalida}`);
      }

      try {
        XlsxPopulate.fromFileAsync(path.normalize(pathArchivoIRPF))
          .then(async (workbook) => {
            console.log("Archivo Cargado: IRPF");
            archivoIRPF = workbook;
            var columnas = archivoIRPF.sheet(0).usedRange()._numColumns;

            var filas = archivoIRPF.sheet(0).usedRange()._numRows;

            var objetoCliente = {};

            var cabeceras = [];
            for (var i = 1; i <= columnas; i++) {
              cabeceras.push(archivoIRPF.sheet(0).cell(2, i).value());
            }

            console.log("Cabeceras: " + cabeceras);

            for (var i = 3; i <= filas; i++) {
              objetoCliente = {};
              for (var j = 1; j <= columnas; j++) {
                if (archivoIRPF.sheet(0).cell(i, j).value() !== undefined) {
                  switch (cabeceras[j - 1]) {
                    case "Emp->Código_de_la_Empresa":
                      objetoCliente["cod_empresa"] = archivoIRPF
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;
                    case "Emp->Nombre_de_la_Empresa":
                      objetoCliente["nombre_empresa"] = archivoIRPF
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;
                    case "Trab->Código_del_Trabajador":
                      objetoCliente["cod_trabajador"] = archivoIRPF
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;
                    case "Trab->DNI_del_Trabajador":
                      objetoCliente["dni_trabajador"] = archivoIRPF
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;
                    case "Trab->Apellidos_y_Nombre_del_Trabajador":
                      objetoCliente["nombre_trabajador"] = archivoIRPF
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;
                    case "Trab->Número_de_hijos":
                      objetoCliente["num_hijos"] = archivoIRPF
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;
                    case "Trab->Porcentaje_retención":
                      objetoCliente["porcentaje_retencion"] = archivoIRPF
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;
                    case "Trab->Tipo_de_retención":
                      objetoCliente["tipo_retencion"] = archivoIRPF
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;
                    case "Trab->Ingresos_anuales":
                      objetoCliente["ingresos_anuales"] = archivoIRPF
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;
                    case "Trab->IRPF_Grado_Discapacidad":
                      objetoCliente["grado_discapacidad"] = archivoIRPF
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;
                    case "Trab->Tipo_Contrato_(3_posiciones)":
                      objetoCliente["tipo_contrato"] = archivoIRPF
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;
                    case "Trab->Edad_Trabajador":
                      objetoCliente["edad_trabajador"] = archivoIRPF
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;
                    case "Trab->Fecha_Nacimiento_(AAAA/MM/DD)":
                      objetoCliente["fecha_nacimiento"] = archivoIRPF
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;
                    case "Trab->Situación_Familiar":
                      objetoCliente["situacion_familiar"] = archivoIRPF
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;
                    case "Trab->DNI_Conyuge":
                      objetoCliente["dni_conyuge"] = archivoIRPF
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;
                    case "Trab->Año_Nacimiento_Hijo_01":
                      objetoCliente["anio_nacimiento_hijo_01"] = archivoIRPF
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;
                    case "Trab->Año_Nacimiento_Hijo_02":
                      objetoCliente["anio_nacimiento_hijo_02"] = archivoIRPF
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;
                    case "Trab->Año_Nacimiento_Hijo_03":
                      objetoCliente["anio_nacimiento_hijo_03"] = archivoIRPF
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;
                    case "Trab->Indicador_Adquisición_Vivienda":
                      objetoCliente["adquisicion_vivienda"] = archivoIRPF
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;
                    case "Trab->Pensiones_Compensatorias_Cónyuge":
                      objetoCliente["pension_conyuge"] = archivoIRPF
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;
                    case "Trab->Anualidades_en_Favor_de_los_Hijos":
                      objetoCliente["anualidades_hijos"] = archivoIRPF
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;
                    case "Trab->Sumatorio_015_de_conceptos_de_paga":
                      objetoCliente["sumatorio_015"] = archivoIRPF
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;
                    case "Trab->Sumatorio_016_de_conceptos_de_paga":
                      objetoCliente["sumatorio_016"] = archivoIRPF
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;
                    case "Trab->Sumatorio_017_de_conceptos_de_paga":
                      objetoCliente["sumatorio_017"] = archivoIRPF
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;
                  }
                }
              }

              objetoCliente["errores"] = [];

              if (
                objetoCliente.dni_trabajador !== "" &&
                objetoCliente.dni_trabajador !== null &&
                objetoCliente.dni_trabajador !== undefined
              ) {
                clientes.push(Object.assign({}, objetoCliente));
              }
            }

            var chromiumExecutablePath = path.normalize(
              argumentos.formularioControl[0],
            );

            //Inicio de procesamiento:
            const browser = await puppeteer.launch({
              executablePath: chromiumExecutablePath,
              headless: false,
            });

            var page = await browser.newPage();

            // Configurar el comportamiento de descarga
            await page._client().send("Page.setDownloadBehavior", {
              behavior: "allow",
              downloadPath: pathSalida,
            });

            await page.setViewport({ width: 1080, height: 1024 });

            var hoy = new Date();
            for (var i = 0; i < clientes.length; i++) {
              //Recargar cada 10 clientes:
              if (i % 10 == 0 && i > 0) {
                //await browser.close();
                await page.close();
                page = await browser.newPage();

                // Configurar el comportamiento de descarga
                await page._client().send("Page.setDownloadBehavior", {
                  behavior: "allow",
                  downloadPath: pathSalida,
                });
                await page.setViewport({ width: 1080, height: 1024 });
              }

              console.log("Procesando cliente: " + i);
              console.log(clientes[i]);

              if (
                clientes[i].dni_trabajador == "" ||
                clientes[i].dni_trabajador == null ||
                clientes[i].dni_trabajador == undefined
              ) {
                clientes[i]["errores"] = ["DNI del trabajador no definido."];
                continue;
              }

              await page.goto(
                "https://prewww2.aeat.es/wlpl/PRET-R200/R250/index.zul",
                { waitUntil: "networkidle0" },
              );

              //Procesado:

              //********
              // DNI
              //********
              await page.locator('input[title="NIF del perceptor"]').wait();
              await page.type(
                'input[title="NIF del perceptor"]',
                String(clientes[i].dni_trabajador),
              );

              //********
              // AÑO DE NACIMIENTO
              //********
              var anioNacimiento = clientes[i].fecha_nacimiento.slice(-4);
              await page.locator('input[title="Año de nacimiento"]').wait();
              await page.type(
                'input[title="Año de nacimiento"]',
                anioNacimiento,
              );

              //********
              //Seleccion de discapacidad:
              //********
              var spanSelector = 'span[title="Sin discapacidad"]';

              if (
                clientes[i].grado_discapacidad == "" ||
                clientes[i].grado_discapacidad == null ||
                clientes[i].grado_discapacidad == undefined
              ) {
                spanSelector = 'span[title="Sin discapacidad"]';
              } else if (clientes[i].grado_discapacidad >= 65) {
                spanSelector = 'span[title="Superior o igual al 65%"]';
              } else if (clientes[i].grado_discapacidad >= 33) {
                spanSelector =
                  'span[title="Superior o igual al 33% e inferior al 65%"]';
              }

              await page.locator(`${spanSelector} input[type="radio"]`).wait();
              var radioButton = await page.$(
                `${spanSelector} input[type="radio"]`,
              );

              if (radioButton) {
                await radioButton.click(); // Hacer clic en el radio button
                console.log("Radio button seleccionado.");
              } else {
                console.log("No se encontró el radio button.");
              }

              //********
              //Seleccion situacion familiar:
              //********
              var spanSelector = "";

              switch (clientes[i].situacion_familiar) {
                case "Soltero,divorciado,v":
                  spanSelector = `span[title='Situación 1: Soltero/a, viudo/a, divorciado/a o separado/a legalmente, con hijos solteros menores de 18 años o incapacitados judicialmente que convivan exclusivamente con el perceptor, sin convivir también con el otro progenitor, siempre que proceda consignar al menos un hijo o descendiente en el apartado "Ascendientes y  Descendientes"']`;
                  break;
                case "Conyuge a Cargo":
                  spanSelector =
                    'span[title="Situación 2: Perceptor casado y no separado legalmente cuyo cónyuge no obtenga rentas superiores a 1.500 euros anuales, excluidas las exentas."]';
                  break;
                case "Sin conyuge a Cargo":
                  spanSelector =
                    'span[title="Situación 3: Perceptor cuya situación familiar es distinta de las dos anteriores (v. gr.: solteros sin hijos; casados cuyo cónyuge obtiene rentas superiores a 1.500 euros anuales, excluidas las exentas, etc.).También se marcará esta casilla cuando el perceptor no desee manifestar su situación familiar"]';
                  break;
              }

              await page.locator(`${spanSelector} input[type="radio"]`).wait();
              await page.locator(`${spanSelector}`).click();

              //Si hay conyuge a cargo pone su DNI:
              if (clientes[i].situacion_familiar == "Conyuge a Cargo") {
                await page.locator('input[title="NIF del cónyuge"]').wait();
                await page.type(
                  'input[title="NIF del cónyuge"]',
                  clientes[i].dni_conyuge,
                );
              }

              //********************
              // TIPO CONTRATO:
              //********************
              spanSelector =
                'span[title="General o relaciones laborales especiales de las personas con discapacidad en centros especiales de empleo, y de los penados en instituciones penitenciarias"]';

              if (clientes[i].tipo_contrato >= 300) {
                spanSelector =
                  'span[title="Duración inferior al año o relación laboral especial de las personas artistas que desarrollan actividades escénicas, audiovisuales y musicales, y de quienes realizan actividades técnicas o auxiliares necesarias para el desarrollo de dicha actividad (excepto relaciones esporádicas: peonadas y jornales diarios)."]';
              }

              await page.locator(`${spanSelector} input[type="radio"]`).wait();
              await page.locator(`${spanSelector}`).click();

              // ******************
              // DATOS ASCENDIENTES / DESCENDIENTES:
              // ******************

              if (
                clientes[i].anio_nacimiento_hijo_01 ||
                clientes[i].anio_nacimiento_hijo_02 ||
                clientes[i].anio_nacimiento_hijo_03
              ) {
                await page
                  .locator("span ::-p-text('Ascendientes y descendientes')")
                  .wait();
                await page
                  .locator("span ::-p-text('Ascendientes y descendientes')")
                  .click();

                //Hijo 01:
                if (
                  clientes[i].anio_nacimiento_hijo_01 &&
                  hoy.getFullYear() -
                    Number(clientes[i].anio_nacimiento_hijo_01) <
                    25
                ) {
                  await page.locator(".z-icon-user-plus").wait();
                  await page.locator(".z-icon-user-plus").click();

                  await page
                    .locator('[role="dialog"] input[title="Año de nacimiento"]')
                    .wait();
                  await page.type(
                    '[role="dialog"] input[title="Año de nacimiento"]',
                    String(clientes[i].anio_nacimiento_hijo_01),
                  );

                  await page.locator("button ::-p-text(' Aceptar')").wait();
                  await page.locator("button ::-p-text(' Aceptar')").click();
                  await page.waitForSelector('[role="dialog"]', {
                    hidden: true,
                  });
                }

                //Hijo 02:
                if (
                  clientes[i].anio_nacimiento_hijo_02 &&
                  hoy.getFullYear() -
                    Number(clientes[i].anio_nacimiento_hijo_02) <
                    25
                ) {
                  await page.locator(".z-icon-user-plus").wait();
                  await page.locator(".z-icon-user-plus").click();

                  await page
                    .locator('[role="dialog"] input[title="Año de nacimiento"]')
                    .wait();
                  await page.type(
                    '[role="dialog"] input[title="Año de nacimiento"]',
                    String(clientes[i].anio_nacimiento_hijo_02),
                  );

                  await page.locator("button ::-p-text(' Aceptar')").click();
                  await page.waitForSelector('[role="dialog"]', {
                    hidden: true,
                  });
                }

                //Hijo 03:
                if (
                  clientes[i].anio_nacimiento_hijo_03 &&
                  hoy.getFullYear() -
                    Number(clientes[i].anio_nacimiento_hijo_03) <
                    25
                ) {
                  await page.locator(".z-icon-user-plus").wait();
                  await page.locator(".z-icon-user-plus").click();

                  await page
                    .locator('[role="dialog"] input[title="Año de nacimiento"]')
                    .wait();
                  await page.type(
                    '[role="dialog"] input[title="Año de nacimiento"]',
                    String(clientes[i].anio_nacimiento_hijo_03),
                  );

                  await page.locator("button ::-p-text(' Aceptar')").click();
                  await page.waitForSelector('[role="dialog"]', {
                    hidden: true,
                  });
                }
              } //Fin ascentientes y descendientes.

              // ******************
              // DATOS ECONOMICOS:
              // ******************
              await page.locator("span ::-p-text('Datos económicos')").wait();
              await page.locator("span ::-p-text('Datos económicos')").click();
              await page
                .locator(
                  'input[title="Retribuciones totales (dinerarias y en especie)."]',
                )
                .wait();
              await page
                .locator(
                  'input[title="Gastos deducibles (Art. 19.2, letras a, b y c de la LIRPF: Seguridad Social, Mutualidades de funcionarios, derechos pasivos, colegios de huérfanos o instituciones similares)"]',
                )
                .wait();
              await page
                .locator(
                  'input[title="Pensión compensatoria a favor del cónyuge. Importe fijado judicialmente"]',
                )
                .wait();
              await page
                .locator(
                  'input[title="Anualidades por alimentos en favor de los hijos. Importe fijado judicialmente"]',
                )
                .wait();
              await page
                .locator(
                  'span[title="El perceptor ha comunicado en el modelo 145 que está efectuando pagos por préstamos destinados a la adquisición o rehabilitación de su vivienda habitual por los que va a tener derecho a deducción por inversión en vivienda habitual en el IRPF y que la suma de los rendimientos íntegros del trabajo procedentes de todos sus pagadores es inferior a 33.007,20 euros anuales."]',
                )
                .wait();

              if (clientes[i].sumatorio_015) {
                await page.type(
                  'input[title="Retribuciones totales (dinerarias y en especie)."]',
                  String(clientes[i].sumatorio_015),
                );
              }

              if (clientes[i].sumatorio_017) {
                await page.type(
                  'input[title="Gastos deducibles (Art. 19.2, letras a, b y c de la LIRPF: Seguridad Social, Mutualidades de funcionarios, derechos pasivos, colegios de huérfanos o instituciones similares)"]',
                  String(clientes[i].sumatorio_017),
                );
              }

              if (clientes[i].pension_conyuge) {
                await page.type(
                  'input[title="Pensión compensatoria a favor del cónyuge. Importe fijado judicialmente"]',
                  String(clientes[i].pension_conyuge),
                );
              }

              if (clientes[i].anualidades_hijos) {
                clientes[i].anualidades_hijos =
                  parseFloat(clientes[i].anualidades_hijos) / 12;

                await page.type(
                  'input[title="Anualidades por alimentos en favor de los hijos. Importe fijado judicialmente"]',
                  String(clientes[i].anualidades_hijos),
                );
              }

              if (clientes[i].adquisicion_vivienda == "Destina (ant.2010)") {
                if (clientes[i].sumatorio_015 < 33007.2) {
                  await page
                    .locator(
                      'span[title="El perceptor ha comunicado en el modelo 145 que está efectuando pagos por préstamos destinados a la adquisición o rehabilitación de su vivienda habitual por los que va a tener derecho a deducción por inversión en vivienda habitual en el IRPF y que la suma de los rendimientos íntegros del trabajo procedentes de todos sus pagadores es inferior a 33.007,20 euros anuales."]',
                    )
                    .click();
                } else {
                  clientes[i]["errores"].push(
                    "WARN: Ingresos superiores a 33.007,20 euros anuales. Omitiendo deducción por vivienda habitual.",
                  );
                }
              }

              // ******************
              // RESULTADOS:
              // ******************

              if (!clientes[i].sumatorio_017) {
                clientes[i]["errores"].push(
                  "ERROR: Faltan datos de sumatorio_017",
                );
                await page.reload();
                continue;
              }
              await page.locator("span ::-p-text('Resultados')").wait();
              await page.locator("span ::-p-text('Resultados')").click();

              await this.esperar(2000);

              const found = await page.evaluate(() => {
                const div = document.querySelector("div");
                return div && div.textContent.includes("Relación de errores");
              });

              if (found) {
                console.log("ERROR EN EL PROCESAMIENTO", i);

                await this.esperar(2000);

                var errores = await page.$$eval(".z-label", (spans) =>
                  spans.map((span) => span.textContent.trim()),
                );

                clientes[i]["errores"].push(...errores);

                console.log("ERRORES", errores);

                await page.reload();
                continue;
              }
              if (
                hoy.getFullYear() -
                  Number(clientes[i].anio_nacimiento_hijo_01) >=
                25
              ) {
                clientes[i]["errores"].push("WARNING: Hijo 1 mayor de 25 años");
              }
              if (
                hoy.getFullYear() -
                  Number(clientes[i].anio_nacimiento_hijo_02) >=
                25
              ) {
                clientes[i]["errores"].push("WARNING: Hijo 2 mayor de 25 años");
              }
              if (
                hoy.getFullYear() -
                  Number(clientes[i].anio_nacimiento_hijo_03) >=
                25
              ) {
                clientes[i]["errores"].push("WARNING: Hijo 3 mayor de 25 años");
              }

              if (clientes[i].num_hijos > 3) {
                clientes[i]["errores"].push(
                  "ERROR: Faltan datos de descendencia (más de 3 hijos)",
                );
              }

              //********************
              // DESCARGA:
              //********************
              await page.locator("button ::-p-text(' Generar PDF')").wait();
              await page.locator("button ::-p-text(' Generar PDF')").click();

              await page.waitForSelector(".resultado");
              var resultados = await page.$$eval(".resultado", (spans) =>
                spans.map((span) => span.textContent.trim()),
              );

              clientes[i]["retencion_aplicable"] = parseFloat(
                resultados[0].replace(/\./g, "").replace(",", "."),
              );
              clientes[i]["resultado"] = parseFloat(
                resultados[1].replace(/\./g, "").replace(",", "."),
              );

              console.log("RESULTADO IRPF", resultados, clientes[i]);

              await this.esperar(2000);
              //await page.reload();
            } // FIN FOR CLIENTES

            //Cerrar navedador
            await browser.close();

            //Procesado de los resultados en XLSX:
            archivoIRPF
              .sheet(0)
              .cell(2, columnas + 1)
              .value("Retención Aplicable");
            archivoIRPF
              .sheet(0)
              .cell(2, columnas + 2)
              .value("Resultado IRPF");
            archivoIRPF
              .sheet(0)
              .cell(2, columnas + 3)
              .value("DIFF");
            archivoIRPF
              .sheet(0)
              .cell(2, columnas + 4)
              .value("Errores");

            var diff = 0;
            for (var i = 0; i < clientes.length; i++) {
              diff =
                (clientes[i].resultado || 0) - (clientes[i].sumatorio_016 || 0);

              archivoIRPF
                .sheet(0)
                .cell(i + 3, columnas + 1)
                .value(clientes[i].retencion_aplicable || 0);
              archivoIRPF
                .sheet(0)
                .cell(i + 3, columnas + 2)
                .value(clientes[i].resultado || 0);
              archivoIRPF
                .sheet(0)
                .cell(i + 3, columnas + 3)
                .value(diff);
              if (
                clientes[i].errores !== undefined &&
                clientes[i].errores !== null &&
                Array.isArray(clientes[i].errores) &&
                clientes[i].errores.length > 0
              ) {
                archivoIRPF
                  .sheet(0)
                  .cell(i + 3, columnas + 4)
                  .value(clientes[i].errores.join(" // "));
              } else {
                if (diff == 0) {
                  archivoIRPF
                    .sheet(0)
                    .cell(i + 3, columnas + 4)
                    .value("OK");
                }
              }
            }

            //ESCRITURA XLSX:
            console.log("Escribiendo archivo...");
            console.log("Path: " + path.normalize(pathSalidaExcel));

            archivoIRPF
              .toFileAsync(
                path.normalize(
                  path.join(pathSalidaExcel, "IRPF-Procesado.xlsx"),
                ),
              )
              .then(() => {
                console.log("Fin del procesamiento");
                //console.log(archivoIRPF)

                resolve(true);
              })
              .catch((err) => {
                console.log("Se ha producido un error interno: ");
                console.log(err);
                var tituloError =
                  "Se ha producido un error escribiendo el archivo: " +
                  path.normalize(pathSalidaExcel);
                resolve(false);
              });

            resolve(true);
          })
          .then(() => {})
          .catch((err) => {
            console.log("ERROR");

            throw err;
          });
      } catch (err) {
        var tituloError = "No se ha podido cargar el archivo";
        var mensajeError =
          "Se ha producido un error interno cargando los archivos.";
        mainProcess.mostrarError(tituloError, mensajeError).then((result) => {
          resolve(false);
        });
      }
    }).catch((err) => {
      console.log("Se ha producido un error interno: ");
      console.log(err);
      var tituloError = "No se ha podido cargar el archivo";
      var mensajeError =
        "Se ha producido un error interno cargando los archivos.";
      mainProcess.mostrarError(tituloError, mensajeError).then((result) => {
        resolve(false);
      });
    });
  }

  async cambioBaseDeCotizacion(argumentos) {
    return new Promise((resolve) => {
      console.log("Cambio de base de cotización...");
      console.log(argumentos.formularioControl[1]);
      console.log("Ruta Google...");
      console.log(argumentos.formularioControl[0]);

      var archivoCambioBase = {};
      var clientes = [];
      var pathArchivoCambioBase = argumentos.formularioControl[1];
      var pathSalidaExcel = path.join(
        path.normalize(argumentos.formularioControl[2]),
        "Cambio de Base - Procesado",
      );
      var pathSalida = path.join(
        path.normalize(argumentos.formularioControl[2]),
        "Cambio de Base - Procesado",
        "Resultados",
      );

      // Verificar si la carpeta "Resultados" existe y crearla si no
      if (!fs.existsSync(pathSalida)) {
        fs.mkdirSync(pathSalida, { recursive: true });
        console.log(`Carpeta creada: ${pathSalida}`);
      } else {
        console.log(`La carpeta ya existe: ${pathSalida}`);
      }

      try {
        XlsxPopulate.fromFileAsync(path.normalize(pathArchivoCambioBase))
          .then(async (workbook) => {
            console.log("Archivo Cargado: Cambio de Base");
            archivoCambioBase = workbook;
            var columnas = archivoCambioBase.sheet(0).usedRange()._numColumns;

            var filas = archivoCambioBase.sheet(0).usedRange()._numRows;

            var objetoCliente = {};

            var cabeceras = [];
            for (var i = 1; i <= columnas; i++) {
              cabeceras.push(archivoCambioBase.sheet(0).cell(1, i).value());
            }

            console.log("Cabeceras: " + cabeceras);

            for (var i = 2; i <= filas; i++) {
              objetoCliente = {};
              for (var j = 1; j <= columnas; j++) {
                if (
                  archivoCambioBase.sheet(0).cell(i, j).value() !== undefined
                ) {
                  switch (cabeceras[j - 1]) {
                    case "EXPT":
                      objetoCliente["expediente"] = archivoCambioBase
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;
                    case "NOMBRE Y APELLIDOS":
                      objetoCliente["nombre"] = archivoCambioBase
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;
                    case "DNI":
                      objetoCliente["dni"] = archivoCambioBase
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;
                    case "NAF":
                      objetoCliente["seguridad_social"] = archivoCambioBase
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;
                    case "BASE MINIMA S/TRAMO":
                      objetoCliente["base_minima"] = archivoCambioBase
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;
                  }
                }
              }

              objetoCliente["errores"] = [];

              if (
                objetoCliente.dni !== "" &&
                objetoCliente.dni !== null &&
                objetoCliente.dni !== undefined
              ) {
                clientes.push(Object.assign({}, objetoCliente));
              }
            }

            console.log("Clientes: ", clientes);
            resolve(true);

            var chromiumExecutablePath = path.normalize(
              argumentos.formularioControl[0],
            );

            //Inicio de procesamiento:
            const browser = await puppeteer.launch({
              executablePath: chromiumExecutablePath,
              headless: false,
              args: [
                `--disable-extensions`,
                `--no-sandbox`,
                `--disable-setuid-sandbox`,
              ],
            });

            var page = await browser.newPage();

            page.on("dialog", async (dialog) => {
              console.log(
                `Se mostró un cuadro de diálogo: ${dialog.message()}`,
              );
              await dialog.accept(); // Acepta el cuadro de diálogo
            });

            // Configurar el comportamiento de descarga
            const client = await page.target().createCDPSession();
            await client.send("Page.setDownloadBehavior", {
              behavior: "allow",
              downloadPath: pathSalida,
            });

            await page.setViewport({ width: 1080, height: 1024 });

            var hoy = new Date();
            for (var i = 0; i < clientes.length; i++) {
              console.log("Procesando cliente: " + i);
              console.log(clientes[i]);

              if (
                clientes[i].dni == "" ||
                clientes[i].dni == null ||
                clientes[i].dni == undefined
              ) {
                clientes[i]["errores"] = ["DNI del trabajador no definido."];
                continue;
              }

              await page.goto(
                "https://w2.seg-social.es/ProsaInternet/OnlineAccess?ARQ.SPM.ACTION=LOGIN&ARQ.SPM.APPTYPE=SERVICE&ARQ.IDAPP=XV26C007",
                {
                  waitUntil: "networkidle0",
                },
              );

              // ******************
              // RESULTADOS:
              // ******************
              console.log("Esperando a que cargue el contenido...");

              //Aceptar terminos y condiciones:
              await page.locator("#CHK_LEIDO").wait();
              await page.locator("#CHK_LEIDO").click();
              await page
                .locator(
                  'button[title="Ejecuta la acción y continúa a la siguiente pantalla."]',
                )
                .wait();
              await page
                .locator(
                  'button[title="Ejecuta la acción y continúa a la siguiente pantalla."]',
                )
                .click();

              //Rellenar Número de segurida social:
              await page
                .locator(
                  'input[title="Número de la Seguridad Social (Númerico 12)"]',
                )
                .wait();
              await page.type(
                'input[title="Número de la Seguridad Social (Númerico 12)"]',
                String(clientes[i].seguridad_social),
              );

              //Selecciona el tipo de documentos:
              await page.locator("#IPF_TIPO").wait();

              const startsWithNumber = (str) => {
                if (!str) return false; // Manejo de cadena vacía
                const firstChar = str[0]; // Primer carácter
                return !isNaN(firstChar); // isNaN -> false si es número
              };

              if (startsWithNumber(clientes[i].dni)) {
                await page.select("#IPF_TIPO", "1");
              } else {
                await page.select("#IPF_TIPO", "6");
              }

              //Rellena el dni
              await page.locator("#IPF_NUMERO").wait();
              await page.type("#IPF_NUMERO", String(clientes[i].dni));

              //Click Continuar
              await page.locator("#ENVIO_3").wait();
              await page.locator("#ENVIO_3").click();

              //Selecciona el tipo de documentos:
              await page.locator("#OPCION_BASE").wait();
              await page.select("#OPCION_BASE", "5");

              //Rellena base de cotización:
              var baseMinima = String(clientes[i].base_minima);
              baseMinima = baseMinima.replace(".", ",");
              await page.locator("#OTRA_BASE").wait();
              await page.type("#OTRA_BASE", baseMinima);

              await this.esperar(1000);

              //Click Continuar
              await page.locator("#ENVIO_3").wait();
              await page.locator("#ENVIO_3").click();

              console.log("CLICK");
              await this.esperar(1000);
              console.log("CLICK");

              const errorYaSolicitada = await page.evaluate(() => {
                console.log("Iniciando");
                const elementos = document.querySelectorAll(".pr_pMensaje");
                console.log("Elementos", elementos);
                return Array.from(elementos).some((el) =>
                  el.textContent.includes(
                    "4913* BASE IGUAL A LA SOLICITADA CON ANTERIORIDAD.",
                  ),
                );
              });

              if (errorYaSolicitada) {
                console.log("Base de cotización igual a la solicitada.", i);
                await this.esperar(1000);
                var errores =
                  "Cotización ya solicitada con anterioridad. No se puede volver a solicitar.";

                clientes[i]["errores"].push(errores);

                console.log("ERRORES", errores);

                await page.reload();
                continue;
              }

              const exito = await page.evaluate(() => {
                console.log("Iniciando");
                const elementos = document.querySelectorAll(".pr_pMensaje");
                console.log("Elementos", elementos);
                return Array.from(elementos).some((el) =>
                  el.textContent.includes("Operación realizada correctamente."),
                );
              });

              if (exito) {
                console.log("Operación Exitosa");
                await page.locator('button[title="Cerrar"]').wait();
                await page.locator('button[title="Cerrar"]').click();

                // Selector del enlace que apunta al archivo PDF
                const selectorEnlace = "a.pr_enlaceDocInforme";

                await this.esperar(2000); // Ajusta según el tamaño del archivo
                // Haz clic en el enlace para iniciar la descarga
                await page.locator("a.pr_enlaceDocInforme").wait();
                await page.locator("a.pr_enlaceDocInforme").click();
                await this.esperar(3000); // Ajusta según el tamaño del archivo
                //await page.locator("a.pr_enlaceDocInforme").click();

                // Espera un tiempo para asegurarte de que la descarga se complete
                console.log("Descargando archivo...");

                var errores = "Realizado con exito.";

                clientes[i]["errores"].push(errores);

                console.log("ERRORES", errores);
                await page.reload();
                continue;
              }

              await this.esperar(2000);
              await page.reload();
            } // FIN FOR CLIENTES

            //Cerrar navedador
            //await browser.close();

            console.log("Clientes: ", clientes);
            console.log("Columnas: ", columnas);

            //Procesado de los resultados en XLSX:
            archivoCambioBase
              .sheet(0)
              .cell(1, 27 + 1)
              .value("Comentarios");

            for (var i = 0; i < clientes.length; i++) {
              if (
                clientes[i].errores !== undefined &&
                clientes[i].errores !== null &&
                Array.isArray(clientes[i].errores) &&
                clientes[i].errores.length > 0
              ) {
                archivoCambioBase
                  .sheet(0)
                  .cell(i + 2, 27 + 1)
                  .value(clientes[i].errores.join(" // "));
              } else {
                if (diff == 0) {
                  archivoCambioBase
                    .sheet(0)
                    .cell(i + 2, 27 + 1)
                    .value("Error");
                }
              }
            }

            //ESCRITURA XLSX:
            console.log("Escribiendo archivo...");
            console.log("Path: " + path.normalize(pathSalidaExcel));

            archivoCambioBase
              .toFileAsync(
                path.normalize(
                  path.join(pathSalidaExcel, "Cambio-Base-Cotizacion.xlsx"),
                ),
              )
              .then(() => {
                console.log("Fin del procesamiento");
                //console.log(archivoIRPF)

                resolve(true);
              })
              .catch((err) => {
                console.log("Se ha producido un error interno: ");
                console.log(err);
                var tituloError =
                  "Se ha producido un error escribiendo el archivo: " +
                  path.normalize(pathSalidaExcel);
                resolve(false);
              });

            resolve(true);
          })
          .then(() => {})
          .catch((err) => {
            console.log("ERROR");

            throw err;
          });
      } catch (err) {
        var tituloError = "No se ha podido cargar el archivo";
        var mensajeError =
          "Se ha producido un error interno cargando los archivos.";
        mainProcess.mostrarError(tituloError, mensajeError).then((result) => {
          resolve(false);
        });
      }
    }).catch((err) => {
      console.log("Se ha producido un error interno: ");
      console.log(err);
      var tituloError = "No se ha podido cargar el archivo";
      var mensajeError =
        "Se ha producido un error interno cargando los archivos.";
      mainProcess.mostrarError(tituloError, mensajeError).then((result) => {
        resolve(false);
      });
    });
  }

  async cartasDePagoEnHacienda(argumentos) {
    return new Promise((resolve) => {
      console.log("Cartas de pago en hacienda");
      console.log(argumentos.formularioControl[1]);
      console.log("Ruta Google...");
      console.log(argumentos.formularioControl[0]);

      var archivoCartas = {};
      var clientes = [];
      var pathArchivoCartas = argumentos.formularioControl[1];
      var pathSalidaExcel = path.join(
        path.normalize(argumentos.formularioControl[2]),
        "Cartas_de_pago-Procesado",
      );
      var pathSalida = path.join(
        path.normalize(argumentos.formularioControl[2]),
        "Cartas_de_pago-Procesado",
        "Resultados",
      );

      // Verificar si la carpeta "Resultados" existe y crearla si no
      if (!fs.existsSync(pathSalida)) {
        fs.mkdirSync(pathSalida, { recursive: true });
        console.log(`Carpeta creada: ${pathSalida}`);
      } else {
        console.log(`La carpeta ya existe: ${pathSalida}`);
      }

      try {
        XlsxPopulate.fromFileAsync(path.normalize(pathArchivoCartas))
          .then(async (workbook) => {
            console.log("Archivo Cargado: Cartas");
            archivoCartas = workbook;
            var columnas = archivoCartas.sheet(0).usedRange()._numColumns;

            var filas = archivoCartas.sheet(0).usedRange()._numRows;

            var objetoCliente = {};

            var cabeceras = [];
            for (var i = 1; i <= columnas; i++) {
              cabeceras.push(archivoCartas.sheet(0).cell(3, i).value());
            }

            console.log("Cabeceras: " + cabeceras);

            for (var i = 4; i <= filas; i++) {
              objetoCliente = {};
              for (var j = 1; j <= columnas; j++) {
                if (archivoCartas.sheet(0).cell(i, j).value() !== undefined) {
                  switch (cabeceras[j - 1]) {
                    case "D.N.I. .":
                      objetoCliente["dni"] = archivoCartas
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;

                    case "D.N.I.":
                      objetoCliente["dni"] = archivoCartas
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;

                    case "Exp":
                      objetoCliente["expediente"] = archivoCartas
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;

                    case "EXP":
                      objetoCliente["expediente"] = archivoCartas
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;

                    case "NIF PAGADOR":
                      objetoCliente["nif_pagador"] = archivoCartas
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;

                    case "APELL.Y NOMBRE":
                      objetoCliente["nombre"] = archivoCartas
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;

                    case "DILIGENCIA":
                      objetoCliente["diligencia"] = archivoCartas
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;
                  }
                }
              }

              objetoCliente["errores"] = [];

              if (
                objetoCliente.dni &&
                objetoCliente.expediente &&
                objetoCliente.nif_pagador &&
                objetoCliente.nombre &&
                objetoCliente.diligencia
              ) {
                objetoCliente["nombreArchivo"] =
                  objetoCliente["expediente"] +
                  objetoCliente["dni"] +
                  "_CARTA_DE_PAGO_" +
                  objetoCliente["nombre"]
                    .replaceAll(" ", "_")
                    .replaceAll(".", "")
                    .replaceAll(",", "") +
                  ".pdf";
                clientes.push(Object.assign({}, objetoCliente));
              }
            }

            console.log("Clientes: ");
            console.log(clientes);

            var chromiumExecutablePath = path.normalize(
              argumentos.formularioControl[0],
            );

            //Inicio de procesamiento:
            const browser = await puppeteer.launch({
              executablePath: chromiumExecutablePath,
              headless: false,
            });

            var page = await browser.newPage();

            // Configurar el comportamiento de descarga
            await page._client().send("Page.setDownloadBehavior", {
              behavior: "allow",
              downloadPath: pathSalida,
            });

            await page.setViewport({ width: 1080, height: 1024 });

            for (var i = 0; i < clientes.length; i++) {
              //Recargar cada 10 clientes:
              if (i % 10 == 0 && i > 0) {
                //await browser.close();
                await page.close();
                page = await browser.newPage();

                // Configurar el comportamiento de descarga
                await page._client().send("Page.setDownloadBehavior", {
                  behavior: "allow",
                  downloadPath: pathSalida,
                });
                await page.setViewport({ width: 1080, height: 1024 });
              }

              console.log("Procesando cliente: " + i);
              console.log(clientes[i]);

              await page.goto(
                "https://www2.agenciatributaria.gob.es/wlpl/inwinvoc/es.aeat.dit.adu.srem.sueldos.SdoQuery?FModo=CP",
                { waitUntil: "networkidle0" },
              );
              //Si es la primera iteracion refresca la pagina para evitar mensaje de alertas pendientes:
              if (i == 1) {
                await page.goto(
                  "https://www2.agenciatributaria.gob.es/wlpl/inwinvoc/es.aeat.dit.adu.srem.sueldos.SdoQuery?FModo=CP",
                  { waitUntil: "networkidle0" },
                );
              }

              //********
              // NIF PAGADOR
              //********
              await page.locator('input[id="FNifPagador"]').wait();
              await page.type(
                'input[id="FNifPagador"]',
                String(clientes[i].nif_pagador),
              );

              //********
              // NIF OBLIGADO
              //********
              await page.locator('input[id="FNifDdr"]').wait();
              await page.type('input[id="FNifDdr"]', String(clientes[i].dni));

              //********
              // DILIGENCIA
              //********
              await page.locator('input[id="FNumDil"]').wait();
              await page.type(
                'input[id="FNumDil"]',
                String(clientes[i].diligencia),
              );

              //*************
              // Buscar
              //*************
              await page.locator('input[name="Buscar"]').wait();
              await page.locator('input[name="Buscar"]').click();

              await this.esperar(2000);

              // Buscar el enlace con JavaScript en el navegador y hacer clic si lo encuentra
              const enlaceEncontrado = await page.evaluate((texto) => {
                const enlaces = Array.from(document.querySelectorAll("a"));
                const enlace = enlaces.find((a) => a.innerText.includes(texto));

                if (enlace) {
                  enlace.click();
                  return enlace; // Retorna true si encontró y clicó el enlace
                }
                return false; // Retorna false si no encontró el enlace
              }, clientes[i]["diligencia"]);

              // Obtener lista de archivos antes de la descarga
              const archivosAntes = new Set(fs.readdirSync(pathSalida));
              await this.esperar(1000);

              if (enlaceEncontrado) {
                console.log(
                  `Enlace encontrado con texto: "${clientes[i]["diligencia"]}"`,
                );

                //*************
                // Generar Documento de ingreso
                //*************
                await page.locator('input[name="Aceptar"]').wait();
                await page.locator('input[name="Aceptar"]').click();

                //*************
                // Generar Documento de ingreso
                //*************

                const [nuevaPagina] = await Promise.all([
                  new Promise((resolve) =>
                    browser.once("targetcreated", (target) =>
                      resolve(target.page()),
                    ),
                  ),
                  await page.locator('input[name="cartapago_pdf"]').wait(),
                  await page.locator('input[name="cartapago_pdf"]').click(),
                ]);

                if (nuevaPagina) {
                  await this.esperar(1000);
                  const pdfUrl = nuevaPagina.url();
                  console.log(`📄 URL del PDF detectada: ${pdfUrl}`);

                  // Descargar el PDF manualmente con Axios
                  const pdfResponse = await axios.get(pdfUrl, {
                    responseType: "arraybuffer",
                  });
                  const filePath = path.join(
                    pathSalida,
                    clientes[i]["nombreArchivo"],
                  );
                  fs.writeFileSync(filePath, pdfResponse.data);

                  console.log(`✅ PDF descargado en: ${filePath}`);

                  await nuevaPagina.close();
                }
              } else {
                console.log(
                  `No se encontró el enlace con texto: "${clientes[i]["diligencia"]}"`,
                );
                continue;
              }

              await this.esperar(1000);

              // Buscar el nuevo archivo descargado
              const archivosDespues = new Set(fs.readdirSync(pathSalida));
              const archivoNuevo = [...archivosDespues].find(
                (file) => !archivosAntes.has(file),
              );

              if (archivoNuevo) {
                const oldPath = path.join(pathSalida, archivoNuevo);
                const newPath = path.join(
                  pathSalida,
                  clientes[i]["nombreArchivo"],
                );

                fs.renameSync(oldPath, newPath);
                console.log(
                  `Archivo renombrado: ${clientes[i]["nombreArchivo"]}`,
                );
              } else {
                console.log(`No se encontró archivo para el cliente: ${i + 1}`);
              }

              await this.esperar(1000);
            } //Fin iteracion de clientes
            //Cerrar navedador
            await browser.close();

            resolve(true);
          })
          .then(() => {})
          .catch((err) => {
            console.log("ERROR");

            throw err;
          });
      } catch (err) {
        var tituloError = "No se ha podido cargar el archivo";
        var mensajeError =
          "Se ha producido un error interno cargando los archivos.";
        mainProcess.mostrarError(tituloError, mensajeError).then((result) => {
          resolve(false);
        });
      }
    }).catch((err) => {
      console.log("Se ha producido un error interno: ");
      console.log(err);
      var tituloError = "No se ha podido cargar el archivo";
      var mensajeError =
        "Se ha producido un error interno cargando los archivos.";
      mainProcess.mostrarError(tituloError, mensajeError).then((result) => {
        resolve(false);
      });
    });
  }

  async etiquetasAEAT(argumentos) {
    return new Promise((resolve) => {
      console.log("Etiquetas AEAT...");
      console.log(argumentos.formularioControl[1]);
      console.log("Ruta Google...");
      console.log(argumentos.formularioControl[0]);

      var archivoEtiquetas = {};
      var clientes = [];
      var pathArchivoEtiquetas = argumentos.formularioControl[1];
      var pathSalidaExcel = path.join(
        path.normalize(argumentos.formularioControl[2]),
        "Etiquetas-Procesado",
      );
      var pathSalida = path.join(
        path.normalize(argumentos.formularioControl[2]),
        "Etiquetas-Procesado",
        "Resultados",
      );

      // Verificar si la carpeta "Resultados" existe y crearla si no
      if (!fs.existsSync(pathSalida)) {
        fs.mkdirSync(pathSalida, { recursive: true });
        console.log(`Carpeta creada: ${pathSalida}`);
      } else {
        console.log(`La carpeta ya existe: ${pathSalida}`);
      }

      try {
        XlsxPopulate.fromFileAsync(path.normalize(pathArchivoEtiquetas))
          .then(async (workbook) => {
            console.log("Archivo Cargado: Etiquetas");
            archivoEtiquetas = workbook;
            var columnas = archivoEtiquetas.sheet(0).usedRange()._numColumns;

            var filas = archivoEtiquetas.sheet(0).usedRange()._numRows;

            var objetoCliente = {};

            var cabeceras = [];
            for (var i = 1; i <= columnas; i++) {
              cabeceras.push(archivoEtiquetas.sheet(0).cell(2, i).value());
            }

            console.log("Cabeceras: " + cabeceras);

            for (var i = 3; i <= filas; i++) {
              objetoCliente = {};
              for (var j = 1; j <= columnas; j++) {
                if (
                  archivoEtiquetas.sheet(0).cell(i, j).value() !== undefined
                ) {
                  switch (cabeceras[j - 1]) {
                    case "'NIE":
                      objetoCliente["dni"] = archivoEtiquetas
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;

                    case "NIE":
                      objetoCliente["dni"] = archivoEtiquetas
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;

                    case "'EXP":
                      objetoCliente["expediente"] = archivoEtiquetas
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;
                    case "EXP":
                      objetoCliente["expediente"] = archivoEtiquetas
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;

                    case "'TRABAJADOR":
                      objetoCliente["nombre"] = archivoEtiquetas
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;

                    case "TRABAJADOR":
                      objetoCliente["nombre"] = archivoEtiquetas
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;
                  }
                }
              }

              objetoCliente["errores"] = [];

              if (
                objetoCliente.dni !== "" &&
                objetoCliente.dni !== null &&
                objetoCliente.dni !== undefined
              ) {
                objetoCliente["nombreArchivo"] =
                  objetoCliente["expediente"] +
                  "_ETIQUETA_" +
                  objetoCliente["nombre"]
                    .replaceAll(" ", "_")
                    .replaceAll(".", "")
                    .replaceAll(",", "") +
                  "_" +
                  objetoCliente["dni"] +
                  ".pdf";
                clientes.push(Object.assign({}, objetoCliente));
              }
            }

            console.log("Clientes: ");
            console.log(clientes);

            var chromiumExecutablePath = path.normalize(
              argumentos.formularioControl[0],
            );

            //Inicio de procesamiento:
            const browser = await puppeteer.launch({
              executablePath: chromiumExecutablePath,
              headless: false,
            });

            var page = await browser.newPage();

            // Configurar el comportamiento de descarga
            await page._client().send("Page.setDownloadBehavior", {
              behavior: "allow",
              downloadPath: pathSalida,
            });

            await page.setViewport({ width: 1080, height: 1024 });

            for (var i = 0; i < clientes.length; i++) {
              //Recargar cada 10 clientes:
              if (i % 10 == 0 && i > 0) {
                //await browser.close();
                await page.close();
                page = await browser.newPage();

                // Configurar el comportamiento de descarga
                await page._client().send("Page.setDownloadBehavior", {
                  behavior: "allow",
                  downloadPath: pathSalida,
                });
                await page.setViewport({ width: 1080, height: 1024 });
              }

              console.log("Procesando cliente: " + i);
              console.log(clientes[i]);

              if (
                clientes[i].dni == "" ||
                clientes[i].dni == null ||
                clientes[i].dni == undefined
              ) {
                clientes[i]["errores"] = ["DNI del trabajador no definido."];
                continue;
              }

              await page.goto(
                "https://www1.agenciatributaria.gob.es/static_files/common/internet/dep/aplicaciones/ov/eticerti.html",
                { waitUntil: "networkidle0" },
              );
              //Si es la primera iteracion refresca la pagina para evitar mensaje de alertas pendientes:
              if (i == 1) {
                await page.goto(
                  "https://www1.agenciatributaria.gob.es/static_files/common/internet/dep/aplicaciones/ov/eticerti.html",
                  { waitUntil: "networkidle0" },
                );
              }

              //********
              // DNI
              //********
              await page.locator('input[id="nif"]').wait();
              await page.type('input[id="nif"]', String(clientes[i].dni));

              //*************
              // APELLIDOS
              //*************
              await page.locator('input[id="ape"]').wait();
              await page.type('input[id="ape"]', "A");

              // Obtener lista de archivos antes de la descarga
              const archivosAntes = new Set(fs.readdirSync(pathSalida));

              //*************
              // Descargar
              //*************
              await page.locator('input[name="ENV"]').wait();
              await page.locator('input[name="ENV"]').click();

              await this.esperar(2000);

              // Buscar el nuevo archivo descargado
              const archivosDespues = new Set(fs.readdirSync(pathSalida));
              const archivoNuevo = [...archivosDespues].find(
                (file) => !archivosAntes.has(file),
              );

              if (archivoNuevo) {
                const oldPath = path.join(pathSalida, archivoNuevo);
                const newPath = path.join(
                  pathSalida,
                  clientes[i]["nombreArchivo"],
                );

                fs.renameSync(oldPath, newPath);
                console.log(
                  `Archivo renombrado: ${clientes[i]["nombreArchivo"]}`,
                );
              } else {
                console.log(`No se encontró archivo para el cliente: ${i + 1}`);
              }

              await this.esperar(2000);
            } //Fin iteracion de clientes
            //Cerrar navedador
            await browser.close();

            resolve(true);
          })
          .then(() => {})
          .catch((err) => {
            console.log("ERROR");

            throw err;
          });
      } catch (err) {
        var tituloError = "No se ha podido cargar el archivo";
        var mensajeError =
          "Se ha producido un error interno cargando los archivos.";
        mainProcess.mostrarError(tituloError, mensajeError).then((result) => {
          resolve(false);
        });
      }
    }).catch((err) => {
      console.log("Se ha producido un error interno: ");
      console.log(err);
      var tituloError = "No se ha podido cargar el archivo";
      var mensajeError =
        "Se ha producido un error interno cargando los archivos.";
      mainProcess.mostrarError(tituloError, mensajeError).then((result) => {
        resolve(false);
      });
    });
  }

  async actualizacionCNAE25(argumentos) {
    return new Promise((resolve) => {
      console.log("Archivo CNAE...");
      console.log(argumentos.formularioControl[1]);
      console.log("Ruta Google...");
      console.log(argumentos.formularioControl[0]);

      var archivoCNAE = {};
      var clientes = [];
      var pathArchivoEtiquetas = argumentos.formularioControl[1];
      var pathSalidaExcel = path.join(
        path.normalize(argumentos.formularioControl[2]),
        "CNAE-Informes-Procesados",
      );
      var pathSalida = path.join(
        path.normalize(argumentos.formularioControl[2]),
        "CNAE-Informes-Procesados",
        "Resultados",
      );

      // Verificar si la carpeta "Resultados" existe y crearla si no
      if (!fs.existsSync(pathSalida)) {
        fs.mkdirSync(pathSalida, { recursive: true });
        console.log(`Carpeta creada: ${pathSalida}`);
      } else {
        console.log(`La carpeta ya existe: ${pathSalida}`);
      }

      try {
        XlsxPopulate.fromFileAsync(path.normalize(pathArchivoEtiquetas))
          .then(async (workbook) => {
            console.log("Archivo Cargado: CNAE");
            archivoCNAE = workbook;
            var columnas = archivoCNAE.sheet(0).usedRange()._numColumns;
            var filas = archivoCNAE.sheet(0).usedRange()._numRows;
            var objetoCliente = {};

            var cabeceras = [];
            for (var i = 1; i <= columnas; i++) {
              cabeceras.push(archivoCNAE.sheet(0).cell(4, i).value());
            }

            console.log("Cabeceras: " + cabeceras);

            for (var i = 5; i <= filas; i++) {
              objetoCliente = {};
              for (var j = 1; j <= columnas; j++) {
                if (archivoCNAE.sheet(0).cell(i, j).value() !== undefined) {
                  switch (cabeceras[j - 1]) {
                    case "Código Cuenta Cotización (CCC)":
                      objetoCliente["ccc"] = archivoCNAE
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      objetoCliente["ccc1"] = objetoCliente["ccc"].substring(
                        0,
                        4,
                      );
                      objetoCliente["ccc2"] = objetoCliente["ccc"].substring(
                        4,
                        6,
                      );
                      objetoCliente["ccc3"] = objetoCliente["ccc"].substring(6);
                      break;

                    case "CNAE25":
                      objetoCliente["cnae25"] = archivoCNAE
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;

                    case "Expediente":
                      objetoCliente["expediente"] = archivoCNAE
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;
                  }
                }
              }

              objetoCliente["errores"] = [];

              if (
                objetoCliente.ccc !== "" &&
                objetoCliente.ccc !== null &&
                objetoCliente.ccc !== undefined &&
                objetoCliente.cnae25 !== "" &&
                objetoCliente.cnae25 !== null &&
                objetoCliente.cnae25 !== undefined
              ) {
                objetoCliente["nombreArchivo"] =
                  objetoCliente["expediente"] +
                  "-" +
                  objetoCliente["ccc"] +
                  ".pdf";
                clientes.push(Object.assign({}, objetoCliente));
              }
            }

            console.log("Clientes: ");
            console.log(clientes);

            var chromiumExecutablePath = path.normalize(
              argumentos.formularioControl[0],
            );

            //Inicio de procesamiento:
            const browser = await puppeteer.launch({
              executablePath: chromiumExecutablePath,
              headless: false,
            });
            console.log(browser.executablePath);

            var page = await browser.newPage();

            // Configurar el comportamiento de descarga
            await page._client().send("Page.setDownloadBehavior", {
              behavior: "allow",
              downloadPath: pathSalida,
            });

            await page.setViewport({ width: 1080, height: 1024 });

            for (var i = 0; i < clientes.length; i++) {
              //Recargar cada 10 clientes:
              if (i % 10 == 0 && i > 0) {
                //await browser.close();
                await page.close();
                page = await browser.newPage();

                // Configurar el comportamiento de descarga
                await page._client().send("Page.setDownloadBehavior", {
                  behavior: "allow",
                  downloadPath: pathSalida,
                });
                await page.setViewport({ width: 1080, height: 1024 });
              }

              console.log("Procesando cliente: " + i);
              console.log(clientes[i]);

              if (
                clientes[i].ccc == "" ||
                clientes[i].ccc == null ||
                clientes[i].ccc == undefined ||
                clientes[i].cnae25 == "" ||
                clientes[i].cnae25 == null ||
                clientes[i].cnae25 == undefined
              ) {
                clientes[i]["errores"] = ["Campos CCC o CNAE25 no definidos."];
                continue;
              }

              await page.goto(
                "https://w2.seg-social.es/Xhtml?JacadaApplicationName=SGIRED&TRANSACCION=ACR82&E=I&AP=AFIR",
                { waitUntil: "networkidle0" },
              );

              await this.esperar(1000);

              //********
              // CCC1
              //********
              await page
                .locator('input[name="txt_SDFA82V0REGKCCOE_ayuda"]')
                .wait();
              await page.type(
                'input[name="txt_SDFA82V0REGKCCOE_ayuda"]',
                String(clientes[i].ccc1),
              );

              //********
              // CCC2
              //********
              await page
                .locator('input[name="txt_SDFA82V0TESCCCOE_ayuda"]')
                .wait();
              await page.type(
                'input[name="txt_SDFA82V0TESCCCOE_ayuda"]',
                String(clientes[i].ccc2),
              );

              //********
              // CCC3
              //********
              await page.locator('input[name="txt_SDFA82V0CCONE"]').wait();
              await page.type(
                'input[name="txt_SDFA82V0CCONE"]',
                String(clientes[i].ccc3),
              );

              await this.esperar(1000);

              //*************
              // Continuar
              //*************
              await page.locator('input[name="btn_Sub2207001004_32"]').wait();
              await page.locator('input[name="btn_Sub2207001004_32"]').click();

              await this.esperar(2000);

              //********
              // CNAE25
              //********
              try {
                await page.waitForSelector(
                  'input[name="txt_SDFA82V1CNAE25E_ayuda"]',
                  { timeout: 5_000 },
                );
              } catch (e) {
                let mensajeError = await page.evaluate((texto) => {
                  try {
                    return Array.from(document.querySelectorAll("#DIL"))[0]
                      .innerText;
                  } catch (error) {
                    return false;
                  }
                });

                archivoCNAE
                  .sheet(0)
                  .cell(i + 5, 11)
                  .value("ERROR: " + mensajeError);

                await this.esperar(1000);
                continue;
              }

              await page
                .locator('input[name="txt_SDFA82V1CNAE25E_ayuda"]')
                .click();

              for (let i = 0; i < 6; i++) {
                await page.keyboard.press("Backspace");
              }

              await page.type(
                'input[name="txt_SDFA82V1CNAE25E_ayuda"]',
                String(clientes[i].cnae25),
              );

              //*************
              // Confirmar
              //*************
              await page.locator('input[name="btn_Sub2207001004_65"]').wait();
              await page.locator('input[name="btn_Sub2207001004_65"]').click();
              await this.esperar(1000);

              //*************
              // Confirmar2
              //*************
              await page.locator('input[name="btn_Sub2204701006_64"]').wait();
              await page.locator('input[name="btn_Sub2204701006_64"]').click();
              await this.esperar(1000);

              const confirmacion = await page.evaluate((texto) => {
                try {
                  return Array.from(document.querySelectorAll("#DIL"))[0]
                    .innerText;
                } catch (error) {
                  return false;
                }
              });

              if (confirmacion) {
                archivoCNAE
                  .sheet(0)
                  .cell(i + 5, 12)
                  .value(confirmacion);
              }

              await this.esperar(1000);
            } //Fin iteracion de clientes

            console.log("Escribiendo archivo...");
            console.log("Path: " + path.normalize(pathSalidaExcel));
            archivoCNAE
              .toFileAsync(
                path.normalize(
                  path.join(pathSalidaExcel, "CNAE-Procesado.xlsx"),
                ),
              )
              .then(() => {
                console.log("XLSX escrito correctamente");
              })
              .catch((err) => {
                console.log("Se ha producido un error interno: ");
                console.log(err);
                var tituloError =
                  "Se ha producido un error escribiendo el archivo: " +
                  path.normalize(pathSalidaExcel);
                resolve(false);
              });

            //Iniciando descarga de informes:
            for (var i = 0; i < clientes.length; i++) {
              //Recargar cada 10 clientes:
              if (i % 10 == 0 && i > 0) {
                //await browser.close();
                await page.close();
                page = await browser.newPage();

                // Configurar el comportamiento de descarga
                await page._client().send("Page.setDownloadBehavior", {
                  behavior: "allow",
                  downloadPath: pathSalida,
                });
                await page.setViewport({ width: 1080, height: 1024 });
              }

              console.log("Procesando cliente: " + i);
              console.log(clientes[i]);

              if (
                clientes[i].ccc == "" ||
                clientes[i].ccc == null ||
                clientes[i].ccc == undefined ||
                clientes[i].cnae25 == "" ||
                clientes[i].cnae25 == null ||
                clientes[i].cnae25 == undefined
              ) {
                clientes[i]["errores"] = ["Campos CCC o CNAE25 no definidos."];
                continue;
              }

              await page.goto(
                "https://w2.seg-social.es/Xhtml?JacadaApplicationName=SGIRED&TRANSACCION=ATR64&E=I&AP=AFIR",
                { waitUntil: "networkidle0" },
              );

              await this.esperar(1000);

              //********
              // CCC1
              //********
              await page.locator('input[name="txt_SDFREG62_ayuda"]').wait();
              await page.type(
                'input[name="txt_SDFREG62_ayuda"]',
                String(clientes[i].ccc1),
              );

              //********
              // CCC2
              //********
              await page.locator('input[name="txt_SDFTESO62"]').wait();
              await page.type(
                'input[name="txt_SDFTESO62"]',
                String(clientes[i].ccc2),
              );

              //********
              // CCC3
              //********
              await page.locator('input[name="txt_SDFNUM62"]').wait();
              await page.type(
                'input[name="txt_SDFNUM62"]',
                String(clientes[i].ccc3),
              );

              await this.esperar(1000);
              await page.select(
                'select[name="cbo_ListaTipoImpresion"]',
                "OnLine",
              );

              await this.esperar(1000);

              //*************
              // Generar Documento de ingreso
              //*************
              let nuevaPagina;
              try {
                [nuevaPagina] = await Promise.all([
                  new Promise((resolvePromise) => {
                    const timeout = setTimeout(() => {
                      resolvePromise(false);
                    }, 5000);

                    browser.once("targetcreated", async (target) => {
                      const newPage = await target.page();
                      newPage.on("response", async (response) => {
                        // Verificar si el contenido es un PDF
                        if (
                          !response.url().endsWith(".js") &&
                          !response.url().endsWith(".css") &&
                          response.url().startsWith("chrome-extension://")
                        ) {
                          console.log("PDF detectado:", response.url());

                          // Intercepta el PDF:
                          const pdfBuffer = await response.buffer();

                          // Guardar el PDF en el sistema de archivos
                          const filePath = path.join(
                            pathSalida,
                            clientes[i]["nombreArchivo"],
                          );
                          fs.writeFileSync(filePath, pdfBuffer);
                          console.log("PDF descargado en:", filePath);
                          resolvePromise(newPage);
                        }
                      });
                    });
                  }),

                  await page.locator('input[name="btn_Sub2207601004"]').wait(),
                  await page.locator('input[name="btn_Sub2207601004"]').click(),
                ]);
              } catch (e) {
                console.log("Error en catch");
              }

              await this.esperar(1000);
              //Comprueba si hubo error
              if (!nuevaPagina) {
                console.log("ERROR EN DESCARGA");
              } else {
                await nuevaPagina.close();
              }
              console.log("Nuevo cliente");
              await this.esperar(1000);
            }

            await browser.close();
            resolve(true);
          })
          .then(() => {})
          .catch((err) => {
            console.log("ERROR");

            throw err;
          });
      } catch (err) {
        var tituloError = "No se ha podido cargar el archivo";
        var mensajeError =
          "Se ha producido un error interno cargando los archivos.";
        mainProcess.mostrarError(tituloError, mensajeError).then((result) => {
          resolve(false);
        });
      }
    }).catch((err) => {
      console.log("Se ha producido un error interno: ");
      console.log(err);
      var tituloError = "No se ha podido cargar el archivo";
      var mensajeError =
        "Se ha producido un error interno cargando los archivos.";
      mainProcess.mostrarError(tituloError, mensajeError).then((result) => {
        resolve(false);
      });
    });
  }

  async cNAE25Autonomos(argumentos) {
    return new Promise((resolve) => {
      console.log("Archivo CNAE...");
      console.log(argumentos.formularioControl[1]);
      console.log("Ruta Google...");
      console.log(argumentos.formularioControl[0]);

      var archivoCNAEAutonomos = {};
      var clientes = [];
      var pathArchivoEtiquetas = argumentos.formularioControl[1];
      var pathSalidaExcel = path.join(
        path.normalize(argumentos.formularioControl[2]),
        "CNAE-Autonomos-Procesados",
      );
      var pathSalida = path.join(
        path.normalize(argumentos.formularioControl[2]),
        "CNAE-Autonomos-Procesados",
        "Resultados",
      );

      // Verificar si la carpeta "Resultados" existe y crearla si no
      if (!fs.existsSync(pathSalida)) {
        fs.mkdirSync(pathSalida, { recursive: true });
        console.log(`Carpeta creada: ${pathSalida}`);
      } else {
        console.log(`La carpeta ya existe: ${pathSalida}`);
      }

      try {
        XlsxPopulate.fromFileAsync(path.normalize(pathArchivoEtiquetas))
          .then(async (workbook) => {
            console.log("Archivo Cargado: CNAE Autonomos");
            archivoCNAEAutonomos = workbook;
            var columnas = archivoCNAEAutonomos
              .sheet(0)
              .usedRange()._numColumns;
            var filas = archivoCNAEAutonomos.sheet(0).usedRange()._numRows;
            var objetoCliente = {};

            var cabeceras = [];
            for (var i = 1; i <= columnas; i++) {
              cabeceras.push(archivoCNAEAutonomos.sheet(0).cell(4, i).value());
            }

            console.log("Cabeceras: " + cabeceras);

            for (var i = 5; i <= filas; i++) {
              objetoCliente = {};
              for (var j = 1; j <= columnas; j++) {
                if (
                  archivoCNAEAutonomos.sheet(0).cell(i, j).value() !== undefined
                ) {
                  switch (cabeceras[j - 1]) {
                    case "NAF (Autónomos)":
                      objetoCliente["naf"] = archivoCNAEAutonomos
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;

                    case "IPF (Autónomos)":
                      objetoCliente["ipf"] = archivoCNAEAutonomos
                        .sheet(0)
                        .cell(i, j)
                        .value()
                        .slice(-9);
                      break;

                    case "Literal CNAE09":
                      objetoCliente["literal"] = archivoCNAEAutonomos
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;

                    case "CNAE09":
                      objetoCliente["cnae09"] = archivoCNAEAutonomos
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;

                    case "CNAE25":
                      objetoCliente["cnae25"] = archivoCNAEAutonomos
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;
                  }
                }
              }

              objetoCliente["errores"] = [];

              if (
                objetoCliente.cnae25 !== "" &&
                objetoCliente.cnae25 !== null &&
                objetoCliente.cnae25 !== undefined
              ) {
                objetoCliente["nombreArchivo"] =
                  objetoCliente["naf"] + "-" + objetoCliente["cnae25"] + ".pdf";
                clientes.push(Object.assign({}, objetoCliente));
              }
            }

            //Creación del objeto clientes agrupado:
            const clientesAgrupados = [];
            const clientesEvitadosPorMismoCNAE09 = [];
            const mensajesError = [];
            clientes.forEach((item, index) => {
              mensajesError.push("");
              const existente = clientesAgrupados.find(
                (x) => x.naf === item.naf && x.ipf === item.ipf,
              );

              if (existente) {
                let mismoCNAE09 = false;
                for (let i = 0; i < existente.cnae25.length; i++) {
                  if (existente.cnae09[i] == item.cnae09) {
                    clientesEvitadosPorMismoCNAE09.push(index);
                    mismoCNAE09 = true;
                  }
                }

                if (!mismoCNAE09) {
                  existente.cnae09.push(item.cnae09);
                  existente.cnae25.push(item.cnae25);
                  existente.literal.push(item.literal);
                  existente.origen.push(index);
                  existente.nombreArchivo.push(
                    item.ipf + "-" + item.cnae25 + ".pdf",
                  );
                }
              } else {
                clientesAgrupados.push({
                  naf: item.naf,
                  ipf: item.ipf,
                  cnae09: [item.cnae09],
                  cnae25: [item.cnae25],
                  literal: [item.literal],
                  origen: [index],
                  nombreArchivo: [item.ipf + "-" + item.cnae25 + ".pdf"],
                });
              }
            });

            console.log("Clientes Agrupados: ");
            console.log(clientesAgrupados);
            console.log("Clientes Evitados: ");
            console.log(clientesEvitadosPorMismoCNAE09);
            console.log("Mensajes Error");
            console.log(mensajesError);

            var chromiumExecutablePath = path.normalize(
              argumentos.formularioControl[0],
            );

            //Inicio de procesamiento:
            const browser = await puppeteer.launch({
              executablePath: chromiumExecutablePath,
              headless: false,
            });
            console.log(browser.executablePath);

            var page = await browser.newPage();

            page.on("dialog", async (dialog) => {
              const tipo = dialog.type();
              if (tipo == "beforeunload") {
                await dialog.accept();
              }
            });
            // Configurar el comportamiento de descarga
            await page._client().send("Page.setDownloadBehavior", {
              behavior: "allow",
              downloadPath: pathSalida,
            });

            await page.setViewport({ width: 1080, height: 1024 });

            for (var i = 0; i < clientesAgrupados.length; i++) {
              //Recargar cada 10 clientes:
              if (i % 10 == 0 && i > 0) {
                //await browser.close();
                await page.close();
                page = await browser.newPage();

                page.on("dialog", async (dialog) => {
                  const tipo = dialog.type();
                  if (tipo == "beforeunload") {
                    await dialog.accept();
                  }
                });
                // Configurar el comportamiento de descarga
                await page._client().send("Page.setDownloadBehavior", {
                  behavior: "allow",
                  downloadPath: pathSalida,
                });
                await page.setViewport({ width: 1080, height: 1024 });
              }

              if (!Array.isArray(clientesAgrupados[i]["cnae25"])) {
                continue;
              }

              for (var j = 0; j < clientesAgrupados[i]["cnae25"].length; j++) {
                console.log("Procesando cliente: " + i + " - Seccion " + j);
                console.log(clientesAgrupados[i]);
                console.log(mensajesError);

                await page.goto(
                  "https://w2.seg-social.es/ProsaInternet/OnlineAccess?ARQ.SPM.ACTION=LOGIN&ARQ.SPM.APPTYPE=SERVICE&ARQ.IDAPP=XV26C010",
                  { waitUntil: "networkidle0" },
                );

                await this.esperar(500);

                //***************************
                // Pagina de condiciones:
                //***************************
                // Check Leido:
                await page.locator('input[id="CHK_LEIDO"]').wait();
                await page.locator('input[id="CHK_LEIDO"]').click();

                await this.esperar(500);

                // Boton Continuar:
                await page
                  .locator('button[name="SPM.ACC.CONTINUAR_CONS"]')
                  .wait();
                await page
                  .locator('button[name="SPM.ACC.CONTINUAR_CONS"]')
                  .click();

                await this.esperar(500);

                //********
                // Numero Seguridad Social
                //********
                await page.locator('input[name="NSS"]').wait();
                await page.type(
                  'input[name="NSS"]',
                  String(clientesAgrupados[i].naf),
                );

                //********
                // Seleccionar DNI
                //********
                await page.locator('select[name="IPF_TIPO"]').wait();
                await page.select('select[name="IPF_TIPO"]', "1");

                //********
                // Introducir DNI
                //********
                await page.locator('input[name="IPF_NUMERO"]').wait();
                await page.type(
                  'input[name="IPF_NUMERO"]',
                  String(clientesAgrupados[i].ipf),
                );

                await this.esperar(1000);

                //*************
                // Continuar
                //*************
                await page
                  .locator('button[name="SPM.ACC.CONTINUAR_TRAB"]')
                  .wait();
                await page
                  .locator('button[name="SPM.ACC.CONTINUAR_TRAB"]')
                  .click();

                await this.esperar(2000);

                //********************************
                // Buscar CNAE en las filas
                //********************************
                const actividadBuscada =
                  clientesAgrupados[i].literal[j].length > 40
                    ? clientesAgrupados[i].literal[j].slice(0, 40)
                    : clientesAgrupados[i].literal[j];

                await page
                  .locator('table[id="titulo_tabla_actividades"]')
                  .wait();
                const botonActividadBuscada = await page.evaluate(
                  (actividadBuscada) => {
                    const filas = Array.from(
                      document.querySelectorAll(
                        "#titulo_tabla_actividades tbody tr",
                      ),
                    );

                    for (const fila of filas) {
                      const celdas = fila.querySelectorAll("td");
                      const actividad = celdas[2]?.textContent?.trim();
                      console.log("evaluando: ", actividad);

                      if (actividad === actividadBuscada) {
                        console.log("Fila encontrada: " + actividadBuscada);
                        // Buscar enlace con texto "Comunicar CNAE25" dentro de esta fila
                        const enlaces = fila.querySelectorAll(
                          'a[title="Comunicar CNAE25"]',
                        );
                        if (enlaces.length > 0) {
                          enlaces[0].click();
                          return enlaces[0];
                        }
                      }
                    }
                    return false;
                  },
                  actividadBuscada,
                );

                if (!botonActividadBuscada) {
                  mensajesError[clientesAgrupados[i].origen[j]] =
                    "Error, no se ha encontrado la actividad económica con el literal " +
                    clientesAgrupados[i].literal[j];
                  continue;
                }

                //********
                // SELECCIONAR CNAE25
                //********
                await page.locator('select[name="CNAE25"]').wait();

                const existeOption = await page.evaluate(() => {
                  const select = document.querySelector(
                    'select[name="CNAE25"]',
                  ); // usa tu selector real
                  if (!select) return false;
                  return true;
                });

                if (existeOption) {
                  await page.select(
                    'select[name="CNAE25"]',
                    String(clientesAgrupados[i].cnae25[j]),
                  );
                  await this.esperar(1000);
                  await page
                    .locator('button[name="SPM.ACC.CONTINUAR_ANOTAR_CNAE25"]')
                    .wait();
                  await page
                    .locator('button[name="SPM.ACC.CONTINUAR_ANOTAR_CNAE25"]')
                    .click();

                  await page.waitForNavigation({ waitUntil: "load" });
                  await this.esperar(1000);

                  const yaNotificado = await page.evaluate(() => {
                    const mensajes = Array.from(
                      document.querySelectorAll(
                        "#lista_mensajes li:not(.pr_oculto) .pr_pMensaje",
                      ),
                    );
                    return mensajes.some(
                      (el) =>
                        el.textContent.trim() ===
                        "La actividad económica 2025 debe ser distinta de la ya comunicada",
                    );
                  });

                  console.log("yaNotificado: ", yaNotificado);
                  if (yaNotificado) {
                    mensajesError[clientesAgrupados[i].origen[j]] =
                      "OK, CNAE25 ya notificado";
                    continue;
                  }

                  await page
                    .locator(
                      'button[name="SPM.ACC.CONTINUAR_CONFIRMAR_CNAE25"]',
                    )
                    .wait();
                  await page
                    .locator(
                      'button[name="SPM.ACC.CONTINUAR_CONFIRMAR_CNAE25"]',
                    )
                    .click();

                  await page.waitForNavigation({ waitUntil: "load" });
                  await this.esperar(1000);

                  const exito = await page.evaluate(() => {
                    const mensajes = Array.from(
                      document.querySelectorAll(
                        "#lista_mensajes li:not(.pr_oculto) .pr_pMensaje",
                      ),
                    );
                    return mensajes.some(
                      (el) =>
                        el.textContent.trim() ===
                        "Operación realizada correctamente.",
                    );
                  });

                  console.log("exito: ", exito);
                  if (!exito) {
                    mensajesError[clientesAgrupados[i].origen[j]] =
                      "ERROR: Error indeterminado";
                    continue;
                  }

                  console.log("✅ Éxito detectado");
                  await page.locator('button[title="Cerrar"]').wait();
                  await page.locator('button[title="Cerrar"]').click();

                  await this.esperar(1000);
                  await page.locator('a[data-pc_tipo="documento"]').wait(),
                    console.log("Descargando Resguardo");
                  //*************
                  // DESCARGANDO RESGUARDO
                  //*************
                  let nuevaPagina;
                  try {
                    [nuevaPagina] = await Promise.all([
                      new Promise((resolvePromise) => {
                        const timeout = setTimeout(() => {
                          resolvePromise(false);
                        }, 5000);

                        browser.once("targetcreated", async (target) => {
                          const newPage = await target.page();
                          newPage.on("response", async (response) => {
                            // Verificar si el contenido es un PDF
                            if (
                              !response.url().endsWith(".js") &&
                              !response.url().endsWith(".css") &&
                              response.url().startsWith("chrome-extension://")
                            ) {
                              console.log("PDF detectado:", response.url());

                              // Intercepta el PDF:
                              const pdfBuffer = await response.buffer();

                              // Guardar el PDF en el sistema de archivos
                              const filePath = path.join(
                                pathSalida,
                                clientesAgrupados[i]["nombreArchivo"][j],
                              );
                              fs.writeFileSync(filePath, pdfBuffer);
                              console.log("PDF descargado en:", filePath);
                              mensajesError[clientesAgrupados[i].origen[j]] =
                                "OK: Resguardo descargado correctamente";
                              resolvePromise(newPage);
                            }
                          });
                        });
                      }),

                      await page.locator('a[data-pc_tipo="documento"]').wait(),
                      await page.locator('a[data-pc_tipo="documento"]').click(),
                    ]);
                  } catch (e) {
                    console.log("Error en catch");
                    mensajesError[clientesAgrupados[i].origen[j]] =
                      "ERROR: Error en la descarga del resguardo";
                    continue;
                  }

                  //Comprueba si hubo error
                  if (!nuevaPagina) {
                    console.log("ERROR EN DESCARGA");
                    mensajesError[clientesAgrupados[i].origen[j]] =
                      "ERROR: Error en la descarga del resguardo";
                    continue;
                  } else {
                    await nuevaPagina.close();
                  }
                } else {
                  mensajesError[clientesAgrupados[i].origen[j]] =
                    "ERROR: No se ha encontrado la opción con el CNAE25 especificado";
                  continue;
                }

                await this.esperar(1000);
              }
            } //Fin iteracion de clientes

            //ESCRIBIR EXCEL FINAL:
            console.log("Escribiendo archivo...");
            for (var k = 0; k < mensajesError.length; k++) {
              archivoCNAEAutonomos
                .sheet(0)
                .cell(k + 5, 10)
                .value(mensajesError[k]);
            }

            for (var k = 0; k < clientesEvitadosPorMismoCNAE09.length; k++) {
              archivoCNAEAutonomos
                .sheet(0)
                .cell(k + 5, 10)
                .value("Evitando por duplicidad en la declaración de CNAE25");
            }

            console.log("Path: " + path.normalize(pathSalidaExcel));
            archivoCNAEAutonomos
              .toFileAsync(
                path.normalize(
                  path.join(pathSalidaExcel, "CNAE25-Autonomos-Procesado.xlsx"),
                ),
              )
              .then(() => {
                console.log("XLSX escrito correctamente");
              })
              .catch((err) => {
                console.log("Se ha producido un error interno: ");
                console.log(err);
                var tituloError =
                  "Se ha producido un error escribiendo el archivo: " +
                  path.normalize(pathSalidaExcel);
                resolve(false);
              });

            await browser.close();
            resolve(true);
          })
          .catch((err) => {
            console.log("ERROR");

            throw err;
          });
      } catch (err) {
        var tituloError = "No se ha podido cargar el archivo";
        var mensajeError =
          "Se ha producido un error interno cargando los archivos.";
        mainProcess.mostrarError(tituloError, mensajeError).then((result) => {
          resolve(false);
        });
      }
    }).catch((err) => {
      console.log("Se ha producido un error interno: ");
      console.log(err);
      var tituloError = "No se ha podido cargar el archivo";
      var mensajeError =
        "Se ha producido un error interno cargando los archivos.";
      mainProcess.mostrarError(tituloError, mensajeError).then((result) => {
        resolve(false);
      });
    });
  }

  async informesITA(argumentos) {
    return new Promise((resolve) => {
      console.log("Archivo ITA...");
      console.log(argumentos.formularioControl[1]);
      console.log("Ruta Google...");
      console.log(argumentos.formularioControl[0]);

      var archivoITA = {};
      var clientes = [];
      var pathArchivoEtiquetas = argumentos.formularioControl[1];
      var pathSalidaExcel = path.join(
        path.normalize(argumentos.formularioControl[2]),
        "ITA-Informes-Procesados",
      );
      var pathSalida = path.join(
        path.normalize(argumentos.formularioControl[2]),
        "ITA-Informes-Procesados",
        "Resultados",
      );

      // Verificar si la carpeta "Resultados" existe y crearla si no
      if (!fs.existsSync(pathSalida)) {
        fs.mkdirSync(pathSalida, { recursive: true });
        console.log(`Carpeta creada: ${pathSalida}`);
      } else {
        console.log(`La carpeta ya existe: ${pathSalida}`);
      }

      try {
        XlsxPopulate.fromFileAsync(path.normalize(pathArchivoEtiquetas))
          .then(async (workbook) => {
            console.log("Archivo Cargado: ITA");
            archivoITA = workbook;
            var columnas = archivoITA.sheet(0).usedRange()._numColumns;
            var filas = archivoITA.sheet(0).usedRange()._numRows;
            var objetoCliente = {};

            var cabeceras = [];
            for (var i = 1; i <= columnas; i++) {
              cabeceras.push(archivoITA.sheet(0).cell(1, i).value());
            }

            console.log("Cabeceras: " + cabeceras);
            for (var i = 2; i <= filas; i++) {
              objetoCliente = {};
              for (var j = 1; j <= columnas; j++) {
                if (archivoITA.sheet(0).cell(i, j).value() !== undefined) {
                  switch (cabeceras[j - 1]) {
                    case "Código Cuenta Cotización (CCC)":
                      objetoCliente["ccc"] = archivoITA
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      objetoCliente["ccc1"] = objetoCliente["ccc"].substring(
                        0,
                        4,
                      );
                      objetoCliente["ccc2"] = objetoCliente["ccc"].substring(
                        4,
                        6,
                      );
                      objetoCliente["ccc3"] = objetoCliente["ccc"].substring(6);
                      break;

                    case "Expediente":
                      objetoCliente["expediente"] = archivoITA
                        .sheet(0)
                        .cell(i, j)
                        .value();
                      break;
                  }
                }
              }

              objetoCliente["errores"] = [];

              if (
                objetoCliente.ccc !== "" &&
                objetoCliente.ccc !== null &&
                objetoCliente.ccc !== undefined
              ) {
                objetoCliente["nombreArchivo"] =
                  objetoCliente["expediente"] +
                  "-" +
                  objetoCliente["ccc"] +
                  ".pdf";
                clientes.push(Object.assign({}, objetoCliente));
              }
            }

            console.log("Clientes: ");
            console.log(clientes);

            var chromiumExecutablePath = path.normalize(
              argumentos.formularioControl[0],
            );

            //Inicio de procesamiento:
            const browser = await puppeteer.launch({
              executablePath: chromiumExecutablePath,
              headless: false,
            });
            console.log(browser.executablePath);

            var page = await browser.newPage();

            // Configurar el comportamiento de descarga
            await page._client().send("Page.setDownloadBehavior", {
              behavior: "allow",
              downloadPath: pathSalida,
            });

            await page.setViewport({ width: 1080, height: 1024 });

            //Iniciando descarga de informes:
            for (var i = 0; i < clientes.length; i++) {
              //Recargar cada 10 clientes:
              if (i % 10 == 0 && i > 0) {
                //await browser.close();
                await page.close();
                page = await browser.newPage();

                // Configurar el comportamiento de descarga
                await page._client().send("Page.setDownloadBehavior", {
                  behavior: "allow",
                  downloadPath: pathSalida,
                });
                await page.setViewport({ width: 1080, height: 1024 });
              }

              console.log("Procesando cliente: " + i);
              console.log(clientes[i]);

              if (
                clientes[i].ccc == "" ||
                clientes[i].ccc == null ||
                clientes[i].ccc == undefined
              ) {
                clientes[i]["errores"] = ["Campo CCC no definidos."];
                continue;
              }

              await page.goto(
                "https://w2.seg-social.es/Xhtml?JacadaApplicationName=SGIRED&TRANSACCION=ATR64&E=I&AP=AFIR",
                { waitUntil: "networkidle0" },
              );

              await this.esperar(1000);

              //********
              // CCC1
              //********
              await page.locator('input[name="txt_SDFREG62_ayuda"]').wait();
              await page.type(
                'input[name="txt_SDFREG62_ayuda"]',
                String(clientes[i].ccc1),
              );

              //********
              // CCC2
              //********
              await page.locator('input[name="txt_SDFTESO62"]').wait();
              await page.type(
                'input[name="txt_SDFTESO62"]',
                String(clientes[i].ccc2),
              );

              //********
              // CCC3
              //********
              await page.locator('input[name="txt_SDFNUM62"]').wait();
              await page.type(
                'input[name="txt_SDFNUM62"]',
                String(clientes[i].ccc3),
              );

              await this.esperar(1000);
              await page.select(
                'select[name="cbo_ListaTipoImpresion"]',
                "OnLine",
              );

              await this.esperar(1000);

              //*************
              // Generar Documento de ingreso
              //*************
              let nuevaPagina;
              try {
                [nuevaPagina] = await Promise.all([
                  new Promise((resolvePromise) => {
                    setTimeout(() => {
                      resolvePromise(false);
                    }, 5000);

                    browser.once("targetcreated", async (target) => {
                      const newPage = await target.page();
                      newPage.on("response", async (response) => {
                        // Verificar si el contenido es un PDF
                        if (
                          !response.url().endsWith(".js") &&
                          !response.url().endsWith(".css") &&
                          response.url().startsWith("chrome-extension://")
                        ) {
                          console.log("PDF detectado:", response.url());

                          // Intercepta el PDF:
                          const pdfBuffer = await response.buffer();

                          // Guardar el PDF en el sistema de archivos
                          const filePath = path.join(
                            pathSalida,
                            clientes[i]["nombreArchivo"],
                          );
                          fs.writeFileSync(filePath, pdfBuffer);
                          console.log("PDF descargado en:", filePath);
                          resolvePromise(newPage);
                        }
                      });
                    });
                  }),

                  await page.locator('input[name="btn_Sub2207601004"]').wait(),
                  await page.locator('input[name="btn_Sub2207601004"]').click(),
                ]);
              } catch (e) {
                console.log("Error en catch");
              }

              await this.esperar(1000);
              //Comprueba si hubo error
              if (!nuevaPagina) {
                console.log("ERROR EN DESCARGA");
                archivoITA
                  .sheet(0)
                  .cell(i + 2, 3)
                  .value("ERROR: No se ha podido descargar el informe.");
              } else {
                archivoITA
                  .sheet(0)
                  .cell(i + 2, 3)
                  .value("OK");
                await nuevaPagina.close();
              }
              console.log("Nuevo cliente");
              await this.esperar(1000);
            }

            await browser.close();

            console.log("Escribiendo archivo...");
            console.log("Path: " + path.normalize(pathSalidaExcel));
            archivoITA
              .toFileAsync(
                path.normalize(
                  path.join(pathSalidaExcel, "ITA-Procesado.xlsx"),
                ),
              )
              .then(() => {
                console.log("XLSX escrito correctamente");
              })
              .catch((err) => {
                console.log("Se ha producido un error interno: ");
                console.log(err);
                var tituloError =
                  "Se ha producido un error escribiendo el archivo: " +
                  path.normalize(pathSalidaExcel);
                resolve(false);
              });

            resolve(true);
          })
          .then(() => {})
          .catch((err) => {
            console.log("ERROR");

            throw err;
          });
      } catch (err) {
        var tituloError = "No se ha podido cargar el archivo";
        var mensajeError =
          "Se ha producido un error interno cargando los archivos.";
        mainProcess.mostrarError(tituloError, mensajeError).then((result) => {
          resolve(false);
        });
      }
    }).catch((err) => {
      console.log("Se ha producido un error interno: ");
      console.log(err);
      var tituloError = "No se ha podido cargar el archivo";
      var mensajeError =
        "Se ha producido un error interno cargando los archivos.";
      mainProcess.mostrarError(tituloError, mensajeError).then((result) => {
        resolve(false);
      });
    });
  }

  async certificadoSeguridadSocial(argumentos) {
    return new Promise((resolve) => {
      console.log("Archivo SS CCC...");
      console.log(argumentos.formularioControl[1]);
      console.log("Ruta Google...");
      console.log(argumentos.formularioControl[0]);

      var archivoSS = {};
      var clientes = [];
      var pathArchivoEtiquetas = argumentos.formularioControl[1];
      var pathSalidaExcel = path.join(
        path.normalize(argumentos.formularioControl[2]),
        "SS-Certificados-Procesados",
      );
      var pathSalida = path.join(
        path.normalize(argumentos.formularioControl[2]),
        "SS-Certificados-Procesados",
        "Resultados",
      );
      var pathSalidaFacturacion = path.join(
        path.normalize(argumentos.formularioControl[2]),
        "SS-Certificados-Procesados",
        "FACTURACIÓN",
      );

      // Verificar si la carpeta "Resultados" existe y crearla si no
      if (!fs.existsSync(pathSalida)) {
        fs.mkdirSync(pathSalida, { recursive: true });
        console.log(`Carpeta creada: ${pathSalida}`);
      } else {
        console.log(`La carpeta ya existe: ${pathSalida}`);
      }

      // Verificar si la carpeta "FACTURACIÓN" existe y crearla si no
      if (!fs.existsSync(pathSalidaFacturacion)) {
        fs.mkdirSync(pathSalidaFacturacion, { recursive: true });
        console.log(`Carpeta creada: ${pathSalidaFacturacion}`);
      } else {
        console.log(`La carpeta ya existe: ${pathSalidaFacturacion}`);
      }

      try {
        XlsxPopulate.fromFileAsync(path.normalize(pathArchivoEtiquetas))
          .then(async (workbook) => {
            console.log("Archivo Cargado: SS");
            archivoSS = workbook;
            var columnas = archivoSS.sheet("DATOS").usedRange()._numColumns;
            var filas = archivoSS.sheet("DATOS").usedRange()._numRows;
            var objetoCliente = {};

            var cabeceras = [];
            for (var i = 1; i <= columnas; i++) {
              cabeceras.push(archivoSS.sheet("DATOS").cell(1, i).value());
            }

            console.log("Cabeceras: " + cabeceras);
            for (var i = 2; i <= filas; i++) {
              objetoCliente = {};
              for (var j = 1; j <= columnas; j++) {
                if (archivoSS.sheet("DATOS").cell(i, j).value() !== undefined) {
                  switch (cabeceras[j - 1]) {
                    case "CCC COMPLETO":
                      objetoCliente["ccc"] = archivoSS
                        .sheet("DATOS")
                        .cell(i, j)
                        .value();
                      break;

                    case "EMPRESA":
                      objetoCliente["empresa"] = archivoSS
                        .sheet("DATOS")
                        .cell(i, j)
                        .value();
                      break;

                    case "CÓDIGO":
                      objetoCliente["codigo"] = archivoSS
                        .sheet("DATOS")
                        .cell(i, j)
                        .value();
                      break;
                  }
                }
              }

              objetoCliente["errores"] = [];

              if (
                objetoCliente.ccc !== "" &&
                objetoCliente.ccc !== null &&
                objetoCliente.ccc !== undefined
              ) {
                objetoCliente["nombreArchivo"] =
                  objetoCliente["codigo"] +
                  " CERTIFICADO ESTAR AL CORRIENTE SS " +
                  objetoCliente["empresa"] +
                  " " +
                  DateTime.now().setZone("Europe/Madrid").toFormat("ddMMyy") +
                  ".pdf";

                objetoCliente["nombreArchivoFacturacion"] =
                  objetoCliente["codigo"] +
                  "-" +
                  objetoCliente["empresa"] +
                  "-3.096-Certificado de estar al corriente AEAT-28.50€-CC.pdf";
                clientes.push(Object.assign({}, objetoCliente));
              }
            }

            console.log("Clientes: ");
            console.log(clientes);

            var chromiumExecutablePath = path.normalize(
              argumentos.formularioControl[0],
            );

            //Inicio de procesamiento:
            const browser = await puppeteer.launch({
              executablePath: chromiumExecutablePath,
              headless: false,
            });
            console.log(browser.executablePath);

            var page = await browser.newPage();

            //Confirma el cambio de pagina:
            page.on("dialog", async (dialog) => {
              const tipo = dialog.type();
              if (tipo == "beforeunload") {
                await dialog.accept();
              }
            });

            // Configurar el comportamiento de descarga
            await page._client().send("Page.setDownloadBehavior", {
              behavior: "allow",
              downloadPath: pathSalida,
            });

            await page.setViewport({ width: 1080, height: 1024 });

            //Iniciando descarga de informes:
            for (var i = 0; i < clientes.length; i++) {
              //Recargar cada 10 clientes:
              if (i % 10 == 0 && i > 0) {
                //await browser.close();
                await page.close();
                page = await browser.newPage();

                // Configurar el comportamiento de descarga
                await page._client().send("Page.setDownloadBehavior", {
                  behavior: "allow",
                  downloadPath: pathSalida,
                });
                await page.setViewport({ width: 1080, height: 1024 });
              }

              console.log("Procesando cliente: " + i);
              console.log(clientes[i]);

              if (
                clientes[i].ccc == "" ||
                clientes[i].ccc == null ||
                clientes[i].ccc == undefined
              ) {
                clientes[i]["errores"] = ["Campo CCC no definidos."];
                continue;
              }

              await page.goto(
                "https://w2.seg-social.es/ProsaInternet/OnlineAccess?ARQ.SPM.ACTION=LOGIN&ARQ.SPM.APPTYPE=SERVICE&ARQ.IDAPP=XV21F001",
                { waitUntil: "networkidle0" },
              );

              //********
              // Pulsar CODIGO ARED
              //********
              await page.locator('a[id="enlace_316077"]').click();

              //********
              // Pulsar Buscar
              //********
              await page
                .locator('button[name="SPM.ACC.AC_BUSCAR_OAR"]')
                .click();

              //********
              // Pulsar opcion CCC/NAF
              //********
              await page.locator(`input[title="Buscar por CCC o NAF"]`).wait();
              var radioButton = await page.$(
                `input[title="Buscar por CCC o NAF"]`,
              );

              if (radioButton) {
                await radioButton.click(); // Hacer clic en el radio button
                console.log("Radio button seleccionado.");
              } else {
                console.log("No se encontró el radio button.");
              }

              //********
              // Introducir CCC
              //********
              await page
                .locator('input[name="criteriosBusquedaCccNaf"]')
                .wait();
              await page.type(
                'input[name="criteriosBusquedaCccNaf"]',
                String(clientes[i].ccc),
              );

              await this.esperar(1000);

              //********
              // Pulsar Buscar
              //********
              await page
                .locator('button[name="SPM.ACC.AC_BUSCAR_OAR"]')
                .click();

              //********
              // Pulsar opcion encontrada
              //********
              await page
                .locator(
                  'a[id="enlace_' + String(Number(clientes[i].ccc)) + '"]',
                )
                .click();

              //********
              // Pulsar boton "Continuar"
              //********
              await page.locator('button[name="SPM.ACC.CONTINUAR"]').click();

              //********
              // Pulsar boton "Imprimir"
              //********
              await page.locator('button[name="SPM.ACC.IMPRIMIR"]').click();
              await page.waitForNavigation({ waitUntil: "load" });

              //*************
              // Generar Documento de certificado
              //*************
              const enlaces = await page.$$("a");
              let enlaceEncontrado = null;

              for (const enlace of enlaces) {
                const texto = await page.evaluate((el) => el.innerText, enlace);
                if (texto.includes("Certificado genérico")) {
                  enlaceEncontrado = enlace;
                  break;
                }
              }

              let nuevaPagina;
              try {
                [nuevaPagina] = await Promise.all([
                  new Promise((resolvePromise) => {
                    setTimeout(() => {
                      resolvePromise(false);
                    }, 10000);

                    browser.once("targetcreated", async (target) => {
                      const newPage = await target.page();
                      newPage.on("response", async (response) => {
                        // Verificar si el contenido es un PDF
                        if (
                          !response.url().endsWith(".js") &&
                          !response.url().endsWith(".css") &&
                          response.url().startsWith("chrome-extension://")
                        ) {
                          console.log("PDF detectado:", response.url());

                          // Intercepta el PDF:
                          const pdfBuffer = await response.buffer();

                          // Guardar el PDF en el sistema de archivos
                          const filePath = path.join(
                            pathSalida,
                            clientes[i]["nombreArchivo"],
                          );
                          const filePathFacturacion = path.join(
                            pathSalidaFacturacion,
                            clientes[i]["nombreArchivoFacturacion"],
                          );

                          fs.writeFileSync(filePath, pdfBuffer);
                          fs.writeFileSync(filePathFacturacion, pdfBuffer);

                          console.log("PDF descargado en:", filePath);
                          resolvePromise(newPage);
                        }
                      });
                    });
                  }),
                  await enlaceEncontrado.click(),
                ]);
              } catch (e) {
                console.log("Error en catch: ", e);
              }

              await this.esperar(1000);

              //Comprueba si hubo error
              if (!nuevaPagina) {
                console.log("ERROR EN DESCARGA");
                archivoSS
                  .sheet("DATOS")
                  .cell(i + 2, 8)
                  .value("ERROR: No se ha podido descargar el certificado.");
              } else {
                archivoSS
                  .sheet("DATOS")
                  .cell(i + 2, 8)
                  .value("OK, certificado descargado.");
                await nuevaPagina.close();
              }
              console.log("Nuevo cliente");
              await this.esperar(1000);
            }

            await browser.close();

            console.log("Escribiendo archivo...");
            console.log("Path: " + path.normalize(pathSalidaExcel));
            archivoSS
              .toFileAsync(
                path.normalize(path.join(pathSalidaExcel, "SS-Procesado.xlsx")),
              )
              .then(() => {
                console.log("XLSX escrito correctamente");
              })
              .catch((err) => {
                console.log("Se ha producido un error interno: ");
                console.log(err);
                var tituloError =
                  "Se ha producido un error escribiendo el archivo: " +
                  path.normalize(pathSalidaExcel);
                resolve(false);
              });

            resolve(true);
          })
          .then(() => {})
          .catch((err) => {
            console.log("ERROR");

            throw err;
          });
      } catch (err) {
        var tituloError = "No se ha podido cargar el archivo";
        var mensajeError =
          "Se ha producido un error interno cargando los archivos.";
        mainProcess.mostrarError(tituloError, mensajeError).then((result) => {
          resolve(false);
        });
      }
    }).catch((err) => {
      console.log("Se ha producido un error interno: ");
      console.log(err);
      var tituloError = "No se ha podido cargar el archivo";
      var mensajeError =
        "Se ha producido un error interno cargando los archivos.";
      mainProcess.mostrarError(tituloError, mensajeError).then((result) => {
        resolve(false);
      });
    });
  }

  async certificadoTributario(argumentos) {
    return new Promise((resolve) => {
      console.log("Archivo Tributario CCC...");
      console.log(argumentos.formularioControl[1]);
      console.log("Ruta Google...");
      console.log(argumentos.formularioControl[0]);

      var archivoTributario = {};
      var clientes = [];
      var pathArchivoEtiquetas = argumentos.formularioControl[1];
      var pathSalidaExcel = path.join(
        path.normalize(argumentos.formularioControl[2]),
        "Certificados_Tributarios-Procesados",
      );
      var pathSalida = path.join(
        path.normalize(argumentos.formularioControl[2]),
        "Certificados_Tributarios-Procesados",
        "Resultados",
      );
      var pathSalidaFacturacion = path.join(
        path.normalize(argumentos.formularioControl[2]),
        "Certificados_Tributarios-Procesados",
        "FACTURACIÓN",
      );

      // Verificar si la carpeta "Resultados" existe y crearla si no
      if (!fs.existsSync(pathSalida)) {
        fs.mkdirSync(pathSalida, { recursive: true });
        console.log(`Carpeta creada: ${pathSalida}`);
      } else {
        console.log(`La carpeta ya existe: ${pathSalida}`);
      }

      // Verificar si la carpeta "Resultados" existe y crearla si no
      if (!fs.existsSync(pathSalidaFacturacion)) {
        fs.mkdirSync(pathSalidaFacturacion, { recursive: true });
        console.log(`Carpeta creada: ${pathSalidaFacturacion}`);
      } else {
        console.log(`La carpeta ya existe: ${pathSalidaFacturacion}`);
      }

      try {
        XlsxPopulate.fromFileAsync(path.normalize(pathArchivoEtiquetas))
          .then(async (workbook) => {
            console.log("Archivo Cargado: Tributario");
            archivoTributario = workbook;
            var columnas = archivoTributario
              .sheet("DATOS")
              .usedRange()._numColumns;
            var filas = archivoTributario.sheet("DATOS").usedRange()._numRows;
            var objetoCliente = {};

            var cabeceras = [];
            for (var i = 1; i <= columnas; i++) {
              cabeceras.push(
                archivoTributario.sheet("DATOS").cell(1, i).value(),
              );
            }

            console.log("Cabeceras: " + cabeceras);
            for (var i = 2; i <= filas; i++) {
              objetoCliente = {};
              for (var j = 1; j <= columnas; j++) {
                if (
                  archivoTributario.sheet("DATOS").cell(i, j).value() !==
                  undefined
                ) {
                  switch (cabeceras[j - 1]) {
                    case "CCC COMPLETO":
                      objetoCliente["ccc"] = archivoTributario
                        .sheet("DATOS")
                        .cell(i, j)
                        .value();
                      break;

                    case "EMPRESA":
                      objetoCliente["empresa"] = archivoTributario
                        .sheet("DATOS")
                        .cell(i, j)
                        .value();
                      break;

                    case "CÓDIGO":
                      objetoCliente["codigo"] = archivoTributario
                        .sheet("DATOS")
                        .cell(i, j)
                        .value();
                      break;

                    case "NIF":
                      objetoCliente["nif"] = archivoTributario
                        .sheet("DATOS")
                        .cell(i, j)
                        .value();
                      break;
                  }
                }
              }

              objetoCliente["errores"] = [];
              objetoCliente["flagEvitarDuplicado"] = false;

              if (
                objetoCliente.ccc !== "" &&
                objetoCliente.ccc !== null &&
                objetoCliente.ccc !== undefined
              ) {
                objetoCliente["nombreArchivo"] =
                  objetoCliente["codigo"] +
                  " CERTIFICADO ESTAR AL CORRIENTE AEAT " +
                  objetoCliente["empresa"] +
                  " " +
                  DateTime.now().setZone("Europe/Madrid").toFormat("ddMMyy") +
                  ".pdf";

                objetoCliente["nombreArchivoFacturacion"] =
                  objetoCliente["codigo"] +
                  "-" +
                  objetoCliente["empresa"] +
                  "-3.096-Certificado de estar al corriente AEAT-28.50€-CC.pdf";
                clientes.push(Object.assign({}, objetoCliente));
              }
            }

            //Procesar duplicados:
            const vistos = new Set();
            clientes = clientes.map((obj) => {
              if (vistos.has(obj.nif)) {
                obj["errores"] = [
                  "Evitando generar certificado por NIF duplicado",
                ];
                return { ...obj, flagEvitarDuplicado: true };
              } else {
                vistos.add(obj.nif);
                return obj;
              }
            });

            console.log("Clientes: ");
            console.log(clientes);

            var chromiumExecutablePath = path.normalize(
              argumentos.formularioControl[0],
            );

            //Inicio de procesamiento:
            const browser = await puppeteer.launch({
              executablePath: chromiumExecutablePath,
              headless: false,
            });
            console.log(browser.executablePath);

            var page = await browser.newPage();

            //Confirma el cambio de pagina:
            page.on("dialog", async (dialog) => {
              console.log("Dialogo: ", dialog.type());
              const tipo = dialog.type();
              if (tipo == "beforeunload") {
                await dialog.accept();
              }
            });

            // Configurar el comportamiento de descarga
            await page._client().send("Page.setDownloadBehavior", {
              behavior: "allow",
              downloadPath: pathSalida,
            });

            await page.setViewport({ width: 1080, height: 1024 });

            //Iniciando descarga de informes:
            for (var i = 0; i < clientes.length; i++) {
              if (clientes[i].flagEvitarDuplicado) {
                archivoTributario
                  .sheet("DATOS")
                  .cell(i + 2, 8)
                  .value("WARNING: Solicitud evitada por duplicidad en NIF.");
                continue;
              }

              //Recargar cada 10 clientes:
              if (i % 10 == 0 && i > 0) {
                //await browser.close();
                await page.close();
                page = await browser.newPage();

                // Configurar el comportamiento de descarga
                await page._client().send("Page.setDownloadBehavior", {
                  behavior: "allow",
                  downloadPath: pathSalida,
                });
                await page.setViewport({ width: 1080, height: 1024 });
              }

              console.log("Procesando cliente: " + i);
              console.log(clientes[i]);

              if (
                clientes[i].ccc == "" ||
                clientes[i].ccc == null ||
                clientes[i].ccc == undefined
              ) {
                clientes[i]["errores"] = ["Campo CCC no definidos."];
                continue;
              }

              await page.goto(
                "https://www1.agenciatributaria.gob.es/wlpl/EMCE-JDIT/ECOTInternetCiudadanosServlet",
                { waitUntil: "networkidle0" },
              );

              try {
                const botonModal = await page.waitForSelector(
                  'button[data-dismiss="modal"]',
                  { timeout: 1000 },
                );
                if (botonModal) {
                  await botonModal.click();
                  console.log("Botón clicado");
                }
              } catch (error) {
                console.log(
                  "Botón no encontrado después de 1 segundo, continuando...",
                );
              }

              //********
              // Pulsar opcion en representacion de terceros
              //********
              await page.locator(`input[id="fTipoRepresentacion1"]`).wait();
              var radioButton = await page.$(
                `input[id="fTipoRepresentacion1"]`,
              );

              if (radioButton) {
                await radioButton.click(); // Hacer clic en el radio button
                console.log("Radio button seleccionado.");
              } else {
                console.log("No se encontró el radio button.");
              }

              //********
              // Introducir NIE
              //********
              await page.locator('input[name="fNifT"]').wait();
              await page.type('input[name="fNifT"]', String(clientes[i].nif));
              await this.esperar(500);

              //********
              // Introducir Nombre y apellidos
              //********
              await page.locator('input[name="fNombreT"]').wait();
              await page.type(
                'input[name="fNombreT"]',
                String(clientes[i].empresa),
              );
              await this.esperar(500);

              //********
              // Pulsar opcion tipo de certificado
              //********
              await page.locator(`input[id="fTipoCertificado4"]`).wait();
              var radioButton2 = await page.$(`input[id="fTipoCertificado4"]`);

              if (radioButton2) {
                await radioButton2.click(); // Hacer clic en el radio button
                console.log("Radio button seleccionado.");
              } else {
                console.log("No se encontró el radio button.");
              }

              //********
              // Pulsar Validar solicitud
              //********
              await page.locator('input[id="validarSolicitud"]').click();
              await page.waitForNavigation({ waitUntil: "load" });

              //********
              // Pulsar opcion encontrada
              //********
              await page.locator('input[value="Firmar Enviar"]').wait();

              let nuevaPagina;
              try {
                [nuevaPagina] = await Promise.all([
                  new Promise((resolvePromise) => {
                    setTimeout(() => {
                      resolvePromise(false);
                    }, 10000);

                    browser.once("targetcreated", async (target) => {
                      const newPage = await target.page();

                      await newPage.locator('input[id="Conforme"]').wait();
                      await newPage.locator('input[id="Conforme"]').click();
                      await this.esperar(500);

                      await newPage.locator('input[name="Firmar"]').wait();
                      await newPage.locator('input[name="Firmar"]').click();

                      await newPage.close();
                      resolvePromise(true);
                    });
                  }),
                  await page.locator('input[value="Firmar Enviar"]').click(),
                ]);
              } catch (e) {
                console.log("Error en catch: ", e);
              }

              await this.esperar(1000);
              console.log("Descargando...");

              //*************
              // Descargar resguardo solicitud
              //*************
              await page.locator('input[id="descarga"]').wait();
              try {
                [nuevaPagina] = await Promise.all([
                  new Promise((resolvePromise) => {
                    setTimeout(() => {
                      resolvePromise(false);
                    }, 5000);

                    browser.once("targetcreated", async (target) => {
                      const newPage = await target.page();
                      newPage.on("response", async (response) => {
                        // Verificar si el contenido es un PDF
                        if (
                          !response.url().endsWith(".js") &&
                          !response.url().endsWith(".css") &&
                          response.url().startsWith("chrome-extension://")
                        ) {
                          console.log("PDF detectado:", response.url());

                          // Intercepta el PDF:
                          const pdfBuffer = await response.buffer();

                          // Guardar el PDF en el sistema de archivos
                          const filePath = path.join(
                            pathSalida,
                            clientes[i]["nombreArchivo"],
                          );
                          const filePathFacturacion = path.join(
                            pathSalidaFacturacion,
                            clientes[i]["nombreArchivoFacturacion"],
                          );
                          fs.writeFileSync(filePath, pdfBuffer);
                          fs.writeFileSync(filePathFacturacion, pdfBuffer);
                          console.log("PDF descargado en:", filePath);
                          resolvePromise(newPage);
                        }
                      });
                    });
                  }),
                  await page.locator('input[id="descarga"]').click(),
                ]);
              } catch (e) {
                console.log("Error en catch");
              }

              //Comprueba si hubo error
              if (!nuevaPagina) {
                console.log("ERROR EN FIRMA DE CONSENTIMIENTO");
                archivoTributario
                  .sheet("DATOS")
                  .cell(i + 2, 8)
                  .value(
                    "ERROR: No se ha podido generar el resguardo de la solicitud.",
                  );
              } else {
                archivoTributario
                  .sheet("DATOS")
                  .cell(i + 2, 8)
                  .value("OK, resguardo de solicitud descargado.");
                await nuevaPagina.close();
              }
              console.log("Nuevo cliente");
              await this.esperar(1000);
            }

            await browser.close();

            console.log("Escribiendo archivo...");
            console.log("Path: " + path.normalize(pathSalidaExcel));
            archivoTributario
              .toFileAsync(
                path.normalize(
                  path.join(pathSalidaExcel, "Tributario-Procesado.xlsx"),
                ),
              )
              .then(() => {
                console.log("XLSX escrito correctamente");
              })
              .catch((err) => {
                console.log("Se ha producido un error interno: ");
                console.log(err);
                var tituloError =
                  "Se ha producido un error escribiendo el archivo: " +
                  path.normalize(pathSalidaExcel);
                resolve(false);
              });

            resolve(true);
          })
          .then(() => {})
          .catch((err) => {
            console.log("ERROR");

            throw err;
          });
      } catch (err) {
        var tituloError = "No se ha podido cargar el archivo";
        var mensajeError =
          "Se ha producido un error interno cargando los archivos.";
        mainProcess.mostrarError(tituloError, mensajeError).then((result) => {
          resolve(false);
        });
      }
    }).catch((err) => {
      console.log("Se ha producido un error interno: ");
      console.log(err);
      var tituloError = "No se ha podido cargar el archivo";
      var mensajeError =
        "Se ha producido un error interno cargando los archivos.";
      mainProcess.mostrarError(tituloError, mensajeError).then((result) => {
        resolve(false);
      });
    });
  }

  async certificadoSubvencionesATC(argumentos) {
    return new Promise((resolve) => {
      console.log("Archivo ATC CCC...");
      console.log(argumentos.formularioControl[1]);
      console.log("Ruta Google...");
      console.log(argumentos.formularioControl[0]);

      var archivoSubvencionesATC = {};
      var clientes = [];
      var pathArchivoEtiquetas = argumentos.formularioControl[1];
      var pathSalidaExcel = path.join(
        path.normalize(argumentos.formularioControl[2]),
        "Certificados_SubvencionesATC-Procesados",
      );
      var pathSalida = path.join(
        path.normalize(argumentos.formularioControl[2]),
        "Certificados_SubvencionesATC-Procesados",
        "Resultados",
      );
      var pathSalidaFacturacion = path.join(
        path.normalize(argumentos.formularioControl[2]),
        "Certificados_SubvencionesATC-Procesados",
        "FACTURACIÓN",
      );

      // Verificar si la carpeta "Resultados" existe y crearla si no
      if (!fs.existsSync(pathSalida)) {
        fs.mkdirSync(pathSalida, { recursive: true });
        console.log(`Carpeta creada: ${pathSalida}`);
      } else {
        console.log(`La carpeta ya existe: ${pathSalida}`);
      }

      // Verificar si la carpeta "Resultados" existe y crearla si no
      if (!fs.existsSync(pathSalidaFacturacion)) {
        fs.mkdirSync(pathSalidaFacturacion, { recursive: true });
        console.log(`Carpeta creada: ${pathSalidaFacturacion}`);
      } else {
        console.log(`La carpeta ya existe: ${pathSalidaFacturacion}`);
      }

      try {
        XlsxPopulate.fromFileAsync(path.normalize(pathArchivoEtiquetas))
          .then(async (workbook) => {
            console.log("Archivo Cargado: CCC");
            archivoSubvencionesATC = workbook;
            var columnas = archivoSubvencionesATC
              .sheet("DATOS")
              .usedRange()._numColumns;
            var filas = archivoSubvencionesATC
              .sheet("DATOS")
              .usedRange()._numRows;
            var objetoCliente = {};

            var cabeceras = [];
            for (var i = 1; i <= columnas; i++) {
              cabeceras.push(
                archivoSubvencionesATC.sheet("DATOS").cell(1, i).value(),
              );
            }

            console.log("Cabeceras: " + cabeceras);
            for (var i = 2; i <= filas; i++) {
              objetoCliente = {};
              for (var j = 1; j <= columnas; j++) {
                if (
                  archivoSubvencionesATC.sheet("DATOS").cell(i, j).value() !==
                  undefined
                ) {
                  switch (cabeceras[j - 1]) {
                    case "CCC COMPLETO":
                      objetoCliente["ccc"] = archivoSubvencionesATC
                        .sheet("DATOS")
                        .cell(i, j)
                        .value();
                      break;

                    case "EMPRESA":
                      objetoCliente["empresa"] = archivoSubvencionesATC
                        .sheet("DATOS")
                        .cell(i, j)
                        .value();
                      break;

                    case "CÓDIGO":
                      objetoCliente["codigo"] = archivoSubvencionesATC
                        .sheet("DATOS")
                        .cell(i, j)
                        .value();
                      break;

                    case "NIF":
                      objetoCliente["nif"] = archivoSubvencionesATC
                        .sheet("DATOS")
                        .cell(i, j)
                        .value();
                      break;
                  }
                }
              }

              objetoCliente["errores"] = [];
              objetoCliente["flagEvitarDuplicado"] = false;

              if (
                objetoCliente.ccc !== "" &&
                objetoCliente.ccc !== null &&
                objetoCliente.ccc !== undefined
              ) {
                objetoCliente["nombreArchivo"] =
                  objetoCliente["codigo"] +
                  " CERTIFICADO ESTAR AL CORRIENTE AEAT " +
                  objetoCliente["empresa"] +
                  " " +
                  DateTime.now().setZone("Europe/Madrid").toFormat("ddMMyy") +
                  ".pdf";

                objetoCliente["nombreArchivoFacturacion"] =
                  objetoCliente["codigo"] +
                  "-" +
                  objetoCliente["empresa"] +
                  "-3.096-Certificado de estar al corriente AEAT-28.50€-CC.pdf";
                clientes.push(Object.assign({}, objetoCliente));
              }
            }

            //Procesar duplicados:
            const vistos = new Set();
            clientes = clientes.map((obj) => {
              if (vistos.has(obj.nif)) {
                obj["errores"] = [
                  "Evitando generar certificado por NIF duplicado",
                ];
                return { ...obj, flagEvitarDuplicado: true };
              } else {
                vistos.add(obj.nif);
                return obj;
              }
            });

            console.log("Clientes: ");
            console.log(clientes);

            var chromiumExecutablePath = path.normalize(
              argumentos.formularioControl[0],
            );

            //Inicio de procesamiento:
            const browser = await puppeteer.launch({
              executablePath: chromiumExecutablePath,
              headless: false,
            });
            console.log(browser.executablePath);

            var page = await browser.newPage();

            //Confirma el cambio de pagina:
            page.on("dialog", async (dialog) => {
              console.log("Dialogo: ", dialog.type());
              const tipo = dialog.type();
              if (tipo == "beforeunload") {
                await dialog.accept();
              }
            });

            // Configurar el comportamiento de descarga
            await page._client().send("Page.setDownloadBehavior", {
              behavior: "allow",
              downloadPath: pathSalida,
            });

            await page.setViewport({ width: 1080, height: 1024 });

            //Iniciando descarga de informes:
            for (var i = 0; i < clientes.length; i++) {
              if (clientes[i].flagEvitarDuplicado) {
                archivoSubvencionesATC
                  .sheet("DATOS")
                  .cell(i + 2, 8)
                  .value("WARNING: Solicitud evitada por duplicidad en NIF.");
                continue;
              }

              //Recargar cada 10 clientes:
              if (i % 10 == 0 && i > 0) {
                //await browser.close();
                await page.close();
                page = await browser.newPage();

                // Configurar el comportamiento de descarga
                await page._client().send("Page.setDownloadBehavior", {
                  behavior: "allow",
                  downloadPath: pathSalida,
                });
                await page.setViewport({ width: 1080, height: 1024 });
              }

              console.log("Procesando cliente: " + i);
              console.log(clientes[i]);

              if (
                clientes[i].ccc == "" ||
                clientes[i].ccc == null ||
                clientes[i].ccc == undefined
              ) {
                clientes[i]["errores"] = ["Campo CCC no definidos."];
                continue;
              }

              await page.goto(
                "https://sede.gobiernodecanarias.org/tributos/ov/seguro/certificados/individual/listado.jsp",
                { waitUntil: "networkidle0" },
              );

              await this.esperar(1000);

              let flagValidacionRequerida = true;
              try {
                const botonEntrar = await page.waitForSelector(
                  'input[id="btnValidar"]',
                  { timeout: 1000 },
                );
                if (botonEntrar) {
                  await botonEntrar.click();
                  flagValidacionRequerida = true;
                  console.log("Botón clicado");
                }
              } catch (error) {
                console.log("No se requiere validacion");
                flagValidacionRequerida = false;
              }

              console.log("Flag validacion: " + flagValidacionRequerida);

              //********
              // Esperar a boton actualizar
              //********
              try {
                const botonSolicitar = await page.waitForSelector(
                  'input[id="btnSolicitar"]',
                  { timeout: 60000 },
                );
                if (botonSolicitar) {
                  await botonSolicitar.click();
                  console.log("Botón clicado");
                }
              } catch (error) {
                console.log(
                  "Botón no encontrado después de 60 segundo, continuando...",
                );
                await browser.close();
                resolve(false);
              }

              await page.locator(`select[name="tiposCertificado"]`).wait();
              await this.esperar(500);
              await page.select('select[name="tiposCertificado"]', "AS");

              await page.locator(`input[id="id_tipo_terceros"]`).wait();
              var radioButton = await page.$(`input[id="id_tipo_terceros"]`);

              if (radioButton) {
                await radioButton.click(); // Hacer clic en el radio button
                console.log("Radio button seleccionado.");
              } else {
                console.log("No se encontró el radio button.");
              }

              await this.esperar(1000);

              //********
              // Introducir NIE
              //********
              await page.locator('input[id="idNifTitular"]').wait();
              await page.type(
                'input[id="idNifTitular"]',
                String(clientes[i].nif),
              );
              await this.esperar(500);

              //********
              // Introducir Nombre y apellidos
              //********
              await page.locator('input[id="idNombreTitular"]').wait();
              await page.type(
                'input[id="idNombreTitular"]',
                String(clientes[i].empresa),
              );
              await this.esperar(500);

              //********
              // Pulsar SOLICITAR
              //********
              await page.locator('input[id="btnSolicitar"]').wait();
              await page.locator('input[id="btnSolicitar"]').click();

              if (
                (await page.evaluate(() => document.readyState)) != "complete"
              ) {
                await page.waitForNavigation({ waitUntil: "load" });
              }

              console.log("Solicitud realizada");

              //********
              // Pulsar Descargar
              //********
              if (
                (await page.evaluate(() => document.readyState)) != "complete"
              ) {
                await page.waitForNavigation({ waitUntil: "load" });
              }

              console.log("Descargando...");
              //*************
              // Descargar resguardo solicitud
              //*************
              try {
                const botonDescargar = await page.waitForSelector(
                  'input[id="btnDescargar"]',
                  { timeout: 40000 },
                );
              } catch (error) {
                clientes[i]["errores"] = [
                  "ERROR: No se ha podido generar la solicitud.",
                ];
                archivoSubvencionesATC
                  .sheet("DATOS")
                  .cell(i + 2, 8)
                  .value("ERROR: No se ha podido generar la solicitud.");
                console.log(
                  "Boton de descarga no encontrado después de 40 segundo, continuando...",
                );
                console.log("Nuevo cliente");
                await this.esperar(1000);
                continue;
              }

              await this.esperar(1000);

              let nuevaPagina;
              try {
                [nuevaPagina] = await Promise.all([
                  new Promise((resolvePromise) => {
                    setTimeout(() => {
                      resolvePromise(false);
                    }, 5000);

                    browser.once("targetcreated", async (target) => {
                      const newPage = await target.page();
                      newPage.on("response", async (response) => {
                        // Verificar si el contenido es un PDF
                        if (
                          !response.url().endsWith(".js") &&
                          !response.url().endsWith(".css") &&
                          response.url().startsWith("chrome-extension://")
                        ) {
                          console.log("PDF detectado:", response.url());
                          // Intercepta el PDF:
                          const pdfBuffer = await response.buffer();

                          // Guardar el PDF en el sistema de archivos
                          const filePath = path.join(
                            pathSalida,
                            clientes[i]["nombreArchivo"],
                          );
                          const filePathFacturacion = path.join(
                            pathSalidaFacturacion,
                            clientes[i]["nombreArchivoFacturacion"],
                          );
                          fs.writeFileSync(filePath, pdfBuffer);
                          fs.writeFileSync(filePathFacturacion, pdfBuffer);
                          console.log("PDF descargado en:", filePath);
                          resolvePromise(newPage);
                        }
                      });
                    });
                  }),
                  await page.locator('input[id="btnDescargar"]').click(),
                ]);
              } catch (e) {
                console.log("Error en catch");
              }

              //Comprueba si hubo error
              if (!nuevaPagina) {
                console.log("ERROR ABRIENDO DESCARGA");
                archivoSubvencionesATC
                  .sheet("DATOS")
                  .cell(i + 2, 8)
                  .value(
                    "ERROR: No se ha podido generar el resguardo de la solicitud.",
                  );
              } else {
                archivoSubvencionesATC
                  .sheet("DATOS")
                  .cell(i + 2, 8)
                  .value("OK, resguardo de solicitud descargado.");
                await nuevaPagina.close();
              }
              console.log("Nuevo cliente");
              await this.esperar(1000);
            }

            await browser.close();

            console.log("Escribiendo archivo...");
            console.log("Path: " + path.normalize(pathSalidaExcel));
            archivoSubvencionesATC
              .toFileAsync(
                path.normalize(
                  path.join(pathSalidaExcel, "SubvencionesATC-Procesado.xlsx"),
                ),
              )
              .then(() => {
                console.log("XLSX escrito correctamente");
              })
              .catch((err) => {
                console.log("Se ha producido un error interno: ");
                console.log(err);
                var tituloError =
                  "Se ha producido un error escribiendo el archivo: " +
                  path.normalize(pathSalidaExcel);
                resolve(false);
              });

            resolve(true);
          })
          .then(() => {})
          .catch((err) => {
            console.log("ERROR");

            throw err;
          });
      } catch (err) {
        var tituloError = "No se ha podido cargar el archivo";
        var mensajeError =
          "Se ha producido un error interno cargando los archivos.";
        mainProcess.mostrarError(tituloError, mensajeError).then((result) => {
          resolve(false);
        });
      }
    }).catch((err) => {
      console.log("Se ha producido un error interno: ");
      console.log(err);
      var tituloError = "No se ha podido cargar el archivo";
      var mensajeError =
        "Se ha producido un error interno cargando los archivos.";
      mainProcess.mostrarError(tituloError, mensajeError).then((result) => {
        resolve(false);
      });
    });
  }

  async spoolToXLSX(argumentos) {
    console.log("Formatear SPOOL");
    console.log("Archivo entrada: " + argumentos[0]);
    console.log("Archivo salida: " + argumentos[1]);

    const pathSpoolInput = path.join(argumentos[0]);
    var pathSpoolOutput;

    if (
      argumentos[2].slice(-4) !== ".txt" &&
      argumentos[2].slice(-4) !== ".TXT"
    ) {
      pathSpoolOutput = path.join(argumentos[1], argumentos[2] + ".txt");
    } else {
      pathSpoolOutput = path.join(argumentos[1], argumentos[2]);
    }

    const readline = require("readline");
    const outputFile = fs.createWriteStream(pathSpoolOutput);

    async function leerSpool() {
      return new Promise((resolve) => {
        const rl = readline.createInterface({
          input: fs.createReadStream(pathSpoolInput),
        });

        // Handle any error that occurs on the write stream
        outputFile.on("err", (err) => {
          // handle error
          console.log(err);
        });

        // Once done writing, rename the output to be the input file name
        outputFile.on("close", () => {
          console.log("done writing");

          /*fs.rename(pathSpoolOutput, pathSpoolInput, err => {
					if (err) {
					  // handle error
					  console.log(err)
					} else {
					  console.log('renamed file')
					}
				})*/
        });

        // Read the file and replace any text that matches
        rl.on("line", (line) => {
          let text = line;

          // Elimina las lineas que no comienzan por tabulador:
          if (!text.startsWith("\t")) {
            return;
          }

          // Elimina las lineas que comienzan por "Md.":
          if (text.startsWith("\tMd.\t")) {
            return;
          }

          // write text to the output file stream with new line character
          outputFile.write(`${text}\n`);
        });

        // Done reading the input, call end() on the write stream
        rl.on("close", () => {
          console.log("FIN DEL PROCESAMIENTO");
          outputFile.end();
          resolve(true);
        });
      });
    }
    var result = await leerSpool();
    return result;
  }

  //********************************
  //  Procesar Report AM
  //********************************

  async generarSeguimientoAM(argumentos) {
    console.log("EJECUTANDO PROCESADO AM");

    var datosNacho = argumentos[1][0];
    var pathArchivoSeguimiento = argumentos[0];

    var archivoSeguimiento = {};

    var configuracion = {
      añoInicioControl: argumentos[2],
      mesInicioControl: argumentos[3],
      añoFinControl: argumentos[4],
      mesFinControl: argumentos[5],
      datosSalidaControl: argumentos[6],
      nombreArchivoSalidaControl: argumentos[7],
    };

    async function generarReportAM(
      archivoSeguimiento,
      datosNacho,
      configuracion,
    ) {
      return new Promise((resolve) => {
        //PROCESAMIENTO:
        XlsxPopulate.fromFileAsync(path.normalize(pathArchivoSeguimiento))
          .then((workbook) => {
            console.log("Archivo Cargado: Seguimiento");
            archivoSeguimiento = workbook;
            return true;
          })
          .then(() => {
            //PASOS DE PROCESAMIENTO:

            // 1) Eliminar Foto semana anterior
            // 2) Mover Foto a semana anterior
            // 3) Pegar Columnas de nacho en foto de seguiemiento
            // 4) Ejecutar formulas campos (Fecha Creación, Week Creación, Fecha Cierre, Week Cierre, Proceso, Subproceso, Últimos 15 días)
            // 5) Procesar recuento General (Entrada,Salida,Cerradas,Canceladas,Backlog,Resolución Neta(E - S));
            // 6) Procesar recuento RMCA (Entrada,Salida,Cerradas,Canceladas,Backlog,Resolución Neta(E - S));
            // 7) Procesar recuento SAP 47 (Entrada,Salida,Cerradas,Canceladas,Backlog,Resolución Neta(E - S));
            // 8) Procesar recuento EDITRAN (Entrada,Salida,Cerradas,Canceladas,Backlog,Resolución Neta(E - S));
            // 9) Procesar recuento Connect Direct (Entrada,Salida,Cerradas,Canceladas,Backlog,Resolución Neta(E - S));

            //Ejecución:

            //  1) ELIMINAR FOTO SEMANA ANTERIOR:
            console.log(archivoSeguimiento.sheet("Foto"));
            console.log(archivoSeguimiento.sheet("Foto").usedRange());

            var columnasFotoAnteriorSeguimiento = archivoSeguimiento
              .sheet("Foto_semana anterior")
              .usedRange()._numColumns;
            var filasFotoAnteriorSeguimiento = archivoSeguimiento
              .sheet("Foto_semana anterior")
              .usedRange()._numRows;

            //Limpia hoja de seguimiento anterior:
            for (var i = 1; i < filasFotoAnteriorSeguimiento; i++) {
              for (var j = 0; j < columnasFotoAnteriorSeguimiento; j++) {
                archivoSeguimiento
                  .sheet("Foto_semana anterior")
                  .row(i + 1)
                  .cell(j + 1)
                  .clear();
              }
            }

            // 2) Mover Foto a semana anterior
            archivoSeguimiento
              .sheet("Foto_semana anterior")
              .name("Provisional");
            archivoSeguimiento.sheet("Foto").name("Foto_semana anterior");
            archivoSeguimiento.sheet("Provisional").name("Foto");
            archivoSeguimiento.moveSheet("Foto", "Foto_semana anterior");

            // 3) Pegar Columnas de nacho en foto de seguimiento:
            console.log("DATOS NACHO: ");
            console.log(datosNacho);

            if (datosNacho == null || datosNacho == undefined) {
              console.log("Se ha producido un error interno: ");
              console.log(err);
              var tituloError =
                "Compruebe que se ha cargado el archivo 'Nacho' en el gestor de datos.";
              mainWindow.webContents.send("onErrorInterno", tituloError, err);
              resolve(false);
            }

            var cabeceraSeleccionada = "";

            for (var i = 1; i < datosNacho.data.length; i++) {
              for (var j = 0; j < columnasFotoAnteriorSeguimiento; j++) {
                cabeceraSeleccionada = String(
                  archivoSeguimiento
                    .sheet("Foto")
                    .row(1)
                    .cell(j + 1)
                    .value(),
                );
                cabeceraSeleccionada = cabeceraSeleccionada.toLowerCase();
                cabeceraSeleccionada = cabeceraSeleccionada.replace(/ /g, "_");

                if (cabeceraSeleccionada == "hub") {
                  cabeceraSeleccionada = "categoria_3";
                }

                //console.log(cabeceraSeleccionada)
                if (cabeceraSeleccionada !== undefined) {
                  if (
                    datosNacho.data[i - 1][cabeceraSeleccionada] !== undefined
                  ) {
                    archivoSeguimiento
                      .sheet("Foto")
                      .row(i + 1)
                      .cell(j + 1)
                      .value(datosNacho.data[i - 1][cabeceraSeleccionada]);
                  } else {
                    console.log(
                      "Warning: Dato no encontrado i=" +
                        i +
                        " j=" +
                        j +
                        " Cabecera: " +
                        cabeceraSeleccionada,
                    );
                  }
                } else {
                  console.log("Warning de cabecera: i=" + i + " j=" + j);
                }
              }
            }

            // 4) Ejecutar formulas campos (Fecha Creación, Week Creación, Fecha Cierre, Week Cierre, Proceso, Subproceso, Últimos 15 días):

            var datoProcesado;
            var indiceBusqueda = 0;
            var indiceBusquedaAnterior = 0;
            var valorEncontrado = false;

            columnasFotoAnteriorSeguimiento = archivoSeguimiento
              .sheet("Foto_semana anterior")
              .usedRange()._numColumns;
            filasFotoAnteriorSeguimiento = archivoSeguimiento
              .sheet("Foto_semana anterior")
              .usedRange()._numRows;

            for (var i = 1; i < datosNacho.data.length; i++) {
              for (var j = 0; j < columnasFotoAnteriorSeguimiento; j++) {
                cabeceraSeleccionada = String(
                  archivoSeguimiento
                    .sheet("Foto")
                    .row(1)
                    .cell(j + 1)
                    .value(),
                );
                cabeceraSeleccionada = cabeceraSeleccionada.toLowerCase();
                cabeceraSeleccionada = cabeceraSeleccionada.replace(/ /g, "_");

                //console.log(cabeceraSeleccionada)
                if (cabeceraSeleccionada === undefined) {
                  console.log("Error de cabecera: i=" + i + " j=" + j);
                } else {
                  switchProcesado: switch (cabeceraSeleccionada) {
                    case "fecha_creacion":
                      datoProcesado = archivoSeguimiento
                        .sheet("Foto")
                        .row(i + 1)
                        .cell(9)
                        .value();
                      if (datoProcesado.indexOf("/") != -1) {
                        datoProcesado = moment(datoProcesado, "DD/MM/YYYY");
                      } else {
                        datoProcesado = moment(datoProcesado, "YYYY-MM-DD");
                      }

                      datoProcesado =
                        datoProcesado.date() +
                        "/" +
                        (datoProcesado.month() + 1) +
                        "/" +
                        datoProcesado.year();

                      archivoSeguimiento
                        .sheet("Foto")
                        .row(i + 1)
                        .cell(j + 1)
                        .value(String(datoProcesado));
                      //console.log(datoProcesado)
                      break;

                    case "week_creacion":
                      datoProcesado = archivoSeguimiento
                        .sheet("Foto")
                        .row(i + 1)
                        .cell(9)
                        .value();
                      //datoProcesado = datoProcesado.replace(/\//g,"-");
                      if (datoProcesado.indexOf("/") != -1) {
                        datoProcesado = moment(datoProcesado, "DD/MM/YYYY");
                      } else {
                        datoProcesado = moment(datoProcesado, "YYYY-MM-DD");
                      }

                      datoProcesado = datoProcesado.week();

                      archivoSeguimiento
                        .sheet("Foto")
                        .row(i + 1)
                        .cell(j + 1)
                        .value(String(datoProcesado));
                      //console.log(datoProcesado)
                      break;

                    case "fecha_cierre":
                      datoProcesado = archivoSeguimiento
                        .sheet("Foto")
                        .row(i + 1)
                        .cell(11)
                        .value();
                      if (
                        datoProcesado === undefined ||
                        datoProcesado == "" ||
                        datoProcesado._error == "#N/A"
                      ) {
                        archivoSeguimiento
                          .sheet("Foto")
                          .row(i + 1)
                          .cell(j + 1)
                          .clear();
                        break;
                      }
                      //datoProcesado = datoProcesado.replace(/\//g,"-");
                      if (datoProcesado.indexOf("/") != -1) {
                        datoProcesado = moment(datoProcesado, "DD/MM/YYYY");
                      } else {
                        datoProcesado = moment(datoProcesado, "YYYY-MM-DD");
                      }

                      datoProcesado =
                        datoProcesado.date() +
                        "/" +
                        (datoProcesado.month() + 1) +
                        "/" +
                        datoProcesado.year();

                      archivoSeguimiento
                        .sheet("Foto")
                        .row(i + 1)
                        .cell(j + 1)
                        .value(String(datoProcesado));
                      //console.log(datoProcesado)
                      break;

                    case "week_cierre":
                      datoProcesado = archivoSeguimiento
                        .sheet("Foto")
                        .row(i + 1)
                        .cell(11)
                        .value();
                      if (
                        datoProcesado === undefined ||
                        datoProcesado == "" ||
                        datoProcesado._error == "#N/A"
                      ) {
                        archivoSeguimiento
                          .sheet("Foto")
                          .row(i + 1)
                          .cell(j + 1)
                          .clear();
                        break;
                      }
                      //datoProcesado = datoProcesado.replace(/\//g,"-");
                      if (datoProcesado.indexOf("/") != -1) {
                        datoProcesado = moment(datoProcesado, "DD/MM/YYYY");
                      } else {
                        datoProcesado = moment(datoProcesado, "YYYY-MM-DD");
                      }

                      datoProcesado = datoProcesado.week();

                      archivoSeguimiento
                        .sheet("Foto")
                        .row(i + 1)
                        .cell(j + 1)
                        .value(String(datoProcesado));
                      //console.log(datoProcesado)
                      break;

                    case "proceso":
                      datoProcesado = archivoSeguimiento
                        .sheet("Foto")
                        .row(i + 1)
                        .cell(2)
                        .value();

                      if (
                        datoProcesado === undefined ||
                        datoProcesado == "" ||
                        datoProcesado._error == "#N/A"
                      ) {
                        archivoSeguimiento
                          .sheet("Foto")
                          .row(i + 1)
                          .cell(j + 1)
                          .clear();
                        break switchProcesado;
                      }

                      indiceBusqueda = indiceBusquedaAnterior;
                      valorEncontrado = false;
                      for (var k = 1; k < filasFotoAnteriorSeguimiento; k++) {
                        if (indiceBusqueda > filasFotoAnteriorSeguimiento) {
                          indiceBusqueda =
                            indiceBusqueda - filasFotoAnteriorSeguimiento;
                        }

                        if (
                          datoProcesado ===
                          archivoSeguimiento
                            .sheet("Foto_semana anterior")
                            .row(indiceBusqueda + 1)
                            .cell(2)
                            .value()
                        ) {
                          datoProcesado = archivoSeguimiento
                            .sheet("Foto_semana anterior")
                            .row(indiceBusqueda + 1)
                            .cell(31)
                            .value();
                          indiceBusquedaAnterior = indiceBusqueda;
                          valorEncontrado = true;
                          break;
                        }
                        indiceBusqueda++;
                      }

                      if (
                        datoProcesado === undefined ||
                        datoProcesado == "" ||
                        datoProcesado._error == "#N/A" ||
                        !valorEncontrado
                      ) {
                        archivoSeguimiento
                          .sheet("Foto")
                          .row(i + 1)
                          .cell(j + 1)
                          .clear();
                        break switchProcesado;
                      }

                      archivoSeguimiento
                        .sheet("Foto")
                        .row(i + 1)
                        .cell(j + 1)
                        .value(String(datoProcesado));
                      //console.log(JSON.stringify(datoProcesado))

                      break switchProcesado;

                    case "subproceso":
                      datoProcesado = archivoSeguimiento
                        .sheet("Foto")
                        .row(i + 1)
                        .cell(2)
                        .value();

                      if (
                        datoProcesado === undefined ||
                        datoProcesado == "" ||
                        datoProcesado._error == "#N/A"
                      ) {
                        archivoSeguimiento
                          .sheet("Foto")
                          .row(i + 1)
                          .cell(j + 1)
                          .clear();
                        break switchProcesado;
                      }

                      indiceBusqueda = indiceBusquedaAnterior;
                      valorEncontrado = false;
                      for (var k = 1; k < filasFotoAnteriorSeguimiento; k++) {
                        if (indiceBusqueda > filasFotoAnteriorSeguimiento) {
                          indiceBusqueda =
                            indiceBusqueda - filasFotoAnteriorSeguimiento;
                        }

                        if (
                          datoProcesado ===
                          archivoSeguimiento
                            .sheet("Foto_semana anterior")
                            .row(indiceBusqueda + 1)
                            .cell(2)
                            .value()
                        ) {
                          datoProcesado = archivoSeguimiento
                            .sheet("Foto_semana anterior")
                            .row(indiceBusqueda + 1)
                            .cell(32)
                            .value();
                          indiceBusquedaAnterior = indiceBusqueda;
                          valorEncontrado = true;
                          break;
                        }
                        indiceBusqueda++;
                      }

                      if (
                        datoProcesado === undefined ||
                        datoProcesado == "" ||
                        datoProcesado._error == "#N/A" ||
                        !valorEncontrado
                      ) {
                        archivoSeguimiento
                          .sheet("Foto")
                          .row(i + 1)
                          .cell(j + 1)
                          .clear();
                        break switchProcesado;
                      }

                      archivoSeguimiento
                        .sheet("Foto")
                        .row(i + 1)
                        .cell(j + 1)
                        .value(String(datoProcesado));
                      //console.log(JSON.stringify(datoProcesado))

                      break switchProcesado;

                    case "ultimos_15_días":
                      datoProcesado = archivoSeguimiento
                        .sheet("Foto")
                        .row(i + 1)
                        .cell(9)
                        .value();

                      if (datoProcesado.indexOf("/") != -1) {
                        datoProcesado = moment(datoProcesado, "DD/MM/YYYY");
                      } else {
                        datoProcesado = moment(datoProcesado, "YYYY-MM-DD");
                      }

                      if (datoProcesado >= moment().subtract(15, "days")) {
                        archivoSeguimiento
                          .sheet("Foto")
                          .row(i + 1)
                          .cell(j + 1)
                          .value("X");
                      } else {
                        archivoSeguimiento
                          .sheet("Foto")
                          .row(i + 1)
                          .cell(j + 1)
                          .clear();
                      }
                      //console.log(datoProcesado)
                      break;
                  }
                }
              }
            }

            // 5) Procesar recuento General (Entrada,Salida,Cerradas,Canceladas,Backlog,Resolución Neta(E - S));

            var filtroServicio = [];
            var filtroTipoIncidente = [];
            var filtroEstado = [];
            var filtroEnOtraCola = [];
            var filtroMes = [];

            var registroResultados = [];
            const resultadosBase = {
              general: {
                entrada: 0,
                salida: 0,
                cerradas: 0,
                canceladas: 0,
                backlog: 0,
                neto: 0,
              },
              rmca: {
                entrada: 0,
                salida: 0,
                cerradas: 0,
                canceladas: 0,
                backlog: 0,
                neto: 0,
              },
              sap: {
                entrada: 0,
                salida: 0,
                cerradas: 0,
                canceladas: 0,
                backlog: 0,
                neto: 0,
              },
              editran: {
                entrada: 0,
                salida: 0,
                cerradas: 0,
                canceladas: 0,
                backlog: 0,
                neto: 0,
              },
              connectDirect: {
                entrada: 0,
                salida: 0,
                cerradas: 0,
                canceladas: 0,
                backlog: 0,
                neto: 0,
              },
            };

            var añoInicio = configuracion.añoInicioControl;
            var mesInicio = configuracion.mesInicioControl - 1;

            var añoFin = configuracion.añoFinControl;
            var mesFin = configuracion.mesFinControl - 1;

            var mesActual = moment().month();
            var añoActual = moment().year();

            console.log("Año actual: " + añoActual);
            console.log("Mes actual: " + mesActual);

            //Verificación de año:
            if (añoInicio > añoFin) {
              añoInicio = añoFin;
            }

            var filasFotoSeguimiento = archivoSeguimiento
              .sheet("Foto")
              .usedRange()._numRows;

            // 5) Procesar recuento General (Entrada,Salida,Cerradas,Canceladas,Backlog,Resolución Neta(E - S));

            //Iteracion por Años
            for (var i = añoInicio; i <= añoFin; i++) {
              var mesStart = 0;
              var mesFin = 11;
              //Iteracion por Meses
              if (i == añoInicio) {
                mesStart = mesInicio;
              }
              if (i == añoFin) {
                mesFin = mesFin;
              }
              for (var j = mesStart; j <= mesFin; j++) {
                registroResultados.push({
                  año: i,
                  mes: j,
                  datos: _.cloneDeep(resultadosBase),
                });

                console.log(JSON.stringify(registroResultados));
                //ITERAR SISTEMA:
                for (var sistema in registroResultados[
                  registroResultados.length - 1
                ].datos) {
                  switch (sistema) {
                    case "general":
                      for (var estado in registroResultados[
                        registroResultados.length - 1
                      ].datos[sistema]) {
                        switch (estado) {
                          case "entrada":
                            filtroServicio = [
                              "VFES-SAP 4.7 SGCYR-PROD",
                              "VFES-SAP 4.7-PROD",
                              "VFES-SEPA CONNECT DIRECT-PROD",
                              "VFES-SAP 4.7 CONNECT DIRECT-PROD",
                              "ES-CONNECT DIRECT",
                              "VFES-EDITRAN-PROD",
                              "VFES-EDITRAN",
                              "VFES-ONO-EDITRAN BANKS-PROD",
                              "VFES-EDITRAN. BANKS-PROD",
                              "VFES-RMCA-PROD",
                              "VFES-RMCA-INFRASTRUCTURE-PROD",
                            ];
                            filtroTipoIncidente = [
                              "Incident",
                              "User Service Restoration",
                            ];
                            filtroEstado = [
                              "Assigned",
                              "Cancelled",
                              "Closed",
                              "In Progress",
                              "Pending",
                              "Resolved",
                            ];
                            filtroEnOtraCola = [undefined];
                            break;
                          case "salida":
                            filtroServicio = [
                              "VFES-SAP 4.7 SGCYR-PROD",
                              "VFES-SAP 4.7-PROD",
                              "VFES-SEPA CONNECT DIRECT-PROD",
                              "VFES-SAP 4.7 CONNECT DIRECT-PROD",
                              "ES-CONNECT DIRECT",
                              "VFES-EDITRAN-PROD",
                              "VFES-EDITRAN",
                              "VFES-ONO-EDITRAN BANKS-PROD",
                              "VFES-EDITRAN. BANKS-PROD",
                              "VFES-RMCA-PROD",
                              "VFES-RMCA-INFRASTRUCTURE-PROD",
                            ];
                            filtroTipoIncidente = [
                              "Incident",
                              "User Service Restoration",
                            ];
                            filtroEstado = ["Cancelled", "Closed"];
                            filtroEnOtraCola = [undefined];
                            break;
                          case "cerradas":
                            filtroServicio = [
                              "VFES-SAP 4.7 SGCYR-PROD",
                              "VFES-SAP 4.7-PROD",
                              "VFES-SEPA CONNECT DIRECT-PROD",
                              "VFES-SAP 4.7 CONNECT DIRECT-PROD",
                              "ES-CONNECT DIRECT",
                              "VFES-EDITRAN-PROD",
                              "VFES-EDITRAN",
                              "VFES-ONO-EDITRAN BANKS-PROD",
                              "VFES-EDITRAN. BANKS-PROD",
                              "VFES-RMCA-PROD",
                              "VFES-RMCA-INFRASTRUCTURE-PROD",
                            ];
                            filtroTipoIncidente = [
                              "Incident",
                              "User Service Restoration",
                            ];
                            filtroEstado = ["Closed"];
                            filtroEnOtraCola = [undefined];
                            break;
                          case "canceladas":
                            filtroServicio = [
                              "VFES-SAP 4.7 SGCYR-PROD",
                              "VFES-SAP 4.7-PROD",
                              "VFES-SEPA CONNECT DIRECT-PROD",
                              "VFES-SAP 4.7 CONNECT DIRECT-PROD",
                              "ES-CONNECT DIRECT",
                              "VFES-EDITRAN-PROD",
                              "VFES-EDITRAN",
                              "VFES-ONO-EDITRAN BANKS-PROD",
                              "VFES-EDITRAN. BANKS-PROD",
                              "VFES-RMCA-PROD",
                              "VFES-RMCA-INFRASTRUCTURE-PROD",
                            ];
                            filtroTipoIncidente = [
                              "Incident",
                              "User Service Restoration",
                            ];
                            filtroEstado = ["Cancelled"];
                            filtroEnOtraCola = [undefined];
                            break;
                          case "backlog":
                            filtroServicio = [
                              "VFES-SAP 4.7 SGCYR-PROD",
                              "VFES-SAP 4.7-PROD",
                              "VFES-SEPA CONNECT DIRECT-PROD",
                              "VFES-SAP 4.7 CONNECT DIRECT-PROD",
                              "ES-CONNECT DIRECT",
                              "VFES-EDITRAN-PROD",
                              "VFES-EDITRAN",
                              "VFES-ONO-EDITRAN BANKS-PROD",
                              "VFES-EDITRAN. BANKS-PROD",
                              "VFES-RMCA-PROD",
                              "VFES-RMCA-INFRASTRUCTURE-PROD",
                            ];
                            filtroTipoIncidente = [
                              "Incident",
                              "User Service Restoration",
                            ];
                            filtroEstado = [
                              "Assigned",
                              "In Progress",
                              "Pending",
                              "Resolved",
                            ];
                            filtroEnOtraCola = [undefined];
                            break;
                          case "neto":
                            filtroServicio = [
                              "VFES-SAP 4.7 SGCYR-PROD",
                              "VFES-SAP 4.7-PROD",
                              "VFES-SEPA CONNECT DIRECT-PROD",
                              "VFES-SAP 4.7 CONNECT DIRECT-PROD",
                              "ES-CONNECT DIRECT",
                              "VFES-EDITRAN-PROD",
                              "VFES-EDITRAN",
                              "VFES-ONO-EDITRAN BANKS-PROD",
                              "VFES-EDITRAN. BANKS-PROD",
                              "VFES-RMCA-PROD",
                              "VFES-RMCA-INFRASTRUCTURE-PROD",
                            ];
                            filtroTipoIncidente = [
                              "Incident",
                              "User Service Restoration",
                            ];
                            filtroEstado = [
                              "Assigned",
                              "Cancelled",
                              "Closed",
                              "In Progress",
                              "Pending",
                              "Resolved",
                            ];
                            filtroEnOtraCola = [undefined];
                            break;
                        }
                        registroResultados[registroResultados.length - 1].datos[
                          sistema
                        ][estado] = procesarRecuentoSeguimientoAM(
                          sistema,
                          estado,
                          i,
                          j,
                          filtroServicio,
                          filtroTipoIncidente,
                          filtroEstado,
                          filtroEnOtraCola,
                        );
                      }
                      break;
                    case "rmca":
                      for (const estado in registroResultados[
                        registroResultados.length - 1
                      ].datos[sistema]) {
                        switch (estado) {
                          case "entrada":
                            filtroServicio = [
                              "VFES-RMCA-PROD",
                              "VFES-RMCA-INFRASTRUCTURE-PROD",
                            ];
                            filtroTipoIncidente = [
                              "Incident",
                              "User Service Restoration",
                            ];
                            filtroEstado = [
                              "Assigned",
                              "Cancelled",
                              "Closed",
                              "In Progress",
                              "Pending",
                              "Resolved",
                            ];
                            filtroEnOtraCola = [undefined];
                            break;
                          case "salida":
                            filtroServicio = [
                              "VFES-RMCA-PROD",
                              "VFES-RMCA-INFRASTRUCTURE-PROD",
                            ];
                            filtroTipoIncidente = [
                              "Incident",
                              "User Service Restoration",
                            ];
                            filtroEstado = ["Cancelled", "Closed"];
                            filtroEnOtraCola = [undefined];
                            break;
                          case "cerradas":
                            filtroServicio = [
                              "VFES-RMCA-PROD",
                              "VFES-RMCA-INFRASTRUCTURE-PROD",
                            ];
                            filtroTipoIncidente = [
                              "Incident",
                              "User Service Restoration",
                            ];
                            filtroEstado = ["Closed"];
                            filtroEnOtraCola = [undefined];
                            break;
                          case "canceladas":
                            filtroServicio = [
                              "VFES-RMCA-PROD",
                              "VFES-RMCA-INFRASTRUCTURE-PROD",
                            ];
                            filtroTipoIncidente = [
                              "Incident",
                              "User Service Restoration",
                            ];
                            filtroEstado = ["Cancelled"];
                            filtroEnOtraCola = [undefined];
                            break;
                          case "backlog":
                            filtroServicio = [
                              "VFES-RMCA-PROD",
                              "VFES-RMCA-INFRASTRUCTURE-PROD",
                            ];
                            filtroTipoIncidente = [
                              "Incident",
                              "User Service Restoration",
                            ];
                            filtroEstado = [
                              "Assigned",
                              "In Progress",
                              "Pending",
                              "Resolved",
                            ];
                            filtroEnOtraCola = [undefined];
                            break;
                          case "neto":
                            filtroServicio = [
                              "VFES-RMCA-PROD",
                              "VFES-RMCA-INFRASTRUCTURE-PROD",
                            ];
                            filtroTipoIncidente = [
                              "Incident",
                              "User Service Restoration",
                            ];
                            filtroEstado = [
                              "Assigned",
                              "Cancelled",
                              "Closed",
                              "In Progress",
                              "Pending",
                              "Resolved",
                            ];
                            filtroEnOtraCola = [undefined];
                            break;
                        }
                        registroResultados[registroResultados.length - 1].datos[
                          sistema
                        ][estado] = procesarRecuentoSeguimientoAM(
                          sistema,
                          estado,
                          i,
                          j,
                          filtroServicio,
                          filtroTipoIncidente,
                          filtroEstado,
                          filtroEnOtraCola,
                        );
                      }
                      break;
                    case "sap":
                      for (var estado in registroResultados[
                        registroResultados.length - 1
                      ].datos[sistema]) {
                        switch (estado) {
                          case "entrada":
                            filtroServicio = [
                              "VFES-SAP 4.7 SGCYR-PROD",
                              "VFES-SAP 4.7-PROD",
                            ];
                            filtroTipoIncidente = [
                              "Incident",
                              "User Service Restoration",
                            ];
                            filtroEstado = [
                              "Assigned",
                              "Cancelled",
                              "Closed",
                              "In Progress",
                              "Pending",
                              "Resolved",
                            ];
                            filtroEnOtraCola = [undefined];
                            break;
                          case "salida":
                            filtroServicio = [
                              "VFES-SAP 4.7 SGCYR-PROD",
                              "VFES-SAP 4.7-PROD",
                            ];
                            filtroTipoIncidente = [
                              "Incident",
                              "User Service Restoration",
                            ];
                            filtroEstado = ["Cancelled", "Closed"];
                            filtroEnOtraCola = [undefined];
                            break;
                          case "cerradas":
                            filtroServicio = [
                              "VFES-SAP 4.7 SGCYR-PROD",
                              "VFES-SAP 4.7-PROD",
                            ];
                            filtroTipoIncidente = [
                              "Incident",
                              "User Service Restoration",
                            ];
                            filtroEstado = ["Closed"];
                            filtroEnOtraCola = [undefined];
                            break;
                          case "canceladas":
                            filtroServicio = [
                              "VFES-SAP 4.7 SGCYR-PROD",
                              "VFES-SAP 4.7-PROD",
                            ];
                            filtroTipoIncidente = [
                              "Incident",
                              "User Service Restoration",
                            ];
                            filtroEstado = ["Cancelled"];
                            filtroEnOtraCola = [undefined];
                            break;
                          case "backlog":
                            filtroServicio = [
                              "VFES-SAP 4.7 SGCYR-PROD",
                              "VFES-SAP 4.7-PROD",
                            ];
                            filtroTipoIncidente = [
                              "Incident",
                              "User Service Restoration",
                            ];
                            filtroEstado = [
                              "Assigned",
                              "In Progress",
                              "Pending",
                              "Resolved",
                            ];
                            filtroEnOtraCola = [undefined];
                            break;
                          case "neto":
                            filtroServicio = [
                              "VFES-SAP 4.7 SGCYR-PROD",
                              "VFES-SAP 4.7-PROD",
                            ];
                            filtroTipoIncidente = [
                              "Incident",
                              "User Service Restoration",
                            ];
                            filtroEstado = [
                              "Assigned",
                              "Cancelled",
                              "Closed",
                              "In Progress",
                              "Pending",
                              "Resolved",
                            ];
                            filtroEnOtraCola = [undefined];
                            break;
                        }
                        registroResultados[registroResultados.length - 1].datos[
                          sistema
                        ][estado] = procesarRecuentoSeguimientoAM(
                          sistema,
                          estado,
                          i,
                          j,
                          filtroServicio,
                          filtroTipoIncidente,
                          filtroEstado,
                          filtroEnOtraCola,
                        );
                      }
                      break;
                    case "editran":
                      for (var estado in registroResultados[
                        registroResultados.length - 1
                      ].datos[sistema]) {
                        switch (estado) {
                          case "entrada":
                            filtroServicio = [
                              "VFES-EDITRAN-PROD",
                              "VFES-EDITRAN",
                              "VFES-ONO-EDITRAN BANKS-PROD",
                              "VFES-EDITRAN. BANKS-PROD",
                            ];
                            filtroTipoIncidente = [
                              "Incident",
                              "User Service Restoration",
                            ];
                            filtroEstado = [
                              "Assigned",
                              "Cancelled",
                              "Closed",
                              "In Progress",
                              "Pending",
                              "Resolved",
                            ];
                            filtroEnOtraCola = [undefined];
                            break;
                          case "salida":
                            filtroServicio = [
                              "VFES-EDITRAN-PROD",
                              "VFES-EDITRAN",
                              "VFES-ONO-EDITRAN BANKS-PROD",
                              "VFES-EDITRAN. BANKS-PROD",
                            ];
                            filtroTipoIncidente = [
                              "Incident",
                              "User Service Restoration",
                            ];
                            filtroEstado = ["Cancelled", "Closed"];
                            filtroEnOtraCola = [undefined];
                            break;
                          case "cerradas":
                            filtroServicio = [
                              "VFES-EDITRAN-PROD",
                              "VFES-EDITRAN",
                              "VFES-ONO-EDITRAN BANKS-PROD",
                              "VFES-EDITRAN. BANKS-PROD",
                            ];
                            filtroTipoIncidente = [
                              "Incident",
                              "User Service Restoration",
                            ];
                            filtroEstado = ["Closed"];
                            filtroEnOtraCola = [undefined];
                            break;
                          case "canceladas":
                            filtroServicio = [
                              "VFES-EDITRAN-PROD",
                              "VFES-EDITRAN",
                              "VFES-ONO-EDITRAN BANKS-PROD",
                              "VFES-EDITRAN. BANKS-PROD",
                            ];
                            filtroTipoIncidente = [
                              "Incident",
                              "User Service Restoration",
                            ];
                            filtroEstado = ["Cancelled"];
                            filtroEnOtraCola = [undefined];
                            break;
                          case "backlog":
                            filtroServicio = [
                              "VFES-EDITRAN-PROD",
                              "VFES-EDITRAN",
                              "VFES-ONO-EDITRAN BANKS-PROD",
                              "VFES-EDITRAN. BANKS-PROD",
                            ];
                            filtroTipoIncidente = [
                              "Incident",
                              "User Service Restoration",
                            ];
                            filtroEstado = [
                              "Assigned",
                              "In Progress",
                              "Pending",
                              "Resolved",
                            ];
                            filtroEnOtraCola = [undefined];
                            break;
                          case "neto":
                            filtroServicio = [
                              "VFES-EDITRAN-PROD",
                              "VFES-EDITRAN",
                              "VFES-ONO-EDITRAN BANKS-PROD",
                              "VFES-EDITRAN. BANKS-PROD",
                            ];
                            filtroTipoIncidente = [
                              "Incident",
                              "User Service Restoration",
                            ];
                            filtroEstado = [
                              "Assigned",
                              "Cancelled",
                              "Closed",
                              "In Progress",
                              "Pending",
                              "Resolved",
                            ];
                            filtroEnOtraCola = [undefined];
                            break;
                        }
                        registroResultados[registroResultados.length - 1].datos[
                          sistema
                        ][estado] = procesarRecuentoSeguimientoAM(
                          sistema,
                          estado,
                          i,
                          j,
                          filtroServicio,
                          filtroTipoIncidente,
                          filtroEstado,
                          filtroEnOtraCola,
                        );
                      }
                      break;
                    case "connectDirect":
                      for (const estado in registroResultados[
                        registroResultados.length - 1
                      ].datos[sistema]) {
                        switch (estado) {
                          case "entrada":
                            filtroServicio = [
                              "VFES-SEPA CONNECT DIRECT-PROD",
                              "VFES-SAP 4.7 CONNECT DIRECT-PROD",
                              "ES-CONNECT DIRECT",
                            ];
                            filtroTipoIncidente = [
                              "Incident",
                              "User Service Restoration",
                            ];
                            filtroEstado = [
                              "Assigned",
                              "Cancelled",
                              "Closed",
                              "In Progress",
                              "Pending",
                              "Resolved",
                            ];
                            filtroEnOtraCola = [undefined];
                            break;
                          case "salida":
                            filtroServicio = [
                              "VFES-SEPA CONNECT DIRECT-PROD",
                              "VFES-SAP 4.7 CONNECT DIRECT-PROD",
                              "ES-CONNECT DIRECT",
                            ];
                            filtroTipoIncidente = [
                              "Incident",
                              "User Service Restoration",
                            ];
                            filtroEstado = ["Cancelled", "Closed"];
                            filtroEnOtraCola = [undefined];
                            break;
                          case "cerradas":
                            filtroServicio = [
                              "VFES-SEPA CONNECT DIRECT-PROD",
                              "VFES-SAP 4.7 CONNECT DIRECT-PROD",
                              "ES-CONNECT DIRECT",
                            ];
                            filtroTipoIncidente = [
                              "Incident",
                              "User Service Restoration",
                            ];
                            filtroEstado = ["Closed"];
                            filtroEnOtraCola = [undefined];
                            break;
                          case "canceladas":
                            filtroServicio = [
                              "VFES-SEPA CONNECT DIRECT-PROD",
                              "VFES-SAP 4.7 CONNECT DIRECT-PROD",
                              "ES-CONNECT DIRECT",
                            ];
                            filtroTipoIncidente = [
                              "Incident",
                              "User Service Restoration",
                            ];
                            filtroEstado = ["Cancelled"];
                            filtroEnOtraCola = [undefined];
                            break;
                          case "backlog":
                            filtroServicio = [
                              "VFES-SEPA CONNECT DIRECT-PROD",
                              "VFES-SAP 4.7 CONNECT DIRECT-PROD",
                              "ES-CONNECT DIRECT",
                            ];
                            filtroTipoIncidente = [
                              "Incident",
                              "User Service Restoration",
                            ];
                            filtroEstado = [
                              "Assigned",
                              "In Progress",
                              "Pending",
                              "Resolved",
                            ];
                            filtroEnOtraCola = [undefined];
                            break;
                          case "neto":
                            filtroServicio = [
                              "VFES-SEPA CONNECT DIRECT-PROD",
                              "VFES-SAP 4.7 CONNECT DIRECT-PROD",
                              "ES-CONNECT DIRECT",
                            ];
                            filtroTipoIncidente = [
                              "Incident",
                              "User Service Restoration",
                            ];
                            filtroEstado = [
                              "Assigned",
                              "Cancelled",
                              "Closed",
                              "In Progress",
                              "Pending",
                              "Resolved",
                            ];
                            filtroEnOtraCola = [undefined];
                            break;
                        }
                        registroResultados[registroResultados.length - 1].datos[
                          sistema
                        ][estado] = procesarRecuentoSeguimientoAM(
                          sistema,
                          estado,
                          i,
                          j,
                          filtroServicio,
                          filtroTipoIncidente,
                          filtroEstado,
                          filtroEnOtraCola,
                        );
                      }
                      break;
                  }
                } //FIN ITERACION SISTEMAS
              } //FIN ITERACION MES
            } //FIN ITERACION AÑO

            //CALCULO NETO:
            for (var i = 0; i < registroResultados.length; i++) {
              for (var sistema in registroResultados[i].datos) {
                registroResultados[i].datos[sistema].neto =
                  registroResultados[i].datos[sistema].entrada -
                  registroResultados[i].datos[sistema].salida;
              }
            }

            function procesarRecuentoSeguimientoAM(
              sistema,
              estado,
              año,
              mes,
              filtroServicio,
              filtroTipoIncidente,
              filtroEstado,
              filtroEnOtraCola,
            ) {
              //Cuenta por fila:
              var cuenta = 0;
              var cumpleFiltro = true;

              var meses = [
                "enero",
                "febrero",
                "marzo",
                "abril",
                "mayo",
                "junio",
                "julio",
                "agosto",
                "septiembre",
                "octubre",
                "noviembre",
                "diciembre",
              ];
              var filasFotoSeguimiento = archivoSeguimiento
                .sheet("Foto")
                .usedRange()._numRows;

              for (var i = 0; i < filasFotoSeguimiento; i++) {
                cumpleFiltro = true;

                //Iteración filtro Año
                if (cumpleFiltro) {
                  cumpleFiltro = false;
                } else {
                  continue;
                }

                switch (estado) {
                  case "salida":
                  case "cerradas":
                  case "canceladas":
                    if (
                      archivoSeguimiento
                        .sheet("Foto")
                        .row(i + 1)
                        .cell(16)
                        .value() == String(año)
                    ) {
                      cumpleFiltro = true;
                    }
                    break;
                  case "backlog":
                    cumpleFiltro = true;
                    break;
                  default:
                    if (
                      archivoSeguimiento
                        .sheet("Foto")
                        .row(i + 1)
                        .cell(14)
                        .value() == String(año)
                    ) {
                      cumpleFiltro = true;
                    }
                    break;
                }

                //Iteración filtro Mes
                if (cumpleFiltro) {
                  cumpleFiltro = false;
                } else {
                  continue;
                }

                switch (estado) {
                  case "salida":
                  case "cerradas":
                  case "canceladas":
                    if (
                      archivoSeguimiento
                        .sheet("Foto")
                        .row(i + 1)
                        .cell(17)
                        .value() == meses[mes]
                    ) {
                      cumpleFiltro = true;
                    }
                    break;
                  case "backlog":
                    cumpleFiltro = true;
                    break;
                  default:
                    if (
                      archivoSeguimiento
                        .sheet("Foto")
                        .row(i + 1)
                        .cell(15)
                        .value() == meses[mes]
                    ) {
                      cumpleFiltro = true;
                    }
                    break;
                }

                //Iteración filtro Servicio
                if (cumpleFiltro) {
                  cumpleFiltro = false;
                } else {
                  continue;
                }
                for (var j = 0; j < filtroServicio.length; j++) {
                  if (
                    archivoSeguimiento
                      .sheet("Foto")
                      .row(i + 1)
                      .cell(1)
                      .value() == filtroServicio[j]
                  ) {
                    cumpleFiltro = true;
                  }
                }

                //Iteración filtro Tipo Incidencia:
                if (cumpleFiltro) {
                  cumpleFiltro = false;
                } else {
                  continue;
                }
                for (var j = 0; j < filtroTipoIncidente.length; j++) {
                  if (
                    archivoSeguimiento
                      .sheet("Foto")
                      .row(i + 1)
                      .cell(5)
                      .value() == filtroTipoIncidente[j]
                  ) {
                    cumpleFiltro = true;
                  }
                }

                //Iteración filtro Estado:
                if (cumpleFiltro) {
                  cumpleFiltro = false;
                } else {
                  continue;
                }
                for (var j = 0; j < filtroEstado.length; j++) {
                  if (
                    archivoSeguimiento
                      .sheet("Foto")
                      .row(i + 1)
                      .cell(6)
                      .value() == filtroEstado[j]
                  ) {
                    cumpleFiltro = true;
                  }
                }

                //Iteración filtro Tipo En Otra Cola:
                if (cumpleFiltro) {
                  cumpleFiltro = false;
                } else {
                  continue;
                }

                for (var j = 0; j < filtroEnOtraCola.length; j++) {
                  if (
                    archivoSeguimiento
                      .sheet("Foto")
                      .row(i + 1)
                      .cell(26)
                      .value() == filtroEnOtraCola[j]
                  ) {
                    cumpleFiltro = true;
                  }
                }

                //Iteración filtro Aplicativo (Solo Backlog):
                if (estado == "backlog") {
                  if (cumpleFiltro) {
                    cumpleFiltro = false;
                  } else {
                    continue;
                  }

                  var filtroAplicativo = [];

                  switch (sistema) {
                    case "general":
                      filtroAplicativo = [
                        "ECC 6.0",
                        "SAP 4.7",
                        "Editran",
                        "Connect Direct",
                      ];
                      break;

                    case "rmca":
                      filtroAplicativo = ["ECC 6.0"];
                      break;

                    case "sap":
                      filtroAplicativo = ["SAP 4.7"];
                      break;

                    case "editran":
                      filtroAplicativo = ["Editran"];
                      break;

                    case "connectDirect":
                      filtroAplicativo = ["Connect Direct"];
                      break;
                  }

                  for (var j = 0; j < filtroAplicativo.length; j++) {
                    if (
                      archivoSeguimiento
                        .sheet("Foto")
                        .row(i + 1)
                        .cell(21)
                        .value() == filtroAplicativo[j]
                    ) {
                      cumpleFiltro = true;
                    }
                  }
                }

                if (cumpleFiltro) {
                  cuenta++;
                }
              } //FIN ITERACIÓN RECUENTO:
              console.log(
                "Cuenta Sistema: " +
                  sistema +
                  " Estado: " +
                  estado +
                  " Cuenta: " +
                  cuenta,
              );
              return cuenta;
            } //FIN FUNCION

            //PROCESAR SALIDA:

            // 6) Procesar recuento RMCA (Entrada,Salida,Cerradas,Canceladas,Backlog,Resolución Neta(E - S));

            // 7) Procesar recuento SAP 47 (Entrada,Salida,Cerradas,Canceladas,Backlog,Resolución Neta(E - S));

            // 8) Procesar recuento EDITRAN (Entrada,Salida,Cerradas,Canceladas,Backlog,Resolución Neta(E - S));

            // 9) Procesar recuento Connect Direct (Entrada,Salida,Cerradas,Canceladas,Backlog,Resolución Neta(E - S));

            //LOG DE RESULTADOS:
            console.log("*******************");
            console.log("    RESULTADOS: ");
            console.log("*******************");

            for (var i = 0; i < registroResultados.length; i++) {
              console.log("");
              console.log("--------------");
              console.log("Año: " + registroResultados[i].año);
              console.log("Mes: " + (registroResultados[i].mes + 1));
              console.log("--------------");
              console.log("");

              //ITERAR SISTEMA:
              for (var sistema in registroResultados[i].datos) {
                console.log("");
                console.log(sistema.toUpperCase());
                console.log("--------------");
                for (var estado in registroResultados[i].datos[sistema]) {
                  console.log(
                    estado +
                      ": " +
                      registroResultados[i].datos[sistema][estado],
                  );
                }
              }
            }

            for (var i = 0; i < registroResultados.length; i++) {
              console.log("");
              console.log("--------------");
              console.log("Año: " + registroResultados[i].año);
              console.log("Mes: " + (registroResultados[i].mes + 1));
              console.log("--------------");
              console.log("");

              console.log(JSON.stringify(registroResultados[i]));
            }

            //10) Añadir columnas graficos:

            //Detectar columna del Mes:
            var mesUltimaColumna;
            var añoUltimaColumna;
            var ultimaColumna;

            function ExcelDateToJSDate(serial) {
              var utc_days = Math.floor(serial - 25569);
              var utc_value = utc_days * 86400;
              var date_info = new Date(utc_value * 1000);
              var fractional_day = serial - Math.floor(serial) + 0.0000001;
              var total_seconds = Math.floor(86400 * fractional_day);
              var seconds = total_seconds % 60;
              total_seconds -= seconds;
              var hours = Math.floor(total_seconds / (60 * 60));
              var minutes = Math.floor(total_seconds / 60) % 60;

              return new Date(
                date_info.getFullYear(),
                date_info.getMonth(),
                date_info.getDate(),
                hours,
                minutes,
                seconds,
              );
            }

            function isValidDate(d) {
              return d instanceof Date && !isNaN(d);
            }

            var hojasResumen = [
              "Resumen General",
              "Resumen RMCA",
              "Resumen SAP 4.7",
              "Resumen Editran",
              "Resumen CD",
            ];
            var sistemas = [
              "general",
              "rmca",
              "sap",
              "editran",
              "connectDirect",
            ];

            //Iteración por hojas resumen:
            for (var k = 0; k < hojasResumen.length; k++) {
              //Inicializacion:
              ultimaColumna = 0;

              // Detección de ultima columna
              for (
                var i = 0;
                i <
                archivoSeguimiento.sheet(hojasResumen[k]).usedRange()
                  ._numColumns;
                i++
              ) {
                if (
                  isValidDate(
                    new Date(
                      ExcelDateToJSDate(
                        archivoSeguimiento
                          .sheet(hojasResumen[k])
                          .row(1)
                          .cell(i + 1)
                          .value(),
                      ),
                    ),
                  )
                ) {
                  ultimaColumna = i + 1;
                  añoUltimaColumna = new Date(
                    ExcelDateToJSDate(
                      archivoSeguimiento
                        .sheet(hojasResumen[k])
                        .row(1)
                        .cell(i + 1)
                        .value(),
                    ),
                  ).getFullYear();
                  mesUltimaColumna = new Date(
                    ExcelDateToJSDate(
                      archivoSeguimiento
                        .sheet(hojasResumen[k])
                        .row(1)
                        .cell(i + 1)
                        .value(),
                    ),
                  ).getMonth();
                }
              }

              console.log("Ultima Columna" + ultimaColumna);
              console.log("Año: " + añoUltimaColumna);
              console.log("Mes: " + mesUltimaColumna);

              if (añoUltimaColumna == añoFin && mesUltimaColumna == mesFin) {
                archivoSeguimiento
                  .sheet(hojasResumen[k])
                  .row(2)
                  .cell(ultimaColumna)
                  .value(
                    registroResultados[registroResultados.length - 1].datos[
                      sistemas[k]
                    ]["entrada"],
                  );
                archivoSeguimiento
                  .sheet(hojasResumen[k])
                  .row(3)
                  .cell(ultimaColumna)
                  .value(
                    registroResultados[registroResultados.length - 1].datos[
                      sistemas[k]
                    ]["salida"],
                  );
                archivoSeguimiento
                  .sheet(hojasResumen[k])
                  .row(4)
                  .cell(ultimaColumna)
                  .value(
                    registroResultados[registroResultados.length - 1].datos[
                      sistemas[k]
                    ]["cerradas"],
                  );
                archivoSeguimiento
                  .sheet(hojasResumen[k])
                  .row(5)
                  .cell(ultimaColumna)
                  .value(
                    registroResultados[registroResultados.length - 1].datos[
                      sistemas[k]
                    ]["canceladas"],
                  );
                archivoSeguimiento
                  .sheet(hojasResumen[k])
                  .row(6)
                  .cell(ultimaColumna)
                  .value(
                    registroResultados[registroResultados.length - 1].datos[
                      sistemas[k]
                    ]["backlog"],
                  );
                archivoSeguimiento
                  .sheet(hojasResumen[k])
                  .row(7)
                  .cell(ultimaColumna)
                  .value(
                    registroResultados[registroResultados.length - 1].datos[
                      sistemas[k]
                    ]["neto"],
                  );
              }

              if (
                añoUltimaColumna == añoFin &&
                mesUltimaColumna == mesFin - 1
              ) {
                console.log(
                  "Modificando Columna: " +
                    registroResultados[registroResultados.length - 1].mes,
                );

                archivoSeguimiento
                  .sheet(hojasResumen[k])
                  .row(2)
                  .cell(ultimaColumna)
                  .value(
                    registroResultados[registroResultados.length - 2].datos[
                      sistemas[k]
                    ]["entrada"],
                  );
                archivoSeguimiento
                  .sheet(hojasResumen[k])
                  .row(3)
                  .cell(ultimaColumna)
                  .value(
                    registroResultados[registroResultados.length - 2].datos[
                      sistemas[k]
                    ]["salida"],
                  );
                archivoSeguimiento
                  .sheet(hojasResumen[k])
                  .row(4)
                  .cell(ultimaColumna)
                  .value(
                    registroResultados[registroResultados.length - 2].datos[
                      sistemas[k]
                    ]["cerradas"],
                  );
                archivoSeguimiento
                  .sheet(hojasResumen[k])
                  .row(5)
                  .cell(ultimaColumna)
                  .value(
                    registroResultados[registroResultados.length - 2].datos[
                      sistemas[k]
                    ]["canceladas"],
                  );
                archivoSeguimiento
                  .sheet(hojasResumen[k])
                  .row(6)
                  .cell(ultimaColumna)
                  .value(
                    registroResultados[registroResultados.length - 2].datos[
                      sistemas[k]
                    ]["backlog"],
                  );
                archivoSeguimiento
                  .sheet(hojasResumen[k])
                  .row(7)
                  .cell(ultimaColumna)
                  .value(
                    registroResultados[registroResultados.length - 2].datos[
                      sistemas[k]
                    ]["neto"],
                  );

                archivoSeguimiento
                  .sheet(hojasResumen[k])
                  .row(1)
                  .cell(ultimaColumna + 1)
                  .value(new Date(añoFin, mesFin, 1))
                  .style("numberFormat", "mmm-yy");

                archivoSeguimiento
                  .sheet(hojasResumen[k])
                  .row(2)
                  .cell(ultimaColumna + 1)
                  .value(
                    registroResultados[registroResultados.length - 1].datos[
                      sistemas[k]
                    ]["entrada"],
                  );
                archivoSeguimiento
                  .sheet(hojasResumen[k])
                  .row(3)
                  .cell(ultimaColumna + 1)
                  .value(
                    registroResultados[registroResultados.length - 1].datos[
                      sistemas[k]
                    ]["salida"],
                  );
                archivoSeguimiento
                  .sheet(hojasResumen[k])
                  .row(4)
                  .cell(ultimaColumna + 1)
                  .value(
                    registroResultados[registroResultados.length - 1].datos[
                      sistemas[k]
                    ]["cerradas"],
                  );
                archivoSeguimiento
                  .sheet(hojasResumen[k])
                  .row(5)
                  .cell(ultimaColumna + 1)
                  .value(
                    registroResultados[registroResultados.length - 1].datos[
                      sistemas[k]
                    ]["canceladas"],
                  );
                archivoSeguimiento
                  .sheet(hojasResumen[k])
                  .row(6)
                  .cell(ultimaColumna + 1)
                  .value(
                    registroResultados[registroResultados.length - 1].datos[
                      sistemas[k]
                    ]["backlog"],
                  );
                archivoSeguimiento
                  .sheet(hojasResumen[k])
                  .row(7)
                  .cell(ultimaColumna + 1)
                  .value(
                    registroResultados[registroResultados.length - 1].datos[
                      sistemas[k]
                    ]["neto"],
                  );
              }
            }

            //11) Guardar Archivo:

            //Fin de procesamiento:
            console.log("Escribiendo archivo...");
            console.log(
              "Path: " +
                path.normalize(
                  path.join(
                    configuracion.datosSalidaControl,
                    configuracion.nombreArchivoSalidaControl + ".xlsx",
                  ),
                ),
            );
            archivoSeguimiento
              .toFileAsync(
                path.normalize(
                  path.join(
                    configuracion.datosSalidaControl,
                    configuracion.nombreArchivoSalidaControl + ".xlsx",
                  ),
                ),
              )
              .then(() => {
                console.log("Fin del procesamiento");
                resolve(true);
              })
              .catch((err) => {
                console.log("Se ha producido un error interno: ");
                console.log(err);
                var tituloError =
                  "Se ha producido un error escribiendo el archivo: " +
                  path.normalize(
                    path.join(
                      configuracion.datosSalidaControl,
                      configuracion.nombreArchivoSalidaControl + ".xlsx",
                    ),
                  );
                resolve(false);
              });

            resolve(true);
          });
      }).catch((err) => {
        console.log("Se ha producido un error interno: ");
        console.log(err);
        var tituloError =
          "Se ha producido un error interno cargando los archivos.";
        mainWindow.webContents.send("onErrorInterno", tituloError, err);
        resolve(false);
      });
    }

    var resultado = await generarReportAM(
      archivoSeguimiento,
      datosNacho,
      configuracion,
    );
    return resultado;
  } //Fin de generación de report AM

  async fusionarObjetos(argumentos) {
    console.log("Fusionar Archivos:");
    var archivoBase = argumentos[0][0];
    var archivoAdd = argumentos[1][0];

    async function fusionarArchivos(archivoBase, archivoAdd) {
      return new Promise((resolve) => {
        for (var i = 0; i < archivoAdd.data.length; i++) {
          archivoBase.data.push(archivoAdd.data[i]);
        }
        resolve(archivoBase);
      });
    }

    var result = await fusionarArchivos(archivoBase, archivoAdd);
    console.log("LOGITUD FINAL: " + result.data.length);
    result["objetoId"] = archivoBase.nombreId;
    return result;
  }

  async procesarIBAN(argumentos) {
    console.log("Procesando Recuento IBAN - Mandato");

    var rutaGuardado = argumentos[1];
    var nombreGuardado = argumentos[2];

    console.log("TAMANO DATOS ARVIVO 1: " + argumentos[0][0].data.length);
    console.log("TAMANO DATOS ARVIVO 2: " + argumentos[1][0].data.length);
    console.log("TAMANO DATOS ARVIVO 3: " + argumentos[2][0].data.length);
    console.log("TAMANO DATOS ARVIVO 4: " + argumentos[3][0].data.length);
    console.log("TAMANO DATOS ARVIVO 5: " + argumentos[4][0].data.length);
    console.log("TAMANO DATOS ARVIVO 6: " + argumentos[5][0].data.length);
    console.log("TAMANO DATOS ARVIVO 7: " + argumentos[6][0].data.length);
    console.log("TAMANO DATOS ARVIVO 8: " + argumentos[7][0].data.length);
    console.log("TAMANO DATOS ARVIVO 9: " + argumentos[8][0].data.length);
    console.log("TAMANO DATOS ARVIVO 10: " + argumentos[9][0].data.length);
    console.log("TAMANO DATOS ARVIVO 11: " + argumentos[10][0].data.length);
    console.log("TAMANO DATOS ARVIVO 12: " + argumentos[11][0].data.length);
    console.log("TAMANO DATOS ARVIVO 13: " + argumentos[12][0].data.length);
    console.log("TAMANO DATOS ARVIVO 14: " + argumentos[13][0].data.length);
    console.log("TAMANO DATOS ARVIVO 15: " + argumentos[14][0].data.length);
    var suma = 0;
    for (var i = 0; i < 1; i++) {
      suma += argumentos[i][0].data.length;
    }

    console.log("Tamaña total: " + suma);

    //Procesado de datos:
    var numeroProcesado = 0;
    var analizandoIBAN;
    var arrayAnalizados = [];
    var objetoSalida;
    var analizado = false;
    var cuentaIbanEncontrado = 0;

    var matrizIban = [];

    //Rellenar matriz de IBAN:
    console.log("Generando Matriz");
    for (
      var iteracionDocumento = 0;
      iteracionDocumento < 15;
      iteracionDocumento++
    ) {
      for (
        var iteracionRegistro = 0;
        iteracionRegistro < argumentos[iteracionDocumento][0].data.length;
        iteracionRegistro++
      ) {
        try {
          matrizIban.push(
            argumentos[iteracionDocumento][0].data[iteracionRegistro]["iban"],
          );
        } catch {
          console.log(
            "Error IBAN; Documento: " +
              iteracionDocumento +
              " Registro: " +
              iteracionRegistro,
          );
        }
      }
    }
    console.log("Matriz finalizada");

    const countOccurrences = (arr) =>
      arr.reduce((prev, curr) => ((prev[curr] = ++prev[curr] || 1), prev), {});

    //Ordenando matriz:
    console.log("Ordenando Matriz: ");
    matrizIban = matrizIban.sort();

    console.log("Calculando ocurrencias");
    objetoSalida = countOccurrences(matrizIban.sort());

    //Formateando objeto salida;
    console.log("Depurando salida");
    for (const property in objetoSalida) {
      if (objetoSalida[property] === 1) {
        delete objetoSalida[property];
      }
    }

    //Iterar en la matriz:
    /*
		for(var i=0; i<1; i++){
			for(var j=0; j<matrizIban[i].length; j++){
				analizandoIBAN = matrizIban[i][j];

				//Verifica si ese IBAN ya ha sido analizado:
				analizado= false;
				for(var iteracionAnalizado = 0; iteracionAnalizado< arrayAnalizados.length; iteracionAnalizado++){
					if(analizandoIBAN === arrayAnalizados[iteracionAnalizado]){analizado = true;}
				}

				if(analizado){continue;}

				//Iteración en cada documento: 
				cuentaIbanEncontrado = 0;
				for(var iteracionDocumento = 0; iteracionDocumento<15; iteracionDocumento++){
					for(var iteracionRegistro = 0; iteracionRegistro< matrizIban[iteracionDocumento].length; iteracionRegistro++){
						if(analizandoIBAN === matrizIban[iteracionDocumento][iteracionRegistro]){
							cuentaIbanEncontrado += 1;
						}
					}
				}

				//Registra los resultados:
				if(cuentaIbanEncontrado > 1){
					arrayAnalizados.push(analizandoIBAN);
					objetoSalida.push({
						iban: analizandoIBAN,
						cuenta: cuentaIbanEncontrado
					})
				}

				numeroProcesado += 1;
				if(numeroProcesado === 15000){
					console.log("PROGRESO PAECIAL");
				}
			}
		}*/
    console.log("Escribiendo archivo");
    var pathSpoolOutput;

    if (
      argumentos[16].slice(-4) !== ".txt" &&
      argumentos[16].slice(-4) !== ".TXT"
    ) {
      pathSpoolOutput = path.join(argumentos[15], argumentos[16] + ".txt");
    } else {
      pathSpoolOutput = path.join(argumentos[15], argumentos[16]);
    }

    const outputFile = fs.createWriteStream(pathSpoolOutput);

    outputFile.on("err", (err) => {
      // handle error
      console.log(err);
    });

    outputFile.on("close", () => {
      console.log("done writing");
    });
    for (const property in objetoSalida) {
      outputFile.write(`${property}\t${objetoSalida[property]}\n`);
    }

    console.log("FIN de procesamiento IBAN");
    var result = true;
    return result;
  }

  async subirCursos(argumentos) {
    console.log("SUBIENDO MONITORIZACIÓN CURSOS");

    const pathMonitorizacionCursos = path.join(argumentos[0]);
    //const pathRaiz = pathMonitorizacionCursos.substring(0, pathMonitorizacionCursos.lastIndexOf("\\"));
    const pathRaiz = path.dirname(pathMonitorizacionCursos);

    var cursos = argumentos[1];
    var formadores = argumentos[3];
    var formadorCurso = argumentos[5];
    var codigosProvincia = argumentos[6];
    var instituciones = argumentos[7];

    console.log("PATH RAIZ: " + pathRaiz);
    console.log("RUTA MONITORIZACIÓN CURSO: " + pathMonitorizacionCursos);

    //Importando XLSX:
    return new Promise((resolve) => {
      var monitorizacionCursos = {};
      XlsxPopulate.fromFileAsync(path.normalize(pathMonitorizacionCursos))
        .then((workbook) => {
          console.log("Archivo Cargado: Monitorización Cursos");
          monitorizacionCursos = workbook;
        })
        .then(() => {
          //IDENTIFICAR CAMBIOS:
          var cambios = [];
          for (var i = 0; i < cursos.length; i++) {
            if (cursos[i].metadatos.flag_cambio && !cursos[i].metadatos.error) {
              cambios.push(cursos[i]);
            }
          }

          var cambiosFormadores = [];
          for (var i = 0; i < formadores[0].data.length; i++) {
            if (
              formadores[0].data[i].metadatos.flag_cambio &&
              !formadores[0].data[i].metadatos.error
            ) {
              cambiosFormadores.push(formadores[0].data[i]);
            }
          }

          var cambiosInstituciones = [];
          for (var i = 0; i < instituciones[0].data.length; i++) {
            if (
              instituciones[0].data[i].metadatos.flag_cambio &&
              !instituciones[0].data[i].metadatos.error
            ) {
              cambiosInstituciones.push(instituciones[0].data[i]);
            }
          }

          console.log("Cambios Cursos Detectados: " + cambios.length);
          console.log(cambios);
          console.log(
            "Cambios Formadores Detectados: " + cambiosFormadores.length,
          );
          console.log(cambiosFormadores);
          console.log(
            "Cambios Instituciones Detectados: " + cambiosInstituciones.length,
          );
          console.log(cambiosInstituciones);

          //Aplicando Cambios Cursos:
          var contadorNuevas = 0;
          var contadorModificacion = 0;
          var columnasCursos = monitorizacionCursos
            .sheet("Cursos")
            .usedRange()._numColumns;
          var filasCursos = monitorizacionCursos
            .sheet("Cursos")
            .usedRange()._numRows;
          var filasFormadoresCursos = monitorizacionCursos
            .sheet("Formador-Curso")
            .usedRange()._numRows;

          //Recalculo de filas usadas:
          while (
            !monitorizacionCursos
              .sheet("Cursos")
              .row(filasCursos)
              .cell(1)
              .value()
          ) {
            filasCursos--;
          }
          while (
            !monitorizacionCursos
              .sheet("Formador-Curso")
              .row(filasFormadoresCursos)
              .cell(1)
              .value()
          ) {
            filasFormadoresCursos--;
          }

          var punteroRegistroFormador = filasFormadoresCursos + 1;
          var encontrado = false;

          for (var i = 0; i < cambios.length; i++) {
            encontrado = false;
            for (var j = 1; j < filasCursos; j++) {
              //Si se encuentra el registro:
              if (
                monitorizacionCursos
                  .sheet("Cursos")
                  .row(j + 1)
                  .cell(1)
                  .value() == cambios[i]["cod_curso"]
              ) {
                contadorModificacion++;
                encontrado = true;

                //Reescribir Registro:
                monitorizacionCursos
                  .sheet("Cursos")
                  .row(j + 1)
                  .cell(2)
                  .value(cambios[i]["cod_grupo"]);
                monitorizacionCursos
                  .sheet("Cursos")
                  .row(j + 1)
                  .cell(3)
                  .value(cambios[i]["cod__postal"]);
                monitorizacionCursos
                  .sheet("Cursos")
                  .row(j + 1)
                  .cell(4)
                  .value(cambios[i]["territorial"]);
                monitorizacionCursos
                  .sheet("Cursos")
                  .row(j + 1)
                  .cell(5)
                  .value(cambios[i]["ccaa_/_pais"]);
                monitorizacionCursos
                  .sheet("Cursos")
                  .row(j + 1)
                  .cell(6)
                  .value(cambios[i]["curso"]);
                monitorizacionCursos
                  .sheet("Cursos")
                  .row(j + 1)
                  .cell(7)
                  .value(cambios[i]["sesión"]);
                monitorizacionCursos
                  .sheet("Cursos")
                  .row(j + 1)
                  .cell(8)
                  .value(cambios[i]["fecha"]);
                monitorizacionCursos
                  .sheet("Cursos")
                  .row(j + 1)
                  .cell(9)
                  .value(cambios[i]["hora_inicio"]);
                monitorizacionCursos
                  .sheet("Cursos")
                  .row(j + 1)
                  .cell(10)
                  .value(cambios[i]["hora_fin"]);
                monitorizacionCursos
                  .sheet("Cursos")
                  .row(j + 1)
                  .cell(11)
                  .value(cambios[i]["duración"]);
                monitorizacionCursos
                  .sheet("Cursos")
                  .row(j + 1)
                  .cell(12)
                  .value(cambios[i]["institución"]);
                monitorizacionCursos
                  .sheet("Cursos")
                  .row(j + 1)
                  .cell(13)
                  .value(cambios[i]["colectivo"]);
                monitorizacionCursos
                  .sheet("Cursos")
                  .row(j + 1)
                  .cell(14)
                  .value(cambios[i]["grupo"]);
                monitorizacionCursos
                  .sheet("Cursos")
                  .row(j + 1)
                  .cell(15)
                  .value(cambios[i]["nºasistentes"]);
                monitorizacionCursos
                  .sheet("Cursos")
                  .row(j + 1)
                  .cell(16)
                  .value(cambios[i]["modalidad"]);
                monitorizacionCursos
                  .sheet("Cursos")
                  .row(j + 1)
                  .cell(17)
                  .value(cambios[i]["estado"]);
                monitorizacionCursos
                  .sheet("Cursos")
                  .row(j + 1)
                  .cell(18)
                  .value(cambios[i]["material"]);
                if (typeof cambios[i]["valoración"] != "undefined") {
                  monitorizacionCursos
                    .sheet("Cursos")
                    .row(j + 1)
                    .cell(19)
                    .value(cambios[i]["valoración"]);
                } else {
                  monitorizacionCursos
                    .sheet("Cursos")
                    .row(j + 1)
                    .cell(19)
                    .value("SIN VALORAR");
                }
                monitorizacionCursos
                  .sheet("Cursos")
                  .row(j + 1)
                  .cell(20)
                  .value(cambios[i]["observaciones"]);
                break;
              }
            }

            if (!encontrado) {
              //Crear Nuevo Curso:
              monitorizacionCursos
                .sheet("Cursos")
                .row(filasCursos + contadorNuevas + 1)
                .cell(1)
                .value(cambios[i]["cod_curso"]);
              monitorizacionCursos
                .sheet("Cursos")
                .row(filasCursos + contadorNuevas + 1)
                .cell(2)
                .value(cambios[i]["cod_grupo"]);
              monitorizacionCursos
                .sheet("Cursos")
                .row(filasCursos + contadorNuevas + 1)
                .cell(3)
                .value(cambios[i]["cod__postal"]);
              monitorizacionCursos
                .sheet("Cursos")
                .row(filasCursos + contadorNuevas + 1)
                .cell(4)
                .value(cambios[i]["territorial"]);
              monitorizacionCursos
                .sheet("Cursos")
                .row(filasCursos + contadorNuevas + 1)
                .cell(5)
                .value(cambios[i]["ccaa_/_pais"]);
              monitorizacionCursos
                .sheet("Cursos")
                .row(filasCursos + contadorNuevas + 1)
                .cell(6)
                .value(cambios[i]["curso"]);
              monitorizacionCursos
                .sheet("Cursos")
                .row(filasCursos + contadorNuevas + 1)
                .cell(7)
                .value(cambios[i]["sesión"]);
              monitorizacionCursos
                .sheet("Cursos")
                .row(filasCursos + contadorNuevas + 1)
                .cell(8)
                .value(cambios[i]["fecha"]);
              monitorizacionCursos
                .sheet("Cursos")
                .row(filasCursos + contadorNuevas + 1)
                .cell(9)
                .value(cambios[i]["hora_inicio"]);
              monitorizacionCursos
                .sheet("Cursos")
                .row(filasCursos + contadorNuevas + 1)
                .cell(10)
                .value(cambios[i]["hora_fin"]);
              monitorizacionCursos
                .sheet("Cursos")
                .row(filasCursos + contadorNuevas + 1)
                .cell(11)
                .value(cambios[i]["duración"]);
              monitorizacionCursos
                .sheet("Cursos")
                .row(filasCursos + contadorNuevas + 1)
                .cell(12)
                .value(cambios[i]["institución"]);
              monitorizacionCursos
                .sheet("Cursos")
                .row(filasCursos + contadorNuevas + 1)
                .cell(13)
                .value(cambios[i]["colectivo"]);
              monitorizacionCursos
                .sheet("Cursos")
                .row(filasCursos + contadorNuevas + 1)
                .cell(14)
                .value(cambios[i]["grupo"]);
              monitorizacionCursos
                .sheet("Cursos")
                .row(filasCursos + contadorNuevas + 1)
                .cell(15)
                .value(cambios[i]["nºasistentes"]);
              monitorizacionCursos
                .sheet("Cursos")
                .row(filasCursos + contadorNuevas + 1)
                .cell(16)
                .value(cambios[i]["modalidad"]);
              monitorizacionCursos
                .sheet("Cursos")
                .row(filasCursos + contadorNuevas + 1)
                .cell(17)
                .value(cambios[i]["estado"]);
              monitorizacionCursos
                .sheet("Cursos")
                .row(filasCursos + contadorNuevas + 1)
                .cell(18)
                .value(cambios[i]["material"]);
              if (typeof cambios[i]["valoración"] != "undefined") {
                monitorizacionCursos
                  .sheet("Cursos")
                  .row(filasCursos + contadorNuevas + 1)
                  .cell(19)
                  .value(cambios[i]["valoración"]);
              } else {
                monitorizacionCursos
                  .sheet("Cursos")
                  .row(filasCursos + contadorNuevas + 1)
                  .cell(19)
                  .value("SIN VALORAR");
              }
              monitorizacionCursos
                .sheet("Cursos")
                .row(filasCursos + contadorNuevas + 1)
                .cell(20)
                .value(cambios[i]["observaciones"]);

              //Actualizar Contador:
              contadorNuevas++;
            }

            //Modificar Curso-Formador:

            // 1) Eliminar Referencias al curso en Curso-Formadores:
            filasFormadoresCursos = monitorizacionCursos
              .sheet("Formador-Curso")
              .usedRange()._numRows;
            for (var k = 1; k < filasFormadoresCursos + 1; k++) {
              if (
                monitorizacionCursos
                  .sheet("Formador-Curso")
                  .row(k)
                  .cell(1)
                  .value() == cambios[i]["cod_curso"]
              ) {
                monitorizacionCursos
                  .sheet("Formador-Curso")
                  .row(k)
                  .cell(1)
                  .value("");
                monitorizacionCursos
                  .sheet("Formador-Curso")
                  .row(k)
                  .cell(2)
                  .value("");
              }
            }

            // 2) Añadiendo Formadores:
            if (typeof cambios[i].metadatos["formadores"] == "object") {
              for (
                var k = 0;
                k < cambios[i].metadatos["formadores"].length;
                k++
              ) {
                monitorizacionCursos
                  .sheet("Formador-Curso")
                  .row(punteroRegistroFormador)
                  .cell(1)
                  .value(cambios[i]["cod_curso"]);
                monitorizacionCursos
                  .sheet("Formador-Curso")
                  .row(punteroRegistroFormador)
                  .cell(2)
                  .value(cambios[i]["metadatos"]["formadores"][k]["id"]);
                punteroRegistroFormador++;
              }
            }
          } //Fin de iteracion de cambios CURSOS

          //Aplicando Cambios Formadores:
          var contadorNuevosFormadores = 0;
          var contadorModificacionFormadores = 0;
          var columnasFormadores = monitorizacionCursos
            .sheet("Formadores")
            .usedRange()._numColumns;
          var filasFormadores = monitorizacionCursos
            .sheet("Formadores")
            .usedRange()._numRows;
          var formadorEncontrado = false;

          //Recalculo de filas usadas:
          while (
            !monitorizacionCursos
              .sheet("Formadores")
              .row(filasFormadores)
              .cell(1)
              .value()
          ) {
            filasFormadores--;
          }

          for (var i = 0; i < cambiosFormadores.length; i++) {
            formadorEncontrado = false;
            for (var j = 1; j < filasFormadores; j++) {
              if (
                monitorizacionCursos
                  .sheet("Formadores")
                  .row(j + 1)
                  .cell(1)
                  .value() == cambiosFormadores[i]["cod__formador"]
              ) {
                contadorModificacionFormadores++;
                formadorEncontrado = true;

                //Reescribir Registro:
                monitorizacionCursos
                  .sheet("Formadores")
                  .row(j + 1)
                  .cell(2)
                  .value(cambiosFormadores[i]["nombre"]);
                monitorizacionCursos
                  .sheet("Formadores")
                  .row(j + 1)
                  .cell(3)
                  .value(cambiosFormadores[i]["email"]);
                monitorizacionCursos
                  .sheet("Formadores")
                  .row(j + 1)
                  .cell(4)
                  .value(cambiosFormadores[i]["telefono"]);
                monitorizacionCursos
                  .sheet("Formadores")
                  .row(j + 1)
                  .cell(5)
                  .value(cambiosFormadores[i]["territorial"]);
                monitorizacionCursos
                  .sheet("Formadores")
                  .row(j + 1)
                  .cell(6)
                  .value(cambiosFormadores[i]["ccaa"]);
                monitorizacionCursos
                  .sheet("Formadores")
                  .row(j + 1)
                  .cell(7)
                  .value(cambiosFormadores[i]["provincia"]);
                monitorizacionCursos
                  .sheet("Formadores")
                  .row(j + 1)
                  .cell(8)
                  .value(cambiosFormadores[i]["fecha"]);
                monitorizacionCursos
                  .sheet("Formadores")
                  .row(j + 1)
                  .cell(9)
                  .value(cambiosFormadores[i]["certificado"]);
                monitorizacionCursos
                  .sheet("Formadores")
                  .row(j + 1)
                  .cell(10)
                  .value(cambiosFormadores[i]["confidencialidad"]);
                monitorizacionCursos
                  .sheet("Formadores")
                  .row(j + 1)
                  .cell(11)
                  .value(cambiosFormadores[i]["consentimiento"]);
                monitorizacionCursos
                  .sheet("Formadores")
                  .row(j + 1)
                  .cell(12)
                  .value(cambiosFormadores[i]["estado"]);
                contadorModificacionFormadores++;
                break;
              }
            }

            //NO ENCONTRADO
            if (!formadorEncontrado) {
              //Nuevo Formador:
              monitorizacionCursos
                .sheet("Formadores")
                .row(filasFormadores + contadorNuevosFormadores + 1)
                .cell(1)
                .value(cambiosFormadores[i]["cod__formador"]);
              monitorizacionCursos
                .sheet("Formadores")
                .row(filasFormadores + contadorNuevosFormadores + 1)
                .cell(2)
                .value(cambiosFormadores[i]["nombre"]);
              monitorizacionCursos
                .sheet("Formadores")
                .row(filasFormadores + contadorNuevosFormadores + 1)
                .cell(3)
                .value(cambiosFormadores[i]["email"]);
              monitorizacionCursos
                .sheet("Formadores")
                .row(filasFormadores + contadorNuevosFormadores + 1)
                .cell(4)
                .value(cambiosFormadores[i]["telefono"]);
              monitorizacionCursos
                .sheet("Formadores")
                .row(filasFormadores + contadorNuevosFormadores + 1)
                .cell(5)
                .value(cambiosFormadores[i]["territorial"]);
              monitorizacionCursos
                .sheet("Formadores")
                .row(filasFormadores + contadorNuevosFormadores + 1)
                .cell(6)
                .value(cambiosFormadores[i]["ccaa"]);
              monitorizacionCursos
                .sheet("Formadores")
                .row(filasFormadores + contadorNuevosFormadores + 1)
                .cell(7)
                .value(cambiosFormadores[i]["provincia"]);
              monitorizacionCursos
                .sheet("Formadores")
                .row(filasFormadores + contadorNuevosFormadores + 1)
                .cell(8)
                .value(cambiosFormadores[i]["fecha"]);
              monitorizacionCursos
                .sheet("Formadores")
                .row(filasFormadores + contadorNuevosFormadores + 1)
                .cell(9)
                .value(cambiosFormadores[i]["certificado"]);
              monitorizacionCursos
                .sheet("Formadores")
                .row(filasFormadores + contadorNuevosFormadores + 1)
                .cell(10)
                .value(cambiosFormadores[i]["confidencialidad"]);
              monitorizacionCursos
                .sheet("Formadores")
                .row(filasFormadores + contadorNuevosFormadores + 1)
                .cell(11)
                .value(cambiosFormadores[i]["consentimiento"]);
              monitorizacionCursos
                .sheet("Formadores")
                .row(filasFormadores + contadorNuevosFormadores + 1)
                .cell(12)
                .value(cambiosFormadores[i]["estado"]);

              //Actualizar Contador:
              contadorNuevosFormadores++;
            }
          } //Fin de iteracion de cambios Formadores

          //Aplicando Cambios Institución:
          var contadorNuevasInstituciones = 0;
          var contadorModificacionInstituciones = 0;
          var columnasInstituciones = monitorizacionCursos
            .sheet("Instituciones")
            .usedRange()._numColumns;
          var filasInstituciones = monitorizacionCursos
            .sheet("Instituciones")
            .usedRange()._numRows;
          var institucionEncontrada = false;

          //Recalculo de filas usadas:
          while (
            !monitorizacionCursos
              .sheet("Instituciones")
              .row(filasInstituciones)
              .cell(1)
              .value()
          ) {
            filasInstituciones--;
          }

          for (var i = 0; i < cambiosInstituciones.length; i++) {
            institucionEncontrada = false;
            for (var j = 1; j < filasInstituciones; j++) {
              if (
                monitorizacionCursos
                  .sheet("Instituciones")
                  .row(j + 1)
                  .cell(1)
                  .value() == cambiosInstituciones[i]["cod_institucion"]
              ) {
                contadorModificacionInstituciones++;
                institucionEncontrada = true;

                //Reescribir Registro:
                monitorizacionCursos
                  .sheet("Instituciones")
                  .row(j + 1)
                  .cell(2)
                  .value(cambiosInstituciones[i]["institucion"]);
                monitorizacionCursos
                  .sheet("Instituciones")
                  .row(j + 1)
                  .cell(3)
                  .value(cambiosInstituciones[i]["tipo"]);
                monitorizacionCursos
                  .sheet("Instituciones")
                  .row(j + 1)
                  .cell(4)
                  .value(cambiosInstituciones[i]["cod__postal"]);
                monitorizacionCursos
                  .sheet("Instituciones")
                  .row(j + 1)
                  .cell(5)
                  .value(cambiosInstituciones[i]["territorial"]);
                monitorizacionCursos
                  .sheet("Instituciones")
                  .row(j + 1)
                  .cell(6)
                  .value(cambiosInstituciones[i]["ccaa_/_pais"]);
                monitorizacionCursos
                  .sheet("Instituciones")
                  .row(j + 1)
                  .cell(7)
                  .value(cambiosInstituciones[i]["provincia"]);
                monitorizacionCursos
                  .sheet("Instituciones")
                  .row(j + 1)
                  .cell(8)
                  .value(cambiosInstituciones[i]["contacto1"]);
                monitorizacionCursos
                  .sheet("Instituciones")
                  .row(j + 1)
                  .cell(9)
                  .value(cambiosInstituciones[i]["email1"]);
                monitorizacionCursos
                  .sheet("Instituciones")
                  .row(j + 1)
                  .cell(10)
                  .value(cambiosInstituciones[i]["telefono1"]);
                monitorizacionCursos
                  .sheet("Instituciones")
                  .row(j + 1)
                  .cell(11)
                  .value(cambiosInstituciones[i]["contacto2"]);
                monitorizacionCursos
                  .sheet("Instituciones")
                  .row(j + 1)
                  .cell(12)
                  .value(cambiosInstituciones[i]["email2"]);
                monitorizacionCursos
                  .sheet("Instituciones")
                  .row(j + 1)
                  .cell(13)
                  .value(cambiosInstituciones[i]["telefono2"]);
                monitorizacionCursos
                  .sheet("Instituciones")
                  .row(j + 1)
                  .cell(14)
                  .value(cambiosInstituciones[i]["contacto3"]);
                monitorizacionCursos
                  .sheet("Instituciones")
                  .row(j + 1)
                  .cell(15)
                  .value(cambiosInstituciones[i]["email3"]);
                monitorizacionCursos
                  .sheet("Instituciones")
                  .row(j + 1)
                  .cell(16)
                  .value(cambiosInstituciones[i]["telefono3"]);
                monitorizacionCursos
                  .sheet("Instituciones")
                  .row(j + 1)
                  .cell(17)
                  .value(cambiosInstituciones[i]["contacto4"]);
                monitorizacionCursos
                  .sheet("Instituciones")
                  .row(j + 1)
                  .cell(18)
                  .value(cambiosInstituciones[i]["email4"]);
                monitorizacionCursos
                  .sheet("Instituciones")
                  .row(j + 1)
                  .cell(19)
                  .value(cambiosInstituciones[i]["telefono4"]);
                monitorizacionCursos
                  .sheet("Instituciones")
                  .row(j + 1)
                  .cell(20)
                  .value(cambiosInstituciones[i]["direccion"]);
                contadorModificacionInstituciones++;
                break;
              }
            }

            //NO ENCONTRADO
            if (!institucionEncontrada) {
              //Nueva Institucion:
              monitorizacionCursos
                .sheet("Instituciones")
                .row(filasInstituciones + contadorNuevasInstituciones + 1)
                .cell(1)
                .value(cambiosInstituciones[i]["cod_institucion"]);
              monitorizacionCursos
                .sheet("Instituciones")
                .row(filasInstituciones + contadorNuevasInstituciones + 1)
                .cell(2)
                .value(cambiosInstituciones[i]["institucion"]);
              monitorizacionCursos
                .sheet("Instituciones")
                .row(filasInstituciones + contadorNuevasInstituciones + 1)
                .cell(3)
                .value(cambiosInstituciones[i]["tipo"]);
              monitorizacionCursos
                .sheet("Instituciones")
                .row(filasInstituciones + contadorNuevasInstituciones + 1)
                .cell(4)
                .value(cambiosInstituciones[i]["cod__postal"]);
              monitorizacionCursos
                .sheet("Instituciones")
                .row(filasInstituciones + contadorNuevasInstituciones + 1)
                .cell(5)
                .value(cambiosInstituciones[i]["territorial"]);
              monitorizacionCursos
                .sheet("Instituciones")
                .row(filasInstituciones + contadorNuevasInstituciones + 1)
                .cell(6)
                .value(cambiosInstituciones[i]["ccaa_/_pais"]);
              monitorizacionCursos
                .sheet("Instituciones")
                .row(filasInstituciones + contadorNuevasInstituciones + 1)
                .cell(7)
                .value(cambiosInstituciones[i]["provincia"]);

              monitorizacionCursos
                .sheet("Instituciones")
                .row(filasInstituciones + contadorNuevasInstituciones + 1)
                .cell(8)
                .value(cambiosInstituciones[i]["contacto1"]);
              monitorizacionCursos
                .sheet("Instituciones")
                .row(filasInstituciones + contadorNuevasInstituciones + 1)
                .cell(9)
                .value(cambiosInstituciones[i]["email1"]);
              monitorizacionCursos
                .sheet("Instituciones")
                .row(filasInstituciones + contadorNuevasInstituciones + 1)
                .cell(10)
                .value(cambiosInstituciones[i]["telefono1"]);

              monitorizacionCursos
                .sheet("Instituciones")
                .row(filasInstituciones + contadorNuevasInstituciones + 1)
                .cell(11)
                .value(cambiosInstituciones[i]["contacto2"]);
              monitorizacionCursos
                .sheet("Instituciones")
                .row(filasInstituciones + contadorNuevasInstituciones + 1)
                .cell(12)
                .value(cambiosInstituciones[i]["email2"]);
              monitorizacionCursos
                .sheet("Instituciones")
                .row(filasInstituciones + contadorNuevasInstituciones + 1)
                .cell(13)
                .value(cambiosInstituciones[i]["telefono2"]);

              monitorizacionCursos
                .sheet("Instituciones")
                .row(filasInstituciones + contadorNuevasInstituciones + 1)
                .cell(14)
                .value(cambiosInstituciones[i]["contacto3"]);
              monitorizacionCursos
                .sheet("Instituciones")
                .row(filasInstituciones + contadorNuevasInstituciones + 1)
                .cell(15)
                .value(cambiosInstituciones[i]["email3"]);
              monitorizacionCursos
                .sheet("Instituciones")
                .row(filasInstituciones + contadorNuevasInstituciones + 1)
                .cell(16)
                .value(cambiosInstituciones[i]["telefono3"]);

              monitorizacionCursos
                .sheet("Instituciones")
                .row(filasInstituciones + contadorNuevasInstituciones + 1)
                .cell(17)
                .value(cambiosInstituciones[i]["contacto4"]);
              monitorizacionCursos
                .sheet("Instituciones")
                .row(filasInstituciones + contadorNuevasInstituciones + 1)
                .cell(18)
                .value(cambiosInstituciones[i]["email4"]);
              monitorizacionCursos
                .sheet("Instituciones")
                .row(filasInstituciones + contadorNuevasInstituciones + 1)
                .cell(19)
                .value(cambiosInstituciones[i]["telefono4"]);
              monitorizacionCursos
                .sheet("Instituciones")
                .row(filasInstituciones + contadorNuevasInstituciones + 1)
                .cell(20)
                .value(cambiosInstituciones[i]["direccion"]);

              //Actualizar Contador:
              contadorNuevasInstituciones++;
            }
          } //Fin de iteracion de cambios Formadores

          console.log("Num Cambios Cursos:" + cambios.length);
          console.log("Modificaciones Cursos:" + contadorModificacion);
          console.log("Nuevos Cursos:" + contadorNuevas);

          console.log("Num Cambios Formadores:" + cambiosFormadores.length);
          console.log(
            "Modificaciones Formadores:" + contadorModificacionFormadores,
          );
          console.log("Nuevos Formadores:" + contadorNuevosFormadores);

          console.log(
            "Num Cambios Instituciones:" + cambiosInstituciones.length,
          );
          console.log(
            "Modificaciones Instituciones:" + contadorModificacionInstituciones,
          );
          console.log("Nuevas Instituciones:" + contadorNuevasInstituciones);

          //Nuevas filas:
          filasCursos = filasCursos + contadorNuevas + 1;
          filasInstituciones =
            filasInstituciones + contadorNuevasInstituciones + 1;
          filasFormadores = filasFormadores + contadorNuevosFormadores + 1;
          filasFormadoresCursos = punteroRegistroFormador + 1;

          //CREAR OBJETOS JSON:
          var jsonCursos = [];
          for (var i = 1; i < filasCursos; i++) {
            jsonCursos.push({
              cod_curso: monitorizacionCursos
                .sheet("Cursos")
                .row(i + 1)
                .cell(1)
                .value(),
              cod_grupo: monitorizacionCursos
                .sheet("Cursos")
                .row(i + 1)
                .cell(2)
                .value(),
              cod__postal: monitorizacionCursos
                .sheet("Cursos")
                .row(i + 1)
                .cell(3)
                .value(),
              territorial: monitorizacionCursos
                .sheet("Cursos")
                .row(i + 1)
                .cell(4)
                .value(),
              "ccaa_/_pais": monitorizacionCursos
                .sheet("Cursos")
                .row(i + 1)
                .cell(5)
                .value(),
              curso: monitorizacionCursos
                .sheet("Cursos")
                .row(i + 1)
                .cell(6)
                .value(),
              "sesi\u00f3n": monitorizacionCursos
                .sheet("Cursos")
                .row(i + 1)
                .cell(7)
                .value(),
              fecha: monitorizacionCursos
                .sheet("Cursos")
                .row(i + 1)
                .cell(8)
                .value(),
              hora_inicio: monitorizacionCursos
                .sheet("Cursos")
                .row(i + 1)
                .cell(9)
                .value(),
              hora_fin: monitorizacionCursos
                .sheet("Cursos")
                .row(i + 1)
                .cell(10)
                .value(),
              "duraci\u00f3n": monitorizacionCursos
                .sheet("Cursos")
                .row(i + 1)
                .cell(11)
                .value(),
              "instituci\u00f3n": monitorizacionCursos
                .sheet("Cursos")
                .row(i + 1)
                .cell(12)
                .value(),
              colectivo: monitorizacionCursos
                .sheet("Cursos")
                .row(i + 1)
                .cell(13)
                .value(),
              grupo: monitorizacionCursos
                .sheet("Cursos")
                .row(i + 1)
                .cell(14)
                .value(),
              "n\u00baasistentes": monitorizacionCursos
                .sheet("Cursos")
                .row(i + 1)
                .cell(15)
                .value(),
              modalidad: monitorizacionCursos
                .sheet("Cursos")
                .row(i + 1)
                .cell(16)
                .value(),
              estado: monitorizacionCursos
                .sheet("Cursos")
                .row(i + 1)
                .cell(17)
                .value(),
              material: monitorizacionCursos
                .sheet("Cursos")
                .row(i + 1)
                .cell(18)
                .value(),
              "valoraci\u00f3n": monitorizacionCursos
                .sheet("Cursos")
                .row(i + 1)
                .cell(19)
                .value(),
              observaciones: monitorizacionCursos
                .sheet("Cursos")
                .row(i + 1)
                .cell(20)
                .value(),
            });
          }

          var jsonFormadores = [];
          for (var i = 1; i < filasFormadores; i++) {
            jsonFormadores.push({
              cod__formador: monitorizacionCursos
                .sheet("Formadores")
                .row(i + 1)
                .cell(1)
                .value(),
              nombre: monitorizacionCursos
                .sheet("Formadores")
                .row(i + 1)
                .cell(2)
                .value(),
              email: monitorizacionCursos
                .sheet("Formadores")
                .row(i + 1)
                .cell(3)
                .value(),
              telefono: monitorizacionCursos
                .sheet("Formadores")
                .row(i + 1)
                .cell(4)
                .value(),
              territorial: monitorizacionCursos
                .sheet("Formadores")
                .row(i + 1)
                .cell(5)
                .value(),
              ccaa: monitorizacionCursos
                .sheet("Formadores")
                .row(i + 1)
                .cell(6)
                .value(),
              provincia: monitorizacionCursos
                .sheet("Formadores")
                .row(i + 1)
                .cell(7)
                .value(),
              fecha: monitorizacionCursos
                .sheet("Formadores")
                .row(i + 1)
                .cell(8)
                .value(),
              certificado: monitorizacionCursos
                .sheet("Formadores")
                .row(i + 1)
                .cell(9)
                .value(),
              confidencialidad: monitorizacionCursos
                .sheet("Formadores")
                .row(i + 1)
                .cell(10)
                .value(),
              consentimiento: monitorizacionCursos
                .sheet("Formadores")
                .row(i + 1)
                .cell(11)
                .value(),
              estado: monitorizacionCursos
                .sheet("Formadores")
                .row(i + 1)
                .cell(12)
                .value(),
            });
          }

          var jsonInstituciones = [];
          for (var i = 1; i < filasInstituciones; i++) {
            jsonInstituciones.push({
              cod_institucion: monitorizacionCursos
                .sheet("Instituciones")
                .row(i + 1)
                .cell(1)
                .value(),
              institucion: monitorizacionCursos
                .sheet("Instituciones")
                .row(i + 1)
                .cell(2)
                .value(),
              tipo: monitorizacionCursos
                .sheet("Instituciones")
                .row(i + 1)
                .cell(3)
                .value(),
              cod__postal: monitorizacionCursos
                .sheet("Instituciones")
                .row(i + 1)
                .cell(4)
                .value(),
              territorial: monitorizacionCursos
                .sheet("Instituciones")
                .row(i + 1)
                .cell(5)
                .value(),
              ccaa: monitorizacionCursos
                .sheet("Instituciones")
                .row(i + 1)
                .cell(6)
                .value(),
              provincia: monitorizacionCursos
                .sheet("Instituciones")
                .row(i + 1)
                .cell(7)
                .value(),

              contacto1: monitorizacionCursos
                .sheet("Instituciones")
                .row(i + 1)
                .cell(8)
                .value(),
              email1: monitorizacionCursos
                .sheet("Instituciones")
                .row(i + 1)
                .cell(9)
                .value(),
              telefono1: monitorizacionCursos
                .sheet("Instituciones")
                .row(i + 1)
                .cell(10)
                .value(),

              contacto2: monitorizacionCursos
                .sheet("Instituciones")
                .row(i + 1)
                .cell(11)
                .value(),
              email2: monitorizacionCursos
                .sheet("Instituciones")
                .row(i + 1)
                .cell(12)
                .value(),
              telefono2: monitorizacionCursos
                .sheet("Instituciones")
                .row(i + 1)
                .cell(13)
                .value(),

              contacto3: monitorizacionCursos
                .sheet("Instituciones")
                .row(i + 1)
                .cell(14)
                .value(),
              email3: monitorizacionCursos
                .sheet("Instituciones")
                .row(i + 1)
                .cell(15)
                .value(),
              telefono3: monitorizacionCursos
                .sheet("Instituciones")
                .row(i + 1)
                .cell(16)
                .value(),

              contacto4: monitorizacionCursos
                .sheet("Instituciones")
                .row(i + 1)
                .cell(17)
                .value(),
              email4: monitorizacionCursos
                .sheet("Instituciones")
                .row(i + 1)
                .cell(18)
                .value(),
              telefono4: monitorizacionCursos
                .sheet("Instituciones")
                .row(i + 1)
                .cell(19)
                .value(),
              direccion: monitorizacionCursos
                .sheet("Instituciones")
                .row(i + 1)
                .cell(20)
                .value(),
            });
          }

          var jsonFormadorCurso = [];
          for (var i = 1; i < filasFormadoresCursos; i++) {
            jsonFormadorCurso.push({
              cod__curso: monitorizacionCursos
                .sheet("Formador-Curso")
                .row(i + 1)
                .cell(1)
                .value(),
              cod__formador: monitorizacionCursos
                .sheet("Formador-Curso")
                .row(i + 1)
                .cell(2)
                .value(),
            });
          }

          //Eliminar Filas Vacias:
          for (var i = 0; i < jsonCursos.length; i++) {
            if (!jsonCursos[i].cod_curso) {
              jsonCursos.splice(i, 1);
              i--;
            }
          }

          //Eliminar Filas Vacias Formadores:
          for (var i = 0; i < jsonFormadores.length; i++) {
            if (!jsonFormadores[i].cod__formador) {
              jsonFormadores.splice(i, 1);
              i--;
            }
          }

          //Eliminar Filas Vacias Instituciones:
          for (var i = 0; i < jsonInstituciones.length; i++) {
            if (!jsonInstituciones[i].cod_institucion) {
              jsonInstituciones.splice(i, 1);
              i--;
            }
          }

          //Eliminar Filas Vacias Curso-Formador:
          for (var i = 0; i < jsonFormadorCurso.length; i++) {
            if (!jsonFormadorCurso[i].cod__curso) {
              jsonFormadorCurso.splice(i, 1);
              i--;
            }
          }

          //Guardar Archivos JSON:
          jsonCursos = JSON.stringify(jsonCursos);
          jsonFormadores = JSON.stringify(jsonFormadores);
          jsonInstituciones = JSON.stringify(jsonInstituciones);
          jsonFormadorCurso = JSON.stringify(jsonFormadorCurso);
          let jsonProvincia = JSON.stringify(codigosProvincia);

          try {
            fs.writeFileSync(
              path.normalize(path.join(pathRaiz, "db/cursos.json")),
              jsonCursos,
            );
            fs.writeFileSync(
              path.normalize(path.join(pathRaiz, "db/formadores.json")),
              jsonFormadores,
            );
            fs.writeFileSync(
              path.normalize(path.join(pathRaiz, "db/instituciones.json")),
              jsonInstituciones,
            );
            fs.writeFileSync(
              path.normalize(path.join(pathRaiz, "db/formador-curso.json")),
              jsonFormadorCurso,
            );
            fs.writeFileSync(
              path.normalize(path.join(pathRaiz, "db/provincia.json")),
              jsonProvincia,
            );
          } catch (err) {
            console.log("Se ha producido un error interno: ");
            console.log(err);
            var tituloError =
              "Se ha producido un error guardando los archivos JSON. ";
            resolve(false);
          }

          //Fin de procesamiento:
          console.log("Escribiendo archivo...");
          console.log("Path: " + path.normalize(pathMonitorizacionCursos));

          monitorizacionCursos
            .toFileAsync(path.normalize(pathMonitorizacionCursos))
            .then(() => {
              console.log("Fin del procesamiento");
              //console.log(monitorizacionCursos)

              resolve(true);
            })
            .catch((err) => {
              console.log("Se ha producido un error interno: ");
              console.log(err);
              var tituloError =
                "Se ha producido un error escribiendo el archivo: " +
                path.normalize(pathMonitorizacionCursos);
              resolve(false);
            });
        });
    });
  }
} //Fin Procesos Asesoria

module.exports = ProcesosAsesoria;
