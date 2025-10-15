const path = require("path");
const fs = require("fs");
const readline = require("readline");
const axios = require("axios");
const moment = require("moment");
const XlsxPopulate = require("xlsx-populate");
const Datastore = require("nedb");
const _ = require("lodash");
const { DateTime } = require("luxon");

const { ipcRenderer } = require("electron");
const puppeteer = require("puppeteer");
const generatePDF = require("./pdf-fie");
const generarDesdePlantilla = require("./emails-fie");

class ProcesosFie {
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

  async fIE(argumentos) {
    return new Promise((resolve) => {
      console.log("Procesamiento de FIE...");

      var pathArchivoFIE = argumentos.formularioControl[0];
      var pathArchivoEmpresas = argumentos.formularioControl[1];
      var pathArchivoEnfermedad = argumentos.formularioControl[2];
      var pathArchivoAccidentes = argumentos.formularioControl[3];
      var archivoFIE = {};

      var pathSalidaPDFBajas = path.join(
        path.normalize(argumentos.formularioControl[4]),
        "Fie-Procesado",
        "Bajas",
      );
      var pathSalidaPDFAltas = path.join(
        path.normalize(argumentos.formularioControl[4]),
        "Fie-Procesado",
        "Altas",
      );
      var pathSalidaPDFConfirmacion = path.join(
        path.normalize(argumentos.formularioControl[4]),
        "Fie-Procesado",
        "Confirmacion",
      );

      // Carpetas de emails:
      var pathSalidaPDFBajasCorreos = path.join(
        path.normalize(argumentos.formularioControl[4]),
        "Fie-Procesado",
        "Bajas-Correos",
      );
      var pathSalidaPDFAltasCorreos = path.join(
        path.normalize(argumentos.formularioControl[4]),
        "Fie-Procesado",
        "Altas-Correos",
      );
      var pathSalidaPDFConfirmacionCorreos = path.join(
        path.normalize(argumentos.formularioControl[4]),
        "Fie-Procesado",
        "Confirmacion-Correos",
      );

      var pathSalidaExcel = path.join(
        path.normalize(argumentos.formularioControl[4]),
        "Fie-Procesado",
      );
      var pathSalida = path.join(
        path.normalize(argumentos.formularioControl[4]),
        "Fie-Procesado",
        "Resultados",
      );

      // Verificar si la carpeta "Resultados" existe y crearla si no
      if (!fs.existsSync(pathSalida)) {
        fs.mkdirSync(pathSalida, { recursive: true });
        console.log(`Carpeta creada: ${pathSalida}`);
      } else {
        console.log(`La carpeta ya existe: ${pathSalida}`);
      }

      if (!fs.existsSync(pathSalidaPDFConfirmacion)) {
        fs.mkdirSync(pathSalidaPDFConfirmacion, { recursive: true });
        console.log(`Carpeta creada: ${pathSalidaPDFConfirmacion}`);
      } else {
        console.log(`La carpeta ya existe: ${pathSalidaPDFConfirmacion}`);
      }

      if (!fs.existsSync(pathSalidaPDFAltas)) {
        fs.mkdirSync(pathSalidaPDFAltas, { recursive: true });
        console.log(`Carpeta creada: ${pathSalidaPDFAltas}`);
      } else {
        console.log(`La carpeta ya existe: ${pathSalidaPDFAltas}`);
      }

      if (!fs.existsSync(pathSalidaPDFBajas)) {
        fs.mkdirSync(pathSalidaPDFBajas, { recursive: true });
        console.log(`Carpeta creada: ${pathSalidaPDFBajas}`);
      } else {
        console.log(`La carpeta ya existe: ${pathSalidaPDFBajas}`);
      }

      if (!fs.existsSync(pathSalidaPDFConfirmacionCorreos)) {
        fs.mkdirSync(pathSalidaPDFConfirmacionCorreos, { recursive: true });
        console.log(`Carpeta creada: ${pathSalidaPDFConfirmacionCorreos}`);
      } else {
        console.log(
          `La carpeta ya existe: ${pathSalidaPDFConfirmacionCorreos}`,
        );
      }

      if (!fs.existsSync(pathSalidaPDFAltasCorreos)) {
        fs.mkdirSync(pathSalidaPDFAltasCorreos, { recursive: true });
        console.log(`Carpeta creada: ${pathSalidaPDFAltasCorreos}`);
      } else {
        console.log(`La carpeta ya existe: ${pathSalidaPDFAltasCorreos}`);
      }

      if (!fs.existsSync(pathSalidaPDFBajasCorreos)) {
        fs.mkdirSync(pathSalidaPDFBajasCorreos, { recursive: true });
        console.log(`Carpeta creada: ${pathSalidaPDFBajasCorreos}`);
      } else {
        console.log(`La carpeta ya existe: ${pathSalidaPDFBajasCorreos}`);
      }

      try {
        //LECTURA EMPRESAS:
        var datosEmpresas = [];
        XlsxPopulate.fromFileAsync(path.normalize(pathArchivoEmpresas))
          .then(async (archivoEmpresas) => {
            console.log("Archivo Cargado: Empresas");
            datosEmpresas = extraccionExcel(archivoEmpresas, 0, 1);
            //resolve(true);
          })
          .then(() => {
            console.log("EMPRESAS:");
            console.log(datosEmpresas[0]);

            XlsxPopulate.fromFileAsync(path.normalize(pathArchivoFIE))
              .then(async (workbook) => {
                console.log("Archivo Cargado: FIE");
                archivoFIE = workbook;

                const datosIncapacidad = extraccionExcel(archivoFIE, 0, 3);
                const partesConfirmacion = extraccionExcel(archivoFIE, 4, 3);

                console.log("Datos incapacidad:");
                console.log(datosIncapacidad[0]);

                console.log("Partes confirmacion:");
                console.log(partesConfirmacion[0]);

                // PASO 1: IDENTIFICACION.
                const altas = [];
                const bajas = [];
                const confirmacion = [];

                var partesDetectados = [];
                //Detectamos si los distintos casos:
                for (var i = 0; i < datosIncapacidad.length; i++) {
                  if (datosIncapacidad[i].fechaFinIt) {
                    altas.push(datosIncapacidad[i]);
                  }

                  //Buscamos parte de confirmacion:
                  partesDetectados = [];
                  for (var j = 0; j < partesConfirmacion.length; j++) {
                    if (partesConfirmacion[j].naf == datosIncapacidad[i].naf) {
                      partesDetectados.push(partesConfirmacion[j]);
                    }
                  }
                  if (partesDetectados.length > 0) {
                    confirmacion.push(datosIncapacidad[i]);
                    confirmacion[confirmacion.length - 1].partesConfirmacion =
                      Object.assign([], partesDetectados);
                  } else {
                    bajas.push(datosIncapacidad[i]);
                  }
                }

                //Asignación de empresas:
                var empresa = {};
                for (var i = 0; i < altas.length; i++) {
                  empresa = {};
                  empresa = datosEmpresas.find(
                    (e) => e.empresa === altas[i].empresa,
                  );
                  if (!empresa) {
                    console.log(
                      "Altas: No se ha encontrado empresa para: ",
                      altas[i].empresa,
                    );
                  }
                  altas[i].expedienteEmpresa = empresa?.codigo || "";
                  altas[i].emailsEmpresa = empresa?.email || "";
                }
                for (var i = 0; i < bajas.length; i++) {
                  empresa = {};
                  empresa = datosEmpresas.find(
                    (e) => e.empresa === bajas[i].empresa,
                  );
                  if (!empresa) {
                    console.log(
                      "Bajas: No se ha encontrado empresa para: ",
                      bajas[i].empresa,
                    );
                  }
                  bajas[i].expedienteEmpresa = empresa?.codigo || "";
                  bajas[i].emailsEmpresa = empresa?.email || "";
                }
                for (var i = 0; i < confirmacion.length; i++) {
                  empresa = {};
                  empresa = datosEmpresas.find(
                    (e) => e.empresa === confirmacion[i].empresa,
                  );
                  if (!empresa) {
                    console.log(
                      "Confirmacion: No se ha encontrado empresa para: ",
                      confirmacion[i].empresa,
                    );
                  }
                  confirmacion[i].expedienteEmpresa = empresa?.codigo || "";
                  confirmacion[i].emailsEmpresa = empresa?.email || "";
                }

                console.log("ALTAS:");
                console.log(altas[0]);

                console.log("CONFIRMACION:");
                console.log(confirmacion[0]);

                console.log("BAJAS:");
                console.log(bajas[0]);

                //PASO 2: GENERACION DE JUSTIFICANTES:

                // Generación
                /*
            const tasks = [];
            for (const r of bajas)
              tasks.push(generatePDF(r, "BAJAS", pathSalidaPDFBajas));
            for (const r of altas)
              tasks.push(generatePDF(r, "ALTAS", pathSalidaPDFAltas));
            for (const r of confirmacion)
              tasks.push(
                generatePDF(r, "CONFIRMACIONES", pathSalidaPDFConfirmacion),
              );

            const generated = await Promise.all(tasks);
            console.log("PDFs generados:");
            generated.forEach((f) => console.log(" -", f));
                        */

                //PASO 3: GENERACION DE CORREOS:
                const toDefault = [
                  {
                    name: "Administración",
                    address: "administracion@tuempresa.com",
                  },
                ];

                const results = [];
                //const altasTest = [altas[0]];
                for (const r of altas) {
                  const file = await generarDesdePlantilla(
                    r,
                    "ALTAS",
                    pathSalidaPDFAltasCorreos,
                    {
                      to: toDefault,
                    },
                  );
                  results.push(file);
                }

                //Correos Bajas:
                for (const r of bajas) {
                  const file = await generarDesdePlantilla(
                    r,
                    "BAJAS",
                    pathSalidaPDFBajasCorreos,
                    {
                      to: toDefault,
                    },
                  );
                  results.push(file);
                }

                //Correos Confirmacion:
                for (const r of confirmacion) {
                  const file = await generarDesdePlantilla(
                    r,
                    "CONFIRMACION",
                    pathSalidaPDFConfirmacionCorreos,
                    {
                      to: toDefault,
                    },
                  );
                  results.push(file);
                }
                //ESCRITURA XLSX:
                console.log("Escribiendo archivo...");
                console.log("Path: " + path.normalize(pathSalidaExcel));

                //resolve(true);
              })
              .then(() => {
                //RESCRITURA DE ENFERMEDADES
                XlsxPopulate.fromFileAsync(
                  path.normalize(pathArchivoEnfermedad),
                )
                  .then(async (archivoEnfermedad) => {
                    console.log("Archivo Cargado: Enfermedad");

                    //const datosIncapacidad = extraccionExcel(archivoFIE, 0, 3);

                    //ESCRITURA XLSX:
                    console.log("Escribiendo archivo...");
                    console.log("Path: " + path.normalize(pathSalidaExcel));

                    archivoEnfermedad
                      .toFileAsync(
                        path.normalize(
                          path.join(
                            pathSalidaExcel,
                            "01 Enfermedad 2025 -Procesado.xlsx",
                          ),
                        ),
                      )
                      .then(() => {
                        //console.log(archivoFIE)
                        //resolve(true);
                      })
                      .catch((err) => {
                        console.log("Se ha producido un error interno: ");
                        console.log(err);
                        var tituloError =
                          "Se ha producido un error escribiendo el archivo: " +
                          path.normalize(pathSalidaExcel);
                        resolve(false);
                      });

                    //resolve(true);
                  })
                  .then(() => {
                    //RESCRITURA DE ACCIDENTES:
                    XlsxPopulate.fromFileAsync(
                      path.normalize(pathArchivoAccidentes),
                    ).then(async (archivoAccidentes) => {
                      console.log("Archivo Cargado: Accidentes");

                      //const datosIncapacidad = extraccionExcel(archivoAccidentes, 0, 3);

                      //ESCRITURA XLSX:
                      console.log("Escribiendo archivo...");
                      console.log("Path: " + path.normalize(pathSalidaExcel));

                      archivoAccidentes
                        .toFileAsync(
                          path.normalize(
                            path.join(
                              pathSalidaExcel,
                              "02 Accidentes 2025 -Procesado.xlsx",
                            ),
                          ),
                        )
                        .then(() => {
                          console.log("Fin del procesamiento");
                          //console.log(archivoFIE)

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
                    });
                  });
              });
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
} //Fin Procesos Fie

function extraccionExcel(workbook, sheet, filaCabecera, columnaCabecera = 1) {
  const columnas = workbook.sheet(sheet).usedRange()._numColumns;
  const filas = workbook.sheet(sheet).usedRange()._numRows;

  //Identificacion de cabeceras:
  const cabeceras = [];
  var cabecera = "";
  for (var i = columnaCabecera; i <= columnas; i++) {
    valor = workbook.sheet(sheet).cell(filaCabecera, i).value();
    if (valor) {
      cabecera = camelize(workbook.sheet(sheet).cell(filaCabecera, i).value());
      cabeceras.push(cabecera);
    } else {
      cabeceras.push(null);
    }
  }

  console.log("Cabeceras (Sheet [" + sheet + "]): ", cabeceras);

  const registros = [];
  var objetoRegistro = {};

  //Asignación de valores:
  for (var i = filaCabecera + 1; i <= filas; i++) {
    objetoRegistro = {};
    for (var j = columnaCabecera; j <= columnas; j++) {
      if (cabeceras[j - 1]) {
        objetoRegistro[cabeceras[j - 1]] = workbook
          .sheet(sheet)
          .cell(i, j)
          .value();
      }
    }

    registros.push(Object.assign({}, objetoRegistro));
  }
  return registros;
}

function camelize(str) {
  str = String(str);
  return str
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/(?:^\w|[A-Z]|\b\w)/g, function (word, index) {
      return index === 0 ? word.toLowerCase() : word.toUpperCase();
    })
    .replace(/\s+/g, "");
}

module.exports = ProcesosFie;
