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
const generarEmailFieDesdePlantilla = require("./emails-fie");

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

  getCurrentDateString() {
    const now = new Date();
    const day = String(now.getDate()).padStart(2, "0");
    const month = String(now.getMonth() + 1).padStart(2, "0"); // los meses van de 0 a 11
    const year = now.getFullYear();
    return `${day}${month}${year}`;
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
        "Fie-Procesado (" + this.getCurrentDateString() + ")",
        "PDFs-Generados",
      );
      var pathSalidaPDFAltas = path.join(
        path.normalize(argumentos.formularioControl[4]),
        "Fie-Procesado (" + this.getCurrentDateString() + ")",
        "PDFs-Generados",
      );
      var pathSalidaPDFConfirmacion = path.join(
        path.normalize(argumentos.formularioControl[4]),
        "Fie-Procesado (" + this.getCurrentDateString() + ")",
        "PDFs-Generados",
      );

      // Carpetas de emails:
      var pathSalidaPDFBajasCorreos = path.join(
        path.normalize(argumentos.formularioControl[4]),
        "Fie-Procesado (" + this.getCurrentDateString() + ")",
        "Bajas-Correos",
      );
      var pathSalidaPDFAltasCorreos = path.join(
        path.normalize(argumentos.formularioControl[4]),
        "Fie-Procesado (" + this.getCurrentDateString() + ")",
        "Altas-Correos",
      );
      var pathSalidaPDFConfirmacionCorreos = path.join(
        path.normalize(argumentos.formularioControl[4]),
        "Fie-Procesado (" + this.getCurrentDateString() + ")",
        "Confirmacion-Correos",
      );

      var pathSalidaExcel = path.join(
        path.normalize(argumentos.formularioControl[4]),
        "Fie-Procesado (" + this.getCurrentDateString() + ")",
      );
      var pathSalida = path.join(
        path.normalize(argumentos.formularioControl[4]),
        "Fie-Procesado (" + this.getCurrentDateString() + ")",
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
        const altas = [];
        const bajas = [];
        const confirmacion = [];

        XlsxPopulate.fromFileAsync(path.normalize(pathArchivoEmpresas))
          .then(async (archivoEmpresas) => {
            console.log("Archivo Cargado: Empresas");
            datosEmpresas = extraccionExcel(archivoEmpresas, 0); //1
            //resolve(true);
          })
          .then(() => {
            console.log("EMPRESAS:");
            console.log(datosEmpresas[0]);

            XlsxPopulate.fromFileAsync(path.normalize(pathArchivoFIE))
              .then(async (workbook) => {
                console.log("Archivo Cargado: FIE");
                archivoFIE = workbook;

                const datosIncapacidad = extraccionExcel(archivoFIE, 0); //3
                const partesConfirmacion = extraccionExcel(archivoFIE, 4); //3

                console.log("Datos incapacidad:");
                console.log(datosIncapacidad[0]);

                console.log("Partes confirmacion:");
                console.log(partesConfirmacion[0]);

                // -- EXTRACCION DNI
                for (var i = 0; i < datosIncapacidad.length; i++) {
                  if (
                    datosIncapacidad[i].nif &&
                    datosIncapacidad[i].nif.trim() != "" &&
                    datosIncapacidad[i].nif.length >= 9
                  ) {
                    datosIncapacidad[i].dni = datosIncapacidad[i].nif
                      .trim()
                      .slice(-9, -1);
                  } else {
                    datosIncapacidad[i].dni = null;
                  }
                }

                // -- IDENTIFICACION (ALTAS, BAJAS, CONFIRMACION)
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
                  if (altas[i].expte) {
                    empresa = datosEmpresas.find(
                      (e) => Number(e.codigo) === Number(altas[i].expte),
                    );
                  } else {
                    empresa = datosEmpresas.find(
                      (e) => e.empresa === altas[i].empresa,
                    );
                  }
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

                  if (bajas[i].expte) {
                    empresa = datosEmpresas.find(
                      (e) => Number(e.codigo) === Number(bajas[i].expte),
                    );
                  } else {
                    empresa = datosEmpresas.find(
                      (e) => e.empresa === bajas[i].empresa,
                    );
                  }
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

                  if (confirmacion[i].expte) {
                    empresa = datosEmpresas.find(
                      (e) => Number(e.codigo) === Number(confirmacion[i].expte),
                    );
                  } else {
                    empresa = datosEmpresas.find(
                      (e) => e.empresa === confirmacion[i].empresa,
                    );
                  }
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

                //PASO 3: GENERACION DE CORREOS:
                const results = [];
                //const altasTest = [altas[0]];
                for (const r of altas) {
                  const file = await generarEmailFieDesdePlantilla(
                    r,
                    "ALTAS",
                    pathSalidaPDFAltasCorreos,
                    {
                      to: r.emailsEmpresa?.split(";") ?? [],
                    },
                  );
                  results.push(file);
                }

                //Correos Bajas:
                for (const r of bajas) {
                  const file = await generarEmailFieDesdePlantilla(
                    r,
                    "BAJAS",
                    pathSalidaPDFBajasCorreos,
                    {
                      to: r.emailsEmpresa?.split(";") ?? [],
                    },
                  );
                  results.push(file);
                }

                //Correos Confirmacion:
                for (const r of confirmacion) {
                  const file = await generarEmailFieDesdePlantilla(
                    r,
                    "CONFIRMACION",
                    pathSalidaPDFConfirmacionCorreos,
                    {
                      to: r.emailsEmpresa?.split(";") ?? [],
                    },
                  );
                  results.push(file);
                }
              })
              .then(() => {
                //RESCRITURA DE ENFERMEDADES
                XlsxPopulate.fromFileAsync(
                  path.normalize(pathArchivoEnfermedad),
                )
                  .then(async (archivoEnfermedad) => {
                    console.log("Archivo Cargado: Enfermedad");
                    const hojas = archivoEnfermedad.sheets();
                    const filas = archivoEnfermedad
                      .sheet(hojas.length - 1)
                      .usedRange()._numRows;
                    const columnas = archivoEnfermedad
                      .sheet(hojas.length - 1)
                      .usedRange()._numColumns;

                    const ultimaHoja = hojas.length - 1;
                    const nuevaHoja = archivoEnfermedad.addSheet(
                      "Procesamiento automatico",
                    );

                    //Identificacion de cabeceras:
                    const { cabeceras, columnaCabecera, filaCabecera } =
                      deteccionCabeceras(archivoEnfermedad, ultimaHoja);

                    //Identificar ultima fila:
                    var filaVacia = 0;
                    var flagVacia = true;
                    for (var i = 3; i < filas; i++) {
                      flagVacia = true;
                      if (!nuevaHoja.cell(i, 1).value()) {
                        for (var j = 1; j < columnas; j++) {
                          if (nuevaHoja.cell(i, j).value()) {
                            nuevaHoja.cell(i, j).value();
                            flagVacia = false;
                          }
                        }
                        if (flagVacia) {
                          filaVacia = i;
                          break;
                        }
                      }
                    }

                    //Creacion de objeto para enfermedades:
                    const enfermedades = [];
                    for (var i = 0; i < bajas.length; i++) {
                      if (bajas[i].contingencia) {
                        if (
                          bajas[i].contingencia[0] == 1 ||
                          bajas[i].contingencia[0] == 2
                        ) {
                          enfermedades.push(bajas[i]);
                          enfermedades[enfermedades.length - 1].tipo = "BAJA";
                        }
                      }
                    }

                    for (var i = 0; i < altas.length; i++) {
                      if (altas[i].contingencia) {
                        if (
                          altas[i].contingencia[0] == 1 ||
                          altas[i].contingencia[0] == 2
                        ) {
                          enfermedades.push(altas[i]);
                          enfermedades[enfermedades.length - 1].tipo = "ALTA";
                        }
                      }
                    }
                    for (var i = 0; i < confirmacion.length; i++) {
                      if (confirmacion[i].contingencia) {
                        if (
                          confirmacion[i].contingencia[0] == 1 ||
                          confirmacion[i].contingencia[0] == 2
                        ) {
                          enfermedades.push(confirmacion[i]);
                          enfermedades[enfermedades.length - 1].tipo =
                            "CONFIRMACION";
                        }
                      }
                    }

                    console.log("ENFERMEDADES [0]:");
                    console.log(enfermedades[0]);

                    //Sobrescribir fila Vacia por ser hoja nueva:
                    filaVacia = 2;

                    var columnasClave = {
                      columnaExpediente: 0,
                      columnaNombre: 0,
                      columnaNaf: 0,
                      columnaDias180: 0,
                      columnaFechaBaja: 0,
                      columnaProximaRev: 0,
                      columnaFechaAlta: 0,
                      columnaDias3: 0,
                      columnaDias5: 0,
                      columnaDias12: 0,
                      columnaDiasResto: 0,
                      columnaDiasTotal: 0,
                      columnaAnotacion: columnas || 23,
                    };

                    //ESCRITURA DE CABECERAS:
                    var columnaMaxima = 0;
                    for (var i = 0; i < cabeceras.length; i++) {
                      switch (cabeceras[i].toLowerCase().trim()) {
                        case "exp":
                          columnasClave.columnaExpediente = i + columnaCabecera;
                          nuevaHoja
                            .cell(filaVacia - 1, columnaCabecera + i)
                            .value(cabeceras[i]);
                          if (columnaMaxima < columnaCabecera + i) {
                            columnaMaxima = columnaCabecera + i;
                          }
                          break;
                        case "apellidos y nombre":
                          columnasClave.columnaNombre = i + columnaCabecera;
                          nuevaHoja
                            .cell(filaVacia - 1, columnaCabecera + i)
                            .value(cabeceras[i]);
                          if (columnaMaxima < columnaCabecera + i) {
                            columnaMaxima = columnaCabecera + i;
                          }
                          break;
                        case "n.a.f.-c.c.c.":
                          columnasClave.columnaNaf = i + columnaCabecera;
                          nuevaHoja
                            .cell(filaVacia - 1, columnaCabecera + i)
                            .value(cabeceras[i]);
                          if (columnaMaxima < columnaCabecera + i) {
                            columnaMaxima = columnaCabecera + i;
                          }
                          break;
                        case "180 dias":
                          columnasClave.columnaDias180 = i + columnaCabecera;
                          nuevaHoja
                            .cell(filaVacia - 1, columnaCabecera + i)
                            .value(cabeceras[i]);
                          if (columnaMaxima < columnaCabecera + i) {
                            columnaMaxima = columnaCabecera + i;
                          }
                          break;
                        case "f.  baja":
                          columnasClave.columnaFechaBaja = i + columnaCabecera;
                          nuevaHoja
                            .cell(filaVacia - 1, columnaCabecera + i)
                            .value(cabeceras[i]);
                          if (columnaMaxima < columnaCabecera + i) {
                            columnaMaxima = columnaCabecera + i;
                          }
                          break;
                        case "próxima revision":
                          columnasClave.columnaProximaRev = i + columnaCabecera;
                          nuevaHoja
                            .cell(filaVacia - 1, columnaCabecera + i)
                            .value(cabeceras[i]);
                          if (columnaMaxima < columnaCabecera + i) {
                            columnaMaxima = columnaCabecera + i;
                          }
                          break;
                        case "f. alta":
                          columnasClave.columnaFechaAlta = i + columnaCabecera;
                          nuevaHoja
                            .cell(filaVacia - 1, columnaCabecera + i)
                            .value(cabeceras[i]);
                          if (columnaMaxima < columnaCabecera + i) {
                            columnaMaxima = columnaCabecera + i;
                          }
                          break;
                        case "dias  50 %(3)":
                          columnasClave.columnaDias3 = i + columnaCabecera;
                          nuevaHoja
                            .cell(filaVacia - 1, columnaCabecera + i)
                            .value(cabeceras[i]);
                          if (columnaMaxima < columnaCabecera + i) {
                            columnaMaxima = columnaCabecera + i;
                          }
                          break;
                        case "dias 60%(12)":
                          columnasClave.columnaDias12 = i + columnaCabecera;
                          nuevaHoja
                            .cell(filaVacia - 1, columnaCabecera + i)
                            .value(cabeceras[i]);
                          if (columnaMaxima < columnaCabecera + i) {
                            columnaMaxima = columnaCabecera + i;
                          }
                          break;
                        case "dias 60%(5)":
                          columnasClave.columnaDias5 = i + columnaCabecera;
                          nuevaHoja
                            .cell(filaVacia - 1, columnaCabecera + i)
                            .value(cabeceras[i]);
                          if (columnaMaxima < columnaCabecera + i) {
                            columnaMaxima = columnaCabecera + i;
                          }
                          break;
                        case "dias 75% (resto)":
                          columnasClave.columnaDiasResto = i + columnaCabecera;
                          nuevaHoja
                            .cell(filaVacia - 1, columnaCabecera + i)
                            .value(cabeceras[i]);
                          if (columnaMaxima < columnaCabecera + i) {
                            columnaMaxima = columnaCabecera + i;
                          }
                          break;
                        case "total dias":
                          columnasClave.columnaDiasTotal = i + columnaCabecera;
                          nuevaHoja
                            .cell(filaVacia - 1, columnaCabecera + i)
                            .value(cabeceras[i]);
                          if (columnaMaxima < columnaCabecera + i) {
                            columnaMaxima = columnaCabecera + i;
                          }
                          break;
                      }
                    }

                    //Columna anotacion:
                    columnasClave.columnaAnotacion = columnaMaxima + 1;

                    //Insertar datos:
                    var fechaSerializada;
                    var diasHastaFinDeMes;
                    var comentario = "";
                    for (var i = 0; i < enfermedades.length; i++) {
                      if (i == 0) {
                        console.log(
                          "Escribiendo enfermedad para [0]: ",
                          enfermedades[i],
                        );
                      }
                      nuevaHoja
                        .cell(filaVacia + i, 1)
                        .value(enfermedades[i].expedienteEmpresa);
                      nuevaHoja
                        .cell(filaVacia + i, 2)
                        .value(enfermedades[i].nombre);

                      nuevaHoja
                        .cell(filaVacia + i, 3)
                        .value(enfermedades[i].naf);

                      if (enfermedades[i].indicadorCarencia[0] == "S") {
                        nuevaHoja.cell(filaVacia + i, 5).value("SI");
                      }

                      nuevaHoja
                        .cell(filaVacia + i, 6)
                        .value(enfermedades[i].fechaBajaIt);

                      nuevaHoja
                        .cell(filaVacia + i, 8)
                        .value(enfermedades[i].fechaFinIt);
                      if (
                        Array.isArray(enfermedades[i].partesConfirmacion) &&
                        enfermedades[i].partesConfirmacion?.length > 0
                      ) {
                        nuevaHoja
                          .cell(filaVacia + i, 7)
                          .value(
                            enfermedades[i].partesConfirmacion[0]
                              .fechaSiguienteRevisionMedica,
                          );
                      }

                      //COMENTARIO:
                      comentario =
                        "Añadido automaticamente: " + enfermedades[i].tipo;

                      nuevaHoja
                        .cell(filaVacia + i, columnasClave.columnaAnotacion)
                        .value(comentario);

                      //Calculo dias hasta fin de mes:
                      //
                      fechaSerializada = excelSerialToUTCDate(
                        enfermedades[i].fechaBajaIt,
                      );
                      diasHastaFinDeMes =
                        diasRestantesFinDeMesInclusive(fechaSerializada);

                      //Todos los dias a =:
                      nuevaHoja.cell(filaVacia + i, 12).value(0);
                      nuevaHoja.cell(filaVacia + i, 13).value(0);
                      nuevaHoja.cell(filaVacia + i, 14).value(0);

                      nuevaHoja.cell(filaVacia + i, 15).value(0);

                      //PRIMER VALOR:
                      if (diasHastaFinDeMes > 3) {
                        nuevaHoja.cell(filaVacia + i, 12).value(3);
                      } else {
                        nuevaHoja
                          .cell(filaVacia + i, 12)
                          .value(diasHastaFinDeMes);
                      }

                      //SEGUNDO VALOR:
                      if (diasHastaFinDeMes > 15) {
                        nuevaHoja.cell(filaVacia + i, 13).value(12);
                      } else {
                        if (diasHastaFinDeMes > 3) {
                          nuevaHoja
                            .cell(filaVacia + i, 13)
                            .value(diasHastaFinDeMes - 3);
                        }
                      }

                      //Tercer VALOR:
                      if (diasHastaFinDeMes > 20) {
                        nuevaHoja.cell(filaVacia + i, 14).value(5);
                      } else {
                        if (diasHastaFinDeMes > 15) {
                          nuevaHoja
                            .cell(filaVacia + i, 14)
                            .value(diasHastaFinDeMes - 15);
                        }
                      }

                      //Tercer VALOR:
                      if (diasHastaFinDeMes > 20) {
                        nuevaHoja
                          .cell(filaVacia + i, 15)
                          .value(diasHastaFinDeMes - 20);
                      }
                    }

                    //ESCRITURA XLSX:
                    console.log("Escribiendo archivo Enfermedad...");
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

                      const hojas = archivoAccidentes.sheets();
                      const filas = archivoAccidentes
                        .sheet(hojas.length - 1)
                        .usedRange()._numRows;
                      const columnas = archivoAccidentes
                        .sheet(hojas.length - 1)
                        .usedRange()._numColumns;

                      const ultimaHoja = hojas.length - 1;
                      const nuevaHoja = archivoAccidentes.addSheet(
                        "Procesamiento automatico",
                      );

                      //Identificacion de cabeceras:
                      const { cabeceras, columnaCabecera, filaCabecera } =
                        deteccionCabeceras(archivoAccidentes, ultimaHoja);

                      //Identificar ultima fila:
                      var filaVacia = 0;
                      var flagVacia = true;
                      for (var i = 4; i < filas; i++) {
                        flagVacia = true;
                        if (
                          !archivoAccidentes
                            .sheet(hojas.length - 1)
                            .cell(i, 1)
                            .value()
                        ) {
                          for (var j = 1; j < columnas; j++) {
                            if (
                              archivoAccidentes
                                .sheet(hojas.length - 1)
                                .cell(i, j)
                                .value()
                            ) {
                              flagVacia = false;
                            }
                          }
                          if (flagVacia) {
                            filaVacia = i;
                            break;
                          }
                        }
                      }

                      //Creacion de objeto para accidentes:
                      const accidentes = [];
                      for (var i = 0; i < bajas.length; i++) {
                        if (bajas[i].contingencia) {
                          if (
                            bajas[i].contingencia[0] != 1 &&
                            bajas[i].contingencia[0] != 2
                          ) {
                            accidentes.push(bajas[i]);
                            accidentes[accidentes.length - 1].tipo = "BAJA";
                          }
                        }
                      }
                      for (var i = 0; i < altas.length; i++) {
                        if (altas[i].contingencia) {
                          if (
                            altas[i].contingencia[0] != 1 &&
                            altas[i].contingencia[0] != 2
                          ) {
                            accidentes.push(altas[i]);
                            accidentes[accidentes.length - 1].tipo = "ALTA";
                          }
                        }
                      }
                      for (var i = 0; i < confirmacion.length; i++) {
                        if (confirmacion[i].contingencia) {
                          if (
                            confirmacion[i].contingencia[0] != 1 &&
                            confirmacion[i].contingencia[0] != 2
                          ) {
                            accidentes.push(confirmacion[i]);
                            accidentes[accidentes.length - 1].tipo =
                              "CONFIRMACION";
                          }
                        }
                      }

                      console.log("ACCIDENTES [0]:");
                      console.log(accidentes[0]);

                      //Sobrescribir fila Vacia por ser hoja nueva:
                      filaVacia = 2;

                      var columnasClave = {
                        columnaExpediente: 0,
                        columnaNombre: 0,
                        columnaNaf: 0,
                        columnaDni: 0,
                        columnaFechaBaja: 0,
                        columnaProximaRev: 0,
                        columnaFechaAlta: 0,
                        columnaDias3: 0,
                        columnaDias5: 0,
                        columnaDias12: 0,
                        columnaDiasResto: 0,
                        columnaDiasTotal: 0,
                        columnaAnotacion: 0,
                      };

                      var columnaMaxima = 0;
                      for (var i = 0; i < cabeceras.length; i++) {
                        switch (cabeceras[i].toLowerCase().trim()) {
                          case "exp":
                            columnasClave.columnaExpediente =
                              i + columnaCabecera;
                            nuevaHoja
                              .cell(filaVacia - 1, columnaCabecera + i)
                              .value(cabeceras[i]);
                            if (columnaMaxima < columnaCabecera + i) {
                              columnaMaxima = columnaCabecera + i;
                            }
                            break;
                          case "apellidos y nombre":
                            columnasClave.columnaNombre = i + columnaCabecera;
                            nuevaHoja
                              .cell(filaVacia - 1, columnaCabecera + i)
                              .value(cabeceras[i]);
                            if (columnaMaxima < columnaCabecera + i) {
                              columnaMaxima = columnaCabecera + i;
                            }
                            break;
                          case "c.c.c.":
                            columnasClave.columnaNaf = i + columnaCabecera;
                            nuevaHoja
                              .cell(filaVacia - 1, columnaCabecera + i)
                              .value(cabeceras[i]);
                            if (columnaMaxima < columnaCabecera + i) {
                              columnaMaxima = columnaCabecera + i;
                            }
                            break;
                          case "dni":
                            columnasClave.columnaDni = i + columnaCabecera;
                            nuevaHoja
                              .cell(filaVacia - 1, columnaCabecera + i)
                              .value(cabeceras[i]);
                            if (columnaMaxima < columnaCabecera + i) {
                              columnaMaxima = columnaCabecera + i;
                            }
                            break;
                          case "f.  baja":
                            columnasClave.columnaFechaBaja =
                              i + columnaCabecera;
                            nuevaHoja
                              .cell(filaVacia - 1, columnaCabecera + i)
                              .value(cabeceras[i]);
                            if (columnaMaxima < columnaCabecera + i) {
                              columnaMaxima = columnaCabecera + i;
                            }
                            break;
                          case "próxima revision":
                            columnasClave.columnaProximaRev =
                              i + columnaCabecera;
                            nuevaHoja
                              .cell(filaVacia - 1, columnaCabecera + i)
                              .value(cabeceras[i]);
                            if (columnaMaxima < columnaCabecera + i) {
                              columnaMaxima = columnaCabecera + i;
                            }
                            break;
                          case "f. alta":
                            columnasClave.columnaFechaAlta =
                              i + columnaCabecera;
                            nuevaHoja
                              .cell(filaVacia - 1, columnaCabecera + i)
                              .value(cabeceras[i]);
                            if (columnaMaxima < columnaCabecera + i) {
                              columnaMaxima = columnaCabecera + i;
                            }
                            break;
                        }
                      }

                      //Columna anotacion:
                      columnasClave.columnaAnotacion = columnaMaxima + 1;

                      var comentario = "";
                      for (var i = 0; i < accidentes.length; i++) {
                        if (i == 0) {
                          console.log(
                            "Escribiendo accidentes para [0]: ",
                            accidentes[i],
                          );
                        }
                        nuevaHoja
                          .cell(filaVacia + i, 1)
                          .value(accidentes[i].expedienteEmpresa);
                        nuevaHoja
                          .cell(filaVacia + i, 2)
                          .value(accidentes[i].nombre);

                        nuevaHoja
                          .cell(filaVacia + i, 3)
                          .value(accidentes[i].naf);

                        if (accidentes[i].indicadorCarencia[0] == "S") {
                          nuevaHoja.cell(filaVacia + i, 5).value("SI");
                        }

                        nuevaHoja
                          .cell(filaVacia + i, 6)
                          .value(accidentes[i].fechaBajaIt);

                        nuevaHoja
                          .cell(filaVacia + i, 8)
                          .value(accidentes[i].fechaFinIt);
                        if (
                          Array.isArray(accidentes[i].partesConfirmacion) &&
                          accidentes[i].partesConfirmacion?.length > 0
                        ) {
                          nuevaHoja
                            .cell(filaVacia + i, 7)
                            .value(
                              accidentes[i].partesConfirmacion[0]
                                .fechaSiguienteRevisionMedica,
                            );
                        }

                        //COMENTARIO:
                        comentario =
                          "Añadido automaticamente: " + accidentes[i].tipo;

                        nuevaHoja
                          .cell(filaVacia + i, columnasClave.columnaAnotacion)
                          .value(comentario);
                      }

                      //ESCRITURA XLSX:
                      console.log("Escribiendo archivo Accidentes...");
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

function extraccionExcel(workbook, sheet, opts = null) {
  var filaCabecera = null;
  var columnaCabecera = null;

  //Activa la deteccion automatica:
  if (!opts) {
    var { columnaCabecera, filaCabecera } = deteccionCabeceras(workbook, sheet);
  } else {
    filaCabecera = opts.filaCabecera;
    columnaCabecera = opts.columnaCabecera;
  }

  if (columnaCabecera == null || filaCabecera == null) {
    console.log("Error, fallo en la extracción del excel.");
    return null;
  }

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

function deteccionCabeceras(workbook, sheet) {
  const columnas = workbook.sheet(sheet).usedRange()._numColumns;
  const filas = workbook.sheet(sheet).usedRange()._numRows;

  //Recorrer las primeras 10 filas e identificar el numero de columnas con valores:
  const filasAnalisis = filas < 10 ? filas : 10;

  const contadoresFilas = [];
  var contadorCampoRelleno = 0;
  for (var i = 1; i <= filasAnalisis; i++) {
    contadorCampoRelleno = 0;
    for (var j = 1; j <= columnas; j++) {
      if (workbook.sheet(sheet).cell(i, j).value()) {
        contadorCampoRelleno++;
      }
    }
    contadoresFilas.push(contadorCampoRelleno);
  }

  //Analisis de filas:
  var valorMedio = 0;
  for (var i = 0; i < contadoresFilas.length; i++) {
    valorMedio += contadoresFilas[i];
  }
  valorMedio = valorMedio / contadoresFilas.length;

  //Detecta primera fila con campos rellenos superior al valor medio -1:
  var filaCabecera = 0;
  for (var i = 0; i < contadoresFilas.length; i++) {
    if (contadoresFilas[i] > valorMedio - 1) {
      filaCabecera = i + 1;
      break;
    }
  }

  //Obtiene las cabeceras y la columna de inicio:
  var columnaCabecera = 0;
  const cabeceras = [];
  for (var i = 1; i < columnas; i++) {
    if (workbook.sheet(sheet).cell(filaCabecera, i).value()) {
      cabeceras.push(workbook.sheet(sheet).cell(filaCabecera, i).value());
      if (columnaCabecera == 0) {
        columnaCabecera = i;
      }
    }
  }

  //Identificacion de cabeceras:
  const objetoReturn = {
    cabeceras: cabeceras,
    columnaCabecera: columnaCabecera,
    filaCabecera: filaCabecera,
  };

  return objetoReturn;
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
// 1) Serial de Excel (sistema 1900) -> Date en UTC
function excelSerialToUTCDate(serial) {
  const dayMs = 24 * 60 * 60 * 1000;
  // Base 30/12/1899: ya contempla el bug del 1900-02-29
  const excelEpoch = Date.UTC(1899, 11, 30); // 1899-12-30
  return new Date(excelEpoch + serial * dayMs);
}

// 2) Días restantes hasta fin de mes (incluyendo hoy)
function diasRestantesFinDeMesInclusive(dateUTC) {
  // Normalizamos a medianoche UTC para evitar desajustes por horas
  const y = dateUTC.getUTCFullYear();
  const m = dateUTC.getUTCMonth();
  const hoyUTC = new Date(Date.UTC(y, m, dateUTC.getUTCDate()));
  const finMesUTC = new Date(Date.UTC(y, m + 1, 0)); // día 0 del mes siguiente = último del mes actual
  const dayMs = 24 * 60 * 60 * 1000;
  return Math.floor((finMesUTC - hoyUTC) / dayMs) + 1; // +1 para incluir el día de hoy
}

module.exports = ProcesosFie;
