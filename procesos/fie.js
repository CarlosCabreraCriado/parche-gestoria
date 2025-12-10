const path = require("path");
const fs = require("fs");
const readline = require("readline");
const axios = require("axios");
const moment = require("moment");
const XlsxPopulate = require("xlsx-populate");
const Datastore = require("nedb");
const _ = require("lodash");
const { DateTime } = require("luxon");

const { registrarEjecucion } = require("../metricas");
const { ipcRenderer } = require("electron");
const puppeteer = require("puppeteer");
const generatePDF = require("./pdf-fie");
const generarEmailFieDesdePlantilla = require("./emails-fie");
const DAY_MS = 24 * 60 * 60 * 1000;

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

      const nombreProceso = "FIE_1";
      let registrosProcesados = 0;

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

                const datosIncapacidad = extraccionExcel(archivoFIE, 0);
                const partesConfirmacion = extraccionExcel(archivoFIE, 4);

                const datosIncapacidad2 = extraccionExcel(archivoFIE, 1);

                console.log("Datos incapacidad:");
                console.log(datosIncapacidad[0]);

                console.log("Partes confirmacion:");
                console.log(partesConfirmacion[0]);

                console.log("Incapadidad 2");
                console.log(datosIncapacidad2[1]);

                // -- EXTRACCION DNI
                for (var i = 0; i < datosIncapacidad.length; i++) {
                  registrosProcesados += 1;
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

                // EXTRACCION FECHA PROXIMA REVISON (BAJA CON SEGUNDA HOJA)
                for (var i = 0; i < datosIncapacidad.length; i++) {
                  datosIncapacidad[i].fechaProximaRevisionParteBaja = null;
                  for (var j = 0; j < datosIncapacidad2.length; j++) {
                    if (
                      datosIncapacidad[i].nif == datosIncapacidad2[j].nif &&
                      datosIncapacidad[i].nif
                    ) {
                      console.log("Encontrado", datosIncapacidad2[j]);

                      datosIncapacidad[i].datosAdicionales =
                        datosIncapacidad2[j];

                      datosIncapacidad[i].fechaProximaRevisionParteBaja =
                        datosIncapacidad2[
                          j
                        ].fechaSiguienteRevisionMedicaParteDeBaja;
                    }
                  }
                }

                // -- IDENTIFICACION (ALTAS, BAJAS, CONFIRMACION)
                var partesDetectados = [];
                //Detectamos si los distintos casos:
                for (var i = 0; i < datosIncapacidad.length; i++) {
                  if (datosIncapacidad[i].fechaFinIt) {
                    altas.push(datosIncapacidad[i]);
                    continue;
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
                    continue;
                  } else {
                    bajas.push(datosIncapacidad[i]);
                    continue;
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

                  // 🔍 Depuración: detectar emails no string
                  if (empresa && typeof empresa.email !== "string") {
                    console.log(
                      "EMAIL NO STRING EN ALTAS:",
                      empresa.codigo,
                      empresa.empresa,
                      empresa.email,
                      typeof empresa.email,
                    );
                  }
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
                  if (empresa && typeof empresa.email !== "string") {
                    console.log(
                      "EMAIL NO STRING EN BAJAS:",
                      empresa.codigo,
                      empresa.empresa,
                      empresa.email,
                      typeof empresa.email,
                    );
                  }
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
                  if (empresa && typeof empresa.email !== "string") {
                    console.log(
                      "EMAIL NO STRING EN CONFIRMACION:",
                      empresa.codigo,
                      empresa.empresa,
                      empresa.email,
                      typeof empresa.email,
                    );
                  }
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
                
                function obtenerEmailsDestino(emailsEmpresa) {
                  if (typeof emailsEmpresa !== "string") {
                    // Si no es string (número, objeto, etc.), no intentamos enviar nada
                    return [];
                  }
                
                  return emailsEmpresa
                    .split(";")
                    .map((e) => e.trim())
                    .filter(Boolean);   // quita vacíos
                }

                for (const r of altas) {
                  const file = await generarEmailFieDesdePlantilla(
                    r,
                    "ALTAS",
                    pathSalidaPDFAltasCorreos,
                    {
                      to: obtenerEmailsDestino(r.emailsEmpresa),
                    },
                  );
                  results.push(file);
                }

                for (const r of bajas) {
                  const file = await generarEmailFieDesdePlantilla(
                    r,
                    "BAJAS",
                    pathSalidaPDFBajasCorreos,
                    {
                      to: obtenerEmailsDestino(r.emailsEmpresa),
                    },
                  );
                  results.push(file);
                }

                for (const r of confirmacion) {
                  const file = await generarEmailFieDesdePlantilla(
                    r,
                    "CONFIRMACION",
                    pathSalidaPDFConfirmacionCorreos,
                    {
                      to: obtenerEmailsDestino(r.emailsEmpresa),
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
                    for (var i = filaCabecera; i < filas; i++) {
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
                    columnasClave.columnaAnotacion = columnaMaxima + 2;
                    columnasClave.columnaMesProcesamiento = columnaMaxima + 1;

                    //Insertar datos:
                    var fechaBajaSerializada;
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
                        .value(enfermedades[i].fechaBajaIt)
                        .style("numberFormat", "dd/mm/yyyy");
                      nuevaHoja
                        .cell(filaVacia + i, 8)
                        .value(enfermedades[i].fechaFinIt)
                        .style("numberFormat", "dd/mm/yyyy");
                      if (
                        Array.isArray(enfermedades[i].partesConfirmacion) &&
                        enfermedades[i].partesConfirmacion?.length > 0
                      ) {
                        nuevaHoja
                          .cell(filaVacia + i, 7)
                          .value(
                            enfermedades[i].partesConfirmacion[0]
                              .fechaSiguienteRevisionMedica,
                          )
                          .style("numberFormat", "dd/mm/yyyy");
                      }

                      //COMENTARIO:
                      comentario =
                        "Añadido automaticamente: " + enfermedades[i].tipo;

                      nuevaHoja
                        .cell(filaVacia + i, columnasClave.columnaAnotacion)
                        .value(comentario);

                      //Calculo dias hasta fin de mes:
                      fechaBajaSerializada = excelSerialToUTCDate(
                        enfermedades[i].fechaBajaIt,
                      );

                      //Obtine el mes en el que se evalua:
                      var fechaRecepcionSerializada = excelSerialToUTCDate(
                        enfermedades[i].fechaRecepcion,
                      );

                      var mesActualIndex = fechaRecepcionSerializada.getMonth();
                      var mesBajaIndex = fechaBajaSerializada.getMonth();
                      var nombreMesActual =
                        obtenerNombreMesByIndex(mesActualIndex);

                      //Marcando mes de procesamiento:
                      nuevaHoja
                        .cell(
                          filaVacia + i,
                          columnasClave.columnaMesProcesamiento,
                        )
                        .value("Mes base: " + nombreMesActual);

                      //Detecta si esta evaluandose en el mismo mes:
                      var diasDeMesAnterior = 0;
                      var primeroDeMes = obtenerPrimeroDeMes(
                        fechaRecepcionSerializada,
                      );
                      var fechaFinal = obtenerUltimoDeMes(
                        fechaRecepcionSerializada,
                      );

                      if (enfermedades[i].fechaFinIt) {
                        fechaFinal = excelSerialToUTCDate(
                          enfermedades[i].fechaFinIt,
                        );
                      }

                      function startOfDayUTC(d) {
                        return Date.UTC(
                          d.getUTCFullYear(),
                          d.getUTCMonth(),
                          d.getUTCDate(),
                        );
                      }

                      function diasEntreFechasUTC(a, b) {
                        const au = startOfDayUTC(a);
                        const bu = startOfDayUTC(b);
                        return Math.round((bu - au) / DAY_MS); // o Math.floor si prefieres
                      }

                      if (mesActualIndex !== mesBajaIndex) {
                        diasDeMesAnterior = diasEntreFechasUTC(
                          fechaBajaSerializada,
                          primeroDeMes,
                        );
                      }

                      //Calculo de dias entre fecha inicio y fin:
                      if (diasDeMesAnterior == 0) {
                        diasHastaFinDeMes =
                          diasEntreFechasUTC(fechaBajaSerializada, fechaFinal) +
                          1;
                      } else {
                        diasHastaFinDeMes =
                          diasEntreFechasUTC(primeroDeMes, fechaFinal) + 1;
                      }

                      console.log("CALCULO FECHAS:");
                      console.log("fecha baja:", fechaBajaSerializada);
                      console.log("fecha primero de mes:", primeroDeMes);
                      console.log("Mes actual:", nombreMesActual);
                      console.log("Dias mes anterior: ", diasDeMesAnterior);
                      console.log(
                        "Dias restantes fin de mes:",
                        diasHastaFinDeMes,
                      );

                      console.log("mes actual:", nombreMesActual);

                      //Todos los dias a =:
                      nuevaHoja.cell(filaVacia + i, 12).value(0);
                      nuevaHoja.cell(filaVacia + i, 13).value(0);
                      nuevaHoja.cell(filaVacia + i, 14).value(0);

                      nuevaHoja.cell(filaVacia + i, 15).value(0);

                      var valorEscritura = 0;
                      var restante = diasHastaFinDeMes;
                      var valoresUmbral = [3, 12, 5]; //Umbrales de dias

                      for (var k = 0; k < valoresUmbral.length; k++) {
                        if (restante <= 0) {
                          continue;
                        }
                        if (restante >= valoresUmbral[k]) {
                          if (diasDeMesAnterior >= valoresUmbral[k]) {
                            valorEscritura = 0;
                            diasDeMesAnterior =
                              diasDeMesAnterior - valoresUmbral[k];
                          } else {
                            valorEscritura =
                              valoresUmbral[k] - diasDeMesAnterior;
                            diasDeMesAnterior = 0;
                          }
                        } else {
                          if (diasDeMesAnterior >= restante) {
                            valorEscritura = 0;
                            diasDeMesAnterior = diasDeMesAnterior - restante;
                          } else {
                            valorEscritura = restante - diasDeMesAnterior;
                            diasDeMesAnterior = 0;
                          }
                        }

                        nuevaHoja
                          .cell(filaVacia + i, 12 + k)
                          .value(valorEscritura);
                        restante = restante - valorEscritura;
                      }

                      //VALOR RESTANTE:
                      nuevaHoja.cell(filaVacia + i, 15).value(restante);
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

                      console.log("Fila Cabedera Accidentes:", filaCabecera);
                      //Identificar ultima fila:
                      var filaVacia = 0;
                      var flagVacia = true;
                      for (var i = filaCabecera; i < filas; i++) {
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
                      console.log("Cabeceras Accidentes:", cabeceras);

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
                        columnaDias: 0,
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
                          case "próxima revisión":
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

                          case "dias 75%":
                            columnasClave.columnaDias = i + columnaCabecera;
                            nuevaHoja
                              .cell(filaVacia - 1, columnaCabecera + i)
                              .value(cabeceras[i]);
                            if (columnaMaxima < columnaCabecera + i) {
                              columnaMaxima = columnaCabecera + i;
                            }
                            break;
                        }
                      }

                      //Cabecera custom para indicador de carencia:
                      nuevaHoja
                        .cell(filaVacia - 1, 5)
                        .value("INDICADOR CARENCIA");

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
                          .cell(filaVacia + i, columnasClave.columnaExpediente)
                          .value(accidentes[i].expedienteEmpresa);
                        nuevaHoja
                          .cell(filaVacia + i, columnasClave.columnaNombre)
                          .value(accidentes[i].nombre);

                        nuevaHoja
                          .cell(filaVacia + i, columnasClave.columnaNaf)
                          .value(accidentes[i].naf);

                        nuevaHoja
                          .cell(filaVacia + i, columnasClave.columnaDni)
                          .value(accidentes[i].dni);

                        if (accidentes[i].indicadorCarencia[0] == "S") {
                          nuevaHoja.cell(filaVacia + i, 5).value("SI");
                        }

                        nuevaHoja
                          .cell(filaVacia + i, columnasClave.columnaFechaBaja)
                          .value(accidentes[i].fechaBajaIt)
                          .style("numberFormat", "dd/mm/yyyy");
                        nuevaHoja
                          .cell(filaVacia + i, columnasClave.columnaFechaAlta)
                          .value(accidentes[i].fechaFinIt)
                          .style("numberFormat", "dd/mm/yyyy");
                        if (
                          Array.isArray(accidentes[i].partesConfirmacion) &&
                          accidentes[i].partesConfirmacion?.length > 0
                        ) {
                          nuevaHoja
                            .cell(
                              filaVacia + i,
                              columnasClave.columnaProximaRev,
                            )
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

                        //Obtine el mes en el que se evalua:
                        var fechaRecepcionSerializada = excelSerialToUTCDate(
                          accidentes[i].fechaRecepcion,
                        );
                        var fechaBajaSerializada = excelSerialToUTCDate(
                          accidentes[i].fechaBajaIt,
                        );

                        var mesActualIndex =
                          fechaRecepcionSerializada.getMonth();
                        var mesBajaIndex = fechaBajaSerializada.getMonth();
                        var nombreMesActual =
                          obtenerNombreMesByIndex(mesActualIndex);

                        //Marcando mes de procesamiento:
                        nuevaHoja
                          .cell(
                            filaVacia + i,
                            columnasClave.columnaAnotacion + 1,
                          )
                          .value("Mes base: " + nombreMesActual);

                        console.log("Escribiendo mes: ", nombreMesActual);
                        console.log(filaVacia + i);

                        //Detecta si esta evaluandose en el mismo mes:
                        var diasDeMesAnterior = 0;
                        var primeroDeMes = obtenerPrimeroDeMes(
                          fechaRecepcionSerializada,
                        );
                        var fechaFinal = obtenerUltimoDeMes(
                          fechaRecepcionSerializada,
                        );

                        if (accidentes[i].fechaFinIt) {
                          fechaFinal = excelSerialToUTCDate(
                            accidentes[i].fechaFinIt,
                          );
                        }

                        function startOfDayUTC(d) {
                          return Date.UTC(
                            d.getUTCFullYear(),
                            d.getUTCMonth(),
                            d.getUTCDate(),
                          );
                        }

                        function diasEntreFechasUTC(a, b) {
                          const au = startOfDayUTC(a);
                          const bu = startOfDayUTC(b);
                          return Math.round((bu - au) / DAY_MS); // o Math.floor si prefieres
                        }

                        if (mesActualIndex !== mesBajaIndex) {
                          diasDeMesAnterior = diasEntreFechasUTC(
                            fechaBajaSerializada,
                            primeroDeMes,
                          );
                        }

                        //Calculo de dias entre fecha inicio y fin:
                        var diasHastaFinDeMes = 0;
                        if (diasDeMesAnterior == 0) {
                          diasHastaFinDeMes =
                            diasEntreFechasUTC(
                              fechaBajaSerializada,
                              fechaFinal,
                            ) + 1;
                        } else {
                          diasHastaFinDeMes =
                            diasEntreFechasUTC(primeroDeMes, fechaFinal) + 1;
                        }

                        nuevaHoja
                          .cell(filaVacia + i, columnasClave.columnaDias)
                          .value(diasHastaFinDeMes);

                        console.log("Escribiendo dias: ", diasHastaFinDeMes);
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
                          //console.log(archivoFIE)
                          registrarEjecucion({
                            nombreProceso,
                            registrosProcesados: registrosProcesados,
                          });
                          console.log("Fin del procesamiento");

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

  async fIE_2(argumentos) {
    const DEBUG = false;
    const logDebug = (...args) => {
      if (DEBUG) console.log(...args);
    };

    console.log(
      "[FIE_2] Iniciando proceso FIE_2 (lectura Excel + automatización web)",
    );

    const nombreProceso = "FIE_2";
    let registrosProcesados = 0;

    return new Promise(async (resolve) => {
      let browser = null;

      try {
        // 1) Entradas (nuevo orden)
        const chromeExePath = argumentos?.formularioControl?.[0];
        const pathArchivoFIE_2 = argumentos?.formularioControl?.[1];
        const pathSalidaBase = argumentos?.formularioControl?.[2];

        if (!chromeExePath || !fs.existsSync(chromeExePath)) {
          console.error("[FIE_2] Ruta a chrome.exe no válida.");
          return resolve(false);
        }
        if (!pathArchivoFIE_2 || typeof pathArchivoFIE_2 !== "string") {
          console.error(
            "[FIE_2] argumentos.formularioControl[1] (Excel) no es una ruta válida.",
          );
          return resolve(false);
        }

        // 2) Carpeta de salida
        let pathSalidaPDFConfirmacion = null;
        if (pathSalidaBase && typeof pathSalidaBase === "string") {
          pathSalidaPDFConfirmacion = path.join(
            path.normalize(pathSalidaBase),
            `TA2 B (${this.getCurrentDateString()})`,
          );
          if (!fs.existsSync(pathSalidaPDFConfirmacion)) {
            fs.mkdirSync(pathSalidaPDFConfirmacion, { recursive: true });
            console.log(`[FIE_2] Carpeta creada: ${pathSalidaPDFConfirmacion}`);
          } else {
            logDebug(
              `[FIE_2] Carpeta ya existente: ${pathSalidaPDFConfirmacion}`,
            );
          }
        } else {
          console.warn(
            "[FIE_2] No se proporcionó carpeta de salida (arg[2]). No se guardarán PDFs.",
          );
        }

        // 3) Lectura Excel
        const rutaNormalizada = path.normalize(pathArchivoFIE_2);
        console.log(`[FIE_2] Cargando Excel: ${rutaNormalizada}`);
        const workbook = await XlsxPopulate.fromFileAsync(rutaNormalizada);
        logDebug("[FIE_2] Archivo Excel cargado correctamente.");

        // --- Lectura de las dos hojas ---
        const datosHoja1 = extraccionExcel(workbook, 0); // hoja principal
        const datosHoja2 = extraccionExcel(workbook, 1); // hoja con Fecha AT/EP

        if (!Array.isArray(datosHoja1)) {
          console.error(
            "[FIE_2] extraccionExcel (hoja 0) no devolvió un array válido.",
          );
          return resolve(false);
        }
        if (!Array.isArray(datosHoja2)) {
          console.warn(
            "[FIE_2] extraccionExcel (hoja 1) no devolvió un array. Continúo sin Fecha AT/EP.",
          );
        }

        // Helper sencillo para comprobar si un campo está relleno (versión JS)
        const campoRelleno = (reg, nombreCampo) => {
          const valor = reg && reg[nombreCampo];
          return (
            valor !== null && valor !== undefined && String(valor).trim() !== ""
          );
        };

        // Filtrar registros de la hoja 1
        const datosHoja1Filtrados = datosHoja1.filter(
          (reg) =>
            campoRelleno(reg, "tipoContrato") &&
            campoRelleno(reg, "base") &&
            campoRelleno(reg, "dia") &&
            campoRelleno(reg, "puestoDeTrabajo") &&
            campoRelleno(reg, "cnoe") &&
            campoRelleno(reg, "detalleCnoe") &&
            campoRelleno(reg, "bajaYAlta"),
        );

        console.log(
          `[FIE_2] Registros hoja 0 totales: ${datosHoja1.length}. Registros válidos tras filtro: ${datosHoja1Filtrados.length}.`,
        );

        if (!datosHoja1Filtrados.length) {
          console.warn(
            "[FIE_2] No hay registros en la hoja 0 que cumplan todos los campos obligatorios.",
          );
          return resolve([]);
        }

        // Helpers para NIF y clave Fecha AT/EP
        const obtenerNifRegistro = (reg) => {
          if (!reg || typeof reg !== "object") return "";
          const keys = Object.keys(reg);
          // buscamos un campo cuyo nombre contenga "nif"
          const nifKey = keys.find((k) => k.toLowerCase().includes("nif"));
          return nifKey ? String(reg[nifKey] ?? "").trim() : "";
        };

        const normalizaNif = (nif) =>
          String(nif ?? "")
            .toUpperCase()
            .replace(/\s+/g, "");

        // Detectamos dinámicamente la clave de "Fecha AT/EP" en la hoja 2
        let fechaATEPKey = null;
        if (Array.isArray(datosHoja2) && datosHoja2.length > 0) {
          const sampleKeys = Object.keys(datosHoja2[0]);
          fechaATEPKey = sampleKeys.find((k) => {
            const norm = k.toLowerCase().replace(/[^a-z0-9]/g, "");
            // algo tipo "fechaat/ep", "fechaatep", etc.
            return norm.includes("fecha") && norm.includes("atep");
          });

          console.log(
            "[FIE_2] Clave detectada para Fecha AT/EP en hoja 2:",
            fechaATEPKey,
          );
          if (!fechaATEPKey) {
            console.warn(
              "[FIE_2] No se pudo detectar automáticamente la columna de Fecha AT/EP en la hoja 2.",
            );
          }
        }

        // Construimos un mapa NIF -> Fecha AT/EP a partir de la hoja 2
        const mapaFechaATEP = new Map();

        if (Array.isArray(datosHoja2) && fechaATEPKey) {
          for (const reg2 of datosHoja2) {
            const nifRaw = obtenerNifRegistro(reg2);
            const nifNorm = normalizaNif(nifRaw);
            if (!nifNorm) continue;

            const valorFecha = reg2[fechaATEPKey];
            if (valorFecha != null && valorFecha !== "") {
              mapaFechaATEP.set(nifNorm, valorFecha);
            }
          }
          console.log(
            `[FIE_2] Mapa Fecha AT/EP construido con ${mapaFechaATEP.size} NIF distintos.`,
          );
        } else {
          console.warn(
            "[FIE_2] No hay datos válidos en hoja 2 o no se encontró clave de Fecha AT/EP; no se fusionarán fechas.",
          );
        }

        // Mezclamos datos de la hoja 1 con la Fecha AT/EP de la hoja 2 (por NIF)
        const datos = datosHoja1Filtrados.map((reg1) => {
          const nifRaw = obtenerNifRegistro(reg1);
          const nifNorm = normalizaNif(nifRaw);
          const fechaDesdeHoja2 = nifNorm ? mapaFechaATEP.get(nifNorm) : null;

          if (fechaDesdeHoja2 != null && fechaDesdeHoja2 !== "") {
            return {
              ...reg1,
              fechaATEP: fechaDesdeHoja2, // esta es la que luego usas en procesarRegistro
            };
          }

          return reg1;
        });

        console.log(`[FIE_2] Filas leídas en Excel (hoja 0): ${datos.length}`);
        if (DEBUG && datos.length > 0) {
          logDebug("[FIE_2] Muestra primer registro fusionado:", datos[0]);
        }
        if (!datos.length) {
          console.warn(
            "[FIE_2] No hay registros en el Excel. Nada que procesar.",
          );
          return resolve(datos);
        }

        // 4) Abrir navegador real (Chrome)
        const urlFS = "https://w2.seg-social.es/fs/indexframes.html";
        try {
          browser = await puppeteer.launch({
            headless: false,
            defaultViewport: null,
            executablePath: chromeExePath,
            args: [
              "--start-maximized",
              "--no-sandbox",
              "--disable-setuid-sandbox",
              "--disable-features=IsolateOrigins,site-per-process",
            ],
          });

          const opened = await browser.pages();
          var page = opened.length ? opened[0] : await browser.newPage();

          // Aceptar automáticamente los popups (alert, confirm, beforeunload...)
          page.on("dialog", async (dialog) => {
            try {
              logDebug(
                "[FIE_2] Dialog detectado:",
                dialog.type(),
                JSON.stringify(dialog.message()),
              );
              await dialog.accept();
              logDebug("[FIE_2] Dialog aceptado automáticamente.");
            } catch (e) {
              console.warn("[FIE_2] Error al aceptar dialog:", e?.message || e);
            }
          });

          await page.goto(urlFS, { waitUntil: "domcontentloaded" });
          console.log(
            "[FIE_2] Chrome abierto en FS. Selecciona el certificado si aparece diálogo.",
          );

          // 5) Helpers + procesamiento secuencial de registros
          if (page && datos.length > 0) {
            const pause = (ms) => new Promise((r) => setTimeout(r, ms));

            const openITOnline = async () => {
              try {
                let clicked = false;
                for (const fr of page.frames()) {
                  const link = await fr.$('a.a2[href*="IWXP0002"]');
                  if (link) {
                    await link.click({ delay: 40 });
                    clicked = true;
                    logDebug(
                      "[FIE_2] Click en 'Incapacidad temporal Online' (href).",
                    );
                    break;
                  }
                }
                if (!clicked) {
                  for (const fr of page.frames()) {
                    const ok = await fr.evaluate(() => {
                      const norm = (s) =>
                        (s || "").trim().toLowerCase().replace(/\s+/g, " ");
                      const target = "incapacidad temporal online";
                      const a = Array.from(document.querySelectorAll("a")).find(
                        (x) => norm(x.textContent).includes(target),
                      );
                      if (a) {
                        a.target = "_self";
                        a.click();
                        return true;
                      }
                      return false;
                    });
                    if (ok) {
                      logDebug(
                        "[FIE_2] Click en 'Incapacidad temporal Online' (texto).",
                      );
                      break;
                    }
                  }
                }
                await this.esperar(1000);
              } catch (e) {
                console.warn(
                  "[FIE_2] No se pudo clicar el enlace de IT Online:",
                  e?.message || e,
                );
              }
            };

            const fillTextWithRetry = async (
              frame,
              selector,
              rawValue,
              {
                tries = 4,
                typeDelay = 60,
                betweenTriesMs = 250,
                commitTab = true,
                digitsOnlyCompare = true,
              } = {},
            ) => {
              const value = String(rawValue ?? "");
              const el = await frame.waitForSelector(selector, {
                visible: true,
                timeout: 15000,
              });
              await el.evaluate((e) => e.scrollIntoView({ block: "center" }));

              const isMac = process.platform === "darwin";
              const modKey = isMac ? "Meta" : "Control";

              for (let i = 1; i <= tries; i++) {
                try {
                  await el.click({ clickCount: 3, delay: 30 });
                  await page.keyboard.down(modKey);
                  await page.keyboard.press("KeyA");
                  await page.keyboard.up(modKey);
                  await page.keyboard.press("Backspace");
                  await pause(40);

                  await el.type(value, { delay: typeDelay });

                  await el.evaluate((e) => {
                    e.dispatchEvent(new Event("input", { bubbles: true }));
                    e.dispatchEvent(new Event("change", { bubbles: true }));
                  });

                  if (commitTab) {
                    await page.keyboard.press("Tab");
                    await pause(120);
                    await page.keyboard.down("Shift");
                    await page.keyboard.press("Tab");
                    await page.keyboard.up("Shift");
                  }

                  const current = await el.evaluate((e) => e.value ?? "");
                  const norm = (s) =>
                    digitsOnlyCompare
                      ? String(s).replace(/\D/g, "")
                      : String(s);
                  logDebug(
                    `[FIE_2] Verificación ${selector} intento ${i}:`,
                    current,
                  );

                  if (norm(current) === norm(value)) return true;

                  await el.evaluate((_, val) => {
                    _.value = val;
                    _.dispatchEvent(new Event("input", { bubbles: true }));
                    _.dispatchEvent(new Event("change", { bubbles: true }));
                    _.blur?.();
                  }, value);

                  const after = await el.evaluate((e) => e.value ?? "");
                  if (norm(after) === norm(value)) return true;
                } catch (e) {
                  console.warn(
                    `[FIE_2] fillTextWithRetry fallo intento ${i} en ${selector}:`,
                    e?.message || e,
                  );
                }
                await pause(betweenTriesMs + i * 150);
              }
              console.warn(
                `[FIE_2] ${selector} no se pudo fijar tras ${tries} intentos`,
              );
              return false;
            };

            const fillIfPresent = async (
              frame,
              selector,
              value,
              opts = { tries: 3, typeDelay: 35, digitsOnlyCompare: false },
            ) => {
              try {
                const val = String(value ?? "");
                const elHandle = await frame.$(selector);

                if (!elHandle) {
                  logDebug(
                    `[FIE_2] Campo opcional NO presente: ${selector}. Continúo.`,
                  );
                  return false;
                }
                if (!val) {
                  logDebug(
                    `[FIE_2] Sin valor para ${selector}. Omite rellenado.`,
                  );
                  return false;
                }

                const isVisible = await elHandle
                  .evaluate((e) => {
                    const s = getComputedStyle(e);
                    const r = e.getBoundingClientRect();
                    return (
                      s.visibility !== "hidden" &&
                      s.display !== "none" &&
                      r.width > 0 &&
                      r.height > 0
                    );
                  })
                  .catch(() => false);

                if (isVisible) {
                  try {
                    await fillTextWithRetry(frame, selector, val, opts);
                    return true;
                  } catch (e) {}
                }

                const ok = await frame.evaluate(
                  (sel, v) => {
                    const el = document.querySelector(sel);
                    if (!el) return false;
                    el.value = v;
                    el.dispatchEvent(new Event("input", { bubbles: true }));
                    el.dispatchEvent(new Event("change", { bubbles: true }));
                    el.blur && el.blur();
                    return true;
                  },
                  selector,
                  val,
                );

                logDebug(
                  ok
                    ? `[FIE_2] ${selector} fijado por JS (fallback, posible campo oculto).`
                    : `[FIE_2] No se pudo fijar ${selector} por JS.`,
                );

                return ok;
              } catch (e) {
                console.warn(
                  `[FIE_2] No pude rellenar opcional ${selector}:`,
                  e?.message || e,
                );
                return false;
              }
            };

            const selectWithRetry = async (
              frame,
              selector,
              rawValue,
              { tries = 4, betweenTriesMs = 250 } = {},
            ) => {
              const value = String(rawValue ?? "");
              await frame.waitForSelector(selector, {
                visible: true,
                timeout: 15000,
              });
              await frame.$eval(selector, (el) =>
                el.scrollIntoView({ block: "center" }),
              );

              for (let i = 1; i <= tries; i++) {
                try {
                  await frame.select(selector, value);
                  await pause(100);
                  let current = await frame.$eval(
                    selector,
                    (el) => el.value ?? "",
                  );
                  logDebug(
                    `[FIE_2] Verificación select ${selector} intento ${i}:`,
                    current,
                  );
                  if (current === value) return true;

                  await frame.evaluate(
                    (sel, val) => {
                      const el = document.querySelector(sel);
                      if (!el) return;
                      el.value = val;
                      el.dispatchEvent(new Event("input", { bubbles: true }));
                      el.dispatchEvent(new Event("change", { bubbles: true }));
                      el.blur?.();
                    },
                    selector,
                    value,
                  );

                  await pause(120);
                  current = await frame.$eval(selector, (el) => el.value ?? "");
                  logDebug(
                    `[FIE_2] Verificación fallback ${selector} intento ${i}:`,
                    current,
                  );
                  if (current === value) return true;
                } catch (e) {
                  console.warn(
                    `[FIE_2] selectWithRetry fallo intento ${i} en ${selector}:`,
                    e?.message || e,
                  );
                }
                await pause(betweenTriesMs + i * 150);
              }
              console.warn(
                `[FIE_2] ${selector} no se pudo seleccionar tras ${tries} intentos`,
              );
              return false;
            };

            const findFrameWithSelector = async (
              selector,
              timeoutMs = 25000,
              pollMs = 400,
            ) => {
              const start = Date.now();
              while (Date.now() - start < timeoutMs) {
                for (const fr of page.frames()) {
                  try {
                    const el = await fr.$(selector);
                    if (el) return fr;
                  } catch (e) {}
                }
                await pause(pollMs);
              }
              return null;
            };

            const toDDMMYYYY = (date) => {
              const dd = String(date.getUTCDate()).padStart(2, "0");
              const mm = String(date.getUTCMonth() + 1).padStart(2, "0");
              const yyyy = String(date.getUTCFullYear());
              return `${dd}/${mm}/${yyyy}`;
            };
            const excelSerialToDDMMYYYY = (serial) =>
              toDDMMYYYY(excelSerialToUTCDate(serial));
            const extraeRegimenYCCC = (cccRaw) => {
              const digits = String(cccRaw ?? "").replace(/\D/g, "");
              return {
                regimen: digits.slice(0, 4).padStart(4, "0"),
                cccResto: digits.slice(4),
              };
            };
            const limpiaDigitos = (n) => String(n ?? "").replace(/\D/g, "");
            const extraeCodigoContingencia = (campo) => {
              const s = String(campo ?? "");
              const m = s.match(/^(\d+)\s*=/);
              return m ? m[1] : "";
            };

            let registrosOk = 0;
            let registrosError = 0;

            // Mapa NIF -> LOG y helper para añadir mensajes
            const logsPorNif = new Map();

            const appendLog = (nifNorm, msg) => {
              if (!nifNorm) return;
              const prev = logsPorNif.get(nifNorm);
              logsPorNif.set(nifNorm, prev ? `${prev} | ${msg}` : msg);
            };

            const procesarRegistro = async (r, indice) => {
              // NIF normalizado del registro actual (para mapear LOG en el Excel)
              const nifNormRegistro = normalizaNif(obtenerNifRegistro(r));

              console.log(
                `[FIE_2] Procesando registro ${indice + 1}/${datos.length} (NAF: ${
                  r?.naf ?? "sin NAF"
                })`,
              );

              // Volvemos siempre a la URL base y entramos de nuevo en IT Online
              try {
                await page.goto(urlFS, { waitUntil: "domcontentloaded" });
              } catch (e) {
                const msg = `[ERROR] No se pudo acceder a FS en el registro ${
                  indice + 1
                }: ${e?.message || e}`;
                console.warn("[FIE_2]", msg);
                appendLog(nifNormRegistro, msg);
                registrosError++;
                return;
              }

              await openITOnline();

              // === Pantalla 1: formulario principal ===
              logDebug("[FIE_2] Buscando frame con el formulario inicial...");
              const formFrame = await findFrameWithSelector(
                "#regimen",
                25000,
                400,
              );
              if (!formFrame) {
                const msg =
                  "[ERROR] No se encontró el formulario inicial (#regimen) en ningún frame.";
                console.warn("[FIE_2]", msg);
                appendLog(nifNormRegistro, msg);
                registrosError++;
                return;
              }

              const { regimen, cccResto } = extraeRegimenYCCC(r?.ccc);
              const naf = limpiaDigitos(r?.naf);
              const contCode = extraeCodigoContingencia(r?.contingencia);
              const fechaBajaStr = r?.fechaBajaIt
                ? excelSerialToDDMMYYYY(r.fechaBajaIt)
                : "";

              if (DEBUG) {
                console.table({
                  "Regimen (4)": regimen,
                  "CCC (resto, 11)": cccResto,
                  "NAF (12)": naf,
                  "Contingencia (1-5)": contCode,
                  "Fecha de baja": fechaBajaStr,
                });
              }

              await fillTextWithRetry(formFrame, "#regimen", regimen);
              await pause(200);
              await fillTextWithRetry(formFrame, "#ccc", cccResto);
              await pause(200);
              await fillTextWithRetry(formFrame, "#naf", naf);
              await pause(200);

              if (["1", "2", "3", "4", "5"].includes(contCode)) {
                await selectWithRetry(formFrame, "#contingencias", contCode);
              } else {
                console.warn(
                  "[FIE_2] Contingencia no reconocida:",
                  r?.contingencia,
                );
              }
              await pause(200);

              if (fechaBajaStr) {
                await fillTextWithRetry(formFrame, "#fechaBaja", fechaBajaStr, {
                  digitsOnlyCompare: false,
                });
              } else {
                console.warn(
                  "[FIE_2] Sin fecha de baja válida; no se rellena #fechaBaja.",
                );
              }

              try {
                await formFrame.waitForSelector("#ENVIO_7", {
                  visible: true,
                  timeout: 8000,
                });
                await formFrame.click("#ENVIO_7", { delay: 60 });
                logDebug("[FIE_2] Click en Aceptar (ENVIO_7).");
              } catch (e) {
                console.warn(
                  "[FIE_2] No se pudo clicar Aceptar:",
                  e?.message || e,
                );
              }

              await this.esperar(1000);

              // === Pantalla 2: Grabación de partes ===
              try {
                await Promise.race([
                  page
                    .waitForNavigation({
                      waitUntil: "domcontentloaded",
                      timeout: 15000,
                    })
                    .catch(() => {}),
                  pause(1500),
                ]);

                const form2 =
                  (await findFrameWithSelector("#FORMULARIO_4", 25000, 400)) ||
                  (await findFrameWithSelector("#puestoTrabajo", 25000, 400));
                if (!form2) {
                  const msg =
                    "[ERROR] No se encontró el formulario de 'Grabación de partes'.";
                  console.warn("[FIE_2]", msg);
                  appendLog(nifNormRegistro, msg);
                  registrosError++;
                  return;
                } else {
                  const puestoDeTrabajo = String(
                    r?.puestoDeTrabajo ?? r?.puestoTrabajo ?? "",
                  );
                  const cnoe = String(r?.cnoe ?? "");
                  const tipoContratoIn = String(r?.tipoContrato ?? "");

                  const baseResto = String(r?.base ?? "");
                  const diasResto = String(r?.dia ?? "");
                  const baseFijoParcial = String(r?.base ?? "");
                  const diasFijoParcial = String(r?.dia ?? "");

                  const detalleCnoe = String(r?.detalleCnoe ?? "");

                  let fechaATEP = "";
                  if (r?.fechaATEP) {
                    try {
                      fechaATEP = toDDMMYYYY(excelSerialToUTCDate(r.fechaATEP));
                    } catch (e) {
                      const msg = `[AVISO] No se pudo convertir la Fecha AT/EP del registro ${indice + 1}: ${e?.message || e}`;
                      console.warn("[FIE_2]", msg);
                      // Aviso pero NO lo consideramos error de registro ni lo llevamos al log de Excel
                    }
                  }

                  if (!fechaATEP) {
                    // Caso normal: muchos registros no tienen Fecha AT/EP.
                    // Sólo lo dejamos como debug para quien quiera activarlo.
                    logDebug(
                      `[FIE_2] Registro ${indice + 1} (NAF: ${
                        r?.naf ?? "sin NAF"
                      }) sin Fecha AT/EP en la hoja 2; se continúa sin rellenar #fechaATEP.`,
                    );
                  }

                  const code = tipoContratoIn.trim();
                  const starts = (p) => code.startsWith(p);

                  let tipoContratoSelect = "";
                  if (starts("2") || starts("3") || starts("5")) {
                    tipoContratoSelect = "1"; // Fijo discontinuo / Tiempo parcial
                  } else if (starts("1") || starts("4")) {
                    tipoContratoSelect = "2"; // Resto
                  }

                  if (DEBUG) {
                    console.table({
                      puestoDeTrabajo,
                      cnoe,
                      tipoContratoIn,
                      tipoContratoSelect,
                      baseResto,
                      diasResto,
                      baseFijoParcial,
                      diasFijoParcial,
                      fechaATEP,
                    });
                  }

                  if (puestoDeTrabajo) {
                    await fillTextWithRetry(
                      form2,
                      "#puestoTrabajo",
                      puestoDeTrabajo,
                      {
                        tries: 3,
                        typeDelay: 35,
                        digitsOnlyCompare: false,
                      },
                    );
                  }

                  if (cnoe) {
                    await selectWithRetry(form2, "#ocupacion", cnoe);
                  }

                  await selectWithRetry(
                    form2,
                    "#tipoContrato",
                    tipoContratoSelect,
                  );
                  await pause(400);

                  if (tipoContratoSelect === "2") {
                    if (baseResto)
                      await fillTextWithRetry(form2, "#BaseCot", baseResto, {
                        tries: 3,
                        typeDelay: 35,
                        digitsOnlyCompare: false,
                      });
                    if (diasResto)
                      await fillTextWithRetry(form2, "#DiasCot", diasResto, {
                        tries: 3,
                        typeDelay: 35,
                      });
                  } else {
                    if (baseFijoParcial)
                      await fillTextWithRetry(
                        form2,
                        "#sumaBaseCot",
                        baseFijoParcial,
                        {
                          tries: 3,
                          typeDelay: 35,
                          digitsOnlyCompare: false,
                        },
                      );
                    if (diasFijoParcial)
                      await fillTextWithRetry(
                        form2,
                        "#sumaDiasCot",
                        diasFijoParcial,
                        {
                          tries: 3,
                          typeDelay: 35,
                        },
                      );
                  }

                  if (
                    !(await fillIfPresent(form2, "#fechaATEP", fechaATEP, {
                      tries: 3,
                      typeDelay: 35,
                      digitsOnlyCompare: false,
                    }))
                  ) {
                    logDebug(
                      "[FIE_2] #fechaATEP ausente o sin valor. Continúo sin error.",
                    );
                  }

                  if (detalleCnoe) {
                    await fillTextWithRetry(
                      form2,
                      "#funcDesempe",
                      detalleCnoe,
                      {
                        tries: 3,
                        typeDelay: 15,
                        digitsOnlyCompare: false,
                      },
                    );
                  }

                  try {
                    await form2.waitForSelector("#ENVIO_14", {
                      visible: true,
                      timeout: 8000,
                    });
                    await form2.click("#ENVIO_14", { delay: 60 });
                    logDebug("[FIE_2] Click en Validar (ENVIO_14).");
                  } catch (e) {
                    console.warn(
                      "[FIE_2] No se pudo clicar Validar:",
                      e?.message || e,
                    );
                  }
                }
              } catch (e) {
                console.warn(
                  "[FIE_2] Error en segunda pantalla:",
                  e?.message || e,
                );
              }

              await this.esperar(1000);

              // === Pantalla de confirmación (Confirmar) ===
              try {
                await Promise.race([
                  page
                    .waitForNavigation({
                      waitUntil: "domcontentloaded",
                      timeout: 15000,
                    })
                    .catch(() => {}),
                  pause(1500),
                ]);

                const confirmFrame1 =
                  (await findFrameWithSelector("#ENVIO_12", 20000, 400)) ||
                  (await findFrameWithSelector(
                    'button[name="SPM.ACC.CONFIRMAR_DATOS_ECONOMICOS"]',
                    20000,
                    400,
                  ));

                if (!confirmFrame1) {
                  const msg =
                    "[ERROR] No se encontró la pantalla de Confirmación (botón #ENVIO_12).";
                  console.warn("[FIE_2]", msg);
                  appendLog(nifNormRegistro, msg);
                  registrosError++;
                  return;
                } else {
                  try {
                    await confirmFrame1.waitForSelector("#ENVIO_12", {
                      visible: true,
                      timeout: 8000,
                    });
                    await confirmFrame1.click("#ENVIO_12", { delay: 60 });
                    logDebug("[FIE_2] Click en Confirmar (ENVIO_12).");
                  } catch (e) {
                    console.warn(
                      "[FIE_2] No se pudo clicar Confirmar (ENVIO_12):",
                      e?.message || e,
                    );
                  }
                }
              } catch (e) {
                console.warn(
                  "[FIE_2] Error en pantalla de confirmación:",
                  e?.message || e,
                );
              }

              await this.esperar(1000);

              // === Pantalla de generación (Generar informe) ===
              try {
                await Promise.race([
                  page
                    .waitForNavigation({
                      waitUntil: "domcontentloaded",
                      timeout: 15000,
                    })
                    .catch(() => {}),
                  pause(1500),
                ]);

                const confirmFrame2 =
                  (await findFrameWithSelector("#ENVIO_8", 20000, 400)) ||
                  (await findFrameWithSelector(
                    'button[name="SPM.ACC.INFORME_DATOS_ECONOMICOS"]',
                    20000,
                    400,
                  ));

                if (!confirmFrame2) {
                  const msg =
                    "[ERROR] No se encontró la pantalla de Generación (botón #ENVIO_8).";
                  console.warn("[FIE_2]", msg);
                  appendLog(nifNormRegistro, msg);
                  registrosError++;
                  return;
                } else {
                  try {
                    await confirmFrame2.waitForSelector("#ENVIO_8", {
                      visible: true,
                      timeout: 8000,
                    });
                    await confirmFrame2.click("#ENVIO_8", { delay: 60 });
                    logDebug("[FIE_2] Click en Generar (ENVIO_8).");
                  } catch (e) {
                    console.warn(
                      "[FIE_2] No se pudo clicar Generar (ENVIO_8):",
                      e?.message || e,
                    );
                  }
                }
              } catch (e) {
                console.warn(
                  "[FIE_2] Error en pantalla de Generación:",
                  e?.message || e,
                );
              }

              await this.esperar(1000);

              // === Enlace "Visualizar informe..." y descargar PDF ===
              try {
                await Promise.race([
                  page
                    .waitForNavigation({
                      waitUntil: "domcontentloaded",
                      timeout: 15000,
                    })
                    .catch(() => {}),
                  pause(1500),
                ]);

                const docFrame = await findFrameWithSelector(
                  'a.pr_enlaceDocInforme[href*="ViewDocUtf8"]',
                  20000,
                  400,
                );

                if (!docFrame) {
                  const msg =
                    "[ERROR] No se encontró el enlace de informe (a.pr_enlaceDocInforme).";
                  console.warn("[FIE_2]", msg);
                  appendLog(nifNormRegistro, msg);
                } else if (!pathSalidaPDFConfirmacion) {
                  console.warn(
                    "[FIE_2] No hay carpeta de salida configurada; no descargo PDF.",
                  );
                } else {
                  const href = await docFrame.$eval(
                    'a.pr_enlaceDocInforme[href*="ViewDocUtf8"]',
                    (el) => el.getAttribute("href") || "",
                  );

                  if (!href) {
                    const msg =
                      "[ERROR] El enlace de informe no tiene href usable.";
                    console.warn("[FIE_2]", msg);
                    appendLog(nifNormRegistro, msg);
                  } else {
                    const baseUrl = page.url();
                    const pdfUrl = new URL(href, baseUrl).toString();
                    logDebug("[FIE_2] URL PDF:", pdfUrl);

                    const pdfBase64 = await docFrame.evaluate(async (url) => {
                      const res = await fetch(url, { credentials: "include" });
                      if (!res.ok) {
                        throw new Error(
                          "Respuesta HTTP no OK al descargar PDF: " +
                            res.status,
                        );
                      }
                      const buf = await res.arrayBuffer();
                      const bytes = new Uint8Array(buf);
                      let binary = "";
                      for (let i = 0; i < bytes.length; i++) {
                        binary += String.fromCharCode(bytes[i]);
                      }
                      return btoa(binary);
                    }, pdfUrl);

                    const buffer = Buffer.from(pdfBase64, "base64");

                    const seqMatch = pdfUrl.match(/[?&]SECUENCIAL=(\d+)/);
                    const seq = (seqMatch && seqMatch[1]) || "1";

                    const nafSafe = (r?.naf ? String(r.naf) : "sinNAF").replace(
                      /\D/g,
                      "",
                    );

                    const fileName = `Informe_Datos_Economicos_${nafSafe}_S${seq}.pdf`;
                    const fullPath = path.join(
                      pathSalidaPDFConfirmacion,
                      fileName,
                    );

                    fs.writeFileSync(fullPath, buffer);
                    console.log("[FIE_2] Informe PDF guardado en:", fullPath);
                  }
                }
              } catch (e) {
                const msg = `[ERROR] Error al localizar/descargar el informe PDF: ${e?.message || e}`;
                console.warn("[FIE_2]", msg);
                appendLog(nifNormRegistro, msg);
              }

              // Si no se ha registrado ningún mensaje de error, marcamos como OK por NIF
              if (nifNormRegistro && !logsPorNif.has(nifNormRegistro)) {
                logsPorNif.set(nifNormRegistro, "OK");
              }
              registrosOk++;
            }; // fin procesarRegistro

            // === Bucle sobre todos los registros del Excel ===
            for (let i = 0; i < datos.length; i++) {
              registrosProcesados += 1;
              try {
                await procesarRegistro(datos[i], i);
              } catch (e) {
                registrosError++;
                const msg = `[ERROR] Error inesperado procesando el registro ${
                  i + 1
                }/${datos.length}: ${e?.message || e}`;
                console.warn("[FIE_2]", msg);

                const nifNorm = normalizaNif(obtenerNifRegistro(datos[i]));
                appendLog(nifNorm, msg);
              }
            }

            console.log(
              `[FIE_2] Proceso completado. Registros OK: ${registrosOk}, con errores: ${registrosError}.`,
            );

            // === Generar copia del Excel de entrada con columna de LOG al principio ===
            try {
              // Carpeta donde dejar el Excel con log
              const carpetaExcelSalida =
                pathSalidaPDFConfirmacion ||
                (pathSalidaBase && path.normalize(pathSalidaBase)) ||
                path.dirname(rutaNormalizada);

              const nombreOriginal = path.basename(pathArchivoFIE_2);
              const nombreSinExt = nombreOriginal.replace(/\.xlsx?$/i, "");
              const nombreCopia = `${nombreSinExt} - LOG.xlsx`;
              const pathCopia = path.join(carpetaExcelSalida, nombreCopia);

              console.log("[FIE_2] Generando copia con LOG en:", pathCopia);

              // Cargamos de nuevo el Excel original para no tocar el archivo de entrada
              const workbookCopia =
                await XlsxPopulate.fromFileAsync(rutaNormalizada);
              const sheet = workbookCopia.sheet(0);

              // --- 1) Detectar dinámicamente la FILA de cabecera (la que tiene "EXPTE") ---
              let filaCabecera = 1;
              let encontradaCabecera = false;

              for (let r = 1; r <= 20 && !encontradaCabecera; r++) {
                for (let c = 1; c <= 40 && !encontradaCabecera; c++) {
                  const val = sheet.cell(r, c).value();
                  if (
                    typeof val === "string" &&
                    val.toLowerCase().includes("expte")
                  ) {
                    filaCabecera = r;
                    encontradaCabecera = true;
                  }
                }
              }

              console.log(
                "[FIE_2] Fila de cabecera detectada en:",
                filaCabecera,
              );

              // --- 2) Calcular cuántas columnas contiguas tiene la cabecera ---
              let numColumnas = 0;
              for (let col = 1; col <= 200; col++) {
                const val = sheet.cell(filaCabecera, col).value();
                if (val === null || val === undefined || val === "") break;
                numColumnas = col;
              }

              if (numColumnas === 0) {
                console.warn(
                  "[FIE_2] No se detectaron columnas en la fila de cabecera al generar el LOG. Se omite inserción.",
                );
              } else {
                // --- 3) Localizar columnas de EXPTE y NIF en la cabecera ---
                let colExpte = null;
                let colNif = null;

                for (let col = 1; col <= numColumnas; col++) {
                  const val = sheet.cell(filaCabecera, col).value();
                  if (typeof val === "string") {
                    const low = val.toLowerCase();
                    if (low.includes("expte") && colExpte === null)
                      colExpte = col;
                    if (low.includes("nif") && colNif === null) colNif = col;
                  }
                }

                if (colExpte == null || colNif == null) {
                  console.warn(
                    "[FIE_2] No se pudieron localizar las columnas de EXPTE o NIF al generar el LOG.",
                  );
                } else {
                  // --- 4) Determinar la última fila de datos usando el nº de registros del Excel ---
                  // datosHoja1.length = nº de registros originales de la hoja 0
                  const totalRegistrosExcel = datosHoja1.length;
                  const ultimaFilaDatos = filaCabecera + totalRegistrosExcel;

                  console.log(
                    "[FIE_2] Última fila de datos calculada en:",
                    ultimaFilaDatos,
                  );

                  // --- 5) Simular "Insertar columna A": desplazar todas las columnas a la derecha
                  //      desde la fila 1 hasta la última fila de datos
                  const filas = sheet.usedRange()._numRows;
                  const columnas = sheet.usedRange()._numColumns;
                  for (let fila = 1; fila <= filas; fila++) {
                    for (let col = columnas; col >= 1; col--) {
                      const valor = sheet.cell(fila, col).value();
                      sheet.cell(fila, col + 1).value(valor);
                    }
                  }

                  // Después de desplazar, la columna NIF pasa a ser colNif + 1
                  const colNifInsertada = colNif + 1;

                  // --- 6) Escribir cabecera de LOG en la columna A de la fila de cabecera ---
                  sheet.cell(filaCabecera, 1).value("LOG FIE_2");

                  // --- 7) Volcar los logs según el NIF de cada fila ---
                  for (
                    let fila = filaCabecera + 1;
                    fila <= ultimaFilaDatos;
                    fila++
                  ) {
                    const nifValor = sheet.cell(fila, colNifInsertada).value();
                    const nifNorm = normalizaNif(nifValor);
                    const log = nifNorm ? logsPorNif.get(nifNorm) || "" : "";
                    sheet.cell(fila, 1).value(log);
                  }
                }
              }

              // Guardamos la copia
              await workbookCopia.toFileAsync(path.normalize(pathCopia));
              console.log(
                "[FIE_2] Copia del Excel con LOG generada correctamente.",
              );
            } catch (e) {
              console.warn(
                "[FIE_2] No se pudo generar la copia del Excel con LOG:",
                e?.message || e,
              );
            }
          }
        } catch (navErr) {
          console.warn(
            "[FIE_2] Aviso: no se pudo abrir el navegador/URL de FS:",
            navErr?.message || navErr,
          );
          return resolve(false);
        } finally {
          if (browser) {
            try {
              await browser.close();
            } catch (_) {}
          }
        }

        registrarEjecucion({
          nombreProceso,
          registrosProcesados: registrosProcesados,
        });

        return resolve(datos);
      } catch (err) {
        console.error("[FIE_2] Error general en el proceso:", err);
        try {
          if (globalThis?.mainProcess?.mostrarError) {
            await globalThis.mainProcess.mostrarError(
              "No se ha podido completar el proceso",
              "Se ha producido un error interno ejecutando FIE_2.",
            );
          }
        } catch (_) {}
        return resolve(false);
      }
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

    // ⬇⬇⬇ NUEVO: saltar filas completamente vacías ⬇⬇⬇
    const hayDatos = Object.values(objetoRegistro).some(
      (v) => v !== undefined && v !== null && v !== "",
    );
    if (!hayDatos) {
      continue; // no añadimos este registro
    }
    // ⬆⬆⬆ FIN NUEVO ⬆⬆⬆

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

//Funcion obtener el mes de una fecha:
function obtenerMesFecha(fecha) {
  const dateObj = excelSerialToUTCDate(fecha);
  return dateObj.getUTCMonth() + 1; // Los meses en JavaScript son 0-indexados
}

function obtenerPrimeroDeMes(date) {
  return new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), 1));
}

function obtenerUltimoDeMes(d) {
  return new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth() + 1, 0));
}

function obtenerNombreMesByIndex(mesIndex) {
  switch (mesIndex) {
    case 0:
      return "Enero";
    case 1:
      return "Febrero";
    case 2:
      return "Marzo";
    case 3:
      return "Abril";
    case 4:
      return "Mayo";
    case 5:
      return "Junio";
    case 6:
      return "Julio";
    case 7:
      return "Agosto";
    case 8:
      return "Septiembre";
    case 9:
      return "Octubre";
    case 10:
      return "Noviembre";
    case 11:
      return "Diciembre";
    default:
      return null;
  }
}

module.exports = ProcesosFie;
