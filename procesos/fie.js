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

  async fie(argumentos) {
    return new Promise((resolve) => {
      console.log("Procesamiento de FIE...");
      console.log(argumentos.formularioControl[1]);

      console.log("Ruta Google...");
      console.log(argumentos.formularioControl[0]);

      var archivoFIE = {};

      var pathArchivoFIE = argumentos.formularioControl[1];
      var pathSalidaExcel = path.join(
        path.normalize(argumentos.formularioControl[2]),
        "Fie-Procesado",
      );
      var pathSalida = path.join(
        path.normalize(argumentos.formularioControl[2]),
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

      //ALMACENAMIENTO TEMPORAL:
      var altas = [];
      var bajas = [];
      var confirmacion = [];

      try {
        XlsxPopulate.fromFileAsync(path.normalize(pathArchivoFIE))
          .then(async (workbook) => {
            console.log("Archivo Cargado: FIE");
            archivoFIE = workbook;

            var datosIncapacidad = extraccionExcel(workbook, 0, 3, 1);
            var partesConfirmacion = extraccionExcel(workbook, 0, 4, 1);

            console.log("Datos incapacidad:");
            console.log(datosIncapacidad[0]);

            console.log("Partes confirmacion:");
            console.log(partesConfirmacion);

            //ESCRITURA XLSX:
            console.log("Escribiendo archivo...");
            console.log("Path: " + path.normalize(pathSalidaExcel));

            archivoFIE
              .toFileAsync(
                path.normalize(
                  path.join(pathSalidaExcel, "FIE-Procesado.xlsx"),
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
