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

class ProcesosPrueba {
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

  async testExcel(argumentos) {
    return new Promise((resolve) => {
      console.log("Leyendo excel de prueba");
      console.log(argumentos.formularioControl[0]);
      console.log("Ruta de salida");
      console.log(argumentos.formularioControl[1]);


      var pathSalida = path.join(
        path.normalize(argumentos.formularioControl[1]),
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
       console.log("Esto es una prueba");
          resolve(true);
      } catch (err) {
        var tituloError = "No se ha podido cargar el archivo";
        var mensajeError =
          "Se ha producido un error interno cargando los archivos.";
        mainProcess.mostrarError(tituloError, mensajeError).then((result) => {
          resolve(false);
        });
      }
      console.log("Test excel funciona de locos")
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


} //Fin Procesos Asesoria

module.exports = ProcesosPrueba;
