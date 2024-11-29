
const path = require("path");
const fs = require("fs");
const readline = require('readline')
const XlsxPopulate = require("xlsx-populate");
const ipcRenderer= require("electron").ipcRenderer;
const ipc = require("electron").ipcMain;
const mainProcess = require("../main.js");

class ProcesosImport {
    
    constructor(pathToDbFolder, nombreProyecto, proyectoDB){

        this.pathToDbFolder = pathToDbFolder;
        this.nombreProyecto = nombreProyecto; 
        this.proyectoDB = proyectoDB;

    }

    async importarExcel(argumentos){
        console.log("Importando Excel:");
        console.log("Argumentos: ");
        console.log(argumentos);

        var rutaArchivo = argumentos[0];
        var numFilaCabecera = argumentos[1];
        var nombreHoja = argumentos[2];
        var nombreObjeto = argumentos[3];

        //PROCESAMIENTO: 
        //Importando XLSX:
        return new Promise((resolve) => {
        XlsxPopulate.fromFileAsync(path.normalize(argumentos[0]))
            .then(workbook => {
                console.log("Cargando Excel:");
                //console.log(workbook);
                //
                if(workbook===undefined){
                    resolve(false)
                }

                //Creación del Objeto:
                var objeto = [{
                    data: [],
                    nombreId: nombreObjeto,
                    objetoId: nombreObjeto,
                }]

                var cabeceraSeleccionada = "";

                var numeroCaberas = workbook
                    .sheet(nombreHoja)
                    .usedRange()._numColumns;

                var numeroRegistros = workbook
                    .sheet(nombreHoja)
                    .usedRange()._numRows;

                //Comprobación de inputs:
                if(workbook.sheet(nombreHoja)===undefined){
                    resolve(false)
                }


                //Relleno de objeto Data:
                for (var i = numFilaCabecera; i < numeroRegistros; i++) {
                    objeto[0].data.push({});
                    for (var j = 0; j < numeroCaberas; j++) {

                        cabeceraSeleccionada = String(
                            workbook
                                .sheet(nombreHoja)
                                .row(numFilaCabecera)
                                .cell(j + 1)
                                .value()
                        );

                        cabeceraSeleccionada = cabeceraSeleccionada.toLowerCase();
                        cabeceraSeleccionada = cabeceraSeleccionada.replace(/ /g, "_");
                        cabeceraSeleccionada = cabeceraSeleccionada.replace(/\./g, "_");

                        if (cabeceraSeleccionada === undefined || cabeceraSeleccionada === "") {
                            console.log(
                                "Error de cabecera: i=" + i + " j=" + j
                            );
                            continue;
                        }else{
                            //Guardado del registro:
                            if(workbook.sheet(nombreHoja).row(i + 1).cell(j+1).value()!==undefined){
                                objeto[0].data[objeto[0].data.length-1][cabeceraSeleccionada]= workbook.sheet(nombreHoja).row(i + 1).cell(j+1).value();
                            }
                                
                        }
                    }

                    if(Object.keys(objeto[0].data[objeto[0].data.length-1]).length === 0){
                        objeto[0].data.pop();
                    }
                }

                mainProcess.guardarDocumento(objeto[0])
                console.log(objeto)

            })
            .then(()=>{
                console.log("Proceso finalizado");
                resolve(true)
                })
        })
    }

    async importarCursos(argumentos){

        console.log("Importando Excel:");
        console.log("Argumentos: ");
        console.log(argumentos);

        var rutaArchivo = argumentos[0];
        var numFilaCabecera = argumentos[1];

        var nombreHojas = ["Cursos","Formador-Curso","Formadores","Codigos Provincia","Instituciones","Correos","Tipología"];
        var nombreObjetos = ["Cursos","Formador-Curso","Formadores","Códigos_Provincia","Instituciones","Correos","Tipología"];

        //PROCESAMIENTO: 
        //Importando XLSX:
        return new Promise((resolve) => {

        if(argumentos[0] === undefined || argumentos[0] ==="" || argumentos[1] === undefined){
            console.log("Error en argumentos");
            resolve(false)
        }

        XlsxPopulate.fromFileAsync(path.normalize(rutaArchivo))
            .then(workbook => {
                console.log("Cargando Excel:");
                if(workbook===undefined){
                    resolve(false)
                }

                //Iteración por Hojas:
                for(var k = 0; k < nombreHojas.length; k++){

                console.log("Procesando Hoja "+nombreHojas[k]+"...")

                //Creación del Objeto:
                var objeto = [{
                    data: [],
                    nombreId: nombreObjetos[k],
                    objetoId: nombreObjetos[k],
                }]

                var cabeceraSeleccionada = "";

                var numeroCaberas = workbook
                    .sheet(nombreHojas[k])
                    .usedRange()._numColumns;

                var numeroRegistros = workbook
                    .sheet(nombreHojas[k])
                    .usedRange()._numRows;

                //Comprobación de inputs:
                if(workbook.sheet(nombreHojas[k])===undefined){
                    resolve(false)
                }

                //Relleno de objeto Data:
                for (var i = numFilaCabecera; i < numeroRegistros; i++) {
                    objeto[0].data.push({});
                    for (var j = 0; j < numeroCaberas; j++) {

                        cabeceraSeleccionada = String(
                            workbook
                                .sheet(nombreHojas[k])
                                .row(numFilaCabecera)
                                .cell(j + 1)
                                .value()
                        );

                        cabeceraSeleccionada = cabeceraSeleccionada.toLowerCase();
                        cabeceraSeleccionada = cabeceraSeleccionada.replace(/ /g, "_");
                        cabeceraSeleccionada = cabeceraSeleccionada.replace(/\./g, "_");

                        if (cabeceraSeleccionada === undefined || cabeceraSeleccionada === "") {
                            console.log(
                                "Error de cabecera: i=" + i + " j=" + j
                            );
                            continue;
                        }else{
                            //Guardado del registro:
                            if(workbook.sheet(nombreHojas[k]).row(i + 1).cell(j+1).value()!==undefined){
                                objeto[0].data[objeto[0].data.length-1][cabeceraSeleccionada]= workbook.sheet(nombreHojas[k]).row(i + 1).cell(j+1).value();
                            }
                                
                        }
                    }

                    if(Object.keys(objeto[0].data[objeto[0].data.length-1]).length === 0){
                        objeto[0].data.pop();
                    }
                }

                mainProcess.guardarDocumento(objeto[0])
                //console.log(objeto)
                } //Fin For Documentos:
            })
            .then(()=>{
                console.log("Proceso finalizado");
                resolve(true)
                })
        })
    }
} 

module.exports = ProcesosImport;


