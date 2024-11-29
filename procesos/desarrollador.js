const path = require("path");
const fs = require("fs");
const readline = require('readline')
const moment = require("moment");
const XlsxPopulate = require("xlsx-populate");
const Datastore = require("nedb");
const _= require("lodash");

class ProcesosDesarrollador {
	
	constructor(pathToDbFolder, nombreProyecto, proyectoDB){
		this.pathToDbFolder = pathToDbFolder;
		this.nombreProyecto = nombreProyecto; 
		this.proyectoDB = proyectoDB;
	}
 	
	async esperar(tiempo){
		return new Promise((resolve)=>{
			setTimeout(resolve, tiempo);
		});
	}


	async unirCarpetaExcel(argumentos){

		console.log("Uniendo Excel en carpeta: ");
		console.log("Ruta entrada: "+argumentos[0])
		console.log("Ruta salida: "+argumentos[1])
		console.log("Nombre salida: "+argumentos[2])
		
		const pathSpoolInput = path.join(argumentos[0]);
		var pathSpoolOutput;	
		pathSpoolOutput = path.join(argumentos[1],argumentos[2]);
		/*
		if(argumentos[2].slice(-4) !== ".txt" &&  argumentos[2].slice(-4) !== ".TXT"){
			pathSpoolOutput = path.join(argumentos[1],argumentos[2]+".txt");
		}else{
			pathSpoolOutput = path.join(argumentos[1],argumentos[2]);
		}
		*/

		const readline = require('readline')

		async function contarArchivos(){
			return new Promise((resolve) => {

				var cuentaArchivos=0;
				var listaArchivos = [];
				
				//Leer los nombres de los archivos en la base de datos:
				fs.readdir(argumentos[0], (err, files) => {
					if (err) {
						console.log(
							"Se ha producido un error leyendo los archivos del directorio: "+rutaCarpeta
						);
						console.log(err);
						resolve(false)
					} else {
						var totalArchivosCargar = 0;
						var cuentaArchivosCargados = 0;
						var archivoProyecto;

						//Excluir archivos:
						var auxFiles = files.slice();

						files.forEach((file,index, array) => {
								if(file.indexOf(".")==-1){
										auxFiles.splice(auxFiles.indexOf(file),1);
								}
						});

						//Si no hay archivos para cargar:
						console.log("Analizando AUX:")
						console.log(auxFiles)

						totalArchivosCargar= auxFiles.length;	
							if(totalArchivosCargar==0){
								console.log("Lista de archivos: ");
								console.log(auxFiles);	
								resolve([]);
							}else{
								console.log("Lista de archivos: ");
								console.log(auxFiles);	
								resolve(auxFiles);
							}
					}
				});
			})
		}

		async function excelToJson(){
			return new Promise((resolve) => {
					XlsxPopulate.fromBlankAsync()
						.then(workbook => {
							console.log("Archivo Cargado: Seguimiento");
							archivoExcel = workbook;
							return true;
						})
						.then(()=>{
							archivoExcel.toFileAsync(path.normalize(path.join(argumentos[1],argumentos[2]+"3.xlsx")));
							resolve(true);
						})
						.catch(err => {
							console.log("ERROR");
							resolve(false);
						})
			})
		}

		async function fusionarArchivos(lista){
			return new Promise((resolve) => {
				
			var archivoExcel;

			XlsxPopulate.fromBlankAsync()
				.then(workbook => {
					console.log("Archivo Cargado: Seguimiento");
					archivoExcel = workbook;
					return true;
				})
				.then(()=>{

					//var resultado= await excelToJson();
					//archivoExcel.toFileAsync(path.normalize(path.join(argumentos[1],argumentos[2]+".xlsx")));

					resolve(true);
				})
				.catch(err => {
					console.log("Se ha producido un error interno: ");
					console.log(err);
					var tituloError =
						"Se ha producido un error interno cargando los archivos.";
					mainWindow.webContents.send(
						"onErrorInterno",
						tituloError,
						err
					);
					resolve(false)
				});
			});
		}

		var lista = await contarArchivos();
		console.log(lista)
		var result = await fusionarArchivos(lista);
		return true;
	}
	
	//********************************
	//  Procesar Report AM
	//********************************

	async generarSeguimientoAM(argumentos) {
		
		console.log("EJECUTANDO PROCESADO AM");

		var datosNacho = argumentos[1][0];
		var pathArchivoSeguimiento = argumentos[0];

		var archivoSeguimiento= {};

		var configuracion = {
			añoInicioControl: argumentos[2],
			mesInicioControl: argumentos[3],
			añoFinControl: argumentos[4],
			mesFinControl: argumentos[5],
			datosSalidaControl: argumentos[6],
			nombreArchivoSalidaControl: argumentos[7] 
		}

	async function generarReportAM(archivoSeguimiento, datosNacho, configuracion) {
		return new Promise((resolve) => {

		//PROCESAMIENTO: 
		XlsxPopulate.fromFileAsync(path.normalize(pathArchivoSeguimiento))
			.then(workbook => {
				console.log("Archivo Cargado: Seguimiento");
				archivoSeguimiento = workbook;
				return true;
			})
			.then(()=>{

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

		var columnasFotoAnteriorSeguimiento = archivoSeguimiento .sheet("Foto_semana anterior") .usedRange()._numColumns;
		var filasFotoAnteriorSeguimiento = archivoSeguimiento .sheet("Foto_semana anterior") .usedRange()._numRows;

		//Limpia hoja de seguimiento anterior:
		for (var i = 1; i < filasFotoAnteriorSeguimiento; i++) {
			for (var j = 0; j < columnasFotoAnteriorSeguimiento; j++) {
				archivoSeguimiento .sheet("Foto_semana anterior") .row(i + 1) .cell(j + 1) .clear();
			}
		}

		// 2) Mover Foto a semana anterior
		archivoSeguimiento.sheet("Foto_semana anterior").name("Provisional");
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
			resolve(false)
		}

		var cabeceraSeleccionada = "";

		for (var i = 1; i < datosNacho.data.length; i++) {
			for (var j = 0; j < columnasFotoAnteriorSeguimiento; j++) {

				cabeceraSeleccionada = String(archivoSeguimiento.sheet("Foto").row(1).cell(j + 1).value());
				cabeceraSeleccionada = cabeceraSeleccionada.toLowerCase();
				cabeceraSeleccionada = cabeceraSeleccionada.replace( / /g, "_");

				if (cabeceraSeleccionada == "hub") {
					cabeceraSeleccionada = "categoria_3";
				}

				//console.log(cabeceraSeleccionada)
				if (cabeceraSeleccionada !== undefined) {
					if (datosNacho.data[i-1][cabeceraSeleccionada] !== undefined) {
						archivoSeguimiento .sheet("Foto") .row(i + 1) .cell(j + 1) .value(datosNacho.data[i-1][cabeceraSeleccionada]);
					} else {
						console.log( "Warning: Dato no encontrado i=" + i + " j=" + j + " Cabecera: " + cabeceraSeleccionada);
					}
				} else {
					console.log( "Warning de cabecera: i=" + i + " j=" + j);
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
						.value()
				);
				cabeceraSeleccionada = cabeceraSeleccionada.toLowerCase();
				cabeceraSeleccionada = cabeceraSeleccionada.replace(
					/ /g,
					"_"
				);

				//console.log(cabeceraSeleccionada)
				if (cabeceraSeleccionada === undefined) {
					console.log(
						"Error de cabecera: i=" + i + " j=" + j
					);
				} else {
					switchProcesado: switch (cabeceraSeleccionada) {
						case "fecha_creacion":
							datoProcesado = archivoSeguimiento
								.sheet("Foto")
								.row(i + 1)
								.cell(9)
								.value();
							if (
								datoProcesado.indexOf(
									"/"
								) != -1
							) {
								datoProcesado = moment(
									datoProcesado,
									"DD/MM/YYYY"
								);
							} else {
								datoProcesado = moment(
									datoProcesado,
									"YYYY-MM-DD"
								);
							}

							datoProcesado =
								datoProcesado.date() +
								"/" +
								(datoProcesado.month() +
									1) +
								"/" +
								datoProcesado.year();

							archivoSeguimiento
								.sheet("Foto")
								.row(i + 1)
								.cell(j + 1)
								.value(
									String(
										datoProcesado
									)
								);
							//console.log(datoProcesado)
							break;

						case "week_creacion":
							datoProcesado = archivoSeguimiento
								.sheet("Foto")
								.row(i + 1)
								.cell(9)
								.value();
							//datoProcesado = datoProcesado.replace(/\//g,"-");
							if (
								datoProcesado.indexOf(
									"/"
								) != -1
							) {
								datoProcesado = moment(
									datoProcesado,
									"DD/MM/YYYY"
								);
							} else {
								datoProcesado = moment(
									datoProcesado,
									"YYYY-MM-DD"
								);
							}

							datoProcesado = datoProcesado.week();

							archivoSeguimiento
								.sheet("Foto")
								.row(i + 1)
								.cell(j + 1)
								.value(
									String(
										datoProcesado
									)
								);
							//console.log(datoProcesado)
							break;

						case "fecha_cierre":
							datoProcesado = archivoSeguimiento
								.sheet("Foto")
								.row(i + 1)
								.cell(11)
								.value();
							if (
								datoProcesado ===
									undefined ||
								datoProcesado == "" ||
								datoProcesado._error ==
									"#N/A"
							) {
								archivoSeguimiento
									.sheet("Foto")
									.row(i + 1)
									.cell(j + 1)
									.clear();
								break;
							}
							//datoProcesado = datoProcesado.replace(/\//g,"-");
							if (
								datoProcesado.indexOf(
									"/"
								) != -1
							) {
								datoProcesado = moment(
									datoProcesado,
									"DD/MM/YYYY"
								);
							} else {
								datoProcesado = moment(
									datoProcesado,
									"YYYY-MM-DD"
								);
							}

							datoProcesado =
								datoProcesado.date() +
								"/" +
								(datoProcesado.month() +
									1) +
								"/" +
								datoProcesado.year();

							archivoSeguimiento
								.sheet("Foto")
								.row(i + 1)
								.cell(j + 1)
								.value(
									String(
										datoProcesado
									)
								);
							//console.log(datoProcesado)
							break;

						case "week_cierre":
							datoProcesado = archivoSeguimiento
								.sheet("Foto")
								.row(i + 1)
								.cell(11)
								.value();
							if (
								datoProcesado ===
									undefined ||
								datoProcesado == "" ||
								datoProcesado._error ==
									"#N/A"
							) {
								archivoSeguimiento
									.sheet("Foto")
									.row(i + 1)
									.cell(j + 1)
									.clear();
								break;
							}
							//datoProcesado = datoProcesado.replace(/\//g,"-");
							if (
								datoProcesado.indexOf(
									"/"
								) != -1
							) {
								datoProcesado = moment(
									datoProcesado,
									"DD/MM/YYYY"
								);
							} else {
								datoProcesado = moment(
									datoProcesado,
									"YYYY-MM-DD"
								);
							}

							datoProcesado = datoProcesado.week();

							archivoSeguimiento
								.sheet("Foto")
								.row(i + 1)
								.cell(j + 1)
								.value(
									String(
										datoProcesado
									)
								);
							//console.log(datoProcesado)
							break;

						case "proceso":
							datoProcesado = archivoSeguimiento
								.sheet("Foto")
								.row(i + 1)
								.cell(2)
								.value();

							if (
								datoProcesado ===
									undefined ||
								datoProcesado == "" ||
								datoProcesado._error ==
									"#N/A"
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
							for (
								var k = 1;
								k <
								filasFotoAnteriorSeguimiento;
								k++
							) {
								if (
									indiceBusqueda >
									filasFotoAnteriorSeguimiento
								) {
									indiceBusqueda =
										indiceBusqueda -
										filasFotoAnteriorSeguimiento;
								}

								if (
									datoProcesado ===
									archivoSeguimiento
										.sheet(
											"Foto_semana anterior"
										)
										.row(
											indiceBusqueda +
												1
										)
										.cell(2)
										.value()
								) {
									datoProcesado = archivoSeguimiento
										.sheet(
											"Foto_semana anterior"
										)
										.row(
											indiceBusqueda +
												1
										)
										.cell(
											31
										)
										.value();
									indiceBusquedaAnterior = indiceBusqueda;
									valorEncontrado = true;
									break;
								}
								indiceBusqueda++;
							}

							if (
								datoProcesado ===
									undefined ||
								datoProcesado == "" ||
								datoProcesado._error ==
									"#N/A" ||
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
								.value(
									String(
										datoProcesado
									)
								);
							//console.log(JSON.stringify(datoProcesado))

							break switchProcesado;

						case "subproceso":
							datoProcesado = archivoSeguimiento
								.sheet("Foto")
								.row(i + 1)
								.cell(2)
								.value();

							if (
								datoProcesado ===
									undefined ||
								datoProcesado == "" ||
								datoProcesado._error ==
									"#N/A"
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
							for (
								var k = 1;
								k <
								filasFotoAnteriorSeguimiento;
								k++
							) {
								if (
									indiceBusqueda >
									filasFotoAnteriorSeguimiento
								) {
									indiceBusqueda =
										indiceBusqueda -
										filasFotoAnteriorSeguimiento;
								}

								if (
									datoProcesado ===
									archivoSeguimiento
										.sheet(
											"Foto_semana anterior"
										)
										.row(
											indiceBusqueda +
												1
										)
										.cell(2)
										.value()
								) {
									datoProcesado = archivoSeguimiento
										.sheet(
											"Foto_semana anterior"
										)
										.row(
											indiceBusqueda +
												1
										)
										.cell(
											32
										)
										.value();
									indiceBusquedaAnterior = indiceBusqueda;
									valorEncontrado = true;
									break;
								}
								indiceBusqueda++;
							}

							if (
								datoProcesado ===
									undefined ||
								datoProcesado == "" ||
								datoProcesado._error ==
									"#N/A" ||
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
								.value(
									String(
										datoProcesado
									)
								);
							//console.log(JSON.stringify(datoProcesado))

							break switchProcesado;

						case "ultimos_15_días":
							datoProcesado = archivoSeguimiento
								.sheet("Foto")
								.row(i + 1)
								.cell(9)
								.value();

							if (
								datoProcesado.indexOf(
									"/"
								) != -1
							) {
								datoProcesado = moment(
									datoProcesado,
									"DD/MM/YYYY"
								);
							} else {
								datoProcesado = moment(
									datoProcesado,
									"YYYY-MM-DD"
								);
							}

							if (
								datoProcesado >=
								moment().subtract(
									15,
									"days"
								)
							) {
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
				neto: 0
			},
			rmca: {
				entrada: 0,
				salida: 0,
				cerradas: 0,
				canceladas: 0,
				backlog: 0,
				neto: 0
			},
			sap: {
				entrada: 0,
				salida: 0,
				cerradas: 0,
				canceladas: 0,
				backlog: 0,
				neto: 0
			},
			editran: {
				entrada: 0,
				salida: 0,
				cerradas: 0,
				canceladas: 0,
				backlog: 0,
				neto: 0
			},
			connectDirect: {
				entrada: 0,
				salida: 0,
				cerradas: 0,
				canceladas: 0,
				backlog: 0,
				neto: 0
			}
		};

		var añoInicio = configuracion.añoInicioControl;
		var mesInicio = configuracion.mesInicioControl-1;
				
		var añoFin = configuracion.añoFinControl;
		var mesFin = configuracion.mesFinControl-1;

		var mesActual = moment().month();
		var añoActual = moment().year();

		console.log("Año actual: " + añoActual);
		console.log("Mes actual: " + mesActual);

		//Verificación de año:
		if (añoInicio > añoFin) {
			añoInicio = añoFin;
		}

		var filasFotoSeguimiento = archivoSeguimiento.sheet("Foto").usedRange()
			._numRows;

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
					datos: _.cloneDeep(resultadosBase)
				});

				console.log(JSON.stringify(registroResultados));
				//ITERAR SISTEMA:
				for (var sistema in registroResultados[
					registroResultados.length - 1
				].datos) {
					switch (sistema) {
						case "general":
							for (var estado in registroResultados[
								registroResultados.length -
									1
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
											"VFES-RMCA-INFRASTRUCTURE-PROD"
										];
										filtroTipoIncidente = [
											"Incident",
											"User Service Restoration"
										];
										filtroEstado = [
											"Assigned",
											"Cancelled",
											"Closed",
											"In Progress",
											"Pending",
											"Resolved"
										];
										filtroEnOtraCola = [
											undefined
										];
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
											"VFES-RMCA-INFRASTRUCTURE-PROD"
										];
										filtroTipoIncidente = [
											"Incident",
											"User Service Restoration"
										];
										filtroEstado = [
											"Cancelled",
											"Closed"
										];
										filtroEnOtraCola = [
											undefined
										];
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
											"VFES-RMCA-INFRASTRUCTURE-PROD"
										];
										filtroTipoIncidente = [
											"Incident",
											"User Service Restoration"
										];
										filtroEstado = [
											"Closed"
										];
										filtroEnOtraCola = [
											undefined
										];
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
											"VFES-RMCA-INFRASTRUCTURE-PROD"
										];
										filtroTipoIncidente = [
											"Incident",
											"User Service Restoration"
										];
										filtroEstado = [
											"Cancelled"
										];
										filtroEnOtraCola = [
											undefined
										];
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
											"VFES-RMCA-INFRASTRUCTURE-PROD"
										];
										filtroTipoIncidente = [
											"Incident",
											"User Service Restoration"
										];
										filtroEstado = [
											"Assigned",
											"In Progress",
											"Pending",
											"Resolved"
										];
										filtroEnOtraCola = [
											undefined
										];
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
											"VFES-RMCA-INFRASTRUCTURE-PROD"
										];
										filtroTipoIncidente = [
											"Incident",
											"User Service Restoration"
										];
										filtroEstado = [
											"Assigned",
											"Cancelled",
											"Closed",
											"In Progress",
											"Pending",
											"Resolved"
										];
										filtroEnOtraCola = [
											undefined
										];
										break;
								}
								registroResultados[
									registroResultados.length -
										1
								].datos[sistema][
									estado
								] = procesarRecuentoSeguimientoAM(
									sistema,
									estado,
									i,
									j,
									filtroServicio,
									filtroTipoIncidente,
									filtroEstado,
									filtroEnOtraCola
								);
							}
							break;
						case "rmca":
							for (const estado in registroResultados[
								registroResultados.length -
									1
							].datos[sistema]) {
								switch (estado) {
									case "entrada":
										filtroServicio = [
											"VFES-RMCA-PROD",
											"VFES-RMCA-INFRASTRUCTURE-PROD"
										];
										filtroTipoIncidente = [
											"Incident",
											"User Service Restoration"
										];
										filtroEstado = [
											"Assigned",
											"Cancelled",
											"Closed",
											"In Progress",
											"Pending",
											"Resolved"
										];
										filtroEnOtraCola = [
											undefined
										];
										break;
									case "salida":
										filtroServicio = [
											"VFES-RMCA-PROD",
											"VFES-RMCA-INFRASTRUCTURE-PROD"
										];
										filtroTipoIncidente = [
											"Incident",
											"User Service Restoration"
										];
										filtroEstado = [
											"Cancelled",
											"Closed"
										];
										filtroEnOtraCola = [
											undefined
										];
										break;
									case "cerradas":
										filtroServicio = [
											"VFES-RMCA-PROD",
											"VFES-RMCA-INFRASTRUCTURE-PROD"
										];
										filtroTipoIncidente = [
											"Incident",
											"User Service Restoration"
										];
										filtroEstado = [
											"Closed"
										];
										filtroEnOtraCola = [
											undefined
										];
										break;
									case "canceladas":
										filtroServicio = [
											"VFES-RMCA-PROD",
											"VFES-RMCA-INFRASTRUCTURE-PROD"
										];
										filtroTipoIncidente = [
											"Incident",
											"User Service Restoration"
										];
										filtroEstado = [
											"Cancelled"
										];
										filtroEnOtraCola = [
											undefined
										];
										break;
									case "backlog":
										filtroServicio = [
											"VFES-RMCA-PROD",
											"VFES-RMCA-INFRASTRUCTURE-PROD"
										];
										filtroTipoIncidente = [
											"Incident",
											"User Service Restoration"
										];
										filtroEstado = [
											"Assigned",
											"In Progress",
											"Pending",
											"Resolved"
										];
										filtroEnOtraCola = [
											undefined
										];
										break;
									case "neto":
										filtroServicio = [
											"VFES-RMCA-PROD",
											"VFES-RMCA-INFRASTRUCTURE-PROD"
										];
										filtroTipoIncidente = [
											"Incident",
											"User Service Restoration"
										];
										filtroEstado = [
											"Assigned",
											"Cancelled",
											"Closed",
											"In Progress",
											"Pending",
											"Resolved"
										];
										filtroEnOtraCola = [
											undefined
										];
										break;
								}
								registroResultados[
									registroResultados.length -
										1
								].datos[sistema][
									estado
								] = procesarRecuentoSeguimientoAM(
									sistema,
									estado,
									i,
									j,
									filtroServicio,
									filtroTipoIncidente,
									filtroEstado,
									filtroEnOtraCola
								);
							}
							break;
						case "sap":
							for (var estado in registroResultados[
								registroResultados.length -
									1
							].datos[sistema]) {
								switch (estado) {
									case "entrada":
										filtroServicio = [
											"VFES-SAP 4.7 SGCYR-PROD",
											"VFES-SAP 4.7-PROD"
										];
										filtroTipoIncidente = [
											"Incident",
											"User Service Restoration"
										];
										filtroEstado = [
											"Assigned",
											"Cancelled",
											"Closed",
											"In Progress",
											"Pending",
											"Resolved"
										];
										filtroEnOtraCola = [
											undefined
										];
										break;
									case "salida":
										filtroServicio = [
											"VFES-SAP 4.7 SGCYR-PROD",
											"VFES-SAP 4.7-PROD"
										];
										filtroTipoIncidente = [
											"Incident",
											"User Service Restoration"
										];
										filtroEstado = [
											"Cancelled",
											"Closed"
										];
										filtroEnOtraCola = [
											undefined
										];
										break;
									case "cerradas":
										filtroServicio = [
											"VFES-SAP 4.7 SGCYR-PROD",
											"VFES-SAP 4.7-PROD"
										];
										filtroTipoIncidente = [
											"Incident",
											"User Service Restoration"
										];
										filtroEstado = [
											"Closed"
										];
										filtroEnOtraCola = [
											undefined
										];
										break;
									case "canceladas":
										filtroServicio = [
											"VFES-SAP 4.7 SGCYR-PROD",
											"VFES-SAP 4.7-PROD"
										];
										filtroTipoIncidente = [
											"Incident",
											"User Service Restoration"
										];
										filtroEstado = [
											"Cancelled"
										];
										filtroEnOtraCola = [
											undefined
										];
										break;
									case "backlog":
										filtroServicio = [
											"VFES-SAP 4.7 SGCYR-PROD",
											"VFES-SAP 4.7-PROD"
										];
										filtroTipoIncidente = [
											"Incident",
											"User Service Restoration"
										];
										filtroEstado = [
											"Assigned",
											"In Progress",
											"Pending",
											"Resolved"
										];
										filtroEnOtraCola = [
											undefined
										];
										break;
									case "neto":
										filtroServicio = [
											"VFES-SAP 4.7 SGCYR-PROD",
											"VFES-SAP 4.7-PROD"
										];
										filtroTipoIncidente = [
											"Incident",
											"User Service Restoration"
										];
										filtroEstado = [
											"Assigned",
											"Cancelled",
											"Closed",
											"In Progress",
											"Pending",
											"Resolved"
										];
										filtroEnOtraCola = [
											undefined
										];
										break;
								}
								registroResultados[
									registroResultados.length -
										1
								].datos[sistema][
									estado
								] = procesarRecuentoSeguimientoAM(
									sistema,
									estado,
									i,
									j,
									filtroServicio,
									filtroTipoIncidente,
									filtroEstado,
									filtroEnOtraCola
								);
							}
							break;
						case "editran":
							for (var estado in registroResultados[
								registroResultados.length -
									1
							].datos[sistema]) {
								switch (estado) {
									case "entrada":
										filtroServicio = [
											"VFES-EDITRAN-PROD",
											"VFES-EDITRAN",
											"VFES-ONO-EDITRAN BANKS-PROD",
											"VFES-EDITRAN. BANKS-PROD"
										];
										filtroTipoIncidente = [
											"Incident",
											"User Service Restoration"
										];
										filtroEstado = [
											"Assigned",
											"Cancelled",
											"Closed",
											"In Progress",
											"Pending",
											"Resolved"
										];
										filtroEnOtraCola = [
											undefined
										];
										break;
									case "salida":
										filtroServicio = [
											"VFES-EDITRAN-PROD",
											"VFES-EDITRAN",
											"VFES-ONO-EDITRAN BANKS-PROD",
											"VFES-EDITRAN. BANKS-PROD"
										];
										filtroTipoIncidente = [
											"Incident",
											"User Service Restoration"
										];
										filtroEstado = [
											"Cancelled",
											"Closed"
										];
										filtroEnOtraCola = [
											undefined
										];
										break;
									case "cerradas":
										filtroServicio = [
											"VFES-EDITRAN-PROD",
											"VFES-EDITRAN",
											"VFES-ONO-EDITRAN BANKS-PROD",
											"VFES-EDITRAN. BANKS-PROD"
										];
										filtroTipoIncidente = [
											"Incident",
											"User Service Restoration"
										];
										filtroEstado = [
											"Closed"
										];
										filtroEnOtraCola = [
											undefined
										];
										break;
									case "canceladas":
										filtroServicio = [
											"VFES-EDITRAN-PROD",
											"VFES-EDITRAN",
											"VFES-ONO-EDITRAN BANKS-PROD",
											"VFES-EDITRAN. BANKS-PROD"
										];
										filtroTipoIncidente = [
											"Incident",
											"User Service Restoration"
										];
										filtroEstado = [
											"Cancelled"
										];
										filtroEnOtraCola = [
											undefined
										];
										break;
									case "backlog":
										filtroServicio = [
											"VFES-EDITRAN-PROD",
											"VFES-EDITRAN",
											"VFES-ONO-EDITRAN BANKS-PROD",
											"VFES-EDITRAN. BANKS-PROD"
										];
										filtroTipoIncidente = [
											"Incident",
											"User Service Restoration"
										];
										filtroEstado = [
											"Assigned",
											"In Progress",
											"Pending",
											"Resolved"
										];
										filtroEnOtraCola = [
											undefined
										];
										break;
									case "neto":
										filtroServicio = [
											"VFES-EDITRAN-PROD",
											"VFES-EDITRAN",
											"VFES-ONO-EDITRAN BANKS-PROD",
											"VFES-EDITRAN. BANKS-PROD"
										];
										filtroTipoIncidente = [
											"Incident",
											"User Service Restoration"
										];
										filtroEstado = [
											"Assigned",
											"Cancelled",
											"Closed",
											"In Progress",
											"Pending",
											"Resolved"
										];
										filtroEnOtraCola = [
											undefined
										];
										break;
								}
								registroResultados[
									registroResultados.length -
										1
								].datos[sistema][
									estado
								] = procesarRecuentoSeguimientoAM(
									sistema,
									estado,
									i,
									j,
									filtroServicio,
									filtroTipoIncidente,
									filtroEstado,
									filtroEnOtraCola
								);
							}
							break;
						case "connectDirect":
							for (const estado in registroResultados[
								registroResultados.length -
									1
							].datos[sistema]) {
								switch (estado) {
									case "entrada":
										filtroServicio = [
											"VFES-SEPA CONNECT DIRECT-PROD",
											"VFES-SAP 4.7 CONNECT DIRECT-PROD",
											"ES-CONNECT DIRECT"
										];
										filtroTipoIncidente = [
											"Incident",
											"User Service Restoration"
										];
										filtroEstado = [
											"Assigned",
											"Cancelled",
											"Closed",
											"In Progress",
											"Pending",
											"Resolved"
										];
										filtroEnOtraCola = [
											undefined
										];
										break;
									case "salida":
										filtroServicio = [
											"VFES-SEPA CONNECT DIRECT-PROD",
											"VFES-SAP 4.7 CONNECT DIRECT-PROD",
											"ES-CONNECT DIRECT"
										];
										filtroTipoIncidente = [
											"Incident",
											"User Service Restoration"
										];
										filtroEstado = [
											"Cancelled",
											"Closed"
										];
										filtroEnOtraCola = [
											undefined
										];
										break;
									case "cerradas":
										filtroServicio = [
											"VFES-SEPA CONNECT DIRECT-PROD",
											"VFES-SAP 4.7 CONNECT DIRECT-PROD",
											"ES-CONNECT DIRECT"
										];
										filtroTipoIncidente = [
											"Incident",
											"User Service Restoration"
										];
										filtroEstado = [
											"Closed"
										];
										filtroEnOtraCola = [
											undefined
										];
										break;
									case "canceladas":
										filtroServicio = [
											"VFES-SEPA CONNECT DIRECT-PROD",
											"VFES-SAP 4.7 CONNECT DIRECT-PROD",
											"ES-CONNECT DIRECT"
										];
										filtroTipoIncidente = [
											"Incident",
											"User Service Restoration"
										];
										filtroEstado = [
											"Cancelled"
										];
										filtroEnOtraCola = [
											undefined
										];
										break;
									case "backlog":
										filtroServicio = [
											"VFES-SEPA CONNECT DIRECT-PROD",
											"VFES-SAP 4.7 CONNECT DIRECT-PROD",
											"ES-CONNECT DIRECT"
										];
										filtroTipoIncidente = [
											"Incident",
											"User Service Restoration"
										];
										filtroEstado = [
											"Assigned",
											"In Progress",
											"Pending",
											"Resolved"
										];
										filtroEnOtraCola = [
											undefined
										];
										break;
									case "neto":
										filtroServicio = [
											"VFES-SEPA CONNECT DIRECT-PROD",
											"VFES-SAP 4.7 CONNECT DIRECT-PROD",
											"ES-CONNECT DIRECT"
										];
										filtroTipoIncidente = [
											"Incident",
											"User Service Restoration"
										];
										filtroEstado = [
											"Assigned",
											"Cancelled",
											"Closed",
											"In Progress",
											"Pending",
											"Resolved"
										];
										filtroEnOtraCola = [
											undefined
										];
										break;
								}
								registroResultados[
									registroResultados.length -
										1
								].datos[sistema][
									estado
								] = procesarRecuentoSeguimientoAM(
									sistema,
									estado,
									i,
									j,
									filtroServicio,
									filtroTipoIncidente,
									filtroEstado,
									filtroEnOtraCola
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
			filtroEnOtraCola
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
				"diciembre"
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

				switch(estado){
					case "salida":
					case "cerradas":
					case "canceladas":
						if (archivoSeguimiento.sheet("Foto").row(i + 1).cell(16).value() == String(año)) {
							cumpleFiltro = true;
						}
						break;
					case "backlog":
						cumpleFiltro = true;
						break;
					default:
						if (archivoSeguimiento.sheet("Foto").row(i + 1).cell(14).value() == String(año)) {
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

				switch(estado){
					case "salida":
					case "cerradas":
					case "canceladas":
						if (archivoSeguimiento .sheet("Foto").row(i + 1).cell(17).value() == meses[mes]) {
							cumpleFiltro = true;
						}
						break;
					case "backlog":
						cumpleFiltro = true;
						break;
					default:
						if (archivoSeguimiento .sheet("Foto").row(i + 1).cell(15).value() == meses[mes]) {
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
							.value() ==
						filtroTipoIncidente[j]
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
					if (archivoSeguimiento.sheet("Foto").row(i + 1).cell(26).value() == filtroEnOtraCola[j]) {
						cumpleFiltro = true;
					}
				}


				//Iteración filtro Aplicativo (Solo Backlog):
				if(estado == "backlog"){

					if (cumpleFiltro) {
						cumpleFiltro = false;
					} else {
						continue;
					}

					var filtroAplicativo = [];

					switch(sistema){
						case "general":
							filtroAplicativo = ["ECC 6.0","SAP 4.7","Editran","Connect Direct"]
							break;

						case "rmca":
							filtroAplicativo = ["ECC 6.0"]
							break;

						case "sap":
							filtroAplicativo = ["SAP 4.7"]
							break;

						case "editran":
							filtroAplicativo = ["Editran"]
							break;

						case "connectDirect":
							filtroAplicativo = ["Connect Direct"]
							break;
					}

					for (var j = 0; j < filtroAplicativo.length; j++) {
						if (archivoSeguimiento.sheet("Foto").row(i + 1).cell(21).value() == filtroAplicativo[j]) {
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
					cuenta
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
			console.log("Mes: " + (registroResultados[i].mes+1));
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
							registroResultados[i].datos[
								sistema
							][estado]
					);
				}
			}
		}

		for (var i = 0; i < registroResultados.length; i++) {
			console.log("");
			console.log("--------------");
			console.log("Año: " + registroResultados[i].año);
			console.log("Mes: " + (registroResultados[i].mes+1));
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
		   var utc_days  = Math.floor(serial - 25569);
		   var utc_value = utc_days * 86400;
		   var date_info = new Date(utc_value * 1000); 
		   var fractional_day = serial - Math.floor(serial) + 0.0000001;
		   var total_seconds = Math.floor(86400 * fractional_day);
		   var seconds = total_seconds % 60;
		   total_seconds -= seconds;
		   var hours = Math.floor(total_seconds / (60 * 60));
		   var minutes = Math.floor(total_seconds / 60) % 60;

		   return new Date(date_info.getFullYear(), date_info.getMonth(), date_info.getDate(), hours, minutes, seconds);
		}
	
		function isValidDate(d) {
		  return d instanceof Date && !isNaN(d);
		}

		var hojasResumen = ["Resumen General","Resumen RMCA","Resumen SAP 4.7","Resumen Editran","Resumen CD"]
		var sistemas = ["general","rmca","sap","editran","connectDirect"]

		//Iteración por hojas resumen:
		for(var k=0; k<hojasResumen.length; k++){
			//Inicializacion:
			ultimaColumna=0;

			// Detección de ultima columna
			for(var i=0; i<archivoSeguimiento.sheet(hojasResumen[k]).usedRange()._numColumns;i++){
				if(isValidDate(new Date(ExcelDateToJSDate(archivoSeguimiento.sheet(hojasResumen[k]).row(1).cell(i+1).value())))){
					ultimaColumna= i+1;	
					añoUltimaColumna= new Date(ExcelDateToJSDate(archivoSeguimiento.sheet(hojasResumen[k]).row(1).cell(i+1).value())).getFullYear()
					mesUltimaColumna= new Date(ExcelDateToJSDate(archivoSeguimiento.sheet(hojasResumen[k]).row(1).cell(i+1).value())).getMonth()
				}
			}

			console.log("Ultima Columna" + ultimaColumna)
			console.log("Año: " + añoUltimaColumna);
			console.log("Mes: " + mesUltimaColumna);

			if((añoUltimaColumna==añoFin) && (mesUltimaColumna==mesFin)){

				archivoSeguimiento.sheet(hojasResumen[k]).row(2).cell(ultimaColumna).value(registroResultados[registroResultados.length-1].datos[sistemas[k]]["entrada"]);
				archivoSeguimiento.sheet(hojasResumen[k]).row(3).cell(ultimaColumna).value(registroResultados[registroResultados.length-1].datos[sistemas[k]]["salida"]);
				archivoSeguimiento.sheet(hojasResumen[k]).row(4).cell(ultimaColumna).value(registroResultados[registroResultados.length-1].datos[sistemas[k]]["cerradas"]);
				archivoSeguimiento.sheet(hojasResumen[k]).row(5).cell(ultimaColumna).value(registroResultados[registroResultados.length-1].datos[sistemas[k]]["canceladas"]);
				archivoSeguimiento.sheet(hojasResumen[k]).row(6).cell(ultimaColumna).value(registroResultados[registroResultados.length-1].datos[sistemas[k]]["backlog"]);
				archivoSeguimiento.sheet(hojasResumen[k]).row(7).cell(ultimaColumna).value(registroResultados[registroResultados.length-1].datos[sistemas[k]]["neto"]);
			}

			if((añoUltimaColumna==añoFin) && (mesUltimaColumna==mesFin-1)){

				console.log("Modificando Columna: "+registroResultados[registroResultados.length-1].mes);	

				archivoSeguimiento.sheet(hojasResumen[k]).row(2).cell(ultimaColumna).value(registroResultados[registroResultados.length-2].datos[sistemas[k]]["entrada"]);
				archivoSeguimiento.sheet(hojasResumen[k]).row(3).cell(ultimaColumna).value(registroResultados[registroResultados.length-2].datos[sistemas[k]]["salida"]);
				archivoSeguimiento.sheet(hojasResumen[k]).row(4).cell(ultimaColumna).value(registroResultados[registroResultados.length-2].datos[sistemas[k]]["cerradas"]);
				archivoSeguimiento.sheet(hojasResumen[k]).row(5).cell(ultimaColumna).value(registroResultados[registroResultados.length-2].datos[sistemas[k]]["canceladas"]);
				archivoSeguimiento.sheet(hojasResumen[k]).row(6).cell(ultimaColumna).value(registroResultados[registroResultados.length-2].datos[sistemas[k]]["backlog"]);
				archivoSeguimiento.sheet(hojasResumen[k]).row(7).cell(ultimaColumna).value(registroResultados[registroResultados.length-2].datos[sistemas[k]]["neto"]);

				archivoSeguimiento.sheet(hojasResumen[k]).row(1).cell(ultimaColumna+1).value(new Date(añoFin, mesFin, 1)).style("numberFormat","mmm-yy");

				archivoSeguimiento.sheet(hojasResumen[k]).row(2).cell(ultimaColumna+1).value(registroResultados[registroResultados.length-1].datos[sistemas[k]]["entrada"]);
				archivoSeguimiento.sheet(hojasResumen[k]).row(3).cell(ultimaColumna+1).value(registroResultados[registroResultados.length-1].datos[sistemas[k]]["salida"]);
				archivoSeguimiento.sheet(hojasResumen[k]).row(4).cell(ultimaColumna+1).value(registroResultados[registroResultados.length-1].datos[sistemas[k]]["cerradas"]);
				archivoSeguimiento.sheet(hojasResumen[k]).row(5).cell(ultimaColumna+1).value(registroResultados[registroResultados.length-1].datos[sistemas[k]]["canceladas"]);
				archivoSeguimiento.sheet(hojasResumen[k]).row(6).cell(ultimaColumna+1).value(registroResultados[registroResultados.length-1].datos[sistemas[k]]["backlog"]);
				archivoSeguimiento.sheet(hojasResumen[k]).row(7).cell(ultimaColumna+1).value(registroResultados[registroResultados.length-1].datos[sistemas[k]]["neto"]);
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
						configuracion.nombreArchivoSalidaControl +
							".xlsx"
					)
				)
		);
		archivoSeguimiento
			.toFileAsync(
				path.normalize(
					path.join(
						configuracion.datosSalidaControl,
						configuracion.nombreArchivoSalidaControl +
							".xlsx"
					)
				)
			)
			.then(() => {
				console.log("Fin del procesamiento");
				resolve(true)
			})
			.catch(err => {
				console.log("Se ha producido un error interno: ");
				console.log(err);
				var tituloError =
					"Se ha producido un error escribiendo el archivo: " +
					path.normalize(
						path.join(
							configuracion.datosSalidaControl,
							configuracion.nombreArchivoSalidaControl +
								".xlsx"
						)
					);
				resolve(false)
			});

		resolve(true);
	})
			})
			.catch(err => {
				console.log("Se ha producido un error interno: ");
				console.log(err);
				var tituloError =
					"Se ha producido un error interno cargando los archivos.";
				mainWindow.webContents.send(
					"onErrorInterno",
					tituloError,
					err
				);
				resolve(false)
			});
	}
	
	var resultado= await generarReportAM(archivoSeguimiento, datosNacho, configuracion); 
	return resultado;	

	}; //Fin de generación de report AM

	async fusionarObjetos(argumentos){

		console.log("Fusionar Archivos:");
		var archivoBase = argumentos[0][0];
		var archivoAdd = argumentos[1][0];

		async function fusionarArchivos(archivoBase,archivoAdd){
			return new Promise((resolve) => {
				
			for(var i = 0; i<archivoAdd.data.length; i++){
				archivoBase.data.push(archivoAdd.data[i]);
			}
				resolve(archivoBase);
			})
		}
		
		var result = await fusionarArchivos(archivoBase, archivoAdd);
		console.log("LOGITUD FINAL: "+ result.data.length);
		result["objetoId"]= archivoBase.nombreId
		return result;
	}

	async procesarIBAN(argumentos){

		console.log("Procesando Recuento IBAN - Mandato");

		var rutaGuardado = argumentos[1];
		var nombreGuardado = argumentos[2];

		console.log("TAMANO DATOS ARVIVO 1: " + argumentos[0][0].data.length)	
		console.log("TAMANO DATOS ARVIVO 2: " + argumentos[1][0].data.length)	
		console.log("TAMANO DATOS ARVIVO 3: " + argumentos[2][0].data.length)	
		console.log("TAMANO DATOS ARVIVO 4: " + argumentos[3][0].data.length)	
		console.log("TAMANO DATOS ARVIVO 5: " + argumentos[4][0].data.length)	
		console.log("TAMANO DATOS ARVIVO 6: " + argumentos[5][0].data.length)	
		console.log("TAMANO DATOS ARVIVO 7: " + argumentos[6][0].data.length)	
		console.log("TAMANO DATOS ARVIVO 8: " + argumentos[7][0].data.length)	
		console.log("TAMANO DATOS ARVIVO 9: " + argumentos[8][0].data.length)	
		console.log("TAMANO DATOS ARVIVO 10: " + argumentos[9][0].data.length)	
		console.log("TAMANO DATOS ARVIVO 11: " + argumentos[10][0].data.length)	
		console.log("TAMANO DATOS ARVIVO 12: " + argumentos[11][0].data.length)	
		console.log("TAMANO DATOS ARVIVO 13: " + argumentos[12][0].data.length)	
		console.log("TAMANO DATOS ARVIVO 14: " + argumentos[13][0].data.length)	
		console.log("TAMANO DATOS ARVIVO 15: " + argumentos[14][0].data.length)	
		var suma = 0;
		for(var i=0; i<1;i++){
			suma += argumentos[i][0].data.length
		}

		console.log("Tamaña total: " + suma) 

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
		for(var iteracionDocumento = 0; iteracionDocumento<15; iteracionDocumento++){
			for(var iteracionRegistro = 0; iteracionRegistro< argumentos[iteracionDocumento][0].data.length; iteracionRegistro++){
				try{
				matrizIban.push(argumentos[iteracionDocumento][0].data[iteracionRegistro]["iban"])	

				}catch{
					console.log("Error IBAN; Documento: "+iteracionDocumento+" Registro: "+iteracionRegistro)
				}
			}
		}
		console.log("Matriz finalizada");

		const countOccurrences = arr => arr.reduce((prev,curr) => (prev[curr] = ++prev[curr] || 1, prev), {});

		//Ordenando matriz:
		console.log("Ordenando Matriz: ");
		matrizIban = matrizIban.sort();

		console.log("Calculando ocurrencias");
		objetoSalida= countOccurrences(matrizIban.sort())

		//Formateando objeto salida;
		console.log("Depurando salida");
		for(const property in objetoSalida){
			if(objetoSalida[property] === 1){
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

		if(argumentos[16].slice(-4) !== ".txt" &&  argumentos[16].slice(-4) !== ".TXT"){
			pathSpoolOutput = path.join(argumentos[15],argumentos[16]+".txt");
		}else{
			pathSpoolOutput = path.join(argumentos[15],argumentos[16]);
		}

		const outputFile = fs.createWriteStream(pathSpoolOutput)

			outputFile.on('err', err => {
				// handle error
				console.log(err)
			})

			outputFile.on('close', () => { 
				console.log('done writing')
			})
		for(const property in objetoSalida){
			outputFile.write(`${property}\t${objetoSalida[property]}\n`)
		}

		console.log("FIN de procesamiento IBAN");
		var result = true; 
		return result;
	}

	async procesarSMS(argumentos){

		console.log("Procesando SMS");
		console.log("Archivo entrada: "+argumentos[0])
		console.log("Archivo salida: "+argumentos[1])
		
		const pathSpoolInput = path.join(argumentos[0]);
		var pathSpoolOutput;	

		if(argumentos[2].slice(-4) !== ".txt" &&  argumentos[2].slice(-4) !== ".TXT"){
			pathSpoolOutput = path.join(argumentos[1],argumentos[2]+".txt");

		}else{
			pathSpoolOutput = path.join(argumentos[1],argumentos[2]);
		}

		const readline = require('readline')
		const outputFile = fs.createWriteStream(pathSpoolOutput,{flags:"a"})

		async function contarRegistros(){
			return new Promise((resolve) => {
			var cuentaRegistros=0;

			const rl = readline.createInterface({
				input: fs.createReadStream(pathSpoolInput)
			})

			rl.on('line', line => {
				cuentaRegistros++;
			})

			rl.on('close', () => {
				console.log("Numero de registros: "+cuentaRegistros);
				resolve(cuentaRegistros);
			})	
			})
		}

		async function leerSpool(registrosTotalesProcesar){
			return new Promise((resolve) => {

			var cuentaRegistroProcesado = 0;

			const rl = readline.createInterface({
				input: fs.createReadStream(pathSpoolInput)
			})

			// Handle any error that occurs on the write stream
			outputFile.on('err', err => {
				// handle error
				console.log(err)
			})

			outputFile.on('close', () => { 
				console.log('done writing')
			})

			var fechaAnterior = "";
			var fechaActual = "";
			var motivo = "";
			var estado = "";
			var indent = "";

			//Variables de recuento:
			var recuentoEstado4 = 0;
			var recuentoEstado5 = 0;

			var recuentoMotivoNoMovil = 0;
			var recuentoMotivoBlanco = 0;
			var recuentoMotivoClienteSinFinanciacion = 0;
			var recuentoMotivoSinInformar = 0;
			var recuentoMotivoImporteInferior = 0;
			var recuentoMotivoConReclamacion = 0;
			var recuentoMotivoOtros = 0;

			rl.on('line', line => {
				let text = line
				
				var indexInicioFecha = 0;
				var indexFinalFecha = 0; 

				var indexInicioIndent = 0;
				var indexFinalIndent = 0;

				var indexInicioMotivo = 0;
				var indexFinalMotivo = 0;

				var indexInicioEstado = 0;
				var indexFinalEstado = 0;
				
				var indexCuentaTabFecha = 2;
				var indexCuentaTabIndent = 3;
				var indexCuentaTabMotivo = 11;
				var indexCuentaTabEstado = 10;
				
				function numberOfTabs(text) {
					var count = 0;

					for( var i= 0; i<text.length; i++){
						if(text.charAt(i) === "|"){
							count++

							//Fija inicio y fin de Fecha:
							if(count == indexCuentaTabFecha){
								indexInicioFecha= i;
							}
							if(count== indexCuentaTabFecha+1){
								indexFinalFecha= i;
							}

							//Fija inicio y fin de Indent:
							if(count == indexCuentaTabIndent){
								indexInicioIndent= i;
							}
							if(count== indexCuentaTabIndent+1){
								indexFinalIndent= i;
							}

							//Fija inicio y fin de Motivo:
							if(count == indexCuentaTabMotivo){
								indexInicioMotivo= i;
							}
							if(count== indexCuentaTabMotivo+1){
								indexFinalMotivo= i;
							}

							//Fija inicio y fin de Estado:
							if(count == indexCuentaTabEstado){
								indexInicioEstado= i;
							}
							if(count== indexCuentaTabEstado+1){
								indexFinalEstado= i;
							}
						}
					}
					return count;
				}

				if(cuentaRegistroProcesado%10000==0){
					console.log("Progreso: "+(cuentaRegistroProcesado/registrosTotalesProcesar*100));
				}

				numberOfTabs(text);
				cuentaRegistroProcesado++; 

				//Extraccion de datos:
				fechaActual = text.substring(indexInicioFecha+1,indexFinalFecha).trim();
				motivo = text.substring(indexInicioMotivo+1,indexFinalMotivo).trim();
				indent = text.substring(indexInicioIndent+1,indexFinalIndent).trim();
				estado = text.substring(indexInicioEstado+1,indexFinalEstado).trim();
				
				// Realiza el proceso de recuento:
				
				// Escribe resultados y Resetea el contador si cambia de fecha:
				if(fechaAnterior != fechaActual){

				//Skip si fecha es incorrecta:
				if(fechaAnterior == "-"||fechaAnterior==""||fechaAnterior=="ID fecha"){
					recuentoEstado4 = 0;
					recuentoEstado5 = 0;
					recuentoMotivoNoMovil = 0;
					recuentoMotivoBlanco = 0;
					recuentoMotivoClienteSinFinanciacion = 0;
					recuentoMotivoSinInformar = 0;
					recuentoMotivoImporteInferior = 0;
					recuentoMotivoConReclamacion = 0;
					recuentoMotivoOtros = 0;
					fechaAnterior = fechaActual;
					return;
				}

				//Skip si indent = PP
				if(indent=="PP"){
					fechaAnterior = fechaActual;
					return;
				}


				outputFile.write(`${fechaAnterior}|${recuentoEstado4}|${recuentoEstado5}|${recuentoMotivoNoMovil}|${recuentoMotivoBlanco}|${recuentoMotivoClienteSinFinanciacion}|${recuentoMotivoSinInformar}|${recuentoMotivoImporteInferior}|${recuentoMotivoConReclamacion}|${recuentoMotivoOtros}\n`)

					recuentoEstado4 = 0;
					recuentoEstado5 = 0;
					recuentoMotivoNoMovil = 0;
					recuentoMotivoBlanco = 0;
					recuentoMotivoClienteSinFinanciacion = 0;
					recuentoMotivoSinInformar = 0;
					recuentoMotivoImporteInferior = 0;
					recuentoMotivoConReclamacion = 0;
					recuentoMotivoOtros = 0;

					fechaAnterior = fechaActual;
					return;
				}

				var motivoContado = false;

				// Realiza en recuento:
				if(estado == "04"){
					recuentoEstado4++;
					motivoContado = true;
				}

				if(estado == "05"){
					recuentoEstado5++;
					motivoContado = true;
				}

				if(motivo.includes("no tiene m")){
					recuentoMotivoNoMovil++;
					motivoContado = true;
				}

				if(motivo.includes("con reclamaci")){
					recuentoMotivoConReclamacion++;
					motivoContado = true;
				}

				if(motivo.includes("Cliente sin financiaci")){
					recuentoMotivoClienteSinFinanciacion++;
					motivoContado = true;
				}

				if(motivo.includes("importe reclamado inferior al m")){
					recuentoMotivoImporteInferior++;
					motivoContado = true;
				}

				if(motivo.includes("El tipo de mensaje est")){
					recuentoMotivoSinInformar++;
					motivoContado = true;
				}

				if(motivo==""){
					recuentoMotivoBlanco++;
					motivoContado = true;
				}
				
				if(!motivoContado){
					recuentoMotivoOtros++;
				}

				//Actualiza la última Fecha
				fechaAnterior = fechaActual;

				return;
			})

			// Done reading the input, call end() on the write stream
			rl.on('close', () => {
				console.log("FIN DEL PROCESAMIENTO");
				outputFile.end()
				resolve(true);
			})	
			})
		}

			var cuentaRegistros = await contarRegistros();
			var result = await leerSpool(cuentaRegistros);
			return result;
	}
}

module.exports = ProcesosDesarrollador;


