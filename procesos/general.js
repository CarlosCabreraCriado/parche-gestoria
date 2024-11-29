const path = require("path");
const fs = require("fs");
const readline = require('readline')
const moment = require("moment");
const XlsxPopulate = require("xlsx-populate");
const Datastore = require("nedb");
const _= require("lodash");
const {ipcRenderer}= require("electron");
const puppeteer = require('puppeteer');

class ProcesosGenerales {
	
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
	
	async extraccionRemedy(argumentos){

		console.log("Extracción Remedy");
		//console.log("Archivo entrada: "+argumentos[0])
		//console.log("Archivo salida: "+argumentos[1])

		const browser = await puppeteer.launch();
		const page = await browser.newPage();

		//(async () => {
			await page.goto('https://oneitsm.onbmc.com/arsys/forms/onbmc-s/SHR%3ALandingConsole/Default+Administrator+View/?cacheid=864a674f');
			await page.screenshot({ path: '/Users/carloscabreracriado/Desktop/prueba.png' });
			await browser.close();
		//})();

		return true;	
	}

	async compensarSpool(argumentos){

		console.log("Formatear SPOOL");
		console.log("Archivo entrada: "+argumentos[0])
		console.log("Archivo salida: "+argumentos[1])
		
		const pathSpoolInput = path.join(argumentos[0]);
		const pathCompensadaInput1 = path.join(argumentos[1]);
		const pathCompensadaInput2 = path.join(argumentos[2]);
		const pathCompensadaInput3 = path.join(argumentos[3]);

		var pathSpoolOutput;	

		if(argumentos[5].slice(-4) !== ".txt" &&  argumentos[5].slice(-4) !== ".TXT"){
			pathSpoolOutput = path.join(argumentos[4],argumentos[5]+".txt");

		}else{
			pathSpoolOutput = path.join(argumentos[4],argumentos[5]);
		}

		const readline = require('readline')
		const outputFile = fs.createWriteStream(pathSpoolOutput)

		var arrayDocumentos = [];
		var arrayDatos = [];

		async function crearArray(pathArray){
			return new Promise((resolve) => {
			var cuentaRegistros=0;

			const rl = readline.createInterface({
				input: fs.createReadStream(pathArray)
			})

			rl.on('line', line => {
				let text = line
				arrayDocumentos.push(parseInt(text.substring(0,text.indexOf("\t")+1)))
				arrayDatos.push((' ' + text).slice(1));
				cuentaRegistros++;
			})

			rl.on('close', () => {
				console.log("Registros añadidos a array: " +cuentaRegistros);
				resolve(true)
			})	

			})
		}

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

		var registrosEncontrados = 0; 
		async function compensar(registrosTotalesProcesar){
			return new Promise((resolve) => {
			const rl = readline.createInterface({
				input: fs.createReadStream(pathSpoolInput)
			})

			outputFile.on('err', err => {
				// handle error
				console.log(err)
			})

			outputFile.on('close', () => { 
				console.log('done writing')
			})

			var cuentaRegistroProcesado = 0; 

			rl.on('line', line => {
				let text = line

				var numDoc = parseInt(text.substring(4, 16));
			
				cuentaRegistroProcesado++;

				//Cuenta numero de tabs:
				var count = (text.match(/\t/g) || []).length;
				
				for(var i=count; i<17; i++){
					text = text+ "\t";
				}
				
				for(var i = 0; i< arrayDocumentos.length; i++){
					if(numDoc === arrayDocumentos[i]){
						text = text+"\t"+arrayDatos[i];
						break;
					}
				}

				if(cuentaRegistroProcesado%10000==0){
					console.log("Progreso: "+(cuentaRegistroProcesado/registrosTotalesProcesar*100));
				}

				outputFile.write(`${text}\n`)
			})

			rl.on('close', () => {
				console.log("FIN DEL PROCESAMIENTO");
				outputFile.end()
				resolve(true);
			})	
			})
		}
			await crearArray(pathCompensadaInput1);
			await crearArray(pathCompensadaInput2);
			await crearArray(pathCompensadaInput3);	

			var numRegistros = await contarRegistros();
			var result = await compensar(numRegistros);

			return result;
	}

	async filtrarFechaSpool(argumentos){

		console.log("Filtrar Fecha SPOOL");
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
		const outputFile = fs.createWriteStream(pathSpoolOutput)

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

			rl.on('line', line => {
				let text = line
				
				var indexInicioFecha = 0;
				var indexFinalFecha = 0; 
				var indexCuentaTab = 23;
				
				function numberOfTabs(text) {
					var count = 0;

					for( var i= 0; i<text.length; i++){
						if(text.charAt(i) === "\t"){
							count++
							if(count == indexCuentaTab){
								indexInicioFecha= i;
							}
							if(count== indexCuentaTab+1){
								indexFinalFecha= i;
							}
						}
					}
					//Si es el ultimo tabulador:
					if(indexCuentaTab == count){
						indexFinalFecha = text.length;	
					}
					/*
					console.log("INDEX INICIO: "+indexInicioFecha);
					console.log("INDEX Fin: "+indexFinalFecha);
					*/
					return count;
				}

				if(cuentaRegistroProcesado%10000==0){
					console.log("Progreso: "+(cuentaRegistroProcesado/registrosTotalesProcesar*100));
				}

				numberOfTabs(text);
				cuentaRegistroProcesado++; 

				//Verifica que la fecha se ajusta con el filtro:
				var day = moment(text.substring(indexInicioFecha,indexFinalFecha),"DD.MM.YYYY");

				if(!day.isValid()){
					return;
				}

				if(day.isAfter(moment('01.11.2020',"DD.MM.YYYY"))){
					return;
				}

				outputFile.write(`${text}\n`)

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

	async formatearSpool(argumentos){

		console.log("Formatear SPOOL");
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
		const outputFile = fs.createWriteStream(pathSpoolOutput)

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

			// Once done writing, rename the output to be the input file name
			outputFile.on('close', () => { 
				console.log('done writing')

				/*fs.rename(pathSpoolOutput, pathSpoolInput, err => {
					if (err) {
					  // handle error
					  console.log(err)
					} else {
					  console.log('renamed file')
					}
				})*/ 
			})

			// Read the file and replace any text that matches

			rl.on('line', line => {
				let text = line

				// Elimina las lineas que no comienzan por tabulador:
				if (!text.startsWith('\t')) {
					return;
				}				

				function numberOfTabs(text) {
					var count = 0;
					for( var i= 0; i<text.length; i++){
						if(text.charAt(i) === "\t"){
							count++
						}
					}
					return count;
				}

				if(cuentaRegistroProcesado%10000==0){
					console.log("Progreso: "+(cuentaRegistroProcesado/registrosTotalesProcesar*100));
				}

				cuentaRegistroProcesado++; 
/*				
				if(numberOfTabs(text)<3){
					return;
				}
*/
				// Elimina las lineas que comienzan por "Md.":
				// Elimina las lineas que comienzan por "N.":
				if (text.startsWith('\tMd.')) {
					return;
				}
				if (text.startsWith('\tN')) {
					return;
				}

				text = text.substr(1);

				outputFile.write(`${text}\n`)
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

	async obtenerObjetoDocumentoSpool(argumentos){

		console.log("Obtener Objeto Documento:");
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
		const outputFile = fs.createWriteStream(pathSpoolOutput)

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

			// Once done writing, rename the output to be the input file name
			outputFile.on('close', () => { 
				console.log('done writing')
			})

			// Read the file and replace any text that matches

			rl.on('line', line => {
				let text = line

				// Elimina las lineas que no comienzan por tabulador:
				if (!text.startsWith('\t')) {
					return;
				}				

				function numberOfTabs(text) {
					var count = 0;
					for( var i= 0; i<text.length; i++){
						if(text.charAt(i) === "\t"){
							count++
						}
					}
					return count;
				}

				if(cuentaRegistroProcesado%10000==0){
					console.log("Progreso: "+(cuentaRegistroProcesado/registrosTotalesProcesar*100));
				}

				cuentaRegistroProcesado++; 
		
				//Elimina las lineas vacias:
				if(numberOfTabs(text)<1){
					return;
				}

				// Elimina las lineas que comienzan por "Md.":
				// Elimina las lineas que comienzan por "N.":
				if (text.startsWith('\tMd.')) {
					return;
				}
				if (text.startsWith('\tN')) {
					return;
				}

				text = text.substr(1);

				text = text.substr(1);


				outputFile.write(`${text}\n`)
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

	async eliminarDuplicadosSpool(argumentos){

		console.log("Eliminando duplicados Spool: ");
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
		const outputFile = fs.createWriteStream(pathSpoolOutput)

		async function leerSpool(){
			return new Promise((resolve) => {
			const rl = readline.createInterface({
				input: fs.createReadStream(pathSpoolInput)
			})

			var lineaAnterior = "";

			outputFile.on('err', err => {
				console.log(err)
			})

			outputFile.on('close', () => { 
				console.log('done writing')
			})
				
			rl.on('line', line => {
				let text = line

				if(lineaAnterior == text){
					return; 
				}

				lineaAnterior = text;

				outputFile.write(`${text}\n`)
			})

			rl.on('close', () => {
				console.log("FIN DEL PROCESAMIENTO");
				outputFile.end()
				resolve(true);
			})	
			})
		}
			var result = await leerSpool();
			return result;
	}

	async dividirArchivoSpool(argumentos){

		console.log("Dividiendo archivo Spool: ");
		console.log("Archivo entrada: "+argumentos[0])
		console.log("Archivo salida: "+argumentos[1])
		
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

		async function dividirArchivo(registrosTotales,numeroArchivos){
			return new Promise((resolve) => {
				
			var archivosNuevos = [];		
			var cuentaLinea= 0;

			for(var i = 0; i<numeroArchivos; i++){
				archivosNuevos.push(fs.createWriteStream(path.join(argumentos[1],argumentos[2]+"_"+i+".txt")));
			}

			const rl = readline.createInterface({
				input: fs.createReadStream(pathSpoolInput)
			})

			rl.on('line', line => {
				let text = line
				archivosNuevos[Math.floor(cuentaLinea/(registrosTotales/numeroArchivos))].write(`${text}\n`)
				cuentaLinea++;
			})

			rl.on('close', () => {
				for(var i = 0; i<numeroArchivos; i++){
					archivosNuevos[i].end();
				}
				console.log("FIN DEL PROCESAMIENTO");
				resolve(true);
			})	
			})
		}
		

		var numeroRegistros = await contarRegistros();
		var registrosDivision = 900000;
		var numeroArchivos = numeroRegistros/registrosDivision;
		numeroArchivos = Math.ceil(numeroArchivos);
		var result = await dividirArchivo(numeroRegistros, numeroArchivos);
		return result;
	}

	async incluirArchivo(){

	}
	
	async spoolToXLSX(argumentos){

		console.log("Formatear SPOOL");
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
		const outputFile = fs.createWriteStream(pathSpoolOutput)

		async function leerSpool(){
			return new Promise((resolve) => {
			const rl = readline.createInterface({
				input: fs.createReadStream(pathSpoolInput)
			})

			// Handle any error that occurs on the write stream
			outputFile.on('err', err => {
				// handle error
				console.log(err)
			})

			// Once done writing, rename the output to be the input file name
			outputFile.on('close', () => { 
				console.log('done writing')

				/*fs.rename(pathSpoolOutput, pathSpoolInput, err => {
					if (err) {
					  // handle error
					  console.log(err)
					} else {
					  console.log('renamed file')
					}
				})*/ 
			})

			// Read the file and replace any text that matches
			rl.on('line', line => {
				let text = line

				// Elimina las lineas que no comienzan por tabulador:
				if (!text.startsWith('\t')) {
					return;
				}
				
				// Elimina las lineas que comienzan por "Md.":
				if (text.startsWith('\tMd.\t')) {
					return;
				}

				// write text to the output file stream with new line character
				outputFile.write(`${text}\n`)
			})

			// Done reading the input, call end() on the write stream
			rl.on('close', () => {
				console.log("FIN DEL PROCESAMIENTO");
				outputFile.end()
				resolve(true);
			})	
			})
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

	async subirCursos(argumentos){

        console.log("SUBIENDO MONITORIZACIÓN CURSOS");

		const pathMonitorizacionCursos = path.join(argumentos[0]);
        //const pathRaiz = pathMonitorizacionCursos.substring(0, pathMonitorizacionCursos.lastIndexOf("\\"));
        const pathRaiz = path.dirname(pathMonitorizacionCursos);

        var cursos = argumentos[1];
        var formadores = argumentos[3];
        var formadorCurso = argumentos[5];
        var codigosProvincia = argumentos[6];
        var instituciones = argumentos[7];

        console.log("PATH RAIZ: "+pathRaiz);
        console.log("RUTA MONITORIZACIÓN CURSO: "+pathMonitorizacionCursos);

        //Importando XLSX:
		return new Promise((resolve) => {
        var monitorizacionCursos = {}
		XlsxPopulate.fromFileAsync(path.normalize(pathMonitorizacionCursos))
			.then(workbook => {
				console.log("Archivo Cargado: Monitorización Cursos");
				monitorizacionCursos = workbook;
			})
			.then(()=>{

                //IDENTIFICAR CAMBIOS:
                var cambios = [];
                for(var i = 0; i < cursos.length; i++){
                   if(cursos[i].metadatos.flag_cambio && !cursos[i].metadatos.error){
                       cambios.push(cursos[i])
                   }
                }

                var cambiosFormadores = [];
                for(var i = 0; i < formadores[0].data.length; i++){
                   if(formadores[0].data[i].metadatos.flag_cambio && !formadores[0].data[i].metadatos.error){
                       cambiosFormadores.push(formadores[0].data[i])
                   }
                }

                var cambiosInstituciones = [];
                for(var i = 0; i < instituciones[0].data.length; i++){
                   if(instituciones[0].data[i].metadatos.flag_cambio && !instituciones[0].data[i].metadatos.error){
                       cambiosInstituciones.push(instituciones[0].data[i])
                   }
                }

                console.log("Cambios Cursos Detectados: "+cambios.length)
                console.log(cambios)
                console.log("Cambios Formadores Detectados: "+ cambiosFormadores.length)
                console.log(cambiosFormadores)
                console.log("Cambios Instituciones Detectados: "+ cambiosInstituciones.length)
                console.log(cambiosInstituciones)

                //Aplicando Cambios Cursos:
                var contadorNuevas = 0;
                var contadorModificacion = 0;
                var columnasCursos = monitorizacionCursos.sheet("Cursos").usedRange()._numColumns;
                var filasCursos = monitorizacionCursos.sheet("Cursos").usedRange()._numRows;
                var filasFormadoresCursos = monitorizacionCursos.sheet("Formador-Curso").usedRange()._numRows;

                //Recalculo de filas usadas:
                while(!monitorizacionCursos.sheet("Cursos").row(filasCursos).cell(1).value()){
                    filasCursos--;
                }
                while(!monitorizacionCursos.sheet("Formador-Curso").row(filasFormadoresCursos).cell(1).value()){
                    filasFormadoresCursos--;
                }
                
                var punteroRegistroFormador = filasFormadoresCursos+1;
                var encontrado = false;

                for(var i = 0; i < cambios.length; i++){
                    encontrado = false;
                    for(var j = 1; j < filasCursos; j++){
                        //Si se encuentra el registro:
                        if(monitorizacionCursos.sheet("Cursos").row(j+1).cell(1).value()== cambios[i]["cod_curso"]){
                            contadorModificacion++;
                            encontrado = true;

                            //Reescribir Registro:
                            monitorizacionCursos.sheet("Cursos").row(j+1).cell(2).value(cambios[i]["cod_grupo"])
                            monitorizacionCursos.sheet("Cursos").row(j+1).cell(3).value(cambios[i]["cod__postal"])
                            monitorizacionCursos.sheet("Cursos").row(j+1).cell(4).value(cambios[i]["territorial"])
                            monitorizacionCursos.sheet("Cursos").row(j+1).cell(5).value(cambios[i]["ccaa_/_pais"])
                            monitorizacionCursos.sheet("Cursos").row(j+1).cell(6).value(cambios[i]["curso"])
                            monitorizacionCursos.sheet("Cursos").row(j+1).cell(7).value(cambios[i]["sesión"])
                            monitorizacionCursos.sheet("Cursos").row(j+1).cell(8).value(cambios[i]["fecha"])
                            monitorizacionCursos.sheet("Cursos").row(j+1).cell(9).value(cambios[i]["hora_inicio"])
                            monitorizacionCursos.sheet("Cursos").row(j+1).cell(10).value(cambios[i]["hora_fin"])
                            monitorizacionCursos.sheet("Cursos").row(j+1).cell(11).value(cambios[i]["duración"])
                            monitorizacionCursos.sheet("Cursos").row(j+1).cell(12).value(cambios[i]["institución"])
                            monitorizacionCursos.sheet("Cursos").row(j+1).cell(13).value(cambios[i]["colectivo"])
                            monitorizacionCursos.sheet("Cursos").row(j+1).cell(14).value(cambios[i]["grupo"])
                            monitorizacionCursos.sheet("Cursos").row(j+1).cell(15).value(cambios[i]["nºasistentes"])
                            monitorizacionCursos.sheet("Cursos").row(j+1).cell(16).value(cambios[i]["modalidad"])
                            monitorizacionCursos.sheet("Cursos").row(j+1).cell(17).value(cambios[i]["estado"])
                            monitorizacionCursos.sheet("Cursos").row(j+1).cell(18).value(cambios[i]["material"])
                            if(typeof cambios[i]["valoración"] != "undefined"){ 
                            monitorizacionCursos.sheet("Cursos").row(j+1).cell(19).value(cambios[i]["valoración"])
                            }else{
                            monitorizacionCursos.sheet("Cursos").row(j+1).cell(19).value("SIN VALORAR")
                            }
                            monitorizacionCursos.sheet("Cursos").row(j+1).cell(20).value(cambios[i]["observaciones"])
                            break;
                        }
                    }

                    if(!encontrado){

                        //Crear Nuevo Curso:
                            monitorizacionCursos.sheet("Cursos").row(filasCursos+contadorNuevas+1).cell(1).value(cambios[i]["cod_curso"])
                            monitorizacionCursos.sheet("Cursos").row(filasCursos+contadorNuevas+1).cell(2).value(cambios[i]["cod_grupo"])
                            monitorizacionCursos.sheet("Cursos").row(filasCursos+contadorNuevas+1).cell(3).value(cambios[i]["cod__postal"])
                            monitorizacionCursos.sheet("Cursos").row(filasCursos+contadorNuevas+1).cell(4).value(cambios[i]["territorial"])
                            monitorizacionCursos.sheet("Cursos").row(filasCursos+contadorNuevas+1).cell(5).value(cambios[i]["ccaa_/_pais"])
                            monitorizacionCursos.sheet("Cursos").row(filasCursos+contadorNuevas+1).cell(6).value(cambios[i]["curso"])
                            monitorizacionCursos.sheet("Cursos").row(filasCursos+contadorNuevas+1).cell(7).value(cambios[i]["sesión"])
                            monitorizacionCursos.sheet("Cursos").row(filasCursos+contadorNuevas+1).cell(8).value(cambios[i]["fecha"])
                            monitorizacionCursos.sheet("Cursos").row(filasCursos+contadorNuevas+1).cell(9).value(cambios[i]["hora_inicio"])
                            monitorizacionCursos.sheet("Cursos").row(filasCursos+contadorNuevas+1).cell(10).value(cambios[i]["hora_fin"])
                            monitorizacionCursos.sheet("Cursos").row(filasCursos+contadorNuevas+1).cell(11).value(cambios[i]["duración"])
                            monitorizacionCursos.sheet("Cursos").row(filasCursos+contadorNuevas+1).cell(12).value(cambios[i]["institución"])
                            monitorizacionCursos.sheet("Cursos").row(filasCursos+contadorNuevas+1).cell(13).value(cambios[i]["colectivo"])
                            monitorizacionCursos.sheet("Cursos").row(filasCursos+contadorNuevas+1).cell(14).value(cambios[i]["grupo"])
                            monitorizacionCursos.sheet("Cursos").row(filasCursos+contadorNuevas+1).cell(15).value(cambios[i]["nºasistentes"])
                            monitorizacionCursos.sheet("Cursos").row(filasCursos+contadorNuevas+1).cell(16).value(cambios[i]["modalidad"])
                            monitorizacionCursos.sheet("Cursos").row(filasCursos+contadorNuevas+1).cell(17).value(cambios[i]["estado"])
                            monitorizacionCursos.sheet("Cursos").row(filasCursos+contadorNuevas+1).cell(18).value(cambios[i]["material"])
                            if(typeof cambios[i]["valoración"] != "undefined"){ 
                            monitorizacionCursos.sheet("Cursos").row(filasCursos+contadorNuevas+1).cell(19).value(cambios[i]["valoración"])
                            }else{
                            monitorizacionCursos.sheet("Cursos").row(filasCursos+contadorNuevas+1).cell(19).value("SIN VALORAR")
                            }
                            monitorizacionCursos.sheet("Cursos").row(filasCursos+contadorNuevas+1).cell(20).value(cambios[i]["observaciones"])

                        //Actualizar Contador:
                        contadorNuevas++;
                    }

                    //Modificar Curso-Formador:
                    
                        // 1) Eliminar Referencias al curso en Curso-Formadores:
                        filasFormadoresCursos = monitorizacionCursos.sheet("Formador-Curso").usedRange()._numRows;
                        for(var k = 1; k < filasFormadoresCursos+1; k++ ){
                           if(monitorizacionCursos.sheet("Formador-Curso").row(k).cell(1).value()==cambios[i]["cod_curso"]){
                                monitorizacionCursos.sheet("Formador-Curso").row(k).cell(1).value("")
                                monitorizacionCursos.sheet("Formador-Curso").row(k).cell(2).value("")
                           }
                        }

                        // 2) Añadiendo Formadores:
                        if(typeof cambios[i].metadatos["formadores"] == "object"){
                            for(var k = 0; k < cambios[i].metadatos["formadores"].length; k++){
                                monitorizacionCursos.sheet("Formador-Curso").row(punteroRegistroFormador).cell(1).value(cambios[i]["cod_curso"])
                                monitorizacionCursos.sheet("Formador-Curso").row(punteroRegistroFormador).cell(2).value(cambios[i]["metadatos"]["formadores"][k]["id"])
                                punteroRegistroFormador++;
                            }
                        }

                } //Fin de iteracion de cambios CURSOS


                //Aplicando Cambios Formadores:
                var contadorNuevosFormadores = 0;
                var contadorModificacionFormadores = 0;
                var columnasFormadores = monitorizacionCursos.sheet("Formadores").usedRange()._numColumns;
                var filasFormadores = monitorizacionCursos.sheet("Formadores").usedRange()._numRows;
                var formadorEncontrado = false;

                //Recalculo de filas usadas:
                while(!monitorizacionCursos.sheet("Formadores").row(filasFormadores).cell(1).value()){
                    filasFormadores--;
                }

                for(var i = 0; i < cambiosFormadores.length; i++){
                    formadorEncontrado = false;
                    for(var j = 1; j < filasFormadores; j++){
                        if(monitorizacionCursos.sheet("Formadores").row(j+1).cell(1).value()== cambiosFormadores[i]["cod__formador"]){
                            contadorModificacionFormadores++;
                            formadorEncontrado = true;

                            //Reescribir Registro:
                            monitorizacionCursos.sheet("Formadores").row(j+1).cell(2).value(cambiosFormadores[i]["nombre"])
                            monitorizacionCursos.sheet("Formadores").row(j+1).cell(3).value(cambiosFormadores[i]["email"])
                            monitorizacionCursos.sheet("Formadores").row(j+1).cell(4).value(cambiosFormadores[i]["telefono"])
                            monitorizacionCursos.sheet("Formadores").row(j+1).cell(5).value(cambiosFormadores[i]["territorial"])
                            monitorizacionCursos.sheet("Formadores").row(j+1).cell(6).value(cambiosFormadores[i]["ccaa"])
                            monitorizacionCursos.sheet("Formadores").row(j+1).cell(7).value(cambiosFormadores[i]["provincia"])
                            monitorizacionCursos.sheet("Formadores").row(j+1).cell(8).value(cambiosFormadores[i]["fecha"])
                            monitorizacionCursos.sheet("Formadores").row(j+1).cell(9).value(cambiosFormadores[i]["certificado"])
                            monitorizacionCursos.sheet("Formadores").row(j+1).cell(10).value(cambiosFormadores[i]["confidencialidad"])
                            monitorizacionCursos.sheet("Formadores").row(j+1).cell(11).value(cambiosFormadores[i]["consentimiento"])
                            monitorizacionCursos.sheet("Formadores").row(j+1).cell(12).value(cambiosFormadores[i]["estado"])
                            contadorModificacionFormadores++;
                            break;
                        }
                    }

                    //NO ENCONTRADO
                    if(!formadorEncontrado){
                            //Nuevo Formador:
                            monitorizacionCursos.sheet("Formadores").row(filasFormadores+contadorNuevosFormadores+1).cell(1).value(cambiosFormadores[i]["cod__formador"])
                            monitorizacionCursos.sheet("Formadores").row(filasFormadores+contadorNuevosFormadores+1).cell(2).value(cambiosFormadores[i]["nombre"])
                            monitorizacionCursos.sheet("Formadores").row(filasFormadores+contadorNuevosFormadores+1).cell(3).value(cambiosFormadores[i]["email"])
                            monitorizacionCursos.sheet("Formadores").row(filasFormadores+contadorNuevosFormadores+1).cell(4).value(cambiosFormadores[i]["telefono"])
                            monitorizacionCursos.sheet("Formadores").row(filasFormadores+contadorNuevosFormadores+1).cell(5).value(cambiosFormadores[i]["territorial"])
                            monitorizacionCursos.sheet("Formadores").row(filasFormadores+contadorNuevosFormadores+1).cell(6).value(cambiosFormadores[i]["ccaa"])
                            monitorizacionCursos.sheet("Formadores").row(filasFormadores+contadorNuevosFormadores+1).cell(7).value(cambiosFormadores[i]["provincia"])
                            monitorizacionCursos.sheet("Formadores").row(filasFormadores+contadorNuevosFormadores+1).cell(8).value(cambiosFormadores[i]["fecha"])
                            monitorizacionCursos.sheet("Formadores").row(filasFormadores+contadorNuevosFormadores+1).cell(9).value(cambiosFormadores[i]["certificado"])
                            monitorizacionCursos.sheet("Formadores").row(filasFormadores+contadorNuevosFormadores+1).cell(10).value(cambiosFormadores[i]["confidencialidad"])
                            monitorizacionCursos.sheet("Formadores").row(filasFormadores+contadorNuevosFormadores+1).cell(11).value(cambiosFormadores[i]["consentimiento"])
                            monitorizacionCursos.sheet("Formadores").row(filasFormadores+contadorNuevosFormadores+1).cell(12).value(cambiosFormadores[i]["estado"])

                        //Actualizar Contador:
                        contadorNuevosFormadores++;
                    }

                } //Fin de iteracion de cambios Formadores

                //Aplicando Cambios Institución:
                var contadorNuevasInstituciones = 0;
                var contadorModificacionInstituciones = 0;
                var columnasInstituciones = monitorizacionCursos.sheet("Instituciones").usedRange()._numColumns;
                var filasInstituciones = monitorizacionCursos.sheet("Instituciones").usedRange()._numRows;
                var institucionEncontrada = false;

                //Recalculo de filas usadas:
                while(!monitorizacionCursos.sheet("Instituciones").row(filasInstituciones).cell(1).value()){
                    filasInstituciones--;
                }

                for(var i = 0; i < cambiosInstituciones.length; i++){
                    institucionEncontrada = false;
                    for(var j = 1; j < filasInstituciones; j++){
                        if(monitorizacionCursos.sheet("Instituciones").row(j+1).cell(1).value()== cambiosInstituciones[i]["cod_institucion"]){
                            contadorModificacionInstituciones++;
                            institucionEncontrada = true;

                            //Reescribir Registro:
                            monitorizacionCursos.sheet("Instituciones").row(j+1).cell(2).value(cambiosInstituciones[i]["institucion"])
                            monitorizacionCursos.sheet("Instituciones").row(j+1).cell(3).value(cambiosInstituciones[i]["tipo"])
                            monitorizacionCursos.sheet("Instituciones").row(j+1).cell(4).value(cambiosInstituciones[i]["cod__postal"])
                            monitorizacionCursos.sheet("Instituciones").row(j+1).cell(5).value(cambiosInstituciones[i]["territorial"])
                            monitorizacionCursos.sheet("Instituciones").row(j+1).cell(6).value(cambiosInstituciones[i]["ccaa_/_pais"])
                            monitorizacionCursos.sheet("Instituciones").row(j+1).cell(7).value(cambiosInstituciones[i]["provincia"])
                            monitorizacionCursos.sheet("Instituciones").row(j+1).cell(8).value(cambiosInstituciones[i]["contacto1"])
                            monitorizacionCursos.sheet("Instituciones").row(j+1).cell(9).value(cambiosInstituciones[i]["email1"])
                            monitorizacionCursos.sheet("Instituciones").row(j+1).cell(10).value(cambiosInstituciones[i]["telefono1"])
                            monitorizacionCursos.sheet("Instituciones").row(j+1).cell(11).value(cambiosInstituciones[i]["contacto2"])
                            monitorizacionCursos.sheet("Instituciones").row(j+1).cell(12).value(cambiosInstituciones[i]["email2"])
                            monitorizacionCursos.sheet("Instituciones").row(j+1).cell(13).value(cambiosInstituciones[i]["telefono2"])
                            monitorizacionCursos.sheet("Instituciones").row(j+1).cell(14).value(cambiosInstituciones[i]["contacto3"])
                            monitorizacionCursos.sheet("Instituciones").row(j+1).cell(15).value(cambiosInstituciones[i]["email3"])
                            monitorizacionCursos.sheet("Instituciones").row(j+1).cell(16).value(cambiosInstituciones[i]["telefono3"])
                            monitorizacionCursos.sheet("Instituciones").row(j+1).cell(17).value(cambiosInstituciones[i]["contacto4"])
                            monitorizacionCursos.sheet("Instituciones").row(j+1).cell(18).value(cambiosInstituciones[i]["email4"])
                            monitorizacionCursos.sheet("Instituciones").row(j+1).cell(19).value(cambiosInstituciones[i]["telefono4"])
                            monitorizacionCursos.sheet("Instituciones").row(j+1).cell(20).value(cambiosInstituciones[i]["direccion"])
                            contadorModificacionInstituciones++;
                            break;
                        }
                    }

                    //NO ENCONTRADO
                    if(!institucionEncontrada){
                            //Nueva Institucion:
                            monitorizacionCursos.sheet("Instituciones").row(filasInstituciones+contadorNuevasInstituciones+1).cell(1).value(cambiosInstituciones[i]["cod_institucion"])
                            monitorizacionCursos.sheet("Instituciones").row(filasInstituciones+contadorNuevasInstituciones+1).cell(2).value(cambiosInstituciones[i]["institucion"])
                            monitorizacionCursos.sheet("Instituciones").row(filasInstituciones+contadorNuevasInstituciones+1).cell(3).value(cambiosInstituciones[i]["tipo"])
                            monitorizacionCursos.sheet("Instituciones").row(filasInstituciones+contadorNuevasInstituciones+1).cell(4).value(cambiosInstituciones[i]["cod__postal"])
                            monitorizacionCursos.sheet("Instituciones").row(filasInstituciones+contadorNuevasInstituciones+1).cell(5).value(cambiosInstituciones[i]["territorial"])
                            monitorizacionCursos.sheet("Instituciones").row(filasInstituciones+contadorNuevasInstituciones+1).cell(6).value(cambiosInstituciones[i]["ccaa_/_pais"])
                            monitorizacionCursos.sheet("Instituciones").row(filasInstituciones+contadorNuevasInstituciones+1).cell(7).value(cambiosInstituciones[i]["provincia"])

                            monitorizacionCursos.sheet("Instituciones").row(filasInstituciones+contadorNuevasInstituciones+1).cell(8).value(cambiosInstituciones[i]["contacto1"])
                            monitorizacionCursos.sheet("Instituciones").row(filasInstituciones+contadorNuevasInstituciones+1).cell(9).value(cambiosInstituciones[i]["email1"])
                            monitorizacionCursos.sheet("Instituciones").row(filasInstituciones+contadorNuevasInstituciones+1).cell(10).value(cambiosInstituciones[i]["telefono1"])

                            monitorizacionCursos.sheet("Instituciones").row(filasInstituciones+contadorNuevasInstituciones+1).cell(11).value(cambiosInstituciones[i]["contacto2"])
                            monitorizacionCursos.sheet("Instituciones").row(filasInstituciones+contadorNuevasInstituciones+1).cell(12).value(cambiosInstituciones[i]["email2"])
                            monitorizacionCursos.sheet("Instituciones").row(filasInstituciones+contadorNuevasInstituciones+1).cell(13).value(cambiosInstituciones[i]["telefono2"])

                            monitorizacionCursos.sheet("Instituciones").row(filasInstituciones+contadorNuevasInstituciones+1).cell(14).value(cambiosInstituciones[i]["contacto3"])
                            monitorizacionCursos.sheet("Instituciones").row(filasInstituciones+contadorNuevasInstituciones+1).cell(15).value(cambiosInstituciones[i]["email3"])
                            monitorizacionCursos.sheet("Instituciones").row(filasInstituciones+contadorNuevasInstituciones+1).cell(16).value(cambiosInstituciones[i]["telefono3"])

                            monitorizacionCursos.sheet("Instituciones").row(filasInstituciones+contadorNuevasInstituciones+1).cell(17).value(cambiosInstituciones[i]["contacto4"])
                            monitorizacionCursos.sheet("Instituciones").row(filasInstituciones+contadorNuevasInstituciones+1).cell(18).value(cambiosInstituciones[i]["email4"])
                            monitorizacionCursos.sheet("Instituciones").row(filasInstituciones+contadorNuevasInstituciones+1).cell(19).value(cambiosInstituciones[i]["telefono4"])
                            monitorizacionCursos.sheet("Instituciones").row(filasInstituciones+contadorNuevasInstituciones+1).cell(20).value(cambiosInstituciones[i]["direccion"])
                        
                        //Actualizar Contador:
                        contadorNuevasInstituciones++;
                    }
                } //Fin de iteracion de cambios Formadores

                console.log("Num Cambios Cursos:"+ cambios.length)
                console.log("Modificaciones Cursos:"+ contadorModificacion)
                console.log("Nuevos Cursos:"+ contadorNuevas)

                console.log("Num Cambios Formadores:"+cambiosFormadores.length)
                console.log("Modificaciones Formadores:"+contadorModificacionFormadores)
                console.log("Nuevos Formadores:"+contadorNuevosFormadores)

                console.log("Num Cambios Instituciones:"+cambiosInstituciones.length)
                console.log("Modificaciones Instituciones:"+contadorModificacionInstituciones)
                console.log("Nuevas Instituciones:"+contadorNuevasInstituciones)

                //Nuevas filas:
                filasCursos = filasCursos+contadorNuevas+1;
                filasInstituciones = filasInstituciones+contadorNuevasInstituciones+1;
                filasFormadores = filasFormadores+contadorNuevosFormadores+1;
                filasFormadoresCursos = punteroRegistroFormador+1;

                //CREAR OBJETOS JSON:
                var jsonCursos = [];
                for(var i = 1; i < filasCursos ; i++){
                    jsonCursos.push({
                        "cod_curso": monitorizacionCursos.sheet("Cursos").row(i+1).cell(1).value(),
                        "cod_grupo": monitorizacionCursos.sheet("Cursos").row(i+1).cell(2).value(),
                        "cod__postal": monitorizacionCursos.sheet("Cursos").row(i+1).cell(3).value(),
                        "territorial": monitorizacionCursos.sheet("Cursos").row(i+1).cell(4).value(),
                        "ccaa_/_pais": monitorizacionCursos.sheet("Cursos").row(i+1).cell(5).value(),
                        "curso": monitorizacionCursos.sheet("Cursos").row(i+1).cell(6).value(),
                        "sesi\u00f3n": monitorizacionCursos.sheet("Cursos").row(i+1).cell(7).value(),
                        "fecha": monitorizacionCursos.sheet("Cursos").row(i+1).cell(8).value(),
                        "hora_inicio": monitorizacionCursos.sheet("Cursos").row(i+1).cell(9).value(),
                        "hora_fin": monitorizacionCursos.sheet("Cursos").row(i+1).cell(10).value(),
                        "duraci\u00f3n": monitorizacionCursos.sheet("Cursos").row(i+1).cell(11).value(),
                        "instituci\u00f3n": monitorizacionCursos.sheet("Cursos").row(i+1).cell(12).value(),
                        "colectivo": monitorizacionCursos.sheet("Cursos").row(i+1).cell(13).value(),
                        "grupo": monitorizacionCursos.sheet("Cursos").row(i+1).cell(14).value(),
                        "n\u00baasistentes": monitorizacionCursos.sheet("Cursos").row(i+1).cell(15).value(),
                        "modalidad": monitorizacionCursos.sheet("Cursos").row(i+1).cell(16).value(),
                        "estado": monitorizacionCursos.sheet("Cursos").row(i+1).cell(17).value(),
                        "material": monitorizacionCursos.sheet("Cursos").row(i+1).cell(18).value(),
                        "valoraci\u00f3n": monitorizacionCursos.sheet("Cursos").row(i+1).cell(19).value(),
                        "observaciones": monitorizacionCursos.sheet("Cursos").row(i+1).cell(20).value()
                    })
                }

                var jsonFormadores = [];
                for(var i = 1; i < filasFormadores ; i++){
                    jsonFormadores.push({
                        "cod__formador": monitorizacionCursos.sheet("Formadores").row(i+1).cell(1).value(),
                        "nombre": monitorizacionCursos.sheet("Formadores").row(i+1).cell(2).value(),
                        "email": monitorizacionCursos.sheet("Formadores").row(i+1).cell(3).value(),
                        "telefono": monitorizacionCursos.sheet("Formadores").row(i+1).cell(4).value(),
                        "territorial": monitorizacionCursos.sheet("Formadores").row(i+1).cell(5).value(),
                        "ccaa": monitorizacionCursos.sheet("Formadores").row(i+1).cell(6).value(),
                        "provincia": monitorizacionCursos.sheet("Formadores").row(i+1).cell(7).value(),
                        "fecha": monitorizacionCursos.sheet("Formadores").row(i+1).cell(8).value(),
                        "certificado": monitorizacionCursos.sheet("Formadores").row(i+1).cell(9).value(),
                        "confidencialidad": monitorizacionCursos.sheet("Formadores").row(i+1).cell(10).value(),
                        "consentimiento": monitorizacionCursos.sheet("Formadores").row(i+1).cell(11).value(),
                        "estado": monitorizacionCursos.sheet("Formadores").row(i+1).cell(12).value()
                    })
                }

                var jsonInstituciones = [];
                for(var i = 1; i < filasInstituciones ; i++){
                    jsonInstituciones.push({
                        "cod_institucion": monitorizacionCursos.sheet("Instituciones").row(i+1).cell(1).value(),
                        "institucion": monitorizacionCursos.sheet("Instituciones").row(i+1).cell(2).value(),
                        "tipo": monitorizacionCursos.sheet("Instituciones").row(i+1).cell(3).value(),
                        "cod__postal": monitorizacionCursos.sheet("Instituciones").row(i+1).cell(4).value(),
                        "territorial": monitorizacionCursos.sheet("Instituciones").row(i+1).cell(5).value(),
                        "ccaa": monitorizacionCursos.sheet("Instituciones").row(i+1).cell(6).value(),
                        "provincia": monitorizacionCursos.sheet("Instituciones").row(i+1).cell(7).value(),

                        "contacto1": monitorizacionCursos.sheet("Instituciones").row(i+1).cell(8).value(),
                        "email1": monitorizacionCursos.sheet("Instituciones").row(i+1).cell(9).value(),
                        "telefono1": monitorizacionCursos.sheet("Instituciones").row(i+1).cell(10).value(),

                        "contacto2": monitorizacionCursos.sheet("Instituciones").row(i+1).cell(11).value(),
                        "email2": monitorizacionCursos.sheet("Instituciones").row(i+1).cell(12).value(),
                        "telefono2": monitorizacionCursos.sheet("Instituciones").row(i+1).cell(13).value(),

                        "contacto3": monitorizacionCursos.sheet("Instituciones").row(i+1).cell(14).value(),
                        "email3": monitorizacionCursos.sheet("Instituciones").row(i+1).cell(15).value(),
                        "telefono3": monitorizacionCursos.sheet("Instituciones").row(i+1).cell(16).value(),

                        "contacto4": monitorizacionCursos.sheet("Instituciones").row(i+1).cell(17).value(),
                        "email4": monitorizacionCursos.sheet("Instituciones").row(i+1).cell(18).value(),
                        "telefono4": monitorizacionCursos.sheet("Instituciones").row(i+1).cell(19).value(),
                        "direccion": monitorizacionCursos.sheet("Instituciones").row(i+1).cell(20).value()
                    })
                }

                var jsonFormadorCurso = [];
                for(var i = 1; i < filasFormadoresCursos ; i++){
                    jsonFormadorCurso.push({
                        "cod__curso": monitorizacionCursos.sheet("Formador-Curso").row(i+1).cell(1).value(),
                        "cod__formador": monitorizacionCursos.sheet("Formador-Curso").row(i+1).cell(2).value(),
                    })
                }

            //Eliminar Filas Vacias:
            for(var i= 0; i < jsonCursos.length; i++){
                if(!jsonCursos[i].cod_curso){
                    jsonCursos.splice(i,1);
                    i--;
                }
            }

            //Eliminar Filas Vacias Formadores:
            for(var i= 0; i < jsonFormadores.length; i++){
                if(!jsonFormadores[i].cod__formador){
                    jsonFormadores.splice(i,1);
                    i--;
                }
            }

            //Eliminar Filas Vacias Instituciones:
            for(var i= 0; i < jsonInstituciones.length; i++){
                if(!jsonInstituciones[i].cod_institucion){
                    jsonInstituciones.splice(i,1);
                    i--;
                }
            }

            //Eliminar Filas Vacias Curso-Formador:
            for(var i= 0; i < jsonFormadorCurso.length; i++){
                if(!jsonFormadorCurso[i].cod__curso){
                    jsonFormadorCurso.splice(i,1);
                    i--;
                }
            }

            //Guardar Archivos JSON:
            jsonCursos = JSON.stringify(jsonCursos);
            jsonFormadores = JSON.stringify(jsonFormadores);
            jsonInstituciones = JSON.stringify(jsonInstituciones);
            jsonFormadorCurso = JSON.stringify(jsonFormadorCurso);
            let jsonProvincia = JSON.stringify(codigosProvincia);

            try{
                fs.writeFileSync(path.normalize(path.join(pathRaiz,'db/cursos.json')), jsonCursos);
                fs.writeFileSync(path.normalize(path.join(pathRaiz,'db/formadores.json')), jsonFormadores);
                fs.writeFileSync(path.normalize(path.join(pathRaiz,'db/instituciones.json')), jsonInstituciones);
                fs.writeFileSync(path.normalize(path.join(pathRaiz,'db/formador-curso.json')), jsonFormadorCurso);
                fs.writeFileSync(path.normalize(path.join(pathRaiz,'db/provincia.json')), jsonProvincia);
            }catch(err){
                console.log("Se ha producido un error interno: ");
                console.log(err);
                var tituloError = "Se ha producido un error guardando los archivos JSON. "
                resolve(false)
            };

            //Fin de procesamiento:
            console.log("Escribiendo archivo...");
            console.log( "Path: " + path.normalize(pathMonitorizacionCursos));

            monitorizacionCursos.toFileAsync(
                    path.normalize(pathMonitorizacionCursos)
                )
                .then(() => {
                    console.log("Fin del procesamiento");
                    //console.log(monitorizacionCursos)

                    resolve(true)
                })
                .catch(err => {
                    console.log("Se ha producido un error interno: ");
                    console.log(err);
                    var tituloError =
                        "Se ha producido un error escribiendo el archivo: " +
                        path.normalize(pathMonitorizacionCursos);
                    resolve(false)
                });

            })
        })
        
    }

}//Fin Procesos Generales

module.exports = ProcesosGenerales;


