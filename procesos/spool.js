const path = require("path");
const fs = require("fs");
const readline = require('readline')
const moment = require("moment");
const XlsxPopulate = require("xlsx-populate");
const Datastore = require("nedb");
const _= require("lodash");
const {ipcRenderer}= require("electron");
const puppeteer = require('puppeteer');

class ProcesosSpool {
	
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
				var indexCuentaTab = 41;
				
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

				if(day.isBefore(moment('01.09.2020',"DD.MM.YYYY"))){
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

	async spoolToXlsx(argumentos){

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
}

module.exports = ProcesosSpool;


