
const path = require("path");
const fs = require("fs");
const readline = require('readline')
const XlsxPopulate = require("xlsx-populate");
const ipcRenderer= require("electron").ipcRenderer;
const ipc = require("electron").ipcMain;
const mainProcess = require("../main.js");
const moment = require("moment");

class ProcesosKPIs {
	
	constructor(pathToDbFolder, nombreProyecto, proyectoDB){

		this.pathToDbFolder = pathToDbFolder;
		this.nombreProyecto = nombreProyecto; 
		this.proyectoDB = proyectoDB;

	}

	async facturacionCicloStep1(argumentos){

		console.log("Importando Excel:");
		console.log("Argumentos: ");
		//console.log(argumentos);

		var pathExcelCiclo = argumentos[0];
		var pathExcelSeguimiento = argumentos[1];
		var numFilaCabecera = 1;

		//PROCESAMIENTO EXCEL Ciclo: 
		XlsxPopulate.fromFileAsync(path.normalize(pathExcelCiclo))
			.then(excelCiclo => {
				console.log("Cargando Excel:");
				//console.log(excelCiclo);
				//
				if(excelCiclo===undefined){
					return false;
				}

				//Creación del Objeto:
				var objetoDatosRaw = [{
					data: [],
					nombreId: "Facturacion_RAW",
					objetoId: "Facturacion_RAW",
				}]

				var cabeceraSeleccionada = "";

				var numeroCaberas = excelCiclo
					.sheet("DATOS_RAW")
					.usedRange()._numColumns;

				var numeroRegistros = excelCiclo
					.sheet("DATOS_RAW")
					.usedRange()._numRows;

				//Comprobación de inputs:
				if(excelCiclo.sheet("DATOS_RAW")===undefined){
					return false;
				}


				//Relleno de objeto Data:
				for (var i = numFilaCabecera; i < numeroRegistros; i++) {
					objetoDatosRaw[0].data.push({});
					for (var j = 0; j < numeroCaberas; j++) {

						cabeceraSeleccionada = String(
							excelCiclo
								.sheet("DATOS_RAW")
								.row(numFilaCabecera)
								.cell(j + 1)
								.value()
						);

						cabeceraSeleccionada = cabeceraSeleccionada.toLowerCase();
						cabeceraSeleccionada = cabeceraSeleccionada.replace( / /g, "_");
						cabeceraSeleccionada = cabeceraSeleccionada.replace( /\./g, "_");

						if (cabeceraSeleccionada === undefined || cabeceraSeleccionada === "") {
							console.log(
								"Error de cabecera: i=" + i + " j=" + j
							);
							continue;
						}else{
							//Guardado del registro:
							if(excelCiclo.sheet("DATOS_RAW").row(i + 1).cell(j+1).value()){
								objetoDatosRaw[0].data[objetoDatosRaw[0].data.length-1][cabeceraSeleccionada]= excelCiclo.sheet("DATOS_RAW").row(i + 1).cell(j+1).value();
							}

						}
					}

					if(Object.keys(objetoDatosRaw[0].data[objetoDatosRaw[0].data.length-1]).length === 0){
						objetoDatosRaw[0].data.pop();
					}
				}

				//Paso 2: Copiar datos en NO SITE:
				
				//Borrar pestaña no site:
				var numeroCaberasNoSite = excelCiclo
					.sheet("NO SITE")
					.usedRange()._numColumns;

				var numeroFilaNoSite = excelCiclo
					.sheet("NO SITE")
					.usedRange()._numRows;

				for(var i= 2;i<=numeroFilaNoSite;i++){
					for(var j= 1;j<=numeroCaberasNoSite;j++){
						excelCiclo.sheet("NO SITE").cell(i,j).value("");
					}
				}

				//Copiar datos no site:
				var cuentaNoSite=2;
				for(var i= 0;i<objetoDatosRaw[0].data.length;i++){
					if(objetoDatosRaw[0].data[i]["subcategoría"]=="No SITE"){
						excelCiclo.sheet("NO SITE").cell(cuentaNoSite,1).value(objetoDatosRaw[0].data[i]["factura"])
						excelCiclo.sheet("NO SITE").cell(cuentaNoSite,2).value(objetoDatosRaw[0].data[i]["id_smart"])
						excelCiclo.sheet("NO SITE").cell(cuentaNoSite,3).value(objetoDatosRaw[0].data[i]["error"])
						excelCiclo.sheet("NO SITE").cell(cuentaNoSite,4).value(objetoDatosRaw[0].data[i]["descripción"])
						excelCiclo.sheet("NO SITE").cell(cuentaNoSite,5).value(objetoDatosRaw[0].data[i]["importe"])
						excelCiclo.sheet("NO SITE").cell(cuentaNoSite,6).value(objetoDatosRaw[0].data[i]["categoría"])
						excelCiclo.sheet("NO SITE").cell(cuentaNoSite,7).value(objetoDatosRaw[0].data[i]["subcategoría"])
						cuentaNoSite++;
					}
				}

				mainProcess.guardarDocumento(objetoDatosRaw[0])
				console.log(objetoDatosRaw[0].data[objetoDatosRaw[0].data.length-1])

				return excelCiclo
					.toFileAsync(path.normalize(pathExcelCiclo))
					.then(() => {
						console.log("Fin del procesamiento");
						return true
					})
					.catch(err => {
						console.log("Se ha producido un error interno: ");
						console.log(err);
						var tituloError =
							"Se ha producido un error escribiendo el archivo: " +
							path.normalize(path.normalize(pathExcelCiclo));
						return false;
					});
			})
			.then(()=>{
				console.log("Proceso finalizado");
				})

		return true;
	}

	async facturacionCicloStep2(argumentos){

		console.log("Importando Excel:");
		console.log("Argumentos: ");
		console.log(argumentos);

		var pathExcelCiclo = argumentos[0];
		var pathExcelSeguimiento = argumentos[1];
		var numFilaCabecera = 1;

		var facturacionNoCargada=0;
		var clasificacionFacturacion = [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]

		//PROCESAMIENTO EXCEL Ciclo: 
		XlsxPopulate.fromFileAsync(path.normalize(pathExcelCiclo))
			.then(excelCiclo => {
				console.log("Cargando Excel:");
				//console.log(excelCiclo);
				//
				if(excelCiclo===undefined){
					return false;
				}

				//Creación del Objeto:
				var objetoDatosRaw = [{
					data: [],
					nombreId: "Facturacion_RAW_NO_SITE",
					objetoId: "Facturacion_RAW_NO_SITE",
				}]

				var cabeceraSeleccionada = "";

				var numeroCabecerasNoSite = excelCiclo
					.sheet("NO SITE")
					.usedRange()._numColumns;

				var numeroRegistrosNoSite = excelCiclo
					.sheet("NO SITE")
					.usedRange()._numRows;

				//Comprobación de inputs:
				if(excelCiclo.sheet("NO SITE")===undefined){
					return false;
				}


				//Relleno de objeto No Site:
				for (var i = numFilaCabecera; i < numeroRegistrosNoSite; i++) {
					objetoDatosRaw[0].data.push({});
					for (var j = 0; j < numeroCabecerasNoSite; j++) {

						cabeceraSeleccionada = String(
							excelCiclo
								.sheet("NO SITE")
								.row(numFilaCabecera)
								.cell(j + 1)
								.value()
						);

						cabeceraSeleccionada = cabeceraSeleccionada.toLowerCase();
						cabeceraSeleccionada = cabeceraSeleccionada.replace( / /g, "_");
						cabeceraSeleccionada = cabeceraSeleccionada.replace( /\./g, "_");

						if (cabeceraSeleccionada === undefined || cabeceraSeleccionada === "") {
							console.log(
								"Error de cabecera: i=" + i + " j=" + j
							);
							continue;
						}else{
							//Guardado del registro:
							if(excelCiclo.sheet("NO SITE").row(i + 1).cell(j+1).value()){
								objetoDatosRaw[0].data[objetoDatosRaw[0].data.length-1][cabeceraSeleccionada]= excelCiclo.sheet("DATOS_RAW").row(i + 1).cell(j+1).value();
							}

						}
					}

					if(Object.keys(objetoDatosRaw[0].data[objetoDatosRaw[0].data.length-1]).length === 0){
						objetoDatosRaw[0].data.pop();
					}
				}

				var numeroRegistrosFacturacionNoCargada = excelCiclo
					.sheet("DATOS_RAW")
					.usedRange()._numRows;

				//CONTAR FACTURACIÓN NO CARGADA:
				for (var i = 2; i < numeroRegistrosFacturacionNoCargada; i++) {
					if(excelCiclo.sheet("DATOS_RAW").cell(i,1).value()!="" && excelCiclo.sheet("DATOS_RAW").cell(i,1).value()!=undefined){
						facturacionNoCargada++;
					}
				}


				//Rellenar datos ICG_Cliente:
				cabeceraSeleccionada="";
				var numFilaCabeceraICG = 1;

				//Creación del Objeto:
				var objetoDatosICG = [{
					data: [],
					nombreId: "Facturacion_ICG",
					objetoId: "Facturacion_ICG",
				}]

				var numeroCabecerasICG = excelCiclo
					.sheet("DATOS_ICG_CLIENTE")
					.usedRange()._numColumns;

				var numeroRegistrosICG = excelCiclo
					.sheet("DATOS_ICG_CLIENTE")
					.usedRange()._numRows;

				for (var i = numFilaCabeceraICG; i < numeroRegistrosICG; i++) {
					objetoDatosICG[0].data.push({});
					for (var j = 0; j < numeroCabecerasICG; j++) {

						cabeceraSeleccionada = String(
							excelCiclo
								.sheet("DATOS_ICG_CLIENTE")
								.row(numFilaCabecera)
								.cell(j + 1)
								.value()
						);

						cabeceraSeleccionada = cabeceraSeleccionada.toLowerCase();
						cabeceraSeleccionada = cabeceraSeleccionada.replace( / /g, "_");
						cabeceraSeleccionada = cabeceraSeleccionada.replace( /\./g, "_");

						if (cabeceraSeleccionada === undefined || cabeceraSeleccionada === "") {
							console.log(
								"Error de cabecera: i=" + i + " j=" + j
							);
							continue;
						}else{
							//Guardado del registro:
							if(excelCiclo.sheet("DATOS_ICG_CLIENTE").row(i + 1).cell(j+1).value()){
								objetoDatosICG[0].data[objetoDatosICG[0].data.length-1][cabeceraSeleccionada]= excelCiclo.sheet("DATOS_ICG_CLIENTE").row(i + 1).cell(j+1).value();
							}

						}
					}

					if(Object.keys(objetoDatosICG[0].data[objetoDatosICG[0].data.length-1]).length === 0){
						objetoDatosICG[0].data.pop();
					}
				}

				//Paso 2: Copiar datos en NO SITE:
				
				//Borrar pestaña no site:
				var numeroCaberasNoSite = excelCiclo
					.sheet("NO SITE")
					.usedRange()._numColumns;

				var numeroFilaNoSite = excelCiclo
					.sheet("NO SITE")
					.usedRange()._numRows;

				for(var i= 2;i<=numeroFilaNoSite;i++){
					for(var j= 1;j<=numeroCaberasNoSite;j++){
						excelCiclo.sheet("NO SITE").cell(i,j).value("");
					}
				}

				//Copiar datos no site:
				var cuentaNoSite=2;
				for(var i= 0;i<objetoDatosRaw[0].data.length;i++){
					if(objetoDatosRaw[0].data[i]["subcategoría"]=="No SITE"){
						excelCiclo.sheet("NO SITE").cell(cuentaNoSite,1).value(objetoDatosRaw[0].data[i]["factura"])
						excelCiclo.sheet("NO SITE").cell(cuentaNoSite,2).value(objetoDatosRaw[0].data[i]["id_smart"])
						excelCiclo.sheet("NO SITE").cell(cuentaNoSite,3).value(objetoDatosRaw[0].data[i]["error"])
						excelCiclo.sheet("NO SITE").cell(cuentaNoSite,4).value(objetoDatosRaw[0].data[i]["descripción"])
						excelCiclo.sheet("NO SITE").cell(cuentaNoSite,5).value(objetoDatosRaw[0].data[i]["importe"])
						excelCiclo.sheet("NO SITE").cell(cuentaNoSite,6).value(objetoDatosRaw[0].data[i]["categoría"])
						excelCiclo.sheet("NO SITE").cell(cuentaNoSite,7).value(objetoDatosRaw[0].data[i]["subcategoría"])
						cuentaNoSite++;
					}
				}

				//Ordenar mensajes ICR:
				
				console.log("Ordenando archivo:")
				objetoDatosICG[0].data.sort((a,b) => {
					return ((parseInt(a.fecha)+parseFloat(a.hora) < parseInt(b.fecha)+parseFloat(b.hora)) ? 1 : -1)
				})

				mainProcess.guardarDocumento(objetoDatosRaw[0])
				mainProcess.guardarDocumento(objetoDatosICG[0])

				//Mapeo de campos NO SITE - ICG
				
				numeroCaberasNoSite = excelCiclo
					.sheet("NO SITE")
					.usedRange()._numColumns;

				numeroFilaNoSite = excelCiclo
					.sheet("NO SITE")
					.usedRange()._numRows;

				var id_smart;
				for(var i= 2;i<=numeroFilaNoSite;i++){
					//Get id_smart:
					id_smart= parseInt(excelCiclo.sheet("NO SITE").cell(i,2).value())

					for(var j= 0;j<objetoDatosICG[0].data.length;j++){
						if(objetoDatosICG[0].data[j]["ctactsisex"]==id_smart){
							excelCiclo.sheet("NO SITE").cell(i,8).value(objetoDatosICG[0].data[j]["iden_msj"]);
							excelCiclo.sheet("NO SITE").cell(i,9).value(objetoDatosICG[0].data[j]["nif_o_cif"]);
							excelCiclo.sheet("NO SITE").cell(i,10).value(objetoDatosICG[0].data[j]["mét_pago"]);
							excelCiclo.sheet("NO SITE").cell(i,11).value(objetoDatosICG[0].data[j]["resp"]);
							excelCiclo.sheet("NO SITE").cell(i,12).value(objetoDatosICG[0].data[j]["car"]);
							excelCiclo.sheet("NO SITE").cell(i,13).value(objetoDatosICG[0].data[j]["estado_cc"]);
							excelCiclo.sheet("NO SITE").cell(i,14).value(objetoDatosICG[0].data[j]["estado"]);
							excelCiclo.sheet("NO SITE").cell(i,15).value(objetoDatosICG[0].data[j]["cod_mot"]);
							excelCiclo.sheet("NO SITE").cell(i,16).value(objetoDatosICG[0].data[j]["descr_motivo_estado"]);
							break;
						}
					}
				}

				function clasificarFacturacionCiclo(){
					var clasificado = false;

					//Si no hay mensaje:
					if(excelCiclo.sheet("NO SITE").cell(i,8).value()==""){
						if(String(excelCiclo.sheet("NO SITE").cell(i,2).value()).charAt(0)=="#"){
							excelCiclo.sheet("NO SITE").cell(i,17).value("SITE COMENTADO")
							clasificacionFacturacion[0]= clasificacionFacturacion[0]+1;
							clasificado= true;
						}else{
							excelCiclo.sheet("NO SITE").cell(i,17).value("NO HAY MENSAJES ICR")
							clasificacionFacturacion[1]= clasificacionFacturacion[1]+1;
							clasificado= true;
						}
					}

					//Responsable N
					if(excelCiclo.sheet("NO SITE").cell(i,11).value()=="N"){
						excelCiclo.sheet("NO SITE").cell(i,17).value("CUENTA NO RESPONSABLE")
						clasificacionFacturacion[2]= clasificacionFacturacion[2]+1;
						clasificado= true;
					}

					//Responsable S
					if(excelCiclo.sheet("NO SITE").cell(i,11).value()=="S"){
						excelCiclo.sheet("NO SITE").cell(i,17).value("CUENTA NO RESPONSABLE")
						clasificacionFacturacion[2]= clasificacionFacturacion[2]+1;
						clasificado= true;
					}

					//Jerarquia 0
					if(String(excelCiclo.sheet("NO SITE").cell(i,12).value())=="0"){
						excelCiclo.sheet("NO SITE").cell(i,17).value("CAMPO JERARQUÍA EN BLANCO (RED CHANNEL)")
						clasificacionFacturacion[3]= clasificacionFacturacion[3]+1;
						clasificado= true;
					}

					//Estado 00
					if(excelCiclo.sheet("NO SITE").cell(i,14).value()=="00"){
						excelCiclo.sheet("NO SITE").cell(i,17).value("PENDIENTE DE PROCESAR")
						clasificacionFacturacion[4]= clasificacionFacturacion[4]+1;
						clasificado= true;
					}

					//Estado 01
					if(excelCiclo.sheet("NO SITE").cell(i,14).value()=="01"){
						if(excelCiclo.sheet("NO SITE").cell(i,11).value()=="R"){
							if(excelCiclo.sheet("NO SITE").cell(i,12).value()==4){
								excelCiclo.sheet("NO SITE").cell(i,17).value("JERARQUÍA 4R")
								clasificacionFacturacion[5]= clasificacionFacturacion[5]+1;
								clasificado= true;
							}
						}
					}

					//Descripción contiene:
					if(String(excelCiclo.sheet("NO SITE").cell(i,16).value()).includes("Error imposible crear CC porque hay nifs duplicados")){
						excelCiclo.sheet("NO SITE").cell(i,17).value("CC CON NIF DUPLICADOS")
						clasificacionFacturacion[6]= clasificacionFacturacion[6]+1;
						clasificado= true;
					}
					
					//Descripción contiene:
					if(String(excelCiclo.sheet("NO SITE").cell(i,16).value()).includes("NIF")||String(excelCiclo.sheet("NO SITE").cell(i,16).value()).includes("Nº id")){
						excelCiclo.sheet("NO SITE").cell(i,17).value("ERROR NIF")
						clasificacionFacturacion[7]= clasificacionFacturacion[7]+1;
						clasificado= true;
					}

					//Descripción contiene:
					if(String(excelCiclo.sheet("NO SITE").cell(i,16).value()).includes("Id site tiene un estado no procesable")){
						excelCiclo.sheet("NO SITE").cell(i,17).value("CUENTA MAESTRA CON ERROR DE JERARQUÍA")
						clasificacionFacturacion[8]= clasificacionFacturacion[8]+1;
						clasificado= true;
					}

					//Descripción contiene:
					if(String(excelCiclo.sheet("NO SITE").cell(i,16).value()).includes("nombre")||String(excelCiclo.sheet("NO SITE").cell(i,16).value()).includes("apellidos")||String(excelCiclo.sheet("NO SITE").cell(i,16).value()).includes("Enter surname for business partner")){
						excelCiclo.sheet("NO SITE").cell(i,17).value("ERROR EN NOMBRES Y APELLIDOS")
						clasificacionFacturacion[9]= clasificacionFacturacion[9]+1;
						clasificado= true;
					}

					//Descripción contiene:
					if(String(excelCiclo.sheet("NO SITE").cell(i,16).value()).includes("Ignorado por GDPR")){
						excelCiclo.sheet("NO SITE").cell(i,17).value("GDPR")
						clasificacionFacturacion[10]= clasificacionFacturacion[10]+1;
						clasificado= true;
					}

					//Descripción contiene:
					if(String(excelCiclo.sheet("NO SITE").cell(i,16).value()).includes("Método de pago no se mapea contra condiciones de pago")){
						excelCiclo.sheet("NO SITE").cell(i,17).value("MÉTODO DE PAGO EN BLANCO")
						clasificacionFacturacion[11]= clasificacionFacturacion[11]+1;
						clasificado= true;
					}

					//Descripción contiene:
					if(String(excelCiclo.sheet("NO SITE").cell(i,16).value()).includes("Tipo de método de pago SMART no se ha podido mapear en RMCA")){
						excelCiclo.sheet("NO SITE").cell(i,17).value("MÉTODO DE PAGO PREPAGO")
						clasificacionFacturacion[12]= clasificacionFacturacion[12]+1;
						clasificado= true;
					}

					//Descripción contiene:
					if(String(excelCiclo.sheet("NO SITE").cell(i,16).value()).includes("Clase de cuenta de contrato no definida")){
						excelCiclo.sheet("NO SITE").cell(i,17).value("TIPO DE CLIENTE NO INFORMADO")
						clasificacionFacturacion[13]= clasificacionFacturacion[13]+1;
						clasificado= true;
					}

					//Descripción contiene:
					if(String(excelCiclo.sheet("NO SITE").cell(i,16).value()).includes("Estado del cliente en SMART no se ha podido mapear con el estado")){
						excelCiclo.sheet("NO SITE").cell(i,17).value("ERROR EN EL MAPEO DE ESTADOS")
						clasificacionFacturacion[14]= clasificacionFacturacion[14]+1;
						clasificado= true;
					}

					//Descripción contiene:
					if(String(excelCiclo.sheet("NO SITE").cell(i,16).value()).includes("Introduzca el nombre del interlocutor comercial")){
						excelCiclo.sheet("NO SITE").cell(i,17).value("INTERLOCUTOR COMERCIAL NO INFORMADO")
						clasificacionFacturacion[15]= clasificacionFacturacion[15]+1;
						clasificado= true;
					}

					//Descripción contiene:
					if(String(excelCiclo.sheet("NO SITE").cell(i,16).value()).includes("Ignorado por mismo tipo_accion con id mensaje mayor para el mismo site")){
						excelCiclo.sheet("NO SITE").cell(i,17).value("NO PROCESADO EL MENSAJE EN QUE APARECE COMO RESPONSABLE")
						clasificacionFacturacion[16]= clasificacionFacturacion[16]+1;
						clasificado= true;
					}

					return clasificado;

				}

				//Clasificacion NO SITE:
				var flagClasificado = false;
				var arrayDescripciones = [];
				var clasificacionAsignada = "";

				for(var i= 2;i<=numeroFilaNoSite;i++){

					flagClasificado=false
					arrayDescripciones=[];

					if(excelCiclo.sheet("NO SITE").cell(i,2).value()==="" || excelCiclo.sheet("NO SITE").cell(i,2).value()===undefined){
						continue;
					}

					flagClasificado = clasificarFacturacionCiclo();

					if(!flagClasificado){
						for(var k=0; k<objetoDatosICG[0].data.length; k++){

							if(objetoDatosICG[0].data[k]["ctactsisex"]==excelCiclo.sheet("NO SITE").cell(i,2).value()){
								arrayDescripciones.push(objetoDatosICG[0].data[k]["descr_motivo_estado"])
							}
						}

						//Check solo 
						var flagSoloMensajeBaja= true;
						var flagIgnorado= false;
						for(var l=0; l<arrayDescripciones.length; l++){
							if(arrayDescripciones[l]!="No existe CC asociada al ID SMART"){
								flagSoloMensajeBaja=false;
								if(l==2 && arrayDescripciones[l]=="Ignorado por mismo tipo_accion con id mensaje mayor para el mismo site"){
									flagIgnorado=true;
								}
							}
						}

						if(flagSoloMensajeBaja){
							excelCiclo.sheet("NO SITE").cell(i,17).value("SOLO LLEGA MENSAJE DE BAJA")
							clasificacionFacturacion[17] = clasificacionFacturacion[17]+1 
						}else{
							//RESTO DE MENSAJES:
							excelCiclo.sheet("NO SITE").cell(i,17).value("OTROS BLOQUEOS")
							clasificacionFacturacion[18] = clasificacionFacturacion[18]+1 
						}

						if(objetoDatosICG[0].data[k]){
						console.log("ID SMART: "+objetoDatosICG[0].data[k]["ctactsisex"])
						}
						console.log(arrayDescripciones);
						console.log("Clasificación: "+ excelCiclo.sheet("NO SITE").cell(i,17).value())
					}

				}//fin for

				//Modificación archivo de seguimiento KPIs:
				return excelCiclo
					.toFileAsync(path.normalize(pathExcelCiclo))
					.then(() => {
						console.log("Fin del procesamiento excel CICLO");
						return true
					})
					.catch(err => {
						console.log("Se ha producido un error interno: ");
						console.log(err);
						var tituloError =
							"Se ha producido un error escribiendo el archivo: " +
							path.normalize(path.normalize(pathExcelCiclo));
						return false;
					});
			})
			.then(()=>{
				console.log("Proceso finalizado");

				XlsxPopulate.fromFileAsync(path.normalize(pathExcelSeguimiento))
					.then(excelSeguimiento => {
						//Tratamiento excel seguimiento:
						
						console.log("Cargando Excel Seguimiento:");

						if(excelSeguimiento===undefined){
							return false;
						}

						//Creación del Objeto:
						var objetoDatosSeguimiento = [{
							data: [],
							nombreId: "Seguimiento KPIs",
							objetoId: "Seguimiento KPIs",
						}]

						var cabeceraSeleccionada = "";

						var numeroCabecerasSeguimiento = excelSeguimiento
							.sheet("Datos KPI 367")
							.usedRange()._numColumns;

						var numeroRegistrosSeguimiento = excelSeguimiento
							.sheet("Datos KPI 367")
							.usedRange()._numRows;

						//Comprobación de inputs:
						if(excelSeguimiento.sheet("Datos KPI 367")===undefined){
							return false;
						}

						var numeroSemana = moment().isoWeek();
						var numeroYear = moment().year();

						if(numeroSemana<10){
							numeroSemana = "W0"+numeroSemana;
						}else{
							numeroSemana = "W"+numeroSemana;
						}

						console.log("Numero de semana actual: "+ numeroSemana);
						console.log("Numero de año actual: "+ numeroYear);

						//Encontrar ultima fila de seguimiento
						var filaSemanaActual= 1	
						for(var i=2; i<=numeroRegistrosSeguimiento+1; i++){
							if(excelSeguimiento.sheet("Datos KPI 367").cell(i,1).value()==numeroSemana && excelSeguimiento.sheet("Datos KPI 367").cell(i,2).value()== numeroYear){
								filaSemanaActual=i;
								break;
							}else 
								if(excelSeguimiento.sheet("Datos KPI 367").cell(i,1).value()=="" || excelSeguimiento.sheet("Datos KPI 367").cell(i,1).value()==undefined){
									filaSemanaActual=i;
									break;
								}
						}

						console.log("Fila Actual: "+ filaSemanaActual)

						//Escribir valores en tabla KPIS:
						if(excelSeguimiento.sheet("Datos KPI 367").cell(filaSemanaActual,1).value()=="" || excelSeguimiento.sheet("Datos KPI 367").cell(filaSemanaActual,1).value()==undefined){
							excelSeguimiento.sheet("Datos KPI 367").cell(filaSemanaActual,1).value(numeroSemana) 
							excelSeguimiento.sheet("Datos KPI 367").cell(filaSemanaActual,2).value(numeroYear) 
						}

						excelSeguimiento.sheet("Datos KPI 367").cell(filaSemanaActual,3).value(facturacionNoCargada) 


						var numeroRegistrosSeguimientoFacturacion = excelSeguimiento
							.sheet("CICLO")
							.usedRange()._numRows;

						var filaSemanaActualFacturacion= 1	
						for(var i=2; i<=numeroRegistrosSeguimientoFacturacion+1; i++){
							if(excelSeguimiento.sheet("CICLO").cell(i,1).value()==numeroSemana && excelSeguimiento.sheet("CICLO").cell(i,2).value()== numeroYear){
								filaSemanaActualFacturacion=i;
								break;
							}else 
								if(excelSeguimiento.sheet("CICLO").cell(i,1).value()=="" || excelSeguimiento.sheet("CICLO").cell(i,1).value()==undefined){
									filaSemanaActualFacturacion=i;
									break;
								}
						}

						//Escribir valores FACTURACION:
						if(excelSeguimiento.sheet("CICLO").cell(filaSemanaActualFacturacion,1).value()=="" || excelSeguimiento.sheet("CICLO").cell(filaSemanaActualFacturacion,1).value()==undefined){
							excelSeguimiento.sheet("CICLO").cell(filaSemanaActualFacturacion,1).value(numeroSemana) 
							excelSeguimiento.sheet("CICLO").cell(filaSemanaActualFacturacion,2).value(numeroYear) 
						}
						for(var i=0; i<clasificacionFacturacion.length;i++){
							excelSeguimiento.sheet("CICLO").cell(filaSemanaActualFacturacion,i+4).value(clasificacionFacturacion[i]) 
						}
						
						return excelSeguimiento
							.toFileAsync(path.normalize(pathExcelSeguimiento))
							.then(() => {
								console.log("Fin del procesamiento excel Seguimiento");
								return true
							})
							.catch(err => {
								console.log("Se ha producido un error interno: ");
								console.log(err);
								var tituloError =
									"Se ha producido un error escribiendo el archivo: " +
											path.normalize(path.normalize(pathExcelSeguimiento));
										return false;
									});
					})
				})

		return true;
	}

	async facturacionHotbillingStep1(argumentos){

		console.log("Importando Excel:");
		console.log("Argumentos: ");
		//console.log(argumentos);

		var pathExcelCiclo = argumentos[0];
		var pathExcelSeguimiento = argumentos[1];
		var numFilaCabecera = 1;

		//PROCESAMIENTO EXCEL Ciclo: 
		XlsxPopulate.fromFileAsync(path.normalize(pathExcelCiclo))
			.then(excelHotbilling => {
				console.log("Cargando Excel:");
				//console.log(excelHotbilling);
				//
				if(excelHotbilling===undefined){
					return false;
				}

				//Creación del Objeto:
				var objetoDatosRaw = [{
					data: [],
					nombreId: "Facturacion_Hotbilling_RAW",
					objetoId: "Facturacion_Hotbilling_RAW",
				}]

				var cabeceraSeleccionada = "";

				var numeroCaberas = excelHotbilling
					.sheet("DATOS_RAW")
					.usedRange()._numColumns;

				var numeroRegistros = excelHotbilling
					.sheet("DATOS_RAW")
					.usedRange()._numRows;

				//Comprobación de inputs:
				if(excelHotbilling.sheet("DATOS_RAW")===undefined){
					return false;
				}


				//Relleno de objeto Data:
				for (var i = numFilaCabecera; i < numeroRegistros; i++) {
					objetoDatosRaw[0].data.push({});
					for (var j = 0; j < numeroCaberas; j++) {

						cabeceraSeleccionada = String(
							excelHotbilling
								.sheet("DATOS_RAW")
								.row(numFilaCabecera)
								.cell(j + 1)
								.value()
						);

						cabeceraSeleccionada = cabeceraSeleccionada.toLowerCase();
						cabeceraSeleccionada = cabeceraSeleccionada.replace( / /g, "_");
						cabeceraSeleccionada = cabeceraSeleccionada.replace( /\./g, "_");

						if (cabeceraSeleccionada === undefined || cabeceraSeleccionada === "") {
							console.log(
								"Error de cabecera: i=" + i + " j=" + j
							);
							continue;
						}else{
							//Guardado del registro:
							if(excelHotbilling.sheet("DATOS_RAW").row(i + 1).cell(j+1).value()){
								objetoDatosRaw[0].data[objetoDatosRaw[0].data.length-1][cabeceraSeleccionada]= excelHotbilling.sheet("DATOS_RAW").row(i + 1).cell(j+1).value();
							}

						}
					}

					if(Object.keys(objetoDatosRaw[0].data[objetoDatosRaw[0].data.length-1]).length === 0){
						objetoDatosRaw[0].data.pop();
					}
				}

				//Paso 2: Copiar datos en NO SITE:
				
				//Borrar pestaña no site VENTAS:
				var numeroCaberasVentasNoSite = excelHotbilling
					.sheet("VENTAS (NO SITE)")
					.usedRange()._numColumns;

				var numeroFilaVentasNoSite = excelHotbilling
					.sheet("VENTAS (NO SITE)")
					.usedRange()._numRows;

				for(var i= 2;i<=numeroFilaVentasNoSite;i++){
					for(var j= 1;j<=numeroCaberasVentasNoSite;j++){
						excelHotbilling.sheet("VENTAS (NO SITE)").cell(i,j).value("");
					}
				}

				//Borrar pestaña no site ABONOS Y DEVOLUCIONES:
				var numeroCaberasAbonosNoSite = excelHotbilling
					.sheet("ABONOS-DEVOLUCIONES (NO SITE)")
					.usedRange()._numColumns;

				var numeroFilaAbonosNoSite = excelHotbilling
					.sheet("ABONOS-DEVOLUCIONES (NO SITE)")
					.usedRange()._numRows;

				for(var i= 2;i<=numeroFilaAbonosNoSite;i++){
					for(var j= 1;j<=numeroCaberasAbonosNoSite;j++){
						excelHotbilling.sheet("ABONOS-DEVOLUCIONES (NO SITE)").cell(i,j).value("");
					}
				}

				//Copiar datos no site:
				var cuentaNoSiteVentas=2;
				var cuentaNoSiteAbonos=2;

				for(var i= 0;i<objetoDatosRaw[0].data.length;i++){
					if(objetoDatosRaw[0].data[i]["subcategoria"]=="No SITE"){
						if(parseFloat(objetoDatosRaw[0].data[i]["total_fact"])>0){
							excelHotbilling.sheet("VENTAS (NO SITE)").cell(cuentaNoSiteVentas,1).value(objetoDatosRaw[0].data[i]["id_men_gnv"])
							excelHotbilling.sheet("VENTAS (NO SITE)").cell(cuentaNoSiteVentas,2).value(objetoDatosRaw[0].data[i]["posicion"])
							excelHotbilling.sheet("VENTAS (NO SITE)").cell(cuentaNoSiteVentas,3).value(objetoDatosRaw[0].data[i]["id_smart"])
							excelHotbilling.sheet("VENTAS (NO SITE)").cell(cuentaNoSiteVentas,4).value(objetoDatosRaw[0].data[i]["id_geneva"])
							excelHotbilling.sheet("VENTAS (NO SITE)").cell(cuentaNoSiteVentas,5).value(objetoDatosRaw[0].data[i]["nif"])
							excelHotbilling.sheet("VENTAS (NO SITE)").cell(cuentaNoSiteVentas,6).value(objetoDatosRaw[0].data[i]["n_factura"])
							excelHotbilling.sheet("VENTAS (NO SITE)").cell(cuentaNoSiteVentas,7).value(objetoDatosRaw[0].data[i]["fecha_fact"])
							excelHotbilling.sheet("VENTAS (NO SITE)").cell(cuentaNoSiteVentas,8).value(objetoDatosRaw[0].data[i]["fecha_fact_c"])
							excelHotbilling.sheet("VENTAS (NO SITE)").cell(cuentaNoSiteVentas,9).value(objetoDatosRaw[0].data[i]["year"])
							excelHotbilling.sheet("VENTAS (NO SITE)").cell(cuentaNoSiteVentas,10).value(objetoDatosRaw[0].data[i]["semana"])
							excelHotbilling.sheet("VENTAS (NO SITE)").cell(cuentaNoSiteVentas,11).value(objetoDatosRaw[0].data[i]["fecha_real_fact"])
							excelHotbilling.sheet("VENTAS (NO SITE)").cell(cuentaNoSiteVentas,12).value(objetoDatosRaw[0].data[i]["fecha_venc_fact"])
							excelHotbilling.sheet("VENTAS (NO SITE)").cell(cuentaNoSiteVentas,13).value(objetoDatosRaw[0].data[i]["total_fact"])
							excelHotbilling.sheet("VENTAS (NO SITE)").cell(cuentaNoSiteVentas,14).value(objetoDatosRaw[0].data[i]["importe"])
							excelHotbilling.sheet("VENTAS (NO SITE)").cell(cuentaNoSiteVentas,15).value(objetoDatosRaw[0].data[i]["fikey"])
							excelHotbilling.sheet("VENTAS (NO SITE)").cell(cuentaNoSiteVentas,16).value(objetoDatosRaw[0].data[i]["minor_code"])
							excelHotbilling.sheet("VENTAS (NO SITE)").cell(cuentaNoSiteVentas,17).value(objetoDatosRaw[0].data[i]["imei"])
							excelHotbilling.sheet("VENTAS (NO SITE)").cell(cuentaNoSiteVentas,18).value(objetoDatosRaw[0].data[i]["msisdn"])
							excelHotbilling.sheet("VENTAS (NO SITE)").cell(cuentaNoSiteVentas,19).value(objetoDatosRaw[0].data[i]["ot_smart"])
							excelHotbilling.sheet("VENTAS (NO SITE)").cell(cuentaNoSiteVentas,20).value(objetoDatosRaw[0].data[i]["estado"])
							excelHotbilling.sheet("VENTAS (NO SITE)").cell(cuentaNoSiteVentas,21).value(objetoDatosRaw[0].data[i]["motivo"])
							excelHotbilling.sheet("VENTAS (NO SITE)").cell(cuentaNoSiteVentas,22).value(objetoDatosRaw[0].data[i]["fecha_mensaje"])
							excelHotbilling.sheet("VENTAS (NO SITE)").cell(cuentaNoSiteVentas,23).value(objetoDatosRaw[0].data[i]["hora_mensaje"])
							excelHotbilling.sheet("VENTAS (NO SITE)").cell(cuentaNoSiteVentas,24).value(objetoDatosRaw[0].data[i]["categoria"])
							excelHotbilling.sheet("VENTAS (NO SITE)").cell(cuentaNoSiteVentas,25).value(objetoDatosRaw[0].data[i]["subcategoria"])
							cuentaNoSiteVentas++;
						}else{
							excelHotbilling.sheet("ABONOS-DEVOLUCIONES (NO SITE)").cell(cuentaNoSiteAbonos,1).value(objetoDatosRaw[0].data[i]["id_men_gnv"])
							excelHotbilling.sheet("ABONOS-DEVOLUCIONES (NO SITE)").cell(cuentaNoSiteAbonos,2).value(objetoDatosRaw[0].data[i]["posicion"])
							excelHotbilling.sheet("ABONOS-DEVOLUCIONES (NO SITE)").cell(cuentaNoSiteAbonos,3).value(objetoDatosRaw[0].data[i]["id_smart"])
							excelHotbilling.sheet("ABONOS-DEVOLUCIONES (NO SITE)").cell(cuentaNoSiteAbonos,4).value(objetoDatosRaw[0].data[i]["id_geneva"])
							excelHotbilling.sheet("ABONOS-DEVOLUCIONES (NO SITE)").cell(cuentaNoSiteAbonos,5).value(objetoDatosRaw[0].data[i]["nif"])
							excelHotbilling.sheet("ABONOS-DEVOLUCIONES (NO SITE)").cell(cuentaNoSiteAbonos,6).value(objetoDatosRaw[0].data[i]["n_factura"])
							excelHotbilling.sheet("ABONOS-DEVOLUCIONES (NO SITE)").cell(cuentaNoSiteAbonos,7).value(objetoDatosRaw[0].data[i]["fecha_fact"])
							excelHotbilling.sheet("ABONOS-DEVOLUCIONES (NO SITE)").cell(cuentaNoSiteAbonos,8).value(objetoDatosRaw[0].data[i]["fecha_fact_c"])
							excelHotbilling.sheet("ABONOS-DEVOLUCIONES (NO SITE)").cell(cuentaNoSiteAbonos,9).value(objetoDatosRaw[0].data[i]["year"])
							excelHotbilling.sheet("ABONOS-DEVOLUCIONES (NO SITE)").cell(cuentaNoSiteAbonos,10).value(objetoDatosRaw[0].data[i]["semana"])
							excelHotbilling.sheet("ABONOS-DEVOLUCIONES (NO SITE)").cell(cuentaNoSiteAbonos,11).value(objetoDatosRaw[0].data[i]["fecha_real_fact"])
							excelHotbilling.sheet("ABONOS-DEVOLUCIONES (NO SITE)").cell(cuentaNoSiteAbonos,12).value(objetoDatosRaw[0].data[i]["fecha_venc_fact"])
							excelHotbilling.sheet("ABONOS-DEVOLUCIONES (NO SITE)").cell(cuentaNoSiteAbonos,13).value(objetoDatosRaw[0].data[i]["total_fact"])
							excelHotbilling.sheet("ABONOS-DEVOLUCIONES (NO SITE)").cell(cuentaNoSiteAbonos,14).value(objetoDatosRaw[0].data[i]["importe"])
							excelHotbilling.sheet("ABONOS-DEVOLUCIONES (NO SITE)").cell(cuentaNoSiteAbonos,15).value(objetoDatosRaw[0].data[i]["fikey"])
							excelHotbilling.sheet("ABONOS-DEVOLUCIONES (NO SITE)").cell(cuentaNoSiteAbonos,16).value(objetoDatosRaw[0].data[i]["minor_code"])
							excelHotbilling.sheet("ABONOS-DEVOLUCIONES (NO SITE)").cell(cuentaNoSiteAbonos,17).value(objetoDatosRaw[0].data[i]["imei"])
							excelHotbilling.sheet("ABONOS-DEVOLUCIONES (NO SITE)").cell(cuentaNoSiteAbonos,18).value(objetoDatosRaw[0].data[i]["msisdn"])
							excelHotbilling.sheet("ABONOS-DEVOLUCIONES (NO SITE)").cell(cuentaNoSiteAbonos,19).value(objetoDatosRaw[0].data[i]["ot_smart"])
							excelHotbilling.sheet("ABONOS-DEVOLUCIONES (NO SITE)").cell(cuentaNoSiteAbonos,20).value(objetoDatosRaw[0].data[i]["estado"])
							excelHotbilling.sheet("ABONOS-DEVOLUCIONES (NO SITE)").cell(cuentaNoSiteAbonos,21).value(objetoDatosRaw[0].data[i]["motivo"])
							excelHotbilling.sheet("ABONOS-DEVOLUCIONES (NO SITE)").cell(cuentaNoSiteAbonos,22).value(objetoDatosRaw[0].data[i]["fecha_mensaje"])
							excelHotbilling.sheet("ABONOS-DEVOLUCIONES (NO SITE)").cell(cuentaNoSiteAbonos,23).value(objetoDatosRaw[0].data[i]["hora_mensaje"])
							excelHotbilling.sheet("ABONOS-DEVOLUCIONES (NO SITE)").cell(cuentaNoSiteAbonos,24).value(objetoDatosRaw[0].data[i]["categoria"])
							excelHotbilling.sheet("ABONOS-DEVOLUCIONES (NO SITE)").cell(cuentaNoSiteAbonos,25).value(objetoDatosRaw[0].data[i]["subcategoria"])
							cuentaNoSiteAbonos++;
						}
					}
				}

				mainProcess.guardarDocumento(objetoDatosRaw[0])
				console.log(objetoDatosRaw[0].data[objetoDatosRaw[0].data.length-1])

				return excelHotbilling
					.toFileAsync(path.normalize(pathExcelCiclo))
					.then(() => {
						console.log("Fin del procesamiento");
						return true
					})
					.catch(err => {
						console.log("Se ha producido un error interno: ");
						console.log(err);
						var tituloError =
							"Se ha producido un error escribiendo el archivo: " +
							path.normalize(path.normalize(pathExcelCiclo));
						return false;
					});
			})
			.then(()=>{
				console.log("Proceso finalizado");
				})

		return true;
	}

	async facturacionHotbillingStep2(argumentos){

		console.log("Importando Excel:");
		console.log("Argumentos: ");
		console.log(argumentos);

		var pathExcelHotbilling = argumentos[0];
		var pathExcelSeguimiento = argumentos[1];

		var numFilaCabecera = 1;

		var ventasNoCargada=0;
		var devolucionesNoCargada=0;
		var clasificacionVentas = [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]
		var clasificacionDevoluciones = [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]

		//PROCESAMIENTO EXCEL Ciclo: 
		XlsxPopulate.fromFileAsync(path.normalize(pathExcelHotbilling))
			.then(excelHotbilling => {
				console.log("Cargando Excel:");
				//console.log(excelHotbilling);
				//
				if(excelHotbilling===undefined){
					return false;
				}

				//Creación del Objeto VENTAS NO SITE:
				var objetoDatosVentasNoSite = [{
					data: [],
					nombreId: "Facturacion_Hotbilling_Ventas_NO_SITE",
					objetoId: "Facturacion_Hotbilling_Ventas_NO_SITE",
				}]

				var cabeceraSeleccionada = "";

				var numeroCabecerasVentasNoSite = excelHotbilling
					.sheet("VENTAS (NO SITE)")
					.usedRange()._numColumns;

				var numeroRegistrosVentasNoSite = excelHotbilling
					.sheet("VENTAS (NO SITE)")
					.usedRange()._numRows;

				//Comprobación de inputs:
				if(excelHotbilling.sheet("VENTAS (NO SITE)")===undefined){
					return false;
				}

				//Calcular facturacion no cargada:
				var numeroRegistrosVentas = excelHotbilling
					.sheet("VENTAS")
					.usedRange()._numRows;

				var numeroRegistrosDevoluciones = excelHotbilling
					.sheet("ABONOS-DEVOLUCIONES")
					.usedRange()._numRows;

				for (var i = 2; i <= numeroRegistrosVentas; i++) {
					if(excelHotbilling.sheet("VENTAS").cell(i,1).value()!="" && excelHotbilling.sheet("VENTAS").cell(i,1).value()!=undefined){
						ventasNoCargada++;
					}
				}

				for (var i = 2; i <= numeroRegistrosDevoluciones; i++) {
					if(excelHotbilling.sheet("ABONOS-DEVOLUCIONES").cell(i,1).value()!="" && excelHotbilling.sheet("ABONOS-DEVOLUCIONES").cell(i,1).value()!=undefined){
						devolucionesNoCargada++;
					}
				}

				//Relleno de objeto Ventas No Site:
				for (var i = numFilaCabecera; i < numeroRegistrosVentasNoSite; i++) {
					objetoDatosVentasNoSite[0].data.push({});
					for (var j = 0; j < numeroCabecerasVentasNoSite; j++) {

						cabeceraSeleccionada = String(
							excelHotbilling
								.sheet("VENTAS (NO SITE)")
								.row(numFilaCabecera)
								.cell(j + 1)
								.value()
						);

						cabeceraSeleccionada = cabeceraSeleccionada.toLowerCase();
						cabeceraSeleccionada = cabeceraSeleccionada.replace( / /g, "_");
						cabeceraSeleccionada = cabeceraSeleccionada.replace( /\./g, "_");

						if (cabeceraSeleccionada === undefined || cabeceraSeleccionada === "") {
							console.log(
								"Error de cabecera: i=" + i + " j=" + j
							);
							continue;
						}else{
							//Guardado del registro:
							if(excelHotbilling.sheet("VENTAS (NO SITE)").row(i + 1).cell(j+1).value()){
								objetoDatosVentasNoSite[0].data[objetoDatosVentasNoSite[0].data.length-1][cabeceraSeleccionada]= excelHotbilling.sheet("VENTAS (NO SITE)").row(i + 1).cell(j+1).value();
							}

						}
					}

					if(Object.keys(objetoDatosVentasNoSite[0].data[objetoDatosVentasNoSite[0].data.length-1]).length === 0){
						objetoDatosVentasNoSite[0].data.pop();
					}
				}

				//Creación del Objeto ABONOS NO SITE:
				var objetoDatosAbonosNoSite = [{
					data: [],
					nombreId: "Facturacion_Hotbilling_Abonos_NO_SITE",
					objetoId: "Facturacion_Hotbilling_Abonos_NO_SITE",
				}]

				var cabeceraSeleccionada = "";

				var numeroCabecerasAbonosNoSite = excelHotbilling
					.sheet("ABONOS-DEVOLUCIONES (NO SITE)")
					.usedRange()._numColumns;

				var numeroRegistrosAbonosNoSite = excelHotbilling
					.sheet("ABONOS-DEVOLUCIONES (NO SITE)")
					.usedRange()._numRows;

				//Comprobación de inputs:
				if(excelHotbilling.sheet("ABONOS-DEVOLUCIONES (NO SITE)")===undefined){
					return false;
				}

				//Relleno de objeto Ventas No Site:
				for (var i = numFilaCabecera; i < numeroRegistrosAbonosNoSite; i++) {
					objetoDatosAbonosNoSite[0].data.push({});
					for (var j = 0; j < numeroCabecerasAbonosNoSite; j++) {

						cabeceraSeleccionada = String(
							excelHotbilling
								.sheet("ABONOS-DEVOLUCIONES (NO SITE)")
								.row(numFilaCabecera)
								.cell(j + 1)
								.value()
						);

						cabeceraSeleccionada = cabeceraSeleccionada.toLowerCase();
						cabeceraSeleccionada = cabeceraSeleccionada.replace( / /g, "_");
						cabeceraSeleccionada = cabeceraSeleccionada.replace( /\./g, "_");

						if (cabeceraSeleccionada === undefined || cabeceraSeleccionada === "") {
							console.log(
								"Error de cabecera: i=" + i + " j=" + j
							);
							continue;
						}else{
							//Guardado del registro:
							if(excelHotbilling.sheet("ABONOS-DEVOLUCIONES (NO SITE)").row(i + 1).cell(j+1).value()){
								objetoDatosAbonosNoSite[0].data[objetoDatosAbonosNoSite[0].data.length-1][cabeceraSeleccionada]= excelHotbilling.sheet("ABONOS-DEVOLUCIONES (NO SITE)").row(i + 1).cell(j+1).value();
							}

						}
					}

					if(Object.keys(objetoDatosAbonosNoSite[0].data[objetoDatosAbonosNoSite[0].data.length-1]).length === 0){
						objetoDatosAbonosNoSite[0].data.pop();
					}
				}

				//Rellenar datos ICG_Cliente_VENTAS:
				cabeceraSeleccionada="";
				var numFilaCabeceraVentasICG = 1;

				//Creación del Objeto:
				var objetoDatosVentasICG = [{
					data: [],
					nombreId: "Facturacion_Hotbilling_Ventas_ICG",
					objetoId: "Facturacion_Hotbilling_Ventas_ICG",
				}]

				var numeroCabecerasVentasICG = excelHotbilling
					.sheet("ICG_CLIENTE (VENTAS)")
					.usedRange()._numColumns;

				var numeroRegistrosVentasICG = excelHotbilling
					.sheet("ICG_CLIENTE (VENTAS)")
					.usedRange()._numRows;

				for (var i = numFilaCabeceraVentasICG; i < numeroRegistrosVentasICG; i++) {
					objetoDatosVentasICG[0].data.push({});
					for (var j = 0; j < numeroCabecerasVentasICG; j++) {

						cabeceraSeleccionada = String(
							excelHotbilling
								.sheet("ICG_CLIENTE (VENTAS)")
								.row(numFilaCabecera)
								.cell(j + 1)
								.value()
						);

						cabeceraSeleccionada = cabeceraSeleccionada.toLowerCase();
						cabeceraSeleccionada = cabeceraSeleccionada.replace( / /g, "_");
						cabeceraSeleccionada = cabeceraSeleccionada.replace( /\./g, "_");

						if (cabeceraSeleccionada === undefined || cabeceraSeleccionada === "") {
							console.log(
								"Error de cabecera: i=" + i + " j=" + j
							);
							continue;
						}else{
							//Guardado del registro:
							if(excelHotbilling.sheet("ICG_CLIENTE (VENTAS)").row(i + 1).cell(j+1).value()){
								objetoDatosVentasICG[0].data[objetoDatosVentasICG[0].data.length-1][cabeceraSeleccionada]= excelHotbilling.sheet("ICG_CLIENTE (VENTAS)").row(i + 1).cell(j+1).value();
							}

						}
					}

					if(Object.keys(objetoDatosVentasICG[0].data[objetoDatosVentasICG[0].data.length-1]).length === 0){
						objetoDatosVentasICG[0].data.pop();
					}
				}
				
				//Rellenar datos ICG_Cliente_ABONOS:
				cabeceraSeleccionada = "";
				var numFilaCabeceraAbonosICG = 1;

				//Creación del Objeto:
				var objetoDatosAbonosICG = [{
					data: [],
					nombreId: "Facturacion_Hotbilling_Abonos_ICG",
					objetoId: "Facturacion_Hotbilling_Abonos_ICG",
				}]

				var numeroCabecerasAbonosICG = excelHotbilling
					.sheet("ICG_CLIENTE (ABONOS Y DEV)")
					.usedRange()._numColumns;

				var numeroRegistrosAbonosICG = excelHotbilling
					.sheet("ICG_CLIENTE (ABONOS Y DEV)")
					.usedRange()._numRows;

				for (var i = numFilaCabeceraAbonosICG; i < numeroRegistrosAbonosICG; i++) {
					objetoDatosAbonosICG[0].data.push({});
					for (var j = 0; j < numeroCabecerasAbonosICG; j++) {

						cabeceraSeleccionada = String(
							excelHotbilling
								.sheet("ICG_CLIENTE (ABONOS Y DEV)")
								.row(numFilaCabecera)
								.cell(j + 1)
								.value()
						);

						cabeceraSeleccionada = cabeceraSeleccionada.toLowerCase();
						cabeceraSeleccionada = cabeceraSeleccionada.replace( / /g, "_");
						cabeceraSeleccionada = cabeceraSeleccionada.replace( /\./g, "_");

						if (cabeceraSeleccionada === undefined || cabeceraSeleccionada === "") {
							console.log(
								"Error de cabecera: i=" + i + " j=" + j
							);
							continue;
						}else{
							//Guardado del registro:
							if(excelHotbilling.sheet("ICG_CLIENTE (ABONOS Y DEV)").row(i + 1).cell(j+1).value()){
								objetoDatosAbonosICG[0].data[objetoDatosAbonosICG[0].data.length-1][cabeceraSeleccionada]= excelHotbilling.sheet("ICG_CLIENTE (ABONOS Y DEV)").row(i + 1).cell(j+1).value();
							}

						}
					}

					if(Object.keys(objetoDatosAbonosICG[0].data[objetoDatosAbonosICG[0].data.length-1]).length === 0){
						objetoDatosAbonosICG[0].data.pop();
					}
				}

				//Paso 2: Copiar datos en NO SITE VENTAS:
				
				//Ordenar mensajes ICR:
				
				console.log("Ordenando archivo:")
				objetoDatosVentasICG[0].data.sort((a,b) => {
					return ((parseInt(a.fecha)+parseFloat(a.hora) < parseInt(b.fecha)+parseFloat(b.hora)) ? 1 : -1)
				})

				console.log("Ordenando archivo:")
				objetoDatosAbonosICG[0].data.sort((a,b) => {
					return ((parseInt(a.fecha)+parseFloat(a.hora) < parseInt(b.fecha)+parseFloat(b.hora)) ? 1 : -1)
				})

				mainProcess.guardarDocumento(objetoDatosVentasNoSite[0])
				mainProcess.guardarDocumento(objetoDatosAbonosNoSite[0])
				mainProcess.guardarDocumento(objetoDatosAbonosICG[0])
				mainProcess.guardarDocumento(objetoDatosVentasICG[0])

				//Mapeo de campos NO SITE - ICG

				function mapeoNoSite(hojaMapeo,objetoICG){
					var numeroCaberasMapeo = excelHotbilling
						.sheet(hojaMapeo)
						.usedRange()._numColumns;

					var numeroFilaMapeo = excelHotbilling
						.sheet(hojaMapeo)
						.usedRange()._numRows;

					var id_smart;
					console.log(objetoICG)
					console.log("OBJETO ICG")
					for(var i= 2;i<=numeroFilaMapeo;i++){

						//Get id_smart:
						id_smart= parseInt(excelHotbilling.sheet(hojaMapeo).cell(i,3).value())

						for(var j= 0;j<objetoICG.length;j++){
							if(objetoICG[j]["cta_contr_sist_exis"]==id_smart){
								excelHotbilling.sheet(hojaMapeo).cell(i,28).value(objetoICG[j]["identificador_del_mensaje"]);
								excelHotbilling.sheet(hojaMapeo).cell(i,29).value(objetoICG[j]["nif_o_cif_asociado_al_cliente"]);
								excelHotbilling.sheet(hojaMapeo).cell(i,30).value(objetoICG[j]["método_pago"]);
								excelHotbilling.sheet(hojaMapeo).cell(i,31).value(objetoICG[j]["responsable"]);
								excelHotbilling.sheet(hojaMapeo).cell(i,32).value(objetoICG[j]["carácter_1"]);
								excelHotbilling.sheet(hojaMapeo).cell(i,33).value(objetoICG[j]["estado_cc_en_smart"]);
								excelHotbilling.sheet(hojaMapeo).cell(i,34).value(objetoICG[j]["estado"]);
								excelHotbilling.sheet(hojaMapeo).cell(i,35).value(objetoICG[j]["cod_mot_estado"]);
								excelHotbilling.sheet(hojaMapeo).cell(i,36).value(objetoICG[j]["descr_motivo_estado"]);
								break;
							}
						}
					}
				}

				mapeoNoSite("VENTAS (NO SITE)",objetoDatosVentasICG[0].data)
				mapeoNoSite("ABONOS-DEVOLUCIONES (NO SITE)",objetoDatosAbonosICG[0].data)

				function clasificarFacturacionHotbilling(hojaMapeo){

					var clasificado = false;

					//Si no hay mensaje:
					if(excelHotbilling.sheet(hojaMapeo).cell(i,28).value()==""){
						if(String(excelHotbilling.sheet(hojaMapeo).cell(i,3).value()).charAt(0)=="#"){
							excelHotbilling.sheet(hojaMapeo).cell(i,37).value("SITE COMENTADO")
							if(hojaMapeo=="VENTAS (NO SITE)"){
								clasificacionVentas[0]= clasificacionVentas[0]+1;
							}else{
								clasificacionDevoluciones[0]= clasificacionDevoluciones[0]+1;
							}
							clasificado= true;
						}else{
							excelHotbilling.sheet(hojaMapeo).cell(i,37).value("NO HAY MENSAJES ICR")
							if(hojaMapeo=="VENTAS (NO SITE)"){
								clasificacionVentas[1]= clasificacionVentas[1]+1;
							}else{
								clasificacionDevoluciones[1]= clasificacionDevoluciones[1]+1;
							}
							clasificado= true;
						}
					}

					//Responsable N
					if(excelHotbilling.sheet(hojaMapeo).cell(i,31).value()=="N"){
						excelHotbilling.sheet(hojaMapeo).cell(i,37).value("CUENTA NO RESPONSABLE")
							if(hojaMapeo=="VENTAS (NO SITE)"){
								clasificacionVentas[2]= clasificacionVentas[2]+1;
							}else{
								clasificacionDevoluciones[2]= clasificacionDevoluciones[2]+1;
							}
						clasificado= true;
					}

					//Responsable S
					if(excelHotbilling.sheet(hojaMapeo).cell(i,31).value()=="S"){
						excelHotbilling.sheet(hojaMapeo).cell(i,37).value("CUENTA NO RESPONSABLE")
							if(hojaMapeo=="VENTAS (NO SITE)"){
								clasificacionVentas[2]= clasificacionVentas[2]+1;
							}else{
								clasificacionDevoluciones[2]= clasificacionDevoluciones[2]+1;
							}
						clasificado= true;
					}

					//Jerarquia 0
					if(String(excelHotbilling.sheet(hojaMapeo).cell(i,32).value())=="0"){
						excelHotbilling.sheet(hojaMapeo).cell(i,37).value("CAMPO JERARQUÍA EN BLANCO (RED CHANNEL)")
						if(hojaMapeo=="VENTAS (NO SITE)"){
							clasificacionVentas[3]= clasificacionVentas[3]+1;
						}else{
							clasificacionDevoluciones[3]= clasificacionDevoluciones[3]+1;
						}
						clasificado= true;
					}

					//Estado 00
					if(excelHotbilling.sheet(hojaMapeo).cell(i,34).value()=="00"){
						excelHotbilling.sheet(hojaMapeo).cell(i,37).value("PENDIENTE DE PROCESAR")
						if(hojaMapeo=="VENTAS (NO SITE)"){
							clasificacionVentas[4]= clasificacionVentas[4]+1;
						}else{
							clasificacionDevoluciones[4]= clasificacionDevoluciones[4]+1;
						}
						clasificado= true;
					}

					//Estado 01
					if(excelHotbilling.sheet(hojaMapeo).cell(i,34).value()=="01"){
						if(excelHotbilling.sheet(hojaMapeo).cell(i,31).value()=="R"){
							if(excelHotbilling.sheet(hojaMapeo).cell(i,32).value()==4){
								excelHotbilling.sheet(hojaMapeo).cell(i,37).value("JERARQUÍA 4R")
						if(hojaMapeo=="VENTAS (NO SITE)"){
							clasificacionVentas[5]= clasificacionVentas[5]+1;
						}else{
							clasificacionDevoluciones[5]= clasificacionDevoluciones[5]+1;
						}
								clasificado= true;
							}
						}
					}

					//Descripción contiene:
					if(String(excelHotbilling.sheet(hojaMapeo).cell(i,36).value()).includes("Error imposible crear CC porque hay nifs duplicados")){
						excelHotbilling.sheet(hojaMapeo).cell(i,37).value("CC CON NIF DUPLICADOS")
						if(hojaMapeo=="VENTAS (NO SITE)"){
							clasificacionVentas[6]= clasificacionVentas[6]+1;
						}else{
							clasificacionDevoluciones[6]= clasificacionDevoluciones[6]+1;
						}
						clasificado= true;
					}
					
					//Descripción contiene:
					if(String(excelHotbilling.sheet(hojaMapeo).cell(i,36).value()).includes("NIF")||String(excelHotbilling.sheet(hojaMapeo).cell(i,36).value()).includes("Nº id")){
						excelHotbilling.sheet(hojaMapeo).cell(i,37).value("ERROR NIF")
						if(hojaMapeo=="VENTAS (NO SITE)"){
							clasificacionVentas[7]= clasificacionVentas[7]+1;
						}else{
							clasificacionDevoluciones[7]= clasificacionDevoluciones[7]+1;
						}
						clasificado= true;
					}

					//Descripción contiene:
					if(String(excelHotbilling.sheet(hojaMapeo).cell(i,36).value()).includes("Id site tiene un estado no procesable")){
						excelHotbilling.sheet(hojaMapeo).cell(i,37).value("CUENTA MAESTRA CON ERROR DE JERARQUÍA")
						if(hojaMapeo=="VENTAS (NO SITE)"){
							clasificacionVentas[8]= clasificacionVentas[8]+1;
						}else{
							clasificacionDevoluciones[8]= clasificacionDevoluciones[8]+1;
						}
						clasificado= true;
					}

					//Descripción contiene:
					if(String(excelHotbilling.sheet(hojaMapeo).cell(i,36).value()).includes("nombre")||String(excelHotbilling.sheet(hojaMapeo).cell(i,36).value()).includes("apellidos")||String(excelHotbilling.sheet(hojaMapeo).cell(i,36).value()).includes("Enter surname for business partner")){
						excelHotbilling.sheet(hojaMapeo).cell(i,37).value("ERROR EN NOMBRES Y APELLIDOS")
						if(hojaMapeo=="VENTAS (NO SITE)"){
							clasificacionVentas[9]= clasificacionVentas[9]+1;
						}else{
							clasificacionDevoluciones[9]= clasificacionDevoluciones[9]+1;
						}
						clasificado= true;
					}

					//Descripción contiene:
					if(String(excelHotbilling.sheet(hojaMapeo).cell(i,36).value()).includes("Ignorado por GDPR")){
						excelHotbilling.sheet(hojaMapeo).cell(i,37).value("GDPR")
						if(hojaMapeo=="VENTAS (NO SITE)"){
							clasificacionVentas[10]= clasificacionVentas[10]+1;
						}else{
							clasificacionDevoluciones[10]= clasificacionDevoluciones[10]+1;
						}
						clasificado= true;
					}

					//Descripción contiene:
					if(String(excelHotbilling.sheet(hojaMapeo).cell(i,36).value()).includes("Método de pago no se mapea contra condiciones de pago")){
						excelHotbilling.sheet(hojaMapeo).cell(i,37).value("MÉTODO DE PAGO EN BLANCO")
						if(hojaMapeo=="VENTAS (NO SITE)"){
							clasificacionVentas[11]= clasificacionVentas[11]+1;
						}else{
							clasificacionDevoluciones[11]= clasificacionDevoluciones[11]+1;
						}
						clasificado= true;
					}

					//Descripción contiene:
					if(String(excelHotbilling.sheet(hojaMapeo).cell(i,36).value()).includes("Tipo de método de pago SMART no se ha podido mapear en RMCA")){
						excelHotbilling.sheet(hojaMapeo).cell(i,37).value("MÉTODO DE PAGO PREPAGO")
						if(hojaMapeo=="VENTAS (NO SITE)"){
							clasificacionVentas[12]= clasificacionVentas[12]+1;
						}else{
							clasificacionDevoluciones[12]= clasificacionDevoluciones[12]+1;
						}
						clasificado= true;
					}

					//Descripción contiene:
					if(String(excelHotbilling.sheet(hojaMapeo).cell(i,36).value()).includes("Clase de cuenta de contrato no definida")){
						excelHotbilling.sheet(hojaMapeo).cell(i,37).value("TIPO DE CLIENTE NO INFORMADO")
						if(hojaMapeo=="VENTAS (NO SITE)"){
							clasificacionVentas[13]= clasificacionVentas[13]+1;
						}else{
							clasificacionDevoluciones[13]= clasificacionDevoluciones[13]+1;
						}
						clasificado= true;
					}

					//Descripción contiene:
					if(String(excelHotbilling.sheet(hojaMapeo).cell(i,36).value()).includes("Estado del cliente en SMART no se ha podido mapear con el estado")){
						excelHotbilling.sheet(hojaMapeo).cell(i,37).value("ERROR EN EL MAPEO DE ESTADOS")
						if(hojaMapeo=="VENTAS (NO SITE)"){
							clasificacionVentas[14]= clasificacionVentas[14]+1;
						}else{
							clasificacionDevoluciones[14]= clasificacionDevoluciones[14]+1;
						}
						clasificado= true;
					}

					//Descripción contiene:
					if(String(excelHotbilling.sheet(hojaMapeo).cell(i,36).value()).includes("Introduzca el nombre del interlocutor comercial")){
						excelHotbilling.sheet(hojaMapeo).cell(i,37).value("INTERLOCUTOR COMERCIAL NO INFORMADO")
						if(hojaMapeo=="VENTAS (NO SITE)"){
							clasificacionVentas[15]= clasificacionVentas[15]+1;
						}else{
							clasificacionDevoluciones[15]= clasificacionDevoluciones[15]+1;
						}
						clasificado= true;
					}

					//Descripción contiene:
					if(String(excelHotbilling.sheet(hojaMapeo).cell(i,36).value()).includes("Ignorado por mismo tipo_accion con id mensaje mayor para el mismo site")){
						excelHotbilling.sheet(hojaMapeo).cell(i,37).value("NO PROCESADO EL MENSAJE EN QUE APARECE COMO RESPONSABLE")
						if(hojaMapeo=="VENTAS (NO SITE)"){
							clasificacionVentas[16]= clasificacionVentas[16]+1;
						}else{
							clasificacionDevoluciones[16]= clasificacionDevoluciones[16]+1;
						}
						clasificado= true;
					}

					return clasificado;
				}

				//Clasificacion NO SITE VENTAS:
				var flagClasificadoVentas = false;
				var arrayDescripcionesVentas = [];
				var clasificacionAsignadaVentas = "";
				var hojaClasificacionVentas = "VENTAS (NO SITE)"

				for(var i= 2;i<=numeroRegistrosVentasNoSite;i++){

					flagClasificadoVentas=false
					arrayDescripcionesVentas=[];

					if(excelHotbilling.sheet("VENTAS (NO SITE)").cell(i,3).value()==="" ||  excelHotbilling.sheet("VENTAS (NO SITE)").cell(i,3).value()===undefined){
						console.log("SKIP")
						continue;
					}

					flagClasificadoVentas = clasificarFacturacionHotbilling(hojaClasificacionVentas);

					if(!flagClasificadoVentas){
						for(var k=0; k<objetoDatosVentasICG[0].data.length; k++){

							if(objetoDatosVentasICG[0].data[k]["cta_contr_sist_exis"]==excelHotbilling.sheet("VENTAS (NO SITE)").cell(i,3).value()){
								arrayDescripcionesVentas.push(objetoDatosVentasICG[0].data[k]["descr_motivo_estado"])
							}
						}

						//Check solo 
						var flagSoloMensajeBajaVentas= true;
						var flagIgnoradoVentas= false;
						for(var l=0; l<arrayDescripcionesVentas.length; l++){
							if(arrayDescripcionesVentas[l]!="No existe CC asociada al ID SMART"){
								flagSoloMensajeBajaVentas=false;
								if(l==2 && arrayDescripcionesVentas[l]=="Ignorado por mismo tipo_accion con id mensaje mayor para el mismo site"){
									flagIgnoradoVentas=true;
								}
							}
						}

						if(flagSoloMensajeBajaVentas){
							excelHotbilling.sheet("VENTAS (NO SITE)").cell(i,37).value("SOLO LLEGA MENSAJE DE BAJA")
							clasificacionVentas[17] = clasificacionVentas[17]+1 
						}else{
							//RESTO DE MENSAJES:
							excelHotbilling.sheet("VENTAS (NO SITE)").cell(i,37).value("OTROS BLOQUEOS")
							clasificacionVentas[18] = clasificacionVentas[18]+1 
						}

						if(objetoDatosVentasICG[0].data[k]){
						console.log("ID SMART: "+objetoDatosVentasICG[0].data[k]["cta_contr_sist_exis"])
						}
						console.log(arrayDescripcionesVentas);
						console.log("Clasificación: "+ excelHotbilling.sheet("VENTAS (NO SITE)").cell(i,37).value())
					}

				}//fin for

				//Clasificacion NO SITE ABONOS:
				var flagClasificadoAbonos = false;
				var arrayDescripcionesAbonos = [];
				var clasificacionAsignadaAbonos = "";
				var hojaClasificacionAbonos = "ABONOS-DEVOLUCIONES (NO SITE)"

				for(var i= 2;i<=numeroRegistrosAbonosNoSite;i++){

					flagClasificadoAbonos=false
					arrayDescripcionesAbonos=[];

					if(excelHotbilling.sheet("ABONOS-DEVOLUCIONES (NO SITE)").cell(i,3).value()==="" || excelHotbilling.sheet("ABONOS-DEVOLUCIONES (NO SITE)").cell(i,3).value()===undefined){
						continue;
					}

					flagClasificadoAbonos = clasificarFacturacionHotbilling(hojaClasificacionAbonos);

					if(!flagClasificadoAbonos){
						for(var k=0; k<objetoDatosAbonosICG[0].data.length; k++){

							if(objetoDatosAbonosICG[0].data[k]["cta_contr_sist_exis"]==excelHotbilling.sheet("ABONOS-DEVOLUCIONES (NO SITE)").cell(i,3).value()){
								arrayDescripcionesAbonos.push(objetoDatosAbonosICG[0].data[k]["descr_motivo_estado"])
							}
						}

						//Check solo 
						var flagSoloMensajeBajaAbonos= true;
						var flagIgnoradoAbonos= false;
						for(var l=0; l<arrayDescripcionesAbonos.length; l++){
							if(arrayDescripcionesAbonos[l]!="No existe CC asociada al ID SMART"){
								flagSoloMensajeBajaAbonos=false;
								if(l==2 && arrayDescripcionesAbonos[l]=="Ignorado por mismo tipo_accion con id mensaje mayor para el mismo site"){
									flagIgnoradoAbonos=true;
								}
							}
						}

						if(flagSoloMensajeBajaAbonos){
							excelHotbilling.sheet("ABONOS-DEVOLUCIONES (NO SITE)").cell(i,37).value("SOLO LLEGA MENSAJE DE BAJA")
							clasificacionDevoluciones[17] = clasificacionDevoluciones[17]+1 
						}else{
							//RESTO DE MENSAJES:
							excelHotbilling.sheet("ABONOS-DEVOLUCIONES (NO SITE)").cell(i,37).value("OTROS BLOQUEOS")
							clasificacionDevoluciones[18] = clasificacionDevoluciones[18]+1 
						}

						if(objetoDatosAbonosICG[0].data[k]){
						console.log("ID SMART: "+objetoDatosAbonosICG[0].data[k]["ctactsisex"])
						}
						console.log(arrayDescripcionesAbonos);
						console.log("Clasificación: "+ excelHotbilling.sheet("ABONOS-DEVOLUCIONES").cell(i,37).value())
					}

				}//fin for

				

				return excelHotbilling
					.toFileAsync(path.normalize(pathExcelHotbilling))
					.then(() => {
						console.log("Fin del procesamiento");
						return true
					})
					.catch(err => {
						console.log("Se ha producido un error interno: ");
						console.log(err);
						var tituloError =
							"Se ha producido un error escribiendo el archivo: " +
							path.normalize(path.normalize(pathExcelHotbilling));
						return false;
					});
			})
			.then(()=>{
				console.log("Proceso finalizado");

				//Modificación de Archivo seguimiento:	
				XlsxPopulate.fromFileAsync(path.normalize(pathExcelSeguimiento))
					.then(excelSeguimiento => {
						//Tratamiento excel seguimiento:
						
						console.log("Cargando Excel Seguimiento:");

						if(excelSeguimiento===undefined){
							return false;
						}

						//Creación del Objeto:
						var objetoDatosSeguimiento = [{
							data: [],
							nombreId: "Seguimiento KPIs",
							objetoId: "Seguimiento KPIs",
						}]

						var cabeceraSeleccionada = "";

						var numeroCabecerasSeguimiento = excelSeguimiento
							.sheet("Datos KPI 367")
							.usedRange()._numColumns;

						var numeroRegistrosSeguimiento = excelSeguimiento
							.sheet("Datos KPI 367")
							.usedRange()._numRows;

						//Comprobación de inputs:
						if(excelSeguimiento.sheet("Datos KPI 367")===undefined){
							return false;
						}

						var numeroSemana = moment().isoWeek();
						var numeroYear = moment().year();

						if(numeroSemana<10){
							numeroSemana = "W0"+numeroSemana;
						}else{
							numeroSemana = "W"+numeroSemana;
						}

						console.log("Numero de semana actual: "+ numeroSemana);
						console.log("Numero de año actual: "+ numeroYear);

						//Encontrar ultima fila de seguimiento
						var filaSemanaActual= 1	
						for(var i=2; i<=numeroRegistrosSeguimiento+1; i++){
							if(excelSeguimiento.sheet("Datos KPI 367").cell(i,1).value()==numeroSemana && excelSeguimiento.sheet("Datos KPI 367").cell(i,2).value()== numeroYear){
								filaSemanaActual=i;
								break;
							}else 
								if(excelSeguimiento.sheet("Datos KPI 367").cell(i,1).value()=="" || excelSeguimiento.sheet("Datos KPI 367").cell(i,1).value()==undefined){
									filaSemanaActual=i;
									break;
								}
						}

						console.log("Fila Actual: "+ filaSemanaActual)

						//Escribir valores en tabla KPIS:
						if(excelSeguimiento.sheet("Datos KPI 367").cell(filaSemanaActual,1).value()=="" || excelSeguimiento.sheet("Datos KPI 367").cell(filaSemanaActual,1).value()==undefined){
							excelSeguimiento.sheet("Datos KPI 367").cell(filaSemanaActual,1).value(numeroSemana) 
							excelSeguimiento.sheet("Datos KPI 367").cell(filaSemanaActual,2).value(numeroYear) 
						}

						excelSeguimiento.sheet("Datos KPI 367").cell(filaSemanaActual,4).value(devolucionesNoCargada) 
						excelSeguimiento.sheet("Datos KPI 367").cell(filaSemanaActual,5).value(ventasNoCargada) 

						//Clasificacion Ventas:
						var numeroRegistrosSeguimientoVentas = excelSeguimiento
							.sheet("VENTAS")
							.usedRange()._numRows;

						var filaSemanaActualVentas= 1	

						for(var i=2; i<=numeroRegistrosSeguimientoVentas+1; i++){
							if(excelSeguimiento.sheet("VENTAS").cell(i,1).value()==numeroSemana && excelSeguimiento.sheet("VENTAS").cell(i,2).value()== numeroYear){
								filaSemanaActualVentas=i;
								break;
							}else 
								if(excelSeguimiento.sheet("VENTAS").cell(i,1).value()=="" || excelSeguimiento.sheet("VENTAS").cell(i,1).value()==undefined){
									filaSemanaActualVentas=i;
									break;
								}
						}

						//Escribir valores FACTURACION:
						if(excelSeguimiento.sheet("VENTAS").cell(filaSemanaActualVentas,1).value()=="" || excelSeguimiento.sheet("VENTAS").cell(filaSemanaActualVentas,1).value()==undefined){
							excelSeguimiento.sheet("VENTAS").cell(filaSemanaActualVentas,1).value(numeroSemana) 
							excelSeguimiento.sheet("VENTAS").cell(filaSemanaActualVentas,2).value(numeroYear) 
						}
						for(var i=0; i<clasificacionVentas.length;i++){
							excelSeguimiento.sheet("VENTAS").cell(filaSemanaActualVentas,i+4).value(clasificacionVentas[i]) 
						}

						//Clasificacion Abonos:
						var numeroRegistrosSeguimientoDevoluciones = excelSeguimiento
							.sheet("ABONOS Y DEVOLUCIONES")
							.usedRange()._numRows;

						var filaSemanaActualDevoluciones= 1	

						for(var i=2; i<=numeroRegistrosSeguimientoDevoluciones+1; i++){
							if(excelSeguimiento.sheet("ABONOS Y DEVOLUCIONES").cell(i,1).value()==numeroSemana && excelSeguimiento.sheet("ABONOS Y DEVOLUCIONES").cell(i,2).value()== numeroYear){
								filaSemanaActualDevoluciones=i;
								break;
							}else 
								if(excelSeguimiento.sheet("ABONOS Y DEVOLUCIONES").cell(i,1).value()=="" || excelSeguimiento.sheet("ABONOS Y DEVOLUCIONES").cell(i,1).value()==undefined){
									filaSemanaActualDevoluciones=i;
									break;
								}
						}

						//Escribir valores FACTURACION:
						if(excelSeguimiento.sheet("ABONOS Y DEVOLUCIONES").cell(filaSemanaActualDevoluciones,1).value()=="" || excelSeguimiento.sheet("ABONOS Y DEVOLUCIONES").cell(filaSemanaActualDevoluciones,1).value()==undefined){
							excelSeguimiento.sheet("ABONOS Y DEVOLUCIONES").cell(filaSemanaActualDevoluciones,1).value(numeroSemana) 
							excelSeguimiento.sheet("ABONOS Y DEVOLUCIONES").cell(filaSemanaActualDevoluciones,2).value(numeroYear) 
						}

						for(var i=0; i<clasificacionDevoluciones.length;i++){
							excelSeguimiento.sheet("ABONOS Y DEVOLUCIONES").cell(filaSemanaActualDevoluciones,i+4).value(clasificacionDevoluciones[i]) 
						}
						
						return excelSeguimiento
							.toFileAsync(path.normalize(pathExcelSeguimiento))
							.then(() => {
								console.log("Fin del procesamiento excel Seguimiento");
								return true
							})
							.catch(err => {
								console.log("Se ha producido un error interno: ");
								console.log(err);
								var tituloError =
									"Se ha producido un error escribiendo el archivo: " +
											path.normalize(path.normalize(pathExcelSeguimiento));
										return false;
									});
					})
				})

		return true;
	}

	async cargarExcel(pathExcel){
		return new Promise((resolve)=>{
			XlsxPopulate.fromFileAsync(path.normalize(pathExcel))
					.then(excel=> {
						if(excel===undefined){
							resolve(false)
						}
						resolve(excel);
					})
		});
	}
			
	async financiaciones(argumentos){

		console.log("Importando Excel:");
		console.log("Argumentos: ");
		console.log(argumentos);

		var pathExcelFinanciaciones = argumentos[0];
		var pathExcelFinanciacionesPasada = argumentos[1];
		var pathExcelSeguimiento = argumentos[2];

		var numFilaCabecera = 1;

		// **************************
		// Importar Objeto de Week -1	
		// **************************
		
		console.log("Cargando Excel Financiaciones W-1...")
		var excelFinanciacionesPasada;
		excelFinanciacionesPasada = await this.cargarExcel(path.normalize(pathExcelFinanciacionesPasada))
		console.log("Cargado OK")

		//PROCESAMIENTO EXCEL FINANCIACIONES: 
		return await XlsxPopulate.fromFileAsync(path.normalize(pathExcelFinanciaciones))
			.then(excelFinanciaciones => {

				console.log("Cargando Excel:");
				//console.log(excelHotbilling);
				//
				if(excelFinanciaciones===undefined){
					return false;
				}


				// **************************
				//	Creación del Objeto DFKKOP-1W:
				// **************************
				
				//RELLENAR DFKKOP SEMANA W-1:
				var numeroRegistrosDFKKOPSemanaPasada = excelFinanciaciones
					.sheet("DFKKOP-1W")
					.usedRange()._numRows;

				var numeroCabecerasDFKKOPSemanaPasada = excelFinanciaciones
					.sheet("DFKKOP-1W")
					.usedRange()._numColumns;
				
				//Limpiar registros:
				for (var i = 2; i <= numeroRegistrosDFKKOPSemanaPasada; i++) {
					for (var j = 1; j <= numeroCabecerasDFKKOPSemanaPasada; j++) {
						excelFinanciaciones.sheet("DFKKOP-1W").cell(i,j).value("")
					}
				}

				//PEGAR VALORES EN DFKKOP-1W:
				numeroRegistrosDFKKOPSemanaPasada = excelFinanciacionesPasada
					.sheet("DFKKOP")
					.usedRange()._numRows;

				numeroCabecerasDFKKOPSemanaPasada = excelFinanciacionesPasada
					.sheet("DFKKOP")
					.usedRange()._numColumns;

				for (var i = 2; i <= numeroRegistrosDFKKOPSemanaPasada; i++) {
					for (var j = 1; j <= numeroCabecerasDFKKOPSemanaPasada; j++) {
						excelFinanciaciones.sheet("DFKKOP-1W").cell(i,j).value(excelFinanciacionesPasada.sheet("DFKKOP").cell(i,j).value())
					}
				}

				var cabeceraSeleccionada = "";

				// **************************
				//	Creación del Objeto DFKKOP:
				// **************************

				var objetoDatosDFKKOP = [{
					data: [],
					nombreId: "DFKKOP - Financiaciones",
					objetoId: "DFKKOP - Financiaciones",
				}]

				cabeceraSeleccionada = "";

				var numeroCabecerasDFKKOP = excelFinanciaciones
					.sheet("DFKKOP")
					.usedRange()._numColumns;

				var numeroRegistrosDFKKOP = excelFinanciaciones
					.sheet("DFKKOP")
					.usedRange()._numRows;

				//Comprobación de inputs:
				if(excelFinanciaciones.sheet("DFKKOP")===undefined){
					return false;
				}

				//Relleno de objeto Ventas No Site:
				for (var i = numFilaCabecera; i < numeroRegistrosDFKKOP; i++) {
					objetoDatosDFKKOP[0].data.push({});
					for (var j = 0; j < numeroCabecerasDFKKOP; j++) {

						cabeceraSeleccionada = String(
							excelFinanciaciones
								.sheet("DFKKOP")
								.row(1)
								.cell(j + 1)
								.value()
						);

						cabeceraSeleccionada = cabeceraSeleccionada.toLowerCase();
						cabeceraSeleccionada = cabeceraSeleccionada.replace( / /g, "_");
						cabeceraSeleccionada = cabeceraSeleccionada.replace( /\./g, "_");

						if (cabeceraSeleccionada === undefined || cabeceraSeleccionada === "") {
							console.log(
								"Error de cabecera: i=" + i + " j=" + j
							);
							continue;
						}else{
							//Guardado del registro:
							if(excelFinanciaciones.sheet("DFKKOP").row(i + 1).cell(j+1).value()){
								objetoDatosDFKKOP[0].data[objetoDatosDFKKOP[0].data.length-1][cabeceraSeleccionada]= excelFinanciaciones.sheet("DFKKOP").row(i + 1).cell(j+1).value();
							}

						}
					}

					//Elimina registros vacios:
					if(Object.keys(objetoDatosDFKKOP[0].data[objetoDatosDFKKOP[0].data.length-1]).length === 0){
						objetoDatosDFKKOP[0].data.pop();
					}
				}

				// **************************
				//	Creación del Objeto DFKKOP-1W:
				// **************************

				var objetoDatosDFKKOPPasada = [{
					data: [],
					nombreId: "DFKKOP-1W - Financiaciones",
					objetoId: "DFKKOP-1W - Financiaciones",
				}]

				cabeceraSeleccionada = "";

				var numeroCabecerasDFKKOPSemanaPasada = excelFinanciaciones
					.sheet("DFKKOP-1W")
					.usedRange()._numColumns;

				var numeroRegistrosDFKKOPSemanaPasada = excelFinanciaciones
					.sheet("DFKKOP-1W")
					.usedRange()._numRows;

				//Comprobación de inputs:
				if(excelFinanciaciones.sheet("DFKKOP-1W")===undefined){
					return false;
				}

				//Relleno de objeto Ventas No Site:
				for (var i = numFilaCabecera; i < numeroRegistrosDFKKOPSemanaPasada; i++) {
					objetoDatosDFKKOPPasada[0].data.push({});
					for (var j = 0; j < numeroCabecerasDFKKOPSemanaPasada; j++) {

						cabeceraSeleccionada = String(
							excelFinanciaciones
								.sheet("DFKKOP-1W")
								.row(1)
								.cell(j + 1)
								.value()
						);

						cabeceraSeleccionada = cabeceraSeleccionada.toLowerCase();
						cabeceraSeleccionada = cabeceraSeleccionada.replace( / /g, "_");
						cabeceraSeleccionada = cabeceraSeleccionada.replace( /\./g, "_");

						if (cabeceraSeleccionada === undefined || cabeceraSeleccionada === "") {
							console.log(
								"Error de cabecera: i=" + i + " j=" + j
							);
							continue;
						}else{
							//Guardado del registro:
							if(excelFinanciaciones.sheet("DFKKOP-1W").row(i + 1).cell(j+1).value()){
								objetoDatosDFKKOPPasada[0].data[objetoDatosDFKKOPPasada[0].data.length-1][cabeceraSeleccionada]= excelFinanciaciones.sheet("DFKKOP-1W").row(i + 1).cell(j+1).value();
							}

						}
					}

					//Elimina registros vacios:
					if(Object.keys(objetoDatosDFKKOPPasada[0].data[objetoDatosDFKKOPPasada[0].data.length-1]).length === 0){
						objetoDatosDFKKOPPasada[0].data.pop();
					}
				}

				// **************************
				//	Creación del Objeto SEMANA PASADA:
				// **************************
				
				var objetoDatosSemanaPasada = [{
					data: [],
					nombreId: "Semana Pasada - Financiaciones",
					objetoId: "Semana Pasada - Financiaciones",
				}]

				cabeceraSeleccionada = "";

				var numeroCabecerasSemanaPasada = excelFinanciacionesPasada
					.sheet("Semana Actual")
					.usedRange()._numColumns;

				var numeroRegistrosSemanaPasada = excelFinanciacionesPasada
					.sheet("Semana Actual")
					.usedRange()._numRows;

				//Comprobación de inputs:
				if(excelFinanciacionesPasada.sheet("Semana Actual")===undefined){
					return false;
				}

				//Relleno de objeto Ventas No Site:
				for (var i = numFilaCabecera; i < numeroRegistrosSemanaPasada; i++) {
					objetoDatosSemanaPasada[0].data.push({});
					for (var j = 0; j < numeroCabecerasSemanaPasada; j++) {

						cabeceraSeleccionada = String(
							excelFinanciacionesPasada
								.sheet("Semana Actual")
								.row(1)
								.cell(j + 1)
								.value()
						);

						cabeceraSeleccionada = cabeceraSeleccionada.toLowerCase();
						cabeceraSeleccionada = cabeceraSeleccionada.replace( / /g, "_");
						cabeceraSeleccionada = cabeceraSeleccionada.replace( /\./g, "_");

						if (cabeceraSeleccionada === undefined || cabeceraSeleccionada === "") {
							console.log(
								"Error de cabecera: i=" + i + " j=" + j
							);
							continue;
						}else{

							//Guardado del registro:
							if(excelFinanciacionesPasada.sheet("Semana Actual").row(i + 1).cell(j+1).value()){
								objetoDatosSemanaPasada[0].data[objetoDatosSemanaPasada[0].data.length-1][cabeceraSeleccionada]= excelFinanciacionesPasada.sheet("Semana Actual").row(i + 1).cell(j+1).value();
							}

						}
					}

					//Elimina registros vacios:
					if(Object.keys(objetoDatosSemanaPasada[0].data[objetoDatosSemanaPasada[0].data.length-1]).length === 0){
						objetoDatosSemanaPasada[0].data.pop();
					}
				}

		console.log(Object.keys(objetoDatosSemanaPasada[0].data[0]))

				// **************************
				//	Creación del Objeto Mapeo:
				// **************************
				
				var objetoDatosMapeo = [{
					data: [],
					nombreId: "Mapeo - Financiaciones",
					objetoId: "Mapeo - Financiaciones",
				}]

				cabeceraSeleccionada = "";

				var numeroCabecerasMapeo = excelFinanciaciones
					.sheet("Mapeo_id_smart")
					.usedRange()._numColumns;

				var numeroRegistrosMapeo = excelFinanciaciones
					.sheet("Mapeo_id_smart")
					.usedRange()._numRows;

				//Comprobación de inputs:
				if(excelFinanciaciones.sheet("Mapeo_id_smart")===undefined){
					return false;
				}

				//Relleno de objeto Ventas No Site:
				for (var i = numFilaCabecera; i < numeroRegistrosMapeo; i++) {
					objetoDatosMapeo[0].data.push({});
					for (var j = 0; j < numeroCabecerasMapeo; j++) {

						cabeceraSeleccionada = String(
							excelFinanciaciones
								.sheet("Mapeo_id_smart")
								.row(1)
								.cell(j + 1)
								.value()
						);

						cabeceraSeleccionada = cabeceraSeleccionada.toLowerCase();
						cabeceraSeleccionada = cabeceraSeleccionada.replace( / /g, "_");
						cabeceraSeleccionada = cabeceraSeleccionada.replace( /\./g, "_");

						if (cabeceraSeleccionada === undefined || cabeceraSeleccionada === "") {
							console.log(
								"Error de cabecera: i=" + i + " j=" + j
							);
							continue;
						}else{
							//Guardado del registro:
							if(excelFinanciaciones.sheet("Mapeo_id_smart").row(i + 1).cell(j+1).value()){
								objetoDatosMapeo[0].data[objetoDatosMapeo[0].data.length-1][cabeceraSeleccionada]= excelFinanciaciones.sheet("Mapeo_id_smart").row(i + 1).cell(j+1).value();
							}

						}
					}

					//Elimina registros vacios:
					if(Object.keys(objetoDatosMapeo[0].data[objetoDatosMapeo[0].data.length-1]).length === 0){
						objetoDatosMapeo[0].data.pop();
					}
				}


				// **************************
				//	Creación del Objeto Facturacion:
				// **************************
				
				var objetoDatosFacturacion = [{
					data: [],
					nombreId: "Facturacion - Financiaciones",
					objetoId: "Facturacion - Financiaciones",
				}]

				cabeceraSeleccionada = "";

				var numeroCabecerasFacturacion = excelFinanciaciones
					.sheet("Facturación")
					.usedRange()._numColumns;

				var numeroRegistrosFacturacion = excelFinanciaciones
					.sheet("Facturación")
					.usedRange()._numRows;

				//Comprobación de inputs:
				if(excelFinanciaciones.sheet("Facturación")===undefined){
					return false;
				}

				//Relleno de objeto Ventas No Site:
				for (var i = numFilaCabecera; i < numeroRegistrosFacturacion; i++) {
					objetoDatosFacturacion[0].data.push({});
					for (var j = 0; j < numeroCabecerasFacturacion; j++) {

						cabeceraSeleccionada = String(
							excelFinanciaciones
								.sheet("Facturación")
								.row(1)
								.cell(j + 1)
								.value()
						);

						cabeceraSeleccionada = cabeceraSeleccionada.toLowerCase();
						cabeceraSeleccionada = cabeceraSeleccionada.replace( / /g, "_");
						cabeceraSeleccionada = cabeceraSeleccionada.replace( /\./g, "_");

						if (cabeceraSeleccionada === undefined || cabeceraSeleccionada === "") {
							console.log(
								"Error de cabecera: i=" + i + " j=" + j
							);
							continue;
						}else{
							//Guardado del registro:
							if(excelFinanciaciones.sheet("Facturación").row(i + 1).cell(j+1).value()){
								objetoDatosFacturacion[0].data[objetoDatosFacturacion[0].data.length-1][cabeceraSeleccionada]= excelFinanciaciones.sheet("Facturación").row(i + 1).cell(j+1).value();
							}

						}
					}

					//Elimina registros vacios:
					if(Object.keys(objetoDatosFacturacion[0].data[objetoDatosFacturacion[0].data.length-1]).length === 0){
						objetoDatosFacturacion[0].data.pop();
					}
				}

				// **************************
				//	Creación del Objeto REINGENIERIA:
				// **************************
				
				var objetoDatosReingenieria = [{
					data: [],
					nombreId: "Reingeniería - Financiaciones",
					objetoId: "Reingeniería - Financiaciones",
				}]

				cabeceraSeleccionada = "";

				var numeroCabecerasReingenieria = excelFinanciaciones
					.sheet("Reingeniería")
					.usedRange()._numColumns;

				var numeroRegistrosReingenieria = excelFinanciaciones
					.sheet("Reingeniería")
					.usedRange()._numRows;

				//Comprobación de inputs:
				if(excelFinanciaciones.sheet("Reingeniería")===undefined){
					return false;
				}

				//Relleno de objeto Ventas No Site:
				for (var i = numFilaCabecera; i < numeroRegistrosReingenieria; i++) {
					objetoDatosReingenieria[0].data.push({});
					for (var j = 0; j < numeroCabecerasReingenieria; j++) {

						cabeceraSeleccionada = String(
							excelFinanciaciones
								.sheet("Reingeniería")
								.row(1)
								.cell(j + 1)
								.value()
						);

						cabeceraSeleccionada = cabeceraSeleccionada.toLowerCase();
						cabeceraSeleccionada = cabeceraSeleccionada.replace( / /g, "_");
						cabeceraSeleccionada = cabeceraSeleccionada.replace( /\./g, "_");

						if (cabeceraSeleccionada === undefined || cabeceraSeleccionada === "") {
							console.log(
								"Error de cabecera: i=" + i + " j=" + j
							);
							continue;
						}else{
							//Guardado del registro:
							if(excelFinanciaciones.sheet("Reingeniería").row(i + 1).cell(j+1).value()){
								objetoDatosReingenieria[0].data[objetoDatosReingenieria[0].data.length-1][cabeceraSeleccionada]= excelFinanciaciones.sheet("Reingeniería").row(i + 1).cell(j+1).value();
							}

						}
					}

					//Elimina registros vacios:
					if(Object.keys(objetoDatosReingenieria[0].data[objetoDatosReingenieria[0].data.length-1]).length === 0){
						objetoDatosReingenieria[0].data.pop();
					}
				}

				// **************************
				//	Creación del Objeto CAT1:
				// **************************
				
				var objetoDatosCAT1 = [{
					data: [],
					nombreId: "CAT1 - Financiaciones",
					objetoId: "CAT1 - Financiaciones",
				}]

				cabeceraSeleccionada = "";

				var numeroCabecerasCAT1 = excelFinanciaciones
					.sheet("CAT_1")
					.usedRange()._numColumns;

				var numeroRegistrosCAT1 = excelFinanciaciones
					.sheet("CAT_1")
					.usedRange()._numRows;

				//Comprobación de inputs:
				if(excelFinanciaciones.sheet("CAT_1")===undefined){
					return false;
				}

				//Relleno de objeto Ventas No Site:
				for (var i = numFilaCabecera; i < numeroRegistrosCAT1; i++) {
					objetoDatosCAT1[0].data.push({});
					for (var j = 0; j < numeroCabecerasCAT1; j++) {

						cabeceraSeleccionada = String(
							excelFinanciaciones
								.sheet("CAT_1")
								.row(1)
								.cell(j + 1)
								.value()
						);

						cabeceraSeleccionada = cabeceraSeleccionada.toLowerCase();
						cabeceraSeleccionada = cabeceraSeleccionada.replace( / /g, "_");
						cabeceraSeleccionada = cabeceraSeleccionada.replace( /\./g, "_");

						if (cabeceraSeleccionada === undefined || cabeceraSeleccionada === "") {
							console.log(
								"Error de cabecera: i=" + i + " j=" + j
							);
							continue;
						}else{
							//Guardado del registro:
							if(excelFinanciaciones.sheet("CAT_1").row(i + 1).cell(j+1).value()){
								objetoDatosCAT1[0].data[objetoDatosCAT1[0].data.length-1][cabeceraSeleccionada]= excelFinanciaciones.sheet("CAT_1").row(i + 1).cell(j+1).value();
							}

						}
					}

					//Elimina registros vacios:
					if(Object.keys(objetoDatosCAT1[0].data[objetoDatosCAT1[0].data.length-1]).length === 0){
						objetoDatosCAT1[0].data.pop();
					}
				}

				// **************************
				//	Creación del Objeto CAT2:
				// **************************
				
				var objetoDatosCAT2 = [{
					data: [],
					nombreId: "CAT2 - Financiaciones",
					objetoId: "CAT2 - Financiaciones",
				}]

				cabeceraSeleccionada = "";

				var numeroCabecerasCAT2 = excelFinanciaciones
					.sheet("CAT_2")
					.usedRange()._numColumns;

				var numeroRegistrosCAT2 = excelFinanciaciones
					.sheet("CAT_2")
					.usedRange()._numRows;

				//Comprobación de inputs:
				if(excelFinanciaciones.sheet("CAT_2")===undefined){
					return false;
				}

				//Relleno de objeto Ventas No Site:
				for (var i = numFilaCabecera; i < numeroRegistrosCAT2; i++) {
					objetoDatosCAT2[0].data.push({});
					for (var j = 0; j < numeroCabecerasCAT2; j++) {

						cabeceraSeleccionada = String(
							excelFinanciaciones
								.sheet("CAT_2")
								.row(1)
								.cell(j + 1)
								.value()
						);

						cabeceraSeleccionada = cabeceraSeleccionada.toLowerCase();
						cabeceraSeleccionada = cabeceraSeleccionada.replace( / /g, "_");
						cabeceraSeleccionada = cabeceraSeleccionada.replace( /\./g, "_");

						if (cabeceraSeleccionada === undefined || cabeceraSeleccionada === "") {
							console.log(
								"Error de cabecera: i=" + i + " j=" + j
							);
							continue;
						}else{
							//Guardado del registro:
							if(excelFinanciaciones.sheet("CAT_2").row(i + 1).cell(j+1).value()){
								objetoDatosCAT2[0].data[objetoDatosCAT2[0].data.length-1][cabeceraSeleccionada]= excelFinanciaciones.sheet("CAT_2").row(i + 1).cell(j+1).value();
							}

						}
					}

					//Elimina registros vacios:
					if(Object.keys(objetoDatosCAT2[0].data[objetoDatosCAT2[0].data.length-1]).length === 0){
						objetoDatosCAT2[0].data.pop();
					}
				}

				// **************************
				//	Creación del Objeto PENDIENTE:
				// **************************
				
				var objetoDatosPendiente = [{
					data: [],
					nombreId: "PENDIENTE - Financiaciones",
					objetoId: "PENDIENTE - Financiaciones",
				}]

				cabeceraSeleccionada = "";

				var numeroCabecerasPendiente = excelFinanciaciones
					.sheet("PENDIENTE")
					.usedRange()._numColumns;

				var numeroRegistrosPendiente = excelFinanciaciones
					.sheet("PENDIENTE")
					.usedRange()._numRows;

				//Comprobación de inputs:
				if(excelFinanciaciones.sheet("PENDIENTE")===undefined){
					return false;
				}

				//Relleno de objeto Ventas No Site:
				for (var i = numFilaCabecera; i < numeroRegistrosPendiente; i++) {
					objetoDatosPendiente[0].data.push({});
					for (var j = 0; j < numeroCabecerasPendiente; j++) {

						cabeceraSeleccionada = String(
							excelFinanciaciones
								.sheet("PENDIENTE")
								.row(1)
								.cell(j + 1)
								.value()
						);

						cabeceraSeleccionada = cabeceraSeleccionada.toLowerCase();
						cabeceraSeleccionada = cabeceraSeleccionada.replace( / /g, "_");
						cabeceraSeleccionada = cabeceraSeleccionada.replace( /\./g, "_");

						if (cabeceraSeleccionada === undefined || cabeceraSeleccionada === "") {
							console.log(
								"Error de cabecera: i=" + i + " j=" + j
							);
							continue;
						}else{
							//Guardado del registro:
							if(excelFinanciaciones.sheet("PENDIENTE").row(i + 1).cell(j+1).value()){
								objetoDatosPendiente[0].data[objetoDatosPendiente[0].data.length-1][cabeceraSeleccionada]= excelFinanciaciones.sheet("PENDIENTE").row(i + 1).cell(j+1).value();
							}
						}
					}

					//Elimina registros vacios:
					if(Object.keys(objetoDatosPendiente[0].data[objetoDatosPendiente[0].data.length-1]).length === 0){
						objetoDatosPendiente[0].data.pop();
					}
				}

				//Paso 2: Copiar datos en NO SITE VENTAS:
				//Ordenar mensajes ICR:
				/*
				console.log("Ordenando archivo:")
				objetoDatosVentasICG[0].data.sort((a,b) => {
					return ((parseInt(a.fecha)+parseFloat(a.hora) < parseInt(b.fecha)+parseFloat(b.hora)) ? 1 : -1)
				})

				console.log("Ordenando archivo:")
				objetoDatosAbonosICG[0].data.sort((a,b) => {
					return ((parseInt(a.fecha)+parseFloat(a.hora) < parseInt(b.fecha)+parseFloat(b.hora)) ? 1 : -1)
				})
				*/

				// **************************
				//	Guardar Objetos:
				// **************************
				
				mainProcess.guardarDocumento(objetoDatosDFKKOP[0])
				mainProcess.guardarDocumento(objetoDatosMapeo[0])
				mainProcess.guardarDocumento(objetoDatosFacturacion[0])
				mainProcess.guardarDocumento(objetoDatosReingenieria[0])
				mainProcess.guardarDocumento(objetoDatosCAT1[0])
				mainProcess.guardarDocumento(objetoDatosCAT2[0])
				mainProcess.guardarDocumento(objetoDatosPendiente[0])

				
				//RELLENAR SEMANA ACTUAL:
				var numeroRegistrosSemanaActual = excelFinanciaciones
					.sheet("Semana Actual")
					.usedRange()._numRows;

				var numeroCabecerasSemanaActual = excelFinanciaciones
					.sheet("Semana Actual")
					.usedRange()._numColumns;
				
				//Limpiar registros:
				for (var i = 2; i <= numeroRegistrosSemanaActual; i++) {
					for (var j = 1; j <= numeroCabecerasSemanaActual; j++) {
						excelFinanciaciones.sheet("Semana Actual").cell(i,j).value("")
					}
				}

				// *********************
				//	COLUMNAS ABC:
				// *********************
				
				
				console.log("Procesando columnas ABC...")

				//Ordenar DFKKOP:
				objetoDatosDFKKOP[0].data.sort((a,b) => {
					return ((parseInt(a["num_doc"]) < parseInt(b["num_doc"])) ? 1 : -1)
				})

				var registrosOPBEL = []

				for (var i = 0; i < objetoDatosDFKKOP[0].data.length; i++) {
					if(objetoDatosDFKKOP[0].data[i]["num_doc"]!="" && objetoDatosDFKKOP[0].data[i]["num_doc"]!=undefined){
						registrosOPBEL.push(objetoDatosDFKKOP[0].data[i]["num_doc"])	
					}
				}

				var registrosOPBELunicos = [...new Set(registrosOPBEL)]
				var registroCuentaPosicion  = 0 
				var registroSumaImporte = 0 

				//Ordenar Registros OPBEL:
				registrosOPBEL.sort((a,b) => {
					return ((parseInt(a) < parseInt(b)) ? 1 : -1)
				})

				var ultimoRegistroAnalizado = 0;
				for (var i = 2; i < registrosOPBELunicos.length+2; i++) {

					excelFinanciaciones.sheet("Semana Actual").cell(i,1).value(registrosOPBELunicos[i-2])

					//Contar Posiciones e importes:
					registroCuentaPosicion = 0
					registroSumaImporte = 0

					if(ultimoRegistroAnalizado>0){
						ultimoRegistroAnalizado--;
					}

					for(var j= 0; j < objetoDatosDFKKOP[0].data.length; j++){

						if((ultimoRegistroAnalizado)>=objetoDatosDFKKOP[0].data.length-1){
							ultimoRegistroAnalizado = 0;
						}
						ultimoRegistroAnalizado++;


						if((j!=0) && (objetoDatosDFKKOP[0].data[ultimoRegistroAnalizado]["num_doc"]!=objetoDatosDFKKOP[0].data[ultimoRegistroAnalizado-1]["num_doc"])&&(objetoDatosDFKKOP[0].data[ultimoRegistroAnalizado-1]["num_doc"]==registrosOPBELunicos[i-2])){
							break;}

						if(parseInt(registrosOPBELunicos[i-2]) == parseInt(objetoDatosDFKKOP[0].data[ultimoRegistroAnalizado]["num_doc"])){
							registroCuentaPosicion++;
						}

						if(parseInt(registrosOPBELunicos[i-2]) == parseInt(objetoDatosDFKKOP[0].data[ultimoRegistroAnalizado]["num_doc"])){
							registroSumaImporte = registroSumaImporte+parseFloat(objetoDatosDFKKOP[0].data[ultimoRegistroAnalizado]["importeml"]);  		
						}

					}

					excelFinanciaciones.sheet("Semana Actual").cell(i,2).value(registroSumaImporte)
					excelFinanciaciones.sheet("Semana Actual").cell(i,3).value(registroCuentaPosicion)
				}

				// *******************
				//	COLUMNAS DEF:
				// *******************
				
				var telefono = 0;
				var factura = "";
				var fechaContabilizacion = 0;

				console.log("Procesando columnas DEF...")

				//Ordenar CAT1:
				objetoDatosCAT1[0].data.sort((a,b) => {
					return ((parseInt(b["número_de_telefono"]) < parseInt(a["número_de_telefono"])) ? 1 : -1)
				})

				ultimoRegistroAnalizado = 0;
				for (var i = 2; i < registrosOPBELunicos.length+2; i++) {

					telefono = 0; 
					factura = 0;

					if(ultimoRegistroAnalizado>0){
						ultimoRegistroAnalizado--;
					}

					for(var j= 0; j < objetoDatosCAT1[0].data.length; j++){

						if((ultimoRegistroAnalizado)>=objetoDatosCAT1[0].data.length-1){
							ultimoRegistroAnalizado = 0;
						}
						ultimoRegistroAnalizado++;

						if(parseInt(registrosOPBELunicos[i-2]) == parseInt(objetoDatosCAT1[0].data[ultimoRegistroAnalizado]["documento_suplente"])){
							telefono = objetoDatosCAT1[0].data[ultimoRegistroAnalizado]["número_de_teléfono"];
							factura = objetoDatosCAT1[0].data[ultimoRegistroAnalizado]["criterio_de_clasificación"];
							fechaContabilizacion = objetoDatosCAT1[0].data[ultimoRegistroAnalizado]["fe_contabilización"];
							break;
						}
					}

					excelFinanciaciones.sheet("Semana Actual").cell(i,4).value(telefono)
					excelFinanciaciones.sheet("Semana Actual").cell(i,5).value(factura)
					excelFinanciaciones.sheet("Semana Actual").cell(i,6).value(fechaContabilizacion)
				}

				// *******************
				//	COLUMNAS GHI:
				// *******************

				var cuentaContrato = ""
				var flagBreak = false;

				console.log("Procesando columnas GHI...")

				ultimoRegistroAnalizado = 0;
				for (var i = 2; i < registrosOPBELunicos.length+2; i++) {

					if(ultimoRegistroAnalizado>0){
						ultimoRegistroAnalizado--;
					}

					for(var j= 0; j < objetoDatosDFKKOP[0].data.length; j++){

						if((ultimoRegistroAnalizado)>=objetoDatosDFKKOP[0].data.length-1){
							ultimoRegistroAnalizado = 0;
						}
						ultimoRegistroAnalizado++;

						if(parseInt(registrosOPBELunicos[i-2]) == parseInt(objetoDatosDFKKOP[0].data[ultimoRegistroAnalizado]["num_doc"])){
							cuentaContrato = objetoDatosDFKKOP[0].data[ultimoRegistroAnalizado]["cta_contr"];
							excelFinanciaciones.sheet("Semana Actual").cell(i,8).value(objetoDatosDFKKOP[0].data[ultimoRegistroAnalizado]["cta_contr"])
							for(var k=0; k < objetoDatosMapeo[0].data.length; k++){
								if(objetoDatosMapeo[0].data[k]["vkont"]==cuentaContrato){
									excelFinanciaciones.sheet("Semana Actual").cell(i,7).value(String(objetoDatosMapeo[0].data[k]["vkona"]))
									excelFinanciaciones.sheet("Semana Actual").cell(i,9).value(String(objetoDatosMapeo[0].data[k]["site"]))
									flagBreak=true;
									break;
								}
							}
						break;
						}
					}
					if(flagBreak){flagBreak=false;continue;}
				}

				// *******************
				//	COLUMNAS JK:
				// *******************

				console.log("Procesando columnas JK...")

				ultimoRegistroAnalizado = 0;
				for (var i = 2; i < registrosOPBELunicos.length+2; i++) {

					if(ultimoRegistroAnalizado>0){
						ultimoRegistroAnalizado--;
					}

					for(var j= 0; j < objetoDatosFacturacion[0].data.length; j++){

						if((ultimoRegistroAnalizado)>=objetoDatosFacturacion[0].data.length-1){
							ultimoRegistroAnalizado = 0;
						}
						ultimoRegistroAnalizado++;

						if(String(excelFinanciaciones.sheet("Semana Actual").cell(i,9).value()) == String(objetoDatosFacturacion[0].data[ultimoRegistroAnalizado]["cta_contr_sist_exis"])){

							excelFinanciaciones.sheet("Semana Actual").cell(i,10).value(objetoDatosFacturacion[0].data[ultimoRegistroAnalizado]["fecha"])
							excelFinanciaciones.sheet("Semana Actual").cell(i,11).value("KPI1");
							flagBreak = true;
						break;
						}
					}
					if(flagBreak){flagBreak=false;continue;}
					excelFinanciaciones.sheet("Semana Actual").cell(i,10).value("")
					excelFinanciaciones.sheet("Semana Actual").cell(i,11).value("KPI2");
				}

				// *******************
				//	COLUMNAS LM:
				// *******************

				console.log("Procesando columnas LM...")

				ultimoRegistroAnalizado = 0;
				for (var i = 2; i < registrosOPBELunicos.length+2; i++) {

					if(ultimoRegistroAnalizado>0){
						ultimoRegistroAnalizado--;
					}

					for(var j= 0; j < objetoDatosFacturacion[0].data.length; j++){

						if((ultimoRegistroAnalizado)>=objetoDatosFacturacion[0].data.length-1){
							ultimoRegistroAnalizado = 0;
						}
						ultimoRegistroAnalizado++;

						if(parseInt(registrosOPBELunicos[i-2]) == parseInt(objetoDatosPendiente[0].data[ultimoRegistroAnalizado]["plpagoplz"])){
							excelFinanciaciones.sheet("Semana Actual").cell(i,13).value(objetoDatosPendiente[0].data[ultimoRegistroAnalizado]["acción"])
							if(objetoDatosPendiente[0].data[ultimoRegistroAnalizado]["acción"]=="MEMO ANTERIOR 2017"){
								excelFinanciaciones.sheet("Semana Actual").cell(i,12).value("Anterior a 2017");
							}else{
								excelFinanciaciones.sheet("Semana Actual").cell(i,12).value("De 2017 y posterior");
							}
							flagBreak = true;
						break;
						}
					}
					if(flagBreak){flagBreak=false;continue;}
					excelFinanciaciones.sheet("Semana Actual").cell(i,13).value("NO")
					excelFinanciaciones.sheet("Semana Actual").cell(i,12).value("De 2017 y posterior");
				}

				// *******************
				//	COLUMNAS NO:
				// *******************

				console.log("Procesando columnas NO...")

				var numDocCon8yOPORD = []; 
				var critClasCon8yOPORD = [];

				//Relleno de array con Restriccion 8 y OPORD:
				for(var j= 0; j < objetoDatosCAT2[0].data.length; j++){
						if(objetoDatosCAT2[0].data[j]["restricción_compens"]==8){
							if(objetoDatosCAT2[0].data[j]["criterio_de_clasificación"]){
								numDocCon8yOPORD.push(objetoDatosCAT2[0].data[j]["número_de_documento"])
								//critClasCon8yOPORD.push(objetoDatosCAT2[0].data[j]["criterio_de_clasificación"])
							}
						}
				}

				//CAMPO N:
				ultimoRegistroAnalizado = 0;
				for (var i = 2; i < registrosOPBELunicos.length+2; i++) {

					if(ultimoRegistroAnalizado>0){
						ultimoRegistroAnalizado--;
					}

					for(var j= 0; j < numDocCon8yOPORD.length; j++){

						if((ultimoRegistroAnalizado)>=numDocCon8yOPORD.length-1){
							ultimoRegistroAnalizado = 0;
						}
						ultimoRegistroAnalizado++;

						if(parseInt(registrosOPBELunicos[i-2]) == parseInt(numDocCon8yOPORD[ultimoRegistroAnalizado])){
							excelFinanciaciones.sheet("Semana Actual").cell(i,14).value("SI");
							flagBreak = true;
							break;
						}
					}

					if(flagBreak){flagBreak=false;continue;}
					excelFinanciaciones.sheet("Semana Actual").cell(i,14).value("NO");
				}

				//CAMPO O:
				
				//Ordenar por número de documento y fecha:
				objetoDatosReingenieria[0].data.sort((a,b) => {
					return ((parseInt(a["nº_plpagoplazos"])+parseFloat(a["día_envío"]) < parseInt(b["nº_plpagoplazos"])+parseFloat(b["día_envío"])) ? 1 : -1)
				})

				ultimoRegistroAnalizado = 0;
				for (var i = 2; i < registrosOPBELunicos.length+2; i++) {

					if(ultimoRegistroAnalizado>0){
						ultimoRegistroAnalizado--;
					}

					for(var j= 0; j < objetoDatosReingenieria[0].data.length; j++){

						if((ultimoRegistroAnalizado)>=objetoDatosReingenieria[0].data.length-1){
							ultimoRegistroAnalizado = 0;
						}
						ultimoRegistroAnalizado++;

						if(parseInt(registrosOPBELunicos[i-2]) == parseInt(objetoDatosReingenieria[0].data[ultimoRegistroAnalizado]["nº_plpagoplazos"])){
							if(objetoDatosReingenieria[0].data[ultimoRegistroAnalizado]["estado_de_cuota_en_financiación"]=="00"){
								excelFinanciaciones.sheet("Semana Actual").cell(i,15).value("Liberado");
							}else{
								excelFinanciaciones.sheet("Semana Actual").cell(i,15).value(objetoDatosReingenieria[0].data[ultimoRegistroAnalizado]["día_envío"]);
							}
							
							flagBreak = true;
							break;
						}
					}
					if(flagBreak){flagBreak=false;continue;}
					excelFinanciaciones.sheet("Semana Actual").cell(i,15).value("Fuera de reingeniería");
				}

				// *******************
				//	COLUMNAS R:
				// *******************
				
				console.log("Procesando columnas R...")

				var numDocRegu = [];

				for(var j= 0; j < objetoDatosCAT2[0].data.length; j++){
					if(parseInt(objetoDatosCAT2[0].data[j]["fecha_de_venta_a_vf_overseas"])>=43831){
						numDocRegu.push(objetoDatosCAT2[0].data[j]["número_de_documento"])
					}					
				}

				ultimoRegistroAnalizado = 0;
				for (var i = 2; i < registrosOPBELunicos.length+2; i++) {

					if(ultimoRegistroAnalizado>0){
						ultimoRegistroAnalizado--;
					}

					for(var j= 0; j < numDocRegu.length; j++){

						if((ultimoRegistroAnalizado)>=numDocRegu.length-1){
							ultimoRegistroAnalizado = 0;
						}
						ultimoRegistroAnalizado++;

						if(numDocRegu[ultimoRegistroAnalizado]==registrosOPBELunicos[i-2]){
							excelFinanciaciones.sheet("Semana Actual").cell(i,18).value("VENDIDO");
							flagBreak = true;
							break;
						}
					}
					if(flagBreak){flagBreak=false;continue;}
					excelFinanciaciones.sheet("Semana Actual").cell(i,18).value("MODIFICABLE");
				}

				// *******************
				//	COLUMNAS OLD:
				// *******************
				
				console.log("Procesando columnas OLD...")

				ultimoRegistroAnalizado = 0;
				for (var i = 2; i < registrosOPBELunicos.length+2; i++) {
					if(ultimoRegistroAnalizado>0){
						ultimoRegistroAnalizado--;
					}
						for(var j= 0; j < objetoDatosSemanaPasada[0].data.length; j++){

							if((ultimoRegistroAnalizado)>=objetoDatosSemanaPasada[0].data.length-1){
								ultimoRegistroAnalizado = 0;
							}
							ultimoRegistroAnalizado++;

							if(parseInt(objetoDatosSemanaPasada[0].data[ultimoRegistroAnalizado]["financiación"])==parseInt(registrosOPBELunicos[i-2])){
								excelFinanciaciones.sheet("Semana Actual").cell(i,19).value(objetoDatosSemanaPasada[0].data[ultimoRegistroAnalizado]["kpi"]);
								excelFinanciaciones.sheet("Semana Actual").cell(i,20).value(objetoDatosSemanaPasada[0].data[ultimoRegistroAnalizado]["creación"]);
								excelFinanciaciones.sheet("Semana Actual").cell(i,21).value(objetoDatosSemanaPasada[0].data[ultimoRegistroAnalizado]["estado"]);
								excelFinanciaciones.sheet("Semana Actual").cell(i,22).value(objetoDatosSemanaPasada[0].data[ultimoRegistroAnalizado]["subestado"]);
								flagBreak = true;
								break;
							}
						}

					if(flagBreak){flagBreak=false;continue;}
					excelFinanciaciones.sheet("Semana Actual").cell(i,21).value("Nuevo");
					excelFinanciaciones.sheet("Semana Actual").cell(i,22).value("Nuevo");
				}

				//CLASIFICACION DE CATEGORIA Y SUBCATEGORIA:
				
				console.log("Realizando clasificación...")
				var flagDespintada = false;

				for (var i = 2; i < registrosOPBELunicos.length+2; i++) {

					//Paso 1: Se conservan devoluciones y traslado antiguo:
					if(String(excelFinanciaciones.sheet("Semana Actual").cell(i,22).value()).includes("Traslado - Pdt. compensar origen")){
								excelFinanciaciones.sheet("Semana Actual").cell(i,16).value("Fuera de reingeniería");
								excelFinanciaciones.sheet("Semana Actual").cell(i,17).value(excelFinanciaciones.sheet("Semana Actual").cell(i,22).value());
					}

					if(String(excelFinanciaciones.sheet("Semana Actual").cell(i,22).value()).includes("Pdt. de aplicar devolución")){
								excelFinanciaciones.sheet("Semana Actual").cell(i,16).value("Fuera de reingeniería");
								excelFinanciaciones.sheet("Semana Actual").cell(i,17).value(excelFinanciaciones.sheet("Semana Actual").cell(i,22).value());
					}

					//Paso 2: Se establecen los nuevos traslados:
					if(String(excelFinanciaciones.sheet("Semana Actual").cell(i,5).value()).includes("TRAS-")){
								excelFinanciaciones.sheet("Semana Actual").cell(i,16).value("Fuera de reingeniería");
								excelFinanciaciones.sheet("Semana Actual").cell(i,17).value("Traslado - Pdt. compensar origen (subcasos)");
					}

					//Paso 3: Estado Cuarentena:
					if(String(excelFinanciaciones.sheet("Semana Actual").cell(i,16).value())==""){
						if(String(excelFinanciaciones.sheet("Semana Actual").cell(i,13).value())!="NO" && String(excelFinanciaciones.sheet("Semana Actual").cell(i,13).value())!=""){
							excelFinanciaciones.sheet("Semana Actual").cell(i,16).value("En cuarentena");
							excelFinanciaciones.sheet("Semana Actual").cell(i,17).value(String(excelFinanciaciones.sheet("Semana Actual").cell(i,13).value()));
						}
					}

					//Paso 4: En reingeniería:
					if(String(excelFinanciaciones.sheet("Semana Actual").cell(i,16).value())==""){
					if(excelFinanciaciones.sheet("Semana Actual").cell(i,15).value()!= "Fuera de reingeniería" && excelFinanciaciones.sheet("Semana Actual").cell(i,15).value()!= "Liberado" && excelFinanciaciones.sheet("Semana Actual").cell(i,15).value()!= ""){
						excelFinanciaciones.sheet("Semana Actual").cell(i,16).value("En reingeniería");
						excelFinanciaciones.sheet("Semana Actual").cell(i,17).value(String(excelFinanciaciones.sheet("Semana Actual").cell(i,19).value()));
					}
					}

					//Paso 5: Devolución o Pdt. Análisis
					if(String(excelFinanciaciones.sheet("Semana Actual").cell(i,16).value())==""){
						excelFinanciaciones.sheet("Semana Actual").cell(i,16).value("Fuera de reingeniería");
						excelFinanciaciones.sheet("Semana Actual").cell(i,17).value("Pdt. Análisis");
					}

					function dateToExcel(date){
						var days = Math.round((date - new Date(1899,11,30))/8.64e7);
						return parseInt((days).toFixed(10));
					}

					//Paso 6: KPI1 
					if(String(excelFinanciaciones.sheet("Semana Actual").cell(i,16).value())=="En reingeniería"){
					if(excelFinanciaciones.sheet("Semana Actual").cell(i,15).value()!= "Fuera de reingeniería" && excelFinanciaciones.sheet("Semana Actual").cell(i,15).value()!= "Liberado" && excelFinanciaciones.sheet("Semana Actual").cell(i,15).value()!= ""){
					
						//ENVIO EN MENOS DE UN MES:
					if(String(excelFinanciaciones.sheet("Semana Actual").cell(i,11).value())=="KPI1"){

						if(excelFinanciaciones.sheet("Semana Actual").cell(i,15).value()>=(dateToExcel(Date.now())-30)){

							excelFinanciaciones.sheet("Semana Actual").cell(i,16).value("En reingeniería");
							excelFinanciaciones.sheet("Semana Actual").cell(i,17).value("Informativas");

					}else{
						//ENVIO HACE MÁS DE UN MES:
						//ANALIZAR DESPINTADA:
						flagDespintada = false;
						for(var j = 0; j<objetoDatosReingenieria[0].data.length; j++){

							if(registrosOPBELunicos[i-2]==objetoDatosReingenieria[0].data[j]["nº_plpagoplazos"] && objetoDatosReingenieria[0].data[j]["despintado"] == "X"){

								excelFinanciaciones.sheet("Semana Actual").cell(i,16).value("Fuera de reingeniería");
								excelFinanciaciones.sheet("Semana Actual").cell(i,17).value("Cuotas despintadas");
								flagDespintada = true;
							}
						}
						if(!flagDespintada){
							excelFinanciaciones.sheet("Semana Actual").cell(i,16).value("En reingeniería");
							excelFinanciaciones.sheet("Semana Actual").cell(i,17).value("Pdt. Análisis");
						}
					}
					}
					}
					}

					//Paso 7: KPI2 
					if(String(excelFinanciaciones.sheet("Semana Actual").cell(i,16).value())=="En reingeniería"){
					
					if(String(excelFinanciaciones.sheet("Semana Actual").cell(i,11).value())=="KPI2"){

						excelFinanciaciones.sheet("Semana Actual").cell(i,16).value("En reingeniería");
						excelFinanciaciones.sheet("Semana Actual").cell(i,17).value("Pdt. Análisis");
					}
					}
					
					//Paso 8: Envio GNV Liberadas Recientemente 
					if(excelFinanciaciones.sheet("Semana Actual").cell(i,15).value()== "Liberado"){
					
						for(var j = 0; j<objetoDatosReingenieria[0].data.length; j++){

							if((registrosOPBELunicos[i-2]==objetoDatosReingenieria[0].data[j]["nº_plpagoplazos"]) && ((dateToExcel(Date.now())-objetoDatosReingenieria[0].data[j]["día_creación"]) <= 3) && (objetoDatosReingenieria[0].data[j]["día_creación"]!= "")){

								excelFinanciaciones.sheet("Semana Actual").cell(i,16).value("En reingeniería");
								excelFinanciaciones.sheet("Semana Actual").cell(i,17).value("Informativas");
							}
						}

					}
				}

					//**********************************
					//	CREACIÓN HOJA SEMANA PASADA:
					//**********************************
					//
					//LIMPIAR HOJA SEMANA PASADA:
					
					var numeroRegistrosSemanaPasada = excelFinanciaciones
						.sheet("Semana Pasada")
						.usedRange()._numRows;

					var numeroCabecerasSemanaPasada = excelFinanciaciones
						.sheet("Semana Pasada")
						.usedRange()._numColumns;
					
					console.log("NUMERO REGISTROS: "+ numeroRegistrosSemanaPasada)
					console.log("NUMERO COLUMNAS: "+ numeroCabecerasSemanaPasada)
					
					//Limpiar registros:
					for (var k = 2; k <= numeroRegistrosSemanaPasada+2; k++) {
						for (var l = 1; l <= numeroCabecerasSemanaPasada+1; l++) {
							excelFinanciaciones.sheet("Semana Pasada").cell(k,l).value("")
						}
					}

					
					//Copiar Valores de Semana Actual en Semana Pasada:
					numeroRegistrosSemanaActual = excelFinanciaciones
						.sheet("Semana Actual")
						.usedRange()._numRows;

					numeroCabecerasSemanaActual = excelFinanciaciones
						.sheet("Semana Actual")
						.usedRange()._numColumns;

					for (var i = 2; i <= numeroRegistrosSemanaActual; i++) {
						for (var j = 1; j <= numeroCabecerasSemanaActual; j++) {
							excelFinanciaciones.sheet("Semana Pasada").cell(i,j).value(excelFinanciaciones.sheet("Semana Actual").cell(i,j).value())
						}
					}

					
					//ITERAR POR NUMERO FINANCIACION W-1:
					var cuentaNumDoc = 0;
					var sumaImporte = 0;

					for (var i = 2; i < registrosOPBELunicos.length+2; i++) {
						cuentaNumDoc = 0;
						sumaImporte = 0; 

						for(var j = 0; j< objetoDatosDFKKOPPasada[0].data.length; j++){
							if(parseInt(registrosOPBELunicos[i-2])==parseInt(objetoDatosDFKKOPPasada[0].data[j]["num_doc"])){
								if(objetoDatosDFKKOPPasada[0].data[j]["importeml"]){
									sumaImporte = sumaImporte + parseFloat(objetoDatosDFKKOPPasada[0].data[j]["importeml"])
								}
								cuentaNumDoc++;
							}
						}
						excelFinanciaciones.sheet("Semana Pasada").cell(i,2).value(sumaImporte)
						excelFinanciaciones.sheet("Semana Pasada").cell(i,3).value(cuentaNumDoc)
					}

					function deleteRow(sheet, rowNumber, count){
						sheet._rows.splice(rowNumber,count);
						sheet._rows.map((row, index) => {
							sheet._node.attributes.r = index;
						});
					}

					//Eliminar Valores cuenta 0 en Semana pasada:
					numeroRegistrosSemanaPasada = excelFinanciaciones
						.sheet("Semana Pasada")
						.usedRange()._numRows;

					numeroCabecerasSemanaPasada = excelFinanciaciones
						.sheet("Semana Pasada")
						.usedRange()._numColumns;
					
					//Limpiar registros:
					var cuentaEliminaciones = 0;
					for (var i = 2; i <= numeroRegistrosSemanaPasada; i++) {
							if(parseInt(excelFinanciaciones.sheet("Semana Pasada").cell(i,3).value())===0){
								deleteRow(excelFinanciaciones.sheet("Semana Pasada"),i-1,1);
								cuentaEliminaciones++;
								i--;
							}
					}
				console.log("Cuenta Eliminaciones: "+cuentaEliminaciones);


				//Modificación archivo de seguimiento KPIs:
				return excelFinanciaciones
					.toFileAsync(path.normalize(pathExcelFinanciaciones))
					.then(() => {
						console.log("Fin del procesamiento excel Financiaciones");
						return true
					})
					.catch(err => {
						console.log("Se ha producido un error interno: ");
						console.log(err);
						var tituloError =
							"Se ha producido un error escribiendo el archivo: " +
							path.normalize(path.normalize(pathExcelFinanciaciones));
						return false;
					});

			})
			.then(()=>{

				//Modificación de Archivo seguimiento:	
				return XlsxPopulate.fromFileAsync(path.normalize(pathExcelSeguimiento))
					.then(excelSeguimiento => {
						//Tratamiento excel seguimiento:
						
						console.log("Cargando Excel Seguimiento:");

						/* Escribir en seguimiento:
						 *
						if(excelSeguimiento===undefined){

							return false;
						}

						//Creación del Objeto:
						var objetoDatosSeguimiento = [{
							data: [],
							nombreId: "Seguimiento KPIs",
							objetoId: "Seguimiento KPIs",
						}]

						var cabeceraSeleccionada = "";

						var numeroCabecerasSeguimiento = excelSeguimiento
							.sheet("Datos KPI 367")
							.usedRange()._numColumns;

						var numeroRegistrosSeguimiento = excelSeguimiento
							.sheet("Datos KPI 367")
							.usedRange()._numRows;

						//Comprobación de inputs:
						if(excelSeguimiento.sheet("Datos KPI 367")===undefined){
							return false;
						}

						var numeroSemana = moment().isoWeek();
						var numeroYear = moment().year();

						if(numeroSemana<10){
							numeroSemana = "W0"+numeroSemana;
						}else{
							numeroSemana = "W"+numeroSemana;
						}

						console.log("Numero de semana actual: "+ numeroSemana);
						console.log("Numero de año actual: "+ numeroYear);

						//Encontrar ultima fila de seguimiento
						var filaSemanaActual= 1	
						for(var i=2; i<=numeroRegistrosSeguimiento+1; i++){
							if(excelSeguimiento.sheet("Datos KPI 367").cell(i,1).value()==numeroSemana && excelSeguimiento.sheet("Datos KPI 367").cell(i,2).value()== numeroYear){
								filaSemanaActual=i;
								break;
							}else 
								if(excelSeguimiento.sheet("Datos KPI 367").cell(i,1).value()=="" || excelSeguimiento.sheet("Datos KPI 367").cell(i,1).value()==undefined){
									filaSemanaActual=i;
									break;
								}
						}

						console.log("Fila Actual: "+ filaSemanaActual)

						//Escribir valores en tabla KPIS:
						if(excelSeguimiento.sheet("Datos KPI 367").cell(filaSemanaActual,1).value()=="" || excelSeguimiento.sheet("Datos KPI 367").cell(filaSemanaActual,1).value()==undefined){
							excelSeguimiento.sheet("Datos KPI 367").cell(filaSemanaActual,1).value(numeroSemana) 
							excelSeguimiento.sheet("Datos KPI 367").cell(filaSemanaActual,2).value(numeroYear) 
						}

						excelSeguimiento.sheet("Datos KPI 367").cell(filaSemanaActual,4).value(devolucionesNoCargada) 
						excelSeguimiento.sheet("Datos KPI 367").cell(filaSemanaActual,5).value(ventasNoCargada) 

						//Clasificacion Ventas:
						var numeroRegistrosSeguimientoVentas = excelSeguimiento
							.sheet("VENTAS")
							.usedRange()._numRows;

						var filaSemanaActualVentas= 1	

						for(var i=2; i<=numeroRegistrosSeguimientoVentas+1; i++){
							if(excelSeguimiento.sheet("VENTAS").cell(i,1).value()==numeroSemana && excelSeguimiento.sheet("VENTAS").cell(i,2).value()== numeroYear){
								filaSemanaActualVentas=i;
								break;
							}else 
								if(excelSeguimiento.sheet("VENTAS").cell(i,1).value()=="" || excelSeguimiento.sheet("VENTAS").cell(i,1).value()==undefined){
									filaSemanaActualVentas=i;
									break;
								}
						}

						//Escribir valores FACTURACION:
						if(excelSeguimiento.sheet("VENTAS").cell(filaSemanaActualVentas,1).value()=="" || excelSeguimiento.sheet("VENTAS").cell(filaSemanaActualVentas,1).value()==undefined){
							excelSeguimiento.sheet("VENTAS").cell(filaSemanaActualVentas,1).value(numeroSemana) 
							excelSeguimiento.sheet("VENTAS").cell(filaSemanaActualVentas,2).value(numeroYear) 
						}
						for(var i=0; i<clasificacionVentas.length;i++){
							excelSeguimiento.sheet("VENTAS").cell(filaSemanaActualVentas,i+4).value(clasificacionVentas[i]) 
						}

						//Clasificacion Abonos:
						var numeroRegistrosSeguimientoDevoluciones = excelSeguimiento
							.sheet("ABONOS Y DEVOLUCIONES")
							.usedRange()._numRows;

						var filaSemanaActualDevoluciones= 1	

						for(var i=2; i<=numeroRegistrosSeguimientoDevoluciones+1; i++){
							if(excelSeguimiento.sheet("ABONOS Y DEVOLUCIONES").cell(i,1).value()==numeroSemana && excelSeguimiento.sheet("ABONOS Y DEVOLUCIONES").cell(i,2).value()== numeroYear){
								filaSemanaActualDevoluciones=i;
								break;
							}else 
								if(excelSeguimiento.sheet("ABONOS Y DEVOLUCIONES").cell(i,1).value()=="" || excelSeguimiento.sheet("ABONOS Y DEVOLUCIONES").cell(i,1).value()==undefined){
									filaSemanaActualDevoluciones=i;
									break;
								}
						}

						//Escribir valores FACTURACION:
						if(excelSeguimiento.sheet("ABONOS Y DEVOLUCIONES").cell(filaSemanaActualDevoluciones,1).value()=="" || excelSeguimiento.sheet("ABONOS Y DEVOLUCIONES").cell(filaSemanaActualDevoluciones,1).value()==undefined){
							excelSeguimiento.sheet("ABONOS Y DEVOLUCIONES").cell(filaSemanaActualDevoluciones,1).value(numeroSemana) 
							excelSeguimiento.sheet("ABONOS Y DEVOLUCIONES").cell(filaSemanaActualDevoluciones,2).value(numeroYear) 
						}

						for(var i=0; i<clasificacionDevoluciones.length;i++){
							excelSeguimiento.sheet("ABONOS Y DEVOLUCIONES").cell(filaSemanaActualDevoluciones,i+4).value(clasificacionDevoluciones[i]) 
						}
						
						*/
						return excelSeguimiento
							.toFileAsync(path.normalize(pathExcelSeguimiento))
							.then(() => {
								console.log("Fin del procesamiento excel Seguimiento");
								return true
							})
							.catch(err => {

								console.log(err);
								var tituloError =
									"Se ha producido un error escribiendo el archivo: " +
											path.normalize(path.normalize(pathExcelSeguimiento));
										return false;
									});
					})
				})

	}
} 

module.exports = ProcesosKPIs;


