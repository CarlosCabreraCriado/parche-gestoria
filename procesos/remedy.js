const path = require("path");
const fs = require("fs");
const readline = require("readline");
const moment = require("moment");
const XlsxPopulate = require("xlsx-populate");
const Datastore = require("nedb");
const _ = require("lodash");
const { ipcRenderer } = require("electron");
const puppeteer = require("puppeteer");

class ProcesosRemedy {
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

  async extraccionRemedy(argumentos) {
    console.log("Extracción Remedy");
    //console.log("Archivo entrada: "+argumentos[0])
    //console.log("Archivo salida: "+argumentos[1])

    const browser = await puppeteer.launch({
      headless: false,
      executablePath: path.join(argumentos[0]),
      args: [
        "--disable-web-security",
        "--disable-features=IsolateOrigins,site-per-process",
      ],
    });

    const page = await browser.newPage();

    //(async () => {
    //
    //await page.goto('https://oneitsm.onbmc.com/arsys');
    await page.goto(
      "https://oneitsm.onbmc.com/arsys/forms/onbmc-s/AR+System+Report+Console/Default+Administrator+View/",
    );

    page.on("dialog", async (dialog) => {
      await dialog.accept();
    });

    //await page.waitForNavigation();
    const tiempoEspera = 2000;
    await page._client.send("Page.setDownloadBehavior", {
      behavior: "allow",
      downloadPath:
        "/Users/carloscabreracriado/OneDrive - Vodafone Group/Procesamiento",
    });

    /*  NAVEGACIÓN DESDE PAGINA PRINCIPAL
			 *
			 *
			await page.waitForTimeout(tiempoEspera);
			await page.waitForSelector("#WIN_0_304316340")
			await page.click("#WIN_0_304316340")

			await page.waitForTimeout(tiempoEspera);
			await page.waitForSelector("#FormContainer > div.FlyoutContainer.Applist.arfid1575.ardbnApplicationListField > div > div:nth-child(9)")
			await page.hover("#FormContainer > div.FlyoutContainer.Applist.arfid1575.ardbnApplicationListField > div > div:nth-child(9)")

			await page.waitForTimeout(tiempoEspera);
			await page.waitForSelector("#FormContainer > div.FlyoutContainer.Applist.arfid1575.ardbnApplicationListField > div > div.root.root_menu.VNavHover0 > div > div:nth-child(1) > a")
			await page.click("#FormContainer > div.FlyoutContainer.Applist.arfid1575.ardbnApplicationListField > div > div.root.root_menu.VNavHover0 > div > div:nth-child(1) > a")

			//Página Incident Manager:
			await page.waitForTimeout(tiempoEspera);
			await page.waitForSelector("#sub-301650500 > div.VNavLeaf.VNavLevel2.arfid301381600.ardbnz2NI_Reports")
			await page.hover("#sub-301650500 > div.VNavLeaf.VNavLevel2.arfid301381600.ardbnz2NI_Reports")

			await page.waitForTimeout(tiempoEspera);
			await page.waitForSelector("#sub-301650500 > div.VNavLeaf.VNavLevel2.arfid301381600.ardbnz2NI_Reports.VNavHover")
			await page.click("#sub-301650500 > div.VNavLeaf.VNavLevel2.arfid301381600.ardbnz2NI_Reports.VNavHover")
			 *
			 *
			 * FIN NAVEGACION DESDE PAGINA PRINCIPAL*/

    //Página reportes:

    //	Seleccion de reporte:
    await page.waitForTimeout(tiempoEspera);
    await page.waitForSelector(
      "#T93250 > tbody > tr:nth-child(2) > td:nth-child(2) > nobr > span",
    );
    await page.click(
      "#T93250 > tbody > tr:nth-child(2) > td:nth-child(2) > nobr > span",
    );

    //Ejecución de reporte:
    await page.waitForTimeout(tiempoEspera);
    await page.waitForSelector("#WIN_0_93272");
    await page.click("#WIN_0_93272");

    await page.waitForTimeout(tiempoEspera);

    //GET IFRAME:
    await page.waitForTimeout(tiempoEspera);
    await page.waitForSelector("#WIN_0_93220 > iframe");
    const elementHandle = await page.$("#WIN_0_93220 > iframe");
    const frame = await elementHandle.contentFrame();

    await page.waitForTimeout(tiempoEspera);
    await frame.waitForSelector(
      "#toolbar > table > tbody > tr:nth-child(2) > td:nth-child(5) > input",
    );
    await frame.click(
      "#toolbar > table > tbody > tr:nth-child(2) > td:nth-child(5) > input",
    );

    await page.waitForTimeout(tiempoEspera);
    await frame.waitForSelector("#exportFormat");
    await frame.select("#exportFormat", "xls");

    await frame.waitForTimeout(tiempoEspera);
    await frame.waitForSelector("#exportReportDialogokButton > input");
    await frame.click("#exportReportDialogokButton > input");

    await page.waitForTimeout(tiempoEspera);
    await browser.close();
    //})();

    return true;
  }

  async procesarReportRemedy(argumentos) {
    console.log("Leyendo archivo Report Remedy");

    console.log("Procesando Report Remedy");
    console.log("Report Remedy:");
    console.log(argumentos[0]);
    console.log("Ruta informe Incidencias: " + argumentos[1]);

    const pathInformeIncidencias = path.join(argumentos[1]);

    var reportIncidencias = argumentos[0];
    var informeIncidencias = {};
    var numColumnasInforme = 0;
    var numFilasInforme = 0;

    //PROCESAMIENTO:
    return await XlsxPopulate.fromFileAsync(
      path.normalize(pathInformeIncidencias),
    )
      .then((workbook) => {
        console.log("Archivo Cargado: Informe Incidencias");
        informeIncidencias = workbook;
        console.log(informeIncidencias);
        return true;
      })
      .then(() => {
        var hojaHistorial = informeIncidencias.sheet("DB-Historial");
        numColumnasInforme = informeIncidencias
          .sheet("DB-Historial")
          .usedRange()._numColumns;
        numFilasInforme = informeIncidencias
          .sheet("DB-Historial")
          .usedRange()._numRows;

        console.log("Filas: " + numFilasInforme);
        console.log("Columnas: " + numColumnasInforme);

        //Obtener Cabeceras Informe Incidencias:
        console.log("Incidencias en Report");

        var incidenciaAnalizada = "";
        var estadoActualIncidencia = [];
        var registrosCreados = 0;
        var registrosModificados = 0;
        var registrosCerrados = 0;

        //MAPEO DE CABECERAS:
        var cabecerasInforme = [
          "Fecha Cambio",
          "Hora Cambio",
          "Incident Number",
          "Servicio RMCA",
          "Service",
          "Assigned Group",
          "Submit Date",
          "Status",
          "Hub Actual",
          "Prioridad",
          "Summary",
          "Descripcion",
          "Solicitante",
          "Estado",
          "PROCESO RMCA",
          "Impacto",
          "Responsable",
          "W.A",
          "Descripción WA",
          "CRQ",
          "FECHA CRQ",
          "TARGET DATE",
          "Submit week",
          "Fecha cierre",
        ];

        var cabecerasReport = [
          "Incident ID*",
          "Service Request ID",
          "Last Name+",
          "First Name",
          "Summary*",
          "Service*+",
          "Priority*",
          "Status*",
          "Assigned Group*+",
          "Assignee+",
          "Target Date",
          "Incident Type*",
          "Submitter*",
          "Submit Date",
          "operational_categorization_tier_3",
        ];

        var mapeoHistorialReport = [
          "incident_id*+",
          "incident_type*",
          "service*+",
          "assigned_group*+",
          "assignee+",
          "submit_date",
          "status*",
          "operational_categorization_tier_3",
          "priority*",
          "summary*",
          "",
          "",
          "",
          "",
          "",
          "",
          "",
          "",
          "",
          "",
          "target_date",
        ];

        var now = moment();
        var time = now.hour() + ":" + now.minutes() + ":" + now.seconds();

        //Loop por incidencias abiertas para analisis:
        for (var i = 0; i < reportIncidencias[0].data.length; i++) {
          numFilasInforme = informeIncidencias
            .sheet("DB-Historial")
            .usedRange()._numRows;

          //Busqueda de Incidencias:
          incidenciaAnalizada = reportIncidencias[0].data[i]["incident_id*+"];
          incidenciaAnalizada = incidenciaAnalizada.trim();

          console.log("ANALIZANDO: " + incidenciaAnalizada);

          //Variables y flags de control:
          var incidenciaEncontrada = false;
          var incidenciaNueva = true;
          var incidenciaModificada = false;

          //Inicializa el Estado Actual:
          estadoActualIncidencia = [];
          for (var j = 0; j < numColumnasInforme - 3; j++) {
            estadoActualIncidencia.push("");
          }

          for (var j = 0; j < numFilasInforme; j++) {
            //Incidencia Encontrada
            if (
              incidenciaAnalizada ==
              hojaHistorial
                .row(j + 1)
                .cell(3)
                .value()
            ) {
              console.log("ENCONTRADO");
              incidenciaEncontrada = true;

              //Registrar campos Incidencia:
              for (var k = 3; k < numColumnasInforme; k++) {
                if (
                  hojaHistorial
                    .row(j + 1)
                    .cell(k)
                    .value()
                ) {
                  estadoActualIncidencia[k - 3] = hojaHistorial
                    .row(j + 1)
                    .cell(k)
                    .value();
                }
              }
            }
          }

          //Estado de incidencia Analizada:
          console.log("Estado Incidencia:" + incidenciaAnalizada);
          console.log(estadoActualIncidencia);

          //Detecta Cambios en la incidencia:
          if (incidenciaEncontrada) {
            console.log("BUSCANDO CAMBIOS...");
            for (var j = 0; j < mapeoHistorialReport.length; j++) {
              if (
                estadoActualIncidencia[j] !=
                  reportIncidencias[0].data[i][mapeoHistorialReport[j]] &&
                mapeoHistorialReport[j] != "" &&
                reportIncidencias[0].data[i][mapeoHistorialReport[j]] !=
                  undefined
              ) {
                console.log("CABECERA " + mapeoHistorialReport[j]);
                console.log(
                  "CAMBIO " +
                    reportIncidencias[0].data[i][mapeoHistorialReport[j]],
                );
                console.log("Marcando Fecha/Hora cambio:");
                informeIncidencias
                  .sheet("DB-Historial")
                  .row(numFilasInforme + 1)
                  .cell(1)
                  .value(moment().toDate());
                informeIncidencias
                  .sheet("DB-Historial")
                  .row(numFilasInforme + 1)
                  .cell(2)
                  .value(time);
                informeIncidencias
                  .sheet("DB-Historial")
                  .row(numFilasInforme + 1)
                  .cell(3)
                  .value(reportIncidencias[0].data[i]["incident_id*+"]);
                informeIncidencias
                  .sheet("DB-Historial")
                  .row(numFilasInforme + 1)
                  .cell(3 + j)
                  .value(reportIncidencias[0].data[i][mapeoHistorialReport[j]]);
                informeIncidencias
                  .sheet("DB-Historial")
                  .row(numFilasInforme + 1)
                  .cell(26)
                  .value("Modificado");
                incidenciaModificada = true;
              }
            }
          }

          //Detecta si la incidencia es nueva:
          for (var j = 0; j < estadoActualIncidencia.length; j++) {
            if (
              estadoActualIncidencia[j] != "" ||
              reportIncidencias[0].data[i]["status*"] == "Resolved" ||
              reportIncidencias[0].data[i]["service*+"] == undefined
            ) {
              incidenciaNueva = false;
            }
          }

          if (incidenciaNueva) {
            //Crea una entrada nueva en el historico:
            console.log("Creando nueva entrada");
            var indiceColumna = 0;
            registrosCreados++;

            console.log("Marcando Fecha/Hora cambio:");
            informeIncidencias
              .sheet("DB-Historial")
              .row(numFilasInforme + 1)
              .cell(1)
              .value(moment().toDate());
            informeIncidencias
              .sheet("DB-Historial")
              .row(numFilasInforme + 1)
              .cell(2)
              .value(time);

            for (var prop in reportIncidencias[0].data[i]) {
              console.log(prop);
              indiceColumna = 0;
              switch (prop) {
                case "incident_id*+":
                  indiceColumna = 3;
                  break;
                case "incident_type*":
                  indiceColumna = 4;
                  break;
                case "service*+":
                  indiceColumna = 5;
                  break;
                case "assigned_group*+":
                  indiceColumna = 6;
                  break;
                case "assignee+":
                  indiceColumna = 7;
                  break;
                case "submit_date":
                  indiceColumna = 8;
                  break;
                case "status*":
                  indiceColumna = 9;
                  break;
                case "operational_categorization_tier_3":
                  indiceColumna = 10;
                  break;
                case "priority*":
                  indiceColumna = 11;
                  break;
                case "summary*":
                  indiceColumna = 12;
                  break;
                case "submitter*":
                  indiceColumna = 14;
                  break;
                case "target_date":
                  indiceColumna = 23;
                  break;
                case "last_name+":
                  indiceColumna = 0;
                  break;
                case "first_name+":
                  indiceColumna = 0;
                  break;
                case "last_modified_date":
                  indiceColumna = 0;
                  break;
                default:
                  console.log("Columna no analizada: " + prop);
                  break;
              }

              if (indiceColumna != 0) {
                informeIncidencias
                  .sheet("DB-Historial")
                  .row(numFilasInforme + 1)
                  .cell(indiceColumna)
                  .value(reportIncidencias[0].data[i][prop]);
                informeIncidencias
                  .sheet("DB-Historial")
                  .row(numFilasInforme + 1)
                  .cell(26)
                  .value("Entrada");
              }
            }
          } else {
          }

          if (incidenciaModificada) registrosModificados++;
          if (incidenciaEncontrada == false) {
            console.log("No ENCONTRADO");
          }
        }

        //MARCADO INCICENCIAS POTENCIALMENTE CERRADAS/CANCELADAS:

        console.log("Detectando Canceladas/Cerradas: ");

        numFilasInforme = informeIncidencias
          .sheet("DB-Historial")
          .usedRange()._numRows;
        var incidenciasCerradas = [];

        for (var i = 1; i < numFilasInforme; i++) {
          incidenciaAnalizada = hojaHistorial
            .row(i + 1)
            .cell(3)
            .value();
          incidenciaEncontrada = false;
          if (incidenciaAnalizada == undefined) continue;
          if (
            reportIncidencias[0].data.find(
              (i) => i["incident_id*+"] == incidenciaAnalizada,
            )
          ) {
            incidenciaEncontrada = true;
          } else {
            if (incidenciasCerradas.indexOf(incidenciaAnalizada) < 0) {
              console.log(
                "No encontrada: " +
                  incidenciaAnalizada +
                  " Marcanco 'Cancelled/Closed'",
              );
              incidenciasCerradas.push(incidenciaAnalizada);
            }
          }
        }

        //Eliminando Marcado de Cancelada Si el ultimo registro historico es cerrado/cancelado
        var estadoFinalIncidencia;
        for (var i = 0; i < incidenciasCerradas.length; i++) {
          estadoFinalIncidencia = "";
          for (var j = 1; j <= numFilasInforme; j++) {
            if (
              informeIncidencias.sheet("DB-Historial").row(j).cell(3).value() ==
                incidenciasCerradas[i] &&
              informeIncidencias.sheet("DB-Historial").row(j).cell(9).value() !=
                ""
            ) {
              estadoFinalIncidencia = informeIncidencias
                .sheet("DB-Historial")
                .row(j)
                .cell(9)
                .value();
            }
          }
          if (estadoFinalIncidencia == "Cancelled/Closed") {
            incidenciasCerradas.splice(i, 1);
            i--;
          }
        }

        //Marcado de INCIDENCIAS CERRADAS/ CANCELADAS:
        /*
				var indexIncidencia = -1;
				for(var i = 1; i<=numFilasInforme; i++){
					if(informeIncidencias.sheet("DB-Historial").row(i).cell(9).value() == "Cancelled/Closed"){
						indexIncidencia = incidenciasCerradas.indexOf(informeIncidencias.sheet("DB-Historial").row(i).cell(3).value())
						if(indexIncidencia>-1){
							incidenciasCerradas.splice(indexIncidencia,1);
						}
					}
				}*/

        for (var i = 0; i < incidenciasCerradas.length; i++) {
          informeIncidencias
            .sheet("DB-Historial")
            .row(numFilasInforme + 1 + i)
            .cell(1)
            .value(moment().toDate());
          informeIncidencias
            .sheet("DB-Historial")
            .row(numFilasInforme + 1 + i)
            .cell(2)
            .value(time);
          informeIncidencias
            .sheet("DB-Historial")
            .row(numFilasInforme + 1 + i)
            .cell(3)
            .value(incidenciasCerradas[i]);
          informeIncidencias
            .sheet("DB-Historial")
            .row(numFilasInforme + 1 + i)
            .cell(9)
            .value("Cancelled/Closed");
          informeIncidencias
            .sheet("DB-Historial")
            .row(numFilasInforme + 1 + i)
            .cell(26)
            .value("Salida");
          registrosCerrados++;
        }

        console.log("**************************");
        console.log("Resumen:");
        console.log("Nuevas: " + registrosCreados);
        console.log("Modificados: " + registrosModificados);
        console.log("Canceladas/Cerradas: " + registrosCerrados);
        console.log("**************************");

        //Fin de procesamiento:
        console.log("Escribiendo archivo...");
        console.log("Path: " + path.normalize(pathInformeIncidencias));

        return informeIncidencias
          .toFileAsync(path.normalize(pathInformeIncidencias))
          .then(() => {
            console.log("Fin del procesamiento");
            return true;
          })
          .catch((err) => {
            console.log("Se ha producido un error interno: ");
            console.log(err);
            var tituloError =
              "Se ha producido un error escribiendo el archivo: " +
              path.normalize(pathInformeIncidencias);
            return false;
          });

        return true;
      })
      .catch((err) => {
        console.log("Se ha producido un error interno: ");
        console.log(err);
        var tituloError =
          "Se ha producido un error interno cargando los archivos.";
        mainWindow.webContents.send("onErrorInterno", tituloError, err);
        return false;
      });
  }

  async cargarExcel(pathExcel) {
    return new Promise((resolve) => {
      XlsxPopulate.fromFileAsync(path.normalize(pathExcel)).then((excel) => {
        if (excel === undefined) {
          resolve(false);
        }
        resolve(excel);
      });
    });
  }

  async procesarExtraccionPowerBI(argumentos) {
    console.log("Leyendo archivo Report Remedy");

    console.log("Procesando Report Remedy");
    console.log("Report Power BI:");
    console.log(argumentos[0]);
    console.log("Ruta informe Incidencias: " + argumentos[1]);

    const pathInformeIncidencias = path.join(argumentos[1]);

    var informeIncidencias = {};
    var numColumnasInforme = 0;
    var numFilasInforme = 0;

    var excelReportPowerBi = {};

    excelReportPowerBi = await this.cargarExcel(path.normalize(argumentos[0]));

    var reportIncidencias = [
      {
        data: [],
      },
    ];

    //**********************
    //CONVERTIR EXCEL POWER BI EN OBJETO:
    //**********************

    var numeroRegistrosPowerBI = excelReportPowerBi
      .sheet("Data_Backlog")
      .usedRange()._numRows;

    var numeroCabecerasPowerBI = excelReportPowerBi
      .sheet("Data_Backlog")
      .usedRange()._numColumns;

    //Limpiar registros:
    var objetoIncidencia = {};
    var cabecerasPowerBi = [];

    for (var j = 1; j <= numeroCabecerasPowerBI; j++) {
      cabecerasPowerBi.push(
        excelReportPowerBi.sheet("Data_Backlog").cell(1, j).value(),
      );
    }

    for (var i = 2; i <= numeroRegistrosPowerBI; i++) {
      objetoIncidencia = {};
      for (var j = 1; j <= numeroCabecerasPowerBI; j++) {
        objetoIncidencia[cabecerasPowerBi[j - 1]] = excelReportPowerBi
          .sheet("Data_Backlog")
          .cell(i, j)
          .value();
      }

      reportIncidencias[0].data.push(objetoIncidencia);
    }

    console.log("OBJETO POWER BI:");
    console.log(reportIncidencias[0].data);

    //**********************
    // FILTRADO OBJETO POWER BI
    //**********************
    console.log("Iniciando filtrado: ");
    for (var i = 0; i <= reportIncidencias[0].data.length; i++) {
      if (reportIncidencias[0].data[i]) {
        if (
          reportIncidencias[0].data[i]["Service"] !=
          "VFES-SAP ECC 6 0 RMCA-PROD"
        ) {
          reportIncidencias[0].data.splice(i, 1);
          i--;
        }
      }
    }

    console.log("OBJETO POWER BI (FILTRADO):");
    console.log(reportIncidencias[0].data);

    //PROCESAMIENTO:
    return await XlsxPopulate.fromFileAsync(
      path.normalize(pathInformeIncidencias),
    )
      .then((workbook) => {
        console.log("Archivo Cargado: Informe Incidencias...");
        informeIncidencias = workbook;
        //console.log(informeIncidencias);
        console.log("OK");
        return true;
      })
      .then(() => {
        var hojaHistorial = informeIncidencias.sheet("DB-Historial");
        numColumnasInforme = informeIncidencias
          .sheet("DB-Historial")
          .usedRange()._numColumns;
        numFilasInforme = informeIncidencias
          .sheet("DB-Historial")
          .usedRange()._numRows;

        console.log("Filas: " + numFilasInforme);
        console.log("Columnas: " + numColumnasInforme);

        //Obtener Cabeceras Informe Incidencias:
        console.log("Incidencias en Report");

        var incidenciaAnalizada = "";
        var estadoActualIncidencia = [];
        var registrosCreados = 0;
        var registrosModificados = 0;
        var registrosCerrados = 0;

        //MAPEO DE CABECERAS:

        //Cabeceras Informe Incidencias:
        var cabecerasInforme = [
          "Fecha Cambio",
          "Hora Cambio",
          "Incident Number",
          "Servicio RMCA",
          "Service",
          "Assigned Group",
          "Submit Date",
          "Status",
          "Hub Actual",
          "Prioridad",
          "Summary",
          "Descripcion",
          "Solicitante",
          "Estado",
          "PROCESO RMCA",
          "Impacto",
          "Responsable",
          "W.A",
          "Descripción WA",
          "CRQ",
          "FECHA CRQ",
          "TARGET DATE",
          "Submit week",
          "Fecha cierre",
        ];

        //Cabeceras de Power BI:
        //var cabecerasReport = ["Incident ID*","Service Request ID","Last Name+","First Name","Summary*","Service*+","Priority*","Status*","Assigned Group*+","Assignee+","Target Date","Incident Type*","Submitter*","Submit Date","operational_categorization_tier_3"]
        var cabecerasReport = [
          "Incident Number",
          "Support Group Name",
          "Submit Date",
          "Summary",
          "Priority",
          "Service Type",
          "Assigned Group",
          "Status",
          "Status Reason",
          "Cause",
          "Estimated Resolution Date",
          "Required Resolution Datetime",
          "Service",
          "Closed Date",
          "Last Resolved Date",
          "Last Modified Date",
          "Categorization Tier 1",
          "Categorization Tier 2",
          "Categorization Tier 3",
          "Product Categorization Tier 1",
          "Product Categorization Tier 2",
          "Product Categorization Tier 3",
          "Reported Source",
          "Closure Product Name",
          "Environment",
          "Type",
          "SPIRIT",
          "Antiguëdad",
          "Rango",
          "Estado",
          "Severity",
          "Inc. Number",
          "Asignado",
          "Group",
          "Group Creator",
          "Product Resp",
          "Product Resp II",
          "Cat",
          "Antiguëdad 2",
          "Antigüedad modificacion",
          "Rango 2",
          "Groups Number",
          "Stack",
          "Inc_Type",
          "Revised_Type",
        ];

        //Mapeo:
        //var mapeoHistorialReport = ["incident_id*+","incident_type*","service*+","assigned_group*+","assignee+","submit_date","status*","operational_categorization_tier_3","priority*","summary*","","","","","","","","","","","target_date"]
        var mapeoHistorialReport = [
          "Incident Number",
          "Type",
          "Service",
          "Assigned Group",
          "Asignado",
          "Submit Date",
          "Status",
          "Categorization Tier 3",
          "Priority",
          "Summary",
          "",
          "",
          "",
          "",
          "",
          "",
          "",
          "",
          "",
          "",
          "Estimated Resolution Date",
        ];

        var now = moment();
        var time = now.hour() + ":" + now.minutes() + ":" + now.seconds();

        //Loop por incidencias abiertas para analisis:
        for (var i = 0; i < reportIncidencias[0].data.length; i++) {
          numFilasInforme = informeIncidencias
            .sheet("DB-Historial")
            .usedRange()._numRows;

          //Busqueda de Incidencias:
          incidenciaAnalizada = reportIncidencias[0].data[i]["Incident Number"];
          incidenciaAnalizada = incidenciaAnalizada.trim();

          console.log("ANALIZANDO: " + incidenciaAnalizada);

          //Variables y flags de control:
          var incidenciaEncontrada = false;
          var incidenciaNueva = true;
          var incidenciaModificada = false;

          //Inicializa el Estado Actual:
          estadoActualIncidencia = [];
          for (var j = 0; j < numColumnasInforme - 3; j++) {
            estadoActualIncidencia.push("");
          }

          for (var j = 0; j < numFilasInforme; j++) {
            //Incidencia Encontrada
            if (
              incidenciaAnalizada ==
              hojaHistorial
                .row(j + 1)
                .cell(3)
                .value()
            ) {
              console.log("ENCONTRADO");
              incidenciaEncontrada = true;

              //Registrar campos Incidencia:
              for (var k = 3; k < numColumnasInforme; k++) {
                if (
                  hojaHistorial
                    .row(j + 1)
                    .cell(k)
                    .value()
                ) {
                  estadoActualIncidencia[k - 3] = hojaHistorial
                    .row(j + 1)
                    .cell(k)
                    .value();
                }
              }
            }
          }

          //Estado de incidencia Analizada:
          console.log("Estado Incidencia:" + incidenciaAnalizada);
          console.log(estadoActualIncidencia);

          //Detecta Cambios en la incidencia:
          if (incidenciaEncontrada) {
            console.log("BUSCANDO CAMBIOS...");
            for (var j = 0; j < mapeoHistorialReport.length; j++) {
              if (
                estadoActualIncidencia[j] !=
                  reportIncidencias[0].data[i][mapeoHistorialReport[j]] &&
                mapeoHistorialReport[j] != "" &&
                reportIncidencias[0].data[i][mapeoHistorialReport[j]] !=
                  undefined
              ) {
                console.log("CABECERA " + mapeoHistorialReport[j]);
                console.log(
                  "CAMBIO " +
                    reportIncidencias[0].data[i][mapeoHistorialReport[j]],
                );
                console.log("Marcando Fecha/Hora cambio:");
                informeIncidencias
                  .sheet("DB-Historial")
                  .row(numFilasInforme + 1)
                  .cell(1)
                  .value(moment().toDate());
                informeIncidencias
                  .sheet("DB-Historial")
                  .row(numFilasInforme + 1)
                  .cell(2)
                  .value(time);
                informeIncidencias
                  .sheet("DB-Historial")
                  .row(numFilasInforme + 1)
                  .cell(3)
                  .value(reportIncidencias[0].data[i]["Incident Number"]);
                informeIncidencias
                  .sheet("DB-Historial")
                  .row(numFilasInforme + 1)
                  .cell(3 + j)
                  .value(reportIncidencias[0].data[i][mapeoHistorialReport[j]]);
                informeIncidencias
                  .sheet("DB-Historial")
                  .row(numFilasInforme + 1)
                  .cell(26)
                  .value("Modificado");
                incidenciaModificada = true;
              }
            }
          }

          //Detecta si la incidencia es nueva:
          for (var j = 0; j < estadoActualIncidencia.length; j++) {
            if (
              estadoActualIncidencia[j] != "" ||
              reportIncidencias[0].data[i]["Status"] == "Resolved" ||
              reportIncidencias[0].data[i]["Service"] == undefined
            ) {
              incidenciaNueva = false;
            }
          }

          if (incidenciaNueva) {
            //Crea una entrada nueva en el historico:
            console.log("Creando nueva entrada");
            var indiceColumna = 0;
            registrosCreados++;

            console.log("Marcando Fecha/Hora cambio:");
            informeIncidencias
              .sheet("DB-Historial")
              .row(numFilasInforme + 1)
              .cell(1)
              .value(moment().toDate());
            informeIncidencias
              .sheet("DB-Historial")
              .row(numFilasInforme + 1)
              .cell(2)
              .value(time);

            for (var prop in reportIncidencias[0].data[i]) {
              console.log(prop);
              indiceColumna = 0;
              switch (prop) {
                case "Incident Number":
                  indiceColumna = 3;
                  break;
                case "Type":
                  indiceColumna = 4;
                  break;
                case "Service":
                  indiceColumna = 5;
                  break;
                case "Assigned Group":
                  indiceColumna = 6;
                  break;
                case "Asignado":
                  indiceColumna = 7;
                  break;
                case "Submit Date":
                  indiceColumna = 8;
                  break;
                case "Status":
                  indiceColumna = 9;
                  break;
                case "Categorization Tier 3":
                  indiceColumna = 10;
                  break;
                case "Priority":
                  indiceColumna = 11;
                  break;
                case "Summary":
                  indiceColumna = 12;
                  break;
                case "submitter*":
                  indiceColumna = 14;
                  break;
                case "Estimated Resolution Date":
                  indiceColumna = 23;
                  break;
                case "last_name+":
                  indiceColumna = 0;
                  break;
                case "first_name+":
                  indiceColumna = 0;
                  break;
                case "last_modified_date":
                  indiceColumna = 0;
                  break;
                default:
                  console.log("Columna no analizada: " + prop);
                  break;
              }

              if (indiceColumna != 0) {
                informeIncidencias
                  .sheet("DB-Historial")
                  .row(numFilasInforme + 1)
                  .cell(indiceColumna)
                  .value(reportIncidencias[0].data[i][prop]);
                informeIncidencias
                  .sheet("DB-Historial")
                  .row(numFilasInforme + 1)
                  .cell(26)
                  .value("Entrada");
              }
            }
          } else {
          }

          if (incidenciaModificada) registrosModificados++;
          if (incidenciaEncontrada == false) {
            console.log("No ENCONTRADO");
          }
        }

        //MARCADO INCICENCIAS POTENCIALMENTE CERRADAS/CANCELADAS:

        console.log("Detectando Canceladas/Cerradas: ");

        numFilasInforme = informeIncidencias
          .sheet("DB-Historial")
          .usedRange()._numRows;
        var incidenciasCerradas = [];

        for (var i = 1; i < numFilasInforme; i++) {
          incidenciaAnalizada = hojaHistorial
            .row(i + 1)
            .cell(3)
            .value();
          incidenciaEncontrada = false;
          if (incidenciaAnalizada == undefined) continue;
          if (
            reportIncidencias[0].data.find(
              (i) => i["Incident Number"] == incidenciaAnalizada,
            )
          ) {
            incidenciaEncontrada = true;
          } else {
            if (incidenciasCerradas.indexOf(incidenciaAnalizada) < 0) {
              console.log(
                "No encontrada: " +
                  incidenciaAnalizada +
                  " Marcanco 'Cancelled/Closed'",
              );
              incidenciasCerradas.push(incidenciaAnalizada);
            }
          }
        }

        //Eliminando Marcado de Cancelada Si el ultimo registro historico es cerrado/cancelado
        var estadoFinalIncidencia;
        for (var i = 0; i < incidenciasCerradas.length; i++) {
          estadoFinalIncidencia = "";
          for (var j = 1; j <= numFilasInforme; j++) {
            if (
              informeIncidencias.sheet("DB-Historial").row(j).cell(3).value() ==
                incidenciasCerradas[i] &&
              informeIncidencias.sheet("DB-Historial").row(j).cell(9).value() !=
                ""
            ) {
              estadoFinalIncidencia = informeIncidencias
                .sheet("DB-Historial")
                .row(j)
                .cell(9)
                .value();
            }
          }
          if (estadoFinalIncidencia == "Cancelled/Closed") {
            incidenciasCerradas.splice(i, 1);
            i--;
          }
        }

        //Marcado de INCIDENCIAS CERRADAS/ CANCELADAS:
        /*
				var indexIncidencia = -1;
				for(var i = 1; i<=numFilasInforme; i++){
					if(informeIncidencias.sheet("DB-Historial").row(i).cell(9).value() == "Cancelled/Closed"){
						indexIncidencia = incidenciasCerradas.indexOf(informeIncidencias.sheet("DB-Historial").row(i).cell(3).value())
						if(indexIncidencia>-1){
							incidenciasCerradas.splice(indexIncidencia,1);
						}
					}
				}*/

        for (var i = 0; i < incidenciasCerradas.length; i++) {
          informeIncidencias
            .sheet("DB-Historial")
            .row(numFilasInforme + 1 + i)
            .cell(1)
            .value(moment().toDate());
          informeIncidencias
            .sheet("DB-Historial")
            .row(numFilasInforme + 1 + i)
            .cell(2)
            .value(time);
          informeIncidencias
            .sheet("DB-Historial")
            .row(numFilasInforme + 1 + i)
            .cell(3)
            .value(incidenciasCerradas[i]);
          informeIncidencias
            .sheet("DB-Historial")
            .row(numFilasInforme + 1 + i)
            .cell(9)
            .value("Cancelled/Closed");
          informeIncidencias
            .sheet("DB-Historial")
            .row(numFilasInforme + 1 + i)
            .cell(26)
            .value("Salida");
          registrosCerrados++;
        }

        console.log("**************************");
        console.log("Resumen:");
        console.log("Nuevas: " + registrosCreados);
        console.log("Modificados: " + registrosModificados);
        console.log("Canceladas/Cerradas: " + registrosCerrados);
        console.log("**************************");

        //Fin de procesamiento:
        console.log("Escribiendo archivo...");
        console.log("Path: " + path.normalize(pathInformeIncidencias));

        return informeIncidencias
          .toFileAsync(path.normalize(pathInformeIncidencias))
          .then(() => {
            console.log("Fin del procesamiento");
            return true;
          })
          .catch((err) => {
            console.log("Se ha producido un error interno: ");
            console.log(err);
            var tituloError =
              "Se ha producido un error escribiendo el archivo: " +
              path.normalize(pathInformeIncidencias);
            return false;
          });

        return true;
      })
      .catch((err) => {
        console.log("Se ha producido un error interno: ");
        console.log(err);
        var tituloError =
          "Se ha producido un error interno cargando los archivos.";
        mainWindow.webContents.send("onErrorInterno", tituloError, err);
        return false;
      });
  }

  async renderizarReportRemedy(argumentos) {
    console.log("Leyendo archivo Report Remedy");

    console.log("Procesando Report Remedy");
    console.log("Ruta informe Incidencias: " + argumentos[0]);

    const pathInformeIncidencias = path.join(argumentos[0]);

    var reportIncidencias = argumentos[0];
    var informeIncidencias = {};
    var numColumnasInforme = 0;
    var numFilasInforme = 0;

    //PROCESAMIENTO:
    return await XlsxPopulate.fromFileAsync(
      path.normalize(pathInformeIncidencias),
    )
      .then((workbook) => {
        console.log("Archivo Cargado: Informe Incidencias");
        informeIncidencias = workbook;
        //console.log(informeIncidencias);
        //return true;
      })
      .then(() => {
        var hojaHistorial = informeIncidencias.sheet("DB-Historial");
        var hojaIncidenciasAbiertas = informeIncidencias.sheet("INC abiertas");
        var hojaIncidenciasCerradas = informeIncidencias.sheet(
          "INC resueltas-cerradas",
        );

        var numColumnasHistorial = informeIncidencias
          .sheet("DB-Historial")
          .usedRange()._numColumns;
        var numFilasHistorial = informeIncidencias
          .sheet("DB-Historial")
          .usedRange()._numRows;

        var numColumnasAbiertas = informeIncidencias
          .sheet("INC abiertas")
          .usedRange()._numColumns;
        var numFilasAbiertas = informeIncidencias
          .sheet("INC abiertas")
          .usedRange()._numRows;

        var numColumnasCerradas = informeIncidencias
          .sheet("INC resueltas-cerradas")
          .usedRange()._numColumns;
        var numFilasCerradas = informeIncidencias
          .sheet("INC resueltas-cerradas")
          .usedRange()._numRows;

        console.log("RENDERIZANDO");
        console.log("Filas Incidencias Abiertas: " + numFilasAbiertas);
        console.log("Filas Historico: " + numFilasHistorial);

        //Borra contenido de Hoja Incidencias Abiertas:
        for (var i = 1; i < numFilasAbiertas; i++) {
          for (var j = 1; j < numColumnasAbiertas; j++) {
            hojaIncidenciasAbiertas
              .row(i + 1)
              .cell(j)
              .clear();
          }
        }

        //Generar Lista de incidencias:
        var listaIncidencias = [];
        var incidenciaAnalizada = "";
        var estadoIncidenciaAnalizada = "";
        var indexIncidencia = -1;
        var arrayIncidencia = [];
        var arrayTiempoCambio = [];
        var incidenciaEncontrada = false;

        //numColumnasHistorial = 23

        //Crear Lista de incidencias:
        for (var i = 1; i < numFilasHistorial; i++) {
          incidenciaAnalizada = hojaHistorial
            .row(1 + i)
            .cell(3)
            .value();
          estadoIncidenciaAnalizada = hojaHistorial
            .row(1 + i)
            .cell(9)
            .value();
          incidenciaEncontrada = false;
          for (var m = 0; m < listaIncidencias.length; m++) {
            if (listaIncidencias[m][0] == incidenciaAnalizada) {
              incidenciaEncontrada = true;
            }
          }
          if (!incidenciaEncontrada) {
            //Creación Entrada Incidencia
            arrayIncidencia = [];
            for (var l = 0; l < numColumnasHistorial; l++) {
              arrayIncidencia.push("");
            }
            arrayIncidencia[0] = incidenciaAnalizada;
            arrayTiempoCambio.push(0);
            listaIncidencias.push(arrayIncidencia);
          }
        }

        //Rellenar lista de incidencias:
        for (var i = 0; i < listaIncidencias.length; i++) {
          for (var j = 1; j < numFilasHistorial; j++) {
            for (var k = 3; k < numColumnasHistorial; k++) {
              if (
                hojaHistorial
                  .row(j + 1)
                  .cell(k + 1)
                  .value() != "" &&
                hojaHistorial
                  .row(j + 1)
                  .cell(k + 1)
                  .value() != undefined &&
                hojaHistorial
                  .row(j + 1)
                  .cell(3)
                  .value() == listaIncidencias[i][0]
              ) {
                listaIncidencias[i][k - 2] = hojaHistorial
                  .row(j + 1)
                  .cell(k + 1)
                  .value();
              }
            }
          }
        }

        //Eliminar Incidencias Cancelled/Closed o Resolved
        var indicesEliminar = [];
        for (var i = 0; i < listaIncidencias.length; i++) {
          if (
            listaIncidencias[i][6] == "Cancelled/Closed" ||
            listaIncidencias[i][6] == "Resolved"
          ) {
            //Escribir en tabla Cerradas/Canceladas;
            for (var j = 0; j < listaIncidencias[i].length; j++) {
              hojaIncidenciasCerradas
                .row(
                  informeIncidencias.sheet("INC resueltas-cerradas").usedRange()
                    ._numRows,
                )
                .cell(j + 1)
                .value(listaIncidencias[i][j]);
            }

            listaIncidencias.splice(i, 1);
            i = i - 1;
          }
        }

        //Escribir Tabla de incidencias Abiertas:
        for (var i = 0; i < listaIncidencias.length; i++) {
          for (var j = 0; j < listaIncidencias[i].length; j++) {
            //listaIncidencias[i][k-3]= hojaHistorial.row(j+1).cell(k+1).value()
            hojaIncidenciasAbiertas
              .row(i + 2)
              .cell(j + 1)
              .value(listaIncidencias[i][j]);
          }
        }

        //console.log("Lista de Incidencias: ");
        //console.log(listaIncidencias)

        //Escribir Archivo:
        return informeIncidencias
          .toFileAsync(path.normalize(pathInformeIncidencias))
          .then(() => {
            console.log("Fin del procesamiento");
            return true;
          })
          .catch((err) => {
            console.log("Se ha producido un error interno: ");
            console.log(err);
            var tituloError =
              "Se ha producido un error escribiendo el archivo: " +
              path.normalize(pathInformeIncidencias);
            return false;
          });
        //return true
      })
      .catch((err) => {
        console.log("Se ha producido un error interno: ");
        console.log(err);
        var tituloError =
          "Se ha producido un error interno cargando los archivos.";
        mainWindow.webContents.send("onErrorInterno", tituloError, err);
        return false;
      });
    //return false;
  }

  async renderizarBacklog(argumentos) {
    console.log("Leyendo archivo Report Remedy");

    console.log("Procesando Report Remedy");
    console.log("Ruta informe Incidencias: " + argumentos[0]);

    const pathInformeIncidencias = path.join(argumentos[0]);

    var reportIncidencias = argumentos[0];
    var informeIncidencias = {};
    var numColumnasInforme = 0;
    var numFilasInforme = 0;

    //PROCESAMIENTO:
    return await XlsxPopulate.fromFileAsync(
      path.normalize(pathInformeIncidencias),
    )
      .then((workbook) => {
        console.log("Archivo Cargado: Informe Incidencias");
        informeIncidencias = workbook;
        //console.log(informeIncidencias);
        //return true;
      })
      .then(() => {
        var hojaHistorial = informeIncidencias.sheet("DB-Historial");

        var hojaIncidenciasAbiertas = informeIncidencias.sheet("INC abiertas");

        var hojaIncidenciasBacklog = informeIncidencias.sheet("DB-Backlog");

        var numColumnasHistorial = informeIncidencias
          .sheet("DB-Historial")
          .usedRange()._numColumns;
        var numFilasHistorial = informeIncidencias
          .sheet("DB-Historial")
          .usedRange()._numRows;

        var numColumnasAbiertas = informeIncidencias
          .sheet("INC abiertas")
          .usedRange()._numColumns;
        var numFilasAbiertas = informeIncidencias
          .sheet("INC abiertas")
          .usedRange()._numRows;

        var numColumnasBacklog = informeIncidencias
          .sheet("DB-Backlog")
          .usedRange()._numColumns;
        var numFilasBacklog = informeIncidencias
          .sheet("DB-Backlog")
          .usedRange()._numRows;

        console.log("RENDERIZANDO BACKLOG");

        //Generar Lista de incidencias:
        var listaIncidencias = [];
        var incidenciaAnalizada = "";
        var estadoIncidenciaAnalizada = "";
        var indexIncidencia = -1;
        var arrayIncidencia = [];
        var arrayTiempoCambio = [];
        var incidenciaEncontrada = false;

        //Importar Array de incidencias Unicas:
        for (var i = 0; i < numFilasHistorial; i++) {
          listaIncidencias.push(
            hojaHistorial
              .row(2 + i)
              .cell(3)
              .value(),
          );
        }

        var incidenciasUnicas = listaIncidencias.filter((item, index) => {
          return listaIncidencias.indexOf(item) === index;
        });

        var fechaUltimaEntrada = new moment();
        var fechaUltimaSalida = new moment();

        function ExcelDateToJSDate(serial) {
          var utc_days = Math.floor(serial - 25569);
          var utc_value = utc_days * 86400;
          var date_info = new Date(utc_value * 1000);

          var fractional_day = serial - Math.floor(serial) + 0.0000001;

          var total_seconds = Math.floor(86400 * fractional_day);

          var seconds = total_seconds % 60;

          total_seconds -= seconds;

          var hours = Math.floor(total_seconds / (60 * 60));
          var minutes = Math.floor(total_seconds / 60) % 60;

          return new Date(
            date_info.getFullYear(),
            date_info.getMonth(),
            date_info.getDate(),
            hours,
            minutes,
            seconds,
          );
        }

        //Borrar Hoja de DB-Backlog:
        for (var i = 0; i < numFilasBacklog; i++) {
          hojaIncidenciasBacklog
            .row(i + 2)
            .cell(1)
            .value("");
          hojaIncidenciasBacklog
            .row(i + 2)
            .cell(2)
            .value("");
          hojaIncidenciasBacklog
            .row(i + 2)
            .cell(3)
            .value("");
        }

        //Iterar por incidencias Unicas:
        var cuentaFila = 2;
        var flagSalida = false;

        for (var i = 0; i < incidenciasUnicas.length; i++) {
          flagSalida = false;
          for (var j = 0; j < numFilasHistorial; j++) {
            //Filtra por incidencia:
            if (
              hojaHistorial
                .row(2 + j)
                .cell(3)
                .value() == incidenciasUnicas[i]
            ) {
              //Actua en Entrada:
              if (
                hojaHistorial
                  .row(2 + j)
                  .cell(26)
                  .value() == "Entrada" &&
                !flagSalida
              ) {
                console.log("Entrada:");
                console.log("INC: " + incidenciasUnicas[i]);
                console.log(
                  hojaHistorial
                    .row(2 + j)
                    .cell(1)
                    .value(),
                );
                var fechaExcel = Number(
                  hojaHistorial
                    .row(2 + j)
                    .cell(1)
                    .value(),
                );
                fechaUltimaEntrada = moment(ExcelDateToJSDate(fechaExcel));
                console.log(fechaUltimaEntrada);
                hojaIncidenciasBacklog
                  .row(cuentaFila)
                  .cell(2)
                  .value(fechaUltimaEntrada.format("DD/MM/YYYY"));
                hojaIncidenciasBacklog
                  .row(cuentaFila)
                  .cell(1)
                  .value(incidenciasUnicas[i]);
              }

              //Actua en Salida:
              if (
                hojaHistorial
                  .row(2 + j)
                  .cell(26)
                  .value() == "Salida" &&
                !flagSalida
              ) {
                console.log("Salida");
                console.log("INC: " + incidenciasUnicas[i]);
                console.log(
                  hojaHistorial
                    .row(2 + j)
                    .cell(1)
                    .value(),
                );
                var fechaExcel = Number(
                  hojaHistorial
                    .row(2 + j)
                    .cell(1)
                    .value(),
                );
                fechaUltimaSalida = moment(ExcelDateToJSDate(fechaExcel));
                console.log(fechaUltimaSalida);
                hojaIncidenciasBacklog
                  .row(cuentaFila)
                  .cell(3)
                  .value(fechaUltimaSalida.format("DD/MM/YYYY"));
                hojaIncidenciasBacklog
                  .row(cuentaFila)
                  .cell(1)
                  .value(incidenciasUnicas[i]);
                flagSalida = true;
              }

              //Detectar flag entrada:
              if (
                flagSalida &&
                (hojaHistorial
                  .row(2 + j)
                  .cell(26)
                  .value() == "Modificado" ||
                  hojaHistorial
                    .row(2 + j)
                    .cell(26)
                    .value() == "Modificación Manual")
              ) {
                //Comprueba que el cambio es posterior a la salida:
                var fechaExcel = Number(
                  hojaHistorial
                    .row(2 + j)
                    .cell(1)
                    .value(),
                );
                if (fechaUltimaSalida.isBefore(ExcelDateToJSDate(fechaExcel))) {
                  cuentaFila++;
                  flagSalida = false;
                  console.log("Reentrada:");
                  console.log("INC: " + incidenciasUnicas[i]);
                  console.log(
                    hojaHistorial
                      .row(2 + j)
                      .cell(1)
                      .value(),
                  );
                  fechaUltimaEntrada = moment(ExcelDateToJSDate(fechaExcel));
                  console.log(fechaUltimaEntrada);
                  hojaIncidenciasBacklog
                    .row(cuentaFila)
                    .cell(2)
                    .value(fechaUltimaEntrada.format("DD/MM/YYYY"));
                  hojaIncidenciasBacklog
                    .row(cuentaFila)
                    .cell(1)
                    .value(incidenciasUnicas[i]);
                }
              }
            }
          }
          cuentaFila++;
        }

        //numColumnasHistorial = 23

        /*
				//Crear Lista de incidencias:
				for( var i = 1; i<numFilasHistorial; i++){
					incidenciaAnalizada = hojaHistorial.row(1+i).cell(3).value();
					estadoIncidenciaAnalizada = hojaHistorial.row(1+i).cell(9).value();
					incidenciaEncontrada= false
					for(var m= 0; m<listaIncidencias.length; m++){
						if(listaIncidencias[m][0]==incidenciaAnalizada){
							incidenciaEncontrada= true;
						}
					}
					if(!incidenciaEncontrada){
						//Creación Entrada Incidencia
						arrayIncidencia = [];
						for(var l = 0; l<numColumnasHistorial; l++){
							arrayIncidencia.push("");
						}
						arrayIncidencia[0] = incidenciaAnalizada; 
						arrayTiempoCambio.push(0);
						listaIncidencias.push(arrayIncidencia);
					}
				}

				//Rellenar lista de incidencias:
				for(var i = 0; i<listaIncidencias.length; i++){
					for(var j = 1; j<numFilasHistorial; j++){
						for(var k=3; k< numColumnasHistorial; k++){
							if( (hojaHistorial.row(j+1).cell(k+1).value()!= "") && 
								(hojaHistorial.row(j+1).cell(k+1).value()!= undefined) &&
								(hojaHistorial.row(j+1).cell(3).value() == listaIncidencias[i][0])){
									listaIncidencias[i][k-2]= hojaHistorial.row(j+1).cell(k+1).value()
							}
						}
					}
				}

				//Eliminar Incidencias Cancelled/Closed o Resolved
				var indicesEliminar = []
				for(var i = 0; i<listaIncidencias.length; i++){
					if(listaIncidencias[i][6]=="Cancelled/Closed" || listaIncidencias[i][6]== "Resolved"){
						//Escribir en tabla Cerradas/Canceladas;
						for(var j = 0; j<listaIncidencias[i].length; j++){
							hojaIncidenciasCerradas.row(informeIncidencias.sheet("INC resueltas-cerradas").usedRange()._numRows).cell(j+1).value(listaIncidencias[i][j])
						}

						listaIncidencias.splice(i,1)
						i=i-1;
					}
				}

				//Escribir Tabla de incidencias Abiertas:
				for(var i = 0; i<listaIncidencias.length; i++){
					for(var j = 0; j<listaIncidencias[i].length; j++){
						//listaIncidencias[i][k-3]= hojaHistorial.row(j+1).cell(k+1).value()
						hojaIncidenciasAbiertas.row(i+2).cell(j+1).value(listaIncidencias[i][j])
					}
				}

			
				//console.log("Lista de Incidencias: ");
				//console.log(listaIncidencias)
				*/

        //Escribir Archivo:
        return informeIncidencias
          .toFileAsync(path.normalize(pathInformeIncidencias))
          .then(() => {
            console.log("Fin del procesamiento");
            return true;
          })
          .catch((err) => {
            console.log("Se ha producido un error interno: ");
            console.log(err);
            var tituloError =
              "Se ha producido un error escribiendo el archivo: " +
              path.normalize(pathInformeIncidencias);
            return false;
          });
        //return true
      })
      .catch((err) => {
        console.log("Se ha producido un error interno: ");
        console.log(err);
        var tituloError =
          "Se ha producido un error interno cargando los archivos.";
        mainWindow.webContents.send("onErrorInterno", tituloError, err);
        return false;
      });
    //return false;
  }
}

module.exports = ProcesosRemedy;
