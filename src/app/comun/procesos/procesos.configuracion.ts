
interface LibreriaProcesos {
	tipo: "proceso" | "directorio" | "redireccion";
	nombre: string;
	categoria: "Remedy" | "Spool" | "Desarrollador" | "Santander" | "Google" | "Despacho" | "Import" | "KPIs" | "Documentos"; 
	descripcion: string;
	autor?: string;
	opciones?: any;
	argumentos?: Argumentos[];
	salida?: Salida[];
	subCategoria?: LibreriaProcesos[];
}

interface Argumentos {
	tipo: tipoArgumento;
	identificador: string;
	obligado: boolean;
	formulario: FormularioArgumento;
	valor?: any;
}

interface Salida {
	tipo: tipoSalida; 
	valor?: any;
}

interface FormularioArgumento {
	titulo: string;
	tipo: "texto" | "ruta" | "numero" | "seleccion" | "fecha" | "boolean" | "archivo" | "objeto";
	placeholder: string;
	valorDefault: any;
	accept?: string;
}

type tipoSalida = "string" | "boolean" | "spool" | "xlsxRaw" | "ruta" | "numero" | "fecha" | "texto";
type tipoArgumento = "string" | "boolean" | "spool" | "xlsxRaw" | "ruta" | "numero" | "fecha" | "texto" | "objeto";

var libreriaProcesos: LibreriaProcesos[]= [
	{
		nombre: "Spool",
		tipo: "directorio",
		categoria: "Spool",
		descripcion: "Procesado de Spools",
		subCategoria: [
			{

				nombre: "Formatear Spool",
				categoria: "Spool",
				tipo: "proceso",
				descripcion: "Elimina las cabeceras y formatea las spools de SAP para su futuro procesamiento.",
				autor: "Carlos Cabrera",
				argumentos: [
					{
						tipo: "texto", 
						obligado: true, 
						identificador: "spoolInput", 
						formulario: {
							titulo: "Ruta Spool entrada",
							tipo: "archivo",
							accept: ".txt, .TXT",
							placeholder: "Introduzca la ruta del archivo de entrada.",
							valorDefault: ""
						}
					},
					{
						tipo: "texto", 
						obligado: true, 
						identificador: "spoolOutput", 
						formulario: {
							titulo: "Directorio de salida:",
							tipo: "ruta",
							placeholder: "Ruta de guardado...",
							valorDefault: ""
						}
					},
					{
						tipo: "texto", 
						obligado: true, 
						identificador: "nombreSpoolOutput", 
						formulario: {
							titulo: "Nombre archivo salida",
							tipo: "texto",
							placeholder: "Introduzca la ruta del archivo de salida.",
							valorDefault: ""
						}
					}],
				opciones: null,
				salida: [{tipo: "boolean", valor: false}]
			},
			{

				nombre: "Filtrar Fecha Spool",
				categoria: "Spool",
				tipo: "proceso",
				descripcion: "Filtra la fecha de compensacion de unas Spool",
				autor: "Carlos Cabrera",
				argumentos: [
					{
						tipo: "texto", 
						obligado: true, 
						identificador: "spoolInput", 
						formulario: {
							titulo: "Ruta Spool entrada",
							tipo: "archivo",
							accept: ".txt, .TXT",
							placeholder: "Introduzca la ruta del archivo de entrada.",
							valorDefault: ""
						}
					},
					{
						tipo: "texto", 
						obligado: true, 
						identificador: "spoolOutput", 
						formulario: {
							titulo: "Directorio de salida:",
							tipo: "ruta",
							placeholder: "Ruta de guardado...",
							valorDefault: ""
						}
					},
					{
						tipo: "texto", 
						obligado: true, 
						identificador: "nombreSpoolOutput", 
						formulario: {
							titulo: "Nombre archivo salida",
							tipo: "texto",
							placeholder: "Introduzca la ruta del archivo de salida.",
							valorDefault: ""
						}
					}],
				opciones: null,
				salida: [{tipo: "boolean", valor: false}]
			},
			{

				nombre: "Obtener Objeto Documento Spool",
				categoria: "Spool",
				tipo: "proceso",
				descripcion: "Obtiene el elemento Objeto Documento de un Spool",
				autor: "Carlos Cabrera",
				argumentos: [
					{
						tipo: "texto", 
						obligado: true, 
						identificador: "spoolInput", 
						formulario: {
							titulo: "Ruta Spool entrada",
							tipo: "archivo",
							accept: ".txt, .TXT",
							placeholder: "Introduzca la ruta del archivo de entrada.",
							valorDefault: ""
						}
					},
					{
						tipo: "texto", 
						obligado: true, 
						identificador: "spoolOutput", 
						formulario: {
							titulo: "Directorio de salida:",
							tipo: "ruta",
							placeholder: "Ruta de guardado...",
							valorDefault: ""
						}
					},
					{
						tipo: "texto", 
						obligado: true, 
						identificador: "nombreSpoolOutput", 
						formulario: {
							titulo: "Nombre archivo salida",
							tipo: "texto",
							placeholder: "Introduzca la ruta del archivo de salida.",
							valorDefault: ""
						}
					}],
				opciones: null,
				salida: [{tipo: "boolean", valor: false}]
			},
			{

				nombre: "Eliminar duplicados spool",
				categoria: "Spool",
				tipo: "proceso",
				descripcion: "Elimina los registros de cuenta contrato duplicadas.",
				autor: "Carlos Cabrera",
				argumentos: [
					{
						tipo: "texto", 
						obligado: true, 
						identificador: "spoolInput", 
						formulario: {
							titulo: "Ruta Spool entrada",
							tipo: "archivo",
							accept: ".txt, .TXT",
							placeholder: "Introduzca la ruta del archivo de entrada.",
							valorDefault: ""
						}
					},
					{
						tipo: "texto", 
						obligado: true, 
						identificador: "spoolOutput", 
						formulario: {
							titulo: "Directorio de salida:",
							tipo: "ruta",
							placeholder: "Ruta de guardado...",
							valorDefault: ""
						}
					},
					{
						tipo: "texto", 
						obligado: true, 
						identificador: "nombreSpoolOutput", 
						formulario: {
							titulo: "Nombre archivo salida",
							tipo: "texto",
							placeholder: "Introduzca la ruta del archivo de salida.",
							valorDefault: ""
						}
					}],
				opciones: null,
				salida: [{tipo: "boolean", valor: false}]
			},
			{

				nombre: "Compensar Spool",
				categoria: "Spool",
				tipo: "proceso",
				descripcion: "Elimina las cabeceras y formatea las spools de SAP para su futuro procesamiento.",
				autor: "Carlos Cabrera",
				argumentos: [
					{
						tipo: "texto", 
						obligado: true, 
						identificador: "spoolInput", 
						formulario: {
							titulo: "Ruta Spool entrada",
							tipo: "archivo",
							accept: ".txt, .TXT",
							placeholder: "Introduzca la ruta del archivo de entrada.",
							valorDefault: ""
						}
					},
					{
						tipo: "texto", 
						obligado: true, 
						identificador: "spoolCompensadas", 
						formulario: {
							titulo: "Spool de compensadas",
							tipo: "archivo",
							accept: ".txt, .TXT",
							placeholder: "Introduzca la ruta del archivo de entrada.",
							valorDefault: ""
						}
					},
					{
						tipo: "texto", 
						obligado: true, 
						identificador: "spoolCompensadas2", 
						formulario: {
							titulo: "Spool de compensadas2",
							tipo: "archivo",
							accept: ".txt, .TXT",
							placeholder: "Introduzca la ruta del archivo de entrada.",
							valorDefault: ""
						}
					},
					{
						tipo: "texto", 
						obligado: true, 
						identificador: "spoolCompensadas3", 
						formulario: {
							titulo: "Spool de compensadas3",
							tipo: "archivo",
							accept: ".txt, .TXT",
							placeholder: "Introduzca la ruta del archivo de entrada.",
							valorDefault: ""
						}
					},
					{
						tipo: "texto", 
						obligado: true, 
						identificador: "spoolOutput", 
						formulario: {
							titulo: "Directorio de salida:",
							tipo: "ruta",
							placeholder: "Ruta de guardado...",
							valorDefault: ""
						}
					},
					{
						tipo: "texto", 
						obligado: true, 
						identificador: "nombreSpoolOutput", 
						formulario: {
							titulo: "Nombre archivo salida",
							tipo: "texto",
							placeholder: "Introduzca la ruta del archivo de salida.",
							valorDefault: ""
						}
					}],
				opciones: null,
				salida: [{tipo: "boolean", valor: false}]
			},

			{

				nombre: "Dividir archivo Spool",
				categoria: "Spool",
				tipo: "proceso",
				descripcion: "Divide el documento de Spool en diferentes archivos",
				autor: "Carlos Cabrera",
				argumentos: [
					{
						tipo: "texto", 
						obligado: true, 
						identificador: "spoolInput", 
						formulario: {
							titulo: "Ruta Spool entrada",
							tipo: "archivo",
							accept: ".txt, .TXT",
							placeholder: "Introduzca la ruta del archivo de entrada.",
							valorDefault: ""
						}
					},
					{
						tipo: "texto", 
						obligado: true, 
						identificador: "spoolOutput", 
						formulario: {
							titulo: "Directorio de salida:",
							tipo: "ruta",
							placeholder: "Ruta de guardado...",
							valorDefault: ""
						}
					},
					{
						tipo: "texto", 
						obligado: true, 
						identificador: "nombreSpoolOutput", 
						formulario: {
							titulo: "Nombre archivo salida",
							tipo: "texto",
							placeholder: "Introduzca la ruta del archivo de salida.",
							valorDefault: ""
						}
					}],
				opciones: null,
				salida: [{tipo: "boolean", valor: false}]
			},

			{
				nombre: "Spool to XLSX",
				categoria: "Spool",
				tipo: "proceso",
				descripcion: "Conivierte un archivo de texto generado por SAP en un archivo excel.",
				autor: "Carlos Cabrera",
				argumentos: [
					{
						tipo: "texto", 
						obligado: true, 
						identificador: "spoolInput", 
						formulario: {
							titulo: "Ruta Spool entrada",
							tipo: "archivo",
							accept: ".txt, .TXT",
							placeholder: "Introduzca la ruta del archivo de entrada.",
							valorDefault: ""
						}
					},
					{
						tipo: "texto", 
						obligado: true, 
						identificador: "spoolOutput", 
						formulario: {
							titulo: "Directorio de salida:",
							tipo: "ruta",
							placeholder: "Ruta de guardado...",
							valorDefault: ""
						}
					},
					{
						tipo: "texto", 
						obligado: true, 
						identificador: "nombreSpoolOutput", 
						formulario: {
							titulo: "Nombre archivo salida",
							tipo: "texto",
							placeholder: "Introduzca la ruta del archivo de salida.",
							valorDefault: ""
						}
					}],
				opciones: null,
				salida: [{tipo: "boolean", valor: false}]
			}
		]
	},
	{
		nombre: "Remedy",
		tipo: "directorio",
		categoria: "Remedy",
		descripcion: "Procesado de datos Remedy",
		subCategoria: [
			{
				nombre: "Extraccion Remedy",
				categoria: "Remedy",
				tipo: "proceso",
				descripcion: "Realiza una extracción de reportes de remedy",
				autor: "Carlos Cabrera",
				argumentos: [
					{
						tipo: "texto", 
						obligado: true, 
						identificador: "ejecutableChrome", 
						formulario: {
							titulo: "Google Chrome",
							tipo: "archivo",
							accept: ".exe",
							placeholder: "Ejecutable Google Chrome",
							valorDefault: ""
						}
				}],
				opciones: null,
				salida: [{tipo: "boolean", valor: false}]
			},
			{
				nombre: "Procesar Report Remedy",
				categoria: "Remedy",
				tipo: "proceso",
				descripcion: "Procesa un reporte de Remedy para incluir los datos en el excel de incidencias",
				autor: "Carlos Cabrera",
				argumentos: [
					{
						tipo: "objeto", 
						obligado: true, 
						identificador: "Report Remedy", 
						formulario: {
							titulo: "Objeto archivo Report Remedy",
							tipo: "objeto",
							accept: ".xlsx",
							placeholder: "Selecceione el objeto",
							valorDefault: ""
						}
					},
					{
						tipo: "texto", 
						obligado: true, 
						identificador: "incidenciasInput", 
						formulario: {
							titulo: "Archivo de monitorización de incidencias",
							tipo: "archivo",
							accept: ".xls, .XLS, .xlsm, .xlsx",
							placeholder: "Introduzca la ruta del archivo de monitorización de incidencias.",
							valorDefault: ""
						}
				}],
				opciones: null,
				salida: [{tipo: "boolean", valor: false}]
			},
			{
				nombre: "Procesar Extraccion Power BI",
				categoria: "Remedy",
				tipo: "proceso",
				descripcion: "Procesa un reporte de Remedy para incluir los datos en el excel de incidencias",
				autor: "Carlos Cabrera",
				argumentos: [
					{
						tipo: "texto", 
						obligado: true, 
						identificador: "reportPowerBI", 
						formulario: {
							titulo: "Extraccion Power BI",
							tipo: "archivo",
							accept: ".xls, .XLS, .xlsm, .xlsx, .xlsb",
							placeholder: "Introduzca la extraccion de Remedy",
							valorDefault: ""
						}
					},
					{
						tipo: "texto", 
						obligado: true, 
						identificador: "incidenciasInput", 
						formulario: {
							titulo: "Archivo de monitorización de incidencias",
							tipo: "archivo",
							accept: ".xls, .XLS, .xlsm, .xlsx",
							placeholder: "Introduzca la ruta del archivo de monitorización de incidencias.",
							valorDefault: ""
						}
				}],
				opciones: null,
				salida: [{tipo: "boolean", valor: false}]
			},
			{
				nombre: "Renderizar Report Remedy",
				categoria: "Remedy",
				tipo: "proceso",
				descripcion: "Transforma los datos historicos del informe de incidencias para su visualización",
				autor: "Carlos Cabrera",
				argumentos: [
					{
						tipo: "texto", 
						obligado: true, 
						identificador: "InformeIncidencias", 
						formulario: {
							titulo: "Archivo de monitorización de incidencias",
							tipo: "archivo",
							accept: ".xls, .XLS, .xlsm, .xlsx",
							placeholder: "Introduzca la ruta del archivo de monitorización de incidencias.",
							valorDefault: ""
						}
				}],
				opciones: null,
				salida: [{tipo: "boolean", valor: false}]
			},
			{
				nombre: "Renderizar Backlog",
				categoria: "Remedy",
				tipo: "proceso",
				descripcion: "Actualiza la tabla de registro de backlog del registro de incidencias.",
				autor: "Carlos Cabrera",
				argumentos: [
					{
						tipo: "texto", 
						obligado: true, 
						identificador: "InformeIncidencias", 
						formulario: {
							titulo: "Archivo de monitorización de incidencias",
							tipo: "archivo",
							accept: ".xls, .XLS, .xlsm, .xlsx",
							placeholder: "Introduzca la ruta del archivo de monitorización de incidencias.",
							valorDefault: ""
						}
				}],
				opciones: null,
				salida: [{tipo: "boolean", valor: false}]
			}
		]
	},
	{
		nombre: "Desarrollador",
		categoria: "Desarrollador",
		tipo: "directorio",
		descripcion: "Procesos de en fase de desarrollo y testeo.",
		subCategoria: [
			{

				nombre: "Unir Carpeta Excel",
				categoria: "Desarrollador",
				tipo: "proceso",
				descripcion: "Fusiona todos los archivos excel de una carpeta en un único archivo",
				autor: "Carlos Cabrera",
				argumentos: [
					{
						tipo: "texto", 
						obligado: true, 
						identificador: "excelInput", 
						formulario: {
							titulo: "Carpeta con excels para fusionar",
							tipo: "ruta",
							placeholder: "Carpeta de archivos Excel...",
							valorDefault: ""
						}
					},
					{
						tipo: "texto", 
						obligado: true, 
						identificador: "excelOutput", 
						formulario: {
							titulo: "Directorio de salida:",
							tipo: "ruta",
							placeholder: "Ruta de guardado...",
							valorDefault: ""
						}
					},
					{
						tipo: "texto", 
						obligado: true, 
						identificador: "nombreExcelOutput", 
						formulario: {
							titulo: "Nombre archivo salida",
							tipo: "texto",
							placeholder: "Nombre del archivo de salida.",
							valorDefault: ""
						}
					}],
				opciones: null,
				salida: [{tipo: "boolean", valor: false}]
			},
			{
				nombre: "Prueba Campos",
				tipo: "proceso",
				categoria: "Desarrollador",
				descripcion: "Divide los datos del Spool en multiples documentos",
				autor: "Carlos Cabrera",
				argumentos: [
					{
						tipo: "texto", 
						obligado: true, 
						identificador: "spoolInput", 
						formulario: {
							titulo: "Formulario seleccion: ",
							tipo: "seleccion",
							placeholder: "Seleccionar...",
							valorDefault: ["Texto","Número"] 
						}
					},
					{
						tipo: "boolean", 
						obligado: true, 
						identificador: "spoolInput", 
						formulario: {
							titulo: "Formulario Boolean:",
							tipo: "boolean",
							placeholder: "Automatico: ",
							valorDefault: ["Texto","Número"] 
						}
					},
					{
						tipo: "numero", 
						obligado: true, 
						identificador: "spoolInput", 
						formulario: {
							titulo: "Formulario número: ",
							tipo: "numero",
							placeholder: "Número de iteraciones. ",
							valorDefault: 1 
						}
					},
					{
						tipo: "fecha", 
						obligado: true, 
						identificador: "spoolInput", 
						formulario: {
							titulo: "Formulario Fecha:",
							tipo: "fecha",
							placeholder: "Introduzca la fecha... ",
							valorDefault: "" 
						}
					},
					{
						tipo: "texto", 
						obligado: true, 
						identificador: "spoolInput", 
						formulario: {
							titulo: "Formulario Texto: ",
							tipo: "texto",
							placeholder: "Nombre",
							valorDefault: "Texto" 
						}
					},
					{
						tipo: "ruta", 
						obligado: true, 
						identificador: "spoolOutput", 
						formulario: {
							titulo: "Formulario Ruta: ",
							tipo: "ruta",
							placeholder: "Introduzca la ruta del archivo de salida.",
							valorDefault: ""
						}
					}],
				opciones: null,
				salida: [{tipo: "boolean", valor: false}]
			},
			{
				nombre: "Ir a ReportAM",
				tipo: "redireccion",
				categoria: "Desarrollador",
				descripcion: "Cambia la interfaz al antiguo sistema ReportAM",
				autor: "Carlos Cabrera",
				argumentos: [],
				opciones: null,
				salida: [{tipo: "texto", valor: "reportAM"}]
			},
			{
				nombre: "Generar Seguimiento AM",
				categoria: "Desarrollador",
				tipo: "proceso",
				descripcion: "Genera un archivo de seguimiento actualizado a partir de un archivo de seguimiento antiguo y los datos de nacho",
				autor: "Carlos Cabrera",
				argumentos: [
					{
						tipo: "texto", 
						obligado: true, 
						identificador: "Seguimiento Base", 
						formulario: {
							titulo: "Ruta seguimiento base",
							tipo: "archivo",
							accept: ".xlsx",
							placeholder: "Introduzca la ruta del archivo",
							valorDefault: ""
						}
					},
					{
						tipo: "objeto", 
						obligado: true, 
						identificador: "Archivo Nacho", 
						formulario: {
							titulo: "Objeto archivo nacho",
							tipo: "objeto",
							accept: ".xlsx",
							placeholder: "Selecceione el objeto",
							valorDefault: ""
						}
					},
					{
						tipo: "numero", 
						obligado: true, 
						identificador: "añoInicio", 
						formulario: {
							titulo: "Año de inicio analisis",
							tipo: "numero",
							placeholder: "Año de inicio:",
							valorDefault: 2020
						}
					},
					{
						tipo: "numero", 
						obligado: true, 
						identificador: "mesInicio", 
						formulario: {
							titulo: "Mes de inicio analisis",
							tipo: "numero",
							placeholder: "Mes de inicio:",
							valorDefault: 9  
						}
					},
					{
						tipo: "numero", 
						obligado: true, 
						identificador: "añoFin", 
						formulario: {
							titulo: "Año fin de analisis",
							tipo: "numero",
							placeholder: "Año fin de analisis:",
							valorDefault: 2020
						}
					},
					{
						tipo: "numero", 
						obligado: true, 
						identificador: "mesFin", 
						formulario: {
							titulo: "Mes fin de analisis",
							tipo: "numero",
							placeholder: "Mes fin de analisis:",
							valorDefault: 10  
						}
					},
					{
						tipo: "texto", 
						obligado: true, 
						identificador: "spoolOutput", 
						formulario: {
							titulo: "Directorio de salida:",
							tipo: "ruta",
							placeholder: "Ruta de guardado...",
							valorDefault: ""
						}
					},
					{
						tipo: "texto", 
						obligado: true, 
						identificador: "nombreSpoolOutput", 
						formulario: {
							titulo: "Nombre archivo salida",
							tipo: "texto",
							placeholder: "Introduzca la ruta del archivo de salida.",
							valorDefault: ""
						}
					}],
				opciones: null,
				salida: [{tipo: "boolean", valor: false}]
			},
			{
				nombre: "Generar Seguimiento AM 2.0",
				categoria: "Desarrollador",
				tipo: "proceso",
				descripcion: "Genera un archivo de seguimiento actualizado a partir de un archivo power bi",
				autor: "Carlos Cabrera",
				argumentos: [
					{
						tipo: "texto", 
						obligado: true, 
						identificador: "Seguimiento Base", 
						formulario: {
							titulo: "Ruta seguimiento base",
							tipo: "archivo",
							accept: ".xlsx",
							placeholder: "Introduzca la ruta del archivo",
							valorDefault: ""
						}
					},
					{
						tipo: "objeto", 
						obligado: true, 
						identificador: "Archivo Nacho", 
						formulario: {
							titulo: "Objeto archivo nacho",
							tipo: "objeto",
							accept: ".xlsx",
							placeholder: "Selecceione el objeto",
							valorDefault: ""
						}
					},
					{
						tipo: "numero", 
						obligado: true, 
						identificador: "añoInicio", 
						formulario: {
							titulo: "Año de inicio analisis",
							tipo: "numero",
							placeholder: "Año de inicio:",
							valorDefault: 2020
						}
					},
					{
						tipo: "numero", 
						obligado: true, 
						identificador: "mesInicio", 
						formulario: {
							titulo: "Mes de inicio analisis",
							tipo: "numero",
							placeholder: "Mes de inicio:",
							valorDefault: 9  
						}
					},
					{
						tipo: "numero", 
						obligado: true, 
						identificador: "añoFin", 
						formulario: {
							titulo: "Año fin de analisis",
							tipo: "numero",
							placeholder: "Año fin de analisis:",
							valorDefault: 2020
						}
					},
					{
						tipo: "numero", 
						obligado: true, 
						identificador: "mesFin", 
						formulario: {
							titulo: "Mes fin de analisis",
							tipo: "numero",
							placeholder: "Mes fin de analisis:",
							valorDefault: 10  
						}
					},
					{
						tipo: "texto", 
						obligado: true, 
						identificador: "spoolOutput", 
						formulario: {
							titulo: "Directorio de salida:",
							tipo: "ruta",
							placeholder: "Ruta de guardado...",
							valorDefault: ""
						}
					},
					{
						tipo: "texto", 
						obligado: true, 
						identificador: "nombreSpoolOutput", 
						formulario: {
							titulo: "Nombre archivo salida",
							tipo: "texto",
							placeholder: "Introduzca la ruta del archivo de salida.",
							valorDefault: ""
						}
					}],
				opciones: null,
				salida: [{tipo: "boolean", valor: false}]
			},
			{
				nombre: "Fusionar Objetos",
				categoria: "Desarrollador",
				tipo: "proceso",
				descripcion: "Añade el contenido de un objeto dentro de otro.",
				autor: "Carlos Cabrera",
				argumentos: [
					{
						tipo: "objeto", 
						obligado: true, 
						identificador: "Objeto Base", 
						formulario: {
							titulo: "Objeto Base",
							tipo: "objeto",
							accept: ".xlsx",
							placeholder: "Selecceione el objeto",
							valorDefault: ""
						}
					},
					{
						tipo: "objeto", 
						obligado: true, 
						identificador: "Objeto añadir", 
						formulario: {
							titulo: "Objeto añadir",
							tipo: "objeto",
							accept: ".xlsx",
							placeholder: "Objeto añadir",
							valorDefault: ""
						}
					}
					],
				opciones: null,
				salida: [{tipo: "boolean", valor: false}]
			},
			{
				nombre: "Procesar IBAN",
				categoria: "Desarrollador",
				tipo: "proceso",
				descripcion: "Procesa el recuento IBAN-Mandato",
				autor: "Carlos Cabrera",
				argumentos: [
					{
						tipo: "objeto", 
						obligado: true, 
						identificador: "Objeto Datos 1", 
						formulario: {
							titulo: "Objeto Datos",
							tipo: "objeto",
							accept: ".xlsx",
							placeholder: "Selecceione el objeto",
							valorDefault: ""
						}
					},
					{
						tipo: "objeto", 
						obligado: true, 
						identificador: "Objeto Datos 2", 
						formulario: {
							titulo: "Objeto Datos",
							tipo: "objeto",
							accept: ".xlsx",
							placeholder: "Selecceione el objeto",
							valorDefault: ""
						}
					},
					{
						tipo: "objeto", 
						obligado: true, 
						identificador: "Objeto Datos 3", 
						formulario: {
							titulo: "Objeto Datos",
							tipo: "objeto",
							accept: ".xlsx",
							placeholder: "Selecceione el objeto",
							valorDefault: ""
						}
					},
					{
						tipo: "objeto", 
						obligado: true, 
						identificador: "Objeto Datos 4", 
						formulario: {
							titulo: "Objeto Datos",
							tipo: "objeto",
							accept: ".xlsx",
							placeholder: "Selecceione el objeto",
							valorDefault: ""
						}
					},
					{
						tipo: "objeto", 
						obligado: true, 
						identificador: "Objeto Datos 5", 
						formulario: {
							titulo: "Objeto Datos",
							tipo: "objeto",
							accept: ".xlsx",
							placeholder: "Selecceione el objeto",
							valorDefault: ""
						}
					},
					{
						tipo: "objeto", 
						obligado: true, 
						identificador: "Objeto Datos 6", 
						formulario: {
							titulo: "Objeto Datos",
							tipo: "objeto",
							accept: ".xlsx",
							placeholder: "Selecceione el objeto",
							valorDefault: ""
						}
					},
					{
						tipo: "objeto", 
						obligado: true, 
						identificador: "Objeto Datos 7", 
						formulario: {
							titulo: "Objeto Datos",
							tipo: "objeto",
							accept: ".xlsx",
							placeholder: "Selecceione el objeto",
							valorDefault: ""
						}
					},
					{
						tipo: "objeto", 
						obligado: true, 
						identificador: "Objeto Datos 8", 
						formulario: {
							titulo: "Objeto Datos",
							tipo: "objeto",
							accept: ".xlsx",
							placeholder: "Selecceione el objeto",
							valorDefault: ""
						}
					},
					{
						tipo: "objeto", 
						obligado: true, 
						identificador: "Objeto Datos 9", 
						formulario: {
							titulo: "Objeto Datos",
							tipo: "objeto",
							accept: ".xlsx",
							placeholder: "Selecceione el objeto",
							valorDefault: ""
						}
					},
					{
						tipo: "objeto", 
						obligado: true, 
						identificador: "Objeto Datos 10", 
						formulario: {
							titulo: "Objeto Datos",
							tipo: "objeto",
							accept: ".xlsx",
							placeholder: "Selecceione el objeto",
							valorDefault: ""
						}
					},
					{
						tipo: "objeto", 
						obligado: true, 
						identificador: "Objeto Datos 11", 
						formulario: {
							titulo: "Objeto Datos",
							tipo: "objeto",
							accept: ".xlsx",
							placeholder: "Selecceione el objeto",
							valorDefault: ""
						}
					},
					{
						tipo: "objeto", 
						obligado: true, 
						identificador: "Objeto Datos 12", 
						formulario: {
							titulo: "Objeto Datos",
							tipo: "objeto",
							accept: ".xlsx",
							placeholder: "Selecceione el objeto",
							valorDefault: ""
						}
					},
					{
						tipo: "objeto", 
						obligado: true, 
						identificador: "Objeto Datos 13", 
						formulario: {
							titulo: "Objeto Datos",
							tipo: "objeto",
							accept: ".xlsx",
							placeholder: "Selecceione el objeto",
							valorDefault: ""
						}
					},
					{
						tipo: "objeto", 
						obligado: true, 
						identificador: "Objeto Datos 14", 
						formulario: {
							titulo: "Objeto Datos",
							tipo: "objeto",
							accept: ".xlsx",
							placeholder: "Selecceione el objeto",
							valorDefault: ""
						}
					},
					{
						tipo: "objeto", 
						obligado: true, 
						identificador: "Objeto Datos 15", 
						formulario: {
							titulo: "Objeto Datos",
							tipo: "objeto",
							accept: ".xlsx",
							placeholder: "Selecceione el objeto",
							valorDefault: ""
						}
					},
					{
						tipo: "texto", 
						obligado: true, 
						identificador: "rutaGuardado", 
						formulario: {
							titulo: "Directorio de salida:",
							tipo: "ruta",
							placeholder: "Ruta de guardado...",
							valorDefault: ""
						}
					},
					{
						tipo: "texto", 
						obligado: true, 
						identificador: "nombreGuardado", 
						formulario: {
							titulo: "Nombre archivo salida",
							tipo: "texto",
							placeholder: "Nombre del archivo",
							valorDefault: ""
						}
					}
					],
				opciones: null,
				salida: [{tipo: "boolean", valor: false}]
			},
			{
				nombre: "Procesar SMS",
				categoria: "Desarrollador",
				tipo: "proceso",
				descripcion: "Realiza un recuento de la Spool de SMS",
				autor: "Carlos Cabrera",
				argumentos: [
					{
						tipo: "texto", 
						obligado: true, 
						identificador: "spoolInput", 
						formulario: {
							titulo: "Ruta Spool entrada",
							tipo: "archivo",
							accept: ".txt, .TXT",
							placeholder: "Introduzca la ruta del archivo de entrada.",
							valorDefault: ""
						}
					},
					{
						tipo: "texto", 
						obligado: true, 
						identificador: "spoolOutput", 
						formulario: {
							titulo: "Directorio de salida:",
							tipo: "ruta",
							placeholder: "Ruta de guardado...",
							valorDefault: ""
						}
					},
					{
						tipo: "texto", 
						obligado: true, 
						identificador: "nombreSpoolOutput", 
						formulario: {
							titulo: "Nombre archivo salida",
							tipo: "texto",
							placeholder: "Introduzca la ruta del archivo de salida.",
							valorDefault: ""
						}
					}],
				opciones: null,
				salida: [{tipo: "boolean", valor: false}]
			}
		]
	},
	
	{
		nombre: "Google",
		categoria: "Google",
		tipo: "directorio",
		descripcion: "Procesos de en fase de desarrollo y testeo.",
		subCategoria: [
			{
				nombre: "Validacion Gmail",
				categoria: "Google",
				tipo: "proceso",
				descripcion: "Realiza la autentificacion de google",
				autor: "Carlos Cabrera",
				argumentos: [],
				opciones: null,
				salida: [{tipo: "boolean", valor: false}]
			},
			{
				nombre: "Obtener Correos",
				categoria: "Google",
				tipo: "proceso",
				descripcion: "Obtiene el listado de correos",
				autor: "Carlos Cabrera",
				argumentos: [
					{
						tipo: "texto", 
						obligado: true, 
						identificador: "querry", 
						formulario: {
							titulo: "Querry de consulta de correo",
							tipo: "texto",
							placeholder: "Querry",
							valorDefault: ""
						}
					}
					],
				opciones: null,
				salida: [{tipo: "boolean", valor: false}]
			}
		]
	},
	{
		nombre: "KPIs",
		categoria: "KPIs",
		tipo: "directorio",
		descripcion: "Procesado de KPIs",
		subCategoria: [
			{
				nombre: "Facturacion Ciclo Step 1",
				categoria: "KPIs",
				tipo: "proceso",
				descripcion: "Primer paso de procesamiento del kpi de facturación",
				autor: "Carlos Cabrera",
				argumentos: [
					{
						tipo: "texto", 
						obligado: true, 
						identificador: "analisisCiclo", 
						formulario: {
							titulo: "Archivo analisis ciclo",
							tipo: "archivo",
							accept: ".xls, .XLS, .xlsm, .xlsx",
							placeholder: "Selecciones el archivo Excel para importar.",
							valorDefault: ""
						}
				},
					{
						tipo: "texto", 
						obligado: true, 
						identificador: "seguimientoKPIs", 
						formulario: {
							titulo: "Archivo analisis ciclo",
							tipo: "archivo",
							accept: ".xls, .XLS, .xlsm, .xlsx",
							placeholder: "Selecciones el archivo Excel para importar.",
							valorDefault: ""
						}
				}
					],
				opciones: null,
				salida: [{tipo: "boolean", valor: false}]
			},
			{
				nombre: "Facturacion Ciclo Step 2",
				categoria: "KPIs",
				tipo: "proceso",
				descripcion: "Segundo paso de procesamiento del kpi de facturación",
				autor: "Carlos Cabrera",
				argumentos: [
					{
						tipo: "texto", 
						obligado: true, 
						identificador: "analisisCiclo", 
						formulario: {
							titulo: "Archivo analisis ciclo",
							tipo: "archivo",
							accept: ".xls, .XLS, .xlsm, .xlsx",
							placeholder: "Selecciones el archivo Excel para importar.",
							valorDefault: ""
						}
				},
					{
						tipo: "texto", 
						obligado: true, 
						identificador: "seguimientoKPIs", 
						formulario: {
							titulo: "Archivo de seguimiento de KPIs",
							tipo: "archivo",
							accept: ".xls, .XLS, .xlsm, .xlsx",
							placeholder: "Selecciones de seguimiento de KPIs.",
							valorDefault: ""
						}
				}
					],
				opciones: null,
				salida: [{tipo: "boolean", valor: false}]
			},
			{
				nombre: "Facturacion Hotbilling Step 1",
				categoria: "KPIs",
				tipo: "proceso",
				descripcion: "Procesado Hotbilling",
				autor: "Carlos Cabrera",
				argumentos: [
					{
						tipo: "texto", 
						obligado: true, 
						identificador: "analisisHotbilling", 
						formulario: {
							titulo: "Archivo analisis Hotbilling",
							tipo: "archivo",
							accept: ".xls, .XLS, .xlsm, .xlsx",
							placeholder: "Selecciones el archivo Excel para importar.",
							valorDefault: ""
						}
				},
					{
						tipo: "texto", 
						obligado: true, 
						identificador: "seguimientoKPIs", 
						formulario: {
							titulo: "Archivo de seguimiento de KPIs",
							tipo: "archivo",
							accept: ".xls, .XLS, .xlsm, .xlsx",
							placeholder: "",
							valorDefault: "Selecciones de seguimiento de KPIs."
						}
				}
					],
				opciones: null,
				salida: [{tipo: "boolean", valor: false}]
			},
			{
				nombre: "Facturacion Hotbilling Step 2",
				categoria: "KPIs",
				tipo: "proceso",
				descripcion: "Procesado Hotbilling",
				autor: "Carlos Cabrera",
				argumentos: [
					{
						tipo: "texto", 
						obligado: true, 
						identificador: "analisisHotbilling", 
						formulario: {
							titulo: "Archivo analisis Hotbilling",
							tipo: "archivo",
							accept: ".xls, .XLS, .xlsm, .xlsx",
							placeholder: "Selecciones el archivo Excel para importar.",
							valorDefault: ""
						}
				},
					{
						tipo: "texto", 
						obligado: true, 
						identificador: "seguimientoKPIs", 
						formulario: {
							titulo: "Archivo de seguimiento de KPIs",
							tipo: "archivo",
							accept: ".xls, .XLS, .xlsm, .xlsx",
							placeholder: "",
							valorDefault: "Selecciones de seguimiento de KPIs."
						}
				}
					],
				opciones: null,
				salida: [{tipo: "boolean", valor: false}]
			},
			{
				nombre: "Financiaciones",
				categoria: "KPIs",
				tipo: "proceso",
				descripcion: "Procesado Financiaciones",
				autor: "Carlos Cabrera",
				argumentos: [
					{
						tipo: "texto", 
						obligado: true, 
						identificador: "analisisFinanciaciones", 
						formulario: {
							titulo: "Archivo analisis Financiaciones",
							tipo: "archivo",
							accept: ".xls, .XLS, .xlsm, .xlsx",
							placeholder: "Selecciones el archivo Excel para importar.",
							valorDefault: ""
						}
				},
					{
						tipo: "texto", 
						obligado: true, 
						identificador: "analisisFinanciacionesW-1", 
						formulario: {
							titulo: "Financiaciones W-1",
							tipo: "archivo",
							accept: ".xls, .XLS, .xlsm, .xlsx",
							placeholder: "Selecciones el archivo Excel para importar.",
							valorDefault: ""
						}
				},
					{
						tipo: "texto", 
						obligado: true, 
						identificador: "seguimientoKPIs", 
						formulario: {
							titulo: "Archivo de seguimiento de KPIs",
							tipo: "archivo",
							accept: ".xls, .XLS, .xlsm, .xlsx",
							placeholder: "",
							valorDefault: "Selecciones de seguimiento de KPIs."
						}
				}
					],
				opciones: null,
				salida: [{tipo: "boolean", valor: false}]
			}
		]
	},
	{
		nombre: "Import",
		categoria: "Import",
		tipo: "directorio",
		descripcion: "Porcesos de importación de datos:",
		subCategoria: [
			{
				nombre: "Importar Excel",
				categoria: "Import",
				tipo: "proceso",
				descripcion: "Importa un archivo Excel",
				autor: "Carlos Cabrera",
				argumentos: [
					{
						tipo: "texto", 
						obligado: true, 
						identificador: "excelImportado", 
						formulario: {
							titulo: "Archivo Excel",
							tipo: "archivo",
							accept: ".xls, .XLS, .xlsm, .xlsx",
							placeholder: "Selecciones el archivo Excel para importar.",
							valorDefault: ""
						}
				},
					{
						tipo: "texto", 
						obligado: true, 
						identificador: "numeroFilaCabecera", 
						formulario: {
							titulo: "Número fila cabecera",
							tipo: "numero",
							placeholder: "Numero Fila Cabecera",
							valorDefault: 1 
						}
					},
					{
						tipo: "texto", 
						obligado: true, 
						identificador: "nombreHoja", 
						formulario: {
							titulo: "Nombre hoja a importar",
							tipo: "texto",
							placeholder: "Nombre Hoja",
							valorDefault: ""
						}
					},
					{
						tipo: "texto", 
						obligado: true, 
						identificador: "nombreObjeto", 
						formulario: {
							titulo: "Nombre del objeto guardado",
							tipo: "texto",
							placeholder: "Nombre Objeto",
							valorDefault: ""
						}
					}
					],
				opciones: null,
				salida: [{tipo: "boolean", valor: false}]
			}
		]
	},
	{
		nombre: "Subir Cursos",
		categoria: "Santander",
		tipo: "directorio",
		descripcion: "Subida de datos de monitorización de cursos",
		subCategoria: [
			{
				nombre: "Subir Monitorización Cursos",
				categoria: "Desarrollador",
				tipo: "proceso",
				descripcion: "Actualizar Archivo monitorización de cursos",
				autor: "Carlos Cabrera",
				argumentos: [
					],
				opciones: null,
				salida: [{tipo: "boolean", valor: false}]
			},
			{
				nombre: "Generar Documento",
				categoria: "Documentos",
				tipo: "proceso",
				descripcion: "Genera un documento mediante una plantilla",
				autor: "Carlos Cabrera",
				argumentos: [
					],
				opciones: null,
				salida: [{tipo: "boolean", valor: false}]
			}
		]
	},
	{
		nombre: "Despacho",
		categoria: "Despacho",
		tipo: "directorio",
		descripcion: "Procesos de despacho",
		subCategoria: [
			{
				nombre: "Correo a Infolex",
				categoria: "Despacho",
				tipo: "proceso",
				descripcion: "Obtiene el listado de correos y realiza el proceso de importación a infolex",
				autor: "Carlos Cabrera",
				argumentos: [
					{
						tipo: "texto", 
						obligado: true, 
						identificador: "querry", 
						formulario: {
							titulo: "Querry de consulta de correo",
							tipo: "texto",
							placeholder: "Querry",
							valorDefault: ""
						}
					}
					],
				opciones: null,
				salida: [{tipo: "boolean", valor: false}]
			},
            {
				nombre: "Get validacion Google",
				categoria: "Despacho",
				tipo: "proceso",
				descripcion: "Obtiene la validacion de Google",
				autor: "Carlos Cabrera",
				argumentos: [],
				opciones: null,
				salida: [{tipo: "boolean", valor: false}]
			}
		]
	}
]



export {LibreriaProcesos,libreriaProcesos};
