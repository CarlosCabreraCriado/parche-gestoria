// Carga de utilidades de rutas del sistema.
// Motivo: formar rutas absolutas y evitar errores entre Windows, macOS y Linux.
const path = require("path");

// Carga de utilidades de acceso a archivos.
// Motivo: comprobar que el archivo existe antes de abrirlo.
const fs = require("fs");

// Carga de la librería para trabajar con ficheros Excel.
// Motivo: abrir el libro, leer celdas y recorrer datos.
const XlsxPopulate = require("xlsx-populate");

// Definición de la clase que agrupa los procesos de prueba.
// Motivo: organizar las funcionalidades del módulo de forma clara.
class ProcesosPrueba {
  // Método que lee un Excel con dos columnas ("nombre" y "nombre_formateado")
  // y muestra por consola el primer nombre, el tamaño del rango usado
  // y el listado completo de nombres.
  async testExcel(argumentos) {
    try {
      // Mensaje inicial para facilitar el seguimiento en consola.
      console.log("Inicio del proceso de lectura de Excel");

      // Obtención de la ruta del archivo elegida por el usuario en la interfaz.
      // Motivo: usar exactamente el archivo que se ha seleccionado.
      // La ? Sirve para acceder a propiedades de forma segura cuando algo puede ser null o undefined, 
      // evitando el típico error TypeError: Cannot read properties of undefined
      const rutaFormulario = argumentos?.formularioControl?.[0]; 

      // Carpeta de salida donde se guardarán los resultados.
      // Motivo: permitir al usuario elegir el destino de los archivos generados.
      const carpetaSalida = argumentos?.formularioControl?.[1];

      // Validaciones básicas de entrada.
      // Motivo: evitar errores al trabajar con rutas incorrectas.
      if (typeof rutaFormulario !== "string" || !rutaFormulario.trim()) {
        console.error("No se ha proporcionado una ruta de archivo válida.");
        return false;
      }
      if (typeof carpetaSalida !== "string" || !carpetaSalida.trim()) {
        console.error("No se ha proporcionado una carpeta de salida válida.");
        return false;
      }

      // Normalización a rutas absolutas.
      // Motivo: garantizar que el sistema localiza los recursos sin ambigüedades.
      const rutaExcel = path.isAbsolute(rutaFormulario) ? rutaFormulario : path.resolve(rutaFormulario);
      const salidaDir = path.isAbsolute(carpetaSalida) ? carpetaSalida : path.resolve(carpetaSalida);

      // Comprobación de existencia del archivo y preparación de la carpeta de salida.
      // Motivo: asegurar que hay fuente y destino disponibles.
      if (!fs.existsSync(rutaExcel)) {
        console.error("El archivo indicado no existe:", rutaExcel);
        return false;
      }
      if (!fs.existsSync(salidaDir)) {
        fs.mkdirSync(salidaDir, { recursive: true });
      }


      // Apertura del libro de Excel de forma asíncrona.
      // Motivo: no bloquear la aplicación durante la lectura del archivo.
      const workbook = await XlsxPopulate.fromFileAsync(rutaExcel);
      console.log("Confirmación: el archivo de Excel se ha abierto correctamente.");

      // Nombres de salida (manteniendo el nombre base del archivo fuente).
      // Motivo: generar archivos reconocibles en la carpeta de salida.
      const baseName = path.basename(rutaExcel, path.extname(rutaExcel));
      const rutaProcesado = path.normalize(path.join(salidaDir, `${baseName}_PROCESADO.xlsx`));


      // Selección de la primera hoja del libro.
      // Motivo: el archivo es sencillo y toda la información está en la primera hoja.
      const hoja = workbook.sheet(0);
      console.log("Confirmación: la hoja de trabajo se ha seleccionado correctamente.");

      // Cálculo del rango usado de la hoja (área que contiene datos).
      // Motivo: conocer cuántas filas y columnas tienen información.
      const rangoUsado = hoja.usedRange();
      const datos = rangoUsado.value();

      // Número total de filas del rango (incluye cabecera).
      // Motivo: informar del volumen de información disponible.
      const totalFilas = Array.isArray(datos) ? datos.length : 0;

      // Número total de columnas del rango.
      // Motivo: verificar que existen las columnas esperadas.
      const totalColumnas = totalFilas > 0 && Array.isArray(datos[0]) ? datos[0].length : 0;

      // Información de tamaño del rango.
      // Motivo: facilitar la revisión por consola.
      console.log("Filas (incluye cabecera):", totalFilas);
      console.log("Columnas:", totalColumnas);
      console.log("Confirmación: el análisis de filas y columnas se ha completado correctamente.");

      // Asegura cabecera de la columna B si estuviera vacía.
      // Motivo: claridad en el resultado final.
      const cabeceraB = String(hoja.cell("B1").value() ?? "").trim();
      if (!cabeceraB) {
        hoja.cell("B1").value("nombre_formateado");
      }

      // Listado completo de la columna "nombre" (A) y escritura en B en MAYÚSCULAS.
      // Motivo: mostrar por consola y generar la columna formateada.
      if (totalFilas < 2) {
        console.log("No se han encontrado nombres para listar.");
      } else {
        for (let fila = 2; fila <= totalFilas; fila++) {
          const celdaA = `A${fila}`;                 // Referencia a la celda de origen (columna A).
          const valor = hoja.cell(celdaA).value();   // Lectura del valor de la celda.

          // Evita filas vacías.
          if (valor !== null && valor !== undefined && String(valor).trim() !== "") {

            // Convierte a texto y a MAYÚSCULAS.
            // Motivo: cumplir el requisito de salida en mayúsculas.
            const textoMayus = String(valor).trim().toUpperCase();

            // Escribe el resultado en la columna B de la misma fila.
            // Motivo: rellenar la columna "nombre_formateado".
            const celdaB = `B${fila}`;
            hoja.cell(celdaB).value(textoMayus);
          }
        }
        console.log("Confirmación: el listado y la escritura de la columna formateada se han realizado correctamente.");
      }

      // Guarda el libro ya modificado con sufijo PROCESADO.
      // Motivo: devolver un archivo con la segunda columna cumplimentada.
      await workbook.toFileAsync(rutaProcesado);
      console.log("Archivo procesado guardado en:", rutaProcesado);

      // Mensaje de fin para indicar que todo el flujo ha concluido con éxito.
      console.log("Proceso finalizado correctamente.");
      return true;
    } catch (error) {
      // Mensaje claro en caso de incidencia durante la ejecución.
      console.error("Incidencia durante la lectura o escritura del Excel:", error);
      return false;
    }
  }
}
// Exportación de la clase o de una instancia según la arquitectura del proyecto.
// Motivo: permitir su utilización desde el resto de la aplicación.
// Si en tu proyecto ya existe la clase y su exportación, conserva ese patrón.
module.exports = ProcesosPrueba;
