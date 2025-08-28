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

      // Validación de que la ruta es un texto no vacío.
      // Motivo: evitar errores al trabajar con rutas incorrectas.
      if (typeof rutaFormulario !== "string" || !rutaFormulario.trim()) {
        console.error("No se ha proporcionado una ruta de archivo válida.");
        return false;
      }

      // Normalización a ruta absoluta.
      // Motivo: garantizar que el sistema localiza el archivo sin ambigüedades.
      const rutaExcel = path.isAbsolute(rutaFormulario) ? rutaFormulario : path.resolve(rutaFormulario);

      // Comprobación de existencia del archivo en disco.
      // Motivo: informar de forma clara si el archivo no está disponible.
      if (!fs.existsSync(rutaExcel)) {
        console.error("El archivo indicado no existe:", rutaExcel);
        return false;
      }

      // Apertura del libro de Excel de forma asíncrona.
      // Motivo: no bloquear la aplicación durante la lectura del archivo.
      const workbook = await XlsxPopulate.fromFileAsync(rutaExcel);
      console.log("Confirmación: el archivo de Excel se ha abierto correctamente.");

      // Selección de la primera hoja del libro.
      // Motivo: el archivo es sencillo y toda la información está en la primera hoja.
      const hoja = workbook.sheet(0);
      console.log("Confirmación: la hoja de trabajo se ha seleccionado correctamente.");

      // Lectura de la celda A2 (A1 es cabecera).
      // Motivo: obtener el primer registro real de la columna "nombre".
      const primerNombre = hoja.cell("A2").value();
      console.log("Primer nombre (A2):", primerNombre ?? "—");
      console.log("Confirmación: la lectura del primer nombre se ha realizado correctamente.");

      // Cálculo del rango usado de la hoja (área que contiene datos).
      // Motivo: conocer cuántas filas y columnas tienen información.
      const rangoUsado = hoja.usedRange();
      const datos = rangoUsado.value();

      // Número total de filas del rango (incluye cabecera).
      // Motivo: informar del volumen de información disponible.
      const totalFilas = Array.isArray(datos) ? datos.length : 0;

      // Número total de columnas del rango.
      // Motivo: verificar que existen las columnas esperadas.
      const totalColumnas =
        totalFilas > 0 && Array.isArray(datos[0]) ? datos[0].length : 0;

      // Información de tamaño del rango.
      // Motivo: facilitar la revisión por consola.
      console.log("Filas (incluye cabecera):", totalFilas);
      console.log("Columnas:", totalColumnas);
      console.log("Confirmación: el análisis de filas y columnas se ha completado correctamente.");

      // Listado completo de la columna "nombre" (columna A) desde la fila 2 hasta la última.
      // Motivo: mostrar por consola todos los nombres existentes en el documento.
      if (totalFilas < 2) {
        // Si no hay filas de datos (solo cabecera o vacío), se informa y se finaliza.
        console.log("No se han encontrado nombres para listar.");
      } else {
        // Recorrido de todas las filas con datos en la columna A.
        for (let fila = 2; fila <= totalFilas; fila++) {
          // Construcción de la referencia de celda en la columna A para la fila actual.
          // Motivo: acceder secuencialmente a cada nombre.
          const celda = `A${fila}`;

          // Lectura del valor de la celda.
          // Motivo: obtener el texto del nombre.
          const valor = hoja.cell(celda).value();

          // Comprobación de que la celda contiene un dato significativo.
          // Motivo: evitar imprimir líneas vacías.
          if (valor !== null && valor !== undefined && String(valor).trim() !== "") {
            console.log(valor);
          }
        }
        console.log("Confirmación: el listado completo de nombres se ha realizado correctamente.");
      }

      // Mensaje de fin para indicar que todo el flujo ha concluido con éxito.
      console.log("Proceso finalizado correctamente.");
      return true;
    } catch (error) {
      // Mensaje claro en caso de incidencia durante la ejecución.
      console.error("Incidencia durante la lectura del Excel:", error);
      return false;
    }
  }
}

// Exportación de la clase o de una instancia según la arquitectura del proyecto.
// Motivo: permitir su utilización desde el resto de la aplicación.
// Si en tu proyecto ya existe la clase y su exportación, conserva ese patrón.
module.exports = ProcesosPrueba;
