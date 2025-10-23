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


  //----------------------------------------------------

  // Método que actuará como MVP para el proceso de posts (Strapi -> Directus).
  // Importante: en esta primera entrega SOLO preparamos las rutas y validaciones básicas,
  // sin implementar todavía la lectura del CSV/XLSX ni la escritura del CSV de salida.
  // Lo hacemos así para confirmar el flujo y los puntos de extensión antes de codificar la lógica.
  async procesoPosts(argumentos) {
    try {
      // 1) Mensaje inicial para localizar el inicio en consola.
      console.log("Inicio del proceso: Proceso Posts (MVP sin transformación)");

      // 2) Recogemos los argumentos que vienen del formulario (UI Angular).
      //    En tu config, el orden es:
      //    0: excelstrapi (ruta de archivo CSV/XLSX),
      //    1: rutaSalida (carpeta donde guardar el resultado).
      const rutaFormulario = argumentos?.formularioControl?.[0]; // excelstrapi
      const carpetaSalida  = argumentos?.formularioControl?.[1]; // rutaSalida

      // 3) Validaciones básicas de entrada para evitar errores comunes.
      if (typeof rutaFormulario !== "string" || !rutaFormulario.trim()) {
        console.error("No se ha proporcionado una ruta de archivo válida para 'excelstrapi'.");
        return false;
      }
      if (typeof carpetaSalida !== "string" || !carpetaSalida.trim()) {
        console.error("No se ha proporcionado una carpeta de salida válida para 'rutaSalida'.");
        return false;
      }

      // 4) Normalizamos rutas a absolutas para que no haya ambigüedades.
      const rutaInput = path.isAbsolute(rutaFormulario) ? rutaFormulario : path.resolve(rutaFormulario);
      const salidaDir = path.isAbsolute(carpetaSalida)  ? carpetaSalida  : path.resolve(carpetaSalida);

      // 5) Comprobamos existencia del archivo de entrada y preparamos la carpeta de salida.
      if (!fs.existsSync(rutaInput)) {
        console.error("El archivo de entrada no existe:", rutaInput);
        return false;
      }
      if (!fs.existsSync(salidaDir)) {
        fs.mkdirSync(salidaDir, { recursive: true });
      }

      // 6) Calculamos el nombre del archivo de salida según el MVP:
      //    "<nombre_entrada>_directus.csv"
      const baseName          = path.basename(rutaInput, path.extname(rutaInput));
      const nombreSalidaCsv   = `${baseName}_directus.csv`;
      const rutaSalidaCsvFull = path.normalize(path.join(salidaDir, nombreSalidaCsv));

      // 7) Solo mostramos por consola la ruta de salida planificada.
      //    AÚN NO escribimos el CSV de salida (se hará en el siguiente paso).
      console.log("Ruta de entrada:", rutaInput);
      console.log("Carpeta de salida:", salidaDir);
      console.log("Nombre de salida planificado:", rutaSalidaCsvFull);

            // ============================================================
      // MVP: Leer el input (CSV o XLSX) y generar <input>_directus.csv
      // ============================================================

      // Función auxiliar: convierte valores "NULL" o nulos a vacío
      // Motivo: en Strapi a veces viene el texto "NULL" y no queremos usarlo.
      const norm = (v) => {
        if (v === null || v === undefined) return "";
        const s = String(v).trim();
        if (s.toUpperCase() === "NULL") return "";
        return s;
      };

      // Función auxiliar: slugify simple para construir url_slug si faltase
      // Motivo: Directus necesita un slug consistente si no viene de Strapi.
      const slugify = (txt) => {
        const s = String(txt ?? "")
          .normalize("NFD")                  // separa acentos
          .replace(/[\u0300-\u036f]/g, "")   // elimina diacríticos
          .toLowerCase()
          .replace(/[^a-z0-9]+/g, "-")       // todo lo que no sea [a-z0-9] -> guión
          .replace(/^-+|-+$/g, "");          // quita guiones al inicio y fin
        return s || "sin-slug";
      };

      // Función auxiliar: formatea fecha/fecha-hora en "YYYY-MM-DDTHH:mm:ss"
      // Motivo: para el MVP replicamos la hora del archivo SIN hacer conversiones.
      const formatDateTime = (input) => {
        const raw = norm(input);
        if (!raw) return "";
        // Si viene con espacio en medio, lo pasamos a formato con 'T'.
        // Cortamos a los primeros 19 caracteres para eliminar microsegundos si los hubiese.
        const compact = raw.replace(" ", "T");
        // Si tiene microsegundos, recortamos:
        // - ejemplo "2025-10-13T14:03:19.750000" -> "2025-10-13T14:03:19"
        const base = compact.length >= 19 ? compact.slice(0, 19) : compact;
        // Si no tiene 'T' y es sólo fecha "YYYY-MM-DD", lo devolvemos tal cual.
        if (/^\d{4}-\d{2}-\d{2}$/.test(base)) return base;
        return base;
      };

      // Función auxiliar: valor CSV seguro con comillas si hace falta
      // Motivo: evitar romper el CSV si hay comas o comillas en el texto.
      const csvSafe = (val) => {
        // Arrays y booleanos tal cual (según formato de Directus que usaremos aquí)
        if (Array.isArray(val)) return JSON.stringify(val);
        if (typeof val === "boolean") return val ? "true" : "false";
        const s = String(val ?? "");
        // Si contiene comillas dobles, las duplicamos (estándar CSV)
        const needsQuotes = /[",\n]/.test(s);
        const escaped = s.replace(/"/g, '""');
        return needsQuotes ? `"${escaped}"` : escaped;
      };

      // 1) Leemos el input en memoria (CSV o XLSX)
      //    - Si es XLSX/XLSM: usamos XlsxPopulate (como testExcel)
      //    - Si es CSV: hacemos un parse sencillo por filas, respetando comillas
      let rows = [];   // aquí dejaremos un array de objetos {columna: valor}
      let headers = []; // cabeceras originales del input

      const ext = path.extname(rutaInput).toLowerCase();

      if (ext === ".xlsx" || ext === ".xlsm" || ext === ".xls") {
        // --- Lectura de Excel con XlsxPopulate (misma forma que testExcel) ---
        const wb = await XlsxPopulate.fromFileAsync(rutaInput);
        const hoja = wb.sheet(0);
        const used = hoja.usedRange();
        const data = used.value(); // matriz [filas][columnas]

        if (!Array.isArray(data) || data.length < 2) {
          console.error("El Excel no contiene filas suficientes (cabecera + datos).");
          return false;
        }

        // La primera fila será la cabecera
        headers = (data[0] || []).map((h) => norm(h));

        // Convertimos el resto de filas en objetos {cabecera: valor}
        for (let i = 1; i < data.length; i++) {
          const fila = data[i] || [];
          if (!Array.isArray(fila)) continue;

          const obj = {};
          for (let c = 0; c < headers.length; c++) {
            obj[headers[c]] = norm(fila[c]);
          }
          // Evitar filas totalmente vacías
          const hayAlgo = Object.values(obj).some((v) => String(v).trim() !== "");
          if (hayAlgo) rows.push(obj);
        }
      } else if (ext === ".csv") {
        // --- Lectura de CSV con parser muy sencillo (MVP) ---
        // NOTA: Para el MVP (1 registro) y porque no usaremos 'post_content',
        // este parser simple nos vale. Si en el futuro hay comas con comillas anidadas,
        // valoraremos endurecerlo o pedir XLSX.
        const raw = fs.readFileSync(rutaInput, "utf8");

        // Función: parsea una línea CSV respetando comillas dobles
        const parseCsvLine = (line) => {
          const out = [];
          let cur = "";
          let inQuotes = false;
          for (let i = 0; i < line.length; i++) {
            const ch = line[i];
            if (ch === '"') {
              // Doble comilla
              if (inQuotes && line[i + 1] === '"') {
                cur += '"'; // comilla escapada
                i++;        // saltar la siguiente
              } else {
                inQuotes = !inQuotes;
              }
            } else if (ch === "," && !inQuotes) {
              out.push(cur);
              cur = "";
            } else {
              cur += ch;
            }
          }
          out.push(cur);
          return out;
        };

        const lines = raw.split(/\r?\n/).filter((l) => l.trim() !== "");
        if (lines.length < 2) {
          console.error("El CSV no contiene filas suficientes (cabecera + datos).");
          return false;
        }

        // Cabeceras:
        headers = parseCsvLine(lines[0]).map((h) => norm(h));

        // Filas de datos:
        for (let i = 1; i < lines.length; i++) {
          const cols = parseCsvLine(lines[i]).map((v) => norm(v));
          // Alinear con nº de cabeceras
          while (cols.length < headers.length) cols.push("");
          const obj = {};
          for (let c = 0; c < headers.length; c++) {
            obj[headers[c]] = cols[c] ?? "";
          }
          // Evitar filas completamente vacías
          const hayAlgo = Object.values(obj).some((v) => String(v).trim() !== "");
          if (hayAlgo) rows.push(obj);
        }
      } else {
        console.error("Formato de entrada no soportado en este MVP. Use .csv o .xlsx");
        return false;
      }

      // 2) Si no hay filas, salimos
      if (rows.length === 0) {
        console.warn("No se encontraron filas de datos en el input.");
        return false;
      }

      // 3) Preparamos cabecera EXACTA de Directus (confirmada)
      const headerDirectus = [
        "categorias",
        "publicacion_automatica",
        "status",
        "url_slug",
        "id",
        "sort",
        "user_created",
        "date_created",
        "user_updated",
        "date_updated",
        "fecha",
        "etiquetas",
        "titulo",
        "contenido",
        "imagenes",
      ];

      // 4) Reglas fijas del MVP
      const UUID_FIJO = "0c839678-2c25-45ea-950a-10c0c9a50195";

      // 5) Transformación por fila (MVP: contenido vacío)
      const salidaLineas = [];
      // Añadimos la cabecera al CSV de salida
      salidaLineas.push(headerDirectus.join(","));

      for (const r of rows) {
        // Orígenes
        const postTitle = norm(r["post_title"]);
        const slug      = norm(r["slug"]);
        const postDate  = formatDateTime(r["post_date"]);
        const createdAt = formatDateTime(r["created_at"]);
        const updatedAt = formatDateTime(r["updated_at"]) || formatDateTime(r["post_modified"]) || formatDateTime(r["post_modified_gmt"]);

        // Fallbacks
        const titulo = postTitle || (norm(r["id"]) ? `Sin título (ID: ${norm(r["id"])})` : "Sin título");
        const urlSlug = slug || slugify(titulo);

        // 1) Fecha “pública” del post: priorizamos post_date y si no, created_at.
        //    Importante: NO usamos la fecha actual como fallback.
        const fecha = postDate || createdAt || "";
              
        // 2) Fecha de creación real en Directus = created_at del origen.
        //    Si faltara (no debería), caemos a post_date o, en último caso, vacío.
        const date_created = createdAt || postDate || "";
              
        // 3) Fecha de actualización: priorizamos updated_at; si no hay, usamos date_created.
        const date_updated = updatedAt || date_created;
              

        // Construimos la fila Directus (contenido vacío, arrays vacíos, status draft, etc.)
        const fila = [
          "[]",                                  // categorias
          false,                                 // publicacion_automatica
          "draft",                               // status
          csvSafe(urlSlug),                      // url_slug
          "",                                    // id (lo genera Directus)
          "",                                    // sort (vacío)
          UUID_FIJO,                             // user_created
          csvSafe(date_created),                 // date_created (misma hora, sin Z)
          UUID_FIJO,                             // user_updated
          csvSafe(date_updated),                 // date_updated (misma hora, sin Z)
          csvSafe(fecha),                        // fecha (misma hora, sin Z)
          "[]",                                  // etiquetas
          csvSafe(titulo),                       // titulo
          csvSafe(""),                           // contenido (MVP: vacío)
          "[]",                                  // imagenes
        ];

        salidaLineas.push(fila.join(","));
      }

      // 6) Escribimos el archivo CSV de salida
      fs.writeFileSync(rutaSalidaCsvFull, salidaLineas.join("\n"), "utf8");
      console.log("CSV de Directus generado en:", rutaSalidaCsvFull);


      // 8) TODO (siguiente paso):
      //    - Detectar si rutaInput es CSV o XLSX.
      //    - Leer el archivo (si CSV: parse; si XLSX: XlsxPopulate).
      //    - Mapear columnas al formato de Directus.
      //    - Dejar 'contenido' vacío y rellenar el resto de campos.
      //    - Escribir el CSV final en 'rutaSalidaCsvFull'.

      // 9) Fin del MVP de estructura.
      console.log("Proceso Posts (MVP): estructura validada. Pendiente implementar transformación.");
      return true;
    } catch (error) {
      console.error("Incidencia en Proceso Posts (MVP):", error);
      return false;
    }
  }

}
// Exportación de la clase o de una instancia según la arquitectura del proyecto.
// Motivo: permitir su utilización desde el resto de la aplicación.
// Si en tu proyecto ya existe la clase y su exportación, conserva ese patrón.
module.exports = ProcesosPrueba;
