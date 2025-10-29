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


// Método que procesa SÓLO JSON (Strapi -> Directus) y genera un JSON importable en Directus.
// Esta versión incluye:
//  - Mapeo de categorias/etiquetas (ya hecho en el paso anterior).
//  - Conversión de "contenido" desde bloques Strapi (Slate-like) → HTML limpio (WYSIWYG).
//  - Limpieza de atributos/etiquetas relacionadas con drag/draggable.
async procesoPosts(argumentos) {
  try {
    console.log("Inicio del proceso: Posts (JSON → JSON) con categorías/etiquetas y contenido WYSIWYG");

    // === 1) Entradas desde la UI (rutas) ===
    const rutaFormulario = argumentos?.formularioControl?.[0]; // JSON Strapi (array de posts)
    const carpetaSalida  = argumentos?.formularioControl?.[1]; // carpeta de salida

    if (typeof rutaFormulario !== "string" || !rutaFormulario.trim()) {
      console.error("No se ha proporcionado una ruta de archivo válida para el JSON de Strapi.");
      return false;
    }
    if (typeof carpetaSalida !== "string" || !carpetaSalida.trim()) {
      console.error("No se ha proporcionado una carpeta de salida válida.");
      return false;
    }

    // Normalizamos rutas a absolutas para evitar ambigüedades del SO.
    const rutaInput = path.isAbsolute(rutaFormulario) ? rutaFormulario : path.resolve(rutaFormulario);
    const salidaDir = path.isAbsolute(carpetaSalida)  ? carpetaSalida  : path.resolve(carpetaSalida);

    // Comprobaciones de existencia de fichero/carpeta.
    if (!fs.existsSync(rutaInput)) {
      console.error("El archivo de entrada no existe:", rutaInput);
      return false;
    }
    if (!fs.existsSync(salidaDir)) {
      fs.mkdirSync(salidaDir, { recursive: true });
    }

    // Solo aceptamos .json (hemos decidido trabajar únicamente con JSON).
    const baseName = path.basename(rutaInput, path.extname(rutaInput));
    const ext      = path.extname(rutaInput).toLowerCase();
    if (ext !== ".json") {
      console.error("Formato no soportado. Este proceso acepta únicamente archivos .json");
      return false;
    }

    // === 2) Utilidades auxiliares simples (para texto, slug y fechas SIN 'Z') ===
    const norm = (v) => {
      if (v === null || v === undefined) return "";
      const s = String(v).trim();
      if (s.toUpperCase() === "NULL") return "";
      return s;
    };

    const slugify = (txt) => {
      const s = String(txt ?? "")
        .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
        .toLowerCase()
        .replace(/[^a-z0-9]+/g, "-")
        .replace(/^-+|-+$/g, "");
      return s || "sin-slug";
    };

    const formatDateTime = (input) => {
      const raw = norm(input);
      if (!raw) return "";
      const compact = raw.replace(" ", "T");  // "YYYY-MM-DD HH:mm:ss" → "YYYY-MM-DDTHH:mm:ss"
      const base = compact.length >= 19 ? compact.slice(0, 19) : compact;
      if (/^\d{4}-\d{2}-\d{2}$/.test(base)) return base;
      return base;
    };

    // === 3) Helpers específicos de CONTENIDO (bloques → HTML) ===

    // 3.1) Intenta parsear una cadena JSON con seguridad. Si falla, devuelve null.
    const safeJsonParse = (str) => {
      try {
        return JSON.parse(str);
      } catch {
        return null;
      }
    };

    // 3.2) Extrae texto concatenado de un array de "children" (Slate):
    //      - Cada child suele tener { text: "..." }.
    //      - Unimos todos los .text conservando espacios básicos.
    const childrenToPlainText = (children) => {
      if (!Array.isArray(children)) return "";
      return children
        .map((ch) => norm(ch?.text))
        .filter((t) => t !== "")
        .join(" ");
    };

    // 3.3) Convierte un array de bloques Strapi (Slate-like) a HTML.
    //      Soporta: "heading" (level 1..6) y "paragraph".
    //      Si aparecen otros tipos, los tratamos como párrafos para no romper.
    const blocksToHtml = (blocks) => {
      if (!Array.isArray(blocks)) return "";

      const parts = [];

      for (const block of blocks) {
        const type  = norm(block?.type).toLowerCase();   // tipo del bloque (heading/paragraph/otro)
        const level = Number(block?.level) || 0;         // nivel de heading si aplica
        const text  = childrenToPlainText(block?.children);

        // Evitamos generar etiquetas vacías con puro texto vacío (pero dejamos <p></p> si quieres huecos).
        const cleanText = text; // aquí podríamos .trim() si necesitas eliminar espacios extremos

        if (type === "heading") {
          // Nivel válido entre 1 y 6; si no, caemos a h2 por defecto muy suave.
          const n = level >= 1 && level <= 6 ? level : 2;
          parts.push(`<h${n}>${cleanText}</h${n}>`);
        } else if (type === "paragraph") {
          parts.push(`<p>${cleanText}</p>`);
        } else {
          // Cualquier tipo no reconocido lo tratamos como párrafo para no romper el flujo.
          parts.push(`<p>${cleanText}</p>`);
        }
      }

      // Unimos el HTML final.
      return parts.join("");
    };

    // 3.4) Limpia HTML de atributos/etiquetas relacionadas con drag/draggable.
    //      - Quita atributos draggable="...", data-drag="..." y similares.
    //      - Elimina etiquetas <drag>…</drag> si existieran.
    //      Nota: usamos regex simples porque el contenido son títulos y párrafos (no HTML complejo).
    const cleanDragFromHtml = (html) => {
      let out = String(html ?? "");

      // Eliminar etiquetas <drag>...</drag> (no suelen aparecer, pero por si acaso)
      out = out.replace(/<\s*drag\b[^>]*>[\s\S]*?<\s*\/\s*drag\s*>/gi, "");

      // Eliminar cualquier atributo draggable="..." (true/false u otros)
      out = out.replace(/\sdraggable\s*=\s*"(?:[^"]*)"/gi, "");
      out = out.replace(/\sdraggable\s*=\s*'(?:[^']*)'/gi, "");
      out = out.replace(/\sdraggable\b/gi, ""); // por si queda sin valor

      // Eliminar atributos data-* que contengan "drag" (p.ej. data-drag, data-drag-id, etc.)
      out = out.replace(/\sdata-[a-z0-9_-]*drag[a-z0-9_-]*\s*=\s*"(?:[^"]*)"/gi, "");
      out = out.replace(/\sdata-[a-z0-9_-]*drag[a-z0-9_-]*\s*=\s*'(?:[^']*)'/gi, "");

      // Eliminar clases que contengan la palabra 'drag' (por si viniera algo tipo class="foo drag bar")
      out = out.replace(/\sclass\s*=\s*"([^"]*)"/gi, (m, cls) => {
        const filtered = cls
          .split(/\s+/)
          .filter((c) => !/drag/i.test(c))
          .join(" ");
        return filtered ? ` class="${filtered}"` : "";
      });
      out = out.replace(/\sclass\s*=\s*'([^']*)'/gi, (m, cls) => {
        const filtered = cls
          .split(/\s+/)
          .filter((c) => !/drag/i.test(c))
          .join(" ");
        return filtered ? ` class='${filtered}'` : "";
      });

      return out;
    };

    // 3.5) Función principal para convertir el campo "contenido" de Strapi a HTML limpio:
    //      - Si viene como JSON string de bloques (tu caso), lo parsea y convierte con blocksToHtml.
    //      - Si viniera como HTML ya, sólo lo limpia de drag/draggable.
    const toWysiwygHtml = (contenidoStrapi) => {
      const raw = norm(contenidoStrapi);
      if (!raw) return "";

      // ¿Parece HTML ya? (comienza por "<")
      if (/^\s*</.test(raw)) {
        return cleanDragFromHtml(raw);
      }

      // Si no parece HTML, intentamos parsear como JSON (array de bloques)
      const parsed = safeJsonParse(raw);
      if (Array.isArray(parsed)) {
        const html = blocksToHtml(parsed);
        return cleanDragFromHtml(html);
      }

      // Si no es HTML ni JSON de bloques, devolvemos tal cual (o lo envolvemos en <p>).
      // Aquí elegimos envolver en <p> para garantizar formato WYSIWYG.
      return cleanDragFromHtml(`<p>${raw}</p>`);
    };

    // === 4) Lectura y validación del JSON de Strapi ===
    const raw = fs.readFileSync(rutaInput, "utf8");
    let arrPosts;
    try {
      arrPosts = JSON.parse(raw);
    } catch (e) {
      console.error("ERROR: El JSON de entrada no es válido. Detalle:", e.message);
      return false;
    }
    if (!Array.isArray(arrPosts)) {
      console.error("ERROR: Se esperaba un ARRAY de posts en el JSON de entrada.");
      return false;
    }

    // === 5) UUID fijo para autoría (confirmado) ===
    const UUID_FIJO = "0c839678-2c25-45ea-950a-10c0c9a50195";

    // === 6) Transformación Strapi → Directus ===
    const salidaDirectus = arrPosts.map((p) => {
      // Campos base Strapi
      const post_title = norm(p.post_title);
      const slug       = norm(p.slug);
      const post_date  = formatDateTime(p.post_date);
      const created_at = formatDateTime(p.created_at);
      const updated_at = formatDateTime(p.updated_at || p.post_modified || p.post_modified_gmt);

      // Título/slug finales
      const titulo   = post_title || (norm(p.id) ? `Sin título (ID: ${norm(p.id)})` : "Sin título");
      const url_slug = slug || slugify(titulo);

      // Fechas Directus (sin Z)
      const fecha        = post_date || created_at || "";
      const date_created = created_at || post_date || "";
      const date_updated = updated_at || date_created;

      // === Categorías / Etiquetas (reutilizamos el parser del paso anterior) ===
      const parseStrapiList = (input) => {
        const rawVal = input ?? "";
        if (Array.isArray(rawVal)) {
          return rawVal.map(norm).filter((x) => x !== "").filter((x, i, a) => a.indexOf(x) === i);
        }
        if (typeof rawVal === "string" && rawVal.trim().startsWith("[")) {
          const arr = safeJsonParse(rawVal);
          if (Array.isArray(arr)) {
            return arr.map(norm).filter((x) => x !== "").filter((x, i, a) => a.indexOf(x) === i);
          }
        }
        if (typeof rawVal === "string") {
          return rawVal.split(",").map(norm).filter((x) => x !== "").filter((x, i, a) => a.indexOf(x) === i);
        }
        return [];
      };

      const categoriasStrapi = parseStrapiList(p.categorias);
      const etiquetasStrapi  = parseStrapiList(p.etiquetas);

      // Mapeo al formato de tu Directus exportado (relación con objeto intermedio):
      const categoria = categoriasStrapi.map((nombre) => ({
        categoria_post_id: { categoria: nombre }
      }));
      const etiqueta = etiquetasStrapi.map((nombre) => ({
        etiqueta_post_id: { etiqueta: nombre }
      }));

      // === Contenido WYSIWYG (AHORA SÍ) ===
      const contenidoHtml = toWysiwygHtml(p.contenido || p.post_content || "");

      // Objeto final para importar en Directus
      return {
        // Contenido principal
        titulo,
        url_slug,
        contenido: contenidoHtml,   // HTML limpio (WYSIWYG)
        // Fechas
        fecha,
        date_created,
        date_updated,
        // Estado / flags
        status: "draft",
        publicacion_automatica: false,
        // Autoría
        user_created: UUID_FIJO,
        user_updated: UUID_FIJO,
        // Relaciones
        categoria,
        etiqueta,
        // Imágenes (pendiente de implementar)
        imagenes: [],
      };
    });

    // === 7) Escritura del JSON de salida (sufijo "_directus.json") ===
    const rutaSalidaJsonFull = path.normalize(path.join(salidaDir, `${baseName}_directus.json`));
    fs.writeFileSync(rutaSalidaJsonFull, JSON.stringify(salidaDirectus, null, 2), "utf8");

    console.log("OK: Generado JSON (con contenido WYSIWYG + limpieza drag) para importar en Directus →", rutaSalidaJsonFull);
    console.log("Proceso Posts finalizado correctamente.");
    return true;

  } catch (error) {
    console.error("Incidencia en Proceso Posts:", error);
    return false;
  }
}

// =============================
// PROCESO: Webscrapping YouTube
// =============================
// Argumentos esperados (en este orden: 0,1,2):
// 0: rutaChromium   -> Ruta al ejecutable de Chrome/Chromium.
// 1: excelEntrada   -> Ruta del Excel (.xlsx/.xlsm) con columna 'url_video' en la primera hoja.
// 2: rutaSalida     -> Carpeta donde guardar el Excel *_METRICAS.xlsx.
//
// Este método abre cada URL de YouTube, extrae métricas básicas y las escribe en un nuevo Excel.
// Está muy comentado (nivel principiante) y con selectores bastante robustos.
async webscrappingYoutube(argumentos) {
  // Cargamos puppeteer aquí para que otros procesos no fallen si no está instalado.
  // Si usas "puppeteer-core", cambia esta línea a require("puppeteer-core")
  const puppeteer = require("puppeteer-core");

  try {
    console.log("Inicio del proceso: Webscrapping YouTube");

    // 1) Leemos los argumentos en el orden acordado.
    const rutaChromium  = argumentos?.formularioControl?.[0]; // ejecutable del navegador
    const rutaExcel     = argumentos?.formularioControl?.[1]; // excel con url_video
    const carpetaSalida = argumentos?.formularioControl?.[2]; // carpeta destino

    // 2) Validaciones básicas para evitar errores tontos.
    if (typeof rutaChromium !== "string" || !rutaChromium.trim()) {
      console.error("Falta la ruta del ejecutable de Chrome/Chromium (argumento 0).");
      return false;
    }
    if (typeof rutaExcel !== "string" || !rutaExcel.trim()) {
      console.error("Falta la ruta del Excel de entrada (argumento 1).");
      return false;
    }
    if (typeof carpetaSalida !== "string" || !carpetaSalida.trim()) {
      console.error("Falta la ruta de la carpeta de salida (argumento 2).");
      return false;
    }

    // 3) Normalizamos rutas para evitar problemas por el sistema operativo.
    const chromiumExecutablePath = path.normalize(rutaChromium);
    const inputPath  = path.isAbsolute(rutaExcel) ? rutaExcel : path.resolve(rutaExcel);
    const outputDir  = path.isAbsolute(carpetaSalida) ? carpetaSalida : path.resolve(carpetaSalida);

    // 4) Comprobamos existencia del Excel y preparamos carpeta de salida.
    if (!fs.existsSync(inputPath)) {
      console.error("El Excel indicado no existe:", inputPath);
      return false;
    }
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir, { recursive: true });
    }

    // 5) Abrimos el Excel con XlsxPopulate y leemos la primera hoja.
    const wb = await XlsxPopulate.fromFileAsync(inputPath);
    const hoja = wb.sheet(0); // primera hoja
    const used = hoja.usedRange();
    const data = used ? used.value() : [];
    if (!Array.isArray(data) || data.length < 2) {
      console.error("El Excel no tiene filas suficientes (se necesita cabecera + al menos 1 URL).");
      return false;
    }

    // 6) Buscamos la cabecera 'url_video' (asumimos que está en la primera fila).
    const headers = (data[0] || []).map((h) => String(h ?? "").trim());
    const colUrlIdx = headers.findIndex((h) => h.toLowerCase() === "url_video");
    if (colUrlIdx === -1) {
      console.error("No se ha encontrado la columna 'url_video' en la fila 1 del Excel.");
      return false;
    }

    // 7) Construimos la lista de URLs desde la columna 'url_video' (filas 2..N).
    const urls = [];
    for (let i = 1; i < data.length; i++) {
      const fila = data[i] || [];
      const url  = String(fila[colUrlIdx] ?? "").trim();
      if (url) urls.push(url);
    }
    if (urls.length === 0) {
      console.error("No se han encontrado URLs en la columna 'url_video'.");
      return false;
    }

    // 8) Definimos los nombres de las columnas nuevas que vamos a escribir.
    const outputCols = [
      "title",
      "channel_name",
      "channel_url",
      "views_text",
      "publish_date_text",
      "likes_text",
      "category_best_effort",
      "ok",
      "error",
    ];

    // 9) Calculamos desde qué columna empezamos a escribir las nuevas métricas.
    const totalColumnas = headers.length;
    const startColIndex = totalColumnas; // 0-based en nuestro array "data"; en Excel será +1
    // Escribimos las cabeceras nuevas en la fila 1.
    outputCols.forEach((colName, i) => {
      hoja.cell(1, startColIndex + 1 + i).value(colName);
    });

    // 10) Lanzamos el navegador (Puppeteer) según tu formato.
    const browser = await puppeteer.launch({
      executablePath: chromiumExecutablePath,
      headless: false,
    });
    const page = await browser.newPage();

    // 11) Ajustamos un viewport estándar para consistencia, y userAgent para parecer un navegador real.
    await page.setViewport({ width: 1366, height: 768 });
    await page.setUserAgent("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120 Safari/537.36");

    // 12) Helper: clic automático en el consentimiento si aparece (YouTube/GDPR).
    const tryAcceptConsent = async () => {
      try {
        // Botones típicos en español/inglés
        const selectors = [
          'button:has-text("Aceptar todo")',
          'button:has-text("Acepto")',
          'button:has-text("I agree")',
          'button:has-text("Agree to all")',
          'form[action*="consent"] button', // genérico
          '#introAgreeButton', // antiguas UIs
        ];

        for (const sel of selectors) {
          const el = await page.$(sel).catch(() => null);
          if (el) { await el.click(); await page.waitForTimeout(500); break; }
        }
      } catch { /* ignoramos */ }
    };

    // 13) Helper: intenta varias estrategias para obtener un texto por selector.
    const getText = async (selectors) => {
      for (const sel of selectors) {
        try {
          const handle = await page.$(sel);
          if (!handle) continue;
          const txt = await page.$eval(sel, (el) => el.textContent?.trim() || "");
          if (txt) return txt;
        } catch { /* probamos el siguiente */ }
      }
      return "";
    };

    // 14) Recorremos cada URL y extraemos métricas.
    for (let i = 0; i < urls.length; i++) {
      const url = urls[i];
      console.log(`→ [${i + 1}/${urls.length}] Navegando a: ${url}`);

      // Fila Excel a escribir (i+2 porque fila 1 es cabecera).
      const excelRow = i + 2;

      let title = "", channel_name = "", channel_url = "", views_text = "", publish_date_text = "", likes_text = "", category_best_effort = "";
      let ok = false, errorMsg = "";

      try {
        // 14.1) Navegamos a la URL y esperamos a que la red esté estable.
        await page.goto(url, { waitUntil: "networkidle2", timeout: 60000 });

        // 14.2) Intentamos aceptar consentimiento si aparece.
        await tryAcceptConsent();

        // 14.3) Esperamos un contenedor principal del watch de YouTube.
        // Nota: YouTube cambia mucho, por eso usamos varios selectores alternativos.
        await page.waitForSelector("ytd-watch-flexy, #columns", { timeout: 15000 });

        // 14.4) Extraemos el título del vídeo (varias rutas + fallback document.title).
        title = await getText([
          "h1.title",                         // UIs antiguas
          "#title h1",                        // UIs intermedias
          "ytd-watch-metadata h1",            // UIs nuevas
        ]);
        if (!title) {
          title = await page.title();         // Fallback final
        }

        // 14.5) Nombre y URL del canal.
        channel_name = await getText([
          "#owner #text",                     // habitual
          "ytd-channel-name a",               // alternativa
          "a.yt-simple-endpoint.style-scope.yt-formatted-string", // genérica
        ]);

        try {
          const channelEl = await page.$("ytd-channel-name a[href], #owner a[href]");
          if (channelEl) {
            channel_url = await page.evaluate((a) => a.getAttribute("href") || "", channelEl);
            if (channel_url && !/^https?:\/\//i.test(channel_url)) {
              channel_url = "https://www.youtube.com" + channel_url;
            }
          }
        } catch { /* nada */ }

        // 14.6) Vistas (texto tal cual muestra la página).
        views_text = await getText([
          "#info ytd-video-view-count-renderer",
          "ytd-video-view-count-renderer span",
          "span.view-count",
          "ytd-watch-metadata #info span",
        ]);

        // 14.7) Fecha de publicación (texto estilo “23 oct 2025”).
        publish_date_text = await getText([
          "#info-strings yt-formatted-string",
          "ytd-watch-metadata #info-strings yt-formatted-string",
          "div#info-strings yt-formatted-string",
        ]);

        // 14.8) Likes (best-effort: YouTube oculta/traslada a veces).
        likes_text = await getText([
          // UIs recientes
          "ytd-segmented-like-dislike-button-renderer yt-formatted-string",
          // UIs intermedias
          "ytd-toggle-button-renderer[is-icon-button][aria-pressed] yt-formatted-string",
          // UIs antiguas
          "ytd-toggle-button-renderer[is-icon-button] #text",
        ]);

        // 14.9) Intento de categoría (si aparece el bloque “Mostrar más” con metadatos).
        // Abrimos "Mostrar más" si existe para descubrir metadata adicional.
        try {
          const showMore = await page.$("tp-yt-paper-button#expand, #description #expand");
          if (showMore) { await showMore.click(); await page.waitForTimeout(300); }
        } catch { /* nada */ }

        // Buscamos pistas de “Categoría” en descripción (cambia según idioma)
        const descText = await getText([
          "#description ytd-text-inline-expander",
          "#description",
          "ytd-text-inline-expander[collapsed] #content",
        ]);
        // Heurística muy simple: busca "Categoría" o "Category" y corta la línea.
        if (descText) {
          const m = descText.match(/Categor[ií]a:\s*([^\n\r]+)/i) || descText.match(/Category:\s*([^\n\r]+)/i);
          category_best_effort = m ? m[1].trim() : "";
        }

        ok = true;
      } catch (err) {
        ok = false;
        errorMsg = err?.message || "Error desconocido navegando o seleccionando elementos.";
        console.warn("   ⚠️  Error:", errorMsg);
      }

      // 14.10) Escribimos las métricas en la misma fila, empezando en la primera columna libre.
      const valores = [
        title,
        channel_name,
        channel_url,
        views_text,
        publish_date_text,
        likes_text,
        category_best_effort,
        ok ? "true" : "false",
        errorMsg,
      ];

      for (let c = 0; c < valores.length; c++) {
        // .cell(fila, columna) es 1-based en XlsxPopulate.
        hoja.cell(excelRow, startColIndex + 1 + c).value(valores[c]);
      }
    }

    // 15) Guardamos el Excel con sufijo *_METRICAS.xlsx en la carpeta de salida.
    const baseName = path.basename(inputPath, path.extname(inputPath));
    const rutaSalida = path.join(outputDir, `${baseName}_METRICAS.xlsx`);
    await wb.toFileAsync(rutaSalida);

    // 16) Cerramos el navegador y listo.
    await browser.close();

    console.log("Archivo con métricas guardado en:", rutaSalida);
    console.log("Proceso Webscrapping YouTube finalizado correctamente.");
    return true;
  } catch (error) {
    console.error("Incidencia en Webscrapping YouTube:", error);
    return false;
  }
}


// =============================
// PROCESO: Imágenes de Posts
// Lee un JSON (post_id, post_date, order, file_url) y DESCARGA las imágenes
// Guardando con el nombre: post_id-YYYYMMDD_HHMMSS-order.ext
// =============================
// =============================
// PROCESO: Imágenes de Posts
// Lee un JSON (post_id, post_date, order, file_url),
// DESCARGA las imágenes con nombre: post_id-YYYYMMDD_HHMMSS-order.ext
// y genera un JSON resumen con la ruta local y el enlace file:///
// =============================
async imagenPost(argumentos) {
  // Cargamos módulos nativos necesarios SÓLO aquí para no afectar a otros procesos.
  // Motivo: escribir código aislado y fácil de mantener.
  const path = require("path");                 // unir rutas (evita problemas entre Windows/Mac/Linux)
  const fs = require("fs");                     // leer archivos, crear carpetas y escribir ficheros
  const http = require("http");                 // descarga si la URL es http://
  const https = require("https");               // descarga si la URL es https://
  const { URL, pathToFileURL } = require("url");// URL para parsear direcciones y pathToFileURL para crear file://

  try {
    console.log("Inicio del proceso: Imágenes de Posts (validación + descarga + resumen)");

    // 1) Recuperamos las entradas del formulario tal y como las definiste en la librería de procesos.
    //    argumentos?.formularioControl es un array con los valores del UI. Orden:
    //    [0] -> json_strapi (ruta del archivo .json de entrada)
    //    [1] -> rutaSalida  (carpeta donde guardar las imágenes y el resumen)
    const rutaJsonStrapi = argumentos?.formularioControl?.[0];
    const carpetaSalida  = argumentos?.formularioControl?.[1];

    // 2) Validamos que el usuario ha introducido textos (evita errores tipo .trim de undefined)
    //    Motivo: evitar fallos de ejecución por parámetros vacíos.
    if (typeof rutaJsonStrapi !== "string" || !rutaJsonStrapi.trim()) {
      console.error("No se ha proporcionado una ruta válida para el JSON de Strapi.");
      return false;
    }
    if (typeof carpetaSalida !== "string" || !carpetaSalida.trim()) {
      console.error("No se ha proporcionado una carpeta de salida válida.");
      return false;
    }

    // 3) Normalizamos rutas a absolutas (así el sistema siempre encuentra bien los recursos)
    //    Motivo: compatibilidad total entre Windows, macOS y Linux.
    const rutaInput = path.isAbsolute(rutaJsonStrapi) ? rutaJsonStrapi : path.resolve(rutaJsonStrapi);
    const salidaDir = path.isAbsolute(carpetaSalida)  ? carpetaSalida  : path.resolve(carpetaSalida);

    // 4) Comprobamos que el JSON existe y preparamos la carpeta de salida (si no existe, la creamos)
    //    Motivo: garantizar fuente y destino disponibles antes de procesar.
    if (!fs.existsSync(rutaInput)) {
      console.error("El archivo JSON indicado no existe:", rutaInput);
      return false;
    }
    if (!fs.existsSync(salidaDir)) {
      fs.mkdirSync(salidaDir, { recursive: true });
    }

    // 5) Leemos y parseamos el JSON. Debe ser un ARRAY de objetos con
    //    { post_id, post_date, order, file_url } por fila.
    //    Motivo: disponer en memoria de la lista de descargas a realizar.
    const raw = fs.readFileSync(rutaInput, "utf8");
    let filas;
    try {
      filas = JSON.parse(raw);
    } catch (e) {
      console.error("El JSON no tiene un formato válido. Detalle:", e.message);
      return false;
    }
    if (!Array.isArray(filas) || filas.length === 0) {
      console.error("Se esperaba un ARRAY con al menos un elemento.");
      return false;
    }

    // 6) Función auxiliar: convertir la fecha a formato simple "YYYYMMDD_HHMMSS"
    //    Entradas como "2025-10-29 10:15:45.841000" -> "20251029_101545"
    //    Motivo: generar nombres de archivo homogéneos y ordenables por fecha.
    const fechaSimple = (isoLike) => {
      if (!isoLike) return "fecha";
      const base = String(isoLike).slice(0, 19).replace("T", " "); // limpiamos posibles 'T' y milisegundos
      const [f, h] = base.split(" ");
      if (!f) return "fecha";
      const y = f.slice(0, 4);
      const m = f.slice(5, 7);
      const d = f.slice(8, 10);
      const hh = (h || "00:00:00").slice(0, 2);
      const mm = (h || "00:00:00").slice(3, 5);
      const ss = (h || "00:00:00").slice(6, 8);
      return `${y}${m}${d}_${hh}${mm}${ss}`;
    };

    // 7) Función auxiliar: deducir la extensión a partir de la URL.
    //    Si no hay extensión en la URL, usaremos ".bin" como acordamos.
    //    Motivo: guardar con la extensión correcta siempre que sea posible.
    const deducirExtension = (fileUrl) => {
      let extension = ".bin";
      try {
        const u = new URL(String(fileUrl));
        const base = u.pathname.split("/").pop() || "";
        const punto = base.lastIndexOf(".");
        if (punto !== -1) extension = base.slice(punto).toLowerCase();
      } catch {
        // Si la URL está mal formada, nos quedamos con ".bin"
      }
      return extension || ".bin";
    };

    // 8) Función auxiliar: descarga una URL en disco, con soporte básico de redirecciones.
    //    Motivo: Cloudinary/CDNs suelen responder con redirecciones temporales/firmadas.
    const descargar = (urlStr, destino, maxRedirects = 5) => {
      return new Promise((resolve, reject) => {
        let redirRestantes = maxRedirects;

        const hacerPeticion = (currentUrl) => {
          const client = currentUrl.startsWith("https") ? https : http;

          const req = client.get(currentUrl, { headers: { "User-Agent": "Mozilla/5.0" } }, (res) => {
            if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
              res.resume();
              if (redirRestantes <= 0) return reject(new Error("Demasiadas redirecciones."));
              redirRestantes--;
              const nuevaUrl = new URL(res.headers.location, currentUrl).toString();
              return hacerPeticion(nuevaUrl);
            }

            if (res.statusCode !== 200) {
              res.resume();
              return reject(new Error(`Respuesta HTTP inesperada: ${res.statusCode}`));
            }

            const fileStream = fs.createWriteStream(destino);
            fileStream.on("finish", () => fileStream.close(() => resolve(true)));
            fileStream.on("error", (err) => {
              req.destroy();
              reject(err);
            });

            res.pipe(fileStream);
          });

          req.on("error", (err) => reject(err));
        };

        hacerPeticion(urlStr);
      });
    };

    // 9) PREPARAMOS todas las tareas (una por imagen) con su nombre final y su ruta destino.
    //    Motivo: disponer de una lista cerrada para iterar y descargar secuencialmente.
    const trabajos = [];
    for (const fila of filas) {
      const postId   = fila?.post_id;
      const postDate = fila?.post_date;
      const order    = fila?.order;
      const fileUrl  = fila?.file_url;

      if (!postId || !postDate || !order || !fileUrl) {
        console.warn("Fila incompleta. Se omite:", fila);
        continue;
      }

      const nombreFecha   = fechaSimple(postDate);
      const extension     = deducirExtension(fileUrl);
      const nombreArchivo = `${postId}-${nombreFecha}-${order}${extension}`;
      const rutaDestino   = path.join(salidaDir, nombreArchivo);

      trabajos.push({ post_id: postId, post_date: postDate, order, file_url: fileUrl, destino: rutaDestino });
    }

    // 10) Si no hay tareas válidas, salimos sin error crítico para que el usuario lo vea claro.
    console.log(`Tareas preparadas: ${trabajos.length}`);
    if (trabajos.length === 0) {
      console.warn("No hay tareas válidas para descargar. Revisa el JSON.");
      return false;
    }
    console.log("Ejemplo de primera tarea:", trabajos[0]);

    // *** NUEVO ***: estructura para acumular resultados y luego escribir el JSON resumen.
    // Motivo: replicar el input y añadir local_path + local_uri + flags de estado.
    const resultados = [];

    // 11) Ejecutamos las descargas de forma SENCILLA y SECUENCIAL (más estable y fácil de depurar).
    //     Motivo: mantenerlo simple. Si se requiere velocidad, luego añadimos concurrencia.
    let ok = 0, ko = 0;
    for (let i = 0; i < trabajos.length; i++) {
      const t = trabajos[i];
      const idx = i + 1;

      try {
        if (fs.existsSync(t.destino)) {
          console.log(`[${idx}/${trabajos.length}] Ya existe, se omite: ${path.basename(t.destino)}`);

          // Registramos en resultados el caso "skip"
          resultados.push({
            post_id: t.post_id,
            post_date: t.post_date,
            order: t.order,
            file_url: t.file_url,
            local_path: t.destino,
            local_uri: pathToFileURL(t.destino).href,
            ok: true,
            skipped: true,
            error: null
          });

          ok++;
          continue;
        }

        console.log(`[${idx}/${trabajos.length}] Descargando: ${t.file_url}`);
        await descargar(t.file_url, t.destino);
        console.log(`    ✔ Guardado como: ${t.destino}`);

        // Registramos en resultados la descarga correcta
        resultados.push({
          post_id: t.post_id,
          post_date: t.post_date,
          order: t.order,
          file_url: t.file_url,
          local_path: t.destino,
          local_uri: pathToFileURL(t.destino).href,
          ok: true,
          skipped: false,
          error: null
        });

        ok++;
      } catch (e) {
        console.warn(`    ✖ Error descargando (se omite): ${e.message}`);

        // Registramos en resultados el error de descarga (dejamos constancia de la ruta prevista)
        resultados.push({
          post_id: t.post_id,
          post_date: t.post_date,
          order: t.order,
          file_url: t.file_url,
          local_path: t.destino,
          local_uri: pathToFileURL(t.destino).href,
          ok: false,
          skipped: false,
          error: e.message || "Error desconocido"
        });

        ko++;
      }
    }

    // 12) Escribimos el JSON resumen en la MISMA carpeta de salida, con sufijo *_con_ruta.json.
    //     Motivo: no sobrescribir el input y disponer de un resultado listo para consultas rápidas.
    const baseNameInput = path.basename(rutaInput, path.extname(rutaInput));
    const rutaResumen   = path.join(salidaDir, `${baseNameInput}_con_ruta.json`);
    fs.writeFileSync(rutaResumen, JSON.stringify(resultados, null, 2), "utf8");
    console.log("Resumen con rutas locales guardado en:", rutaResumen);

    // 13) Resumen final por consola y retorno
    console.log(`Finalizado. Correctas: ${ok}, Fallidas: ${ko}, Total: ${trabajos.length}`);
    return ok > 0; // devolvemos true si al menos una descarga fue correcta

  } catch (error) {
    console.error("Incidencia en Imágenes de Posts:", error);
    return false;
  }
}






}
// Exportación de la clase o de una instancia según la arquitectura del proyecto.
// Motivo: permitir su utilización desde el resto de la aplicación.
// Si en tu proyecto ya existe la clase y su exportación, conserva ese patrón.
module.exports = ProcesosPrueba;
