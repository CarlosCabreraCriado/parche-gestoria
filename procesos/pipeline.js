const path = require("path");
const fs = require("fs");
const { spawn } = require("child_process");
const { registrarEjecucion } = require("../metricas");
const ProcesosDuplicados = require("./duplicados");

/**
 * Pipelines integrados: genera informes desde analisis-a3 (Python)
 * y los encadena con procesos de parche-gestoria.
 *
 * MVP: Format 8 (Altas) -> Duplicados TA2+IDC
 */
class ProcesosPipeline {
  constructor(pathToDbFolder, nombreProyecto, proyectoDB) {
    this.pathToDbFolder = pathToDbFolder;
    this.nombreProyecto = nombreProyecto;
    this.proyectoDB = proyectoDB;
    this.TAG = "[PIPELINE]";
  }

  // ==========================================================
  // Utils
  // ==========================================================

  log(msg, ...rest) {
    console.log(`${this.TAG} ${msg}`, ...rest);
  }

  logWarn(msg, ...rest) {
    console.warn(`${this.TAG} ${msg}`, ...rest);
  }

  logErr(msg, ...rest) {
    console.error(`${this.TAG} ${msg}`, ...rest);
  }

  ensureDir(dir) {
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
  }

  // ==========================================================
  // Python subprocess helper (reutilizable para futuros pipelines)
  // ==========================================================

  /**
   * Ejecuta un script Python como subprocess.
   * @param {string} pythonPath  - Ruta al ejecutable de Python (o "python")
   * @param {string} cwd         - Directorio de trabajo (raíz de analisis-a3)
   * @param {string[]} args      - Argumentos para el script
   * @param {number} [timeoutMs] - Timeout en ms (default: 5 min)
   * @returns {Promise<{code: number, stdout: string, stderr: string}>}
   */
  _spawnPython(pythonPath, cwd, args, timeoutMs = 300_000) {
    return new Promise((resolve, reject) => {
      this.log(`Ejecutando: ${pythonPath} ${args.join(" ")}`);
      this.log(`CWD: ${cwd}`);

      const proc = spawn(pythonPath, args, {
        cwd: path.normalize(cwd),
        windowsHide: true,
        stdio: ["ignore", "pipe", "pipe"],
      });

      let stdout = "";
      let stderr = "";
      let killed = false;

      proc.stdout.on("data", (chunk) => {
        const text = chunk.toString("utf8");
        stdout += text;
        this.log(`[Python stdout] ${text.trimEnd()}`);
      });

      proc.stderr.on("data", (chunk) => {
        const text = chunk.toString("utf8");
        stderr += text;
        this.logWarn(`[Python stderr] ${text.trimEnd()}`);
      });

      const timer = setTimeout(() => {
        killed = true;
        proc.kill("SIGTERM");
        reject(
          new Error(
            `Python subprocess excedió el timeout de ${timeoutMs / 1000}s`,
          ),
        );
      }, timeoutMs);

      proc.on("error", (err) => {
        clearTimeout(timer);
        if (err.code === "ENOENT") {
          reject(
            new Error(
              `Python no encontrado en "${pythonPath}". Verifique la ruta o que Python esté en el PATH.`,
            ),
          );
        } else {
          reject(err);
        }
      });

      proc.on("close", (code) => {
        clearTimeout(timer);
        if (!killed) {
          resolve({ code, stdout, stderr });
        }
      });
    });
  }

  // ==========================================================
  // Pipeline: Altas (Formato 8) -> Duplicados TA2+IDC
  // ==========================================================

  async pipelineAltasDuplicados(argumentos) {
    this.log("Iniciando pipeline: Altas (Fmt 8) -> Duplicados TA2+IDC");

    const nombreProceso = "PIPELINE_ALTAS_DUPLICADOS";
    let registrosProcesados = 0;

    return new Promise(async (resolve) => {
      try {
        // --- Extraer parámetros del formulario ---
        const chromeExePath = argumentos?.formularioControl?.[0];
        const empresaCodes = argumentos?.formularioControl?.[1];
        const regimenManual = argumentos?.formularioControl?.[2];
        const pathSalidaBase = argumentos?.formularioControl?.[3];
        const pythonPath = argumentos?.formularioControl?.[4] || "python";
        const analisisA3Path = argumentos?.formularioControl?.[5];

        // --- Validaciones ---
        if (!chromeExePath || !fs.existsSync(chromeExePath)) {
          this.logErr("[INPUT] Ruta a chrome.exe no válida.");
          return resolve(false);
        }
        if (
          !empresaCodes ||
          typeof empresaCodes !== "string" ||
          !empresaCodes.trim()
        ) {
          this.logErr(
            "[INPUT] Códigos de empresa vacíos. Use formato: 00008 o 00008,01378",
          );
          return resolve(false);
        }
        if (
          !pathSalidaBase ||
          typeof pathSalidaBase !== "string" ||
          !pathSalidaBase.trim()
        ) {
          this.logErr("[INPUT] Ruta de salida no válida.");
          return resolve(false);
        }
        if (
          !analisisA3Path ||
          !fs.existsSync(
            path.join(path.normalize(analisisA3Path), "scripts", "modules", "nomv5e", "generador_informes.py"),
          )
        ) {
          this.logErr(
            '[INPUT] Ruta de analisis-a3 no válida. No se encontró "scripts/modules/nomv5e/generador_informes.py".',
          );
          return resolve(false);
        }

        const regimen4 = String(regimenManual || "0111")
          .replace(/\D/g, "")
          .padStart(4, "0");
        if (!/^\d{4}$/.test(regimen4)) {
          this.logErr("[INPUT] Régimen inválido. Debe ser 4 dígitos (ej: 0111).");
          return resolve(false);
        }

        // --- Normalizar códigos de empresa ---
        const codes = empresaCodes
          .split(",")
          .map((c) => c.trim())
          .filter((c) => c.length > 0)
          .map((c) => c.replace(/\D/g, "").padStart(5, "0"));

        if (codes.length === 0) {
          this.logErr("[INPUT] No se encontraron códigos de empresa válidos.");
          return resolve(false);
        }

        this.log(`Empresas a procesar: ${codes.join(", ")}`);

        // --- Paso 1: Generar XLSX con Python ---
        const tempDir = path.join(path.normalize(pathSalidaBase), "pipeline-temp");
        this.ensureDir(tempDir);

        const tempXlsx = path.join(tempDir, "altas_fmt8.xlsx");

        // Eliminar XLSX previo si existe (evitar datos stale)
        if (fs.existsSync(tempXlsx)) {
          fs.unlinkSync(tempXlsx);
          this.log("XLSX previo eliminado.");
        }

        const pyArgs = [
          path.join("scripts", "modules", "nomv5e", "generador_informes.py"),
          "--empresa",
          codes.join(","),
          "--formato",
          "8",
          "--output",
          tempXlsx,
        ];

        this.log("=== PASO 1: Generando listado Altas (Formato 8) ===");

        let pyResult;
        try {
          pyResult = await this._spawnPython(
            pythonPath,
            path.normalize(analisisA3Path),
            pyArgs,
          );
        } catch (err) {
          this.logErr(`Error ejecutando Python: ${err.message}`);
          return resolve(false);
        }

        if (pyResult.code !== 0) {
          this.logErr(
            `Python finalizó con código ${pyResult.code}. Stderr:\n${pyResult.stderr}`,
          );
          return resolve(false);
        }

        // --- Paso 2: Verificar XLSX ---
        if (!fs.existsSync(tempXlsx)) {
          this.logErr(
            "El script Python finalizó OK pero no generó el XLSX. " +
              "Posiblemente no hay trabajadores activos para las empresas indicadas.",
          );
          return resolve(false);
        }

        const stats = fs.statSync(tempXlsx);
        this.log(
          `XLSX generado: ${tempXlsx} (${(stats.size / 1024).toFixed(1)} KB)`,
        );

        // --- Paso 3: Ejecutar Duplicados TA2+IDC ---
        this.log("=== PASO 2: Ejecutando Duplicados TA2+IDC ===");

        const dup = new ProcesosDuplicados(
          this.pathToDbFolder,
          this.nombreProyecto,
          this.proyectoDB,
        );

        const dupArgs = {
          formularioControl: [
            chromeExePath,
            tempXlsx,
            regimen4,
            path.normalize(pathSalidaBase),
          ],
        };

        const dupResult = await dup.duplicadosTa2(dupArgs);

        registrosProcesados = codes.length;
        registrarEjecucion({
          nombreProceso,
          registrosProcesados,
          empresas: codes.map((c) => ({ codigo: c, nombre: "", registrosProcesados: 1 })),
        });

        this.log(
          `Pipeline finalizado. Resultado duplicados: ${dupResult}`,
        );
        return resolve(dupResult);
      } catch (err) {
        this.logErr(`Error inesperado en pipeline: ${err.message}`);
        console.error(err);
        return resolve(false);
      }
    });
  }

  // Alias para el dispatch de main.js (camelize de "PIPELINE ALTAS DUPLICADOS")
  async ["pIPELINEALTASDUPLICADOS"](argumentos) {
    return this.pipelineAltasDuplicados(argumentos);
  }
}

module.exports = ProcesosPipeline;
