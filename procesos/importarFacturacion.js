const path = require("path");
const fs = require("fs");
const { registrarEjecucion } = require("../metricas");

const { Mapeos } = require("./importarFacturacion/mapeos");
const nominas = require("./importarFacturacion/nominas");
const notificaciones = require("./importarFacturacion/notificaciones");
const tramites = require("./importarFacturacion/tramites");
const generateTraspaso = require("./importarFacturacion/generateTraspaso");
const {
  _toDate,
  ensureDir,
  stampYYYYMMDDHHmm,
} = require("./importarFacturacion/utils");

const DEFAULT_MAPEOS_DIR = path.join(__dirname, "inputs", "mapeos");
const DEFAULT_PLANTILLA = "M:\\A3\\A3GESW\\PLANTILLA DE TRASPASO DE DATOS A A3GES.XLSX";

class ProcesosImportarFacturacion {
  constructor(pathToDbFolder, nombreProyecto, proyectoDB) {
    this.pathToDbFolder = pathToDbFolder;
    this.nombreProyecto = nombreProyecto;
    this.proyectoDB = proyectoDB;
    this.TAG = "[IMPORTAR-FACTURACION]";
  }

  log(msg, ...rest) {
    console.log(`${this.TAG} ${msg}`, ...rest);
  }
  logWarn(msg, ...rest) {
    console.warn(`${this.TAG} ${msg}`, ...rest);
  }
  logErr(msg, ...rest) {
    console.error(`${this.TAG} ${msg}`, ...rest);
  }

  _parseArgs(argumentos, tipo) {
    // Orden esperado (procesos.configuracion.ts):
    // [0] archivoInput, [1] rutaSalida, [2] fechaFactura,
    // [3] carpetaMapeos, [4] rutaPlantillaA3
    const c = argumentos?.formularioControl || [];
    return {
      archivoInput: c[0],
      rutaSalida: c[1],
      fechaFactura: c[2],
      carpetaMapeos: c[3] && String(c[3]).trim() ? String(c[3]) : DEFAULT_MAPEOS_DIR,
      rutaPlantillaA3: c[4] && String(c[4]).trim() ? String(c[4]) : DEFAULT_PLANTILLA,
      tipo,
    };
  }

  _validate(args) {
    if (!args.archivoInput || !fs.existsSync(args.archivoInput)) {
      this.logErr(`Archivo de entrada no válido: ${args.archivoInput}`);
      return false;
    }
    if (!args.rutaSalida || !String(args.rutaSalida).trim()) {
      this.logErr("Ruta de salida vacía.");
      return false;
    }
    if (!fs.existsSync(args.carpetaMapeos)) {
      this.logErr(`Carpeta de mapeos no encontrada: ${args.carpetaMapeos}`);
      return false;
    }
    if (!fs.existsSync(args.rutaPlantillaA3)) {
      this.logErr(`Plantilla A3 no encontrada: ${args.rutaPlantillaA3}`);
      return false;
    }
    return true;
  }

  async _run(tipo, transformer, argumentos) {
    return new Promise(async (resolve) => {
      try {
        const args = this._parseArgs(argumentos, tipo);
        if (!this._validate(args)) return resolve(false);

        const now = new Date();
        const stamp = stampYYYYMMDDHHmm(now);
        const outDir = path.join(
          path.normalize(args.rutaSalida),
          `importacion-${tipo}-${stamp}`
        );
        ensureDir(outDir);

        const fechaFactura = _toDate(args.fechaFactura, now);

        this.log(`Iniciando importación tipo=${tipo}`);
        this.log(`  Input:    ${args.archivoInput}`);
        this.log(`  Output:   ${outDir}`);
        this.log(`  Mapeos:   ${args.carpetaMapeos}`);
        this.log(`  Plantilla:${args.rutaPlantillaA3}`);

        // 1. Cargar mapeos
        const mapeos = await Mapeos.fromDir(args.carpetaMapeos);
        const summary = mapeos.summary();
        this.log(`Mapeos cargados: ${JSON.stringify(summary)}`);
        for (const w of mapeos.allWarnings()) this.logWarn(w);

        // 2. Transformar
        const result = await transformer.transform(
          path.normalize(args.archivoInput),
          mapeos,
          outDir,
          fechaFactura
        );
        this.log(`Transformación: ${JSON.stringify(result)}`);

        // 3. Generar XLSX de traspaso
        const xlsxOut = path.join(outDir, `TRASPASO_A3_${tipo}_${stamp}.xlsx`);
        this.log(`Inyectando plantilla A3 → ${xlsxOut}`);
        const gen = await generateTraspaso.run(
          outDir,
          xlsxOut,
          path.normalize(args.rutaPlantillaA3)
        );
        this.log(`Plantilla generada: ${gen.total_rows} filas en ${gen.sheets.length} hoja(s).`);

        // 4. Resumen JSON
        const resumen = {
          tipo,
          timestamp: now.toISOString(),
          input: args.archivoInput,
          output_dir: outDir,
          xlsx_generado: xlsxOut,
          mapeos: summary,
          transform: result,
          traspaso: {
            output: gen.output,
            total_rows: gen.total_rows,
            sheets: gen.sheets,
          },
        };
        fs.writeFileSync(
          path.join(outDir, "resumen.json"),
          JSON.stringify(resumen, null, 2),
          "utf8"
        );

        // 5. Métricas
        registrarEjecucion({
          nombreProceso: `IMPORTAR_${tipo.toUpperCase()}`,
          registrosProcesados: result.conceptos,
        });

        this.log(
          `Fin importación tipo=${tipo}. Conceptos=${result.conceptos}, Incidencias=${result.incidencias}, WarningsQC=${result.warnings_qc}`
        );
        return resolve(true);
      } catch (err) {
        this.logErr(`Error en importación: ${err.message}`);
        console.error(err);
        return resolve(false);
      }
    });
  }

  async importarNominas(argumentos) {
    return this._run("nominas", nominas, argumentos);
  }

  async importarNotificaciones(argumentos) {
    return this._run("notificaciones", notificaciones, argumentos);
  }

  async importarTramites(argumentos) {
    return this._run("tramites", tramites, argumentos);
  }
}

module.exports = ProcesosImportarFacturacion;
