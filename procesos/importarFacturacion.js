const path = require("path");
const fs = require("fs");
const { registrarEjecucion } = require("../metricas");

const { Mapeos } = require("./importarFacturacion/mapeos");
const nominas = require("./importarFacturacion/nominas");
const notificaciones = require("./importarFacturacion/notificaciones");
const tramites = require("./importarFacturacion/tramites");
const generateTraspaso = require("./importarFacturacion/generateTraspaso");
const {
  ensureDir,
  stampYYYYMMDDHHmm,
} = require("./importarFacturacion/utils");

// Plantilla A3: siempre la misma, versionada en el propio proyecto (no configurable por el usuario).
const PLANTILLA_A3 = path.join(
  __dirname,
  "inputs",
  "PLANTILLA DE TRASPASO DE DATOS A A3GES.xlsx"
);

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
    // [0] archivoInput, [1] rutaSalida, [2] archivoMapeos
    const c = argumentos?.formularioControl || [];
    return {
      archivoInput: c[0],
      rutaSalida: c[1],
      archivoMapeos: c[2],
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
    if (!args.archivoMapeos || !fs.existsSync(args.archivoMapeos)) {
      this.logErr(`Archivo de mapeos no encontrado: ${args.archivoMapeos}`);
      return false;
    }
    if (!fs.existsSync(PLANTILLA_A3)) {
      this.logErr(`Plantilla A3 no encontrada en el proyecto: ${PLANTILLA_A3}`);
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

        this.log(`Iniciando importación tipo=${tipo}`);
        this.log(`  Input:    ${args.archivoInput}`);
        this.log(`  Output:   ${outDir}`);
        this.log(`  Mapeos:   ${args.archivoMapeos}`);
        this.log(`  Plantilla:${PLANTILLA_A3}`);

        // 1. Cargar mapeos
        const mapeos = await Mapeos.fromFile(args.archivoMapeos);
        const summary = mapeos.summary();
        this.log(`Mapeos cargados: ${JSON.stringify(summary)}`);
        for (const w of mapeos.allWarnings()) this.logWarn(w);

        // 2. Transformar
        const result = await transformer.transform(
          path.normalize(args.archivoInput),
          mapeos,
          outDir
        );
        this.log(`Transformación: ${JSON.stringify(result)}`);

        // 3. Generar XLSX de traspaso
        const xlsxOut = path.join(outDir, `TRASPASO_A3_${tipo}_${stamp}.xlsx`);
        this.log(`Inyectando plantilla A3 → ${xlsxOut}`);
        const gen = await generateTraspaso.run(outDir, xlsxOut, PLANTILLA_A3);
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
