const DEBUG = process.env.NODE_ENV === "dev";
console.log(process.env.NODE_ENV);
console.log("DESARROLLO: " + DEBUG);

const axios = require("axios");

const METRICS_ENDPOINT =
  "https://nodus-backend-production.up.railway.app/registrarEjecucion";

async function registrarEjecucion({
  nombreProceso,
  fechaEjecucion = new Date(),
  registrosProcesados = 0,
  empresas = [], // [{ codigo, nombre, registrosProcesados }]
}) {
  try {
    const payload = {
      nombreProceso,
      fechaEjecucion,
      registrosProcesados,
      ...(empresas.length > 0 && { empresas }),
    };

    if (!DEBUG) {
      await axios.post(METRICS_ENDPOINT, payload);
    }
  } catch (err) {
    console.error("Error enviando métricas:", err.message);
  }
}

function agruparPorEmpresa(clientes, camposCodigo = ["cod_empresa"], camposNombre = ["nombre_empresa"]) {
  const mapa = {};
  for (const c of clientes) {
    const codigo = camposCodigo.map((k) => c[k]).find((v) => v != null) ?? "";
    const nombre = camposNombre.map((k) => c[k]).find((v) => v != null) ?? "";
    const key = String(codigo);
    if (!mapa[key]) {
      mapa[key] = { codigo: String(codigo), nombre: String(nombre), registrosProcesados: 0 };
    }
    mapa[key].registrosProcesados += 1;
  }
  return Object.values(mapa);
}

module.exports = { registrarEjecucion, agruparPorEmpresa };
