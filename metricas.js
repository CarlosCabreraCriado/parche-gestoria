//DETERMINA SI ES DESARROLLO O PRODUCCION:
const DEBUG = process.env.NODE_ENV === "dev"; //Verifica si esta en producción
console.log(process.env.NODE_ENV);
console.log("DESARROLLO: " + DEBUG);

// metricsClient.js
const axios = require("axios"); // npm install axios

const METRICS_ENDPOINT =
  "https://nodus-backend-production.up.railway.app/registrarEjecucion";

async function registrarEjecucion({
  nombreProceso,
  fechaEjecucion = new Date(),
  registrosProcesados,
}) {
  try {
    const payload = {
      nombreProceso,
      fechaEjecucion, // se enviará en ISO (axios lo serializa)
      registrosProcesados,
    };

    const headers = {};

    if (!DEBUG) {
      await axios.post(METRICS_ENDPOINT, payload, { headers });
    }
  } catch (err) {
    // Importante: nunca revientes el proceso sólo por fallo al enviar métricas
    console.error("Error enviando métricas:", err.message);
  }
}

module.exports = { registrarEjecucion };
