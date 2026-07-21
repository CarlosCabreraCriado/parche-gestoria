// Reescribe la hoja LEEME de una plantilla A3 de pagos ya generada, dejando el
// resto del libro intacto.
//
// Existe porque la columna IMPORTE (v2) se añadió a mano sobre una plantilla v1:
// las hojas de modelo quedaron bien (centinela A3PAGOS v2 + columna IMPORTE),
// pero el LEEME siguió siendo el de la v1, documentando una Zona A de A–F, una
// Zona B desde H, ninguna columna IMPORTE y un centinela "A3PAGOS v1" que
// contradice el que llevan las hojas. Regenerar la plantilla entera no es opción
// hoy: las anclas legacy por número de fila de los specs están calibradas contra
// 4PAGOS2026 (1).xls y `migrarPlantilla.js` falla contra el 4PAGOS del 2T.
//
// El texto NO se duplica aquí: se reutiliza `writeLeeme` de migrarPlantilla.js,
// así que este script y una plantilla recién generada dicen exactamente lo mismo.
//
// Uso: node actualizarLeeme.js <plantilla.xlsx> [salida.xlsx]
//      (sin salida, reescribe la plantilla en el sitio)

const path = require("path");
const XlsxPopulate = require("xlsx-populate");
const { writeLeeme, SPECS, VERBATIM } = require("./migrarPlantilla");

const HOJA_LEEME = "LEEME";

// "Generada el 2026-07-09 a partir de "4PAGOS2026 (1).xls"." — la escribe
// `writeLeeme` y es el único dato del LEEME que no se puede deducir del libro.
const GENERADA_RX = /^Generada el (.+?) a partir de "(.+)"\.\s*$/;

// De qué archivo y cuándo salió la plantilla. Se conserva tal cual: el LEEME se
// actualiza, pero los datos siguen siendo los del 4PAGOS de origen y decir otra
// cosa haría creer que la plantilla se ha refrescado.
function leerProcedencia(sheet) {
  for (let fila = 1; fila <= 5; fila++) {
    const v = sheet.cell(fila, 1).value();
    const m = typeof v === "string" ? v.match(GENERADA_RX) : null;
    if (m) return { fechaGen: m[1], inputName: m[2] };
  }
  throw new Error(
    `La hoja ${HOJA_LEEME} no tiene la línea "Generada el ... a partir de ..." en las primeras 5 filas: ` +
      `¿es una plantilla A3 de pagos generada por migrarPlantilla.js?`
  );
}

async function main() {
  const [input, output] = process.argv.slice(2);
  if (!input) throw new Error("Uso: node actualizarLeeme.js <plantilla.xlsx> [salida.xlsx]");
  const destino = output || input;

  const wb = await XlsxPopulate.fromFileAsync(path.normalize(input));

  const leemeViejo = wb.sheet(HOJA_LEEME);
  if (!leemeViejo) throw new Error(`El libro no tiene hoja '${HOJA_LEEME}'.`);
  const { fechaGen, inputName } = leerProcedencia(leemeViejo);

  // Solo las hojas que esta plantilla trae de verdad: si se generó con --hojas,
  // el LEEME no debe prometer hojas que no están.
  const presentes = new Set(wb.sheets().map((s) => s.name()));
  const specsSel = SPECS.filter((s) => presentes.has(s.hoja));
  const verbatimSel = VERBATIM.filter((v) => presentes.has(v.hoja));

  // Hoja nueva en lugar de sobrescribir celda a celda: el LEEME viejo tiene
  // negritas en filas que en el nuevo texto son otra cosa, y esos estilos
  // sobrevivirían a la reescritura.
  wb.deleteSheet(leemeViejo);
  const leeme = wb.addSheet(HOJA_LEEME, 0);
  writeLeeme(leeme, fechaGen, inputName, specsSel, verbatimSel);

  // Rastro de que el LEEME es más nuevo que el resto del libro: sin esto, la
  // fecha de generación de arriba parece desmentir lo que el texto documenta.
  const ultima = leeme.usedRange().endCell().rowNumber();
  const hoy = new Date().toISOString().slice(0, 10);
  leeme
    .cell(ultima + 2, 1)
    .value(`LEEME actualizado el ${hoy}: documenta la columna IMPORTE (plantilla A3PAGOS v2).`)
    .style({ italic: true, fontColor: "595959" });
  leeme
    .cell(ultima + 3, 1)
    .value("Los datos de las hojas de modelo siguen siendo los del archivo de origen indicado arriba.")
    .style({ italic: true, fontColor: "595959" });

  leeme.active(true);
  await wb.toFileAsync(path.normalize(destino));
  console.log(`LEEME actualizado en ${destino}`);
  console.log(`  Procedencia conservada: ${fechaGen} · ${inputName}`);
  console.log(`  Hojas de modelo documentadas: ${specsSel.map((s) => s.hoja).join(", ") || "ninguna"}`);
  console.log(`  Hojas copiadas tal cual: ${verbatimSel.map((v) => v.hoja).join(", ") || "ninguna"}`);
}

main().catch((err) => {
  console.error(err.message);
  process.exit(1);
});
