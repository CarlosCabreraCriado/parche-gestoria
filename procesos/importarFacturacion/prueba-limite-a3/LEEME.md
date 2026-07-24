# Prueba: límite de caracteres de A3 (Descripción y Descripción Ampliada)

## Objetivo
Medir cuántos caracteres admiten en A3 los campos **Descripción** y **Descripción
Ampliada** de un concepto pendiente de facturar, para saber a partir de qué
longitud A3 los recorta.

## Qué contiene `PRUEBA_LIMITE_A3_tramites.xlsx`
Es un traspaso normal (mismo formato que genera la app), con **6 líneas de
prueba** en la hoja *Conceptos Pendientes Facturar*. Todas usan claves reales y
válidas de trámites para que A3 las acepte:

- Empresa `14`, Cliente `01378`, Concepto `1.159`, Expediente `01378-00407417`
- Importe **0,01 €** (simbólico: son líneas de PRUEBA, **no** para facturar)

La Descripción y la Descripción Ampliada de cada línea son una **regla
medidora**: cada 10 caracteres aparece impresa la **posición** (010, 020, 030 …).
Los rellenos son distintos en cada columna para no confundirlas:

- **Descripción** → puntos: `.......010.......020.......030 …`
- **Descripción Ampliada** → equis: `xxxxxxx010xxxxxxx020xxxxxxx030 …`

Las 6 líneas tienen longitudes **30, 80, 160, 240, 320 y 400** caracteres.

## Cómo leer el resultado
1. Importa el archivo en A3 (entorno de pruebas si es posible).
2. Abre las líneas importadas y **mira el último número que quede visible** en
   cada campo. Ese número **es el límite** (o el múltiplo de 10 más cercano).
   - Si tras el último número hay puntos/equis sueltas, súmalas: p. ej.
     `…200.....` = 200 + 5 = **205 caracteres**.
3. Pásame lo que veas (o exporta las líneas y mándame el Excel). Con eso fijamos
   el límite exacto de cada campo.

> La línea de 30 caracteres debería entrar entera (verás `…030` completo): sirve
> de control de que la regla se lee bien.

## Resultado (2026-07-24)
Importado y re-extraído desde A3 (`A3DatosFact_tamaño.xls`):

| Descripción enviada | En A3 | Ampliada enviada | En A3 |
|---:|---:|---:|---:|
| 30 | 30 (+relleno a 50) | 30 | 30 ✓ |
| 80 | **50** ✂ | 80 | 80 ✓ |
| 160 | **50** ✂ | 160 | 160 ✓ |
| 240 | **50** ✂ | 240 | 240 ✓ |
| 320 | **50** ✂ | 320 | 320 ✓ |
| 400 | **50** ✂ | 400 | 400 ✓ |

- **Descripción → 50 caracteres** (campo de ancho fijo: rellena con espacios
  hasta 50 y corta el resto).
- **Descripción Ampliada → ≥ 400** (sin recorte hasta 400).

## Arreglo implementado (2026-07-24)
Regla común en `utils.js` (`repartirDescripcion`) aplicada a los 4 importadores:

- **Descripción** = concepto recortado a **50** caracteres **por palabra** (nunca
  parte una palabra).
- **Descripción Ampliada** = texto ÍNTEGRO del concepto si no cupo en 50 + el
  resto del detalle (nombre del trabajador / razón social, observación, fecha).
  Se normalizan espacios y saltos de línea a un solo espacio.

Así no se pierde nada y la línea no acaba a mitad de palabra. La fecha del trámite
ya no va en la Descripción (A3 tiene su columna Fecha); se conserva en la
Ampliada.

`VERIFICACION_tramites_50.xlsx` (en esta carpeta) tiene 4 líneas reales de
trámites —incluidos los 2 conceptos que antes se cortaban (Accidente 95 car.,
Modificación de contrato 93 car.)— con importe 0,01 € de PRUEBA. Al importarlo,
la Descripción entra entera (≤50, corte por palabra) y el detalle completo queda
en la Ampliada. Verificado sobre la corrida real: máx Descripción = 46, ninguna
línea supera 50.
