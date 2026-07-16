# Informe de migración — 4PAGOS → plantilla A3

- **Origen:** `4PAGOS2026_2T.xls`
- **Generado:** `4PAGOS2026_2T - PLANTILLA A3 v2.xlsx` (2026-07-16)
- **Hojas migradas:** 111

## Resumen por hoja

| Hoja | Filas SI | Filas NO | Filas REVISAR | Sin expte | Notas conservadas | P1 | P2 | P3 | P4 |
|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|
| 111 | 205 | 18 | 0 | 0 | 8 | 187 | 188 | 3 | 3 |

P1–P4 es un **diagnóstico del archivo original**: cuenta las filas con FACTURAR=SI que traían el periodo relleno. No es una regla de facturación — el importador no lee esas columnas (ver LEEME).

## Columna FRECUENCIA

| Hoja | Frecuencia por defecto | Excepciones por fila detectadas |
|---|---|---:|
| 111 | TRIMESTRAL | 13 |

En la hoja 111, 13 empresas tenían "MENSUAL" escrito en la columna BAJA del original — se promovió a FRECUENCIA=MENSUAL y F.BAJA se dejó intacta en Zona B.

## Columna OBSERVACIONES

Unifica bajo un único nombre la columna de notas de cada hoja (OBS/OBSERV/OBSERVACIONES según el modelo).
Todas las hojas traían columna de notas en el original.

## Secciones aplicadas (decisión FACTURAR heredada del bloque del original)

Cada sección se localiza por el texto de su título en el archivo de origen. Se listan la fila donde se resolvió y el texto crudo que hizo match: es el rastro para comprobar que el bloque es el que se esperaba.

| Hoja | Sección | → FACTURAR | Fila en el origen | Texto encontrado |
|---|---|---|---:|---|
| 111 | MODELO 111 TRIMESTRAL _(inicial)_ | SI | 3 | — |
| 111 | BAJA ASESORIA 2026 | NO | 208 | `BAJA ASESORIA 2026` |
| 111 | EMPRESAS REALIZAN ELLOS MODELOS 111 - 190 | NO | 226 | `EMPRESAS REALIZAN ELLOS MODELOS 111 - 190` |

## Celdas de periodo con texto (revisar FACTURAR a mano)

Diagnóstico del archivo original: el importador no lee P1–P4, así que estas celdas no bloquean nada por sí solas. Importan porque contradicen la fila: una empresa con FACTURAR=SI cuyo periodo pone "NO" o "??" es candidata a poner en REVISAR o NO. La decisión es manual — el script no la toma.

En filas **SI/REVISAR**: 18

| Hoja | Fila original | Periodo | Valor | Empresa | FACTURAR |
|---|---|---|---|---|---|
| 111 | 20 | P1 | ?? | VALTESOL PROMOCIONES SL | SI |
| 111 | 20 | P2 | NO | VALTESOL PROMOCIONES SL | SI |
| 111 | 20 | P3 | NO | VALTESOL PROMOCIONES SL | SI |
| 111 | 20 | P4 | NO | VALTESOL PROMOCIONES SL | SI |
| 111 | 43 | P3 | NO | TODOCEMENTO SL | SI |
| 111 | 43 | P4 | NO | TODOCEMENTO SL | SI |
| 111 | 136 | P1 | NO | RODRIGUEZ VARGAS SOFIA | SI |
| 111 | 148 | P2 | NO | LA PINOCHERIA SL | SI |
| 111 | 148 | P3 | NO | LA PINOCHERIA SL | SI |
| 111 | 148 | P4 | NO | LA PINOCHERIA SL | SI |
| 111 | 150 | P1 | NO | LOPEZ GONZALEZ ORLEXNYS | SI |
| 111 | 150 | P2 | NO | LOPEZ GONZALEZ ORLEXNYS | SI |
| 111 | 199 | P1 | NO | VULKASOFT SLU | SI |
| 111 | 199 | P2 | NO | VULKASOFT SLU | SI |
| 111 | 205 | P1 | NO | ASOC CULTURAL COLOMBOFILA ADEXE | SI |
| 111 | 206 | P1 | NO | DENIZ SCHLEUPNER, JONATHAN MIGUEL | SI |
| 111 | 207 | P1 | NO | ENIDENSSON SL | SI |
| 111 | 207 | P2 | NO | ENIDENSSON SL | SI |

## Filas de cliente sin EXPTE (migradas con EXPTE vacío — no facturables hasta asignarlo)

Ninguna.

## Hojas copiadas tal cual (sin Zona A)

Ninguna.

## Hojas del original no migradas

- **No seleccionadas** (tienen spec, excluidas con `--hojas`): 130, 420, 303-349-369-340, 123-193, 202, 210, 131, 421, 115, alcohol, 184, EXP
- **Vacías en el original**: 1, Hoja1, Hoja2, Hoja3
- **Con datos y sin spec** (se pierden si hacen falta): ninguna

## Columnas descartadas

Ninguna: las columnas marcadas para descartar en los specs no traían datos en este archivo.
