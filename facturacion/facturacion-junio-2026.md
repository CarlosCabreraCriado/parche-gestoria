# Informe de trabajo — Junio 2026

**Cliente:** Del Castillo Asesores
**Período:** 1 – 25 de junio de 2026
**Proyectos:** Backend (analisis-a3), Aplicación de gestión (parche-gestoria)

---

## Resumen de conceptos

| Concepto / Subconcepto                                                                         | Horas    |
| ---------------------------------------------------------------------------------------------- | -------- |
| **1. Módulo de análisis de facturación A3GES — empresa 14**                                   | **28 h** |
| — 1.1. Integración del módulo gesw en analisis-a3 y consolidación de fixtures                 | 6 h      |
| — 1.2. Ingeniería inversa del layout ISAM FAC/COF y replicación del listado A3DatosFact       | 10 h     |
| — 1.3. Hito A — fórmula de importe: precio × unidades × (1 − descuento)                      | 6 h      |
| — 1.4. Hito B — identificación y tratamiento de Proformas                                     | 3 h      |
| — 1.5. Listado de clientes con expedientes y guía operativa de importación                    | 3 h      |
| **2. Art.42 — mejoras de estabilidad y rendimiento**                                           | **3 h**  |
| — 2.1. Corrección de falso error, detección de resultado y optimización de tiempos            | 3 h      |
| **Total**                                                                                       | **31 h** |

---

## Detalle por concepto

### 1. Módulo de análisis de facturación A3GES — empresa 14 (~28 h)

Desarrollo completo del módulo de análisis e ingeniería inversa de los ficheros ISAM de A3GES para la empresa 14, con el objetivo de replicar con exactitud el listado de facturación `A3DatosFact` y preparar la migración de los datos al nuevo sistema.

#### 1.1. Integración del módulo gesw en analisis-a3 y consolidación de fixtures (~6 h)

El módulo `gesw` (análisis de A3GES) se ha migrado desde el repositorio de exploración (`A3GESW_analisis`) al proyecto principal `analisis-a3`, siguiendo el patrón ya establecido por el módulo de nóminas:

- Lectores binarios validados de los ficheros ISAM: `gfFCUOTA`, `gfFEXPED`, `gfFAC`, `gfCOF`
- Schemas de familias A y B en `family_schemas.py` con tipos comunes entre ambas
- Round-trip `A3DatosFact` + comparativa automática contra el XLS oficial
- Consolidador de `A3DiarioVentas*.xls` y generador de plantilla de traspaso
- Documentación completa: análisis del módulo, mapeo de origen de datos, cobertura por tabla y arquitectura de maestros
- Consolidación de todos los fixtures del proyecto anterior bajo `gesw/fixtures/{export,test_import,samples}`

#### 1.2. Ingeniería inversa del layout ISAM y replicación del listado A3DatosFact (~10 h)

Decodificación byte a byte de los ficheros binarios `GFF14FAC.DAT` (cabeceras de factura, 544 bytes/registro) y `GFF14COF.DAT` (líneas de factura, 236 bytes/registro):

- Cadena de claves foráneas `FK_3B + fecha_documento` que enlaza cabeceras con líneas
- Resolución de NIFs desde el índice de clientes `cliente_by_fk` (corrección de offset NIF: posición 15–25 para cubrir NIFs de 10 caracteres)
- Resultado final: **2.558/2.558 líneas** y **444/444 documentos** replicados (100% de cobertura) contra el periodo abril–junio 2026

#### 1.3. Hito A — fórmula de importe: precio × unidades × (1 − descuento) (~6 h)

Cierre del 19% de brecha en el total de importes. Tres descubrimientos byte-perfect en el layout de `GFF14COF`:

- **Precio base:** offset `+77` (ya conocido)
- **Unidades:** offsets `+97..+99`, uint24 LE / 100 (ejemplo: 184400 → 1844 nóminas)
- **Descuento %:** offsets `+100..+101`, uint16 LE / 100 (800 → 8%, 5000 → 50%)
- **id_seq:** offsets `+2..+6`, uint40 BE, orden cronológico de inserción (necesario para partir lotes compartidos)

Resultados vs XLS oficial tras Hito A:
- Match exacto por tupla: **68,3% → 97,8%**
- Total importe: **113.264 € → 141.141 €** (de 142.399 €, 99,1%)
- Documentos con suma idéntica: **291/622 → 602/622**

#### 1.4. Hito B — identificación y tratamiento de Proformas (~3 h)

Las 30 líneas de tipo «Proforma» del listado oficial corresponden 1:1 a líneas en `COF` que tienen `fecha_documento` rellena pero no existe cabecera `FAC` para esa pareja `(FK, fdoc)`. Afecta a 5 clientes concretos.

Resultados acumulados tras Hito A + B:
- Filas con NIF: **2.521 → 2.553** (de 2.558 = **99,8%**)
- Match exacto por tupla: **97,8% → 99,0%**
- Total importe: **141.141 € → 142.143 €** (de 142.399 € = **99,8%**)
- Documentos con suma idéntica: **602/622 → 607/622**

#### 1.5. Listado de clientes con expedientes y guía operativa de importación (~3 h)

- Nuevo subcomando `datosfact list-expedientes`: cruce de `GFF14FAC.DAT` con `gfFEXPED.DAT` por FK; genera CSV con NIF, razón social, expediente y responsable — **945 clientes, cobertura 100%**
- Excel con tres hojas: datos, análisis de estructura y NIFs con expediente duplicado
- Guía operativa paso a paso para cumplimentar la plantilla oficial de traspaso con las **530 líneas pendientes** (176 clientes, 176 expedientes, **23.269,70 €**), reutilizando las reglas validadas en el PoC de empresa 26

---

### 2. Art.42 — mejoras de estabilidad y rendimiento (~3 h)

Correcciones y optimizaciones sobre el proceso Art.42 implantado en mayo:

- **Corrección de falso error:** el proceso marcaba como error casos en los que el alta ya había sido realizada previamente; ahora se detecta correctamente como éxito
- **Detección de resultado mejorada:** mejor identificación de respuestas de validación y mensajes de error del portal, reduciendo falsos negativos
- **Eliminación de espera innecesaria:** se ha suprimido la espera de navegación tras la confirmación del alta, acortando el tiempo total del proceso por empresa
- Versión de la aplicación: **0.93.1**

---

_Documento generado el 25 de junio de 2026_
