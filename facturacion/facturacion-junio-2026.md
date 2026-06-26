# Informe de trabajo — Junio 2026

**Cliente:** Del Castillo Asesores
**Período:** 1 – 25 de junio de 2026
**Proyectos:** Backend (analisis-a3), Aplicación de gestión (parche-gestoria)

---

## Resumen de conceptos

| Concepto / Subconcepto                                                                              | Horas    |
| --------------------------------------------------------------------------------------------------- | -------- |
| **1. Análisis del módulo de facturación GESW**                                                     | **20 h** |
| — 1.1. Integración del módulo y lectores binarios validados                                         | 6 h      |
| — 1.2. Ingeniería inversa del layout FAC/COF: fórmula de importe y Proformas                       | 14 h     |
| **2. Adaptación de la plantilla de importación de facturación**                                     | **5 h**  |
| — 2.1. Guía operativa de importación para empresa 14                                                | 2 h      |
| — 2.2. Adaptador Sección 3 y validación import → export                                            | 3 h      |
| **3. Análisis de clientes y expedientes**                                                           | **3 h**  |
| — 3.1. Listado cruzado FAC ↔ EXPED y generación de CSVs                                           | 3 h      |
| **4. Art.42 — mejoras de estabilidad y rendimiento**                                                | **3 h**  |
| — 4.1. Corrección de falso error, detección de resultado y optimización de tiempos                 | 3 h      |
| **Total**                                                                                            | **31 h** |

---

## Detalle por concepto

### 1. Análisis del módulo de facturación GESW (~20 h)

Análisis estructurado e ingeniería inversa del módulo de facturación de A3GES para la empresa 14, con el objetivo de replicar con exactitud el listado `A3DatosFact` desde los ficheros ISAM binarios.

#### 1.1. Integración del módulo y lectores binarios validados (~6 h)

El módulo `gesw` se ha migrado desde el repositorio de exploración al proyecto principal `analisis-a3`, siguiendo el patrón del módulo de nóminas:

- Lectores binarios validados de los ficheros ISAM: `gfFCUOTA`, `gfFEXPED`, `gfFAC`, `gfCOF`
- Schemas de familias A y B con tipos comunes en `family_schemas.py`
- Round-trip `A3DatosFact` + comparativa automática contra el XLS oficial
- Consolidador de `A3DiarioVentas*.xls` y generador de plantilla de traspaso
- Documentación completa: análisis del módulo, mapeo de origen de datos, cobertura por tabla y arquitectura de maestros
- Consolidación de fixtures bajo `gesw/fixtures/{export,test_import,samples}`

#### 1.2. Ingeniería inversa del layout FAC/COF: fórmula de importe y Proformas (~14 h)

Decodificación byte a byte de `GFF14FAC.DAT` (cabeceras, 544 bytes/registro) y `GFF14COF.DAT` (líneas, 236 bytes/registro).

**Estructura de claves y resolución de NIFs:**

- Cadena de claves foráneas `FK_3B + fecha_documento` enlaza cabeceras con líneas
- Corrección de offset NIF (posición 15–25) para cubrir NIFs de 10 caracteres
- Resultado base: **2.558/2.558 líneas** y **444/444 documentos** replicados (100% de cobertura)

**Hito A — fórmula de importe:** `precio × unidades × (1 − descuento)`

Cierre del 19% de brecha en el total de importes. Tres campos descubiertos en `GFF14COF`:

- **Unidades:** offsets `+97..+99`, uint24 LE / 100
- **Descuento %:** offsets `+100..+101`, uint16 LE / 100
- **id_seq:** offsets `+2..+6`, uint40 BE — orden cronológico para partir lotes compartidos entre facturas

Resultados tras Hito A: match exacto por tupla **68,3% → 97,8%** · total importe **113.264 € → 141.141 €** (99,1% de 142.399 €)

**Hito B — Proformas:**

Las 30 líneas de tipo «Proforma» son líneas `COF` con `fecha_documento` rellena pero sin cabecera `FAC` para la pareja `(FK, fdoc)`. Afecta a 5 clientes concretos.

Resultados acumulados A + B: **99,0% de match** · **142.143 € de 142.399 € (99,8%)** · **607/622 documentos** con suma idéntica

---

### 2. Adaptación de la plantilla de importación de facturación (~5 h)

Trabajo orientado a preparar los datos para su importación al nuevo sistema, partiendo del análisis anterior.

#### 2.1. Guía operativa de importación para empresa 14 (~2 h)

Documento paso a paso para cumplimentar la plantilla oficial de traspaso con las **530 líneas pendientes** (176 clientes, 176 expedientes, **23.269,70 €**), reutilizando las reglas de tipos, mapeo IGIC y criterios de idempotencia validados en el PoC de empresa 26.

#### 2.2. Adaptador Sección 3 y validación import → export (~3 h)

- `build_traspaso_seccion3.py`: convierte el listado «Sección 3» (XLSX, 32 líneas de servicio) a los 3 CSVs que consume el generador de plantilla de traspaso — Clientes, Expedientes y Conceptos Pendientes de Facturar
- `validate_import_vs_export.py`: verifica que cada registro importado aparezca correctamente en el `A3DatosFact` de salida, cerrando el ciclo de validación extremo a extremo

---

### 3. Análisis de clientes y expedientes (~3 h)

- Nuevo subcomando `datosfact list-expedientes`: cruce de `GFF14FAC.DAT` con `gfFEXPED.DAT` por FK; genera CSV con NIF, razón social, expediente y responsable — **945 clientes, cobertura 100%**
- Excel de análisis con tres hojas: datos completos, estructura y NIFs con expediente duplicado
- Identificación de los 176 clientes con facturación pendiente de importar

---

### 4. Art.42 — mejoras de estabilidad y rendimiento (~3 h)

Correcciones y optimizaciones sobre el proceso Art.42 implantado en mayo:

- **Corrección de falso error:** el proceso marcaba como error casos en los que el alta ya había sido realizada; ahora se detecta correctamente como éxito
- **Detección de resultado mejorada:** mejor identificación de respuestas de validación y mensajes de error del portal
- **Eliminación de espera innecesaria:** suprimida la espera de navegación tras la confirmación, acortando el tiempo total por empresa
- Versión de la aplicación: **0.93.1**

---

_Documento generado el 25 de junio de 2026_
