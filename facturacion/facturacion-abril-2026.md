# Informe de trabajo — Abril 2026

**Cliente:** Del Castillo Asesores
**Período:** 1 – 30 de abril de 2026
**Proyectos:** Aplicación de gestión (parche-gestoria), Backend (Nodus-Backend), Análisis A3

---

## Resumen de conceptos

| Concepto / Subconcepto                                                             | Horas    |
| ---------------------------------------------------------------------------------- | -------- |
| **1. Unificación y mejora del proceso de Certificados de estar al corriente**      | **18 h** |
| — 1.1. Proceso unificado SS + ATC + ITA con selector multi-organismo               | 10 h     |
| — 1.2. Correcciones de robustez (reintentos, CDP, detección PDF)                   | 8 h      |
| **2. Nuevos informes en el sistema de Análisis A3**                                | **16 h** |
| — 2.1. Informes Enfermedad, Accidentes, Embargos y Extranjeros (Fmt. 6, 7, 10, 15) | 12 h     |
| — 2.2. Consola unificada de informes e integración en la aplicación                | 4 h      |
| **3. Módulo de extracción de nóminas desde A3** _(en curso)_                       | **12 h** |
| — 3.1. Diseño, Excel maestro de configuración y modelos de datos                   | 5 h      |
| — 3.2. Extracción directa desde ficheros DAT y verificación de resultados          | 7 h      |
| **4. Adaptaciones IRPF 2026**                                                      | **2 h**  |
| — 4.1. IRPF 2026 — adaptación y corrección campaña Renta                           | 2 h      |

| **5. Integración y mejoras menores** | **5 h** |
| — 5.1. Pipeline de integración Análisis A3 con la aplicación | 3 h |
| — 5.2. Migración fuente de datos y ajustes en proceso de autónomos | 2 h |
| **Total** | **xx h** |

---

## Detalle por concepto

### 1. Unificación y mejora del proceso de Certificados de estar al corriente (~18 h)

Los tres procesos independientes de obtención de certificados se han fusionado en un único flujo con selector multi-organismo, permitiendo lanzar los tres de forma conjunta o individual desde una sola pantalla:

- **Seguridad Social (SS):** certificado de estar al corriente en el pago de cuotas a la Seguridad Social
- **Agencia Tributaria (ATC):** certificado de estar al corriente en obligaciones tributarias
- **Inspección de Trabajo (ITA):** certificado de corriente en obligaciones laborales, gestionado a través del portal InSeNaCoder

**Nuevas funcionalidades:**

- Proceso unificado con selector multi-check (SS, ATC, ITA) — anteriormente eran 3 procesos separados
- Pre-selección automática del certificado digital adecuado antes de procesar cada empresa
- Agrupación automática de todos los PDFs descargados en una carpeta con la fecha de ejecución

**Correcciones y robustez:**

- Corrección de checkboxes en ATC y flujo de autenticación
- Corrección en la selección del certificado digital en ATC
- Detección robusta de PDFs y reintentos automáticos en la descarga
- Captura de la pestaña InSeNaCoder del ITA mediante CDP (Chrome DevTools Protocol) para mayor estabilidad
- Reintentos automáticos en ITA al detectar fallos de conexión

---

### 2. Nuevos informes en el sistema de Análisis A3 (~16 h)

Desarrollo de cuatro nuevos formatos de informe a partir de los ficheros DAT exportados desde A3, y creación de una consola unificada que los agrupa todos.

**Nuevos informes generados en Excel:**

- **Listado Enfermedad (Fmt. 10):** trabajadores en situación de IT, con base reguladora y conceptos de cobro por enfermedad común y accidente (C450/C451)
- **Listado Accidentes (Fmt. 7):** trabajadores con partes de accidente de trabajo, con lectura directa de registros NIN
- **Listado Embargos (Fmt. 15):** trabajadores con embargo de nómina, incluyendo columna de paga extra proporcional
- **Listado Extranjeros (Fmt. 6):** trabajadores de nacionalidad extranjera con sus datos de contrato

**Mejoras de sistema:**

- Consola unificada `generador_informes` que agrupa todos los formatos disponibles (6, 7, 8, 10, 15) en una sola herramienta
- Integración del generador en la TUI (interfaz de terminal) de la aplicación principal
- Optimización del pipeline de lectura de ficheros DAT: fusión de lecturas NTR/NPT para reducir tiempos
- Limpieza de scripts obsoletos y unificación de la documentación de origen de datos

---

### 3. Módulo de extracción de nóminas desde A3 — _en curso_ (~12 h)

Inicio del desarrollo de un módulo para extraer nóminas directamente de los ficheros DAT de A3, sin depender de exportaciones manuales desde el programa. Trabajo realizado en abril:

- Diseño y planificación detallada del proceso (documento de plan y especificaciones)
- Excel maestro de configuración: relación de empresas, trabajadores y parámetros de envío
- Modelos de datos para la lectura del Excel maestro
- Extracción directa de datos de nómina desde ficheros DAT
- Verificación del output generado contra PDFs de nómina de referencia

_Módulo pendiente de completar en próximas iteraciones._

---

### 4. Adaptaciones IRPF 2026, Seguridad Social y ATC (~10 h)

Actualizaciones necesarias para mantener la compatibilidad con los cambios en las plataformas de la administración y con la campaña de la Renta 2026.

**IRPF 2026:**

- Desarrollo y adaptación del proceso de presentación del IRPF para el ejercicio 2026
- Corrección de bloqueo en la sección de datos económicos del formulario, que impedía avanzar en determinados casos

**Seguridad Social:**

- Adaptación del proceso automatizado al navegador Microsoft Edge (la sede electrónica de la SS dejó de ser compatible con el flujo anterior)
- Corrección de fallos en la interacción con los controles de la web de la SS

**ATC:**

- Adaptación del proceso de autenticación al navegador Chrome
- Corrección en la selección del certificado digital durante el inicio de sesión

**Aplicación general:**

- Corrección de solapamiento visual de campos en el formulario de procesos

---

### 5. Integración y mejoras menores (~5 h)

- **Pipeline A3 + aplicación:** nuevo flujo que encadena automáticamente la generación de informes de Análisis A3 con la ejecución del proceso en la aplicación, eliminando el paso manual entre ambos
- **Migración de fuente de datos:** el input de empresas para el proceso de autónomos se ha unificado en el fichero `Excel CCC Empresas.xlsx`, que ya se usaba en otros procesos
- **Ajustes menores:** renombrado de etiquetas y argumentos en el proceso de autónomos para mayor claridad

---

_Documento generado el 30 de abril de 2026_
