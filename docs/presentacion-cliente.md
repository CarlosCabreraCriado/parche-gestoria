# Automatización de Procesos · Del Castillo Asesores
### Informe de Progreso y Hoja de Ruta
**Nodus · Abril 2026**

---

## 1. El Ecosistema de la Gestoría: Los 3 Ejes

Una gestoría trabaja constantemente moviendo información entre tres actores:

```
                    ┌──────────────────────┐
                    │      CLIENTES        │
                    │  Empresas · Autónomos│
                    └──────────┬───────────┘
                               │
               Datos y         │         Tramitaciones
               encargos        │         y resultados
                               │
              ┌────────────────▼─────────────────┐
              │                                  │
              │       DEL CASTILLO ASESORES      │
              │     (nodo central de gestión)    │
              │                                  │
              └──────┬───────────────┬───────────┘
                     │               │
          Consulta y │               │ Presentación
          extracción │               │ de trámites
          de datos   │               │
                     ▼               ▼
         ┌───────────────┐   ┌────────────────────┐
         │   A3 ASESOR   │   │  ADMINISTRACIONES  │
         │ (Nóminas,     │   │     PÚBLICAS       │
         │  Contabilidad,│   │ AEAT · Seg. Social │
         │  IRPF, SS...) │   │ DGT · INSS · SEPE  │
         └───────────────┘   └────────────────────┘
```

**El trabajo diario de la gestoría consiste en tramitar el flujo de información entre estos tres ejes:**
- Recoger datos de los clientes y registrarlos en A3
- Extraer información de A3 para preparar trámites
- Presentar esos trámites ante las administraciones
- Devolver resultados y certificados a los clientes

**El objetivo del proyecto de automatización es eliminar los pasos manuales y repetitivos de ese flujo, para que el equipo pueda centrarse en las tareas que realmente aportan valor.**

---

## 2. Dónde Estábamos: El Proceso Manual

Antes de la automatización, cada proceso seguía un ciclo manual:

| Paso | Acción manual | Tiempo estimado |
|------|--------------|-----------------|
| 1 | Entrar en A3 y generar el informe correspondiente | 5–10 min |
| 2 | Descargar el Excel o listado generado | 2 min |
| 3 | Abrir el navegador y acceder a la plataforma de la SS/AEAT | 2 min |
| 4 | Buscar cada trabajador/empresa uno a uno | Variable (muy alto) |
| 5 | Descargar o registrar los documentos | Variable |
| 6 | Organizar los archivos en carpetas | 5–15 min |

**El cuello de botella era el paso 4:** con lotes de 50–200 trabajadores, este proceso podía ocupar horas de trabajo.

---

## 3. Dónde Estamos Ahora: Lo Que Ya Funciona

### 3.1 · parche-gestoria (el .exe)

La herramienta principal que usa el equipo hoy: una aplicación de escritorio que automatiza los procesos más repetitivos de la gestoría.

El operador selecciona el proceso, introduce los parámetros necesarios (ruta del Excel de A3, año, empresa...) y la herramienta ejecuta todo el ciclo automáticamente.

**Procesos actualmente automatizados:**

| Proceso | Qué hace | Antes era... |
|---------|----------|--------------|
| **Duplicados TA2 + IDC** | Descarga automática de documentos TA2 e IDC por trabajador desde la Seguridad Social | Búsqueda manual uno a uno |
| **Bases y Recibos de Autónomos** | Descarga de bases cotizadas y recibos al cobro por NAF y año | Acceso y descarga manual |
| **Informes FIE** | Generación de documentos de altas, bajas y confirmaciones de IT | Preparación manual de PDFs |
| **Certificados AEAT / SS** | Descarga de certificados tributarios, de estar al corriente, de cotización | Solicitudes manuales |
| **Gestiones IRPF** | Procesos de campaña de renta, etiquetas fiscales, cambios en base de cotización | Tramitación manual vía web |
| **Análisis KPIs** | Informes de ciclo de facturación y métricas de la asesoría | Hoja de cálculo manual |

> Todos estos procesos se ejecutan con un solo clic, sin que el empleado tenga que interactuar con las plataformas web.

---

### 3.2 · analisis-a3 (el extractor de datos)

Paralelamente, se ha desarrollado un sistema en Python que lee directamente los archivos de datos de A3 Asesor, sin necesidad de que el empleado entre en A3 a generar informes manualmente.

**Qué extrae hoy:**

| Módulo A3 | Datos extraídos | Formatos de salida |
|-----------|-----------------|--------------------|
| A3NOMV5E (Nóminas) | Listado de altas, bajas por IT, accidentes, embargos, trabajadores extranjeros | Excel (.xlsx) |
| A3ENTORNO (CRM) | Datos de clientes, domicilios, contactos, empresas | SQLite, CSV |

Hasta ahora, estos informes se generaban **manualmente desde la interfaz de A3** y luego se alimentaban a la herramienta. Con este extractor, **ese paso desaparece**.

---

### 3.3 · El Pipeline: Cerrando el Ciclo

El paso más importante en desarrollo: el **pipeline** que conecta el extractor de datos con los procesos automáticos.

```
  ┌──────────┐     ┌───────────────────┐     ┌──────────────────┐     ┌──────────────┐
  │          │     │   analisis-a3     │     │  parche-gestoria │     │              │
  │  A3      │────▶│  (lee los datos   │────▶│  (ejecuta el     │────▶│ Administra-  │
  │ Asesor   │     │   directamente)   │     │   proceso)       │     │ ción Pública │
  │          │     └───────────────────┘     └──────────────────┘     │              │
  └──────────┘                                                         └──────────────┘
        │                                                                      │
        │                                                                      │
        └──────────────────────────── ─ ─ ─ ─ ─ ─ ─ ─ ─────────────────────▶│
                           (flujo completamente automático)
```

**Primer pipeline en producción:** Listado de Altas (Formato 8) → Descarga automática de TA2 + IDC

> Sin intervención del empleado: A3 tiene los datos → el sistema los extrae, los procesa y descarga los documentos de la Seguridad Social.

---

## 4. Analíticas: Impacto y Ahorro de Tiempo

> **[DATOS PENDIENTES — rellenar con métricas reales del sistema]**

### Estructura prevista para esta sección:

#### Ejecuciones por proceso (último mes)

| Proceso | Nº ejecuciones | Registros procesados |
|---------|---------------|---------------------|
| Duplicados TA2 + IDC | `___` | `___` trabajadores |
| Bases Autónomos | `___` | `___` clientes |
| Informes FIE | `___` | `___` documentos |
| ... | | |

#### Ahorro de tiempo estimado

| Proceso | Tiempo manual (estimado) | Tiempo automático | Ahorro por ejecución |
|---------|--------------------------|-------------------|----------------------|
| Duplicados TA2 (lote 100 trabajadores) | ~3 horas | ~8 min | ~2h 50min |
| Bases Autónomos (20 clientes) | ~1 hora | ~5 min | ~55 min |
| Certificado AEAT por empresa | ~15 min | ~2 min | ~13 min |

#### Equivalente en horas de empleado (mensual)

```
  Horas ahorradas al mes:  ████████████████████  XX h
  Coste equivalente:       ████████████████████  XXX €
```

> *Los datos exactos se extraen del sistema de métricas integrado en la herramienta, que registra cada ejecución en producción.*

---

## 5. Próximos Pasos y Potencial Futuro

### Corto plazo — Completar el Pipeline

Hoy el pipeline automatiza el flujo de **Altas → TA2+IDC**. El siguiente paso es conectar el resto de informes:

- **Formato 6 (Extranjeros)** → gestiones de NIE y permisos de trabajo
- **Formato 7 (Accidentes)** → partes de accidente a la Mutua / SS
- **Formato 10 (Bajas por IT)** → confirmaciones de incapacidad temporal
- **Formato 15 (Embargos)** → gestión de retenciones salariales

Resultado: **cualquier proceso que hoy depende de un Excel generado manualmente, pasará a ejecutarse de forma completamente automática**.

---

### Medio plazo — Más Datos de A3

A3 Asesor contiene muchos más módulos de los que hoy se extraen. Cada uno representa una categoría de procesos que se puede automatizar:

| Módulo | Qué contiene | Qué se puede automatizar |
|--------|-------------|--------------------------|
| **A3ECO** (Contabilidad) | Asientos, facturas, modelos 303/347/390 | Generación y presentación de modelos fiscales |
| **A3REN** (IRPF) | Declaraciones de renta, datos de clientes | Revisión automática de borradores, detección de errores |
| **A3SOCW** (Sociedades) | Impuesto de Sociedades, modelos 200/220 | Verificación y preparación de modelos |
| **A3BANK** (Tesorería) | Movimientos bancarios, conciliaciones | Conciliación automática y alertas de descuadres |
| **A3GESW** (Facturación) | Facturas emitidas, VeriFactu | Factura electrónica automática (obligatoria desde 2025) |

---

### Largo plazo — Cerrar el Ciclo Completo

La visión final es que **el flujo entre los 3 ejes funcione de forma autónoma**, con mínima intervención manual:

#### Portal del Cliente
Los clientes podrán consultar el estado de sus trámites en tiempo real, sin necesidad de llamar o enviar un email para preguntar. Visibilidad de documentos, vencimientos y gestiones pendientes.

#### Notificaciones Automáticas
El sistema detecta eventos importantes y avisa automáticamente:
- "El certificado de estar al corriente de la empresa X está listo"
- "El trabajador Y tiene un parte de baja que vence el día Z"
- "Hay un embargo activo que requiere atención"

#### Conexión Directa con Sede Electrónica
En lugar de que un empleado acceda manualmente a la web de la AEAT o la Seguridad Social, el sistema conecta directamente (vía API o automatización web) para **presentar documentos sin intervención humana**.

#### Análisis Predictivo e Inteligencia de Datos
Con todos los datos de A3 accesibles de forma estructurada, se puede construir:
- Alertas de vencimientos (modelos trimestrales, declaraciones anuales)
- Comparativas entre ejercicios por cliente
- Detección de anomalías en nóminas o contabilidad
- Informes ejecutivos automáticos para los clientes

---

## 6. La Visión: Del Ciclo Manual al Ciclo Automático

```
  HOY                                        MAÑANA

  Empleado                                   Sistema
     │                                          │
     ├─ Entra en A3                             ├─ Lee A3 directamente
     ├─ Genera informe                          │
     ├─ Descarga Excel                          ├─ Ejecuta proceso
     ├─ Abre herramienta                        │
     ├─ Carga Excel                             ├─ Presenta a la administración
     ├─ Ejecuta proceso                         │
     ├─ Accede a administración                 ├─ Notifica al cliente
     ├─ Gestiona trámite                        │
     └─ Informa al cliente                      └─ Registra métricas
```

**El equipo de Del Castillo Asesores no desaparece — se transforma.** Las horas que hoy se dedican a tareas repetitivas y mecánicas se liberan para atención al cliente, resolución de casos complejos y asesoramiento de valor.

---

*Documento elaborado por Nodus · Abril 2026*  
*Para más información sobre los desarrollos en curso, contactar con el equipo técnico.*
