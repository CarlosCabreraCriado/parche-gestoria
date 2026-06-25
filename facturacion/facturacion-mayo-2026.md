# Informe de trabajo — Mayo 2026

**Cliente:** Del Castillo Asesores
**Período:** 1 – 31 de mayo de 2026
**Proyectos:** Aplicación de gestión (parche-gestoria)

---

## Resumen de conceptos

| Concepto / Subconcepto                                                                       | Horas    |
| -------------------------------------------------------------------------------------------- | -------- |
| **1. Nuevo subproceso Art.42 — Contratas y subcontratas**                                   | **8 h**  |
| — 1.1. Subproceso completo Art.42 integrado en el formulario de certificados                 | 5 h      |
| — 1.2. Modo automático con flags Excel e inputs manuales de empresa autorizada               | 3 h      |
| **2. Selección automática de certificado digital en AEAT**                                   | **20 h** |
| — 2.1. Mecanismo Chrome-por-empresa con política AutoSelect vía almacén Windows              | 12 h     |
| — 2.2. Elevación UAC, persistencia de clave y correcciones de compatibilidad                 | 8 h      |
| **3. Generación de correos con resultados de certificados**                                   | **8 h**  |
| — 3.1. Borradores de correo por empresa con plantilla del cliente                            | 5 h      |
| — 3.2. Acumulación de logs y correcciones en descarga de PDF TRIB                            | 3 h      |
| **4. Correcciones de robustez y mejoras menores**                                             | **4 h**  |
| — 4.1. Correcciones SS / ATC / Art42 standalone y FIE régimen autónomos                     | 2 h      |
| — 4.2. TC2 en Formatear Recibos, métricas, facturación y seguridad                          | 2 h      |
| **Total**                                                                                     | **40 h** |

---

## Detalle por concepto

### 1. Nuevo subproceso Art.42 — Contratas y subcontratas (~8 h)

Desarrollo del nuevo flujo para obtener el certificado de empresa autorizada/contratista (Art. 42 LISOS), integrado en el proceso unificado de certificados de estar al corriente:

- Nuevo checkbox «Art. 42» en el formulario de certificados; proceso disponible tanto de forma conjunta como independiente
- Subproceso completo de obtención del certificado, con campos de empresa autorizada configurables desde el formulario o desde el Excel de empresas
- **Modo automático:** activación desde flags del Excel, deshabilitando los inputs manuales en ese modo
- Homogeneización del formato de nombres de archivo para ITA y Art.42 (alineado con el resto de certificados)
- Refactorización del frontend: `mostrarSi` por proceso para ocultar/mostrar cada sección de forma independiente

---

### 2. Selección automática de certificado digital en AEAT (~20 h)

El portal de la AEAT requería seleccionar manualmente el certificado digital de cada empresa antes de iniciar la gestión, bloqueando la automatización. El nuevo mecanismo elimina completamente la intervención manual:

**Núcleo del mecanismo (semana del 11 al 22 de mayo):**

- Por cada empresa, se lanza una instancia de Chrome limpia con una política de auto-selección registrada en Windows que apunta al CN del certificado de esa empresa
- El certificado se inyecta en el almacén de Windows (`X509Store`, buscando en `CurrentUser` y `LocalMachine`) en lugar de pasarlo como argumento de Playwright, evitando el diálogo de selección del navegador
- La política AutoSelect se escribe en el registro de Windows mediante PowerShell Security, usando BOM UTF-8 para compatibilidad con PowerShell 5.x

**Correcciones y robustez:**

- Elevación de privilegios UAC al inicio de la ejecución para garantizar acceso de escritura al registro; la clave se preserva entre ejecuciones para no elevar UAC repetidamente
- Corrección de encoding UTF-8 en nombres de CN con acentos (certificados de personas físicas)
- Corrección para aceptar certificados personales en la solicitud AEAT y seleccionar «En nombre propio» sin rellenar datos del titular
- Soporte de formato `.p12` como alternativa a `.pfx`
- Búsqueda automática de certificados sin necesidad de `config.json`; búsqueda de fallback en `CurrentUser` si la política de auto-selección falla
- Evitar reprocesar empresas con múltiples CCC (primer CCC válido por empresa)

---

### 3. Generación de correos con resultados de certificados (~8 h)

Al finalizar cada empresa, el proceso genera automáticamente un borrador de correo listo para enviar al cliente con el resumen del resultado:

- Generación de borradores de correo por empresa en el cliente de correo instalado
- Plantilla de correo actualizada con los textos definitivos proporcionados por el cliente
- El correo se omite cuando todos los certificados de una empresa han fallado (para no enviar mensajes vacíos)
- El Excel de salida acumula ahora los logs de ejecuciones anteriores, manteniendo el historial completo sin sobrescribir
- Corrección de errores en la descarga de PDF TRIB por CDP y eliminación de timeout fantasma que provocaba falsos errores

---

### 4. Correcciones de robustez y mejoras menores (~4 h)

**Certificados — desajuste de índices:**

- Corrección de desajuste de índices en los procesos standalone de SS, ITA, AEAT y Art.42 (bug introducido al separar los procesos en opciones independientes)

**Otros procesos:**

- **Formatear Recibos de Liquidación:** soporte de documentos TC2 junto a los TC1 ya existentes
- **FIE:** omitir el campo CCC en el régimen de autónomos (código 0521), donde no aplica
- **Seguridad:** eliminación de credenciales de base de datos del historial de permisos git y exclusión del fichero de configuración local mediante `.gitignore`
- **Métricas:** corrección para mostrar el nombre de empresa en el reporte de Formatear Recibos
- **Facturación:** deshabilitar la escritura manual en los campos de fecha (solo selección por picker)
- Renombrado del archivo de certificado ITA de «CERT CORRIENTE» a «Informe ITA»

---

_Documento generado el 25 de junio de 2026_
