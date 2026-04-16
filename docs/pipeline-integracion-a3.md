# Pipeline de Integracion: analisis-a3 + parche-gestoria

## Indice

1. [Contexto](#contexto)
2. [Arquitectura general](#arquitectura-general)
3. [Mapeo de formatos y procesos](#mapeo-de-formatos-y-procesos)
4. [Pipeline MVP: Altas -> Duplicados TA2+IDC](#pipeline-mvp-altas---duplicados-ta2idc)
5. [Archivos involucrados](#archivos-involucrados)
6. [Flujo de datos detallado](#flujo-de-datos-detallado)
7. [Configuracion del formulario UI](#configuracion-del-formulario-ui)
8. [Manejo de errores](#manejo-de-errores)
9. [Guia de uso](#guia-de-uso)
10. [Extension a futuros pipelines](#extension-a-futuros-pipelines)

---

## Contexto

### El problema

Los proyectos `analisis-a3` y `parche-gestoria` operan de forma aislada:

- **analisis-a3** (Python) lee archivos binarios COBOL de A3 Asesor (`M:\A3`) y genera listados XLSX con datos de trabajadores, accidentes, enfermedades, etc.
- **parche-gestoria** (Electron/Node.js) consume esos XLSX para automatizar tramites en la Seguridad Social via Puppeteer (descarga de TA2, IDC, procesamiento FIE, etc.).

Hasta ahora, el usuario debia:

1. Abrir una terminal y ejecutar `python scripts/modules/nomv5e/generador_informes.py` con los parametros adecuados
2. Esperar a que se genere el XLSX
3. Abrir parche-gestoria
4. Cargar manualmente el XLSX como input del proceso correspondiente
5. Ejecutar el proceso

### La solucion

Un **pipeline integrado dentro de parche-gestoria** que encadena ambos pasos automaticamente: genera el XLSX invocando Python como subprocess y luego lo pasa directamente al proceso JS correspondiente. El usuario solo rellena un formulario y pulsa ejecutar.

---

## Arquitectura general

```
+---------------------------+          +---------------------------+
|      analisis-a3          |          |     parche-gestoria       |
|        (Python)           |          |   (Electron + Angular)    |
|                           |          |                           |
|  M:\A3 (COBOL binario)   |          |  UI Angular (formulario)  |
|         |                 |          |         |                 |
|         v                 |          |         v                 |
|  scripts/.../              |  spawn   |  pipeline.js              |
|   generador_informes.py   | <-----  |  _spawnPython()           |
|         |                 |          |         |                 |
|         v                 |          |         v                 |
|  altas_fmt8.xlsx          | -------> |  duplicados.js            |
|  (output temporal)        |  path    |  duplicadosTa2()          |
|                           |          |         |                 |
+---------------------------+          |         v                 |
                                       |  Puppeteer -> PDFs        |
                                       +---------------------------+
```

### Tecnologias y comunicacion

| Capa | Tecnologia | Rol |
|------|-----------|-----|
| UI | Angular 17 + Material | Formulario de parametros |
| IPC | Electron ipcMain/ipcRenderer | Comunicacion UI <-> Backend |
| Orquestador | Node.js (`pipeline.js`) | Coordina subprocess Python + proceso JS |
| Generador | Python 3.14 (`generador_informes.py`) | Lee COBOL, genera XLSX |
| Automatizador | Puppeteer (`duplicados.js`) | Navega Seguridad Social, descarga PDFs |

---

## Mapeo de formatos y procesos

### Conexiones confirmadas

| analisis-a3 | Formato | Output | parche-gestoria | Proceso | Estado |
|---|---|---|---|---|---|
| `parse_altas` | 8 - Listado Altas | `altas_fmt8.xlsx` (23 cols) | `duplicados.js` | `duplicadosTa2()` | **Implementado** |
| `parse_accidentes` | 7 - Listado Accidentes | `accidentes_fmt7.xlsx` (17 cols) | `fie.js` | `fIE()` [input 3] | Futuro |
| `parse_enfermedad` | 10 - Listado Enfermedad | `enfermedad_fmt10.xlsx` (18 cols) | `fie.js` | `fIE()` [input 2] | Futuro |
| `parse_extranjeros` | 6 - Listado Extranjeros | `extranjeros_fmt6.xlsx` (10 cols) | - | Sin proceso asociado | - |
| `parse_embargos` | 15 - Listado Embargos | `embargos_fmt15.xlsx` (7 cols) | - | Sin proceso asociado | - |

### Detalle del mapeo Format 8 -> Duplicados

Las columnas del XLSX generado por Format 8 coinciden **exactamente** con las que espera `duplicados.js`:

| Columna XLSX (analisis-a3) | Campo en duplicados.js | Uso |
|---|---|---|
| `Emp->Codigo_de_la_Empresa` | `exp` | Identificador de empresa |
| `Emp->Nombre_de_la_Empresa` | `empresa` | Nombre para logs/carpetas |
| `Cent->2_primeras_cifras_Segsoc` | `prov_ccc` | Provincia CCC (2 digitos) |
| `Cent->7_siguientes_cifras_SegSoc` | `ccc7` | CCC parte 1 (7 digitos) |
| `Cent->2_ultimas_cifras_SegSoc` | `ccc2` | CCC parte 2 (2 digitos) |
| `Trab->Apellidos_y_Nombre_del_Trabajador` | `trabajador` | Nombre en PDFs |
| `Trab->DNI_del_Trabajador` | `dni` | Clave unica por trabajador |
| `Trab->2_primeras_cifras_Segsoc` | `prov_naf` | Provincia NAF (2 digitos) |
| `Trab->8_siguientes_cifras_SegSoc` | `naf8` | NAF parte 1 (8 digitos) |
| `Trab->2_ultimas_cifras_SegSoc` | `naf2` | NAF parte 2 (2 digitos) |

Las 13 columnas restantes del Format 8 (fechas, contrato, categoria, etc.) no son usadas por duplicados.js pero se incluyen en el XLSX igualmente.

---

## Pipeline MVP: Altas -> Duplicados TA2+IDC

### Que hace

1. **Genera** el listado de trabajadores activos (altas) desde los ficheros COBOL de A3 Asesor
2. **Ejecuta** el proceso de descarga de duplicados TA2 e IDC desde la Seguridad Social para cada trabajador

### Parametros de entrada (formulario UI)

| Campo | Tipo | Obligatorio | Descripcion | Ejemplo |
|-------|------|-------------|-------------|---------|
| Google .exe | Archivo | Si | Ruta al ejecutable de Chrome/Chromium | `C:\Program Files\Google\Chrome\Application\chrome.exe` |
| Codigos de empresa | Texto | Si | Codigos A3 separados por coma (5 digitos) | `00008, 01378` |
| Regimen | Texto | Si | Codigo de regimen SS (4 digitos) | `0111` |
| Directorio de salida | Ruta | Si | Carpeta donde se guardan los PDFs | `C:\Salida\Duplicados` |
| Ruta Python | Texto | No | Ejecutable de Python (default: `python`) | `python` |
| Ruta analisis-a3 | Ruta | No | Carpeta raiz del proyecto analisis-a3 | `C:\Users\preprod\Documents\Proyectos\analisis-a3` |

### Output generado

```
<Directorio de salida>/
  pipeline-temp/
    altas_fmt8.xlsx              # XLSX intermedio generado por Python
  Duplicados (YYYY-MM-DD)/
    LOGS/
      detalle_ta2.log            # Log de descarga TA2 por DNI
      detalle_idc.log            # Log de descarga IDC por DNI
    <DNI_1>/
      TA2 ADDMMYY Nombre.pdf    # Documento TA2 del trabajador
      IDC ADDMMYY Nombre.pdf    # Informe de datos de cotizacion
    <DNI_2>/
      ...
```

---

## Archivos involucrados

### Archivos creados

| Archivo | Proyecto | Descripcion |
|---------|----------|-------------|
| `procesos/pipeline.js` | parche-gestoria | Clase `ProcesosPipeline` con el orquestador |

### Archivos modificados

| Archivo | Proyecto | Cambios |
|---------|----------|---------|
| `main.js` | parche-gestoria | Import, variable, instanciacion y case "pipeline" en switch |
| `src/app/comun/procesos/procesos.configuracion.ts` | parche-gestoria | Tipo "Pipeline" y definicion del formulario |

### Archivos de referencia (sin modificar)

| Archivo | Proyecto | Rol |
|---------|----------|-----|
| `scripts/modules/nomv5e/generador_informes.py` | analisis-a3 | Script Python invocado como subprocess |
| `src/nomv5e_reader.py` | analisis-a3 | Parser de datos COBOL de nominas |
| `scripts/modules/nomv5e/export_altas.py` | analisis-a3 | Define cabeceras XLSX del Format 8 |
| `scripts/modules/nomv5e/xlsx_exporter.py` | analisis-a3 | Genera XLSX sin dependencias externas |
| `procesos/duplicados.js` | parche-gestoria | Proceso Puppeteer de descarga TA2/IDC |

---

## Flujo de datos detallado

### Paso 0: Dispatch desde la UI

```
Angular UI
  -> usuario selecciona "Pipeline > PIPELINE ALTAS DUPLICADOS"
  -> rellena formulario (chrome, empresas, regimen, salida, python, ruta-a3)
  -> click Ejecutar
  -> ipcRenderer.invoke("onEjecutarProceso", proceso, argumentos)

main.js :: onEjecutarProceso
  -> camelize("PIPELINE ALTAS DUPLICADOS") = "pIPELINEALTASDUPLICADOS"
  -> switch(categoria === "pipeline")
  -> procesosPipeline["pIPELINEALTASDUPLICADOS"](argumentos)
  -> delega a pipelineAltasDuplicados()
```

### Paso 1: Generacion del XLSX (Python)

```
pipeline.js :: pipelineAltasDuplicados()
  -> Valida inputs (chrome existe, empresas no vacias, ruta A3 contiene scripts/modules/nomv5e/generador_informes.py)
  -> Normaliza codigos de empresa: "8, 01378" -> ["00008", "01378"]
  -> Crea directorio: <salida>/pipeline-temp/
  -> Elimina XLSX previo si existe (evitar datos stale)

  -> _spawnPython("python", "<ruta-a3>", [
       "scripts/modules/nomv5e/generador_informes.py",
       "--empresa", "00008,01378",
       "--formato", "8",
       "--output", "<salida>/pipeline-temp/altas_fmt8.xlsx"
     ])

Python (scripts/modules/nomv5e/generador_informes.py):
  -> Parsea --empresa, --formato, --output
  -> Instancia NomV5EReader(base_path="M:/A3/A3NOMV5E")
  -> Llama parse_altas(["00008", "01378"])
     -> Lee NOMCCC.DAT (centros + trabajadores)
     -> Lee NEM00008.DAT, NEM01378.DAT (nombres de empresa)
     -> Lee NTR00008.DAT, NTR01378.DAT (fechas, tarifa, categoria)
     -> Lee NTN00008.DAT, NTN01378.DAT (tipo contrato)
     -> Lee A3TRT*.DAT (transmisiones SS: motivo baja, fecha fin contrato)
     -> Filtra: solo fecha_baja == "00/00/0000" AND cod_trabajador 1-999999
  -> export_xls(workers, output_path)
     -> Genera XLSX con 23 columnas y N filas de trabajadores activos

pipeline.js:
  -> Verifica que el XLSX existe y tiene tamanio > 0
  -> Log: "XLSX generado: <path> (XX.X KB)"
```

### Paso 2: Ejecucion de Duplicados (Puppeteer)

```
pipeline.js:
  -> Instancia ProcesosDuplicados(pathToDbFolder, nombreProyecto, proyectoDB)
  -> Llama duplicadosTa2({
       formularioControl: [chromeExePath, tempXlsx, regimen4, pathSalidaBase]
     })

duplicados.js :: duplicadosTa2():
  -> Lee el XLSX generado via leerExcelDuplicados()
     -> Detecta fila de cabecera (busca columnas DNI, NAF, CCC)
     -> Extrae registros: exp, empresa, provCCC, ccc, trabajador, dni, provNAF, naf
  -> Normaliza y valida cada registro
  -> Deduplica por DNI
  -> Lanza Puppeteer con Chrome
  -> Navega a https://w2.seg-social.es/fs/indexframes.html
  -> Por cada trabajador:
     a) Proceso TA2 (ATR65):
        -> Rellena formulario: NAF, CCC, Regimen
        -> Selecciona el ALTA mas reciente
        -> Descarga PDF via CDP
        -> Guarda: <DNI>/TA2 ADDMMYY Nombre.pdf
     b) Proceso IDC (ATR37):
        -> Rellena formulario similar
        -> Selecciona F.R. Alta mas reciente
        -> Descarga PDF via CDP
        -> Guarda: <DNI>/IDC ADDMMYY Nombre.pdf
  -> Escribe logs cada 5 registros
  -> Cierra browser
  -> Retorna true/false
```

### Paso 3: Resultado

```
pipeline.js:
  -> Registra metricas: registrarEjecucion("PIPELINE_ALTAS_DUPLICADOS", N)
  -> Log: "Pipeline finalizado. Resultado duplicados: true/false"
  -> resolve(resultado) -> devuelve a main.js -> devuelve a Angular UI
```

---

## Configuracion del formulario UI

La configuracion vive en `src/app/comun/procesos/procesos.configuracion.ts`.

### Estructura en el arbol de procesos

```
Libreria de procesos
  |-- Asesoria
  |-- FIE
  |-- Duplicados
  |     |-- DUPLICADOS TA2+IDC          (proceso existente, pide Excel manual)
  |-- Bases y recibos al cobro autonomos
  |-- Pipeline                            (NUEVO directorio)
        |-- PIPELINE ALTAS DUPLICADOS    (NUEVO proceso integrado)
```

### Relacion nombre -> metodo

El dispatch de `main.js` transforma el nombre del proceso con `camelize()`:

| nombre en configuracion | Tras normalize + camelize | Metodo en pipeline.js |
|---|---|---|
| `"PIPELINE ALTAS DUPLICADOS"` | `"pIPELINEALTASDUPLICADOS"` | `["pIPELINEALTASDUPLICADOS"]()` |

Esto sigue el mismo patron que los procesos existentes:

| nombre | camelize | Metodo |
|---|---|---|
| `"DUPLICADOS TA2+IDC"` | `"dUPLICADOSTA2+IDC"` | `["dUPLICADOSTA2+IDC"]()` |

---

## Manejo de errores

### Errores en Paso 1 (Python)

| Error | Causa | Mensaje en consola | Resultado |
|-------|-------|-------------------|-----------|
| ENOENT | Python no esta en PATH ni en la ruta indicada | `Python no encontrado en "python"` | `false` |
| Exit code != 0 | Script falla (M: no montada, empresa invalida, etc.) | Muestra stderr completo | `false` |
| Timeout (5 min) | Demasiadas empresas o disco lento | `Python subprocess excedio el timeout` | `false` |
| XLSX no existe | Script OK pero 0 trabajadores (todos dados de baja) | `No genero el XLSX` | `false` |

### Errores en Paso 2 (Duplicados)

Los errores de Puppeteer son gestionados por `duplicados.js` internamente:

| Error | Comportamiento |
|-------|---------------|
| Frame desconectado | Reintento automatico (3 intentos) |
| PDF no descargado | Reintento con reapertura de popup |
| DIL error (SS) | Detectado y registrado en log |
| Trabajador ya procesado | Se salta (resume via logs) |

### Recuperacion

Si el pipeline falla en el Paso 2 (Puppeteer), el XLSX ya esta generado en `pipeline-temp/altas_fmt8.xlsx`. El usuario puede:

1. Ejecutar el pipeline de nuevo (reusara el XLSX si no se elimina, aunque por defecto lo regenera)
2. Usar el proceso "DUPLICADOS TA2+IDC" original pasando el XLSX manualmente
3. Los logs de TA2/IDC permiten resume automatico (se saltan DNIs ya procesados OK)

---

## Guia de uso

### Prerrequisitos

1. **Python 3.14+** instalado y accesible desde PATH (o indicar la ruta)
2. **Unidad M:** montada con acceso a `M:\A3\A3NOMV5E` (datos de nominas)
3. **Google Chrome** instalado
4. **Certificado digital** activo para acceso a la Seguridad Social
5. **Proyecto analisis-a3** descargado en la ruta por defecto o indicar ruta alternativa

### Pasos

1. Abrir **parche-gestoria** (Electron)
2. Abrir o crear un proyecto
3. Navegar a **Pipeline** en el arbol de procesos
4. Seleccionar **"PIPELINE ALTAS DUPLICADOS"**
5. Rellenar el formulario:
   - **Google .exe**: Seleccionar `chrome.exe`
   - **Codigos de empresa**: Escribir los codigos separados por coma (ej: `00008, 01378`)
   - **Regimen**: Normalmente `0111` (General)
   - **Directorio de salida**: Elegir carpeta donde guardar los PDFs
   - (Opcional) **Ruta Python**: Dejar `python` si esta en PATH
   - (Opcional) **Ruta analisis-a3**: Dejar el valor por defecto si no se ha movido
6. Pulsar **Ejecutar**
7. Esperar a que:
   - Se genere el XLSX (segundos a minutos segun numero de empresas)
   - Se descarguen los TA2/IDC (minutos a horas segun numero de trabajadores)
8. Revisar la carpeta de salida:
   - `pipeline-temp/altas_fmt8.xlsx` - listado generado
   - `Duplicados (YYYY-MM-DD)/` - PDFs organizados por DNI
   - `Duplicados (YYYY-MM-DD)/LOGS/` - detalle de cada descarga

### Monitorizacion

Abrir la consola de Electron (Ctrl+Shift+I > Console) para ver el progreso en tiempo real:

```
[PIPELINE] Iniciando pipeline: Altas (Fmt 8) -> Duplicados TA2+IDC
[PIPELINE] Empresas a procesar: 00008, 01378
[PIPELINE] === PASO 1: Generando listado Altas (Formato 8) ===
[PIPELINE] Ejecutando: python scripts/modules/nomv5e/generador_informes.py --empresa 00008,01378 --formato 8 --output ...
[PIPELINE] [Python stdout] Exportados 142 trabajadores a altas_fmt8.xlsx
[PIPELINE] XLSX generado: ...\pipeline-temp\altas_fmt8.xlsx (45.2 KB)
[PIPELINE] === PASO 2: Ejecutando Duplicados TA2+IDC ===
[DUPLICADOS] Iniciando proceso DUPLICADOS TA2 + IDC (por empleado)
[DUPLICADOS] Registros validos: 140 de 142
...
[PIPELINE] Pipeline finalizado. Resultado duplicados: true
```

---

## Extension a futuros pipelines

### Anadir un nuevo pipeline

El sistema esta disenado para que anadir un nuevo pipeline requiera solo 3 pasos:

#### 1. Anadir metodo en `procesos/pipeline.js`

```js
async pipelineNuevoProceso(argumentos) {
  // 1. Extraer parametros de argumentos.formularioControl
  // 2. Llamar _spawnPython() con el formato adecuado
  // 3. Verificar output
  // 4. Instanciar y llamar al proceso destino
}

// Alias (resultado de camelize sobre el nombre en configuracion)
async ["pIPELINENUEVOPROCESO"](args) {
  return this.pipelineNuevoProceso(args);
}
```

#### 2. Anadir entrada en `procesos.configuracion.ts`

Dentro del array `subCategoria` del directorio "Pipeline":

```ts
{
  nombre: "PIPELINE NUEVO PROCESO",
  categoria: "Pipeline",
  tipo: "proceso",
  descripcion: "Descripcion del pipeline",
  argumentos: [ /* campos del formulario */ ],
  opciones: null,
  salida: [{ tipo: "boolean", valor: false }],
}
```

#### 3. (Opcional) Anadir nuevo formato en analisis-a3

Si el formato no existe, se anade en `scripts/modules/nomv5e/generador_informes.py` > `FORMATOS` dict y se crea su parser en `nomv5e_reader.py` y exportador en `scripts/modules/nomv5e/`.

### Pipeline FIE (proximo candidato)

El proceso FIE (`fie.js`) recibe 5 inputs:

```
[0] pathArchivoFIE          -> proporcionado por el usuario
[1] pathArchivoEmpresas     -> proporcionado por el usuario
[2] pathArchivoEnfermedad   -> GENERADO por Format 10 de analisis-a3
[3] pathArchivoAccidentes   -> GENERADO por Format 7 de analisis-a3
[4] pathSalida              -> proporcionado por el usuario
```

El pipeline FIE ejecutaria Format 10 y Format 7 en paralelo (`Promise.all`) y luego pasaria los XLSX generados junto con los archivos manuales a `fIE()`.

---

## Referencia tecnica

### Funcion `_spawnPython`

```
_spawnPython(pythonPath, cwd, args, timeoutMs?)
```

| Parametro | Tipo | Default | Descripcion |
|-----------|------|---------|-------------|
| `pythonPath` | string | - | Ruta al ejecutable de Python |
| `cwd` | string | - | Directorio de trabajo (raiz de analisis-a3) |
| `args` | string[] | - | Argumentos para el script |
| `timeoutMs` | number | 300000 | Timeout en milisegundos (5 min) |

**Retorna**: `Promise<{code: number, stdout: string, stderr: string}>`

**Errores**: Rechaza la promesa si Python no se encuentra (ENOENT) o si se excede el timeout.

### CLI de generador_informes.py

```
python scripts/modules/nomv5e/generador_informes.py [opciones]

--empresa CODIGO[,CODIGO...]   Codigos de empresa (5 digitos, separados por coma)
--formato {6,7,8,10,15}        Numero de formato
--periodo YYYYMM               Periodo (solo formatos 7, 10, 15)
--inicio YYYYMMDD              Fecha inicio rango (alternativa a --periodo)
--fin YYYYMMDD                 Fecha fin rango (alternativa a --periodo)
--output RUTA                  Ruta completa del XLSX de salida
```

### Formatos disponibles en analisis-a3

| Formato | Nombre | Necesita periodo | Parser | Datos de origen (M:\A3) |
|---------|--------|-----------------|--------|------------------------|
| 6 | Listado Extranjeros | No | `parse_extranjeros` | NOMCCC, NEM, NTR |
| 7 | Listado Accidentes | Si | `parse_accidentes` | NOMCCC, NEM, NTR, NIN, NPT |
| 8 | Listado Altas | No | `parse_altas` | NOMCCC, NEM, NTR, NTN, A3TRT |
| 10 | Listado Enfermedad | Si | `parse_enfermedad` | NOMCCC, NEM, NTR, NIN, NPT |
| 15 | Listado Embargos | Si | `parse_embargos` | NIN, NPT |
