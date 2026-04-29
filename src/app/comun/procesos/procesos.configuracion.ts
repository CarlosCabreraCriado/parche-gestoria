interface LibreriaProcesos {
  tipo: "proceso" | "directorio" | "redireccion";
  nombre: string;
  categoria:
    | "Remedy"
    | "Spool"
    | "Desarrollador"
    | "Santander"
    | "Google"
    | "Despacho"
    | "Import"
    | "KPIs"
    | "Documentos"
    | "Asesoria"
    | "Prueba"
    | "Fie"
    | "Duplicados"
    | "Autonomos"
    | "Pipeline"
    | "Facturacion";

  descripcion: string;
  autor?: string;
  opciones?: any;
  argumentos?: Argumentos[];
  salida?: Salida[];
  subCategoria?: LibreriaProcesos[];
}

interface Argumentos {
  tipo: tipoArgumento;
  identificador: string;
  obligado: boolean;
  formulario: FormularioArgumento;
  valor?: any;
}

interface Salida {
  tipo: tipoSalida;
  valor?: any;
}

interface FormularioArgumento {
  titulo: string;
  tipo:
    | "texto"
    | "ruta"
    | "numero"
    | "seleccion"
    | "fecha"
    | "boolean"
    | "archivo"
    | "objeto";
  placeholder: string;
  valorDefault: any;
  accept?: string;
}

type tipoSalida =
  | "string"
  | "boolean"
  | "spool"
  | "xlsxRaw"
  | "ruta"
  | "numero"
  | "fecha"
  | "texto";
type tipoArgumento =
  | "string"
  | "boolean"
  | "spool"
  | "xlsxRaw"
  | "ruta"
  | "numero"
  | "fecha"
  | "texto"
  | "objeto";

var libreriaProcesos: LibreriaProcesos[] = [
  {
    nombre: "Asesoria",
    categoria: "Asesoria",
    tipo: "directorio",
    descripcion: "Procesos de asesoría",
    subCategoria: [
      {
        nombre: "IRPF 2024",
        categoria: "Asesoria",
        tipo: "proceso",
        descripcion:
          "Obtiene los datos de los clientes mediante excel y calcular el IRPF correspondiente a la calculadora de la Agencia Tributaria 2024",
        autor: "Carlos Cabrera",
        argumentos: [
          {
            tipo: "texto",
            obligado: true,
            identificador: "googleChrome",
            formulario: {
              titulo: "Google .exe",
              tipo: "archivo",
              accept: ".exe, .EXE",
              placeholder: "Introduzca la ruta del ejecutable de Google",
              valorDefault: "",
            },
          },
          {
            tipo: "texto",
            obligado: true,
            identificador: "excelClientes",
            formulario: {
              titulo: "Excel de clientes",
              tipo: "archivo",
              accept: ".xlsm, .xlsx, .XLSX",
              placeholder:
                "Introduzca la ruta del archivo de datos de clientes.",
              valorDefault: "",
            },
          },
          {
            tipo: "texto",
            obligado: true,
            identificador: "rutaSalida",
            formulario: {
              titulo: "Directorio de salida:",
              tipo: "ruta",
              placeholder: "Ruta de guardado...",
              valorDefault: "",
            },
          },
        ],
        opciones: null,
        salida: [{ tipo: "boolean", valor: false }],
      },
      {
        nombre: "IRPF 2025",
        categoria: "Asesoria",
        tipo: "proceso",
        descripcion:
          "Obtiene los datos de los clientes mediante excel y calcular el IRPF correspondiente a la calculadora de la Agencia Tributaria de 2025",
        autor: "Carlos Cabrera",
        argumentos: [
          {
            tipo: "texto",
            obligado: true,
            identificador: "googleChrome",
            formulario: {
              titulo: "Google .exe",
              tipo: "archivo",
              accept: ".exe, .EXE",
              placeholder: "Introduzca la ruta del ejecutable de Google",
              valorDefault: "",
            },
          },
          {
            tipo: "texto",
            obligado: true,
            identificador: "excelClientes",
            formulario: {
              titulo: "Excel de clientes",
              tipo: "archivo",
              accept: ".xlsm, .xlsx, .XLSX",
              placeholder:
                "Introduzca la ruta del archivo de datos de clientes.",
              valorDefault: "",
            },
          },
          {
            tipo: "texto",
            obligado: true,
            identificador: "rutaSalida",
            formulario: {
              titulo: "Directorio de salida:",
              tipo: "ruta",
              placeholder: "Ruta de guardado...",
              valorDefault: "",
            },
          },
        ],
        opciones: null,
        salida: [{ tipo: "boolean", valor: false }],
      },
      {
        nombre: "IRPF 2026",
        categoria: "Asesoria",
        tipo: "proceso",
        descripcion:
          "Obtiene los datos de los clientes mediante excel y calcular el IRPF correspondiente a la calculadora de la Agencia Tributaria de 2026",
        autor: "Carlos Cabrera",
        argumentos: [
          {
            tipo: "texto",
            obligado: true,
            identificador: "googleChrome",
            formulario: {
              titulo: "Google .exe",
              tipo: "archivo",
              accept: ".exe, .EXE",
              placeholder: "Introduzca la ruta del ejecutable de Google",
              valorDefault: "",
            },
          },
          {
            tipo: "texto",
            obligado: true,
            identificador: "excelClientes",
            formulario: {
              titulo: "Excel de clientes",
              tipo: "archivo",
              accept: ".xlsm, .xlsx, .XLSX",
              placeholder:
                "Introduzca la ruta del archivo de datos de clientes.",
              valorDefault: "",
            },
          },
          {
            tipo: "texto",
            obligado: true,
            identificador: "rutaSalida",
            formulario: {
              titulo: "Directorio de salida:",
              tipo: "ruta",
              placeholder: "Ruta de guardado...",
              valorDefault: "",
            },
          },
        ],
        opciones: null,
        salida: [{ tipo: "boolean", valor: false }],
      },
      {
        nombre: "Cartas de pago en Hacienda",
        categoria: "Asesoria",
        tipo: "proceso",
        descripcion:
          "Obtiene los datos de los clientes mediante excel y genera las cartas de pago correspondientes",
        autor: "Carlos Cabrera",
        argumentos: [
          {
            tipo: "texto",
            obligado: true,
            identificador: "googleChrome",
            formulario: {
              titulo: "Google .exe",
              tipo: "archivo",
              accept: ".exe, .EXE",
              placeholder: "Introduzca la ruta del ejecutable de Google",
              valorDefault: "",
            },
          },
          {
            tipo: "texto",
            obligado: true,
            identificador: "excelClientes",
            formulario: {
              titulo: "Excel de clientes",
              tipo: "archivo",
              accept: ".xlsm, .xlsx, .XLSX",
              placeholder:
                "Introduzca la ruta del archivo de datos de clientes.",
              valorDefault: "",
            },
          },
          {
            tipo: "texto",
            obligado: true,
            identificador: "rutaSalida",
            formulario: {
              titulo: "Directorio de salida:",
              tipo: "ruta",
              placeholder: "Ruta de guardado...",
              valorDefault: "",
            },
          },
        ],
        opciones: null,
        salida: [{ tipo: "boolean", valor: false }],
      },
      {
        nombre: "Etiquetas AEAT",
        categoria: "Asesoria",
        tipo: "proceso",
        descripcion: "Descarga automática de etiquetas de la AEAT",
        autor: "Carlos Cabrera",
        argumentos: [
          {
            tipo: "texto",
            obligado: true,
            identificador: "googleChrome",
            formulario: {
              titulo: "Google .exe",
              tipo: "archivo",
              accept: ".exe, .EXE",
              placeholder: "Introduzca la ruta del ejecutable de Google",
              valorDefault: "",
            },
          },
          {
            tipo: "texto",
            obligado: true,
            identificador: "excelClientes",
            formulario: {
              titulo: "Excel de clientes",
              tipo: "archivo",
              accept: ".xlsm, .xlsx, .XLSX",
              placeholder:
                "Introduzca la ruta del archivo de datos de clientes.",
              valorDefault: "",
            },
          },
          {
            tipo: "texto",
            obligado: true,
            identificador: "rutaSalida",
            formulario: {
              titulo: "Directorio de salida:",
              tipo: "ruta",
              placeholder: "Ruta de guardado...",
              valorDefault: "",
            },
          },
        ],
        opciones: null,
        salida: [{ tipo: "boolean", valor: false }],
      },
      {
        nombre: "Cambio base de cotización",
        categoria: "Asesoria",
        tipo: "proceso",
        descripcion:
          "Cambia la base de cotización de los trabajadores en la Seguridad Social, mediante una plantilla de excel.",
        autor: "Carlos Cabrera",
        argumentos: [
          {
            tipo: "texto",
            obligado: true,
            identificador: "googleChrome",
            formulario: {
              titulo: "Google .exe",
              tipo: "archivo",
              accept: ".exe, .EXE",
              placeholder: "Introduzca la ruta del ejecutable de Google",
              valorDefault: "",
            },
          },
          {
            tipo: "texto",
            obligado: true,
            identificador: "excelClientes",
            formulario: {
              titulo: "Excel de clientes",
              tipo: "archivo",
              accept: ".xlsm, .xlsx, .XLSX",
              placeholder:
                "Introduzca la ruta del archivo de datos de clientes.",
              valorDefault: "",
            },
          },
          {
            tipo: "texto",
            obligado: true,
            identificador: "rutaSalida",
            formulario: {
              titulo: "Directorio de salida:",
              tipo: "ruta",
              placeholder: "Ruta de guardado...",
              valorDefault: "",
            },
          },
        ],
        opciones: null,
        salida: [{ tipo: "boolean", valor: false }],
      },
      {
        nombre: "Actualización CNAE25",
        categoria: "Asesoria",
        tipo: "proceso",
        descripcion: "",
        autor: "Carlos Cabrera",
        argumentos: [
          {
            tipo: "texto",
            obligado: true,
            identificador: "googleChrome",
            formulario: {
              titulo: "Google .exe",
              tipo: "archivo",
              accept: ".exe, .EXE",
              placeholder: "Introduzca la ruta del ejecutable de Google",
              valorDefault: "",
            },
          },
          {
            tipo: "texto",
            obligado: true,
            identificador: "excelClientes",
            formulario: {
              titulo: "Excel de clientes",
              tipo: "archivo",
              accept: ".xlsm, .xlsx, .XLSX",
              placeholder:
                "Introduzca la ruta del archivo de datos de clientes.",
              valorDefault: "",
            },
          },
          {
            tipo: "texto",
            obligado: true,
            identificador: "rutaSalida",
            formulario: {
              titulo: "Directorio de salida:",
              tipo: "ruta",
              placeholder: "Ruta de guardado...",
              valorDefault: "",
            },
          },
        ],
        opciones: null,
        salida: [{ tipo: "boolean", valor: false }],
      },
      {
        nombre: "CNAE25 Autónomos",
        categoria: "Asesoria",
        tipo: "proceso",
        descripcion: "",
        autor: "Carlos Cabrera",
        argumentos: [
          {
            tipo: "texto",
            obligado: true,
            identificador: "googleChrome",
            formulario: {
              titulo: "Google .exe",
              tipo: "archivo",
              accept: ".exe, .EXE",
              placeholder: "Introduzca la ruta del ejecutable de Google",
              valorDefault: "",
            },
          },
          {
            tipo: "texto",
            obligado: true,
            identificador: "excelClientes",
            formulario: {
              titulo: "Excel de clientes",
              tipo: "archivo",
              accept: ".xlsm, .xlsx, .XLSX",
              placeholder:
                "Introduzca la ruta del archivo de datos de clientes.",
              valorDefault: "",
            },
          },
          {
            tipo: "texto",
            obligado: true,
            identificador: "rutaSalida",
            formulario: {
              titulo: "Directorio de salida:",
              tipo: "ruta",
              placeholder: "Ruta de guardado...",
              valorDefault: "",
            },
          },
        ],
        opciones: null,
        salida: [{ tipo: "boolean", valor: false }],
      },
      {
        nombre: "Informes ITA",
        categoria: "Asesoria",
        tipo: "proceso",
        descripcion: "",
        autor: "Carlos Cabrera",
        argumentos: [
          {
            tipo: "texto",
            obligado: true,
            identificador: "googleChrome",
            formulario: {
              titulo: "Google .exe",
              tipo: "archivo",
              accept: ".exe, .EXE",
              placeholder: "Introduzca la ruta del ejecutable de Google",
              valorDefault: "",
            },
          },
          {
            tipo: "texto",
            obligado: true,
            identificador: "excelClientes",
            formulario: {
              titulo: "Excel con CCC en primera columna",
              tipo: "archivo",
              accept: ".xlsm, .xlsx, .XLSX",
              placeholder:
                "Introduzca la ruta del archivo de datos de clientes.",
              valorDefault: "",
            },
          },
          {
            tipo: "texto",
            obligado: false,
            identificador: "codigoEmpresa",
            formulario: {
              titulo: "Código de empresa (Dejar vacío para procesar todos)",
              tipo: "texto",
              placeholder: "Ej: 0061, 52; 8 0140-72",
              valorDefault: "",
            },
          },
          {
            tipo: "texto",
            obligado: true,
            identificador: "rutaSalida",
            formulario: {
              titulo: "Directorio de salida:",
              tipo: "ruta",
              placeholder: "Ruta de guardado...",
              valorDefault: "",
            },
          },
        ],
        opciones: null,
        salida: [{ tipo: "boolean", valor: false }],
      },
      {
        nombre: "Certificados de estar al corriente",
        categoria: "Asesoria",
        tipo: "proceso",
        descripcion:
          "Descarga 1..3 certificados (Seguridad Social, Tributario, ATC) en una sola ejecución usando el mismo Excel.",
        autor: "Carlos Cabrera",
        argumentos: [
          {
            tipo: "texto",
            obligado: true,
            identificador: "googleChrome",
            formulario: {
              titulo: "Google .exe",
              tipo: "archivo",
              accept: ".exe, .EXE",
              placeholder: "Introduzca la ruta del ejecutable de Google",
              valorDefault: "",
            },
          },
          {
            tipo: "texto",
            obligado: true,
            identificador: "excelClientes",
            formulario: {
              titulo: "Excel con CCC",
              tipo: "archivo",
              accept: ".xlsm, .xlsx, .XLSX",
              placeholder:
                "Introduzca la ruta del archivo de datos de clientes.",
              valorDefault: "",
            },
          },
          {
            tipo: "texto",
            obligado: false,
            identificador: "codigoEmpresa",
            formulario: {
              titulo: "Código de empresa (Dejar vacío para procesar todos)",
              tipo: "texto",
              placeholder: "Ej: 0061, 52; 8 0140-72",
              valorDefault: "",
            },
          },
          {
            tipo: "texto",
            obligado: true,
            identificador: "rutaSalida",
            formulario: {
              titulo: "Directorio de salida:",
              tipo: "ruta",
              placeholder: "Ruta de guardado...",
              valorDefault: "",
            },
          },
          {
            tipo: "boolean",
            obligado: false,
            identificador: "certSS",
            formulario: {
              titulo: "Seguridad Social",
              tipo: "boolean",
              placeholder: "",
              valorDefault: true,
            },
          },
          {
            tipo: "boolean",
            obligado: false,
            identificador: "certTributario",
            formulario: {
              titulo: "Tributario (AEAT)",
              tipo: "boolean",
              placeholder: "",
              valorDefault: true,
            },
          },
          {
            tipo: "boolean",
            obligado: false,
            identificador: "certATC",
            formulario: {
              titulo: "Subvenciones ATC",
              tipo: "boolean",
              placeholder: "",
              valorDefault: true,
            },
          },
        ],
        opciones: null,
        salida: [{ tipo: "boolean", valor: false }],
      },
      {
        nombre: "Certificado Seguridad Social",
        categoria: "Asesoria",
        tipo: "proceso",
        descripcion:
          "[Deprecado: usa 'Certificados de estar al corriente']",
        autor: "Carlos Cabrera",
        argumentos: [
          {
            tipo: "texto",
            obligado: true,
            identificador: "googleChrome",
            formulario: {
              titulo: "Google .exe",
              tipo: "archivo",
              accept: ".exe, .EXE",
              placeholder: "Introduzca la ruta del ejecutable de Google",
              valorDefault: "",
            },
          },
          {
            tipo: "texto",
            obligado: true,
            identificador: "excelClientes",
            formulario: {
              titulo: "Excel con CCC",
              tipo: "archivo",
              accept: ".xlsm, .xlsx, .XLSX",
              placeholder:
                "Introduzca la ruta del archivo de datos de clientes.",
              valorDefault: "",
            },
          },
          {
            tipo: "texto",
            obligado: false,
            identificador: "codigoEmpresa",
            formulario: {
              titulo: "Código de empresa (Dejar vacío para procesar todos)",
              tipo: "texto",
              placeholder: "Ej: 0061, 52; 8 0140-72",
              valorDefault: "",
            },
          },
          {
            tipo: "texto",
            obligado: true,
            identificador: "rutaSalida",
            formulario: {
              titulo: "Directorio de salida:",
              tipo: "ruta",
              placeholder: "Ruta de guardado...",
              valorDefault: "",
            },
          },
        ],
        opciones: null,
        salida: [{ tipo: "boolean", valor: false }],
      },
      {
        nombre: "Certificado Tributario",
        categoria: "Asesoria",
        tipo: "proceso",
        descripcion:
          "[Deprecado: usa 'Certificados de estar al corriente']",
        autor: "Carlos Cabrera",
        argumentos: [
          {
            tipo: "texto",
            obligado: true,
            identificador: "googleChrome",
            formulario: {
              titulo: "Google .exe",
              tipo: "archivo",
              accept: ".exe, .EXE",
              placeholder: "Introduzca la ruta del ejecutable de Google",
              valorDefault: "",
            },
          },
          {
            tipo: "texto",
            obligado: true,
            identificador: "excelClientes",
            formulario: {
              titulo: "Excel con CCC",
              tipo: "archivo",
              accept: ".xlsm, .xlsx, .XLSX",
              placeholder:
                "Introduzca la ruta del archivo de datos de clientes.",
              valorDefault: "",
            },
          },
          {
            tipo: "texto",
            obligado: false,
            identificador: "codigoEmpresa",
            formulario: {
              titulo: "Código de empresa (Dejar vacío para procesar todos)",
              tipo: "texto",
              placeholder: "Ej: 0061, 52; 8 0140-72",
              valorDefault: "",
            },
          },
          {
            tipo: "texto",
            obligado: true,
            identificador: "rutaSalida",
            formulario: {
              titulo: "Directorio de salida:",
              tipo: "ruta",
              placeholder: "Ruta de guardado...",
              valorDefault: "",
            },
          },
        ],
        opciones: null,
        salida: [{ tipo: "boolean", valor: false }],
      },
      {
        nombre: "Certificado Subvenciones ATC",
        categoria: "Asesoria",
        tipo: "proceso",
        descripcion:
          "[Deprecado: usa 'Certificados de estar al corriente']",
        autor: "Carlos Cabrera",
        argumentos: [
          {
            tipo: "texto",
            obligado: true,
            identificador: "googleChrome",
            formulario: {
              titulo: "Google .exe",
              tipo: "archivo",
              accept: ".exe, .EXE",
              placeholder: "Introduzca la ruta del ejecutable de Google",
              valorDefault: "",
            },
          },
          {
            tipo: "texto",
            obligado: true,
            identificador: "excelClientes",
            formulario: {
              titulo: "Excel con CCC",
              tipo: "archivo",
              accept: ".xlsm, .xlsx, .XLSX",
              placeholder:
                "Introduzca la ruta del archivo de datos de clientes.",
              valorDefault: "",
            },
          },
          {
            tipo: "texto",
            obligado: false,
            identificador: "codigoEmpresa",
            formulario: {
              titulo: "Código de empresa (Dejar vacío para procesar todos)",
              tipo: "texto",
              placeholder: "Ej: 0061, 52; 8 0140-72",
              valorDefault: "",
            },
          },
          {
            tipo: "texto",
            obligado: true,
            identificador: "rutaSalida",
            formulario: {
              titulo: "Directorio de salida:",
              tipo: "ruta",
              placeholder: "Ruta de guardado...",
              valorDefault: "",
            },
          },
        ],
        opciones: null,
        salida: [{ tipo: "boolean", valor: false }],
      },
      {
        nombre: "Formatear recibos de liquidacion",
        categoria: "Asesoria",
        tipo: "proceso",
        descripcion: "",
        autor: "Carlos Cabrera",
        argumentos: [
          {
            tipo: "texto",
            obligado: true,
            identificador: "excelClientes",
            formulario: {
              titulo: "Excel de clientes",
              tipo: "archivo",
              accept: ".xlsm, .xlsx, .XLSX",
              placeholder:
                "Introduzca la ruta del archivo de datos de clientes.",
              valorDefault: "",
            },
          },
          {
            tipo: "texto",
            obligado: true,
            identificador: "rutaSalida",
            formulario: {
              titulo: "Directorio de PDFs:",
              tipo: "ruta",
              placeholder: "Carpeta con los PDFs de SILTRA",
              valorDefault: "",
            },
          },
        ],
        opciones: null,
        salida: [{ tipo: "boolean", valor: false }],
      },
    ],
  },
  /*
  {
    nombre: "Test Excel",
    categoria: "Prueba",
    tipo: "directorio",
    descripcion: "Procesos de asesoría",
    subCategoria: [
      {
        nombre: "Test Excel",
        categoria: "Prueba",
        tipo: "proceso",
        descripcion:
          "Obtiene los nombres de los clientes mediante excel y los pasa a mayúsculas",
        autor: "Gonzalo",
        argumentos: [
          {
            tipo: "texto",
            obligado: true,
            identificador: "excelClientes",
            formulario: {
              titulo: "Excel de clientes",
              tipo: "archivo",
              accept: ".xlsm, .xlsx, .XLSX",
              placeholder:
                "Introduzca la ruta del archivo de datos de clientes.",
              valorDefault: "",
            },
          },
          {
            tipo: "texto",
            obligado: true,
            identificador: "rutaSalida",
            formulario: {
              titulo: "Directorio de salida:",
              tipo: "ruta",
              placeholder: "Ruta de guardado...",
              valorDefault: "",
            },
          },
        ],
        opciones: null,
        salida: [{ tipo: "boolean", valor: false }],
      },
    ],
  },
    */
  {
    nombre: "FIE",
    categoria: "Fie",
    tipo: "directorio",
    descripcion: "Procesos de asesoría",
    subCategoria: [
      {
        nombre: "FIE",
        categoria: "Fie",
        tipo: "proceso",
        descripcion: "Proceso FIE",
        autor: "Carlos Cabrera",
        argumentos: [
          {
            tipo: "texto",
            obligado: true,
            identificador: "exelFie",
            formulario: {
              titulo: "Excel FIE",
              tipo: "archivo",
              accept: ".xlsm, .xlsx, .XLSX",
              placeholder: "Introduzca la ruta del archivo de datos FIE.",
              valorDefault: "",
            },
          },

          {
            tipo: "texto",
            obligado: true,
            identificador: "exelEmpresas",
            formulario: {
              titulo: "Excel con CCC",
              tipo: "archivo",
              accept: ".xlsm, .xlsx, .XLSX",
              placeholder: "Introduzca la ruta del archivo...",
              valorDefault: "",
            },
          },

          {
            tipo: "texto",
            obligado: true,
            identificador: "exelEnfermedad",
            formulario: {
              titulo: "Excel 01 Enfermedad",
              tipo: "archivo",
              accept: ".xlsm, .xlsx, .XLSX",
              placeholder: "Introduzca la ruta del archivo...",
              valorDefault: "",
            },
          },
          {
            tipo: "texto",
            obligado: true,
            identificador: "exelAccidentes",
            formulario: {
              titulo: "Excel 02 Accidentes",
              tipo: "archivo",
              accept: ".xlsm, .xlsx, .XLSX",
              placeholder: "Introduzca la ruta del archivo...",
              valorDefault: "",
            },
          },
          {
            tipo: "texto",
            obligado: true,
            identificador: "rutaSalida",
            formulario: {
              titulo: "Directorio de salida:",
              tipo: "ruta",
              placeholder: "Ruta de guardado...",
              valorDefault: "",
            },
          },
        ],
        opciones: null,
        salida: [{ tipo: "boolean", valor: false }],
      },
      {
        nombre: "FIE_2",
        categoria: "Fie",
        tipo: "proceso",
        descripcion: "Proceso FIE_2",
        autor: "Gonzalo",
        argumentos: [
          {
            tipo: "texto",
            obligado: true,
            identificador: "googleChrome",
            formulario: {
              titulo: "Google .exe",
              tipo: "archivo",
              accept: ".exe, .EXE",
              placeholder: "Introduzca la ruta del ejecutable de Google",
              valorDefault: "",
            },
          },
          {
            tipo: "texto",
            obligado: true,
            identificador: "exelFie",
            formulario: {
              titulo: "Excel FIE",
              tipo: "archivo",
              accept: ".xlsm, .xlsx, .XLSX",
              placeholder: "Introduzca la ruta del archivo de datos FIE.",
              valorDefault: "",
            },
          },
          {
            tipo: "texto",
            obligado: true,
            identificador: "rutaSalida",
            formulario: {
              titulo: "Directorio de salida:",
              tipo: "ruta",
              placeholder: "Ruta de guardado...",
              valorDefault: "",
            },
          },
        ],
        opciones: null,
        salida: [{ tipo: "boolean", valor: false }],
      },
    ],
  },
  {
    nombre: "Duplicados",
    categoria: "Duplicados",
    tipo: "directorio",
    descripcion: "Procesos de asesoría",
    subCategoria: [
      {
        // OJO: este "nombre" es el que se intenta ejecutar según tus logs
        nombre: "DUPLICADOS TA2+IDC",
        categoria: "Duplicados",
        tipo: "proceso",
        descripcion:
          "Por cada trabajador: descarga TA2 y a continuación IDC (mismo Excel).",
        autor: "Gonzalo Martín",
        argumentos: [
          {
            tipo: "texto",
            obligado: true,
            identificador: "googleChrome",
            formulario: {
              titulo: "Google .exe",
              tipo: "archivo",
              accept: ".exe, .EXE",
              placeholder: "Introduzca la ruta del ejecutable de Google",
              valorDefault: "",
            },
          },
          {
            tipo: "texto",
            obligado: true,
            identificador: "excelClientes",
            formulario: {
              titulo: "Excel de clientes",
              tipo: "archivo",
              accept: ".xlsm, .xlsx, .XLSX",
              placeholder:
                "Introduzca la ruta del archivo de datos de clientes.",
              valorDefault: "",
            },
          },
          {
            tipo: "texto",
            obligado: true,
            identificador: "regimen",
            formulario: {
              titulo: "Régimen (4 dígitos)",
              tipo: "texto",
              placeholder: "Ej: 0111",
              valorDefault: "0111",
            },
          },
          {
            tipo: "texto",
            obligado: true,
            identificador: "rutaSalida",
            formulario: {
              titulo: "Directorio de PDFs:",
              tipo: "ruta",
              placeholder: "Carpeta con los PDFs",
              valorDefault: "",
            },
          },
        ],
        opciones: null,
        salida: [{ tipo: "boolean", valor: false }],
      },
    ],
  },
  {
    nombre: "Bases y recibos al cobro autónomos",
    categoria: "Autonomos",
    tipo: "directorio",
    descripcion: "Procesos de asesoría",
    subCategoria: [
      {
        // OJO: este "nombre" es el que se intenta ejecutar según tus logs
        nombre: "Bases y recibos al cobro autónomos",
        categoria: "Autonomos",
        tipo: "proceso",
        descripcion: "Bases y recibos al cobro autónomos",
        autor: "Gonzalo Martín",
        argumentos: [
          {
            tipo: "texto",
            obligado: true,
            identificador: "googleChrome",
            formulario: {
              titulo: "Google .exe",
              tipo: "archivo",
              accept: ".exe, .EXE",
              placeholder: "Introduzca la ruta del ejecutable de Google",
              valorDefault: "",
            },
          },
          {
            tipo: "texto",
            obligado: true,
            identificador: "excelClientes",
            formulario: {
              titulo: "Excel de clientes",
              tipo: "archivo",
              accept: ".xlsm, .xlsx, .XLSX",
              placeholder:
                "Introduzca la ruta del archivo de datos de clientes.",
              valorDefault: "",
            },
          },
          {
            tipo: "texto",
            obligado: true,
            identificador: "ejercicio_economico",
            formulario: {
              titulo: "Ejercicio económico (AAAA)",
              tipo: "texto",
              placeholder: "",
              valorDefault: "2025",
            },
          },
          {
            tipo: "texto",
            obligado: true,
            identificador: "rutaSalida",
            formulario: {
              titulo: "Directorio de PDFs:",
              tipo: "ruta",
              placeholder: "Carpeta con los PDFs",
              valorDefault: "",
            },
          },
        ],
        opciones: null,
        salida: [{ tipo: "boolean", valor: false }],
      },
    ],
  },
  {
    nombre: "Pipeline",
    categoria: "Pipeline",
    tipo: "directorio",
    descripcion: "Pipelines integrados: genera informe A3 + ejecuta proceso",
    subCategoria: [
      {
        nombre: "PIPELINE ALTAS DUPLICADOS",
        categoria: "Pipeline",
        tipo: "proceso",
        descripcion:
          "Genera listado Altas (Fmt 8) desde A3 y ejecuta Duplicados TA2+IDC automáticamente",
        autor: "Integración A3",
        argumentos: [
          {
            tipo: "texto",
            obligado: true,
            identificador: "googleChrome",
            formulario: {
              titulo: "Google .exe",
              tipo: "archivo",
              accept: ".exe, .EXE",
              placeholder: "Introduzca la ruta del ejecutable de Google",
              valorDefault: "",
            },
          },
          {
            tipo: "texto",
            obligado: true,
            identificador: "empresaCodes",
            formulario: {
              titulo: "Códigos de empresa (separados por coma)",
              tipo: "texto",
              placeholder: "Ej: 00008, 01378",
              valorDefault: "",
            },
          },
          {
            tipo: "texto",
            obligado: true,
            identificador: "regimen",
            formulario: {
              titulo: "Régimen (4 dígitos)",
              tipo: "texto",
              placeholder: "Ej: 0111",
              valorDefault: "0111",
            },
          },
          {
            tipo: "texto",
            obligado: true,
            identificador: "rutaSalida",
            formulario: {
              titulo: "Directorio de salida:",
              tipo: "ruta",
              placeholder: "Carpeta de salida",
              valorDefault: "",
            },
          },
          {
            tipo: "texto",
            obligado: false,
            identificador: "pythonPath",
            formulario: {
              titulo: "Ruta Python (vacío = usar PATH)",
              tipo: "texto",
              placeholder: "python",
              valorDefault: "python",
            },
          },
          {
            tipo: "texto",
            obligado: false,
            identificador: "analisisA3Path",
            formulario: {
              titulo: "Ruta proyecto analisis-a3",
              tipo: "ruta",
              placeholder: "Carpeta raíz de analisis-a3",
              valorDefault:
                "C:\\Users\\preprod\\Documents\\Proyectos\\analisis-a3",
            },
          },
        ],
        opciones: null,
        salida: [{ tipo: "boolean", valor: false }],
      },
    ],
  },
  {
    nombre: "Facturación",
    categoria: "Facturacion",
    tipo: "directorio",
    descripcion: "Reportes de facturación por empresa",
    subCategoria: [
      {
        nombre: "Reporte de Facturación",
        categoria: "Facturacion",
        tipo: "proceso",
        descripcion:
          "Genera un Excel con todos los procesos ejecutados en el periodo, desglosados por empresa, para facturación.",
        autor: "Gonzalo Martín",
        argumentos: [
          {
            tipo: "fecha",
            obligado: true,
            identificador: "desde",
            formulario: {
              titulo: "Desde (fecha inicio)",
              tipo: "fecha",
              placeholder: "Fecha de inicio del periodo",
              valorDefault: "",
            },
          },
          {
            tipo: "fecha",
            obligado: true,
            identificador: "hasta",
            formulario: {
              titulo: "Hasta (fecha fin)",
              tipo: "fecha",
              placeholder: "Fecha de fin del periodo",
              valorDefault: "",
            },
          },
        ],
        opciones: null,
        salida: [{ tipo: "boolean", valor: false }],
      },
    ],
  },
];

export { LibreriaProcesos, libreriaProcesos };
