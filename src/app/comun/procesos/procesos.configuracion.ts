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
    | "Asesoria";
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
        nombre: "Cálculo de IRPF",
        categoria: "Asesoria",
        tipo: "proceso",
        descripcion:
          "Obtiene los datos de los clientes mediante excel y calcular el IRPF correspondiente a la calculadora de la Agencia Tributaria",
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
    ],
  },
];

export { LibreriaProcesos, libreriaProcesos };
