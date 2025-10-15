interface LibreriaPlantillas {
  tipo: "proceso" | "directorio" | "redireccion";
  nombre: string;
  categoria: "Excel" | "Docx" | "Personalizado";
  descripcion: string;
  autor?: string;
  opciones?: any;
  argumentos?: Argumentos[];
  salida?: Salida[];
  subCategoria?: LibreriaPlantillas[];
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

var libreriaPlantillas: LibreriaPlantillas[] = [
  {
    nombre: "Excel",
    tipo: "directorio",
    categoria: "Excel",
    descripcion: "Plantillas Excel",
    subCategoria: [
      {
        nombre: "Infolex",
        categoria: "Excel",
        tipo: "proceso",
        descripcion: "Añade un registro de Infolex a un archivo Excel.",
        autor: "Carlos Cabrera",
        argumentos: [
          {
            tipo: "texto",
            obligado: true,
            identificador: "spoolInput",
            formulario: {
              titulo: "Ruta del Archivo Excel",
              tipo: "archivo",
              accept: ".xlsx",
              placeholder: "Introduzca la ruta del archivo de entrada.",
              valorDefault: "",
            },
          },
          {
            tipo: "texto",
            obligado: true,
            identificador: "nombreSpoolOutput",
            formulario: {
              titulo: "Nombre archivo salida",
              tipo: "texto",
              placeholder: "Introduzca la ruta del archivo de salida.",
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
    nombre: "Docx",
    categoria: "Docx",
    tipo: "directorio",
    descripcion: "Plantillas Microsoft Word",
    subCategoria: [
      {
        nombre: "Generica",
        categoria: "Docx",
        tipo: "proceso",
        descripcion: "Genera un documento a partir de un plantilla DOCX",
        autor: "Carlos Cabrera",
        argumentos: [
          {
            tipo: "texto",
            obligado: true,
            identificador: "Documento Docx",
            formulario: {
              titulo: "Ruta del Archivo Docx",
              tipo: "archivo",
              accept: ".docx",
              placeholder: "Introduzca la ruta del archivo de entrada.",
              valorDefault: "",
            },
          },
          {
            tipo: "texto",
            obligado: true,
            identificador: "nombreSpoolOutput",
            formulario: {
              titulo: "Nombre archivo salida",
              tipo: "texto",
              placeholder: "Introduzca la ruta del archivo de salida.",
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
    nombre: "Personalizado",
    categoria: "Personalizado",
    tipo: "directorio",
    descripcion: "Plantillas personalizadas",
    subCategoria: [
      {
        nombre: "Plantilla InfoLex",
        categoria: "Personalizado",
        tipo: "proceso",
        descripcion: "Generación de excel para importación en InfoLex.",
        autor: "Carlos Cabrera",
        argumentos: [
          {
            tipo: "texto",
            obligado: true,
            identificador: "querry",
            formulario: {
              titulo: "Querry de consulta de correo",
              tipo: "texto",
              placeholder: "Querry",
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

export { LibreriaPlantillas, libreriaPlantillas };
