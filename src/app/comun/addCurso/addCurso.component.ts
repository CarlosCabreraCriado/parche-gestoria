import {
  Component,
  OnInit,
  Input,
  ViewChild,
  AfterViewInit,
  ViewEncapsulation,
} from "@angular/core";
import { AppService } from "../../app.service";
import {
  animate,
  state,
  style,
  transition,
  trigger,
} from "@angular/animations";
import { MatTableDataSource } from "@angular/material/table";
import { MatPaginator } from "@angular/material/paginator";
import { DialogoComponent } from "../dialogos/dialogos.component";
import { MatDialog, MatDialogRef } from "@angular/material/dialog";
import { Observable } from "rxjs";
import { map, startWith } from "rxjs/operators";
import {
  UntypedFormControl,
  UntypedFormArray,
  FormGroup,
} from "@angular/forms";
import { FormControl, FormsModule, ReactiveFormsModule } from "@angular/forms";
import { MatSlideToggleModule } from "@angular/material/slide-toggle";

//Modulos:
import { CommonModule } from "@angular/common";
import { MatDatepickerModule } from "@angular/material/datepicker";
import { MatSelectModule } from "@angular/material/select";
import { MatTableModule } from "@angular/material/table";
import { MatIconModule } from "@angular/material/icon";
import { MatPaginatorModule } from "@angular/material/paginator";
import { MatAutocompleteModule } from "@angular/material/autocomplete";
import { MatInputModule } from "@angular/material/input";
import { MatBadgeModule } from "@angular/material/badge";
import { MatButtonToggleModule } from "@angular/material/button-toggle";
import { MatButtonModule } from "@angular/material/button";

import moment from "moment";

export interface Formador {
  nombre: string;
  id: string;
}

export interface Institucion {
  institucion: string;
  id: string;
}

export interface FiltroCursos {
  filtroMaestro: boolean;
  filtroGeneral: null | string;
  filtroError: boolean;
  filtroWarning: boolean;
  filtroProgramada: boolean;
  filtroModificado: boolean;
  filtroFecha: any;
  filtroCodigoCurso: null | string;
}

export interface FiltroFormadores {
  filtroMaestro: boolean;
  filtroGeneral: null | string;
  filtroError: boolean;
  filtroWarning: boolean;
  filtroModificado: boolean;
  filtroCodigoFormador: null | string;
}

export interface FiltroInstituciones {
  filtroMaestro: boolean;
  filtroGeneral: null | string;
  filtroError: boolean;
  filtroWarning: boolean;
  filtroModificado: boolean;
  filtroCodigoInstitucion: null | string;
}

@Component({
  standalone: true,
  encapsulation: ViewEncapsulation.None,
  imports: [
    MatButtonToggleModule,
    MatButtonModule,
    MatBadgeModule,
    FormsModule,
    ReactiveFormsModule,
    MatInputModule,
    MatDatepickerModule,
    MatAutocompleteModule,
    MatPaginatorModule,
    MatIconModule,
    CommonModule,
    MatSelectModule,
    MatTableModule,
    MatSlideToggleModule,
  ],
  selector: "addCursoComponent",
  templateUrl: "./addCurso.component.html",
  styleUrls: ["./addCurso.component.sass"],
  animations: [
    trigger("detailExpand", [
      state("collapsed", style({ height: "0px", minHeight: "0" })),
      state("expanded", style({ height: "*" })),
      transition(
        "expanded <=> collapsed",
        animate("225ms cubic-bezier(0.4, 0.0, 0.2, 1)"),
      ),
    ]),
  ],
})
export class AddCursoComponent implements OnInit {
  private cursos: any = [];
  private parametros: any = [];
  private codigoProvincia: any = [];
  private rutaArchivoCursos: any = "";
  private datosProyecto: any = {};
  public columnsToDisplay = ["cod_curso", "curso", "sesion", "institucion"];
  public columnsToDisplayFormador = ["cod_formador", "nombre", "territorial"];
  public columnsToDisplayInstitucion = [
    "cod_institucion",
    "nombre",
    "territorial",
  ];
  public columnsToDisplayCorreos = ["cod_curso", "institucion", "estado"];
  public columnsToDisplayWithExpand = [...this.columnsToDisplay, "expand"];
  public columnsToDisplayWithExpandFormador = [
    ...this.columnsToDisplayFormador,
    "expand",
  ];
  public columnsToDisplayWithExpandInstitucion = [
    ...this.columnsToDisplayInstitucion,
    "expand",
  ];
  public columnsToDisplayWithExpandCorreos = [
    ...this.columnsToDisplayCorreos,
    "expand",
  ];
  public expandedElement: any | null;

  //Tablas:
  public dataTable: any = new MatTableDataSource([]);
  public tablaFormadores: any = new MatTableDataSource([]);
  public tablaInstituciones: any = new MatTableDataSource([]);
  public tablaCorreos: any = new MatTableDataSource([]);

  //Metadatos:
  private metadatosCursos: any = {};
  private metadatosFormadores: any = {};
  private metadatosInstituciones: any = {};

  //Formadores:
  private formadores: any = [];
  private formadoresCurso: any = [];
  public autoFormadorControl = new UntypedFormControl("");
  private opcionesFormador: Formador[] = [];
  public filteredOptionsFormador: Observable<Formador[]>;

  //Instituciones:
  private instituciones: any = [];
  public autoInstitucionControl = new UntypedFormControl("");
  private opcionesInstituciones: Institucion[] = [];
  public filteredOptionsInstitucion: Observable<Institucion[]>;

  //Correos:
  private correos: any = [];
  public correosTemplate: any = [];
  public correosActualizados: boolean = false;
  public indexCorreoVisualizado: number = 0;
  public correosVisualizados: any = [];
  private plantillaCorreoInstitucion: any = null;
  private plantillaCorreoMaterial: any = null;
  private plantillaCorreoRecordatorio: any = null;
  private listaBorradores: any = [];

  //Tipología:
  private tipología: any = [];

  //CONFIGURACIONES:
  public reducirTablaCursos: boolean = true;
  public reducirTablaCorreos: boolean = true;
  private numeroCursosReducidos: number = 500;

  //Control de formularios:
  private formularioControl = new UntypedFormArray([]);

  //Filtros:
  private filtroCursos: FiltroCursos = {
    filtroMaestro: false,
    filtroGeneral: null,
    filtroError: false,
    filtroWarning: false,
    filtroProgramada: false,
    filtroModificado: false,
    filtroFecha: null,
    filtroCodigoCurso: null,
  };

  private filtroFormadores: FiltroFormadores = {
    filtroMaestro: false,
    filtroGeneral: null,
    filtroError: false,
    filtroWarning: false,
    filtroModificado: false,
    filtroCodigoFormador: null,
  };

  private filtroInstituciones: FiltroInstituciones = {
    filtroMaestro: false,
    filtroGeneral: null,
    filtroError: false,
    filtroWarning: false,
    filtroModificado: false,
    filtroCodigoInstitucion: null,
  };

  public filterButtonControl = new FormControl("");
  public filterButtonControlFormadores = new FormControl("");
  public filterButtonControlInstituciones = new FormControl("");

  public filtroSAP: boolean = false;
  public filtroRMCA: boolean = false;

  public filtroModificados: boolean = false;
  public filtroWarning: boolean = false;
  public filtroError: boolean = false;
  public filtroProgramadas: boolean = false;

  @Input() pantallaSeleccionada: any;

  @ViewChild("paginatorCursos", { static: false })
  paginatorCursos: MatPaginator;
  @ViewChild("paginatorFormadores", { static: false })
  paginatorFormadores: MatPaginator;
  @ViewChild("paginatorInstituciones", { static: false })
  paginatorInstituciones: MatPaginator;
  @ViewChild("paginatorCorreos", { static: false })
  paginatorCorreos: MatPaginator;

  @ViewChild("buscador", { static: false }) buscador: any;

  constructor(
    private appService: AppService,
    public dialog: MatDialog,
  ) {}

  async ngOnInit() {
    //Inicializa la tabla:
    //await this.inicializarDatos();

    const dialogoProcesandoInicializacion = this.dialog.open(DialogoComponent, {
      disableClose: true,
      data: {
        tipoDialogo: "procesando",
        titulo: "Importando Datos",
        contenido: "",
      },
    });

    dialogoProcesandoInicializacion.afterOpened().subscribe(() => {
      this.inicializarDatos(dialogoProcesandoInicializacion).then(
        (dialogo: MatDialogRef<DialogoComponent>) => {
          console.warn("FIN inicializacion");
          this.dataTable.filter = this.filtroCursos;
          this.tablaCorreos.filter = {};

          //Obtener Correo:
          this.correosTemplate = this.appService.obtenerCorreo();

          this.correosTemplate = [];

          this.datosProyecto = this.appService.getDatosProyecto();

          console.warn("DATOS PROYECTO: ", this.datosProyecto);
          console.error("CORREOS TEMPLATE: ", this.correosTemplate);

          dialogo.close();
        },
      );
    });
  } //Fin OnInit

  displayFn(user: Formador): string {
    return user && user.nombre ? user.nombre : "";
  }

  private _filterFormador(nombre: string): Formador[] {
    const filterValue = nombre.toLowerCase();
    return this.opcionesFormador.filter((option) =>
      option.nombre.toLowerCase().includes(filterValue),
    );
  }

  private _filterInstitucion(nombre: string): Institucion[] {
    const filterValue = nombre.toLowerCase();
    return this.opcionesInstituciones.filter((option) =>
      option.institucion.toLowerCase().includes(filterValue),
    );
  }

  comandoHerramienta(comando: string) {
    console.warn("Comando Recibido: ", comando);

    switch (comando) {
      case "nuevo":
        if (this.pantallaSeleccionada == "Cursos") {
          this.addCurso();
        }
        if (this.pantallaSeleccionada == "Formadores") {
          this.addFormador();
        }
        if (this.pantallaSeleccionada == "Instituciones") {
          this.addInstitucion();
        }
        break;
      case "verificar":
        this.checkError();
        break;
      case "descargar":
        this.descargarDatos2();
        break;
      case "subir":
        this.subirCursos();
        break;
      case "guardar":
        this.guardarCursos();
        break;
      case "refreshAutomatizacion":
        if (this.pantallaSeleccionada == "Automatizacion") {
          this.refreshAutomatizacion();
        }
        break;
      case "generarDocumentoComunidades":
        this.pedirDatosGeneracionDocumento();
        break;
    }
    return;
  }

  async inicializarDatos(dialogo?, config?) {
    return new Promise(async (resolve) => {
      console.warn(" --> INICIALIZANDO DATOS <--");

      if (!config || config["omitirRecarga"] == false) {
        this.parametros = await this.appService.getDato(
          "/Archivos",
          "Parametros",
        );
      }

      console.log("Parametros: ", this.parametros);
      console.log("Códigos Provincias: ", this.codigoProvincia);
      console.log("Formadores: ", this.formadores);
      console.log("Formadores-Curso: ", this.formadoresCurso);

      //Comprueba si hay ruta de cursos en parametros:
      if (!this.parametros || this.parametros.length == 0) {
        //Preguntar por la ruta del archivo:
        const dialogRef = this.dialog.open(DialogoComponent, {
          disableClose: true,
          data: {
            tipoDialogo: "inputFile",
            titulo: "Ruta del archivo excel",
            contenido: "Introduzca la ruta del archivo 'Monitorización.xlsx'.",
          },
        });

        dialogRef.afterClosed().subscribe((result) => {
          //Guardar ruta curso en parametros:
          console.log("Ruta del archivo: ");
          console.log(result);

          this.parametros = {
            nombreId: "Parametros",
            rutaCursos: result[0].path,
          };

          const dialogoProcesando = this.dialog.open(DialogoComponent, {
            disableClose: true,
            data: {
              tipoDialogo: "procesando",
              titulo: "Procesando",
              contenido: "",
            },
          });

          this.appService.guardarArchivo(this.parametros, dialogoProcesando);
          this.rutaArchivoCursos = this.parametros.rutaCursos;
          this.descargarDatos2();
        });

        console.warn("RESOLVE");
        resolve(dialogo);
        return;
      } else {
        console.warn("ENTRANDO");
        this.rutaArchivoCursos = this.parametros[0].rutaCursos;
      }
      console.warn("CONTINUANDO");

      if (!config || config["omitirRecarga"] == false) {
        console.time("Descargar Datos");
        this.codigoProvincia = await this.appService.getDato(
          "/Archivos",
          "Códigos_Provincia",
        );

        //Cargar Datos Cursos:
        this.cursos = await this.appService.getDato("/Archivos", "Cursos");
        this.metadatosCursos = await this.appService.getDato(
          "/Archivos",
          "Metadatos Cursos",
        );

        //Cargar Datos Formadores:
        this.formadores = await this.appService.getDato(
          "/Archivos",
          "Formadores",
        );
        this.formadoresCurso = await this.appService.getDato(
          "/Archivos",
          "Formador-Curso",
        );
        this.metadatosFormadores = await this.appService.getDato(
          "/Archivos",
          "Metadatos Formadores",
        );

        //Cargar Datos Instituciones:
        this.instituciones = await this.appService.getDato(
          "/Archivos",
          "Instituciones",
        );
        this.metadatosInstituciones = await this.appService.getDato(
          "/Archivos",
          "Metadatos Instituciones",
        );

        //Carga de Datos Tipologías:
        this.tipología = await this.appService.getDato(
          "/Archivos",
          "Tipología",
        );

        console.warn("Tipología: ", this.tipología);

        //Cargar Datos de los Correos:
        this.correos = await this.appService.getDato("/Archivos", "Correos");

        //Formateo de objeto Cursos:
        if (this.cursos[0]) {
          this.cursos = this.cursos[0].data;
        }

        //Formateo de objeto Correos:
        if (this.correos[0]) {
          this.correos = this.correos[0].data;
        }

        console.timeEnd("Descargar Datos");
      }

      console.warn("Correos: ", this.correos);

      //Crea los objetos de metadatos si no existen:
      if (!this.metadatosCursos[0]) {
        this.metadatosCursos = [
          {
            data: [],
            nombreId: "Metadatos Cursos",
            objetoId: "Metadatos",
          },
        ];
      }
      if (!this.metadatosFormadores[0]) {
        this.metadatosFormadores = [
          {
            data: [],
            nombreId: "Metadatos Formadores",
            objetoId: "Metadatos Formadores",
          },
        ];
      }
      if (!this.metadatosInstituciones[0]) {
        this.metadatosInstituciones = [
          {
            data: [],
            nombreId: "Metadatos Instituciones",
            objetoId: "Metadatos Instituciones",
          },
        ];
      }

      //Preprocesado del objeto de cursos:
      for (var i = 0; i < this.cursos.length; i++) {
        //Eliminar Registros Sin Codigo De Curso:
        if (this.cursos[i]["cod_curso"] === undefined) {
          this.cursos.splice(i, 1);
          i--;
        }
      }

      //Transformación de codigos a números:
      /*
        for(var i = 0; i < this.cursos.length ; i++){
            this.cursos[i]["cod_curso"] = parseInt(this.cursos[i]["cod_curso"])
        }

        for(var i = 0; i < this.instituciones[0].data.length ; i++){
            this.instituciones[0].data[i]["cod_institucion"] = parseInt(this.instituciones[0].data[i]["cod_institucion"])
        }

        for(var i = 0; i < this.metadatosInstituciones[0].data.length ; i++){
            this.metadatosInstituciones[0].data[i]["cod_institucion"] = parseInt(this.metadatosInstituciones[0].data[i]["cod_institucion"])
        }

        for(var i = 0; i < this.formadores[0].data.length ; i++){
            this.formadores[0].data[i]["cod__formador"] = parseInt(this.formadores[0].data[i]["cod__formador"])
        }

        for(var i = 0; i < this.metadatosFormadores[0].data.length ; i++){
            this.metadatosFormadores[0].data[i]["cod__formador"] = parseInt(this.metadatosFormadores[0].data[i]["cod__formador"])
        }
        */

      //Ordenado de Formadores, Cursos y Formadores-Curso:
      console.time("Ordenado Datos");
      try {
        this.cursos.sort(
          (a, b) => Number(a["cod_curso"]) - Number(b["cod_curso"]),
        );
        this.formadores[0].data.sort(
          (a, b) => Number(a["cod__formador"]) - Number(b["cod__formador"]),
        );
        this.formadoresCurso[0].data.sort(
          (a, b) => Number(a["cod_curso"]) - Number(b["cod_curso"]),
        );
        this.metadatosCursos[0].data.sort(
          (a, b) => Number(a["cod_curso"]) - Number(b["cod_curso"]),
        );
        this.metadatosFormadores[0].data.sort(
          (a, b) => Number(a["cod__formador"]) - Number(b["cod__formador"]),
        );
        this.instituciones[0].data.sort(
          (a, b) => Number(a["cod_institucion"]) - Number(b["cod_institucion"]),
        );
        this.metadatosInstituciones[0].data.sort(
          (a, b) => Number(a["cod_institucion"]) - Number(b["cod_institucion"]),
        );
      } catch (e) {
        console.error("Error al ordenar los datos...  Solicitando Descarga");
        this.descargarDatos2();
        resolve(dialogo);
      }
      console.timeEnd("Ordenado Datos");

      console.time("Añadir Metadatos Formadores e Instituciones");
      //Añadir metadatos de Formadores no encontrados:
      for (var i = 0; i < this.formadores[0].data.length; i++) {
        if (
          this.binarySearchObject(
            this.metadatosFormadores[0].data,
            "cod__formador",
            this.formadores[0].data[i]["cod__formador"],
          ) == -1
        ) {
          this.metadatosFormadores[0].data.push({
            cod__formador: this.formadores[0].data[i]["cod__formador"],
            formadorAdicional: false,
            flag_cambio: false,
            flag_eliminar: false,
            error: false,
            errores: {
              nombre: false,
              email: false,
              telefono: false,
              territorial: false,
              ccaa: false,
              fecha: false,
              estado: false,
              certificado: false,
              confidencialidad: false,
              consentimiento: false,
            },
            comentarios: [],
          });
        }
      }

      //Añadir metadatos de Instituciones no encontradas:
      for (var i = 0; i < this.instituciones[0].data.length; i++) {
        if (
          this.binarySearchObject(
            this.metadatosInstituciones[0].data,
            "cod_institucion",
            this.instituciones[0].data[i]["cod_institucion"],
          ) == -1
        ) {
          this.metadatosInstituciones[0].data.push({
            cod_institucion: this.instituciones[0].data[i]["cod_institucion"],
            intitucionAdicional: false,
            flag_cambio: false,
            flag_eliminar: false,
            error: false,
            errores: {
              nombre: false,
              email: false,
              telefono: false,
              territorial: false,
              ccaa: false,
              fecha: false,
              estado: false,
            },
            comentarios: [],
          });
        }
      }

      console.timeEnd("Añadir Metadatos Formadores e Instituciones");

      console.log("Meta Formadores:");
      console.log(this.metadatosFormadores);

      console.log("Meta Instituciones:");
      console.log(this.metadatosInstituciones);

      console.time("Añadir Metadatos Cursos");

      //Añadir metadatos de cursos no encontrados:
      for (var i = 0; i < this.cursos.length; i++) {
        if (
          this.binarySearchObject(
            this.metadatosCursos[0].data,
            "cod_curso",
            this.cursos[i]["cod_curso"],
          ) == -1
        ) {
          this.metadatosCursos[0].data.push({
            cod_curso: this.cursos[i]["cod_curso"],
            incidenciaAdicional: false,
            revisado: false,
            flag: 0,
            flag_cambio: false,
            flag_eliminar: false,
            modificado: false,
            formadorModificado: false,
            modificaciones: {},
            error: false,
            errores: {
              postal: false,
              ccaa: false,
              cod_grupo: false,
              territorial: false,
              institucion: false,
              fecha: false,
              hora_inicio: false,
              hora_fin: false,
              curso: false,
              sesion: false,
              colectivo: false,
              grupo: false,
              alumnos: false,
              formadores: false,
            },
            comentarios: [],
          });
        }
      }

      console.timeEnd("Añadir Metadatos Cursos");

      console.time("Relacion Formadores-Cursos");
      console.warn("Formador-Curso", this.formadoresCurso);

      //Cargar Datos De Formadores en Metadatos de Curso:
      if (typeof this.formadoresCurso[0] != "undefined") {
        var metadatosFormador = [];
        var nombre = "";
        var lastIndex = 0;

        for (var i = 0; i < this.metadatosCursos[0].data.length; i++) {
          //Si ya ha sido modificado no hacer nada:
          if (this.metadatosCursos[0].data[i]["formadorModificado"]) {
            continue;
          }

          metadatosFormador = [];

          for (
            var j = lastIndex;
            j < this.formadoresCurso[0].data.length;
            j++
          ) {
            //Si coinciden los codigos de curso:
            if (
              this.formadoresCurso[0].data[j]["cod_curso"] ==
              this.metadatosCursos[0].data[i]["cod_curso"]
            ) {
              //Buscar datos de formador:
              var indexFormador = this.binarySearchObject(
                this.formadores[0].data,
                "cod__formador",
                this.formadoresCurso[0].data[j]["cod__formador"],
              );

              if (indexFormador != -1) {
                metadatosFormador.push({
                  id: this.formadoresCurso[0].data[j]["cod__formador"],
                  nombre: this.formadores[0].data[indexFormador]["nombre"],
                });
              } else {
                //Si error buscando nombre por ID:
                metadatosFormador.push({
                  id: this.formadoresCurso[0].data[j]["cod__formador"],
                  nombre: "ERROR",
                });
              }
            } else if (
              this.formadoresCurso[0].data[j]["cod_curso"] >
              this.metadatosCursos[0].data[i]["cod_curso"]
            ) {
              lastIndex = j - 1;
              break;
            }
          }

          this.metadatosCursos[0].data[i]["formadores"] = metadatosFormador;
        }
      }
      //FIN RELACION FORMADORES-CURSOS:
      console.timeEnd("Relacion Formadores-Cursos");

      if (this.reducirTablaCursos) {
        console.warn("REDUCIENDO TABLA DE CURSOS");
        this.dataTable = Object.assign(
          [],
          this.cursos.slice(
            Math.max(this.cursos.length - this.numeroCursosReducidos, 0),
          ),
        );
      } else {
        this.dataTable = Object.assign([], this.cursos);
      }

      this.tablaFormadores = Object.assign([], this.formadores[0].data);
      this.tablaInstituciones = Object.assign([], this.instituciones[0].data);

      //CARGA DE DATOS DE CORREOS:
      this.procesarDatosCorreos();
      //this.tablaCorreos = Object.assign([],this.correos);

      console.time("Cursos Adicionales");
      //Cargar Cursos adicionales:
      for (var i = 0; i < this.metadatosCursos[0].data.length; i++) {
        if (
          this.metadatosCursos[0].data[i]["incidenciaAdicional"] &&
          this.dataTable.find(
            (j) => j.cod_curso == this.metadatosCursos[0].data[i]["cod_curso"],
          ) == undefined
        ) {
          this.dataTable.push({
            cod_curso: this.metadatosCursos[0].data[i]["cod_curso"],
            modalidad: this.metadatosCursos[0].data[i]["modalidad"],
            estado: this.metadatosCursos[0].data[i]["estado"],
            material: this.metadatosCursos[0].data[i]["material"],
            valoración: this.metadatosCursos[0].data[i]["valoración"],
            observaciones: this.metadatosCursos[0].data[i]["observaciones"],
            cod__postal: this.metadatosCursos[0].data[i]["cod__postal"],
            territorial: this.metadatosCursos[0].data[i]["territorial"],
            "ccaa_/_pais": this.metadatosCursos[0].data[i]["ccaa_/_pais"],
            institución: this.metadatosCursos[0].data[i]["institución"],
            fecha: this.metadatosCursos[0].data[i]["fecha"],
            hora_inicio: this.metadatosCursos[0].data[i]["hora_inicio"],
            hora_fin: this.metadatosCursos[0].data[i]["hora_fin"],
            fecha_formateada:
              this.metadatosCursos[0].data[i]["fecha_formateada"],
            hora_inicio_formateada:
              this.metadatosCursos[0].data[i]["hora_inicio_formateada"],
            hora_fin_formateada:
              this.metadatosCursos[0].data[i]["hora_fin_formateada"],
            duracion_formateada:
              this.metadatosCursos[0].data[i]["duracion_formateada"],
            cod_grupo: this.metadatosCursos[0].data[i]["cod_grupo"],
            curso: this.metadatosCursos[0].data[i]["curso"],
            sesión: this.metadatosCursos[0].data[i]["sesión"],
            colectivo: this.metadatosCursos[0].data[i]["colectivo"],
            grupo: this.metadatosCursos[0].data[i]["grupo"],
            nºasistentes: this.metadatosCursos[0].data[i]["nºasistentes"],
            metadatos: {
              incidenciaAdicional: false,
              flag_cambio: true,
              flag_eliminar: false,
              formadores: this.metadatosCursos[0].data[i]["formadores"],
              error: true,
            },
          });
          this.cursos.push({
            cod_curso: this.metadatosCursos[0].data[i]["cod_curso"],
            modalidad: this.metadatosCursos[0].data[i]["modalidad"],
            estado: this.metadatosCursos[0].data[i]["estado"],
            material: this.metadatosCursos[0].data[i]["material"],
            valoración: this.metadatosCursos[0].data[i]["valoración"],
            observaciones: this.metadatosCursos[0].data[i]["observaciones"],
            cod__postal: this.metadatosCursos[0].data[i]["cod__postal"],
            territorial: this.metadatosCursos[0].data[i]["territorial"],
            "ccaa_/_pais": this.metadatosCursos[0].data[i]["ccaa_/_pais"],
            institución: this.metadatosCursos[0].data[i]["institución"],
            fecha: this.metadatosCursos[0].data[i]["fecha"],
            hora_inicio: this.metadatosCursos[0].data[i]["hora_inicio"],
            hora_fin: this.metadatosCursos[0].data[i]["hora_fin"],
            fecha_formateada:
              this.metadatosCursos[0].data[i]["fecha_formateada"],
            hora_inicio_formateada:
              this.metadatosCursos[0].data[i]["hora_inicio_formateada"],
            hora_fin_formateada:
              this.metadatosCursos[0].data[i]["hora_fin_formateada"],
            duracion_formateada:
              this.metadatosCursos[0].data[i]["duracion_formateada"],
            cod_grupo: this.metadatosCursos[0].data[i]["cod_grupo"],
            curso: this.metadatosCursos[0].data[i]["curso"],
            sesión: this.metadatosCursos[0].data[i]["sesión"],
            colectivo: this.metadatosCursos[0].data[i]["colectivo"],
            grupo: this.metadatosCursos[0].data[i]["grupo"],
            nºasistentes: this.metadatosCursos[0].data[i]["nºasistentes"],
            metadatos: {
              incidenciaAdicional: false,
              flag_cambio: true,
              flag_eliminar: false,
              formadores: this.metadatosCursos[0].data[i]["formadores"],
              error: true,
            },
          });
          console.log(
            "Añadiendo " + this.metadatosCursos[0].data[i]["cod_curso"],
          );
        }
      }
      console.timeEnd("Cursos Adicionales");

      console.time("Formadores Adicionales");
      //Cargar Formadores Adicionales:
      for (var i = 0; i < this.metadatosFormadores[0].data.length; i++) {
        if (
          this.metadatosFormadores[0].data[i]["formadorAdicional"] &&
          this.tablaFormadores.find(
            (j) =>
              j.cod__formador ==
              this.metadatosFormadores[0].data[i]["cod__formador"],
          ) == undefined
        ) {
          this.tablaFormadores.push({
            cod__formador: this.metadatosFormadores[0].data[i]["cod__formador"],
            nombre: this.metadatosFormadores[0].data[i]["nombre"],
            estado: this.metadatosFormadores[0].data[i]["estado"],
            fecha: this.metadatosFormadores[0].data[i]["fecha"],
            territorial: this.metadatosFormadores[0].data[i]["territorial"],
            ccaa: this.metadatosFormadores[0].data[i]["ccaa"],
            email: this.metadatosFormadores[0].data[i]["email"],
            telefono: this.metadatosFormadores[0].data[i]["telefono"],
            certificado: this.metadatosFormadores[0].data[i]["certificado"],
            consentimiento:
              this.metadatosFormadores[0].data[i]["consentimiento"],
            confidencialidad:
              this.metadatosFormadores[0].data[i]["confidencialidad"],
            metadatos: {
              formadorAdicional: true,
              flag_cambio: true,
              flag_eliminar: false,
              error: true,
            },
          });

          this.formadores[0].data.push({
            cod__formador: this.metadatosFormadores[0].data[i]["cod__formador"],
            nombre: this.metadatosFormadores[0].data[i]["nombre"],
            estado: this.metadatosFormadores[0].data[i]["estado"],
            fecha: this.metadatosFormadores[0].data[i]["fecha"],
            territorial: this.metadatosFormadores[0].data[i]["territorial"],
            ccaa: this.metadatosFormadores[0].data[i]["ccaa"],
            email: this.metadatosFormadores[0].data[i]["email"],
            telefono: this.metadatosFormadores[0].data[i]["telefono"],
            certificado: this.metadatosFormadores[0].data[i]["certificado"],
            consentimiento:
              this.metadatosFormadores[0].data[i]["consentimiento"],
            confidencialidad:
              this.metadatosFormadores[0].data[i]["confidencialidad"],
            metadatos: {
              formadorAdicional: true,
              flag_cambio: true,
              flag_eliminar: false,
              error: true,
            },
          });
          console.log(
            "Añadiendo Formador: " +
              this.metadatosFormadores[0].data[i]["cod__formador"],
          );
        }
      }

      console.timeEnd("Formadores Adicionales");

      console.time("Instituciones Adicionales");
      //Cargar Instituciones Adicionales:
      for (var i = 0; i < this.metadatosInstituciones[0].data.length; i++) {
        if (
          this.metadatosInstituciones[0].data[i]["institucionAdicional"] &&
          this.tablaInstituciones.find(
            (j) =>
              j.cod_institucion ==
              this.metadatosInstituciones[0].data[i]["cod_institucion"],
          ) == undefined
        ) {
          this.tablaInstituciones.push({
            cod_institucion:
              this.metadatosInstituciones[0].data[i]["cod_institucion"],
            nombre: this.metadatosInstituciones[0].data[i]["nombre"],
            estado: this.metadatosInstituciones[0].data[i]["estado"],
            fecha: this.metadatosInstituciones[0].data[i]["fecha"],
            territorial: this.metadatosInstituciones[0].data[i]["territorial"],
            ccaa: this.metadatosInstituciones[0].data[i]["ccaa"],
            email: this.metadatosInstituciones[0].data[i]["email"],
            telefono: this.metadatosInstituciones[0].data[i]["telefono"],
            metadatos: {
              institucionAdicional: true,
              flag_cambio: true,
              flag_eliminar: false,
              error: true,
            },
          });

          this.instituciones[0].data.push({
            cod_institucion:
              this.metadatosInstituciones[0].data[i]["cod_institucion"],
            nombre: this.metadatosInstituciones[0].data[i]["nombre"],
            estado: this.metadatosInstituciones[0].data[i]["estado"],
            fecha: this.metadatosInstituciones[0].data[i]["fecha"],
            territorial: this.metadatosInstituciones[0].data[i]["territorial"],
            ccaa: this.metadatosInstituciones[0].data[i]["ccaa"],
            email: this.metadatosInstituciones[0].data[i]["email"],
            telefono: this.metadatosInstituciones[0].data[i]["telefono"],
            metadatos: {
              institucionAdicional: true,
              flag_cambio: true,
              flag_eliminar: false,
              error: true,
            },
          });
          console.log(
            "Añadiendo Insititución: " +
              this.metadatosInstituciones[0].data[i]["cod_institucion"],
          );
        }
      }

      console.timeEnd("Instituciones Adicionales");

      console.time("Formatos Fechas");
      //Añadir Campos Fechas Formateadas:
      for (var i = 0; i < this.cursos.length; i++) {
        //Fecha:
        if (!this.cursos[i]["fecha_formateada"]) {
          if (typeof this.cursos[i]["fecha"] == "number") {
            this.cursos[i]["fecha_formateada"] = this.ExcelDateToJSDate(
              this.cursos[i]["fecha"],
            );
          }
        }
        if (!this.cursos[i]["hora_inicio_formateada"]) {
          //Hora Inicio:
          var horaInicio = moment({
            hour: this.cursos[i]["hora_inicio"] * 24,
            minute:
              (this.cursos[i]["hora_inicio"] * 24 -
                Math.floor(this.cursos[i]["hora_inicio"] * 24)) *
              60,
          });
          this.cursos[i]["hora_inicio_formateada"] = horaInicio.format("HH:mm");
        }
        if (!this.cursos[i]["hora_fin_formateada"]) {
          //Hora Fin:
          var horaFin = moment({
            hour: this.cursos[i]["hora_fin"] * 24,
            minute:
              (this.cursos[i]["hora_fin"] * 24 -
                Math.floor(this.cursos[i]["hora_fin"] * 24)) *
              60,
          });
          this.cursos[i]["hora_fin_formateada"] = horaFin.format("HH:mm");
        }

        if (!this.cursos[i]["duracion_formateada"]) {
          //Duración:
          var horaInicio = moment({
            hour: this.cursos[i]["hora_inicio"] * 24,
            minute:
              (this.cursos[i]["hora_inicio"] * 24 -
                Math.floor(this.cursos[i]["hora_inicio"] * 24)) *
              60,
          });
          var horaFin = moment({
            hour: this.cursos[i]["hora_fin"] * 24,
            minute:
              (this.cursos[i]["hora_fin"] * 24 -
                Math.floor(this.cursos[i]["hora_fin"] * 24)) *
              60,
          });
          var diff = horaFin.diff(horaInicio);

          //Formateo Horas:
          this.cursos[i]["duracion_formateada"] = moment
            .utc(diff)
            .format("HH:mm");
        }
      }

      console.timeEnd("Formatos Fechas");

      console.time("Formularios Archivos Formadores");
      //Incluir formularios de archivos Formadores:
      for (var i = 0; i < this.tablaFormadores.length; i++) {
        this.formularioControl.push(
          new UntypedFormControl({ value: "", disabled: true }),
        );
      }
      console.timeEnd("Formularios Archivos Formadores");

      //Cargar Objeto Opciones Formadores:
      this.refreshFormadores({ omitirCarga: true });

      //Cargar Objeto Opciones Instituciones:
      this.refreshInstituciones({ omitirCarga: true });

      //Incluir Metadatos Cursos en DataTable:
      console.time("Insertar Metadatos en Tabla");
      var indexBusqueda = -1;
      var indexBusquedaInstitucion = -1;
      for (var i = 0; i < this.dataTable.length; i++) {
        indexBusqueda = this.binarySearchObject(
          this.metadatosCursos[0].data,
          "cod_curso",
          this.dataTable[i]["cod_curso"],
        );
        indexBusquedaInstitucion = this.binarySearchObject(
          this.instituciones[0].data,
          "cod_institucion",
          this.dataTable[i]["institución"],
        );
        if (indexBusqueda != -1) {
          this.dataTable[i]["metadatos"] =
            this.metadatosCursos[0].data[indexBusqueda];
        }
        if (indexBusquedaInstitucion != -1) {
          this.dataTable[i]["nombreInstitucion"] =
            this.instituciones[0].data[indexBusquedaInstitucion]["institucion"];
        }
      }

      //Incluir Metadatos Formadores en DataTable:
      indexBusqueda = -1;
      for (var i = 0; i < this.tablaFormadores.length; i++) {
        indexBusqueda = this.binarySearchObject(
          this.metadatosFormadores[0].data,
          "cod__formador",
          this.tablaFormadores[i]["cod__formador"],
        );
        if (indexBusqueda != -1) {
          this.tablaFormadores[i]["metadatos"] =
            this.metadatosFormadores[0].data[indexBusqueda];
        }
      }

      //Incluir Metadatos Instituciones en DataTable:
      indexBusqueda = -1;
      for (var i = 0; i < this.tablaInstituciones.length; i++) {
        indexBusqueda = this.binarySearchObject(
          this.metadatosInstituciones[0].data,
          "cod_institucion",
          this.tablaInstituciones[i]["cod_institucion"],
        );
        if (indexBusqueda != -1) {
          this.tablaInstituciones[i]["metadatos"] =
            this.metadatosInstituciones[0].data[indexBusqueda];
        }
      }
      console.timeEnd("Insertar Metadatos en Tabla");

      //Declaración de buscador de Formadores:
      console.time("Autocompletado Formadores");
      this.filteredOptionsFormador = this.autoFormadorControl.valueChanges.pipe(
        startWith(""),
        map((value) => {
          const nombre = typeof value === "string" ? value : value["nombre"];
          return nombre
            ? this._filterFormador(nombre as string)
            : this.opcionesFormador.slice();
        }),
      );
      console.timeEnd("Autocompletado Formadores");

      //Declaración de buscador de Instituciones:
      console.time("Autocompletado Instituciones");
      this.filteredOptionsInstitucion =
        this.autoInstitucionControl.valueChanges.pipe(
          startWith(""),
          map((value) => {
            const institucion =
              typeof value === "string" ? value : value["institucion"];
            return institucion
              ? this._filterInstitucion(institucion as string)
              : this.opcionesInstituciones.slice();
          }),
        );
      console.timeEnd("Autocompletado Instituciones");

      //COMPLETAR METADATOS INTEGRADOS EN DATOS PRINCIPALES (FORMADORES):
      for (var i = 0; i < this.formadores[0].data.length; i++) {
        if (!this.formadores[0].data[i]["metadatos"]) {
          this.formadores[0].data[i]["metadatos"] = {
            formadorAdicional: false,
            flag_cambio: false,
            flag_eliminar: false,
            error: false,
          };
        }
      }

      //COMPLETAR METADATOS INTEGRADOS EN DATOS PRINCIPALES (INSTITUCIONES):
      for (var i = 0; i < this.instituciones[0].data.length; i++) {
        if (!this.instituciones[0].data[i]["metadatos"]) {
          this.instituciones[0].data[i]["metadatos"] = {
            institucionAdicional: false,
            flag_cambio: false,
            flag_eliminar: false,
            error: false,
          };
        }
      }

      //Montaje de tablas:
      console.time("Montaje de tablas");
      //this.dataTable.paginator = this.paginatorCursos;
      //this.tablaFormadores.paginator = this.paginatorFormadores;

      this.dataTable = new MatTableDataSource(this.dataTable);
      this.dataTable.paginator = this.paginatorCursos;
      this.dataTable.filterPredicate = this.filtradoCursos;

      this.tablaFormadores = new MatTableDataSource(this.tablaFormadores);
      this.tablaFormadores.paginator = this.paginatorFormadores;
      this.tablaFormadores.filterPredicate = this.filtradoFormadores;

      this.tablaInstituciones = new MatTableDataSource(this.tablaInstituciones);
      this.tablaInstituciones.paginator = this.paginatorInstituciones;
      this.tablaInstituciones.filterPredicate = this.filtradoInstituciones;

      //this.tablaCorreos = new MatTableDataSource(this.tablaCorreos)
      //this.tablaCorreos.paginator = this.paginatorCorreos;
      //this.tablaCorreos.filterPredicate = this.filtradoCorreos;

      console.timeEnd("Montaje de tablas");

      console.warn("Cursos: ", this.cursos);
      console.warn("Metadatos Cursos:", this.metadatosCursos);
      console.warn("Formadores: ", this.formadores);
      console.warn("Metadatos Formadores:", this.metadatosFormadores);
      console.warn("Formadores-Cursos: ", this.formadoresCurso);
      console.warn("Instituciones: ", this.instituciones);
      console.warn("Metadatos Instituciones: ", this.metadatosInstituciones);

      console.warn("Tabla Cursos: ", this.dataTable);
      console.warn("Tabla Formadores: ", this.tablaFormadores);
      console.warn("Tabla Instituciones: ", this.tablaInstituciones);
      //console.warn("Tabla Correos: ",this.tablaCorreos);

      resolve(dialogo);
    }); //Fin resturn Promesa
  }

  async cargarCursos(codCurso, dialogo?, tipo?) {
    return new Promise((resolve) => {
      //Comprueba si hay ruta de cursos en parametros:
      if (!this.parametros[0]) {
        //Preguntar por la ruta del archivo:
        const dialogRef = this.dialog.open(DialogoComponent, {
          disableClose: true,
          data: {
            tipoDialogo: "inputFile",
            titulo: "Ruta del archivo excel",
            contenido: "Introduzca la ruta del archivo 'Monitorización.xlsx'.",
          },
        });

        dialogRef.afterClosed().subscribe((result) => {
          //Guardar ruta curso en parametros:
          console.log("Ruta del archivo: ");
          console.log(result);

          this.parametros = {
            nombreId: "Parametros",
            rutaCursos: result[0].path,
          };

          const dialogoProcesando = this.dialog.open(DialogoComponent, {
            disableClose: true,
            data: {
              tipoDialogo: "procesando",
              titulo: "Procesando",
              contenido: "",
            },
          });

          this.appService.guardarArchivo(this.parametros, dialogoProcesando);
          this.rutaArchivoCursos = this.parametros.rutaCursos;
          this.descargarDatos();
        });
        resolve(dialogo);
      } else {
      } //Fin del check

      console.time("Descargar Datos");
      this.cursos = this.appService.getDato("/Archivos", "Cursos");
      this.metadatosCursos = this.appService.getDato(
        "/Archivos",
        "Metadatos Cursos",
      );
      this.metadatosFormadores = this.appService.getDato(
        "/Archivos",
        "Metadatos Formadores",
      );
      this.formadores = this.appService.getDato("/Archivos", "Formadores");
      this.formadoresCurso = this.appService.getDato(
        "/Archivos",
        "Formador-Curso",
      );
      console.timeEnd("Descargar Datos");

      if (this.cursos[0]) {
        this.cursos = this.cursos[0].data;
      }

      //Crea los objetos de metadatos si no existen:
      if (!this.metadatosCursos[0]) {
        this.metadatosCursos = [
          {
            data: [],
            nombreId: "Metadatos Cursos",
            objetoId: "Metadatos",
          },
        ];
      }
      if (!this.metadatosFormadores[0]) {
        this.metadatosFormadores = [
          {
            data: [],
            nombreId: "Metadatos Formadores",
            objetoId: "Metadatos Formadores",
          },
        ];
      }

      //Preprocesado del objeto de cursos:
      for (var i = 0; i < this.cursos.length; i++) {
        //Eliminar Registros Sin Codigo De Curso:
        if (this.cursos[i]["cod_curso"] === undefined) {
          this.cursos.splice(i, 1);
          i--;
        }
      }

      //Transformación de codigos a números:
      /*
        for(var i = 0; i < this.cursos.length ; i++){
            this.cursos[i]["cod_curso"] = parseInt(this.cursos[i]["cod_curso"])
        }

        for(var i = 0; i < this.instituciones[0].data.length ; i++){
            this.instituciones[0].data[i]["cod_institucion"] = parseInt(this.instituciones[0].data[i]["cod_institucion"])
        }

        for(var i = 0; i < this.metadatosInstituciones[0].data.length ; i++){
            this.metadatosInstituciones[0].data[i]["cod_institucion"] = parseInt(this.metadatosInstituciones[0].data[i]["cod_institucion"])
        }

        for(var i = 0; i < this.formadores[0].data.length ; i++){
            this.formadores[0].data[i]["cod__formador"] = parseInt(this.formadores[0].data[i]["cod__formador"])
        }

        for(var i = 0; i < this.metadatosFormadores[0].data.length ; i++){
            this.metadatosFormadores[0].data[i]["cod__formador"] = parseInt(this.metadatosFormadores[0].data[i]["cod__formador"])
        }
        */

      //Ordenado de Formadores, Cursos y Formadores-Curso:
      console.time("Ordenado Datos");
      this.cursos.sort(
        (a, b) => Number(a["cod_curso"]) - Number(b["cod_curso"]),
      );
      this.formadores[0].data.sort(
        (a, b) => Number(a["cod__formador"]) - Number(b["cod__formador"]),
      );
      this.formadoresCurso[0].data.sort(
        (a, b) => Number(a["cod_curso"]) - Number(b["cod_curso"]),
      );
      this.metadatosCursos[0].data.sort(
        (a, b) => Number(a["cod_curso"]) - Number(b["cod_curso"]),
      );
      this.metadatosFormadores[0].data.sort(
        (a, b) => Number(a["cod__formador"]) - Number(b["cod__formador"]),
      );
      this.instituciones[0].data.sort(
        (a, b) => Number(a["cod_institucion"]) - Number(b["cod_institucion"]),
      );
      this.metadatosInstituciones[0].data.sort(
        (a, b) => Number(a["cod_institucion"]) - Number(b["cod_institucion"]),
      );

      console.timeEnd("Ordenado Datos");

      console.log("Formadores:");
      console.log(this.formadores);

      console.time("Añadir Metadatos Formadores");
      //Añadir metadatos de Formadores no encontrados:
      for (var i = 0; i < this.formadores[0].data.length; i++) {
        if (
          this.binarySearchObject(
            this.metadatosFormadores[0].data,
            "cod__formador",
            this.formadores[0].data[i]["cod__formador"],
          ) == -1
        ) {
          this.metadatosFormadores[0].data.push({
            cod__formador: this.formadores[0].data[i]["cod__formador"],
            formadorAdicional: false,
            flag_cambio: false,
            flag_eliminar: false,
            error: false,
            errores: {
              nombre: false,
              email: false,
              telefono: false,
              territorial: false,
              ccaa: false,
              fecha: false,
              estado: false,
              certificado: false,
              confidencialidad: false,
              consentimiento: false,
            },
            comentarios: [],
          });
        }
      }

      console.timeEnd("Añadir Metadatos Formadores");

      console.log("Meta Formadores:");
      console.log(this.metadatosFormadores);

      console.time("Añadir Metadatos Cursos");

      //Añadir metadatos de cursos no encontrados:
      for (var i = 0; i < this.cursos.length; i++) {
        if (
          this.binarySearchObject(
            this.metadatosCursos[0].data,
            "cod_curso",
            this.cursos[i]["cod_curso"],
          ) == -1
        ) {
          this.metadatosCursos[0].data.push({
            cod_curso: this.cursos[i]["cod_curso"],
            incidenciaAdicional: false,
            revisado: false,
            flag: 0,
            flag_cambio: false,
            flag_eliminar: false,
            modificado: false,
            formadorModificado: false,
            modificaciones: {},
            error: false,
            errores: {
              postal: false,
              ccaa: false,
              cod_grupo: false,
              territorial: false,
              institucion: false,
              fecha: false,
              hora_inicio: false,
              hora_fin: false,
              curso: false,
              sesion: false,
              colectivo: false,
              grupo: false,
              alumnos: false,
              formadores: false,
            },
            comentarios: [],
          });
        }
      }

      console.timeEnd("Añadir Metadatos Cursos");

      console.time("Relacion Formadores-Cursos");

      console.warn("Formador-Curso", this.formadoresCurso);

      //Cargar Datos De Formadores en Metadatos:
      if (typeof this.formadoresCurso[0] != "undefined") {
        var metadatosFormador = [];
        var nombre = "";
        var lastIndex = 0;

        for (var i = 0; i < this.metadatosCursos[0].data.length; i++) {
          //Si ya ha sido modificado no hacer nada:
          if (this.metadatosCursos[0].data[i]["formadorModificado"]) {
            continue;
          }

          metadatosFormador = [];

          for (
            var j = lastIndex;
            j < this.formadoresCurso[0].data.length;
            j++
          ) {
            //Si coinciden los codigos de curso:
            if (
              this.formadoresCurso[0].data[j]["cod_curso"] ==
              this.metadatosCursos[0].data[i]["cod_curso"]
            ) {
              //Buscar datos de formador:
              var indexFormador = this.binarySearchObject(
                this.formadores[0].data,
                "cod__formador",
                this.formadoresCurso[0].data[j]["cod__formador"],
              );

              if (indexFormador != -1) {
                metadatosFormador.push({
                  id: this.formadoresCurso[0].data[j]["cod__formador"],
                  nombre: this.formadores[0].data[indexFormador]["nombre"],
                });
              } else {
                //Si error buscando nombre por ID:
                metadatosFormador.push({
                  id: this.formadoresCurso[0].data[j]["cod__formador"],
                  nombre: "ERROR",
                });
              }

              //Buscar datos de formador:
              /*
                        if(typeof this.formadores[0].data.find(k => k["cod__formador"] == this.formadoresCurso[0].data[j]["cod__formador"])!="undefined"){
                            metadatosFormador.push({
                                id: this.formadoresCurso[0].data[j]["cod__formador"],
                                nombre: this.formadores[0].data.find(k => k["cod__formador"]==this.formadoresCurso[0].data[j]["cod__formador"])["nombre"]
                            })

                        }else{
                            //Si error buscando nombre por ID:
                            metadatosFormador.push({
                                id: this.formadoresCurso[0].data[j]["cod__formador"],
                                nombre: "ERROR"
                            })
                        }
                        */
            } else if (
              this.formadoresCurso[0].data[j]["cod_curso"] >
              this.metadatosCursos[0].data[i]["cod_curso"]
            ) {
              lastIndex = j - 1;
              break;
            }
          }

          this.metadatosCursos[0].data[i]["formadores"] = metadatosFormador;
        }
      }

      console.timeEnd("Relacion Formadores-Cursos");

      //Filtrar Cursos:
      console.log("Cursos: ");
      console.log(this.cursos);

      console.log("Metadatos: ");
      console.log(this.metadatosCursos);

      this.dataTable = Object.assign([], this.cursos);
      this.tablaFormadores = Object.assign([], this.formadores[0].data);

      console.time("Incidencias Adicionales");

      //Cargar Cursos adicionales:
      for (var i = 0; i < this.metadatosCursos[0].data.length; i++) {
        if (
          this.metadatosCursos[0].data[i]["incidenciaAdicional"] &&
          this.dataTable.find(
            (j) => j.cod_curso == this.metadatosCursos[0].data[i]["cod_curso"],
          ) == undefined
        ) {
          this.dataTable.push({
            cod_curso: this.metadatosCursos[0].data[i]["cod_curso"],
            modalidad: this.metadatosCursos[0].data[i]["modalidad"],
            estado: this.metadatosCursos[0].data[i]["estado"],
            material: this.metadatosCursos[0].data[i]["material"],
            valoración: this.metadatosCursos[0].data[i]["valoración"],
            observaciones: this.metadatosCursos[0].data[i]["observaciones"],
            cod__postal: this.metadatosCursos[0].data[i]["cod__postal"],
            territorial: this.metadatosCursos[0].data[i]["territorial"],
            "ccaa_/_pais": this.metadatosCursos[0].data[i]["ccaa_/_pais"],
            institución: this.metadatosCursos[0].data[i]["institución"],
            fecha: this.metadatosCursos[0].data[i]["fecha"],
            hora_inicio: this.metadatosCursos[0].data[i]["hora_inicio"],
            hora_fin: this.metadatosCursos[0].data[i]["hora_fin"],
            fecha_formateada:
              this.metadatosCursos[0].data[i]["fecha_formateada"],
            hora_inicio_formateada:
              this.metadatosCursos[0].data[i]["hora_inicio_formateada"],
            hora_fin_formateada:
              this.metadatosCursos[0].data[i]["hora_fin_formateada"],
            duracion_formateada:
              this.metadatosCursos[0].data[i]["duracion_formateada"],
            cod_grupo: this.metadatosCursos[0].data[i]["cod_grupo"],
            curso: this.metadatosCursos[0].data[i]["curso"],
            sesión: this.metadatosCursos[0].data[i]["sesión"],
            colectivo: this.metadatosCursos[0].data[i]["colectivo"],
            grupo: this.metadatosCursos[0].data[i]["grupo"],
            nºasistentes: this.metadatosCursos[0].data[i]["nºasistentes"],
            metadatos: {
              incidenciaAdicional: true,
              flag_cambio: true,
              flag_eliminar: false,
              formadores: this.metadatosCursos[0].data[i]["formadores"],
              error: true,
            },
          });
          this.cursos.push({
            cod_curso: this.metadatosCursos[0].data[i]["cod_curso"],
            modalidad: this.metadatosCursos[0].data[i]["modalidad"],
            estado: this.metadatosCursos[0].data[i]["estado"],
            material: this.metadatosCursos[0].data[i]["material"],
            valoración: this.metadatosCursos[0].data[i]["valoración"],
            observaciones: this.metadatosCursos[0].data[i]["observaciones"],
            cod__postal: this.metadatosCursos[0].data[i]["cod__postal"],
            territorial: this.metadatosCursos[0].data[i]["territorial"],
            "ccaa_/_pais": this.metadatosCursos[0].data[i]["ccaa_/_pais"],
            institución: this.metadatosCursos[0].data[i]["institución"],
            fecha: this.metadatosCursos[0].data[i]["fecha"],
            hora_inicio: this.metadatosCursos[0].data[i]["hora_inicio"],
            hora_fin: this.metadatosCursos[0].data[i]["hora_fin"],
            fecha_formateada:
              this.metadatosCursos[0].data[i]["fecha_formateada"],
            hora_inicio_formateada:
              this.metadatosCursos[0].data[i]["hora_inicio_formateada"],
            hora_fin_formateada:
              this.metadatosCursos[0].data[i]["hora_fin_formateada"],
            duracion_formateada:
              this.metadatosCursos[0].data[i]["duracion_formateada"],
            cod_grupo: this.metadatosCursos[0].data[i]["cod_grupo"],
            curso: this.metadatosCursos[0].data[i]["curso"],
            sesión: this.metadatosCursos[0].data[i]["sesión"],
            colectivo: this.metadatosCursos[0].data[i]["colectivo"],
            grupo: this.metadatosCursos[0].data[i]["grupo"],
            nºasistentes: this.metadatosCursos[0].data[i]["nºasistentes"],
            metadatos: {
              incidenciaAdicional: true,
              flag_cambio: true,
              flag_eliminar: false,
              formadores: this.metadatosCursos[0].data[i]["formadores"],
              error: true,
            },
          });
          console.log(
            "Añadiendo " + this.metadatosCursos[0].data[i]["cod_curso"],
          );
        }
      }

      console.timeEnd("Incidencias Adicionales");

      console.time("Formadores Adicionales");
      //Cargar Formadores Adicionales:
      for (var i = 0; i < this.metadatosFormadores[0].data.length; i++) {
        if (
          this.metadatosFormadores[0].data[i]["formadorAdicional"] &&
          this.tablaFormadores.find(
            (j) =>
              j.cod__formador ==
              this.metadatosFormadores[0].data[i]["cod__formador"],
          ) == undefined
        ) {
          this.tablaFormadores.push({
            cod__formador: this.metadatosFormadores[0].data[i]["cod__formador"],
            nombre: this.metadatosFormadores[0].data[i]["nombre"],
            estado: this.metadatosFormadores[0].data[i]["estado"],
            fecha: this.metadatosFormadores[0].data[i]["fecha"],
            territorial: this.metadatosFormadores[0].data[i]["territorial"],
            ccaa: this.metadatosFormadores[0].data[i]["ccaa"],
            email: this.metadatosFormadores[0].data[i]["email"],
            telefono: this.metadatosFormadores[0].data[i]["telefono"],
            certificado: this.metadatosFormadores[0].data[i]["certificado"],
            consentimiento:
              this.metadatosFormadores[0].data[i]["consentimiento"],
            confidencialidad:
              this.metadatosFormadores[0].data[i]["confidencialidad"],
            metadatos: {
              formadorAdicional: true,
              flag_cambio: true,
              flag_eliminar: false,
              error: true,
            },
          });

          this.formadores[0].data.push({
            cod__formador: this.metadatosFormadores[0].data[i]["cod__formador"],
            nombre: this.metadatosFormadores[0].data[i]["nombre"],
            estado: this.metadatosFormadores[0].data[i]["estado"],
            fecha: this.metadatosFormadores[0].data[i]["fecha"],
            territorial: this.metadatosFormadores[0].data[i]["territorial"],
            ccaa: this.metadatosFormadores[0].data[i]["ccaa"],
            email: this.metadatosFormadores[0].data[i]["email"],
            telefono: this.metadatosFormadores[0].data[i]["telefono"],
            certificado: this.metadatosFormadores[0].data[i]["certificado"],
            consentimiento:
              this.metadatosFormadores[0].data[i]["consentimiento"],
            confidencialidad:
              this.metadatosFormadores[0].data[i]["confidencialidad"],
            metadatos: {
              formadorAdicional: true,
              flag_cambio: true,
              flag_eliminar: false,
              error: true,
            },
          });
          console.log(
            "Añadiendo Formador: " +
              this.metadatosFormadores[0].data[i]["cod__formador"],
          );
        }
      }
      console.timeEnd("Formadores Adicionales");

      console.time("Formatos Fechas");
      //Añadir Campos Fechas Formateadas:
      for (var i = 0; i < this.cursos.length; i++) {
        //Fecha:
        if (!this.cursos[i]["fecha_formateada"]) {
          if (typeof this.cursos[i]["fecha"] == "number") {
            this.cursos[i]["fecha_formateada"] = this.ExcelDateToJSDate(
              this.cursos[i]["fecha"],
            );
          }
        }

        if (!this.cursos[i]["hora_inicio_formateada"]) {
          //Hora Inicio:
          var horaInicio = moment({
            hour: this.cursos[i]["hora_inicio"] * 24,
            minute:
              (this.cursos[i]["hora_inicio"] * 24 -
                Math.floor(this.cursos[i]["hora_inicio"] * 24)) *
              60,
          });
          this.cursos[i]["hora_inicio_formateada"] = horaInicio.format("HH:mm");
        }
        if (!this.cursos[i]["hora_fin_formateada"]) {
          //Hora Fin:
          var horaFin = moment({
            hour: this.cursos[i]["hora_fin"] * 24,
            minute:
              (this.cursos[i]["hora_fin"] * 24 -
                Math.floor(this.cursos[i]["hora_fin"] * 24)) *
              60,
          });
          this.cursos[i]["hora_fin_formateada"] = horaFin.format("HH:mm");
        }

        if (!this.cursos[i]["duracion_formateada"]) {
          //Duración:
          var horaInicio = moment({
            hour: this.cursos[i]["hora_inicio"] * 24,
            minute:
              (this.cursos[i]["hora_inicio"] * 24 -
                Math.floor(this.cursos[i]["hora_inicio"] * 24)) *
              60,
          });
          var horaFin = moment({
            hour: this.cursos[i]["hora_fin"] * 24,
            minute:
              (this.cursos[i]["hora_fin"] * 24 -
                Math.floor(this.cursos[i]["hora_fin"] * 24)) *
              60,
          });
          var diff = horaFin.diff(horaInicio);

          //Formateo Horas:
          this.cursos[i]["duracion_formateada"] = moment
            .utc(diff)
            .format("HH:mm");
        }
      }

      console.timeEnd("Formatos Fechas");

      if (codCurso == "inicio") {
        this.dataTable = new MatTableDataSource([]);
        this.tablaFormadores = new MatTableDataSource([]);
        this.dataTable.paginator = this.paginatorCursos;
        this.tablaFormadores.paginator = this.paginatorFormadores;
        this.tablaInstituciones.paginator = this.paginatorInstituciones;
        this.tablaCorreos.paginator = this.paginatorCorreos;
        resolve(dialogo);
        console.warn("Return Inicio");
        return;
      }

      console.time("Resto Procesamiento");

      console.time("Formularios Control");
      //Incluir formularios de archivos Formadores:
      for (var i = 0; i < this.tablaFormadores.length; i++) {
        this.formularioControl.push(
          new UntypedFormControl({ value: "", disabled: true }),
        );
      }
      console.timeEnd("Formularios Control");

      console.log("FORM CONTROL:");
      console.log(this.formularioControl);

      //Incluir Metadatos Cursos en DataTable:
      var indexBusqueda = -1;
      for (var i = 0; i < this.dataTable.length; i++) {
        indexBusqueda = this.binarySearchObject(
          this.metadatosCursos[0].data,
          "cod_curso",
          this.dataTable[i]["cod_curso"],
        );
        if (indexBusqueda != -1) {
          this.dataTable[i]["metadatos"] =
            this.metadatosCursos[0].data[indexBusqueda];
        }
      }

      indexBusqueda = -1;
      //Incluir Metadatos Formadores en DataTable:
      for (var i = 0; i < this.tablaFormadores.length; i++) {
        indexBusqueda = this.binarySearchObject(
          this.metadatosFormadores[0].data,
          "cod__formador",
          this.tablaFormadores[i]["cod__formador"],
        );
        if (indexBusqueda != -1) {
          this.tablaFormadores[i]["metadatos"] =
            this.metadatosFormadores[0].data[indexBusqueda];
        }
      }

      console.timeEnd("Resto Procesamiento");

      //Aplicar Filtros:
      if (
        this.filtroModificados ||
        this.filtroWarning ||
        this.filtroError ||
        this.filtroProgramadas
      ) {
        if (this.filtroModificados) {
          for (var i = 0; i < this.dataTable.length; i++) {
            if (!this.dataTable[i]["metadatos"]["flag_cambio"]) {
              this.dataTable.splice(i, 1);
              i--;
            }
          }
        } else if (this.filtroWarning) {
          for (var i = 0; i < this.dataTable.length; i++) {
            if (!this.dataTable[i]["metadatos"]["flagTareaPendiente"]) {
              this.dataTable.splice(i, 1);
              i--;
            }
          }
        } else if (this.filtroError) {
          for (var i = 0; i < this.dataTable.length; i++) {
            if (!this.dataTable[i]["metadatos"]["error"]) {
              this.dataTable.splice(i, 1);
              i--;
            }
          }
        } else if (this.filtroProgramadas) {
          for (var i = 0; i < this.dataTable.length; i++) {
            if (
              this.dataTable[i]["estado"] != "PROGRAMADA" &&
              this.dataTable[i]["estado"] != "Programada"
            ) {
              this.dataTable.splice(i, 1);
              i--;
            }
          }
        }
      }

      //Si CodCurso Vacio:
      if (codCurso == "") {
        //this.dataTable = this.cursos;
      } else if (codCurso && tipo == "Formador") {
        this.tablaFormadores = [
          this.tablaFormadores.find((i) => i.cod__formador == codCurso),
        ];
      } else if (codCurso) {
        this.dataTable = [this.dataTable.find((i) => i.cod_curso == codCurso)];
      }

      if (this.dataTable[0] == undefined) {
        this.dataTable = [];
        this.appService.openDialog("warning", {
          titulo: "No encontrado",
          contenido: "No se han encontrado cursos con el codigo especificado.",
        });
        resolve(dialogo);
        return;
      }

      console.time("Montaje tablas");
      console.log("Cursos:");
      console.log(this.cursos);

      console.log("Metadatos:");
      console.log(this.metadatosCursos);

      console.log("Tabla Cursos:");
      this.dataTable = new MatTableDataSource(this.dataTable);
      this.dataTable.paginator = this.paginatorCursos;
      console.log(this.dataTable);

      console.log("Tabla Formadores:");
      this.tablaFormadores = new MatTableDataSource(this.tablaFormadores);
      this.tablaFormadores.paginator = this.paginatorFormadores;
      console.log(this.tablaFormadores);

      console.log("Tabla Instituciones:");
      this.tablaInstituciones = new MatTableDataSource(this.tablaInstituciones);
      this.tablaInstituciones.paginator = this.paginatorInstituciones;
      console.log(this.tablaInstituciones);

      //console.log("Tabla Correos:");
      //this.tablaCorreos = new MatTableDataSource(this.tablaCorreos)
      //this.tablaCorreos.paginator = this.paginatorInstituciones;
      //console.log(this.tablaCorreos);

      console.timeEnd("Montaje tablas");

      resolve(dialogo);
      return;
    }); //Fin resturn Promesa
  }

  /*
    async recargarCursos(){
        console.log("Cargando...");
        const dialogoProcesandoCarga = await this.dialog.open(DialogoComponent,{ disableClose: true,
              data: {tipoDialogo: "procesando", titulo: "Procesando", contenido: ""}
          });

        dialogoProcesandoCarga.afterOpened().subscribe(() => {
            this.cargarCursos("",dialogoProcesandoCarga).then((dialogo: MatDialogRef<DialogoComponent>) => {
                console.log("CERRANDO")
                dialogo.close();
            })
        })
    }
    */

  async guardarCursos(omitirMensaje?: boolean) {
    return new Promise((resolve) => {
      const dialogoProcesandoGuardado = this.dialog.open(DialogoComponent, {
        disableClose: true,
        data: {
          tipoDialogo: "procesando",
          titulo: "Procesando",
          contenido: "",
        },
      });

      //OBJETO CURSOS:
      var guardadoCursos = {};
      if (this.cursos.length > 1) {
        guardadoCursos = {
          data: this.cursos,
          nombreId: "Cursos",
          objetoId: "Cursos",
        };
      } else {
        guardadoCursos = {
          data: [],
          nombreId: "Cursos",
          objetoId: "Cursos",
        };
      }

      //OBJETO FORMADORES:
      var guardadoFormadores = {};
      if (this.formadores[0].data.length > 1) {
        guardadoFormadores = {
          data: this.formadores[0]["data"],
          nombreId: "Formadores",
          objetoId: "Formadores",
        };
      } else {
        guardadoFormadores = {
          data: [],
          nombreId: "Formadores",
          objetoId: "Formadores",
        };
      }

      //OBJETO INSTITUCIONES:
      var guardadoInstituciones = {};
      if (this.formadores[0].data.length > 1) {
        guardadoInstituciones = {
          data: this.instituciones[0]["data"],
          nombreId: "Instituciones",
          objetoId: "Instituciones",
        };
      } else {
        guardadoInstituciones = {
          data: [],
          nombreId: "Instituciones",
          objetoId: "Instituciones",
        };
      }

      //Guardar METADATOS CURSOS:
      console.log("Guardando METADATOS CURSOS: ");
      console.log(this.metadatosCursos);
      this.appService.guardarArchivo(this.metadatosCursos[0]).then((result) => {
        if (!result) {
          console.log("Error guardando archivo");
          dialogoProcesandoGuardado.close("error");
          this.appService.openDialog("error", {
            titulo: "Error",
            contenido: "Error guardando los archivos de cursos",
          });
          resolve(false);
        }

        //GUARDANDO CURSOS:
        console.log("Guardando CURSOS: ");
        console.log(guardadoCursos);
        this.appService.guardarArchivo(guardadoCursos).then((result) => {
          if (!result) {
            console.log("Error guardando archivo");
            dialogoProcesandoGuardado.close("error");
            this.appService.openDialog("error", {
              titulo: "Error",
              contenido: "Error guardando los archivos de cursos",
            });
            resolve(false);
          }

          //GUARDANDO METADATOS FORMADOR:
          console.log("Guardando METADATOS FORMADORES...");
          console.log(this.metadatosFormadores);
          this.appService
            .guardarArchivo(this.metadatosFormadores[0])
            .then((result) => {
              if (!result) {
                console.log("Error guardando archivo");
                dialogoProcesandoGuardado.close("error");
                this.appService.openDialog("error", {
                  titulo: "Error",
                  contenido: "Error guardando los archivos de cursos",
                });
                resolve(false);
              }

              console.log("Guardando FORMADOR...");
              console.log(guardadoFormadores);
              this.appService
                .guardarArchivo(guardadoFormadores)
                .then((result) => {
                  if (!result) {
                    console.log("Error guardando archivo");
                    dialogoProcesandoGuardado.close("error");
                    this.appService.openDialog("error", {
                      titulo: "Error",
                      contenido: "Error guardando los archivos de cursos",
                    });
                    resolve(false);
                  }

                  console.log("Guardando METADATOS INSTITUCION...");
                  console.log(this.metadatosInstituciones);
                  this.appService
                    .guardarArchivo(this.metadatosInstituciones[0])
                    .then((result) => {
                      if (!result) {
                        console.log("Error guardando archivo");
                        dialogoProcesandoGuardado.close("error");
                        this.appService.openDialog("error", {
                          titulo: "Error",
                          contenido: "Error guardando los archivos de cursos",
                        });
                        resolve(false);
                      }

                      console.log("Guardando INSTITUCION...");
                      console.log(guardadoInstituciones);
                      this.appService
                        .guardarArchivo(guardadoInstituciones)
                        .then((result) => {
                          if (!result) {
                            console.log("Error guardando archivo");
                            dialogoProcesandoGuardado.close("error");
                            this.appService.openDialog("error", {
                              titulo: "Error",
                              contenido:
                                "Error guardando los archivos de cursos",
                            });
                            resolve(false);
                          }

                          console.log("GUARDADO CON EXITO");
                          dialogoProcesandoGuardado.close(true);
                          resolve(true);
                        }); //FIN CALLBACK INSTITUCIONES.
                    }); //FIN CALLBACK METADATOS INSTITUCIONES.
                }); //FIN CALLBACK FORMADORES.
            }); //FIN CALLBACK METADATOS FORMADORES.
        }); //FIN CALLBACK CURSOS
      }); //FIN CALLBACK METADATOS CURSOS.

      //Guardar Metadatos Formador:
      //console.log("Guardando METADATOS FORMADOR: ");
      //console.log(this.metadatosFormadores);
      //this.appService.guardarArchivo(this.metadatosFormadores[0])

      //Guardar Formador:
      //this.appService.guardarArchivo(this.metadatosFormadores[0], dialogoProcesando)
    }); //Fin return Promesa
  }

  filtradoCursos(data: any, filter: any): boolean {
    //console.warn("filtrado",data,filter)
    if (!filter.filtroMaestro) {
      return true;
    }

    if (filter.filtroModificado) {
      if (!data["metadatos"]["flag_cambio"]) {
        return false;
      }
    }

    if (filter.filtroWarning) {
      if (!data["metadatos"]["flagTareaPendiente"]) {
        return false;
      }
    }

    if (filter.filtroError) {
      if (!data["metadatos"]["error"]) {
        return false;
      }
    }

    if (filter.filtroProgramada) {
      if (data["estado"] != "PROGRAMADA" && data["estado"] != "Programada") {
        return false;
      }
    }

    if (filter.filtroCodigoCurso) {
      if (data["cod_curso"] != filter.filtroCodigoCurso) {
        return false;
      }
    }

    if (filter.filtroFecha) {
      var fecha = moment(data["fecha_formateada"]);
      if (
        (fecha.isBefore(filter.filtroFecha["fechaFin"]) &&
          fecha.isAfter(filter.filtroFecha["fechaInicio"])) ||
        fecha.isSame(filter.filtroFecha["fechaInicio"]) ||
        fecha.isSame(filter.filtroFecha["fechaFin"])
      ) {
      } else {
        return false;
      }
    }

    //False --> No se muestra
    //True --> Se muestra
    return true;
  }

  filtradoFormadores(data: any, filter: any): boolean {
    if (!filter.filtroMaestro) {
      return true;
    }
    //console.warn("filtrado",data,filter)
    if (filter.filtroCodigoFormador) {
      if (data["cod__formador"] != filter.filtroCodigoFormador) {
        return false;
      }
    }
    if (filter.filtroModificado) {
      if (!data["metadatos"]["flag_cambio"]) {
        return false;
      }
    }
    if (filter.filtroError) {
      if (!data["metadatos"]["error"]) {
        return false;
      }
    }
    return true;
  }

  filtradoInstituciones(data: any, filter: any): boolean {
    if (!filter.filtroMaestro) {
      return true;
    }
    if (filter.filtroCodigoInstitucion) {
      if (data["cod_institucion"] != filter.filtroCodigoInstitucion) {
        return false;
      }
    }
    if (filter.filtroModificado) {
      if (!data["metadatos"]["flag_cambio"]) {
        return false;
      }
    }
    if (filter.filtroError) {
      if (!data["metadatos"]["error"]) {
        return false;
      }
    }
    return true;
  }

  filtradoCorreos(data: any, filter: any): boolean {
    if (data["estado"] == "COMPLETADO") {
      return false;
    }
    return true;
  }

  resetFiltro() {
    console.warn(this.buscador);
    switch (this.pantallaSeleccionada) {
      case "Cursos":
        this.filtroCursos = {
          filtroMaestro: false,
          filtroGeneral: null,
          filtroError: false,
          filtroWarning: false,
          filtroProgramada: false,
          filtroModificado: false,
          filtroFecha: null,
          filtroCodigoCurso: null,
        };
        this.buscador.nativeElement.value = null;
        this.dataTable.filter = Object.assign(this.filtroCursos, {});
        this.filterButtonControl.setValue("");
        break;
      case "Formadores":
        this.filtroFormadores = {
          filtroMaestro: false,
          filtroGeneral: null,
          filtroError: false,
          filtroWarning: false,
          filtroModificado: false,
          filtroCodigoFormador: null,
        };
        this.autoFormadorControl.setValue({
          nombre: "",
          id: "",
        });
        this.tablaFormadores.filter = Object.assign(this.filtroFormadores, {});
        //this.filterButtonControl.setValue("")
        break;
      case "Instituciones":
        this.filtroInstituciones = {
          filtroMaestro: false,
          filtroGeneral: null,
          filtroError: false,
          filtroWarning: false,
          filtroModificado: false,
          filtroCodigoInstitucion: null,
        };
        this.autoInstitucionControl.setValue({
          institucion: "",
          id: "",
        });
        this.tablaInstituciones.filter = Object.assign(
          this.filtroInstituciones,
          {},
        );
        break;
    }
  }

  filtrar(filtro?: string, valor?: any) {
    //this.filtroModificados = false;
    //this.filtroWarning = false;
    //this.filtroError = false;
    //this.filtroProgramadas = false;

    var filtroProvisional:
      | FiltroCursos
      | FiltroFormadores
      | FiltroInstituciones;
    if (this.pantallaSeleccionada == "Cursos") {
      filtroProvisional = Object.assign(this.filtroCursos, {});
    } else if (this.pantallaSeleccionada == "Formadores") {
      filtroProvisional = Object.assign(this.filtroFormadores, {});
    } else if (this.pantallaSeleccionada == "Instituciones") {
      filtroProvisional = Object.assign(this.filtroInstituciones, {});
    }

    switch (filtro) {
      case "Todas":
        this.filtroModificados = false;
        this.filtroWarning = false;
        this.filtroError = false;
        this.filtroProgramadas = false;
        this.guardarCursos(true).then(() => {
          console.log("Realizando Carga");
          this.cargarCursos("", dialogoProcesandoCarga).then(
            (dialogo: MatDialogRef<DialogoComponent>) => {
              //dialogo.close();
            },
          );
        });
        return;
        break;

      case "Modificado":
        filtroProvisional["filtroModificado"] =
          !filtroProvisional["filtroModificado"];
        break;
      case "Warning":
        filtroProvisional["filtroWarning"] =
          !filtroProvisional["filtroWarning"];
        break;
      case "Error":
        filtroProvisional["filtroError"] = !filtroProvisional["filtroError"];
        break;
      case "Programadas":
        //this.filtroProgramadas = !this.filtroProgramadas;
        filtroProvisional["filtroProgramada"] =
          !filtroProvisional["filtroProgramada"];
        break;
      case "Codigo Curso":
        filtroProvisional["filtroCodigoCurso"] = valor;
        break;
      case "Codigo Formador":
        filtroProvisional["filtroCodigoFormador"] = valor;
        break;
      case "Codigo Institucion":
        filtroProvisional["filtroCodigoInstitucion"] = valor;
        break;
      case "Fecha Cursos":
        filtroProvisional["filtroFecha"] = valor;
        break;
    }

    //Check de filtro Maestro:

    switch (this.pantallaSeleccionada) {
      case "Cursos":
        if (
          filtroProvisional["filtroProgramada"] == false &&
          filtroProvisional["filtroModificado"] == false &&
          filtroProvisional["filtroWarning"] == false &&
          filtroProvisional["filtroError"] == false &&
          filtroProvisional["filtroGeneral"] == null &&
          filtroProvisional["filtroFecha"] == null &&
          filtroProvisional["filtroCodigoCurso"] == null
        ) {
          filtroProvisional["filtroMaestro"] = false;
        } else {
          filtroProvisional["filtroMaestro"] = true;
        }
        this.dataTable.filter = filtroProvisional;
        break;
      case "Formadores":
        if (
          filtroProvisional["filtroModificado"] == false &&
          filtroProvisional["filtroWarning"] == false &&
          filtroProvisional["filtroError"] == false &&
          filtroProvisional["filtroGeneral"] == null &&
          filtroProvisional["filtroCodigoFormador"] == null
        ) {
          filtroProvisional["filtroMaestro"] = false;
        } else {
          filtroProvisional["filtroMaestro"] = true;
        }
        this.tablaFormadores.filter = filtroProvisional;
        break;
      case "Instituciones":
        if (
          filtroProvisional["filtroModificado"] == false &&
          filtroProvisional["filtroWarning"] == false &&
          filtroProvisional["filtroError"] == false &&
          filtroProvisional["filtroGeneral"] == null &&
          filtroProvisional["filtroCodigoInstitucion"] == null
        ) {
          filtroProvisional["filtroMaestro"] = false;
        } else {
          filtroProvisional["filtroMaestro"] = true;
        }
        this.tablaInstituciones.filter = filtroProvisional;
        break;
    }

    //Aplicación del filtro:
    return;

    console.log("Buscando Cursos...");
    const dialogoProcesandoCarga = this.dialog.open(DialogoComponent, {
      disableClose: true,
      data: { tipoDialogo: "procesando", titulo: "Procesando", contenido: "" },
    });

    dialogoProcesandoCarga.afterOpened().subscribe(() => {
      //Aplicar Filtros:
      if (
        this.filtroModificados ||
        this.filtroWarning ||
        this.filtroError ||
        this.filtroProgramadas
      ) {
        if (this.filtroModificados) {
          for (var i = 0; i < this.dataTable.data.length; i++) {
            if (!this.dataTable.data[i]["metadatos"]["flag_cambio"]) {
              this.dataTable.data.splice(i, 1);
              i--;
            }
          }
        } else if (this.filtroWarning) {
          for (var i = 0; i < this.dataTable.data.length; i++) {
            if (!this.dataTable.data[i]["metadatos"]["flagTareaPendiente"]) {
              this.dataTable.splice(i, 1);
              i--;
            }
          }
        } else if (this.filtroError) {
          for (var i = 0; i < this.dataTable.data.length; i++) {
            if (!this.dataTable.data[i]["metadatos"]["error"]) {
              this.dataTable.data.splice(i, 1);
              i--;
            }
          }
        } else if (this.filtroProgramadas) {
          for (var i = 0; i < this.dataTable.data.length; i++) {
            if (
              this.dataTable.data[i]["estado"] != "PROGRAMADA" &&
              this.dataTable.data[i]["estado"] != "Programada"
            ) {
              this.dataTable.data.splice(i, 1);
              i--;
            }
          }
        }
      }

      //athis.changeDetectorRefs.detectChanges();
      this.dataTable._updateChangeSubscription();
      console.log(this.dataTable);
      dialogoProcesandoCarga.close();
    });
    /*
        this.guardarCursos(true).then(() =>{
            console.log("Realizando Carga");
            this.cargarCursos("",dialogoProcesandoCarga).then((dialogo: MatDialogRef<DialogoComponent>) => {
                dialogo.close();
            })
        });
        */
  }

  filtrarRMCA() {
    this.filtroRMCA = true;
    this.filtroSAP = false;
    this.cargarCursos("");
  }

  filtrarSAP() {
    this.filtroSAP = true;
    this.filtroRMCA = false;
    this.cargarCursos("");
  }

  filtrarGeneral(event: Event) {
    this.filtroRMCA = false;
    this.filtroSAP = false;
    this.cargarCursos("");
  }

  resetREV(event: Event) {
    for (var i = 0; i < this.cursos.length; i++) {
      this.cursos[i]["metadatos"]["revisado"] = false;
    }
    this.cargarCursos("");
  }

  applyFilter(event: Event) {
    const filterValue = (event.target as HTMLInputElement).value;
    this.dataTable.filter = filterValue.trim().toLowerCase();
  }

  addCampoInput(event: Event, element, campo: string) {
    //Asignación del valor:
    if (
      typeof event["source"] !== "undefined" &&
      event["source"]["controlType"] == "mat-select"
    ) {
      var valor = event["value"];
    } else {
      var valor = (event.target as HTMLInputElement).value;
    }

    //Actuación para FORMADORES:
    if (this.pantallaSeleccionada == "Formadores") {
      this.metadatosFormadores[0].data.find(
        (i) => i.cod__formador == element.cod__formador,
      )["flag_cambio"] = true;

      this.formadores[0].data.find(
        (i) => i.cod__formador == element.cod__formador,
      )[campo] = valor;
      this.tablaFormadores.data.find(
        (i) => i.cod__formador == element.cod__formador,
      )[campo] = valor;

      console.log("FORMADORES: ");
      console.log(this.formadores);
      this.comprobarFormador(element.cod__formador);
    } else if (this.pantallaSeleccionada == "Instituciones") {
      this.metadatosInstituciones[0].data.find(
        (i) => i.cod_institucion == element.cod_institucion,
      )["flag_cambio"] = true;

      this.instituciones[0].data.find(
        (i) => i.cod_institucion == element.cod_institucion,
      )[campo] = valor;
      this.tablaInstituciones.data.find(
        (i) => i.cod_institucion == element.cod_institucion,
      )[campo] = valor;

      //Autorelleno de COD_POSTAL:
      if (campo == "cod__postal") {
        console.log("Codigo Postal: " + valor);
        if (typeof valor == "number") {
          valor = valor.toString();
        }
        if (valor.length == 5) {
          for (var i = 0; i < this.codigoProvincia[0].data.length; i++) {
            if (
              Number(this.codigoProvincia[0].data[i]["cod__provincia"]) ==
              Number(valor.slice(0, 2))
            ) {
              console.log(
                "Asignando Provincia: " +
                  this.codigoProvincia[0].data[i]["provincia"],
              );

              //Actualiza Territorial:
              console.log(
                "Asignando Territorial: " +
                  this.codigoProvincia[0].data[i]["territorial"],
              );
              this.metadatosInstituciones[0].data.find(
                (j) => j.cod_institucion == element.cod_institucion,
              )["flag_cambio"] = true;
              this.metadatosInstituciones[0].data.find(
                (j) => j.cod_institucion == element.cod_institucion,
              )["territorial"] = this.codigoProvincia[0].data[i]["territorial"];

              console.log(this.instituciones);
              console.log(this.metadatosInstituciones);

              this.instituciones[0].data.find(
                (j) => j.cod_institucion == element.cod_institucion,
              )["territorial"] = this.codigoProvincia[0].data[i]["territorial"];
              this.tablaInstituciones.data.find(
                (j) => j.cod_institucion == element.cod_institucion,
              )["territorial"] = this.codigoProvincia[0].data[i]["territorial"];

              //Actualiza CCAA:
              console.log(
                "Asignando CCAA: " + this.codigoProvincia[0].data[i]["ccaa"],
              );
              this.metadatosInstituciones[0].data.find(
                (j) => j.cod_institucion == element.cod_institucion,
              )["flag_cambio"] = true;
              this.metadatosInstituciones[0].data.find(
                (j) => j.cod_institucion == element.cod_institucion,
              )["ccaa/_pais"] = this.codigoProvincia[0].data[i]["ccaa"];
              this.instituciones[0].data.find(
                (j) => j.cod_institucion == element.cod_institucion,
              )["ccaa_/_pais"] = this.codigoProvincia[0].data[i]["ccaa"];
              this.tablaInstituciones.data.find(
                (j) => j.cod_institucion == element.cod_institucion,
              )["ccaa_/_pais"] = this.codigoProvincia[0].data[i]["ccaa"];
            }
          }
        }
      }

      this.comprobarInstitucion(element.cod_institucion);

      console.log("INSTITUCIONES: ");
      console.log(this.instituciones);

      //Actuación para CURSOS:
    } else if (this.pantallaSeleccionada == "Cursos") {
      switch (campo) {
        case "fecha_formateada":
          console.log(event);
          console.log(element);
          var fechaExcel = this.JSDateToExcelDate(
            moment(event["value"]).toDate(),
          );
          console.log("Fecha formateada: ");
          console.log(fechaExcel);

          this.metadatosCursos[0].data.find(
            (i) => i.cod_curso == element.cod_curso,
          )["flag_cambio"] = true;
          this.metadatosCursos[0].data.find(
            (i) => i.cod_curso == element.cod_curso,
          )["fecha_formateada"] = moment(event["value"]).toDate();
          this.metadatosCursos[0].data.find(
            (i) => i.cod_curso == element.cod_curso,
          )["fecha"] = fechaExcel;
          this.dataTable.data.find((i) => i.cod_curso == element.cod_curso)[
            "fecha"
          ] = fechaExcel;
          this.cursos.find((i) => i.cod_curso == element.cod_curso)["fecha"] =
            fechaExcel;
          console.log(
            "FECHA ACTUALIZADA: " +
              this.JSDateToExcelDate(
                moment(
                  this.cursos.find((i) => i.cod_curso == element.cod_curso)[
                    "metadatos"
                  ]["fecha_formateada"],
                ).toDate(),
              ),
          );
          break;

        case "hora_inicio_formateada":
          console.log("Hora formateada: ");
          console.log(this.JSDateToExcelDate(moment(valor, "HH:mm").toDate()));

          var offset = Math.floor(this.JSDateToExcelDate(moment().toDate()));

          console.log("Offset: " + offset);

          this.metadatosCursos[0].data.find(
            (i) => i.cod_curso == element.cod_curso,
          )["flag_cambio"] = true;

          this.metadatosCursos[0].data.find(
            (i) => i.cod_curso == element.cod_curso,
          )["hora_inicio"] =
            this.JSDateToExcelDate(moment(valor, "HH:mm").toDate()) - offset;
          this.dataTable.data.find((i) => i.cod_curso == element.cod_curso)[
            "hora_inicio"
          ] = this.JSDateToExcelDate(moment(valor, "HH:mm").toDate()) - offset;
          this.cursos.find((i) => i.cod_curso == element.cod_curso)[
            "hora_inicio"
          ] = this.JSDateToExcelDate(moment(valor, "HH:mm").toDate()) - offset;

          //Hora Inicio:
          var horaSinOffset =
            this.JSDateToExcelDate(moment(valor, "HH:mm").toDate()) - offset;
          var horaInicio = moment({
            hour: horaSinOffset * 24,
            minute: (horaSinOffset * 24 - Math.floor(horaSinOffset * 24)) * 60,
          });

          this.metadatosCursos[0].data.find(
            (i) => i.cod_curso == element.cod_curso,
          )["hora_inicio_formateada"] = horaInicio.format("HH:mm");

          break;

        case "hora_fin_formateada":
          var horaFormateada = this.JSDateToExcelDate(
            moment(valor, "HH:mm").toDate(),
          );

          console.log("Hora formateada: ");
          console.log(horaFormateada);

          var offset = Math.floor(this.JSDateToExcelDate(moment().toDate()));

          console.log("Offset: " + offset);

          this.metadatosCursos[0].data.find(
            (i) => i.cod_curso == element.cod_curso,
          )["flag_cambio"] = true;

          this.metadatosCursos[0].data.find(
            (i) => i.cod_curso == element.cod_curso,
          )["hora_fin_formateada"] = horaFormateada;
          this.metadatosCursos[0].data.find(
            (i) => i.cod_curso == element.cod_curso,
          )["hora_fin"] =
            this.JSDateToExcelDate(moment(valor, "HH:mm").toDate()) - offset;
          this.dataTable.data.find((i) => i.cod_curso == element.cod_curso)[
            "hora_fin"
          ] = this.JSDateToExcelDate(moment(valor, "HH:mm").toDate()) - offset;
          this.cursos.find((i) => i.cod_curso == element.cod_curso)[
            "hora_fin"
          ] = this.JSDateToExcelDate(moment(valor, "HH:mm").toDate()) - offset;

          //Hora Inicio:
          var horaSinOffset =
            this.JSDateToExcelDate(moment(valor, "HH:mm").toDate()) - offset;
          var horaFin = moment({
            hour: horaSinOffset * 24,
            minute: (horaSinOffset * 24 - Math.floor(horaSinOffset * 24)) * 60,
          });

          this.metadatosCursos[0].data.find(
            (i) => i.cod_curso == element.cod_curso,
          )["hora_fin_formateada"] = horaFin.format("HH:mm");

          break;

        default:
          this.metadatosCursos[0].data.find(
            (i) => i.cod_curso == element.cod_curso,
          )[campo] = valor;
          break;
      }

      console.log("Valor");
      console.log(valor);

      //Cambiar Flag de modificado:
      this.metadatosCursos[0].data.find(
        (i) => i.cod_curso == element.cod_curso,
      )["flag_cambio"] = true;

      this.cursos.find((i) => i.cod_curso == element.cod_curso)[campo] = valor;
      this.dataTable.data.find((i) => i.cod_curso == element.cod_curso)[campo] =
        valor;

      //Autorelleno de DURACIÓN:
      if (campo == "hora_inicio_formateada" || campo == "hora_fin_formateada") {
        if (element["hora_fin_formateada"] && element["hora_fin_formateada"]) {
          var end = moment(element["hora_fin_formateada"], "HH:mm");
          var startTime = moment(element["hora_inicio_formateada"], "HH:mm");
          var duration = moment.duration(end.diff(startTime));
          var hours = duration.asHours();
          console.log("Inicio: ");
          console.log(startTime);
          console.log(element["hora_inicio_formateada"]);
          console.log("Fin: ");
          console.log(end);
          console.log("Duracion: ");
          console.log(duration);
          if (duration.isValid()) {
            element["duracion_formateada"] = moment
              .utc(duration.asMilliseconds())
              .format("HH:mm");
            console.log("Formateado: " + element["duracion_formateada"]);
            this.metadatosCursos[0].data.find(
              (i) => i.cod_curso == element.cod_curso,
            )["flag_cambio"] = true;
            this.metadatosCursos[0].data.find(
              (i) => i.cod_curso == element.cod_curso,
            )["duracion_formateada"] = element["duracion_formateada"];
            this.cursos.find((i) => i.cod_curso == element.cod_curso)[
              "duracion_formateada"
            ] = element["duracion_formateada"];
            this.dataTable.data.find((i) => i.cod_curso == element.cod_curso)[
              "duracion_formateada"
            ] = element["duracion_formateada"];
          }
        }
      }

      //Autorelleno de COD_POSTAL:
      if (campo == "cod__postal") {
        console.log("Codigo Postal: " + valor);
        if (typeof valor == "number") {
          valor = valor.toString();
        }
        if (valor.length == 5) {
          for (var i = 0; i < this.codigoProvincia[0].data.length; i++) {
            if (
              Number(this.codigoProvincia[0].data[i]["cod__provincia"]) ==
              Number(valor.slice(0, 2))
            ) {
              console.log(
                "Asignando Provincia: " +
                  this.codigoProvincia[0].data[i]["provincia"],
              );

              //Actualiza Territorial:
              console.log(
                "Asignando Territorial: " +
                  this.codigoProvincia[0].data[i]["territorial"],
              );
              this.metadatosCursos[0].data.find(
                (j) => j.cod_curso == element.cod_curso,
              )["flag_cambio"] = true;
              this.metadatosCursos[0].data.find(
                (j) => j.cod_curso == element.cod_curso,
              )["territorial"] = this.codigoProvincia[0].data[i]["territorial"];
              this.cursos.find((j) => j.cod_curso == element.cod_curso)[
                "territorial"
              ] = this.codigoProvincia[0].data[i]["territorial"];
              this.dataTable.data.find((j) => j.cod_curso == element.cod_curso)[
                "territorial"
              ] = this.codigoProvincia[0].data[i]["territorial"];

              //Actualiza CCAA:
              console.log(
                "Asignando CCAA: " + this.codigoProvincia[0].data[i]["ccaa"],
              );
              this.metadatosCursos[0].data.find(
                (j) => j.cod_curso == element.cod_curso,
              )["flag_cambio"] = true;
              this.metadatosCursos[0].data.find(
                (j) => j.cod_curso == element.cod_curso,
              )["ccaa_/_pais"] = this.codigoProvincia[0].data[i]["ccaa"];
              this.cursos.find((j) => j.cod_curso == element.cod_curso)[
                "ccaa_/_pais"
              ] = this.codigoProvincia[0].data[i]["ccaa"];
              this.dataTable.data.find((j) => j.cod_curso == element.cod_curso)[
                "ccaa_/_pais"
              ] = this.codigoProvincia[0].data[i]["ccaa"];
            }
          }
        }
      }

      this.comprobarDatos(element.cod_curso);
    } //Fin modificación campo Curso

    console.log("Cursos: ");
    console.log(this.cursos);
    console.log("Metadatos: ");
    console.log(this.metadatosCursos);
    console.log("Tabla: ");
    console.log(this.dataTable);
  }

  cambioHora() {
    console.log("CAMBIO FECHA");
  }

  formatearFecha(fechaSinFormatear: number) {
    console.log("Formateando Fecha");
    console.log(fechaSinFormatear);
    return;
  }

  ExcelDateToJSDate(excelDate) {
    return new Date(Math.round((excelDate - 25569) * 86400 * 1000));
  }

  JSDateToExcelDate(fecha) {
    let date = new Date(fecha);
    var converted =
      25569.0 +
      (date.getTime() - date.getTimezoneOffset() * 60 * 1000) /
        (1000 * 60 * 60 * 24);
    return converted;
  }

  marcarRevisado(element) {
    this.metadatosCursos[0].data.find(
      (i) => i.incidencia == element.cod_curso,
    ).revisado = true;
    this.dataTable.data.find(
      (i) => i.cod_curso == element.cod_curso,
    ).metadatos.revisado = true;
    console.log(this.metadatosCursos);
  }

  mostrarComentarios(codigoCurso) {
    console.log("Mostrando comentarios de " + codigoCurso);
    var comentarios = this.metadatosCursos[0].data.find(
      (i) => i.cod_curso == codigoCurso,
    ).comentarios;

    console.log(comentarios);

    const dialogRef = this.dialog.open(DialogoComponent, {
      data: {
        tipoDialogo: "comentarioInc",
        titulo: codigoCurso,
        contenido: comentarios,
      },
    });

    dialogRef.afterClosed().subscribe((result) => {
      console.log("Fin del dialogo comentario");
      console.log(result);

      //Asignar Metadatos de incidencia:
      this.metadatosCursos[0].data.find(
        (i) => i.cod_curso == result.titulo,
      ).comentarios = result.contenido;

      console.log("META: ");
      console.log(this.metadatosCursos);

      //Check de tareas pendientes:
      var checkPendiente = false;
      for (var i = 0; i < this.metadatosCursos[0].data.length; i++) {
        for (
          var j = 0;
          j < this.metadatosCursos[0].data[i].comentarios.length;
          j++
        ) {
          if (this.metadatosCursos[0].data[i].comentarios[j].completo) {
            this.metadatosCursos[0].data[i]["flagTareaPendiente"] = true;
            checkPendiente = true;
          }
        }
        if (!checkPendiente) {
          this.metadatosCursos[0].data[i]["flagTareaPendiente"] = false;
        }
      }

      console.log(result);
    });
  }

  addCurso() {
    console.log("Añadiendo Incidencia");

    //Buscar un codigo sin asignar:
    var codigoBase = moment().format("YYYYMMDD");
    var codigoNuevo = codigoBase;

    var codigoSincronizacion = this.datosProyecto.idSincronizacion;
    if (this.datosProyecto.idSincronizacion < 10) {
      codigoSincronizacion = "0" + codigoSincronizacion;
    } else {
      codigoSincronizacion = String(codigoSincronizacion);
    }

    console.error("Codigo Sincronizacion: " + codigoSincronizacion);
    //Iterar entre los codigos disponibles:
    var indice = "00";
    for (var i = 0; i < 100; i++) {
      if (i < 10) {
        codigoNuevo = codigoBase + "0" + String(i) + codigoSincronizacion;
      } else {
        codigoNuevo = codigoBase + String(i) + codigoSincronizacion;
      }

      console.log(this.metadatosCursos);
      if (
        !this.metadatosCursos[0].data.find((j) => j.cod_curso == codigoNuevo)
      ) {
        break;
      }
    }

    console.log("Codigo Nuevo: " + codigoNuevo);
    this.addIncidenciaAdicional(codigoNuevo);
  }

  addFormador() {
    console.log("Añadiendo Formador:");
    //Obtener codigo nuevo Codigo de formador:
    var codigoFormador = 0;

    var codigoSincronizacion = this.datosProyecto.idSincronizacion;
    if (this.datosProyecto.idSincronizacion < 10) {
      codigoSincronizacion = "0" + codigoSincronizacion;
    } else {
      codigoSincronizacion = String(codigoSincronizacion);
    }

    for (var i = 0; i < this.formadores[0].data.length; i++) {
      if (
        Number(this.formadores[0].data[i]["cod__formador"]) > codigoFormador
      ) {
        codigoFormador = Number(this.formadores[0].data[i]["cod__formador"]);
      }
    }

    if (codigoFormador > 9999) {
      codigoFormador = Number(codigoFormador.toString().slice(0, -2));
    }

    codigoFormador++;

    //Añade Codigo Sincronizacion:
    codigoFormador = Number(
      String(codigoFormador).concat(codigoSincronizacion),
    );

    console.log("Nuevo Formador: " + codigoFormador);
    this.addFormadorAdicional(codigoFormador);

    return;
  }

  eliminarFormador(codigoFormador) {
    //Eliminando Formador:
    var indexFormador = this.formadores[0].data.findIndex(
      (i) => i.cod__formador == codigoFormador,
    );
    var indexMetadatosFormador = this.metadatosFormadores[0].data.findIndex(
      (i) => i.cod__formador == codigoFormador,
    );

    const dialogoProcesandoCarga = this.dialog.open(DialogoComponent, {
      disableClose: true,
      data: { tipoDialogo: "procesando", titulo: "Procesando", contenido: "" },
    });

    if (indexFormador == -1) {
      dialogoProcesandoCarga.close();
      return;
    }

    //Si el formador es adicional:
    if (
      this.formadores[0].data[indexFormador]["metadatos"]["formadorAdicional"]
    ) {
      console.log("Eliminando Formador Adicional");
      console.log("Index Formador: " + indexFormador);
      console.log("Index Metadatos Formador: " + indexMetadatosFormador);
      this.formadores[0].data.splice(indexFormador, 1);
      this.metadatosFormadores[0].data.splice(indexMetadatosFormador, 1);

      console.log(this.formadores);
      console.log(this.metadatosFormadores);

      this.guardarCursos(true).then(() => {
        console.log("Realizando Carga");
        this.inicializarDatos().then(() => {
          this.resetFiltro();
          dialogoProcesandoCarga.close();
        });
      });
    } else {
      //No es Formador Adicional:
      this.metadatosFormadores[0].data[indexMetadatosFormador][
        "flag_eliminar"
      ] = true;
      this.metadatosFormadores[0].data[indexMetadatosFormador][
        "flagTareaPendiente"
      ] = true;
      this.metadatosFormadores[0].data[indexMetadatosFormador][
        "comentarios"
      ].push({
        fecha: new Date(Date.now()),
        completo: false,
        comentario: "Marcado para eliminación",
      });
      this.guardarCursos();
      dialogoProcesandoCarga.close();
      this.appService.openDialog("warning", {
        titulo: "Marcado para eliminación",
        contenido:
          "El curso ha sido marcado para su eliminación y será eliminado de la base de datos en la proxima subida.",
      });
    }
  }

  addInstitucion() {
    console.log("Añadiendo Institucion:");
    //Obtener codigo nuevo Codigo de formador:
    var codigoInstitucion = 0;

    var codigoSincronizacion = this.datosProyecto.idSincronizacion;
    if (this.datosProyecto.idSincronizacion < 10) {
      codigoSincronizacion = "0" + codigoSincronizacion;
    } else {
      codigoSincronizacion = String(codigoSincronizacion);
    }

    for (var i = 0; i < this.instituciones[0].data.length; i++) {
      if (
        Number(this.instituciones[0].data[i]["cod_institucion"]) >
        codigoInstitucion
      ) {
        codigoInstitucion = Number(
          this.instituciones[0].data[i]["cod_institucion"],
        );
      }
    }

    if (codigoInstitucion > 9999) {
      codigoInstitucion = Number(codigoInstitucion.toString().slice(0, -2));
    }

    codigoInstitucion++;

    //Añade Codigo Sincronizacion:
    codigoInstitucion = Number(
      String(codigoInstitucion).concat(codigoSincronizacion),
    );

    console.log("Nuevo Formador: " + codigoInstitucion);
    this.addInstitucionAdicional(codigoInstitucion);
    return;
  }

  eliminarInstitucion(codigoInstitucion) {
    //Eliminando Institucion:
    var indexInstitucion = this.instituciones[0].data.findIndex(
      (i) => i.cod_institucion == codigoInstitucion,
    );
    var indexMetadatosInstitucion =
      this.metadatosInstituciones[0].data.findIndex(
        (i) => i.cod_institucion == codigoInstitucion,
      );

    const dialogoProcesandoCarga = this.dialog.open(DialogoComponent, {
      disableClose: true,
      data: { tipoDialogo: "procesando", titulo: "Procesando", contenido: "" },
    });

    if (indexInstitucion == -1) {
      dialogoProcesandoCarga.close();
      return;
    }

    //Si la institucion es adicional:
    if (
      this.instituciones[0].data[indexInstitucion]["metadatos"][
        "institucionAdicional"
      ]
    ) {
      console.log("Eliminando Institucion Adicional");
      console.log("Index Institucion: " + indexInstitucion);
      console.log("Index Metadatos Institucion: " + indexMetadatosInstitucion);
      this.instituciones[0].data.splice(indexInstitucion, 1);
      this.metadatosInstituciones[0].data.splice(indexMetadatosInstitucion, 1);

      console.log(this.instituciones);
      console.log(this.metadatosInstituciones);

      this.guardarCursos(true).then(() => {
        console.log("Realizando Carga");
        this.inicializarDatos().then(() => {
          this.resetFiltro();
          dialogoProcesandoCarga.close();
        });
      });
    } else {
      //No es Insititución Adicional:
      this.metadatosInstituciones[0].data[indexMetadatosInstitucion][
        "flag_eliminar"
      ] = true;
      this.metadatosInstituciones[0].data[indexMetadatosInstitucion][
        "flagTareaPendiente"
      ] = true;
      this.metadatosInstituciones[0].data[indexMetadatosInstitucion][
        "comentarios"
      ].push({
        fecha: new Date(Date.now()),
        completo: false,
        comentario: "Marcado para eliminación",
      });
      this.guardarCursos();
      dialogoProcesandoCarga.close();
      this.appService.openDialog("warning", {
        titulo: "Marcado para eliminación",
        contenido:
          "El curso ha sido marcado para su eliminación y será eliminado de la base de datos en la proxima subida.",
      });
    }

    return;
  }

  mostrarAddIncidencia() {
    console.log("Panel add Incidencia");

    const dialogRef = this.dialog.open(DialogoComponent, {
      data: {
        tipoDialogo: "inputText",
        titulo: "Añadir Incidencia",
        contenido: "Introduce el código de incidencia: ",
      },
    });

    dialogRef.afterClosed().subscribe((result) => {
      console.log("Fin del dialogo comentario");
      //Asignar Metadatos de incidencia:
      this.addIncidenciaAdicional(result);
      console.log(result);
    });
  }

  addFormadorAdicional(codigoFormador) {
    //Formateo del código
    var nuevoCodigo = codigoFormador.toLocaleString("es-ES", {
      minimumIntegerDigits: 5,
      useGrouping: false,
    });

    //Añadiendo incidencia a metadatos:
    this.metadatosFormadores[0].data.push({
      cod__formador: nuevoCodigo,
      formadorAdicional: true,
      flag_cambio: true,
      error: true,
      errores: {
        nombre: true,
        email: false,
        telefono: false,
        postal: false,
        territorial: false,
        ccaa: false,
        fecha: false,
        estado: false,
        certificado: false,
        confidencialidad: false,
        consentimiento: false,
      },
      comentarios: [],
    });

    var objetoCodigo = {
      option: {
        value: {
          id: nuevoCodigo,
        },
      },
    };

    const dialogoProcesandoCarga = this.dialog.open(DialogoComponent, {
      disableClose: true,
      data: { tipoDialogo: "procesando", titulo: "Procesando", contenido: "" },
    });
    this.guardarCursos(true).then(() => {
      console.log("Realizando Carga");
      this.inicializarDatos().then(() => {
        this.buscarFormador(objetoCodigo);
        dialogoProcesandoCarga.close();
      });
    });
  }

  addInstitucionAdicional(codigoInstitucion) {
    //Formateo del código
    var nuevoCodigo = codigoInstitucion.toLocaleString("es-ES", {
      minimumIntegerDigits: 5,
      useGrouping: false,
    });

    //Añadiendo incidencia a metadatos:
    this.metadatosInstituciones[0].data.push({
      cod_institucion: nuevoCodigo,
      institucionAdicional: true,
      flag_cambio: true,
      error: true,
      errores: {
        institucion: true,
        contacto: false,
        email: false,
        telefono: false,
        postal: false,
        territorial: false,
        ccaa: false,
        fecha: false,
        estado: false,
      },
      comentarios: [],
    });

    var objetoCodigo = {
      option: {
        value: {
          id: nuevoCodigo,
        },
      },
    };

    const dialogoProcesandoCarga = this.dialog.open(DialogoComponent, {
      disableClose: true,
      data: { tipoDialogo: "procesando", titulo: "Procesando", contenido: "" },
    });
    this.guardarCursos(true).then(() => {
      console.log("Realizando Carga");
      this.inicializarDatos().then(() => {
        this.buscarInstitucion(objetoCodigo);
        dialogoProcesandoCarga.close();
      });
    });
  }

  addIncidenciaAdicional(codigoCurso) {
    //Añadiendo incidencia a metadatos:
    this.metadatosCursos[0].data.push({
      cod_curso: codigoCurso,
      cod_grupo: codigoCurso,
      modalidad: "Presencial",
      estado: "PROGRAMADA",
      material: "SI",
      curso: "AF",
      incidenciaAdicional: true,
      servicio: "",
      equipo: "",
      log: [],
      revisado: false,
      tipo_cambio: "Cola",
      ultima_revision: 44833,
      wa: false,
      week_resolution: "-",
      flag: 0,
      flag_cambio: true,
      error: true,
      errores: {
        postal: true,
        ccaa: true,
        cod_grupo: true,
        territorial: true,
        institucion: true,
        fecha: true,
        hora_inicio: true,
        hora_fin: true,
        curso: false,
        sesion: true,
        colectivo: true,
        grupo: true,
        alumnos: true,
        formadores: true,
      },
      comentarios: [],
      descripcion_ejecutiva: "Descripción ejecutiva",
      en_crq: false,
      responsable: "",
    });

    const dialogoProcesandoCarga = this.dialog.open(DialogoComponent, {
      disableClose: true,
      data: { tipoDialogo: "procesando", titulo: "Procesando", contenido: "" },
    });

    console.time("nuevoCurso");
    dialogoProcesandoCarga.afterOpened().subscribe((result) => {
      //CREACION ANTIGUA DE CURSO
      /*
            this.guardarCursos(true).then(() =>{
                console.log("Realizando Carga");
                this.inicializarDatos().then(() => {
                    this.buscarCodigo(codigoCurso);
                    dialogoProcesandoCarga.close();
                    console.timeEnd("nuevoCurso");
                })
            });
            */

      this.inicializarDatos(dialogoProcesandoCarga, {
        omitirRecarga: true,
      }).then((dialogo: MatDialogRef<DialogoComponent>) => {
        this.buscarCodigo(codigoCurso);
        dialogo.close();
        console.timeEnd("nuevoCurso");
      });
    });

    /*
        this.guardarCursos(true).then(() =>{
            console.log("Realizando Carga");
            this.inicializarDatos().then(() => {
                this.buscarCodigo(codigoCurso);
                dialogoProcesandoCarga.close();
                console.timeEnd("nuevoCurso");
            })
        });
        */
  }

  eliminarCurso(codigo) {
    const dialogoProcesandoEliminarCurso = this.dialog.open(DialogoComponent, {
      disableClose: true,
      data: { tipoDialogo: "procesando", titulo: "Procesando", contenido: "" },
    });

    dialogoProcesandoEliminarCurso.afterOpened().subscribe(() => {
      //Determinar si es incidencia Adicional:
      var indexMetadatos = this.metadatosCursos[0].data.findIndex(
        (i) => i.cod_curso == codigo,
      );
      var indexCursos = this.cursos.findIndex((i) => i.cod_curso == codigo);

      if (this.metadatosCursos[0].data[indexMetadatos]["incidenciaAdicional"]) {
        this.metadatosCursos[0].data.splice(indexMetadatos, 1);
        this.cursos.splice(indexCursos, 1);

        this.guardarCursos(true).then(() => {
          console.log("Realizando Carga");
          this.inicializarDatos().then(() => {
            dialogoProcesandoEliminarCurso.close();
            this.appService.openDialog("exito", {
              titulo: "Curso eliminado con exito",
              contenido:
                "El curso se ha eliminado satisfactoriamente del gestor.",
            });
          });
        });
      } else {
        //No es curso Adicional:
        this.metadatosCursos[0].data[indexMetadatos]["flag_eliminar"] = true;
        this.metadatosCursos[0].data[indexMetadatos]["flagTareaPendiente"] =
          true;
        this.metadatosCursos[0].data[indexMetadatos]["comentarios"].push({
          fecha: new Date(Date.now()),
          completo: false,
          comentario: "Marcado para eliminación",
        });
        this.guardarCursos();
        dialogoProcesandoEliminarCurso.close();
        this.appService.openDialog("warning", {
          titulo: "Marcado para eliminación",
          contenido:
            "El curso ha sido marcado para su eliminación y será eliminado de la base de datos en la proxima subida.",
        });
      }
    });
  }

  async buscarCodigo(busqueda) {
    console.log("Buscando Codigo");
    const dialogoProcesandoCarga = this.dialog.open(DialogoComponent, {
      disableClose: true,
      data: { tipoDialogo: "procesando", titulo: "Procesando", contenido: "" },
    });

    this.filtrar("Codigo Curso", busqueda);

    dialogoProcesandoCarga.close();

    /*
        this.guardarCursos(true).then(() =>{
            console.log("Realizando Carga");
            this.cargarCursos(busqueda,dialogoProcesandoCarga).then((dialogo: MatDialogRef<DialogoComponent>) => {
                if(busqueda!= "" && busqueda!="inicio"){
                    this.comprobarDatos(busqueda);
                }
                dialogo.close();
            })
        });
        */
  }

  async buscarFormador(event) {
    console.log("Buscando Formador");
    console.log(event);
    var busqueda = event["option"]["value"]["id"];

    const dialogoProcesandoCarga = this.dialog.open(DialogoComponent, {
      disableClose: true,
      data: { tipoDialogo: "procesando", titulo: "Procesando", contenido: "" },
    });

    console.warn("Buscando Formador: " + busqueda);

    this.filtrar("Codigo Formador", busqueda);
    dialogoProcesandoCarga.close();

    /*
        this.guardarCursos(true).then(() =>{
            console.log("Realizando Busqueda...");

            this.cargarCursos(busqueda,dialogoProcesandoCarga,"Formador").then((dialogo: MatDialogRef<DialogoComponent>) => {
                if(busqueda!= "" && busqueda!="inicio"){
                    this.comprobarDatos(busqueda,"Formador");
                }
                dialogo.close();
            })
        });
        */
  }

  async buscarInstitucion(event) {
    console.log("Buscando Institución");
    console.log(event);
    var busqueda = event["option"]["value"]["id"];

    const dialogoProcesandoCarga = this.dialog.open(DialogoComponent, {
      disableClose: true,
      data: { tipoDialogo: "procesando", titulo: "Procesando", contenido: "" },
    });

    console.warn("Buscando Institucion: " + busqueda);

    this.filtrar("Codigo Institucion", busqueda);
    dialogoProcesandoCarga.close();

    /*
        this.guardarCursos(true).then(() =>{
            console.log("Realizando Busqueda...");

            this.cargarCursos(busqueda,dialogoProcesandoCarga,"Formador").then((dialogo: MatDialogRef<DialogoComponent>) => {
                if(busqueda!= "" && busqueda!="inicio"){
                    this.comprobarDatos(busqueda,"Formador");
                }
                dialogo.close();
            })
        });
        */
  }

  comprobarFormador(codigoFormador) {
    console.log("Comprobando Formador: " + codigoFormador);

    if (
      this.formadores[0].data.find((i) => i.cod__formador == codigoFormador) ==
      undefined
    ) {
      console.log("Error comprobando datos: Código de Formador no encontrado");
      return;
    }

    this.metadatosFormadores[0].data.find(
      (i) => i.cod__formador == codigoFormador,
    )["error"] = false;
    this.tablaFormadores.data.find((i) => i.cod__formador == codigoFormador)[
      "metadatos"
    ]["error"] = false;

    //Comprobando NOMBRE:
    if (
      !this.formadores[0].data.find((i) => i.cod__formador == codigoFormador)[
        "nombre"
      ]
    ) {
      console.log("ERROR NOMBRE: " + codigoFormador);
      this.tablaFormadores.data.find((i) => i.cod__formador == codigoFormador)[
        "metadatos"
      ]["errores"]["nombre"] = true;
      this.metadatosFormadores[0].data.find(
        (i) => i.cod__formador == codigoFormador,
      )["errores"]["nombre"] = true;
      this.metadatosFormadores[0].data.find(
        (i) => i.cod__formador == codigoFormador,
      )["error"] = true;
    } else {
      this.tablaFormadores.data.find((i) => i.cod__formador == codigoFormador)[
        "metadatos"
      ]["errores"]["nombre"] = false;
      this.metadatosFormadores[0].data.find(
        (i) => i.cod__formador == codigoFormador,
      )["errores"]["nombre"] = false;
    }

    //Comprobando ESTADO:
    if (
      !this.formadores[0].data.find((i) => i.cod__formador == codigoFormador)[
        "estado"
      ]
    ) {
      console.log("ERROR ESTADO: " + codigoFormador);
      this.tablaFormadores.data.find((i) => i.cod__formador == codigoFormador)[
        "metadatos"
      ]["errores"]["estado"] = true;
      this.metadatosFormadores[0].data.find(
        (i) => i.cod__formador == codigoFormador,
      )["errores"]["estado"] = true;
      this.metadatosFormadores[0].data.find(
        (i) => i.cod__formador == codigoFormador,
      )["error"] = true;
    } else {
      this.tablaFormadores.data.find((i) => i.cod__formador == codigoFormador)[
        "metadatos"
      ]["errores"]["estado"] = false;
      this.metadatosFormadores[0].data.find(
        (i) => i.cod__formador == codigoFormador,
      )["errores"]["estado"] = false;
    }

    return;
  }

  comprobarInstitucion(codigoInstitucion) {
    if (
      this.instituciones[0].data.find(
        (i) => i.cod_institucion == codigoInstitucion,
      ) == undefined
    ) {
      console.log("Error comprobando datos: Código de Formador no encontrado");
      return;
    }

    this.metadatosInstituciones[0].data.find(
      (i) => i.cod_institucion == codigoInstitucion,
    )["error"] = false;
    this.tablaInstituciones.data.find(
      (i) => i.cod_institucion == codigoInstitucion,
    )["metadatos"]["error"] = false;

    //Comprobando NOMBRE INSTITUCION:
    if (
      !this.instituciones[0].data.find(
        (i) => i.cod_institucion == codigoInstitucion,
      )["institucion"]
    ) {
      this.tablaInstituciones.data.find(
        (i) => i.cod_institucion == codigoInstitucion,
      )["metadatos"]["errores"]["institucion"] = true;
      this.metadatosInstituciones[0].data.find(
        (i) => i.cod_institucion == codigoInstitucion,
      )["errores"]["institucion"] = true;
      this.metadatosInstituciones[0].data.find(
        (i) => i.cod_institucion == codigoInstitucion,
      )["error"] = true;
      console.error("ERROR INSTITUCION: " + codigoInstitucion);
    } else {
      this.tablaInstituciones.data.find(
        (i) => i.cod_institucion == codigoInstitucion,
      )["metadatos"]["errores"]["institucion"] = false;
      this.metadatosInstituciones[0].data.find(
        (i) => i.cod_institucion == codigoInstitucion,
      )["errores"]["institucion"] = false;
    }

    //Comprobando COD POSTAL:
    if (
      !this.instituciones[0].data.find(
        (i) => i.cod_institucion == codigoInstitucion,
      )["cod__postal"]
    ) {
      this.tablaInstituciones.data.find(
        (i) => i.cod_institucion == codigoInstitucion,
      )["metadatos"]["errores"]["postal"] = true;
      this.metadatosInstituciones[0].data.find(
        (i) => i.cod_institucion == codigoInstitucion,
      )["errores"]["postal"] = true;
      this.metadatosInstituciones[0].data.find(
        (i) => i.cod_institucion == codigoInstitucion,
      )["error"] = true;
      console.error("ERROR POSTAL: " + codigoInstitucion);
    } else {
      this.tablaInstituciones.data.find(
        (i) => i.cod_institucion == codigoInstitucion,
      )["metadatos"]["errores"]["postal"] = false;
      this.metadatosInstituciones[0].data.find(
        (i) => i.cod_institucion == codigoInstitucion,
      )["errores"]["postal"] = false;
    }

    //Comprobando CONTACTO 1:
    if (
      !this.instituciones[0].data.find(
        (i) => i.cod_institucion == codigoInstitucion,
      )["contacto1"]
    ) {
      this.tablaInstituciones.data.find(
        (i) => i.cod_institucion == codigoInstitucion,
      )["metadatos"]["errores"]["contacto"] = true;
      this.metadatosInstituciones[0].data.find(
        (i) => i.cod_institucion == codigoInstitucion,
      )["errores"]["contacto"] = true;
      this.metadatosInstituciones[0].data.find(
        (i) => i.cod_institucion == codigoInstitucion,
      )["error"] = true;
      console.error("ERROR CONTACTO 1: " + codigoInstitucion);
    } else {
      this.tablaInstituciones.data.find(
        (i) => i.cod_institucion == codigoInstitucion,
      )["metadatos"]["errores"]["contacto"] = false;
      this.metadatosInstituciones[0].data.find(
        (i) => i.cod_institucion == codigoInstitucion,
      )["errores"]["contacto"] = false;
    }

    //Comprobando EMAIL 1:
    if (
      !this.instituciones[0].data.find(
        (i) => i.cod_institucion == codigoInstitucion,
      )["email1"]
    ) {
      this.tablaInstituciones.data.find(
        (i) => i.cod_institucion == codigoInstitucion,
      )["metadatos"]["errores"]["email"] = true;
      this.metadatosInstituciones[0].data.find(
        (i) => i.cod_institucion == codigoInstitucion,
      )["errores"]["email"] = true;
      this.metadatosInstituciones[0].data.find(
        (i) => i.cod_institucion == codigoInstitucion,
      )["error"] = true;
      console.error("ERROR EMAIL: " + codigoInstitucion);
    } else {
      this.tablaInstituciones.data.find(
        (i) => i.cod_institucion == codigoInstitucion,
      )["metadatos"]["errores"]["email"] = false;
      this.metadatosInstituciones[0].data.find(
        (i) => i.cod_institucion == codigoInstitucion,
      )["errores"]["email"] = false;
    }
  }

  comprobarDatos(codigoCurso: string, tipo?: string) {
    if (tipo == "Formador") {
      this.comprobarFormador(codigoCurso);
      return;
    }

    if (this.cursos.find((i) => i.cod_curso == codigoCurso) == undefined) {
      console.log("Error comprobando datos: Código no encontrado");
      return;
    }

    this.metadatosCursos[0].data.find((i) => i.cod_curso == codigoCurso)[
      "error"
    ] = false;
    this.dataTable.data.find((i) => i.cod_curso == codigoCurso)["metadatos"][
      "error"
    ] = false;

    //Comprobando COD. POSTAL:
    var codigoPostalCursos = this.cursos.find(
      (i) => i.cod_curso == codigoCurso,
    )["cod__postal"];
    if (typeof codigoPostalCursos == "number") {
      codigoPostalCursos = codigoPostalCursos.toString();
    }

    if (!codigoPostalCursos || codigoPostalCursos.length < 5) {
      this.dataTable.data.find((i) => i.cod_curso == codigoCurso)["metadatos"][
        "errores"
      ]["postal"] = true;
      this.metadatosCursos[0].data.find((i) => i.cod_curso == codigoCurso)[
        "errores"
      ]["postal"] = true;
      this.metadatosCursos[0].data.find((i) => i.cod_curso == codigoCurso)[
        "error"
      ] = true;
    } else {
      this.dataTable.data.find((i) => i.cod_curso == codigoCurso)["metadatos"][
        "errores"
      ]["postal"] = false;
      this.metadatosCursos[0].data.find((i) => i.cod_curso == codigoCurso)[
        "errores"
      ]["postal"] = false;
    }

    //Comprobando CONCORDANCIA LOCALIDAD:
    var valor = this.cursos.find((i) => i.cod_curso == codigoCurso)[
      "cod__postal"
    ];

    if (typeof valor == "number") {
      valor = valor.toString();
    }

    if (typeof valor == "string") {
      for (var i = 0; i < this.codigoProvincia[0].data.length; i++) {
        if (
          Number(this.codigoProvincia[0].data[i]["cod__provincia"]) ==
          Number(valor.slice(0, 2))
        ) {
          //Comprobando CCAA:
          if (
            this.cursos.find((i) => i.cod_curso == codigoCurso)[
              "ccaa_/_pais"
            ] != this.codigoProvincia[0].data[i]["ccaa"]
          ) {
            this.dataTable.data.find((i) => i.cod_curso == codigoCurso)[
              "metadatos"
            ]["errores"]["ccaa"] = true;
            this.metadatosCursos[0].data.find(
              (i) => i.cod_curso == codigoCurso,
            )["errores"]["ccaa"] = true;
            this.metadatosCursos[0].data.find(
              (i) => i.cod_curso == codigoCurso,
            )["error"] = true;
          } else {
            this.dataTable.data.find((i) => i.cod_curso == codigoCurso)[
              "metadatos"
            ]["errores"]["ccaa"] = false;
            this.metadatosCursos[0].data.find(
              (i) => i.cod_curso == codigoCurso,
            )["errores"]["ccaa"] = false;
          }

          //Comprobando TERRITORIAL:
          if (
            this.cursos.find((i) => i.cod_curso == codigoCurso)[
              "territorial"
            ] != this.codigoProvincia[0].data[i]["territorial"]
          ) {
            this.dataTable.data.find((i) => i.cod_curso == codigoCurso)[
              "metadatos"
            ]["errores"]["territorial"] = true;
            this.metadatosCursos[0].data.find(
              (i) => i.cod_curso == codigoCurso,
            )["errores"]["territorial"] = true;
            this.metadatosCursos[0].data.find(
              (i) => i.cod_curso == codigoCurso,
            )["error"] = true;
          } else {
            this.dataTable.data.find((i) => i.cod_curso == codigoCurso)[
              "metadatos"
            ]["errores"]["territorial"] = false;
            this.metadatosCursos[0].data.find(
              (i) => i.cod_curso == codigoCurso,
            )["errores"]["territorial"] = false;
          }
        }
      } //FIN FOR
    } else {
      this.dataTable.data.find((i) => i.cod_curso == codigoCurso)["metadatos"][
        "errores"
      ]["postal"] = true;
      this.metadatosCursos[0].data.find((i) => i.cod_curso == codigoCurso)[
        "errores"
      ]["postal"] = true;
      this.metadatosCursos[0].data.find((i) => i.cod_curso == codigoCurso)[
        "error"
      ] = true;
    }

    if (valor == "") {
      this.dataTable.data.find((i) => i.cod_curso == codigoCurso)["metadatos"][
        "errores"
      ]["postal"] = true;
      this.metadatosCursos[0].data.find((i) => i.cod_curso == codigoCurso)[
        "errores"
      ]["postal"] = true;
      this.metadatosCursos[0].data.find((i) => i.cod_curso == codigoCurso)[
        "error"
      ] = true;
    }

    //Comprobando INSTITUCION:
    console.warn(
      "Comprobando institución: ",
      this.cursos.find((i) => i.cod_curso == codigoCurso),
    );
    if (!this.cursos.find((i) => i.cod_curso == codigoCurso)["institución"]) {
      this.dataTable.data.find((i) => i.cod_curso == codigoCurso)["metadatos"][
        "errores"
      ]["institucion"] = true;
      this.metadatosCursos[0].data.find((i) => i.cod_curso == codigoCurso)[
        "errores"
      ]["institucion"] = true;
      this.metadatosCursos[0].data.find((i) => i.cod_curso == codigoCurso)[
        "error"
      ] = true;
    } else {
      this.dataTable.data.find((i) => i.cod_curso == codigoCurso)["metadatos"][
        "errores"
      ]["institucion"] = false;
      this.metadatosCursos[0].data.find((i) => i.cod_curso == codigoCurso)[
        "errores"
      ]["institucion"] = false;
    }

    //Comprobando FECHA:
    var fecha = moment(
      this.cursos.find((i) => i.cod_curso == codigoCurso)["fecha_formateada"],
    );
    if (
      !fecha.isValid() ||
      typeof this.cursos.find((i) => i.cod_curso == codigoCurso)[
        "fecha_formateada"
      ] == "undefined"
    ) {
      this.dataTable.data.find((i) => i.cod_curso == codigoCurso)["metadatos"][
        "errores"
      ]["fecha"] = true;
      this.metadatosCursos[0].data.find((i) => i.cod_curso == codigoCurso)[
        "errores"
      ]["fecha"] = true;
      this.metadatosCursos[0].data.find((i) => i.cod_curso == codigoCurso)[
        "error"
      ] = true;
    } else {
      this.dataTable.data.find((i) => i.cod_curso == codigoCurso)["metadatos"][
        "errores"
      ]["fecha"] = false;
      this.metadatosCursos[0].data.find((i) => i.cod_curso == codigoCurso)[
        "errores"
      ]["fecha"] = false;
    }

    //Comprobando HORA INICIO:
    var horaInicio = moment(
      this.cursos.find((i) => i.cod_curso == codigoCurso)[
        "hora_inicio_formateada"
      ],
      "HH:mm",
    );
    if (!horaInicio.isValid()) {
      this.dataTable.data.find((i) => i.cod_curso == codigoCurso)["metadatos"][
        "errores"
      ]["hora_inicio"] = true;
      this.metadatosCursos[0].data.find((i) => i.cod_curso == codigoCurso)[
        "errores"
      ]["hora_inicio"] = true;
      this.metadatosCursos[0].data.find((i) => i.cod_curso == codigoCurso)[
        "error"
      ] = true;
    } else {
      this.dataTable.data.find((i) => i.cod_curso == codigoCurso)["metadatos"][
        "errores"
      ]["hora_inicio"] = false;
      this.metadatosCursos[0].data.find((i) => i.cod_curso == codigoCurso)[
        "errores"
      ]["hora_inicio"] = false;
    }

    //Comprobando HORA FIN:
    var horaFin = moment(
      this.cursos.find((i) => i.cod_curso == codigoCurso)[
        "hora_fin_formateada"
      ],
      "HH:mm",
    );
    if (!horaFin.isValid()) {
      this.dataTable.data.find((i) => i.cod_curso == codigoCurso)["metadatos"][
        "errores"
      ]["hora_fin"] = true;
      this.metadatosCursos[0].data.find((i) => i.cod_curso == codigoCurso)[
        "errores"
      ]["hora_fin"] = true;
      this.metadatosCursos[0].data.find((i) => i.cod_curso == codigoCurso)[
        "error"
      ] = true;
    } else {
      this.dataTable.data.find((i) => i.cod_curso == codigoCurso)["metadatos"][
        "errores"
      ]["hora_fin"] = false;
      this.metadatosCursos[0].data.find((i) => i.cod_curso == codigoCurso)[
        "errores"
      ]["hora_fin"] = false;
    }

    //Comprobando HORA INICIO < HORA FIN:
    if (horaFin.isBefore(horaInicio)) {
      this.dataTable.data.find((i) => i.cod_curso == codigoCurso)["metadatos"][
        "errores"
      ]["hora_inicio"] = true;
      this.metadatosCursos[0].data.find((i) => i.cod_curso == codigoCurso)[
        "errores"
      ]["hora_inicio"] = true;
      this.metadatosCursos[0].data.find((i) => i.cod_curso == codigoCurso)[
        "error"
      ] = true;
    } else {
      this.dataTable.data.find((i) => i.cod_curso == codigoCurso)["metadatos"][
        "errores"
      ]["hora_inicio"] = false;
      this.metadatosCursos[0].data.find((i) => i.cod_curso == codigoCurso)[
        "errores"
      ]["hora_inicio"] = false;
    }

    //Comprobando Duracion:
    var duracion = moment(
      this.cursos.find((i) => i.cod_curso == codigoCurso)[
        "duracion_formateada"
      ],
      "HH:mm",
    );
    if (!duracion.isValid()) {
      this.dataTable.data.find((i) => i.cod_curso == codigoCurso)["metadatos"][
        "errores"
      ]["duracion"] = true;
      this.metadatosCursos[0].data.find((i) => i.cod_curso == codigoCurso)[
        "errores"
      ]["duracion"] = true;
      this.metadatosCursos[0].data.find((i) => i.cod_curso == codigoCurso)[
        "error"
      ] = true;
    } else {
      this.dataTable.data.find((i) => i.cod_curso == codigoCurso)["metadatos"][
        "errores"
      ]["duracion"] = false;
      this.metadatosCursos[0].data.find((i) => i.cod_curso == codigoCurso)[
        "errores"
      ]["duracion"] = false;
    }

    //Comprobando COD GRUPO:
    if (!this.cursos.find((i) => i.cod_curso == codigoCurso)["cod_grupo"]) {
      this.dataTable.data.find((i) => i.cod_curso == codigoCurso)["metadatos"][
        "errores"
      ]["cod_grupo"] = true;
      this.metadatosCursos[0].data.find((i) => i.cod_curso == codigoCurso)[
        "errores"
      ]["cod_grupo"] = true;
      this.metadatosCursos[0].data.find((i) => i.cod_curso == codigoCurso)[
        "error"
      ] = true;
    } else {
      this.dataTable.data.find((i) => i.cod_curso == codigoCurso)["metadatos"][
        "errores"
      ]["cod_grupo"] = false;
      this.metadatosCursos[0].data.find((i) => i.cod_curso == codigoCurso)[
        "errores"
      ]["cod_grupo"] = false;
    }

    //Comprobando CURSO:
    /*
        if(!this.cursos.find(i => i.cod_curso==codigoCurso)["curso"]){
            this.dataTable.data.find(i => i.cod_curso==codigoCurso)["metadatos"]["errores"]["curso"]= true
            this.metadatosCursos[0].data.find(i => i.cod_curso==codigoCurso)["errores"]["curso"]= true
            this.metadatosCursos[0].data.find(i => i.cod_curso==codigoCurso)["error"]= true
        }else{
            this.dataTable.data.find(i => i.cod_curso==codigoCurso)["metadatos"]["errores"]["curso"]= false
            this.metadatosCursos[0].data.find(i => i.cod_curso==codigoCurso)["errores"]["curso"]= false
        }
        */

    //Comprobando SESION:
    if (!this.cursos.find((i) => i.cod_curso == codigoCurso)["sesión"]) {
      this.dataTable.data.find((i) => i.cod_curso == codigoCurso)["metadatos"][
        "errores"
      ]["sesion"] = true;
      this.metadatosCursos[0].data.find((i) => i.cod_curso == codigoCurso)[
        "errores"
      ]["sesion"] = true;
      this.metadatosCursos[0].data.find((i) => i.cod_curso == codigoCurso)[
        "error"
      ] = true;
    } else {
      this.dataTable.data.find((i) => i.cod_curso == codigoCurso)["metadatos"][
        "errores"
      ]["sesion"] = false;
      this.metadatosCursos[0].data.find((i) => i.cod_curso == codigoCurso)[
        "errores"
      ]["sesion"] = false;
    }

    //Comprobando COLECTIVO:
    if (!this.cursos.find((i) => i.cod_curso == codigoCurso)["colectivo"]) {
      this.dataTable.data.find((i) => i.cod_curso == codigoCurso)["metadatos"][
        "errores"
      ]["colectivo"] = true;
      this.metadatosCursos[0].data.find((i) => i.cod_curso == codigoCurso)[
        "errores"
      ]["colectivo"] = true;
      this.metadatosCursos[0].data.find((i) => i.cod_curso == codigoCurso)[
        "error"
      ] = true;
    } else {
      this.dataTable.data.find((i) => i.cod_curso == codigoCurso)["metadatos"][
        "errores"
      ]["colectivo"] = false;
      this.metadatosCursos[0].data.find((i) => i.cod_curso == codigoCurso)[
        "errores"
      ]["colectivo"] = false;
    }

    //Comprobando GRUPO:
    if (!this.cursos.find((i) => i.cod_curso == codigoCurso)["grupo"]) {
      this.dataTable.data.find((i) => i.cod_curso == codigoCurso)["metadatos"][
        "errores"
      ]["grupo"] = true;
      this.metadatosCursos[0].data.find((i) => i.cod_curso == codigoCurso)[
        "errores"
      ]["grupo"] = true;
      this.metadatosCursos[0].data.find((i) => i.cod_curso == codigoCurso)[
        "error"
      ] = true;
    } else {
      this.dataTable.data.find((i) => i.cod_curso == codigoCurso)["metadatos"][
        "errores"
      ]["grupo"] = false;
      this.metadatosCursos[0].data.find((i) => i.cod_curso == codigoCurso)[
        "errores"
      ]["grupo"] = false;
    }

    //Comprobando ASISTENTES:
    if (
      Number(
        this.cursos.find((i) => i.cod_curso == codigoCurso)["nºasistentes"],
      ) <= 0 ||
      typeof this.cursos.find((i) => i.cod_curso == codigoCurso)[
        "nºasistentes"
      ] == "undefined"
    ) {
      this.dataTable.data.find((i) => i.cod_curso == codigoCurso)["metadatos"][
        "errores"
      ]["alumnos"] = true;
      this.metadatosCursos[0].data.find((i) => i.cod_curso == codigoCurso)[
        "errores"
      ]["alumnos"] = true;
      this.metadatosCursos[0].data.find((i) => i.cod_curso == codigoCurso)[
        "error"
      ] = true;
    } else {
      this.dataTable.data.find((i) => i.cod_curso == codigoCurso)["metadatos"][
        "errores"
      ]["alumnos"] = false;
      this.metadatosCursos[0].data.find((i) => i.cod_curso == codigoCurso)[
        "errores"
      ]["alumnos"] = false;
    }

    //Comprobando FORMADOR:
    if (
      typeof this.cursos.find((i) => i.cod_curso == codigoCurso)["metadatos"][
        "formadores"
      ] == "undefined" ||
      Number(
        this.cursos.find((i) => i.cod_curso == codigoCurso)["metadatos"][
          "formadores"
        ].length,
      ) <= 0
    ) {
      this.dataTable.data.find((i) => i.cod_curso == codigoCurso)["metadatos"][
        "errores"
      ]["formadores"] = true;
      this.metadatosCursos[0].data.find((i) => i.cod_curso == codigoCurso)[
        "errores"
      ]["formadores"] = true;
      this.metadatosCursos[0].data.find((i) => i.cod_curso == codigoCurso)[
        "error"
      ] = true;
    } else {
      this.dataTable.data.find((i) => i.cod_curso == codigoCurso)["metadatos"][
        "errores"
      ]["formadores"] = false;
      this.metadatosCursos[0].data.find((i) => i.cod_curso == codigoCurso)[
        "errores"
      ]["formadores"] = false;
    }
  } //Fin Comprobaciones

  reformatearFechas() {
    console.log("Formateando...");
    //ITERAR POR CURSOS:
    for (var i = 0; i < this.cursos.length; i++) {
      if (
        this.cursos[i]["fecha"] &&
        moment(this.cursos[i]["fecha"]).isValid()
      ) {
        this.cursos[i]["fecha"] = this.JSDateToExcelDate(
          moment(this.cursos[i]["fecha"]).toDate(),
        );
      }
    }
  } //FIN REFORMATEO

  revertirFechas() {
    console.log("Revirtiendo Fechas...");
    for (var i = 0; i < this.cursos.length; i++) {
      //Fecha:
      if (
        this.dataTable.data[i] &&
        typeof this.dataTable.data[i]["fecha"] == "number"
      ) {
        this.dataTable.data[i]["fecha"] = this.ExcelDateToJSDate(
          this.cursos[i]["fecha"],
        );
        this.cursos[i]["fecha"] = this.ExcelDateToJSDate(
          this.cursos[i]["fecha"],
        );
      }
    }
  }

  async subirCursos() {
    //this.appService.openDialog("error",{titulo: "Error",contenido: "Error: No se ha podido conectar con los servicios One Drive"})

    const dialogoProcesando = this.dialog.open(DialogoComponent, {
      disableClose: true,
      data: { tipoDialogo: "procesando", titulo: "Procesando", contenido: "" },
    });

    console.log("Subiendo Cursos...");
    console.log("FORMADORES: ");
    console.log(this.formadores);
    console.log("META FORMADORES: ");
    console.log(this.metadatosFormadores);
    console.log("INSTITUCIONES: ");
    console.log(this.instituciones);
    console.log("META INSTITUCIONES: ");
    console.log(this.metadatosInstituciones);

    //Asignar Metadatos en los objetos:
    var indexFormador = -1;
    for (var i = 0; i < this.formadores[0].data.length; i++) {
      indexFormador = this.binarySearchObject(
        this.metadatosFormadores[0].data,
        "cod__formador",
        this.formadores[0].data[i]["cod__formador"],
      );
      if (indexFormador >= 0) {
        this.formadores[0].data[i]["metadatos"] =
          this.metadatosFormadores[0].data[indexFormador];
      }
    }
    var indexInstitucion = -1;
    for (var i = 0; i < this.instituciones[0].data.length; i++) {
      indexInstitucion = this.binarySearchObject(
        this.metadatosInstituciones[0].data,
        "cod_institucion",
        this.instituciones[0].data[i]["cod_institucion"],
      );
      if (indexInstitucion >= 0) {
        this.instituciones[0].data[i]["metadatos"] =
          this.metadatosInstituciones[0].data[indexInstitucion];
      }
    }
    var indexCursos = -1;
    for (var i = 0; i < this.cursos.length; i++) {
      indexCursos = this.binarySearchObject(
        this.metadatosCursos[0].data,
        "cod_curso",
        this.cursos[i]["cod_curso"],
      );
      if (indexCursos >= 0) {
        this.cursos[i]["metadatos"] = this.metadatosCursos[0].data[indexCursos];
      }
    }

    this.appService
      .ejecutarProceso({ nombre: "Subir Cursos", categoria: "general" }, [
        this.rutaArchivoCursos,
        this.cursos,
        this.metadatosCursos,
        this.formadores,
        this.metadatosFormadores,
        this.formadoresCurso,
        this.codigoProvincia,
        this.instituciones,
        this.metadatosInstituciones,
      ])
      .then((result) => {
        if (result) {
          //Cambio de Flags CURSOS:
          for (var i = 0; i < this.metadatosCursos[0].data.length; i++) {
            if (
              this.metadatosCursos[0].data[i]["flag_cambio"] &&
              !this.metadatosCursos[0].data[i]["error"]
            ) {
              this.metadatosCursos[0].data[i]["flag_cambio"] = false;
              this.cursos.find(
                (j) =>
                  j.cod_curso == this.metadatosCursos[0].data[i]["cod_curso"],
              )["metadatos"] = this.metadatosCursos[0].data[i];
            }
          }

          //Cambio de Flags FORMADORES:
          for (var i = 0; i < this.metadatosFormadores[0].data.length; i++) {
            if (
              this.metadatosFormadores[0].data[i]["flag_cambio"] &&
              !this.metadatosFormadores[0].data[i]["error"]
            ) {
              this.metadatosFormadores[0].data[i]["flag_cambio"] = false;
              this.metadatosFormadores[0].data[i]["formadorAdicional"] = false;
              this.formadores[0].data.find(
                (j) =>
                  j.cod__formador ==
                  this.metadatosFormadores[0].data[i]["cod__formador"],
              )["metadatos"] = this.metadatosFormadores[0].data[i];
            }
          }

          //Cambio de Flags INSTITUCIONES:
          for (var i = 0; i < this.metadatosInstituciones[0].data.length; i++) {
            if (
              this.metadatosInstituciones[0].data[i]["flag_cambio"] &&
              !this.metadatosInstituciones[0].data[i]["error"]
            ) {
              this.metadatosInstituciones[0].data[i]["flag_cambio"] = false;
              this.metadatosInstituciones[0].data[i]["institucionAdicional"] =
                false;
              this.instituciones[0].data.find(
                (j) =>
                  j.cod_institucion ==
                  this.metadatosInstituciones[0].data[i]["cod_institucion"],
              )["metadatos"] = this.metadatosInstituciones[0].data[i];
            }
          }

          this.guardarCursos();
          console.log("Proceso finalizado: ");
          console.log(result);
          dialogoProcesando.close();
        } else {
          dialogoProcesando.close();
          this.dialog.open(DialogoComponent, {
            disableClose: true,
            data: {
              tipoDialogo: "error",
              titulo: "Se ha producido un error inesperado.",
              contenido: "",
            },
          });
        }
      });
  }

  copiarSesion(element) {
    console.log(element);
    if (!element["cod_grupo"]) {
      this.appService.openDialog("error", {
        titulo: "Error",
        contenido:
          "Error: Se debe especificar el codigo de grupo antes de copiar",
      });
      return;
    }

    console.log("Añadiendo Incidencia");
    /*
            const dialogoProcesando = this.dialog.open(DialogoComponent,{ disableClose: true,
                  data: {tipoDialogo: "procesando", titulo: "Procesando", contenido: ""}
            });
            */

    //Buscar un codigo sin asignar:
    var codigoBase = moment().format("YYYYMMDD");
    var codigoNuevo = codigoBase;

    //Iterar entre los codigos disponibles:
    var indice = "00";

    var codigoSincronizacion = this.datosProyecto.idSincronizacion;
    if (this.datosProyecto.idSincronizacion < 10) {
      codigoSincronizacion = "0" + codigoSincronizacion;
    } else {
      codigoSincronizacion = String(codigoSincronizacion);
    }

    for (var i = 0; i < 100; i++) {
      if (i < 10) {
        codigoNuevo = codigoBase + "0" + String(i) + codigoSincronizacion;
      } else {
        codigoNuevo = codigoBase + String(i) + codigoSincronizacion;
      }

      console.log(this.metadatosCursos);
      if (
        !this.metadatosCursos[0].data.find((j) => j.cod_curso == codigoNuevo)
      ) {
        break;
      }
    }

    console.log("Codigo Nuevo: " + codigoNuevo);

    //Añadiendo incidencia a metadatos:
    this.metadatosCursos[0].data.push({
      cod_curso: codigoNuevo,
      cod_grupo: element["cod_grupo"],
      curso: element["curso"],
      sesión: element["sesión"],
      colectivo: element["colectivo"],
      nºasistentes: element["nºasistentes"],
      modalidad: "Presencial",
      estado: "PROGRAMADA",
      material: "SI",
      incidenciaAdicional: true,
      servicio: "",
      equipo: "",
      log: [],
      revisado: false,
      tipo_cambio: "Cola",
      ultima_revision: 44833,
      wa: false,
      week_resolution: "-",
      flag: 0,
      flag_cambio: true,
      error: true,
      errores: {
        postal: true,
        ccaa: true,
        cod_grupo: true,
        territorial: true,
        institucion: true,
        fecha: true,
        hora_inicio: true,
        hora_fin: true,
        curso: true,
        sesion: true,
        colectivo: true,
        grupo: true,
        alumnos: true,
        formadores: true,
      },
      comentarios: [],
      descripcion_ejecutiva: "Descripción ejecutiva",
      en_crq: false,
      responsable: "",
    });

    const dialogoProcesandoCarga = this.dialog.open(DialogoComponent, {
      disableClose: true,
      data: { tipoDialogo: "procesando", titulo: "Procesando", contenido: "" },
    });

    dialogoProcesandoCarga.afterOpened().subscribe((result) => {
      this.inicializarDatos(dialogoProcesandoCarga, {
        omitirRecarga: true,
      }).then((dialogo: MatDialogRef<DialogoComponent>) => {
        this.buscarCodigo(codigoNuevo);
        dialogo.close();
      });
    });
  } //FIN COPIA

  seleccionInstitucion(event, element) {
    //Añadir Formador:
    console.log("Añadiendo Institucion:");
    console.log(event);
    console.log(element);

    var indexDataTable = this.dataTable.data.findIndex(
      (i) => i.cod_curso == element.cod_curso,
    );

    //Actualización de Metadatos:
    this.metadatosCursos[0].data.find((i) => i.cod_curso == element.cod_curso)[
      "institucion"
    ] = event.option.value;
    //this.metadatosCursos[0].data.find(i => i.cod_curso==element.cod_curso)["institucionModificada"] = true;
    this.metadatosCursos[0].data.find((i) => i.cod_curso == element.cod_curso)[
      "flag_cambio"
    ] = true;

    var indexBusquedaInstitucion = this.binarySearchObject(
      this.instituciones[0].data,
      "cod_institucion",
      event.option.value["id"],
    );
    console.warn("Busqueda: ", indexBusquedaInstitucion);
    if (indexBusquedaInstitucion != -1) {
      this.dataTable.data[indexDataTable]["nombreInstitucion"] =
        this.instituciones[0].data[indexBusquedaInstitucion]["institucion"];
      this.cursos.find((i) => i.cod_curso == element.cod_curso)["institución"] =
        this.instituciones[0].data[indexBusquedaInstitucion]["cod_institucion"];
    }

    //Actualización de Render:
    this.cursos.find((i) => i.cod_curso == element.cod_curso)["metadatos"] =
      this.metadatosCursos[0].data.find(
        (i) => i.cod_curso == element.cod_curso,
      );
    this.dataTable.data[indexDataTable]["metadatos"] =
      this.metadatosCursos[0].data.find(
        (i) => i.cod_curso == element.cod_curso,
      );
    this.dataTable.data[indexDataTable]["metadatos"]["flag_cambio"] = true;

    this.comprobarDatos(element.cod_curso);
  } //Fin Seleccion Formador

  seleccionFormador(event, element) {
    //Añadir Formador:
    console.log("Añadiendo Formador:");
    console.log(event);
    console.log(element);

    //Actualización de Metadatos:
    this.metadatosCursos[0].data
      .find((i) => i.cod_curso == element.cod_curso)
      ["formadores"].push(event.option.value);
    this.metadatosCursos[0].data.find((i) => i.cod_curso == element.cod_curso)[
      "formadorModificado"
    ] = true;
    this.metadatosCursos[0].data.find((i) => i.cod_curso == element.cod_curso)[
      "flag_cambio"
    ] = true;

    //Actualización de Render:
    this.cursos.find((i) => i.cod_curso == element.cod_curso)["metadatos"] =
      this.metadatosCursos[0].data.find(
        (i) => i.cod_curso == element.cod_curso,
      );
    this.dataTable.data.find((i) => i.cod_curso == element.cod_curso)[
      "metadatos"
    ] = this.metadatosCursos[0].data.find(
      (i) => i.cod_curso == element.cod_curso,
    );
    this.dataTable.data.find((i) => i.cod_curso == element.cod_curso)[
      "metadatos"
    ]["flag_cambio"] = true;
    this.autoFormadorControl.setValue({
      nombre: "",
      id: "",
    });
    this.comprobarDatos(element.cod_curso);
  } //Fin Seleccion Formador

  eliminarFormadorLista(id, element) {
    //Eliminando Formador:
    console.log("Eliminando Formador:");
    for (
      var i = 0;
      this.metadatosCursos[0].data.find(
        (i) => i.cod_curso == element.cod_curso,
      )["formadores"].length;
      i++
    ) {
      if (
        this.metadatosCursos[0].data.find(
          (j) => j.cod_curso == element.cod_curso,
        )["formadores"][i]["id"] == id
      ) {
        this.metadatosCursos[0].data
          .find((j) => j.cod_curso == element.cod_curso)
          ["formadores"].splice(i, 1);
      }
    }

    this.metadatosCursos[0].data.find((i) => i.cod_curso == element.cod_curso)[
      "formadorModificado"
    ] = true;
    this.metadatosCursos[0].data.find((i) => i.cod_curso == element.cod_curso)[
      "flag_cambio"
    ] = true;

    //Actualización de Render:
    this.cursos.find((i) => i.cod_curso == element.cod_curso)["metadatos"] =
      this.metadatosCursos[0].data.find(
        (i) => i.cod_curso == element.cod_curso,
      );
    this.dataTable.data.find((i) => i.cod_curso == element.cod_curso)[
      "metadatos"
    ] = this.metadatosCursos[0].data.find(
      (i) => i.cod_curso == element.cod_curso,
    );
    this.dataTable.data.find((i) => i.cod_curso == element.cod_curso)[
      "metadatos"
    ]["flag_cambio"] = true;

    this.comprobarDatos(element.cod_curso);
  }

  descargarDatos2() {
    console.log("Descargando Cursos...");
    const dialogoProcesando = this.dialog.open(DialogoComponent, {
      disableClose: true,
      data: { tipoDialogo: "procesando", titulo: "Procesando", contenido: "" },
    });

    console.log(this.rutaArchivoCursos);

    if (this.rutaArchivoCursos == "") {
      console.error(
        "Error: No se ha especificado la ruta de descarga de los cursos.",
      );
      this.dialog.open(DialogoComponent, {
        disableClose: true,
        data: {
          tipoDialogo: "error",
          titulo:
            "Se ha producido un error inesperado. La ruta al archivo de monitorización es incorrecta.",
          contenido: "",
        },
      });
    }

    //Descargar Cursos:
    this.appService
      .ejecutarProceso({ nombre: "Importar Cursos", categoria: "import" }, [
        this.rutaArchivoCursos,
        1,
      ])
      .then((result) => {
        this.inicializarDatos();
        dialogoProcesando.close();
        if (!result) {
          dialogoProcesando.close();
          this.dialog.open(DialogoComponent, {
            disableClose: true,
            data: {
              tipoDialogo: "error",
              titulo: "Se ha producido un error inesperado.",
              contenido: "",
            },
          });
        }
      });
  }

  descargarDatos() {
    console.log("Descargando Cursos...");
    const dialogoProcesando = this.dialog.open(DialogoComponent, {
      disableClose: true,
      data: { tipoDialogo: "procesando", titulo: "Procesando", contenido: "" },
    });

    //Descargar Cursos:
    this.appService
      .ejecutarProceso({ nombre: "Importar Excel", categoria: "import" }, [
        this.rutaArchivoCursos,
        1,
        "Cursos",
        "Cursos",
      ])
      .then((result) => {
        if (!result) {
          dialogoProcesando.close();
          this.dialog.open(DialogoComponent, {
            disableClose: true,
            data: {
              tipoDialogo: "error",
              titulo: "Se ha producido un error inesperado.",
              contenido: "",
            },
          });
        }

        //Descargar Formadores-Curso
        this.appService
          .ejecutarProceso({ nombre: "Importar Excel", categoria: "import" }, [
            this.rutaArchivoCursos,
            1,
            "Formador-Curso",
            "Formador-Curso",
          ])
          .then((result) => {
            if (!result) {
              dialogoProcesando.close();
              this.dialog.open(DialogoComponent, {
                disableClose: true,
                data: {
                  tipoDialogo: "error",
                  titulo: "Se ha producido un error inesperado.",
                  contenido: "",
                },
              });
            }
            //Descargar Formadores:
            this.appService
              .ejecutarProceso(
                { nombre: "Importar Excel", categoria: "import" },
                [this.rutaArchivoCursos, 1, "Formadores", "Formadores"],
              )
              .then((result) => {
                if (!result) {
                  dialogoProcesando.close();
                  this.dialog.open(DialogoComponent, {
                    disableClose: true,
                    data: {
                      tipoDialogo: "error",
                      titulo: "Se ha producido un error inesperado.",
                      contenido: "",
                    },
                  });
                }
                //Descargar Códigos Postal:
                this.appService
                  .ejecutarProceso(
                    { nombre: "Importar Excel", categoria: "import" },
                    [
                      this.rutaArchivoCursos,
                      1,
                      "Codigos Provincia",
                      "Códigos_Provincias",
                    ],
                  )
                  .then((result) => {
                    this.inicializarDatos();
                    dialogoProcesando.close();
                    if (!result) {
                      dialogoProcesando.close();
                      this.dialog.open(DialogoComponent, {
                        disableClose: true,
                        data: {
                          tipoDialogo: "error",
                          titulo: "Se ha producido un error inesperado.",
                          contenido: "",
                        },
                      });
                    }
                  });
              });
          });
      });
  }

  incluirRuta(evt: any, indexControl: number) {
    //Lectura de evento Input
    const target: DataTransfer = <DataTransfer>evt.target;

    console.log("Objeto Ruta:");
    console.log(target.files);

    var formularioTemporal: any;
    //formularioTemporal= this.formularioProcesoGroup.getRawValue();
    //formularioTemporal.formularioControl[indexControl]= target.files[0].path;
    this.formularioControl[indexControl] = target.files[0]["path"];
    //this.formularioProcesoGroup.setValue(formularioTemporal);
  }

  checkError() {
    console.log("Comprobando Errores");
    const dialogoProcesando = this.dialog.open(DialogoComponent, {
      disableClose: true,
      data: { tipoDialogo: "procesando", titulo: "Procesando", contenido: "" },
    });

    dialogoProcesando.afterOpened().subscribe(() => {
      if (this.dataTable.data.length) {
        for (var i = 0; i < this.dataTable.data.length; i++) {
          this.comprobarDatos(this.dataTable.data[i]["cod_curso"]);
        }
      }
      this.filtroError = false;
      this.filtrar("Error");

      dialogoProcesando.close();
    });
  }

  abrirFiltroFecha() {
    console.log("Comprobando Errores");
    const dialogoProcesando = this.dialog.open(DialogoComponent, {
      disableClose: true,
      data: { tipoDialogo: "filtroFecha", titulo: "Procesando", contenido: "" },
    });

    dialogoProcesando.afterClosed().subscribe((result) => {
      console.log(result);

      if (typeof result["filtroFecha"] == "undefined") {
        return;
      }

      var fechaInicio = moment(result["filtroFecha"]["fechaInicio"]);
      var fechaFin = moment(result["filtroFecha"]["fechaFin"]);
      //var fecha = moment();

      if (!fechaInicio.isValid() || !fechaFin.isValid()) {
        this.appService.openDialog("error", {
          titulo: "Fechas no Validas",
          contenido:
            "Para poder filtrar los datos se requiere dos fechas validas",
        });
        return;
      }

      var objetoFiltroFecha = {
        fechaInicio: fechaInicio,
        fechaFin: fechaFin,
      };

      this.filtrar("Fecha Cursos", objetoFiltroFecha);

      //this.dataTable._updateChangeSubscription();
    });
  }

  binarySearchObject(arr, propiedad, val) {
    let start = 0;
    let end = arr.length - 1;
    val = Number(val);

    while (start <= end) {
      let mid = Math.floor((start + end) / 2);

      if (Number(arr[mid][propiedad]) === val) {
        return mid;
      }

      if (val < Number(arr[mid][propiedad])) {
        end = mid - 1;
      } else {
        start = mid + 1;
      }
    }
    return -1;
  }

  procesarDatosCorreos() {
    //INCLUYE CURSO E INSTITUCIÓN EN CORREOS:
    for (var i = 0; i < this.correos.length; i++) {
      this.correos[i]["curso"] = this.cursos.find(
        (j) => j.cod_curso == this.correos[i]["cod_curso"],
      );
      if (this.correos[i]["curso"] != undefined) {
        this.correos[i]["institucion"] = this.instituciones[0].data.find(
          (j) => j.cod_institucion == this.correos[i]["curso"]["institución"],
        );
      }
    }
    console.error("CORREOS: ", this.correos);
  }

  async getCorreo(asunto: string) {
    switch (asunto) {
      case "institucion":
        asunto = '"[2024032365]"';
        break;
      case "material":
        break;
      case "recordatorio":
        break;
    }

    var argumentos = {
      formularioControl: [asunto],
    };

    var proceso = {
      nombre: "getCorreosAsunto",
      categoria: "Google",
      argumentos: argumentos,
    };
    var correos = await this.appService.ejecutarProceso(proceso, argumentos);

    if (correos) {
      var plantillaInstitucion = {
        data: correos,
        nombreId: "PlantillaRecordatorio",
        objetoId: "PlantillaRecordatorio",
      };

      this.appService.guardarArchivo(plantillaInstitucion).then((result) => {
        console.error("GUARDADO CON EXITO: ", result);
      });

      console.error("CORREOS: ", correos);
      this.indexCorreoVisualizado = 0;
      this.correosVisualizados = correos["data"];
      console.error(this.correosVisualizados[0]);
    }
    return;
  }

  async guardarPlantillaCorreo(tipo: string) {
    var asunto = '"[2024032368]"';
    var nombreIdDocumento = "PlantillaInstitucion";

    var argumentos = {
      formularioControl: [asunto],
    };

    var proceso = {
      nombre: "getCorreosAsunto",
      categoria: "Google",
      argumentos: argumentos,
    };
    var correos = await this.appService.ejecutarProceso(proceso, argumentos);
    console.warn("CORREO: ", correos);

    if (correos) {
      var plantillaInstitucion = {
        data: correos,
        nombreId: nombreIdDocumento,
        objetoId: nombreIdDocumento,
      };
      this.appService.guardarArchivo(plantillaInstitucion).then((result) => {
        console.error("GUARDADO CON EXITO: ", result);
      });
    }
    return;
  }

  async crearBorrador(cuerpoEmail, asunto, destinatario, adjunto?) {
    var argumentos = {
      formularioControl: [cuerpoEmail, asunto, destinatario, adjunto],
    };

    var proceso = {
      nombre: "crearBorrador",
      categoria: "Google",
      argumentos: argumentos,
    };
    var result = await this.appService.ejecutarProceso(proceso, argumentos);

    console.warn("Crear Draft RESULT: ", result);
    return result;
  }

  async procesarDatos() {
    var codigosConflictivos = [];
    var ultimoCodigoProcesado = "";
    var encontrado = false;

    //Ordenado de cursos:

    for (var i = 0; i < this.cursos.length; i++) {
      if (this.cursos[i]["cod_grupo"] != ultimoCodigoProcesado) {
        encontrado = false;
        ultimoCodigoProcesado = this.cursos[i]["cod_grupo"];
        for (var j = 0; j < codigosConflictivos.length; j++) {
          if (codigosConflictivos[j] == this.cursos[i]["cod_grupo"]) {
            encontrado = true;
            break;
          }
        }

        //Iteración por todos los cursos:
        if (!encontrado) {
          for (var j = 0; j < this.cursos.length; j++) {
            if (this.cursos[j]["cod_grupo"] == this.cursos[i]["cod_grupo"]) {
              //         console.log("SAME")
              if (this.cursos[j]["colectivo"] != this.cursos[i]["colectivo"]) {
                codigosConflictivos.push(this.cursos[i]["cod_grupo"]);
                break;
              }
            }
          }
        }
      }
    }

    console.warn("CODIGOS CONFLICTIVOS: ", codigosConflictivos);
  }

  async generarCorreo(tipoCorreo: string, codCurso: string, parametros?: any) {
    console.error("GENERANDO CORREO... Tipo: ", tipoCorreo);

    var plantilla;
    const dialogoProcesando = this.dialog.open(DialogoComponent, {
      disableClose: true,
      data: {
        tipoDialogo: "procesando",
        titulo: "Generando Correo",
        contenido: "",
      },
    });

    switch (tipoCorreo) {
      case "institucion":
        console.error("NO HAY PLANTILLA");
        this.plantillaCorreoInstitucion = await this.appService.getCorreo(
          "PlantillaInstitucion",
        );
        plantilla = this.plantillaCorreoInstitucion;
        break;
      case "material":
        this.plantillaCorreoMaterial =
          await this.appService.getCorreo("PlantillaMaterial");
        plantilla = this.plantillaCorreoMaterial;
        break;
      case "recordatorio":
        this.plantillaCorreoRecordatorio = await this.appService.getCorreo(
          "PlantillaRecordatorio",
        );
        plantilla = this.plantillaCorreoRecordatorio;
        break;
      case "graciasInstitucion":
        this.plantillaCorreoRecordatorio = await this.appService.getCorreo(
          "PlantillaGraciasInstitucion",
        );
        plantilla = this.plantillaCorreoRecordatorio;
        break;
      case "graciasFormador":
        this.plantillaCorreoRecordatorio = await this.appService.getCorreo(
          "PlantillaGraciasFormador",
        );
        plantilla = this.plantillaCorreoRecordatorio;
        break;
    }

    console.warn("PLANTILLA SIN MODIFICAR: ", plantilla);

    if (!plantilla) {
      return false;
    }

    //Modificar la plantilla:
    //var htmlCorreo = atob(plantilla[0].data.payload.parts[0].parts[1].body.data.replace(/-/g, '+').replace(/_/g, '/'));
    //var imagenCorreo = plantilla[0].data.payload.parts[1].body.attachmentId;

    var htmlCorreo = atob(
      plantilla[0].data.mensajeHtml.replace(/-/g, "+").replace(/_/g, "/"),
    );
    var imagenesCorreo = plantilla[0].data.imagenes;

    //Corrección de Tildes:
    htmlCorreo = htmlCorreo.replaceAll("Ã¡", "á");
    htmlCorreo = htmlCorreo.replaceAll("Ã©", "é");
    htmlCorreo = htmlCorreo.replaceAll("Ã­", "í");
    htmlCorreo = htmlCorreo.replaceAll("Ã³", "ó");
    htmlCorreo = htmlCorreo.replaceAll("Ãº", "ú");
    htmlCorreo = htmlCorreo.replaceAll("Â", "");
    htmlCorreo = htmlCorreo.replaceAll("Ã±", "ñ");
    htmlCorreo = htmlCorreo.replaceAll("â", "-");
    htmlCorreo = htmlCorreo.replaceAll("â", "");
    htmlCorreo = htmlCorreo.replaceAll("â", "");
    htmlCorreo = htmlCorreo.replaceAll("â¢", "");

    //Buscar Datos del Curso:
    var curso = this.cursos.find((j) => j.cod_curso == codCurso);
    var institucion = this.instituciones[0].data.find(
      (j) => j.cod_institucion == curso["institución"],
    );

    console.warn("CURSO CORREO: ", curso);
    console.warn("INSTITUCIÓN CORREO: ", institucion);

    //Asignacion de Correo SYNC:
    var correoSync = "";
    switch (Number(this.datosProyecto.idSincronizacion)) {
      case 1:
        correoSync = "natalia.u@sanfi.es";
        break;
      case 2:
        correoSync = "laura.seco@sanfi.es";
        break;
      case 3:
        correoSync = "";
        break;
      default:
        correoSync = "noreply@finanzasparamortales.com";
        break;
    }

    console.warn("CORREO SIN REPLACE: ", htmlCorreo);

    //Formateo de Correos:
    switch (tipoCorreo) {
      case "institucion":
        htmlCorreo = htmlCorreo
          .replace("{cod_sesion}", codCurso)
          .replaceAll("{institucion}", institucion.institucion)

          .replace("{programa}", this.getTituloPrograma(curso["sesión"]))
          .replace("{hora}", curso["hora_inicio_formateada"])
          .replace("{contacto_institucion}", institucion["contacto1"])
          .replace("{duracion}", curso["duracion_formateada"])

          .replace("{correoSync}", correoSync)
          .replace(
            "{fecha}",
            curso.fecha_formateada.toLocaleDateString("es-ES"),
          );
        break;
      case "material":
        if (!institucion["telefono1"]) {
          institucion["telefono1"] = "Teléfono no especificado";
        }
        if (!institucion["direccion"]) {
          institucion["direccion"] = "Dirección no especificado";
        }
        htmlCorreo = htmlCorreo
          .replace("[cod_curso]", codCurso)
          .replaceAll("{institucion}", institucion.institucion)

          .replace(
            "{fecha}",
            curso.fecha_formateada.toLocaleDateString("es-ES"),
          )
          .replace("{hora}", curso["hora_inicio_formateada"])

          .replace("{contacto_institucion}", institucion["contacto1"])
          .replace("{material}", this.getTituloPrograma(curso["sesión"]))
          .replace("{correo_institucion}", institucion["email1"])
          .replace("{telefono_institucion}", institucion["telefono1"])
          .replace("{direccion_institucion}", institucion["direccion"])
          .replace("{correoSync}", correoSync);
        break;

      case "recordatorio":
        var htmlFormadores = "";
        var urlEncuestaBeneficiario =
          "https://form.typeform.com/to/NRDiAJBX#cod_sesion=" +
          codCurso +
          "&fecha_sesion=" +
          curso.fecha_formateada.toLocaleDateString("es-ES");
        if (!institucion["telefono1"]) {
          institucion["telefono1"] = "Teléfono no especificado";
        }
        if (!institucion["direccion"]) {
          institucion["direccion"] = "Dirección no especificado";
        }
        for (var i = 0; i < curso.metadatos.formadores.length; i++) {
          if (
            curso.metadatos.formadores[i]["nombre"] &&
            curso.metadatos.formadores[i]["email"]
          ) {
            htmlFormadores +=
              '<p><span style="font-family:&quot;Arial&quot;,sans-serif">' +
              curso.metadatos.formadores[i]["nombre"] +
              '<o:p></o:p></span></p><p><span style="font-family:&quot;Arial&quot;,sans-serif">' +
              curso.metadatos.formadores[i]["email"] +
              "<o:p></o:p></span></p>";
          }
        }

        htmlCorreo = htmlCorreo
          .replaceAll("{cod_sesion}", codCurso)
          .replaceAll("{institucion}", institucion.institucion)
          .replace(
            "{fecha}",
            curso.fecha_formateada.toLocaleDateString("es-ES"),
          )
          .replace("{hora}", curso["hora_inicio_formateada"])
          .replace("{duracion}", curso["duracion_formateada"])
          .replace("{encuesta_alumnos}", urlEncuestaBeneficiario)

          .replace("{contacto_institucion}", institucion["contacto1"])
          .replace("{emailInstitucion}", institucion["email1"])
          .replace("{telefono_institucion}", institucion["telefono1"])
          .replace("{direccion_institucion}", institucion["direccion"])

          .replace("{formador}", htmlFormadores);

        break;

      case "graciasFormador":
        htmlFormadores = "";

        var urlEncuestaFormador =
          "https://form.typeform.com/to/yxdU1rlU#cod_sesion=" +
          codCurso +
          "&fecha_sesion=" +
          curso.fecha_formateada.toLocaleDateString("es-ES");
        for (var i = 0; i < curso.metadatos.formadores.length; i++) {
          if (
            curso.metadatos.formadores[i]["nombre"] &&
            curso.metadatos.formadores[i]["email"]
          ) {
            htmlFormadores += "" + curso.metadatos.formadores[i]["nombre"] + "";
          }
          if (i < curso.metadatos.formadores.length - 1) {
            htmlFormadores += ", ";
          }
          if (i == curso.metadatos.formadores.length - 1) {
            htmlFormadores += " ";
          }
        }
        htmlCorreo = htmlCorreo
          .replace("{formador}", htmlFormadores)
          .replace("{programa}", this.getTituloPrograma(curso["sesión"]))
          .replace("{encuesta_formadores}", urlEncuestaFormador);
        break;

      case "graciasInstitucion":
        var urlEncuestaInstitucion =
          "https://form.typeform.com/to/BRejHo2Y#cod_sesion=" +
          codCurso +
          "&fecha_sesion=" +
          curso.fecha_formateada.toLocaleDateString("es-ES");
        htmlCorreo = htmlCorreo
          .replace("{institucion}", institucion.institucion)
          .replace("{contacto_institucion}", institucion["contacto1"])
          .replace("{programa}", this.getTituloPrograma(curso["sesión"]))
          .replace("{encuesta_institucion}", urlEncuestaInstitucion);
        break;
    }

    //Visualizar el correo:
    this.correosVisualizados[0] = htmlCorreo;
    this.indexCorreoVisualizado = 0;
    console.warn(this.correosVisualizados[0]);

    //Creación del Borrador:
    var asunto = "";
    var destinatario = [];
    var adjuntos = [];
    switch (tipoCorreo) {
      case "institucion":
        asunto =
          "Propuesta Sesiones FxM - " +
          curso.nombreInstitucion +
          " [#" +
          parametros[1] +
          "]";
        if (institucion["email1"]) {
          destinatario.push(institucion["email1"]);
        }
        if (institucion["email2"]) {
          destinatario.push(institucion["email2"]);
        }
        if (institucion["email3"]) {
          destinatario.push(institucion["email3"]);
        }
        if (institucion["email4"]) {
          destinatario.push(institucion["email4"]);
        }

        adjuntos.push({
          inline: true,
          filename: "logo_sanfi_correo.png",
          contentType: "image/png;base64",
          data: imagenesCorreo[0],
          headers: { "Content-ID": "image001" },
        });

        break;
      case "material":
        asunto =
          "Material y contacto FxM - " +
          curso.nombreInstitucion +
          " [#" +
          parametros[1] +
          "]";
        for (var i = 0; i < curso.metadatos.formadores.length; i++) {
          if (!curso.metadatos.formadores[i]["email"]) {
            continue;
          }
          destinatario.push(curso.metadatos.formadores[i]["email"]);
        }

        adjuntos.push({
          inline: true,
          filename: "logo_sanfi_correo.png",
          contentType: "image/png;base64",
          data: imagenesCorreo[0],
          headers: { "Content-ID": "image001" },
        });

        break;
      case "recordatorio":
        asunto =
          "Recordatorio sesiones y encuestas - " +
          curso.nombreInstitucion +
          " [#" +
          parametros[1] +
          "]";
        if (institucion["email1"]) {
          destinatario.push(institucion["email1"]);
        }
        if (institucion["email2"]) {
          destinatario.push(institucion["email2"]);
        }
        if (institucion["email3"]) {
          destinatario.push(institucion["email3"]);
        }
        if (institucion["email4"]) {
          destinatario.push(institucion["email4"]);
        }

        for (var i = 0; i < curso.metadatos.formadores.length; i++) {
          if (!curso.metadatos.formadores[i]["email"]) {
            continue;
          }
          destinatario.push(curso.metadatos.formadores[i]["email"]);
        }

        adjuntos.push({
          inline: true,
          filename: "logo_sanfi_correo.png",
          contentType: "image/png;base64",
          data: imagenesCorreo[0],
          headers: { "Content-ID": "image001" },
        });

        adjuntos.push({
          inline: true,
          qr: true,
          urlQR: urlEncuestaBeneficiario,
          filename: "qr_encuesta.png",
          contentType: "image/png;base64",
          data: imagenesCorreo[1],
          headers: { "Content-ID": "image002" },
        });

        break;

      case "graciasFormador":
        asunto =
          "Encuestas Formador - " +
          curso.nombreInstitucion +
          " [#" +
          parametros[1] +
          "]";

        for (var i = 0; i < curso.metadatos.formadores.length; i++) {
          if (!curso.metadatos.formadores[i]["email"]) {
            continue;
          }
          destinatario.push(curso.metadatos.formadores[i]["email"]);
        }

        adjuntos.push({
          inline: true,
          filename: "logo_sanfi_correo.png",
          contentType: "image/png;base64",
          data: imagenesCorreo[0],
          headers: { "Content-ID": "image001" },
        });

        adjuntos.push({
          inline: true,
          qr: true,
          urlQR: urlEncuestaFormador,
          filename: "qr_encuesta.png",
          contentType: "image/png;base64",
          data: imagenesCorreo[1],
          headers: { "Content-ID": "image002" },
        });

        break;

      case "graciasInstitucion":
        asunto =
          "Encuestas Institución - " +
          curso.nombreInstitucion +
          " [#" +
          parametros[1] +
          "]";
        if (institucion["email1"]) {
          destinatario.push(institucion["email1"]);
        }
        if (institucion["email2"]) {
          destinatario.push(institucion["email2"]);
        }
        if (institucion["email3"]) {
          destinatario.push(institucion["email3"]);
        }
        if (institucion["email4"]) {
          destinatario.push(institucion["email4"]);
        }

        adjuntos.push({
          inline: true,
          filename: "logo_sanfi_correo.png",
          contentType: "image/png;base64",
          data: imagenesCorreo[0],
          headers: { "Content-ID": "image001" },
        });

        adjuntos.push({
          inline: true,
          qr: true,
          urlQR: urlEncuestaInstitucion,
          filename: "qr_encuesta.png",
          contentType: "image/png;base64",
          data: imagenesCorreo[1],
          headers: { "Content-ID": "image002" },
        });

        break;
    }

    var idBorrador = await this.crearBorrador(
      htmlCorreo,
      asunto,
      destinatario,
      adjuntos,
    );

    if (idBorrador) {
      var indexMetadatosCurso = this.binarySearchObject(
        this.metadatosCursos[0].data,
        "cod_curso",
        curso.cod_curso,
      );
      var indexCorreo = -1;
      indexCorreo = this.binarySearchObject(
        this.correos,
        "cod_curso",
        curso.cod_curso,
      );

      if (indexCorreo != -1) {
        switch (tipoCorreo) {
          case "institucion":
            this.correos[indexCorreo].idDraftInstitucion = parametros[1];
            break;
          case "material":
            this.correos[indexCorreo].idDraftMaterial = parametros[1];
            break;
          case "recordatorio":
            this.correos[indexCorreo].idDraftRecordatorio = parametros[1];
            break;
          case "graciasInstitucion":
            this.correos[indexCorreo].idDraftGraciasInstitucion = parametros[1];
            break;
          case "graciasFormador":
            this.correos[indexCorreo].idDraftGraciasFormador = parametros[1];
            break;
        }
      }

      if (indexCorreo != -1) {
        this.checkEstadoCorreo(indexCorreo);
        console.warn("ACTUALIZANDO TABLA CORREO: ", this.correos[indexCorreo]);
        this.tablaCorreos = new MatTableDataSource(this.correos);
        this.tablaCorreos.filterPredicate = this.filtradoCorreos;
        this.tablaCorreos.filter = {};
      }

      if (indexMetadatosCurso != -1) {
        var nombrePropiedad = "borradorId_" + tipoCorreo;
        this.metadatosCursos[0].data[indexMetadatosCurso][nombrePropiedad] =
          idBorrador;
        this.appService
          .guardarArchivo(this.metadatosCursos[0])
          .then((result) => {
            dialogoProcesando.close();
          });
        return;
      } else {
        console.error(
          "NO SE HA ENCONTRADO EL CURSO... Error asignando idBorrador al curso",
        );
        dialogoProcesando.close();
      }
    } else {
    }

    console.error("ERROR AL CREAR BORRADOR... ID BORRADOR: ", idBorrador);
    dialogoProcesando.close();
    return;
  }

  reGenerarCorreo(tipoCorreo: string, codCurso: string, parametros?: any) {
    const dialogRef = this.dialog.open(DialogoComponent, {
      disableClose: true,
      data: {
        tipoDialogo: "confirmacion",
        titulo: "El correo ya ha sido generado",
        contenido:
          "Este correo ya ha sido previamente generado.Esta operación puede acabar generando un correo duplicado. ¿Desea volver a generarlo?",
      },
    });

    dialogRef.afterClosed().subscribe((result) => {
      //Guardar ruta curso en parametros:
      console.log("Ruta del archivo: ");
      console.log(result);
      console.log("Parametros: ", tipoCorreo, codCurso, parametros);
      if (result) {
        this.generarCorreo(tipoCorreo, codCurso, parametros);
      } else {
        return;
      }
    });
  }

  getTituloPrograma(sesion: string): string {
    console.warn("TIPOLOGIA: ", this.tipología);
    console.warn("Sesion: ", sesion);
    var tipologia = this.tipología[0].data.find((i) => i.código == sesion);
    console.warn("TIPOLOGIA: ", tipologia);
    if (tipologia) {
      return tipologia["material"];
    } else {
      return "Sin definir";
    }
  } //Fin getTituloPrograma

  //Refresca el autocompletado de los formadores
  refreshFormadores(config?: any) {
    if (!config) {
      config = { omitirCarga: false };
    }
    if (!config["omitirCarga"]) {
      var dialogoProcesandoRefresh = this.dialog.open(DialogoComponent, {
        disableClose: true,
        data: {
          tipoDialogo: "procesando",
          titulo: "Recargando Formadores",
          contenido: "",
        },
      });
    }
    this.opcionesFormador = [];
    for (var i = 0; i < this.formadores[0].data.length; i++) {
      this.opcionesFormador.push({
        nombre: this.formadores[0].data[i].nombre,
        id: this.formadores[0].data[i]["cod__formador"],
      });
      if (!this.opcionesFormador[i].nombre) {
        this.opcionesFormador[i].nombre = "Sin Nombre";
      }
    }
    if (!config["omitirCarga"]) {
      dialogoProcesandoRefresh.close();
    }
    return;
  }

  //Refresca el autocompletado de las instituciones
  refreshInstituciones(config?: any) {
    if (!config) {
      config = { omitirCarga: false };
    }
    if (!config["omitirCarga"]) {
      var dialogoProcesandoRefresh = this.dialog.open(DialogoComponent, {
        disableClose: true,
        data: {
          tipoDialogo: "procesando",
          titulo: "Recargando Instituciones",
          contenido: "",
        },
      });
    }
    this.opcionesInstituciones = [];
    for (var i = 0; i < this.instituciones[0].data.length; i++) {
      this.opcionesInstituciones.push({
        institucion: this.instituciones[0].data[i].institucion,
        id: this.instituciones[0].data[i]["cod_institucion"],
      });
      if (!this.opcionesInstituciones[i].institucion) {
        this.opcionesInstituciones[i].institucion = "Sin Nombre";
      }
    }
    if (!config["omitirCarga"]) {
      dialogoProcesandoRefresh.close();
    }
    return;
  }

  toggleCargaCursosCompleta() {
    this.reducirTablaCursos = !this.reducirTablaCursos;

    const dialogoProcesandoCarga = this.dialog.open(DialogoComponent, {
      disableClose: true,
      data: { tipoDialogo: "procesando", titulo: "Procesando", contenido: "" },
    });

    dialogoProcesandoCarga.afterOpened().subscribe(() => {
      this.inicializarDatos(dialogoProcesandoCarga).then(
        (dialogo: MatDialogRef<DialogoComponent>) => {
          dialogo.close();
        },
      );
    });
  }

  toggleCargaCorreosCompleta() {
    this.reducirTablaCorreos = !this.reducirTablaCorreos;

    const dialogoProcesandoCarga = this.dialog.open(DialogoComponent, {
      disableClose: true,
      data: { tipoDialogo: "procesando", titulo: "Procesando", contenido: "" },
    });

    dialogoProcesandoCarga.afterOpened().subscribe(() => {
      this.refreshAutomatizacion().then(() => {
        dialogoProcesandoCarga.close();
      });
    });
  }

  async refreshAutomatizacion() {
    console.warn("----> Actualizando automatización <---");

    //Identifica los cursos programados
    var cursosProgramados = [];
    this.correos = [];

    for (var i = 0; i < this.cursos.length; i++) {
      if (this.cursos[i].estado == "PROGRAMADA") {
        if (this.reducirTablaCorreos) {
          if (
            Number(this.cursos[i].cod_curso) > 9999999999 &&
            Number(String(this.cursos[i].cod_curso).slice(-2)) ==
              Number(this.datosProyecto.idSincronizacion)
          ) {
            cursosProgramados.push(this.cursos[i]);
          }
        } else {
          cursosProgramados.push(this.cursos[i]);
        }
      }
    }

    console.warn("Cursos programados: ", cursosProgramados);
    console.warn("Correos PRE:", this.correos);
    console.warn("Instituciones:", this.instituciones);
    console.warn("Formadores:", this.formadores);

    //Generar archivo correos:
    var institucion = {};
    var formadores = [];
    var metadatos = {};
    var indexInstitucion = -1;
    var indexMetadatosCurso = -1;
    var idDraftInstitucion = null;
    var idDraftMaterial = null;
    var idDraftRecordatorio = null;
    var idDraftGraciasFormador = null;
    var idDraftGraciasInstitucion = null;
    var forzarCorreoInstitucion = false;

    //Check de borradores:
    var checkBorradoresID = [];
    var checkBorradoresCodCurso = [];

    //Incluir Correos de nuevos cursos programados:
    for (var i = 0; i < cursosProgramados.length; i++) {
      if (
        this.binarySearchObject(
          this.correos,
          "cod_curso",
          cursosProgramados[i]["cod_curso"],
        ) == -1
      ) {
        metadatos = null;
        institucion = null;
        idDraftInstitucion = null;
        idDraftMaterial = null;
        idDraftRecordatorio = null;
        idDraftGraciasFormador = null;
        idDraftGraciasInstitucion = null;
        forzarCorreoInstitucion = false;

        //Buscar Institución:
        indexInstitucion = this.binarySearchObject(
          this.instituciones[0].data,
          "cod_institucion",
          cursosProgramados[i]["institución"],
        );
        if (indexInstitucion != -1) {
          institucion = this.instituciones[0].data[indexInstitucion];
        }

        //Buscar Metadatos Curso:
        indexMetadatosCurso = this.binarySearchObject(
          this.metadatosCursos[0].data,
          "cod_curso",
          cursosProgramados[i]["cod_curso"],
        );
        if (indexMetadatosCurso != -1) {
          metadatos = this.metadatosCursos[0].data[indexMetadatosCurso];
        } else {
          console.error(
            "Metadatos no encontrados para curso: ",
            cursosProgramados[i]["cod_curso"],
          );
        }

        //Buscar formador:
        if (metadatos) {
          formadores = metadatos["formadores"];
          idDraftInstitucion = metadatos["borradorId_institucion"];
          idDraftMaterial = metadatos["borradorId_material"];
          idDraftRecordatorio = metadatos["borradorId_recordatorio"];
          idDraftGraciasInstitucion =
            metadatos["borradorId_graciasInstitucion"];
          idDraftGraciasFormador = metadatos["borradorId_graciasFormador"];
          forzarCorreoInstitucion = metadatos["forzarCorreoInstitucion"];
        }

        //Asignar email a formadores:
        if (!formadores) {
          formadores = [];
        } else {
          for (var j = 0; j < formadores.length; j++) {
            try {
              formadores[j]["email"] = this.formadores[0].data.find(
                (k) => k.cod__formador == formadores[j]["id"],
              )["email"];
            } catch (e) {
              console.error(
                "Formador no encontrado en la asignacion de email... ",
              );
            }
          }
        }

        //Añadir curso no encontrado en correos:
        this.correos.push({
          cod_curso: cursosProgramados[i]["cod_curso"],
          cod_correo_institucion: cursosProgramados[i]["cod_curso"] + "C1",
          cod_correo_material: cursosProgramados[i]["cod_curso"] + "C2",
          cod_correo_recordatorio: cursosProgramados[i]["cod_curso"] + "C3",
          cod_correo_graciasInstitucion:
            cursosProgramados[i]["cod_curso"] + "C4",
          cod_correo_graciasFormador: cursosProgramados[i]["cod_curso"] + "C5",
          estado: "PROGRAMADA",
          metadatos: metadatos,
          institucion: institucion,
          formadores: formadores,
          fecha_formateada: cursosProgramados[i]["fecha_formateada"],
          fecha: moment(cursosProgramados[i]["fecha_formateada"]).format(
            "DD/MM/YYYY",
          ),
          confirmacion_centro: "SIN GENERAR",
          confirmacion_material: "SIN GENERAR",
          confirmacion_graciasInstitucion: "SIN GENERAR",
          confirmacion_graciasFormador: "SIN GENERAR",
          recordatorio: "SIN GENERAR",
          fecha_confirmacion_centro: 0,
          fecha_confirmacion_material: 0,
          fecha_confirmacion_graciasFormador: 0,
          fecha_confirmacion_graciasInstitucion: 0,
          id_recordatorio: 0,
          mensajeFaseI: "",
          mensajeFaseII: "",
          mensajeFaseIII: "",
          mensajeFaseIVFormador: "",
          mensajeFaseIVInstitucion: "",
          idDraftInstitucion: idDraftInstitucion,
          idDraftMaterial: idDraftMaterial,
          idDraftRecordatorio: idDraftRecordatorio,
          idDraftGraciasFormador: idDraftGraciasFormador,
          idDraftGraciasInstitucion: idDraftGraciasInstitucion,
          flagsEstados: ["OK", "warn", "error"],
          flagEstadoGeneral: "OK",
          forzarCorreoInstitucion: forzarCorreoInstitucion,
        });
      }
    }

    //Lista de borradores:
    for (var i = 0; i < this.correos.length; i++) {
      if (this.correos[i].idDraftInstitucion) {
        checkBorradoresID.push(this.correos[i].cod_correo_institucion);
        checkBorradoresCodCurso.push(this.correos[i]["cod_correo_institucion"]);
      }
      if (this.correos[i].idDraftMaterial) {
        checkBorradoresID.push(this.correos[i].cod_correo_material);
        checkBorradoresCodCurso.push(this.correos[i]["cod_correo_material"]);
      }
      if (this.correos[i].idDraftRecordatorio) {
        checkBorradoresID.push(this.correos[i].cod_correo_recordatorio);
        checkBorradoresCodCurso.push(
          this.correos[i]["cod_correo_recordatorio"],
        );
      }
      if (this.correos[i].idDraftGraciasFormador) {
        checkBorradoresID.push(this.correos[i].cod_correo_graciasFormador);
        checkBorradoresCodCurso.push(
          this.correos[i]["cod_correo_graciasFormador"],
        );
      }
      if (this.correos[i].idDraftGraciasInstitucion) {
        checkBorradoresID.push(this.correos[i].cod_correo_graciasInstitucion);
        checkBorradoresCodCurso.push(
          this.correos[i]["cod_correo_graciasInstitucion"],
        );
      }
    }

    //CHECK BORRADORES:
    var querryListadoBorradores = "in:sent (";

    if (checkBorradoresID.length > 0) {
      for (var i = 0; i < checkBorradoresID.length; i++) {
        if (i != checkBorradoresID.length - 1) {
          querryListadoBorradores += "[" + checkBorradoresID[i] + "]|";
        } else {
          querryListadoBorradores += "[" + checkBorradoresID[i] + "])";
        }
      }

      console.warn("Querry Borradores: ", querryListadoBorradores);

      this.listaBorradores = await this.appService.ejecutarProceso(
        {
          nombre: "obtenerCorreos",
          categoria: "Google",
          argumentos: { formularioControl: [querryListadoBorradores] },
        },
        { formularioControl: [querryListadoBorradores] },
      );
    } else {
      this.listaBorradores = [];
    }

    console.warn("Lista Borradores: ", this.listaBorradores);

    //Comprobación de estados:
    var checkBorradores = [];
    for (var i = 0; i < this.correos.length; i++) {
      if (this.correos[i].estado != "PROGRAMADA") {
        continue;
      }
      this.checkEstadoCorreo(i);
    }

    console.warn("Correos POST:", this.correos);

    //Creación de la tabla de correos:
    this.tablaCorreos = new MatTableDataSource(this.correos);
    this.tablaCorreos.paginator = this.paginatorCorreos;
    this.tablaCorreos.filterPredicate = this.filtradoCorreos;
    this.tablaCorreos.filter = {};
    this.correosActualizados = true;

    //this.tablaCorreos._updateChangeSubscription();
  }

  checkEstadoCorreo(indexCorreo: number) {
    if (!this.correos[indexCorreo].formadores) {
      this.correos[indexCorreo]["formadores"] = [];
    }

    this.correos[indexCorreo]["flagsEstados"] = ["OK", "OK", "OK", "OK", "OK"];
    this.correos[indexCorreo].confirmacion_centro = "OK";
    this.correos[indexCorreo].confirmacion_material = "OK";
    this.correos[indexCorreo].confirmacion_recordatorio = "OK";
    this.correos[indexCorreo].confirmacion_graciasFormador = "OK";
    this.correos[indexCorreo].confirmacion_graciasInstitucion = "OK";
    this.correos[indexCorreo].mensajeFaseI = "";
    this.correos[indexCorreo].mensajeFaseII = "";
    this.correos[indexCorreo].mensajeFaseIII = "";
    this.correos[indexCorreo].mensajeFaseIVFormador = "";
    this.correos[indexCorreo].mensajeFaseIVInstitucion = "";

    //SI NO HAY INSTITUCIÓN:
    if (
      this.correos[indexCorreo].institucion == null ||
      this.correos[indexCorreo].institucion == undefined
    ) {
      this.correos[indexCorreo].institucion = {};
      this.correos[indexCorreo].flagsEstados[0] = "error";
      this.correos[indexCorreo].flagsEstados[1] = "error";
      this.correos[indexCorreo].flagsEstados[2] = "error";
      this.correos[indexCorreo].flagsEstados[4] = "error";
      this.correos[indexCorreo].confirmacion_centro = "ERROR";
      this.correos[indexCorreo].confirmacion_material = "ERROR";
      this.correos[indexCorreo].confirmacion_recordatorio = "ERROR";
      this.correos[indexCorreo].confirmacion_graciasInstitucion = "ERROR";
      this.correos[indexCorreo].mensajeFaseI =
        "No hay institución asociada al curso. ";
      this.correos[indexCorreo].mensajeFaseII =
        "No hay institución asociada al curso. ";
      this.correos[indexCorreo].mensajeFaseIII =
        "No hay institución asociada al curso. ";
      this.correos[indexCorreo].mensajeFaseIVInstitucion =
        "No hay institución asociada al curso. ";
    } else if (!this.correos[indexCorreo].institucion.email1) {
      this.correos[indexCorreo].flagsEstados[0] = "error";
      this.correos[indexCorreo].flagsEstados[1] = "error";
      this.correos[indexCorreo].flagsEstados[2] = "error";
      this.correos[indexCorreo].flagsEstados[4] = "error";
      this.correos[indexCorreo].confirmacion_centro = "ERROR";
      this.correos[indexCorreo].confirmacion_material = "ERROR";
      this.correos[indexCorreo].confirmacion_recordatorio = "ERROR";
      this.correos[indexCorreo].confirmacion_graciasInstitucion = "ERROR";
      this.correos[indexCorreo].mensajeFaseI +=
        "No se encuentra el correo de la institución.";
      this.correos[indexCorreo].mensajeFaseII +=
        "No se encuentra el correo de la institución.";
      this.correos[indexCorreo].mensajeFaseIII +=
        "No se encuentra el correo de la institución.";
      this.correos[indexCorreo].mensajeFaseIVInstitucion +=
        "No se encuentra el correo de la institución.";
    }

    //Si falta el formador:
    if (this.correos[indexCorreo].formadores.length == 0) {
      this.correos[indexCorreo].institucion = {};
      this.correos[indexCorreo].flagsEstados[1] = "error";
      this.correos[indexCorreo].flagsEstados[2] = "error";
      this.correos[indexCorreo].flagsEstados[3] = "error";
      this.correos[indexCorreo].confirmacion_material = "ERROR";
      this.correos[indexCorreo].confirmacion_recordatorio = "ERROR";
      this.correos[indexCorreo].confirmacion_graciasFormador = "ERROR";
      this.correos[indexCorreo].mensajeFaseII =
        "No hay formadores asociados al curso.";
      this.correos[indexCorreo].mensajeFaseIII =
        "No hay formadores asociados al curso.";
      this.correos[indexCorreo].mensajeFaseIVFormador =
        "No hay formadores asociados al curso.";
    } else if (!this.correos[indexCorreo].formadores[0].email) {
      this.correos[indexCorreo].flagsEstados[1] = "error";
      this.correos[indexCorreo].flagsEstados[2] = "error";
      this.correos[indexCorreo].flagsEstados[3] = "error";
      this.correos[indexCorreo].confirmacion_material = "ERROR";
      this.correos[indexCorreo].confirmacion_recordatorio = "ERROR";
      this.correos[indexCorreo].confirmacion_graciasFormador = "ERROR";
      this.correos[indexCorreo].mensajeFaseII +=
        "El formador no tiene email asociado.";
      this.correos[indexCorreo].mensajeFaseIII +=
        "El formador no tiene email asociado.";
      this.correos[indexCorreo].mensajeFaseIVFormador +=
        "El formador no tiene email asociado.";
    }

    //Comprueba si hay sido Generado FASE I:
    if (this.correos[indexCorreo].flagsEstados[0] != "error") {
      if (this.correos[indexCorreo].idDraftInstitucion == undefined) {
        this.correos[indexCorreo].flagsEstados[0] = "warn";
        this.correos[indexCorreo].confirmacion_centro = "SIN GENERAR";
        this.correos[indexCorreo].mensajeFaseI = "Pendiente de ser generado.";
      } else {
        var flagEncontrado = false;
        for (var j = 0; j < this.listaBorradores.length; j++) {
          if (
            this.correos[indexCorreo].idDraftInstitucion ==
            this.listaBorradores[j]["threadId"]
          ) {
            flagEncontrado = true;
            break;
          }
        }
        if (!flagEncontrado) {
          this.correos[indexCorreo].flagsEstados[0] = "warn";
          this.correos[indexCorreo].confirmacion_centro = "NO ENVIADO";
          this.correos[indexCorreo].mensajeFaseI =
            "El correo se ha generado pero no se ha enviado. Puedes buscarlo en Gmail con la siguiente consulta: in:draft ([#" +
            this.correos[indexCorreo].cod_correo_institucion +
            "])";
        } else {
          this.correos[indexCorreo].confirmacion_centro = "OK";
          this.correos[indexCorreo].mensajeFaseI =
            "El correo se ha enviado correctamente. Puedes consultarlo con la siguiente busqueda [#" +
            this.correos[indexCorreo].cod_correo_institucion +
            "]";
        }
      }
    }

    //Comprueba si hay sido Generado FASE II:
    if (this.correos[indexCorreo].flagsEstados[1] != "error") {
      if (this.correos[indexCorreo].idDraftMaterial == undefined) {
        this.correos[indexCorreo].flagsEstados[1] = "warn";
        this.correos[indexCorreo].confirmacion_material = "SIN GENERAR";
        this.correos[indexCorreo].mensajeFaseII = "Pendiente de ser generado.";
      } else {
        var flagEncontrado = false;
        for (var j = 0; j < this.listaBorradores.length; j++) {
          if (
            this.correos[indexCorreo].idDraftMaterial ==
            this.listaBorradores[j]["threadId"]
          ) {
            flagEncontrado = true;
            break;
          }
        }
        if (!flagEncontrado) {
          this.correos[indexCorreo].flagsEstados[1] = "warn";
          this.correos[indexCorreo].confirmacion_material = "NO ENVIADO";
          this.correos[indexCorreo].mensajeFaseII =
            "El correo se ha generado pero no se ha enviado. Puedes buscarlo en Gmail con la siguiente consulta: in:draft ([#" +
            this.correos[indexCorreo].cod_correo_material +
            "])";
        } else {
          this.correos[indexCorreo].confirmacion_material = "OK";
          this.correos[indexCorreo].mensajeFaseII =
            "El correo se ha enviado correctamente. Puedes consultarlo con la siguiente busqueda [#" +
            this.correos[indexCorreo].cod_correo_material +
            "]";
        }
      }
    }

    //Comprueba si hay sido Generado FASE III:
    if (this.correos[indexCorreo].flagsEstados[2] != "error") {
      if (this.correos[indexCorreo].idDraftRecordatorio == undefined) {
        this.correos[indexCorreo].flagsEstados[2] = "warn";
        this.correos[indexCorreo].confirmacion_recordatorio = "SIN GENERAR";
        this.correos[indexCorreo].mensajeFaseIII = "Pendiente de ser generado.";
      } else {
        var flagEncontrado = false;
        for (var j = 0; j < this.listaBorradores.length; j++) {
          if (
            this.correos[indexCorreo].idDraftRecordatorio ==
            this.listaBorradores[j]["threadId"]
          ) {
            flagEncontrado = true;
            break;
          }
        }
        if (!flagEncontrado) {
          this.correos[indexCorreo].flagsEstados[2] = "warn";
          this.correos[indexCorreo].confirmacion_recordatorio = "NO ENVIADO";
          this.correos[indexCorreo].mensajeFaseIII =
            "El correo se ha generado pero no se ha enviado. Puedes buscarlo en Gmail con la siguiente consulta: in:draft ([#" +
            this.correos[indexCorreo].cod_correo_recordatorio +
            "])";
        } else {
          this.correos[indexCorreo].confirmacion_recordatorio = "OK";
          this.correos[indexCorreo].mensajeFaseIII =
            "El correo se ha enviado correctamente. Puedes consultarlo con la siguiente busqueda [#" +
            this.correos[indexCorreo].cod_correo_recordatorio +
            "]";
        }
      }

      var hoy: any = moment().startOf("day");
      var fechaCurso: any = moment(
        this.correos[indexCorreo]["fecha_formateada"],
      );
      var diasHastaCurso = Math.round(
        moment.duration(fechaCurso - hoy).asDays(),
      );
      if (diasHastaCurso > 10) {
        this.correos[indexCorreo].flagsEstados[2] = "PENDIENTE";
        this.correos[indexCorreo].confirmacion_recordatorio = "EN FECHA";
        this.correos[indexCorreo].mensajeFaseIII =
          "Quedan " + diasHastaCurso + " días para el curso.";
      }
    }

    //Comprueba si hay sido Generado FASE IV (Formador):
    if (this.correos[indexCorreo].flagsEstados[3] != "error") {
      var hoy: any = moment().startOf("day");
      var fechaCurso: any = moment(
        this.correos[indexCorreo]["fecha_formateada"],
      );
      var diasHastaCurso = Math.round(
        moment.duration(fechaCurso - hoy).asDays(),
      );

      //Si todavía no ha llegado la fecha del curso -> OK
      if (diasHastaCurso > 0) {
        this.correos[indexCorreo].flagsEstados[3] = "PENDIENTE";
        this.correos[indexCorreo].confirmacion_graciasFormador = "PENDIENTE";
        this.correos[indexCorreo].mensajeFaseIVFormador =
          "Quedan " + diasHastaCurso + " días para que se realize el curso";
      } else {
        //Si ya ha pasado la fecha del curso y no se ha generado el correo -> WARN
        if (this.correos[indexCorreo].idDraftGraciasFormador == undefined) {
          this.correos[indexCorreo].flagsEstados[3] = "warn";
          this.correos[indexCorreo].confirmacion_graciasFormador =
            "SIN GENERAR";
          this.correos[indexCorreo].mensajeFaseIVFormador =
            "Han pasado " +
            Math.abs(diasHastaCurso) +
            " días desde que se realizó el curso.";
        } else {
          var flagEncontrado = false;
          for (var j = 0; j < this.listaBorradores.length; j++) {
            if (
              this.correos[indexCorreo].idDraftGraciasFormador ==
              this.listaBorradores[j]["threadId"]
            ) {
              flagEncontrado = true;
              break;
            }
          }
          if (!flagEncontrado) {
            //Si ya ha pasado la fecha del curso y no se ha enviado el correo -> WARN
            this.correos[indexCorreo].flagsEstados[3] = "warn";
            this.correos[indexCorreo].confirmacion_graciasFormador =
              "NO ENVIADO";
            this.correos[indexCorreo].mensajeFaseIII =
              "El correo se ha generado pero no se ha enviado. Puedes buscarlo en Gmail con la siguiente consulta: in:draft ([#" +
              this.correos[indexCorreo].cod_correo_graciasFormador +
              "])";
          } else {
            //Si ya ha pasado la fecha del curso y se ha enviado el correo -> OK
            this.correos[indexCorreo].confirmacion_graciasFormador = "OK";
            this.correos[indexCorreo].mensajeFaseIII =
              "El correo se ha enviado correctamente. Puedes consultarlo con la siguiente busqueda [#" +
              this.correos[indexCorreo].cod_correo_graciasFormador +
              "]";
          }
        }
      }
    }

    //Comprueba si hay sido Generado FASE IV (Institucion):
    if (this.correos[indexCorreo].flagsEstados[4] != "error") {
      var hoy: any = moment().startOf("day");
      var fechaCurso: any = moment(
        this.correos[indexCorreo]["fecha_formateada"],
      );
      var diasHastaCurso = Math.round(
        moment.duration(fechaCurso - hoy).asDays(),
      );

      //Si todavía no ha llegado la fecha del curso -> OK
      if (diasHastaCurso > 0) {
        this.correos[indexCorreo].flagsEstados[4] = "PENDIENTE";
        this.correos[indexCorreo].confirmacion_graciasInstitucion = "PENDIENTE";
        this.correos[indexCorreo].mensajeFaseIVInstitucion =
          "Quedan " + diasHastaCurso + " días para que se realize el curso";
      } else {
        //Si ya ha pasado la fecha del curso y no se ha generado el correo -> WARN
        if (this.correos[indexCorreo].idDraftGraciasInstitucion == undefined) {
          this.correos[indexCorreo].flagsEstados[4] = "warn";
          this.correos[indexCorreo].confirmacion_graciasInstitucion =
            "SIN GENERAR";
          this.correos[indexCorreo].mensajeFaseIVInstitucion =
            "Han pasado " +
            Math.abs(diasHastaCurso) +
            " días desde que se realizó el curso.";
        } else {
          var flagEncontrado = false;
          for (var j = 0; j < this.listaBorradores.length; j++) {
            if (
              this.correos[indexCorreo].idDraftGraciasInstitucion ==
              this.listaBorradores[j]["threadId"]
            ) {
              flagEncontrado = true;
              break;
            }
          }
          if (!flagEncontrado) {
            //Si ya ha pasado la fecha del curso y no se ha enviado el correo -> WARN
            this.correos[indexCorreo].flagsEstados[4] = "warn";
            this.correos[indexCorreo].confirmacion_graciasInstitucion =
              "NO ENVIADO";
            this.correos[indexCorreo].mensajeFaseIII =
              "El correo se ha generado pero no se ha enviado. Puedes buscarlo en Gmail con la siguiente consulta: in:draft ([#" +
              this.correos[indexCorreo].cod_correo_graciasInstitucion +
              "])";
          } else {
            //Si ya ha pasado la fecha del curso y se ha enviado el correo -> OK
            this.correos[indexCorreo].confirmacion_graciasInstitucion = "OK";
            this.correos[indexCorreo].mensajeFaseIII =
              "El correo se ha enviado correctamente. Puedes consultarlo con la siguiente busqueda [#" +
              this.correos[indexCorreo].cod_correo_graciasInstitucion +
              "]";
          }
        }
      }
    }

    //CHECK FORZADO:
    if (this.correos[indexCorreo]["forzarCorreoInstitucion"]) {
      this.correos[indexCorreo].flagsEstados[0] = "OK";
    }

    this.correos[indexCorreo]["flagEstadoGeneral"] = "OK";

    if (this.correos[indexCorreo].flagsEstados.indexOf("PENDIENTE") != -1) {
      this.correos[indexCorreo]["flagEstadoGeneral"] = "PENDIENTE";
    }
    if (this.correos[indexCorreo].flagsEstados.indexOf("warn") != -1) {
      this.correos[indexCorreo]["flagEstadoGeneral"] = "warn";
    }
    if (this.correos[indexCorreo].flagsEstados.indexOf("error") != -1) {
      this.correos[indexCorreo]["flagEstadoGeneral"] = "error";
    }
  } //FIN checkEstadoCorreo

  forzarCorreoOK(cod_curso, tipo: string) {
    const dialogoProcesando = this.dialog.open(DialogoComponent, {
      disableClose: true,
      data: {
        tipoDialogo: "procesando",
        titulo: "Forzando OK Correo",
        contenido: "",
      },
    });

    //Buscar codigo Curso:
    var indexMetadatosCurso = this.binarySearchObject(
      this.metadatosCursos[0].data,
      "cod_curso",
      cod_curso,
    );
    var indexCorreo = this.binarySearchObject(
      this.correos,
      "cod_curso",
      cod_curso,
    );

    if (indexMetadatosCurso < 0) {
      dialogoProcesando.close();
      return false;
    }

    switch (tipo) {
      case "institucion":
        this.metadatosCursos[0].data[indexMetadatosCurso][
          "forzarCorreoInstitucion"
        ] =
          !this.metadatosCursos[0].data[indexMetadatosCurso][
            "forzarCorreoInstitucion"
          ];
        break;
    }

    if (indexCorreo >= 0) {
      this.checkEstadoCorreo(indexCorreo);
    }

    this.appService.guardarArchivo(this.metadatosCursos[0]).then((result) => {
      dialogoProcesando.close();
    });
  }

  pedirDatosGeneracionDocumento() {
    const dialogoProcesando = this.dialog.open(DialogoComponent, {
      disableClose: true,
      data: {
        tipoDialogo: "filtroGenerarDocumento",
        titulo: "Procesando",
        contenido: "",
      },
    });

    dialogoProcesando.afterClosed().subscribe((result) => {
      console.log(result);
      if (!result) {
        return false;
      }
      this.generarDocumentoComunidades(result);
    });
  }

  generarDocumentoComunidades(opciones) {
    console.warn("Generando Documento...");
    console.warn(opciones);

    var fechaInicio = moment(opciones.filtroFecha.fechaInicio);
    var fechaFin = moment(opciones.filtroFecha.fechaFin);

    var objetoDocumento = {};

    var cursosFiltrados = [];

    //Filtrado del periodo de fechas:
    var fecha = moment();
    for (var i = 0; i < this.cursos.length; i++) {
      fecha = moment(this.cursos[i].fecha_formateada);
      if (
        (fecha.isBefore(fechaFin) &&
          fecha.isAfter(fechaInicio) &&
          this.cursos[i].estado == "REALIZADA") ||
        fecha.isSame(fechaInicio) ||
        fecha.isSame(fechaFin)
      ) {
        cursosFiltrados.push(this.cursos[i]);
      }
    }

    console.warn("Cursos Filtrados:", cursosFiltrados);

    //Creación de parametros:
    var tipoDatos = ["Nses", "Nform", "Htotal", "Himpacto", "Nbene", "Percent"];

    //Creación por provincia y por comunidad autónoma:
    for (var i = 0; i < this.codigoProvincia[0].data.length; i++) {
      for (var j = 0; j < tipoDatos.length; j++) {
        objetoDocumento[
          tipoDatos[j] + "_" + this.codigoProvincia[0].data[i].cod__provincia
        ] = 0;
        if (String(this.codigoProvincia[0].data[i].ccaa.slice(0, 2)) == "C.") {
          objetoDocumento[tipoDatos[j] + "_VA"] = 0;
        } else {
          objetoDocumento[
            tipoDatos[j] +
              "_" +
              String(this.codigoProvincia[0].data[i].ccaa).slice(0, 2)
          ] = 0;
        }
      }
    }

    //Creación General:
    for (var j = 0; j < tipoDatos.length; j++) {
      objetoDocumento[tipoDatos[j] + "_ES"] = 0;
      objetoDocumento[tipoDatos[j] + "_ES"] = 0;
    }

    //Calculo Secuencial NSes:
    var duracionCurso;
    var codGrupoAnterior = 0;
    var cod_provincia = "";
    cursosFiltrados.sort(
      (a, b) => Number(a["cod_grupo"]) - Number(b["cod_grupo"]),
    );
    for (var i = 0; i < cursosFiltrados.length; i++) {
      duracionCurso =
        (cursosFiltrados[i].hora_fin - cursosFiltrados[i].hora_inicio) * 24;
      cod_provincia = String(cursosFiltrados[i]["cod__postal"]).slice(0, 2);

      if (isNaN(duracionCurso)) {
        duracionCurso = 1;
      }

      //ESPAÑA
      objetoDocumento["Nses_ES"]++;
      objetoDocumento["Htotal_ES"] += duracionCurso;
      objetoDocumento["Himpacto_ES"] +=
        duracionCurso * cursosFiltrados[i]["nºasistentes"];
      if (codGrupoAnterior != cursosFiltrados[i].cod_grupo) {
        objetoDocumento["Nbene_ES"] += cursosFiltrados[i]["nºasistentes"];
      }

      //PROVINCIA:
      objetoDocumento["Nses_" + cod_provincia]++;
      objetoDocumento["Htotal_" + cod_provincia] += duracionCurso;
      objetoDocumento["Himpacto_" + cod_provincia] +=
        duracionCurso * cursosFiltrados[i]["nºasistentes"];
      if (codGrupoAnterior != cursosFiltrados[i].cod_grupo) {
        objetoDocumento["Nbene_impacto" + cod_provincia] +=
          cursosFiltrados[i]["nºasistentes"];
      }

      //CCAA:
      if (String(cursosFiltrados[i]["ccaa_/_pais"]).slice(0, 2) == "C.") {
        objetoDocumento["Nses_VA"]++;
        objetoDocumento["Htotal_VA"] += duracionCurso;
        objetoDocumento["Himpacto_VA"] +=
          duracionCurso * cursosFiltrados[i]["nºasistentes"];
        if (codGrupoAnterior != cursosFiltrados[i].cod_grupo) {
          objetoDocumento["Nbene_VA"] += cursosFiltrados[i]["nºasistentes"];
        }
      } else {
        objetoDocumento[
          "Nses_" + String(cursosFiltrados[i]["ccaa_/_pais"]).slice(0, 2)
        ]++;
        objetoDocumento[
          "Htotal_" + String(cursosFiltrados[i]["ccaa_/_pais"]).slice(0, 2)
        ] += duracionCurso;
        objetoDocumento[
          "Himpacto_" + String(cursosFiltrados[i]["ccaa_/_pais"]).slice(0, 2)
        ] += duracionCurso * cursosFiltrados[i]["nºasistentes"];
        if (codGrupoAnterior != cursosFiltrados[i].cod_grupo) {
          objetoDocumento[
            "Nbene_" + String(cursosFiltrados[i]["ccaa_/_pais"]).slice(0, 2)
          ] += cursosFiltrados[i]["nºasistentes"];
        }
      }

      codGrupoAnterior = cursosFiltrados[i].cod_grupo;
    } //Fin calculo secuencial

    //Generación de calculo formadores:
    var objetoFormadores = [];
    var indexBusqueda = 0;
    var iteracionBusqueda = 0;
    var flagEncontrado = false;

    this.formadoresCurso[0].data.sort(
      (a, b) => Number(a["cod_curso"]) - Number(b["cod_curso"]),
    );
    cursosFiltrados.sort(
      (a, b) => Number(a["cod_curso"]) - Number(b["cod_curso"]),
    );
    for (var i = 0; i < cursosFiltrados.length; i++) {
      flagEncontrado = false;
      while (iteracionBusqueda < this.formadoresCurso[0].data.length + 1) {
        if (this.formadoresCurso[0].data[indexBusqueda] == undefined) {
          console.error("Error en la busqueda de formadores");
          console.error("IndexBusqueda:", indexBusqueda);
          console.error("Formadores:", this.formadoresCurso[0].data);
        }

        if (
          Number(this.formadoresCurso[0].data[indexBusqueda]["cod_curso"]) ==
          Number(cursosFiltrados[i].cod_curso)
        ) {
          //Encontrado
          flagEncontrado = true;

          objetoFormadores.push({
            idFormador:
              this.formadoresCurso[0].data[indexBusqueda]["cod__formador"],
            codCurso: cursosFiltrados[i].cod_curso,
            cod_postal: cursosFiltrados[i].cod__postal,
            ccaa: cursosFiltrados[i]["ccaa_/_pais"],
          });

          indexBusqueda++;
          if (indexBusqueda >= this.formadoresCurso[0].data.length) {
            indexBusqueda = 0;
          }
        } else if (flagEncontrado) {
          break;
        } else {
          indexBusqueda++;
          if (indexBusqueda >= this.formadoresCurso[0].data.length) {
            indexBusqueda = 0;
          }
          iteracionBusqueda++;
        }
      } //Fin bucle busqueda
    } //Fin generación de calculo formadores

    //Calculo de porcentajes:
    for (var i = 0; i < cursosFiltrados.length; i++) {
      //CCAA y Provincia:
      if (String(cursosFiltrados[i]["ccaa_/_pais"]).slice(0, 2) == "C.") {
        objetoDocumento["Percent_VA"] =
          objetoDocumento["Nses_VA"] / objetoDocumento["Nses_ES"];
        objetoDocumento[
          "Percent_" + String(cursosFiltrados[i].cod__provincia)
        ] =
          (objetoDocumento[
            "Nses_" + String(cursosFiltrados[i].cod__provincia)
          ] /
            objetoDocumento["Nses_VA"]) *
          100;
      } else {
        objetoDocumento[
          "Percent_" + String(cursosFiltrados[i]["ccaa_/_pais"]).slice(0, 2)
        ] =
          (objetoDocumento[
            "Nses_" + String(cursosFiltrados[i]["ccaa_/_pais"]).slice(0, 2)
          ] /
            objetoDocumento["Nses_ES"]) *
          100;
        objetoDocumento[
          "Percent_" + String(cursosFiltrados[i].cod__provincia)
        ] =
          (objetoDocumento[
            "Nses_" + String(cursosFiltrados[i].cod__provincia)
          ] /
            objetoDocumento[
              "Nses_" + String(cursosFiltrados[i]["ccaa_/_pais"]).slice(0, 2)
            ]) *
          100;
      }
    } //Fin calculo porcentajes

    console.warn("Formadores-Curso:", this.formadoresCurso[0].data);

    //var indexFormador = this.binarySearchObject(this.formadoresCurso[0].data[j], "cod__formador", this.formadoresCurso[0].data[j]["cod__formador"])

    //Ordenación de formadores idFormador --> CodPostal:
    objetoFormadores.sort(
      (a, b) =>
        Number(a["idFormador"]) * 10000 +
        Number(a["cod_postal"]) -
        (Number(b["idFormador"]) * 10000 + Number(b["cod_postal"])),
    );

    console.warn("Objeto Formadores", objetoFormadores);

    var lastFormador = 0;
    var lastCCAA = "00";
    var lastProvincia = "00";
    var arrayformadores = [];
    for (var i = 0; i < objetoFormadores.length; i++) {
      if (!Number(objetoFormadores[i]["idFormador"])) {
        continue;
      }

      if (lastFormador != objetoFormadores[i]["idFormador"]) {
        objetoDocumento["Nform_ES"]++;
      }

      if (
        lastProvincia !=
          String(objetoFormadores[i]["cod_postal"]).slice(0, 2) ||
        lastFormador != objetoFormadores[i]["idFormador"]
      ) {
        objetoDocumento[
          "Nform_" + String(objetoFormadores[i]["cod_postal"]).slice(0, 2)
        ]++;
        if (String(objetoFormadores[i]["cod_postal"]).slice(0, 2) == "33") {
          arrayformadores.push(objetoFormadores[i]["idFormador"]);
        }
      }

      lastFormador = objetoFormadores[i]["idFormador"];
      lastProvincia = String(objetoFormadores[i]["cod_postal"]).slice(0, 2);
    } //Fin calculo secuencial

    console.log("array formadores: ", arrayformadores);
    objetoFormadores.sort(
      (a, b) => Number(a["idFormador"]) - Number(b["idFormador"]),
    );
    objetoFormadores.sort((a, b) => Number(a["ccaa"]) - Number(b["ccaa"]));

    console.warn("Objeto Formadores 2", objetoFormadores);
    for (var i = 0; i < objetoFormadores.length; i++) {
      if (!Number(objetoFormadores[i]["idFormador"])) {
        continue;
      }

      if (lastCCAA != String(objetoFormadores[i]["ccaa"]).slice(0, 2)) {
        if (String(objetoFormadores[i]["ccaa"]).slice(0, 2) == "C.") {
          objetoDocumento["Nform_VA"]++;
        } else {
          objetoDocumento[
            "Nform_" + String(objetoFormadores[i]["ccaa"]).slice(0, 2)
          ]++;
        }
      }
      lastCCAA = String(objetoFormadores[i]["ccaa"]).slice(0, 2);
    }

    //Formateo de numeros:
    for (const [key, value] of Object.entries(objetoDocumento)) {
      //objetoDocumento[key] = Math.round(Number(value));
      if (key.slice(7) == "Percent") {
        objetoDocumento[key] = Math.round(Number(value) * 10) / 10;
      } else {
        objetoDocumento[key] = Math.round(Number(value))
          .toString()
          .replace(/\B(?=(\d{3})+(?!\d))/g, ".");
      }
    }

    console.warn("Cursos Filtrados:", cursosFiltrados);
    console.warn("ObjetoDocumento", objetoDocumento);

    objetoDocumento["Titulo_Periodo"] = opciones.tituloDialogo;
    objetoDocumento["Periodo"] =
      "De " +
      moment(opciones.filtroFecha.fechaInicio).format("DD/MM/YYYY") +
      " a " +
      moment(opciones.filtroFecha.fechaFin).format("DD/MM/YYYY");

    var argumentos = [
      opciones.valorInput[0].path,
      opciones.pathOutput,
      objetoDocumento,
    ];
    var parametrosProceso = {
      tipo: "proceso",
      nombre: "Generar Documento",
      categoria: "Documentos",
      argumentos: argumentos,
    };

    this.appService.ejecutarProceso(parametrosProceso, argumentos);
  }
}
