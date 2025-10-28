import { Component, OnInit, ViewChild, ViewEncapsulation } from "@angular/core";
import { AppService } from "../../app.service";
import { DashboardService } from "./dashboard.service";
import { MatDialog } from "@angular/material/dialog";
import { DialogoComponent } from "../dialogos/dialogos.component";
import { AddDatoComponent } from "../addDato/addDato.component";
import { AddPlantillaComponent } from "../addPlantilla/addPlantilla.component";
import { EjecutarProcesoComponent } from "../ejecutarProceso/ejecutarProceso.component";
import { VisualizarDato } from "../visualizarDato/visualizarDato.component";
import { GestionarDato } from "../gestionarDato/gestionarDato.component";
import { MatSidenav } from "@angular/material/sidenav";
import { FlatTreeControl } from "@angular/cdk/tree";
import {
  MatTreeFlatDataSource,
  MatTreeFlattener,
} from "@angular/material/tree";
import { InsertarElementoComponent } from "../insertarElemento/insertarElemento.component";
import { AddCursoComponent } from "../addCurso/addCurso.component";
import { EditorDocumentoComponent } from "../editorDocumento/editorDocumento.component";

//Modulos:
import { MatTableModule } from "@angular/material/table";
import { MatIconModule } from "@angular/material/icon";
import { MatDialogModule } from "@angular/material/dialog";
import { CommonModule } from "@angular/common";
import { MatButtonToggleModule } from "@angular/material/button-toggle";
import { MatSidenavModule } from "@angular/material/sidenav";
import { MatTreeModule } from "@angular/material/tree";
import { MatMenuModule } from "@angular/material/menu";
import { MatTabsModule } from "@angular/material/tabs";
import { MatButtonModule } from "@angular/material/button";

import { environment } from "src/environments/environment";

interface ArchivoNode {
  nombre: string;
  tipo: string;
  direccion: string;
  subDirectorio?: ArchivoNode[];
}

interface ExpansibleNode {
  expandable: boolean;
  nombre: string;
  level: number;
}

type EstadoDashboard = "Normal" | "Borrado";

@Component({
  standalone: true,
  imports: [
    MatButtonModule,
    MatMenuModule,
    MatTabsModule,
    MatButtonToggleModule,
    MatSidenavModule,
    MatTreeModule,
    CommonModule,
    MatIconModule,
    MatTableModule,
    MatDialogModule,
    EditorDocumentoComponent,
    AddCursoComponent,
  ],
  encapsulation: ViewEncapsulation.None,
  selector: "app-dashboard",
  templateUrl: "./dashboard.component.html",
  styleUrls: ["./dashboard.component.sass"],
})
export class DashboardComponent implements OnInit {
  constructor(
    public appService: AppService,
    public dashboardService: DashboardService,
    public dialog: MatDialog,
  ) {
    this.arbolArchivosDataSource.data = this.datosArbolArchivos;
    this.elementosRender = {};
  }

  public selectedTab = 2;
  public proyectoActivo = false;
  public nombreProyectoActivo = "Selecciona un proyecto";
  public datosProyecto: any = {};
  public listaProyectos: any = [];
  private estadoDashboard: EstadoDashboard = "Normal";

  //Menu:
  public mostrarMenu: boolean = false;
  public pantallaSeleccionada: string = "Cursos";
  public claseAddCursoComponent: string = "visible";
  public claseEditorCorreoComponent: string = "oculto";
  public clasePowerBi: string = "visible";

  //Elementos Render:
  public claseElementos: string = "Normal";
  public elementosRender: any = {};

  @ViewChild("drawer", { static: false }) drawer: MatSidenav;
  @ViewChild("AddCursoComponent", { static: false })
  addCursoComponent: AddCursoComponent;

  //Arbol de archivosArchivoNode
  private _transformer = (node: ArchivoNode, level: number) => {
    return {
      expandable: !!node.subDirectorio && node.subDirectorio.length > 0,
      nombre: node.nombre,
      tipo: node.tipo,
      level: level,
    };
  };

  public arbolArchivosControl = new FlatTreeControl<ExpansibleNode>(
    (node) => node.level,
    (node) => node.expandable,
  );
  private reductorArbolArchivos = new MatTreeFlattener(
    this._transformer,
    (node) => node.level,
    (node) => node.expandable,
    (node) => node.subDirectorio,
  );
  public arbolArchivosDataSource = new MatTreeFlatDataSource(
    this.arbolArchivosControl,
    this.reductorArbolArchivos,
  );
  public hasChild = (_: number, node: ExpansibleNode) => node.expandable;

  public datosArbolArchivos: ArchivoNode[] = [];

  ngOnInit() {
    //Comprueba que se ha seleccionado un proyecto:
    this.appService.getProyecto();

    if (!this.appService.proyectoActivo) {
      console.log("SALIENDO DE DASHBOARD");
      this.appService.cambiarUrl("/index");
    } else {
      this.nombreProyectoActivo = this.appService.proyectoConfig.nombre;
      this.proyectoActivo = true;
    }

    //Inicializar Arbol de proyecto:
    this.appService.getArbolProyecto().then((result: ArchivoNode[]) => {
      console.log("Promesa:");
      console.log(result);

      this.arbolArchivosDataSource.data = result;
    });

    //Pantalla por defecto:
    this.pantallaSeleccionada = "Monitorizacion";
    this.clasePowerBi = "visible";
    this.claseAddCursoComponent = "oculto";
    this.claseEditorCorreoComponent = "oculto";

    //Inicializar Render Elementos:
    this.elementosRender = this.appService.obtenerElementosRender();

    console.log("ELEMENTOS: ");
    console.log(this.elementosRender);

    this.cambioPestana(this.selectedTab);
  }

  reloadArbolProyecto() {
    //Reload del arbol de proyecto:
    this.appService.getArbolProyecto().then((result: ArchivoNode[]) => {
      console.log("Promesa:");
      console.log(result);
      this.arbolArchivosDataSource.data = result;
    });
  }

  cambioPestana(event: any) {
    switch (event) {
      case 0:
        this.pantallaSeleccionada = "Monitorizacion";
        this.clasePowerBi = "visible";
        this.claseAddCursoComponent = "oculto";
        this.claseEditorCorreoComponent = "oculto";
        break;
      case 1:
        this.pantallaSeleccionada = "Automatizacion";
        this.clasePowerBi = "oculto";
        this.claseAddCursoComponent = "visible";
        this.claseEditorCorreoComponent = "oculto";
        break;
      case 2:
        this.pantallaSeleccionada = "Cursos";
        this.clasePowerBi = "oculto";
        this.claseAddCursoComponent = "visible";
        this.claseEditorCorreoComponent = "oculto";
        break;
      case 3:
        this.pantallaSeleccionada = "Formadores";
        this.clasePowerBi = "oculto";
        this.claseAddCursoComponent = "visible";
        this.claseEditorCorreoComponent = "oculto";
        break;
      case 4:
        this.pantallaSeleccionada = "Instituciones";
        this.clasePowerBi = "oculto";
        this.claseAddCursoComponent = "visible";
        this.claseEditorCorreoComponent = "oculto";
        break;
    }
    console.log("Cambio de pestaña: " + this.pantallaSeleccionada);
    return;
  }

  //Funciones de herramientas:
  addDato(opcionesArg: any) {
    const dialogRef = this.dialog.open(AddDatoComponent, {
      disableClose: true,
      width: "50%",
      data: {
        opciones: opcionesArg,
        titulo: "Añadir dato",
        contenido: "Aqui se añaden datos.",
      },
    });

    dialogRef.afterClosed().subscribe((result) => {
      if (result == "exito") {
        const dialogExito = this.dialog.open(DialogoComponent, {
          data: {
            tipoDialogo: "exito",
            titulo: "Archivo guardado con exito",
            contenido: "El archivo se ha guardado con exito.",
          },
        });
      }

      if (result == "error") {
        const dialogExito = this.dialog.open(DialogoComponent, {
          data: {
            tipoDialogo: "error",
            titulo: "Se ha producido un error.",
            contenido:
              "No se ha podido guardar el archivo debido a un error inesperado.",
          },
        });
      }

      console.log("Fin de Herramienta AddDato: " + result);
      return;
    });
  }

  addPlantilla(opcionesArg: any) {
    const dialogRef = this.dialog.open(AddPlantillaComponent, {
      disableClose: true,
      width: "50%",
      data: {
        opciones: opcionesArg,
        titulo: "Añadir dato",
        contenido: "Aqui se añaden datos.",
      },
    });

    dialogRef.afterClosed().subscribe((result) => {
      if (result == "exito") {
        const dialogExito = this.dialog.open(DialogoComponent, {
          data: {
            tipoDialogo: "exito",
            titulo: "Archivo guardado con exito",
            contenido: "El archivo se ha guardado con exito.",
          },
        });
      }

      if (result == "error") {
        const dialogExito = this.dialog.open(DialogoComponent, {
          data: {
            tipoDialogo: "error",
            titulo: "Se ha producido un error.",
            contenido:
              "No se ha podido guardar el archivo debido a un error inesperado.",
          },
        });
      }

      console.log("Fin de Herramienta AddDato: " + result);
      return;
    });
  }

  //Gestionar Dato:
  gestionarDato() {
    const dialogRef = this.dialog.open(GestionarDato, {
      disableClose: true,
      width: "70%",
      data: {
        opciones: {},
        titulo: "Gestionar datos",
        contenido: "Panel de gestion de datos",
      },
    });

    dialogRef.afterClosed().subscribe((result) => {
      if (result == "error") {
        const dialogExito = this.dialog.open(DialogoComponent, {
          data: {
            tipoDialogo: "error",
            titulo: "Se ha producido un error.",
            contenido: "Error desconocido.",
          },
        });
      }

      console.log("Fin Gestion de datos: " + result);
      return;
    });
  }

  //Visualizar Dato:
  visualizarDato() {
    const dialogRef = this.dialog.open(VisualizarDato, {
      disableClose: true,
      width: "70%",
      data: {
        opciones: {},
        titulo: "Visualizador de datos",
        contenido: "Panel de visualización de datos",
      },
    });

    dialogRef.afterClosed().subscribe((result) => {
      if (result == "error") {
        const dialogExito = this.dialog.open(DialogoComponent, {
          data: {
            tipoDialogo: "error",
            titulo: "Se ha producido un error.",
            contenido: "Error desconocido.",
          },
        });
      }

      console.log("Fin visulazardor de datos: " + result);
      return;
    });
  }

  //Funciones de herramientas:
  insertarElemento(opcionesArg: any) {
    //Añadir estado Drawer
    opcionesArg.estadoDrawer = this.drawer.opened;

    var configuracionDialogo = {};
    switch (opcionesArg.tipo) {
      case "texto":
        configuracionDialogo = {
          disableClose: true,
          width: "50%",
          data: {
            opciones: opcionesArg,
            titulo: "Insertar Elemento",
            contenido: this.elementosRender,
          },
        };
        break;
      case "tabla":
        configuracionDialogo = {
          disableClose: true,
          width: "90%",
          data: {
            opciones: opcionesArg,
            titulo: "Insertar Elemento",
            contenido: this.elementosRender,
          },
        };
        break;

      case "externo":
        configuracionDialogo = {
          disableClose: true,
          width: "90%",
          data: {
            opciones: opcionesArg,
            titulo: "Insertar Elemento",
            contenido: this.elementosRender,
          },
        };
        break;
    }

    const dialogRef = this.dialog.open(
      InsertarElementoComponent,
      configuracionDialogo,
    );

    dialogRef.afterClosed().subscribe((result) => {
      if (result == "exito") {
        const dialogExito = this.dialog.open(DialogoComponent, {
          data: {
            tipoDialogo: "exito",
            titulo: "Archivo guardado con exito",
            contenido: "El archivo se ha guardado con exito.",
          },
        });
      }

      if (result == "error") {
        const dialogExito = this.dialog.open(DialogoComponent, {
          data: {
            tipoDialogo: "error",
            titulo: "Se ha producido un error.",
            contenido: "No se ha podido insertar el elemento.",
          },
        });
      }

      console.log("Fin de Herramienta InsertarElemento: " + result);
      return;
    });
  }

  activarBorradoElementos() {
    this.claseElementos = "borradoActivado";
    this.estadoDashboard = "Borrado";
  }

  clickElemento(index: number) {
    switch (this.estadoDashboard) {
      case "Normal":
        break;
      case "Borrado":
        if (index > -1) {
          this.elementosRender["elementos"].splice(index, 1);
          this.appService.guardarDocumentoElementos(this.elementosRender);
          this.claseElementos = "Normal";
          this.estadoDashboard = "Normal";
        }
        break;
    }
  }

  ejecutarProceso(opcionesArg: any) {
    const dialogRef = this.dialog.open(EjecutarProcesoComponent, {
      disableClose: true,
      width: "70%",
      data: {
        opciones: opcionesArg,
        titulo: "Añadir dato",
        contenido: "Aqui se añaden datos.",
      },
    });

    dialogRef.afterClosed().subscribe((result) => {
      if (result == "exito") {
        const dialogExito = this.dialog.open(DialogoComponent, {
          data: {
            tipoDialogo: "exito",
            titulo: "Archivo guardado con exito",
            contenido: "El archivo se ha guardado con exito.",
          },
        });
      }

      if (result == "error") {
        const dialogExito = this.dialog.open(DialogoComponent, {
          data: {
            tipoDialogo: "error",
            titulo: "Se ha producido un error.",
            contenido:
              "No se ha podido guardar el archivo debido a un error inesperado.",
          },
        });
      }

      console.log("Fin de Herramienta AddDato: " + result);
      return;
    });
  }

  analizarSpool() {
    this.appService.analizarSpool().then((result) => {
      console.log("Analisis Spool:");
      console.log(result);
    });
  }

  listarProyectos() {
    this.listaProyectos = this.appService.listarProyectos();

    for (var i = 0; i < this.listaProyectos.length; ++i) {
      //this.listaProyectos[i]= this.listaProyectos[i].replace(/.db/gi,"")
      //this.listaProyectos[i]= this.listaProyectos[i].nombreProyecto;
    }
    return;
  }

  seleccionar(herramienta: string): void {
    this.appService.cambiarUrl("/" + herramienta);
  }

  cargarProyecto(nombreProyecto) {
    this.datosProyecto = this.appService.abrirProyecto(nombreProyecto);
    console.log(this.datosProyecto);
    if (this.datosProyecto) {
      this.proyectoActivo = true;
      this.nombreProyectoActivo = nombreProyecto;
      this.drawer.open();
    } else {
      console.log("Se ha producido un error en la carga del Proyecto");
      this.nombreProyectoActivo = "Selecciona un proyecto";
      this.proyectoActivo = false;
    }
    return;
  }

  cerrarProyecto() {
    this.appService.cerrarProyecto();
    this.nombreProyectoActivo = "Selecciona un proyecto";
    this.proyectoActivo = false;
    this.appService.cambiarUrl("/index");
    return;
  }

  eliminarProyecto(nombreProyecto) {
    const dialogRef = this.dialog.open(DialogoComponent, {
      data: {
        tipoDialogo: "confirmacion",
        titulo:
          "¿Seguro que quiere eliminar el proyecto '" + nombreProyecto + "'?",
        contenido:
          "Si elimina el proyecto se eliminarán todos los archivos importados en el mismo.",
      },
    });

    dialogRef.afterClosed().subscribe((result) => {
      console.log("Fin del dialogo: " + result);

      if (result === true) {
        this.cerrarProyecto();
        this.appService.eliminarProyecto(nombreProyecto);
        this.listarProyectos();
      } else {
        return;
      }
    });
    return;
  }

  crearProyecto(nombreProyecto) {
    console.log("Creando Proyecto");

    const dialogRef = this.dialog.open(DialogoComponent, {
      data: { tipoDialogo: "crearProyecto", data: {} },
    });

    dialogRef.afterClosed().subscribe((result) => {
      console.log("Fin del dialogo");
      console.log(result);
      if (result === undefined) {
        return;
      }

      if (result.nombre === undefined || result.nombre === null) {
        return;
      }

      this.datosProyecto = this.appService.crearProyecto(result);

      if (this.datosProyecto) {
        this.proyectoActivo = true;
        this.nombreProyectoActivo = nombreProyecto;
        this.drawer.open();
      } else {
        console.log("Se ha producido un error en la carga del Proyecto");
        this.nombreProyectoActivo = "Selecciona un proyecto";
        this.proyectoActivo = false;
      }
    });
    return;
  }

  openDialog(tipoDialogoArg: string, dataArg: any) {
    const dialogRef = this.dialog.open(DialogoComponent, {
      data: { tipoDialogo: tipoDialogoArg, data: dataArg },
    });

    dialogRef.afterClosed().subscribe((result) => {
      console.log("Fin del dialogo");
      console.log(result);
    });
    return;
  }

  comandoAddCurso(comando: string) {
    this.addCursoComponent.comandoHerramienta(comando);
  }

  abrirSoporte(): void {
    const w = window.open(environment.URL_SOPORTE, "_blank");
    if (w) {
      w.opener = null;
    }
  }
}
