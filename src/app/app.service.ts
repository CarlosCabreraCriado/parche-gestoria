import { Injectable, EventEmitter, Output } from "@angular/core";
import { Router, ActivatedRoute } from "@angular/router";
import { Subject, Observable } from "rxjs";
import { HttpClient } from "@angular/common/http";
import { DialogoComponent } from "./comun/dialogos/dialogos.component";
import { MatDialog } from "@angular/material/dialog";

@Injectable({
  providedIn: "root",
})
export class AppService {
  constructor(
    private route: ActivatedRoute,
    private router: Router,
    private http: HttpClient,
    public dialog: MatDialog,
  ) {}

  //MODO DEBUG:
  public debug: boolean = false;

  //public ipRemota: string= "http://www.carloscabreracriado.com";
  public ipRemota: string = "http://127.0.0.1:8000";

  // Observable string sources
  private observarAppService = new Subject<string>();

  // Observable string streams
  observarAppService$ = this.observarAppService.asObservable();

  //Variables de Proyecto:
  public proyectoActivo = false;
  public nombreProyectoActivo = "Selecciona un proyecto";
  public datosProyecto: any = {};
  public listaProyectos: any = [];
  public proyectoConfig: any = {};
  public parametrosProyecto: any = {};

  //Variables Globales:
  public version: string = "0.9.93";

  cambiarUrl(url: string): void {
    console.log("CAMBIANDO A URL: " + url);
    this.router.navigateByUrl(url);
  }

  async inicializarAppService() {
    //Obtener Proyecto si esta cargado en electron:
    this.proyectoConfig = await window.electronAPI.getProyecto();
    console.warn("Proyecto Config: ", this.proyectoConfig);

    if (this.proyectoConfig !== null || this.proyectoConfig !== undefined) {
      this.proyectoActivo = true;
    }

    this.cargarParametros();

    window.electronAPI.on("onAbrirModo", (evnt, modo, data) => {
      console.log("Abriendo Modo: ");
      console.log(modo);
      this.dialog.closeAll();
      this.openDialog("informativo", {
        tipoDialogo: "informativo",
        titulo: "Abriendo Importador Mail",
        contenido:
          "El servicio de importación se abrirá en una ventana emergente",
      });
    });

    window.electronAPI.on("onErrorInterno", (evnt, tituloError, err) => {
      console.log("Se ha producido un error interno: ");
      console.log(err);
      this.dialog.closeAll();
      this.openDialog("error", {
        tipoDialogo: "error",
        titulo: tituloError,
        contenido: err,
      });
    });

    window.electronAPI.on("onMostrarError", (evnt, tituloError, err) => {
      console.log("Se ha producido un error interno: ");
      console.log(err);
      this.dialog.closeAll();
      this.openDialog("error", {
        tipoDialogo: "error",
        titulo: tituloError,
        contenido: err,
      });
    });

    window.electronAPI.on(
      "onMostrarWarning",
      (evnt, tituloWarning, warning) => {
        console.log("Se ha producido un error interno: ");
        console.log(warning);
        this.dialog.closeAll();
        this.openDialog("warning", {
          tipoDialogo: "warning",
          titulo: tituloWarning,
          contenido: warning,
        });
      },
    );

    window.electronAPI.on("dialogoAutentificarGoogle", (event, url) => {
      console.log("Solicitando codigo Google");

      const dialogoAutentificarGoogle = this.dialog.open(DialogoComponent, {
        data: {
          tipoDialogo: "googleAuth",
          titulo: "Valide el uso de esta aplicación con el siguiente Link: ",
          contenido: url,
          codigo: "",
        },
      });

      dialogoAutentificarGoogle.afterClosed().subscribe((result) => {
        console.log("Fin de Autentificación Google:");
        console.log(result);
        window.electronAPI.setCodigoGoogle(result).then((result) => {});
      });
    });

    return;
  }

  openDialog(tipoDialogoArg: string, dataArg: any) {
    const dialogRef = this.dialog.open(DialogoComponent, {
      data: {
        tipoDialogo: tipoDialogoArg,
        titulo: dataArg["titulo"],
        contenido: dataArg["contenido"],
      },
    });

    dialogRef.afterClosed().subscribe((result) => {
      console.log("Fin del dialogo");
      console.log(result);
    });
    return;
  }

  getProyecto() {
    this.proyectoConfig = window.electronAPI.getProyecto();

    if (this.proyectoConfig === null || this.proyectoConfig === undefined) {
      console.log("ERROR DE PROYECTO CONFIG");
      this.proyectoActivo = false;
    } else {
      this.proyectoActivo = true;
      this.nombreProyectoActivo = this.proyectoConfig.nombre;
    }

    console.log("Proyecto activo: " + this.proyectoActivo);
    if (this.proyectoActivo) {
      console.log(this.proyectoConfig);
    }
    return;
  }

  getArbolProyecto() {
    return new Promise((resolve) => {
      console.log("Obteniendo arbol de proyecto:" + this.nombreProyectoActivo);

      window.electronAPI
        .getArbolProyecto(this.nombreProyectoActivo)
        .then((result) => {
          resolve(result);
        });
    });
  }

  obtenerCorreo() {
    var correo = window.electronAPI.getCorreo();
    console.log("Correo: ");
    console.log(correo);
    return correo;
  }

  importarSpool(rutaSpool: string, nombreGuardado: string) {
    return new Promise((resolve) => {
      console.log("Importando SPOOL");
      window.electronAPI
        .invoke("onImportarSpool", rutaSpool, nombreGuardado)
        .then((result) => {
          resolve(result);
        });
    });
  }

  incluirDirectorio() {
    return new Promise((resolve) => {
      console.log("Incluir directorio");
      window.electronAPI.incluirDirectorio().then((result) => {
        resolve(result);
      });
    });
  }

  procesarSpool() {
    return new Promise((resolve) => {
      console.log("Analizando SPOOL");
      window.electronAPI.invoke("onProcesarSpool").then((result) => {
        resolve(result);
      });
    });
  }

  ejecutarProceso(proceso, argumentos) {
    return new Promise((resolve) => {
      console.log("Ejecutando proceso...");
      console.log(proceso);
      console.log(argumentos);
      window.electronAPI.ejecutarProceso(proceso, argumentos).then((result) => {
        resolve(result);
      });
    });
  }

  ejecutarPlantilla(proceso, argumentos) {
    return new Promise((resolve) => {
      console.log("Generando plantilla...");
      window.electronAPI
        .ejecutarPlantilla(proceso, argumentos)
        .then((result) => {
          resolve(result);
        });
    });
  }

  analizarSpool() {
    return new Promise((resolve) => {
      console.log("Analizando SPOOL");
      window.electronAPI.invoke("onAnalizarSpool").then((result) => {
        resolve(result);
      });
    });
  }

  //Antiguo guardarDocumentoElementos
  guardarEnConfiguracion(objetoConfiguracion): boolean {
    return window.electronAPI.guardarEnConfiguracion(objetoConfiguracion);
  }

  guardarDocumentoElementos(objetoElementos): boolean {
    if (objetoElementos.nombreId !== "Elementos") {
      console.log(
        "ERROR: El objeto no contiene el campo nombreId: 'Elementos'",
      );
      return false;
    }

    return window.electronAPI.setDocumento(objetoElementos);
  }

  guardarArchivo(objetoArchivo, dialogo?): Promise<any> {
    return new Promise((resolve) => {
      window.electronAPI
        .guardarDocumento(objetoArchivo, "archivos")
        .then((result) => {
          if (dialogo) {
            dialogo.close(result);
          }
          resolve(result);
        });
    });
  }

  guardarPlantilla(objetoPlantilla, dialogo): void {
    window.electronAPI
      .guardarDocumento(objetoPlantilla, "plantillas")
      .then((result) => {
        dialogo.close(result);
      });
    return;
  }

  eliminarArchivo(path: string, nombre: string, dialogo): void {
    window.electronAPI.eliminarDocumento(path, nombre).then((result) => {
      dialogo.close(result);
    });
    return;
  }

  obtenerElementosRender() {
    return window.electronAPI.getDocumento({ nombreId: "Elementos" });
  }

  getCamposDocx(pathPlantilla: string) {
    return window.electronAPI.getCamposDocx(pathPlantilla);
  }

  obtenerDatos(filtroParam) {
    if (
      filtroParam === undefined ||
      filtroParam === null ||
      filtroParam === ""
    ) {
      console.log("ERROR: Filtro obtener datos invalido");
      return false;
    }

    var filtro = filtroParam;
    if (typeof filtroParam == "string") {
      filtro = { nombreId: filtro };
    }
    console.log("Filtro: ");
    console.log(filtro);

    return window.electronAPI.getDocumento(filtro);
  }

  getListaObjetosEnColeccion(path: string, nombreArchivo: string) {
    return window.electronAPI.getListaObjetosEnColeccion(path, nombreArchivo);
  }

  getObjetoEnColeccion(path: string, nombreArchivo: string, objetoId: string) {
    return window.electronAPI.getObjetoEnColeccion(
      path,
      nombreArchivo,
      objetoId,
    );
  }

  getCorreo(plantilla: string) {
    var filtro = {};

    if (
      plantilla !== "PlantillaInstitucion" &&
      plantilla !== "PlantillaMaterial" &&
      plantilla !== "PlantillaRecordatorio" &&
      plantilla !== "PlantillaGraciasInstitucion" &&
      plantilla !== "PlantillaGraciasFormador"
    ) {
      return false;
    } else {
      filtro = { nombreId: plantilla };
    }

    console.log("Obteniendo Correo: ", plantilla, "Filtro: ", filtro);

    return window.electronAPI.getDocumentoPath(
      "../../src/plantillas-correos",
      plantilla,
      filtro,
    );
  }

  getDato(path: string, nombreArchivo: string, nombreId?: string) {
    var filtro = {};

    if (
      nombreId !== undefined &&
      nombreId !== null &&
      nombreId !== "" &&
      typeof nombreId == "string"
    ) {
      console.log("ERROR: Filtro obtener datos invalido");
      filtro = { nombreId: nombreId };
    }

    console.log("Obteniendo dato: ", nombreArchivo, "Filtro: ", filtro);

    return window.electronAPI.getDocumentoPath(path, nombreArchivo, filtro);
  }

  listarProyectos(): any[] {
    this.listaProyectos = window.electronAPI.listarProyectos();

    for (var i = 0; i < this.listaProyectos.length; i++) {
      if (typeof this.listaProyectos[i] == "undefined") {
        this.listaProyectos.splice(i, 1);
      }
    }
    console.log("Lista Proyectos: ");
    console.log(this.listaProyectos);
    return this.listaProyectos;
  }

  crearProyecto(dataProyecto) {
    if (
      dataProyecto["nombre"] === undefined ||
      dataProyecto["nombre"] === null ||
      dataProyecto["nombre"] === ""
    ) {
      console.log("Error creando proyecto: nombre no valido.");
      return false;
    }

    if (window.electronAPI.crearProyecto(dataProyecto)) {
      return this.abrirProyecto(dataProyecto.nombre);
    } else {
      return false;
    }
  }

  eliminarProyecto(nombreProyecto: string) {
    return window.electronAPI.eliminarProyecto(nombreProyecto);
  }

  cargarParametros() {
    //Obtener del localStorage los parametros guardados:
    this.parametrosProyecto = window.localStorage.getItem("parametrosProyecto");

    //Obtener en LocalStorage:
    const stringParametros = window.localStorage.getItem("parametrosProyecto");
    if (stringParametros) {
      this.parametrosProyecto = JSON.parse(stringParametros);
    } else {
      //Guardar en LocalStorage:
      this.parametrosProyecto = {
        procesos: [],
      };
      window.localStorage.setItem(
        "parametrosProyecto",
        JSON.stringify(this.parametrosProyecto),
      );
    }
  }

  abrirProyecto(nombreProyecto: string) {
    console.log("Abriendo Proyecto" + nombreProyecto);
    this.cargarParametros();
    this.proyectoActivo = true;
    this.nombreProyectoActivo = nombreProyecto;
    this.datosProyecto = window.electronAPI.abrirProyecto(nombreProyecto);
    return this.datosProyecto;
  }

  importarValoresPorDefecto(nombreProceso) {
    function camelize(str) {
      return str
        .replace(/(?:^\w|[A-Z]|\b\w)/g, function (word, index) {
          return index === 0 ? word.toLowerCase() : word.toUpperCase();
        })
        .replace(/\s+/g, "");
    }

    nombreProceso = camelize(nombreProceso);

    if (this.parametrosProyecto.procesos) {
      for (var i = 0; i < this.parametrosProyecto.procesos.length; i++) {
        if (
          this.parametrosProyecto.procesos[i].nombreProceso == nombreProceso
        ) {
          return this.parametrosProyecto.procesos[i].valorParametros;
        }
      }
    }

    return [];
  }

  setValoresPorDefecto(nombreProceso: string, valores) {
    function camelize(str) {
      return str
        .replace(/(?:^\w|[A-Z]|\b\w)/g, function (word, index) {
          return index === 0 ? word.toLowerCase() : word.toUpperCase();
        })
        .replace(/\s+/g, "");
    }

    //Normalizar nombreProceso:
    nombreProceso = camelize(nombreProceso);

    //Buscar si ya existe el proceso:
    for (var i = 0; i < this.parametrosProyecto.procesos.length; i++) {
      if (this.parametrosProyecto.procesos[i].nombreProceso == nombreProceso) {
        this.parametrosProyecto.procesos[i].valorParametros = valores;
        window.localStorage.setItem(
          "parametrosProyecto",
          JSON.stringify(this.parametrosProyecto),
        );
        return;
      }
    }

    //Si no existe, crearlo:
    this.parametrosProyecto.procesos.push({
      nombreProceso: nombreProceso,
      valorParametros: valores,
    });

    window.localStorage.setItem(
      "parametrosProyecto",
      JSON.stringify(this.parametrosProyecto),
    );

    return;
  }

  getDatosProyecto() {
    this.datosProyecto = window.electronAPI.getProyecto();
    return this.datosProyecto;
  }

  cerrarProyecto() {
    this.proyectoActivo = false;
    this.nombreProyectoActivo = "";
    this.datosProyecto = {};
    this.listaProyectos = this.listarProyectos();
    return window.electronAPI.cerrarProyecto();
  }

  abrirEditorPrograma() {
    return window.electronAPI.abrirEditorPrograma();
  }

  abrirEditorDocumento() {
    return window.electronAPI.abrirEditorDocumento();
  }

  //*************************************************
  //    SISTEMA AUTOUPDATE:
  //*************************************************

  setVersion(version) {
    console.log("Versión: ");
    console.log(version);
    this.version = version;
    return;
  }

  getVersion() {
    return this.version;
  }

  buscarActualizacion() {
    console.log("Buscando Actualizacion");
    this.openDialog("buscarActualizacion", {});
    //this.electronService.ipcRenderer.send("buscarActualizacion");

    return;
  }

  inicializarAutoupdate() {
    /*
        this.electronService.ipcRenderer.on("errorInterno", err => {
            console.log("Se ha producido un error interno: ");
            console.log(err);
            this.openDialog("informativo", {
                tipoDialogo: "informativo",
                titulo: "Se ha producido un errro interno",
                contenido:
                    "Si el problema persiste, contacte con el administrador de la aplicación."
            });
        });

        //Gestión de autoupdate:
        this.electronService.ipcRenderer.on("app_version", (event, arg) => {
            this.electronService.ipcRenderer.removeAllListeners("app_version");
            console.log("Vesion de app: " + arg.version);
            this.version = arg.version;
        });

        this.electronService.ipcRenderer.on("updateActual", event => {
            console.log("Update encontrada. ");

            this.dialog.closeAll();
            var dialogoConfirmarActualizacion = this.dialog.open(
                DialogoComponent,
                {
                    disableClose: true,
                    data: { tipoDialogo: "actualizacionActual", data: {} }
                }
            );

            dialogoConfirmarActualizacion.afterClosed().subscribe(result => {
                if (result == true) {
                    console.log("CLOSE TRUE:");
                    console.log(result);
                    this.observarAppService.next("descargarActualizacion");
                } else {
                    console.log("Descarga Cancelada");
                }
            });
        });

        this.electronService.ipcRenderer.on("updateEncontrada", event => {
            console.log("Update encontrada. ");

            this.dialog.closeAll();
            var dialogoConfirmarActualizacion = this.dialog.open(
                DialogoComponent,
                {
                    disableClose: true,
                    data: { tipoDialogo: "actualizacionEncontrada", data: {} }
                }
            );

            dialogoConfirmarActualizacion.afterClosed().subscribe(result => {
                if (result == true) {
                    console.log("CLOSE TRUE:");
                    console.log(result);
                    this.observarAppService.next("descargarActualizacion");
                } else {
                    console.log("Descarga Cancelada");
                }
            });
        });

        this.electronService.ipcRenderer.on(
            "actualizacionNoEncontrada",
            event => {
                console.log("Actualización no encontrada");
                this.dialog.closeAll();
                this.openDialog("informativo", {
                    tipoDialogo: "informativo",
                    titulo: "No hay actualizaciones disponibles.",
                    contenido: "La aplicación esta actualizada."
                });
            }
        );

        this.electronService.ipcRenderer.on(
            "dialogoAutentificarGoogle",
            (event, url) => {
                console.log("Solicitando codigo Google");

                const dialogoAutentificarGoogle = this.dialog.open(DialogoComponent, {
                    data: {
                        tipoDialogo: "googleAuth",
                        titulo: "Valide el uso de esta aplicación con el siguiente Link: ",
                        contenido: url,
                        codigo: ""  
                    }
                });

                dialogoAutentificarGoogle.afterClosed().subscribe(result => {
                    console.log("Fin de Autentificación Google:");
                    console.log(result);
                    this.electronService.ipcRenderer.send("setCodigoGoogle",result);
                });
            }
        );

        this.electronService.ipcRenderer.on(
            "progresoDescarga",
            (event, progreso) => {
                console.log("Progreso: ");
                console.log(progreso);

                //this.observarAppService.next("descargaCompletada");
            }
        );

        this.electronService.ipcRenderer.on("descargaCompletada", event => {
            console.log("Descarga Completada");
            //this.observarAppService.next("descargaCompletada");
            this.instalarActualizacion();
        });

        this.electronService.ipcRenderer.send("app_version");
       */
  }

  setProgresoDescarga(progreso) {}

  descargaCompletada() {
    this.dialog.closeAll();
    var dialogoDescargaCompletada = this.dialog.open(DialogoComponent, {
      disableClose: true,
      data: { tipoDialogo: "descargaCompletada", data: {} },
    });

    dialogoDescargaCompletada.afterClosed().subscribe((result) => {
      if (result == true) {
        this.instalarActualizacion();
      }
    });
  }

  descargarActualizacion() {
    //Gestión de autoupdate:

    console.log("Descargando Actualizacion");

    this.dialog.closeAll();
    var dialogoDescargandoActualizacion = this.dialog.open(DialogoComponent, {
      disableClose: true,
      id: "descargandoActualizacion",
      data: { tipoDialogo: "descargandoActualizacion", data: {} },
    });

    //this.electronService.ipcRenderer.send("descargarActualizacion");
    //this.aumentarProgreso(10);
  }

  instalarActualizacion() {
    console.log("Instalando Actualizacion...");
    //this.electronService.ipcRenderer.send("instalarActualizacion");
  }
} //Fin Component
