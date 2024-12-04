import { Component, Inject, OnInit } from "@angular/core";
import {
  MatDialog,
  MatDialogRef,
  MAT_DIALOG_DATA,
} from "@angular/material/dialog";
import {
  UntypedFormBuilder,
  UntypedFormGroup,
  UntypedFormControl,
  Validators,
  UntypedFormArray,
} from "@angular/forms";
import { DialogoComponent } from "../dialogos/dialogos.component";
import { AppService } from "../../app.service";
import {
  MatTreeFlatDataSource,
  MatTreeFlattener,
} from "@angular/material/tree";
import { FlatTreeControl } from "@angular/cdk/tree";
import {
  LibreriaProcesos,
  libreriaProcesos,
} from "../procesos/procesos.configuracion";

//Modulos:
import { MatSelectModule } from "@angular/material/select";
import { FormsModule, ReactiveFormsModule } from "@angular/forms";
import { MatDatepickerModule } from "@angular/material/datepicker";
import { MatIconModule } from "@angular/material/icon";
import { MatSidenavModule } from "@angular/material/sidenav";
import { MatTreeModule } from "@angular/material/tree";
import { MatDialogModule } from "@angular/material/dialog";
import { MatButtonModule } from "@angular/material/button";

import { MatInputModule } from "@angular/material/input";
import { MatFormFieldModule } from "@angular/material/form-field";

export interface AddDatoData {
  opciones: any;
  data: any;
}

interface ProcesosNode {
  nombre: string;
  tipo: string;
  argumentos: [];
  salida: [];
  subDirectorio?: ProcesosNode[];
}

interface ExpansibleNode {
  expandable: boolean;
  categoria: string;
  nombre: string;
  level: number;
}

interface objetosSeleccionados {
  indexControl: number;
  node: any;
}

@Component({
  standalone: true,
  imports: [
    MatInputModule,
    MatFormFieldModule,
    MatButtonModule,
    MatDialogModule,
    MatSidenavModule,
    MatIconModule,
    MatSelectModule,
    FormsModule,
    ReactiveFormsModule,
    MatTreeModule,
    MatDatepickerModule,
  ],
  selector: "ejecutarProceso",
  templateUrl: "./ejecutarProceso.component.html",
  styleUrls: ["./ejecutarProceso.component.sass"],
})
export class EjecutarProcesoComponent implements OnInit {
  //Arbol de Procesos para ejecutar:
  private _transformer = (node: LibreriaProcesos, level: number) => {
    return {
      expandable: !!node.subCategoria && node.subCategoria.length > 0,
      categoria: node.categoria,
      level: level,
      nombre: node.nombre,
      tipo: node.tipo,
      autor: node.autor,
      opciones: node.opciones,
      salida: node.salida,
      argumentos: node.argumentos,
      descripcion: node.descripcion,
    };
  };

  public arbolProcesosControl = new FlatTreeControl<ExpansibleNode>(
    (node) => node.level,
    (node) => node.expandable,
  );
  private reductorArbolProcesos = new MatTreeFlattener(
    this._transformer,
    (node) => node.level,
    (node) => node.expandable,
    (node) => node.subCategoria,
  );
  public arbolProcesosDataSource = new MatTreeFlatDataSource(
    this.arbolProcesosControl,
    this.reductorArbolProcesos,
  );
  public hasChild = (_: number, node: ExpansibleNode) => node.expandable;

  public datosArbolProcesos: LibreriaProcesos[] = libreriaProcesos;

  //Variables generales:
  public procesoSeleccionado: any = null;
  private objetosSeleccionados: objetosSeleccionados[] = [];

  //Control de formularios:
  public formularioProcesoGroup: UntypedFormGroup;

  public formularioControl = new UntypedFormArray([]);
  public formularioPruebaControl = new UntypedFormControl("");
  public formularioCargado: boolean = false;
  public argumentosActivos: any;

  constructor(
    private appService: AppService,
    public dialogRef: MatDialogRef<EjecutarProcesoComponent>,
    @Inject(MAT_DIALOG_DATA) public data: AddDatoData,
    public formBuilder: UntypedFormBuilder,
    private dialog: MatDialog,
  ) {
    //Inicializacion del arbol de procesos:
    this.arbolProcesosDataSource.data = this.datosArbolProcesos;

    /*
    this.formularioProcesoGroup = formBuilder.group({
      formularioControl: new UntypedFormArray([]),
    });
    */
  }

  ngOnInit() {
    this.datosArbolProcesos = libreriaProcesos;
  }

  generarArbolProcesos(): ProcesosNode[] {
    var arbolProcesoTemporal: ProcesosNode[] = [];
    console.log(libreriaProcesos);
    for (var i = 0; i < libreriaProcesos.length; i++) {}

    return arbolProcesoTemporal;
  }

  ejecutarProceso(proceso: LibreriaProcesos) {
    //Gestionar procesos de redireccion:
    if (proceso.tipo == "redireccion") {
      this.appService.cambiarUrl("" + proceso.salida[0].valor);
      return;
    }

    //Extraer parametros:
    console.log("Argumentos proceso: ");
    var argumentos = this.formularioProcesoGroup.getRawValue();

    //Extraccion de los objetos:
    var objeto = {};
    for (var i = 0; i < this.objetosSeleccionados.length; i++) {
      objeto = this.appService.getDato(
        this.objetosSeleccionados[i].node.direccion,
        this.objetosSeleccionados[i].node.nombre,
      );
      argumentos.formularioControl[this.objetosSeleccionados[i].indexControl] =
        objeto;
    }

    console.log(argumentos);

    const dialogoProcesando = this.dialog.open(DialogoComponent, {
      disableClose: true,
      data: { tipoDialogo: "procesando", titulo: "Procesando", contenido: "" },
    });

    this.appService.ejecutarProceso(proceso, argumentos).then((result) => {
      console.log("Proceso finalizado: ");
      console.log(result);
      dialogoProcesando.close();
      if (!result) {
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

  incluirDirectorio(indexControl) {
    const dialogoBloqueo = this.dialog.open(DialogoComponent, {
      disableClose: true,
      data: {
        tipoDialogo: "bloqueoVentana",
        titulo: "Cierre las ventanas para continuar",
        contenido:
          "Hay una ventana de sistema abierta. Cierrela para continuar.",
      },
    });

    this.appService.incluirDirectorio().then((result: any) => {
      console.log("Incluyendo directorio:");
      console.log(result);
      if (result.canceled) {
        dialogoBloqueo.close();
        return;
      }
      var formularioTemporal: any;
      formularioTemporal = this.formularioProcesoGroup.getRawValue();
      formularioTemporal.formularioControl[indexControl] = result.filePaths[0];
      //this.addTargetFile("rutaArchivoControl", target.files[0]);
      this.formularioProcesoGroup.setValue(formularioTemporal);
      dialogoBloqueo.close();
      return;
    });
  }
  incluirRuta(evt: any, indexControl: number) {
    //Lectura de evento Input
    const target: DataTransfer = <DataTransfer>evt.target;

    console.log("Objeto Ruta:");
    console.log(target.files);

    var formularioTemporal: any;
    formularioTemporal = this.formularioProcesoGroup.getRawValue();
    formularioTemporal.formularioControl[indexControl] =
      target.files[0]["path"];
    this.formularioProcesoGroup.setValue(formularioTemporal);
  }

  incluirObjeto(objeto: any, indexControl: number) {
    console.log(objeto);
    console.log(indexControl);

    //Asignacion de array de objetos seleccionados:
    var objetoEncontrado = false;
    for (var i = 0; i < this.objetosSeleccionados.length; i++) {
      if (this.objetosSeleccionados[i].indexControl == indexControl) {
        objetoEncontrado = true;
        this.objetosSeleccionados[i] = {
          indexControl: indexControl,
          node: objeto,
        };
      }
    }

    if (!objetoEncontrado) {
      this.objetosSeleccionados.push({
        indexControl: indexControl,
        node: objeto,
      });
    }

    var formularioTemporal: any;
    formularioTemporal = this.formularioProcesoGroup.getRawValue();
    console.log(this.formularioProcesoGroup);
    console.log(formularioTemporal);
    formularioTemporal.formularioControl[indexControl] = objeto.nombre;
    this.formularioProcesoGroup.setValue(formularioTemporal);
  }

  abrirProceso(node) {
    console.log("Abriendo proceso: ", node.argumentos);

    this.argumentosActivos = node.argumentos;
    this.formularioProcesoGroup = this.formBuilder.group({});

    this.formularioControl = new UntypedFormArray([]);
    this.objetosSeleccionados = [];

    //DEFINICIONES DE FORMULARIO:
    for (var i = 0; i < node.argumentos.length; i++) {
      if (node.argumentos[i].tipo == "objeto") {
        this.formularioProcesoGroup.addControl(
          node.argumentos[i].identificador,
          new UntypedFormControl({ value: "", disabled: true }),
        );
        /*
        this.formularioControl.push(
          new UntypedFormControl({ value: "", disabled: true }),
        );
                */
      } else {
        if (node.argumentos[i].formulario.valorDefault) {
          this.formularioProcesoGroup.addControl(
            node.argumentos[i].identificador,
            new UntypedFormControl({
              value: node.argumentos[i].formulario.valorDefault,
              disabled: false,
            }),
          );

          /*
          this.formularioControl.push(
            new UntypedFormControl({
              value: node.argumentos[i].formulario.valorDefault,
              disabled: false,
            }),
          );
                        */
        } else {
          console.log("Agregando control: ");

          this.formularioProcesoGroup.addControl(
            node.argumentos[i].identificador,
            new UntypedFormControl({
              value: "",
              disabled: false,
            }),
          );

          /*
          this.formularioControl.push(
            new UntypedFormControl({ 
                            value: "", 
                            disabled: false 
                        }),
          );

                    */
        }
      }
    }

    /*
    this.formularioProcesoGroup = this.formBuilder.group({
      formularioControl: this.formularioControl,
    });
        */

    console.log(this.formularioProcesoGroup);
    //console.log(this.formularioControl);

    this.procesoSeleccionado = node;
    console.log(node);
    this.formularioCargado = true;
    return;
  }

  getFormArray(): UntypedFormArray {
    return this.formularioProcesoGroup.get(
      "formularioControl",
    ) as UntypedFormArray;
  }

  seleccionarObjeto(indexControl: number) {
    this.appService.getArbolProyecto().then((result) => {
      const dialogoSeleccionarObjeto = this.dialog.open(DialogoComponent, {
        disableClose: false,
        data: {
          tipoDialogo: "seleccionarObjeto",
          titulo: "Seleccione un objeto",
          contenido: result,
        },
      });

      dialogoSeleccionarObjeto.afterClosed().subscribe((result) => {
        if (result == "error") {
          const dialogExito = this.dialog.open(DialogoComponent, {
            data: {
              tipoDialogo: "error",
              titulo: "Se ha producido un error.",
              contenido: "No se ha podido realizar la selección.",
            },
          });
        } else if (result != undefined && result != false) {
          //Asignando Objeto:
          this.incluirObjeto(result, indexControl);
        }

        console.log("Fin dialogo seleccionar objeto: ");
        console.log(result);
        return;
      });
    });
  }
}
