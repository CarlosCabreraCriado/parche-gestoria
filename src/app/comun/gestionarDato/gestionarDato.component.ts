import { Component , Inject,  OnInit, ViewChild} from '@angular/core';
import { MatDialog, MatDialogRef, MAT_DIALOG_DATA} from '@angular/material/dialog';
import { UntypedFormBuilder, UntypedFormGroup, UntypedFormControl, Validators, UntypedFormArray} from '@angular/forms'; 
import { JsonEditorComponent, JsonEditorOptions } from 'ang-jsoneditor';
import { AppService } from '../../app.service';
import { FlatTreeControl } from '@angular/cdk/tree';
import { DialogoComponent } from '../dialogos/dialogos.component';
import { MatTreeFlatDataSource, MatTreeFlattener } from '@angular/material/tree';
import { LibreriaProcesos,libreriaProcesos} from '../procesos/procesos.configuracion';
    
//Modulos:
import { MatIconModule } from '@angular/material/icon';
import { MatSidenavModule } from '@angular/material/sidenav';
import { MatTreeModule } from '@angular/material/tree';
import { MatDialogModule } from '@angular/material/dialog';
import { MatButtonModule } from '@angular/material/button';

export interface AddDatoData {
    opciones: any;
    data: any;
}

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

@Component({
  selector: 'gestionarDato',
  standalone: true,
  imports: [
      MatButtonModule,
      MatIconModule,
      MatDialogModule,
      MatTreeModule,
      MatSidenavModule
  ],
  templateUrl: './gestionarDato.component.html',
  styleUrls: ['./gestionarDato.component.sass']
})

export class GestionarDato implements OnInit{

    constructor(private appService: AppService, public dialogRef: MatDialogRef<GestionarDato>, @Inject(MAT_DIALOG_DATA) public data: AddDatoData, public formBuilder: UntypedFormBuilder, private dialog: MatDialog)  {
        
        //Inicializacion del arbol de procesos: 
        this.arbolArchivosDataSource.data = this.datosArbolArchivos;

        //Inicializar JSON Viewer:
        this.editorVerOptions = new JsonEditorOptions()
        this.editorModificarOptions = new JsonEditorOptions()
    }

    //Arbol de Procesos para ejecutar:
    private _transformer = (node: ArchivoNode, level: number) => {
        return {
          expandable: !!node.subDirectorio && node.subDirectorio.length > 0,
          nombre: node.nombre,
          tipo: node.tipo,
          direccion: node.direccion,
          level: level,
        };
      }

    public arbolArchivosControl = new FlatTreeControl<ExpansibleNode>(node => node.level, node => node.expandable);
    private reductorArbolArchivos = new MatTreeFlattener(this._transformer, node => node.level, node => node.expandable, node => node.subDirectorio);
    public arbolArchivosDataSource = new MatTreeFlatDataSource(this.arbolArchivosControl, this.reductorArbolArchivos);
    public hasChild = (_: number, node: ExpansibleNode) => node.expandable;

    public  datosArbolArchivos: ArchivoNode[] = [];

    //Variables generales:
    public archivoSeleccionado: any= null;

    //Control de formularios: 
    private formularioProcesoGroup: UntypedFormGroup;

    private formularioControl  = new UntypedFormArray([]) 
    private formularioPruebaControl  = new UntypedFormControl("") 

    //JSON Viewer:
    @ViewChild(JsonEditorComponent, { static: true }) editor: JsonEditorComponent;
    public editorVerOptions: JsonEditorOptions;
    public editorModificarOptions: JsonEditorOptions;
    public archivoJSONviewer: any = {};

    ngOnInit() {

        this.appService.getArbolProyecto().then((result:ArchivoNode[])=>{
            console.log("Promesa:");
            console.log(result);
            this.arbolArchivosDataSource.data = result; 
        })
    }

    reloadArbolProyecto(){
        this.appService.getArbolProyecto().then((result:ArchivoNode[])=>{
            console.log("Promesa:");
            console.log(result);
            this.arbolArchivosDataSource.data = result; 
        })
    }
    
    gestionarArchivo(node){
        this.archivoSeleccionado = node;
        console.log(node)
        return;
    }

    eliminarArchivo(){

        const dialogoProcesando = this.dialog.open(DialogoComponent,{ disableClose: true,
              data: {tipoDialogo: "procesando", titulo: "Procesando", contenido: ""}
          });

        dialogoProcesando.afterClosed().subscribe(result => {
            if(result){
                const dialogoExito = this.dialog.open(DialogoComponent,{ disableClose: true,
                      data: {tipoDialogo: "exito", titulo: "Archivo eliminado con exito", contenido: ""}
                  });
                  this.reloadArbolProyecto();
                  this.archivoSeleccionado= null;
            }else{
                const dialogoError = this.dialog.open(DialogoComponent,{ disableClose: true,
                      data: {tipoDialogo: "error", titulo: "Se ha producido un error desconocido", contenido: ""}
                  });
            }
        });

        if(this.archivoSeleccionado){
            this.appService.eliminarArchivo(this.archivoSeleccionado.direccion, this.archivoSeleccionado.nombre, dialogoProcesando);
        }
    }
    
}





