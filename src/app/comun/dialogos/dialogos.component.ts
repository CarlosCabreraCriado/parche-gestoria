import { FlatTreeControl } from '@angular/cdk/tree';
import { MatTreeFlatDataSource, MatTreeFlattener } from '@angular/material/tree';
import { Component , Inject, ViewChild, OnInit, ViewEncapsulation} from '@angular/core';
import { MatDialog, MatDialogRef, MAT_DIALOG_DATA} from '@angular/material/dialog';
import { JsonEditorComponent, JsonEditorOptions } from 'ang-jsoneditor';

//Modulos:
import { MatDialogModule } from '@angular/material/dialog';
import { MatDatepickerModule } from '@angular/material/datepicker';
import { FormsModule, ReactiveFormsModule } from '@angular/forms';
import { MatIconModule } from '@angular/material/icon';
import { MatSlideToggleModule } from '@angular/material/slide-toggle';
import { MatListModule } from '@angular/material/list';
import { MatCardModule } from '@angular/material/card';
import { MatDividerModule } from '@angular/material/divider';
import { MatTreeModule } from '@angular/material/tree';
import { MatProgressBarModule } from '@angular/material/progress-bar';
import { MatProgressSpinnerModule} from '@angular/material/progress-spinner';
import { NgJsonEditorModule } from "ang-jsoneditor";
import { MatSelectModule} from '@angular/material/select';
import { MatInputModule } from '@angular/material/input';
import { MatButtonModule } from '@angular/material/button';

export interface DialogData {
  tipoDialogo: string;
  contenido: any;
  data: any;
  titulo: string;
  codigo: any;
  inputLabel: string;
  valorInput: any;
  tipoDocumento: string;
  tituloDialogo: string;
  pathOutput: string;
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
  standalone: true,
  encapsulation: ViewEncapsulation.None,
  imports: [
      MatSelectModule,
      NgJsonEditorModule,
      MatProgressBarModule,
      MatProgressSpinnerModule,
      MatDividerModule,
      MatCardModule,
      MatSlideToggleModule,
      MatIconModule,
      MatListModule,
      FormsModule,
      ReactiveFormsModule,
      MatDialogModule, 
      MatTreeModule,
      MatDatepickerModule,
      MatInputModule,
      MatButtonModule
  ],
  selector: 'dialog-elements-example-dialog',
  templateUrl: './dialogos.component.html',
  styleUrls: ['./dialogos.component.sass']
})

export class DialogoComponent implements OnInit{

    public editorVerOptions: JsonEditorOptions;
    public editorModificarOptions: JsonEditorOptions;

    @ViewChild(JsonEditorComponent, { static: true }) editor: JsonEditorComponent;

    constructor(public dialogRef: MatDialogRef<DialogoComponent>, @Inject(MAT_DIALOG_DATA) public data: DialogData) {

        this.arbolArchivosDataSource.data = this.datosArbolArchivos;

        this.editorVerOptions = new JsonEditorOptions()
        this.editorModificarOptions = new JsonEditorOptions()
        this.editorModificarOptions.mode = 'tree'; // set all allowed modes
        this.editorVerOptions.mode = 'view'; // set all allowed modes
        console.log("Abriendo Dialogo: "+this.data.tipoDialogo)
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
    public reductorArbolArchivos = new MatTreeFlattener(this._transformer, node => node.level, node => node.expandable, node => node.subDirectorio);
    public arbolArchivosDataSource = new MatTreeFlatDataSource(this.arbolArchivosControl, this.reductorArbolArchivos);
    public hasChild = (_: number, node: ExpansibleNode) => node.expandable;

    public  datosArbolArchivos: ArchivoNode[] = [];

    ngOnInit(){

        //Inicializacion del arbol de procesos: 
        var arbol: ArchivoNode[]= this.data.contenido
        
        if(this.data.tipoDialogo== "seleccionarObjeto"){
            this.arbolArchivosDataSource.data = arbol;
        }

        if(this.data.tipoDialogo== "filtroFecha"){
           this.data["filtroFecha"]= {
               fechaInicio: 0,
               fechaFin:0
           }
        }

        if(this.data.tipoDialogo== "filtroGenerarDocumento"){
           this.data["filtroFecha"]= {
               fechaInicio: 0,
               fechaFin:0,
           }
           this.data["tipoDocumento"]= "documento";
           this.data["tituloDialogo"]= "Titulo";
           this.data["pathOutput"]= "";
        }

    }

    onNoClick(): void {
        this.dialogRef.close(false);
    }

    seleccionarObjeto(node: ArchivoNode){
        this.dialogRef.close(node);
    }

    incluirTituloPeriodo(evento){
        console.log("Evento: ",evento);
        this.data["tituloDialogo"]= evento.value;
    }

    incluirRutaDocumento(evento){
        this.data["pathOutput"]= evento.value;
    }

    incluirRuta(evt: any){

        //Lectura de evento Input
        const target: DataTransfer = <DataTransfer>(evt.target);

        console.log("Objeto Ruta:");
        console.log(target.files);

        this.data["valorInput"]= target.files

    }   

    addComentario(comentario: string){
        this.data.contenido.push({
            fecha: new Date(Date.now()),
            completo: false,
            comentario: comentario
        })
    }

    eliminarComentario(comentario){
        this.data.contenido.splice(this.data.contenido.indexOf(comentario),1);
    }

    toggleComentario(comentario){
        comentario.completo = !comentario.completo;
    }

    cambiarFechaInicio(event){

       if(this.data["filtroFecha"] == undefined){
           this.data["filtroFecha"]= {
               fechaInicio: 0,
               fechaFin:0
           }
       }
       this.data["filtroFecha"]["fechaInicio"] = event.value;


    }

    cambiarFechaFin(event){
        if(this.data["filtroFecha"] == undefined){
           this.data["filtroFecha"]= {
               fechaInicio: 0,
               fechaFin:0
           }
        }
        this.data["filtroFecha"]["fechaFin"] = event.value;
    }

}





