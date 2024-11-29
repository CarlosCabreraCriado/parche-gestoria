import { Component , Inject,  OnInit} from '@angular/core';
import { MatDialog, MatDialogRef, MAT_DIALOG_DATA} from '@angular/material/dialog';
import { UntypedFormBuilder, UntypedFormGroup, UntypedFormControl, Validators, UntypedFormArray} from '@angular/forms'; 
import { DialogoComponent } from '../dialogos/dialogos.component';
import { AppService } from '../../app.service';
import { FlatTreeControl } from '@angular/cdk/tree';
import { MatTreeFlatDataSource, MatTreeFlattener } from '@angular/material/tree';
import { LibreriaProcesos,libreriaProcesos} from '../procesos/procesos.configuracion';
    

//Modulos:
import { MatIconModule } from '@angular/material/icon';
import { MatTreeModule } from '@angular/material/tree';
import { MatDialogModule } from '@angular/material/dialog';
import { FormsModule, ReactiveFormsModule } from '@angular/forms';
import { MatSidenavModule } from '@angular/material/sidenav';

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
    indexControl : number;
    node: any;
}

@Component({
  standalone: true,
  imports: [
      MatSidenavModule,
      MatDialogModule, 
      FormsModule,
      ReactiveFormsModule,
      MatIconModule, 
      MatTreeModule
  ],
  selector: 'addProceso',
  templateUrl: './addProceso.component.html',
  styleUrls: ['./addProceso.component.sass']
})

export class AddProcesoComponent implements OnInit{

    constructor(private appService: AppService, public dialogRef: MatDialogRef<AddProcesoComponent>, @Inject(MAT_DIALOG_DATA) public data: AddDatoData, public formBuilder: UntypedFormBuilder, private dialog: MatDialog)  {
        
        //Inicializacion del arbol de procesos: 
        this.arbolProcesosDataSource.data = this.datosArbolProcesos;

        this.formularioProcesoGroup = formBuilder.group({
            formularioPruebaControl: this.formularioPruebaControl
        })
    }
    
    //Arbol de Procesos para ejecutar:
    private _transformer = (node: LibreriaProcesos, level: number) => {
        return {
          expandable: (!!node.subCategoria && node.subCategoria.length > 0),
          categoria: node.categoria,
          level: level,
          nombre: node.nombre,
          tipo: node.tipo,
          autor: node.autor,
          opciones: node.opciones,
          salida: node.salida,
          argumentos: node.argumentos,
          descripcion: node.descripcion
        };
      }

    public arbolProcesosControl = new FlatTreeControl<ExpansibleNode>(node => node.level, node => node.expandable);
    private reductorArbolProcesos = new MatTreeFlattener(this._transformer, node => node.level, node => node.expandable, node => node.subCategoria);
    public arbolProcesosDataSource = new MatTreeFlatDataSource(this.arbolProcesosControl, this.reductorArbolProcesos);
    public hasChild = (_: number, node: ExpansibleNode) => node.expandable;

    public  datosArbolProcesos:LibreriaProcesos[] = libreriaProcesos;

    //Variables generales:
    public procesoSeleccionado: any= null;
    public objetosSeleccionados: objetosSeleccionados[] = []

    //Control de formularios: 
    public formularioProcesoGroup: UntypedFormGroup;

    public formularioControl  = new UntypedFormArray([]) 
    public formularioPruebaControl  = new UntypedFormControl("") 

    ngOnInit() {
        this.datosArbolProcesos= libreriaProcesos;
    }
    

    generarArbolProcesos():ProcesosNode[]{

        var arbolProcesoTemporal:ProcesosNode[] = [];
        console.log(libreriaProcesos);
        for(var i=0; i<libreriaProcesos.length; i++){
        }

        return arbolProcesoTemporal;    
    }


    abrirProceso(node){
        
        console.log("Abriendo proceso");
        this.formularioControl = new UntypedFormArray([]);
        this.objetosSeleccionados = [];

        //DEFINICIONES DE FORMULARIO:
        for(var i=0; i< node.argumentos.length; i++){
            if(node.argumentos[i].tipo == "objeto"){
                this.formularioControl.push(new UntypedFormControl({value: '', disabled: true}));
            }else{
                if(node.argumentos[i].formulario.valorDefault){
                    this.formularioControl.push(new UntypedFormControl({value: node.argumentos[i].formulario.valorDefault, disabled: false}));
                }else{
                    this.formularioControl.push(new UntypedFormControl({value: '', disabled: false}));
                }
            }
        }
        
        this.formularioProcesoGroup = this.formBuilder.group({
          formularioControl: this.formularioControl
        });

        console.log(this.formularioControl)

        this.procesoSeleccionado = node;
        console.log(node);
        return;
    }

}





