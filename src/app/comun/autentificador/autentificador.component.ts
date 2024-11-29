
import { Component, OnInit, ViewChild,ElementRef} from '@angular/core';
import { AppService } from '../../app.service';
import { MatDialog} from '@angular/material/dialog';
import { DialogoComponent } from '../dialogos/dialogos.component'
import { MatSidenav} from '@angular/material/sidenav';
import { UntypedFormBuilder, UntypedFormGroup, UntypedFormControl, Validators, UntypedFormArray} from '@angular/forms'; 
import { FlatTreeControl } from '@angular/cdk/tree';
import { MatTreeFlatDataSource, MatTreeFlattener } from '@angular/material/tree';
import { LibreriaPlantillas,libreriaPlantillas } from '../../../../plantillas/plantillas.configuracion';
	
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
  selector: 'app-editorDocumento',
  templateUrl: './autentificador.component.html',
  styleUrls: ['./autentificador.component.sass']
})

export class EditorDocumentoComponent implements OnInit{

	constructor(public appService: AppService, public dialog: MatDialog ,public formBuilder: UntypedFormBuilder) {
		//Inicialización del arbol de procesos:	
		this.arbolProcesosDataSource.data = this.datosArbolProcesos;

		this.formularioProcesoGroup = formBuilder.group({
			formularioPruebaControl: this.formularioPruebaControl
		})
	}

	//Arbol de Procesos para ejecutar:
	private _transformer = (node: LibreriaPlantillas, level: number) => {
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

	private arbolProcesosControl = new FlatTreeControl<ExpansibleNode>(node => node.level, node => node.expandable);
	private reductorArbolProcesos = new MatTreeFlattener(this._transformer, node => node.level, node => node.expandable, node => node.subCategoria);
	private arbolProcesosDataSource = new MatTreeFlatDataSource(this.arbolProcesosControl, this.reductorArbolProcesos);
	private hasChild = (_: number, node: ExpansibleNode) => node.expandable;

	public  datosArbolProcesos:LibreriaPlantillas[] = libreriaPlantillas;

	//Variables generales:
	private procesoSeleccionado: any= null;
	private objetosSeleccionados: objetosSeleccionados[] = []

	//Control de formularios: 
	private formularioProcesoGroup: UntypedFormGroup;

    private formularioControl  = new UntypedFormArray([]) 
    private formularioPruebaControl  = new UntypedFormControl("") 
	public proyectoActivo = false;
	public nombreProyectoActivo = "Selecciona un proyecto";
	public datosProyecto: any = {};
	private editor: any;
	private correo: any;

	@ViewChild('drawer',{static: false}) drawer: MatSidenav;

	ngOnInit(){

		//Inicilizar:
		console.log("Cargando Editor Documento: ");
		
		//Obtener Correo:
		this.correo = this.appService.obtenerCorreo()
		console.log("Correos: ")
		console.log(this.correo)

		console.log("ARBOL DE PLANTILLAS");
		console.log(this.datosArbolProcesos);
		this.datosArbolProcesos= libreriaPlantillas;
	}

	openDialog(tipoDialogoArg: string, dataArg: any) {

    	const dialogRef = this.dialog.open(DialogoComponent,{
      		data: {tipoDialogo: tipoDialogoArg, data: dataArg}
    	});

    	dialogRef.afterClosed().subscribe(result => {
      		console.log('Fin del dialogo');
      		console.log(result)
    	});
    	return;
  	}

	abrirProceso(node){
		
		console.log("Abriendo proceso");
		this.formularioControl = new UntypedFormArray([]);

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




