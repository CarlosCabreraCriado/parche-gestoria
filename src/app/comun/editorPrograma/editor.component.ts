
import { Component, Inject,  OnInit, ViewChild,ElementRef} from '@angular/core';
import { AppService } from '../../app.service';
import { MatDialog} from '@angular/material/dialog';
import { DialogoComponent } from '../dialogos/dialogos.component'
import { MatSidenav} from '@angular/material/sidenav';
import { AddProcesoComponent } from '../addProceso/addProceso.component'
import { Subject } from 'rxjs';
//import { Node, Edge, ClusterNode } from '@swimlane/ngx-graph';
import { SeleccionarProgramaComponent } from '../seleccionarPrograma/seleccionarPrograma.component'
import {MatSnackBar} from '@angular/material/snack-bar';

declare let LeaderLine: any;

//Modulos:
import { MatIconModule } from '@angular/material/icon';
import { MatSidenavModule } from '@angular/material/sidenav';
import {MatMenuModule} from '@angular/material/menu';

@Component({
  standalone: true,
  imports: [
      MatIconModule,
      MatMenuModule,
      MatSidenavModule
  ],
  selector: 'app-editor',
  templateUrl: './editor.component.html',
  styleUrls: ['./editor.component.sass','./drawflow.min.css']
})

export class EditorPrograma implements OnInit{

	constructor(public appService: AppService, public dialog: MatDialog, private _snackBar: MatSnackBar) { }

	public proyectoActivo = false;
	public nombreProyectoActivo = "Selecciona un proyecto";
	public datosProyecto: any = {};
	public modoEditor: string = "normal";
	public origenLink: string = "";
	public destinoLink: string = "";
	private editor: any;
	private arrow: any;
    public showfiller:boolean = false;

	//Variables de Programa:
	public nodes: any = [];
	public links: any = [];

	public update$: Subject<boolean> = new Subject();
	
	@ViewChild('drawer',{static: false}) drawer: MatSidenav;

	ngOnInit(){

		//Inicilizar:;
		console.log("Cargando Editor");

		const dialogRef = this.dialog.open(SeleccionarProgramaComponent,{
			disableClose: true,
			width: "50%",
      		data: {opciones: {}, titulo:"Seleccione un programa",contenido:"Dialogo de selección de programas"}
    	});

    	dialogRef.afterClosed().subscribe(result => {
			if(result=="exito"){
				/*
				const dialogExito = this.dialog.open(DialogoComponent,{
      				data: {tipoDialogo: "exito", titulo:"Archivo guardado con exito",contenido:"El archivo se ha guardado con exito."}
    			});
			*/
			}

			if(result=="error"){
				const dialogExito = this.dialog.open(DialogoComponent,{
      				data: {tipoDialogo: "error", titulo:"Se ha producido un error.",contenido:"Error desconocido"}
    			});
			}

      		console.log('Fin de Herramienta AddDato: '+result);
			return;
    	});

		//Inicializa un Programa de prueba:
		this.nodes = [
				{
					id: 'proceso_1',
					label: 'A',
					dimension: {width: 30, height: 30},
					titulo: "Renderizar Report Remedy",
					idProceso: "renderizarReportRemedy",

					entradas: [
						{tipo: "texto"},
						{tipo: "texto"},
						{tipo: "texto"}
					],

					salidas: [
						{tipo: "texto"},
						{tipo: "texto"}
					]
				}
			  ]

		this.links = [
				{
				  id: 'link_1',
				  source: 'proceso_1',
				  target: 'proceso_2',
				  label: 'is parent of'
				}
			  ]

		new LeaderLine(
			document.getElementById('proceso_1'),
			document.getElementById('proceso_2'),
			{color: "red"}
		);

	}

	//Funciones de herramientas:
	seleccionarPrograma(opcionesArg:any){

		const dialogRef = this.dialog.open(SeleccionarProgramaComponent,{
			disableClose: true,
			width: "50%",
      		data: {opciones: opcionesArg, titulo:"Seleccione un programa",contenido:"Dialogo de selección de programas"}
    	});

    	dialogRef.afterClosed().subscribe(result => {
			if(result=="exito"){
				/*
				const dialogExito = this.dialog.open(DialogoComponent,{
      				data: {tipoDialogo: "exito", titulo:"Archivo guardado con exito",contenido:"El archivo se ha guardado con exito."}
    			});
			*/
			}

			if(result=="error"){
				const dialogExito = this.dialog.open(DialogoComponent,{
      				data: {tipoDialogo: "error", titulo:"Se ha producido un error.",contenido:"Error desconocido"}
    			});
			}

      		console.log('Fin de Herramienta AddDato: '+result);
			return;
    	});
	}

	seleccionarHerramienta(herramienta:string, opcionesArg){
		switch(herramienta){

			//Seleccion de herramienta AddProceso
			case "addProceso":
				const dialogRef = this.dialog.open(AddProcesoComponent,{
					disableClose: true,
					width: "70%",
					data: {opciones: opcionesArg, titulo:"Insertar Proceso",contenido:"Heramienta de procesos."}
				});

				dialogRef.afterClosed().subscribe(result => {
					if(result!=="error"){
						console.log("Añadiendo Proceso: ")
						console.log(result)
						this.addProceso(result)
					}

					if(result=="error"){
						const dialogExito = this.dialog.open(DialogoComponent,{
							data: {tipoDialogo: "error", titulo:"Se ha producido un error.",contenido:"No se he podido incluir el proceso en el programa"}
						});
					}
				});
				break; //Fin de herramienta AddProceso

			case "addLink":
		
				if(this.modoEditor!="link"){
					this.abrirSnack("Selecciona el primer proceso a conectar", "Cancelar")
					this.modoEditor="link"
				}else{
					this.modoEditor="normal"
					this.cerrarSnack()
				}
				break;

			case "borrar":
		
				if(this.modoEditor!="borrar"){
					this.abrirSnack("Seleccione el elemento a borrar", "Cancelar")
					this.modoEditor="borrar"
				}else{
					this.modoEditor="normal"
					this.cerrarSnack()
				}
				break;
		}	

			return;
	}

	guardarPrograma(){
		var programa = {
			nombreId: "programa_1",
			objetoId: "programa_1",
			data: {
				nodes: this.nodes,
				links: this.links
			}
		}

		const dialogoProcesando = this.dialog.open(DialogoComponent,{ disableClose: true,
    	      data: {tipoDialogo: "procesando", titulo: "Procesando", contenido: ""}
    	  });

		dialogoProcesando.afterClosed().subscribe(result => {
			if(result==true){
				console.log("Archivo guardando con exito");
				//this.dialogRef.close("exito");
			}else{
				console.log("Error guardando archivo");
				//this.dialogRef.close("error");
			}
      	});

		this.appService.guardarArchivo(programa, dialogoProcesando)
		
		return;
	}

	abrirSnack(message: string, action: string) {
	   this._snackBar.open(message, action);
	}

	cerrarSnack() {
	   this._snackBar.dismiss();
	}

	addProceso(proceso){

		//Añadiendo nodo de proceso:
		this.nodes.push(
			{
				id: "proceso_"+this.nodes.length,
				label: proceso.nombre ,
				dimension: {width: 30, height: 30},
				titulo: "",
				idProceso: "",

				entradas: [
					{tipo: "texto"},
					{tipo: "texto"}
				],

				salidas: [
					{tipo: "texto"}
				]
			})

		this.update$.next(true)	
		console.log(this.nodes)

	}

	onNodeSelect(event){
		return;	
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

	clickNode(data, nodeIndex: number){
		console.log("click",data);

		switch(this.modoEditor){

			case "link":

				if(this.origenLink==""){
					this.origenLink = data.id;	
					console.log("Origen: ",this.origenLink);
				}else if(this.destinoLink==""){
					this.destinoLink = data.id;
					console.log("Destino: ",this.destinoLink);

					//Crear Link:
					var lengthLink= this.links.length+1
					this.links.push({
					  id: "link_"+lengthLink.toString(),
					  source: this.origenLink,
					  target: this.destinoLink,
					  label: "" 
					})

					console.log("Link creado:",this.links)

					var leader= new LeaderLine(
						document.getElementById(this.origenLink),
						document.getElementById(this.destinoLink),
						{color: "blue"}
					);

					this.modoEditor="normal"
					this.origenLink=""
					this.destinoLink=""
					this.cerrarSnack()
					this.update$.next(true)	
				}
			break;

			case "borrar":
				this.nodes.splice(nodeIndex,1);
				this.modoEditor="normal"
				this.cerrarSnack()
				this.update$.next(true)	
				break;
		}
	}
	clickSnap(nodeData,nodeIndex,tipoIO,indexIO){

		switch(this.modoEditor){

			case "link":

				if(this.origenLink==""){
					this.origenLink = nodeData.id;	
					console.log("Origen: ",this.origenLink);
				}else if(this.destinoLink==""){
					this.destinoLink = nodeData.id;
					console.log("Destino: ",this.destinoLink);

					//Crear Link:
					var lengthLink= this.links.length+1
					this.links.push({
					  id: "link_"+lengthLink.toString(),
					  source: this.origenLink,
					  target: this.destinoLink,
					  label: "" 
					})

					console.log("Link creado:",this.links)

					var leader= new LeaderLine(
						document.getElementById(this.origenLink),
						document.getElementById(this.destinoLink),
						{color: "blue"}
					);

					this.modoEditor="normal"
					this.origenLink=""
					this.destinoLink=""
					this.cerrarSnack()
					this.update$.next(true)	
				}
			break;

		}


	}

}




