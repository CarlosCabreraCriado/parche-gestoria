
import { Component, OnInit, ViewChild} from '@angular/core';
import { AppService } from '../../app.service';
import { MatDialog} from '@angular/material/dialog';
import { DialogoComponent } from '../dialogos/dialogos.component'
import { MatSidenav} from '@angular/material/sidenav';

//Modulos:
import { MatIconModule } from '@angular/material/icon';
import { MatCardModule } from '@angular/material/card';
import { MatSidenavModule } from '@angular/material/sidenav';
import { MatButtonModule } from '@angular/material/button';

@Component({
  standalone: true,
  imports: [
      MatButtonModule,
      MatIconModule,
      MatSidenavModule,
      MatCardModule
  ],
  selector: 'app-index',
  templateUrl: './index.component.html',
  styleUrls: ['./index.component.sass']
})

export class IndexComponent implements OnInit{

	constructor(public appService: AppService, public dialog: MatDialog) { }

	public proyectoActivo = false;
	public nombreProyectoActivo = "Selecciona un proyecto";
	public datosProyecto: any = {};
	public listaProyectos: any =[];
    public showFiller = false;

	@ViewChild('drawer',{static: false}) drawer: MatSidenav;

	ngOnInit(){
		//Inicilizar:
		console.log("Cargando Index");
		this.appService.getProyecto();
		this.listarProyectos();
	}

	listarProyectos(){
		this.listaProyectos= this.appService.listarProyectos();

        console.warn("Proyectos cargados: ",this.listaProyectos);
		for (var i = 0; i < this.listaProyectos.length; ++i) {
			//this.listaProyectos[i]= this.listaProyectos[i].replace(/.db/gi,"")
			//this.listaProyectos[i]= this.listaProyectos[i].nombreProyecto;
		}
		return;
	}

	seleccionar(herramienta: string):void{
       	this.appService.cambiarUrl("/"+herramienta);
	}

	cargarProyecto(nombreProyecto){
		this.datosProyecto = this.appService.abrirProyecto(nombreProyecto);
       	this.appService.cambiarUrl("/dashboard");
		console.log("Datos del proyecto");
		console.log(this.datosProyecto)
		if(this.datosProyecto){
			this.proyectoActivo= true
			this.nombreProyectoActivo= nombreProyecto;
			this.drawer.open();
		}else{
			console.log("Se ha producido un error en la carga del Proyecto")
			this.nombreProyectoActivo= "Selecciona un proyecto";
			this.proyectoActivo= false
		}
		return;
	}

	cerrarProyecto(){
		this.appService.cerrarProyecto();
		this.nombreProyectoActivo= "Selecciona un proyecto";
		this.proyectoActivo= false;
		this.listarProyectos();
		return;
	}

	eliminarProyecto(nombreProyecto){
		
		const dialogRef = this.dialog.open(DialogoComponent,{
      		data: {tipoDialogo: "confirmacion", titulo:"¿Seguro que quiere eliminar el proyecto '"+nombreProyecto+"'?",contenido:"Si elimina el proyecto se eliminarán todos los archivos importados en el mismo."}
    	});

    	dialogRef.afterClosed().subscribe(result => {
      		console.log('Fin del dialogo: '+result);
 
      		if(result===true){
      			//this.cerrarProyecto();
				this.appService.eliminarProyecto(nombreProyecto);
				this.listarProyectos();
      		}else{return}
      			
    	});
		return;   		
	}

	crearProyecto(){
		console.log("Creando Proyecto")
		
		const dialogRef = this.dialog.open(DialogoComponent,{
      		data: {tipoDialogo: "crearProyecto", data: {}}
    	});

    	dialogRef.afterClosed().subscribe(result => {
      		console.log('Fin del dialogo');
      		console.log(result);
      		if(result===undefined){return;}
      		
      		if(result.nombre=== undefined || result.nombre=== null){
      			return;
      		}
      			
      		this.datosProyecto=this.appService.crearProyecto(result);

      		if(this.datosProyecto){
				this.listarProyectos();
			}else{
				console.log("Se ha producido un error en la carga del Proyecto")
				this.nombreProyectoActivo= "Selecciona un proyecto";
				this.proyectoActivo= false
			}
    	});
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
}




