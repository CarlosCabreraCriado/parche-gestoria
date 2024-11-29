import { Component, OnInit, ViewChild,ElementRef} from '@angular/core';
import { AppService } from '../../app.service';
import { MatDialog} from '@angular/material/dialog';
import { DialogoComponent } from '../dialogos/dialogos.component'
import { MatSidenav} from '@angular/material/sidenav';
import { UntypedFormBuilder, UntypedFormGroup, UntypedFormControl, Validators, UntypedFormArray} from '@angular/forms'; 
import { FlatTreeControl } from '@angular/cdk/tree';
import { MatTreeFlatDataSource, MatTreeFlattener } from '@angular/material/tree';
import { LibreriaPlantillas,libreriaPlantillas } from '../../../../plantillas/plantillas.configuracion';
import { EjecutarProcesoComponent} from '../ejecutarProceso/ejecutarProceso.component'

//Modulos:
import { MatSelectModule} from '@angular/material/select';
import { FormsModule, ReactiveFormsModule } from '@angular/forms';
import { MatDatepickerModule } from '@angular/material/datepicker';
import { MatIconModule } from '@angular/material/icon';
import { MatDialogModule } from '@angular/material/dialog';
import { MatInputModule } from '@angular/material/input';
    
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
      MatDialogModule,
      MatIconModule,
      MatSelectModule,
      FormsModule,
      MatInputModule,
      ReactiveFormsModule,
      MatDatepickerModule
  ],
  selector: 'app-editorDocumento',
  templateUrl: './editorDocumento.component.html',
  styleUrls: ['./editorDocumento.component.sass']
})

export class EditorDocumentoComponent implements OnInit{

    constructor(public appService: AppService, public dialog: MatDialog ,public formBuilder: UntypedFormBuilder) {
        //Inicialización del arbol de procesos: 
        this.arbolProcesosDataSource.data = this.datosArbolProcesos;

        this.formularioProcesoGroup = formBuilder.group({
            formularioPruebaControl: this.formularioPruebaControl
        })

        this.formularioPlantillaGroup = formBuilder.group({
            formularioPlantillaControl: this.formularioPlantillaControl
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
    public procesoSeleccionado: any= null;
    private objetosSeleccionados: objetosSeleccionados[] = []
    private addParametro: string= "";

    //Control de formularios: 
    public formularioProcesoGroup: UntypedFormGroup;
    public formularioPlantillaGroup: UntypedFormGroup;

    public formularioControl  = new UntypedFormArray([]) 
    public formularioPlantillaControl = new UntypedFormControl({value: '', disabled: true})
    private formularioPruebaControl  = new UntypedFormControl("") 
    public proyectoActivo = false;
    public nombreProyectoActivo = "Selecciona un proyecto";
    public datosProyecto: any = {};
    private editor: any;

    public correo: any = [];
    public indexCorreo: number= 0;

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

    pulsarTecla(event){
        if (event.keyCode === 13) {
            alert('you just pressed the enter key');
        }
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

    siguienteCorreo(){
        if(this.indexCorreo < this.correo.length-1){
            this.indexCorreo++;
        }
        console.log("Correo: "+this.indexCorreo)
    }

    anteriorCorreo(){
        if(this.indexCorreo > 0){
            this.indexCorreo--;
        }
        console.log("Correo: "+this.indexCorreo)
    }

    incluirRuta(evt: any, indexControl: number){
        //Lectura de evento Input
        const target: DataTransfer = <DataTransfer>(evt.target);

        console.log("Objeto Ruta:");
        console.log(target.files);

        var formularioTemporal: any;
        formularioTemporal= this.formularioProcesoGroup.getRawValue();
        formularioTemporal.formularioControl[indexControl]= target.files[0]["path"];
        this.formularioProcesoGroup.setValue(formularioTemporal);
    }   

    ejecutarPlantilla(proceso: LibreriaPlantillas){

        //Extraer parametros:
        console.log("Argumentos proceso: ");
        var argumentos = this.formularioProcesoGroup.getRawValue();

        //Extraccion de los objetos:
        var objeto = {}
        for(var i=0; i<this.objetosSeleccionados.length; i++){
            objeto = this.appService.getDato(this.objetosSeleccionados[i].node.direccion, this.objetosSeleccionados[i].node.nombre) 
            argumentos.formularioControl[this.objetosSeleccionados[i].indexControl] = objeto;
        }

        console.log(proceso)
        console.log(argumentos);

        const dialogoProcesando = this.dialog.open(DialogoComponent,{ disableClose: true,
              data: {tipoDialogo: "procesando", titulo: "Procesando", contenido: ""}
          });

        this.appService.ejecutarPlantilla(proceso, argumentos).then((result)=>{
            console.log("Proceso finalizado: ");
            console.log(result);
            dialogoProcesando.close();
            if(!result){
                this.dialog.open(DialogoComponent,{ disableClose: true,
                      data: {tipoDialogo: "error", titulo: "Se ha producido un error inesperado.", contenido: ""}
                  });
            }
        })

    }

    incluirPlantilla(objeto: any){

        console.log(objeto)
        var indexControl = 0;

        //Asignacion de array de objetos seleccionados: 
        var objetoEncontrado = false; 
        for(var i=0; i<this.objetosSeleccionados.length; i++){
            if(this.objetosSeleccionados[i].indexControl==indexControl){
                objetoEncontrado= true;     
                this.objetosSeleccionados[i]= {
                    indexControl: indexControl,
                    node: objeto
                }
            }
        }

        if(!objetoEncontrado){
            this.objetosSeleccionados.push({
                indexControl: indexControl,
                node: objeto
            })
        }

        var formularioTemporal: any;
        formularioTemporal= this.formularioPlantillaGroup.getRawValue();
        console.log(this.formularioPlantillaGroup)
        console.log(formularioTemporal)
        formularioTemporal.formularioPlantillaControl= objeto.nombre;
        this.formularioPlantillaGroup.setValue(formularioTemporal);

        objeto = this.appService.getDato(objeto.direccion, objeto.nombre)
        objeto = objeto[0]

        console.log(objeto)

        //Contruir Objeto de procesado Documento:
        this.procesoSeleccionado = {
            nombre: objeto.objetoId,
            path: objeto.path,
            argumentos: []
        }

        for(var i = 0; i<objeto.data.length; i++){
            this.procesoSeleccionado.argumentos.push({
                identificador: objeto.data[i].campo,
                tipo: objeto.data[i].opciones.tipo,
                obligado: false,
                formulario: {
                    tipo: objeto.data[i].opciones.tipo,
                    placeholder: objeto.data[i].opciones.descripcion,
                    titulo: objeto.data[i].campo,
                    valorDefault: ""
                }
            })
        }

        this.formularioControl = new UntypedFormArray([]);

        //DEFINICIONES DE FORMULARIO:
        for(var i=0; i< this.procesoSeleccionado.argumentos.length; i++){
                this.formularioControl.push(new UntypedFormControl({value: this.procesoSeleccionado.argumentos[i].formulario.valorDefault, disabled: false}));
        }
        
        this.formularioProcesoGroup = this.formBuilder.group({
          formularioControl: this.formularioControl
        });

        console.log(this.procesoSeleccionado)
        return true;
        
    }   

    seleccionarPlantilla(){

        this.appService.getArbolProyecto().then((result)=>{

            const dialogoSeleccionarObjeto = this.dialog.open(DialogoComponent,{ disableClose: false,
                  data: {tipoDialogo: "seleccionarObjeto", titulo: "Seleccione un objeto", contenido: result}
              });

            dialogoSeleccionarObjeto.afterClosed().subscribe(result => {
                if(result=="error"){
                    const dialogExito = this.dialog.open(DialogoComponent,{
                        data: {tipoDialogo: "error", titulo:"Se ha producido un error.",contenido:"No se ha podido realizar la selección."}
                    });
                }else 
                    if(result != undefined && result != false){
                        //Asignando Objeto:
                        this.incluirPlantilla(result);
                    }

                console.log('Fin dialogo seleccionar objeto: ');
                console.log(result);
                return;
            });
        })
    }

    incluirDirectorio(directorio:any){
        console.warn("Funcion no implementada");
        return;
    }

    seleccionarObjeto(objeto:any){
        console.warn("Funcion no implementada");
        return;
    }

}




