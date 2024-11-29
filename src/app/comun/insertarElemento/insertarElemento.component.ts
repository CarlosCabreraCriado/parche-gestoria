import { Component , Inject, ViewChild, OnInit} from '@angular/core';
import { MatDialog, MatDialogRef, MAT_DIALOG_DATA} from '@angular/material/dialog';
import { UntypedFormBuilder, UntypedFormGroup, UntypedFormControl, FormArray, Validators} from '@angular/forms'; 
import { MatStepper } from '@angular/material/stepper';
import { DialogoComponent } from '../dialogos/dialogos.component';
import { AppService } from '../../app.service';
import * as XLSX from 'xlsx';

import { MatDialogModule } from '@angular/material/dialog';
import { MatTableModule} from '@angular/material/table';
import { MatFormFieldModule } from '@angular/material/form-field';
import { MatStepperModule } from '@angular/material/stepper';
import { MatSidenavModule } from '@angular/material/sidenav';
import { MatIconModule } from '@angular/material/icon';
import { MatDatepickerModule } from '@angular/material/datepicker';
import { MatExpansionModule } from '@angular/material/expansion';
import { MatListModule } from '@angular/material/list';
import { FormsModule, ReactiveFormsModule } from '@angular/forms';
import { MatButtonToggleModule } from '@angular/material/button-toggle';
import { MatButtonModule } from '@angular/material/button';
import { CommonModule } from '@angular/common';

export interface InsertarElementoData {
    opciones: any;
    data: any;
    contenido: any;
}

@Component({
  standalone: true,
  imports: [
      MatButtonModule,
      MatButtonToggleModule,
      FormsModule,
      ReactiveFormsModule,
      MatListModule,
      MatDialogModule, 
      MatTableModule, 
      CommonModule, 
      MatFormFieldModule,
      MatStepperModule, 
      MatSidenavModule,
      MatIconModule,
      MatExpansionModule,
      MatDatepickerModule
  ],
  selector: 'insertarElemento',
  templateUrl: './insertarElemento.component.html',
  styleUrls: ['./insertarElemento.component.sass']
})

export class InsertarElementoComponent implements OnInit{

    constructor(private appService: AppService, public dialogRef: MatDialogRef<InsertarElementoComponent>, @Inject(MAT_DIALOG_DATA) public data: InsertarElementoData, public formBuilder: UntypedFormBuilder, private dialog: MatDialog)  {

        //INICIALIZACION DE FORMULARIO: Texto
        this.formularioTextoGroup = formBuilder.group({
            formularioTextoControl: this.formularioTextoControl,
            formularioTextoSizeControl: this.formularioTextoSizeControl
        })
        
        //INICIALIZACION DE FORMULARIO: Tabla
        this.formularioTablaGroup = formBuilder.group({
            formularioObjetoTablaControl: this.formularioObjetoTablaControl,
        })
    }

    //Declaracion de Stepper:
    @ViewChild('Stepper',{static: false}) private stepper: MatStepper;

    //Variables Generales: 
    private elementosRender: any;
    public valorTexto: string= "";
    public formatoTexto: any= {};
    public formatoTabla: any= {};
    public cabecerasTabla: string[]= [];
    public cabecerasTablaTemporal: string[] = [];
    public objetosColeccion: any[] = [];
    private objetoImportado: any[] = [];
    public datosTabla: any[] = [];
    private objetoTablaSeleccionado: any = {};

    private posicionElementoPosicionado= {top: 100,left: 0};
    public estadoDrawer: string = "abierto";

    //DECLARACION DE FORMULARIOS:
    
    //Formulario: TEXTO
    public formularioTextoGroup: UntypedFormGroup;
    public formularioTextoControl = new UntypedFormControl({value: 'Texto', disabled: false});
    public formularioTextoSizeControl = new UntypedFormControl({value: 16, disabled: false});

    //Formulario: TABLA
    public formularioTablaGroup: UntypedFormGroup;
    public formularioObjetoTablaControl = new UntypedFormControl({value: '', disabled: true});

    ngOnInit() {

        //Inicializar posicionador:
        this.data.opciones.posicionador= false;

        //Get documento Elementos:  
        console.log("Abriendo Insertar Elemento");
        console.log(this.data);
        this.elementosRender = this.data.contenido;

        //Inicialización segun tipo de elemento:
        switch(this.data.opciones.tipo){
            case "texto":
                this.formatoTexto= {
                    "color": "black",
                    "text-align": "center",
                    "font-weight": "normal",
                    "font-style": "normal",
                    "text-decoration": "none"
                }
                break;
        }

    }

    cambiarAlineamientoTexto(alineamiento:string){
        this.formatoTexto["justify-content"]= alineamiento; 
    }
    cambiarColorTexto(color:string){
        this.formatoTexto.color= color; 
    }

    cambiarFontSize(){
        var fontSize= this.formularioTextoGroup.getRawValue();
        this.formatoTexto["font-size"]= fontSize.formularioTextoSizeControl+"px";
    }

    cambiarFormatoTexto(formato:string){
        switch(formato){
            case "bold":
                if(this.formatoTexto["font-weight"]=="normal"){
                    this.formatoTexto["font-weight"]= "bold";
                }else{
                    this.formatoTexto["font-weight"]= "normal";
                }
                break;
            case "italic":
                if(this.formatoTexto["font-style"]=="normal"){
                    this.formatoTexto["font-style"]= "underline";
                }else{
                    this.formatoTexto["font-style"]= "normal";
                }
                break;
            case "underline":
                if(this.formatoTexto["text-decoration"]=="none"){
                    this.formatoTexto["text-decoration"]= "underline";
                }else{
                    this.formatoTexto["text-decoration"]= "none";
                }
                break;
        }
    }

    avanzarStepper(parametro?: any){

        console.log("Avanzando Stepper: ");
        console.log("Index: "+ this.stepper.selectedIndex);
        if(parametro){
            console.log("Parametros: ");
            console.log(parametro);
        }   

        var dialogoProcesando = this.dialog.open(DialogoComponent,{ disableClose: true,
              data: {tipoDialogo: "procesando", titulo: "Procesando", contenido: ""}
        });
        
        //Verificacion de ruta de archivo:
        switch(this.stepper.selectedIndex){

            //Incluir Colección:
            case 0:
                dialogoProcesando.afterClosed().subscribe(result => {
                    if(result){
                        this.stepper.next();
                    }
                });

                this.incluirColeccion(dialogoProcesando);
                
                break;

            //Selección de Objeto: 
            case 1:
                if(parametro.selectedOptions.selected.length<=0){
                    this.dialog.open(DialogoComponent,{ disableClose: true,
                          data: {tipoDialogo: "informativo", titulo: "Debe seleccionar un objeto", contenido: "No ha seleccinado ningun objeto para la importación."}
                      });
                return;}

                dialogoProcesando.afterClosed().subscribe(result => {
                    if(result){
                        this.stepper.next();
                    }
                });

                this.objetoTablaSeleccionado = parametro.selectedOptions.selected[0]

                console.log("Objeto seleccionado: "+this.objetoTablaSeleccionado);
                console.log(this.objetoTablaSeleccionado);

                this.incluirObjeto(this.objetoTablaSeleccionado.value, dialogoProcesando);
                break;

            //Selección de Cabecera: 
            case 2:
                if(parametro.selectedOptions.selected.length<=0){
                    this.dialog.open(DialogoComponent,{ disableClose: true,
                          data: {tipoDialogo: "informativo", titulo: "Debe seleccionar por lo menos una", contenido: "No ha seleccinado ninguna cabecera."}
                      });
                return;}

                //Construyendo cabeceras seleccionadas
                for(var i=0; i<parametro.selectedOptions.selected.length; i++){
                    this.cabecerasTabla.push(parametro.selectedOptions.selected[i].value);
                }

                dialogoProcesando.afterClosed().subscribe(result => {
                    if(result){
                        this.stepper.next();
                    }
                });

                console.log("Cabeceras tabla seleccionadas: ");
                console.log(this.cabecerasTabla);

                this.construirTabla(dialogoProcesando);
                
                break;
        } //Fin switch

    }

    incluirColeccion(dialogoProcesando?: any) {
        try{

            this.objetosColeccion= this.appService.getListaObjetosEnColeccion(this.data.contenido.direccion, this.data.contenido.nombre)

            console.log(this.objetosColeccion);
        }catch(err){
            const dialogoError = this.dialog.open(DialogoComponent,{ disableClose: true,
                  data: {tipoDialogo: "error", titulo: "Se ha producido un error", contenido: "No se ha podido realizar la importación del objeto."}
              });
            dialogoError.afterClosed().subscribe(result => {
                dialogoProcesando.close(false);
            });
            return;
        }
        if(dialogoProcesando){
            dialogoProcesando.close(true);
        }
        return true;
  }


    incluirObjeto(objeto: any, dialogoProcesando?){
        try{
            this.objetoImportado = this.appService.getObjetoEnColeccion(this.data.contenido.direccion, this.data.contenido.nombre, objeto.objetoId);
        }catch(err){
            console.log(err);
            const dialogoError = this.dialog.open(DialogoComponent,{ disableClose: true,
                  data: {tipoDialogo: "error", titulo: "Se ha producido un error", contenido: "No se ha podido realizar la importación del objeto."}
              });
            dialogoError.afterClosed().subscribe(result => {
                dialogoProcesando.close(false);
            });
            return;
        }

        this.extraerCabecerasObjeto(this.objetoImportado[0]["data"]);

        if(dialogoProcesando){
            dialogoProcesando.close(true);
        }

        return true;
    }

    extraerCabecerasObjeto(objeto:any){
            
        //Iterar por todos los registros del objeto:
        this.cabecerasTablaTemporal=[];

        for(var i=0; i<objeto.length; i++){
            Object.getOwnPropertyNames(objeto[i]).forEach((val)=>{
                if(this.cabecerasTablaTemporal.indexOf(val)==-1){
                    this.cabecerasTablaTemporal.push(val);
                }       
            })
        }
        console.log("Cabeceras: ");
        console.log(this.cabecerasTablaTemporal);
        return;
    }

    construirTabla(dialogoProcesando?){

        this.datosTabla = [];
        for(var i=0; i< this.objetoImportado[0]["data"].length; i++){
            this.datosTabla.push({});
            for(var j=0; j<this.cabecerasTabla.length; j++){
                this.datosTabla[this.datosTabla.length-1][this.cabecerasTabla[j]]= this.objetoImportado[0]["data"][i][this.cabecerasTabla[j]];  
            }
        }
        console.log("Objeto tabla:");
        console.log(this.datosTabla);

        if(dialogoProcesando){
            dialogoProcesando.close(true);
        }

        return true;
    }

    posicionarElemento(tipo: string){

        this.data.opciones.posicionador = true;
        this.dialogRef.updatePosition({top: "0"})

        switch(tipo){
            case "texto":
                var valorTexto= this.formularioTextoGroup.getRawValue();
                this.valorTexto= valorTexto.formularioTextoControl;
                break;

        }

        if(this.data.opciones.estadoDrawer){
            this.estadoDrawer = "abierto";
        }else{
            this.estadoDrawer = "cerrado";
        }

    }
    
    onDragEnded(event) {
        let element = event.source.getRootElement();
        let boundingClientRect = element.getBoundingClientRect();
        let parentPosition = this.getCoordenadas(element);
        console.log('x: ' + (boundingClientRect.x - parentPosition.left), 'y: ' + (boundingClientRect.y - parentPosition.top));
        this.posicionElementoPosicionado= {
            left: (boundingClientRect.x - parentPosition.left),
            top: (boundingClientRect.y - parentPosition.top)
        }
        console.log(this.posicionElementoPosicionado);
      }

    getCoordenadas(el) {
        let x = 0;
        let y = 0;
        while(el && !isNaN(el.offsetLeft) && !isNaN(el.offsetTop)) {
          x += el.offsetLeft - el.scrollLeft;
          y += el.offsetTop - el.scrollTop;
          el = el.offsetParent;
        }
        return { top: y, left: x };
    }

    seleccionarObjeto(indexControl: number){

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
                        //this.incluirObjeto(result);
                        console.log(result);
                        this.data.contenido= result;    
                        this.formularioObjetoTablaControl.setValue(this.data.contenido.nombre)
                    }

                console.log('Fin dialogo seleccionar objeto: ');
                console.log(result);
                return;
            });

        })
    }

    insertarElemento(){
        //Extraer parametros:
        console.log("Argumentos proceso: ");
        var argumentos = this.formularioTextoGroup.getRawValue();
        console.log(argumentos);

        //Obtener Posicion:
        console.log("Posicion")
        console.log(this.posicionElementoPosicionado)

        switch(this.data.opciones.tipo){
            case "texto":
                this.elementosRender.elementos.push( {
                    tipo: "texto",
                    valor: argumentos.formularioTextoControl,
                    estilo: {
                        "font-size": argumentos.formularioTextoSizeControl+"px",
                        top: this.posicionElementoPosicionado.top+"px",
                        left: this.posicionElementoPosicionado.left+"px",
                    }
                })
                break;
                case "tabla":
                    this.elementosRender.elementos.push( {
                        tipo: "tabla",
                        valor: this.datosTabla,
                        cabeceras: this.cabecerasTabla,
                        estilo: {
                            top: this.posicionElementoPosicionado.top+"px",
                            left: this.posicionElementoPosicionado.left+"px",
                        }
                    })
                    break;
        }

        this.appService.guardarDocumentoElementos(this.elementosRender)
        this.dialogRef.close(true);
        return;
    }

    importarSpool(){
        console.warn("Funcion no implementada");
    }
    
}





