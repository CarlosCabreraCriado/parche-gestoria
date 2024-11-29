import { Component , Inject, ViewChild, OnInit} from '@angular/core';
import { MatDialog, MatDialogRef, MAT_DIALOG_DATA} from '@angular/material/dialog';
import { UntypedFormBuilder, UntypedFormGroup, UntypedFormControl, Validators} from '@angular/forms'; 
import { MatStepper } from '@angular/material/stepper';
import { DialogoComponent } from '../dialogos/dialogos.component';
import { AppService } from '../../app.service';
//import {listCommands} from 'docx-templates';
import * as XLSX from 'xlsx';

//Modulos:
import { MatFormFieldModule } from '@angular/material/form-field';
import { FormsModule, ReactiveFormsModule } from '@angular/forms';
import { MatStepperModule} from '@angular/material/stepper';
import { MatDialogModule } from '@angular/material/dialog';
import { MatExpansionModule } from '@angular/material/expansion';
import { MatListModule} from '@angular/material/list';
import { MatButtonToggleModule } from '@angular/material/button-toggle';

import { Buffer } from 'buffer/';

export interface AddDatoData {
    opciones: any;
    data: any;
}

@Component({
  standalone: true,
  imports: [
      MatButtonToggleModule,
      MatListModule,
      MatDialogModule,
      MatExpansionModule,
      MatFormFieldModule,
      FormsModule,
      ReactiveFormsModule,
      MatStepperModule
  ],
  selector: 'addPlantilla',
  templateUrl: './addPlantilla.component.html',
  styleUrls: ['./addPlantilla.component.sass']
})

export class AddPlantillaComponent implements OnInit{

    public camposPlantilla = [];
    public hojasSeleccionadasTemporal = [];
    private targetFiles: any= [];
    public addParametro: string = "";

    @ViewChild('rawLibreStepper',{static: false}) private rawLibreStepper: MatStepper;

    public rutaArchivoGroup: UntypedFormGroup;
    public hojasArchivoGroup: UntypedFormGroup;
    public guardadoArchivoGroup: UntypedFormGroup;
    public opcionesArchivoGroup: UntypedFormGroup;

    rutaArchivoControl = new UntypedFormControl("");
    guardadoArchivoControl = new UntypedFormControl("");
    nombreGuardadoArchivoControl = new UntypedFormControl("");
    rutaGuardadoArchivoControl = new UntypedFormControl("");

    constructor(private appService: AppService, public dialogRef: MatDialogRef<AddPlantillaComponent>, @Inject(MAT_DIALOG_DATA) public data: AddDatoData, public formBuilder: UntypedFormBuilder, private dialog: MatDialog)  {

        //DEFINICIONES DE FORMULARIO:
        this.rutaArchivoGroup = formBuilder.group({
          rutaArchivoControl: this.rutaArchivoControl
        });

        this.hojasArchivoGroup = formBuilder.group({
        });

        this.guardadoArchivoGroup = formBuilder.group({
          rutaGuardadoArchivoControl: this.rutaGuardadoArchivoControl,
          nombreGuardadoArchivoControl: this.nombreGuardadoArchivoControl
        });
    }

    ngOnInit() {
        /*
      this.rutaArchivoGroup = this.formBuilder.group({
        firstCtrl: ['', Validators.required]
      });

      this.datosArchivoGroup = this.formBuilder.group({
        secondCtrl: ['', Validators.required]
      });

      this.guardadoArchivoGroup = this.formBuilder.group({
        secondCtrl: ['', Validators.required]
      });
     */

    //Subscripciones de formularios:
        this.rutaArchivoGroup.valueChanges.subscribe((val) =>{
          //this.reportAMService.formularioNacho = val;
          console.log("Formulario Ruta Archivo:")
          console.log(val)
          console.log(this.rutaArchivoGroup);
        });
    }

    goBack(){
        this.rawLibreStepper.previous();
    }
    
    pulsarTecla(event){
        if (event.keyCode === 13) {
            if(this.addParametro!= ""){
                this.camposPlantilla.push({
                    type: "INS",
                    code: this.addParametro
                })
            }   
        }
        return;
    }

    incluirRuta(evt: any, nombreControl: string){
        //Lectura de evento Input
        const target: DataTransfer = <DataTransfer>(evt.target);

        console.log("Objeto Ruta:");
        console.log(target.files);
        var formularioTemporal: any;
        switch(nombreControl){
            case "rutaArchivoControl":
                formularioTemporal= this.rutaArchivoGroup.value;
                formularioTemporal.rutaArchivoControl= target.files[0]["path"];
                this.addTargetFile("rutaArchivoControl", target.files[0]);
                this.rutaArchivoGroup.setValue(formularioTemporal);
                break;
        }
    }   

    addTargetFile(control: string, target: Blob){
        if(this.targetFiles.find(i=>i.control)==-1 || this.targetFiles.find(i=>i.control)==undefined){
            this.targetFiles.push({control: control, target: target});
            return;
        }else{
            this.targetFiles.find(i=>i.control).target= target; 
            return;
        }
    }

    eliminarTargetFile(control: string){
        if(this.targetFiles.find(i=>i.control)==-1 || this.targetFiles.find(i=>i.control)==undefined){
            return;
        }else{
            this.targetFiles= this.targetFiles.splice(this.targetFiles.indexOf(this.targetFiles.find(i=>i.control)),1);
            return;
        }
    }

    avanzarStepper(){
        
        //Verificacion de ruta de archivo:
        if(this.rawLibreStepper.selectedIndex==0){
            const dialogoProcesando = this.dialog.open(DialogoComponent,{ disableClose: true,
                  data: {tipoDialogo: "procesando", titulo: "Procesando", contenido: ""}
              });
            dialogoProcesando.afterClosed().subscribe(result => {
                this.rawLibreStepper.next();
            });
            this.incluirArchivo("rutaArchivoControl",dialogoProcesando);
        }
    }

    incluirArchivo(controlArg: string, dialogoProcesando?: any) {

        //Verifica el parametro de control:
        try{
            if(this.targetFiles.find(i => i.control==controlArg).target == undefined) throw new Error('No se pueden seleccionar varios archivos');
        }catch(err){
            console.log("Error incluyendo el archivo");
            return;
        }

      const reader: FileReader = new FileReader();
      reader.onload = (e: any) => {
        
        const template_buffer: Buffer = e.target.result;

        /* Leer Plantilla */

        /*
        listCommands(template_buffer, ['{', '}']).then((result)=>{
                
        this.camposPlantilla =  result;
        this.camposPlantilla['path'] = this.targetFiles.find(i=> i.control==controlArg).target.path;
        console.log(this.camposPlantilla)

        //Configuracion de cabeceras:

        console.log("Archivo includo con exito");
        if(dialogoProcesando){
            dialogoProcesando.close();
        }
        return true;
        });
        */
      };
      reader.readAsBinaryString(this.targetFiles.find(i=> i.control==controlArg).target);
  }

    seleccionCampos(hojas: any){
        if(hojas.selectedOptions.selected.length<=0){
            const dialogoProcesando = this.dialog.open(DialogoComponent,{ disableClose: true,
                  data: {tipoDialogo: "informativo", titulo: "Hojas insuficientes", contenido: "Debe seleccionar alguna hoja para poder realizar la importación."}
              });
        return;}
        this.hojasSeleccionadasTemporal= [];
        for(var i= 0; i<hojas.selectedOptions.selected.length; i++){
            this.hojasSeleccionadasTemporal.push({campo: hojas.selectedOptions.selected[i]._element.nativeElement.innerText, opciones: {tipo: "texto",descripcion: ""}});
            
        }   
        this.rawLibreStepper.next();
    }

    seleccionOpciones(){
        console.log(this.hojasSeleccionadasTemporal);
        this.rawLibreStepper.next();
    }

    guardarArchivoRaw(){

        var objetoPlantilla= {};

        objetoPlantilla["data"]= this.hojasSeleccionadasTemporal;
        objetoPlantilla["path"]= this.camposPlantilla['path']; 

        //Incluir nombreID:
        objetoPlantilla["nombreId"]= this.nombreGuardadoArchivoControl.value; 
        objetoPlantilla["objetoId"]= this.nombreGuardadoArchivoControl.value;

        console.log("Archivo Verificado: ");
        console.log(objetoPlantilla);

        const dialogoProcesando = this.dialog.open(DialogoComponent,{ disableClose: true,
              data: {tipoDialogo: "procesando", titulo: "Procesando", contenido: ""}
          });

        dialogoProcesando.afterClosed().subscribe(result => {
            if(result==true){
                console.log("Archivo guardando con exito");
                this.dialogRef.close("exito");
            }else{
                console.log("Error guardando archivo");
                this.dialogRef.close("error");
            }
        });

        this.appService.guardarPlantilla(objetoPlantilla, dialogoProcesando);
/*
      if(this.appService.guardarDocumento(objetoArchivo)){
        console.log("Archivo guardado con exito");
      }else{
        console.log("Error guardando el archivo");
      } 
     */
      return;
      
    } //Fin guardar Archivo Raw
    
/*
    subirArchivo(){
      
      this.documentosHerramienta[this.configuracion.render.indexArchivoSeleccionado].estado="subido";
      this.log("Guardando archivo...","orange");
      this.mostrarMensaje= true;
      this.mostrarSpinner= true;
      this.mensaje= "Guardando archivo: "+this.configuracion.render.nombreArchivoSeleccionado;

      if(this.appService.guardarDocumento(this.documentosHerramienta[this.configuracion.render.indexArchivoSeleccionado].objetoArchivo)){
        this.log("Archivo guardado con exito.","green");
        this.mostrarMensaje= false;
        //this.reloadDatos(false);
      }else{
        this.log("Error guardando el archivo","red");
        this.mostrarMensaje= false;
      } 
    }*/

    importarSpool(){
        console.log("Importarndo Spool: ");
        console.log(this.rutaArchivoControl.value);
        console.log(this.nombreGuardadoArchivoControl.value);
        
        this.appService.importarSpool(this.rutaArchivoControl.value, this.nombreGuardadoArchivoControl.value).then((result)=>{
            console.log("Spool importado:");
            console.log(result);
        })
    }

    onNoClick(): void {
        this.dialogRef.close();
    }
}




