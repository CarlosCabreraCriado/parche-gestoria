import { Component , Inject, ViewChild, OnInit} from '@angular/core';
import { MatDialog, MatDialogRef, MAT_DIALOG_DATA} from '@angular/material/dialog';
import { UntypedFormBuilder, UntypedFormGroup, UntypedFormControl, Validators} from '@angular/forms'; 
import { MatStepper } from '@angular/material/stepper';
import { DialogoComponent } from '../dialogos/dialogos.component';
import { AppService } from '../../app.service';
import * as XLSX from 'xlsx';

//Modulos:
import { FormsModule, ReactiveFormsModule } from '@angular/forms';
import { MatFormFieldModule } from '@angular/material/form-field';
import { MatStepperModule} from '@angular/material/stepper';
import { MatDialogModule } from '@angular/material/dialog';
import { MatIconModule } from '@angular/material/icon';
import { MatExpansionModule } from '@angular/material/expansion';
import { MatListModule} from '@angular/material/list';
import { MatButtonToggleModule } from '@angular/material/button-toggle';

export interface AddDatoData {
    opciones: any;
    data: any;
}

@Component({
  standalone: true,
  imports: [
      MatButtonToggleModule,
      MatListModule,
      MatExpansionModule,
      MatIconModule,
      MatDialogModule, 
      FormsModule,
      MatStepperModule,
      ReactiveFormsModule,
      MatFormFieldModule,
  ],
  selector: 'addDato',
  templateUrl: './addDato.component.html',
  styleUrls: ['./addDato.component.sass']
})

export class AddDatoComponent implements OnInit{

    private workbookTemporal: XLSX.WorkBook;
    public hojasTemporal = [];
    public hojasSeleccionadasTemporal = [];
    public opcionesHojasSeleccionadasTemporal = [];
    private targetFiles: any= [];

    @ViewChild('rawLibreStepper',{static: false}) private rawLibreStepper: MatStepper;

    public rutaArchivoGroup: UntypedFormGroup;
    public hojasArchivoGroup: UntypedFormGroup;
    public guardadoArchivoGroup: UntypedFormGroup;
    public opcionesArchivoGroup: UntypedFormGroup;

    rutaArchivoControl = new UntypedFormControl("");
    guardadoArchivoControl = new UntypedFormControl("");
    nombreGuardadoArchivoControl = new UntypedFormControl("");
    rutaGuardadoArchivoControl = new UntypedFormControl("");

    constructor(private appService: AppService, public dialogRef: MatDialogRef<AddDatoComponent>, @Inject(MAT_DIALOG_DATA) public data: AddDatoData, public formBuilder: UntypedFormBuilder, private dialog: MatDialog)  {

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

      //Lectura de archivo
      const reader: FileReader = new FileReader();
      reader.onload = (e: any) => {
        
        /* Leer workbook */
        const bstr: string = e.target.result;
        const wb: XLSX.WorkBook = XLSX.read(bstr, {type: 'binary', cellStyles: true});

        /* grab first sheet */
        const wsname: string = wb.SheetNames[1];
        const ws: XLSX.WorkSheet = wb.Sheets[wsname];
        this.hojasTemporal = wb.SheetNames;
        console.log(this.hojasTemporal)

        //Configuracion de cabeceras:
        this.workbookTemporal = wb;
        console.log("Archivo includo con exito");
        if(dialogoProcesando){
            dialogoProcesando.close();
        }
        return true;
      };
      reader.readAsBinaryString(this.targetFiles.find(i=> i.control==controlArg).target);
  }

    seleccionHojas(hojas: any){
        if(hojas.selectedOptions.selected.length<=0){
            const dialogoProcesando = this.dialog.open(DialogoComponent,{ disableClose: true,
                  data: {tipoDialogo: "informativo", titulo: "Hojas insuficientes", contenido: "Debe seleccionar alguna hoja para poder realizar la importación."}
              });
        return;}
        this.hojasSeleccionadasTemporal= [];
        for(var i= 0; i<hojas.selectedOptions.selected.length; i++){
            this.hojasSeleccionadasTemporal.push({hoja: hojas.selectedOptions.selected[i]._element.nativeElement.innerText, opciones: {cabecera: 1}});
            
        }   
        this.rawLibreStepper.next();
    }

    seleccionOpciones(){
        console.log(this.hojasSeleccionadasTemporal);
        this.rawLibreStepper.next();
    }

    guardarArchivoRaw(){
        var worksheet: XLSX.WorkSheet;
        var cabecerasHojas= [];
        var hojasArchivo= [];
        var errorGuardado= false;
        var objetoArchivo= {};

        //Creacion del Objeto Archivo:
        for(var i=0; i<this.hojasSeleccionadasTemporal.length; i++){
          
          try{
            //Obtención de hoja:
            worksheet = this.workbookTemporal.Sheets[this.hojasSeleccionadasTemporal[i].hoja];

            //Procesado de datos TIPO FILA:
              cabecerasHojas.push(XLSX.utils.sheet_to_json(worksheet, {header: 1 , range: this.hojasSeleccionadasTemporal[i].opciones.cabecera-1})[0]);
              
              console.log(cabecerasHojas);

              //Formateo de cabeceras (TIPO FILA):
                for(var j= 0; j<cabecerasHojas[i].length;j++){
                    cabecerasHojas[i][j]= cabecerasHojas[i][j].toLowerCase();
                    cabecerasHojas[i][j]= cabecerasHojas[i][j].replace(/ /g,"_");
                    cabecerasHojas[i][j]= cabecerasHojas[i][j].replace(/\./g," ");
                }

              hojasArchivo.push((XLSX.utils.sheet_to_json(worksheet, {header: cabecerasHojas[i],range: this.hojasSeleccionadasTemporal[i].opciones.cabecera })));

            //Añadir objetoId:
            //hojasArchivo[hojasArchivo.length-1]["objetoId"]= this.hojasSeleccionadasTemporal[i].hoja;

            console.log("Cabeceras:");
            console.log(cabecerasHojas);

            console.log("Datos:");
            console.log(hojasArchivo);

          }catch(error){
            console.log("ERROR obteniendo hoja: "+this.hojasSeleccionadasTemporal[i].hoja);
            errorGuardado= true;
            console.log(error);
          }
        }

        if(errorGuardado){
            console.log("Se ha producido un error guardando el archivo");
        return;}else{
            console.log("Verificación superada con exito");
        }

        //CONSTRUCTOR DE ARCHIVO:
        for(var i=0; i<this.hojasSeleccionadasTemporal.length; i++){
            if(i==0){
                objetoArchivo["data"]=hojasArchivo[i];
            }
        }

        //Incluir nombreID:
        objetoArchivo["nombreId"]= this.nombreGuardadoArchivoControl.value; 
        objetoArchivo["objetoId"]= this.hojasSeleccionadasTemporal[0].hoja.toLowerCase().replace(/ /g,"_");

        console.log("Archivo Verificado: ");
        console.log(objetoArchivo);

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

        this.appService.guardarArchivo(objetoArchivo, dialogoProcesando);
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





