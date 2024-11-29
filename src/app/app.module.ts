import { BrowserModule } from "@angular/platform-browser";
import { NgModule } from "@angular/core";


import { NgJsonEditorModule } from "ang-jsoneditor";

import { HttpClientModule } from "@angular/common/http";

import { AppRoutingModule } from "./app-routing.module";
import { AppComponent } from "./app.component";

import { MAT_DATE_LOCALE } from '@angular/material/core'

//Materials:
import { DemoMaterialModule } from "./material-module";
import { FormsModule, ReactiveFormsModule } from "@angular/forms";
import { MatFormFieldModule } from "@angular/material/form-field";
import { MatNativeDateModule} from '@angular/material/core';
import { MatDatepickerModule} from '@angular/material/datepicker';

import { MatIconModule } from '@angular/material/icon'


//Declaración de componentes:
import { IndexComponent } from "./comun/index/index.component";
import { DashboardComponent } from "./comun/dashboard/dashboard.component";
import { DialogoComponent } from "./comun/dialogos/dialogos.component";
import { AddDatoComponent } from "./comun/addDato/addDato.component";
import { InsertarElementoComponent } from "./comun/insertarElemento/insertarElemento.component";
import { EjecutarProcesoComponent } from "./comun/ejecutarProceso/ejecutarProceso.component";
import { VisualizarDato } from "./comun/visualizarDato/visualizarDato.component";
import { GestionarDato } from "./comun/gestionarDato/gestionarDato.component";
import { ConfiguracionComponent } from "./comun/configuracion/configuracion.component";
import { SplashComponent } from "./comun/splash/splash.component";
import { EditorPrograma } from "./comun/editorPrograma/editor.component";
import { SeleccionarProgramaComponent } from "./comun/seleccionarPrograma/seleccionarPrograma.component";
import { AddProcesoComponent } from "./comun/addProceso/addProceso.component";
import { BrowserAnimationsModule } from "@angular/platform-browser/animations";
import { EditorDocumentoComponent } from "./comun/editorDocumento/editorDocumento.component";
import { AddPlantillaComponent } from "./comun/addPlantilla/addPlantilla.component";
import { AddCursoComponent } from "./comun/addCurso/addCurso.component";


import {MatDialogModule, MatDialog, MatDialogRef, MAT_DIALOG_DATA} from '@angular/material/dialog';

@NgModule({
    declarations: [
        AppComponent,
        ConfiguracionComponent,
        SplashComponent,
    ],
    imports: [
        MatDialogModule,
        NgJsonEditorModule,
        BrowserModule,
        AppRoutingModule,
        HttpClientModule,
        NgJsonEditorModule,
        BrowserAnimationsModule,
        DemoMaterialModule,
        MatIconModule,
        FormsModule,
        MatFormFieldModule,
        ReactiveFormsModule,
        MatFormFieldModule,
        MatDatepickerModule,
        MatNativeDateModule
    ],
    providers: [{ provide: MAT_DATE_LOCALE, useValue: 'es-ES' }],
    bootstrap: [AppComponent]
})

export class AppModule {}



