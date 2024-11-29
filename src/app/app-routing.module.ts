
import { NgModule } from '@angular/core';
import { Routes, RouterModule } from '@angular/router';
import {AppComponent} from './app.component';

//Declaración de componentes:
import { IndexComponent } from './comun/index/index.component';
import { DashboardComponent } from './comun/dashboard/dashboard.component';
import { EditorPrograma} from './comun/editorPrograma/editor.component';
import { EditorDocumentoComponent} from './comun/editorDocumento/editorDocumento.component';

const routes: Routes = [
	{path: '' ,redirectTo: "/index", pathMatch: "full" },
	{path: 'index',component: IndexComponent},
	//{path: 'dashboard', redirectTo: "/dashboard", pathMatch: "full"},
	{path: 'dashboard', component: DashboardComponent},
	{path: 'editor', redirectTo: "/editor", pathMatch: "full"},
	{path: 'editor', component: EditorPrograma},
	{path: 'documento', redirectTo: "/documento", pathMatch: "full"},
	{path: 'documento', component: EditorDocumentoComponent}
	];

@NgModule({
  imports: [RouterModule.forRoot(routes)],
  exports: [RouterModule]
})

export class AppRoutingModule { }
