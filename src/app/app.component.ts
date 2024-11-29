import { Component, OnInit } from "@angular/core";
import { AppService } from "./app.service";
//import { ElectronService } from "ngx-electron";
import { Subscription } from "rxjs";

@Component({
	selector: "app-root",
	templateUrl: "./app.component.html",
	styleUrls: ["./app.component.sass"]
})

export class AppComponent implements OnInit {
	title = "GestorVF";

	constructor(
		private appService: AppService,
//		private electronService: ElectronService
	){}

	//Declara Suscripcion para Logger:
	private appServiceSuscripcion: Subscription;

	ngOnInit() {
		console.log("Iniciando aplicación: ");
		//Suscripcion AppService:
		this.appServiceSuscripcion = this.appService.observarAppService$.subscribe(
			val => {
				switch (val) {
					case "descargarActualizacion":
						this.appService.descargarActualizacion();
						break;
					case "descargaCompletada":
						this.appService.descargaCompletada();
						break;

					default:
						break;
				}
			}
		); //Fin AppServiceSuscription.
		this.appService.inicializarAppService();
		this.appService.inicializarAutoupdate();
	} //Fin OnInit

} //Fin component
