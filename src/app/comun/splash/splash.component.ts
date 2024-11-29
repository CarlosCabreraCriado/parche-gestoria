
import { Component, OnInit, ElementRef } from '@angular/core';
import { Router, ActivatedRoute } from '@angular/router';
import { AppService } from '../../app.service';
import { SplashService } from './splash.service';

@Component({
  selector: 'app-splash',
  templateUrl: './splash.component.html',
  styleUrls: ['./splash.component.sass']
})

export class SplashComponent implements OnInit{

	constructor(public appService: AppService, public SplashService: SplashService) { }

	public mostrarSplash= true;
	public tiempoSplash= 1000;
	public version= "0.0.0";

	ngOnInit(){

		setTimeout(()=>{  
 					this.mostrarSplash = false;
 			}, this.tiempoSplash);	

		this.version= this.appService.getVersion();
	}
	
}




