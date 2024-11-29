
import { Component , Input } from '@angular/core';

@Component({
  selector: 'plantillaComponent',
  templateUrl: './plantilla.component.html',
  styleUrls: ['./plantilla.component.sass']
})

export class PlantillaComponent {

	@Input() texto: string; 

	constructor() {}

}





