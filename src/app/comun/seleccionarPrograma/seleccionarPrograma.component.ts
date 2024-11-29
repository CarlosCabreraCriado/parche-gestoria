
import { Component , Input } from '@angular/core';
import { MatDialogModule } from '@angular/material/dialog';
import { DemoMaterialModule } from "../../material-module";

@Component({
  standalone: true,
  imports: [DemoMaterialModule, MatDialogModule],
  selector: 'seleccionarProgramaComponent',
  templateUrl: './seleccionarPrograma.component.html',
  styleUrls: ['./seleccionarPrograma.component.sass']
})

export class SeleccionarProgramaComponent {

	@Input() texto: string; 
    public panelOpenState: boolean = false;

	constructor() {}

}





