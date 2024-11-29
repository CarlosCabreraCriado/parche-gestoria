
import { Component , Inject, ViewChild} from '@angular/core';
import {MatDialog, MatDialogRef, MAT_DIALOG_DATA} from '@angular/material/dialog';
import { JsonEditorComponent, JsonEditorOptions } from 'ang-jsoneditor';

export interface DialogData {
  tipoDialogo: string;
  data: any;
}

@Component({
  selector: 'configuracion-component',
  templateUrl: './configuracion.component.html',
  styleUrls: ['./configuracion.component.sass']
})

export class ConfiguracionComponent {

  public editorVerOptions: JsonEditorOptions;
  public editorModificarOptions: JsonEditorOptions;

  @ViewChild(JsonEditorComponent, { static: true }) editor: JsonEditorComponent;

	constructor(public dialogRef: MatDialogRef<ConfiguracionComponent>, @Inject(MAT_DIALOG_DATA) public data: DialogData) {

    this.editorVerOptions = new JsonEditorOptions()
    this.editorModificarOptions = new JsonEditorOptions()
    this.editorModificarOptions.mode = 'tree'; // set all allowed modes
    this.editorVerOptions.mode = 'view'; // set all allowed modes

  }

  	onNoClick(): void {
    	this.dialogRef.close();
  	}
}





