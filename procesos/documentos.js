const path = require("path");
const fs = require("fs");
const moment = require("moment");
const {ipcRenderer}= require("electron");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");


class ProcesosGenerales {
    
    constructor(pathToDbFolder, nombreProyecto, proyectoDB){
        this.pathToDbFolder = pathToDbFolder;
        this.nombreProyecto = nombreProyecto; 
        this.proyectoDB = proyectoDB;
    }
    
    async esperar(tiempo){
        return new Promise((resolve)=>{
            setTimeout(resolve, tiempo);
        });
    }

    async generarDocumento(argumentos){

        //ARGUMENTOS:
        // 0: Ruta Plantilla 
        // 1: Parametros de plantilla  
        //argumentos = [__dirname, {first_name: "John", last_name: "Doe", phone: "0652455478", description: "New Website"}]

        console.log("Generando Documento...", argumentos[0]);
        console.log("Parametros:");

        console.log(argumentos[2]);


        // Load the docx file as binary content
        const content = fs.readFileSync(
            path.resolve(argumentos[0]),
            "binary"
        );

        // Unzip the content of the file
        const zip = new PizZip(content);

        // This will parse the template, and will throw an error if the template is
        // invalid, for example, if the template is "{user" (no closing tag)
        const doc = new Docxtemplater(zip, {
            paragraphLoop: true,
            linebreaks: true,
        });

        // Render the document (Replace {first_name} by John, {last_name} by Doe, ...)
        doc.render(argumentos[2]);

        // Get the zip document and generate it as a nodebuffer
        const buf = doc.getZip().generate({
            type: "nodebuffer",
            // compression: DEFLATE adds a compression step.
            // For a 50MB output document, expect 500ms additional CPU time
            compression: "DEFLATE",
        });

        // buf is a nodejs Buffer, you can either write it to a
        // file or res.send it with express for example.
        fs.writeFileSync(path.resolve(argumentos[1], "comunidad_autonoma_output.pptx"), buf);

        return true;
    }


}//Fin Procesos Generales

module.exports = ProcesosGenerales;


