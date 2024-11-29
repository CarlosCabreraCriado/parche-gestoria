const path = require("path");
const fs = require("fs").promises;
const readline = require('readline')
const {Base64} = require('js-base64');

const {createMimeMessage} = require('mimetext')
const process = require('process');
const {authenticate} = require('@google-cloud/local-auth');
const {google} = require('googleapis');
const mainProcess = require("../main.js");
const QRCode = require('qrcode');


var codigoGoogle;
var oAuth2Client; //Deprecated??
var forzarToken;

// If modifying these scopes, delete token.json.
const SCOPES = ['https://www.googleapis.com/auth/gmail.readonly','https://www.googleapis.com/auth/gmail.compose'];
// The file token.json stores the user's access and refresh tokens, and is
// created automatically when the authorization flow completes for the first
// time.

class ProcesosGoogle {
    
    constructor(pathToDbFolder, nombreProyecto, proyectoDB){

        this.pathToDbFolder = pathToDbFolder;
        this.nombreProyecto = nombreProyecto; 
        this.proyectoDB = proyectoDB;
        this.forzarToken = false;

        this.TOKEN_PATH = path.join(pathToDbFolder,nombreProyecto,'token.json');
        this.CREDENTIALS_PATH = path.join(pathToDbFolder,nombreProyecto,'credenciales.json');
        this.pathmessage = path.join(this.pathToDbFolder,nombreProyecto,'correo.txt');

        console.log("SET CREDENTIALS_PATH")
        console.log(this.CREDENTIALS_PATH)

    }

    //*********************************************** GOOGLE GMAIL API*********************************************

    setCodigoGoogle(codigo){
        console.log("SET COGIGO GOOGLE");
        this.codigoGoogle= codigo
        return;
    }

    /**
     * Reads previously authorized credentials from the save file.
     *
     * @return {Promise<OAuth2Client|null>}
     */
    async loadSavedTokenIfExist() {
      try {
        if(!this.forzarToken){
            const content = await fs.readFile(this.TOKEN_PATH);
            console.log("TOKEN",content)
            const credentials = JSON.parse(content);
            return google.auth.fromJSON(credentials);
        }else{
            const content = null; //Null para forzar la autentificacion desde navegador.
            this.forzarToken = false;
            const credentials = JSON.parse(content);
            return google.auth.fromJSON(credentials);
        }
      } catch (err) {
        return null;
      }
    }

    /**
     * Serializes credentials to a file compatible with GoogleAUth.fromJSON.
     *
     * @param {OAuth2Client} client
     * @return {Promise<void>}
     */
    async saveCredentials(client) {
      const content = await fs.readFile(this.CREDENTIALS_PATH);
      const keys = JSON.parse(content);
      const key = keys.installed || keys.web;
      const payload = JSON.stringify({
        type: 'authorized_user',
        client_id: key.client_id,
        client_secret: key.client_secret,
        refresh_token: client.credentials.refresh_token,
      });
      await fs.writeFile(this.TOKEN_PATH, payload);
    }

    /**
     * Load or request or authorization to call APIs.
     *
     */
    async authorize() {

      console.log("Autorizando...")
      //Busca el Token si existe y es valido:
      let client = await this.loadSavedTokenIfExist();
      if (client) {
        console.log("Token guardado encontrado.")
        return client;
      }

      console.log("No hay token guardado.")
      //Si tiene credenciales intenta autentificar:
      try{
          client = await authenticate({
            scopes: SCOPES,
            keyfilePath: this.CREDENTIALS_PATH,
          });
      }catch(err){
          console.log("Error al autentificar: No hay archivo de credenciales.")

          return false;
      }

    console.log("CREDENCIALES: ",client.credentials)
      if (client.credentials) {
        await this.saveCredentials(client);
      }

      return client;
    }

    /**
     * Lists the labels in the user's account.
     *
     * @param {google.auth.OAuth2} auth An authorized OAuth2 client.
     */
    async listLabels(auth) {
      const gmail = google.gmail({version: 'v1', auth});
      const res = await gmail.users.labels.list({
        userId: 'me',
      });
      const labels = res.data.labels;
      if (!labels || labels.length === 0) {
        console.log('No labels found.');
        return;
      }
      console.log('Labels:');
      labels.forEach((label) => {
        console.log(`- ${label.name}`);
      });
    }

    //**********************************************FIN GOOGLE GMAIL API*******************************************
    
    async esperar(tiempo){
        return new Promise((resolve)=>{
            setTimeout(resolve, tiempo);
        });
    }

    async obtenerCorreos(argumentos){
        console.log("Ejecutando OBTENER CORREOS");
        console.log("Querry:" + argumentos[0]);
        this.oAuth2Client = await this.authorize();
        if(!this.oAuth2Client){return false;}
        
        try{
            const messages = await this.listarMensajes(this.oAuth2Client, argumentos[0]);
            console.log("Mensajes")
            console.log(messages)
            return messages;
        }catch(err){
            console.log("ERRROR: ",err);
            
        }
    }

    listarMensajes(auth, query) {
      return new Promise((resolve, reject) => {
        const gmail = google.gmail({version: 'v1', auth});
        gmail.users.messages.list({
            userId: 'me',
            q: query,
          },(err, res) => {

            if (err) {
                if(err.response.data.error == "invalid_grant"){
                    console.log("Error de Token... Regenerando token...")
                    this.forzarToken = true;
                    this.authorize();
                    //reject(err);
                    mainProcess.mostrarWarning("Autenificación Necesaria","Completa el proceso de Login en la pestaña emergente para autentificarte. Una vez cumplimentado vuelve a realizar la operación solicitada.").then((result) => {
                    });
                    return false;
                }
                reject(err);
                return;
            }

            if (!res.data.messages) {
                resolve([]);
                return;
            }
            resolve(res.data.messages);
          }
        );
      })
    ;}

    async getCorreosAsunto(argumentos){

        //var validacion = await this.getValidacionGoogle();
        this.oAuth2Client = await this.authorize();
        console.log(this.oAuth2Client);
        console.log("Ejecutando OBTENER CORREOS");
        console.log("Querry:" + argumentos[0]);

        try{
            const mensajes = await this.listarMensajes(this.oAuth2Client, argumentos[0]);
            //Get mensajes:
            const gmail = google.gmail({version: 'v1', auth: this.oAuth2Client});
        
            var contenidoMensajes = [];
            var mensaje;
            var part;

            //Obtención de mensajes:
            console.log("CORREOS ENCONTRADOS: "+mensajes.length);
            for(var i=0; i<mensajes.length;i++){
                //Obtiene mensaje:
                mensaje=await gmail.users.messages.get({
                    id: mensajes[i].id,
                    userId: 'me',
                    //format: 'RAW'
                });

                console.log(mensaje)
                return mensaje.data;

                if(mensaje.data.payload.parts=== undefined){
                    console.log("SIN PARTE:")
                    if(mensaje.data.payload.mimeType=="text/html"){
                        //contenidoMensajes.push(Base64.decode(mensaje.data.payload.body.data.replace(/-/g, '+').replace(/_/g, '/')))
                    }
                    continue;
                }

                return mensaje.data;
                part = mensaje.data.payload.parts.filter(function(part) {
                  return part.mimeType == 'text/html';
                });

                if(part[0]=== undefined){
                    console.log("SIN HTML: ")
                    console.log(mensaje.data.payload.parts)
                    continue;}
                    contenidoMensajes.push(Base64.decode(part[0].body.data.replace(/-/g, '+').replace(/_/g, '/')))
                
            }
            
            var result= {
                tipo: "correo",
                data: contenidoMensajes
            }

        }catch(err){
            console.log("ERROR DE VALIDACIÓN")
            console.log(err)
            result= false
        }

        return result;
    }

    async crearBorrador(argumentos) {

        //ARGUMENTOS:
        //0: HTML (String)
        //1: Asunto
        //2: Destinatario 
        //3: Adjunto {
        // filename: 'sample.jpg',
        // contentType: 'image/jpg',
        // data: '...base64 encoded data...'
        // }

        console.log("Ejecutando CREAR BORRADOR");
        console.log("Asunto: " + argumentos[1]);
        console.log("Destinatario: " + argumentos[2]);

        this.oAuth2Client = await this.authorize();
        if(!this.oAuth2Client){return false;}

        console.log("OAUTH2 CLIENT: ")
        console.log(this.oAuth2Client)

        const gmail = google.gmail({ version: 'v1', auth: this.oAuth2Client });

        //Construcción de mensaje:
        const message = createMimeMessage()

        message.setSender("me")
        message.setTo(argumentos[2])
        message.setSubject(argumentos[1])

        message.addMessage({
            contentType: 'text/html',
            data: argumentos[0] 
        })

        /*
        message.addAttachment({
            inline: true,
            filename: 'logo_sanfi_correo.png',
            contentType: 'image/png',
            data: '',
            headers: {'Content-ID': 'image00'}
        })
        */
        
        if(argumentos[3]!==undefined){
            for(var i=0; i<argumentos[3].length; i++){
                if(argumentos[3][i].qr){
                    //Creación de QR:
                    let QRbase64 = await new Promise((resolve, reject) => {
                    QRCode.toDataURL(argumentos[3][i].urlQR, function (err, code) {
                            if (err) {
                                reject(reject);
                                return;
                            }
                            resolve(code);
                        });
                    });

                    console.log("QR BASE64: ")
                    console.log(QRbase64)

                    //Eliminacion de cabecera: (data:image/png;base64,)
                    argumentos[3][i].data = QRbase64.slice(22);
                    message.addAttachment(argumentos[3][i])

                }else{
                    message.addAttachment(argumentos[3][i])
                }
            }
        }
    
        /*
        fs.writeFile(this.pathmessage, raw, err => {
          if (err) {
            console.error(err);
          } else {
            // file written successfully
          }
        });
        */

        var drafts = await gmail.users.drafts.create({
            userId: "me",
            requestBody: {
                message: {
                    raw: message.asEncoded()
                }
            }
        }).catch((v) => {
            console.log(v)
            console.log("ERROR MENSAJE")
            return false;
        })

        console.log("DRAFTS: ")
        console.log(drafts)
        var idDraft = drafts.data.message.threadId;
        return idDraft;
    }


} //Fin Procesos Google


module.exports = ProcesosGoogle;


