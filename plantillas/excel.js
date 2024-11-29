
const path = require("path");
const fs = require("fs");
const readline = require('readline')
const {google} = require('googleapis');
const {Base64} = require('js-base64');
const electron = require("electron");
const ipc = require("electron").ipcMain;
const mainProcess = require("../main.js");

// If modifying these scopes, delete token.json.
const SCOPES = ['https://www.googleapis.com/auth/gmail.readonly'];
// The file token.json stores the user's access and refresh tokens, and is
// created automatically when the authorization flow completes for the first
// time.
const TOKEN_PATH = 'token.json';
var oAuth2Client;
var codigoGoogle;

class PlantillaExcel {
	
	constructor(pathToDbFolder, nombreProyecto, proyectoDB){

		this.pathToDbFolder = pathToDbFolder;
		this.nombreProyecto = nombreProyecto; 
		this.proyectoDB = proyectoDB;

	}

	//*********************************************** GOOGLE GMAIL API*********************************************

	/**
	 * Create an OAuth2 client with the given credentials, and then execute the
	 * given callback function.
	 * @param {Object} credentials The authorization client credentials.
	 * @param {function} callback The callback to call with the authorized client.
	 */

	authorize(credentials, callback) {

	   const {client_secret, client_id, redirect_uris} = credentials.installed;
	   this.oAuth2Client = new google.auth.OAuth2(
		  client_id, client_secret, redirect_uris[0]);

	  // Check if we have previously stored a token.
	  fs.readFile(TOKEN_PATH, (err, token) => {
		if (err) return this.getNewToken(this.oAuth2Client, callback);
		this.oAuth2Client.setCredentials(JSON.parse(token));
		callback(this.oAuth2Client);
	  });
	}

	/**
	 * Get and store new token after prompting for user authorization, and then
	 * execute the given callback with the authorized OAuth2 client.
	 * @param {google.auth.OAuth2} oAuth2Client The OAuth2 client to get token for.
	 * @param {getEventsCallback} callback The callback for the authorized client.
	 */

	setCodigoGoogle(codigo){
		console.log("SET COGIGO GOOGLE");
		this.codigoGoogle= codigo
		return;
	}

	getNewToken(oAuth2Client, callback) {
		console.log("Refrescando Token");
		const authUrl = this.oAuth2Client.generateAuthUrl({
			access_type: 'offline',
			scope: SCOPES,
			approval_prompt: 'force'
		});

			
		const rl = readline.createInterface({
			input: process.stdin,
			output: process.stdout
		});

		console.log("Probando autorización con código: "+this.codigoGoogle)

		oAuth2Client.getToken(this.codigoGoogle, (err, token) => {
		  if (err){
			console.log('Authorize this app by visiting this url:', authUrl);
			mainProcess.autentificarGoogle(authUrl).then((result) => {
				console.log(result)
			})
			return console.error('Error retrieving access token', err);
		  }else{
			  oAuth2Client.setCredentials(token);
			  // Store the token to disk for later program executions
			  fs.writeFile(TOKEN_PATH, JSON.stringify(token), (err) => {
				if (err) return console.error(err);
				console.log('Token stored to', TOKEN_PATH);
			  });
			  callback(oAuth2Client);
		  }
		});
	}

	/**
	 * Lists the labels in the user's account.
	 *
	 * @param {google.auth.OAuth2} auth An authorized OAuth2 client.
	 */

	listLabels(auth) {
	  const gmail = google.gmail({version: 'v1', auth});
	  gmail.users.labels.list({
		userId: 'me',
	  }, (err, res) => {
		if (err) return console.log('The API returned an error: ' + err);
		const labels = res.data.labels;
		if (labels.length) {
		  console.log('Labels:');
		  labels.forEach((label) => {
			console.log(`- ${label.name}`);
		  });
		} else {
		  console.log('No labels found.');
		}
	  });
	}

 	
	async esperar(tiempo){
		return new Promise((resolve)=>{
			setTimeout(resolve, tiempo);
		});
	}

	async getValidacionGoogle(){

		// Load client secrets from a local file.
		fs.readFile(path.join(__dirname,"../",'credentials.json'), (err, content) => {
		  if (err) return console.log('Error loading client secret file:', err);
		  // Authorize a client with credentials, then call the Gmail API.
		  this.authorize(JSON.parse(content), this.listLabels);
		});

	}

	async correoAInfolex(argumentos){
		var validacion = await this.getValidacionGoogle();
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
				});

				if(mensaje.data.payload.parts=== undefined){
					console.log("SIN PARTE:")
					if(mensaje.data.payload.mimeType=="text/html"){
						contenidoMensajes.push(Base64.decode(mensaje.data.payload.body.data.replace(/-/g, '+').replace(/_/g, '/')))
					}
					continue;
				}

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

	listarMensajes(auth, query) {
	  return new Promise((resolve, reject) => {
		const gmail = google.gmail({version: 'v1', auth});
		gmail.users.messages.list({
			userId: 'me',
			q: query,
		  },(err, res) => {

			if (err) {
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

	//**********************************************FIN GOOGLE GMAIL API*******************************************
	addExcel(argumentos){
		console.log("REALIZANDO ADD EXCEL: ");
		console.log("Argumentos");
		console.log(argumentos);
	}
} //Fin Procesos Excel

module.exports = PlantillaExcel;


