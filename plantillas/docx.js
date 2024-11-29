const path = require("path");
const fs = require("fs");
const readline = require('readline')
const {google} = require('googleapis');

//const {listCommands} = require('docx-templates');


	// If modifying these scopes, delete token.json.
	const SCOPES = ['https://www.googleapis.com/auth/gmail.readonly'];
	// The file token.json stores the user's access and refresh tokens, and is
	// created automatically when the authorization flow completes for the first
	// time.
	const TOKEN_PATH = 'token.json';
	var oAuth2Client;

class PlantillaDocx {
	
	constructor(pathToDbFolder, nombreProyecto, proyectoDB){

		this.pathToDbFolder = pathToDbFolder;
		this.nombreProyecto = nombreProyecto; 
		this.proyectoDB = proyectoDB;

		// Load client secrets from a local file.
		/*
		fs.readFile(path.join(__dirname,"../",'credentials.json'), (err, content) => {
		  if (err) return console.log('Error loading client secret file:', err);
		  // Authorize a client with credentials, then call the Gmail API.
		  this.authorize(JSON.parse(content), this.listLabels);
		});
		*/
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
		if (err) return getNewToken(oAuth2Client, callback);
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

	getNewToken(oAuth2Client, callback) {
	  const authUrl = oAuth2Client.generateAuthUrl({
		access_type: 'offline',
		scope: SCOPES,
		prompt: "consent"
	  });
		
	  console.log('Authorize this app by visiting this url:', authUrl);
	  const rl = readline.createInterface({
		input: process.stdin,
		output: process.stdout,
	  });
	  rl.question('Enter the code from that page here: ', (code) => {
		rl.close();
		oAuth2Client.getToken(code, (err, token) => {
		  if (err) return console.error('Error retrieving access token', err);
		  oAuth2Client.setCredentials(token);
		  // Store the token to disk for later program executions
		  fs.writeFile(TOKEN_PATH, JSON.stringify(token), (err) => {
			if (err) return console.error(err);
			console.log('Token stored to', TOKEN_PATH);
		  });
		  callback(oAuth2Client);
		});
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

	//**********************************************FIN GOOGLE GMAIL API*******************************************
 	
	async esperar(tiempo){
		return new Promise((resolve)=>{
			setTimeout(resolve, tiempo);
		});
	}

	async obtenerCorreos(argumentos){
		console.log("Ejecutando OBTENER CORREOS");
		console.log("Querry:" + argumentos[0]);
		const messages = await this.listarMensajes(this.oAuth2Client, 'label:inbox subject:reminder');
		console.log("Mensajes")
		console.log(messages)
		return true;
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
	

	async generarPlantillaDocx(argumentos){

			var pathPlantilla = path.normalize(argumentos[0])
			console.log("PATH PLANTILLA:" + argumentos[0])
			
			const template_buffer = fs.readFileSync(pathPlantilla);
        /*
			const commands = await listCommands(template_buffer, ['{', '}']);
			console.log("Parametros: ");
			console.log(commands)
			return commands
            */
        return false;
	}

} //Fin Procesos Google

module.exports = PlantillaDocx;


