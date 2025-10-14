//DETERMINA SI ES DESARROLLO O PRODUCCION:
const DEBUG = process.env.NODE_ENV === "dev"; //Verifica si esta en producción
console.log(process.env.NODE_ENV);
console.log("DESARROLLO: " + DEBUG);

// Modules to control application life and create native browser window
const electron = require("electron");
const { dialog } = require("electron");
const path = require("path");
const url = require("url");
const fs = require("fs");
const ipc = require("electron").ipcMain;

var https = require("https");
const readline = require("readline");
const moment = require("moment");

const { autoUpdater } = require("electron-updater");
var _ = require("lodash");

if (process.env !== undefined) {
  console.warn("USANDO WEBPACK: ");
}

var procesosGenerales;
var procesosAsesoria;
var procesosFie;
var procesosPrueba;
var procesosDocumentos;
var procesosSpool;
var procesosRemedy;
var procesosDesarrollador;
var procesosGoogle;
var procesosDespacho;
var procesosKPIs;
var procesosImport;

var procesarPlantillaDocx;
var procesarPlantillaExcel;

var correo = [];

autoUpdater.autoDownload = false;
autoUpdater.autoInstallOnAppQuit = false;

//Inicialización del sistema de almacenamiento local:
var Datastore = require("nedb");

const pathToDbFolder = path.join(
  DEBUG ? __dirname : __dirname,
  DEBUG ? "" : "../",
  "db",
);

console.log("RUTA A BASE DE DATOS: ");
console.log(pathToDbFolder);

// Module to control application life.
const app = electron.app;

//app.disableHardwareAcceleration();
// Module to create native browser window.
const BrowserWindow = electron.BrowserWindow;

// Keep a global reference of the window object, if you don't, the window will
// be closed automatically when the JavaScript object is garbage collected.
let mainWindow;
let editorWindow;
let documentoWindow;
let autentificacion;

function createWindow() {
  // Create the browser window.
  if (DEBUG) {
    console.log("CARGANDOOOO");
    mainWindow = new BrowserWindow({
      width: 1600,
      height: 800,
      webPreferences: {
        enableRemoteModule: false,
        nodeIntegration: false,
        preload: path.join(__dirname, "preload.js"),
        contextIsolation: true,
        nativeWindowOpen: true,
        affinity: "main-window",
      },
      //webPreferences: {
      //preload: path.join(__dirname, 'preload.js')
      //}
    });
  } else {
    mainWindow = new BrowserWindow({
      width: 1000,
      height: 600,
      webPreferences: {
        enableRemoteModule: false,
        nodeIntegration: false,
        preload: path.join(__dirname, "preload.js"),
        contextIsolation: true,
        nativeWindowOpen: true,
        affinity: "main-window",
      },
      //webPreferences: {
      //preload: path.join(__dirname, 'preload.js')
      //}
    });
  }

  // and load the index.html of the app.
  if (DEBUG) {
    mainWindow.loadURL("http://localhost:4200/");
  } else {
    mainWindow.loadURL(
      url.format({
        pathname: path.join(__dirname, "compilado/browser/index.html"),
        protocol: "file:",
        slashes: true,
      }),
    );

    console.log("RUTA DE PAGINA: ");
    console.log(path.join(__dirname, "compilado/browser/index.html"));
    //console.log('file://${__dirname}/compilado/index.html');
    //mainWindow.loadFile(path.join(__dirname, 'dist/GestorSantanderFormacion/index.html'));

    //mainWindow.loadURL('file://${__dirname}/compilado/index.html');
    //mainWindow.webContents.openDevTools();
  }

  //mainWindow.removeMenu()
  // Open the DevTools.
  if (DEBUG) {
    mainWindow.webContents.openDevTools();
  }

  mainWindow.webContents.on(
    "new-window",
    (event, url, frameName, disposition, options) => {
      //event.preventDefault()
      options.webPreferences.affinity = "main-window";
      Object.assign(options, {});
      /*
          const win = new BrowserWindow({
            show: false,
            webPreferences: { 
                nodeIntegration: false,
                nativeWindowOpen: true,
                affinity: 'main-window'
            }
          })

          win.once('ready-to-show', () => win.show())
          win.loadURL(url)
          event.newGuest = win
          */
    },
  );
}

function createEditorPrograma() {
  // Create the browser window.
  if (DEBUG) {
    editorWindow = new BrowserWindow({
      width: 1600,
      height: 800,
      webPreferences: { enableRemoteModule: true, nodeIntegration: true },
      //webPreferences: {
      //preload: path.join(__dirname, 'preload.js')
      //}
    });
  } else {
    editorWindow = new BrowserWindow({
      width: 1000,
      height: 600,
      webPreferences: { enableRemoteModule: true, nodeIntegration: true },
      //webPreferences: {
      //preload: path.join(__dirname, 'preload.js')
      //}
    });
  }

  // and load the index.html of the app.
  if (DEBUG) {
    editorWindow.loadURL("http://localhost:4200/editor");
  } else {
    editorWindow.loadURL(
      url.format({
        pathname: path.join(__dirname, "compilado/browser/index.html"),
        protocol: "file:",
        slashes: true,
      }),
    );

    console.log("RUTA DE PAGINA: ");
    console.log(path.join(__dirname, "compilado/browser/index.html"));
    //console.log('file://${__dirname}/compilado/index.html');
    //mainWindow.loadFile(path.join(__dirname, 'dist/Santander-Fomacion/index.html'));

    //mainWindow.loadURL('file://${__dirname}/compilado/index.html');
    //mainWindow.webContents.openDevTools();
  }

  //mainWindow.removeMenu()
  // Open the DevTools.
  if (DEBUG) {
    editorWindow.webContents.openDevTools();
  }

  editorWindow.webContents.on("new-window", (event, url) => {
    event.preventDefault();
    const win = new BrowserWindow({
      show: false,
      webPreferences: { nodeIntegration: false },
    });
    win.once("ready-to-show", () => win.show());
    win.loadURL(url);
    event.newGuest = win;
  });
}

function createEditorDocumento() {
  // Create the browser window.
  if (DEBUG) {
    documentoWindow = new BrowserWindow({
      width: 1800,
      height: 1000,
      webPreferences: { enableRemoteModule: true, nodeIntegration: true },
      //webPreferences: {
      //preload: path.join(__dirname, 'preload.js')
      //}
    });
  } else {
    documentoWindow = new BrowserWindow({
      width: 1200,
      height: 800,
      webPreferences: { enableRemoteModule: true, nodeIntegration: true },
      //webPreferences: {
      //preload: path.join(__dirname, 'preload.js')
      //}
    });
  }

  // and load the index.html of the app.
  if (DEBUG) {
    documentoWindow.loadURL("http://localhost:4200/documento");
  } else {
    documentoWindow.loadURL(
      url.format({
        pathname: path.join(__dirname, "compilado/browser/index.html"),
        protocol: "file:",
        slashes: true,
      }),
    );

    console.log("RUTA DE PAGINA: ");
    console.log(path.join(__dirname, "compilado/browser/index.html"));
    //console.log('file://${__dirname}/compilado/index.html');
    //mainWindow.loadFile(path.join(__dirname, 'dist/Santander-Fomacion/index.html'));

    //mainWindow.loadURL('file://${__dirname}/compilado/index.html');
    //mainWindow.webContents.openDevTools();
  }

  //mainWindow.removeMenu()
  // Open the DevTools.

  if (DEBUG) {
    documentoWindow.webContents.openDevTools();
  }

  documentoWindow.webContents.on("new-window", (event, url) => {
    event.preventDefault();
    const win = new BrowserWindow({
      show: false,
      webPreferences: {
        nodeIntegration: false,
        preload: path.join(__dirname, "preload.js"),
        contextIsolation: true,
      },
    });
    win.once("ready-to-show", () => win.show());
    win.loadURL(url);
    event.newGuest = win;
  });
}

function crearAutentificacion(url) {
  // Create the browser window.
  if (DEBUG) {
    autentificacion = new BrowserWindow({
      width: 1600,
      height: 800,
      webPreferences: { userAgent: "Chrome" },
      //webPreferences: {
      //preload: path.join(__dirname, 'preload.js')
      //}
    });
  } else {
    autentificacion = new BrowserWindow({
      width: 1000,
      height: 600,
      webPreferences: { userAgent: "Chrome" },
      //webPreferences: {
      //preload: path.join(__dirname, 'preload.js')
      //}
    });
  }

  // and load the index.html of the app.
  if (DEBUG) {
    autentificacion.loadURL(url);
  } else {
    autentificacion.loadURL(url);
  }

  //mainWindow.removeMenu()
  // Open the DevTools.
  if (DEBUG) {
    autentificacion.webContents.openDevTools();
  }

  autentificacion.webContents.on("new-window", (event, url) => {
    event.preventDefault();
    const win = new BrowserWindow({
      show: false,
      webPreferences: { nodeIntegration: false },
    });
    win.once("ready-to-show", () => win.show());
    win.loadURL(url);
    event.newGuest = win;
  });
}
// This method will be called when Electron has finished
// initialization and is ready to create browser windows.
// Some APIs can only be used after this event occurs.

app.on("ready", createWindow);

// Quit when all windows are closed.
app.on("window-all-closed", function () {
  // On macOS it is common for applications and their menu bar
  // to stay active until the user quits explicitly with Cmd + Q
  if (process.platform !== "darwin") app.quit();
});

app.on("activate", function () {
  // On macOS it's common to re-create a window in the app when the
  // dock icon is clicked and there are no other windows open.
  if (BrowserWindow.getAllWindows().length === 0) createWindow();
});

//********************************
//  FUNCIONES AUTOUPDATE:
//********************************

ipc.on("app_version", (event) => {
  console.log("Version: " + app.getVersion());
  event.sender.send("app_version", { version: app.getVersion() });
});

ipc.on("buscarActualizacion", () => {
  console.log("Buscando actualización");
  //autoUpdater.checkForUpdatesAndNotify();
  autoUpdater
    .checkForUpdates()
    .then((result, err) => {
      if (err) {
        mainWindow.webContents.send("actualizacionNoEncontrada");
        console.log("ERROR: ");
        console.log(err);
        return;
      } else {
        console.log("RESULT: ");
        console.log(result);
      }
    })
    .catch((err) => {
      console.log("ERROR: ");
      console.error(err);
      mainWindow.webContents.send("actualizacionNoEncontrada");
    });
});

autoUpdater.on("update-available", () => {
  console.log("ACTUALIZACION DISPONIBLE");
  mainWindow.webContents.send("updateEncontrada");
});

autoUpdater.on("update-not-available", () => {
  console.log("La actualización no esta disponible:");
  mainWindow.webContents.send("updateActual");
});

autoUpdater.on("download-progress", (progress) => {
  //mainWindow.webContents.send('progresoDescarga',progress);
  console.log(progress);
});

autoUpdater.on("update-downloaded", () => {
  mainWindow.webContents.send("descargaCompletada");
});

ipc.on("descargarActualizacion", () => {
  autoUpdater.downloadUpdate().catch((err) => {
    console.error(err);
    mainWindow.webContents.send("errorInterno", err);
  });
});

ipc.on("instalarActualizacion", () => {
  try {
    autoUpdater.quitAndInstall();
  } catch (err) {
    console.error(err);
    mainWindow.webContents.send("errorInterno", err);
  }
});

//********************************
//  Declaracion de base de datos:
//********************************

var listaProyectos = [];
var nombreProyecto = "";

var masterDB = {};
var proyecto = null;
var proyectoConfig = null;

inicializarMasterDB();

function inicializarMasterDB() {
  console.log("Buscando Ruta:");
  console.log(path.join(pathToDbFolder, "masterDB.db"));

  masterDB = new Datastore(path.join(pathToDbFolder, "masterDB.db"));

  masterDB.loadDatabase(function (err) {
    if (err) {
      console.log(err);
      console.log("Se ha producido un error cargando MasterDB.");
    }
  });

  var configuracionMaestra = {
    nombreId: "configuracionMaestra",
  };

  //Verifica si existe:
  masterDB.find(
    { nombreId: configuracionMaestra.nombreId },
    function (err, docs) {
      if (err) {
        console.log("Se ha producido un error");
        console.log(err);
        return;
      }

      if (docs.length != 0) {
        console.log("Configuracion Maestra: ");
        console.log(docs);
        return true;
      } else {
        console.log("Creando configuracion maestra: ");
        masterDB.insert(configuracionMaestra, function (err, newDocs) {
          if (err) {
            console.log("Se ha producido un error");
            console.log(err);
            return false;
          } else {
            return true;
          }
        });
      }
    },
  );
} //Cierre de inicializar MasterDB

ipc.on("listaArchivos", function (event, rutaCarpeta) {
  var listaArchivos = [];

  console.log("Listando Archivos en ruta: " + rutaCarpeta);

  //Leer los nombres de los archivos en la base de datos:
  fs.readdir(rutaCarpeta, (err, files) => {
    if (err) {
      console.log(
        "Se ha producido un error leyendo los archivos del directorio: " +
          rutaCarpeta,
      );
      console.log(err);
      event.returnValue = false;
    } else {
      var totalArchivosCargar = 0;
      var cuentaArchivosCargados = 0;
      var archivoProyecto;

      //Excluir archivos:
      var auxFiles = files.slice();

      files.forEach((file, index, array) => {
        if (file.indexOf(".") != -1) {
          auxFiles.splice(auxFiles.indexOf(file), 1);
        }
      });

      //Si no hay archivos para cargar:
      console.log("Analizando AUX:");
      console.log(auxFiles);

      totalArchivosCargar = auxFiles.length;
      if (totalArchivosCargar == 0) {
        event.returnValue = [];
      } else {
        event.returnValue = auxFiles;
      }
    }
  });
});

ipc.on("listarProyectos", function (event) {
  listaProyectos = [];

  /*masterDB.find({nombreId: "Proyecto"}, function (err, docs) {

    if(err){
      console.log("Se ha producido un error");
      console.log(err);
      event.returnValue = null;
    }else{
      console.log("Documentos: ")
      console.log(docs)
      event.returnValue = docs;
    }
  });*/
  console.log("Listando Proyectos: ");
  //Leer los nombres de los archivos en la base de datos:
  fs.readdir(pathToDbFolder, (err, files) => {
    if (err) {
      console.log(
        "Se ha producido un error leyendo los archivos del directorio db",
      );
      console.log(err);
      event.returnValue = false;
    } else {
      var totalArchivosCargar = 0;
      var cuentaArchivosCargados = 0;
      var archivoProyecto;
      //Excluir archivos:
      var auxFiles = files.slice();

      files.forEach((file, index, array) => {
        if (file.indexOf(".") != -1) {
          auxFiles.splice(auxFiles.indexOf(file), 1);
        }
      });

      totalArchivosCargar = auxFiles.length;
      if (totalArchivosCargar == 0) {
        event.returnValue = [];
      }
      console.log("Analizando AUX:");
      console.log(auxFiles);

      auxFiles.forEach((file, index, array) => {
        //Ignorar si el archivo empieza por "."
        if (file.indexOf(".") == -1) {
          //Entra en cada directorio y extrae la informacion de proyecto:
          archivoProyecto = new Datastore(
            path.join(pathToDbFolder, file, "configuracion.db"),
          );

          archivoProyecto.loadDatabase(function (err) {
            if (err) {
              console.log(err);
              console.log(
                "Se ha producido un error cargando la base de datos.",
              );
              event.returnValue = false;
            }
          });

          //Leer Archivo config:
          archivoProyecto.find(
            { nombreId: "proyectoConfig" },
            function (err, docs) {
              if (err) {
                console.log("Se ha producido un error");
                console.log(err);
                event.returnValue = false;
              } else {
                cuentaArchivosCargados++;
                listaProyectos.push(docs[0]);
              }
              if (totalArchivosCargar == cuentaArchivosCargados) {
                console.log("Lista Proyectos: ");
                console.log(listaProyectos);
                event.returnValue = listaProyectos;
              }
            },
          );
        }
      });
    }
  });
});

ipc.on("crearProyecto", function (event, dataProyecto) {
  //Comprueba el nombre de proyecto:
  if (
    dataProyecto.nombre === undefined ||
    dataProyecto.nombre === null ||
    dataProyecto.nombre === ""
  ) {
    console.log("Error al crear proyecto: nombre invalido");
  }

  nombreProyecto = dataProyecto.nombre;

  //Creacion de directorio:
  if (!fs.existsSync(path.join(pathToDbFolder, nombreProyecto))) {
    fs.mkdirSync(path.join(pathToDbFolder, nombreProyecto));
    fs.mkdirSync(path.join(pathToDbFolder, nombreProyecto, "Imagenes"));
    fs.mkdirSync(path.join(pathToDbFolder, nombreProyecto, "Archivos"));
    fs.mkdirSync(path.join(pathToDbFolder, nombreProyecto, "Informes"));
    fs.mkdirSync(path.join(pathToDbFolder, nombreProyecto, "Plantillas"));
    fs.mkdirSync(path.join(pathToDbFolder, nombreProyecto, "Programas"));
  }

  //Crea el archivo de configuracion:
  proyecto = new Datastore(
    path.join(pathToDbFolder, nombreProyecto, "configuracion.db"),
  );

  proyecto.loadDatabase(function (err) {
    if (err) {
      nombreProyecto = "";
      console.log(err);
      console.log("Se ha producido un error cargando la base de datos.");
      event.returnValue = false;
    }
  });

  dataProyecto["nombreId"] = "proyectoConfig";
  //Guardar objeto de configuración:

  //Verifica si existe:
  proyecto.find({ nombreId: dataProyecto.nombreId }, function (err, docs) {
    if (err) {
      console.log("Se ha producido un error");
      console.log(err);
      event.returnValue = false;
    }

    //Si existe un documento hace Update:
    if (docs.length != 0) {
      console.log("Archivo encontrado");
      proyecto.update(
        { nombreId: dataProyecto.nombreId },
        dataProyecto,
        { upsert: true },
        function (err, numReplaced, upsert) {
          if (err) {
            console.log("Se ha producido remplazando el archivo");
            console.log(err);
            event.returnValue = false;
          } else {
            console.log("Documento Actualizado con exito");
            console.log("Remplazos: " + numReplaced);
            console.log("Upsert: " + upsert);

            //SE CREA LA ENTRADA EN CONFIGURACION MAESTRA:
            masterDB.update(
              {
                nombreId: "Proyecto",
                nombreProyecto: dataProyecto.nombreId,
              },
              {
                nombreId: "Proyecto",
                nombreProyecto: dataProyecto.nombre,
                descripcion: dataProyecto["descripcion"],
                autor: dataProyecto["autor"],
              },
              { upsert: true },
              function (err, numReplaced, upsert) {
                if (err) {
                  console.log(
                    "Se ha producido remplazando el archivo en CONFIGURACION MAESTRA",
                  );
                  console.log(err);
                  event.returnValue = false;
                } else {
                  console.log("Documento Actualizado con exito");
                  console.log("Remplazos: " + numReplaced);
                  console.log("Upsert: " + upsert);

                  event.returnValue = true;
                }
              },
            );

            var documentoElementos = {
              nombreId: "Elementos",
              elementos: [],
            };

            proyecto.insert(documentoElementos, function (err) {
              if (err) {
                console.log("Se ha producido un error");
                console.log(err);
                event.returnValue = false;
              } else {
                event.returnValue = true;
              }
            });
          }
        },
      );

      //Si no existe documento crea uno nuevo:
    } else {
      console.log("Archivo no encontrado");

      proyecto.insert(dataProyecto, function (err, newDocs) {
        if (err) {
          console.log("Se ha producido un error");
          console.log(err);
          event.returnValue = false;
        } else {
          //SE CREA LA ENTRADA EN CONFIGURACION MAESTRA:
          masterDB.update(
            {
              nombreId: "Proyecto",
              nombreProyecto: dataProyecto.nombreId,
            },
            {
              nombreId: "Proyecto",
              nombreProyecto: dataProyecto.nombre,
              descripcion: dataProyecto["descripcion"],
              autor: dataProyecto["autor"],
            },
            { upsert: true },
            function (err, numReplaced, upsert) {
              if (err) {
                console.log(
                  "Se ha producido remplazando el archivo en CONFIGURACION MAESTRA",
                );
                console.log(err);
                event.returnValue = false;
              } else {
                console.log("Documento Actualizado con exito");
                console.log("Remplazos: " + numReplaced);
                console.log("Upsert: " + upsert);

                var documentoElementos = {
                  nombreId: "Elementos",
                  elementos: [],
                };

                proyecto.insert(documentoElementos, function (err) {
                  if (err) {
                    console.log("Se ha producido un error");
                    console.log(err);
                    event.returnValue = false;
                  } else {
                    event.returnValue = true;
                  }
                });
              }
            },
          );

          event.returnValue = true;
        }
      });
    }
  });
});

ipc.on("eliminarProyecto", function (event, nombreProyecto) {
  try {
    //Eliminar directorio Proyecto:
    var deleteFolderRecursive = function (path) {
      if (fs.existsSync(path)) {
        fs.readdirSync(path).forEach(function (file) {
          var curPath = path + "/" + file;
          if (fs.lstatSync(curPath).isDirectory()) {
            // recurse
            deleteFolderRecursive(curPath);
          } else {
            // delete file
            fs.unlinkSync(curPath);
          }
        });
        fs.rmdirSync(path);
      }
    };

    deleteFolderRecursive(path.join(pathToDbFolder, nombreProyecto));

    //file removed
    console.log(
      "Proyecto eliminado: " + path.join(pathToDbFolder, nombreProyecto),
    );

    masterDB.remove(
      {
        nombreId: "Proyecto",
        nombreProyecto: nombreProyecto,
      },
      {},
      function (err, numRemoved) {
        if (err) {
          console.log(err);
          event.returnValue = false;
        }
        event.returnValue = true;
      },
    );
  } catch (err) {
    console.error(err);
    masterDB.remove(
      {
        nombreId: "Proyecto",
        nombreProyecto: nombreProyecto,
      },
      {},
      function (err) {
        if (err) {
          console.log(err);
          event.returnValue = false;
        }
        event.returnValue = true;
      },
    );
    event.returnValue = false;
  }
});

ipc.handle("onEliminarDocumento", async (event, ruta, nombre) => {
  try {
    var pathArchivoEliminar = path.join(ruta, nombre);

    //Eliminar el archivo;
    if (fs.existsSync(pathArchivoEliminar)) {
      fs.unlinkSync(pathArchivoEliminar);
    }

    // Archivo eliminado:
    console.log("Proyecto eliminado: " + pathArchivoEliminar);
    return true;
  } catch (err) {
    console.error(err);
    return false;
  }
});

ipc.handle("onIncluirDirectorio", async (event) => {
  async function abrirPanelDirectorio() {
    return dialog.showOpenDialog({ properties: ["openDirectory"] });
  }

  var result = await abrirPanelDirectorio();
  return result;
});

ipc.on("abrirProyecto", function (event, nombreProy) {
  console.log("Abriendo proyecto: " + nombreProy);
  nombreProyecto = nombreProy;

  proyecto = new Datastore(
    path.join(pathToDbFolder, nombreProyecto, "configuracion.db"),
  );

  proyecto.loadDatabase(function (err) {
    if (err) {
      nombreProyecto = "";
      console.log(err);
      console.log("Se ha producido un error cargando la base de datos.");
      event.returnValue = false;
    }
  });

  proyecto.find({ nombreId: "proyectoConfig" }, function (err, docs) {
    if (err) {
      console.log(
        "Se ha producido un error buscando los documentos en la base de datos.",
      );
      console.log(err);
      event.returnValue = false;
    } else {
      console.log("Enviando documentos de base de datos: " + nombreProy);
      console.log(docs);
      proyectoConfig = docs[0];
      nombreProyecto = nombreProy;
      event.returnValue = docs;
    }
  });

  //Creacion de instancia de ejecucion de procesos:
  procesosGenerales = new ProcesosGenerales(
    pathToDbFolder,
    nombreProyecto,
    proyecto,
  );
  procesosAsesoria = new ProcesosAsesoria(
    pathToDbFolder,
    nombreProyecto,
    proyecto,
  );
  procesosFie = new ProcesosFie(pathToDbFolder, nombreProyecto, proyecto);
  procesosDocumentos = new ProcesosDocumentos(
    pathToDbFolder,
    nombreProyecto,
    proyecto,
  );
  procesosSpool = new ProcesosSpool(pathToDbFolder, nombreProyecto, proyecto);
  procesosRemedy = new ProcesosRemedy(pathToDbFolder, nombreProyecto, proyecto);
  procesosDesarrollador = new ProcesosDesarrollador(
    pathToDbFolder,
    nombreProyecto,
    proyecto,
  );
  procesosGoogle = new ProcesosGoogle(pathToDbFolder, nombreProyecto, proyecto);
  procesosDespacho = new ProcesosDespacho(
    pathToDbFolder,
    nombreProyecto,
    proyecto,
  );
  procesosImport = new ProcesosImport(pathToDbFolder, nombreProyecto, proyecto);
  procesosKPIs = new ProcesosKPIs(pathToDbFolder, nombreProyecto, proyecto);

  //Creacion de instancia de ejecucion de procesos:
  procesarPlantillaExcel = new PlantillaExcel(
    pathToDbFolder,
    nombreProyecto,
    proyecto,
  );
  procesarPlantillaDocx = new PlantillaDocx(
    pathToDbFolder,
    nombreProyecto,
    proyecto,
  );
});

ipc.on("onCerrarProyecto", function (event) {
  if (proyectoConfig == null || proyectoConfig == undefined) {
    proyecto = null;
    proyectoConfig = null;
    event.returnValue = true;
    return;
  }

  if (proyectoConfig.nombre) {
    console.log("Cerrando proyecto: " + proyectoConfig.nombre);
  }

  proyecto = null;
  proyectoConfig = null;

  event.returnValue = true;
});

ipc.on("getDocumentoPath", function (event, pathDato, nombreArchivo, filtro) {
  console.log("Obteniendo documento en base de datos " + nombreProyecto + ":");

  //Formateo del Path y nombre de archivo:
  nombreArchivo = nombreArchivo.replace(/\.[^/.]+$/, "");
  pathDato = pathDato.replace(path.join(pathToDbFolder, nombreProyecto), "");

  console.log("Path:");
  console.log(
    path.join(pathToDbFolder, nombreProyecto, pathDato, nombreArchivo + ".db"),
  );
  console.log("Filtro:");
  console.log(filtro);

  var archivoCargado = new Datastore(
    path.join(pathToDbFolder, nombreProyecto, pathDato, nombreArchivo + ".db"),
  );

  archivoCargado.loadDatabase(function (err) {
    if (err) {
      console.log(err);
      console.log("Se ha producido un error cargando la base de datos.");
      event.returnValue = false;
    }
  });

  archivoCargado.find(filtro, function (err, docs) {
    if (err) {
      console.log("Se ha producido un error");
      console.log(err);
      event.returnValue = null;
    } else {
      console.log("Documentos: ");
      console.log(docs);
      event.returnValue = docs;
    }
  });
});

ipc.on("getListaObjetosEnColeccion", function (event, pathDato, nombreArchivo) {
  console.log(
    "Obteniendo objetos en coleccion: " + nombreProyecto + "path: " + path,
  );

  var archivoCargado = new Datastore(path.join(pathDato, nombreArchivo));

  archivoCargado.loadDatabase(function (err) {
    if (err) {
      console.log(err);
      console.log("Se ha producido un error cargando la base de datos.");
      event.returnValue = false;
    }
  });

  archivoCargado.find(
    {},
    { nombreId: 1, objetoId: 1, _id: 0 },
    function (err, docs) {
      if (err) {
        console.log("Se ha producido un error");
        console.log(err);
        event.returnValue = null;
      } else {
        console.log("Documentos: ");
        event.returnValue = docs;
      }
    },
  );
});

ipc.on(
  "getObjetoEnColeccion",
  function (event, pathDato, nombreArchivo, objetoId) {
    console.log(
      "Obteniendo objetos en coleccion: " + nombreProyecto + "path: " + path,
    );

    var archivoCargado = new Datastore(path.join(pathDato, nombreArchivo));

    archivoCargado.loadDatabase(function (err) {
      if (err) {
        console.log(err);
        console.log("Se ha producido un error cargando la base de datos.");
        event.returnValue = false;
      }
    });

    archivoCargado.find(
      { objetoId: objetoId },
      { data: 1, _id: 0 },
      function (err, docs) {
        if (err) {
          console.log("Se ha producido un error");
          console.log(err);
          event.returnValue = null;
        } else {
          console.log("Documentos: ");
          event.returnValue = docs;
        }
      },
    );
  },
);

ipc.on("getDocumento", function (event, filtro) {
  console.log(
    "Obteniendo documento en base de datos " +
      nombreProyecto +
      ". NombreId: " +
      filtro["nombreId"],
  );

  proyecto.findOne(filtro, function (err, docs) {
    if (err) {
      console.log("Se ha producido un error");
      console.log(err);
      event.returnValue = null;
    } else {
      console.log("Documentos: ");
      console.log(docs);
      event.returnValue = docs;
    }
  });
});

ipc.on("getProyecto", function (event) {
  console.log(proyectoConfig);
  event.returnValue = proyectoConfig;
});

ipc.handle("onImportarSpool", async (event, rutaInput, nombreOutput) => {
  console.log("Importando SPOOL...");

  const pathSpoolInput = path.join(rutaInput);
  const pathSpoolOutput = path.join(
    pathToDbFolder,
    nombreProyecto,
    "/archivos",
    nombreOutput + ".txt",
  );

  console.log("Ruta de salida:");
  console.log(pathSpoolOutput);
  console.log("Ruta de entrada:");
  console.log(pathSpoolInput);

  const readline = require("readline");
  const outputFile = fs.createWriteStream(pathSpoolOutput);

  const rl = readline.createInterface({
    input: fs.createReadStream(pathSpoolInput),
  });

  // Handle any error that occurs on the write stream
  outputFile.on("err", (err) => {
    // handle error
    console.log(err);
  });

  // Once done writing, rename the output to be the input file name
  outputFile.on("close", () => {
    console.log("done writing");

    fs.rename(pathSpoolOutput, pathSpoolInput, (err) => {
      if (err) {
        // handle error
        console.log(err);
        return false;
      } else {
        console.log("renamed file");
        return true;
      }
    });
  });

  // Read the file and replace any text that matches
  rl.on("line", (line) => {
    let text = line;

    // Elimina las lineas que no comienzan por tabulador:
    if (!text.startsWith("\t")) {
      return;
    }

    // Elimina las lineas que comienzan por "Md.":
    if (text.startsWith("\tMd.\t")) {
      return;
    }

    // write text to the output file stream with new line character
    outputFile.write(`${text}\n`);
  });

  // Done reading the input, call end() on the write stream
  rl.on("close", () => {
    console.log("FIN DE IMPORTACION SPOOL");
    outputFile.end();
    return true;
  });
});

ipc.handle("onEjecutarProceso", async (event, proceso, argumentos) => {
  function camelize(str) {
    return str
      .replace(/(?:^\w|[A-Z]|\b\w)/g, function (word, index) {
        return index === 0 ? word.toLowerCase() : word.toUpperCase();
      })
      .replace(/\s+/g, "");
  }

  proceso.nombre = proceso.nombre
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");
  var identificador = camelize(proceso.nombre);
  var categoria = camelize(proceso.categoria);

  console.log("Ejecutando proceso: " + identificador);
  console.log("Categoria : " + categoria);
  console.log("Argumentos");
  console.log(argumentos);

  //Selección de archivo de Ejecucion:
  var result;
  switch (categoria) {
    case "general":
      //result = await procesosGenerales[identificador](argumentos.formularioControl);
      result = await procesosGenerales[identificador](argumentos);
      break;

    case "asesoria":
      //result = await procesosGenerales[identificador](argumentos.formularioControl);
      result = await procesosAsesoria[identificador](argumentos);
      break;
    case "fie":
      //result = await procesosGenerales[identificador](argumentos.formularioControl);
      result = await procesosFie[identificador](argumentos);
      break;
    case "documentos":
      //result = await procesosDocumentos[identificador](argumentos.formularioControl);
      result = await procesosDocumentos[identificador](argumentos);
      break;
    case "spool":
      result = await procesosSpool[identificador](argumentos.formularioControl);
      break;
    case "remedy":
      result = await procesosRemedy[identificador](
        argumentos.formularioControl,
      );
      break;
    case "desarrollador":
      result = await procesosDesarrollador[identificador](
        argumentos.formularioControl,
      );
      break;
    case "google":
      result = await procesosGoogle[identificador](
        argumentos.formularioControl,
      );
      break;
    case "import":
      result = await procesosImport[identificador](argumentos);
      //result = await procesosImport[identificador](argumentos.formularioControl);
      break;
    case "kPIs":
      result = await procesosKPIs[identificador](argumentos.formularioControl);
      break;
    case "despacho":
      result = await procesosDespacho[identificador](
        argumentos.formularioControl,
      );
      break;
  }

  if (result == undefined) {
    console.log("WARNING: El proceso ha devuelto un resultado 'undefined'");
    return false;
  }

  if (result["nombreId"]) {
    console.log("Guardando archivo: " + result.nombreId);
    result = await guardarDocumento(result);
  }

  if (result.tipo == "correo") {
    correo = result.data;
  }

  console.log("FIN PROCESO - DEVOLVIENDO A APPSERVICE");

  return result;
});

ipc.handle("onEjecutarPlantilla", async (event, proceso, argumentos) => {
  function camelize(str) {
    return str
      .replace(/(?:^\w|[A-Z]|\b\w)/g, function (word, index) {
        return index === 0 ? word.toLowerCase() : word.toUpperCase();
      })
      .replace(/\s+/g, "");
  }

  var identificador = proceso.nombre;
  var categoria = proceso.categoria;

  console.log("Ejecutando plantilla: " + identificador);
  console.log("Categoria : " + categoria);
  console.log("Argumentos");
  console.log(argumentos);
  console.log("proceso");
  console.log(proceso);

  //Selección de archivo de Ejecucion:
  var result;
  switch (categoria) {
    case "excel":
      //result = await procesarPlantillaExcel.addExcel(argumentos.formularioControl);
      break;
    case "docx":
      //result = await procesarPlantillaDocx.generarPlantillaDocx(argumentos.formularioControl);
      break;
    case "personalizado":
      break;
  }

  var datosPlantilla = {
    data: {},
    cmdDelimiter: ["{", "}"],
  };
  //Insertar Campos:

  for (var i = 0; i < proceso.argumentos.length; i++) {
    datosPlantilla.data[proceso.argumentos[i].identificador] =
      proceso.argumentos[i].formulario;
    console.log(proceso.argumentos[i].formulario);
  }

  const template = fs.readFileSync(
    "/Users/carloscabreracriado/Desktop/simple.docx",
  );

  const buffer = await createReport({
    template,
    data: {
      first_name: "Carlos",
      last_name: "Cabrera",
      phone: "690168013",
      description: "Esto es una Prueba",
    },
    cmdDelimiter: ["{", "}"],
  });

  fs.writeFileSync("/Users/carloscabreracriado/Desktop/salida.docx", buffer);

  result = true;
  if (result == undefined) {
    console.log("WARNING: La plantilla ha devuelto un resultado 'undefined'");
    return false;
  }

  if (result.nombreId) {
    console.log("Guardando archivo: " + result.nombreId);
    result = await guardarDocumento(result);
  }

  console.log("FIN PLANTILLA - DEVOLVIENDO A APPSERVICE");
  if (result.tipo == "correo") {
    correo = result.data;
  }

  return result;
});

ipc.handle("onProcesarSpool", async (event, datos) => {
  console.log("Procesando SPOOL");

  const pathSpoolInput = path.join(
    pathToDbFolder,
    nombreProyecto,
    "spoolInput.txt",
  );
  const pathSpoolOutput = path.join(
    pathToDbFolder,
    nombreProyecto,
    "spoolOutput.txt",
  );

  const readline = require("readline");
  const outputFile = fs.createWriteStream(pathSpoolOutput);

  const rl = readline.createInterface({
    input: fs.createReadStream(pathSpoolInput),
  });

  // Handle any error that occurs on the write stream
  outputFile.on("err", (err) => {
    // handle error
    console.log(err);
  });

  // Once done writing, rename the output to be the input file name
  outputFile.on("close", () => {
    console.log("done writing");

    /*fs.rename(pathSpoolOutput, pathSpoolInput, err => {
            if (err) {
              // handle error
              console.log(err)
            } else {
              console.log('renamed file')
            }
        })*/
    return false;
  });

  // Read the file and replace any text that matches
  rl.on("line", (line) => {
    let text = line;

    // Elimina las lineas que no comienzan por tabulador:
    if (!text.startsWith("\t")) {
      return;
    }

    // Elimina las lineas que comienzan por "Md.":
    if (text.startsWith("\tMd.\t")) {
      return;
    }

    // write text to the output file stream with new line character
    outputFile.write(`${text}\n`);
  });

  // Done reading the input, call end() on the write stream
  rl.on("close", () => {
    console.log("FIN DEL PROCESAMIENTO");
    outputFile.end();
  });

  return true;
});

ipc.handle("onAnalizarSpool", async (event, datos) => {
  console.log("Analizando SPOOL");

  const pathSpoolInput = path.join(
    pathToDbFolder,
    nombreProyecto,
    "spoolInput.txt",
  );
  const pathSpoolOutput = path.join(
    pathToDbFolder,
    nombreProyecto,
    "spoolOutput.txt",
  );

  const outputFile = fs.createWriteStream(pathSpoolOutput);
  const readline = require("readline");

  const rl = readline.createInterface({
    input: fs.createReadStream(pathSpoolInput),
  });

  var cuentaWeek = [];
  for (var i = 0; i < 53; i++) {
    cuentaWeek.push(0);
  }

  // Handle any error that occurs on the write stream
  outputFile.on("err", (err) => {
    // handle error
    console.log(err);
  });

  // Once done writing, rename the output to be the input file name
  outputFile.on("close", () => {
    console.log("done writing");

    /*fs.rename(pathSpoolOutput, pathSpoolInput, err => {
            if (err) {
              // handle error
              console.log(err)
            } else {
              console.log('renamed file')
            }
        })*/
  });

  // Read the file and replace any text that matches
  rl.on("line", (line) => {
    let text = line;

    // Elimina las lineas que no comienzan por tabulador:
    if (!text.startsWith("\t")) {
      return;
    }

    // Elimina las lineas que comienzan por "Md.":
    if (text.startsWith("\tMd.\t")) {
      return;
    }

    let parseado = text.split("\t");
    /*
        let semanaInicioBloqueo = moment(parseado[7], "DD.MM.YYYY");
        let semanaFinBloqueo = moment(parseado[8], "DD.MM.YYYY");

            for(let i = 0; i< 53; i++){
                if( ((semanaInicioBloqueo.year()<2020) || ((semanaInicioBloqueo.isoWeek()<=(i+1)) && (semanaInicioBloqueo.year()==2020))) && ((semanaFinBloqueo.year()>2020) || ((semanaFinBloqueo.isoWeek() >= (i+1))&& (semanaFinBloqueo.year()==2020)))){
                    cuentaWeek[i]++;
                }
            }
            */
    outputFile.write(`${parseado[2].substr(0, 12)}\n`);
  });

  outputFile.on("close", () => {
    console.log("done writing");

    /*fs.rename(pathSpoolOutput, pathSpoolInput, err => {
            if (err) {
              // handle error
              console.log(err)
            } else {
              console.log('renamed file')
            }
        })*/
  });

  // Done reading the input, call end() on the write stream
  rl.on("close", () => {
    console.log("FIN DEL PROCESAMIENTO");
    console.log("Cuenta: ");
    //Console.log(cuentaWeek);
  });

  return true;
});

ipc.handle("onGetArbolProyecto", async (event, nombreProyectoArg) => {
  console.log("Obteniendo arbol de proyecto");

  var arbolProyecto = {};

  async function obtenerNivelDirectorio(rutaDirectorio) {
    return new Promise((resolve) => {
      var nivelArbol = [];

      //Lee el directorio raiz:
      fs.readdir(rutaDirectorio, async (err, files) => {
        if (err) {
          console.log(
            "Se ha producido un error en la generacion del arbol de archivos",
          );
          console.log(err);
          event.returnValue = false;
        } else {
          //Excluir archivos que comiencen por '.':
          var auxFiles = files.slice();
          files.forEach((file, index, array) => {
            if (file.indexOf(".") == 0) {
              auxFiles.splice(auxFiles.indexOf(file), 1);
            }
          });
          files = auxFiles.slice();

          //Salir si no hay archivos:
          if (files.length == 0) {
            resolve(nivelArbol);
          }

          //Obtiene los Archivos:
          files.forEach((file) => {
            if (file.indexOf(".") != -1) {
              nivelArbol.push({
                nombre: file,
                tipo: file.substr(file.lastIndexOf(".") + 1),
                direccion: rutaDirectorio,
              });
            }
          });

          //Obtiene los directorios:
          files.forEach((file) => {
            if (file.indexOf(".") == -1) {
              nivelArbol.push({
                nombre: file,
                tipo: "dir",
                direccion: rutaDirectorio,
                subDirectorio: null,
              });
            }
          });
        }
        resolve(nivelArbol);
      });
    }); //Fin de la promesa
  }

  //Lee los archivos en directorio Proyecto:
  async function explorarDirectorio(ruta, nivelArbol = null) {
    console.log("Explorando: " + ruta);

    if (nivelArbol == null) {
      nivelArbol = await obtenerNivelDirectorio(ruta);
      console.log("Nivel completo:");
      console.log(nivelArbol);
    }

    var indicesDirectorioSinExplorar = [];

    nivelArbol.forEach((nivel, index) => {
      if (nivel.tipo == "dir" && nivel.subDirectorio == null) {
        indicesDirectorioSinExplorar.push(index);
      }
    });

    if (indicesDirectorioSinExplorar.length == 0) {
      return nivelArbol;
    }

    for (var i = 0; i < indicesDirectorioSinExplorar.length; i++) {
      nivelArbol[indicesDirectorioSinExplorar[i]].subDirectorio =
        await explorarDirectorio(
          path.join(ruta, nivelArbol[indicesDirectorioSinExplorar[i]].nombre),
        );
    }

    console.log("Fin nivel: " + ruta);
    return nivelArbol;
  } //Fin de explorar directorio

  arbolProyecto = await explorarDirectorio(
    path.join(pathToDbFolder, nombreProyectoArg),
  );
  console.log("ARBOL DE PROYECTO");
  console.log(arbolProyecto);
  return arbolProyecto;
});

ipc.handle("onGuardarDocumento", async (event, datos, tipo) => {
  if (tipo != "archivos" && tipo != "plantillas") {
    console.log(
      "Se ha producido un error en el guardado de documento: Tipo incorrecto.",
    );
    return false;
  }

  console.log("Guardando Documento: " + datos.nombreId);
  //Verifica que existe un proyecto abierto:
  if (
    nombreProyecto == "" ||
    nombreProyecto == null ||
    nombreProyecto == undefined
  ) {
    console.log("Nombre de proyecto no encontrado");
    return false;
  }

  //Verifica si el nombreId del documento es correcto:
  if (datos.nombreId === undefined || datos.nombreId === null) {
    console.log("No se puede insertar un documento sin nombreId");
    return false;
  }

  //Abre el nuevo documento:

  var documento = new Datastore(
    path.join(pathToDbFolder, nombreProyecto, tipo, datos.nombreId + ".db"),
  );

  documento.loadDatabase(function (err) {
    if (err) {
      console.log(err);
      console.log("Se ha producido un error cargando la base de datos.");
      return false;
    }
  });
  const promesa = new Promise((resolve) => {
    documento.find({ nombreId: datos.nombreId }, function (err, docs) {
      if (err) {
        console.log("Se ha producido un error");
        console.log(err);
        resolve(false);
      }

      if (docs.length != 0) {
        console.log("Archivo encontrado");
        documento.update(
          { nombreId: datos.nombreId },
          datos,
          { upsert: true },
          function (err, numReplaced, upsert) {
            if (err) {
              console.log("Se ha producido un error remplazando el archivo");
              console.log(err);
              resolve(false);
              return false;
            } else {
              console.log("Documento Actualizado con exito");
              console.log("Remplazos: " + numReplaced);
              resolve(true);
            }
          },
        );
      } else {
        console.log("Archivo no encontrado");

        documento.insert(datos, function (err, newDocs) {
          if (err) {
            console.log("Se ha producido un error");
            console.log(err);
            resolve(false);
          } else {
            console.log("Archivo guardado con exito");
            resolve(true);
          }
        });
      }
    });
  }); // Fin Promesa

  var respuesta = await promesa;

  return respuesta;
});

ipc.on("guardarEnConfiguracion", function (event, datos) {
  if (datos.nombreId === undefined || datos.nombreId === null) {
    console.log("No se puede insertar un documento sin nombreId");
    event.returnValue = false;
  }

  //Verifica si existe:
  proyecto.find({ nombreId: datos.nombreId }, function (err, docs) {
    if (err) {
      console.log("Se ha producido un error");
      console.log(err);
      event.returnValue = false;
    }

    if (docs.length != 0) {
      console.log("Archivo encontrado");
      proyecto.update(
        { nombreId: datos.nombreId },
        datos,
        {},
        function (err, numReplaced, upsert) {
          if (err) {
            console.log("Se ha producido un error remplazando el archivo");
            console.log(err);
            event.returnValue = false;
          } else {
            console.log("Documento Actualizado con exito");
            console.log("Remplazos: " + numReplaced);
            console.log("Upsert: " + upsert);

            event.returnValue = true;
          }
        },
      );
    } else {
      console.log("Archivo no encontrado");

      proyecto.insert(datos, function (err, newDocs) {
        if (err) {
          console.log("Se ha producido un error");
          console.log(err);
          event.returnValue = false;
        } else {
          event.returnValue = true;
        }
      });
    }
  });
});

//Se utiliza para guardar documentos dentro del scope de un proceso:
async function guardarDocumento(datos) {
  console.log("Guardando Documento: " + datos.nombreId);
  //console.log(datos)
  //Verifica que existe un proyecto abierto:
  if (
    nombreProyecto == "" ||
    nombreProyecto == null ||
    nombreProyecto == undefined
  ) {
    console.log("Nombre de proyecto no encontrado");
    return false;
  }

  //Verifica si el nombreId del documento es correcto:
  if (datos.nombreId === undefined || datos.nombreId === null) {
    console.log("No se puede insertar un documento sin nombreId");
    return false;
  }

  //Abre el nuevo documento:

  var documento = new Datastore(
    path.join(
      pathToDbFolder,
      nombreProyecto,
      "archivos",
      datos.nombreId + ".db",
    ),
  );

  documento.loadDatabase(function (err) {
    if (err) {
      console.log(err);
      console.log("Se ha producido un error cargando la base de datos.");
      return false;
    }
  });

  const promesa = new Promise((resolve) => {
    documento.find({ nombreId: datos.nombreId }, function (err, docs) {
      if (err) {
        console.log("Se ha producido un error");
        console.log(err);
        resolve(false);
      }

      if (docs.length != 0) {
        console.log("Archivo encontrado");
        documento.update(
          { nombreId: datos.nombreId },
          datos,
          { upsert: true },
          function (err, numReplaced, affectedDocuments, upsert) {
            if (err) {
              console.log("Se ha producido remplazando el archivo");
              console.log(err);
              resolve(false);
              return false;
            } else {
              console.log("Documento Actualizado con exito");
              console.log("Num Remplazos: " + numReplaced);
              resolve(true);
            }
          },
        );
      } else {
        console.log("Archivo no encontrado");
        documento.insert(datos, function (err, newDocs) {
          if (err) {
            console.log("Se ha producido un error");
            console.log(err);
            resolve(false);
          } else {
            console.log("Archivo guardado con exito");
            resolve(true);
          }
        });
      }
    });
  }); // Fin Promesa

  var respuesta = await promesa;
  return respuesta;
}

ipc.on("abrirEditorPrograma", function (event) {
  createEditorPrograma();
  event.returnValue = true;
});

ipc.on("abrirEditorDocumento", function (event) {
  createEditorDocumento();
  event.returnValue = true;
});

ipc.on("getCamposDocx", function (event, rutaArchivo) {
  /* Leer Plantilla */
  /*
    console.log(rutaArchivo)

    const template_buffer = fs.readFileSync(path.join(rutaArchivo));
    fs.readFile(path.join(rutaArchivo),(err,data)=>{
        
        listCommands(data, ['{', '}']).then((result,err)=>{
            if(err){
                console.log(err);
                event.returnvalue = false;
            }
            event.returnValue= result;
        });     

    })
    */
  event.returnvalue = false;
});

ipc.on("getCorreo", function (event) {
  event.returnValue = correo;
});

ipc.on("setCodigoGoogle", function (event, codigo) {
  console.log("CODIGO: " + codigo);
  procesosDespacho.setCodigoGoogle(codigo);
  event.returnValue = true;
});

function mostrarError(error, mensajeError) {
  return new Promise((resolve) => {
    mainWindow.webContents.send("mostrarError", error, mensajeError);
    resolve(true);
  });
}

function mostrarWarning(warning, mensajeWarning) {
  return new Promise((resolve) => {
    mainWindow.webContents.send("mostrarWarning", warning, mensajeWarning);
    resolve(true);
  });
}

function autentificarGoogle(url) {
  return new Promise((resolve) => {
    mainWindow.webContents.send("dialogoAutentificarGoogle", url);
    resolve(true);
  });
}

module.exports = {
  autentificarGoogle,
  guardarDocumento,
  mostrarError,
  mostrarWarning,
};

//Importar Procesos:
const ProcesosGenerales = require("./procesos/general.js");
const ProcesosAsesoria = require("./procesos/asesoria.js");
const ProcesosFie = require("./procesos/fie.js");
const ProcesosDocumentos = require("./procesos/documentos.js");
const ProcesosSpool = require("./procesos/spool.js");
const ProcesosRemedy = require("./procesos/remedy.js");
const ProcesosDesarrollador = require("./procesos/desarrollador.js");
const ProcesosGoogle = require("./procesos/google.js");
const ProcesosDespacho = require("./procesos/despacho.js");
const ProcesosImport = require("./procesos/import.js");
const ProcesosKPIs = require("./procesos/kpis.js");

const PlantillaExcel = require("./plantillas/excel.js");
const PlantillaDocx = require("./plantillas/docx.js");
