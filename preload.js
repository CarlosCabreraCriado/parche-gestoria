// All of the Node.js APIs are available in the preload process.
// It has the same sandbox as a Chrome extension.
window.addEventListener('DOMContentLoaded', () => {
  const replaceText = (selector, text) => {
    const element = document.getElementById(selector)
    if (element) element.innerText = text
  }

  for (const type of ['chrome', 'node', 'electron']) {
    replaceText(`${type}-version`, process.versions[type])
  }
})

const { contextBridge, ipcRenderer } = require('electron')

const $canalesOnIPC = ["onAbrirModo","onErrorInterno"];
const $canalesInvokeIPC = [
    "onGetArbolProyecto",
    "onImportarSpool",
    "onIncluirDirectorio",
    "onProcesarSpool",
    "onEjecutarProceso",
    "onEjecutarPlantilla",
    "onAnalizarSpool",
    "onGuardarDocumento",
    "onEliminarArchivo"
];

contextBridge.exposeInMainWorld('electronAPI', {

    //Gestion de Canales SendSync IPC (Angular --> Electron):
    getProyecto: () => ipcRenderer.sendSync('getProyecto'),
    getCorreo: () => ipcRenderer.sendSync('getCorreo'),

    guardarEnConfiguracion: (objetoConfiguracion) => ipcRenderer.sendSync('guardarEnConfiguracion', objetoConfiguracion),
    guardarEnProyecto: (objetoProyecto) => ipcRenderer.sendSync('guardarEnProyecto', objetoProyecto),
    setDocumento: (objetoDocumento) => ipcRenderer.sendSync('setDocumento', objetoDocumento),
    getDocumento: (objetoDocumento) => ipcRenderer.sendSync('getDocumento', objetoDocumento),
    getCampos: (objetoDocumento) => ipcRenderer.sendSync('getCampos', objetoDocumento),
    getListaObjetosEnColeccion: (path, nombreArchivo) => ipcRenderer.sendSync('getListaObjetosEnColeccion', path, nombreArchivo),
    getDocumentoPath: (path, nombreArchivo, filtro) => ipcRenderer.sendSync('getDocumentoPath', path, nombreArchivo, filtro),
    listarProyectos: () => ipcRenderer.sendSync('listarProyectos'),
    crearProyecto: (objetoNuevoProyecto) => ipcRenderer.sendSync('crearProyecto', objetoNuevoProyecto),
    eliminarProyecto: (nombreProyecto) => ipcRenderer.sendSync('eliminarProyecto', nombreProyecto),
    abrirProyecto: (nombreProyecto) => ipcRenderer.sendSync('abrirProyecto', nombreProyecto),
    cerrarProyecto: () => ipcRenderer.sendSync('cerrarProyecto'),
    abrirEditorPrograma: () => ipcRenderer.sendSync('abrirEditorPrograma'),
    abrirEditorDocumento: () => ipcRenderer.sendSync('abrirEditorDocumento'),
    

    //Gestion de Canales Invoke IPC (Angular --> Electron --> Angular):
    ejecutarProceso: (proceso, argumentos) => ipcRenderer.invoke('onEjecutarProceso', proceso, argumentos),

    guardarDocumento: (datos, tipo) => ipcRenderer.invoke('onGuardarDocumento', datos, tipo),
    getArbolProyecto: (nombreProyecto) => ipcRenderer.invoke('onGetArbolProyecto', nombreProyecto),
    incluirDirectorio: () => ipcRenderer.invoke('onIncluirDirectorio'),
    ejecutarPlantilla: (proceso, argumentos) => ipcRenderer.invoke('onEjecutarPlantilla', proceso, argumentos),
    eliminarDocumento: (path, nombreArchivo) => ipcRenderer.invoke('onEliminarDocumento', path, nombreArchivo),
    setCodigoGoogle: (codigo) => ipcRenderer.invoke('setCodigoGoogle', codigo),


   

    invoke: (channel, args) => {
            let validChannels = $canalesInvokeIPC;
            if (validChannels.includes(channel)) {
                return ipcRenderer.invoke(channel, args);
            }
        },


    //Gestion de Canales ON IPC (Electron --> Angular):
    on: (channel, listener) => {
            let validChannels = $canalesOnIPC;
            if (validChannels.includes(channel)) {
                // Deliberately strip event as it includes `sender`.
                ipcRenderer.on(channel, (event, ...args) => listener(...args));
            }
        }

})
