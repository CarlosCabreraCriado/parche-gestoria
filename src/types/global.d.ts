export interface IElectronAPI {

    //Gestion de Peticiones Sinc::
    getProyecto: () => Promise<string>,
    getCorreo: () => Promise<string>,
    guardarEnConfiguracion: (objetoConfiguracion: any) => boolean,
    setDocumento: (objetoDocumento: any) => boolean,
    getDocumento: (objetoDocumento: any) => Promise<any>,
    getCamposDocx: (pathPlantilla: string) => Promise<any>,
    getListaObjetosEnColeccion: (path: string, nombreArchivo: string) => any[],
    getObjetoEnColeccion: (path: string, nombreArchivo: string, objetoId: string) => any,
    getDocumentoPath: (path: string, nombreArchivo: string, filtro?:any) => Promise<any>,
    listarProyectos: () => Promise<any[]>,
    crearProyecto: (objetoNuevoProyecto: any) => Promise<boolean>,
    eliminarProyecto: (nombreProyecto: string) => Promise<boolean>,
    abrirProyecto: (nombreProyecto: string) => Promise<any>,
    cerrarProyecto: () => Promise<boolean>,
    abrirEditorPrograma: () => Promise<boolean>,
    abrirEditorDocumento: () => Promise<boolean>,
    

    //Gestion de Peticiones On:
    on: (channel: string, listener: (event: any, ...args: any[]) => void) => void,

    //Gestion de Peticiones Invoke::
    invoke: (channel: string, ...args: any[]) => Promise<any>,
    ejecutarProceso: (proceso: any, argumentos: any[]) => Promise<any>,
    guardarDocumento: (datos: any, tipo: string) => Promise<any>,
    getArbolProyecto: (nombreProyecto: string) => Promise<any>,
    incluirDirectorio: () => Promise<any>,
    ejecutarPlantilla: (proceso: any, argumentos: any[]) => Promise<any>,
    eliminarDocumento: (path: string, nombre: string) => Promise<boolean>,
    setCodigoGoogle: (codigo: string) => Promise<boolean>

}

declare global {
  interface Window {
    electronAPI: IElectronAPI
  }
}
