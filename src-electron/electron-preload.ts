import { contextBridge, IpcRendererEvent, ipcRenderer } from 'electron';
import { RegParseModel } from './modules/reg-parse-model';

type cb = (event: IpcRendererEvent, res:{state: string, progress: string}) => void

declare global {
  interface Window  {
    electronAPI: {
      openFile: () => Promise<string>,
      parseColumns: (filePath: string) => Promise<[]>,
      executeRegMatch: (config: RegParseModel) => void,
      onUpdateState: (callback: cb) => void
    }
  }
}

contextBridge.exposeInMainWorld('electronAPI',{
  openFile: () => ipcRenderer.invoke('dialog:openFile'),
  parseColumns: (filePath: string) => ipcRenderer.invoke('parseColumns', filePath),
  executeRegMatch: (config: RegParseModel) => ipcRenderer.send('executeRegMatch', config),
  onUpdateState: (callback: cb) => ipcRenderer.on('updateState', callback)
})


