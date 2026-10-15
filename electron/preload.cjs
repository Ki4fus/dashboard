const {contextBridge, ipcRenderer} = require('electron');
contextBridge.exposeInMainWorld('dashboard', {
  restore: () => ipcRenderer.invoke('excel:restore'),
  open: () => ipcRenderer.invoke('excel:open'),
  refresh: () => ipcRenderer.invoke('excel:refresh')
});
