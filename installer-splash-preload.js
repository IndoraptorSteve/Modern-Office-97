const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
  installerComplete: () => ipcRenderer.send('installer-complete')
});
