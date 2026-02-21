'use strict';
const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('launcherAPI', {
  launchApp: (name) => ipcRenderer.send('launch-app', name),
});
