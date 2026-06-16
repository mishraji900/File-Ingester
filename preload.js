const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('api', {
  selectFiles: () => ipcRenderer.invoke('select-files'),
  convertFiles: (filePaths) => ipcRenderer.invoke('convert-files', filePaths),
  getSheets: (filePath) => ipcRenderer.invoke('get-sheets', filePath),
  previewSheet: (filePath, sheetName) => ipcRenderer.invoke('preview-sheet', filePath, sheetName)
});