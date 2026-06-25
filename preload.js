// filename: preload.js
const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("api", {
  selectWorkspaceDir: () => ipcRenderer.invoke("select-workspace-dir"),
  getWorkspaceDir: () => ipcRenderer.invoke("get-workspace-dir"),
  selectFiles: () => ipcRenderer.invoke("select-files"),
  convertFiles: (filePaths) => ipcRenderer.invoke("convert-files", filePaths),
  getSheets: (filePath) => ipcRenderer.invoke("get-sheets", filePath),
  previewSheet: (filePath, sheetName) => ipcRenderer.invoke("preview-sheet", filePath, sheetName),
  runValidation: (payload) => ipcRenderer.invoke("run-validation", payload),
  openReport: (reportPath) => ipcRenderer.invoke("open-report", reportPath),
  selectTemplateFile: () => ipcRenderer.invoke("select-template-file"),
  runTrialBalance: (payload) => ipcRenderer.invoke("run-trial-balance", payload)
});