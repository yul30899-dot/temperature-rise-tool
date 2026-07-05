const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
  saveExcelFile: (buffer, defaultFileName) => ipcRenderer.invoke('save-excel-file', buffer, defaultFileName),
  openFile: (filePath) => ipcRenderer.invoke('open-file', filePath),
  selectDirectory: () => ipcRenderer.invoke('select-directory'),
  saveFileToPath: (dirPath, fileName, buffer) => ipcRenderer.invoke('save-file-to-path', dirPath, fileName, buffer)
});
