const { app, BrowserWindow, ipcMain, dialog, shell } = require('electron');
const path = require('path');
const fs = require('fs');

ipcMain.handle('save-excel-file', async (event, buffer, defaultFileName) => {
  const { canceled, filePath } = await dialog.showSaveDialog({
    title: '保存文件',
    defaultPath: defaultFileName,
    filters: [
      { name: 'Excel Workbook', extensions: ['xlsx'] }
    ]
  });

  if (canceled || !filePath) return null;

  fs.writeFileSync(filePath, Buffer.from(buffer));
  return filePath;
});

ipcMain.handle('open-file', async (event, filePath) => {
  if (filePath) {
    await shell.openPath(filePath);
  }
});

ipcMain.handle('select-directory', async (event) => {
  const { canceled, filePaths } = await dialog.showOpenDialog({
    title: '选择导出文件夹',
    properties: ['openDirectory']
  });
  if (canceled || filePaths.length === 0) return null;
  return filePaths[0];
});

ipcMain.handle('save-file-to-path', async (event, dirPath, fileName, buffer) => {
  const fullPath = path.join(dirPath, fileName);
  fs.writeFileSync(fullPath, Buffer.from(buffer));
  return fullPath;
});


function createWindow() {
  const win = new BrowserWindow({
    width: 1400,
    height: 900,
    minWidth: 1024,
    minHeight: 768,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      preload: path.join(__dirname, 'preload.cjs'),
      webSecurity: false // allow local files loading for some plugins if needed
    },
    autoHideMenuBar: true, // Hide menu bar
    show: false // Wait until ready-to-show
  });

  // Load the React build
  win.loadFile(path.join(__dirname, '../dist/index.html'));

  win.once('ready-to-show', () => {
    win.show();
  });
}

app.whenReady().then(() => {
  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});
