const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const { spawn } = require('child_process');
const axios = require('axios');

let mainWindow;
let pyProcess;
const PY_PORT = 5123;
const PY_BASE = `http://127.0.0.1:${PY_PORT}`;

function startPythonBackend() {
  // Dev mode: run python script directly
  // Prod mode: run bundled app.exe (PyInstaller output)
  const isDev = !app.isPackaged;

  if (isDev) {
    pyProcess = spawn('python', [path.join(__dirname, 'python_backend', 'app.py')], {
      cwd: path.join(__dirname, 'python_backend')
    });
  } else {
    const exePath = path.join(process.resourcesPath, 'python_backend', 'app.exe');
    pyProcess = spawn(exePath, [], {
      cwd: path.join(process.resourcesPath, 'python_backend')
    });
  }

  pyProcess.stdout.on('data', (d) => console.log(`[PY] ${d}`));
  pyProcess.stderr.on('data', (d) => console.error(`[PY-ERR] ${d}`));
  pyProcess.on('close', (code) => console.log(`Python backend exited: ${code}`));
}

function waitForBackend(retries = 20) {
  return new Promise((resolve, reject) => {
    const tryPing = (n) => {
      axios.get(`${PY_BASE}/health`)
        .then(() => resolve(true))
        .catch(() => {
          if (n <= 0) return reject(new Error('Python backend did not start'));
          setTimeout(() => tryPing(n - 1), 500);
        });
    };
    tryPing(retries);
  });
}

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1100,
    height: 750,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false
    }
  });

  if (process.env.ELECTRON_DEV === '1') {
    mainWindow.loadURL('http://localhost:5173');
  } else {
    mainWindow.loadFile(path.join(__dirname, 'dist-renderer', 'index.html'));
  }
}

app.whenReady().then(async () => {
  startPythonBackend();
  try {
    await waitForBackend();
  } catch (e) {
    console.error(e);
  }
  createWindow();
});

app.on('window-all-closed', () => {
  if (pyProcess) pyProcess.kill();
  if (process.platform !== 'darwin') app.quit();
});

// ---- IPC Handlers ----

ipcMain.handle('select-files', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    properties: ['openFile', 'multiSelections'],
    filters: [{ name: 'Excel Files', extensions: ['xls', 'xlsx', 'xlsm'] },
  { name: 'All Files', extensions: ['*'] }]
  });
  return result.filePaths;
});

ipcMain.handle('convert-files', async (event, filePaths) => {
  const res = await axios.post(`${PY_BASE}/convert`, { files: filePaths });
  return res.data;
});

ipcMain.handle('get-sheets', async (event, filePath) => {
  const res = await axios.post(`${PY_BASE}/sheets`, { file: filePath });
  return res.data;
});

ipcMain.handle('preview-sheet', async (event, filePath, sheetName) => {
  const res = await axios.post(`${PY_BASE}/preview`, { file: filePath, sheet: sheetName });
  return res.data;
});