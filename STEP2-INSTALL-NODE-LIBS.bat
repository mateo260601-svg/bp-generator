const { app, BrowserWindow, ipcMain, dialog, shell, Menu } = require('electron');
const path = require('path');
const { spawn } = require('child_process');
const http = require('http');
const fs = require('fs');

const PORT = 8765;
let mainWindow = null;
let pythonProcess = null;

const backendDir = app.isPackaged
  ? path.join(process.resourcesPath, 'backend')
  : path.join(__dirname, '..', 'backend');

function startBackend() {
  const python = process.platform === 'win32' ? 'python' : 'python3';
  pythonProcess = spawn(python, [
    '-m', 'uvicorn', 'server:app',
    '--host', '127.0.0.1',
    '--port', String(PORT)
  ], {
    cwd: backendDir,
    windowsHide: true
  });
  pythonProcess.stdout.on('data', d => console.log('[py]', d.toString().trim()));
  pythonProcess.stderr.on('data', d => console.log('[py]', d.toString().trim()));
}

function waitForBackend(tries) {
  tries = tries || 40;
  return new Promise(function(resolve, reject) {
    function attempt() {
      http.get('http://127.0.0.1:' + PORT + '/api/health', function(res) {
        if (res.statusCode === 200) resolve();
        else retry();
      }).on('error', retry);
    }
    function retry() {
      if (--tries <= 0) return reject(new Error('Backend timeout'));
      setTimeout(attempt, 500);
    }
    attempt();
  });
}

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1200,
    height: 800,
    minWidth: 900,
    minHeight: 600,
    icon: path.join(__dirname, '..', 'assets', 'icon.ico'),
    title: 'BP Generator',
    show: false,
    backgroundColor: '#f0f2f5',
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false
    }
  });

  Menu.setApplicationMenu(null);
  mainWindow.loadURL('http://127.0.0.1:' + PORT);
  mainWindow.once('ready-to-show', function() {
    mainWindow.show();
  });
  mainWindow.on('closed', function() {
    mainWindow = null;
  });
}

ipcMain.handle('download-file', async function(event, opts) {
  var filename = opts.filename;
  var licenseKey = opts.licenseKey;
  var result = await dialog.showSaveDialog(mainWindow, {
    defaultPath: filename,
    filters: [{ name: 'Excel Files', extensions: ['xlsx'] }]
  });
  if (result.canceled || !result.filePath) return { ok: false };
  return new Promise(function(resolve) {
    var req = http.get(
      'http://127.0.0.1:' + PORT + '/api/download/' + filename,
      { headers: { 'x-license': licenseKey } },
      function(res) {
        if (res.statusCode !== 200) {
          resolve({ ok: false });
          return;
        }
        var out = fs.createWriteStream(result.filePath);
        res.pipe(out);
        out.on('finish', function() {
          out.close();
          shell.showItemInFolder(result.filePath);
          resolve({ ok: true, path: result.filePath });
        });
      }
    );
    req.on('error', function(e) { resolve({ ok: false, error: e.message }); });
  });
});

app.whenReady().then(async function() {
  startBackend();
  try { await waitForBackend(); } catch(e) { console.error(e.message); }
  createWindow();
});

app.on('window-all-closed', function() {
  if (pythonProcess) pythonProcess.kill();
  if (process.platform !== 'darwin') app.quit();
});

app.on('before-quit', function() {
  if (pythonProcess) pythonProcess.kill();
});
