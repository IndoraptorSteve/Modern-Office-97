'use strict';

const { app, BrowserWindow, ipcMain, dialog, Menu } = require('electron');
const path = require('path');
const fs = require('fs');

const args = process.argv.slice(2);
let mainWin = null;

/* ---- Pixelated font rendering - disable ALL antialiasing ---- */
app.commandLine.appendSwitch('disable-lcd-text');
app.commandLine.appendSwitch('disable-font-subpixel-positioning');
app.commandLine.appendSwitch('disable-gpu-driver-bug-workarounds');
app.commandLine.appendSwitch('disable-gpu-rasterization');
app.commandLine.appendSwitch('disable-accelerated-2d-canvas');
app.commandLine.appendSwitch('font-render-hinting', 'none');

/* ---- First-run flag ---- */
const SETUP_DONE_FLAG = path.join(app.getPath ? app.getPath('userData') : __dirname, '.office97-setup-done');

function isFirstRun() {
  // Skip setup if --no-setup flag given or if running a specific app
  if (args.includes('--no-setup') || args.includes('--word') || args.includes('--excel') ||
      args.includes('--powerpoint') || args.includes('--access')) {
    return false;
  }
  return !fs.existsSync(SETUP_DONE_FLAG);
}

function markSetupDone() {
  try {
    fs.writeFileSync(SETUP_DONE_FLAG, new Date().toISOString(), 'utf8');
  } catch(e) {
    // Ignore errors writing flag
  }
}

/* ---- Installer Splash Window ---- */
function createInstallerSplash(nextWindow) {
  const win = new BrowserWindow({
    width: 640,
    height: 480,
    title: 'Microsoft Office 97 Setup',
    resizable: false,
    frame: false,
    icon: path.join(__dirname, 'icons/office.ico'),
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      preload: path.join(__dirname, 'installer-splash-preload.js'),
      zoomFactor: 1.0,
      defaultFontSize: 13,
      defaultMonospaceFontSize: 13,
    }
  });

  win.setMenu(null);
  win.loadFile(path.join(__dirname, 'installer-splash.html'));

  // When installer completes, close splash and show next window
  ipcMain.once('installer-complete', () => {
    win.close();
    if (nextWindow === 'setup') {
      createSetupWindow();
    } else if (nextWindow === 'launcher') {
      mainWin = createLauncherWindow();
      setupIpc(mainWin);
    }
  });

  return win;
}

/* ---- Setup window ---- */
function createSetupWindow() {
  const win = new BrowserWindow({
    width: 560,
    height: 480,
    title: 'Microsoft Office 97 Setup',
    resizable: false,
    icon: path.join(__dirname, 'icons/office.ico'),
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      preload: path.join(__dirname, 'setup', 'setup-preload.js'),
      zoomFactor: 1.0,
      defaultFontSize: 13,
      defaultMonospaceFontSize: 13,
    }
  });

  win.setMenu(null);
  win.loadFile(path.join(__dirname, 'setup', 'setup.html'));

  // When setup finishes, close setup and open launcher
  ipcMain.once('setup:finish', () => {
    markSetupDone();
    win.close();
    mainWin = createLauncherWindow();
    setupIpc(mainWin);
  });

  // When setup exits
  ipcMain.once('setup:exit', () => {
    app.quit();
  });

  return win;
}

function getAppToLaunch() {
  if (args.includes('--word')) return 'word';
  if (args.includes('--excel')) return 'excel';
  if (args.includes('--powerpoint') || args.includes('--ppt')) return 'powerpoint';
  if (args.includes('--access')) return 'access';
  if (args.includes('--outlook')) return 'outlook';
  return null;
}

/* ---- Shared IPC Handlers ---- */
function setupIpc(win) {
  ipcMain.handle('dialog:open', async (event, options) => {
    return await dialog.showOpenDialog(win, options);
  });

  ipcMain.handle('dialog:save', async (event, options) => {
    return await dialog.showSaveDialog(win, options);
  });

  ipcMain.handle('file:read', async (event, filePath) => {
    try {
      const data = fs.readFileSync(filePath);
      return { success: true, data: data.toString('base64'), path: filePath };
    } catch (e) {
      return { success: false, error: e.message };
    }
  });

  ipcMain.handle('file:write', async (event, filePath, data) => {
    try {
      fs.writeFileSync(filePath, Buffer.from(data, 'base64'));
      return { success: true };
    } catch (e) {
      return { success: false, error: e.message };
    }
  });

  ipcMain.handle('file:readText', async (event, filePath) => {
    try {
      return { success: true, data: fs.readFileSync(filePath, 'utf8'), path: filePath };
    } catch (e) {
      return { success: false, error: e.message };
    }
  });

  ipcMain.handle('file:writeText', async (event, filePath, data) => {
    try {
      fs.writeFileSync(filePath, data, 'utf8');
      return { success: true };
    } catch (e) {
      return { success: false, error: e.message };
    }
  });

  ipcMain.handle('ppt:export', async (event, data) => {
    return { success: true };
  });
}

/* ---- Launch Specific App ---- */
function launchApp(appName) {
  const appDir = path.join(__dirname, appName);
  const htmlFile = path.join(appDir, appName + '.html');
  const preloadFile = path.join(appDir, appName + '-preload.js');

  const titles = {
    word: 'Microsoft Word',
    excel: 'Microsoft Excel',
    powerpoint: 'Microsoft PowerPoint',
    access: 'Microsoft Access',
    outlook: 'Inbox - Microsoft Outlook'
  };

  const sizes = {
    word: { width: 900, height: 700 },
    excel: { width: 1000, height: 720 },
    powerpoint: { width: 1000, height: 720 },
    access: { width: 900, height: 680 },
    outlook: { width: 950, height: 720 }
  };

  const win = new BrowserWindow({
    ...( sizes[appName] || { width: 900, height: 700 }),
    title: titles[appName] || 'Microsoft Office 97',
    icon: path.join(__dirname, 'icons', appName + '97.ico'),
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      preload: fs.existsSync(preloadFile) ? preloadFile : undefined,
      zoomFactor: 1.0,
      defaultFontSize: 13,
      defaultMonospaceFontSize: 13,
    }
  });

  win.loadFile(htmlFile);
  win.setMenu(null);
  return win;
}

/* ---- Launcher Window ---- */
function createLauncher() {
  const win = new BrowserWindow({
    width: 320,
    height: 240,
    title: 'Microsoft Office 97',
    resizable: false,
    icon: path.join(__dirname, 'icons/office.ico'),
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      zoomFactor: 1.0,
      defaultFontSize: 13,
      defaultMonospaceFontSize: 13,
    }
  });

  win.setMenu(null);

  // Inline HTML for launcher
  const launcherHtml = `<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>Microsoft Office 97</title>
  <link rel="stylesheet" href="shared/win95-ui.css">
  <style>
    body { display:flex;flex-direction:column;align-items:center;justify-content:center;height:100%;gap:8px;padding:16px;background:#c0c0c0; }
    h2 { font-size:13px;font-weight:bold;color:#000080;margin-bottom:8px;font-family:Tahoma,Arial,sans-serif; }
    .app-grid { display:grid;grid-template-columns:1fr 1fr;gap:12px;width:100%; }
    .app-btn {
      background:#c0c0c0;
      border:2px solid;
      border-color:#fff #808080 #808080 #fff;
      padding:12px 8px;
      cursor:pointer;
      text-align:center;
      font-family:Tahoma,Arial,sans-serif;
      font-size:11px;
      display:flex;
      flex-direction:column;
      align-items:center;
      gap:6px;
    }
    .app-btn:hover { background:#d4d0c8; }
    .app-btn:active { border-color:#808080 #fff #fff #808080; }
    .app-icon { width:32px;height:32px;display:block; }
    .word-icon { background:#2b579a; }
    .excel-icon { background:#217346; }
    .ppt-icon { background:#d04423; }
    .access-icon { background:#a4373a; }
    .app-icon span { display:block;color:#fff;font-size:20px;font-weight:bold;line-height:32px;text-align:center;font-family:Times New Roman,serif; }
  </style>
</head>
<body>
  <h2>Microsoft Office 97</h2>
  <div class="app-grid">
    <div class="app-btn" onclick="launch('word')">
      <div class="app-icon word-icon"><span>W</span></div>
      <span>Word</span>
    </div>
    <div class="app-btn" onclick="launch('excel')">
      <div class="app-icon excel-icon"><span>X</span></div>
      <span>Excel</span>
    </div>
    <div class="app-btn" onclick="launch('powerpoint')">
      <div class="app-icon ppt-icon"><span>P</span></div>
      <span>PowerPoint</span>
    </div>
    <div class="app-btn" onclick="launch('access')">
      <div class="app-icon access-icon"><span>A</span></div>
      <span>Access</span>
    </div>
  </div>
  <script>
    function launch(app) {
      // Signal to main process via URL hash
      window.location.hash = app;
    }
    // Poll for hash change
    window.addEventListener('hashchange', () => {
      const app = window.location.hash.slice(1);
      if (app) {
        fetch('http://localhost:9999/launch/' + app).catch(() => {});
      }
    });
  </script>
</body>
</html>`;

  win.loadURL('data:text/html,' + encodeURIComponent(launcherHtml));

  return win;
}

/* ---- App Startup ---- */
app.whenReady().then(() => {
  const appToLaunch = getAppToLaunch();

  if (appToLaunch) {
    // Direct app launch - skip splash and setup
    const win = launchApp(appToLaunch);
    setupIpc(win);
  } else if (isFirstRun()) {
    // Show installer splash, then setup
    createInstallerSplash('setup');
  } else {
    // Show installer splash, then launcher
    createInstallerSplash('launcher');
  }
});

function createLauncherWindow() {
  const win = new BrowserWindow({
    width: 420,
    height: 280,
    title: 'Microsoft Office 97',
    resizable: false,
    icon: path.join(__dirname, 'icons/office.ico'),
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      preload: path.join(__dirname, 'launcher-preload.js'),
      zoomFactor: 1.0,
      defaultFontSize: 13,
      defaultMonospaceFontSize: 13,
    }
  });

  win.setMenu(null);
  win.loadFile(path.join(__dirname, 'launcher.html'));
  return win;
}

// IPC: launch another Office app
ipcMain.on('launch-app', (event, appName, appArgs) => {
  launchApp(appName, appArgs);
});

// IPC: send-to-mail opens Outlook compose with doc info
ipcMain.on('send-to-mail', (event, docInfo) => {
  // Check if outlook window already open
  const wins = BrowserWindow.getAllWindows();
  const outlookWin = wins.find(w => w.getTitle && w.getTitle().includes('Outlook'));
  if (outlookWin) {
    outlookWin.webContents.send('open-compose', docInfo);
    outlookWin.show();
    outlookWin.focus();
  } else {
    // Launch outlook and then open compose
    const newWin = launchApp('outlook');
    newWin.webContents.on('did-finish-load', () => {
      setTimeout(() => newWin.webContents.send('open-compose', docInfo), 500);
    });
  }
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) {
    const appToLaunch = getAppToLaunch();
    if (appToLaunch) launchApp(appToLaunch);
    else mainWin = createLauncherWindow();
  }
});
