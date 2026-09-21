const { app, BrowserWindow, Menu, dialog, shell, ipcMain } = require('electron');
const path = require('path');
const fs = require('fs');
const os = require('os');

let mainWindow = null;

// Set the app icon for the taskbar (Windows)
// This must be done before the app is ready
if (process.platform === 'win32') {
  app.setAppUserModelId('com.elitefoods.manager');
}

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1400,
    height: 900,
    minWidth: 1024,
    minHeight: 700,
    title: 'Elite Foods Manager',
    icon: path.join(__dirname, 'icon.ico'),
    backgroundColor: '#0B0D12',
    show: false,
    autoHideMenuBar: false,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      webSecurity: false,
      allowRunningInsecureContent: true,
      sandbox: false
    }
  });

  // Load the HTML file
  mainWindow.loadFile(path.join(__dirname, 'assets', 'EliteFoodsManager.html'));

  // Show window when ready (prevents white flash)
  mainWindow.once('ready-to-show', () => {
    mainWindow.show();
    // Maximize on first launch
    if (process.platform === 'win32') {
      mainWindow.maximize();
    }
  });

  // Handle external links - open in default browser
  mainWindow.webContents.setWindowOpenHandler(({ url }) => {
    if (url.startsWith('http://') || url.startsWith('https://') || url.startsWith('raast://')) {
      shell.openExternal(url);
      return { action: 'deny' };
    }
    return { action: 'allow' };
  });

  // Handle file downloads (Excel, JSON, SVG, etc.)
  mainWindow.webContents.session.on('will-download', (event, item, webContents) => {
    const fileName = item.getFilename();
    const downloadsPath = app.getPath('downloads');
    const eliteDir = path.join(downloadsPath, 'EliteFoods');

    // Create EliteFoods subfolder in Downloads
    if (!fs.existsSync(eliteDir)) {
      fs.mkdirSync(eliteDir, { recursive: true });
    }

    const filePath = path.join(eliteDir, fileName);
    item.setSavePath(filePath);

    item.on('done', (e, state) => {
      if (state === 'completed') {
        // Show notification
        if (mainWindow && !mainWindow.isDestroyed()) {
          mainWindow.webContents.executeJavaScript(
            `if(typeof toast==='function'){toast('Downloaded: ${fileName} to EliteFoods folder','s')}`
          );
        }
      } else {
        if (mainWindow && !mainWindow.isDestroyed()) {
          mainWindow.webContents.executeJavaScript(
            `if(typeof toast==='function'){toast('Download failed','e')}`
          );
        }
      }
    });
  });

  // Handle data: URL downloads (base64 encoded files from SheetJS etc.)
  // This intercepts the download trigger from the HTML app
  mainWindow.webContents.on('did-finish-load', () => {
    // Inject a download handler for data: URLs
    mainWindow.webContents.executeJavaScript(`
      (function() {
        // Override the default download behavior for data: URLs
        var originalCreateElement = document.createElement.bind(document);
        
        // Intercept <a download> clicks for data: URLs
        document.addEventListener('click', function(e) {
          var a = e.target.closest('a[download]');
          if (a && a.href && a.href.startsWith('data:')) {
            e.preventDefault();
            e.stopPropagation();
            // Send to main process for saving
            window.electronAPI.saveDataUrl(a.href, a.download || 'download');
            return false;
          }
        }, true);
        
        // Also intercept XLSX.writeFile and similar
        if (typeof XLSX !== 'undefined') {
          var originalWriteFile = XLSX.writeFile;
          XLSX.writeFile = function(wb, filename, opts) {
            try {
              var wbout = XLSX.write(wb, {bookType:'xlsx', type:'base64'});
              var dataUrl = 'data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,' + wbout;
              window.electronAPI.saveDataUrl(dataUrl, filename);
              if(typeof toast==='function'){toast('Excel saved: ' + filename,'s')}
            } catch(e) {
              console.error('XLSX save error:', e);
              if(typeof toast==='function'){toast('Excel save failed: ' + e.message,'e')}
            }
          };
        }
        
        // Override window.open for print windows
        var originalWindowOpen = window.open;
        window.open = function(url, target, features) {
          if (url && (url.startsWith('data:text/html') || (url === '' && target === '_blank'))) {
            // For print windows, create a new Electron window
            if (url.startsWith('data:text/html')) {
              window.electronAPI.openPrintWindow(url);
              return { document: { write: function(){}, close: function(){} }, print: function(){}, focus: function(){} };
            }
          }
          // For other URLs, use the original
          return originalWindowOpen.call(window, url, target, features);
        };
      })();
    `).catch(() => {});
  });

  mainWindow.on('closed', () => {
    mainWindow = null;
  });

  // Build the application menu
  buildMenu();
}

function buildMenu() {
  const template = [
    {
      label: 'File',
      submenu: [
        {
          label: 'New Order',
          accelerator: 'CmdOrCtrl+N',
          click: () => { mainWindow.webContents.executeJavaScript("if(typeof go==='function'){go('orders');if(typeof nOrder==='function'){nOrder()}}"); }
        },
        {
          label: 'Export All Data (Excel)',
          accelerator: 'CmdOrCtrl+E',
          click: () => { mainWindow.webContents.executeJavaScript("if(typeof expAll==='function'){expAll()}"); }
        },
        { type: 'separator' },
        {
          label: 'Backup Data (JSON)',
          accelerator: 'CmdOrCtrl+B',
          click: () => { mainWindow.webContents.executeJavaScript("if(typeof bkpJSON==='function'){bkpJSON()}"); }
        },
        {
          label: 'Import Data (JSON)',
          click: () => {
            dialog.showOpenDialog(mainWindow, {
              title: 'Import Backup Data',
              filters: [{ name: 'JSON Backup', extensions: ['json'] }],
              properties: ['openFile']
            }).then(result => {
              if (!result.canceled && result.filePaths.length > 0) {
                const filePath = result.filePaths[0];
                fs.readFile(filePath, 'utf-8', (err, data) => {
                  if (err) {
                    dialog.showErrorBox('Import Error', 'Failed to read file: ' + err.message);
                    return;
                  }
                  mainWindow.webContents.executeJavaScript(
                    `try{var p=JSON.parse(${JSON.stringify(data)});D={...D,...p};save();rAll();bdg();toast('Data imported','s')}catch(e){toast('Invalid file','e')}`
                  );
                });
              }
            });
          }
        },
        { type: 'separator' },
        { role: 'quit', label: 'Exit' }
      ]
    },
    {
      label: 'View',
      submenu: [
        { label: 'Dashboard', click: () => mainWindow.webContents.executeJavaScript("go('dashboard')") },
        { label: 'Orders', click: () => mainWindow.webContents.executeJavaScript("go('orders')") },
        { label: 'Menu', click: () => mainWindow.webContents.executeJavaScript("go('dishes')") },
        { label: 'Delivery', click: () => mainWindow.webContents.executeJavaScript("go('delivery')") },
        { label: 'Tax', click: () => mainWindow.webContents.executeJavaScript("go('tax')") },
        { label: 'PRA', click: () => mainWindow.webContents.executeJavaScript("go('pra')") },
        { label: 'SMS', click: () => mainWindow.webContents.executeJavaScript("go('sms')") },
        { label: 'Digital Menu', click: () => mainWindow.webContents.executeJavaScript("go('digitalmenu')") },
        { type: 'separator' },
        { label: 'Settings', click: () => mainWindow.webContents.executeJavaScript("go('settings')") },
        { type: 'separator' },
        { role: 'reload', label: 'Reload' },
        { role: 'toggleDevTools', label: 'Developer Tools' },
        { type: 'separator' },
        { role: 'resetZoom', label: 'Reset Zoom' },
        { role: 'zoomIn', label: 'Zoom In' },
        { role: 'zoomOut', label: 'Zoom Out' },
        { role: 'togglefullscreen', label: 'Fullscreen' }
      ]
    },
    {
      label: 'Theme',
      submenu: [
        { label: 'Light', click: () => mainWindow.webContents.executeJavaScript("sTh('light')") },
        { label: 'Dark', click: () => mainWindow.webContents.executeJavaScript("sTh('dark')") },
        { label: 'Elite (Black & Gold)', click: () => mainWindow.webContents.executeJavaScript("sTh('elite')") },
        { label: 'Ocean Deep', click: () => mainWindow.webContents.executeJavaScript("sTh('oceandeep')") },
        { label: 'Sunset', click: () => mainWindow.webContents.executeJavaScript("sTh('sunset')") },
        { label: 'Forest', click: () => mainWindow.webContents.executeJavaScript("sTh('forest')") },
        { type: 'separator' },
        { label: 'Cycle Themes', accelerator: 'CmdOrCtrl+T', click: () => mainWindow.webContents.executeJavaScript("cycleTh()") }
      ]
    },
    {
      label: 'Help',
      submenu: [
        {
          label: 'About Elite Foods Manager',
          click: () => {
            dialog.showMessageBox(mainWindow, {
              type: 'info',
              title: 'About',
              message: 'Elite Foods Manager v9.0',
              detail: 'Complete Restaurant Management System\nBuilt for Pakistan\n\nFeatures:\n• Orders & Menu Management\n• FBR Tax + PRA e-IMS Integration\n• Delivery (foodpanda/Cheetay)\n• SMS Notifications\n• Digital Menu QR Code\n• Urdu Language Support\n• Thermal Printer Support\n• 19+ Themes\n• Multi-Device Sync\n\n© 2026 Elite Foods',
              icon: path.join(__dirname, 'icon.ico'),
              buttons: ['OK']
            });
          }
        },
        { label: 'Guide', click: () => mainWindow.webContents.executeJavaScript("go('guide')") },
        { type: 'separator' },
        { label: 'Open Data Folder', click: () => { shell.openPath(app.getPath('userData')); } },
        { label: 'Open Downloads Folder', click: () => { shell.openPath(path.join(app.getPath('downloads'), 'EliteFoods')); } }
      ]
    }
  ];

  const menu = Menu.buildFromTemplate(template);
  Menu.setApplicationMenu(menu);
}

// Print window handler
let printWindows = [];

ipcMain.handle('save-data-url', async (event, dataUrl, filename) => {
  try {
    const downloadsPath = app.getPath('downloads');
    const eliteDir = path.join(downloadsPath, 'EliteFoods');
    if (!fs.existsSync(eliteDir)) {
      fs.mkdirSync(eliteDir, { recursive: true });
    }

    // Parse data URL
    const base64Data = dataUrl.split('base64,')[1];
    if (!base64Data) {
      // Handle non-base64 data URLs
      const textData = decodeURIComponent(dataUrl.split(',')[1]);
      const filePath = path.join(eliteDir, filename);
      fs.writeFileSync(filePath, textData, 'utf-8');
      return { success: true, path: filePath };
    }

    const buffer = Buffer.from(base64Data, 'base64');
    const filePath = path.join(eliteDir, filename);
    fs.writeFileSync(filePath, buffer);
    return { success: true, path: filePath };
  } catch (error) {
    console.error('Save error:', error);
    return { success: false, error: error.message };
  }
});

ipcMain.handle('open-print-window', async (event, htmlContent) => {
  try {
    const printWin = new BrowserWindow({
      width: 400,
      height: 700,
      title: 'Receipt',
      parent: mainWindow,
      modal: false,
      webPreferences: {
        contextIsolation: true,
        nodeIntegration: false
      }
    });

    printWindows.push(printWin);

    // Load the HTML content
    const dataUrl = htmlContent;
    await printWin.loadURL(dataUrl);

    // Auto-print after a short delay
    setTimeout(() => {
      printWin.webContents.print({ silent: false, printBackground: true }, (success) => {
        if (!success) {
          console.log('Print was canceled or failed');
        }
      });
    }, 500);

    printWin.on('closed', () => {
      printWindows = printWindows.filter(w => w !== printWin);
    });

    return { success: true };
  } catch (error) {
    console.error('Print window error:', error);
    return { success: false, error: error.message };
  }
});

ipcMain.handle('get-app-path', async () => {
  return app.getPath('userData');
});

ipcMain.handle('show-save-dialog', async (event, options) => {
  const result = await dialog.showSaveDialog(mainWindow, options);
  return result;
});

ipcMain.handle('show-open-dialog', async (event, options) => {
  const result = await dialog.showOpenDialog(mainWindow, options);
  return result;
});

// App lifecycle
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

// Prevent multiple instances
const gotTheLock = app.requestSingleInstanceLock();
if (!gotTheLock) {
  app.quit();
} else {
  app.on('second-instance', () => {
    if (mainWindow) {
      if (mainWindow.isMinimized()) mainWindow.restore();
      mainWindow.focus();
    }
  });
}
