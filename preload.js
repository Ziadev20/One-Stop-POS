const { contextBridge, ipcRenderer } = require('electron');

// Expose a safe API to the renderer (HTML app)
contextBridge.exposeInMainWorld('electronAPI', {
  // Save data: URLs (base64 files like Excel, JSON, SVG) to Downloads/EliteFoods/
  saveDataUrl: (dataUrl, filename) => ipcRenderer.invoke('save-data-url', dataUrl, filename),

  // Open a print window for receipts
  openPrintWindow: (htmlContent) => ipcRenderer.invoke('open-print-window', htmlContent),

  // Get the app data path
  getAppPath: () => ipcRenderer.invoke('get-app-path'),

  // Show native save dialog
  showSaveDialog: (options) => ipcRenderer.invoke('show-save-dialog', options),

  // Show native open dialog
  showOpenDialog: (options) => ipcRenderer.invoke('show-open-dialog', options),

  // Check if running in Electron
  isElectron: true
});

// Also expose a flag that the HTML app can check
window.addEventListener('DOMContentLoaded', () => {
  // Add a class to <html> so the CSS can detect Electron
  document.documentElement.classList.add('electron-app');

  // Override XLSX.writeFile to use Electron's file save instead of browser download
  if (typeof XLSX !== 'undefined') {
    const originalWriteFile = XLSX.writeFile;
    XLSX.writeFile = function(wb, filename, opts) {
      try {
        const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'base64' });
        const dataUrl = 'data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,' + wbout;
        window.electronAPI.saveDataUrl(dataUrl, filename).then(result => {
          if (result.success) {
            if (typeof toast === 'function') {
              toast('Excel saved: ' + filename + ' → Downloads/EliteFoods/', 's');
            }
          } else {
            // Fallback to original
            originalWriteFile.call(XLSX, wb, filename, opts);
          }
        });
      } catch (e) {
        console.error('Electron XLSX save error:', e);
        originalWriteFile.call(XLSX, wb, filename, opts);
      }
    };
  }

  // Override window.open for print windows — route through Electron
  const originalWindowOpen = window.open;
  window.open = function(url, target, features) {
    // If it's a data:text/html URL (receipt print), use Electron's print window
    if (url && url.startsWith('data:text/html')) {
      window.electronAPI.openPrintWindow(url);
      // Return a mock window object so the HTML code doesn't crash
      return {
        document: {
          write: function() {},
          close: function() {},
          open: function() { return this; }
        },
        print: function() {},
        focus: function() {},
        close: function() {},
        onload: null
      };
    }
    // For other URLs, use the original
    return originalWindowOpen.call(window, url, target, features);
  };

  // Override the printOrd function's window.open behavior
  // The HTML app calls w.document.write(html); w.document.close(); w.print();
  // We need to intercept this pattern
  let printWindowContent = '';

  // Monitor for document.write calls to print windows
  const origDocumentWrite = document.write;
  document.write = function(content) {
    origDocumentWrite.apply(document, arguments);
  };
});
