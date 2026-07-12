const { app, BrowserWindow, Menu, shell } = require('electron');

// Points at the hosted deployment — the Anthropic API key and MSAL config
// live server-side there and never ship inside this desktop app.
const APP_URL = 'https://afocusday.onrender.com';

function createWindow() {
  const win = new BrowserWindow({
    width: 1280,
    height: 860,
    title: 'A Focus Day',
    autoHideMenuBar: true,
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: false,
    },
  });

  // Microsoft's loginPopup() and Google's sign-in button both open their
  // auth screens via window.open(). Let those become real popup windows so
  // the OAuth redirect flow completes normally; anything else (e.g. a link
  // to an external site) opens in the system browser instead.
  win.webContents.setWindowOpenHandler(({ url }) => {
    // MSAL's loginPopup() opens the popup at about:blank first, then
    // navigates it to the real login URL a moment later — that initial
    // about:blank open must be allowed too, or MSAL never gets a window
    // handle to navigate.
    if (url === 'about:blank' || url.startsWith(APP_URL) || url.startsWith('https://login.live.com') ||
        url.startsWith('https://login.microsoftonline.com') || url.startsWith('https://accounts.google.com')) {
      return { action: 'allow' };
    }
    shell.openExternal(url);
    return { action: 'deny' };
  });

  win.loadURL(APP_URL);
}

app.whenReady().then(() => {
  Menu.setApplicationMenu(null);
  createWindow();
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) createWindow();
});
