// Processus principal Electron. Ne charge que des ressources locales.
'use strict';

const { app, BrowserWindow, Menu } = require('electron');
const path = require('path');
const { checkForUpdates } = require('./update-check');

// Délai avant la vérification des mises à jour au démarrage (laisse la fenêtre
// s'afficher et se stabiliser avant d'éventuellement notifier l'utilisateur).
const STARTUP_UPDATE_DELAY_MS = 6000;

function createWindow() {
  const win = new BrowserWindow({
    width: 1180,
    height: 1020,
    minWidth: 900,
    minHeight: 640,
    backgroundColor: '#efe6d0',
    title: 'Super Constellation — Horaires L-1049',
    icon: path.join(__dirname, '..', 'build', 'icon.ico'),
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false,
    },
  });

  Menu.setApplicationMenu(null);
  win.loadFile(path.join(__dirname, 'index.html'));

  // Vérification des mises à jour, après un délai au démarrage. Silencieuse
  // s'il n'y a rien de neuf ou en cas d'absence de réseau.
  setTimeout(() => {
    if (!win.isDestroyed()) checkForUpdates(win, { silent: true });
  }, STARTUP_UPDATE_DELAY_MS);

  return win;
}

app.whenReady().then(() => {
  createWindow();
  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});
