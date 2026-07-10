// Preload : expose les données compilées au renderer via un pont sécurisé.
// Le JSON est empaqueté dans l'application : aucun accès réseau ni fichier
// externe n'est requis au runtime.
'use strict';

const { contextBridge, shell } = require('electron');
const data = require('./data/timetables.json');
const worldMap = require('./data/worldmap.json');
const pkg = require('../package.json');

contextBridge.exposeInMainWorld('timetables', {
  getData: () => data,
  getMap: () => worldMap,
  version: pkg.version,
  // Ouvre un lien http(s) dans le navigateur par défaut (jamais dans l'app).
  openExternal: (url) => {
    if (typeof url === 'string' && /^https?:\/\//i.test(url)) shell.openExternal(url);
  },
});
