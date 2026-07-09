// Preload : expose les données compilées au renderer via un pont sécurisé.
// Le JSON est empaqueté dans l'application : aucun accès réseau ni fichier
// externe n'est requis au runtime.
'use strict';

const { contextBridge } = require('electron');
const data = require('./data/timetables.json');
const worldMap = require('./data/worldmap.json');

contextBridge.exposeInMainWorld('timetables', {
  getData: () => data,
  getMap: () => worldMap,
});
