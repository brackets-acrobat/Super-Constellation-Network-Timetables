// Compile les 3 fichiers Excel en un unique timetables.json embarqué dans l'app.
//
//   node scripts/compile-data.js
//
// À exécuter UNIQUEMENT à la phase de préparation (build). L'application
// Electron livrée ne lit que src/data/timetables.json : elle ne dépend ni de
// la librairie xlsx, ni des fichiers .xlsx d'origine.

'use strict';

const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');
const { parseSheet } = require('./parse-lib');
const { geocode } = require('./place-coords');
const T = require('./translations');

const SRC_DIR = path.resolve(__dirname, '../../liaisons_Super_Constellation');
const OUT = path.resolve(__dirname, '../src/data/timetables.json');

const AIRLINES = [
  { id: 'twa', name: 'TWA', file: 'Reseau_TWA_L1049_sourced.xlsx' },
  { id: 'air-france', name: 'Air France', file: 'Reseau_Air_France_L1049_sourced.xlsx' },
  { id: 'lufthansa', name: 'Lufthansa', file: 'Reseau_Lufthansa_L1049G_sourced.xlsx' },
  { id: 'eastern-air-lines', name: 'Eastern Air Lines', file: 'Reseau_Eastern_Air_Lines_L1049_sourced.xlsx' },
  { id: 'flying-tiger', name: 'Flying Tiger', file: 'Reseau_Flying_Tiger_L1049H_sourced.xlsx' },
  { id: 'irish-airlines', name: 'Irish Airlines', file: 'Reseau_Irish_Airlines_L1049_sourced.xlsx' },
  { id: 'pan-am', name: 'Pan Am', file: 'Reseau_Pan_Am_L1049_sourced.xlsx' },
];

function slug(s) {
  return String(s)
    .toLowerCase()
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/(^-|-$)/g, '');
}

const missing = new Set();
const trMisses = new Set();

function withCoords(itinerary) {
  return itinerary.map((stop) => {
    const g = geocode(stop.place);
    if (!g) missing.add(stop.place);
    return { ...stop, lat: g ? g.lat : null, lon: g ? g.lon : null };
  });
}

// Traduit en anglais une étape d'itinéraire / escale (lieu + horaires).
function translateStop(s) {
  const out = { ...s };
  if (s.place != null) out.place = T.translatePlace(s.place);
  if (s.arrival != null) out.arrival = T.translateTime(s.arrival);
  if (s.departure != null) out.departure = T.translateTime(s.departure);
  return out;
}

// Construit la version anglaise d'un vol (mêmes id / lat / lon / marqueurs).
function translateFlight(f) {
  return {
    ...f,
    line: T.lookup(T.LINES, f.line, trMisses, 'LINE'),
    lineHeader: T.lookup(T.LINE_HEADERS, f.lineHeader, trMisses, 'LINE_HEADER'),
    frequency: T.lookup(T.FREQ, f.frequency, trMisses, 'FREQ'),
    frequencyHeader: T.lookup(T.FREQ_HEADERS, f.frequencyHeader, trMisses, 'FREQ_HEADER'),
    meta: (f.meta || []).map((m) => ({
      header: T.lookup(T.META_HEADERS, m.header, trMisses, 'META_HEADER'),
      value: m.value,
    })),
    origin: T.translatePlace(f.origin),
    destination: T.translatePlace(f.destination),
    departure: T.translateTime(f.departure),
    arrival: T.translateTime(f.arrival),
    stops: (f.stops || []).map(translateStop),
    itinerary: (f.itinerary || []).map(translateStop),
  };
}

function translateSheet(s) {
  return {
    ...s,
    name: T.lookup(T.SHEETS, s.name, trMisses, 'SHEET'),
    notes: (s.notes || []).map((n) => T.lookup(T.NOTES, n, trMisses, 'NOTE')),
    flights: s.flights.map(translateFlight),
  };
}

function translateAirlines(airlines) {
  // Le nom de compagnie est conservé tel quel (TWA, Pan Am, ...).
  return airlines.map((a) => ({ ...a, sheets: a.sheets.map(translateSheet) }));
}

function main() {
  const airlines = AIRLINES.map((a) => {
    const wb = XLSX.readFile(path.join(SRC_DIR, a.file));
    const sheets = wb.SheetNames.map((sn) => {
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, defval: '', raw: false });
      const { flights, notes } = parseSheet(rows);
      const enriched = flights.map((f, i) => ({
        id: `${a.id}-${slug(sn)}-${i}`,
        ...f,
        itinerary: withCoords(f.itinerary),
      }));
      return { id: slug(sn), name: sn, notes, flights: enriched };
    });
    return { id: a.id, name: a.name, sourceFile: a.file, sheets };
  });

  const airlinesEn = translateAirlines(airlines);

  const data = {
    generatedAt: new Date().toISOString(),
    aircraft: 'Lockheed L-1049 Super Constellation',
    airlines,        // version française (par défaut historique des données)
    airlinesEn,      // version anglaise, mêmes identifiants
  };

  fs.mkdirSync(path.dirname(OUT), { recursive: true });
  fs.writeFileSync(OUT, JSON.stringify(data, null, 2), 'utf8');

  const flightCount = airlines.reduce(
    (n, a) => n + a.sheets.reduce((m, s) => m + s.flights.length, 0), 0);

  console.log(`✓ ${OUT}`);
  console.log(`  ${airlines.length} compagnies, ${flightCount} vols compilés (FR + EN).`);
  if (missing.size) {
    console.warn(`\n⚠ Lieux SANS coordonnées (${missing.size}) :`);
    [...missing].sort().forEach((p) => console.warn('   - ' + p));
    process.exitCode = 1;
  } else {
    console.log('  Tous les lieux sont géocodés ✓');
  }
  if (trMisses.size) {
    console.warn(`\n⚠ Chaînes SANS traduction EN (${trMisses.size}) :`);
    [...trMisses].sort().forEach((p) => console.warn('   - ' + p));
    process.exitCode = 1;
  } else {
    console.log('  Toutes les chaînes sont traduites en anglais ✓');
  }
}

main();
