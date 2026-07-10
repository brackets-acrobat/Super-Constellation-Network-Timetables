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
  // Extension : 12 compagnies supplémentaires (générées par gen-new-airlines.js).
  { id: 'klm', name: 'KLM', file: 'Reseau_KLM_L1049_sourced.xlsx' },
  { id: 'sabena', name: 'Sabena', file: 'Reseau_Sabena_L1049_sourced.xlsx' },
  { id: 'qantas', name: 'Qantas', file: 'Reseau_Qantas_L1049_sourced.xlsx' },
  { id: 'trans-canada', name: 'Trans-Canada Air Lines', file: 'Reseau_Trans_Canada_L1049_sourced.xlsx' },
  { id: 'buffalo', name: 'Buffalo Airways', file: 'Reseau_Buffalo_L1049_sourced.xlsx' },
  { id: 'nordair', name: 'Nordair', file: 'Reseau_Nordair_L1049_sourced.xlsx' },
  { id: 'aeropostal', name: 'Línea Aeropostal Venezolana', file: 'Reseau_Aeropostal_Venezolana_L1049_sourced.xlsx' },
  { id: 'air-india', name: 'Air India', file: 'Reseau_Air_India_L1049_sourced.xlsx' },
  { id: 'pia', name: 'Pakistan International', file: 'Reseau_PIA_L1049_sourced.xlsx' },
  { id: 'varig', name: 'Varig', file: 'Reseau_Varig_L1049_sourced.xlsx' },
  { id: 'saa', name: 'South African Airways', file: 'Reseau_South_African_L1049_sourced.xlsx' },
  // 2e lot : 11 compagnies supplémentaires (générées par gen-new-airlines.js).
  { id: 'american-overseas', name: 'American Overseas Airlines', file: 'Reseau_American_Overseas_L1049_sourced.xlsx' },
  { id: 'swissair', name: 'Swissair', file: 'Reseau_Swissair_L1049_sourced.xlsx' },
  { id: 'iberia', name: 'Iberia', file: 'Reseau_Iberia_L1049_sourced.xlsx' },
  { id: 'alaska', name: 'Alaska Airlines', file: 'Reseau_Alaska_L1049_sourced.xlsx' },
  { id: 'capitol', name: 'Capitol Airways', file: 'Reseau_Capitol_L1049_sourced.xlsx' },
  { id: 'seaboard', name: 'Seaboard & Western', file: 'Reseau_Seaboard_Western_L1049_sourced.xlsx' },
  { id: 'avianca', name: 'Avianca', file: 'Reseau_Avianca_L1049_sourced.xlsx' },
  { id: 'boac', name: 'BOAC', file: 'Reseau_BOAC_L1049_sourced.xlsx' },
  { id: 'united', name: 'United Air Lines', file: 'Reseau_United_L1049_sourced.xlsx' },
  { id: 'luxair', name: 'Luxair', file: 'Reseau_Luxair_L1049_sourced.xlsx' },
  { id: 'delta', name: 'Delta Air Lines', file: 'Reseau_Delta_L1049_sourced.xlsx' },
  // 3e lot : opérateurs réels du L-1049 encore manquants.
  { id: 'northwest', name: 'Northwest Orient Airlines', file: 'Reseau_Northwest_Orient_L1049_sourced.xlsx' },
  { id: 'cubana', name: 'Cubana de Aviación', file: 'Reseau_Cubana_L1049_sourced.xlsx' },
  { id: 'real', name: 'Real Transportes Aéreos', file: 'Reseau_Real_L1049_sourced.xlsx' },
  // 4e lot : lignes FICTIVES (vraies routes DC-4 / DC-6, hors L-1049).
  { id: 'aerolineas-argentinas', name: 'Aerolíneas Argentinas', file: 'Reseau_Aerolineas_Argentinas_L1049_sourced.xlsx' },
  { id: 'olympic', name: 'Olympic Airways', file: 'Reseau_Olympic_L1049_sourced.xlsx' },
  { id: 'aviateca', name: 'Aviateca', file: 'Reseau_Aviateca_L1049_sourced.xlsx' },
  { id: 'cathay-pacific', name: 'Cathay Pacific', file: 'Reseau_Cathay_Pacific_L1049_sourced.xlsx' },
  { id: 'icelandair', name: 'Icelandair (Loftleiðir)', file: 'Reseau_Icelandair_L1049_sourced.xlsx' },
  { id: 'alitalia', name: 'Alitalia', file: 'Reseau_Alitalia_L1049_sourced.xlsx' },
  { id: 'el-al', name: 'El Al', file: 'Reseau_El_Al_L1049_sourced.xlsx' },
  { id: 'jal', name: 'Japan Airlines', file: 'Reseau_JAL_L1049_sourced.xlsx' },
  { id: 'aeromexico', name: 'Aeroméxico', file: 'Reseau_Aeromexico_L1049_sourced.xlsx' },
  { id: 'sata', name: 'SATA Air Açores', file: 'Reseau_SATA_L1049_sourced.xlsx' },
];

// Compagnies « fictives » (feuillets [R] illustratifs : appareil/compagnie non
// historiques pour le L-1049). Toutes les autres sont « historiques » (vrais
// opérateurs [S]). Sert au filtre « Lignes historiques / fictives » de l'appli.
const FICTIONAL = new Set([
  // Lignes « fictives » : vraies routes DC-4 / DC-6, mais compagnies n'ayant pas
  // exploité le L-1049 (appareil de même catégorie).
  'pan-am', 'buffalo', 'nordair', 'saa', 'american-overseas', 'swissair',
  'alaska', 'united', 'luxair', 'delta',
  'aerolineas-argentinas', 'olympic', 'aviateca', 'cathay-pacific', 'icelandair',
  'alitalia', 'el-al', 'jal', 'aeromexico', 'sata',
]);

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
    const kind = FICTIONAL.has(a.id) ? 'fictional' : 'historical';
    return { id: a.id, name: a.name, kind, sourceFile: a.file, sheets };
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
