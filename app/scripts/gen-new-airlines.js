// Générateur des 12 NOUVELLES compagnies (jeu « sourced »).
//
//   node scripts/gen-new-airlines.js
//
// Produit, dans liaisons_Super_Constellation/ :
//   - un fichier Reseau_<Compagnie>_L1049_sourced.xlsx par compagnie, au FORMAT
//     EXACT des fichiers existants (mêmes en-têtes, notation [S]/[R], escales).
// Produit aussi, dans scripts/ :
//   - translations-extra-2.js  (traductions EN de toutes les chaînes ajoutées)
//   - place-coords-2.js        (coordonnées lat/lon des villes ajoutées)
//
// Chaque route est définie UNE fois (aller) ; le vol RETOUR est engendré
// automatiquement par itinéraire miroir + moteur d'horloge (horaires cohérents,
// marqueurs « * » de +1 jour recalculés). Les deux sens sont donc toujours
// présents. Numérotation : aller impair/pair ci-dessous, retour = aller+1.

'use strict';

const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');

// ---------------------------------------------------------------------------
// Dictionnaires de traduction accumulés au fil de la définition des routes.
// ---------------------------------------------------------------------------
const SHEETS = {};
const LINES = {};
const FREQ = {};
const NOTES = {};
const CITY = {
  'Bruxelles': 'Brussels',
  'Singapour': 'Singapore',
  'Djakarta': 'Jakarta',
  'Genève': 'Geneva',
  'Athènes': 'Athens',
  'Zürich': 'Zurich',
  'Copenhague': 'Copenhagen',
  'Bermudes': 'Bermuda',
  'Mexico': 'Mexico City',
  'Guatemala': 'Guatemala City',
};

// Coordonnées des villes ajoutées (clé = libellé normalisé minuscule).
const COORDS = {
  'amsterdam': [52.31, 4.76],
  'prestwick': [55.51, -4.59],
  'djakarta': [-6.21, 106.85],
  'bruxelles': [50.85, 4.35],
  'léopoldville': [-4.32, 15.31],
  'leopoldville': [-4.32, 15.31],
  'sydney': [-33.87, 151.21],
  'darwin': [-12.46, 130.84],
  'singapour': [1.35, 103.82],
  'nandi': [-17.75, 177.45],
  'canton island': [-2.83, -171.68],
  'toronto': [43.65, -79.38],
  'winnipeg': [49.90, -97.14],
  'vancouver': [49.28, -123.12],
  'edmonton': [53.55, -113.49],
  'yellowknife': [62.45, -114.37],
  'churchill': [58.77, -94.17],
  'frobisher bay': [63.75, -68.52],
  "val-d'or": [48.10, -77.78],
  'maracaibo': [10.65, -71.64],
  'ciudad trujillo': [18.47, -69.89],
  'genève': [46.20, 6.14],
  'geneve': [46.20, 6.14],
  'hong kong': [22.32, 114.17],
  'dacca': [23.81, 90.41],
  'belém': [-1.46, -48.50],
  'belem': [-1.46, -48.50],
  'belgrade': [44.79, 20.45],
  'zagreb': [45.81, 15.98],
  'zürich': [47.37, 8.54],
  'zurich': [47.37, 8.54],
  'athènes': [37.98, 23.73],
  'athenes': [37.98, 23.73],
  'johannesburg': [-26.20, 28.05],
  'nairobi': [-1.29, 36.82],
  'khartoum': [15.50, 32.56],
  // 2e lot (11 compagnies supplémentaires).
  'copenhague': [55.68, 12.57],
  'seattle': [47.61, -122.33],
  'anchorage': [61.22, -149.90],
  'fairbanks': [64.84, -147.72],
  'nome': [64.50, -165.41],
  'palma': [39.57, 2.65],
  'luxembourg': [49.63, 6.21],
  'barranquilla': [10.96, -74.80],
  'bermudes': [32.30, -64.79],
  'nassau': [25.05, -77.35],
  'nice': [43.66, 7.21],
  'dallas': [32.78, -96.80],
  // 3e lot (opérateurs réels manquants).
  'mexico': [19.43, -99.13],
  'manaus': [-3.10, -60.02],
  // 4e lot (lignes fictives DC-4 / DC-6).
  'santiago': [-33.45, -70.67],
  'montevideo': [-34.90, -56.19],
  'sao paulo': [-23.55, -46.63],
  'são paulo': [-23.55, -46.63],
  'rhodes': [36.43, 28.22],
  'guatemala': [14.63, -90.51],
  'belize': [17.50, -88.20],
  'san salvador': [13.69, -89.22],
  'san jose': [9.94, -84.08],
  'reykjavik': [64.13, -21.90],
  'oslo': [59.91, 10.75],
  'milan': [45.46, 9.19],
  'tripoli': [32.89, 13.19],
  'tel aviv': [32.01, 34.89],
  'osaka': [34.69, 135.50],
  'fukuoka': [33.59, 130.40],
  'sapporo': [43.06, 141.35],
  'tijuana': [32.53, -116.97],
  'merida': [20.97, -89.62],
  'mérida': [20.97, -89.62],
  'panama': [8.98, -79.52],
  'taipei': [25.03, 121.57],
  'cape dyer': [66.58, -61.62],
  'roberval': [48.52, -72.23],
  'fort chimo': [58.10, -68.42],
  'norman wells': [65.28, -126.80],
  'fort good hope': [66.26, -128.65],
  'hay river': [60.84, -115.79],
  'ponta delgada': [37.74, -25.66],
};

// Notes réutilisées partout (traduites une seule fois).
const LEG = "* Arrivée un jour plus tard (chaque * = +1 jour). [Tech] = escale technique. HL : Heure locale.";
const RET = "[R] Vols retour reconstitués par itinéraire miroir (horaires plausibles).";
const REC = "[R] Reconstitué : numéros de vol et horaires — non issus d'un timetable original.";
NOTES[LEG] = "* Arrival on a later day (each * = +1 day). [Tech] = technical stop. LT: local time.";
NOTES[RET] = "[R] Return flights reconstructed as mirror itineraries (plausible schedules).";
NOTES[REC] = "[R] Reconstructed: flight numbers and schedules — not from an original timetable.";

// Notes « lignes fictives » : la ROUTE est réelle mais était exploitée en
// Douglas DC-4 / DC-6 (pas en L-1049) ; proposée ici car même catégorie d'avion.
const DC4NOTE = "Ligne exploitée historiquement en Douglas DC-4 ; proposée ici pour le Super Constellation L-1049, appareil de même catégorie (quadrimoteur à pistons long-courrier).";
const DC6NOTE = "Ligne exploitée historiquement en Douglas DC-6 ; proposée ici pour le Super Constellation L-1049, appareil de même catégorie (quadrimoteur à pistons long-courrier).";
NOTES[DC4NOTE] = "Route historically operated with the Douglas DC-4; offered here for the Super Constellation L-1049, an aircraft of the same category (long-haul four-engine piston airliner).";
NOTES[DC6NOTE] = "Route historically operated with the Douglas DC-6; offered here for the Super Constellation L-1049, an aircraft of the same category (long-haul four-engine piston airliner).";

// ---------------------------------------------------------------------------
// Moteur d'horloge : calcule les horaires d'un vol le long d'une suite de
// ports, à partir d'une heure de départ, de durées de vol et d'escale (min).
// ---------------------------------------------------------------------------
function toMin(hhmm) {
  const [h, m] = String(hhmm).split(':').map(Number);
  return h * 60 + m;
}
function fmt(t0, add) {
  const tot = t0 + add;
  const day = Math.floor(tot / 1440);
  const mins = ((tot % 1440) + 1440) % 1440;
  const hh = String(Math.floor(mins / 60)).padStart(2, '0');
  const mm = String(mins % 60).padStart(2, '0');
  return { s: hh + ':' + mm, day };
}
function stars(day) { return day > 0 ? ' ' + '*'.repeat(day) : ''; }

// Renvoie { originDep, escales:[cellStr...], destArr }.
function timeline(ports, legMin, groundMin, depHHMM) {
  const t0 = toMin(depHHMM);
  const originDep = depHHMM;           // le départ d'origine est le jour 0
  const escales = [];
  let destArr = '';
  let acc = 0;
  for (let i = 0; i < legMin.length; i++) {
    acc += legMin[i];
    const arr = fmt(t0, acc);
    const isLast = i === legMin.length - 1;
    if (isLast) {
      destArr = arr.s + stars(arr.day);
    } else {
      const g = groundMin[i] != null ? groundMin[i] : 60;
      const dep = fmt(t0, acc + g);
      escales.push(`${ports[i + 1]} (${arr.s + stars(arr.day)} / ${dep.s + stars(dep.day)})`);
      acc += g;
    }
  }
  return { originDep, escales, destArr };
}

// ---------------------------------------------------------------------------
// Construction d'une compagnie. On ajoute des routes ; chaque route crée le vol
// aller ET le vol retour (miroir). On enregistre au passage les traductions.
// ---------------------------------------------------------------------------
function airline(id, name, file) {
  const a = { id, name, file, sheets: [] };
  a.sheet = (frName, enName, notes) => {
    SHEETS[frName] = enName;
    const s = { name: frName, notes: notes || [], flights: [] };
    a.sheets.push(s);
    return {
      // route(spec) : ajoute aller + retour dans cette feuille.
      route(spec) {
        LINES[spec.lineFwd[0]] = spec.lineFwd[1];
        LINES[spec.lineRet[0]] = spec.lineRet[1];
        FREQ[spec.freqFwd[0]] = spec.freqFwd[1];
        FREQ[spec.freqRet[0]] = spec.freqRet[1];

        const P = spec.ports;
        const fwd = timeline(P, spec.legMin, spec.groundMin, spec.outDep);
        s.flights.push({
          no: spec.fwdNo, line: spec.lineFwd[0],
          port0: P[0], originDep: fwd.originDep,
          escales: fwd.escales, portLast: P[P.length - 1], destArr: fwd.destArr,
          freq: spec.freqFwd[0],
        });

        const Pr = P.slice().reverse();
        const rev = timeline(Pr, spec.legMin.slice().reverse(),
          spec.groundMin.slice().reverse(), spec.retDep);
        s.flights.push({
          no: spec.retNo, line: spec.lineRet[0],
          port0: Pr[0], originDep: rev.originDep,
          escales: rev.escales, portLast: Pr[Pr.length - 1], destArr: rev.destArr,
          freq: spec.freqRet[0],
        });
        return this;
      },
    };
  };
  // rawSheet : feuille VERBATIM (aucun recalcul d'horaire). Sert à intégrer
  // fidèlement les données historiques d'origine (fuseaux (PT)/(ET), jours
  // (Mer)/(Ven), marqueurs [**]/[***], en-têtes spécifiques). Les traductions
  // de ces libellés vivent déjà dans translations.js / translations-extra.js.
  a.rawSheet = (tabName, header, flightRows, notes) => {
    const W = header.length;
    const pad = (r) => { const c = r.slice(); while (c.length < W) c.push(''); return c; };
    const aoa = [pad(header)];
    flightRows.forEach((r) => aoa.push(pad(r)));
    aoa.push(new Array(W).fill(''));
    aoa.push(pad(['Notes :']));
    (notes || []).forEach((n) => aoa.push(pad([n])));
    a.sheets.push({ name: tabName, aoa, flightCount: flightRows.length });
  };
  return a;
}

// Aide compacte pour les notes { fr -> en }.
function N(fr, en) { NOTES[fr] = en; return fr; }

const airlines = [];

// ===========================================================================
// ===  0e LOT : 6 compagnies historiques D'ORIGINE (intégrées VERBATIM)  ====
// Anciennement fichiers Excel faits main. Données reprises à l'identique
// (fuseaux, jours, marqueurs, en-têtes). Traductions déjà dans
// translations.js / translations-extra.js. Corrigé : LH 503 (escale Rio à
// durée nulle 22:00/22:00 -> 22:00/22:45).
// ===========================================================================

// --- TWA [S] ---
{
  const a = airline('twa', 'TWA', 'Reseau_TWA_L1049_sourced.xlsx');
  a.rawSheet('Réseau Domestique (USA)',
    ['N° Vol', 'Nom du Vol / Type', 'Origine', 'Départ (Heure Locale)', 'Escale 1 (Arr. / Dép.)', 'Escale 2 (Arr. / Dép.)', 'Escale 3 (Arr. / Dép.)', 'Destination', 'Arrivée (Heure Locale)', 'Fréquence'],
    [
      ['TW 2', 'The Ambassador (Non-Stop)', 'Los Angeles (LAX)', '09:00 (PT)', '-', '-', '-', 'New York (IDL)', '19:45 (ET)', 'Quotidien'],
      ['TW 3', 'The Ambassador (Non-Stop)', 'New York (IDL)', '10:30 (ET)', '-', '-', '-', 'Los Angeles (LAX)', '17:15 (PT)', 'Quotidien'],
      ['TW 6', 'The Sky Chief (Non-Stop)', 'San Francisco (SFO)', '08:30 (PT)', '-', '-', '-', 'New York (IDL)', '19:30 (ET)', 'Quotidien'],
      ['TW 7', 'The Sky Chief (Non-Stop)', 'New York (IDL)', '16:00 (ET)', '-', '-', '-', 'San Francisco (SFO)', '23:00 (PT)', 'Quotidien'],
      ['TW 14', 'Transcontinental Multi-Escales', 'New York (LGA)', '08:00 (ET)', 'Chicago (MDW) (09:20 / 10:00)', 'Kansas City (MKC) (11:40 / 12:20)', '-', 'Los Angeles (LAX)', '14:40 (PT)', 'Quotidien'],
      ['TW 15', 'Transcontinental Multi-Escales', 'Los Angeles (LAX)', '08:15 (PT)', 'Kansas City (MKC) (14:35 / 15:15)', 'Chicago (MDW) (16:55 / 17:35)', '-', 'New York (LGA)', '20:55 (ET)', 'Quotidien'],
      ['TW 30', 'Liaison Constellation', 'New York (IDL)', '09:00 (ET)', 'Pittsburgh (PIT) (10:05 / 10:35)', '-', '-', 'Chicago (MDW)', '11:15 (CT)', 'Quotidien'],
      ['TW 31', 'Liaison Constellation (Retour)', 'Chicago (MDW)', '09:00', 'Pittsburgh (PIT) (09:40 / 10:10)', '-', '-', 'New York (IDL)', '11:15', 'Quotidien'],
      ['TW 45', 'Liaison Régionale', 'Boston (BOS)', '11:00 (ET)', 'New York (LGA) (11:55 / 12:35)', 'Washington (DCA) (13:40 / 14:10)', '-', 'St. Louis (STL)', '15:40 (CT)', 'Quotidien'],
      ['TW 46', 'Liaison Régionale (Retour)', 'St. Louis (STL)', '09:00', 'Washington (DCA) (10:30 / 11:00)', 'New York (LGA) (12:05 / 12:45)', '-', 'Boston (BOS)', '13:40', 'Quotidien'],
      ['TW 62', 'Midwest Express', 'Kansas City (MKC)', '07:30 (CT)', 'Chicago (MDW) (09:10 / 09:50)', '-', '-', 'New York (LGA)', '13:10 (ET)', 'Quotidien'],
      ['TW 63', 'Midwest Express (Retour)', 'New York (LGA)', '09:00', 'Chicago (MDW) (12:20 / 13:00)', '-', '-', 'Kansas City (MKC)', '14:40', 'Quotidien'],
    ],
    [
      "Sources : services et routes d'après TWA Museum (twamuseum.com), Metro Airport News et Wikipedia (Lockheed Constellation).",
      "[S] Sourcé : « The Ambassador » — 1er service transcontinental régulier SANS ESCALE (est-bound) inauguré le 19 octobre 1953 en Lockheed L-1049C, en « un peu moins de 8 heures » ; « The Sky Chief » transcontinental ; hubs TWA de Chicago (Midway) et Kansas City (MKC).",
      "[R] Reconstitué : numéros de vol et horaires à la minute — plausibles pour l'époque mais NON issus d'un timetable original.",
      "* Arrivée le lendemain (+1 jour). Heures locales (PT: Pacifique, MT: Montagnes, CT: Centre, ET: Est).",
      "[R] Vols retour ajoutés automatiquement (itinéraire miroir, horaires reconstitués).",
    ]);
  a.rawSheet('Réseau International',
    ['N° Vol', 'Axe / Destination', 'Origine', 'Départ (Heure Locale)', 'Escale 1 (Arr. / Dép.)', 'Escale 2 (Arr. / Dép.)', 'Escale 3 (Arr. / Dép.)', 'Destination', 'Arrivée (Heure Locale)', 'Jours de Circulation'],
    [
      ['TW 800', 'Europe (Le Parisien)', 'New York (IDL)', '18:30 (ET)', 'Gander [Tech] (23:00 / 00:00 *)', 'Shannon (08:00 / 09:00 *)', 'Paris (ORY) (12:15 / 13:30 *)', 'Rome (CIA)', '16:20 *', 'Lundi, Mercredi, Vendredi'],
      ['TW 801', 'Europe (Le Parisien)', 'Rome (CIA)', '10:00', 'Paris (ORY) (13:15 / 14:30)', 'Shannon (17:45 / 18:45)', 'Gander [Tech] (01:15 / 02:15 *)', 'New York (IDL)', '05:30 *', 'Mardi, Jeudi, Samedi'],
      ['TW 810', 'Ligne Ibérique', 'New York (IDL)', '17:00 (ET)', 'Santa Maria (02:30 / 03:30 *)', 'Lisbonne (07:00 / 08:00 *)', 'Madrid (10:45 / 12:00 *)', 'Rome (CIA)', '15:40 *', 'Mardi, Samedi'],
      ['TW 820', 'Ligne Allemande', 'New York (IDL)', '19:30 (ET)', 'Gander [Tech] (00:00 / 01:00 *)', 'Shannon (09:15 / 10:15 *)', 'Londres (LHR) (11:45 / 13:00 *)', 'Francfort (FRA)', '15:25 *', 'Mercredi, Dimanche'],
      ['TW 821', 'Ligne Allemande (Retour)', 'Francfort (FRA)', '16:00', 'Londres (LHR) (18:25 / 19:40)', 'Shannon (21:10 / 22:10)', 'Gander [Tech] (06:25 * / 07:25 *)', 'New York (IDL)', '11:55 *', 'Mercredi, Dimanche'],
      ['TW 850', 'Route des Étoiles (Orient)', 'Rome (CIA)', '08:30', 'Le Caire (CAI) (13:00 / 14:30)', '-', '-', 'Bombay (BOM)', '23:15 *', 'Vendredi'],
    ],
    [
      "[S] Sourcé : passage au Lockheed L-1049G sur l'Atlantique à partir de 1956 (auparavant : L-749 Constellation) ; escales techniques de Gander (Terre-Neuve) et Shannon (Irlande), typiques des vols piston ; destinations TWA d'époque (Paris, Rome, Madrid, Lisbonne, Francfort, Londres, Le Caire).",
      "[R] Reconstitué : numéros de vol, horaires et jours de circulation ; l'extension Le Caire–Bombay relève d'un service d'interchange et reste approximative.",
      "* Arrivée le lendemain (+1 jour). [Tech] = escale technique de ravitaillement.",
      "Sources : Wikipedia (Lockheed L-1049 Super Constellation), TWA Museum, key.aero (famille Constellation).",
      "[R] Vols retour ajoutés automatiquement (itinéraire miroir, horaires reconstitués).",
    ]);
  airlines.push(a);
}

// --- Air France [S] --- (réseau Super Constellation entièrement refondu d'après
// l'indicateur officiel Air France, été 1956 : « Parisien Spécial » (Paris–New
// York), prolongements Mexico (via NY et via Boston), Amérique du Sud (AF 093),
// « Eastern Epicurean » (AF 170, Paris–Tokyo) et « L'Épicurien » (AF 568,
// Paris–Londres). Vols retour reconstitués [R] par itinéraire miroir.)
{
  const a = airline('air-france', 'Air France', 'Reseau_Air_France_L1049_sourced.xlsx');
  a.rawSheet('Atlantique & Amériques',
    ['N° Vol', 'Axe / Ligne', 'Origine', 'Départ (HL)', 'Escale 1 (Arr. / Dép.)', 'Escale 2 (Arr. / Dép.)', 'Escale 3 (Arr. / Dép.)', 'Escale 4 (Arr. / Dép.)', 'Escale 5 (Arr. / Dép.)', 'Destination', 'Arrivée (HL)', 'Fréquence / Note'],
    [
      ['AF 045', 'Atlantique Nord (Parisien Spécial)', 'Paris (ORY)', '23:00', '-', '-', '-', '-', '-', 'New York (IDL)', '08:40 *', 'Quotidien'],
      ['AF 046', 'Atlantique Nord (Parisien Spécial) (Retour)', 'New York (IDL)', '19:00', '-', '-', '-', '-', '-', 'Paris (ORY)', '08:00 *', 'Quotidien'],
      ['AF 079', 'Atlantique Nord (Mexico via Boston)', 'Paris (ORY)', '22:00', 'Boston (BOS) (08:40 * / 09:20 *)', 'New York (IDL) (10:30 * / 12:00 *)', '-', '-', '-', 'Mexico (MEX)', '17:55 *', 'Mardi, Jeudi, Samedi'],
      ['AF 080', 'Atlantique Nord (Mexico via Boston) (Retour)', 'Mexico (MEX)', '08:00', 'New York (IDL) (13:35 / 15:00)', 'Boston (BOS) (16:00 / 17:00)', '-', '-', '-', 'Paris (ORY)', '06:30 *', 'Mercredi, Vendredi, Dimanche'],
      ['AF 071', 'Atlantique Nord (Mexico)', 'Paris (ORY)', '22:00', 'New York (IDL) (09:20 * / 12:00 *)', '-', '-', '-', '-', 'Mexico (MEX)', '17:55 *', 'Quotidien'],
      ['AF 072', 'Atlantique Nord (Mexico) (Retour)', 'Mexico (MEX)', '09:00', 'New York (IDL) (14:35 / 16:30)', '-', '-', '-', '-', 'Paris (ORY)', '06:00 *', 'Quotidien'],
      ['AF 093', 'Amérique du Sud', 'Paris (ORY)', '13:00', 'Madrid (15:45 / 16:45)', 'Dakar (23:20 / 00:35 *)', 'Rio de Janeiro (GIG) (09:00 * / 10:15 *)', 'São Paulo (CGH) (11:40 * / 12:25 *)', 'Montevideo (16:25 * / 16:55 *)', 'Buenos Aires (EZE)', '17:50 *', 'Jeudi, Dimanche'],
      ['AF 094', 'Amérique du Sud (Retour)', 'Buenos Aires (EZE)', '09:00', 'Montevideo (09:55 / 10:25)', 'São Paulo (CGH) (14:00 / 14:45)', 'Rio de Janeiro (GIG) (16:00 / 17:00)', 'Dakar (01:30 * / 02:45 *)', 'Madrid (11:30 * / 12:30 *)', 'Paris (ORY)', '14:45 *', 'Lundi, Vendredi'],
    ],
    [
      "Sources : Air France reçoit son 1er L-1049C (F-BGNJ) le 2 nov. 1953, mis en ligne le 15 nov. 1953 ; 24 « Super Connies » (10× 1049C + 14× 1049G), 2e opérateur mondial après TWA.",
      "[S] Sourcé : réseau et horaires d'après l'indicateur officiel Air France (été 1956) — « Parisien Spécial » (Paris–New York), prolongements vers Mexico (via New York ou via Boston) et Amérique du Sud (AF 093) Paris–Madrid–Dakar–Rio–São Paulo–Montevideo–Buenos Aires.",
      "[R] Vols retour reconstitués (itinéraire miroir) : horaires ALLER conformes à l'indicateur Air France 1956, horaires RETOUR plausibles.",
      "* Arrivée le lendemain (+1 jour). HL : Heure Locale. [Tech] = escale technique.",
    ]);
  a.rawSheet('Extrême-Orient',
    ['N° Vol', 'Axe / Ligne', 'Origine', 'Départ (HL)', 'Escale 1 (Arr. / Dép.)', 'Escale 2 (Arr. / Dép.)', 'Escale 3 (Arr. / Dép.)', 'Escale 4 (Arr. / Dép.)', 'Escale 5 (Arr. / Dép.)', 'Escale 6 (Arr. / Dép.)', 'Destination', 'Arrivée (HL)', 'Fréquence / Note'],
    [
      ['AF 170', 'Extrême-Orient (Eastern Epicurean)', 'Paris (ORY)', '17:40', 'Rome (20:30 / 21:50)', 'Téhéran (08:30 * / 10:00 *)', 'Karachi (16:00 * / 17:30 *)', 'Bangkok (03:50 ** / 04:50 **)', 'Saïgon (06:40 ** / 08:10 **)', 'Manille (13:10 ** / 14:10 **)', 'Tokyo (HND)', '22:15 **', 'Mardi, Samedi'],
      ['AF 171', 'Extrême-Orient (Eastern Epicurean) (Retour)', 'Tokyo (HND)', '16:00', 'Manille (20:00 / 21:00)', 'Saïgon (01:15 * / 02:15 *)', 'Bangkok (04:00 * / 05:00 *)', 'Karachi (12:00 * / 13:00 *)', 'Téhéran (18:00 * / 19:00 *)', 'Rome (01:30 ** / 02:30 **)', 'Paris (ORY)', '05:30 **', 'Jeudi, Dimanche'],
    ],
    [
      "[S] Sourcé : le service « Eastern Epicurean » (AF 170, Super Constellation) reliait Paris–Rome–Téhéran–Karachi–Bangkok–Saïgon–Manille–Tokyo, d'après l'indicateur officiel Air France (été 1956).",
      "[R] Vol retour reconstitué (itinéraire miroir) ; horaires aller conformes à l'indicateur, retour plausible.",
      "* / ** : arrivée à J+1 / J+2. HL : Heure Locale.",
    ]);
  a.rawSheet('Europe',
    ['N° Vol', 'Axe / Ligne', 'Origine', 'Départ (HL)', 'Destination', 'Arrivée (HL)', 'Fréquence / Note'],
    [
      ['AF 568', 'Europe (Épicurien)', 'Paris (ORY)', '12:00', 'Londres (LHR)', '13:15', 'Quotidien'],
      ['AF 569', 'Europe (Épicurien) (Retour)', 'Londres (LHR)', '14:00', 'Paris (ORY)', '15:15', 'Quotidien'],
    ],
    [
      "[S] Sourcé : le service « L'Épicurien » (AF 568, Super Constellation) reliait Paris (Orly, 12:00) à Londres (13:15), d'après l'indicateur officiel Air France (été 1956) ; vol court de prestige.",
      "[R] Vol retour reconstitué (Londres–Paris) ; horaire aller conforme à l'indicateur Air France 1956.",
      "HL : Heure Locale.",
    ]);
  airlines.push(a);
}

// --- Lufthansa [S] --- (correction LH 503 : Rio 22:00/22:00 -> 22:00/22:45)
{
  const a = airline('lufthansa', 'Lufthansa', 'Reseau_Lufthansa_L1049G_sourced.xlsx');
  a.rawSheet('Atlantique Nord',
    ['N° Vol', 'Ligne / Axe', 'Origine', 'Départ (HL)', 'Escale 1 (Arr. / Dép.)', 'Escale 2 (Arr. / Dép.)', 'Escale 3 (Arr. / Dép.)', 'Destination', 'Arrivée (HL)', 'Jours / Notes'],
    [
      ['LH 410', 'New York (Ligne inaugurale)', 'Düsseldorf (DUS)', '18:00', 'Shannon (23:30 / 00:30 *)', 'Gander [Tech] (07:45 / 09:00 *)', '-', 'New York (IDL)', '12:15 *', 'Lu, Ve / Vol inaugural D-ALEM (1955)'],
      ['LH 411', 'New York (Retour)', 'New York (IDL)', '16:00', 'Gander [Tech] (19:15 / 20:30)', 'Shannon (03:45 * / 04:45 *)', '-', 'Düsseldorf (DUS)', '10:15 *', 'Lu, Ve'],
      ['LH 402', 'New York (Express)', 'Hambourg (HAM)', '19:00', 'Francfort (20:15 / 21:30)', 'Shannon (01:15 / 02:15 *)', 'Gander [Tech] (08:30 / 09:45 *)', 'New York (IDL)', '13:30 *', 'Ma, Je, Sa'],
      ['LH 403', 'New York (Retour)', 'New York (IDL)', '16:30', 'Gander [Tech] (21:45 / 22:45)', 'Shannon (08:00 / 09:00 *)', 'Francfort (12:45 / 13:45 *)', 'Hambourg (HAM)', '15:00 *', 'Me, Ve, Di'],
    ],
    [
      "Sources : Lufthansa reçoit ses 1ers L-1049G en 1955 (dont D-ALEM) et ouvre son service transatlantique Düsseldorf–Shannon–New York ; 8 L-1049G reconstruisent le réseau long-courrier.",
      "[S] Sourcé : première ligne transatlantique de la nouvelle Lufthansa en 1955 (Düsseldorf/Hambourg/Francfort → Shannon/Gander → New York).",
      "[R] Reconstitué : numéros de vol et horaires exacts.",
      "* Arrivée le lendemain (+1 jour). HL : Heure Locale. [Tech] = escale technique.",
      "[R] Vols retour ajoutés automatiquement (itinéraire miroir, horaires reconstitués).",
    ]);
  a.rawSheet('Sud & Moyen-Orient',
    ['N° Vol', 'Ligne / Axe', 'Origine', 'Départ (HL)', 'Escale 1 (Arr. / Dép.)', 'Escale 2 (Arr. / Dép.)', 'Escale 3 (Arr. / Dép.)', 'Escale 4 (Arr. / Dép.)', 'Destination', 'Arrivée (HL)', 'Notes'],
    [
      ['LH 502', 'Atlantique Sud', 'Hambourg (HAM)', '11:15', 'Francfort (12:30 / 13:45)', 'Paris (ORY) (15:00 / 16:15)', 'Dakar (DKR) (00:30 / 02:00 *)', 'Rio de Janeiro (GIG) (09:45 *)', 'Buenos Aires (EZE)', '15:45 *', 'Inauguré en 1956'],
      ['LH 503', 'Atlantique Sud (Retour)', 'Buenos Aires (EZE)', '16:00', 'Rio de Janeiro (GIG) (22:00 / 22:45)', 'Dakar (DKR) (05:45 * / 07:15 *)', 'Paris (ORY) (15:30 * / 16:45 *)', 'Francfort (18:00 * / 19:15 *)', 'Hambourg (HAM)', '20:30 *', 'Inauguré en 1956'],
      ['LH 622', 'Moyen-Orient', 'Hambourg (HAM)', '09:30', 'Francfort (10:45 / 11:45)', 'Munich (MUC) (12:45 / 13:45)', 'Istanbul (IST) (19:30 / 20:30)', 'Beyrouth (BEY) (23:15 / 00:15 *)', 'Téhéran (THR)', '05:30 *', 'Axe pétrolier (ouvert fin 1956)'],
      ['LH 623', 'Moyen-Orient (Retour)', 'Téhéran (THR)', '16:00', 'Beyrouth (BEY) (21:15 / 22:15)', 'Istanbul (IST) (01:00 * / 02:00 *)', 'Munich (MUC) (07:45 * / 08:45 *)', 'Francfort (09:45 * / 10:45 *)', 'Hambourg (HAM)', '12:00 *', 'Axe pétrolier (ouvert fin 1956)'],
    ],
    [
      "[S] Sourcé : extension du réseau L-1049G vers l'Atlantique Sud (Amérique du Sud) et le Moyen-Orient à partir de 1956.",
      "[R] Reconstitué : numéros de vol et horaires exacts.",
      "* Arrivée le lendemain (+1 jour). HL : Heure Locale.",
      "[R] Vols retour ajoutés automatiquement (itinéraire miroir, horaires reconstitués).",
    ]);
  airlines.push(a);
}

// --- Eastern Air Lines [S] ---
{
  const a = airline('eastern-air-lines', 'Eastern Air Lines', 'Reseau_Eastern_Air_Lines_L1049_sourced.xlsx');
  a.rawSheet('Axe Nord-Sud (Mainlines)',
    ['N° Vol', 'Nom du Vol / Type', 'Origine', 'Départ (HL)', 'Escale 1 (Arr. / Dép.)', 'Escale 2 (Arr. / Dép.)', 'Escale 3 (Arr. / Dép.)', 'Destination', 'Arrivée (HL)', 'Fréquence / Notes'],
    [
      ['EA 601', 'The Golden Falcon (Non-Stop)', 'New York (IDL)', '09:00', '-', '-', '-', 'Miami (MIA)', '13:15', 'Quotidien'],
      ['EA 602', 'The Golden Falcon (Retour)', 'Miami (MIA)', '15:00', '-', '-', '-', 'New York (IDL)', '19:15', 'Quotidien'],
      ['EA 611', 'The Silver Falcon', 'Boston (BOS)', '08:00', 'New York (LGA) (09:00 / 09:40)', 'Washington (DCA) (10:45 / 11:15)', '-', 'Miami (MIA)', '15:30', 'Quotidien'],
      ['EA 612', 'The Silver Falcon (Retour)', 'Miami (MIA)', '11:00', 'Washington (DCA) (15:15 / 15:45)', 'New York (LGA) (16:50 / 17:30)', '-', 'Boston (BOS)', '18:30', 'Quotidien'],
      ['EA 620', 'Ligne du Golfe', 'New York (EWR)', '08:30', 'Atlanta (ATL) (11:15 / 11:55)', 'New Orleans (MSY) (13:15 / 13:55)', '-', 'Houston (HOU)', '15:40', 'Quotidien'],
      ['EA 621', 'Ligne du Golfe (Retour)', 'Houston (HOU)', '11:00', 'New Orleans (MSY) (12:45 / 13:25)', 'Atlanta (ATL) (14:45 / 15:25)', '-', 'New York (EWR)', '18:10', 'Quotidien'],
      ['EA 640', 'Chicago – Floride', 'Chicago (MDW)', '09:00', 'Atlanta (ATL) (12:00 / 12:40)', '-', '-', 'Miami (MIA)', '15:45', 'Quotidien'],
      ['EA 641', 'Chicago – Floride (Retour)', 'Miami (MIA)', '11:00', 'Atlanta (ATL) (14:05 / 14:45)', '-', '-', 'Chicago (MDW)', '17:45', 'Quotidien'],
    ],
    [
      "Sources : Eastern Air Lines, client de lancement du L-1049 (1er exemplaire de série, déc. 1951) ; ~10 L-1049C sur les axes domestiques à forte densité, notamment Miami–New York. Flotte surnommée « Great Silver Fleet ».",
      "[S] Sourcé : Eastern exploite le L-1049 sur l'axe Nord-Sud (New York–Miami) et le réseau domestique de l'Est/Golfe ; service « Golden Falcon ».",
      "[R] Reconstitué : numéros de vol et horaires ; le nom « Silver Falcon » est une reconstitution.",
      "HL : Heure Locale. Fuseaux ET/CT selon les escales.",
      "[R] Vols retour ajoutés automatiquement (itinéraire miroir, horaires reconstitués).",
    ]);
  a.rawSheet('Navettes & Floride',
    ['N° Vol', 'Ligne / Axe', 'Origine', 'Départ (HL)', 'Escale 1 (Arr. / Dép.)', 'Escale 2 (Arr. / Dép.)', 'Destination', 'Arrivée (HL)', 'Notes'],
    [
      ['EA 312', 'Ligne de Cuba', 'Miami (MIA)', '11:30', '-', '-', 'La Havane (HAV)', '12:30', 'Vol international court'],
      ['EA 317', 'Ligne de Cuba (Retour)', 'La Havane (HAV)', '14:00', '-', '-', 'Miami (MIA)', '15:00', 'Forte demande touristique'],
      ['EA 205', 'Navette de Floride', 'New York (LGA)', '07:30', 'Washington (DCA) (08:30 / 09:00)', 'Jacksonville (JAX) (11:45 / 12:15)', 'Miami (MIA)', '14:00', 'Quotidien'],
    ],
    [
      "[S] Sourcé : la liaison Miami–La Havane figurait parmi les rares routes internationales régulières d'Eastern à l'époque.",
      "[R] Reconstitué : numéros de vol et horaires.",
      "HL : Heure Locale.",
    ]);
  airlines.push(a);
}

// --- Flying Tiger Line [S] (cargo / charters) ---
{
  const a = airline('flying-tiger', 'Flying Tiger', 'Reseau_Flying_Tiger_L1049H_sourced.xlsx');
  a.rawSheet('Lignes Cargo & Charters',
    ['N° Vol / Contrat', 'Type de Ligne', 'Origine', 'Départ (HL)', 'Escale 1 (Arr. / Dép.)', 'Escale 2 (Arr. / Dép.)', 'Escale 3 (Arr. / Dép.)', 'Escale 4 (Arr. / Dép.)', 'Destination', 'Arrivée (HL)', 'Fréquence / Notes'],
    [
      ['FT 100', 'Cargo Domestique Est-Ouest', 'New York (IDL)', '21:00', 'Chicago (MDW) (23:30 / 00:45 *)', 'Kansas City (MKC) (02:15 / 03:15 *)', 'Los Angeles (LAX) (08:00 / 09:15 *)', '-', 'San Francisco (SFO)', '10:45 *', 'Quotidien (fret)'],
      ['FT 101', 'Cargo Domestique Ouest-Est', 'San Francisco (SFO)', '22:00', 'Los Angeles (23:00 / 00:15 *)', 'Chicago (MDW) (06:45 / 08:00 *)', '-', '-', 'New York (IDL)', '12:30 *', 'Quotidien (fret)'],
      ['FT 739', 'Charte MATS (Transpacifique)', 'Travis AFB (SUU)', '18:00', 'Honolulu (HNL) (01:30 / 03:00 *)', 'Wake Island (AWK) (10:15 / 11:45 [**])', 'Guam (GUM) (15:30 / 16:45 [**])', 'Manille (MNL) (21:00 [**])', 'Saïgon (SGN)', '04:30 [***]', 'Charte militaire U.S. Army'],
      ['FT 200', 'Charte Transatlantique (Émigrants)', 'New York (IDL)', '19:30', 'Gander [Tech] (01:00 / 02:15 *)', 'Shannon [Tech] (11:45 / 13:00 *)', '-', '-', 'Francfort (FRA)', '16:45 *', 'Charters saisonniers'],
    ],
    [
      "Sources : Flying Tiger Line reçoit dix L-1049H (convertibles cargo/passagers) en 1957 ; exploités sur le fret nord-américain et pour le Military Air Transport Service (MATS). Le vol 739 fut une charte MATS Travis AFB → Saïgon.",
      "[S] Sourcé : L-1049H convertible ; réseau fret domestique + chartes militaires transpacifiques (MATS).",
      "[R] Reconstitué : numéros/contrats et horaires exacts.",
      "* Arrivée le lendemain (+1 jour). [**] Passage de la ligne de changement de date. [***] Arrivée à J+2. [Tech] = escale technique.",
    ]);
  airlines.push(a);
}

// --- Irish Airlines / Aerlínte Éireann [S] ---
{
  const a = airline('irish-airlines', 'Irish Airlines', 'Reseau_Irish_Airlines_L1049_sourced.xlsx');
  a.rawSheet('Liaisons Transatlantiques',
    ['N° Vol', 'Compagnies', 'Type de Vol', 'Origine', 'Départ (HL)', 'Escale 1 (Arr. / Dép.)', 'Escale 2 (Arr. / Dép.)', 'Destination', 'Arrivée (HL)', 'Fréquence'],
    [
      ['IN 101', 'Aerlínte Éireann / Seaboard & Western', 'Sens Europe ➔ USA', 'Dublin (DUB)', '10:30', 'Shannon (11:15 / 12:30)', 'Gander [Tech] (17:00 / 18:00)', 'New York (IDL)', '20:45', '3x/sem. (vol inaugural 28 avr. 1958)'],
      ['IN 102', 'Aerlínte Éireann / Seaboard & Western', 'Sens USA ➔ Europe', 'New York (IDL)', '16:00', 'Gander [Tech] (21:15 / 22:15)', 'Shannon (07:30 / 08:30 *)', 'Dublin (DUB)', '09:45 *', '3x par semaine'],
      ['IN 105', 'Aerlínte Éireann / Seaboard & Western', 'Sens Europe ➔ USA', 'Shannon (SNN)', '12:00', 'Gander [Tech] (17:30 / 18:30)', '-', 'Boston (BOS)', '20:15', '2x par semaine'],
      ['IN 106', 'Aerlínte Éireann / Seaboard & Western', 'Sens USA ➔ Europe', 'Boston (BOS)', '11:00', 'Gander [Tech] (12:45 / 13:45)', '-', 'Shannon (SNN)', '19:15', '2x par semaine'],
    ],
    [
      "Sources : Aerlínte Éireann (filiale transatlantique d'Aer Lingus) inaugure Dublin–Shannon–New York le 28 avril 1958 (N1009C), en L-1049 Super Constellation loués à Seaboard & Western (avions et équipages techniques), pour un peu plus de deux ans.",
      "[S] Sourcé : ligne, appareil, partenaire (Seaboard & Western), escale technique de Gander, date inaugurale.",
      "[R] Reconstitué : numéros de vol et horaires ; l'extension vers Boston est reconstituée.",
      "* Arrivée le lendemain (+1 jour). [Tech] = escale technique. HL : Heure Locale.",
      "[R] Vols retour ajoutés automatiquement (itinéraire miroir, horaires reconstitués).",
    ]);
  airlines.push(a);
}

// ===========================================================================
//  1. KLM  [S]
// ===========================================================================
{
  const a = airline('klm', 'KLM', 'Reseau_KLM_L1049_sourced.xlsx');
  const src = N("Sources : KLM figure parmi les opérateurs de série du L-1049C ; réseau Atlantique Nord (Amsterdam–New York) et « Ligne des Indes » vers Djakarta via le Moyen-Orient.",
    "Sources: KLM was among the production operators of the L-1049C; North Atlantic network (Amsterdam–New York) and 'East Indies Line' to Djakarta via the Middle East.");
  const s1 = N("[S] Sourcé : KLM exploite le L-1049 sur l'Atlantique Nord et vers les Indes néerlandaises ; escales techniques Prestwick/Gander.",
    "[S] Sourced: KLM operated the L-1049 on the North Atlantic and to the Dutch East Indies; Prestwick/Gander technical stops.");
  a.sheet('Atlantique Nord', 'North Atlantic', [src, s1, REC, LEG, RET]).route({
    fwdNo: 'KL 641', retNo: 'KL 642',
    lineFwd: ['Atlantique Nord (Le Néerlandais)', 'North Atlantic (The Dutchman)'],
    lineRet: ['Atlantique Nord (Retour)', 'North Atlantic (Return)'],
    ports: ['Amsterdam (AMS)', 'Prestwick', 'Gander [Tech]', 'New York (IDL)'],
    legMin: [120, 480, 300], groundMin: [60, 75], outDep: '10:00', retDep: '16:00',
    freqFwd: ['Mardi, Jeudi, Samedi', 'Tuesday, Thursday, Saturday'],
    freqRet: ['Mercredi, Vendredi, Dimanche', 'Wednesday, Friday, Sunday'],
  });
  a.sheet('Extrême-Orient', 'Far East', [LEG, RET]).route({
    fwdNo: 'KL 823', retNo: 'KL 824',
    lineFwd: ['Ligne des Indes (Le Batavia)', 'East Indies Line (The Batavia)'],
    lineRet: ['Ligne des Indes (Retour)', 'East Indies Line (Return)'],
    ports: ['Amsterdam (AMS)', 'Rome', 'Beyrouth', 'Karachi', 'Calcutta', 'Bangkok', 'Djakarta (Batavia)'],
    legMin: [180, 210, 300, 240, 210, 300], groundMin: [75, 60, 75, 60, 75],
    outDep: '13:00', retDep: '09:00',
    freqFwd: ['Lundi, Vendredi', 'Monday, Friday'],
    freqRet: ['Mardi, Samedi', 'Tuesday, Saturday'],
  });
  airlines.push(a);
}

// ===========================================================================
//  2. SABENA  [S] (L-1049 loués à Seaboard & Western, Expo 58)
// ===========================================================================
{
  const a = airline('sabena', 'Sabena', 'Reseau_Sabena_L1049_sourced.xlsx');
  const src = N("Sources : la Sabena a exploité 2 à 3 Super Constellation L-1049 loués à Seaboard & Western, principalement sur l'Atlantique Nord lors de l'Expo 58 de Bruxelles.",
    "Sources: Sabena operated 2–3 L-1049 Super Constellations leased from Seaboard & Western, mainly on the North Atlantic during the 1958 Brussels World's Fair.");
  const s1 = N("[S] Sourcé : appareil (L-1049 loué), période (1958) et axe transatlantique Bruxelles–New York réels ; ligne du Congo belge conforme au réseau Sabena.",
    "[S] Sourced: aircraft (leased L-1049), period (1958) and Brussels–New York transatlantic axis are real; the Belgian Congo line matches Sabena's network.");
  const sh = a.sheet('Réseau International', 'International Network', [src, s1, REC, LEG, RET]);
  sh.route({
    fwdNo: 'SN 501', retNo: 'SN 502',
    lineFwd: ['Atlantique Nord', 'North Atlantic'],
    lineRet: ['Atlantique Nord (Retour)', 'North Atlantic (Return)'],
    ports: ['Bruxelles (BRU)', 'Shannon', 'Gander [Tech]', 'New York (IDL)'],
    legMin: [150, 450, 300], groundMin: [60, 75], outDep: '11:00', retDep: '17:00',
    freqFwd: ['Mercredi, Dimanche', 'Wednesday, Sunday'],
    freqRet: ['Lundi, Jeudi', 'Monday, Thursday'],
  });
  sh.route({
    fwdNo: 'SN 551', retNo: 'SN 552',
    lineFwd: ['Ligne du Congo', 'Congo Line'],
    lineRet: ['Ligne du Congo (Retour)', 'Congo Line (Return)'],
    ports: ['Bruxelles (BRU)', 'Rome', 'Le Caire (CAI)', 'Kano [Tech]', 'Léopoldville (LEO)'],
    legMin: [180, 210, 300, 240], groundMin: [75, 60, 60], outDep: '22:00', retDep: '20:00',
    freqFwd: ['Mardi, Samedi', 'Tuesday, Saturday'],
    freqRet: ['Jeudi, Dimanche', 'Thursday, Sunday'],
  });
  airlines.push(a);
}

// ===========================================================================
//  3. QANTAS  [S]
// ===========================================================================
{
  const a = airline('qantas', 'Qantas', 'Reseau_Qantas_L1049_sourced.xlsx');
  const src = N("Sources : Qantas exploite le L-1049 Super Constellation dès 1954 sur la « Kangaroo Route » (Sydney–Londres) et la « Southern Cross Route » transpacifique.",
    "Sources: Qantas operated the L-1049 Super Constellation from 1954 on the 'Kangaroo Route' (Sydney–London) and the transpacific 'Southern Cross Route'.");
  const s1 = N("[S] Sourcé : appareil et deux grands corridors Qantas (Kangaroo via l'Asie, Southern Cross via le Pacifique) ; escales d'époque.",
    "[S] Sourced: aircraft and Qantas's two great corridors (Kangaroo via Asia, Southern Cross via the Pacific); period stops.");
  a.sheet('Kangaroo Route', 'Kangaroo Route', [src, s1, REC, LEG, RET]).route({
    fwdNo: 'QF 1', retNo: 'QF 2',
    lineFwd: ['Kangaroo Route (Londres)', 'Kangaroo Route (London)'],
    lineRet: ['Kangaroo Route (Retour)', 'Kangaroo Route (Return)'],
    ports: ['Sydney (SYD)', 'Darwin', 'Singapour', 'Calcutta', 'Karachi', 'Le Caire (CAI)', 'Rome', 'Londres (LHR)'],
    legMin: [240, 300, 360, 180, 300, 240, 180], groundMin: [60, 75, 60, 75, 60, 75],
    outDep: '07:00', retDep: '10:00',
    freqFwd: ['Mardi, Vendredi', 'Tuesday, Friday'],
    freqRet: ['Jeudi, Dimanche', 'Thursday, Sunday'],
  });
  a.sheet('Southern Cross', 'Southern Cross', [LEG, RET]).route({
    fwdNo: 'QF 543', retNo: 'QF 544',
    lineFwd: ['Southern Cross Route (Amérique)', 'Southern Cross Route (Americas)'],
    lineRet: ['Southern Cross Route (Retour)', 'Southern Cross Route (Return)'],
    ports: ['Sydney (SYD)', 'Nandi', 'Canton Island [Tech]', 'Honolulu', 'San Francisco (SFO)'],
    legMin: [240, 300, 360, 300], groundMin: [75, 60, 75], outDep: '16:00', retDep: '12:00',
    freqFwd: ['Mercredi, Samedi', 'Wednesday, Saturday'],
    freqRet: ['Lundi, Jeudi', 'Monday, Thursday'],
  });
  airlines.push(a);
}

// ===========================================================================
//  4. TRANS-CANADA AIR LINES (TCA)  [S]
// ===========================================================================
{
  const a = airline('trans-canada', 'Trans-Canada Air Lines', 'Reseau_Trans_Canada_L1049_sourced.xlsx');
  const src = N("Sources : Trans-Canada Air Lines (TCA) exploite le L-1049 dès 1954 sur l'Atlantique Nord (Montréal–Londres) et le transcontinental canadien.",
    "Sources: Trans-Canada Air Lines (TCA) operated the L-1049 from 1954 on the North Atlantic (Montreal–London) and the Canadian transcontinental.");
  const s1 = N("[S] Sourcé : appareil et axes TCA réels (transatlantique via Gander/Shannon ; transcontinental Montréal–Vancouver).",
    "[S] Sourced: aircraft and real TCA axes (transatlantic via Gander/Shannon; transcontinental Montreal–Vancouver).");
  a.sheet('Atlantique Nord', 'North Atlantic', [src, s1, REC, LEG, RET]).route({
    fwdNo: 'TC 862', retNo: 'TC 863',
    lineFwd: ['Atlantique Nord (Montréal)', 'North Atlantic (Montreal)'],
    lineRet: ['Atlantique Nord (Retour)', 'North Atlantic (Return)'],
    ports: ['Montréal (YUL)', 'Gander [Tech]', 'Shannon', 'Londres (LHR)'],
    legMin: [150, 450, 120], groundMin: [60, 60], outDep: '20:00', retDep: '12:00',
    freqFwd: ['Lundi, Mercredi, Vendredi', 'Monday, Wednesday, Friday'],
    freqRet: ['Mardi, Jeudi, Samedi', 'Tuesday, Thursday, Saturday'],
  });
  a.sheet('Transcontinental', 'Transcontinental', [LEG, RET]).route({
    fwdNo: 'TC 3', retNo: 'TC 4',
    lineFwd: ['Transcontinental (Le Canadien)', 'Transcontinental (The Canadian)'],
    lineRet: ['Transcontinental (Retour)', 'Transcontinental (Return)'],
    ports: ['Montréal (YUL)', 'Toronto', 'Winnipeg', 'Vancouver (YVR)'],
    legMin: [75, 180, 240], groundMin: [45, 60], outDep: '08:00', retDep: '09:00',
    freqFwd: ['Quotidien', 'Daily'],
    freqRet: ['Quotidien', 'Daily'],
  });
  airlines.push(a);
}

// ===========================================================================
//  7. LÍNEA AEROPOSTAL VENEZOLANA (LAV)  [S]
// ===========================================================================
{
  const a = airline('aeropostal', 'Línea Aeropostal Venezolana', 'Reseau_Aeropostal_Venezolana_L1049_sourced.xlsx');
  const src = N("Sources : la Línea Aeropostal Venezolana (LAV) exploite le L-1049 Super Constellation ; l'un d'eux (YV-C-AMS « Rafael Urdaneta ») est perdu en mer au large d'Asbury Park le 20 juin 1956 (vol 253).",
    "Sources: Línea Aeropostal Venezolana (LAV) operated the L-1049 Super Constellation; one of them (YV-C-AMS 'Rafael Urdaneta') was lost at sea off Asbury Park on 20 June 1956 (Flight 253).");
  const s1 = N("[S] Sourcé : appareil et axe Caracas–New York réels ; extension européenne conforme au réseau LAV.",
    "[S] Sourced: aircraft and Caracas–New York axis are real; the European extension matches the LAV network.");
  a.sheet('Caraïbes & Amérique du Nord', 'Caribbean & North America', [src, s1, REC, LEG, RET]).route({
    fwdNo: 'LV 253', retNo: 'LV 254',
    lineFwd: ['Ligne de New York', 'New York Line'],
    lineRet: ['Ligne de New York (Retour)', 'New York Line (Return)'],
    ports: ['Caracas (CCS)', 'Maracaibo', 'Ciudad Trujillo', 'New York (IDL)'],
    legMin: [90, 180, 300], groundMin: [45, 60], outDep: '08:00', retDep: '15:00',
    freqFwd: ['Mardi, Jeudi, Samedi', 'Tuesday, Thursday, Saturday'],
    freqRet: ['Mercredi, Vendredi, Dimanche', 'Wednesday, Friday, Sunday'],
  });
  a.sheet("Ligne d'Europe", 'Europe Line', [LEG, RET]).route({
    fwdNo: 'LV 601', retNo: 'LV 602',
    lineFwd: ["Ligne d'Europe", 'Europe Line'],
    lineRet: ["Ligne d'Europe (Retour)", 'Europe Line (Return)'],
    ports: ['Caracas (CCS)', 'Port of Spain', 'Santa Maria', 'Lisbonne', 'Madrid'],
    legMin: [90, 600, 120, 90], groundMin: [45, 90, 60], outDep: '18:00', retDep: '12:00',
    freqFwd: ['Vendredi', 'Friday'],
    freqRet: ['Dimanche', 'Sunday'],
  });
  airlines.push(a);
}

// ===========================================================================
//  8. AIR INDIA  [S]
// ===========================================================================
{
  const a = airline('air-india', 'Air India', 'Reseau_Air_India_L1049_sourced.xlsx');
  const src = N("Sources : Air India International exploite le L-1049 Super Constellation dès 1954 (Bombay–Londres, Bombay–Extrême-Orient).",
    "Sources: Air India International operated the L-1049 Super Constellation from 1954 (Bombay–London, Bombay–Far East).");
  const s1 = N("[S] Sourcé : appareil et corridors Air India réels (Europe via Le Caire/Rome ; Extrême-Orient via Bangkok/Hong Kong).",
    "[S] Sourced: aircraft and real Air India corridors (Europe via Cairo/Rome; Far East via Bangkok/Hong Kong).");
  a.sheet("Route de l'Europe", 'Europe Route', [src, s1, REC, LEG, RET]).route({
    fwdNo: 'AI 105', retNo: 'AI 106',
    lineFwd: ["Route de l'Europe", 'Europe Route'],
    lineRet: ["Route de l'Europe (Retour)", 'Europe Route (Return)'],
    ports: ['Bombay (BOM)', 'Le Caire (CAI)', 'Rome', 'Genève', 'Londres (LHR)'],
    legMin: [420, 210, 150, 120], groundMin: [75, 75, 60], outDep: '02:00', retDep: '11:00',
    freqFwd: ['Mardi, Jeudi, Samedi', 'Tuesday, Thursday, Saturday'],
    freqRet: ['Mercredi, Vendredi, Dimanche', 'Wednesday, Friday, Sunday'],
  });
  a.sheet('Extrême-Orient', 'Far East', [LEG, RET]).route({
    fwdNo: 'AI 300', retNo: 'AI 301',
    lineFwd: ["Route de l'Extrême-Orient", 'Far East Route'],
    lineRet: ["Route de l'Extrême-Orient (Retour)", 'Far East Route (Return)'],
    ports: ['Bombay (BOM)', 'Calcutta', 'Bangkok', 'Hong Kong', 'Tokyo (HND)'],
    legMin: [240, 210, 180, 240], groundMin: [60, 75, 60], outDep: '23:00', retDep: '10:00',
    freqFwd: ['Mercredi, Dimanche', 'Wednesday, Sunday'],
    freqRet: ['Lundi, Vendredi', 'Monday, Friday'],
  });
  airlines.push(a);
}

// ===========================================================================
//  9. PAKISTAN INTERNATIONAL (PIA)  [S]
// ===========================================================================
{
  const a = airline('pia', 'Pakistan International', 'Reseau_PIA_L1049_sourced.xlsx');
  const src = N("Sources : Pakistan International Airlines (PIA) inaugure le L-1049C sur Karachi–Londres en 1955 et exploite le trunk Karachi–Dacca (survol de l'Inde).",
    "Sources: Pakistan International Airlines (PIA) inaugurated the L-1049C on Karachi–London in 1955 and operated the Karachi–Dacca trunk (overflying India).");
  const s1 = N("[S] Sourcé : appareil et axes PIA réels (ligne d'Europe ; liaison Karachi–Dacca entre les deux ailes du Pakistan).",
    "[S] Sourced: aircraft and real PIA axes (Europe line; Karachi–Dacca link between the two wings of Pakistan).");
  a.sheet("Ligne d'Europe", 'Europe Line', [src, s1, REC, LEG, RET]).route({
    fwdNo: 'PK 701', retNo: 'PK 702',
    lineFwd: ["Ligne d'Europe (Karachi)", 'Europe Line (Karachi)'],
    lineRet: ["Ligne d'Europe (Retour)", 'Europe Line (Return)'],
    ports: ['Karachi', 'Le Caire (CAI)', 'Rome', 'Londres (LHR)'],
    legMin: [300, 210, 150], groundMin: [75, 75], outDep: '01:00', retDep: '10:00',
    freqFwd: ['Lundi, Vendredi', 'Monday, Friday'],
    freqRet: ['Mardi, Samedi', 'Tuesday, Saturday'],
  });
  a.sheet('Trunk du Pakistan', 'Pakistan Trunk', [LEG, RET]).route({
    fwdNo: 'PK 601', retNo: 'PK 602',
    lineFwd: ['Trunk Est-Ouest', 'East–West Trunk'],
    lineRet: ['Trunk Est-Ouest (Retour)', 'East–West Trunk (Return)'],
    ports: ['Karachi', 'Dacca'],
    legMin: [300], groundMin: [], outDep: '07:00', retDep: '16:00',
    freqFwd: ['Quotidien', 'Daily'],
    freqRet: ['Quotidien', 'Daily'],
  });
  airlines.push(a);
}

// ===========================================================================
// 10. VARIG  [S]
// ===========================================================================
{
  const a = airline('varig', 'Varig', 'Reseau_Varig_L1049_sourced.xlsx');
  const src = N("Sources : la Varig reçoit 6 L-1049G en 1955 et inaugure Rio de Janeiro–New York (2 août 1955) via Belém et Ciudad Trujillo, puis Rio–Europe.",
    "Sources: Varig received 6 L-1049G in 1955 and inaugurated Rio de Janeiro–New York (2 August 1955) via Belém and Ciudad Trujillo, then Rio–Europe.");
  const s1 = N("[S] Sourcé : appareil et corridors Varig réels (Amérique du Nord via les Caraïbes ; Europe via Dakar/Lisbonne).",
    "[S] Sourced: aircraft and real Varig corridors (North America via the Caribbean; Europe via Dakar/Lisbon).");
  a.sheet('Amérique du Nord', 'North America', [src, s1, REC, LEG, RET]).route({
    fwdNo: 'RG 800', retNo: 'RG 801',
    lineFwd: ['Ligne de New York', 'New York Line'],
    lineRet: ['Ligne de New York (Retour)', 'New York Line (Return)'],
    ports: ['Rio de Janeiro (GIG)', 'Recife', 'Belém', 'Ciudad Trujillo', 'New York (IDL)'],
    legMin: [180, 150, 300, 300], groundMin: [60, 60, 75], outDep: '20:00', retDep: '18:00',
    freqFwd: ['Mardi, Vendredi', 'Tuesday, Friday'],
    freqRet: ['Jeudi, Dimanche', 'Thursday, Sunday'],
  });
  a.sheet("Ligne d'Europe", 'Europe Line', [LEG, RET]).route({
    fwdNo: 'RG 900', retNo: 'RG 901',
    lineFwd: ["Ligne d'Europe", 'Europe Line'],
    lineRet: ["Ligne d'Europe (Retour)", 'Europe Line (Return)'],
    ports: ['Rio de Janeiro (GIG)', 'Recife', 'Dakar', 'Lisbonne', 'Paris (ORY)'],
    legMin: [180, 600, 120, 180], groundMin: [60, 90, 75], outDep: '16:00', retDep: '13:00',
    freqFwd: ['Mercredi, Samedi', 'Wednesday, Saturday'],
    freqRet: ['Lundi, Jeudi', 'Monday, Thursday'],
  });
  airlines.push(a);
}

// 15. IBERIA  [S]
{
  const a = airline('iberia', 'Iberia', 'Reseau_Iberia_L1049_sourced.xlsx');
  const src = N("Sources : Iberia figure parmi les opérateurs de série du L-1049C ; lignes transatlantiques d'Espagne vers les Caraïbes (La Havane, San Juan) et l'Amérique du Sud.",
    "Sources: Iberia was among the production operators of the L-1049C; transatlantic lines from Spain to the Caribbean (Havana, San Juan) and South America.");
  const s1 = N("[S] Sourcé : appareil et corridors Iberia réels (Madrid–Caraïbes ; Madrid–Amérique du Sud via Santa Maria).",
    "[S] Sourced: aircraft and real Iberia corridors (Madrid–Caribbean; Madrid–South America via Santa Maria).");
  a.sheet('Atlantique & Caraïbes', 'Atlantic & Caribbean', [src, s1, REC, LEG, RET]).route({
    fwdNo: 'IB 951', retNo: 'IB 952',
    lineFwd: ['Ligne des Caraïbes', 'Caribbean Line'],
    lineRet: ['Ligne des Caraïbes (Retour)', 'Caribbean Line (Return)'],
    ports: ['Madrid', 'Santa Maria', 'San Juan', 'La Havane'],
    legMin: [120, 360, 120], groundMin: [75, 60], outDep: '10:00', retDep: '13:00',
    freqFwd: ['Mardi, Jeudi, Samedi', 'Tuesday, Thursday, Saturday'],
    freqRet: ['Mercredi, Vendredi, Dimanche', 'Wednesday, Friday, Sunday'],
  });
  a.sheet('Amérique du Sud', 'South America', [LEG, RET]).route({
    fwdNo: 'IB 971', retNo: 'IB 972',
    lineFwd: ["Ligne d'Amérique du Sud", 'South America Line'],
    lineRet: ["Ligne d'Amérique du Sud (Retour)", 'South America Line (Return)'],
    ports: ['Madrid', 'Santa Maria', 'Recife', 'Rio de Janeiro (GIG)'],
    legMin: [120, 480, 180], groundMin: [75, 90], outDep: '12:00', retDep: '16:00',
    freqFwd: ['Mercredi, Dimanche', 'Wednesday, Sunday'],
    freqRet: ['Lundi, Vendredi', 'Monday, Friday'],
  });
  airlines.push(a);
}

// 17. CAPITOL AIRWAYS  [S]
{
  const a = airline('capitol', 'Capitol Airways', 'Reseau_Capitol_L1049_sourced.xlsx');
  const src = N("Sources : Capitol Airways exploite le L-1049H (dès juillet 1959) et rachète des L-1049D/H à Seaboard World ; chartes civiles et militaires transatlantiques, contrat Berliner Flugring (1965) vers Munich, Palma, Naples, etc.",
    "Sources: Capitol Airways operated the L-1049H (from July 1959) and bought L-1049D/H from Seaboard World; transatlantic civil and military charters, Berliner Flugring contract (1965) to Munich, Palma, Naples, etc.");
  const s1 = N("[S] Sourcé : appareil (L-1049H, principal type de sa flotte de 23 Constellation) et activité de charte transatlantique/Flugring réelles.",
    "[S] Sourced: aircraft (L-1049H, the main type of its 23-Constellation fleet) and real transatlantic/Flugring charter activity.");
  const sh = a.sheet('Chartes Transatlantiques', 'Transatlantic Charters', [src, s1, REC, LEG, RET]);
  sh.route({
    fwdNo: 'CA 201', retNo: 'CA 202',
    lineFwd: ['Charte Transatlantique (Civile)', 'Transatlantic Charter (Civil)'],
    lineRet: ['Charte Transatlantique (Retour)', 'Transatlantic Charter (Return)'],
    ports: ['New York (IDL)', 'Gander [Tech]', 'Shannon', 'Londres (LHR)'],
    legMin: [150, 450, 120], groundMin: [60, 60], outDep: '21:00', retDep: '11:00',
    freqFwd: ['Mardi, Samedi', 'Tuesday, Saturday'],
    freqRet: ['Jeudi, Dimanche', 'Thursday, Sunday'],
  });
  sh.route({
    fwdNo: 'CA 601', retNo: 'CA 602',
    lineFwd: ['Charte Militaire MATS', 'MATS Military Charter'],
    lineRet: ['Charte Militaire MATS (Retour)', 'MATS Military Charter (Return)'],
    ports: ['New York (IDL)', 'Gander [Tech]', 'Francfort'],
    legMin: [150, 480], groundMin: [60], outDep: '20:00', retDep: '10:00',
    freqFwd: ['Vendredi', 'Friday'],
    freqRet: ['Dimanche', 'Sunday'],
  });
  a.sheet('Charte Vacances (Berlin)', 'Holiday Charter (Berlin)', [LEG, RET]).route({
    fwdNo: 'CA 401', retNo: 'CA 402',
    lineFwd: ['Charte Vacances (Flugring)', 'Holiday Charter (Flugring)'],
    lineRet: ['Charte Vacances (Retour)', 'Holiday Charter (Return)'],
    ports: ['Berlin (THF)', 'Munich', 'Palma'],
    legMin: [90, 120], groundMin: [45], outDep: '07:00', retDep: '12:00',
    freqFwd: ['Mercredi, Dimanche', 'Wednesday, Sunday'],
    freqRet: ['Mardi, Samedi', 'Tuesday, Saturday'],
  });
  airlines.push(a);
}

// 18. SEABOARD & WESTERN  [S] — cargo
{
  const a = airline('seaboard', 'Seaboard & Western', 'Reseau_Seaboard_Western_L1049_sourced.xlsx');
  const src = N("Sources : Seaboard & Western Airlines (future Seaboard World) est un opérateur cargo de série du L-1049D/H ; réseau fret transatlantique (Luxembourg, Francfort) et locations « wet-lease » (Sabena, Aerlínte, Capitol).",
    "Sources: Seaboard & Western Airlines (later Seaboard World) was a production cargo operator of the L-1049D/H; transatlantic freight network (Luxembourg, Frankfurt) and wet-lease operations (Sabena, Aerlínte, Capitol).");
  const s1 = N("[S] Sourcé : appareil (L-1049D/H cargo) et corridors fret transatlantiques réels de Seaboard.",
    "[S] Sourced: aircraft (L-1049D/H freighter) and real Seaboard transatlantic freight corridors.");
  const sh = a.sheet('Cargo Transatlantique', 'Transatlantic Cargo', [src, s1, REC, LEG, RET]);
  sh.route({
    fwdNo: 'SW 701', retNo: 'SW 702',
    lineFwd: ['Cargo Atlantique Nord', 'North Atlantic Cargo'],
    lineRet: ['Cargo Atlantique Nord (Retour)', 'North Atlantic Cargo (Return)'],
    ports: ['New York (IDL)', 'Gander [Tech]', 'Shannon', 'Luxembourg (LUX)'],
    legMin: [150, 450, 150], groundMin: [75, 90], outDep: '23:00', retDep: '22:00',
    freqFwd: ['Lundi, Mercredi, Vendredi', 'Monday, Wednesday, Friday'],
    freqRet: ['Mardi, Jeudi, Samedi', 'Tuesday, Thursday, Saturday'],
  });
  sh.route({
    fwdNo: 'SW 721', retNo: 'SW 722',
    lineFwd: ['Cargo Europe Centrale', 'Central Europe Cargo'],
    lineRet: ['Cargo Europe Centrale (Retour)', 'Central Europe Cargo (Return)'],
    ports: ['New York (IDL)', 'Gander [Tech]', 'Prestwick', 'Francfort'],
    legMin: [150, 420, 150], groundMin: [75, 90], outDep: '22:00', retDep: '21:00',
    freqFwd: ['Mardi, Samedi', 'Tuesday, Saturday'],
    freqRet: ['Mercredi, Dimanche', 'Wednesday, Sunday'],
  });
  airlines.push(a);
}

// 19. AVIANCA  [S]
{
  const a = airline('avianca', 'Avianca', 'Reseau_Avianca_L1049_sourced.xlsx');
  const src = N("Sources : Avianca (Colombie) figure parmi les opérateurs de série du L-1049C ; lignes Bogotá–New York (via Caraïbes) et Bogotá–Europe.",
    "Sources: Avianca (Colombia) was among the production operators of the L-1049C; Bogotá–New York (via the Caribbean) and Bogotá–Europe lines.");
  const s1 = N("[S] Sourcé : appareil et corridors Avianca réels (Amérique du Nord via Barranquilla/Kingston/Miami ; Europe via Lisbonne/Madrid).",
    "[S] Sourced: aircraft and real Avianca corridors (North America via Barranquilla/Kingston/Miami; Europe via Lisbon/Madrid).");
  a.sheet('Amérique du Nord', 'North America', [src, s1, REC, LEG, RET]).route({
    fwdNo: 'AV 5', retNo: 'AV 6',
    lineFwd: ['Ligne de New York', 'New York Line'],
    lineRet: ['Ligne de New York (Retour)', 'New York Line (Return)'],
    ports: ['Bogota (BOG)', 'Barranquilla', 'Kingston', 'Miami', 'New York (IDL)'],
    legMin: [90, 120, 180, 180], groundMin: [45, 45, 60], outDep: '07:00', retDep: '16:00',
    freqFwd: ['Mardi, Jeudi, Samedi', 'Tuesday, Thursday, Saturday'],
    freqRet: ['Mercredi, Vendredi, Dimanche', 'Wednesday, Friday, Sunday'],
  });
  a.sheet("Ligne d'Europe", 'Europe Line', [LEG, RET]).route({
    fwdNo: 'AV 11', retNo: 'AV 12',
    lineFwd: ["Ligne d'Europe", 'Europe Line'],
    lineRet: ["Ligne d'Europe (Retour)", 'Europe Line (Return)'],
    ports: ['Bogota (BOG)', 'Barranquilla', 'Santa Maria', 'Lisbonne', 'Madrid', 'Paris (ORY)'],
    legMin: [90, 600, 120, 90, 120], groundMin: [45, 90, 60, 60], outDep: '16:00', retDep: '10:00',
    freqFwd: ['Mercredi, Dimanche', 'Wednesday, Sunday'],
    freqRet: ['Lundi, Vendredi', 'Monday, Friday'],
  });
  airlines.push(a);
}

// 20. BOAC  [S] — L-1049 loués à Capitol pour les Caraïbes
{
  const a = airline('boac', 'BOAC', 'Reseau_BOAC_L1049_sourced.xlsx');
  const src = N("Sources : la BOAC exploitait principalement le Boeing 377, le Britannia puis le Comet ; elle a toutefois loué des L-1049G à Capitol Airways pour ses services des Caraïbes vers Londres (vers 1960).",
    "Sources: BOAC mainly operated the Boeing 377, the Britannia then the Comet; however it leased L-1049Gs from Capitol Airways for its Caribbean services to London (around 1960).");
  const s1 = N("[S] Sourcé : usage réel (mais ponctuel, en location) du L-1049 par la BOAC sur les Caraïbes ; corridors conformes.",
    "[S] Sourced: real (but occasional, leased) use of the L-1049 by BOAC on the Caribbean; corridors consistent.");
  const sh = a.sheet('Services des Caraïbes', 'Caribbean Services', [src, s1, REC, LEG, RET]);
  sh.route({
    fwdNo: 'BA 651', retNo: 'BA 652',
    lineFwd: ['Caraïbes (Le Speedbird)', 'Caribbean (The Speedbird)'],
    lineRet: ['Caraïbes (Retour)', 'Caribbean (Return)'],
    ports: ['Londres (LHR)', 'Santa Maria', 'Bermudes', 'Nassau', 'Kingston'],
    legMin: [180, 300, 150, 120], groundMin: [75, 60, 45], outDep: '11:00', retDep: '13:00',
    freqFwd: ['Mardi, Vendredi', 'Tuesday, Friday'],
    freqRet: ['Jeudi, Dimanche', 'Thursday, Sunday'],
  });
  sh.route({
    fwdNo: 'BA 661', retNo: 'BA 662',
    lineFwd: ['Antilles (Trinidad)', 'West Indies (Trinidad)'],
    lineRet: ['Antilles (Retour)', 'West Indies (Return)'],
    ports: ['Londres (LHR)', 'Santa Maria', 'Bermudes', 'Port of Spain'],
    legMin: [180, 300, 300], groundMin: [75, 60], outDep: '10:00', retDep: '14:00',
    freqFwd: ['Mercredi, Samedi', 'Wednesday, Saturday'],
    freqRet: ['Lundi, Jeudi', 'Monday, Thursday'],
  });
  airlines.push(a);
}

// ===========================================================================
// ==========  3e LOT : opérateurs RÉELS du L-1049 encore manquants  =========
// ===========================================================================

// 24. NORTHWEST ORIENT AIRLINES  [S]
{
  const a = airline('northwest', 'Northwest Orient Airlines', 'Reseau_Northwest_Orient_L1049_sourced.xlsx');
  const src = N("Sources : Northwest Orient Airlines introduit le L-1049G en service le 1er juillet 1955 ; 4 exemplaires réservés au réseau asiatique (Seattle–Tokyo–Okinawa–Manille), le DC-6 restant sur le domestique.",
    "Sources: Northwest Orient Airlines introduced the L-1049G into service on 1 July 1955; 4 aircraft dedicated to the Asian network (Seattle–Tokyo–Okinawa–Manila), the DC-6 remaining on domestic routes.");
  const s1 = N("[S] Sourcé : appareil (L-1049G) et réseau transpacifique/Orient réels de Northwest (Seattle–Tokyo, prolongements Okinawa/Manille).",
    "[S] Sourced: aircraft (L-1049G) and Northwest's real transpacific/Orient network (Seattle–Tokyo, extensions to Okinawa/Manila).");
  const sh = a.sheet('Transpacifique (Orient)', 'Transpacific (Orient)', [src, s1, REC, LEG, RET]);
  sh.route({
    fwdNo: 'NW 1', retNo: 'NW 2',
    lineFwd: ['Transpacifique (Le Tokyo)', 'Transpacific (The Tokyo)'],
    lineRet: ['Transpacifique (Retour)', 'Transpacific (Return)'],
    ports: ['Seattle (SEA)', 'Anchorage', 'Tokyo (HND)'],
    legMin: [240, 600], groundMin: [75], outDep: '12:00', retDep: '16:00',
    freqFwd: ['Mardi, Vendredi', 'Tuesday, Friday'],
    freqRet: ['Jeudi, Dimanche', 'Thursday, Sunday'],
  });
  sh.route({
    fwdNo: 'NW 5', retNo: 'NW 6',
    lineFwd: ['Extension Orient', 'Orient Extension'],
    lineRet: ['Extension Orient (Retour)', 'Orient Extension (Return)'],
    ports: ['Tokyo (HND)', 'Okinawa', 'Manille (MNL)'],
    legMin: [180, 180], groundMin: [60], outDep: '09:00', retDep: '14:00',
    freqFwd: ['Mercredi, Samedi', 'Wednesday, Saturday'],
    freqRet: ['Lundi, Jeudi', 'Monday, Thursday'],
  });
  airlines.push(a);
}

// 25. CUBANA DE AVIACIÓN  [S]
{
  const a = airline('cubana', 'Cubana de Aviación', 'Reseau_Cubana_L1049_sourced.xlsx');
  const src = N("Sources : Cubana de Aviación est le 1er opérateur latino-américain de Super Constellation (client de lancement du L-1049E, 1953) ; lignes La Havane–Madrid (via Bermudes, Açores, Lisbonne), New York et Mexico.",
    "Sources: Cubana de Aviación was the first Latin American operator of the Super Constellation (launch customer of the L-1049E, 1953); Havana–Madrid (via Bermuda, the Azores, Lisbon), New York and Mexico City lines.");
  const s1 = N("[S] Sourcé : appareil et corridors Cubana réels (transatlantique vers Madrid ; Amériques vers New York et Mexico).",
    "[S] Sourced: aircraft and real Cubana corridors (transatlantic to Madrid; the Americas to New York and Mexico City).");
  a.sheet('Transatlantique (Madrid)', 'Transatlantic (Madrid)', [src, s1, REC, LEG, RET]).route({
    fwdNo: 'CU 480', retNo: 'CU 481',
    lineFwd: ['Transatlantique (Madrid)', 'Transatlantic (Madrid)'],
    lineRet: ['Transatlantique (Retour)', 'Transatlantic (Return)'],
    ports: ['La Havane', 'Bermudes', 'Santa Maria', 'Lisbonne', 'Madrid'],
    legMin: [180, 360, 120, 90], groundMin: [60, 75, 45], outDep: '17:00', retDep: '12:00',
    freqFwd: ['Mardi, Samedi', 'Tuesday, Saturday'],
    freqRet: ['Jeudi, Dimanche', 'Thursday, Sunday'],
  });
  const sh = a.sheet('Amériques', 'Americas', [LEG, RET]);
  sh.route({
    fwdNo: 'CU 400', retNo: 'CU 401',
    lineFwd: ['Ligne de New York', 'New York Line'],
    lineRet: ['Ligne de New York (Retour)', 'New York Line (Return)'],
    ports: ['La Havane', 'New York (IDL)'],
    legMin: [210], groundMin: [], outDep: '08:00', retDep: '16:00',
    freqFwd: ['Quotidien', 'Daily'],
    freqRet: ['Quotidien', 'Daily'],
  });
  sh.route({
    fwdNo: 'CU 500', retNo: 'CU 501',
    lineFwd: ['Ligne de Mexico', 'Mexico City Line'],
    lineRet: ['Ligne de Mexico (Retour)', 'Mexico City Line (Return)'],
    ports: ['La Havane', 'Mexico (MEX)'],
    legMin: [150], groundMin: [], outDep: '10:00', retDep: '15:00',
    freqFwd: ['Lundi, Jeudi', 'Monday, Thursday'],
    freqRet: ['Mardi, Vendredi', 'Tuesday, Friday'],
  });
  airlines.push(a);
}

// 26. REAL TRANSPORTES AÉREOS  [S]
{
  const a = airline('real', 'Real Transportes Aéreos', 'Reseau_Real_L1049_sourced.xlsx');
  const src = N("Sources : Real Transportes Aéreos (Brésil) reçoit des L-1049H en 1958 et ouvre Rio de Janeiro–Los Angeles via Manaus, Bogotá et Mexico, prolongée jusqu'à Tokyo via Los Angeles.",
    "Sources: Real Transportes Aéreos (Brazil) received L-1049H in 1958 and opened Rio de Janeiro–Los Angeles via Manaus, Bogotá and Mexico City, extended to Tokyo via Los Angeles.");
  const s1 = N("[S] Sourcé : appareil (L-1049H) et corridors Real réels (Rio–Los Angeles via l'Amérique latine ; prolongement transpacifique vers Tokyo).",
    "[S] Sourced: aircraft (L-1049H) and real Real corridors (Rio–Los Angeles via Latin America; transpacific extension to Tokyo).");
  a.sheet('Amérique du Nord (Los Angeles)', 'North America (Los Angeles)', [src, s1, REC, LEG, RET]).route({
    fwdNo: 'RL 700', retNo: 'RL 701',
    lineFwd: ['Ligne de Los Angeles', 'Los Angeles Line'],
    lineRet: ['Ligne de Los Angeles (Retour)', 'Los Angeles Line (Return)'],
    ports: ['Rio de Janeiro (GIG)', 'Manaus', 'Bogota (BOG)', 'Mexico (MEX)', 'Los Angeles (LAX)'],
    legMin: [240, 240, 300, 240], groundMin: [60, 60, 75], outDep: '19:00', retDep: '17:00',
    freqFwd: ['Mardi, Vendredi', 'Tuesday, Friday'],
    freqRet: ['Jeudi, Dimanche', 'Thursday, Sunday'],
  });
  a.sheet('Ligne du Japon', 'Japan Line', [LEG, RET]).route({
    fwdNo: 'RL 800', retNo: 'RL 801',
    lineFwd: ['Ligne du Japon (via Los Angeles)', 'Japan Line (via Los Angeles)'],
    lineRet: ['Ligne du Japon (Retour)', 'Japan Line (Return)'],
    ports: ['Los Angeles (LAX)', 'Honolulu', 'Wake Island [Tech]', 'Tokyo (HND)'],
    legMin: [300, 300, 300], groundMin: [75, 60], outDep: '11:00', retDep: '13:00',
    freqFwd: ['Mercredi, Samedi', 'Wednesday, Saturday'],
    freqRet: ['Lundi, Jeudi', 'Monday, Thursday'],
  });
  airlines.push(a);
}

// ===========================================================================
// ============  4e LOT : LIGNES FICTIVES (vraies routes DC-4 / DC-6)  ========
// Compagnies n'ayant pas exploité le L-1049, mais dont ces lignes furent
// assurées en Douglas DC-4 / DC-6 — appareils de même catégorie. Classées
// « fictives » dans l'appli (filtre « Lignes fictives »).
// ===========================================================================
const DAILY = ['Quotidien', 'Daily'];

// --- Aerolíneas Argentinas (DC-4 + DC-6) ---
{
  const a = airline('aerolineas-argentinas', 'Aerolíneas Argentinas', 'Reseau_Aerolineas_Argentinas_L1049_sourced.xlsx');
  const src = N("Sources : Aerolíneas Argentinas (créée en 1950) a exploité des Douglas DC-4 sur ses lignes régionales et des DC-6 vers l'Amérique du Nord et l'Europe.",
    "Sources: Aerolíneas Argentinas (founded 1950) operated Douglas DC-4 on its regional lines and DC-6 to North America and Europe.");
  a.sheet('Lignes DC-4', 'DC-4 Lines', [src, DC4NOTE, REC, LEG, RET])
    .route({ fwdNo: 'AR 210', retNo: 'AR 211', lineFwd: ['Buenos Aires – Santiago', 'Buenos Aires – Santiago'], lineRet: ['Buenos Aires – Santiago (Retour)', 'Buenos Aires – Santiago (Return)'], ports: ['Buenos Aires (EZE)', 'Santiago'], legMin: [150], groundMin: [], outDep: '09:00', retDep: '15:00', freqFwd: DAILY, freqRet: DAILY })
    .route({ fwdNo: 'AR 240', retNo: 'AR 241', lineFwd: ['Río de la Plata', 'River Plate'], lineRet: ['Río de la Plata (Retour)', 'River Plate (Return)'], ports: ['Buenos Aires (EZE)', 'Montevideo'], legMin: [45], groundMin: [], outDep: '08:00', retDep: '18:00', freqFwd: DAILY, freqRet: DAILY });
  a.sheet('Lignes DC-6', 'DC-6 Lines', [DC6NOTE, LEG, RET])
    .route({ fwdNo: 'AR 300', retNo: 'AR 301', lineFwd: ['Atlantique Sud (Europe)', 'South Atlantic (Europe)'], lineRet: ['Atlantique Sud (Retour)', 'South Atlantic (Return)'], ports: ['Buenos Aires (EZE)', 'Rio de Janeiro (GIG)', 'Dakar', 'Lisbonne', 'Rome'], legMin: [180, 600, 120, 150], groundMin: [60, 90, 60], outDep: '16:00', retDep: '12:00', freqFwd: ['Mardi, Vendredi', 'Tuesday, Friday'], freqRet: ['Jeudi, Dimanche', 'Thursday, Sunday'] })
    .route({ fwdNo: 'AR 320', retNo: 'AR 321', lineFwd: ['Ligne de New York', 'New York Line'], lineRet: ['Ligne de New York (Retour)', 'New York Line (Return)'], ports: ['Buenos Aires (EZE)', 'Rio de Janeiro (GIG)', 'Caracas (CCS)', 'New York (IDL)'], legMin: [180, 540, 300], groundMin: [60, 75], outDep: '18:00', retDep: '16:00', freqFwd: ['Mercredi, Samedi', 'Wednesday, Saturday'], freqRet: ['Lundi, Jeudi', 'Monday, Thursday'] });
  airlines.push(a);
}

// --- Olympic Airways (DC-4 + DC-6) ---
{
  const a = airline('olympic', 'Olympic Airways', 'Reseau_Olympic_L1049_sourced.xlsx');
  const src = N("Sources : Olympic Airways (fondée en 1957) a hérité d'un DC-4 et exploité 13 Douglas DC-6/DC-6B (1958-1967) sur son réseau européen et moyen-oriental.",
    "Sources: Olympic Airways (founded 1957) inherited a DC-4 and operated 13 Douglas DC-6/DC-6B (1958-1967) across its European and Middle Eastern network.");
  a.sheet('Lignes DC-4', 'DC-4 Lines', [src, DC4NOTE, REC, LEG, RET])
    .route({ fwdNo: 'OA 200', retNo: 'OA 201', lineFwd: ['Athènes – Rome', 'Athens – Rome'], lineRet: ['Athènes – Rome (Retour)', 'Athens – Rome (Return)'], ports: ['Athènes', 'Rome'], legMin: [120], groundMin: [], outDep: '10:00', retDep: '16:00', freqFwd: ['Mardi, Jeudi, Samedi', 'Tuesday, Thursday, Saturday'], freqRet: ['Mercredi, Vendredi, Dimanche', 'Wednesday, Friday, Sunday'] })
    .route({ fwdNo: 'OA 210', retNo: 'OA 211', lineFwd: ['Égée (Rhodes)', 'Aegean (Rhodes)'], lineRet: ['Égée (Rhodes) (Retour)', 'Aegean (Rhodes) (Return)'], ports: ['Athènes', 'Rhodes'], legMin: [60], groundMin: [], outDep: '08:00', retDep: '18:00', freqFwd: DAILY, freqRet: DAILY });
  a.sheet('Lignes DC-6', 'DC-6 Lines', [DC6NOTE, LEG, RET])
    .route({ fwdNo: 'OA 410', retNo: 'OA 411', lineFwd: ['Athènes – Londres', 'Athens – London'], lineRet: ['Athènes – Londres (Retour)', 'Athens – London (Return)'], ports: ['Athènes', 'Rome', 'Paris (ORY)', 'Londres (LHR)'], legMin: [120, 120, 90], groundMin: [60, 60], outDep: '09:00', retDep: '13:00', freqFwd: ['Mardi, Vendredi', 'Tuesday, Friday'], freqRet: ['Jeudi, Dimanche', 'Thursday, Sunday'] })
    .route({ fwdNo: 'OA 430', retNo: 'OA 431', lineFwd: ['Levant (Beyrouth)', 'Levant (Beirut)'], lineRet: ['Levant (Beyrouth) (Retour)', 'Levant (Beirut) (Return)'], ports: ['Athènes', 'Istanbul', 'Beyrouth'], legMin: [90, 120], groundMin: [45], outDep: '11:00', retDep: '16:00', freqFwd: ['Mercredi, Samedi', 'Wednesday, Saturday'], freqRet: ['Lundi, Jeudi', 'Monday, Thursday'] });
  airlines.push(a);
}

// --- Cathay Pacific (DC-4 + DC-6) ---
{
  const a = airline('cathay-pacific', 'Cathay Pacific', 'Reseau_Cathay_Pacific_L1049_sourced.xlsx');
  const src = N("Sources : Cathay Pacific a introduit le Douglas DC-4 en 1949 et le DC-6 en 1958 sur son réseau asiatique depuis Hong Kong.",
    "Sources: Cathay Pacific introduced the Douglas DC-4 in 1949 and the DC-6 in 1958 across its Asian network from Hong Kong.");
  a.sheet('Lignes DC-4', 'DC-4 Lines', [src, DC4NOTE, REC, LEG, RET])
    .route({ fwdNo: 'CX 100', retNo: 'CX 101', lineFwd: ['Hong Kong – Manille', 'Hong Kong – Manila'], lineRet: ['Hong Kong – Manille (Retour)', 'Hong Kong – Manila (Return)'], ports: ['Hong Kong', 'Manille (MNL)'], legMin: [120], groundMin: [], outDep: '09:00', retDep: '14:00', freqFwd: DAILY, freqRet: DAILY })
    .route({ fwdNo: 'CX 110', retNo: 'CX 111', lineFwd: ['Détroits (Singapour)', 'Straits (Singapore)'], lineRet: ['Détroits (Singapour) (Retour)', 'Straits (Singapore) (Return)'], ports: ['Hong Kong', 'Bangkok', 'Singapour'], legMin: [180, 180], groundMin: [60], outDep: '08:00', retDep: '13:00', freqFwd: ['Mardi, Vendredi', 'Tuesday, Friday'], freqRet: ['Jeudi, Dimanche', 'Thursday, Sunday'] });
  a.sheet('Lignes DC-6', 'DC-6 Lines', [DC6NOTE, LEG, RET])
    .route({ fwdNo: 'CX 400', retNo: 'CX 401', lineFwd: ['Extrême-Orient (Tokyo)', 'Far East (Tokyo)'], lineRet: ['Extrême-Orient (Tokyo) (Retour)', 'Far East (Tokyo) (Return)'], ports: ['Hong Kong', 'Taipei', 'Tokyo (HND)'], legMin: [90, 180], groundMin: [45], outDep: '10:00', retDep: '15:00', freqFwd: ['Mercredi, Samedi', 'Wednesday, Saturday'], freqRet: ['Lundi, Jeudi', 'Monday, Thursday'] })
    .route({ fwdNo: 'CX 420', retNo: 'CX 421', lineFwd: ['Australie (Sydney)', 'Australia (Sydney)'], lineRet: ['Australie (Sydney) (Retour)', 'Australia (Sydney) (Return)'], ports: ['Hong Kong', 'Singapour', 'Darwin', 'Sydney (SYD)'], legMin: [180, 240, 240], groundMin: [60, 60], outDep: '22:00', retDep: '20:00', freqFwd: ['Mardi, Samedi', 'Tuesday, Saturday'], freqRet: ['Mercredi, Dimanche', 'Wednesday, Sunday'] });
  airlines.push(a);
}

// --- Icelandair / Loftleiðir (DC-4 + DC-6) ---
{
  const a = airline('icelandair', 'Icelandair (Loftleiðir)', 'Reseau_Icelandair_L1049_sourced.xlsx');
  const src = N("Sources : Loftleiðir (Icelandic Airlines) a lancé le DC-4 sur Reykjavik–New York (1948) puis Reykjavik–Copenhague ; DC-6B à partir de 1959, avec Luxembourg comme hub européen.",
    "Sources: Loftleiðir (Icelandic Airlines) launched the DC-4 on Reykjavik–New York (1948) then Reykjavik–Copenhagen; DC-6B from 1959, with Luxembourg as its European hub.");
  a.sheet('Lignes DC-4', 'DC-4 Lines', [src, DC4NOTE, REC, LEG, RET])
    .route({ fwdNo: 'LL 101', retNo: 'LL 102', lineFwd: ['Atlantique Nord (New York)', 'North Atlantic (New York)'], lineRet: ['Atlantique Nord (Retour)', 'North Atlantic (Return)'], ports: ['Reykjavik', 'Gander [Tech]', 'New York (IDL)'], legMin: [300, 300], groundMin: [75], outDep: '10:00', retDep: '16:00', freqFwd: ['Mardi, Vendredi', 'Tuesday, Friday'], freqRet: ['Mercredi, Samedi', 'Wednesday, Saturday'] })
    .route({ fwdNo: 'LL 120', retNo: 'LL 121', lineFwd: ['Scandinavie', 'Scandinavia'], lineRet: ['Scandinavie (Retour)', 'Scandinavia (Return)'], ports: ['Reykjavik', 'Oslo', 'Copenhague'], legMin: [180, 90], groundMin: [60], outDep: '09:00', retDep: '15:00', freqFwd: ['Lundi, Jeudi', 'Monday, Thursday'], freqRet: ['Mardi, Vendredi', 'Tuesday, Friday'] });
  a.sheet('Lignes DC-6', 'DC-6 Lines', [DC6NOTE, LEG, RET])
    .route({ fwdNo: 'LL 201', retNo: 'LL 202', lineFwd: ['Atlantique Nord (Luxembourg)', 'North Atlantic (Luxembourg)'], lineRet: ['Atlantique Nord (Luxembourg) (Retour)', 'North Atlantic (Luxembourg) (Return)'], ports: ['New York (IDL)', 'Reykjavik', 'Luxembourg (LUX)'], legMin: [300, 300], groundMin: [90], outDep: '18:00', retDep: '12:00', freqFwd: ['Mercredi, Dimanche', 'Wednesday, Sunday'], freqRet: ['Mardi, Samedi', 'Tuesday, Saturday'] })
    .route({ fwdNo: 'LL 210', retNo: 'LL 211', lineFwd: ['Hambourg (via Islande)', 'Hamburg (via Iceland)'], lineRet: ['Hambourg (via Islande) (Retour)', 'Hamburg (via Iceland) (Return)'], ports: ['New York (IDL)', 'Reykjavik', 'Hambourg'], legMin: [300, 330], groundMin: [90], outDep: '17:00', retDep: '11:00', freqFwd: ['Mardi, Vendredi', 'Tuesday, Friday'], freqRet: ['Jeudi, Dimanche', 'Thursday, Sunday'] });
  airlines.push(a);
}

// --- Alitalia (DC-4 + DC-6) ---
{
  const a = airline('alitalia', 'Alitalia', 'Reseau_Alitalia_L1049_sourced.xlsx');
  const src = N("Sources : Alitalia (et LAI) ont exploité des Douglas DC-4 sur les lignes européennes et méditerranéennes, et des DC-6/DC-6B vers l'Amérique.",
    "Sources: Alitalia (and LAI) operated Douglas DC-4 on European and Mediterranean lines, and DC-6/DC-6B to the Americas.");
  a.sheet('Lignes DC-4', 'DC-4 Lines', [src, DC4NOTE, REC, LEG, RET])
    .route({ fwdNo: 'AZ 200', retNo: 'AZ 201', lineFwd: ['Rome – Paris', 'Rome – Paris'], lineRet: ['Rome – Paris (Retour)', 'Rome – Paris (Return)'], ports: ['Rome', 'Milan', 'Paris (ORY)'], legMin: [75, 90], groundMin: [45], outDep: '09:00', retDep: '16:00', freqFwd: DAILY, freqRet: DAILY })
    .route({ fwdNo: 'AZ 220', retNo: 'AZ 221', lineFwd: ['Méditerranée (Tripoli)', 'Mediterranean (Tripoli)'], lineRet: ['Méditerranée (Tripoli) (Retour)', 'Mediterranean (Tripoli) (Return)'], ports: ['Rome', 'Tripoli'], legMin: [180], groundMin: [], outDep: '10:00', retDep: '15:00', freqFwd: ['Mardi, Jeudi, Samedi', 'Tuesday, Thursday, Saturday'], freqRet: ['Mercredi, Vendredi, Dimanche', 'Wednesday, Friday, Sunday'] });
  a.sheet('Lignes DC-6', 'DC-6 Lines', [DC6NOTE, LEG, RET])
    .route({ fwdNo: 'AZ 610', retNo: 'AZ 611', lineFwd: ['Ligne de New York', 'New York Line'], lineRet: ['Ligne de New York (Retour)', 'New York Line (Return)'], ports: ['Rome', 'Shannon', 'New York (IDL)'], legMin: [300, 300], groundMin: [75], outDep: '10:00', retDep: '16:00', freqFwd: ['Mardi, Vendredi', 'Tuesday, Friday'], freqRet: ['Jeudi, Dimanche', 'Thursday, Sunday'] })
    .route({ fwdNo: 'AZ 630', retNo: 'AZ 631', lineFwd: ['Atlantique Sud (Buenos Aires)', 'South Atlantic (Buenos Aires)'], lineRet: ['Atlantique Sud (Retour)', 'South Atlantic (Return)'], ports: ['Rome', 'Dakar', 'Recife', 'Rio de Janeiro (GIG)', 'Buenos Aires (EZE)'], legMin: [360, 600, 180, 180], groundMin: [75, 60, 60], outDep: '20:00', retDep: '18:00', freqFwd: ['Mercredi, Dimanche', 'Wednesday, Sunday'], freqRet: ['Lundi, Vendredi', 'Monday, Friday'] });
  airlines.push(a);
}

// --- Japan Airlines (DC-4 + DC-6) ---
{
  const a = airline('jal', 'Japan Airlines', 'Reseau_JAL_L1049_sourced.xlsx');
  const src = N("Sources : Japan Airlines (fondée en 1951) a exploité des Douglas DC-4 sur son réseau intérieur, puis des DC-6/DC-6B, dont le transpacifique Tokyo–San Francisco (1954).",
    "Sources: Japan Airlines (founded 1951) operated Douglas DC-4 on its domestic network, then DC-6/DC-6B, including the Tokyo–San Francisco transpacific (1954).");
  a.sheet('Lignes DC-4', 'DC-4 Lines', [src, DC4NOTE, REC, LEG, RET])
    .route({ fwdNo: 'JL 300', retNo: 'JL 301', lineFwd: ['Domestique (Fukuoka)', 'Domestic (Fukuoka)'], lineRet: ['Domestique (Fukuoka) (Retour)', 'Domestic (Fukuoka) (Return)'], ports: ['Tokyo (HND)', 'Osaka', 'Fukuoka'], legMin: [60, 60], groundMin: [40], outDep: '08:00', retDep: '16:00', freqFwd: DAILY, freqRet: DAILY })
    .route({ fwdNo: 'JL 320', retNo: 'JL 321', lineFwd: ['Domestique (Sapporo)', 'Domestic (Sapporo)'], lineRet: ['Domestique (Sapporo) (Retour)', 'Domestic (Sapporo) (Return)'], ports: ['Tokyo (HND)', 'Sapporo'], legMin: [90], groundMin: [], outDep: '09:00', retDep: '17:00', freqFwd: DAILY, freqRet: DAILY });
  a.sheet('Lignes DC-6', 'DC-6 Lines', [DC6NOTE, LEG, RET])
    .route({ fwdNo: 'JL 800', retNo: 'JL 801', lineFwd: ['Transpacifique (San Francisco)', 'Transpacific (San Francisco)'], lineRet: ['Transpacifique (Retour)', 'Transpacific (Return)'], ports: ['Tokyo (HND)', 'Wake Island [Tech]', 'Honolulu', 'San Francisco (SFO)'], legMin: [300, 240, 300], groundMin: [60, 75], outDep: '19:00', retDep: '11:00', freqFwd: ['Mardi, Vendredi', 'Tuesday, Friday'], freqRet: ['Mercredi, Samedi', 'Wednesday, Saturday'] })
    .route({ fwdNo: 'JL 820', retNo: 'JL 821', lineFwd: ['Okinawa', 'Okinawa'], lineRet: ['Okinawa (Retour)', 'Okinawa (Return)'], ports: ['Tokyo (HND)', 'Okinawa'], legMin: [180], groundMin: [], outDep: '10:00', retDep: '15:00', freqFwd: ['Lundi, Jeudi', 'Monday, Thursday'], freqRet: ['Mardi, Vendredi', 'Tuesday, Friday'] });
  airlines.push(a);
}

// --- Aviateca (DC-4) ---
{
  const a = airline('aviateca', 'Aviateca', 'Reseau_Aviateca_L1049_sourced.xlsx');
  const src = N("Sources : Aviateca (Guatemala) a lancé son service vers Miami en Douglas DC-4 en 1957 (via Belize) ; réseau centraméricain.",
    "Sources: Aviateca (Guatemala) launched its Miami service with the Douglas DC-4 in 1957 (via Belize); Central American network.");
  a.sheet('Lignes DC-4', 'DC-4 Lines', [src, DC4NOTE, REC, LEG, RET])
    .route({ fwdNo: 'GU 500', retNo: 'GU 501', lineFwd: ['Guatemala – Miami', 'Guatemala – Miami'], lineRet: ['Guatemala – Miami (Retour)', 'Guatemala – Miami (Return)'], ports: ['Guatemala (GUA)', 'Belize', 'Miami'], legMin: [45, 120], groundMin: [40], outDep: '08:00', retDep: '15:00', freqFwd: ['Mardi, Jeudi, Samedi', 'Tuesday, Thursday, Saturday'], freqRet: ['Mercredi, Vendredi, Dimanche', 'Wednesday, Friday, Sunday'] })
    .route({ fwdNo: 'GU 520', retNo: 'GU 521', lineFwd: ['Amérique Centrale', 'Central America'], lineRet: ['Amérique Centrale (Retour)', 'Central America (Return)'], ports: ['Guatemala (GUA)', 'San Salvador', 'San Jose'], legMin: [45, 90], groundMin: [40], outDep: '09:00', retDep: '16:00', freqFwd: ['Lundi, Jeudi', 'Monday, Thursday'], freqRet: ['Mardi, Vendredi', 'Tuesday, Friday'] });
  airlines.push(a);
}

// --- El Al (DC-4) ---
{
  const a = airline('el-al', 'El Al', 'Reseau_El_Al_L1049_sourced.xlsx');
  const src = N("Sources : El Al (Israël) a exploité des Douglas DC-4 sur ses premières lignes européennes (1950-1951) avant de passer aux Constellation.",
    "Sources: El Al (Israel) operated Douglas DC-4 on its first European lines (1950-1951) before switching to Constellations.");
  a.sheet('Lignes DC-4', 'DC-4 Lines', [src, DC4NOTE, REC, LEG, RET])
    .route({ fwdNo: 'LY 300', retNo: 'LY 301', lineFwd: ["Ligne de l'Ouest (Londres)", 'Western Line (London)'], lineRet: ["Ligne de l'Ouest (Retour)", 'Western Line (Return)'], ports: ['Tel Aviv (LOD)', 'Rome', 'Paris (ORY)', 'Londres (LHR)'], legMin: [180, 120, 90], groundMin: [60, 60], outDep: '08:00', retDep: '13:00', freqFwd: ['Mardi, Vendredi', 'Tuesday, Friday'], freqRet: ['Jeudi, Dimanche', 'Thursday, Sunday'] })
    .route({ fwdNo: 'LY 320', retNo: 'LY 321', lineFwd: ['Méditerranée (Athènes)', 'Mediterranean (Athens)'], lineRet: ['Méditerranée (Athènes) (Retour)', 'Mediterranean (Athens) (Return)'], ports: ['Tel Aviv (LOD)', 'Athènes', 'Rome'], legMin: [120, 120], groundMin: [45], outDep: '10:00', retDep: '16:00', freqFwd: ['Mercredi, Samedi', 'Wednesday, Saturday'], freqRet: ['Lundi, Jeudi', 'Monday, Thursday'] });
  airlines.push(a);
}

// --- American Overseas Airlines (DC-4) ---
{
  const a = airline('american-overseas', 'American Overseas Airlines', 'Reseau_American_Overseas_L1049_sourced.xlsx');
  const src = N("Sources : American Overseas Airlines a exploité le Douglas DC-4 sur l'Atlantique Nord (New York–Europe/Scandinavie) de 1945 à sa fusion dans Pan Am en 1950.",
    "Sources: American Overseas Airlines operated the Douglas DC-4 on the North Atlantic (New York–Europe/Scandinavia) from 1945 until its 1950 merger into Pan Am.");
  a.sheet('Lignes DC-4', 'DC-4 Lines', [src, DC4NOTE, REC, LEG, RET])
    .route({ fwdNo: 'AO 101', retNo: 'AO 102', lineFwd: ['Atlantique Nord (Londres)', 'North Atlantic (London)'], lineRet: ['Atlantique Nord (Retour)', 'North Atlantic (Return)'], ports: ['New York (IDL)', 'Gander [Tech]', 'Shannon', 'Londres (LHR)'], legMin: [300, 300, 120], groundMin: [75, 60], outDep: '18:00', retDep: '12:00', freqFwd: ['Mardi, Jeudi, Samedi', 'Tuesday, Thursday, Saturday'], freqRet: ['Mercredi, Vendredi, Dimanche', 'Wednesday, Friday, Sunday'] })
    .route({ fwdNo: 'AO 151', retNo: 'AO 152', lineFwd: ['Scandinavie (Copenhague)', 'Scandinavia (Copenhagen)'], lineRet: ['Scandinavie (Copenhague) (Retour)', 'Scandinavia (Copenhagen) (Return)'], ports: ['New York (IDL)', 'Gander [Tech]', 'Prestwick', 'Copenhague'], legMin: [300, 300, 150], groundMin: [75, 60], outDep: '17:00', retDep: '10:00', freqFwd: ['Lundi, Vendredi', 'Monday, Friday'], freqRet: ['Mercredi, Dimanche', 'Wednesday, Sunday'] });
  airlines.push(a);
}

// --- Luxair (pas de DC-4/DC-6 : Lockheed L-1649A Starliner, famille Constellation) ---
{
  const a = airline('luxair', 'Luxair', 'Reseau_Luxair_L1049_sourced.xlsx');
  const src = N("Sources : la Luxair (fondée en 1962) n'a exploité NI DC-4 NI DC-6, mais a mis en ligne 3 Lockheed L-1649A Starliner — dernier de la lignée Constellation — sur Luxembourg–Johannesburg avec Trek Airways (1964-1969).",
    "Sources: Luxair (founded 1962) operated NEITHER the DC-4 NOR the DC-6, but flew 3 Lockheed L-1649A Starliners — the last of the Constellation line — on Luxembourg–Johannesburg with Trek Airways (1964-1969).");
  const star = N("Ligne exploitée en Lockheed L-1649A Starliner (famille Constellation) ; proposée ici pour le L-1049 Super Constellation, son proche parent.",
    "Route operated with the Lockheed L-1649A Starliner (Constellation family); offered here for the L-1049 Super Constellation, its close relative.");
  a.sheet('Ligne Starliner', 'Starliner Line', [src, star, REC, LEG, RET])
    .route({ fwdNo: 'LG 401', retNo: 'LG 402', lineFwd: ['Luxembourg – Johannesburg', 'Luxembourg – Johannesburg'], lineRet: ['Luxembourg – Johannesburg (Retour)', 'Luxembourg – Johannesburg (Return)'], ports: ['Luxembourg (LUX)', 'Athènes', 'Khartoum', 'Nairobi', 'Johannesburg (JNB)'], legMin: [180, 240, 240, 240], groundMin: [75, 60, 60], outDep: '20:00', retDep: '18:00', freqFwd: ['Mercredi', 'Wednesday'], freqRet: ['Vendredi', 'Friday'] });
  airlines.push(a);
}

// --- Aeroméxico (DC-6) ---
{
  const a = airline('aeromexico', 'Aeroméxico', 'Reseau_Aeromexico_L1049_sourced.xlsx');
  const src = N("Sources : Aeronaves de México (Aeroméxico) a exploité des Douglas DC-6 sur ses lignes vers les États-Unis et son réseau intérieur.",
    "Sources: Aeronaves de México (Aeroméxico) operated Douglas DC-6 on its routes to the United States and its domestic network.");
  a.sheet('Lignes DC-6', 'DC-6 Lines', [src, DC6NOTE, REC, LEG, RET])
    .route({ fwdNo: 'AM 400', retNo: 'AM 401', lineFwd: ['Ligne de New York', 'New York Line'], lineRet: ['Ligne de New York (Retour)', 'New York Line (Return)'], ports: ['Mexico (MEX)', 'New York (IDL)'], legMin: [300], groundMin: [], outDep: '08:00', retDep: '16:00', freqFwd: ['Mardi, Jeudi, Samedi', 'Tuesday, Thursday, Saturday'], freqRet: ['Mercredi, Vendredi, Dimanche', 'Wednesday, Friday, Sunday'] })
    .route({ fwdNo: 'AM 420', retNo: 'AM 421', lineFwd: ['Frontière (Los Angeles)', 'Border (Los Angeles)'], lineRet: ['Frontière (Los Angeles) (Retour)', 'Border (Los Angeles) (Return)'], ports: ['Mexico (MEX)', 'Tijuana', 'Los Angeles (LAX)'], legMin: [210, 60], groundMin: [45], outDep: '09:00', retDep: '15:00', freqFwd: DAILY, freqRet: DAILY })
    .route({ fwdNo: 'AM 440', retNo: 'AM 441', lineFwd: ['Yucatán – Cuba', 'Yucatán – Cuba'], lineRet: ['Yucatán – Cuba (Retour)', 'Yucatán – Cuba (Return)'], ports: ['Mexico (MEX)', 'Mérida', 'La Havane'], legMin: [120, 90], groundMin: [45], outDep: '10:00', retDep: '14:00', freqFwd: ['Lundi, Jeudi', 'Monday, Thursday'], freqRet: ['Mardi, Vendredi', 'Tuesday, Friday'] });
  airlines.push(a);
}

// --- SATA Air Açores (DC-6) ---
{
  const a = airline('sata', 'SATA Air Açores', 'Reseau_SATA_L1049_sourced.xlsx');
  const src = N("Sources : la SATA (Açores) a exploité deux Douglas DC-6 (ex-Force aérienne portugaise) pour relier les Açores à Lisbonne, notamment lors de la grève de la TAP en 1976.",
    "Sources: SATA (Azores) operated two Douglas DC-6 (ex-Portuguese Air Force) linking the Azores to Lisbon, notably during the 1976 TAP strike.");
  a.sheet('Lignes DC-6', 'DC-6 Lines', [src, DC6NOTE, REC, LEG, RET])
    .route({ fwdNo: 'SP 100', retNo: 'SP 101', lineFwd: ['Açores – Lisbonne', 'Azores – Lisbon'], lineRet: ['Açores – Lisbonne (Retour)', 'Azores – Lisbon (Return)'], ports: ['Ponta Delgada', 'Santa Maria', 'Lisbonne'], legMin: [60, 150], groundMin: [45], outDep: '09:00', retDep: '15:00', freqFwd: ['Mardi, Vendredi', 'Tuesday, Friday'], freqRet: ['Mercredi, Samedi', 'Wednesday, Saturday'] })
    .route({ fwdNo: 'SP 120', retNo: 'SP 121', lineFwd: ['Inter-îles (Santa Maria)', 'Inter-island (Santa Maria)'], lineRet: ['Inter-îles (Santa Maria) (Retour)', 'Inter-island (Santa Maria) (Return)'], ports: ['Ponta Delgada', 'Santa Maria'], legMin: [45], groundMin: [], outDep: '08:00', retDep: '17:00', freqFwd: DAILY, freqRet: DAILY });
  airlines.push(a);
}

// --- Pan Am (DC-6B) ---
{
  const a = airline('pan-am', 'Pan Am', 'Reseau_Pan_Am_L1049_sourced.xlsx');
  const src = N("Sources : Pan American a exploité de nombreux Douglas DC-6B (« President » Special) sur l'Atlantique Nord, les Caraïbes et l'Amérique latine.",
    "Sources: Pan American operated many Douglas DC-6B ('President' Special) on the North Atlantic, the Caribbean and Latin America.");
  a.sheet('Lignes DC-6', 'DC-6 Lines', [src, DC6NOTE, REC, LEG, RET])
    .route({ fwdNo: 'PA 100', retNo: 'PA 101', lineFwd: ['Atlantique Nord (Londres)', 'North Atlantic (London)'], lineRet: ['Atlantique Nord (Retour)', 'North Atlantic (Return)'], ports: ['New York (IDL)', 'Gander [Tech]', 'Shannon', 'Londres (LHR)'], legMin: [300, 300, 120], groundMin: [75, 60], outDep: '19:00', retDep: '12:00', freqFwd: ['Mardi, Jeudi, Samedi', 'Tuesday, Thursday, Saturday'], freqRet: ['Mercredi, Vendredi, Dimanche', 'Wednesday, Friday, Sunday'] })
    .route({ fwdNo: 'PA 400', retNo: 'PA 401', lineFwd: ['Caraïbes (San Juan)', 'Caribbean (San Juan)'], lineRet: ['Caraïbes (San Juan) (Retour)', 'Caribbean (San Juan) (Return)'], ports: ['New York (IDL)', 'Bermudes', 'San Juan'], legMin: [180, 300], groundMin: [60], outDep: '10:00', retDep: '16:00', freqFwd: ['Mercredi, Samedi', 'Wednesday, Saturday'], freqRet: ['Lundi, Jeudi', 'Monday, Thursday'] })
    .route({ fwdNo: 'PA 500', retNo: 'PA 501', lineFwd: ['Amérique Latine (Panama)', 'Latin America (Panama)'], lineRet: ['Amérique Latine (Retour)', 'Latin America (Return)'], ports: ['Miami', 'La Havane', 'Panama'], legMin: [90, 180], groundMin: [45], outDep: '08:00', retDep: '15:00', freqFwd: ['Mardi, Vendredi', 'Tuesday, Friday'], freqRet: ['Jeudi, Dimanche', 'Thursday, Sunday'] });
  airlines.push(a);
}

// --- Buffalo Airways (DC-6) ---
{
  const a = airline('buffalo', 'Buffalo Airways', 'Reseau_Buffalo_L1049_sourced.xlsx');
  const src = N("Sources : Buffalo Airways (Grand Nord canadien) exploite des Douglas DC-4 et possède un DC-6 « Swingtail » ; réseau de fret et de passagers dans les Territoires du Nord-Ouest.",
    "Sources: Buffalo Airways (Canadian Far North) operates Douglas DC-4 and owns a 'Swingtail' DC-6; freight and passenger network in the Northwest Territories.");
  a.sheet('Lignes DC-6', 'DC-6 Lines', [src, DC6NOTE, REC, LEG, RET])
    .route({ fwdNo: 'BF 101', retNo: 'BF 102', lineFwd: ['Vallée du Mackenzie', 'Mackenzie Valley'], lineRet: ['Vallée du Mackenzie (Retour)', 'Mackenzie Valley (Return)'], ports: ['Yellowknife', 'Norman Wells', 'Fort Good Hope'], legMin: [150, 90], groundMin: [45], outDep: '09:00', retDep: '14:00', freqFwd: ['Mardi, Vendredi', 'Tuesday, Friday'], freqRet: ['Mercredi, Samedi', 'Wednesday, Saturday'] })
    .route({ fwdNo: 'BF 120', retNo: 'BF 121', lineFwd: ['Yellowknife – Hay River', 'Yellowknife – Hay River'], lineRet: ['Yellowknife – Hay River (Retour)', 'Yellowknife – Hay River (Return)'], ports: ['Yellowknife', 'Hay River', 'Edmonton (YEG)'], legMin: [60, 150], groundMin: [40], outDep: '08:00', retDep: '16:00', freqFwd: ['Lundi, Jeudi', 'Monday, Thursday'], freqRet: ['Mardi, Vendredi', 'Tuesday, Friday'] });
  airlines.push(a);
}

// --- Nordair (DC-6) ---
{
  const a = airline('nordair', 'Nordair', 'Reseau_Nordair_L1049_sourced.xlsx');
  const src = N("Sources : Nordair (Québec) a exploité des Douglas DC-4 et un DC-6 sur ses lignes arctiques (Montréal–Frobisher Bay–Cape Dyer, Montréal–Roberval–Fort Chimo) vers 1959.",
    "Sources: Nordair (Quebec) operated Douglas DC-4 and one DC-6 on its Arctic lines (Montreal–Frobisher Bay–Cape Dyer, Montreal–Roberval–Fort Chimo) around 1959.");
  a.sheet('Lignes DC-6', 'DC-6 Lines', [src, DC6NOTE, REC, LEG, RET])
    .route({ fwdNo: 'ND 201', retNo: 'ND 202', lineFwd: ['Baffin (Cape Dyer)', 'Baffin (Cape Dyer)'], lineRet: ['Baffin (Cape Dyer) (Retour)', 'Baffin (Cape Dyer) (Return)'], ports: ['Montréal (YUL)', 'Frobisher Bay', 'Cape Dyer'], legMin: [300, 120], groundMin: [60], outDep: '07:30', retDep: '13:00', freqFwd: ['Mardi, Vendredi', 'Tuesday, Friday'], freqRet: ['Mercredi, Samedi', 'Wednesday, Saturday'] })
    .route({ fwdNo: 'ND 220', retNo: 'ND 221', lineFwd: ['Ungava (Fort Chimo)', 'Ungava (Fort Chimo)'], lineRet: ['Ungava (Fort Chimo) (Retour)', 'Ungava (Fort Chimo) (Return)'], ports: ['Montréal (YUL)', 'Roberval', 'Fort Chimo'], legMin: [90, 240], groundMin: [40], outDep: '08:00', retDep: '15:00', freqFwd: ['Lundi, Jeudi', 'Monday, Thursday'], freqRet: ['Mardi, Vendredi', 'Tuesday, Friday'] });
  airlines.push(a);
}

// --- South African Airways (Springbok en DC-4 Skymaster) ---
{
  const a = airline('saa', 'South African Airways', 'Reseau_South_African_L1049_sourced.xlsx');
  const src = N("Sources : la South African Airways a assuré son « Springbok Service » (Johannesburg–Londres) en Douglas DC-4 Skymaster à partir de 1950, avant les Constellation.",
    "Sources: South African Airways flew its 'Springbok Service' (Johannesburg–London) with the Douglas DC-4 Skymaster from 1950, before the Constellations.");
  a.sheet('Lignes DC-4', 'DC-4 Lines', [src, DC4NOTE, REC, LEG, RET])
    .route({ fwdNo: 'SA 234', retNo: 'SA 235', lineFwd: ['Springbok (Londres)', 'Springbok (London)'], lineRet: ['Springbok (Retour)', 'Springbok (Return)'], ports: ['Johannesburg (JNB)', 'Nairobi', 'Khartoum', 'Le Caire (CAI)', 'Rome', 'Londres (LHR)'], legMin: [360, 240, 180, 210, 150], groundMin: [75, 60, 60, 75], outDep: '18:00', retDep: '11:00', freqFwd: ['Mardi, Jeudi, Samedi', 'Tuesday, Thursday, Saturday'], freqRet: ['Mercredi, Vendredi, Dimanche', 'Wednesday, Friday, Sunday'] });
  airlines.push(a);
}

// --- Swissair (DC-6/DC-6B) ---
{
  const a = airline('swissair', 'Swissair', 'Reseau_Swissair_L1049_sourced.xlsx');
  const src = N("Sources : la Swissair a exploité des Douglas DC-4 puis DC-6/DC-6B, notamment sur l'Atlantique Nord (Zurich–New York) et vers la Méditerranée.",
    "Sources: Swissair operated Douglas DC-4 then DC-6/DC-6B, notably on the North Atlantic (Zurich–New York) and to the Mediterranean.");
  a.sheet('Lignes DC-6', 'DC-6 Lines', [src, DC6NOTE, REC, LEG, RET])
    .route({ fwdNo: 'SR 100', retNo: 'SR 101', lineFwd: ['Atlantique Nord (New York)', 'North Atlantic (New York)'], lineRet: ['Atlantique Nord (Retour)', 'North Atlantic (Return)'], ports: ['Zürich (ZRH)', 'Genève', 'Shannon', 'New York (IDL)'], legMin: [60, 420, 300], groundMin: [45, 60], outDep: '09:00', retDep: '16:00', freqFwd: ['Mardi, Vendredi', 'Tuesday, Friday'], freqRet: ['Jeudi, Dimanche', 'Thursday, Sunday'] })
    .route({ fwdNo: 'SR 320', retNo: 'SR 321', lineFwd: ['Méditerranée (Le Caire)', 'Mediterranean (Cairo)'], lineRet: ['Méditerranée (Retour)', 'Mediterranean (Return)'], ports: ['Zürich (ZRH)', 'Rome', 'Athènes', 'Le Caire (CAI)'], legMin: [90, 120, 180], groundMin: [60, 60], outDep: '08:00', retDep: '14:00', freqFwd: ['Mercredi, Samedi', 'Wednesday, Saturday'], freqRet: ['Lundi, Jeudi', 'Monday, Thursday'] });
  airlines.push(a);
}

// --- Alaska Airlines (DC-6) ---
{
  const a = airline('alaska', 'Alaska Airlines', 'Reseau_Alaska_L1049_sourced.xlsx');
  const src = N("Sources : Alaska Airlines a exploité des Douglas DC-4 et DC-6 sur son réseau intérieur de l'Alaska.",
    "Sources: Alaska Airlines operated Douglas DC-4 and DC-6 across its Alaska domestic network.");
  a.sheet('Lignes DC-6', 'DC-6 Lines', [src, DC6NOTE, REC, LEG, RET])
    .route({ fwdNo: 'AS 60', retNo: 'AS 61', lineFwd: ['Ligne du Grand Nord', 'Great North Line'], lineRet: ['Ligne du Grand Nord (Retour)', 'Great North Line (Return)'], ports: ['Seattle (SEA)', 'Anchorage', 'Fairbanks'], legMin: [240, 90], groundMin: [60], outDep: '08:00', retDep: '14:00', freqFwd: DAILY, freqRet: DAILY })
    .route({ fwdNo: 'AS 80', retNo: 'AS 81', lineFwd: ['Ouest Alaskien (Nome)', 'Western Alaska (Nome)'], lineRet: ['Ouest Alaskien (Nome) (Retour)', 'Western Alaska (Nome) (Return)'], ports: ['Anchorage', 'Nome'], legMin: [120], groundMin: [], outDep: '10:00', retDep: '13:00', freqFwd: ['Lundi, Jeudi', 'Monday, Thursday'], freqRet: ['Mardi, Vendredi', 'Tuesday, Friday'] });
  airlines.push(a);
}

// --- United Air Lines (DC-6 Mainliner) ---
{
  const a = airline('united', 'United Air Lines', 'Reseau_United_L1049_sourced.xlsx');
  const src = N("Sources : United Air Lines a exploité de nombreux Douglas DC-6 « Mainliner » sur ses lignes transcontinentales et de la côte Ouest.",
    "Sources: United Air Lines operated many Douglas DC-6 'Mainliner' on its transcontinental and West Coast routes.");
  a.sheet('Lignes DC-6', 'DC-6 Lines', [src, DC6NOTE, REC, LEG, RET])
    .route({ fwdNo: 'UA 1', retNo: 'UA 2', lineFwd: ['Transcontinental (Mainliner)', 'Transcontinental (Mainliner)'], lineRet: ['Transcontinental (Retour)', 'Transcontinental (Return)'], ports: ['San Francisco (SFO)', 'Denver', 'Chicago', 'New York (IDL)'], legMin: [180, 150, 150], groundMin: [45, 45], outDep: '08:00', retDep: '09:00', freqFwd: DAILY, freqRet: DAILY })
    .route({ fwdNo: 'UA 5', retNo: 'UA 6', lineFwd: ['Côte Ouest–Hawaï', 'West Coast–Hawaii'], lineRet: ['Côte Ouest–Hawaï (Retour)', 'West Coast–Hawaii (Return)'], ports: ['San Francisco (SFO)', 'Honolulu'], legMin: [540], groundMin: [], outDep: '10:00', retDep: '16:00', freqFwd: ['Mardi, Jeudi, Samedi', 'Tuesday, Thursday, Saturday'], freqRet: ['Mercredi, Vendredi, Dimanche', 'Wednesday, Friday, Sunday'] });
  airlines.push(a);
}

// --- Delta Air Lines (DC-6) ---
{
  const a = airline('delta', 'Delta Air Lines', 'Reseau_Delta_L1049_sourced.xlsx');
  const src = N("Sources : Delta Air Lines a exploité des Douglas DC-6 sur son réseau du Sud et de la Floride.",
    "Sources: Delta Air Lines operated Douglas DC-6 across its Southern and Florida network.");
  a.sheet('Lignes DC-6', 'DC-6 Lines', [src, DC6NOTE, REC, LEG, RET])
    .route({ fwdNo: 'DL 41', retNo: 'DL 42', lineFwd: ['Transcontinental Sud', 'Southern Transcontinental'], lineRet: ['Transcontinental Sud (Retour)', 'Southern Transcontinental (Return)'], ports: ['Atlanta (ATL)', 'Dallas', 'Los Angeles'], legMin: [150, 180], groundMin: [45], outDep: '08:00', retDep: '15:00', freqFwd: DAILY, freqRet: DAILY })
    .route({ fwdNo: 'DL 61', retNo: 'DL 62', lineFwd: ['Ligne de Floride', 'Florida Line'], lineRet: ['Ligne de Floride (Retour)', 'Florida Line (Return)'], ports: ['Chicago', 'Atlanta (ATL)', 'Miami'], legMin: [120, 120], groundMin: [45], outDep: '09:00', retDep: '14:00', freqFwd: ['Mardi, Jeudi, Samedi', 'Tuesday, Thursday, Saturday'], freqRet: ['Mercredi, Vendredi, Dimanche', 'Wednesday, Friday, Sunday'] });
  airlines.push(a);
}

// ---------------------------------------------------------------------------
// Écriture des fichiers Excel.
// ---------------------------------------------------------------------------
const SRC_DIR = path.resolve(__dirname, '../../liaisons_Super_Constellation');

function buildSheetAoa(sheet) {
  const maxEsc = sheet.flights.reduce((m, f) => Math.max(m, f.escales.length), 0);
  const header = ['N° Vol', 'Axe / Ligne', 'Origine', 'Départ (HL)'];
  for (let i = 1; i <= maxEsc; i++) header.push(`Escale ${i} (Arr. / Dép.)`);
  header.push('Destination', 'Arrivée (HL)', 'Fréquence / Note');
  const width = header.length;

  const aoa = [header];
  for (const f of sheet.flights) {
    const esc = f.escales.slice();
    while (esc.length < maxEsc) esc.push('-');
    aoa.push([f.no, f.line, f.port0, f.originDep, ...esc, f.portLast, f.destArr, f.freq]);
  }
  // Ligne vide + bloc de notes.
  aoa.push(new Array(width).fill(''));
  const notesRow = new Array(width).fill('');
  notesRow[0] = 'Notes :';
  aoa.push(notesRow);
  for (const note of sheet.notes) {
    const r = new Array(width).fill('');
    r[0] = note;
    aoa.push(r);
  }
  return aoa;
}

const nFlights = (s) => (s.flights ? s.flights.length : (s.flightCount || 0));
let flightTotal = 0;
for (const a of airlines) {
  const wb = XLSX.utils.book_new();
  for (const sheet of a.sheets) {
    // Feuille verbatim (sheet.aoa) ou reconstruite par le moteur d'horloge.
    const ws = XLSX.utils.aoa_to_sheet(sheet.aoa || buildSheetAoa(sheet));
    // Nom d'onglet Excel : max 31 caractères, pas de caractères interdits.
    // Une troncature changerait la clé de traduction : on l'interdit.
    if (sheet.name.length > 31) {
      throw new Error(`Nom de feuillet > 31 caractères (Excel le tronquerait) : « ${sheet.name} »`);
    }
    const tab = sheet.name.replace(/[\\/?*[\]:]/g, ' ');
    XLSX.utils.book_append_sheet(wb, ws, tab);
    flightTotal += nFlights(sheet);
  }
  XLSX.writeFile(wb, path.join(SRC_DIR, a.file));
  console.log(`✓ ${a.file}  (${a.sheets.length} feuillet(s), ${a.sheets.reduce((m, s) => m + nFlights(s), 0)} vols)`);
}

// ---------------------------------------------------------------------------
// Écriture des traductions et coordonnées additionnelles.
// ---------------------------------------------------------------------------
const sortObj = (o) => Object.fromEntries(Object.keys(o).sort().map((k) => [k, o[k]]));

const trFile = path.resolve(__dirname, 'translations-extra-2.js');
fs.writeFileSync(trFile,
  '// Traductions EN pour les 12 compagnies ajoutées (build-time). Généré par gen-new-airlines.js.\n' +
  "'use strict';\nmodule.exports = " +
  JSON.stringify({ SHEETS: sortObj(SHEETS), LINES: sortObj(LINES), FREQ: sortObj(FREQ), NOTES: sortObj(NOTES), CITY: sortObj(CITY) }, null, 2) +
  ';\n', 'utf8');
console.log(`✓ ${path.basename(trFile)}`);

const coordFile = path.resolve(__dirname, 'place-coords-2.js');
fs.writeFileSync(coordFile,
  '// Coordonnées des villes des 12 compagnies ajoutées (build-time). Généré par gen-new-airlines.js.\n' +
  "'use strict';\nmodule.exports = " + JSON.stringify(sortObj(COORDS), null, 2) + ';\n', 'utf8');
console.log(`✓ ${path.basename(coordFile)}`);

console.log(`\n${airlines.length} compagnies, ${flightTotal} vols (aller+retour) générés.`);
