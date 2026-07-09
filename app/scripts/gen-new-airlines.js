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
};

// Notes réutilisées partout (traduites une seule fois).
const LEG = "* Arrivée un jour plus tard (chaque * = +1 jour). [Tech] = escale technique. HL : Heure locale.";
const RET = "[R] Vols retour reconstitués par itinéraire miroir (horaires plausibles).";
const REC = "[R] Reconstitué : numéros de vol et horaires — non issus d'un timetable original.";
NOTES[LEG] = "* Arrival on a later day (each * = +1 day). [Tech] = technical stop. LT: local time.";
NOTES[RET] = "[R] Return flights reconstructed as mirror itineraries (plausible schedules).";
NOTES[REC] = "[R] Reconstructed: flight numbers and schedules — not from an original timetable.";

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
  return a;
}

// Aide compacte pour les notes { fr -> en }.
function N(fr, en) { NOTES[fr] = en; return fr; }

const airlines = [];

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
//  5. BUFFALO AIRWAYS  [R] — n'a JAMAIS exploité le L-1049 (illustratif)
// ===========================================================================
{
  const a = airline('buffalo', 'Buffalo Airways', 'Reseau_Buffalo_L1049_sourced.xlsx');
  const warn = N("ATTENTION : Buffalo Airways est un transporteur canadien moderne (Grand Nord) qui n'a JAMAIS exploité le L-1049. Ce feuillet est fourni pour parité avec l'appli ; il est ENTIÈREMENT RECONSTITUÉ [R] et illustratif.",
    "WARNING: Buffalo Airways is a modern Canadian northern operator that NEVER operated the L-1049. This sheet is provided for parity with the app; it is ENTIRELY RECONSTRUCTED [R] and illustrative.");
  const st = N("[R] Aucune donnée sourcée : compagnie/appareil non historiques pour le L-1049.",
    "[R] No sourced data: airline/aircraft not historical for the L-1049.");
  const sh = a.sheet('Grand Nord (illustratif)', 'Far North (illustrative)', [warn, st, LEG, RET]);
  sh.route({
    fwdNo: 'BF 101', retNo: 'BF 102',
    lineFwd: ['Liaison du Grand Nord', 'Far North Link'],
    lineRet: ['Liaison du Grand Nord (Retour)', 'Far North Link (Return)'],
    ports: ['Edmonton (YEG)', 'Yellowknife', 'Churchill'],
    legMin: [150, 180], groundMin: [60], outDep: '09:00', retDep: '14:00',
    freqFwd: ['Lundi, Jeudi', 'Monday, Thursday'],
    freqRet: ['Mardi, Vendredi', 'Tuesday, Friday'],
  });
  sh.route({
    fwdNo: 'BF 103', retNo: 'BF 104',
    lineFwd: ['Navette Arctique', 'Arctic Shuttle'],
    lineRet: ['Navette Arctique (Retour)', 'Arctic Shuttle (Return)'],
    ports: ['Yellowknife', 'Churchill', 'Frobisher Bay'],
    legMin: [180, 300], groundMin: [60], outDep: '10:00', retDep: '13:00',
    freqFwd: ['Mercredi', 'Wednesday'],
    freqRet: ['Samedi', 'Saturday'],
  });
  airlines.push(a);
}

// ===========================================================================
//  6. NORDAIR  [R] — DC-3/DC-4/Electra, PAS de L-1049 (illustratif)
// ===========================================================================
{
  const a = airline('nordair', 'Nordair', 'Reseau_Nordair_L1049_sourced.xlsx');
  const warn = N("ATTENTION : Nordair (régional québécois, fondé en 1957) a exploité DC-3/DC-4 puis Electra et 737 — PAS le L-1049. Feuillet illustratif, ENTIÈREMENT RECONSTITUÉ [R].",
    "WARNING: Nordair (a Quebec regional carrier founded in 1957) operated DC-3/DC-4 then the Electra and 737 — NOT the L-1049. Illustrative sheet, ENTIRELY RECONSTRUCTED [R].");
  const st = N("[R] Aucune donnée sourcée pour le L-1049 : réseau nordique reconstitué à titre illustratif.",
    "[R] No sourced L-1049 data: northern network reconstructed for illustration only.");
  const sh = a.sheet('Québec & Arctique (illustratif)', 'Quebec & Arctic (illustrative)', [warn, st, LEG, RET]);
  sh.route({
    fwdNo: 'ND 201', retNo: 'ND 202',
    lineFwd: ["Liaison de l'Ungava", 'Ungava Link'],
    lineRet: ["Liaison de l'Ungava (Retour)", 'Ungava Link (Return)'],
    ports: ['Montréal (YUL)', "Val-d'Or", 'Frobisher Bay'],
    legMin: [90, 300], groundMin: [45], outDep: '07:30', retDep: '13:00',
    freqFwd: ['Mardi, Vendredi', 'Tuesday, Friday'],
    freqRet: ['Mercredi, Samedi', 'Wednesday, Saturday'],
  });
  sh.route({
    fwdNo: 'ND 203', retNo: 'ND 204',
    lineFwd: ['Navette de la Baie', 'Bay Shuttle'],
    lineRet: ['Navette de la Baie (Retour)', 'Bay Shuttle (Return)'],
    ports: ['Montréal (YUL)', 'Churchill'],
    legMin: [300], groundMin: [], outDep: '08:00', retDep: '15:00',
    freqFwd: ['Jeudi', 'Thursday'],
    freqRet: ['Dimanche', 'Sunday'],
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

// ===========================================================================
// 11. JAT YUGOSLAV AIRLINES  [R] — Convair/DC-6B/Caravelle (illustratif)
// ===========================================================================
{
  const a = airline('jat', 'JAT Yugoslav Airlines', 'Reseau_JAT_L1049_sourced.xlsx');
  const warn = N("ATTENTION : la JAT Yugoslav Airlines n'a PAS exploité le L-1049 (flotte Convair 340/440, DC-6B puis Caravelle). Feuillet illustratif, ENTIÈREMENT RECONSTITUÉ [R].",
    "WARNING: JAT Yugoslav Airlines did NOT operate the L-1049 (fleet of Convair 340/440, DC-6B then Caravelle). Illustrative sheet, ENTIRELY RECONSTRUCTED [R].");
  const st = N("[S] Sourcé : uniquement les villes réelles du réseau JAT — pas l'appareil.",
    "[S] Sourced: only the real cities of the JAT network — not the aircraft.");
  const sh = a.sheet('Réseau Européen (illustratif)', 'European Network (illustrative)', [warn, st, LEG, RET]);
  sh.route({
    fwdNo: 'JU 350', retNo: 'JU 351',
    lineFwd: ["Ligne de l'Ouest", 'Western Line'],
    lineRet: ["Ligne de l'Ouest (Retour)", 'Western Line (Return)'],
    ports: ['Belgrade (BEG)', 'Zagreb', 'Zürich', 'Paris (ORY)'],
    legMin: [45, 90, 120], groundMin: [30, 45], outDep: '08:00', retDep: '17:00',
    freqFwd: ['Lundi, Mercredi, Vendredi', 'Monday, Wednesday, Friday'],
    freqRet: ['Mardi, Jeudi, Samedi', 'Tuesday, Thursday, Saturday'],
  });
  sh.route({
    fwdNo: 'JU 450', retNo: 'JU 451',
    lineFwd: ['Ligne du Levant', 'Levant Line'],
    lineRet: ['Ligne du Levant (Retour)', 'Levant Line (Return)'],
    ports: ['Belgrade (BEG)', 'Athènes', 'Le Caire (CAI)'],
    legMin: [120, 180], groundMin: [60], outDep: '09:00', retDep: '15:00',
    freqFwd: ['Mardi, Samedi', 'Tuesday, Saturday'],
    freqRet: ['Mercredi, Dimanche', 'Wednesday, Sunday'],
  });
  airlines.push(a);
}

// ===========================================================================
// 12. SOUTH AFRICAN AIRWAYS (SAA)  [R] — Springbok en L-749/DC-7B (illustratif)
// ===========================================================================
{
  const a = airline('saa', 'South African Airways', 'Reseau_South_African_L1049_sourced.xlsx');
  const warn = N("ATTENTION : le « Springbok Service » de la SAA (Johannesburg–Londres) était assuré en L-749 Constellation puis DC-7B ; l'exploitation du L-1049 n'est pas documentée. Données [R] reconstituées.",
    "WARNING: SAA's 'Springbok Service' (Johannesburg–London) was flown with the L-749 Constellation then the DC-7B; L-1049 operation is not documented. Data is [R] reconstructed.");
  const st = N("[S] Sourcé : uniquement l'itinéraire Springbok réel (via l'Afrique de l'Est) — pas l'appareil L-1049.",
    "[S] Sourced: only the real Springbok routing (via East Africa) — not the L-1049 aircraft.");
  a.sheet('Springbok Service (illustratif)', 'Springbok Service (illustrative)', [warn, st, LEG, RET]).route({
    fwdNo: 'SA 234', retNo: 'SA 235',
    lineFwd: ['Springbok Service (Londres)', 'Springbok Service (London)'],
    lineRet: ['Springbok Service (Retour)', 'Springbok Service (Return)'],
    ports: ['Johannesburg (JNB)', 'Nairobi', 'Khartoum', 'Le Caire (CAI)', 'Rome', 'Londres (LHR)'],
    legMin: [360, 240, 180, 210, 150], groundMin: [75, 60, 60, 75], outDep: '18:00', retDep: '11:00',
    freqFwd: ['Mardi, Jeudi, Samedi', 'Tuesday, Thursday, Saturday'],
    freqRet: ['Mercredi, Vendredi, Dimanche', 'Wednesday, Friday, Sunday'],
  });
  airlines.push(a);
}

// ===========================================================================
// ====================  2e LOT : 11 compagnies supplémentaires  =============
// ===========================================================================

// 13. AMERICAN OVERSEAS AIRLINES  [R] — fusionnée dans Pan Am en 1950
{
  const a = airline('american-overseas', 'American Overseas Airlines', 'Reseau_American_Overseas_L1049_sourced.xlsx');
  const warn = N("ATTENTION : American Overseas Airlines a fusionné avec Pan Am en septembre 1950, AVANT l'entrée en service du L-1049 (déc. 1951) ; elle exploitait des Constellation L-049/L-749. Feuillet illustratif, ENTIÈREMENT RECONSTITUÉ [R].",
    "WARNING: American Overseas Airlines merged into Pan Am in September 1950, BEFORE the L-1049 entered service (Dec. 1951); it operated L-049/L-749 Constellations. Illustrative sheet, ENTIRELY RECONSTRUCTED [R].");
  const st = N("[S] Sourcé : uniquement les corridors nord-atlantiques réels d'AOA — pas l'appareil L-1049.",
    "[S] Sourced: only AOA's real North Atlantic corridors — not the L-1049 aircraft.");
  const sh = a.sheet('Atlantique Nord (illustratif)', 'North Atlantic (illustrative)', [warn, st, LEG, RET]);
  sh.route({
    fwdNo: 'AO 101', retNo: 'AO 102',
    lineFwd: ['Atlantique Nord (Flagship)', 'North Atlantic (Flagship)'],
    lineRet: ['Atlantique Nord (Retour)', 'North Atlantic (Return)'],
    ports: ['New York (IDL)', 'Gander [Tech]', 'Shannon', 'Londres (LHR)'],
    legMin: [150, 450, 120], groundMin: [60, 60], outDep: '18:00', retDep: '12:00',
    freqFwd: ['Mardi, Jeudi, Samedi', 'Tuesday, Thursday, Saturday'],
    freqRet: ['Mercredi, Vendredi, Dimanche', 'Wednesday, Friday, Sunday'],
  });
  sh.route({
    fwdNo: 'AO 151', retNo: 'AO 152',
    lineFwd: ['Route de Scandinavie', 'Scandinavian Route'],
    lineRet: ['Route de Scandinavie (Retour)', 'Scandinavian Route (Return)'],
    ports: ['New York (IDL)', 'Gander [Tech]', 'Prestwick', 'Copenhague'],
    legMin: [150, 420, 150], groundMin: [60, 60], outDep: '17:00', retDep: '10:00',
    freqFwd: ['Lundi, Vendredi', 'Monday, Friday'],
    freqRet: ['Mercredi, Dimanche', 'Wednesday, Sunday'],
  });
  airlines.push(a);
}

// 14. SWISSAIR  [R] — cliente Douglas (DC-6B/DC-7C)
{
  const a = airline('swissair', 'Swissair', 'Reseau_Swissair_L1049_sourced.xlsx');
  const warn = N("ATTENTION : la Swissair était une cliente Douglas (DC-6B puis DC-7C) et n'a PAS exploité le L-1049. Feuillet illustratif, ENTIÈREMENT RECONSTITUÉ [R].",
    "WARNING: Swissair was a Douglas customer (DC-6B then DC-7C) and did NOT operate the L-1049. Illustrative sheet, ENTIRELY RECONSTRUCTED [R].");
  const st = N("[S] Sourcé : uniquement les villes/corridors réels du réseau Swissair — pas l'appareil.",
    "[S] Sourced: only the real cities/corridors of the Swissair network — not the aircraft.");
  const sh = a.sheet('International (illustratif)', 'International Network (illustrative)', [warn, st, LEG, RET]);
  sh.route({
    fwdNo: 'SR 100', retNo: 'SR 101',
    lineFwd: ['Atlantique Nord (Genève)', 'North Atlantic (Geneva)'],
    lineRet: ['Atlantique Nord (Retour)', 'North Atlantic (Return)'],
    ports: ['Zürich (ZRH)', 'Genève', 'Shannon', 'New York (IDL)'],
    legMin: [60, 420, 300], groundMin: [45, 60], outDep: '09:00', retDep: '16:00',
    freqFwd: ['Mardi, Vendredi', 'Tuesday, Friday'],
    freqRet: ['Jeudi, Dimanche', 'Thursday, Sunday'],
  });
  sh.route({
    fwdNo: 'SR 320', retNo: 'SR 321',
    lineFwd: ['Ligne du Levant', 'Levant Line'],
    lineRet: ['Ligne du Levant (Retour)', 'Levant Line (Return)'],
    ports: ['Zürich (ZRH)', 'Rome', 'Athènes', 'Le Caire (CAI)'],
    legMin: [90, 120, 180], groundMin: [60, 60], outDep: '08:00', retDep: '14:00',
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

// 16. FLY ALASKA / Alaska Airlines  [R] — L-749 (MATS), L-1049 non documenté
{
  const a = airline('alaska', 'Alaska Airlines', 'Reseau_Alaska_L1049_sourced.xlsx');
  const warn = N("ATTENTION : Alaska Airlines a exploité des Constellation L-749 (charters militaires MATS) mais l'usage du L-1049 n'est pas documenté. Feuillet illustratif, données [R] reconstituées.",
    "WARNING: Alaska Airlines operated L-749 Constellations (MATS military charters) but L-1049 use is not documented. Illustrative sheet, [R] reconstructed data.");
  const st = N("[S] Sourcé : uniquement les liaisons réelles du réseau alaskien — pas l'appareil L-1049.",
    "[S] Sourced: only the real Alaska network links — not the L-1049 aircraft.");
  const sh = a.sheet("Alaska (illustratif)", 'Alaska Network (illustrative)', [warn, st, LEG, RET]);
  sh.route({
    fwdNo: 'AS 60', retNo: 'AS 61',
    lineFwd: ['Ligne du Grand Nord', 'Great North Line'],
    lineRet: ['Ligne du Grand Nord (Retour)', 'Great North Line (Return)'],
    ports: ['Seattle (SEA)', 'Anchorage', 'Fairbanks'],
    legMin: [240, 90], groundMin: [60], outDep: '08:00', retDep: '14:00',
    freqFwd: ['Quotidien', 'Daily'],
    freqRet: ['Quotidien', 'Daily'],
  });
  sh.route({
    fwdNo: 'AS 80', retNo: 'AS 81',
    lineFwd: ["Route de l'Ouest Alaskien", 'Western Alaska Route'],
    lineRet: ["Route de l'Ouest Alaskien (Retour)", 'Western Alaska Route (Return)'],
    ports: ['Anchorage', 'Nome'],
    legMin: [120], groundMin: [], outDep: '10:00', retDep: '13:00',
    freqFwd: ['Lundi, Jeudi', 'Monday, Thursday'],
    freqRet: ['Mardi, Vendredi', 'Tuesday, Friday'],
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

// 21. UNITED AIR LINES  [R] — cliente Douglas (DC-6/DC-7 Mainliner)
{
  const a = airline('united', 'United Air Lines', 'Reseau_United_L1049_sourced.xlsx');
  const warn = N("ATTENTION : United Air Lines était une grande cliente Douglas (DC-6/DC-7 « Mainliner ») et n'a PAS exploité le L-1049. Feuillet illustratif, ENTIÈREMENT RECONSTITUÉ [R].",
    "WARNING: United Air Lines was a major Douglas customer (DC-6/DC-7 'Mainliner') and did NOT operate the L-1049. Illustrative sheet, ENTIRELY RECONSTRUCTED [R].");
  const st = N("[S] Sourcé : uniquement les corridors réels du réseau United — pas l'appareil.",
    "[S] Sourced: only the real corridors of the United network — not the aircraft.");
  const sh = a.sheet('Transcontinental (illustratif)', 'Transcontinental (illustrative)', [warn, st, LEG, RET]);
  sh.route({
    fwdNo: 'UA 1', retNo: 'UA 2',
    lineFwd: ['Transcontinental (Mainliner)', 'Transcontinental (Mainliner)'],
    lineRet: ['Transcontinental (Retour)', 'Transcontinental (Return)'],
    ports: ['San Francisco (SFO)', 'Denver', 'Chicago', 'New York (IDL)'],
    legMin: [180, 150, 150], groundMin: [45, 45], outDep: '08:00', retDep: '09:00',
    freqFwd: ['Quotidien', 'Daily'],
    freqRet: ['Quotidien', 'Daily'],
  });
  sh.route({
    fwdNo: 'UA 5', retNo: 'UA 6',
    lineFwd: ['Côte Ouest–Hawaï', 'West Coast–Hawaii'],
    lineRet: ['Côte Ouest–Hawaï (Retour)', 'West Coast–Hawaii (Return)'],
    ports: ['San Francisco (SFO)', 'Honolulu'],
    legMin: [540], groundMin: [], outDep: '10:00', retDep: '16:00',
    freqFwd: ['Mardi, Jeudi, Samedi', 'Tuesday, Thursday, Saturday'],
    freqRet: ['Mercredi, Vendredi, Dimanche', 'Wednesday, Friday, Sunday'],
  });
  airlines.push(a);
}

// 22. LUXAIR  [R] — fondée en 1962 (F27/Viscount/Caravelle)
{
  const a = airline('luxair', 'Luxair', 'Reseau_Luxair_L1049_sourced.xlsx');
  const warn = N("ATTENTION : la Luxair a été fondée en 1962 (Fokker F27, Viscount, Caravelle) — bien après l'ère du L-1049, qu'elle n'a jamais exploité. Feuillet illustratif, ENTIÈREMENT RECONSTITUÉ [R].",
    "WARNING: Luxair was founded in 1962 (Fokker F27, Viscount, Caravelle) — well after the L-1049 era, which it never operated. Illustrative sheet, ENTIRELY RECONSTRUCTED [R].");
  const st = N("[S] Sourcé : uniquement les villes réelles desservies par Luxair — pas l'appareil ni l'époque.",
    "[S] Sourced: only the real cities served by Luxair — not the aircraft or the period.");
  const sh = a.sheet('Réseau Européen (illustratif)', 'European Network (illustrative)', [warn, st, LEG, RET]);
  sh.route({
    fwdNo: 'LG 201', retNo: 'LG 202',
    lineFwd: ['Ligne de Paris', 'Paris Line'],
    lineRet: ['Ligne de Paris (Retour)', 'Paris Line (Return)'],
    ports: ['Luxembourg (LUX)', 'Paris (ORY)', 'Nice'],
    legMin: [75, 90], groundMin: [45], outDep: '08:00', retDep: '17:00',
    freqFwd: ['Quotidien', 'Daily'],
    freqRet: ['Quotidien', 'Daily'],
  });
  sh.route({
    fwdNo: 'LG 301', retNo: 'LG 302',
    lineFwd: ['Ligne du Sud', 'Southern Line'],
    lineRet: ['Ligne du Sud (Retour)', 'Southern Line (Return)'],
    ports: ['Luxembourg (LUX)', 'Munich', 'Rome'],
    legMin: [90, 120], groundMin: [45], outDep: '09:00', retDep: '15:00',
    freqFwd: ['Mardi, Jeudi, Samedi', 'Tuesday, Thursday, Saturday'],
    freqRet: ['Mercredi, Vendredi, Dimanche', 'Wednesday, Friday, Sunday'],
  });
  airlines.push(a);
}

// 23. DELTA AIR LINES  [R] — cliente Douglas (DC-6/DC-7)
{
  const a = airline('delta', 'Delta Air Lines', 'Reseau_Delta_L1049_sourced.xlsx');
  const warn = N("ATTENTION : Delta Air Lines était une cliente Douglas (DC-6/DC-7) et n'a PAS exploité le L-1049. Feuillet illustratif, ENTIÈREMENT RECONSTITUÉ [R].",
    "WARNING: Delta Air Lines was a Douglas customer (DC-6/DC-7) and did NOT operate the L-1049. Illustrative sheet, ENTIRELY RECONSTRUCTED [R].");
  const st = N("[S] Sourcé : uniquement les corridors réels du réseau Delta — pas l'appareil.",
    "[S] Sourced: only the real corridors of the Delta network — not the aircraft.");
  const sh = a.sheet('Réseau du Sud (illustratif)', 'Southern Network (illustrative)', [warn, st, LEG, RET]);
  sh.route({
    fwdNo: 'DL 41', retNo: 'DL 42',
    lineFwd: ['Transcontinental Sud', 'Southern Transcontinental'],
    lineRet: ['Transcontinental Sud (Retour)', 'Southern Transcontinental (Return)'],
    ports: ['Atlanta (ATL)', 'Dallas', 'Los Angeles'],
    legMin: [150, 180], groundMin: [45], outDep: '08:00', retDep: '15:00',
    freqFwd: ['Quotidien', 'Daily'],
    freqRet: ['Quotidien', 'Daily'],
  });
  sh.route({
    fwdNo: 'DL 61', retNo: 'DL 62',
    lineFwd: ['Ligne de Floride', 'Florida Line'],
    lineRet: ['Ligne de Floride (Retour)', 'Florida Line (Return)'],
    ports: ['Chicago', 'Atlanta (ATL)', 'Miami'],
    legMin: [120, 120], groundMin: [45], outDep: '09:00', retDep: '14:00',
    freqFwd: ['Mardi, Jeudi, Samedi', 'Tuesday, Thursday, Saturday'],
    freqRet: ['Mercredi, Vendredi, Dimanche', 'Wednesday, Friday, Sunday'],
  });
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

let flightTotal = 0;
for (const a of airlines) {
  const wb = XLSX.utils.book_new();
  for (const sheet of a.sheets) {
    const ws = XLSX.utils.aoa_to_sheet(buildSheetAoa(sheet));
    // Nom d'onglet Excel : max 31 caractères, pas de caractères interdits.
    // Une troncature changerait la clé de traduction : on l'interdit.
    if (sheet.name.length > 31) {
      throw new Error(`Nom de feuillet > 31 caractères (Excel le tronquerait) : « ${sheet.name} »`);
    }
    const tab = sheet.name.replace(/[\\/?*[\]:]/g, ' ');
    XLSX.utils.book_append_sheet(wb, ws, tab);
    flightTotal += sheet.flights.length;
  }
  XLSX.writeFile(wb, path.join(SRC_DIR, a.file));
  console.log(`✓ ${a.file}  (${a.sheets.length} feuillet(s), ${a.sheets.reduce((m, s) => m + s.flights.length, 0)} vols)`);
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
