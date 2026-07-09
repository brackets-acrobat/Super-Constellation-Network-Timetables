// Dictionnaire de traduction FR -> EN (build-time). Utilisé par compile-data.js
// pour générer l'arbre de données anglais. L'app charge le JSON compilé et ne
// dépend donc pas de ce fichier au runtime.
'use strict';

const SHEETS = {
  'Atlantique & Amériques': 'Atlantic & Americas',
  'Atlantique Nord': 'North Atlantic',
  'Axe Nord-Sud (Mainlines)': 'North–South Axis (Mainlines)',
  'Liaisons Transatlantiques': 'Transatlantic Routes',
  'Lignes Atlantique & LatAm': 'Atlantic & Latin America Lines',
  'Lignes Cargo & Charters': 'Cargo & Charter Lines',
  'Navettes & Floride': 'Shuttles & Florida',
  'Orient & Afrique': 'Orient & Africa',
  'Réseau Domestique (USA)': 'Domestic Network (USA)',
  'Réseau Inter-Allemagne (IGS)': 'Inter-German Network (IGS)',
  'Réseau International': 'International Network',
  'Sud & Moyen-Orient': 'South & Middle East',
};

const LINE_HEADERS = {
  'Axe / Destination': 'Axis / Destination',
  'Axe / Ligne': 'Axis / Line',
  'Ligne / Axe': 'Line / Axis',
  'Nom du Vol / Type': 'Flight Name / Type',
  'Secteur': 'Sector',
  'Type de Ligne': 'Line Type',
  'Type de Ligne / Nom': 'Line Type / Name',
  'Type de Vol': 'Flight Type',
};

const FREQ_HEADERS = {
  'Fréquence': 'Frequency',
  'Fréquence / Note': 'Frequency / Note',
  'Fréquence / Notes': 'Frequency / Notes',
  'Jours / Notes': 'Days / Notes',
  'Jours de Circulation': 'Days of Operation',
  'Note': 'Note',
  'Notes': 'Notes',
  'Notes / Contexte Historique': 'Notes / Historical Context',
};

const META_HEADERS = {
  'Compagnies': 'Airlines',
};

const LINES = {
  'Afrique Centrale': 'Central Africa',
  'Amérique Centrale': 'Central America',
  'Amérique Latine (Retour)': 'Latin America (Return)',
  'Amérique Latine - Navette': 'Latin America - Shuttle',
  'Anglo-Allemand': 'Anglo-German',
  'Atlantique Nord (Le Parisien)': 'North Atlantic (Le Parisien)',
  'Atlantique Nord (Retour)': 'North Atlantic (Return)',
  'Atlantique Sud': 'South Atlantic',
  'Atlantique Sud (Retour)': 'South Atlantic (Return)',
  'Caraïbes / Amérique du Sud': 'Caribbean / South America',
  'Cargo Domestique Est-Ouest': 'Domestic Cargo East–West',
  'Cargo Domestique Ouest-Est': 'Domestic Cargo West–East',
  'Charte Militaire (Transpacifique)': 'Military Charter (Transpacific)',
  'Charte Transatlantique (Saisonnier)': 'Transatlantic Charter (Seasonal)',
  'Chicago (Retour)': 'Chicago (Return)',
  'Chicago (via Canada)': 'Chicago (via Canada)',
  "Courrier d'Asie": 'Asia Mail',
  'Europe du Nord': 'Northern Europe',
  'IGS Corridor Centre': 'IGS Central Corridor',
  'IGS Corridor Centre (Rtr)': 'IGS Central Corridor (Rtn)',
  'IGS Corridor Nord': 'IGS Northern Corridor',
  'IGS Corridor Nord (Rtr)': 'IGS Northern Corridor (Rtn)',
  'IGS Corridor Sud': 'IGS Southern Corridor',
  'Le Grand Sud-Ouest': 'The Great Southwest',
  'Liaison Régionale': 'Regional Link',
  'Ligne des Épices': 'The Spice Route',
  'Ligne des Épices (Retour)': 'The Spice Route (Return)',
  'Moyen-Orient': 'Middle East',
  'Multi-Escales Atlantique': 'Multi-Stop Atlantic',
  'Multi-Escales Atlantique (Rtr)': 'Multi-Stop Atlantic (Rtn)',
  'Multi-Escales Centre': 'Multi-Stop Central',
  'Multi-Escales Nord': 'Multi-Stop North',
  'Méditerranéen': 'Mediterranean',
  'Navette Inter-Metropoles': 'Inter-City Shuttle',
  'Navette Inter-Metropoles (Rtr)': 'Inter-City Shuttle (Rtn)',
  'Navette Régionale': 'Regional Shuttle',
  'New York (Express)': 'New York (Express)',
  'New York (Retour)': 'New York (Return)',
  "Route de l'Inde": 'The India Route',
  'Sens Europe ➔ USA': 'Europe ➔ USA',
  'Sens USA ➔ Europe': 'USA ➔ Europe',
  'The Ambassador (Non-Stop)': 'The Ambassador (Non-Stop)',
  'The Flying Fisherman (Prestige)': 'The Flying Fisherman (Prestige)',
  'The Golden Falcon (Non-Stop)': 'The Golden Falcon (Non-Stop)',
  'The Golden Falcon (Retour)': 'The Golden Falcon (Return)',
  'The Sky Chief (Non-Stop)': 'The Sky Chief (Non-Stop)',
  'Transatlantique Cargo/Mixte': 'Transatlantic Cargo/Mixed',
  'Transcontinental Sud': 'Southern Transcontinental',
};

const FREQ = {
  '2x par semaine': '2× weekly',
  '3x par semaine': '3× weekly',
  '3x par semaine (Quotidien en été)': '3× weekly (Daily in summer)',
  'Axe transsaharien': 'Trans-Saharan route',
  'Charters Émigrants / Troupes': 'Emigrant / Troop charters',
  'Connexion majeure avec les vols transatlantiques': 'Major connection with transatlantic flights',
  'Connexion vers la Floride': 'Connection to Florida',
  'Dimanche': 'Sunday',
  'Escale Hong Kong incluse': 'Hong Kong stop included',
  "Fort trafic d'affaires et postal": 'Heavy business and mail traffic',
  'Horaires soumis aux fuseaux CT/ET': 'Times subject to CT/ET time zones',
  'Inauguré en 1956. Escale Rio très importante.': 'Inaugurated in 1956. Rio stop very important.',
  'Ligne Prestige (Nuit)': 'Prestige Line (Night)',
  'Ligne ouverte fin 1956. Axe pétrolier.': 'Line opened late 1956. Oil route.',
  'Lu, Ve / Ligne inaugurée en 1956': 'Mon, Fri / Line inaugurated in 1956',
  'Lundi, Jeudi': 'Monday, Thursday',
  'Lundi, Vendredi': 'Monday, Friday',
  'Ma, Je, Sa / Premier vol LH Transat': 'Tue, Thu, Sat / First LH transatlantic flight',
  'Ma, Sa': 'Tue, Sat',
  'Mardi, Jeudi, Samedi': 'Tuesday, Thursday, Saturday',
  'Mardi, Samedi': 'Tuesday, Saturday',
  'Me, Ve, Di': 'Wed, Fri, Sun',
  'Mercredi, Dimanche': 'Wednesday, Sunday',
  'Mercredi, Vendredi, Dimanche': 'Wednesday, Friday, Sunday',
  'Navette quotidienne - Corridor aérien Allied': 'Daily shuttle - Allied air corridor',
  'Opéré sous pavillon américain (accords quadripartites)': 'Operated under U.S. flag (quadripartite agreements)',
  'Quotidien': 'Daily',
  'Quotidien (Sauf Dimanche)': 'Daily (Except Sunday)',
  'Quotidien (Sauf Lundi)': 'Daily (Except Monday)',
  'Quotidien / Config Luxe': 'Daily / Luxury config',
  'Quotidien / L-1049C Super G': 'Daily / L-1049C Super G',
  'Quotidien / Vol Direct': 'Daily / Direct flight',
  'Route des Antilles': 'West Indies route',
  'Saut océanique de nuit Dakar-Rio': 'Overnight ocean crossing Dakar–Rio',
  "Très forte demande touristique d'époque": 'Very high tourist demand at the time',
  'Utilisé pour contourner le blocage terrestre de la RDA': 'Used to bypass the GDR land blockade',
  'Vendredi (Ex NY Jeudi)': 'Friday (ex NY Thursday)',
  'Vol International Court (Cuba)': 'Short international flight (Cuba)',
  'Vol direct si météo OK': 'Direct flight if weather permits',
  'Vols contractuels réguliers U.S. Army': 'Regular U.S. Army contract flights',
  'via Rio de Janeiro': 'via Rio de Janeiro',
};

const NOTES = {
  '* Arrivée le lendemain (+1 jour).': '* Arrival next day (+1 day).',
  '* Arrivée le lendemain (+1 jour). HL : Heure Locale.': '* Arrival next day (+1 day). LT: Local Time.',
  '* Arrivée le lendemain (+1 jour). [Tech] indique une escale purement technique de ravitaillement.':
    '* Arrival next day (+1 day). [Tech] indicates a purely technical refuelling stop.',
  "Eastern Air Lines appelait sa flotte de Super Constellations la 'Great Silver Fleet'.":
    "Eastern Air Lines called its Super Constellation fleet the 'Great Silver Fleet'.",
  'HL : Heure Locale. Tous les vols sont quotidiens sous réserve des conditions météorologiques en Floride.':
    'LT: Local Time. All flights are daily, subject to weather conditions in Florida.',
  "La liaison Miami - La Havane (Cuba) constituait l'une des rares routes internationales régulières opérées en gros porteur par Eastern à cette époque.":
    'The Miami – Havana (Cuba) service was one of the few regular international routes operated by Eastern with a large aircraft at the time.',
  "Le réseau IGS (Internal German Services) était exploité en raison des accords de l'immédiat après-guerre : seules les compagnies alliées (Pan Am, BEA, Air France) pouvaient desservir Berlin-Tempelhof (THF).":
    'The IGS (Internal German Services) network was operated under the immediate post-war agreements: only Allied carriers (Pan Am, BEA, Air France) could serve Berlin-Tempelhof (THF).',
  "Le vol LH 502 vers Buenos Aires s'arrête également à Rio de Janeiro (Galeão) de 09:45 à 11:15 le deuxième jour avant le tronçon final.":
    'Flight LH 502 to Buenos Aires also stops at Rio de Janeiro (Galeão) from 09:45 to 11:15 on the second day before the final leg.',
  "Note sur la flotte : Appareils exploités sous un accord de location d'avions et de navigants techniques avec Seaboard & Western Airlines.":
    'Fleet note: Aircraft operated under a wet-lease agreement (aircraft and technical crew) with Seaboard & Western Airlines.',
  "Note sur la flotte : FTL exploitait la version convertible L-1049H (Cargo / Passagers), permettant de basculer d'un réseau fret à un vol charter militaire en quelques heures.":
    'Fleet note: FTL operated the convertible L-1049H version (Cargo / Passenger), allowing it to switch from a freight network to a military charter flight within hours.',
  'Pan Am utilisait ses Super Constellations en complément de sa flotte principale de Boeing 377 Stratocruiser et Douglas DC-6B/DC-7C.':
    'Pan Am used its Super Constellations to supplement its main fleet of Boeing 377 Stratocruiser and Douglas DC-6B/DC-7C.',
  'Tous les horaires sont exprimés en heures locales (PT: Pacifique, MT: Montagnes, CT: Centre, ET: Est).':
    'All times are local (PT: Pacific, MT: Mountain, CT: Central, ET: Eastern).',
  '[***] Arrivée le surlendemain (+2 jours au total depuis la Californie).':
    '[***] Arrival two days later (+2 days total from California).',
  "[***] Le vol AF 182 fait escale à Saïgon le jeudi soir (23:50), l'équipage y est relayé et l'appareil redécolle le vendredi à 01:30. Il transite ensuite par Hong Kong (06:00/07:15) avant Tokyo.":
    '[***] Flight AF 182 stops at Saigon on Thursday evening (23:50); the crew is relieved there and the aircraft departs again on Friday at 01:30. It then transits via Hong Kong (06:00/07:15) before Tokyo.',
  "[**] Le vol AF 501 s'arrête également à Rio de Janeiro (Galeão) de 22:30 à 23:59 le deuxième jour avant d'atteindre Buenos Aires le troisième jour.":
    '[**] Flight AF 501 also stops at Rio de Janeiro (Galeão) from 22:30 to 23:59 on the second day before reaching Buenos Aires on the third day.',
  '[**] Le vol TW 49 comporte une escale supplémentaire à Phoenix (PHX) de 01:10 à 01:45 avant son arrivée finale.':
    '[**] Flight TW 49 includes an additional stop at Phoenix (PHX) from 01:10 to 01:45 before its final arrival.',
  '[**] Passage de la ligne de changement de date (Dateline) sur le tronçon Honolulu - Wake Island (+1 jour géographique).':
    '[**] Crossing of the International Date Line on the Honolulu – Wake Island leg (+1 geographic day).',
  "[Tech] Gander (Terre-Neuve) servait d'escale technique obligatoire pour le ravitaillement en carburant.":
    '[Tech] Gander (Newfoundland) served as a mandatory technical stop for refuelling.',
  '[Tech] indique une escale technique de ravitaillement obligatoire. Hambourg et Francfort servaient de hubs de départ alternés.':
    '[Tech] indicates a mandatory technical refuelling stop. Hamburg and Frankfurt served as alternating departure hubs.',
  '[Tech] indique une escale technique de ravitaillement.': '[Tech] indicates a technical refuelling stop.',
};

// Villes dont le nom français diffère de l'anglais (base, hors code aéroport).
const CITY = {
  'Alger': 'Algiers',
  'Bassorah': 'Basra',
  'Beyrouth': 'Beirut',
  'Le Caire': 'Cairo',
  'Lisbonne': 'Lisbon',
  'Londres': 'London',
  'Hambourg': 'Hamburg',
  'Francfort': 'Frankfurt',
  'Saïgon': 'Saigon',
  'Téhéran': 'Tehran',
  'La Havane': 'Havana',
  'Manille': 'Manila',
  'Montréal': 'Montreal',
};

const DAY = { Lun: 'Mon', Mar: 'Tue', Mer: 'Wed', Jeu: 'Thu', Ven: 'Fri', Sam: 'Sat', Dim: 'Sun' };

// Traduit un libellé de lieu en préservant code aéroport « (XXX) » / « XXX »
// et marqueur « [Tech] / [**] ».
function translatePlace(fr) {
  let s = String(fr || '').trim();
  let marker = '';
  const mm = s.match(/\s*(\[[^\]]*\])\s*$/);
  if (mm) { marker = ' ' + mm[1]; s = s.slice(0, mm.index).trim(); }
  let paren = '';
  const pm = s.match(/\s*(\([^)]*\))\s*$/);
  if (pm) { paren = ' ' + pm[1]; s = s.slice(0, pm.index).trim(); }
  let code = '';
  const cm = s.match(/\s+([A-Z]{3,})$/);
  if (cm) { code = ' ' + cm[1]; s = s.slice(0, cm.index).trim(); }
  const base = CITY[s] || s;
  return (base + code + paren + marker).trim();
}

// Traduit les abréviations de jour éventuellement présentes dans un horaire.
function translateTime(t) {
  if (!t) return t;
  return String(t).replace(/\((Lun|Mar|Mer|Jeu|Ven|Sam|Dim)\)/g, (_, d) => '(' + DAY[d] + ')');
}

// Recherche dans un dictionnaire ; enregistre les manques et renvoie la valeur FR par défaut.
function lookup(dict, s, misses, label) {
  const v = String(s || '');
  if (!v) return v;
  if (Object.prototype.hasOwnProperty.call(dict, v)) return dict[v];
  if (misses) misses.add(label + ': ' + JSON.stringify(v));
  return v;
}

// Fusionne les traductions additionnelles du jeu « sourced » si présentes.
try {
  const extra = require('./translations-extra');
  Object.assign(SHEETS, extra.SHEETS || {});
  Object.assign(LINE_HEADERS, extra.LINE_HEADERS || {});
  Object.assign(FREQ_HEADERS, extra.FREQ_HEADERS || {});
  Object.assign(META_HEADERS, extra.META_HEADERS || {});
  Object.assign(LINES, extra.LINES || {});
  Object.assign(FREQ, extra.FREQ || {});
  Object.assign(NOTES, extra.NOTES || {});
} catch (e) { /* pas d'extension */ }

module.exports = {
  SHEETS, LINE_HEADERS, FREQ_HEADERS, META_HEADERS, LINES, FREQ, NOTES, CITY, DAY,
  translatePlace, translateTime, lookup,
};
