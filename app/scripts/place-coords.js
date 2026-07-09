// Table de géocodage (build-time uniquement). Les coordonnées lat/lon sont
// figées dans timetables.json ; l'app finale n'a donc besoin d'aucun service.

'use strict';

// Normalise un libellé de lieu en clé de recherche :
//  "Le Caire (CAI)" -> "le caire", "Chicago MDW" -> "chicago",
//  "Gander [Tech]" -> "gander", "New York (IDL)" -> "new york".
function normalizePlace(raw) {
  let s = String(raw || '');
  s = s.replace(/\[[^\]]*\]/g, ' ');      // retire [Tech], [**], ...
  s = s.replace(/\([^)]*\)/g, ' ');        // retire (IDL), (ORY), ...
  s = s.replace(/\s+[A-Z]{3}\b/g, ' ');    // retire un code aéroport accolé (ORY, MDW, LHR)
  s = s.replace(/\s+/g, ' ').trim().toLowerCase();
  return s;
}

// Villes -> [latitude, longitude].
const COORDS = {
  'alger': [36.75, 3.06],
  'bassorah': [30.51, 47.78],
  'beyrouth': [33.89, 35.50],
  'bogota': [4.71, -74.07],
  'bombay': [19.08, 72.88],
  'brazzaville': [-4.26, 15.28],
  'buenos aires': [-34.61, -58.38],
  'calcutta': [22.57, 88.36],
  'caracas': [10.48, -66.90],
  'casablanca': [33.57, -7.59],
  'chicago': [41.85, -87.65],
  'dakar': [14.72, -17.47],
  'dayton': [39.76, -84.19],
  'denver': [39.74, -104.99],
  'dhahran': [26.29, 50.11],
  'düsseldorf': [51.23, 6.78],
  'dusseldorf': [51.23, 6.78],
  'francfort': [50.11, 8.68],
  'gander': [48.95, -54.61],
  'hambourg': [53.55, 9.99],
  'istanbul': [41.01, 28.98],
  'kano': [12.00, 8.52],
  'kansas city': [39.10, -94.58],
  'karachi': [24.86, 67.00],
  'le caire': [30.04, 31.24],
  'lisbonne': [38.72, -9.14],
  'londres': [51.51, -0.13],
  'los angeles': [34.05, -118.24],
  'madrid': [40.42, -3.70],
  'montréal': [45.50, -73.57],
  'montreal': [45.50, -73.57],
  'munich': [48.14, 11.58],
  'new york': [40.71, -74.01],
  'paris': [48.86, 2.35],
  'pittsburgh': [40.44, -79.996],
  'pointe-à-pitre': [16.24, -61.53],
  'pointe-a-pitre': [16.24, -61.53],
  'recife': [-8.05, -34.88],
  'rio de janeiro': [-22.91, -43.17],
  'rome': [41.90, 12.50],
  'san francisco': [37.77, -122.42],
  'santa maria': [36.97, -25.10],
  'saïgon': [10.82, 106.63],
  'saigon': [10.82, 106.63],
  'shannon': [52.70, -8.92],
  'st. louis': [38.63, -90.20],
  'tokyo': [35.68, 139.69],
  'téhéran': [35.69, 51.39],
  'teheran': [35.69, 51.39],
  'washington': [38.90, -77.04],
  // Ajouts : Eastern Air Lines, Flying Tiger, Irish Airlines, Pan Am.
  'atlanta': [33.75, -84.39],
  'bangkok': [13.75, 100.50],
  'guam': [13.47, 144.75],
  'berlin': [52.52, 13.40],
  'binghamton': [42.10, -75.91],
  'boston': [42.36, -71.06],
  'camagüey': [21.38, -77.92],
  'camaguey': [21.38, -77.92],
  'cleveland': [41.50, -81.69],
  'detroit': [42.33, -83.05],
  'dublin': [53.35, -6.26],
  'georgetown': [6.80, -58.16],
  'honolulu': [21.31, -157.86],
  'houston': [29.76, -95.37],
  'indianapolis': [39.77, -86.16],
  'jacksonville': [30.33, -81.66],
  'kingston': [17.97, -76.79],
  'la havane': [23.11, -82.37],
  'louisville': [38.25, -85.76],
  'manille': [14.60, 120.98],
  'miami': [25.76, -80.19],
  'montego bay': [18.47, -77.92],
  'new orleans': [29.95, -90.07],
  'okinawa': [26.21, 127.68],
  'port of spain': [10.65, -61.51],
  'san juan': [18.47, -66.11],
  'travis': [38.27, -121.93],
  'travis afb': [38.27, -121.93],
  'wake island': [19.28, 166.65],
};

function geocode(raw) {
  const key = normalizePlace(raw);
  const c = COORDS[key];
  return c ? { lat: c[0], lon: c[1] } : null;
}

module.exports = { normalizePlace, geocode, COORDS };
