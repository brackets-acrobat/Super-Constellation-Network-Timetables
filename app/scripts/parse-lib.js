// Bibliothèque de parsing partagée : transforme les feuilles Excel brutes
// en structures de données exploitables. Utilisée par compile-data.js (build)
// uniquement — l'application finale ne dépend PAS de ce fichier ni de xlsx.

'use strict';

// Découpe une cellule d'escale "Lieu (Arr. / Dép.)" en { place, arrival, departure, nextDay }.
// Exemples gérés :
//   "Shannon (08:30 / 09:30 *)"        -> place Shannon, arr 08:30, dép 09:30, nextDay(dep)
//   "Gander [Tech] (23:30 / 00:30)"    -> place "Gander [Tech]"
//   "Paris (ORY) (15:00 / 16:15)"      -> place "Paris (ORY)"
//   "Rio de Janeiro (GIG) (09:45 *)"   -> une seule heure (arrivée seule)
function parseStopCell(raw) {
  const text = String(raw || '').trim();
  if (!text || text === '-') return null;

  // Le dernier groupe entre parenthèses contenant un horaire (chiffres:chiffres)
  // représente les heures ; tout ce qui précède est le nom du lieu.
  // Un marqueur optionnel ([**], [Tech], ...) peut suivre le groupe d'horaires.
  const timeGroupRe = /\(([^()]*\d{1,2}:\d{2}[^()]*)\)\s*(\[[^\]]*\])?\s*$/;
  const m = text.match(timeGroupRe);
  if (!m) {
    // Pas d'horaire détecté -> tout est le nom du lieu.
    return { place: text, arrival: null, departure: null };
  }
  const marker = m[2] ? ' ' + m[2].trim() : '';
  const place = (text.slice(0, m.index).trim() + marker).trim();
  const inner = m[1].trim();

  const parts = inner.split('/').map((s) => s.trim());
  const clean = (t) => {
    if (t == null) return { time: null, nextDay: false, note: null };
    // Un marqueur de note ([***], [**], ...) peut être accolé à l'horaire :
    // on le retire de l'heure et on le conserve à part.
    const noteMatch = t.match(/\[[^\]]*\]/);
    const note = noteMatch ? noteMatch[0] : null;
    let s = t.replace(/\[[^\]]*\]/g, '');
    const nextDay = /\*/.test(s);
    const time = s.replace(/\*/g, '').replace(/\s+/g, ' ').trim() || null;
    return { time, nextDay, note };
  };

  if (parts.length >= 2) {
    const a = clean(parts[0]);
    const d = clean(parts[1]);
    const note = a.note || d.note || null;
    return {
      place,
      arrival: a.time,
      arrivalNextDay: a.nextDay,
      departure: d.time,
      departureNextDay: d.nextDay,
      ...(note ? { note } : {}),
    };
  }
  // Une seule heure -> considérée comme arrivée à cette escale.
  const a = clean(parts[0]);
  return {
    place, arrival: a.time, arrivalNextDay: a.nextDay, departure: null,
    ...(a.note ? { note: a.note } : {}),
  };
}

// Repère l'index des colonnes à partir de la ligne d'en-tête.
function mapColumns(header) {
  const cols = {
    flightNo: -1,
    line: -1,
    origin: -1,
    departure: -1,
    destination: -1,
    arrival: -1,
    frequency: -1,
    escales: [],
  };
  const labels = {};
  header.forEach((hRaw, i) => {
    const h = String(hRaw || '').trim();
    labels[i] = h;
    const low = h.toLowerCase();
    if (/^n°\s*vol/.test(low)) cols.flightNo = i;
    else if (low.startsWith('origine')) cols.origin = i;
    else if (low.startsWith('départ') || low.startsWith('depart')) cols.departure = i;
    else if (low.startsWith('escale')) cols.escales.push(i);
    else if (low.startsWith('destination')) cols.destination = i;
    else if (low.startsWith('arrivée') || low.startsWith('arrivee')) cols.arrival = i;
  });
  // La colonne « ligne / axe » est celle juste avant l'origine (c'est en
  // général le nom ou le type de la ligne). Les éventuelles colonnes situées
  // entre le n° de vol et cette colonne (ex. « Compagnies » pour Irish) sont
  // conservées comme descripteurs additionnels (meta).
  cols.meta = [];
  if (cols.origin > 0) {
    cols.line = cols.origin - 1;
    for (let i = cols.flightNo + 1; i < cols.line; i++) cols.meta.push(i);
  } else if (cols.flightNo !== -1) {
    cols.line = cols.flightNo + 1;
  }
  // La dernière colonne renseignée dans l'en-tête est la fréquence / notes.
  const filled = header.map((h, i) => (String(h || '').trim() ? i : -1)).filter((i) => i >= 0);
  cols.frequency = filled[filled.length - 1];
  cols.labels = labels;
  return cols;
}

// Parse une feuille (tableau de lignes) -> { flights, notes }.
function parseSheet(rows) {
  if (!rows.length) return { flights: [], notes: [] };
  const header = rows[0];
  const cols = mapColumns(header);
  const flights = [];
  const notes = [];

  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    const cell = (i) => (i >= 0 ? String(row[i] || '').trim() : '');
    const first = cell(cols.flightNo) || String(row[0] || '').trim();

    const origin = cell(cols.origin);
    const destination = cell(cols.destination);

    // Une VRAIE ligne de vol possède à la fois une origine et une destination.
    // (Plus robuste qu'un motif de n° de vol : certaines lignes n'en ont pas,
    //  ex. « MATS Pacifique », « FT-Euro Charter ».)
    const isFlight = origin && destination;
    if (!isFlight) {
      const joined = row.map((c) => String(c || '').trim()).filter(Boolean).join(' ');
      if (!joined) continue;
      // Ignore les en-têtes de section de notes (« Notes : », « Contexte … : »).
      if (/^(notes|contexte)\b.*:\s*$/i.test(joined)) continue;
      notes.push(joined);
      continue;
    }

    // Ligne de vol.
    const stops = cols.escales
      .map((ci) => parseStopCell(row[ci]))
      .filter(Boolean);

    const departure = cell(cols.departure);
    const arrival = cell(cols.arrival);

    // Itinéraire complet : origine -> escales -> destination.
    const itinerary = [];
    itinerary.push({ place: origin, arrival: null, departure: departure, role: 'origin' });
    stops.forEach((s) => itinerary.push({ ...s, role: 'stop' }));
    itinerary.push({ place: destination, arrival: arrival, departure: null, role: 'destination' });

    // Descripteurs additionnels (colonnes entre le n° de vol et la ligne).
    const meta = (cols.meta || [])
      .map((i) => ({ header: cols.labels[i] || '', value: cell(i) }))
      .filter((m) => m.value);

    flights.push({
      flightNo: first,
      line: cell(cols.line),
      lineHeader: cols.labels[cols.line] || '',
      meta,
      origin,
      departure,
      destination,
      arrival,
      frequency: cell(cols.frequency),
      frequencyHeader: cols.labels[cols.frequency] || '',
      stops,
      itinerary,
    });
  }

  return { flights, notes };
}

module.exports = { parseStopCell, mapColumns, parseSheet };
