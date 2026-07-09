'use strict';

const DATA = window.timetables.getData();
const WORLD = window.timetables.getMap();
let currentFlight = null;

// ---------- Internationalisation (interface) ----------
const I18N = {
  en: {
    title: 'Super Constellation — L-1049 Timetables',
    brandTitle: 'Super Constellation Network Timetables',
    brandEra: 'For the Red Wing Lockheed Super Constellation',
    lblAirline: '1 · Airline',
    lblSheet: '2 · Network / Sheet',
    lblFlight: '3 · Flight / Line',
    phAirline: '— Select an airline —',
    phSheet: '— Select a network —',
    phFlight: '— Select a flight —',
    placeholder: 'Select an airline, a network and a flight to display the route details.',
    mapTitle: 'ROUTE MAP',
    mapUnavailable: 'Map unavailable.',
    timelineTitle: 'Itinerary & schedule (local times)',
    roleOrigin: 'Origin', roleDest: 'Destination', roleStop: 'Stop', roleTech: 'Technical stop',
    arr: 'Arr.', dep: 'Dep.', nextday: '+1 d',
    stopsLabel: (n) => `${n} ${n === 1 ? 'stop' : 'stops'}`,
    notesTitle: (name) => `Network notes “${name}”`,
    footer: (aircraft, date) => `${aircraft} · data compiled on ${date}`,
    dateLocale: 'en-GB',
  },
  fr: {
    title: 'Super Constellation — Horaires L-1049',
    brandTitle: 'Super Constellation Network Timetables',
    brandEra: 'Pour le Red Wing Lockheed Super Constellation',
    lblAirline: '1 · Compagnie',
    lblSheet: '2 · Réseau / Feuillet',
    lblFlight: '3 · Vol / Ligne',
    phAirline: '— Choisir une compagnie —',
    phSheet: '— Choisir un réseau —',
    phFlight: '— Choisir un vol —',
    placeholder: 'Sélectionnez une compagnie, un réseau puis un vol pour afficher le détail de la ligne.',
    mapTitle: 'CARTE DE LA LIGNE',
    mapUnavailable: 'Carte indisponible.',
    timelineTitle: 'Itinéraire & horaires (heures locales)',
    roleOrigin: 'Origine', roleDest: 'Destination', roleStop: 'Escale', roleTech: 'Escale technique',
    arr: 'Arr.', dep: 'Dép.', nextday: '+1 j',
    stopsLabel: (n) => `${n} escale${n > 1 ? 's' : ''}`,
    notesTitle: (name) => `Notes du réseau « ${name} »`,
    footer: (aircraft, date) => `${aircraft} · données compilées le ${date}`,
    dateLocale: 'fr-FR',
  },
};
let LANG = 'en';                                   // anglais par défaut à l'ouverture
const T = () => I18N[LANG];
let AIRLINES = DATA.airlinesEn;                    // arbre de données courant (par langue)

const els = {
  airline: document.getElementById('sel-airline'),
  sheet: document.getElementById('sel-sheet'),
  flight: document.getElementById('sel-flight'),
  detail: document.getElementById('detail'),
  footerMeta: document.getElementById('footer-meta'),
};

// ---------- Utilitaires ----------
function opt(value, label) {
  const o = document.createElement('option');
  o.value = value;
  o.textContent = label;
  return o;
}

function placeholderOption(text) {
  const o = opt('', text);
  o.disabled = true;
  o.selected = true;
  return o;
}

function findAirline(id) { return AIRLINES.find((a) => a.id === id); }
function findSheet(airline, id) { return airline && airline.sheets.find((s) => s.id === id); }
function findFlight(sheet, id) { return sheet && sheet.flights.find((f) => f.id === id); }

function esc(s) {
  return String(s ?? '').replace(/[&<>"']/g, (c) =>
    ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

const isTech = (place) => /\[tech\]/i.test(place || '');
const shortName = (s) => String(s || '').replace(/\s*\[.*\]/, '').replace(/\s*\(.*\)/, '');

// ---------- Remplissage des menus ----------
function resetSelect(select, placeholder) {
  select.innerHTML = '';
  select.appendChild(placeholderOption(placeholder));
  select.disabled = true;
}

function flightOptionLabel(f) {
  return `${f.flightNo} — ${f.line}  ·  ${shortName(f.origin)} → ${shortName(f.destination)}`;
}

function fillAirlines() {
  els.airline.innerHTML = '';
  els.airline.appendChild(placeholderOption(T().phAirline));
  AIRLINES.forEach((a) => els.airline.appendChild(opt(a.id, a.name)));
}

function fillSheets(airline) {
  resetSelect(els.sheet, T().phSheet);
  if (!airline) return;
  airline.sheets.forEach((s) => els.sheet.appendChild(opt(s.id, s.name)));
  els.sheet.disabled = false;
}

function fillFlights(sheet) {
  resetSelect(els.flight, T().phFlight);
  if (!sheet) return;
  sheet.flights.forEach((f) => els.flight.appendChild(opt(f.id, flightOptionLabel(f))));
  els.flight.disabled = false;
}

// ---------- Cascade ----------
els.airline.addEventListener('change', () => {
  const airline = findAirline(els.airline.value);
  fillSheets(airline);
  resetSelect(els.flight, T().phFlight);
  showPlaceholder();
});

els.sheet.addEventListener('change', () => {
  const airline = findAirline(els.airline.value);
  const sheet = findSheet(airline, els.sheet.value);
  fillFlights(sheet);
  showPlaceholder();
});

els.flight.addEventListener('change', () => {
  const airline = findAirline(els.airline.value);
  const sheet = findSheet(airline, els.sheet.value);
  const flight = findFlight(sheet, els.flight.value);
  if (flight) renderFlight(airline, sheet, flight);
});

// ---------- Textes statiques & langue ----------
function setText(id, txt) { const e = document.getElementById(id); if (e) e.textContent = txt; }

function updateFooter() {
  const t = T();
  els.footerMeta.textContent =
    t.footer(DATA.aircraft, new Date(DATA.generatedAt).toLocaleDateString(t.dateLocale));
}

function applyStaticI18n() {
  const t = T();
  document.title = t.title;
  setText('brand-title', t.brandTitle);
  setText('brand-era', t.brandEra);
  setText('lbl-airline', t.lblAirline);
  setText('lbl-sheet', t.lblSheet);
  setText('lbl-flight', t.lblFlight);
  setText('placeholder-msg', t.placeholder);
  updateFooter();
  document.querySelectorAll('.lang-btn').forEach((b) =>
    b.classList.toggle('active', b.dataset.lang === LANG));
}

// Change la langue en conservant la sélection courante (les identifiants sont
// communs aux deux arbres de données).
function setLang(lang) {
  if (!I18N[lang]) return;
  LANG = lang;
  AIRLINES = lang === 'en' ? DATA.airlinesEn : DATA.airlines;
  document.documentElement.lang = lang;
  applyStaticI18n();

  const aId = els.airline.value, sId = els.sheet.value, fId = els.flight.value;
  fillAirlines();
  els.airline.value = aId;
  const airline = findAirline(aId);
  fillSheets(airline);
  if (airline && sId) els.sheet.value = sId;
  const sheet = findSheet(airline, sId);
  fillFlights(sheet);
  if (sheet && fId) els.flight.value = fId;
  const flight = findFlight(sheet, fId);
  if (flight) renderFlight(airline, sheet, flight);
  else showPlaceholder();
}

// ---------- Rendu ----------
function showPlaceholder() {
  currentFlight = null;
  els.detail.innerHTML =
    `<div class="placeholder"><div class="placeholder-star">✦</div>
     <p>${esc(T().placeholder)}</p></div>`;
}

function timeCell(label, time, nextDay) {
  if (!time) return '';
  return `<span><span class="t-label">${esc(label)}</span> ${esc(time)}${
    nextDay ? ` <span class="nextday">${esc(T().nextday)}</span>` : ''}</span>`;
}

function renderTimeline(flight) {
  const items = flight.itinerary.map((s) => {
    const roleClass =
      s.role === 'origin' ? 'origin' :
      s.role === 'destination' ? 'destination' :
      'stop' + (isTech(s.place) ? ' tech' : '');
    const roleLabel =
      s.role === 'origin' ? T().roleOrigin :
      s.role === 'destination' ? T().roleDest :
      (isTech(s.place) ? T().roleTech : T().roleStop);
    const times = [
      timeCell(T().arr, s.arrival, s.arrivalNextDay),
      timeCell(T().dep, s.departure, s.departureNextDay),
    ].filter(Boolean).join('');
    return `<li class="tl-item ${roleClass}">
      <span class="tl-dot"></span>
      <div class="tl-place">${esc(s.place)}${
        s.note ? ` <span class="tl-note">${esc(s.note)}</span>` : ''}<span class="tl-role">${roleLabel}</span></div>
      ${times ? `<div class="tl-times">${times}</div>` : ''}
    </li>`;
  }).join('');
  return `<ul class="timeline">${items}</ul>`;
}

// --- Vraie carte vintage : géométrie mondiale sépia, zoomée sur le vol ---
// Projection équirectangulaire identique à celle de worldmap.json :
//   x = (lon + 180) * pxPerDeg ; y = (90 - lat) * pxPerDeg.
function projX(lon) { return (lon + 180) * WORLD.pxPerDeg; }
function projY(lat) { return (90 - lat) * WORLD.pxPerDeg; }

// boxW = largeur réelle du panneau (px), boxH = hauteur voulue (px, ici 900).
function renderMapSVG(flight, boxW, boxH) {
  const pts = flight.itinerary.filter((s) => s.lat != null && s.lon != null);
  if (pts.length < 2) return `<p style="color:var(--ink-soft);font-size:13px">${esc(T().mapUnavailable)}</p>`;

  const W = WORLD.width;
  const xy = pts.map((p) => ({ x: projX(p.lon), y: projY(p.lat), ref: p }));

  // « Déroulage » de l'antiméridien : rend la suite des longitudes continue
  // pour qu'une route transpacifique (ex. Honolulu → Wake → Tokyo) soit tracée
  // du bon côté au lieu de faire le tour du globe.
  for (let i = 1; i < xy.length; i++) {
    while (xy[i].x - xy[i - 1].x > W / 2) xy[i].x -= W;
    while (xy[i].x - xy[i - 1].x < -W / 2) xy[i].x += W;
  }

  // Cadre de vue : englobe toutes les escales, avec une marge proportionnelle.
  let minX = Math.min(...xy.map((p) => p.x));
  let maxX = Math.max(...xy.map((p) => p.x));
  let minY = Math.min(...xy.map((p) => p.y));
  let maxY = Math.max(...xy.map((p) => p.y));
  let vw = Math.max(maxX - minX, 1);
  let vh = Math.max(maxY - minY, 1);

  // Marge : 16 % autour + un minimum -> tracé zoomé, villes bien dégagées.
  const margin = Math.max(vw, vh) * 0.16 + 10;
  minX -= margin; maxX += margin; minY -= margin; maxY += margin;

  // Ajuste le cadre au ratio EXACT du panneau (largeur/hauteur) : pas de
  // déformation ni de bandes vides.
  const targetRatio = boxW / boxH;
  vw = maxX - minX; vh = maxY - minY;
  if (vw / vh < targetRatio) {
    const nw = vh * targetRatio; const dx = (nw - vw) / 2;
    minX -= dx; maxX += dx;
  } else {
    const nh = vw / targetRatio; const dy = (nh - vh) / 2;
    minY -= dy; maxY += dy;
  }
  vw = maxX - minX; vh = maxY - minY;

  // Échelle px→unités : permet des tailles constantes exprimées en pixels écran.
  const scale = boxH / vh;            // pixels par unité de projection
  const U = (px) => px / scale;       // pixels écran -> unités du viewBox

  const rEnd = U(4.5), rStop = U(3);
  const strokeRoute = U(1.6);
  const coast = U(0.6), border = U(0.45);
  const fontPx = 9.5;                 // taille de police à l'écran (petite)
  const font = U(fontPx);
  const mapTitle = T().mapTitle;

  // Route : arcs légèrement bombés (esthétique carte de ligne aérienne).
  let route = `M ${xy[0].x.toFixed(1)} ${xy[0].y.toFixed(1)}`;
  for (let i = 1; i < xy.length; i++) {
    const a = xy[i - 1], b = xy[i];
    const mx = (a.x + b.x) / 2;
    const my = (a.y + b.y) / 2 - Math.hypot(b.x - a.x, b.y - a.y) * 0.14;
    route += ` Q ${mx.toFixed(1)} ${my.toFixed(1)} ${b.x.toFixed(1)} ${b.y.toFixed(1)}`;
  }

  const gap = U(3);
  const markers = xy.map((p, i) => {
    const isEnd = i === 0 || i === xy.length - 1;
    const tech = isTech(p.ref.place);
    const r = isEnd ? rEnd : rStop;
    const fill = isEnd ? 'var(--red)' : (tech ? '#f7f0dd' : 'var(--gold)');
    const stroke = isEnd ? 'var(--red-deep)' : 'var(--gold)';
    const name = esc(p.ref.place.replace(/\s*\[.*\]/, '').replace(/\s*\(.*\)/, ''));
    // Libellé alterné au-dessus / en-dessous du point : évite de croiser la
    // route (globalement horizontale) et le chevauchement entre voisins.
    const above = i % 2 === 0;
    const ty = above ? p.y - r - gap : p.y + r + gap + font * 0.8;
    return `<g>
      <circle cx="${p.x.toFixed(1)}" cy="${p.y.toFixed(1)}" r="${r.toFixed(2)}"
              fill="${fill}" stroke="${stroke}" stroke-width="${(coast*1.4).toFixed(2)}"/>
      <text x="${p.x.toFixed(1)}" y="${ty.toFixed(1)}" text-anchor="middle"
            class="m-label" style="font-size:${font.toFixed(2)}px" stroke-width="${(font*0.28).toFixed(2)}">${name}</text>
    </g>`;
  }).join('');

  // Fond de carte répété horizontalement (pas de W) pour couvrir les vues
  // qui débordent de [0, W] après déroulage de l'antiméridien.
  const kStart = Math.floor(minX / W);
  const kEnd = Math.floor(maxX / W);
  let landLayers = '';
  for (let k = kStart; k <= kEnd; k++) {
    landLayers += `<g transform="translate(${(k * W).toFixed(0)},0)">`
      + `<path class="land" d="${WORLD.land}"/>`
      + `<path class="borders" d="${WORLD.borders}"/></g>`;
  }

  // Rose des vents ancrée en bas à droite du cadre.
  const cr = U(24);
  const cx = maxX - cr * 1.5, cy = maxY - cr * 1.5;

  return `<svg class="route-map" width="${boxW}" height="${boxH}"
       viewBox="${minX.toFixed(1)} ${minY.toFixed(1)} ${vw.toFixed(1)} ${vh.toFixed(1)}"
       preserveAspectRatio="xMidYMid meet" xmlns="http://www.w3.org/2000/svg">
    <defs>
      <style>
        .ocean { fill: #e4ecec; }
        .land  { fill: #e9dcbb; stroke: #bfa877; stroke-width: ${coast.toFixed(2)};
                 stroke-linejoin: round; stroke-linecap: round; }
        .borders { fill: none; stroke: #c7b083; stroke-width: ${border.toFixed(2)};
                   stroke-linejoin: round; opacity: .45; }
        .route { fill: none; stroke: var(--red); stroke-width: ${strokeRoute.toFixed(2)};
                 stroke-dasharray: ${U(5).toFixed(1)} ${U(4).toFixed(1)}; stroke-linecap: round; }
        .m-label { font-family: "Century Gothic","Futura",sans-serif; fill: var(--ink);
                   font-weight: 700; paint-order: stroke; stroke: #e4ecec; }
        .frame { fill: none; stroke: var(--gold); }
      </style>
      <clipPath id="mapclip"><rect x="${minX.toFixed(1)}" y="${minY.toFixed(1)}" width="${vw.toFixed(1)}" height="${vh.toFixed(1)}"/></clipPath>
    </defs>
    <rect class="ocean" x="${minX.toFixed(1)}" y="${minY.toFixed(1)}" width="${vw.toFixed(1)}" height="${vh.toFixed(1)}"/>
    <g clip-path="url(#mapclip)">
      ${landLayers}
      <path class="route" d="${route}"/>
      ${markers}
      <g transform="translate(${cx.toFixed(1)},${cy.toFixed(1)})" opacity="0.7">
        <circle r="${cr.toFixed(1)}" fill="#f4ead0" stroke="var(--gold)" stroke-width="${(coast*1.2).toFixed(2)}"/>
        <path d="M0,${(-cr*1.15).toFixed(1)} L${(cr*0.22).toFixed(1)},0 L0,${(cr*0.3).toFixed(1)} L${(-cr*0.22).toFixed(1)},0 Z" fill="var(--red)"/>
        <path d="M0,${(cr*1.15).toFixed(1)} L${(cr*0.22).toFixed(1)},0 L${(-cr*0.22).toFixed(1)},0 Z" fill="var(--ink-soft)"/>
        <text y="${(-cr*1.3).toFixed(1)}" text-anchor="middle" style="font-size:${(cr*0.55).toFixed(1)}px;fill:var(--red);font-weight:700">N</text>
      </g>
    </g>
    <rect class="frame" x="${(minX+U(4)).toFixed(1)}" y="${(minY+U(4)).toFixed(1)}"
          width="${(vw-U(8)).toFixed(1)}" height="${(vh-U(8)).toFixed(1)}"
          stroke-width="${U(1.2).toFixed(2)}"/>
    <g transform="translate(${(minX+U(12)).toFixed(1)},${(minY+U(12)).toFixed(1)})">
      <rect x="0" y="0" width="${U(mapTitle.length * 7.4 + 20).toFixed(1)}" height="${U(24).toFixed(1)}" rx="${U(2).toFixed(1)}"
            fill="#f4ead0" stroke="var(--gold)" stroke-width="${U(0.8).toFixed(2)}" opacity="0.92"/>
      <text x="${U(10).toFixed(1)}" y="${U(16).toFixed(1)}"
            style="font-size:${U(10.5).toFixed(2)}px;fill:var(--red-deep);font-weight:700;letter-spacing:${U(0.6).toFixed(2)}px;font-family:'Century Gothic','Futura',sans-serif">${esc(mapTitle)}</text>
    </g>
  </svg>`;
}

function renderNotes(sheet) {
  if (!sheet.notes || !sheet.notes.length) return '';
  const items = sheet.notes.map((n) => `<li>${esc(n)}</li>`).join('');
  return `<section class="notes"><h3>${esc(T().notesTitle(sheet.name))}</h3><ul>${items}</ul></section>`;
}

function renderFlight(airline, sheet, flight) {
  const stopsCount = flight.itinerary.length - 2;

  els.detail.innerHTML = `
    <article class="flight-card">
      <div class="flight-head">
        <div>
          <span class="flight-no">${esc(flight.flightNo)}</span>
          <h2 class="flight-title">${esc(flight.line)}</h2>
          <p class="flight-sub">${esc(airline.name)} · ${esc(sheet.name)}${
            flight.lineHeader ? ` · ${esc(flight.lineHeader)}` : ''}${
            (flight.meta && flight.meta.length)
              ? ' · ' + flight.meta.map((m) => esc(m.value)).join(' · ') : ''}</p>
        </div>
        <div class="flight-route-summary">
          <strong>${esc(shortName(flight.origin))} → ${esc(shortName(flight.destination))}</strong><br/>
          ${esc(T().stopsLabel(stopsCount))}
          ${flight.frequency ? `<br/><span class="freq-badge">${esc(flight.frequency)}</span>` : ''}
        </div>
      </div>
      <div class="flight-body">
        <div class="map-pane">
          <div class="map-holder"></div>
        </div>
        <div class="timeline-pane">
          <h3>${esc(T().timelineTitle)}</h3>
          ${renderTimeline(flight)}
          ${renderNotes(sheet)}
        </div>
      </div>
    </article>`;
  els.detail.scrollTop = 0;
  currentFlight = flight;
  drawMap();
}

// Trace (ou re-trace) la carte du vol courant à la largeur réelle du panneau,
// hauteur fixe de 900 px.
const MAP_HEIGHT = 600;
function drawMap() {
  const holder = els.detail.querySelector('.map-holder');
  if (!holder || !currentFlight) return;
  const w = Math.max(Math.round(holder.clientWidth), 200);
  holder.innerHTML = renderMapSVG(currentFlight, w, MAP_HEIGHT);
}

// Re-trace au redimensionnement de la fenêtre (largeur variable).
let resizeTimer = null;
window.addEventListener('resize', () => {
  clearTimeout(resizeTimer);
  resizeTimer = setTimeout(drawMap, 120);
});

// ---------- Bascule de langue ----------
document.getElementById('lang-toggle').addEventListener('click', (e) => {
  const btn = e.target.closest('.lang-btn');
  if (btn && btn.dataset.lang !== LANG) setLang(btn.dataset.lang);
});

// ---------- Go (anglais par défaut) ----------
applyStaticI18n();
fillAirlines();
