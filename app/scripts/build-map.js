// Génère src/data/worldmap.json : la géométrie mondiale (côtes + frontières)
// projetée une fois en équirectangulaire, sous forme de chemins SVG.
//
//   node scripts/build-map.js
//
// Build-time uniquement. L'app ne charge que le JSON produit (aucune dépendance
// à world-atlas / topojson au runtime).
'use strict';

const path = require('path');
const fs = require('fs');
const topojson = require('topojson-client');

const land = require('world-atlas/land-50m.json');
const countries = require('world-atlas/countries-50m.json');

// Projection équirectangulaire : 10 px par degré -> monde 3600 x 1800.
const PXPERDEG = 10;
const W = 360 * PXPERDEG;
const H = 180 * PXPERDEG;
const projX = (lon) => +((lon + 180) * PXPERDEG).toFixed(1);
const projY = (lat) => +((90 - lat) * PXPERDEG).toFixed(1);

// Convertit une géométrie GeoJSON (Polygon/MultiPolygon/LineString/MultiLineString)
// en une liste de chaînes de commandes SVG.
function toPaths(geom, close) {
  const out = [];
  // Découpe un anneau chaque fois que deux points consécutifs traversent
  // l'antiméridien (saut de longitude > 180°). Sans cela, la projection
  // équirectangulaire tracerait un trait horizontal en travers de la carte.
  const ringToPath = (ring) => {
    const subpaths = [];
    let cur = [];
    let prevLon = null;
    for (const [lon, lat] of ring) {
      if (prevLon !== null && Math.abs(lon - prevLon) > 180) {
        if (cur.length) subpaths.push(cur);
        cur = [];
      }
      cur.push([lon, lat]);
      prevLon = lon;
    }
    if (cur.length) subpaths.push(cur);
    return subpaths
      .filter((sp) => sp.length >= 2)
      .map((sp) => {
        let d = '';
        sp.forEach((pt, i) => { d += (i === 0 ? 'M' : 'L') + projX(pt[0]) + ' ' + projY(pt[1]); });
        if (close && sp.length > 2) d += 'Z';
        return d;
      })
      .join(' ');
  };
  const walk = (g) => {
    if (!g) return;
    switch (g.type) {
      case 'Polygon':
        g.coordinates.forEach((ring) => out.push(ringToPath(ring)));
        break;
      case 'MultiPolygon':
        g.coordinates.forEach((poly) => poly.forEach((ring) => out.push(ringToPath(ring))));
        break;
      case 'LineString':
        out.push(ringToPath(g.coordinates));
        break;
      case 'MultiLineString':
        g.coordinates.forEach((line) => out.push(ringToPath(line)));
        break;
      case 'GeometryCollection':
        g.geometries.forEach(walk);
        break;
      default:
        break;
    }
  };
  walk(geom);
  return out;
}

function main() {
  const landFc = topojson.feature(land, land.objects.land);
  const landGeoms = landFc.features
    ? landFc.features.map((f) => f.geometry)
    : [landFc.geometry || landFc];
  const landPaths = landGeoms.flatMap((g) => toPaths(g, true));

  // Frontières internes uniquement (a !== b) : évite de redessiner les côtes.
  const borderMesh = topojson.mesh(
    countries, countries.objects.countries, (a, b) => a !== b);
  const borderPaths = toPaths(borderMesh, false);

  const data = {
    projection: 'equirectangular',
    pxPerDeg: PXPERDEG,
    width: W,
    height: H,
    land: landPaths.join(' '),
    borders: borderPaths.join(' '),
  };

  const OUT = path.resolve(__dirname, '../src/data/worldmap.json');
  fs.writeFileSync(OUT, JSON.stringify(data), 'utf8');
  const kb = (fs.statSync(OUT).size / 1024).toFixed(0);
  console.log(`✓ ${OUT}`);
  console.log(`  ${landPaths.length} anneaux de terres, ${borderPaths.length} tronçons de frontières, ${kb} Ko`);
}

main();
