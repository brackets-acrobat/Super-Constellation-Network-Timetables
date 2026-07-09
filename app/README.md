# Super Constellation — Horaires L-1049

Application de bureau **Electron** présentant les horaires (timetables) du réseau
exploité en Lockheed **L-1049 Super Constellation** par sept compagnies :
**TWA**, **Air France**, **Lufthansa**, **Eastern Air Lines**, **Flying Tiger**,
**Irish Airlines** et **Pan Am**.

Trois menus déroulants en cascade :

1. **Compagnie** → TWA · Air France · Lufthansa · Eastern Air Lines · Flying Tiger · Irish Airlines · Pan Am
2. **Réseau / Feuillet** → les feuillets de la compagnie choisie
3. **Vol / Ligne** → numéros de vol + intitulés des lignes

La sélection d'un vol affiche le détail : une **vraie carte vintage** (continents
en sépia, côtes et frontières) zoomée sur la région du vol avec le tracé
origine → escales → destination, et une **timeline** des horaires (arrivées /
départs, mentions « +1 jour »), plus les notes du réseau.

### Bilingue EN / FR

Une bascule **EN / FR** (en haut à droite) traduit toute l'application, interface
**et** données (feuillets, intitulés de lignes, fréquences, notes, noms de villes :
Londres→London, Le Caire→Cairo, …). **L'anglais est la langue par défaut à
l'ouverture.** Les deux versions sont compilées dans `timetables.json`
(`airlines` = FR, `airlinesEn` = EN, identifiants communs) ; la bascule conserve
le vol sélectionné.

## Indépendance vis-à-vis des fichiers Excel

L'application **ne dépend pas** des fichiers `.xlsx` d'origine ni des librairies
de build (`xlsx`, `world-atlas`, `topojson-client`) au runtime. Tout est
**compilé une fois** dans deux fichiers embarqués :

- [`src/data/timetables.json`](src/data/timetables.json) — vols, escales, horaires
  et coordonnées géographiques des villes ;
- [`src/data/worldmap.json`](src/data/worldmap.json) — géométrie mondiale (côtes +
  frontières) projetée en chemins SVG pour la carte.

Aucun accès réseau n'est nécessaire.

## Structure

```
app/
  src/
    main.js              Processus principal Electron
    preload.js           Pont sécurisé exposant les données au renderer
    index.html           Interface (3 menus + zone de détail)
    styles.css           Thème rétro aviation années 1950
    renderer.js          Cascade des menus, timeline, carte SVG
    data/timetables.json Données compilées (SEULE dépendance runtime)
    data/worldmap.json   Carte du monde en chemins SVG (dépendance runtime)
  scripts/
    parse-lib.js         Parsing des feuilles Excel  (build)
    place-coords.js      Table de géocodage des villes (build)
    compile-data.js      Génère src/data/timetables.json (build)
    build-map.js         Génère src/data/worldmap.json (build)
    smoke-test.js        Test de rendu automatisé
```

## Développement

```bash
npm install            # installe les outils de build (dev uniquement)
npm run build-all      # (re)génère timetables.json ET worldmap.json
npm start              # lance l'application
```

`npm run compile-data` lit les fichiers dans `../liaisons_Super_Constellation/`
et `npm run build-map` génère la carte du monde. À ne relancer que si les
fichiers Excel sources changent — les données compilées sont déjà présentes
dans `src/data/`.

## Créer l'installateur Windows

```bash
npm run dist           # produit un installateur .exe (electron-builder / NSIS)
```

Le résultat se trouve dans `dist/`. L'installateur ne contient que `src/**`
(dont `timetables.json`) et le runtime Electron — pas les fichiers Excel.
