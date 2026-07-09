// Test de fumée : charge l'UI, simule les 3 sélections, capture les erreurs
// console et produit des PNG du rendu. Usage : electron scripts/smoke-test.js
'use strict';
const { app, BrowserWindow } = require('electron');
const path = require('path');
const fs = require('fs');

const OUTDIR = process.env.SMOKE_OUT || path.join(app.getPath('temp'), 'sc-smoke');
fs.mkdirSync(OUTDIR, { recursive: true });
const errors = [];

app.whenReady().then(async () => {
  const win = new BrowserWindow({
    width: 1180, height: Number(process.env.SMOKE_H || 820), show: false,
    webPreferences: { preload: path.join(__dirname, '../src/preload.js'), contextIsolation: true, nodeIntegration: false, sandbox: false },
  });
  win.webContents.on('console-message', (_e, level, msg) => { if (level >= 2) errors.push(msg); });
  win.webContents.on('render-process-gone', (_e, d) => errors.push('render-gone: ' + JSON.stringify(d)));

  await win.loadFile(path.join(__dirname, '../src/index.html'));
  await new Promise((r) => setTimeout(r, 400));
  // Désactive les animations pour des captures nettes.
  await win.webContents.insertCSS('*{animation:none!important;transition:none!important}');

  // Vérifie que les données sont accessibles et que les menus se peuplent.
  const airlineCount = await win.webContents.executeJavaScript(
    `document.getElementById('sel-airline').options.length`);

  // Contrôle i18n : langue par défaut + textes statiques + bascule.
  const i18n = await win.webContents.executeJavaScript(`(function(){
    const active=document.querySelector('.lang-btn.active');
    return { htmlLang: document.documentElement.lang,
      activeBtn: active ? active.dataset.lang : null,
      lblAirline: (document.getElementById('lbl-airline')||{}).textContent,
      brandEra: (document.getElementById('brand-era')||{}).textContent,
      airlinePh: document.getElementById('sel-airline').options[0].textContent };
  })()`);
  console.log('I18N_DEFAULT=' + JSON.stringify(i18n));

  // Sélection : TWA -> Réseau International -> premier vol (TW 800).
  await win.webContents.executeJavaScript(`(function(){
    const a=document.getElementById('sel-airline'); a.value='twa'; a.dispatchEvent(new Event('change'));
    const s=document.getElementById('sel-sheet'); s.value=s.options[2].value; s.dispatchEvent(new Event('change'));
    const f=document.getElementById('sel-flight'); f.value=f.options[1].value; f.dispatchEvent(new Event('change'));
    return true;
  })()`);
  await new Promise((r) => setTimeout(r, 300));

  const report = await win.webContents.executeJavaScript(`(function(){
    const d=document.getElementById('detail');
    return {
      title: (document.querySelector('.flight-title')||{}).textContent||null,
      flightNo: (document.querySelector('.flight-no')||{}).textContent||null,
      stops: document.querySelectorAll('.tl-item').length,
      hasMap: !!document.querySelector('svg.route-map'),
      mapMarkers: document.querySelectorAll('svg.route-map circle').length,
      notes: document.querySelectorAll('.notes li').length,
      detailScrollH: d.scrollHeight, detailClientH: d.clientHeight,
      vScroll: d.scrollHeight > d.clientHeight + 1,
    };
  })()`);
  await win.webContents.capturePage().then((img) =>
    fs.writeFileSync(path.join(OUTDIR, 'twa-intl.png'), img.toPNG()));

  // Sélectionne par id de compagnie / index de feuillet / index de vol.
  async function pick(airlineId, sheetIdx, flightIdx, shot) {
    await win.webContents.executeJavaScript(`(function(){
      const a=document.getElementById('sel-airline'); a.value='${airlineId}'; a.dispatchEvent(new Event('change'));
      const s=document.getElementById('sel-sheet'); s.value=s.options[${sheetIdx}].value; s.dispatchEvent(new Event('change'));
      const f=document.getElementById('sel-flight'); f.value=f.options[${flightIdx}].value; f.dispatchEvent(new Event('change'));
      return true;
    })()`);
    await new Promise((r) => setTimeout(r, 250));
    const rep = await win.webContents.executeJavaScript(`(function(){
      const d=document.getElementById('detail');
      return { airlineSheets: document.getElementById('sel-sheet').options.length,
        title:(document.querySelector('.flight-title')||{}).textContent||null,
        sub:(document.querySelector('.flight-sub')||{}).textContent||null,
        stops:document.querySelectorAll('.tl-item').length,
        markers:document.querySelectorAll('svg.route-map circle').length,
        vScroll: d.scrollHeight > d.clientHeight + 1 };
    })()`);
    if (shot) await win.webContents.capturePage().then((img) =>
      fs.writeFileSync(path.join(OUTDIR, shot + '.png'), img.toPNG()));
    return rep;
  }

  // Test de bascule EN -> FR sur un vol sélectionné.
  await pick('twa', 2, 1, 'twa-en');
  const enTitle = await win.webContents.executeJavaScript(
    `(document.querySelector('.flight-title')||{}).textContent`);
  await win.webContents.executeJavaScript(
    `document.querySelector('.lang-btn[data-lang="fr"]').click()`);
  await new Promise((r) => setTimeout(r, 250));
  const toggled = await win.webContents.executeJavaScript(`(function(){
    return { htmlLang: document.documentElement.lang,
      active: (document.querySelector('.lang-btn.active')||{}).dataset ? document.querySelector('.lang-btn.active').dataset.lang : null,
      title:(document.querySelector('.flight-title')||{}).textContent,
      airlineVal: document.getElementById('sel-airline').value,
      flightVal: document.getElementById('sel-flight').value,
      mapCartouche: (document.querySelector('svg.route-map text')||{}).textContent };
  })()`);
  await win.webContents.capturePage().then((img) =>
    fs.writeFileSync(path.join(OUTDIR, 'twa-fr.png'), img.toPNG()));
  console.log('TOGGLE_EN_TITLE=' + JSON.stringify(enTitle));
  console.log('TOGGLE_FR=' + JSON.stringify(toggled));
  // Repasse en anglais pour les captures suivantes.
  await win.webContents.executeJavaScript(
    `document.querySelector('.lang-btn[data-lang="en"]').click()`);
  await new Promise((r) => setTimeout(r, 150));

  const report2 = await pick('air-france', 2, 1, 'af-orient');
  // 1re option de feuillet = index 1 (index 0 = placeholder).
  const ft = await pick('flying-tiger', 1, 3, 'flying-tiger');   // MATS Pacifique (dateline)
  const ir = await pick('irish-airlines', 1, 1, 'irish');        // 1 feuillet + col Compagnies
  const pa = await pick('pan-am', 2, 1, 'panam-igs');            // Pan Am 2e feuillet (IGS, direct)

  console.log('AIRLINE_OPTIONS=' + airlineCount);
  console.log('REPORT1=' + JSON.stringify(report));
  console.log('REPORT2=' + JSON.stringify(report2));
  console.log('FLYING_TIGER=' + JSON.stringify(ft));
  console.log('IRISH=' + JSON.stringify(ir));
  console.log('PANAM=' + JSON.stringify(pa));
  console.log('ERRORS=' + JSON.stringify(errors));
  console.log('OUTDIR=' + OUTDIR);
  app.quit();
});
