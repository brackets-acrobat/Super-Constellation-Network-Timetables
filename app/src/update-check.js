// Vérification des mises à jour au démarrage (check & notify), façon
// NavXpressVFR : interroge la dernière release GitHub, compare la version et,
// si une version plus récente existe, propose d'ouvrir la page de
// téléchargement. Aucune installation silencieuse (app non signée).
'use strict';

const { app, dialog, shell, net } = require('electron');

const OWNER = 'brackets-acrobat';
const REPO = 'Super-Constellation-Network-Timetables';
const LATEST_URL = `https://api.github.com/repos/${OWNER}/${REPO}/releases/latest`;

// Compare deux versions « x.y.z ». Renvoie 1 si a>b, -1 si a<b, 0 sinon.
function compareVersions(a, b) {
  const pa = String(a).replace(/^v/i, '').split('.').map((n) => parseInt(n, 10) || 0);
  const pb = String(b).replace(/^v/i, '').split('.').map((n) => parseInt(n, 10) || 0);
  for (let i = 0; i < Math.max(pa.length, pb.length); i++) {
    const d = (pa[i] || 0) - (pb[i] || 0);
    if (d) return d > 0 ? 1 : -1;
  }
  return 0;
}

// Récupère la dernière release publiée (JSON). Résout null en cas d'échec.
function fetchLatest() {
  return new Promise((resolve) => {
    let done = false;
    const finish = (v) => { if (!done) { done = true; resolve(v); } };
    try {
      const req = net.request({ method: 'GET', url: LATEST_URL });
      req.setHeader('User-Agent', `${REPO}/${app.getVersion()}`);
      req.setHeader('Accept', 'application/vnd.github+json');
      req.on('response', (res) => {
        if (res.statusCode < 200 || res.statusCode >= 300) { finish(null); return; }
        let body = '';
        res.on('data', (c) => { body += c; });
        res.on('end', () => {
          try { finish(JSON.parse(body)); } catch { finish(null); }
        });
        res.on('error', () => finish(null));
      });
      req.on('error', () => finish(null));
      req.end();
    } catch {
      finish(null);
    }
    // Filet de sécurité : ne jamais rester bloqué.
    setTimeout(() => finish(null), 15000);
  });
}

// Vérifie et notifie. `silent` (défaut true) : ne rien afficher si à jour ou
// en cas d'erreur réseau (adapté au démarrage). Passer false pour un contrôle
// manuel qui confirme « à jour ».
async function checkForUpdates(parentWin, { silent = true } = {}) {
  const rel = await fetchLatest();
  if (!rel || rel.draft || !rel.tag_name) {
    if (!silent) {
      dialog.showMessageBox(parentWin, {
        type: 'info',
        title: 'Mises à jour',
        message: 'Impossible de vérifier les mises à jour. / Could not check for updates.',
      });
    }
    return;
  }

  const current = app.getVersion();
  const latest = String(rel.tag_name).replace(/^v/i, '');

  if (compareVersions(latest, current) > 0) {
    const choice = dialog.showMessageBoxSync(parentWin, {
      type: 'info',
      title: 'Mise à jour disponible / Update available',
      message: `Une nouvelle version (${latest}) est disponible.\nA new version (${latest}) is available.`,
      detail: `Vous utilisez la version ${current}. / You are using version ${current}.`,
      buttons: ['Télécharger / Download', 'Plus tard / Later'],
      defaultId: 0,
      cancelId: 1,
      noLink: true,
    });
    if (choice === 0) {
      shell.openExternal(rel.html_url || `https://github.com/${OWNER}/${REPO}/releases/latest`);
    }
  } else if (!silent) {
    dialog.showMessageBox(parentWin, {
      type: 'info',
      title: 'Mises à jour / Updates',
      message: `Vous utilisez la dernière version (${current}).\nYou are running the latest version (${current}).`,
    });
  }
}

module.exports = { checkForUpdates, compareVersions, fetchLatest };
