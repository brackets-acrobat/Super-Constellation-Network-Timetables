// Génère l'icône de l'application (build/icon.ico + build/icon.png) à partir
// d'une image source PNG, redimensionnée en plusieurs tailles puis convertie
// en ICO.
//   electron scripts/make-icon.js [chemin_source.png]
'use strict';

const { app, nativeImage } = require('electron');
const path = require('path');
const fs = require('fs');

const SRC = process.argv[2]
  ? path.resolve(process.argv[2])
  : path.resolve(__dirname, '../l1049ico.png');
const OUT_DIR = path.resolve(__dirname, '../build');
const SIZES = [256, 128, 64, 48, 32, 16];

app.whenReady().then(async () => {
  if (!fs.existsSync(SRC)) {
    console.error('Image source introuvable : ' + SRC);
    app.exit(1);
    return;
  }
  fs.mkdirSync(OUT_DIR, { recursive: true });

  const base = nativeImage.createFromPath(SRC);
  if (base.isEmpty()) {
    console.error('Image illisible : ' + SRC);
    app.exit(1);
    return;
  }

  // PNG 256 (aperçu / usage multiplateforme).
  const png256 = base.resize({ width: 256, height: 256, quality: 'best' });
  fs.writeFileSync(path.join(OUT_DIR, 'icon.png'), png256.toPNG());

  // ICO multi-tailles (les dimensions d'un ICO doivent être ≤ 256).
  const buffers = SIZES.map((s) =>
    base.resize({ width: s, height: s, quality: 'best' }).toPNG());
  const pngToIco = (await import('png-to-ico')).default;
  const ico = await pngToIco(buffers);
  fs.writeFileSync(path.join(OUT_DIR, 'icon.ico'), ico);

  console.log('✓ source : ' + SRC);
  console.log('✓ build/icon.png  (256×256)');
  console.log('✓ build/icon.ico  (' + SIZES.join(', ') + ')');
  app.quit();
});
