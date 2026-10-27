import fs from 'node:fs';
import path from 'node:path';

const required = ['Logo SuryaKuCan.png', 'manifest.json', 'sw.js'];
const missing = required.filter((name) => !fs.existsSync(path.join('public', name)));
if (missing.length) {
  console.error('Build dihentikan agar aset PWA tidak hilang. Salin aset asli ke public/:');
  for (const name of missing) console.error(`- public/${name}`);
  process.exit(1);
}
console.log('PASS: logo, manifest, dan service worker tersedia.');
