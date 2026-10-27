import fs from 'node:fs';
import assert from 'node:assert/strict';

const html = fs.readFileSync(new URL('../index.html', import.meta.url), 'utf8');
const source = fs.readFileSync(new URL('../src/main.jsx', import.meta.url), 'utf8');
const pkg = JSON.parse(fs.readFileSync(new URL('../package.json', import.meta.url), 'utf8'));

assert.match(html, /<div id="root"><\/div>/, 'React root must remain present');
assert.match(html, /type="module" src="\.\/src\/main\.jsx"/, 'Vite module entry must be present');
assert.doesNotMatch(html, /text\/babel|@babel\/standalone/i, 'Babel Standalone must not load in production source');
assert.match(html, /react\.production\.min\.js/, 'React global runtime retained');
assert.match(html, /react-dom@18\/umd\/react-dom\.production\.min\.js/, 'ReactDOM global runtime retained');
assert.match(html, /leaflet@1\.9\.4\/dist\/leaflet\.js/, 'Leaflet runtime retained');
assert.match(html, /supabase-js@2\/dist\/umd\/supabase\.js/, 'Supabase runtime retained');
assert.match(source, /ReactDOM\.createRoot\(rootElement\)\.render\(<App \/>\)/, 'App entry must remain');
assert.match(source, /function ensureXLSXLoaded\(/, 'Lazy Excel loader retained');
assert.match(source, /await ensureXLSXLoaded\(\)/, 'Excel export must await lazy loader');
assert.match(source, /isFetchingAll/, 'Refresh single-flight guard retained');
assert.match(source, /refreshQueued/, 'Queued refresh retained');
assert.match(source, /beforeinstallprompt/, 'PWA install prompt retained');
assert.match(source, /serviceWorker\.register\('\.\/sw\.js'\)/, 'Service worker registration retained');
assert.match(source, /handleExportRekapanHarian/, 'Daily report export retained');
assert.match(source, /presensiList/, 'Attendance data flow retained');
assert.match(source, /supabase/, 'Supabase integration retained');
assert.match(source, /More Apps/, 'More Apps interface retained');
assert.match(source, /lembur/i, 'Overtime flow retained');
assert.match(source, /klaim/i, 'Claims flow retained');
assert.ok(pkg.scripts?.build, 'Build script must be configured');
console.log('PASS: 18 static smoke checks. These checks verify source markers only, not live Supabase or browser runtime behavior.');
