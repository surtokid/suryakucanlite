# SuryaKuCan — Vite production-build migration

This project uses the latest confirmed-working SuryaKuCan v16 source as its basis. The goal of this first migration stage is to compile JSX ahead of time with Vite and remove Babel Standalone from browser startup, while keeping the existing global React, ReactDOM, Leaflet, Supabase, and Tailwind CDN setup to minimize behavioral changes.

## Requirements
- Node.js 20.19+ (or a compatible current LTS)
- npm
- Internet access during `npm install` and first runtime load, because this cautious first stage still uses existing CDN dependencies.

## Build
1. Copy the existing `Logo SuryaKuCan.png`, `manifest.json`, and `sw.js` from the current production repository into `public/`. Keep the original contents of `manifest.json` and `sw.js`.
2. Run `npm install`.
3. Run `npm run test:smoke`.
4. Run `npm run build`.
5. Inspect `dist/` and confirm it contains `index.html`, `assets/`, `Logo SuryaKuCan.png`, `manifest.json`, and `sw.js`.
6. Review the original service worker's `APP_SHELL` list. Since Vite emits a hashed JavaScript filename under `assets/`, make sure the worker's offline precache strategy covers that file before publishing. Do not publish until this is confirmed.
7. Deploy the *contents* of `dist/` to the existing `/suryakucan/` hosting location, preserving the current custom domain and route.

`base: './'` is intentional so built assets resolve under a nested path such as `https://suryatokoku.my.id/suryakucan/`.

## What is preserved in this migration stage
- Existing JSX/React app logic and root render.
- Current CDN React/ReactDOM globals, Leaflet, Supabase JS, Tailwind CDN, external font/icon links.
- Lazy loading for SheetJS/Excel.
- Existing PWA install handler and registration of `./sw.js`.
- Existing Supabase and attendance/operational flows in source.

## Important limitations
- This stage does not bundle React, Tailwind, Leaflet, or Supabase into local assets; doing that is a separate step with more regression risk.
- Static smoke checks are not a substitute for testing with real Supabase credentials and a real browser.
- Do not replace the production site until the build succeeds and the regression checklist below has been completed.

## Regression checklist before production
- [ ] Login/logout for employee and admin
- [ ] Presensi masuk and pulang; status libur/izin/sakit/cuti behavior
- [ ] Realtime refresh and offline/online recovery
- [ ] Master karyawan, divisi, jabatan, lokasi, and shift changes
- [ ] Klaim and lembur flows
- [ ] Rekapan harian filters and Excel export
- [ ] More Apps links and permissions
- [ ] PWA install, refresh, and offline reopen (including service-worker cache)
- [ ] Camera/location flows on the Android wrapper if used

Rollback: keep the current production `index.html` and PWA assets backed up; restore them if any regression appears.
