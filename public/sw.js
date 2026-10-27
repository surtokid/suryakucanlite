
const CACHE_NAME = 'app-suryakucan-v4';

const APP_SHELL = [
  './',
  './index.html',
  './manifest.json',
  './Logo%20SuryaKuCan.png'
];

self.addEventListener('install', (event) => {
  event.waitUntil((async () => {
    const cache = await caches.open(CACHE_NAME);

    for (const path of APP_SHELL) {
      try {
        const response = await fetch(path, { cache: 'reload' });
        if (response.ok) {
          await cache.put(path, response);
        }
      } catch (error) {
        console.warn('Tidak dapat cache:', path, error);
      }
    }

    await self.skipWaiting();
  })());
});

self.addEventListener('activate', (event) => {
  event.waitUntil((async () => {
    const keys = await caches.keys();

    await Promise.all(
      keys
        .filter((key) =>
          key.startsWith('app-suryakucan-') &&
          key !== CACHE_NAME
        )
        .map((key) => caches.delete(key))
    );

    await self.clients.claim();
  })());
});

self.addEventListener('fetch', (event) => {
  const request = event.request;

  if (request.method !== 'GET') return;

  const url = new URL(request.url);

  if (url.origin !== self.location.origin) return;

  event.respondWith((async () => {
    try {
      const response = await fetch(request);

      if (response && response.ok) {
        const cache = await caches.open(CACHE_NAME);
        cache.put(request, response.clone()).catch(() => {});
      }

      return response;
    } catch (error) {
      const cached = await caches.match(request);

      if (cached) return cached;

      if (request.mode === 'navigate') {
        const appPage = await caches.match('./index.html');
        if (appPage) return appPage;
      }

      return new Response(
        'Halaman tidak tersedia saat offline.',
        {
          status: 503,
          headers: {
            'Content-Type': 'text/plain; charset=utf-8'
          }
        }
      );
    }
  })());
});
