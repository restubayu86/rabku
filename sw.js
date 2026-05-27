const CACHE_NAME = 'rab-v1.0.0';
const ASSETS = [
  './',
  './index.html',
  './manifest.json',
  'https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap'
];

self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE_NAME).then(cache => cache.addAll(ASSETS)).then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k)))
    ).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);
  
  // Network-first for API calls (Apps Script)
  if (url.hostname.includes('script.google.com')) {
    e.respondWith(
      fetch(e.request).catch(() => new Response(JSON.stringify({ ok: false, error: 'offline' }), {
        headers: { 'Content-Type': 'application/json' }
      }))
    );
    return;
  }

  // Cache-first for static assets
  e.respondWith(
    caches.match(e.request).then(cached => {
      if (cached) return cached;
      return fetch(e.request).then(resp => {
        if (resp.ok) {
          const clone = resp.clone();
          caches.open(CACHE_NAME).then(cache => cache.put(e.request, clone));
        }
        return resp;
      }).catch(() => {
        if (e.request.mode === 'navigate') {
          return caches.match('./index.html');
        }
      });
    })
  );
});

// Background sync untuk data yang pending
self.addEventListener('sync', e => {
  if (e.tag === 'sync-rab') {
    e.waitUntil(syncPendingData());
  }
});

async function syncPendingData() {
  const cache = await caches.open('rab-pending');
  const requests = await cache.keys();
  for (const req of requests) {
    try {
      const data = await cache.match(req);
      const body = await data.json();
      await fetch(body.url, { method: 'POST', body: JSON.stringify(body.payload) });
      await cache.delete(req);
    } catch (e) {}
  }
}
