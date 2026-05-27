const CACHE_NAME = 'rab-v1.0.1';
const ASSETS = [
  '/',
  '/index.html',
  '/manifest.json'
];

// No external font caching to avoid CORS issues
const EXTERNAL_URLS = [
  'https://fonts.googleapis.com',
  'https://fonts.gstatic.com'
];

self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE_NAME).then(async cache => {
      // Cache core assets
      for (const asset of ASSETS) {
        try {
          await cache.add(asset);
        } catch (err) {
          console.log('Failed to cache:', asset, err);
        }
      }
      return self.skipWaiting();
    })
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
      fetch(e.request)
        .catch(() => new Response(JSON.stringify({ ok: false, error: 'offline' }), {
          headers: { 'Content-Type': 'application/json' }
        }))
    );
    return;
  }

  // Stale-while-revalidate for external fonts
  if (EXTERNAL_URLS.some(external => url.hostname.includes(external))) {
    e.respondWith(
      caches.match(e.request).then(cached => {
        const fetchPromise = fetch(e.request).then(response => {
          if (response.ok) {
            caches.open(CACHE_NAME).then(cache => cache.put(e.request, response.clone()));
          }
          return response;
        }).catch(() => null);
        return cached || fetchPromise;
      })
    );
    return;
  }

  // Cache-first for static assets, fallback to index.html for navigation
  e.respondWith(
    caches.match(e.request).then(cached => {
      if (cached) return cached;
      
      return fetch(e.request).then(response => {
        if (response && response.ok && e.request.method === 'GET') {
          const responseToCache = response.clone();
          caches.open(CACHE_NAME).then(cache => {
            cache.put(e.request, responseToCache);
          });
        }
        return response;
      }).catch(() => {
        // Return index.html for navigation requests when offline
        if (e.request.mode === 'navigate') {
          return caches.match('/index.html');
        }
        return new Response('Offline - Content not available', { status: 404 });
      });
    })
  );
});
