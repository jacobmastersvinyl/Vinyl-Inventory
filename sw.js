/* Funkified — Service Worker v6
   - Caches the app shell so the PWA loads offline.
   - Network-first for same-origin GETs (so deploys land fast).
   - Passes through POSTs and external endpoints without touching them.
   - Supports skipWaiting on demand.
*/
const VERSION = 'funkified-v9-2026-04-28-create-return-lightbox';
const APP_SHELL = [
  './',
  './index.html',
  './manifest.json',
  'https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap'
];

const PASSTHROUGH_HOSTS = [
  'script.google.com',
  'script.googleusercontent.com',
  'drive.google.com',
  'googleusercontent.com',
  'nominatim.openstreetmap.org'
];

self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(VERSION).then(cache => cache.addAll(APP_SHELL)).catch(() => {})
  );
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys => Promise.all(
      keys.filter(k => k !== VERSION).map(k => caches.delete(k))
    )).then(() => self.clients.claim())
  );
});

self.addEventListener('message', event => {
  if (event.data && event.data.type === 'SKIP_WAITING') self.skipWaiting();
});

self.addEventListener('fetch', event => {
  const req = event.request;
  const url = new URL(req.url);

  // Never intercept non-GET. Apps Script POSTs must go straight to the network
  // so the client-side drain / idempotency layer stays in charge.
  if (req.method !== 'GET') return;

  // Passthrough for third-party endpoints we don't want to cache.
  if (PASSTHROUGH_HOSTS.some(h => url.hostname.endsWith(h))) return;

  // Network-first strategy for same-origin + allowlisted GETs.
  event.respondWith(
    fetch(req).then(res => {
      // Cache successful basic responses for next offline session.
      if (res && res.status === 200 && (res.type === 'basic' || res.type === 'cors')) {
        const copy = res.clone();
        caches.open(VERSION).then(c => c.put(req, copy)).catch(() => {});
      }
      return res;
    }).catch(() => caches.match(req).then(hit => hit || caches.match('./index.html')))
  );
});
