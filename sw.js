// ══════════════════════════════════════════════════════════════════
// Funkified — Service Worker
//
// Responsibilities:
//   • Cache the app shell so it loads instantly even fully offline
//   • Do NOT intercept POSTs to Apps Script (the page handles queuing
//     via IndexedDB; the SW has no business touching those).
//   • Network-first for same-origin so shell updates deploy cleanly.
// ══════════════════════════════════════════════════════════════════

const VERSION = 'funkified-v5-2026-04-20';
const APP_SHELL = [
  './',
  './index.html',
  './manifest.json',
  'https://fonts.googleapis.com/css2?family=Bebas+Neue&family=DM+Mono:wght@400;500&family=Barlow:wght@400;500;600&display=swap',
];

// ── INSTALL: cache app shell ────────────────────────────────────
self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(VERSION).then((cache) => cache.addAll(APP_SHELL)).catch(() => {})
  );
  self.skipWaiting();
});

// ── ACTIVATE: clean old caches ──────────────────────────────────
self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== VERSION).map((k) => caches.delete(k)))
    )
  );
  self.clients.claim();
});

// ── FETCH: strategy ─────────────────────────────────────────────
self.addEventListener('fetch', (event) => {
  const req = event.request;
  const url = new URL(req.url);

  // Never intercept Apps Script POSTs — the page owns that path.
  if (req.method !== 'GET') return;

  // Never intercept Apps Script GETs either (search/lookup).
  // The page has its own IndexedDB cache for offline browse.
  if (url.hostname.includes('script.google.com') ||
      url.hostname.includes('googleusercontent.com')) {
    return;
  }

  // Don't cache photo URLs from Drive (they're large; let browser handle)
  if (url.hostname.includes('drive.google.com')) return;

  // Don't try to cache Nominatim (reverse geocode) responses
  if (url.hostname.includes('nominatim.openstreetmap.org')) return;

  // Same-origin + Google Fonts: network-first, fall back to cache
  event.respondWith(
    fetch(req)
      .then((resp) => {
        // Only cache successful GETs
        if (resp && resp.ok && resp.status === 200) {
          const copy = resp.clone();
          caches.open(VERSION).then((cache) => cache.put(req, copy)).catch(() => {});
        }
        return resp;
      })
      .catch(() => caches.match(req).then((cached) => cached || caches.match('./index.html')))
  );
});

// ── MESSAGE: allow page to request cache refresh ────────────────
self.addEventListener('message', (event) => {
  if (event.data === 'skipWaiting') self.skipWaiting();
});
