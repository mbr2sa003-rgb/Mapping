// RC Location Map - Service Worker
// This only caches the app's own static files (so the app shell loads fast/offline).
// It never caches Google Sign-In, Google Maps, Apps Script data, or OpenStreetMap
// requests - those always go to the network, since pin data must always be fresh
// and Google's own scripts should not be cached by us.

const CACHE_NAME = 'rc-map-shell-v1';
const APP_SHELL = [
  './index.html',
  './manifest.json',
  './icon-192.png',
  './icon-512.png'
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(APP_SHELL))
  );
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((names) =>
      Promise.all(names.filter((n) => n !== CACHE_NAME).map((n) => caches.delete(n)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', (event) => {
  const url = new URL(event.request.url);

  // Only handle same-origin requests for our own app shell files.
  // Everything else (Google APIs, Apps Script, Nominatim, etc.) goes straight
  // to the network as normal, untouched.
  if (url.origin !== self.location.origin) return;

  event.respondWith(
    caches.match(event.request).then((cached) => {
      const networkFetch = fetch(event.request)
        .then((response) => {
          if (response && response.status === 200) {
            const clone = response.clone();
            caches.open(CACHE_NAME).then((cache) => cache.put(event.request, clone));
          }
          return response;
        })
        .catch(() => cached); // offline fallback to cached shell if network fails
      return cached || networkFetch;
    })
  );
});
