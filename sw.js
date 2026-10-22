// TMC Go service worker: makes the app installable and opens instantly, even on a weak signal.
// Pages are fetched from the network first (so updates arrive straight away) and fall back to the last copy.
// Data calls to Google are never cached.
const CACHE = 'tmcgo-v156';
const SHELL = ['./', 'index.html', 'manifest.webmanifest', 'icon-192.png', 'icon-512.png'];
self.addEventListener('install', e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).then(() => self.skipWaiting())); });
self.addEventListener('activate', e => { e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim())); });
self.addEventListener('fetch', e => {
  const u = new URL(e.request.url);
  if (e.request.method !== 'GET' || u.origin !== location.origin) return;           // Google data, CDNs: straight to network
  e.respondWith(fetch(e.request).then(r => { if (r.ok && !u.search) { const copy = r.clone(); caches.open(CACHE).then(c => c.put(e.request, copy)); } return r; })
    .catch(() => caches.match(e.request).then(m => m || (e.request.mode === 'navigate' ? caches.match('index.html') : Response.error()))));
});
