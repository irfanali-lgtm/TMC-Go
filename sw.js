// TMC Go service worker: makes the app installable and opens instantly, even on a weak signal.
// Pages are fetched from the network first (so updates arrive straight away) and fall back to the last copy;
// opening the app gives the network 3 seconds, then the last copy opens and the fresh one is kept for next time.
// Data calls to Google are never cached.
const CACHE = 'tmcgo-v193';
const SHELL = ['./', 'index.html', 'manifest.webmanifest', 'icon-192.png', 'icon-512.png'];
const PAGE_WAIT = 3000;
self.addEventListener('install', e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).then(() => self.skipWaiting())); });
self.addEventListener('activate', e => { e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim())); });
self.addEventListener('fetch', e => {
  const u = new URL(e.request.url);
  if (e.request.method !== 'GET' || u.origin !== location.origin) return;           // Google data, CDNs: straight to network
  let saved = Promise.resolve();
  const net = fetch(e.request).then(r => { if (r.ok && !u.search) { const copy = r.clone(); saved = caches.open(CACHE).then(c => c.put(e.request, copy)).catch(() => {}); } return r; });
  e.waitUntil(net.then(() => saved, () => {}));                                     // the fresh copy is saved even when the last one opened first
  const last = () => caches.match(e.request).then(m => m || (e.request.mode === 'navigate' ? caches.match('index.html') : undefined));
  if (e.request.mode !== 'navigate') { e.respondWith(net.catch(() => last().then(m => m || Response.error()))); return; }
  e.respondWith(new Promise(done => {
    let sent = false; const send = r => { if (!sent) { sent = true; done(r); } };
    const slow = setTimeout(() => last().then(m => { if (m) send(m); }), PAGE_WAIT);   // weak signal: open the last copy
    net.then(r => { clearTimeout(slow); send(r); }, () => { clearTimeout(slow); last().then(m => send(m || Response.error())); });
  }));
});
