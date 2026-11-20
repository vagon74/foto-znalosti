// Cache jen pro vlastní soubory aplikace; přihlášení a Graph jdou vždy přímo na síť.
const CACHE = 'foto-znalosti-1.0.1';
const ASSETS = ['./', './index.html', './app.js', './app.css', './config.js', './manifest.json',
  './icons/icon-192.png', './icons/icon-512.png', './icons/apple-touch-icon.png'];

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(ASSETS)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', (e) => {
  e.waitUntil(caches.keys()
    .then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
    .then(() => self.clients.claim()));
});

// network-first: nová verze se projeví hned, offline se použije cache
self.addEventListener('fetch', (e) => {
  const url = new URL(e.request.url);
  if (e.request.method !== 'GET' || url.origin !== location.origin) return;
  const key = url.searchParams.has('code') || url.searchParams.has('demo') ? './index.html' : e.request;
  e.respondWith(
    fetch(e.request)
      .then((res) => {
        if (res.ok && !url.search) { const copy = res.clone(); caches.open(CACHE).then((c) => c.put(e.request, copy)); }
        return res;
      })
      .catch(() => caches.match(key, { ignoreSearch: true }))
  );
});
