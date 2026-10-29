const CACHE = 'omr-v5';
const FILES = ['./', './index.html', './app.js', './omr.js', './manifest.webmanifest', './icons/icon-192.png', './icons/icon-512.png', './lib/opencv.js', './lib/xlsx.full.min.js'];
self.addEventListener('install', e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(FILES)).then(() => self.skipWaiting())); });
self.addEventListener('activate', e => { e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim())); });
self.addEventListener('fetch', e => {
  if (e.request.method !== 'GET' || new URL(e.request.url).origin !== location.origin) return; // Gemini API সরাসরি নেটওয়ার্কে
  e.respondWith(caches.match(e.request).then(r => r || fetch(e.request)));
});
