/* ITR Generator service worker: caches the app shell for fast, installable loads. */
const VERSION = 'itr-v1.4.0';
const SHELL = [
  './', './index.html', './css/styles.css',
  './js/config.js', './js/tax.js', './js/excel.js', './js/app.js',
  './manifest.webmanifest', './icons/icon-192.png', './icons/icon-512.png', './icons/apple-touch-icon.png', './icons/logo.png', './icons/logo-mark.png'
];

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(VERSION).then((c) => c.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', (e) => {
  e.waitUntil(caches.keys().then((keys) => Promise.all(keys.filter((k) => k !== VERSION).map((k) => caches.delete(k)))).then(() => self.clients.claim()));
});

self.addEventListener('fetch', (e) => {
  const req = e.request;
  if (req.method !== 'GET') return; // never cache API calls (POST)
  const url = new URL(req.url);
  if (url.hostname.includes('script.google') || url.hostname.includes('googleusercontent')) return;

  // Same-origin app files: network first (always fresh after a GitHub update), cache fallback offline
  if (url.origin === location.origin) {
    e.respondWith(
      fetch(req).then((res) => {
        const copy = res.clone();
        caches.open(VERSION).then((c) => c.put(req, copy));
        return res;
      }).catch(() => caches.match(req).then((r) => r || caches.match('./index.html')))
    );
    return;
  }
  // CDN libraries and fonts: cache first
  if (/cdn\.sheetjs\.com|cdn\.jsdelivr\.net|fonts\.(googleapis|gstatic)\.com/.test(url.hostname)) {
    e.respondWith(
      caches.match(req).then((hit) => hit || fetch(req).then((res) => {
        const copy = res.clone();
        caches.open(VERSION).then((c) => c.put(req, copy));
        return res;
      }))
    );
  }
});
