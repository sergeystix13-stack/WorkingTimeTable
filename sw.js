// Офлайн-режим: после первого открытия приложение работает без интернета.
// При обновлении сайта меняется VERSION — телефон скачает новые файлы при следующем запуске с интернетом.
const VERSION = "4778804631";
const CACHE = "tabel-" + VERSION;
const FILES = ["./", "app.css", "app.js", "calc.js", "fonts/Monomakh-Regular.ttf", "fonts/OldStandard-Bold.ttf", "fonts/RuslanDisplay-Regular.ttf", "fonts/Shafarik-Regular.ttf", "icons/apple-touch-icon.png", "icons/icon-192.png", "icons/icon-512.png", "icons/icon-maskable-512.png", "index.html", "install-hint.js", "manifest.webmanifest", "planner.js", "pwa.css", "sw-register.js", "sync.js", "xlsx.js"];

self.addEventListener("install", event => {
  event.waitUntil(caches.open(CACHE).then(c => c.addAll(FILES)).then(() => self.skipWaiting()));
});
self.addEventListener("activate", event => {
  event.waitUntil(caches.keys()
    .then(keys => Promise.all(keys.filter(k => k.startsWith("tabel-") && k !== CACHE).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});
self.addEventListener("fetch", event => {
  const req = event.request;
  if (req.method !== "GET" || new URL(req.url).origin !== location.origin) return;
  event.respondWith(
    caches.match(req, { ignoreSearch: true }).then(hit => {
      if (hit) return hit;
      if (req.mode === "navigate") return caches.match("index.html").then(r => r || fetch(req));
      return fetch(req);
    })
  );
});
