// RSA Driver Attendance — Service Worker
// Caches the GitHub Pages shell so the install prompt works offline.
// The actual GAS app content is always fetched live (never cached).

var CACHE = 'rsa-attend-v2';
var SHELL = [
  '/rsa-driver-attendance/',
  '/rsa-driver-attendance/index.html',
  '/rsa-driver-attendance/manifest.json',
  '/rsa-driver-attendance/icon.svg'
];

self.addEventListener('install', function(e) {
  e.waitUntil(
    caches.open(CACHE).then(function(cache) {
      return cache.addAll(SHELL);
    })
  );
  self.skipWaiting();
});

self.addEventListener('activate', function(e) {
  // Delete any old cache versions
  e.waitUntil(
    caches.keys().then(function(keys) {
      return Promise.all(
        keys.filter(function(k) { return k !== CACHE; })
            .map(function(k) { return caches.delete(k); })
      );
    })
  );
  e.waitUntil(clients.claim());
});

self.addEventListener('fetch', function(e) {
  // Never intercept GAS requests — always needs live network
  if (e.request.url.includes('script.google.com') ||
      e.request.url.includes('googleusercontent.com') ||
      e.request.url.includes('googleapis.com')) {
    return;
  }

  // Cache-first for shell assets (HTML, manifest, icon)
  e.respondWith(
    caches.match(e.request).then(function(cached) {
      return cached || fetch(e.request);
    })
  );
});
