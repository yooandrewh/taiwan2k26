/* OneSignal push (loads only when online) + offline app-shell cache.
   The importScripts is wrapped so the worker still installs with no network. */
try { importScripts("https://cdn.onesignal.com/sdks/web/v16/OneSignalSDK.sw.js"); } catch (e) {}

var CACHE = 'tw2k26-v260';
var SHELL = ['./', './index.html', './manifest.json', './icon-180.png', './icon-512.png', './park-map-v.webp', './park-map-v.jpg'];

self.addEventListener('install', function (e) {
  self.skipWaiting();
  e.waitUntil(caches.open(CACHE).then(function (c) {
    return Promise.all(SHELL.map(function (u) { return c.add(u).catch(function () {}); }));
  }));
});

self.addEventListener('activate', function (e) {
  e.waitUntil(
    caches.keys().then(function (keys) {
      return Promise.all(keys.map(function (k) { if (k !== CACHE) return caches.delete(k); }));
    }).then(function () { return self.clients.claim(); })
  );
});

self.addEventListener('fetch', function (e) {
  var req = e.request;
  if (req.method !== 'GET') return;
  var url = new URL(req.url);
  if (url.origin !== self.location.origin) return; // let CDN / OneSignal / maps go straight to network

  // App page: network-first (so online users always get the latest), fall back to cache offline.
  if (req.mode === 'navigate') {
    e.respondWith(
      fetch(req).then(function (res) {
        var copy = res.clone();
        caches.open(CACHE).then(function (c) { c.put('./index.html', copy); });
        return res;
      }).catch(function () {
        return caches.match('./index.html').then(function (m) { return m || caches.match('./'); });
      })
    );
    return;
  }

  // version.json: always go straight to network, never cache — the in-app update
  // checker polls this to know if a newer build has been deployed, so a cached
  // (stale) answer would defeat the entire point.
  if (url.pathname.endsWith('/version.json')) {
    e.respondWith(fetch(req, { cache: 'no-store' }).catch(function () { return new Response('{}'); }));
    return;
  }

  // Other same-origin assets: cache-first, then network (and cache it for next time).
  e.respondWith(
    caches.match(req).then(function (cached) {
      return cached || fetch(req).then(function (res) {
        var copy = res.clone();
        caches.open(CACHE).then(function (c) { c.put(req, copy); });
        return res;
      }).catch(function () { return cached; });
    })
  );
});
