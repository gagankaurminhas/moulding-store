const C="ms-v4-light-luxury";
const A=["./","./index.html","./styles.css","./app.js","./config.js","./manifest.webmanifest","./assets/icon.svg"];
self.addEventListener("install",event=>event.waitUntil(caches.open(C).then(cache=>cache.addAll(A)).then(()=>self.skipWaiting())));
self.addEventListener("activate",event=>event.waitUntil(caches.keys().then(keys=>Promise.all(keys.filter(k=>k!==C).map(k=>caches.delete(k)))).then(()=>self.clients.claim())));
self.addEventListener("fetch",event=>event.respondWith(caches.match(event.request).then(cached=>cached||fetch(event.request))));
