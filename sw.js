// QE Pharmacy Seminar Room Booking - Service Worker (PWA Offline Cache)
const CACHE_NAME = "qe-room-booking-v1.0.0";
const PRECACHE_ASSETS = [
  "./",
  "./index.html",
  "./styles.css",
  "./app.js",
  "./sync-engine.js",
  "./qrcode.js",
  "./xlsx.full.min.js",
  "./manifest.webmanifest",
  "./assets/icon-192.png",
  "./assets/icon-512.png",
  "./assets/apple-touch-icon.png"
];

// Install Event: Precaches the application shell
self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      return cache.addAll(PRECACHE_ASSETS).catch((err) => {
        console.warn("SW precache partial fallback:", err);
      });
    }).then(() => self.skipWaiting())
  );
});

// Activate Event: Cleans up obsolete caches
self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) => {
      return Promise.all(
        keys.map((key) => {
          if (key !== CACHE_NAME) {
            return caches.delete(key);
          }
        })
      );
    }).then(() => self.clients.claim())
  );
});

// Fetch Event: Stale-While-Revalidate for app shell, Network-only for cloud sync APIs
self.addEventListener("fetch", (event) => {
  const req = event.request;
  const url = new URL(req.url);

  // Bypass service worker for non-GET requests or Google Apps Script / Supabase API calls
  if (req.method !== "GET" || url.origin.includes("script.google.com") || url.origin.includes("supabase.co")) {
    return;
  }

  // Handle local application assets
  event.respondWith(
    caches.match(req).then((cachedResponse) => {
      if (cachedResponse) {
        // Fetch from network in background to update cache (Stale-While-Revalidate)
        fetch(req).then((networkResponse) => {
          if (networkResponse && networkResponse.status === 200) {
            caches.open(CACHE_NAME).then((cache) => cache.put(req, networkResponse));
          }
        }).catch(() => {
          // Offline, use cached
        });
        return cachedResponse;
      }

      // Not in cache: fetch from network and cache
      return fetch(req).then((networkResponse) => {
        if (!networkResponse || networkResponse.status !== 200 || networkResponse.type !== "basic") {
          return networkResponse;
        }
        const responseToCache = networkResponse.clone();
        caches.open(CACHE_NAME).then((cache) => {
          cache.put(req, responseToCache);
        });
        return networkResponse;
      }).catch(() => {
        // If offline and request is for page, return cached index.html
        if (req.headers.get("accept") && req.headers.get("accept").includes("text/html")) {
          return caches.match("./index.html");
        }
      });
    })
  );
});

