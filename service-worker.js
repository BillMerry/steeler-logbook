// service-worker.js
// Simple offline-first cache for STEELER Logbook

const CACHE_NAME = "steeler-logbook-v1";

const ASSETS = [
  "./",
  "./index.html",
  "./styles.css",
  "./app.js",
  "./manifest.json"
  // Later we can add:
  // "./icons/steeler-192.png",
  // "./icons/steeler-512.png"
];

// Install: cache the app shell
self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(ASSETS))
  );
  self.skipWaiting();
});

// Activate: clean up old caches if we bump CACHE_NAME
self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(
        keys.map((key) => {
          if (key !== CACHE_NAME) {
            return caches.delete(key);
          }
        })
      )
    )
  );
  self.clients.claim();
});

// Fetch: cache-first for app shell, network-first fallback for others
self.addEventListener("fetch", (event) => {
  const request = event.request;
  if (request.method !== "GET") return;

  event.respondWith(
    caches.match(request).then((cached) => {
      if (cached) {
        // Serve from cache, try to update in background
        fetch(request).catch(() => {});
        return cached;
      }

      // Not cached: try network, fall back to index.html for navigation
      return fetch(request).catch(() => {
        if (request.mode === "navigate") {
          return caches.match("./index.html");
        }
        return new Response("Offline", {
          status: 503,
          statusText: "Offline"
        });
      });
    })
  );
});
