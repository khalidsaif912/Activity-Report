/* Activity Report PWA — cache static assets; always fetch JSON fresh. */
const CACHE_VERSION = "activity-report-pwa-v1";

const PRECACHE_URLS = [
  "/data/report/offload_report.html",
  "/index.html",
  "/manifest.webmanifest",
  "/assets/icons/icon-192.png",
  "/assets/icons/icon-512.png",
  "/js/offload-loader.js",
  "/js/flight-hint-cache.js",
  "/js/recipients-cache.js",
  "/js/manpower-role-hint-cache.js",
  "/js/manpower-autocomplete.js",
  "/js/flight-autocomplete.js",
  "/js/csd-route-hint-cache.js",
  "/js/phrase-usage-cache.js",
  "/js/phrase-autocomplete.js",
  "/js/activity-report-app.js",
  "/js/pwa-register.js",
  "/assets/export-roster-plane.png",
  "/assets/transom-logo.png"
];

function isJsonRequest(url) {
  return url.pathname.endsWith(".json");
}

function isNavigation(request) {
  return request.mode === "navigate";
}

async function networkFirst(request) {
  try {
    const response = await fetch(request);
    return response;
  } catch (err) {
    const cached = await caches.match(request);
    if (cached) return cached;
    throw err;
  }
}

async function staleWhileRevalidate(request) {
  const cache = await caches.open(CACHE_VERSION);
  const cached = await cache.match(request);
  const networkPromise = fetch(request)
    .then((response) => {
      if (response && response.ok) {
        cache.put(request, response.clone());
      }
      return response;
    })
    .catch(() => null);

  if (cached) {
    networkPromise.catch(() => {});
    return cached;
  }

  const network = await networkPromise;
  if (network) return network;
  throw new Error("Offline and no cache for " + request.url);
}

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches
      .open(CACHE_VERSION)
      .then((cache) => cache.addAll(PRECACHE_URLS))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches
      .keys()
      .then((keys) =>
        Promise.all(keys.filter((key) => key !== CACHE_VERSION).map((key) => caches.delete(key)))
      )
      .then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", (event) => {
  const request = event.request;
  if (request.method !== "GET") return;

  const url = new URL(request.url);
  if (url.origin !== self.location.origin) return;

  if (isJsonRequest(url)) {
    event.respondWith(networkFirst(request));
    return;
  }

  if (isNavigation(request)) {
    event.respondWith(
      fetch(request).catch(async () => {
        const cached =
          (await caches.match("/data/report/offload_report.html")) ||
          (await caches.match("/index.html"));
        return cached || Response.error();
      })
    );
    return;
  }

  event.respondWith(staleWhileRevalidate(request));
});
