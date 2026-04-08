/* ===== Service Worker: cache offline + control total ===== */

const CACHE_NAME = "asistencia-v2"; // ⬅️ sube versión cuando cambies algo
const ASSETS = [
  "./",
  "./index.html",
  "./styles.css",
  "./app.js",
  "./manifest.webmanifest",
  "./icon-192.png",
  "./icon-512.png",
  "./apple-touch-icon.png",
  "./favicon-32.png",
  "./favicon-16.png"
];

/* ===== INSTALL ===== */
self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache => cache.addAll(ASSETS))
  );

  // 🔥 fuerza activación inmediata
  self.skipWaiting();
});

/* ===== ACTIVATE ===== */
self.addEventListener("activate", (event) => {
  event.waitUntil(
    (async () => {
      // 🧹 borrar TODAS las caches viejas
      const keys = await caches.keys();
      await Promise.all(
        keys.map(k => (k !== CACHE_NAME ? caches.delete(k) : null))
      );

      // 🧠 tomar control inmediato de todas las pestañas
      await self.clients.claim();
    })()
  );
});

/* ===== FETCH ===== */
self.addEventListener("fetch", (event) => {
  const req = event.request;

  // Solo GET
  if (req.method !== "GET") return;

  // ✅ Evitar cachear esquemas raros (chrome-extension:, data:, blob:, etc.)
  const url = new URL(req.url);
  if (url.protocol !== "http:" && url.protocol !== "https:") return;

  // ✅ NO cachear Google Apps Script (evita errores CORS/opaque y ensuciar cache)
  if (url.hostname.includes("script.google.com")) {
    // solo pasa directo a red
    event.respondWith(fetch(req));
    return;
  }

  event.respondWith(
    caches.match(req).then((cached) => {
      return (
        cached ||
        fetch(req)
          .then((res) => {
            // ✅ Solo cachear respuestas válidas y del mismo origen (más seguro)
            if (!res || !res.ok) return res;

            // Si la respuesta es opaque (no-cors), mejor no cachearla
            if (res.type === "opaque") return res;

            const copy = res.clone();
            caches.open(CACHE_NAME).then((cache) => cache.put(req, copy));
            return res;
          })
          .catch(() => cached)
      );
    })
  );
});


/* ===== KILL SWITCH (desde la app) ===== */
self.addEventListener("message", (event) => {
  if (event.data === "KILL_SW") {
    self.registration.unregister().then(() => {
      self.clients.matchAll({ includeUncontrolled: true }).then(clients => {
        clients.forEach(client => {
          // 🔄 fuerza recarga sin SW
          client.navigate(client.url);
        });
      });
    });
  }
});


