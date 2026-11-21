// לוז התקנות: network-first. הנתונים (Supabase) לא נשמרים במטמון, רק קבצי הדף לפתיחה בלי רשת.
const CACHE = "lz-1.15";
const SHELL = ["./", "./index.html", "./manifest.json?v=14", "./icon-192-v2.png", "./icon-512-v2.png"];
self.addEventListener("install", e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).then(() => self.skipWaiting())); });
self.addEventListener("activate", e => { e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim())); });
self.addEventListener("fetch", e => {
  const u = new URL(e.request.url);
  if (e.request.method !== "GET" || u.origin !== location.origin) return;   // Supabase ו-CDN: ישירות לרשת
  e.respondWith(fetch(e.request).then(r => { const copy = r.clone(); caches.open(CACHE).then(c => c.put(e.request, copy)); return r; })
    .catch(() => caches.match(e.request).then(m => m || caches.match("./index.html"))));
});
