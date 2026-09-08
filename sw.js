const CACHE_NAME = 'esta-110sppc-v32';
const ASSETS = [
  './',
  './index.html',
  './data.json',
  './manifest.json',
  './icon-192.png',
  './icon-512.png',
  './icon-shield.png',
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then((cache) => cache.addAll(ASSETS))
      .catch(() => {})
  );
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)))
    )
  );
  self.clients.claim();
});

// Ağ varsa her zaman ağdan çek (tarayıcının kendi HTTP önbelleğini de
// atlayarak) ve en güncel dosyayı önbelleğe yaz. Ağ yoksa (çevrimdışı),
// son bilinen (önbellekteki) sürümü göster. Böylece kullanıcı hiçbir
// zaman "site verilerini temizle" yapmak zorunda kalmaz — bu da parola
// hatırlama kaydını (localStorage) korumuş olur.
self.addEventListener('fetch', (event) => {
  if (event.request.method !== 'GET') return;
  event.respondWith(
    fetch(event.request, { cache: 'no-store' })
      .then((res) => {
        const resClone = res.clone();
        caches.open(CACHE_NAME).then((cache) => cache.put(event.request, resClone));
        return res;
      })
      .catch(() => caches.match(event.request))
  );
});
