// „Къде ми е?“ — работа без интернет
const CACHE = 'kade-mi-e-v1';
const SHELL = ['./', './index.html', './manifest.webmanifest', './icon-192.png', './icon-512.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys()
    .then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;               // синхронизацията с Google не минава оттук
  const url = new URL(req.url);
  if (url.hostname.endsWith('google.com') || url.hostname.endsWith('googleusercontent.com')) return;
  const isFont = /fonts\.(googleapis|gstatic)\.com$/.test(url.hostname);
  if (isFont) { // шрифтове: първо от кеша
    e.respondWith(caches.match(req).then(hit => hit || fetch(req).then(res => {
      const copy = res.clone(); caches.open(CACHE).then(c => c.put(req, copy)); return res;
    })));
    return;
  }
  if (url.origin !== location.origin) return;
  // приложението: първо от мрежата (за да идват новите версии), иначе от кеша
  e.respondWith(fetch(req).then(res => {
    const copy = res.clone(); caches.open(CACHE).then(c => c.put(req, copy)); return res;
  }).catch(() => caches.match(req, { ignoreSearch: true }).then(hit => hit || caches.match('./index.html'))));
});
