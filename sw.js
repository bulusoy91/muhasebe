// Service Worker: Network-First Strategy 
const cacheName = 'takvim-v5'; // Sürümü v5 yaptık, uygulama güncellenecek
const assets = ['./', './index.html', './app.js', './style.css'];

self.addEventListener('install', (e) => {
    e.waitUntil(caches.open(cacheName).then((cache) => cache.addAll(assets)));
});

// Eski cache'leri sil
self.addEventListener('activate', (e) => {
    e.waitUntil(
        caches.keys().then((keys) => {
            return Promise.all(keys.filter(k => k !== cacheName).map(k => caches.delete(k)));
        })
    );
});

self.addEventListener('fetch', (e) => {
    e.respondWith(
        fetch(e.request).catch(() => caches.match(e.request))
    );
});
