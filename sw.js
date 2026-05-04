const CACHE_NAME = 'v1_flashcards';
const ASSETS = [
    './',
    './index.html',
    './app.js',
    './manifest.json',
    'https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js'
];

// Instalar y guardar en caché
self.addEventListener('install', e => {
    e.waitUntil(
        caches.open(CACHE_NAME).then(cache => cache.addAll(ASSETS))
    );
});

// Interceptar peticiones para usar la caché si no hay red
self.addEventListener('fetch', e => {
    e.respondWith(
        caches.match(e.request).then(res => res || fetch(e.request))
    );
});
