const CACHE_NAME = 'flashcards-v13'; // Cambia el nombre cada vez que edites la app
const ASSETS = [
    './',
    './index.html',
    './app.js',
    './manifest.json',
    './favicon.png', // Añadido para evitar error 404 en el registro
    'https://jsdelivr.net'
];

// Instalar y guardar en caché
self.addEventListener('install', e => {
    e.waitUntil(
        caches.open(CACHE_NAME).then(cache => cache.addAll(ASSETS))
    );
    self.skipWaiting();
});

// Limpiar cachés antiguas al activar
self.addEventListener('activate', e => {
    e.waitUntil(
        caches.keys().then(keys => {
            return Promise.all(
                keys.filter(key => key !== CACHE_NAME).map(key => caches.delete(key))
            );
        })
    );
});

// Estrategia: Cache First, luego Red
self.addEventListener('fetch', e => {
    e.respondWith(
        caches.match(e.request).then(res => {
            return res || fetch(e.request);
        })
    );
});
