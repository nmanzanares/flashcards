const CACHE_NAME = 'flashcards-v44'; // Cambia el nombre cada vez que edites la app
const ASSETS = [
    './',
    './index.html',
    './app.js',
    './manifest.json',
    './favicon.png',
    'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js',
    'https://cdn.jsdelivr.net/npm/epubjs/dist/epub.min.js'
];

// Instalar y guardar en caché
self.addEventListener('install', e => {
    e.waitUntil(
        caches.open(CACHE_NAME)
            .then(cache => {
                console.log('Cacheando archivos...');
                return cache.addAll(ASSETS);
            })
            .then(() => self.skipWaiting())
    );
});

// Limpiar cachés antiguas al activar
self.addEventListener('activate', e => {
    e.waitUntil(
        caches.keys().then(keys => {
            return Promise.all(
                keys.map(key => {
                    if (key !== CACHE_NAME) {
                        console.log('Borrando caché antigua:', key);
                        return caches.delete(key);
                    }
                })
            );
        }).then(() => self.clients.claim())
    );
});

// Estrategia: Cache First, luego Red
self.addEventListener('fetch', e => {
    // FILTRO CRÍTICO: Si la petición va a la API de Gemini, NO la interceptes con la caché
    if (e.request.url.includes('generativelanguage.googleapis.com')) {
        return; // El Service Worker se aparta y deja que el navegador use internet nativo 
    }
    e.respondWith(
        caches.match(e.request).then(res => {
            return res || fetch(e.request).catch(() => {
                // Si falla la red y no hay caché, podrías devolver una página offline aquí
                console.log('Error de red y sin caché para:', e.request.url);
            });
        })
    );
});

