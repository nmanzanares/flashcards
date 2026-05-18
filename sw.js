const CACHE_NAME = 'flashcards-v58'; // Cambia el nombre cada vez que edites la app
const ASSETS = [
    './',
    './index.html',
    './style.css',
    './app.js',
    './manifest.json',
    './favicon.png',
    /*'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js',
    'https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js',
    'https://cdn.jsdelivr.net/npm/epubjs/dist/epub.min.js',
    'https://cdn.jsdelivr.net/npm/pdfjs-dist@4.4.168/build/pdf.min.js'
    */
];

self.addEventListener('install', e => {
    e.waitUntil(
        caches.open(CACHE_NAME)
            .then(cache => cache.addAll(ASSETS))
            .then(() => self.skipWaiting())
    );
});

self.addEventListener('activate', e => {
    e.waitUntil(
        caches.keys().then(keys => {
            return Promise.all(
                keys.map(key => {
                    if (key !== CACHE_NAME) {
                        return caches.delete(key);
                    }
                })
            );
        }).then(() => self.clients.claim())
    );
});

self.addEventListener('fetch', e => {
    // No interceptar Gemini
    if (e.request.url.includes('generativelanguage.googleapis.com')) {
        return;
    }

    // No cachear CDN dinámico
    if (e.request.url.includes('cdn.jsdelivr.net') ||
        e.request.url.includes('cdnjs.cloudflare.com')) {

        e.respondWith(fetch(e.request));
        return;
    }

    e.respondWith(
        caches.match(e.request).then(res => {
            return res || fetch(e.request)
                .then(networkRes => {
                    return caches.open(CACHE_NAME).then(cache => {
                        cache.put(e.request, networkRes.clone());
                        return networkRes;
                    });
                });
        })
    );
});
