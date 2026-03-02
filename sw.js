const CACHE_NAME = 'paradise-vendas-v2';

const STATIC_ASSETS = [
    './',
    './index.html',
    './css/styles.css',
    './js/app.js',
    './manifest.json',
    './icons/icon-192.svg',
    './icons/icon-512.svg',
    'https://cdn.jsdelivr.net/npm/chart.js@4'
];

// Install: cache static assets
self.addEventListener('install', event => {
    event.waitUntil(
        caches.open(CACHE_NAME).then(cache => cache.addAll(STATIC_ASSETS))
    );
    self.skipWaiting();
});

// Activate: clean old caches
self.addEventListener('activate', event => {
    event.waitUntil(
        caches.keys().then(keys =>
            Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k)))
        )
    );
    self.clients.claim();
});

// Fetch: cache-first for static, network-first for Google Sheets data
self.addEventListener('fetch', event => {
    const url = event.request.url;

    // Google Sheets API requests: network-first (always try fresh data)
    if (url.includes('docs.google.com/spreadsheets')) {
        event.respondWith(
            fetch(event.request)
                .then(response => {
                    const clone = response.clone();
                    caches.open(CACHE_NAME).then(cache => cache.put(event.request, clone));
                    return response;
                })
                .catch(() => caches.match(event.request))
        );
        return;
    }

    // Static assets: cache-first
    event.respondWith(
        caches.match(event.request).then(cached => cached || fetch(event.request))
    );
});
