// Service Worker สำหรับระบบตรวจงานแม่บ้าน
const CACHE_NAME = 'maid-inspection-v1';
const urlsToCache = [
  './',
  './index.html',
  './manifest.json'
];

// Install event
self.addEventListener('install', event => {
  console.log('Service Worker: Installing...');
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(cache => {
        console.log('Service Worker: Caching files');
        return cache.addAll(urlsToCache);
      })
      .then(() => self.skipWaiting())
  );
});

// Activate event
self.addEventListener('activate', event => {
  console.log('Service Worker: Activated');
  event.waitUntil(
    caches.keys().then(cacheNames => {
      return Promise.all(
        cacheNames.map(cache => {
          if (cache !== CACHE_NAME) {
            console.log('Service Worker: Clearing old cache');
            return caches.delete(cache);
          }
        })
      );
    })
  );
});

// Fetch event
self.addEventListener('fetch', event => {
  // Skip non-GET requests
  if (event.request.method !== 'GET') return;
  
  // Network first, then cache
  event.respondWith(
    fetch(event.request)
      .then(response => {
        // Cache the response
        const responseToCache = response.clone();
        caches.open(CACHE_NAME)
          .then(cache => {
            cache.put(event.request, responseToCache);
          });
        return response;
      })
      .catch(() => {
        // If network fails, return from cache
        return caches.match(event.request)
          .then(response => {
            if (response) {
              return response;
            }
            // If not in cache, return offline page
            return caches.match('./index.html');
          });
      })
  );
});