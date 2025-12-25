// Service Worker สำหรับ PWA
const CACHE_NAME = 'maid-inspection-v2.1';
const urlsToCache = [
    './',
    './index.html',
    './manifest.json',
    'https://cdn.tailwindcss.com',
    'https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css',
    'https://fonts.googleapis.com/css2?family=Noto+Sans+Thai:wght@300;400;500;600;700&display=swap'
];

// ติดตั้ง Service Worker
self.addEventListener('install', event => {
    event.waitUntil(
        caches.open(CACHE_NAME)
            .then(cache => {
                console.log('Opened cache');
                return cache.addAll(urlsToCache);
            })
            .then(() => self.skipWaiting())
    );
});

// เปิดใช้งาน Service Worker
self.addEventListener('activate', event => {
    event.waitUntil(
        caches.keys().then(cacheNames => {
            return Promise.all(
                cacheNames.map(cacheName => {
                    if (cacheName !== CACHE_NAME) {
                        console.log('Deleting old cache:', cacheName);
                        return caches.delete(cacheName);
                    }
                })
            );
        }).then(() => self.clients.claim())
    );
});

// ดึงข้อมูลจาก cache หรือ network
self.addEventListener('fetch', event => {
    // ข้ามสำหรับ Google Script requests
    if (event.request.url.includes('script.google.com')) {
        return;
    }
    
    event.respondWith(
        caches.match(event.request)
            .then(response => {
                // ถ้ามีใน cache ให้ใช้ cache
                if (response) {
                    return response;
                }
                
                // ถ้าไม่มีให้เรียกจาก network
                return fetch(event.request)
                    .then(response => {
                        // ตรวจสอบว่า response ถูกต้อง
                        if (!response || response.status !== 200 || response.type !== 'basic') {
                            return response;
                        }
                        
                        // Clone response สำหรับ cache
                        const responseToCache = response.clone();
                        
                        caches.open(CACHE_NAME)
                            .then(cache => {
                                cache.put(event.request, responseToCache);
                            });
                        
                        return response;
                    })
                    .catch(() => {
                        // ถ้า offline และเป็น HTML request ให้แสดง offline page
                        if (event.request.headers.get('accept').includes('text/html')) {
                            return caches.match('./index.html');
                        }
                    });
            })
    );
});

// รับข้อความจาก client
self.addEventListener('message', event => {
    if (event.data.action === 'skipWaiting') {
        self.skipWaiting();
    }
});