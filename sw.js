// 生成物 —— 由 tools/build.mjs 生成，请勿手改
const CACHE_NAME = 'exam-memo-8bc62a10';
const PRECACHE = [
  "./",
  "./index.html",
  "./app.js?v=1.8.0",
  "./style.css?v=1.8.0",
  "./manifest.webmanifest",
  "./icon.svg",
  "./icon-180.png",
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then((cache) => cache.addAll(PRECACHE))
      .then(() => self.skipWaiting()),
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((key) => key !== CACHE_NAME).map((key) => caches.delete(key))))
      .then(() => self.clients.claim()),
  );
});

self.addEventListener('fetch', (event) => {
  if (event.request.method !== 'GET') return;
  const request = event.request;
  // 导航请求 = 联网优先（US-41）：联网时打开就是最新版；断网回退缓存
  if (request.mode === 'navigate') {
    event.respondWith(
      fetch(request).then((response) => {
        const copy = response.clone();
        caches.open(CACHE_NAME).then((cache) => cache.put(request, copy));
        return response;
      }).catch(() => caches.match(request).then((hit) => hit || caches.match('./index.html'))),
    );
    return;
  }
  // 其余静态资源 = 缓存优先（命中即返回；未命中走网络并写回）
  event.respondWith(
    caches.match(request).then((hit) => hit || fetch(request).then((response) => {
      const copy = response.clone();
      caches.open(CACHE_NAME).then((cache) => cache.put(request, copy));
      return response;
    }).catch(() => caches.match('./index.html'))),
  );
});
