/* 金拾煎 PWA Service Worker：離線可開啟首頁，外部資源（字體、Tailwind、Three.js、圖片）邊用邊快取 */
const VERSION = 'v1';
const SHELL = `shell-${VERSION}`;
const RUNTIME = `runtime-${VERSION}`;
const SHELL_FILES = ['./', './index.html', './manifest.webmanifest',
  './icons/icon-192.png', './icons/icon-512.png', './icons/icon-maskable-512.png', './icons/apple-touch-icon.png'];

self.addEventListener('install', (e) => {
  /* 逐一快取，某個檔案缺少時不會讓整個安裝失敗 */
  e.waitUntil(caches.open(SHELL)
    .then((c) => Promise.all(SHELL_FILES.map((f) => c.add(f).catch(() => null))))
    .then(() => self.skipWaiting()));
});

self.addEventListener('activate', (e) => {
  e.waitUntil(caches.keys()
    .then((keys) => Promise.all(keys.filter((k) => k !== SHELL && k !== RUNTIME).map((k) => caches.delete(k))))
    .then(() => self.clients.claim()));
});

self.addEventListener('fetch', (e) => {
  const req = e.request;
  if (req.method !== 'GET') return;

  /* 頁面：先抓最新版，離線時用快取 */
  if (req.mode === 'navigate') {
    e.respondWith(fetch(req).then((res) => {
      const copy = res.clone();
      caches.open(SHELL).then((c) => c.put('./index.html', copy));
      return res;
    }).catch(() => caches.match('./index.html')));
    return;
  }

  /* 其他資源：先用快取立即顯示，同時在背景更新 */
  e.respondWith(caches.match(req).then((cached) => {
    const fresh = fetch(req).then((res) => {
      if (res && (res.ok || res.type === 'opaque')) {
        const copy = res.clone();
        caches.open(RUNTIME).then((c) => c.put(req, copy));
      }
      return res;
    }).catch(() => cached);
    return cached || fresh;
  }));
});
