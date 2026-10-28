/*
 * Service worker DataStaf Import Tool.
 *
 * - File aplikasi (index.html, app.js, style.css, dst.): JARINGAN DULU, cache hanya cadangan saat offline.
 *   Jadi setelah app.js diperbarui di GitHub Pages, aplikasi langsung memakai versi baru (bukan versi lama dari cache).
 * - Library SheetJS dari jsDelivr (versi dipatok @0.18.5): cache dulu supaya halaman tetap terbuka saat sinyal buruk.
 * - Permintaan POST ke Google Apps Script (proses impor) dan semua permintaan lain TIDAK disentuh sama sekali.
 *
 * Ubah CACHE_VERSION bila daftar file di ASSETS berubah, supaya cache lama dibersihkan.
 */
const CACHE_VERSION = 'datastaf-v1';
const SCOPE = self.registration.scope;

const ASSETS = [
  './',
  'index.html',
  'style.css',
  'app.js',
  'pwa.js',
  'manifest.json',
  'icons/icon-192.png',
  'icons/icon-512.png',
  'icons/icon-maskable-512.png',
  'icons/apple-touch-icon.png'
].map(path => new URL(path, SCOPE).href);

const CDN_HOST = 'cdn.jsdelivr.net';

self.addEventListener('install', event => {
  event.waitUntil((async () => {
    const cache = await caches.open(CACHE_VERSION);
    await cache.addAll(ASSETS);
    // Library CDN: dicoba, tetapi kegagalan tidak boleh membatalkan instalasi. no-cors -> respons opaque (boleh disimpan).
    try {
      const cdnScript = 'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js';
      const response = await fetch(new Request(cdnScript, { mode: 'no-cors' }));
      await cache.put(cdnScript, response);
    } catch (error) { /* abaikan */ }
    await self.skipWaiting();
  })());
});

self.addEventListener('activate', event => {
  event.waitUntil((async () => {
    const keys = await caches.keys();
    await Promise.all(keys.filter(key => key !== CACHE_VERSION).map(key => caches.delete(key)));
    await self.clients.claim();
  })());
});

self.addEventListener('fetch', event => {
  const request = event.request;
  if (request.method !== 'GET') return; // POST ke Apps Script dll. lewat langsung

  const url = new URL(request.url);

  if (url.origin === self.location.origin) {
    event.respondWith(networkFirst(request));
  } else if (url.hostname === CDN_HOST) {
    event.respondWith(cacheFirst(request));
  }
  // Selain itu (script.google.com, dll.): tidak di-intercept.
});

async function networkFirst(request) {
  const cache = await caches.open(CACHE_VERSION);
  try {
    // no-cache: selalu validasi ke server, tidak memakai cache HTTP bawaan GitHub Pages (10 menit).
    const response = await fetch(request, { cache: 'no-cache' });
    if (response && response.ok) cache.put(request, response.clone());
    return response;
  } catch (error) {
    const cached = await cache.match(request, { ignoreSearch: true });
    if (cached) return cached;
    if (request.mode === 'navigate') {
      const shell = await cache.match(new URL('index.html', SCOPE).href);
      if (shell) return shell;
    }
    throw error;
  }
}

async function cacheFirst(request) {
  const cache = await caches.open(CACHE_VERSION);
  const cached = await cache.match(request);
  if (cached) return cached;
  const response = await fetch(request);
  // ok = respons biasa; type "opaque" = skrip lintas-origin tanpa CORS (status 0), tetap valid disimpan.
  if (response && (response.ok || response.type === 'opaque')) cache.put(request, response.clone());
  return response;
}
