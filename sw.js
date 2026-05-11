// SheetFilter Service Worker v3
// Caches app shell for offline use + handles Web Share Target POST requests.

const CACHE = 'sheetfilter-v4';

const PRECACHE = [
  './index.html',
  './manifest.json',
  'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js',
  'https://fonts.googleapis.com/css2?family=DM+Mono:wght@400;500&family=Sora:wght@300;400;500;600&display=swap'
];

// ── Install: cache app shell ───────────────────────────────────────────────
self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE).then(cache => {
      return cache.addAll(['./index.html', './manifest.json']).then(() => {
        return Promise.allSettled(
          PRECACHE.slice(2).map(url =>
            fetch(url, { mode: 'cors' })
              .then(res => res.ok ? cache.put(url, res) : null)
              .catch(() => null)
          )
        );
      });
    }).then(() => self.skipWaiting())
  );
});

// ── Activate: remove old caches ────────────────────────────────────────────
self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))
    ).then(() => self.clients.claim())
  );
});

// ── Fetch: intercept share target POST + cache-first for everything else ───
self.addEventListener('fetch', event => {
  const url = new URL(event.request.url);

  // ── Share Target: browser POSTs the shared file here ──────────────────
  // Triggered when user picks SheetFilter from the OS share sheet.
  if (
    event.request.method === 'POST' &&
    url.pathname.endsWith('index.html')
  ) {
    event.respondWith(
      (async () => {
        const formData = await event.request.formData();
        const file = formData.get('file');           // matches "name" in manifest share_target

        if (file && file instanceof File) {
          // Read the file as ArrayBuffer so we can pass it to the page
          const arrayBuffer = await file.arrayBuffer();

          // Find all open windows/tabs for this SW scope and send the file
          const allClients = await self.clients.matchAll({
            includeUncontrolled: true,
            type: 'window'
          });

          // Redirect to index.html first (GET), then postMessage once the
          // client signals it is ready — or post to whichever client opens.
          // We store the file in a short-lived SW-level variable so the
          // redirected page can claim it immediately on load.
          self._sharedFile = { name: file.name, buffer: arrayBuffer };

          // If a window is already open, send it right now
          for (const client of allClients) {
            client.postMessage({
              type: 'SHARE_TARGET_FILE',
              name: file.name,
              buffer: arrayBuffer
            }, [arrayBuffer.slice(0)]);   // transfer a copy, keep original for redirect
          }
        }

        // Always redirect to the app after receiving the share
        return Response.redirect('./index.html', 303);
      })()
    );
    return;
  }

  // Skip non-GET beyond this point
  if (event.request.method !== 'GET') return;
  if (!url.protocol.startsWith('http')) return;

  // ── Cache-first for everything else ─────────────────────────────────────
  event.respondWith(
    caches.match(event.request).then(cached => {
      if (cached) return cached;
      return fetch(event.request).then(response => {
        if (
          response.ok &&
          (url.origin === self.location.origin ||
           url.hostname === 'cdnjs.cloudflare.com' ||
           url.hostname === 'fonts.googleapis.com' ||
           url.hostname === 'fonts.gstatic.com')
        ) {
          const clone = response.clone();
          caches.open(CACHE).then(cache => cache.put(event.request, clone));
        }
        return response;
      }).catch(() => {
        if (event.request.mode === 'navigate') {
          return caches.match('./index.html');
        }
      });
    })
  );
});

// ── Message relay: page says "ready", SW sends the pending shared file ─────
self.addEventListener('message', event => {
  if (event.data && event.data.type === 'SHARE_TARGET_READY') {
    if (self._sharedFile) {
      event.source.postMessage({
        type: 'SHARE_TARGET_FILE',
        name: self._sharedFile.name,
        buffer: self._sharedFile.buffer
      }, [self._sharedFile.buffer]);
      self._sharedFile = null;
    }
  }
});
