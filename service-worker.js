/*
 * SmartTrade Pro - Service Worker
 * ============================================
 * Goals:
 *  1. Let the app be "installed" (Add to Home Screen / desktop install).
 *  2. Keep the app usable OFFLINE (app shell cached).
 *  3. Always pick up new deployments automatically.
 *  4. NEVER interfere with calls to the Google Apps Script backend.
 *
 * Cache version is intentionally preserved. Normal deployments use
 * network-first for same-origin app files; bump only when old cached assets
 * must be forcibly discarded.
 */

const CACHE_VERSION = 'v23';
const CACHE_NAME = `smarttrade-cache-${CACHE_VERSION}`;

const APP_SHELL = [
  './',
  './index.html',
  './manifest.json',
  './icon-192.png',
  './icon-512.png'
];

const BYPASS_HOSTS = [
  'script.google.com',
  'script.googleusercontent.com',
  'drive.google.com',
  'googleusercontent.com',
  'docs.google.com'
];

const CDN_HOSTS = [
  'cdnjs.cloudflare.com',
  'fonts.googleapis.com',
  'fonts.gstatic.com'
];

function hostMatches(hostname, allowedHost) {
  return hostname === allowedHost || hostname.endsWith(`.${allowedHost}`);
}

function matchesAnyHost(hostname, hosts) {
  return hosts.some((host) => hostMatches(hostname, host));
}

self.addEventListener('install', (event) => {
  event.waitUntil((async () => {
    const cache = await caches.open(CACHE_NAME);
    await Promise.all(APP_SHELL.map(async (url) => {
      try {
        const response = await fetch(url, { cache: 'reload' });
        if (response && response.ok && response.type === 'basic') {
          await cache.put(url, response);
        }
      } catch (err) {
        // Best effort: one missing shell asset must not block installation.
        console.warn('[SW] Could not pre-cache', url, err);
      }
    }));
    await self.skipWaiting();
  })());
});

self.addEventListener('activate', (event) => {
  event.waitUntil((async () => {
    const names = await caches.keys();
    await Promise.all(
      names
        .filter((name) => name.startsWith('smarttrade-cache-') && name !== CACHE_NAME)
        .map((name) => caches.delete(name))
    );

    await self.clients.claim();

    const clients = await self.clients.matchAll({ type: 'window' });
    clients.forEach((client) => {
      client.postMessage({ type: 'SW_ACTIVATED', version: CACHE_VERSION });
    });
  })());
});

self.addEventListener('message', (event) => {
  if (event.data && event.data.type === 'SKIP_WAITING') {
    event.waitUntil(self.skipWaiting());
  }
});

self.addEventListener('fetch', (event) => {
  const request = event.request;

  // Only GET requests are eligible for caching. POST/PUT/etc. pass through.
  if (request.method !== 'GET') return;

  const url = new URL(request.url);

  // Never intercept backend calls, Drive files, or other explicitly bypassed hosts.
  if (matchesAnyHost(url.hostname, BYPASS_HOSTS)) return;

  // Only cache known CDN assets and this app's own origin. Other external
  // requests pass through untouched (e.g. external images or integrations).
  if (matchesAnyHost(url.hostname, CDN_HOSTS)) {
    event.respondWith(staleWhileRevalidate(request, event));
    return;
  }

  if (url.origin !== self.location.origin) return;

  event.respondWith(networkFirst(request));
});

async function networkFirst(request) {
  const cache = await caches.open(CACHE_NAME);

  try {
    const freshResponse = await fetch(request);
    if (freshResponse && freshResponse.ok && freshResponse.type === 'basic') {
      // Await the cache write so the latest successful response is safely stored.
      await cache.put(request, freshResponse.clone());
    }
    return freshResponse;
  } catch (err) {
    const cached = await cache.match(request, { ignoreSearch: true });
    if (cached) return cached;

    if (request.mode === 'navigate') {
      const shell = await cache.match('./index.html');
      if (shell) return shell;
    }
    throw err;
  }
}

// CDN cache hits return immediately for speed; the cached copy is refreshed
// in the background so libraries can still update without delaying page load.
async function staleWhileRevalidate(request, event) {
  const cache = await caches.open(CACHE_NAME);
  const cached = await cache.match(request);

  const refreshPromise = fetch(request).then(async (freshResponse) => {
    if (freshResponse && freshResponse.ok &&
        (freshResponse.type === 'basic' || freshResponse.type === 'cors')) {
      await cache.put(request, freshResponse.clone());
    }
    return freshResponse;
  });

  if (cached) {
    event.waitUntil(refreshPromise.then(() => undefined).catch((err) => {
      console.warn('[SW] CDN refresh deferred:', request.url, err);
    }));
    return cached;
  }

  try {
    return await refreshPromise;
  } catch (err) {
    throw err;
  }
}
