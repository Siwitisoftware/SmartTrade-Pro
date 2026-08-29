/*
 * SmartTrade Pro - Service Worker
 * ============================================
 * Goals:
 *  1. Let the app be "installed" (Add to Home Screen / desktop install).
 *  2. Keep the app usable OFFLINE (app shell cached).
 *  3. Always pick up new deployments automatically - a user who opens the
 *     app while online gets the newest version without needing to
 *     uninstall/reinstall or manually clear cache.
 *  4. NEVER interfere with calls to the Google Apps Script backend - the
 *     app already has its own offline queueing for that data, and it must
 *     always go straight to the network (or fail naturally) so that logic
 *     keeps working exactly as before.
 *
 * Bump CACHE_VERSION any time you want to force every installed client to
 * throw away old cached assets (e.g. after removing a file). You do NOT
 * need to bump it for normal content updates - the network-first strategy
 * below already fetches fresh HTML whenever the device is online.
 * ============================================
 */

const CACHE_VERSION = 'v7';
const CACHE_NAME = `smarttrade-cache-${CACHE_VERSION}`;

// Core "app shell" files to have ready immediately after install, so the
// very first load (even offline, right after installing the PWA) works.
// Adjust this list if your deployed file names differ.
const APP_SHELL = [
  './',
  './index.html',
  './manifest.json',
  './icon-192.png',
  './icon-512.png'
];

// Third-party origins that the app already handles its own way and that
// the Service Worker must always leave completely alone (pass straight
// through to the network, never cached, never intercepted):
//  - script.google.com / script.googleusercontent.com -> the Apps Script
//    backend (all business data read/write, session tokens, etc.)
//  - drive.google.com / googleusercontent.com (lh3...) -> uploaded
//    business logo and debtor photos served from Google Drive
const BYPASS_HOSTS = [
  'script.google.com',
  'script.googleusercontent.com',
  'drive.google.com',
  'googleusercontent.com',
  'docs.google.com'
];

// CDN-hosted static libraries (Font Awesome, etc.) - these rarely change,
// so a cache-first strategy keeps the app fast and fully usable offline
// once they've been fetched once.
const CDN_HOSTS = [
  'cdnjs.cloudflare.com',
  'fonts.googleapis.com',
  'fonts.gstatic.com'
];

self.addEventListener('install', (event) => {
  // Activate the new Service Worker immediately instead of waiting for
  // every open tab to close first - this is what makes updates roll out
  // right away rather than "on next full app restart".
  self.skipWaiting();

  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      // Best-effort: don't let one missing file block installation.
      return Promise.all(
        APP_SHELL.map((url) =>
          cache.add(url).catch((err) => {
            console.warn('[SW] Could not pre-cache', url, err);
          })
        )
      );
    })
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    (async () => {
      // Remove any caches from previous versions of this Service Worker.
      const names = await caches.keys();
      await Promise.all(
        names
          .filter((name) => name.startsWith('smarttrade-cache-') && name !== CACHE_NAME)
          .map((name) => caches.delete(name))
      );

      // Take control of any already-open tabs immediately, without
      // requiring a reload, so the update takes effect right away.
      await self.clients.claim();

      // Let every open tab know a fresh version is now active, in case
      // the page wants to show a "reload to refresh" toast (optional -
      // see the SW_ACTIVATED message handled in index.html, if present).
      const clients = await self.clients.matchAll({ type: 'window' });
      clients.forEach((client) => client.postMessage({ type: 'SW_ACTIVATED', version: CACHE_VERSION }));
    })()
  );
});

// Lets the page force an immediate update check/activation on demand,
// e.g. from an "Update available" button, by calling:
//   navigator.serviceWorker.controller.postMessage({ type: 'SKIP_WAITING' })
self.addEventListener('message', (event) => {
  if (event.data && event.data.type === 'SKIP_WAITING') {
    self.skipWaiting();
  }
});

self.addEventListener('fetch', (event) => {
  const request = event.request;

  // Only ever handle simple GET requests - never touch POST/PUT/etc,
  // which is exactly how the app talks to the Apps Script backend.
  if (request.method !== 'GET') {
    return;
  }

  const url = new URL(request.url);

  // Never intercept the Apps Script backend or Google Drive-hosted
  // images - always go straight to the network for these, untouched.
  if (BYPASS_HOSTS.some((host) => url.hostname.includes(host))) {
    return;
  }

  // CDN libraries: cache-first (fast, and still available offline once
  // fetched at least once).
  if (CDN_HOSTS.some((host) => url.hostname.includes(host))) {
    event.respondWith(cacheFirst(request));
    return;
  }

  // Everything else same-origin (the app shell itself: the HTML page,
  // manifest, icons): network-first, so an online user always gets the
  // latest deployed version, while an offline user still gets the last
  // good copy instead of an error page.
  event.respondWith(networkFirst(request));
});

async function networkFirst(request) {
  const cache = await caches.open(CACHE_NAME);
  try {
    const freshResponse = await fetch(request);
    // Only cache good, basic (same-origin) responses.
    if (freshResponse && freshResponse.ok) {
      cache.put(request, freshResponse.clone());
    }
    return freshResponse;
  } catch (err) {
    const cached = await cache.match(request, { ignoreSearch: true });
    if (cached) return cached;

    // Last resort for a full-page navigation while offline with nothing
    // cached yet - fall back to whatever app shell page we do have.
    if (request.mode === 'navigate') {
      const shell = await cache.match('./index.html');
      if (shell) return shell;
    }
    throw err;
  }
}

async function cacheFirst(request) {
  const cache = await caches.open(CACHE_NAME);
  const cached = await cache.match(request);
  if (cached) return cached;

  try {
    const freshResponse = await fetch(request);
    if (freshResponse && freshResponse.ok) {
      cache.put(request, freshResponse.clone());
    }
    return freshResponse;
  } catch (err) {
    throw err;
  }
}
