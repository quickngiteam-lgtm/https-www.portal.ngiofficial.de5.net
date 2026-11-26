const CACHE = "ngi-shell-v3";
const SHELL = ["./", "index.html", "manifest.json", "icons/app-icon-192.png", "icons/app-icon-512.png"];
const SHELL_URLS = new Set(SHELL.map(path => new URL(path, self.registration.scope).href));

self.addEventListener("install", event => {
  event.waitUntil(caches.open(CACHE).then(cache => cache.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener("activate", event => {
  event.waitUntil(caches.keys().then(keys => Promise.all(
    keys.filter(key => key.startsWith("ngi-shell-") && key !== CACHE).map(key => caches.delete(key))
  )).then(() => self.clients.claim()));
});

self.addEventListener("fetch", event => {
  if (event.request.method !== "GET") return;
  const url = new URL(event.request.url);
  url.search = "";
  if (!SHELL_URLS.has(url.href)) return;

  event.respondWith(loadShellAsset(event.request, url.href));
});

async function loadShellAsset(request, cacheUrl) {
  const response = await fetch(request).catch(() => undefined);
  if (response && response.status < 500) return response;

  const cache = await caches.open(CACHE);
  return (await cache.match(cacheUrl)) || response || Response.error();
}
