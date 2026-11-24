try { importScripts('https://cdn.onesignal.com/sdks/web/v16/OneSignalSDK.sw.js'); } catch (e) {}   // push (OneSignal); без интернета оболочка работает как раньше
// Кэшируется только «оболочка» приложения. Данные таблиц (Apps Script) не кэшируются здесь —
// они хранятся в браузере самим приложением и обновляются раз в 30 минут.
const V = 'pro-shell-v3';
const SHELL = ['./', 'index.html', 'manifest.webmanifest', 'icons/icon-192.png', 'icons/icon-512.png', 'icons/apple-touch-icon.png'];
self.addEventListener('install', e => { e.waitUntil(caches.open(V).then(c => c.addAll(SHELL)).then(() => self.skipWaiting())); });
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== V).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
// «Сначала кэш, обновление в фоне»: страница открывается мгновенно из кэша, свежая версия скачивается тихо и применяется при следующем открытии (или по кнопке в баннере)
let SHELL_NEW = Promise.resolve(false);                 // закончилась ли проверка обновления оболочки и нашлась ли новая версия
self.addEventListener('fetch', e => {
  const r = e.request, u = new URL(r.url);
  if (r.method !== 'GET' || u.origin !== location.origin) return;     // скрипт таблиц и CDN — мимо кэша
  const isShell = r.mode === 'navigate' || /index\.html$|\/$/.test(u.pathname);
  e.respondWith(caches.open(V).then(async c => {
    const hit = await c.match(r, { ignoreSearch: r.mode === 'navigate' });
    let changed = false;
    const net = fetch(r, { cache: 'no-cache' }).then(res => {
      if (res && res.ok){
        const sig = x => x ? (x.headers.get('etag') || x.headers.get('last-modified') || x.headers.get('content-length') || '') : '';
        changed = !!hit && sig(hit) !== sig(res);
        c.put(r, res.clone());
      }
      return res;
    }).catch(() => null);
    if (isShell) SHELL_NEW = net.then(() => changed);
    if (hit) { e.waitUntil(net); return hit; }
    return (await net) || (await caches.match('index.html')) || Response.error();
  }));
});
self.addEventListener('message', e => { if (e.data === 'hello' && e.source) e.waitUntil(SHELL_NEW.then(ch => { if (ch) e.source.postMessage('shell-updated'); })); });
self.addEventListener('notificationclick', e => { e.notification.close(); e.waitUntil(self.clients.matchAll({ type: 'window', includeUncontrolled: true }).then(l => l.length ? l[0].focus() : self.clients.openWindow('./'))); });
