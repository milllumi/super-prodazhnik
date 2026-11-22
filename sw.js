try { importScripts('https://cdn.onesignal.com/sdks/web/v16/OneSignalSDK.sw.js'); } catch (e) {}   // push (OneSignal); без интернета оболочка работает как раньше
// Кэшируется только «оболочка» приложения. Данные таблиц (Apps Script) не кэшируются здесь —
// они хранятся в браузере самим приложением и обновляются раз в 30 минут.
const V = 'pro-shell-v2';
const SHELL = ['./', 'index.html', 'manifest.webmanifest', 'icons/icon-192.png', 'icons/icon-512.png', 'icons/apple-touch-icon.png'];
self.addEventListener('install', e => { e.waitUntil(caches.open(V).then(c => c.addAll(SHELL)).then(() => self.skipWaiting())); });
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== V).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const r = e.request, u = new URL(r.url);
  if (r.method !== 'GET' || u.origin !== location.origin) return;     // скрипт таблиц и CDN — мимо кэша
  e.respondWith(fetch(r).then(res => { const cp = res.clone(); caches.open(V).then(c => c.put(r, cp)); return res; }).catch(() => caches.match(r).then(m => m || caches.match('index.html'))));
});
self.addEventListener('notificationclick', e => { e.notification.close(); e.waitUntil(self.clients.matchAll({ type: 'window', includeUncontrolled: true }).then(l => l.length ? l[0].focus() : self.clients.openWindow('./'))); });
