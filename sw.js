// Odmit service worker: offline app shell + push reminders (Firebase Cloud Messaging) + notification clicks.
// One service worker for everything: the app passes this registration to Firebase's getToken(), so no
// firebase-messaging-sw.js is needed and no Firebase code runs here (works offline, nothing to download).
// API calls (to script.google.com) are never cached here; the app keeps its own offline copy of your data.
const CACHE = 'odmit-v8'; // v6: push notifications; v7: shared-completion alerts; v8: Microsoft To Do import
// Big picker data (full Myna UI icon set, emoji list): fetched lazily by the app, pre-cached here in the background for offline use.
const LAZY = ['./myna-icons-outline.json', './myna-icons-solid.json', './emoji.json'];
const SHELL = ['./', './index.html', './manifest.json', './icons/icon-192.png', './icons/icon-512.png',
  './icons/icon-maskable-512.png', './icons/apple-touch-icon.png', './icons/favicon-32.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL).then(() => { c.addAll(LAZY).catch(() => {}); })).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const req = e.request;
  const url = new URL(req.url);
  if (req.method !== 'GET' || url.origin !== self.location.origin) return; // API etc: straight to network
  // Microsoft sign-in lands on msal-redirect.html with the sign-in response in its URL: never cache or answer it here.
  if (/\/msal-redirect(\.html|-bridge\.min\.js)$/.test(url.pathname)) return;
  if (req.mode === 'navigate') { // network first, so updates arrive; cached page when offline
    const scope = new URL(self.registration.scope).pathname;
    if (url.pathname !== scope && url.pathname !== scope + 'index.html') return; // some other page: not the app shell
    e.respondWith(fetch(req, { cache: 'no-cache' }).then(res => { const copy = res.clone(); caches.open(CACHE).then(c => c.put('./index.html', copy)); return res; })
      .catch(() => caches.match('./index.html').then(r => r || caches.match('./'))));
    return;
  }
  e.respondWith(caches.match(req, { ignoreSearch: true }).then(hit => {
    const net = fetch(req).then(res => { if (res.ok) { const copy = res.clone(); caches.open(CACHE).then(c => c.put(req, copy)); } return res; }).catch(() => hit);
    return hit || net;
  }));
});
self.addEventListener('notificationclick', e => {
  e.notification.close();
  const nd = e.notification.data || {}, taskId = nd.taskId || '', listId = nd.listId || '';
  const q = [taskId && 'task=' + encodeURIComponent(taskId), listId && 'list=' + encodeURIComponent(listId)].filter(Boolean).join('&');
  const url = new URL('./', self.registration.scope).href + (q ? '?' + q : '');
  e.waitUntil(self.clients.matchAll({ type: 'window', includeUncontrolled: true }).then(list => {
    for (const c of list) { if ('focus' in c) { c.postMessage({ type: 'open-task', taskId, listId }); return c.focus(); } }
    return self.clients.openWindow(url);
  }));
});

// Push from FCM. The server sends data-only messages: {data: {type, taskId, title, body, tag}}
// type 'reminder' (sticky) | 'completed' ("<Name> completed a task", also carries listId) | 'test'.
// (A "notification" payload is handled too, just in case.) Chrome requires a visible notification for every push.
self.addEventListener('push', e => {
  let p = {};
  try { p = e.data ? e.data.json() : {}; } catch (err) { p = { data: { body: e.data ? e.data.text() : '' } }; }
  const d = p.data || {}, n = p.notification || {};
  const title = d.title || n.title || 'Odmit reminder';
  const taskId = d.taskId || '', listId = d.listId || '';
  // same tag as the in-app reminder (task id), so a push and the open app's own notification merge into one
  const opts = { body: d.body || n.body || '', icon: './icons/icon-192.png', badge: './icons/icon-192.png',
    tag: d.tag || taskId || 'odmit', data: { taskId, listId }, requireInteraction: d.type === 'reminder', renotify: false };
  e.waitUntil(self.registration.getNotifications({ tag: opts.tag }).then(open => {
    if (open.length && d.type === 'reminder') return; // already showing (the open app notified first)
    return self.registration.showNotification(title, opts);
  }).then(() => self.clients.matchAll({ type: 'window', includeUncontrolled: true }))
    .then(list => list.forEach(c => c.postMessage({ type: 'push', kind: d.type || '', taskId }))));
});
