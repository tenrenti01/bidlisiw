/* minimal service worker: lets Chrome offer "Install app"; always goes to the network (data stays live) */
self.addEventListener('install', e => self.skipWaiting());
self.addEventListener('activate', e => e.waitUntil(self.clients.claim()));
self.addEventListener('fetch', () => {});
