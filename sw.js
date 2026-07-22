// Service worker mínimo para que la Plataforma LOGÍSTICA sea instalable como app (PWA).
// IMPORTANTE: NO cachea nada a propósito — todas las peticiones van siempre a la red,
// para que nunca se sirvan páginas ni data.json desatualizados (ver incidentes de
// "pestaña con versión antigua" que pisaban datos).
self.addEventListener('install', function(e){ self.skipWaiting(); });
self.addEventListener('activate', function(e){ e.waitUntil(self.clients.claim()); });
self.addEventListener('fetch', function(e){ /* passthrough: red directa, sin caché */ });
