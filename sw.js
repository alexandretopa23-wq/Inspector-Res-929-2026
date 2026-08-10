/* Service worker del Inspector de Piscinas.
   Estrategia: network-first para lo que cambia con cada despliegue (el
   propio app shell y config.json, que trae la URL del webhook), cache-first
   para assets estáticos (íconos, manifest). Las peticiones POST al webhook
   de Sheets NO se interceptan aquí — esas ya las reintenta la cola offline
   de index.html (IndexedDB); el service worker solo se encarga de que la
   app cargue sin señal.

   Sube CACHE_NAME en cada despliegue con cambios relevantes de assets, para
   que las pestañas viejas no se queden sirviendo un shell desactualizado
   indefinidamente — junto con el `updateViaCache:'none'` del registro en
   index.html, esto hace que el navegador revise el sw.js nuevo pronto. */
var CACHE_NAME = 'inspector-piscinas-v2';
var APP_SHELL = [
  './',
  './index.html',
  './dashboard.html',
  './config.json',
  './manifest.json',
  './icon-192.png',
  './icon-512.png'
];

self.addEventListener('install', function(event){
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(function(cache){ return cache.addAll(APP_SHELL); })
      .then(function(){ return self.skipWaiting(); })
  );
});

self.addEventListener('activate', function(event){
  event.waitUntil(
    caches.keys().then(function(keys){
      return Promise.all(keys.filter(function(k){ return k!==CACHE_NAME; }).map(function(k){ return caches.delete(k); }));
    }).then(function(){ return self.clients.claim(); })
  );
});

function networkFirst(req){
  return fetch(req).then(function(res){
    var copia = res.clone();
    caches.open(CACHE_NAME).then(function(cache){ cache.put(req, copia); });
    return res;
  }).catch(function(){ return caches.match(req); });
}

function cacheFirst(req){
  return caches.match(req).then(function(cached){
    if(cached) return cached;
    return fetch(req).then(function(res){
      var copia = res.clone();
      caches.open(CACHE_NAME).then(function(cache){ cache.put(req, copia); });
      return res;
    });
  });
}

self.addEventListener('fetch', function(event){
  var req = event.request;
  if(req.method !== 'GET') return; // POST al webhook de Sheets: no se intercepta
  var url = new URL(req.url);
  if(url.origin !== self.location.origin) return; // dominios externos (script.google.com, etc.): tampoco

  var esShellDinamico = req.mode==='navigate' || url.pathname.endsWith('/index.html') || url.pathname.endsWith('/config.json') || url.pathname.endsWith('/');
  event.respondWith(esShellDinamico ? networkFirst(req) : cacheFirst(req));
});
