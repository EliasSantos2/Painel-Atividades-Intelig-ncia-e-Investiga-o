const CACHE = 'rumo-intel-v1';
const ASSETS = [
  './',
  './index.html',
  './Painel_Investigacoes.html',
  './Painel_Produtividade.html',
  './Indicadores_Inteligencia.html',
  './Biblioteca_Inteligencia.html',
  './Estrutura_Organizacional.html',
  './Mapa_Ocorrencias.html',
  './Consulta_Oculta.html',
  './dados_apreensoes.js',
];

self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE).then(c => c.addAll(ASSETS)).then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))
    ).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  // Para requisições ao Supabase, sempre vai para a rede (dados em tempo real)
  if (e.request.url.includes('supabase.co')) {
    e.respondWith(fetch(e.request).catch(() => new Response('{}', { headers: { 'Content-Type': 'application/json' } })));
    return;
  }

  // Cache-first para assets locais; fallback para rede
  e.respondWith(
    caches.match(e.request).then(cached => cached || fetch(e.request).then(res => {
      if (res && res.status === 200 && e.request.method === 'GET') {
        const clone = res.clone();
        caches.open(CACHE).then(c => c.put(e.request, clone));
      }
      return res;
    }))
  );
});
