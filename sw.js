const V = 'brollerpro-v12';
const SHELL = ['./', 'xlsx.js', 'manifest.webmanifest', 'icon-192.png', 'icon-512.png', 'icon-maskable-512.png'];

// Yo'naltirilgan (redirected) javobni tozalaydi: aks holda navigatsiyada ERR_FAILED beradi
const clean = async r => {
  if (!r.redirected) return r;
  return new Response(await r.blob(), { status: r.status, statusText: r.statusText, headers: r.headers });
};

self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(V)
      .then(c => Promise.all(SHELL.map(u => fetch(u).then(clean).then(r => { if (!r.ok) throw new Error(u); return c.put(u, r); }))))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(k => k !== V).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);

  // Shrift fayllari: keshdan beradi, fonda yangilaydi
  if (url.hostname === 'fonts.googleapis.com' || url.hostname === 'fonts.gstatic.com') {
    e.respondWith(
      caches.open(V).then(async c => {
        const hit = await c.match(req);
        const net = fetch(req).then(r => { c.put(req, r.clone()); return r; }).catch(() => hit);
        return hit || net;
      })
    );
    return;
  }

  if (url.origin !== location.origin) return;

  // /index.html ni aniq yo'naltiramiz (keshga yo'naltirilgan javob tushmasligi uchun)
  if (req.mode === 'navigate' && /\/index\.html$/.test(url.pathname)) {
    e.respondWith(Response.redirect(new URL('./' + url.search, url).href, 302));
    return;
  }

  // Ilova fayllari: avval keshdan, internet bo'lsa yangilab qo'yadi
  e.respondWith(
    caches.match(req, { ignoreSearch: true }).then(hit => {
      const net = fetch(req).then(r => {
        if (r.ok) clean(r.clone()).then(x => caches.open(V).then(c => c.put(req, x)));
        return r;
      }).catch(() => hit || (req.mode === 'navigate' ? caches.match('./') : undefined));
      return hit || net;
    })
  );
});