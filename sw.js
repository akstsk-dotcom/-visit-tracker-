/* SSPL PR Desk service worker - app shell only. NEVER touches Google Apps Script / data calls. */
const V='prdesk-v1';
const SHELL=['SSPL_PR_Desk_Mobile.html','SSPL_PR_Desk_Desktop.html','icon-192.png','icon-512.png'];
self.addEventListener('install',e=>{e.waitUntil(caches.open(V).then(c=>Promise.all(SHELL.map(u=>c.add(u).catch(()=>{})))).then(()=>self.skipWaiting()))});
self.addEventListener('activate',e=>{e.waitUntil(caches.keys().then(k=>Promise.all(k.filter(x=>x!==V).map(x=>caches.delete(x)))).then(()=>self.clients.claim()))});
self.addEventListener('fetch',e=>{
  const r=e.request;if(r.method!=='GET')return;
  const u=new URL(r.url);
  const same=u.origin===self.location.origin;
  const stat=/(fonts\.googleapis\.com|fonts\.gstatic\.com|cdnjs\.cloudflare\.com)$/.test(u.hostname);
  if(!same&&!stat)return;                       // data/API/maps/geocoding: straight to network, never cached
  if(same){                                      // network-first: users always get the newest app, cache only if offline
    e.respondWith(fetch(r).then(res=>{if(res.ok){const c=res.clone();caches.open(V).then(x=>x.put(r,c))}return res}).catch(()=>caches.match(r).then(m=>m||caches.match('SSPL_PR_Desk_Mobile.html'))));
  }else{                                         // fonts / pdf libs: cache-first
    e.respondWith(caches.match(r).then(m=>m||fetch(r).then(res=>{const c=res.clone();caches.open(V).then(x=>x.put(r,c));return res})));
  }
});
