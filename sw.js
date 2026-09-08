importScripts('data-core.js');
const CACHE_NAME='esta-110sppc-v40';
const ASSETS=['./','./index.html','./styles.css','./i18n.js','./data-core.js','./app.js','./pdf.js','./vendor/jspdf.umd.min.js','./vendor/report-font.ttf','./manifest.json','./icon-shield.png','./icon-192.png','./icon-512.png'];
self.addEventListener('install',event=>{event.waitUntil(caches.open(CACHE_NAME).then(cache=>cache.addAll(ASSETS)).then(()=>self.skipWaiting()))});
self.addEventListener('activate',event=>{event.waitUntil(caches.keys().then(keys=>Promise.all(keys.filter(k=>k.startsWith('esta-110sppc-')&&k!==CACHE_NAME).map(k=>caches.delete(k)))).then(()=>self.clients.claim()))});
self.addEventListener('fetch',event=>{
 const req=event.request,url=new URL(req.url);
 if(req.method!=='GET'||url.origin!==self.location.origin||!url.href.startsWith(self.registration.scope)||url.pathname.includes('/cdn-cgi/'))return;
 const isData=url.pathname.endsWith('/data.json');
 const known=isData||ASSETS.some(a=>new URL(a,self.registration.scope).pathname===url.pathname);
 if(!known)return;
 const work=(async()=>{
   const cache=await caches.open(CACHE_NAME),controller=new AbortController(),timer=setTimeout(()=>controller.abort(),8000);
   try{
     const res=await fetch(req,{cache:'no-store',signal:controller.signal});
     // Never replace an existing report with an Access page or an HTTP error.
     if(res.status===401||res.status===403)return res;
     if(!res.ok)throw Error('HTTP '+res.status);
     if(res.redirected||(isData&&(res.headers.get('Content-Type')||'').includes('text/html')))return res;
     if(isData)Core.validate(await res.clone().json());
     await cache.put(req,res.clone());return res;
   }catch(error){
     const cached=await cache.match(req);
     if(cached){const headers=new Headers(cached.headers);headers.set('X-ESTA-Source','cache');return new Response(await cached.arrayBuffer(),{status:cached.status,statusText:cached.statusText,headers})}
     return new Response(JSON.stringify({error:'unavailable'}),{status:503,headers:{'Content-Type':'application/json'}});
   }finally{clearTimeout(timer)}
 })();
 event.respondWith(work);event.waitUntil(work.then(()=>{}).catch(()=>{}));
});
