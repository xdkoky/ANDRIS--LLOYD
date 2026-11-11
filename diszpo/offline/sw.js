const CACHE="kap-ad-69bb2e072c";
const CORE=["./", "fonts/courierprime-v11-u-450q2lgwslOqpF-6gQ8kELaw9pWt-.woff2", "fonts/courierprime-v11-u-450q2lgwslOqpF-6gQ8kELawFpWg.woff2", "fonts/courierprime-v11-u-4k0q2lgwslOqpF-6gQ8kELY7pMT-7fq8Ho.woff2", "fonts/courierprime-v11-u-4k0q2lgwslOqpF-6gQ8kELY7pMT-Dfqw.woff2", "fonts/courierprime-v11-u-4n0q2lgwslOqpF-6gQ8kELawRZVsf6lvg.woff2", "fonts/courierprime-v11-u-4n0q2lgwslOqpF-6gQ8kELawRZWMf6.woff2", "fonts/inter-v20-UcC73FwrK3iLTeHuS-nVMrMxCp50SjIa1ZL7.woff2", "fonts/inter-v20-UcC73FwrK3iLTeHuS-nVMrMxCp50SjIa25L7SUc.woff2", "icon-180.png", "icon-192.png", "icon-512.png", "manifest.json"];
const MEDIA=["storyboard/1-1.jpg", "storyboard/1-2.jpg", "storyboard/1-3A.jpg", "storyboard/1-3B.jpg", "storyboard/2-1A.jpg", "storyboard/2-1B.jpg", "storyboard/2-1C.jpg", "storyboard/2-1E.jpg", "storyboard/2-2.jpg", "storyboard/2-3A.jpg", "storyboard/2-3B.jpg", "storyboard/2-4.jpg", "storyboard/2-5.jpg", "storyboard/2-6.jpg", "storyboard/full/1-1.jpg", "storyboard/full/1-2.jpg", "storyboard/full/1-3A.jpg", "storyboard/full/1-3B.jpg", "storyboard/full/2-1A.jpg", "storyboard/full/2-1B.jpg", "storyboard/full/2-1C.jpg", "storyboard/full/2-1E.jpg", "storyboard/full/2-2.jpg", "storyboard/full/2-3A.jpg", "storyboard/full/2-3B.jpg", "storyboard/full/2-4.jpg", "storyboard/full/2-5.jpg", "storyboard/full/2-6.jpg", "vilko/full/1-1.jpg", "vilko/full/1-2.jpg", "vilko/full/1-3A.jpg", "vilko/full/1-3B.jpg", "vilko/full/2-1A.jpg", "vilko/full/2-1B.jpg", "vilko/full/2-1C.jpg", "vilko/full/2-1E.jpg", "vilko/full/2-2.jpg", "vilko/full/2-3A.jpg", "vilko/full/2-3B.jpg", "vilko/full/2-4.jpg", "vilko/full/2-5.jpg", "vilko/full/2-6.jpg", "vilko/full/3-1.jpg", "vilko/full/3-2.jpg", "vilko/t/1-1.jpg", "vilko/t/1-2.jpg", "vilko/t/1-3A.jpg", "vilko/t/1-3B.jpg", "vilko/t/2-1A.jpg", "vilko/t/2-1B.jpg", "vilko/t/2-1C.jpg", "vilko/t/2-1E.jpg", "vilko/t/2-2.jpg", "vilko/t/2-3A.jpg", "vilko/t/2-3B.jpg", "vilko/t/2-4.jpg", "vilko/t/2-5.jpg", "vilko/t/2-6.jpg", "vilko/t/3-1.jpg", "vilko/t/3-2.jpg"];
self.addEventListener("install",e=>{e.waitUntil((async()=>{
  const c=await caches.open(CACHE);await c.addAll(CORE);
  const small=MEDIA.filter(u=>!u.includes("/full/")),big=MEDIA.filter(u=>u.includes("/full/"));
  await Promise.allSettled(small.map(u=>c.add(u)));
  for(let i=0;i<big.length;i+=3)await Promise.allSettled(big.slice(i,i+3).map(u=>c.add(u)));
  await self.skipWaiting();
})())});
self.addEventListener("activate",e=>{e.waitUntil(caches.keys().then(ks=>Promise.all(ks.filter(k=>k!==CACHE).map(k=>caches.delete(k)))).then(()=>self.clients.claim()))});
self.addEventListener("fetch",e=>{
  if(e.request.method!=="GET"||new URL(e.request.url).origin!==location.origin)return;
  e.respondWith((async()=>{
    const hit=await caches.match(e.request,{ignoreSearch:true});if(hit)return hit;
    try{const r=await fetch(e.request);if(r.ok){const c=await caches.open(CACHE);c.put(e.request,r.clone())}return r}
    catch(err){if(e.request.mode==="navigate")return caches.match("./");throw err}
  })());
});
