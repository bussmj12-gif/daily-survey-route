// Minimal service worker whose only job is to receive files shared to this
// installed app via the OS "공유하기" sheet (Web Share Target API), stash the
// file in Cache Storage, then hand off to index.html to actually read it.
// It intentionally does not cache app files or work offline -- that's not
// the goal here, just removing the manual "download then upload" steps.

const SHARE_CACHE_NAME = 'shared-workbook-cache';

self.addEventListener('install', () => {
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(self.clients.claim());
});

self.addEventListener('fetch', (event) => {
  const url = new URL(event.request.url);
  const isShareTarget = event.request.method === 'POST' && url.pathname.endsWith('/share-target.html');
  if(!isShareTarget) return; // let every other request pass through untouched

  event.respondWith((async () => {
    try{
      const formData = await event.request.formData();
      const file = formData.get('workbook');
      if(file){
        const cache = await caches.open(SHARE_CACHE_NAME);
        await cache.put(url.origin + '/shared-workbook-data', new Response(file));
      }
    }catch(e){
      // If parsing fails for any reason, we still redirect -- the app will
      // simply find nothing shared and behave like a normal cold start.
    }
    return Response.redirect('./index.html?shared=1', 303);
  })());
});
