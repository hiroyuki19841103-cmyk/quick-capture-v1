import {defineConfig} from "vite";
import {readdirSync,readFileSync,writeFileSync} from "node:fs";
import {createHash} from "node:crypto";
import {join} from "node:path";
export default defineConfig({
  base:"/",build:{target:"es2022"},
  plugins:[{name:"qc-offline-build",apply:"build",closeBundle(){
    const root="dist";
    const walk=(dir)=>readdirSync(join(root,dir),{withFileTypes:true}).flatMap(e=>e.isDirectory()?walk(join(dir,e.name)):[join(dir,e.name)]);
    const files=walk("").filter(f=>/\.(html|js|css|svg|png|ico|woff2?)$/.test(f)&&f!=="sw.js").sort();
    const hash=createHash("sha256").update(readFileSync("vite.config.js"));
    for(const f of files)hash.update(f).update(readFileSync(join(root,f)));
    const version=hash.digest("hex").slice(0,16);
    writeFileSync(join(root,"sw.js"),`const CACHE="qc-shell-${version}";
const ASSETS=${JSON.stringify(files.map(f=>"/"+f))};
self.addEventListener("install",event=>event.waitUntil((async()=>{
  const cache=await caches.open(CACHE);
  await cache.addAll(ASSETS.map(url=>new Request(url,{cache:"reload"})));
})()));
self.addEventListener("activate",event=>event.waitUntil((async()=>{
  for(const key of await caches.keys())if(key.startsWith("qc-shell-")&&key!==CACHE)await caches.delete(key);
  await self.clients.claim();
})()));
self.addEventListener("fetch",event=>{
  const url=new URL(event.request.url);
  if(event.request.method!=="GET"||url.origin!==self.location.origin)return;
  if(event.request.mode==="navigate"&&(url.pathname==="/"||url.pathname==="/index.html")){
    event.respondWith(caches.open(CACHE).then(cache=>cache.match("/index.html")).then(hit=>hit||fetch(event.request)));return;
  }
  if(ASSETS.includes(url.pathname))event.respondWith(caches.open(CACHE).then(cache=>cache.match(url.pathname)).then(hit=>hit||fetch(event.request)));
});
`);
  }}]
});
