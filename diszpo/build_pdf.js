// Nyomtatható PDF-ek készítése: node diszpo/build_pdf.js
// - Kapaszkodni_diszpo_2026-11-12.pdf   (nyomtathato.html; az INFO objektumban tölthető ki)
// - Kapaszkodni_shotlist.pdf            (shotlist_nyomtathato.html)
// - Kapaszkodni_utemterv_2026-11-12.pdf (utemterv.html)
// - Kapaszkodni_forgatokonyv.pdf        (forgatokonyv_nyomtathato.html)
// A shotlist és a szkript adatai az index.html-ből jönnek, így egy helyen kell javítani őket.
const {chromium}=require('playwright');
const fs=require('fs'),path=require('path');

const src=fs.readFileSync(path.join(__dirname,'index.html'),'utf8');
const grab=name=>{const m=src.match(new RegExp(`const ${name}=(\\[[\\s\\S]*?\\n\\]);`));if(!m)throw new Error(name+' nem található');return eval(m[1])};
const DATA={SHOTS:grab('SHOTS'),SCENES:grab('SCENES'),SCRIPT:grab('SCRIPT')};

const foot=(left,size=7)=>`<div style="font-family:Courier New,monospace;font-size:${size}px;width:100%;padding:0 9mm;display:flex;justify-content:space-between"><span>${left}</span><span><span class="pageNumber"></span>/<span class="totalPages"></span></span></div>`;
const JOBS=[
  {html:'nyomtathato.html',out:'Kapaszkodni_diszpo_2026-11-12.pdf',footer:foot('KAPASZKODNI · napi diszpó · 2026.11.12. · kérdés esetén: 2. asszisztens (Balázs)')},
  {html:'shotlist_nyomtathato.html',out:'Kapaszkodni_shotlist.pdf',render:true,footer:foot('KAPASZKODNI · shotlist · 2026.11.12.')},
  {html:'utemterv.html',out:'Kapaszkodni_utemterv_2026-11-12.pdf',footer:foot('KAPASZKODNI · napi ütemterv · TERVEZET · 2026.11.12.')},
  // A forgatókönyv címlapja külön készül, hogy ne kapjon oldalszámot; utána pdfunite fűzi össze.
  {html:'forgatokonyv_nyomtathato.html',out:'_cim.pdf',render:true,mode:'only-title',header:'<span></span>',footer:'<span></span>'},
  {html:'forgatokonyv_nyomtathato.html',out:'_torzs.pdf',render:true,mode:'only-body',
   header:'<div style="font-family:Courier New,monospace;font-size:9px;width:100%;padding:0 22mm;text-align:right"><span class="pageNumber"></span>.</div>',footer:'<span></span>'}
];

(async()=>{
  const proxy=process.env.HTTPS_PROXY?{server:process.env.HTTPS_PROXY}:undefined;
  const b=await chromium.launch({proxy});
  for(const j of JOBS){
    const p=await b.newPage();
    await p.goto('file://'+path.join(__dirname,j.html),{waitUntil:'networkidle'});
    if(j.render)await p.evaluate(d=>window.render(d),DATA);
    if(j.mode)await p.evaluate(m=>document.body.classList.add(m),j.mode);
    await p.evaluate(async()=>{await document.fonts.ready;await Promise.all([...document.images].map(i=>i.complete?0:new Promise(r=>{i.onload=i.onerror=r})))});
    await p.pdf({path:path.join(__dirname,j.out),printBackground:true,preferCSSPageSize:true,displayHeaderFooter:true,
      headerTemplate:j.header||'<span></span>',footerTemplate:j.footer});
    if(!j.out.startsWith('_'))console.log('kész:',j.out);
    await p.close();
  }
  await b.close();
  const d=f=>path.join(__dirname,f);
  require('child_process').execFileSync('pdfunite',[d('_cim.pdf'),d('_torzs.pdf'),d('Kapaszkodni_forgatokonyv.pdf')]);
  fs.unlinkSync(d('_cim.pdf'));fs.unlinkSync(d('_torzs.pdf'));
  console.log('kész: Kapaszkodni_forgatokonyv.pdf');
})();
