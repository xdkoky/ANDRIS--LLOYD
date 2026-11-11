// A három nyomtatandó PDF elkészítése: python3 diszpo/build_print.py && node diszpo/build_pdf.js
// - Kapaszkodni_diszpo_2026-11-12.pdf   (print_diszpo.html, A4 álló)
// - Kapaszkodni_shotlist.pdf            (print_shotlist.html, A4 fekvő)
// - Kapaszkodni_forgatokonyv.pdf        (print_forgatokonyv.html, A4 álló, számozatlan címlappal)
// A print_*.html fájlokat a build_print.py állítja elő a legfrissebb táblázatból.
const {chromium}=require('playwright');
const fs=require('fs'),path=require('path');

const foot=left=>`<div style="font-family:Courier New,monospace;font-size:7px;width:100%;padding:0 9mm;display:flex;justify-content:space-between"><span>${left}</span><span><span class="pageNumber"></span>/<span class="totalPages"></span></span></div>`;
const JOBS=[
  {html:'print_diszpo.html',out:'Kapaszkodni_diszpo_2026-11-12.pdf',footer:foot('KAPASZKODNI · napi diszpó · 2026.11.12.')},
  {html:'print_shotlist.html',out:'Kapaszkodni_shotlist.pdf',footer:foot('KAPASZKODNI · shotlist · 2026.11.12.')},
  // A forgatókönyv címlapja külön készül, hogy ne kapjon oldalszámot; utána pdfunite fűzi össze.
  {html:'print_forgatokonyv.html',out:'_cim.pdf',mode:'only-title',header:'<span></span>',footer:'<span></span>'},
  {html:'print_forgatokonyv.html',out:'_torzs.pdf',mode:'only-body',
   header:'<div style="font-family:Courier New,monospace;font-size:9px;width:100%;padding:0 22mm;text-align:right"><span class="pageNumber"></span>.</div>',footer:'<span></span>'}
];

(async()=>{
  const proxy=process.env.HTTPS_PROXY?{server:process.env.HTTPS_PROXY}:undefined;
  const b=await chromium.launch({proxy});
  for(const j of JOBS){
    const p=await b.newPage();
    await p.goto('file://'+path.join(__dirname,j.html),{waitUntil:'networkidle'});
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
