// Nyomtatható diszpó PDF készítése: node diszpo/build_pdf.js
// Az adatokat a nyomtathato.html elején lévő INFO objektumban lehet kitölteni.
const {chromium}=require('playwright');
const path=require('path');
(async()=>{
  const proxy=process.env.HTTPS_PROXY?{server:process.env.HTTPS_PROXY}:undefined;
  const b=await chromium.launch({proxy});const p=await b.newPage();
  await p.goto('file://'+path.join(__dirname,'nyomtathato.html'),{waitUntil:'networkidle'});
  await p.evaluate(()=>document.fonts.ready);
  await p.pdf({path:path.join(__dirname,'Kapaszkodni_diszpo_2026-11-12.pdf'),format:'A4',printBackground:true,preferCSSPageSize:true,
    displayHeaderFooter:true,headerTemplate:'<span></span>',
    footerTemplate:'<div style="font-family:Courier New,monospace;font-size:7px;width:100%;padding:0 9mm;display:flex;justify-content:space-between"><span>KAPASZKODNI · napi diszpó · 2026.11.12. · kérdés esetén: 2. asszisztens (Balázs)</span><span><span class="pageNumber"></span>/<span class="totalPages"></span></span></div>'});
  await b.close();
})();
