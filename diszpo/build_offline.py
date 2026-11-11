# Offline (PWA) változat az 1st AD appból.
# Futtatás (a build_ad_app.py után): python3 diszpo/build_offline.py
#   -> diszpo/offline/   (index.html, sw.js, manifest.json, ikonok, fonts/, storyboard/, vilko/)
#   -> diszpo/Kapaszkodni_AD_offline.zip   (ezt kell feltölteni pl. Netlify Drop-ra)
# A nyilvános tárhely miatt a telefonszámok kimaradnak ebből a változatból.
import hashlib, json, os, re, shutil, zipfile
from PIL import Image, ImageDraw, ImageFilter, ImageFont

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "offline")
os.makedirs(OUT, exist_ok=True)

src = open(os.path.join(HERE, "ad_app.html"), encoding="utf-8").read()

# 1) adatok: telefonszámok ki
m = re.search(r"const D=(\{.*?\});\n", src, flags=re.S)
data = json.loads(m.group(1))
for c in data["contacts"]:
    c["phone"] = ""
src = src[:m.start(1)] + json.dumps(data, ensure_ascii=False) + src[m.end(1):]

# 2) betűk helyben: a sablon Google Fonts hivatkozásából letöltve (latin + latin-ext), a hivatkozások ki
import subprocess
FONTS = os.path.join(OUT, "fonts")
shutil.rmtree(FONTS, ignore_errors=True); os.makedirs(FONTS)
href = re.search(r'<link rel="stylesheet" href="(https://fonts\.googleapis\.com[^"]+)"', src).group(1).replace("&amp;", "&")
UA = "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124 Safari/537.36"
gcss = subprocess.check_output(["curl", "-sS", "-A", UA, href]).decode()
faces = []
for subset, ff in re.findall(r"/\*\s*([\w-]+)\s*\*/\s*(@font-face\s*{[^}]*})", gcss):
    if subset not in ("latin", "latin-ext"):
        continue
    url = re.search(r"url\((https://[^)]+)\)", ff).group(1)
    fn = re.sub(r"[^A-Za-z0-9]+", "-", url.split("/s/")[1].rsplit(".", 1)[0])[-60:] + ".woff2"
    if not os.path.exists(os.path.join(FONTS, fn)):
        subprocess.run(["curl", "-sS", "-o", os.path.join(FONTS, fn), url], check=True)
    faces.append(ff.replace(url, "fonts/" + fn))
fonts_css = "\n".join(faces)
src = re.sub(r'<link rel="preconnect"[^>]*>\n?', "", src)
src = re.sub(r'<link rel="stylesheet" href="https://fonts\.googleapis\.com[^>]*>\n?', "", src)
src = src.replace("<style>", "<style>\n" + fonts_css + "\n", 1)
src = re.sub(r'<meta name="claude:viewport-fit-sides"[^>]*>\n?', "", src)

# 3) teljes HTML-dokumentum PWA-fejléccel
title = re.search(r"<title>(.*?)</title>", src).group(1)
src = src.replace(f"<title>{title}</title>\n", "", 1)
head = f"""<!doctype html>
<html lang="hu"><head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">
<title>{title}</title>
<meta name="theme-color" content="#07090d">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
<meta name="apple-mobile-web-app-title" content="Kapaszkodni AD">
<link rel="manifest" href="manifest.json">
<link rel="apple-touch-icon" href="icon-180.png">
<link rel="icon" href="icon-192.png">
<style>html{{padding-top:env(safe-area-inset-top,0px)}}[hidden]{{display:none!important}}img{{max-width:100%}}</style>
</head><body>
"""
tail = """
<script>
if("serviceWorker" in navigator){
  navigator.serviceWorker.register("sw.js").then(()=>navigator.serviceWorker.ready).then(()=>{
    try{if(!localStorage.getItem("kap-offline-ok")){localStorage.setItem("kap-offline-ok","1");toast("Offline használatra kész","var(--ok)")}}catch(e){}
  }).catch(()=>{});
}
</script>
</body></html>
"""
html = head + src + tail
open(os.path.join(OUT, "index.html"), "w", encoding="utf-8").write(html)

# 4) képek
for d in ("storyboard", "vilko"):
    shutil.copytree(os.path.join(HERE, d), os.path.join(OUT, d), dirs_exist_ok=True)

# 5) ikonok: sötét alap, fénylő gyűrű, K betű
def icon(size):
    im = Image.new("RGB", (size, size), (7, 9, 13))
    glow = Image.new("RGB", (size, size), (7, 9, 13)); g = ImageDraw.Draw(glow)
    pad = size * 0.16; w = max(2, int(size * 0.075))
    g.ellipse([pad, pad, size - pad, size - pad], outline=(53, 216, 232), width=w * 2)
    glow = glow.filter(ImageFilter.GaussianBlur(size * 0.05))
    im = Image.blend(im, glow, 0.9); d = ImageDraw.Draw(im)
    d.arc([pad, pad, size - pad, size - pad], start=-90, end=200, fill=(53, 216, 232), width=w)
    d.arc([pad, pad, size - pad, size - pad], start=200, end=270, fill=(255, 82, 99), width=w)
    f = None
    for fp in (os.path.join(OUT, "fonts", "BarlowCondensed-800-latin.woff2"),
               "/usr/share/fonts/truetype/dejavu/DejaVuSansCondensed-Bold.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"):
        try:
            f = ImageFont.truetype(fp, int(size * 0.42)); break
        except Exception:
            pass
    f = f or ImageFont.load_default()
    d.text((size / 2, size / 2), "K", fill=(234, 240, 245), font=f, anchor="mm")
    return im
for s in (180, 192, 512):
    icon(s).save(os.path.join(OUT, f"icon-{s}.png"))

manifest = {"name": "Kapaszkodni 1st AD", "short_name": "Kapaszkodni AD", "start_url": "./", "scope": "./",
            "display": "standalone", "orientation": "any", "background_color": "#07090d", "theme_color": "#07090d",
            "lang": "hu", "icons": [{"src": "icon-192.png", "sizes": "192x192", "type": "image/png"},
                                     {"src": "icon-512.png", "sizes": "512x512", "type": "image/png", "purpose": "any maskable"}]}
open(os.path.join(OUT, "manifest.json"), "w").write(json.dumps(manifest, ensure_ascii=False, indent=1))

# 6) service worker: minden fájl előre a gyorsítótárba
files = []
for root, _, fs in os.walk(OUT):
    for f in fs:
        rel = os.path.relpath(os.path.join(root, f), OUT).replace(os.sep, "/")
        if rel == "sw.js":
            continue
        files.append(rel)
files = sorted(files)
digest = hashlib.sha1("".join(open(os.path.join(OUT, f), "rb").read().hex()[:2000] + f for f in files).encode()).hexdigest()[:10]
sw = f"""const CACHE="kap-ad-{digest}";
const FILES={json.dumps(["./"] + files)};
self.addEventListener("install",e=>{{e.waitUntil(caches.open(CACHE).then(c=>c.addAll(FILES)).then(()=>self.skipWaiting()))}});
self.addEventListener("activate",e=>{{e.waitUntil(caches.keys().then(ks=>Promise.all(ks.filter(k=>k!==CACHE).map(k=>caches.delete(k)))).then(()=>self.clients.claim()))}});
self.addEventListener("fetch",e=>{{
  if(e.request.method!=="GET")return;
  e.respondWith(caches.match(e.request,{{ignoreSearch:true}}).then(r=>r||fetch(e.request).catch(()=>e.request.mode==="navigate"?caches.match("index.html"):undefined)));
}});
"""
open(os.path.join(OUT, "sw.js"), "w").write(sw)

# 7) zip
zp = os.path.join(HERE, "Kapaszkodni_AD_offline.zip")
with zipfile.ZipFile(zp, "w", zipfile.ZIP_DEFLATED) as z:
    for root, _, fs in os.walk(OUT):
        for f in fs:
            full = os.path.join(root, f)
            z.write(full, os.path.join("kapaszkodni-ad", os.path.relpath(full, OUT)))
print("kész:", OUT, "·", len(files), "fájl ·", round(os.path.getsize(zp) / 1e6, 1), "MB zip")
