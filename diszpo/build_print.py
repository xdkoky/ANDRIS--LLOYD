# A három nyomtatandó dokumentum a legfrissebb táblázatból (Kapaszkodni_utemterv_V*.xlsx):
#   print_diszpo.html        -> Kapaszkodni_diszpo_2026-11-12.pdf   (A4 álló)
#   print_shotlist.html      -> Kapaszkodni_shotlist.pdf            (A4 fekvő, forgatási sorrend, storyboarddal)
#   print_forgatokonyv.html  -> Kapaszkodni_forgatokonyv.pdf        (A4 álló, történet szerinti sorrend, a margón a felsnitt shotjai)
# Futtatás: python3 diszpo/build_print.py && node diszpo/build_pdf.js
# Az adatok ugyanonnan jönnek, mint az appé (build_ad_app.py), így a nyomtatott és az appbeli terv egyezik.
import html, json, os, re, subprocess, sys

HERE = os.path.dirname(os.path.abspath(__file__))
subprocess.run([sys.executable, os.path.join(HERE, "build_ad_app.py")], check=True, stdout=subprocess.DEVNULL)
D = json.loads(re.search(r"const D=(\{.*?\});\n", open(os.path.join(HERE, "ad_app.html"), encoding="utf-8").read(), re.S).group(1))
js = r"""const s=require('fs').readFileSync(process.argv[1],'utf8');
console.log(JSON.stringify(eval(s.match(/const SHOTS=(\[[\s\S]*?\n\]);/)[1])));"""
SL = {x["id"].replace("-", "/", 1): x for x in json.loads(subprocess.check_output(["node", "-e", js, os.path.join(HERE, "index.html")]))}

# az író elérhetősége a forgatókönyv címlapjára: a meglévő diszpó-oldal adataiból, nem ide beírva
_m = re.search(r'tel:"([^"]+)",mail:"([^"]+)"', open(os.path.join(HERE, "index.html"), encoding="utf-8").read())
AUTHOR = (_m.group(1), _m.group(2)) if _m else ("", "")
e = lambda v: html.escape(str(v if v is not None else ""))
hm = lambda m: f"{int(m) // 60}:{int(m) % 60:02d}"
B = D["blocks"]
for b in B:
    b["e"] = b["s"] + b["d"]
SHOOT = [b for b in B if b["type"] == "shoot"]
WRAP = B[-1]["e"]
DATE = "2026.11.12."
FONT = '<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Courier+Prime:ital,wght@0,400;0,700;1,400&family=IBM+Plex+Sans+Condensed:wght@400;500;600;700&display=swap">'
STAMP = f"{D['src'].replace('.xlsx', '').replace('Kapaszkodni_utemterv_', '')} időrend"

def ids_of(b):
    return D["shots"].get(b["label"], {}).get("ids", [])
def next_shoot(i):
    return next((x for x in B[i + 1:] if x["type"] == "shoot"), None)
MONT = {"loc": "Nappali", "dn": "NAPPAL", "type": "Montázs", "desc": "Bicepsz a kézisúllyal, nyújtás a matracon, Nagyi segít (kiegészítő jelenet).",
        "lens": "", "angle": "", "move": "Svenk", "sup": "Sticks", "talent": "Soma, Nagyi", "eff": "", "com": ""}
info = lambda i: SL.get(i, MONT)
CAST_ID = {c["role"]: n + 1 for n, c in enumerate(D["cast"])}
lunch = next((b for b in B if b["type"] == "lunch"), None)
FIX_POS = {"1ST AD": "1st AD", "2ND AD": "2nd AD", "Világoító": "Világosító"}
FIX_NAME = {"Volninasky Kata": "Volniansky Kata"}
phone = lambda role: next((c for c in D["contacts"] if c["role"].lower().startswith(role)), {"name": "", "phone": ""})

# ====================== 1) DISZPÓ ======================
def diszpo():
    rows = []
    for i, b in enumerate(B):
        t = f'{hm(b["s"])}–<br>{hm(b["e"])}'
        if b["type"] in ("prep", "pres"):
            txt = f'<b>{e(b["label"])}</b>' + (f' — {e(" · ".join(x for x in (b["setup"], b["note"]) if x))}' if b["setup"] or b["note"] else "")
            rows.append(f'<tr class="pre"><td class="time">{t}</td><td colspan="7">{txt}</td></tr>')
        elif b["type"] == "lunch":
            rows.append(f'<tr class="sep"><td class="time">{t}</td><td colspan="7">Ebédszünet: {hm(b["s"])}–{hm(b["e"])}</td></tr>')
        else:
            tgt = b if b["type"] == "shoot" else next_shoot(i)
            ids = ids_of(tgt) if tgt else []
            sc = "<br>".join("MONT." if x == "MONTAGE" else e(x) for x in ids)
            dn = "E" if ids and ids[0].startswith("3/") else "N"
            loc = e(info(ids[0])["loc"]) if ids else ""
            if b["type"] == "shoot":
                syn = e(" / ".join(dict.fromkeys(info(x)["desc"] for x in ids)))
                kar = ",".join(str(CAST_ID[r]) for r in sorted(D["shots"][b["label"]]["talent"], key=lambda r: CAST_ID.get(r, 9)) if r in CAST_ID)
                note = " — ".join(x for x in (b["setup"], b["note"]) if x)
                rows.append(f'<tr class="rec"><td class="time">{t}</td><td class="no">{sc}</td><td>B<br>{dn}</td><td><b>Felvétel</b></td><td>{syn}</td><td>{kar}</td><td>{loc}</td><td>{e(note)}</td></tr>')
            else:
                generic = re.fullmatch(r"(?i)átállás", b["label"].strip())
                note = " — ".join(x for x in (b["setup"], b["note"]) if x)
                rows.append(f'<tr class="mv"><td class="time">{t}</td><td class="no">{sc}</td><td>B<br>{dn}</td><td>Átállás</td><td>{"" if generic else e(b["label"])}</td><td></td><td>{loc}</td><td>{e(note)}</td></tr>')
    rows.append(f'<tr class="sep"><td class="time">{hm(WRAP)}</td><td colspan="7">WRAP — bontás, pakolás, napi riport</td></tr>')
    cast = "".join(f'<tr><td><b>{n + 1}</b></td><td>{e(c["name"])}</td><td>{e(c["role"])}</td><td>{e(c["arr"])}</td><td>{e(c["ruha"])}</td><td>{e(c["smink"])}</td><td>{e(c["haj"])}</td><td>{e(c["onset"])}</td></tr>'
                   for n, c in enumerate(D["cast"]))
    crew, dept = "", None
    for c in D["crew"]:
        if c["dept"] != dept:
            dept = c["dept"]; crew += f'<tr><td colspan="3" class="bar">{e(dept.upper())}</td></tr>'
        crew += f'<tr><td>{e(FIX_POS.get(c["pos"], c["pos"]))}</td><td>{e(FIX_NAME.get(c["name"], c["name"]))}</td><td>{e(c["arr"])}</td></tr>'
    PROPS = [("1", "Teli tálak étellel (több take-re) + sok üres tányér"), ("1", "Villa, kanál · mosogatószivacs"), ("1", "Cetli és ceruza (Nagyi zsebében)"),
             ("1", "Hűtő (mágnesek, rajzok), bicikli, polc"), ("1", "Tányérszárító, konyhai eszközök, fém eszköztartó"),
             ("2", "Fekvenyomó pad, rúd, könnyített súlyok + pótsúly a földön"), ("2", "Kézisúlyok · kulacs · sporttáskák · matrac a nyújtáshoz"),
             ("2", "Anyag a plafonra (2/3A, 2/3B)"), ("2", "Nagy szőnyeg, bútorok, székek, fogasok"), ("2", "Érmek, trófeák, poszterek, kép, régi naptár"),
             ("2", "Praktikal lámpák, Andris lámpája, NAGY ventilátor"), ("2", "Húzódzkodó rúd, darts tábla, kosárpalánk az ajtón"),
             ("2", "Rádió, magnó, VHS, kazetták, bakelit"), ("2", "Könyvek, magazinok, magyar zászló, hamutartó, cigi"),
             ("2", "Előtér: ventilátor + növény (tágban és közeliben is)"), ("3", "Működő TV + régi súlyemelő tudósítás"),
             ("J/S", "Fiú: fehér trikó, fekete rövidnadrág · hajzselé / sea salt"), ("J/S", "Nagyi: otthonka/kardigán zsebbel, papucs, láncos szemüveg")]
    props = "".join(f'<tr><td class="c">{j}</td><td>{e(p)}</td><td></td></tr>' for j, p in PROPS)
    r, o, a = phone("rendező"), phone("operatőr"), phone("első")
    first_shoot = SHOOT[0]["s"]
    return f"""<!doctype html><html lang="hu"><head><meta charset="utf-8"><title>Kapaszkodni – napi diszpó</title>{FONT}
<style>
@page{{size:A4;margin:9mm 9mm 12mm}}
*{{box-sizing:border-box}}html,body{{margin:0;background:#fff;color:#000}}
body{{font-family:"Courier Prime","Courier New",monospace;font-size:8.4pt;line-height:1.22}}
table{{border-collapse:collapse;width:100%}}th,td{{border:1px solid #000;padding:2.5px 4px;vertical-align:top}}
th{{background:#c4c4c4;font-weight:700;text-align:center;font-size:7.4pt}}
tr{{break-inside:avoid}}thead{{display:table-header-group}}
.bar{{background:#c4c4c4;border:1.5px solid #000;text-align:center;font-weight:700;padding:3px;font-size:9pt}}
.titlebox{{border:1.5px solid #000;width:62%;margin-bottom:4mm}}
.titlebox .t{{font-size:24pt;font-weight:700;text-align:center;text-decoration:underline;padding:3mm 0 2mm}}
.titlebox .d{{background:#e2e2e2;border-top:1.5px solid #000;text-align:center;font-weight:700;font-size:13pt;padding:2mm 0}}
.head{{border:1.5px solid #000}}.head td{{width:33.3%;padding:5px 6px}}.head .c{{text-align:center}}
.head .big{{font-size:12.5pt;font-weight:700;text-decoration:underline;display:block}}.head .huge{{font-size:13pt;font-weight:700;display:block;margin:2px 0 6px}}
.head .row{{margin:2px 0}}.line{{display:inline-block;min-width:44mm;border-bottom:1px solid #000;height:1.05em}}
.red{{border:1.5px solid #000;text-align:center;color:#c00000;font-weight:700;padding:5px;margin:4mm 0}}
table.order{{font-size:7.5pt}}table.order td{{text-align:center}}
td.time{{white-space:nowrap;width:13.5mm;font-weight:700}}td.no{{font-weight:700;white-space:nowrap;width:13mm}}
tr.rec td{{border-bottom:2px solid #000}}tr.mv td{{color:#222}}
tr.sep td{{background:#e2e2e2;font-weight:700;font-size:8.6pt}}tr.sep td.time{{background:#fff}}
tr.pre td{{text-align:left;background:#f3f3f3}}tr.pre td.time{{text-align:center;background:#fff}}
.legend{{font-size:7pt;margin-top:1.5mm}}.mt{{margin-top:4mm}}.c{{text-align:center}}
.cols{{display:grid;grid-template-columns:1fr 1fr;gap:5mm;align-items:start}}.page{{break-before:page}}
.box{{border:1.5px solid #000;padding:5px 7px}}.box b.h{{display:block;margin-bottom:2px}}.box li{{margin:1.5px 0}}.box ul{{margin:0;padding-left:14px}}
</style></head><body>
<div class="titlebox"><div class="t">Kapaszkodni</div><div class="d">// {DATE}<br>Csütörtök //</div></div>
<table class="head"><tr>
 <td><div class="row"><b>Rendező:</b> {e(r["name"])}</div><div class="row">{e(r["phone"])}</div><div class="row"><b>Operatőr:</b> {e(o["name"])}</div><div class="row">{e(o["phone"])}</div></td>
 <td class="c"><span class="big">NAPI DISZPÓ</span><span class="big">Crew Call</span><span class="huge">{hm(D["plan"]["call"])}</span></td>
 <td><div class="row"><b>Első asszisztens:</b></div><div class="row">{e(a["name"])}</div><div class="row">{e(a["phone"])}</div></td>
</tr><tr>
 <td><div class="row"><b>HELYSZÍN:</b> stúdió</div><div class="row"><span class="line"></span></div><div class="row" style="margin-top:6px"><b>Parkolás:</b> <span class="line" style="min-width:28mm"></span></div></td>
 <td class="c"><span class="big">MUNKAIDŐ:</span><span class="huge">{e(D["info"].get("munkaidő", "").replace(" - ", " – "))}</span><span class="big">FORGATÁS:</span><span class="huge" style="margin-bottom:0">{hm(first_shoot)} – {hm(WRAP)}</span></td>
 <td><div class="row"><b>Ebéd:</b> {hm(lunch["s"])}–{hm(lunch["e"])}</div><div class="row">{lunch["d"]} perc</div><div class="row" style="margin-top:6px"><b>Wrap:</b> {hm(WRAP)}</div><div class="row">Stúdió, végig mesterséges fény</div></td>
</tr></table>
<div class="red">Technikai setup és berendezés: {hm(B[1]["s"])}–{hm(B[1]["e"])} · Színészek on set: {e(D["cast"][0]["onset"]) if D["cast"] else ""}</div>
<table class="order"><thead>
 <tr><th colspan="8" class="bar">FORGATÁSI SORREND</th></tr>
 <tr><th>IDŐ</th><th>SC /<br>SH</th><th>B/K<br>N/E</th><th>Mi<br>történik?</th><th style="width:29%">SZINOPSZIS</th><th>KAR.</th><th>Hely</th><th style="width:27%">MEGJEGYZÉS</th></tr>
</thead><tbody>{"".join(rows)}</tbody></table>
<div class="legend">B = belső · N = nappal · E = este · Karakter: {" · ".join(f'{n} = {e(c["role"])} ({e(c["name"])})' for n, c in enumerate(D["cast"], 1))} · Forrás: {e(STAMP)}</div>
<table class="mt"><thead><tr><th>ID</th><th>SZÍNÉSZ</th><th>KARAKTER</th><th>ÉRKEZÉS</th><th>RUHA</th><th>SMINK</th><th>HAJ</th><th>ON SET</th></tr></thead><tbody>{cast}</tbody></table>
<div class="page"></div>
<div class="cols"><div>
 <table><thead><tr><th>POZÍCIÓ</th><th>NÉV</th><th>ÉRKEZÉS</th></tr></thead><tbody>{crew}</tbody></table>
 <div class="box mt"><b class="h">FONTOS</b><ul>
  <li><b>Biztonság:</b> a rúd ráesése (2/6) csak üres vagy könnyített rúddal, mellette biztosító.</li>
  <li>Soma: a 2/2-ben legfeljebb 3 teljes erőkifejtéses take, a próbák könnyített rúddal.</li>
  <li>12K és kábelek: eligazítás reggel.</li></ul></div>
</div><div>
 <table><thead><tr><th colspan="3" class="bar">Kelléklista</th></tr><tr><th style="width:13mm">Jelenet</th><th>Kellék / berendezés</th><th style="width:8mm">✓</th></tr></thead>
 <tbody>{props}<tr><td>&nbsp;</td><td></td><td></td></tr><tr><td>&nbsp;</td><td></td><td></td></tr></tbody></table>
</div></div>
</body></html>"""

# ====================== 2) SHOTLIST ======================
def shotlist():
    rows, n = [], 0
    for b in B:
        if b["type"] == "lunch":
            rows.append(f'<tr class="sep"><td colspan="11">EBÉD · {hm(b["s"])}–{hm(b["e"])}</td></tr>'); continue
        if b["type"] != "shoot":
            continue
        ids = ids_of(b)
        for i in ids:
            n += 1; s = info(i)
            sb = next((x for x in D["shots"][b["label"]]["sb"] if x["id"] == i), None)
            pic = f'<img src="{sb["f"]}" alt="">' if sb else '<div class="none">nincs rajz</div>'
            v5 = " — ".join(x for x in (b["setup"], b["note"]) if x)
            cam = [f'<b>{e(s["lens"])}</b>' if s["lens"] else "", e(s["angle"]), e(" · ".join(x for x in (s["move"], s["sup"]) if x))]
            rows.append(f'''<tr><td class="ord">{n}</td><td class="tm">{hm(b["s"])}<br>–{hm(b["e"])}{"<small>egyben</small>" if len(ids) > 1 else ""}</td>
<td class="no">{"MONT." if i == "MONTAGE" else e(i)}<small>{e(s["loc"])} · {"este" if s["dn"] == "ESTE" else "nappal"}</small></td><td class="sb">{pic}</td>
<td class="desc"><b>{e(s["type"])}</b>{e(s["desc"])}{f'<span class="com">{e(s["com"])}</span>' if s.get("com") else ""}{f'<span class="v5">{e(v5)}</span>' if v5 else ""}</td>
<td class="spec">{"".join(f"<span>{x}</span>" for x in cam if x)}</td><td class="tal">{e(s["talent"])}</td>
<td class="eff{" uh" if s["eff"] == "Ultra High" else ""}">{e(s["eff"])}</td><td class="takes">1 2 3 4 5<br>6 7 8 9 10</td><td class="ok"></td><td></td></tr>''')
    return f"""<!doctype html><html lang="hu"><head><meta charset="utf-8"><title>Kapaszkodni – shotlist</title>{FONT}
<style>
@page{{size:A4 landscape;margin:8mm 8mm 11mm}}
*{{box-sizing:border-box}}html,body{{margin:0;background:#fff;color:#000}}
body{{font-family:"IBM Plex Sans Condensed","Arial Narrow",Arial,sans-serif;font-size:8.6pt;line-height:1.25}}
.top{{display:grid;grid-template-columns:auto 1fr auto;align-items:end;gap:6mm;border-bottom:2px solid #000;padding-bottom:2mm;margin-bottom:3mm}}
.top h1{{font-family:"Courier Prime",monospace;font-size:20pt;margin:0;letter-spacing:.03em}}
.top .meta{{font-family:"Courier Prime",monospace;font-size:9pt;line-height:1.35}}.top .legend{{font-size:7.4pt;text-align:right;color:#333}}
table{{border-collapse:collapse;width:100%}}th,td{{border:1px solid #000;padding:2.5px 4px;vertical-align:middle}}
thead th{{background:#c4c4c4;font-size:7.2pt;text-transform:uppercase;letter-spacing:.04em;text-align:center}}
thead{{display:table-header-group}}tr{{break-inside:avoid}}
td.ord{{width:6mm;text-align:center;font-family:"Courier Prime",monospace;color:#444}}
td.tm{{width:15mm;text-align:center;font-family:"Courier Prime",monospace;font-weight:700;font-size:9.5pt;line-height:1.15}}
td.tm small,td.no small{{display:block;font-size:6.8pt;font-weight:400}}
td.no{{width:19mm;text-align:center;font-family:"Courier Prime",monospace;font-weight:700;font-size:13pt;line-height:1.1}}
td.sb{{width:39mm;padding:1px;text-align:center}}td.sb img{{display:block;width:100%;height:27mm;object-fit:contain}}
td.sb .none{{height:27mm;display:flex;align-items:center;justify-content:center;color:#777;font-size:7pt}}
td.desc{{width:66mm}}td.desc b{{display:block;font-size:9pt}}
.com{{color:#c00000;font-weight:700;display:block;margin-top:2px}}.v5{{display:block;margin-top:3px;padding-top:2px;border-top:1px dotted #777;font-style:italic}}
td.spec{{width:30mm;font-size:7.8pt}}td.spec span{{display:block}}td.tal{{width:15mm;text-align:center}}
td.eff{{width:12mm;text-align:center;font-size:7.2pt;text-transform:uppercase}}td.eff.uh{{font-weight:700;color:#c00000}}
td.takes{{width:29mm;text-align:center;font-family:"Courier Prime",monospace;font-size:9.5pt;letter-spacing:.15em;line-height:1.7}}td.ok{{width:8mm}}
tr.sep td{{background:#e2e2e2;font-family:"Courier Prime",monospace;font-weight:700;text-align:center;padding:2px}}
</style></head><body>
<div class="top"><h1>KAPASZKODNI · shotlist</h1>
 <div class="meta">{DATE} csütörtök &nbsp;·&nbsp; Rendező: Ágoston András &nbsp;·&nbsp; DP: Détári Szabolcs &nbsp;·&nbsp; forgatási sorrend ({e(STAMP)})</div>
 <div class="legend">take-ek: karikázd a jót · dőlt betű: megjegyzés az időrendből</div></div>
<table><thead><tr><th>#</th><th>Idő</th><th>Sc / Shot</th><th>Storyboard</th><th>Leírás</th><th>Kamera</th><th>Szereplő</th><th>Effort</th><th>Take</th><th>OK</th><th>Szkriptes jegyzet</th></tr></thead>
<tbody>{"".join(rows)}</tbody></table></body></html>"""

# ====================== 3) FORGATÓKÖNYV ======================
def forgatokonyv():
    S = D["script"]
    scene, out, prev = 0, [], None
    def when(prefix):
        got = [b for b in SHOOT if any(x.startswith(prefix) or (prefix == "M" and x == "MONTAGE") for x in ids_of(b))]
        return " · ".join(f'{e(b["label"].split(" – ")[0])} {hm(b["s"])}' for b in sorted(got, key=lambda b: b["s"]))
    for l in S:
        t, x = l["t"], l["x"]
        if t == "sh":
            scene += 1
            num = str(scene) if scene <= 3 else "M"
            title = x if scene <= 3 else "BELSŐ. NAPPALI – NAPPAL · MONTÁZS"
            out.append(f'<div class="sh" data-n="{num}">{e(title)}</div>')
            w = when(f"{scene}/" if scene <= 3 else "M")
            out.append(f'<div class="shots">Forgatjuk: {w}</div>' if w else "")
            prev = None; continue
        if t == "end":
            out.append(f'<div class="end">{e(x)}</div>'); continue
        ids = list(dict.fromkeys(i for c in l["cov"] for i in c["ids"]))
        key = ",".join(ids)
        tag = ""
        if key != prev and ids:
            labs = [c["lab"] for c in l["cov"] if len(c["lab"]) <= 26 and not c["lab"].startswith(("a shotlist", "kiegészítő"))]
            tag = f'<span class="tag"><b>{e(" · ".join("MONT." if i == "MONTAGE" else i for i in ids))}</b>{("<i>" + e(labs[0]) + "</i>") if len(labs) == 1 else ""}</span>'
        prev = key
        cls = {"ac": "ac", "ch": "ch", "pa": "pa", "dl": "dl"}.get(t, "ac")
        out.append(f'<div class="{cls}">{tag}{e(x)}</div>')
    return f"""<!doctype html><html lang="hu"><head><meta charset="utf-8"><title>Kapaszkodni – forgatókönyv</title>{FONT}
<style>
@page{{size:A4;margin:22mm 12mm 20mm 12mm}}
*{{box-sizing:border-box}}html,body{{margin:0;background:#fff;color:#000}}
body{{padding:0 10mm 0 36mm;font-family:"Courier Prime","Courier New",Courier,monospace;font-size:12pt;line-height:1.18}}
.title{{height:250mm;margin-left:-26mm;display:flex;flex-direction:column;justify-content:center;align-items:center;text-align:center;position:relative}}
.title h1{{font-size:12pt;font-weight:700;text-decoration:underline;margin:0 0 12pt;letter-spacing:.05em}}.title p{{margin:0 0 12pt}}
.title .contact{{position:absolute;left:0;bottom:0;text-align:left;font-size:10pt}}.title .prod{{position:absolute;right:0;bottom:0;text-align:right;font-size:10pt}}
.sh{{font-weight:700;text-transform:uppercase;margin:24pt 0 4pt;position:relative;break-after:avoid}}
.sh::before{{content:attr(data-n);position:absolute;left:-36mm}}.sh::after{{content:attr(data-n);position:absolute;right:-9mm}}
#script>.sh:first-child{{margin-top:0}}
.shots{{font-family:"IBM Plex Sans Condensed",Arial,sans-serif;font-size:8.5pt;margin:0 0 12pt;border:1px solid #000;padding:2pt 5pt;display:inline-block;break-after:avoid}}
.ac,.ch,.pa,.dl{{position:relative}}
.ac{{margin:0 0 12pt}}.ch{{margin:12pt 0 0 37mm;text-transform:uppercase;break-after:avoid}}
.pa{{margin:0 0 0 27mm;width:60mm;break-after:avoid;break-before:avoid}}.dl{{margin:0 0 12pt 17mm;width:88mm;break-before:avoid}}
.end{{text-align:center;margin:24pt 0}}
.tag{{position:absolute;top:1pt;width:27mm;text-align:right;font-family:"IBM Plex Sans Condensed",Arial,sans-serif;font-size:7.6pt;line-height:1.15;text-transform:none;font-style:normal;padding-right:2.5mm;border-right:1.5px solid #000}}
.tag b{{display:block;font-weight:700}}.tag i{{display:block;font-style:italic;font-weight:400}}
.ac>.tag{{left:-30mm}}.ch>.tag{{left:-67mm}}.pa>.tag{{left:-57mm}}.dl>.tag{{left:-47mm}}
body.only-title #script,body.only-body .title{{display:none}}
</style></head><body>
<section class="title"><h1>KAPASZKODNI</h1><p>Írta</p><p>Ágoston András</p>
 <div class="contact">{e(AUTHOR[1])}<br>{e(AUTHOR[0])}</div><div class="prod">Forgatási példány · {DATE}<br>felsnitt és {e(STAMP)} szerint</div></section>
<div id="script">{"".join(out)}</div></body></html>"""

for name, fn in (("print_diszpo.html", diszpo), ("print_shotlist.html", shotlist), ("print_forgatokonyv.html", forgatokonyv)):
    open(os.path.join(HERE, name), "w", encoding="utf-8").write(fn())
    print("kész:", name)
