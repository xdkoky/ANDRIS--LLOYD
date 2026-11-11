# Az 1st AD forgatási app adatai a szerkesztett táblázatból.
# Futtatás: python3 diszpo/build_ad_app.py [táblázat.xlsx]
#   bemenet: ad_app_template.html  ->  kimenet: ad_app.html (ezt publikáljuk; képek: storyboard/, vilko/)
import glob, json, os, re, shutil, subprocess, sys, tempfile
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from openpyxl import load_workbook
from felsnitt import SEG, el

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.abspath(sys.argv[1]) if len(sys.argv) > 1 else max(
    glob.glob(os.path.join(HERE, "Kapaszkodni_utemterv_V*.xlsx")), key=lambda p: int(re.search(r"_V(\d+)\.xlsx$", p).group(1)))

tmp = tempfile.mkdtemp()
subprocess.run(["soffice", "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", tmp, SRC],
               check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
ws = load_workbook(os.path.join(tmp, os.path.basename(SRC)), data_only=True)["Időrend"]
shutil.rmtree(tmp)

def tm(v):
    return None if v is None else v.hour * 60 + v.minute
def txt(v):
    return "" if v is None else re.sub(r"\s+", " ", str(v)).strip()

TYPES = {"Előkészület": "prep", "Prezentálás": "pres", "Forgatás": "shoot", "Átállás": "move", "Ebéd": "lunch", "Tartalék": "res"}
HR = next(r for r in range(1, 60) if txt(ws.cell(r, 1).value) == "Kezdés")

# fejléc-információk (a táblázat fölötti blokk)
info, contacts = {}, []
for r in range(1, HR):
    a, d = txt(ws.cell(r, 1).value), txt(ws.cell(r, 4).value)
    if d.endswith(":"):
        info[d[:-1].lower()] = txt(ws.cell(r + 1, 4).value)
    if ":" in a and r > 2 and not a.lower().startswith(("crew call", "tervezett", "számolt", "túllépés", "plusz")):
        role, name = [x.strip() for x in a.split(":", 1)]
        phone = txt(ws.cell(r + 1, 1).value)
        contacts.append({"role": role, "name": name, "phone": phone if re.search(r"\d{3}", phone) else ""})

blocks, r = [], HR + 1
while r <= ws.max_row:
    label = txt(ws.cell(r, 7).value)
    if label.upper() == "WRAP":
        wrap_row = r
        break
    typ = txt(ws.cell(r, 6).value)
    if typ in TYPES and ws.cell(r, 3).value not in (None, ""):
        blocks.append({"s": tm(ws.cell(r, 1).value), "d": int(ws.cell(r, 3).value), "type": TYPES[typ], "label": label,
                       "setup": txt(ws.cell(r, 9).value).replace("–", "").strip(), "note": txt(ws.cell(r, 10).value).replace("–", "").strip(),
                       "group": txt(ws.cell(r, 11).value)})
    r += 1

cast, crew, dept = [], [], ""
for rr in range(wrap_row + 1, ws.max_row + 1):
    a = txt(ws.cell(rr, 1).value)
    if a == "SZÍNÉSZ":
        k = rr + 1
        while txt(ws.cell(k, 1).value):
            g = lambda c: (ws.cell(k, c).value.strftime("%H:%M") if hasattr(ws.cell(k, c).value, "strftime") else txt(ws.cell(k, c).value))
            cast.append({"name": txt(ws.cell(k, 1).value).title(), "role": txt(ws.cell(k, 3).value).capitalize(),
                         "arr": g(5), "ruha": g(6), "smink": g(7), "haj": g(8), "onset": g(9)})
            k += 1
    if a == "POZÍCIÓ":
        for k in range(rr + 1, ws.max_row + 1):
            pos, name = txt(ws.cell(k, 1).value), txt(ws.cell(k, 4).value)
            if not pos:
                continue
            if not name:
                dept = pos.capitalize(); continue
            v = ws.cell(k, 7).value
            crew.append({"dept": dept, "pos": pos if "AD" in pos.split() else pos.capitalize(), "name": name.title(),
                         "arr": v.strftime("%H:%M") if hasattr(v, "strftime") else txt(v)})
        break

# ---- shotonkénti kiegészítő tudás ----
def ids_of(label):
    if label.upper().startswith("MONTAGE"):
        return ["MONTAGE"]
    return re.findall(r"\d/\d[A-E]?", label)

TALENT = {"1/1": ["Fiú", "Nagyi"], "1/3A": ["Fiú"], "1/3B": ["Fiú"], "1/2": ["Nagyi"], "2/2": ["Fiú", "Nagyi"],
          "MONTAGE": ["Fiú", "Nagyi"], "2/3A": ["Nagyi"], "2/3B": ["Nagyi"], "2/5": ["Nagyi", "Fiú"], "2/4": ["Fiú", "Nagyi"],
          "2/1C": ["Fiú"], "2/1B": ["Fiú"], "2/1A": ["Fiú"], "2/1E": ["Fiú"], "2/6": ["Fiú", "Nagyi"], "3/1": ["Nagyi"], "3/2": ["Nagyi"]}
OFFCAM = {"1/3A": "Nagyi a kamera mögül adja a szöveget", "1/2": "a Fiú a kamera mögül adja a szöveget",
          "2/3A": "a Fiú a pad alatt / kamera mögött", "2/1C": "Nagyi szövege a kamera mellől",
          "2/1B": "Nagyi szövege a kamera mellől", "2/1A": "Nagyi számolása a kamera mellől"}
PROPS = {
    "1/1": ["Teli tálak, több take-re elég étel", "Sok üres tányér az asztalon", "Villa", "A srác szája maszatos", "Take-enként tányérok vissza"],
    "1/3A": ["Tányérok úgy, ahogy az 1/1 végén", "Kanál", "Cetli és ceruza Nagyi zsebében", "Mosogatószivacs + tányér Nagyinál"],
    "1/3B": ["Az ajtón át a nappaliban látszik a fekvenyomó pad: legyen berendezve"],
    "1/2": ["Mosogatószivacs + tányér", "Cetli és ceruza", "Vetítés: fontos!"],
    "2/2": ["Fekvenyomó pad, rúd, könnyített súlyok", "Pótsúly a földön a pad két oldalán", "Soma: max. 3 teljes erőkifejtéses take"],
    "MONTAGE": ["Kézisúlyok a földön", "Matrac / polifom a nyújtáshoz"],
    "2/3A": ["Plafon-anyag kifeszítve", "Nagyi szemüvege láncon"],
    "2/3B": ["Plafon-anyag kifeszítve"],
    "2/5": ["Ölelés ugyanabban a pozícióban, mint a 2/4-ben"],
    "2/4": ["2,4 m sín, A + B kamera", "Könnyített súlyok a rúdon", "Oda-vissza dolly: felülés → ölelés → visszafekvés"],
    "2/1C": ["Felülő pozíció = a 2/4 közepe"],
    "2/1B": ["Fekvő pozíció, súly a rúdon", "Szövegjavítás: »lefekszik«, »lock in«"],
    "2/1A": ["Csak a jelenet eleje kell", "Izzadás fokozódik"],
    "2/1E": ["Csak a jelenet vége kell", "A kar belóg az arc elé"],
    "2/6": ["Magic arm a rúdon", "BIZTONSÁG: a rúd ráesik – csak üres/könnyített rúd, biztosító mellette!"],
    "3/1": ["Sötét szoba, csak a TV fénye", "TV-n: régi súlyemelő sporttudósítás", "Gyakorlati lámpa"],
    "3/2": ["Sötét szoba, TV-fény", "Ugyanaz a jelmez/haj, mint a 3/1-ben"],
}
STORY = {
    "1/1": "1. jelenet eleje: a Fiú eszik, Nagyi a negyedik tálat hozza, majd leül mellé.",
    "1/3A": "A teljes konyhai jelenet a Fiú arcán.", "1/3B": "A konyhai jelenet vége: a Fiú hátrafordul, a nappaliban a pad (átmenet a 2. jelenetbe).",
    "1/2": "A teljes konyhai jelenet Nagyi oldaláról, az ordításig.",
    "2/2": "Nappali, tág: az első sorozat vége, majd együtt rakják fel a plusz súlyt.",
    "MONTAGE": "Kiegészítő: bicepsz a kézisúllyal, nyújtás a matracon, Nagyi segít.",
    "2/3A": "Nagyi reakciói alulról a nappali jelenet egészében.", "2/3B": "A film vége: Nagyi rémült reakciója a balesetre.",
    "2/5": "Az ölelés után Nagyi arca: »Újra!«", "2/4": "A hatodiknál elbukik, felül, ölelés, »Újra!«, visszafekszik.",
    "2/1C": "Felülve: »Mennyi volt?« – »Hét bogaram.«", "2/1B": "Fekve: sorozatok, »Sakkozik«, »Nem«, lock in.",
    "2/1A": "A nappali jelenet nyitánya és a hatodik ismétlés küzdelme.", "2/1E": "Az utolsó push, csak az arc.",
    "2/6": "A baleset: a rúd ráesik (a rúd nézőpontjából).", "3/1": "Este: Nagyi egyedül nézi a TV-t.",
    "3/2": "Este: Nagyi végső reakciója.",
}
GROUPS = {
    "K1": ["Konyha, Fiú felől", "Ablakfények (ARRI T12 12K, T5 5K, 3× T2), 5 PavoTube (4 Half Skirt, 1 SnapGrid), LiteMat 4 + 2. Sín: 1/1-hez 2,4 m, 1/3-hoz 1,2 m."],
    "K2": ["Konyha, ellenirány", "Ugyanazok az ablakfények; PavoTube-ok átrendezve (3 Half Skirt, 2 SnapGrid); sín nincs."],
    "N1": ["Nappali alapfény", "Ablakfények + 2 PavoTube (Half Skirt), LiteMat 2. A gyakorlati lámpa NEM ég."],
    "N2": ["Nappali + gyakorlati lámpa", "+1 PavoTube (20%) és Andris gyakorlati lámpája BE. A 2/4-hez B kamera is a sínen."],
    "N3": ["Nagyi alulról", "Gyakorlati lámpa KI, PavoTube állványon áthelyezve, 2× 4×4 Full Frost, 2× 18×24 Double Net. Plafon-anyag kell."],
    "E": ["Este", "Nagy ablakfények KI. 6 PavoTube, LiteMat 4, Aputure MC Pro, gyakorlati 60W, TV-tartalom. 3/2: csak 3 PavoTube + prakti."],
    "N1 + E": ["Nappali alap + esti előkészítve", "Reggel a nappali esti fénye előre beállítva, közben a konyha nappali fénye."],
}
def img(kind, i):
    """Előnézet (t) és teljes minőségű (f) kép egy shothoz; a teljes a build_images.py-ból jön."""
    n = i.replace("/", "-") + ".jpg"
    ex = lambda *a: os.path.exists(os.path.join(HERE, kind, *a))
    if not (ex(n) or ex("full", n)):
        return None
    return {"id": i, "t": f"{kind}/t/{n}" if ex("t", n) else f"{kind}/{n}", "f": f"{kind}/full/{n}" if ex("full", n) else f"{kind}/{n}"}

def lines(a, b):
    return [{"t": el(i)[0], "x": el(i)[1]} for i in range(a, b + 1)]

shots = {}
for b in blocks:
    if b["type"] != "shoot":
        continue
    ids = ids_of(b["label"])
    segs = [{"lab": lab, "a": a, "b": bb, "ids": [x for x in sh if x in ids], "lines": lines(a, bb)}
            for a, bb, lab, sh in SEG if set(sh) & set(ids)]
    talent = list(dict.fromkeys(t for i in ids for t in TALENT.get(i, [])))
    shots[b["label"]] = {
        "ids": ids, "talent": talent, "offcam": [OFFCAM[i] for i in ids if i in OFFCAM],
        "props": [p for i in ids for p in PROPS.get(i, [])], "story": " ".join(STORY[i] for i in ids if i in STORY),
        "sb": [x for x in (img("storyboard", i) for i in ids) if x], "vk": [x for x in (img("vilko", i) for i in ids) if x],
        "segs": segs, "pos": min((s["a"] for s in segs), default=999),
    }

# a teljes szkript (sztori-sorrend) a SZKRIPT nézethez: elemek + melyik shotok fedik
from felsnitt import SCRIPT
cover = {}
for a, bb, lab, sh in SEG:
    for i in range(a, bb + 1):
        cover.setdefault(i, []).append({"lab": lab, "ids": sh})
script = [{"t": el(i)[0], "x": el(i)[1], "cov": cover.get(i, [])} for i in range(len(SCRIPT))]

vilko_all = []
for b in blocks:
    if b["type"] == "shoot":
        for x in shots[b["label"]]["vk"]:
            if x not in vilko_all:
                vilko_all.append(x)
boards_all = []
for b in blocks:
    if b["type"] == "shoot":
        for x in shots[b["label"]]["sb"]:
            if x not in boards_all:
                boards_all.append(x)
DATA = {"vilko": vilko_all, "boards": boards_all, "src": os.path.basename(SRC), "date": "2026-11-12", "info": info, "contacts": contacts,
        "plan": {"call": blocks[0]["s"], "targetWrap": tm(ws["B4"].value)},
        "blocks": blocks, "shots": shots, "groups": GROUPS, "cast": cast, "crew": crew, "script": script}
tpl = open(os.path.join(HERE, "ad_app_template.html"), encoding="utf-8").read()
out = tpl.replace("/*__DATA__*/null", json.dumps(DATA, ensure_ascii=False))
open(os.path.join(HERE, "ad_app.html"), "w", encoding="utf-8").write(out)
print(f"kész: {len(blocks)} blokk, {len(shots)} shot, {len(cast)} színész, {len(crew)} stábtag, forrás: {DATA['src']}")
print("info:", info); print("contacts:", contacts)
