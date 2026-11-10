# Az ütemterv-PDF adatait a szerkesztett táblázatból veszi át.
# Futtatás: python3 diszpo/sync_utemterv.py [táblázat.xlsx]   (alapból a legújabb Kapaszkodni_utemterv_V*.xlsx)
# Utána: node diszpo/build_pdf.js  ->  Kapaszkodni_utemterv_2026-11-12.pdf
import glob, html, json, os, re, shutil, subprocess, sys, tempfile
from openpyxl import load_workbook

HERE = os.path.dirname(os.path.abspath(__file__))
HTML = os.path.join(HERE, "utemterv.html")
if len(sys.argv) > 1:
    SRC = os.path.abspath(sys.argv[1])
else:
    SRC = max(glob.glob(os.path.join(HERE, "Kapaszkodni_utemterv_V*.xlsx")),
              key=lambda p: int(re.search(r"_V(\d+)\.xlsx$", p).group(1)))

TYPES = {"Előkészület": "prep", "Prezentálás": "pres", "Forgatás": "shoot", "Átállás": "move", "Ebéd": "lunch", "Tartalék": "res"}

# képletek kiértékelése LibreOffice-szal (csak olvasáshoz)
tmp = tempfile.mkdtemp()
subprocess.run(["soffice", "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", tmp, SRC],
               check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
ws = load_workbook(os.path.join(tmp, os.path.basename(SRC)), data_only=True)["Időrend"]
shutil.rmtree(tmp)

t = lambda v: v.strftime("%H:%M")
esc = lambda v: html.escape(str(v).strip()) if v not in (None, "") else ""
R, CREW, r = [], [], 9
while r <= ws.max_row:
    label = ws.cell(r, 7).value
    if str(label).strip().upper() == "WRAP":
        break
    typ = ws.cell(r, 6).value
    if typ in TYPES and ws.cell(r, 1).value is not None and ws.cell(r, 3).value not in (None, ""):
        row = [t(ws.cell(r, 1).value), t(ws.cell(r, 2).value), TYPES[typ], esc(label), esc(ws.cell(r, 9).value) or "–",
               esc(ws.cell(r, 10).value)]
        if ws.cell(r, 11).value:
            row.append(str(ws.cell(r, 11).value).strip())
        R.append(row)
    r += 1
for rr in range(r + 1, ws.max_row + 1):
    for col in (2, 7):
        v = ws.cell(rr, col).value
        if v and ":" in str(v) and not str(v).strip().endswith(":"):
            CREW.append(esc(re.sub(r"\s+", " ", str(v))))
PLAN = {"call": t(ws["B3"].value), "wrap": t(ws["B4"].value),
        "src": f"a szerkesztett táblázat alapján ({os.path.basename(SRC)})"}

block = ("// ADAT-ELEJE (a sync_utemterv.py írja felül a táblázatból)\n"
         "// [kezdés, vég, típus, feladat, beállítás, megjegyzés, fénycsoport]\nconst R=[\n"
         + ",\n".join(" " + json.dumps(x, ensure_ascii=False) for x in R)
         + "\n];\nconst CREW=" + json.dumps(CREW, ensure_ascii=False)
         + ";\nconst PLAN=" + json.dumps(PLAN, ensure_ascii=False) + ";\n// ADAT-VEGE")
s = open(HTML, encoding="utf-8").read()
s = re.sub(r"// ADAT-ELEJE.*?// ADAT-VEGE", lambda _: block, s, flags=re.S)
open(HTML, "w", encoding="utf-8").write(s)
print(f"kész: {len(R)} sor, {len(CREW)} stábtag, forrás: {os.path.basename(SRC)}")
