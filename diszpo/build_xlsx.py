# Szerkeszthető Excel-ütemterv az utemterv.html adataiból.
# Futtatás: python3 diszpo/build_xlsx.py  ->  diszpo/Kapaszkodni_utemterv_2026-11-12.xlsx
# Az Időrend lapon csak a „Perc” oszlopot kell átírni: a kezdés/vég időpontok és az összesítők képletből számolódnak.
import json, os, re, subprocess
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "Kapaszkodni_utemterv_2026-11-12.xlsx")

# Az R tömb kiolvasása az utemterv.html-ből (egy helyen kell javítani az időrendet)
js = r"""
const fs=require('fs');const s=fs.readFileSync(process.argv[1],'utf8');
const g=n=>eval(s.match(new RegExp('const '+n+'=(\\[[\\s\\S]*?\\n\\]);'))[1]);
console.log(JSON.stringify({R:g('R')}));
"""
data = json.loads(subprocess.check_output(["node", "-e", js, os.path.join(HERE, "utemterv.html")]))
R = data["R"]

TYPES = {"prep": "Előkészület", "pres": "Prezentálás", "shoot": "Forgatás", "move": "Átállás", "lunch": "Ebéd", "res": "Tartalék"}
FILL = {"Előkészület": "D9D9D9", "Prezentálás": "E4DAF0", "Forgatás": "D5E8DC", "Átállás": "FCEBC4", "Ebéd": "D6E2F2", "Tartalék": "EDEDED"}
GROUPS = [("K1", "Konyha, Fiú felől"), ("K2", "Konyha, ellenirány"), ("N1", "Nappali alapfény"),
          ("N2", "Nappali + prakti lámpa"), ("N3", "Nappali, Nagyi alulról"), ("E", "Nappali este")]

def clean(t):
    t = re.sub(r"<span[^>]*>(.*?)</span>", r"– \1", t or "")
    return re.sub(r"<[^>]+>", "", t).strip()

def mins(t):
    h, m = map(int, t.split(":"))
    return h * 60 + m

thin = Side(style="thin", color="999999")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
HEAD_FILL = PatternFill("solid", fgColor="C4C4C4")
BOLD = Font(bold=True)
WRAP = Alignment(wrap_text=True, vertical="top")
TIME = "h:mm"

def header(ws, row, titles, widths=None):
    for i, t in enumerate(titles, 1):
        c = ws.cell(row=row, column=i, value=t)
        c.font = BOLD; c.fill = HEAD_FILL; c.border = BOX
        c.alignment = Alignment(vertical="center", wrap_text=True)
    if widths:
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[chr(64 + i)].width = w

wb = Workbook()

# ---------- Időrend ----------
ws = wb.active
ws.title = "Időrend"
ws["A1"] = "KAPASZKODNI – napi ütemterv"; ws["A1"].font = Font(bold=True, size=16)
ws["A2"] = "2026.11.12. csütörtök · stúdió · Rendező: Ágoston András · DP: Détári Szabolcs · 2. asszisztens: Balázs · Script: Bence"
LBL = [("A3", "Crew call"), ("A4", "Tervezett wrap"), ("A5", "Számolt wrap"), ("A6", "Eltérés (perc)")]
for c, t in LBL:
    ws[c] = t; ws[c].font = BOLD
ws["B3"] = mins(R[0][0]) / 1440; ws["B3"].number_format = TIME
ws["B4"] = mins(R[-1][1]) / 1440; ws["B4"].number_format = TIME
for c in ("B3", "B4"):
    ws[c].fill = PatternFill("solid", fgColor="FFF2CC")
ws["D3"] = "Csak a sárga cellákat írd át. Az időrendben a Rászánt idő oszlopot,"
ws["D4"] = "a Kezdés és a Vég magától számolódik. Új sornál másold le a fölötte lévő sor képleteit."
for c in ("D3", "D4"):
    ws[c].font = Font(italic=True, color="666666")
HR = 8
header(ws, HR, ["Kezdés", "Vég", "Rászánt idő (perc)", "Típus", "Jel. / shot", "Beállítás", "Megjegyzés", "Fény"],
       [15, 9, 11, 13, 18, 34, 52, 8])
first = HR + 1
for i, r in enumerate(R):
    row = first + i
    ws.cell(row=row, column=1, value="=$B$3" if i == 0 else f"=B{row-1}")
    ws.cell(row=row, column=2, value=f"=A{row}+C{row}/1440")
    ws.cell(row=row, column=3, value=mins(r[1]) - mins(r[0]))
    ws.cell(row=row, column=4, value=TYPES[r[2]])
    ws.cell(row=row, column=5, value=clean(r[3]))
    ws.cell(row=row, column=6, value=clean(r[4]))
    ws.cell(row=row, column=7, value=clean(r[5]))
    ws.cell(row=row, column=8, value=r[6] if len(r) > 6 else None)
    for col in range(1, 9):
        c = ws.cell(row=row, column=col)
        c.border = BOX; c.alignment = WRAP
    pc = ws.cell(row=row, column=3)
    pc.fill = PatternFill("solid", fgColor="FFF2CC"); pc.font = BOLD; pc.alignment = Alignment(horizontal="center", vertical="top")
    ws.cell(row=row, column=1).number_format = TIME
    ws.cell(row=row, column=2).number_format = TIME
    if r[2] in ("shoot", "pres", "lunch"):
        ws.cell(row=row, column=5).font = BOLD
last = first + len(R) - 1
wrow = last + 1
ws.cell(row=wrow, column=1, value=f"=B{last}").number_format = TIME
ws.cell(row=wrow, column=5, value="WRAP")
ws.cell(row=wrow, column=7, value="Bontás, pakolás, napi riport (Bence)")
for col in range(1, 9):
    c = ws.cell(row=wrow, column=col)
    c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="6B6B6B"); c.border = BOX
ws["B5"] = f"=B{last}"; ws["B5"].number_format = TIME; ws["B5"].font = BOLD
ws["B6"] = "=ROUND((B5-B4)*1440,0)"
ws.conditional_formatting.add("B6", FormulaRule(formula=["$B$6>0"], font=Font(bold=True, color="B3261E")))

dv = DataValidation(type="list", formula1='"' + ",".join(TYPES.values()) + '"', allow_blank=True)
ws.add_data_validation(dv); dv.add(f"D{first}:D{last+40}")
for name, color in FILL.items():
    ws.conditional_formatting.add(f"A{first}:B{last+40} D{first}:H{last+40}",
        FormulaRule(formula=[f'$D{first}="{name}"'], fill=PatternFill("solid", fgColor=color)))
ws.freeze_panes = f"A{first}"
ws.page_setup.orientation = "landscape"; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.print_title_rows = f"{HR}:{HR}"

RNG = lambda col: f"'Időrend'!${col}${first}:${col}${last+40}"

# ---------- Összesítés ----------
s = wb.create_sheet("Összesítés")
s["A1"] = "Összesítés"; s["A1"].font = Font(bold=True, size=14)
header(s, 3, ["Típus", "Perc", "Óra:perc"], [28, 10, 10])
row = 4
for name in TYPES.values():
    s.cell(row=row, column=1, value=name)
    s.cell(row=row, column=2, value=f'=SUMIF({RNG("D")},A{row},{RNG("C")})')
    s.cell(row=row, column=3, value=f"=B{row}/1440").number_format = "[h]:mm"
    row += 1
s.cell(row=row, column=1, value="Összesen").font = BOLD
s.cell(row=row, column=2, value=f"=SUM(B4:B{row-1})").font = BOLD
s.cell(row=row, column=3, value=f"=B{row}/1440").number_format = "[h]:mm"
row += 2
s.cell(row=row, column=1, value="Crew call"); s.cell(row=row, column=2, value="='Időrend'!B3").number_format = TIME
s.cell(row=row + 1, column=1, value="Tervezett wrap"); s.cell(row=row + 1, column=2, value="='Időrend'!B4").number_format = TIME
s.cell(row=row + 2, column=1, value="Számolt wrap"); s.cell(row=row + 2, column=2, value="='Időrend'!B5").number_format = TIME
s.cell(row=row + 3, column=1, value="Eltérés (perc, + = túllépés)"); s.cell(row=row + 3, column=2, value="='Időrend'!B6")
s.cell(row=row + 4, column=1, value="Forgatási sorok száma")
s.cell(row=row + 4, column=2, value=f'=COUNTIF({RNG("D")},"Forgatás")')

# ---------- Fénybeállítások ----------
f = wb.create_sheet("Fénybeállítások")
f["A1"] = "Fénybeállítások (vilkó)"; f["A1"].font = Font(bold=True, size=14)
header(f, 3, ["Kód", "Fény", "Shotok", "Forgatás (perc)", "Átállás (perc)", "Összesen"], [7, 26, 46, 14, 14, 11])
row = 4
for code, name in GROUPS:
    shots = [clean(r[3]).replace(" – Ultra High", "") for r in R if len(r) > 6 and r[6] == code and r[2] == "shoot"]
    f.cell(row=row, column=1, value=code); f.cell(row=row, column=2, value=name); f.cell(row=row, column=3, value=", ".join(shots))
    f.cell(row=row, column=4, value=f'=SUMIFS({RNG("C")},{RNG("H")},A{row},{RNG("D")},"Forgatás")')
    f.cell(row=row, column=5, value=f'=SUMIFS({RNG("C")},{RNG("H")},A{row},{RNG("D")},"Átállás")')
    f.cell(row=row, column=6, value=f"=D{row}+E{row}")
    for col in range(1, 7):
        f.cell(row=row, column=col).border = BOX
    row += 1

# ---------- Optika ----------
o = wb.create_sheet("Optika")
o["A1"] = "Optika: shotlist / vilkó"; o["A1"].font = Font(bold=True, size=14)
header(o, 3, ["Shot", "Shotlist", "Vilkó", "Végleges (Szabolcs)"], [14, 12, 12, 22])
for i, (a, b, c) in enumerate([("1/3A, 1/3B", "75 mm", "50 mm"), ("2/1B, 2/1C", "35 mm", "50 mm"),
                               ("2/5", "50 mm", "75 mm"), ("3/1, 3/2", "50 mm", "30 mm")]):
    for j, v in enumerate((a, b, c, None), 1):
        o.cell(row=4 + i, column=j, value=v).border = BOX

# ---------- Megjegyzések ----------
n = wb.create_sheet("Megjegyzések")
n.column_dimensions["A"].width = 100
blocks = [
    ("Megjegyzések", [
        "Rúd ráesik (2. jel.): csak üres vagy könnyített rúddal, mellette biztosító.",
        "Soma: a 2/2-ben max. 3 teljes take, próba könnyített rúddal. Közeliken műanyag tárcsák.",
        "1/1: legalább 3 sor teli és üres tányér.",
        "A nappali 11:00-ra legyen berendezve (1/3B-ben látszik a pad).",
        "Plafon-anyag: reggel, a konyhai forgatás alatt (Kristóf, Gyuri).",
        "A 2/4 sínje lerakható a 2/1C alatt, ha nem látszik.",
        "Konyhai fények bontása az esti jelenet alatt."]),
    ("Ha csúszunk", [
        "1. Tartalékok felhasználása (15′ + 5′).",
        "2. Kiegészítő jelenet kimarad (–10′).",
        "3. 2/1E B kamerával a 2/1B alatt (–15′).",
        "4. 2/5 B kamerával a 2/4 alatt (–20′).",
        "5. 3/2 a 3/1-gyel egy beállításból (–10′)."]),
    ("Nyitott kérdések", [
        "B kamera egész nap? (Szabolcs)",
        "Optikák: lásd az Optika lapot. (Szabolcs)",
        "Plafon-anyag csak a 2/3A–B-hez kell? (Szabolcs)",
        "Kiegészítő jelenet (bicepsz, nyújtás) marad? (Andris)",
        "3/2: alulról vagy eye level? 1/3A: dolly mozog vagy static? (Andris)",
        "Stúdió címe, parkolás, meddig van meg a stúdió (bontás 45–60′). (gyártás)",
        "Színészhívás 9:00, smink Kata? (gyártás)"]),
]
row = 1
for title, items in blocks:
    n.cell(row=row, column=1, value=title).font = Font(bold=True, size=13)
    row += 1
    for it in items:
        n.cell(row=row, column=1, value=it).alignment = WRAP
        row += 1
    row += 1

wb.calculation.fullCalcOnLoad = True
wb.save(OUT)
import shutil, tempfile
tmp = tempfile.mkdtemp()
subprocess.run(["soffice", "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", tmp, OUT],
               check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
shutil.move(os.path.join(tmp, os.path.basename(OUT)), OUT)
print("kész:", OUT)
