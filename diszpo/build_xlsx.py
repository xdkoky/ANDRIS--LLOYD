# Szerkeszthető Excel-ütemterv az utemterv.html adataiból.
# Futtatás: python3 diszpo/build_xlsx.py  ->  diszpo/Kapaszkodni_utemterv_2026-11-12.xlsx
# Az Időrend lapon csak a „Rászánt idő” oszlopot kell átírni: a kezdés/vég időpontok és az összesítők képletből számolódnak.
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
YEL = PatternFill("solid", fgColor="FFF2CC")
RED_FILL = PatternFill("solid", fgColor="F4C7C3")

# ---------- Időrend ----------
ws = wb.active
ws.title = "Időrend"
ws["A1"] = "KAPASZKODNI – napi ütemterv"; ws["A1"].font = Font(bold=True, size=16)
ws["A2"] = "2026.11.12. csütörtök · stúdió · Rendező: Ágoston András · DP: Détári Szabolcs · 2. asszisztens: Balázs · Script: Bence"
for c, t in [("A3", "Crew call"), ("A4", "Tervezett wrap"), ("A5", "Számolt wrap"), ("A6", "Túllépés (perc)")]:
    ws[c] = t; ws[c].font = BOLD
ws["B3"] = mins(R[0][0]) / 1440; ws["B3"].number_format = TIME
ws["B4"] = mins(R[-1][1]) / 1440; ws["B4"].number_format = TIME
for c in ("B3", "B4"):
    ws[c].fill = YEL
ws["D3"] = "Csak a sárga cellákat írd át (Rászánt idő, crew call, tervezett wrap). A Kezdés és a Vég ebből számolódik."
ws["D4"] = "Az Eltérés oszlop mutatja, hol adtatok több (piros) vagy kevesebb (zöld) időt, mint az eredeti terv."
ws["D5"] = "A wrap után eső sorok pirosak: ennyit kell lecsípni. Ötletek az Összesítés lapon."
for c in ("D3", "D4", "D5"):
    ws[c].font = Font(italic=True, color="666666")
HR = 8
header(ws, HR, ["Kezdés", "Vég", "Rászánt idő (perc)", "Eredeti terv (perc)", "Eltérés (perc)", "Típus",
                "Jel. / shot", "Beállítás", "Megjegyzés", "Fény"],
       [15, 9, 11, 11, 10, 13, 18, 34, 46, 7])
first = HR + 1
for i, r in enumerate(R):
    row = first + i
    minutes = mins(r[1]) - mins(r[0])
    ws.cell(row=row, column=1, value="=$B$3" if i == 0 else f"=A{row-1}+C{row-1}/1440")
    ws.cell(row=row, column=2, value=f"=A{row}+C{row}/1440")
    ws.cell(row=row, column=3, value=minutes)
    ws.cell(row=row, column=4, value=minutes)
    ws.cell(row=row, column=5, value=f"=C{row}-D{row}")
    ws.cell(row=row, column=6, value=TYPES[r[2]])
    ws.cell(row=row, column=7, value=clean(r[3]))
    ws.cell(row=row, column=8, value=clean(r[4]))
    ws.cell(row=row, column=9, value=clean(r[5]))
    ws.cell(row=row, column=10, value=r[6] if len(r) > 6 else None)
    for col in range(1, 11):
        c = ws.cell(row=row, column=col)
        c.border = BOX; c.alignment = WRAP
    pc = ws.cell(row=row, column=3)
    pc.fill = YEL; pc.font = BOLD; pc.alignment = Alignment(horizontal="center", vertical="top")
    for col in (4, 5):
        ws.cell(row=row, column=col).alignment = Alignment(horizontal="center", vertical="top")
    ws.cell(row=row, column=4).font = Font(color="777777")
    ws.cell(row=row, column=1).number_format = TIME
    ws.cell(row=row, column=2).number_format = TIME
    ws.cell(row=row, column=5).number_format = '+0;-0;0'
    if r[2] in ("shoot", "pres", "lunch"):
        ws.cell(row=row, column=7).font = BOLD
last = first + len(R) - 1
wrow = last + 1
ws.cell(row=wrow, column=1, value=f"=B{last}").number_format = TIME
ws.cell(row=wrow, column=3, value=f"=SUM(C{first}:C{last})")
ws.cell(row=wrow, column=4, value=f"=SUM(D{first}:D{last})")
ws.cell(row=wrow, column=5, value=f"=C{wrow}-D{wrow}").number_format = '+0;-0;0'
ws.cell(row=wrow, column=7, value="WRAP")
ws.cell(row=wrow, column=9, value="Bontás, pakolás, napi riport (Bence)")
for col in range(1, 11):
    c = ws.cell(row=wrow, column=col)
    c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="6B6B6B"); c.border = BOX
ws["B5"] = f"=B{last}"; ws["B5"].number_format = TIME; ws["B5"].font = BOLD
ws["B6"] = "=MAX(0,ROUND((B5-B4)*1440,0))"; ws["B6"].font = BOLD
ws["C6"] = '=IF(B6>0,"Túl van az időn – csípj le "&B6&" percet","Belefér")'
ws.conditional_formatting.add("B6:C6", FormulaRule(formula=["$B$6>0"], font=Font(bold=True, color="B3261E"), fill=RED_FILL))
ws.conditional_formatting.add("B6:C6", FormulaRule(formula=["$B$6=0"], font=Font(bold=True, color="1E7B3C")))

END = last + 40
# 1. wrap után eső sor: piros (ez az első szabály, ez nyer)
ws.conditional_formatting.add(f"A{first}:J{END}",
    FormulaRule(formula=[f'AND($C{first}<>"",$B{first}>$B$4+0.00001)'], fill=RED_FILL, font=Font(color="B3261E", bold=True), stopIfTrue=True))
# 2. eltérés színe
ws.conditional_formatting.add(f"E{first}:E{END}", FormulaRule(formula=[f"$E{first}>0"], font=Font(bold=True, color="B3261E")))
ws.conditional_formatting.add(f"E{first}:E{END}", FormulaRule(formula=[f"$E{first}<0"], font=Font(bold=True, color="1E7B3C")))
# 3. típus szerinti háttér
for name, color in FILL.items():
    ws.conditional_formatting.add(f"A{first}:B{END} D{first}:J{END}",
        FormulaRule(formula=[f'$F{first}="{name}"'], fill=PatternFill("solid", fgColor=color)))
dv = DataValidation(type="list", formula1='"' + ",".join(TYPES.values()) + '"', allow_blank=True)
ws.add_data_validation(dv); dv.add(f"F{first}:F{END}")
ws.freeze_panes = f"A{first}"
ws.page_setup.orientation = "landscape"; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.print_title_rows = f"{HR}:{HR}"

RNG = lambda col: f"'Időrend'!${col}${first}:${col}${END}"

# ---------- Összesítés ----------
s = wb.create_sheet("Összesítés")
s["A1"] = "Összesítés"; s["A1"].font = Font(bold=True, size=14)
header(s, 3, ["Típus", "Eredeti (perc)", "Most (perc)", "Eltérés"], [34, 14, 12, 10])
row = 4
for name in TYPES.values():
    s.cell(row=row, column=1, value=name)
    s.cell(row=row, column=2, value=f'=SUMIF({RNG("F")},A{row},{RNG("D")})')
    s.cell(row=row, column=3, value=f'=SUMIF({RNG("F")},A{row},{RNG("C")})')
    s.cell(row=row, column=4, value=f"=C{row}-B{row}").number_format = '+0;-0;0'
    row += 1
tot = row
s.cell(row=tot, column=1, value="Összesen").font = BOLD
for col, L in ((2, "B"), (3, "C"), (4, "D")):
    s.cell(row=tot, column=col, value=f"=SUM({L}4:{L}{tot-1})").font = BOLD
s.cell(row=tot, column=4).number_format = '+0;-0;0'
s.conditional_formatting.add(f"D4:D{tot}", FormulaRule(formula=["D4>0"], font=Font(bold=True, color="B3261E")))
s.conditional_formatting.add(f"D4:D{tot}", FormulaRule(formula=["D4<0"], font=Font(bold=True, color="1E7B3C")))
row = tot + 2
s.cell(row=row, column=1, value="Wrap").font = Font(bold=True, size=12); row += 1
pairs = [("Crew call", "='Időrend'!B3", TIME), ("Tervezett wrap", "='Időrend'!B4", TIME),
         ("Számolt wrap", "='Időrend'!B5", TIME), ("Túllépés (perc)", "='Időrend'!B6", None)]
for label, fml, fmt in pairs:
    s.cell(row=row, column=1, value=label); c = s.cell(row=row, column=2, value=fml)
    if fmt: c.number_format = fmt
    row += 1
row += 1
s.cell(row=row, column=1, value="Honnan lehet lecsípni?").font = Font(bold=True, size=12); row += 1
cut = [("Tartalék (teljesen elvehető)", f'=SUMIF({RNG("F")},"Tartalék",{RNG("C")})'),
       ("Kiegészítő jelenet (kihagyható)", f'=SUMIF({RNG("G")},"Kiegészítő",{RNG("C")})'),
       ("2/1E B kamerával a 2/1B alatt (ha van B kamera)", "=15"),
       ("2/5 B kamerával a 2/4 alatt (ha van B kamera)", "=20"),
       ("3/2 a 3/1-gyel egy beállításból", "=10")]
cs = row
for label, fml in cut:
    s.cell(row=row, column=1, value=label); s.cell(row=row, column=2, value=fml); row += 1
s.cell(row=row, column=1, value="Összesen lecsípható").font = BOLD
s.cell(row=row, column=2, value=f"=SUM(B{cs}:B{row-1})").font = BOLD
s.cell(row=row + 1, column=1, value="Ezután is túllépés (perc)").font = BOLD
s.cell(row=row + 1, column=2, value=f"=MAX(0,'Időrend'!B6-B{row})").font = BOLD
s.conditional_formatting.add(f"B{row+1}", FormulaRule(formula=[f"B{row+1}>0"], font=Font(bold=True, color="B3261E")))
s.cell(row=row + 2, column=1, value="A B kamerás és a 3/2-es lehetőség becslés; a rendező és az operatőr dönt róla.").font = Font(italic=True, color="666666")

# ---------- Fénybeállítások ----------
f = wb.create_sheet("Fénybeállítások")
f["A1"] = "Fénybeállítások (vilkó)"; f["A1"].font = Font(bold=True, size=14)
header(f, 3, ["Kód", "Fény", "Shotok", "Eredeti (perc)", "Most (perc)", "Eltérés"], [7, 26, 46, 14, 12, 10])
row = 4
for code, name in GROUPS:
    shots = [clean(r[3]).replace(" – Ultra High", "") for r in R if len(r) > 6 and r[6] == code and r[2] == "shoot"]
    f.cell(row=row, column=1, value=code); f.cell(row=row, column=2, value=name); f.cell(row=row, column=3, value=", ".join(shots))
    f.cell(row=row, column=4, value=f'=SUMIF({RNG("J")},A{row},{RNG("D")})')
    f.cell(row=row, column=5, value=f'=SUMIF({RNG("J")},A{row},{RNG("C")})')
    f.cell(row=row, column=6, value=f"=E{row}-D{row}").number_format = '+0;-0;0'
    for col in range(1, 7):
        f.cell(row=row, column=col).border = BOX
    row += 1
f.conditional_formatting.add(f"F4:F{row-1}", FormulaRule(formula=["F4>0"], font=Font(bold=True, color="B3261E")))
f.conditional_formatting.add(f"F4:F{row-1}", FormulaRule(formula=["F4<0"], font=Font(bold=True, color="1E7B3C")))

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
    ("Ha csúszunk (lásd még: Összesítés lap)", [
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
print("kész:", OUT)
