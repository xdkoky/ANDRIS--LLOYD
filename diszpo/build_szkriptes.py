# Egyszerű szkriptes napló (Excel / Google Táblázatok): EGY lap, a shotok forgatási sorrendben egymás alatt.
# A szkriptes csak két dolgot csinál: x a jó take mellé, és megjegyzés.
# Futtatás: python3 diszpo/build_szkriptes.py  ->  diszpo/Kapaszkodni_szkriptes_2026-11-12.xlsx
import json, os, re, subprocess, sys
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.formatting.rule import FormulaRule

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "Kapaszkodni_szkriptes_2026-11-12.xlsx")
TAKES = 6                                           # előre kiírt take-sorok shotonként (lehet újat beszúrni)
subprocess.run([sys.executable, os.path.join(HERE, "build_ad_app.py")], check=True, stdout=subprocess.DEVNULL)
D = json.loads(re.search(r"const D=(\{.*?\});\n", open(os.path.join(HERE, "ad_app.html"), encoding="utf-8").read(), re.S).group(1))
js = r"""const s=require('fs').readFileSync(process.argv[1],'utf8');
console.log(JSON.stringify(eval(s.match(/const SHOTS=(\[[\s\S]*?\n\]);/)[1])));"""
SL = {x["id"].replace("-", "/", 1): x for x in json.loads(subprocess.check_output(["node", "-e", js, os.path.join(HERE, "index.html")]))}
hm = lambda m: f"{int(m) // 60}:{int(m) % 60:02d}"
NAME = {"3/1": "Nagyi TV-t néz", "1/1": "Fiú eszik, dolly out", "1/3A": "Fiú közeli", "1/3B": "Fiú közeli + ajtó-reveal", "1/2": "Nagyi a konyhában, svenk",
        "2/2": "Tág", "MONTAGE": "Montázs", "2/3A": "Nagyi alulról, közeli", "2/3B": "Nagyi alulról, végső reakció", "2/5": "Inzert: Nagyi arca",
        "2/4": "Dolly: felülés, ölelés", "2/1C": "Fiú felülő medium", "2/1B": "Fiú fekvő medium", "2/1A": "Fiú közeli, profil",
        "2/1E": "Fiú extrém közeli, utolsó push", "2/6": "Bench bar POV", "3/2": "Nagyi végső reakció (este)"}
SHOTS = []
for b in D["blocks"]:
    if b["type"] == "shoot":
        for i in D["shots"][b["label"]]["ids"]:
            SHOTS.append((hm(b["s"]), "MONTÁZS" if i == "MONTAGE" else i, NAME.get(i, SL.get(i, {}).get("type", ""))))

thin = Side(style="thin", color="B5B5B5")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
F = lambda c: PatternFill("solid", fgColor=c)
DARK, INPUT, GREEN = F("303843"), F("FFF7D6"), F("BFE6C9")
MID = Alignment(vertical="center", horizontal="center")

wb = Workbook(); ws = wb.active; ws.title = "Szkript napló"
for col, w in zip("ABC", (9, 9, 110)):
    ws.column_dimensions[col].width = w
ws["A1"] = "KAPASZKODNI · szkriptes napló · 2026.11.12."; ws["A1"].font = Font(bold=True, size=15)
ws["A2"] = "A shotok forgatási sorrendben jönnek. Ha egy take jó, írj x-et a „Jó?” oszlopba. A megjegyzésbe bármit írhatsz. Ha elfogy a hely, szúrj be új sort."
ws["A2"].font = Font(color="444444")
ws["A3"] = "Kész:"; ws["A3"].font = Font(bold=True, size=12)
ws["C3"].font = Font(bold=True, size=12)
r = 5; heads = []
for time, sid, name in SHOTS:
    first, last = r + 1, r + TAKES
    ws.cell(row=r, column=1, value=time); ws.cell(row=r, column=2, value=f'=IF(COUNTA(B{first}:B{last})>0,"KÉSZ","")')
    ws.cell(row=r, column=3, value=f"{sid}  ·  {name}")
    for col in (1, 2, 3):
        c = ws.cell(row=r, column=col); c.fill = DARK; c.font = Font(bold=True, color="FFFFFF", size=12); c.border = BOX
        c.alignment = Alignment(vertical="center", horizontal="left" if col == 3 else "center")
    ws.row_dimensions[r].height = 24; heads.append(r)
    for k in range(TAKES):
        rr = first + k
        a = ws.cell(row=rr, column=1, value=k + 1); a.alignment = MID; a.font = Font(color="555555"); a.border = BOX
        b = ws.cell(row=rr, column=2); b.fill = INPUT; b.alignment = MID; b.font = Font(bold=True, size=12); b.border = BOX
        c = ws.cell(row=rr, column=3); c.fill = INPUT; c.alignment = Alignment(wrap_text=True, vertical="center"); c.border = BOX
        ws.row_dimensions[rr].height = 20
    kr = last + 1
    ws.cell(row=kr, column=1, value="kont.").alignment = MID; ws.cell(row=kr, column=1).font = Font(italic=True, color="555555"); ws.cell(row=kr, column=1).border = BOX
    ws.cell(row=kr, column=2).border = BOX
    c = ws.cell(row=kr, column=3); c.fill = INPUT; c.alignment = Alignment(wrap_text=True, vertical="center"); c.border = BOX
    ws.row_dimensions[kr].height = 20
    r = kr + 2
end = r
ws["B3"] = f'=COUNTIF(B5:B{end},"KÉSZ")&" / {len(SHOTS)}"'; ws["B3"].font = Font(bold=True, size=12); ws["B3"].alignment = MID
# fejléc a take-sorok fölé: a legelső shot előtt egyszer, és befagyasztva
ws["A4"] = "Take"; ws["B4"] = "Jó?"; ws["C4"] = "Megjegyzés (játék, technika, szövegeltérés) · a „kont.” sorba: jelmez, kellék, pozíció"
for col in "ABC":
    ws[f"{col}4"].font = Font(bold=True); ws[f"{col}4"].fill = F("E3E6EA"); ws[f"{col}4"].border = BOX
    ws[f"{col}4"].alignment = Alignment(horizontal="center" if col != "C" else "left", vertical="center")
ws.freeze_panes = "A5"
# zöld: a jó take sora, és a kész shot fejléce
ws.conditional_formatting.add(f"A5:C{end}", FormulaRule(formula=['AND($B5<>"",ISNUMBER($A5))'], fill=GREEN))
ws.conditional_formatting.add(f"A5:C{end}", FormulaRule(formula=['$B5="KÉSZ"'], fill=F("2E8B57"), font=Font(bold=True, color="FFFFFF")))
ws.page_setup.orientation = "portrait"; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0; ws.sheet_properties.pageSetUpPr.fitToPage = True
wb.calculation.fullCalcOnLoad = True
wb.save(OUT)
print("kész:", OUT, "·", len(SHOTS), "shot ·", TAKES, "take-sor shotonként")
