# Szkriptes munkafüzet (Excel / Google Táblázatok) a legfrissebb időrendből, shotlistből és a felsnittből.
# Futtatás: python3 diszpo/build_szkriptes.py  ->  diszpo/Kapaszkodni_szkriptes_2026-11-12.xlsx
# Lapok: Napló (take-enként, ide ír a szkriptes) · Shotok (állapot, magától számol) · Kontinuitás ·
#        Lefedettség (szkript soronként × shotok) · Napi jelentés (összesítő)
import json, os, re, subprocess, sys
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "Kapaszkodni_szkriptes_2026-11-12.xlsx")
subprocess.run([sys.executable, os.path.join(HERE, "build_ad_app.py")], check=True, stdout=subprocess.DEVNULL)
D = json.loads(re.search(r"const D=(\{.*?\});\n", open(os.path.join(HERE, "ad_app.html"), encoding="utf-8").read(), re.S).group(1))
js = r"""const s=require('fs').readFileSync(process.argv[1],'utf8');
console.log(JSON.stringify(eval(s.match(/const SHOTS=(\[[\s\S]*?\n\]);/)[1])));"""
SL = {x["id"].replace("-", "/", 1): x for x in json.loads(subprocess.check_output(["node", "-e", js, os.path.join(HERE, "index.html")]))}
hm = lambda m: f"{int(m) // 60}:{int(m) % 60:02d}"

# ---------- shotok forgatási sorrendben ----------
NAME = {"3/1": "Nagyi TV-t néz (CU)", "1/1": "Fiú eszik, dolly out", "1/3A": "Fiú közeli", "1/3B": "Fiú közeli + ajtó-reveal", "1/2": "Nagyi cowboy, svenk",
        "2/2": "Tág (master)", "MONTAGE": "Montázs", "2/3A": "Nagyi alulról, közeli", "2/3B": "Nagyi alulról, végső reakció", "2/5": "Inzert: Nagyi arca",
        "2/4": "Dolly: felülés, ölelés", "2/1C": "Fiú felülő medium", "2/1B": "Fiú fekvő medium", "2/1A": "Fiú MCU, profil",
        "2/1E": "Fiú ECU, utolsó push", "2/6": "Bench bar POV", "3/2": "Nagyi végső reakció (este)"}
MONT = {"type": "Montázs", "desc": "Bicepsz a kézisúllyal, nyújtás a matracon, Nagyi segít.", "lens": "", "angle": "", "move": "Svenk", "sup": "Sticks", "talent": "Soma, Nagyi"}
SHOTS = []
for b in D["blocks"]:
    if b["type"] != "shoot":
        continue
    sh = D["shots"][b["label"]]
    for i in sh["ids"]:
        s = SL.get(i, MONT); short = "MONT" if i == "MONTAGE" else i
        SHOTS.append({"id": short, "raw": i, "key": f"{short} · {NAME.get(i, s['type'])}", "scene": "M" if i == "MONTAGE" else i.split("/")[0],
                      "time": f'{hm(b["s"])}–{hm(b["s"] + b["d"])}', "desc": f'{s["type"]}: {s["desc"]}',
                      "cam": " · ".join(x for x in (s["lens"], s["angle"], s["move"], s["sup"]) if x), "talent": s["talent"],
                      "v5": " — ".join(x for x in (b["setup"], b["note"]) if x), "props": "\n".join("• " + p for p in sh["props"]), "story": sh["story"]})
NS = len(SHOTS)

# ---------- szkript sorai (szereplő + instrukció + dialógus egy sorban) ----------
ROWS, cur = [], None
for l in D["script"]:
    t, x = l["t"], l["x"]
    ids = [("MONT" if i == "MONTAGE" else i) for c in l["cov"] for i in c["ids"]]
    if t == "sh":
        ROWS.append({"head": True, "text": x.upper()}); cur = None
    elif t == "end":
        ROWS.append({"head": True, "text": x})
    elif t == "ch":
        cur = {"who": x.upper(), "par": "", "ids": set(ids)}
    elif t == "pa" and cur:
        cur["par"] += (" " if cur["par"] else "") + x; cur["ids"] |= set(ids)
    elif t == "dl" and cur:
        if cur.get("row"):                       # ugyanannak a megszólalásnak a folytatása (instrukció után)
            cur["row"]["text"] += f' {cur["par"]} {x}'.replace("  ", " "); cur["row"]["ids"] |= cur["ids"] | set(ids); cur["par"] = ""
        else:
            row = {"text": f'{cur["who"]}{(" " + cur["par"]) if cur["par"] else ""}: {x}', "ids": cur["ids"] | set(ids), "dl": True}
            ROWS.append(row); cur["row"] = row; cur["par"] = ""
    else:
        ROWS.append({"text": x, "ids": set(ids)}); cur = None
n = 0
for r in ROWS:
    if not r.get("head"):
        n += 1; r["n"] = n
NL = n
def ranges(nums):
    nums = sorted(nums); out = []
    for v in nums:
        if out and v == out[-1][1] + 1:
            out[-1][1] = v
        else:
            out.append([v, v])
    return ", ".join(str(a) if a == b else f"{a}–{b}" for a, b in out)
for s in SHOTS:
    mine = [r["n"] for r in ROWS if not r.get("head") and s["id"] in r["ids"]]
    s["lines"] = ranges(mine); s["first"] = min(mine) if mine else 999
for k, s in enumerate(sorted(SHOTS, key=lambda s: s["first"]), 1):
    s["story_no"] = k

# ---------- stílusok ----------
thin = Side(style="thin", color="B0B0B0")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
F = lambda c: PatternFill("solid", fgColor=c)
HEAD, INPUT, AUTO, GREEN, AMBER, GREY = F("303843"), F("FFF7D6"), F("EEF1F4"), F("CDEBD5"), F("FFE9B8"), F("E3E6EA")
WHITE_B = Font(bold=True, color="FFFFFF")
WRAP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(horizontal="center", vertical="top", wrap_text=True)
def head(ws, row, titles, widths):
    for i, (t, w) in enumerate(zip(titles, widths), 1):
        c = ws.cell(row=row, column=i, value=t); c.font = WHITE_B; c.fill = HEAD; c.border = BOX
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        ws.column_dimensions[L(i)].width = w
    ws.row_dimensions[row].height = 32
def title(ws, text, sub):
    ws["A1"] = text; ws["A1"].font = Font(bold=True, size=15)
    ws["A2"] = sub; ws["A2"].font = Font(italic=True, color="666666")

wb = Workbook()
N1, N2 = 6, 205                                   # a Napló adatsorai
R = lambda col: f"Napló!${col}${N1}:${col}${N2}"
S1, S2 = 5, 5 + NS - 1                            # a Shotok adatsorai

# ====================== 1) NAPLÓ ======================
ws = wb.active; ws.title = "Napló"
title(ws, "KAPASZKODNI · szkriptes napló", "2026.11.12. · minden take egy sor. Csak a sárga cellákba írj, a szürkék maguktól töltődnek.")
ws["A3"] = "Így megy: 1) Shot: válaszd a legördülőből → a take száma magától jön.  2) Idő: óra:perc, pl. 10:42.  3) Hossz: perc:mp, pl. 0:42.  4) Minősítés: a végleges take legyen KOPÍR (több is lehet)."
ws["A3"].font = Font(color="333333"); ws["A3"].alignment = Alignment(wrap_text=False)
head(ws, 5, ["Idő", "Shot (válassz)", "Take", "Minősítés", "Hossz (perc:mp)", "Kártya / klip", "Hang", "Optika, blende (tényleges)",
             "Játék · szöveg", "Technika (fókusz, mozgás, hang)", "Kontinuitás", "Vágónak", "Jel.", "Terv: kamera", "Terv: mit veszünk"],
     [8, 30, 6, 13, 9, 13, 10, 16, 34, 30, 34, 28, 5, 26, 44])
for r in range(N1, N2 + 1):
    ws.cell(row=r, column=3, value=f'=IF($B{r}="",0,COUNTIF($B${N1}:$B{r},$B{r}))').number_format = "0;;"
    ws.cell(row=r, column=13, value=f'=IF($B{r}="","",IFERROR(VLOOKUP($B{r},Shotok!$A${S1}:$H${S2},2,FALSE),""))')
    ws.cell(row=r, column=14, value=f'=IF($B{r}="","",IFERROR(VLOOKUP($B{r},Shotok!$A${S1}:$H${S2},6,FALSE),""))')
    ws.cell(row=r, column=15, value=f'=IF($B{r}="","",IFERROR(VLOOKUP($B{r},Shotok!$A${S1}:$H${S2},5,FALSE),""))')
    for col in range(1, 16):
        c = ws.cell(row=r, column=col); c.border = BOX; c.alignment = WRAP
        c.fill = AUTO if col in (3, 13, 14, 15) else INPUT
        if col in (3, 13, 14, 15):
            c.font = Font(color="555555", bold=(col == 3))
        if col in (14, 15):
            c.alignment = Alignment(wrap_text=False, vertical="top")   # ne nyújtsa meg a sort
    ws.cell(row=r, column=1).number_format = "h:mm"; ws.cell(row=r, column=5).number_format = "[h]:mm"
    ws.cell(row=r, column=2).number_format = "@"
    for col in (1, 3, 4, 5):
        ws.cell(row=r, column=col).alignment = CENTER
dv = DataValidation(type="list", formula1=f"=Shotok!$A${S1}:$A${S2}", allow_blank=True); ws.add_data_validation(dv); dv.add(f"B{N1}:B{N2}")
RATE = ["KOPÍR", "jó", "tartalék", "rossz", "megszakadt"]
dv2 = DataValidation(type="list", formula1='"' + ",".join(RATE) + '"', allow_blank=True); ws.add_data_validation(dv2); dv2.add(f"D{N1}:D{N2}")
ws.conditional_formatting.add(f"A{N1}:L{N2}", FormulaRule(formula=[f'$D{N1}="KOPÍR"'], fill=GREEN, font=Font(bold=True)))
ws.conditional_formatting.add(f"A{N1}:L{N2}", FormulaRule(formula=[f'OR($D{N1}="rossz",$D{N1}="megszakadt")'], fill=GREY, font=Font(color="808080")))
ws.freeze_panes = f"C{N1}"

# ====================== 2) SHOTOK ======================
sh = wb.create_sheet("Shotok")
title(sh, "Shotok · állapot", "Forgatási sorrendben. Az utolsó oszlopok a Naplóból számolódnak: nem kell ide írni, csak a Megjegyzésbe.")
head(sh, 4, ["Shot", "Jel.", "Terv idő", "Sztori-sorrend", "Mit veszünk", "Kamera (terv)", "Szereplő", "Szkript-sorok (Lefedettség #)",
             "Megjegyzés az időrendből", "Take", "Kopír", "Utolsó kopír take", "Kopír hossz", "Állapot", "Szkriptes megjegyzés"],
     [30, 5, 12, 8, 46, 28, 12, 22, 40, 6, 6, 9, 9, 11, 36])
for k, s in enumerate(SHOTS):
    r = S1 + k
    vals = [s["key"], s["scene"], s["time"], s["story_no"], s["desc"], s["cam"], s["talent"], s["lines"], s["v5"],
            f'=COUNTIF({R("B")},$A{r})', f'=COUNTIFS({R("B")},$A{r},{R("D")},"KOPÍR")',
            f'=IF(K{r}=0,"",SUMPRODUCT(MAX(({R("B")}=$A{r})*({R("D")}="KOPÍR")*{R("C")})))',
            f'=SUMIFS({R("E")},{R("B")},$A{r},{R("D")},"KOPÍR")', f'=IF(K{r}>0,"KÉSZ",IF(J{r}>0,"forog","hátravan"))', None]
    for col, v in enumerate(vals, 1):
        c = sh.cell(row=r, column=col, value=v); c.border = BOX; c.alignment = WRAP
        if col in (2, 3, 4, 10, 11, 12, 13, 14):
            c.alignment = CENTER
        if 10 <= col <= 14:
            c.fill = AUTO
    sh.cell(row=r, column=1).font = Font(bold=True); sh.cell(row=r, column=1).number_format = "@"
    sh.cell(row=r, column=13).number_format = "[h]:mm"; sh.cell(row=r, column=15).fill = INPUT
sh.conditional_formatting.add(f"N{S1}:N{S2}", FormulaRule(formula=[f'$N{S1}="KÉSZ"'], fill=GREEN, font=Font(bold=True)))
sh.conditional_formatting.add(f"N{S1}:N{S2}", FormulaRule(formula=[f'$N{S1}="forog"'], fill=AMBER, font=Font(bold=True)))
sh.conditional_formatting.add(f"A{S1}:A{S2}", FormulaRule(formula=[f'$N{S1}="KÉSZ"'], fill=GREEN))
sh.freeze_panes = f"B{S1}"

# ====================== 3) KONTINUITÁS ======================
ko = wb.create_sheet("Kontinuitás")
title(ko, "Kontinuitás", "A sorok forgatási sorrendben vannak; a Sztori-sorrend oszlop mutatja, hol lesz a filmben. Szomszédos sztori-számoknál kell egyeznie mindennek.")
FIX = ["1. jel.: Nagyi a NEGYEDIK tál ételt teszi le; az asztalon sok üres tányér.", "Nagyi cetlije: „Harminc, ötvenkettő… hetvenhárom”. Az apa „30-al többet” vitt be.",
       "Nagyi a mosogatószivaccsal és egy tányérral fordul hátra.", "2. jel., 1. sorozat: Nagyi kilencig számol, de azt mondja: HÉT.",
       "A plusz súlyt KÉT OLDALRÓL, majdnem szinkronban rakják fel.", "2. sorozat: a HATODIKNÁL bukik el → ölelés → „Újra!”.",
       "3. sorozat: a rúd RÁESIK; Nagyi „Tomi”-nak szólítja.", "3. jel.: sötét szoba, csak a TV fénye Nagyi arcán.",
       "Terv szerinti jelmez – Fiú: fehér trikó, fekete rövidnadrág · Nagyi: otthonka/kardigán zsebbel, papucs, láncos szemüveg."]
ko["A4"] = "Fix pontok a szkriptből"; ko["A4"].font = Font(bold=True, size=12)
for i, t in enumerate(FIX):
    c = ko.cell(row=5 + i, column=1, value="• " + t); ko.merge_cells(start_row=5 + i, start_column=1, end_row=5 + i, end_column=8)
KH = 5 + len(FIX) + 1
head(ko, KH, ["Shot", "Sztori-sorrend", "Mit látunk", "Figyelni (kellék, biztonság)", "Jelmez – Fiú", "Jelmez – Nagyi", "Haj · smink · izzadás",
              "Kellékek állása (tányér, étel, súly a rúdon)", "Pozíció · kéz · tekintet", "Fotó #", "Megjegyzés"],
     [30, 8, 34, 38, 22, 22, 24, 30, 30, 9, 34])
for k, s in enumerate(SHOTS):
    r = KH + 1 + k
    for col, v in enumerate([s["key"], s["story_no"], s["story"], s["props"]] + [None] * 7, 1):
        c = ko.cell(row=r, column=col, value=v); c.border = BOX; c.alignment = WRAP
        c.fill = INPUT if col >= 5 else AUTO
    ko.cell(row=r, column=1).font = Font(bold=True); ko.cell(row=r, column=2).alignment = CENTER
    ko.row_dimensions[r].height = max(48, 15 * (s["props"].count("\n") + 1))
ko.freeze_panes = ko.cell(row=KH + 1, column=2)

# ====================== 4) LEFEDETTSÉG ======================
le = wb.create_sheet("Lefedettség")
title(le, "Lefedettség · szkript soronként", "■ = a felsnitt szerint ez a shot veszi a sort. Zöld, ha a shotnak már van KOPÍR take-je. A ✓ oszlop mutatja, hogy a sor megvan-e már.")
C0 = 6                                             # az első shot-oszlop
head(le, 5, ["#", "Szöveg", "Megvan", "Szövegeltérés · megjegyzés", ""] + [s["id"] for s in SHOTS], [5, 70, 8, 34, 2] + [5.2] * NS)
le.cell(row=6, column=2, value="A shot állapota →").alignment = Alignment(horizontal="right"); le.cell(row=6, column=2).font = Font(italic=True, color="666666")
for k in range(NS):
    c = le.cell(row=5, column=C0 + k); c.alignment = Alignment(text_rotation=90, horizontal="center", vertical="bottom"); c.number_format = "@"
    st = le.cell(row=6, column=C0 + k, value=f"=Shotok!$N${S1 + k}"); st.alignment = Alignment(text_rotation=90, horizontal="center"); st.font = Font(size=8, color="555555")
le.row_dimensions[5].height = 46; le.row_dimensions[6].height = 52
LR1 = 7
a, z = L(C0), L(C0 + NS - 1)
for k, row in enumerate(ROWS):
    r = LR1 + k
    if row.get("head"):
        c = le.cell(row=r, column=2, value=row["text"]); c.font = Font(bold=True)
        for col in range(1, C0 + NS):
            le.cell(row=r, column=col).fill = GREY
        continue
    le.cell(row=r, column=1, value=row["n"]).alignment = CENTER
    t = le.cell(row=r, column=2, value=row["text"]); t.alignment = WRAP
    if row.get("dl"):
        t.font = Font(color="1A3E72")
    m = le.cell(row=r, column=3, value=f'=IF(COUNTIF({a}{r}:{z}{r},"■")=0,"–",IF(SUMPRODUCT(({a}{r}:{z}{r}="■")*({a}$6:{z}$6="KÉSZ"))>0,"✓",""))')
    m.alignment = CENTER; m.fill = AUTO; m.font = Font(bold=True, color="1E7B3C")
    n4 = le.cell(row=r, column=4); n4.fill = INPUT; n4.alignment = WRAP
    for j, s in enumerate(SHOTS):
        c = le.cell(row=r, column=C0 + j, value="■" if s["id"] in row["ids"] else None); c.alignment = Alignment(horizontal="center", vertical="center")
    for col in range(1, C0 + NS):
        le.cell(row=r, column=col).border = BOX
LR2 = LR1 + len(ROWS) - 1
le.conditional_formatting.add(f"{a}{LR1}:{z}{LR2}", FormulaRule(formula=[f'AND({a}{LR1}="■",{a}$6="KÉSZ")'], fill=F("5DBB7A"), font=Font(color="1E7B3C")))
le.conditional_formatting.add(f"{a}{LR1}:{z}{LR2}", FormulaRule(formula=[f'{a}{LR1}="■"'], fill=AMBER, font=Font(color="A06A00")))
le.conditional_formatting.add(f"A{LR1}:B{LR2}", FormulaRule(formula=[f'$C{LR1}="✓"'], fill=GREEN))
le["B3"] = f'="Megvan: "&COUNTIF($C${LR1}:$C${LR2},"✓")&" / {NL} sor"'; le["B3"].font = Font(bold=True)
le.freeze_panes = le.cell(row=LR1, column=3)

# ====================== 5) NAPI JELENTÉS ======================
nj = wb.create_sheet("Napi jelentés")
title(nj, "Napi jelentés · szkript", "A szürke cellák a Naplóból számolódnak, a sárgákat te töltöd ki.")
for col, w in zip("ABCDEF", (30, 18, 16, 16, 16, 30)):
    nj.column_dimensions[col].width = w
def kv(r, label, value=None, fmt=None, inp=False):
    a_ = nj.cell(row=r, column=1, value=label); a_.font = Font(bold=True); a_.border = BOX
    b_ = nj.cell(row=r, column=2, value=value); b_.border = BOX; b_.fill = INPUT if inp else AUTO; b_.alignment = Alignment(horizontal="left")
    if fmt:
        b_.number_format = fmt
kv(4, "Produkció", "Kapaszkodni"); kv(5, "Dátum", "2026.11.12. (csütörtök)"); kv(6, "Rendező", "Ágoston András"); kv(7, "Operatőr", "Détári Szabolcs")
kv(8, "Szkriptes", next((c["name"] for c in D["crew"] if c["pos"].lower().startswith("script")), ""))
kv(10, "Crew call (terv)", hm(D["plan"]["call"])); kv(11, "Crew call (tényleges)", None, "h:mm", True)
kv(12, "Első csapó", f'=IF(COUNT({R("A")})=0,"",MIN({R("A")}))', "h:mm"); kv(13, "Ebéd (tól–ig)", None, None, True)
kv(14, "Utolsó take", f'=IF(COUNT({R("A")})=0,"",MAX({R("A")}))', "h:mm"); kv(15, "Wrap (tényleges)", None, "h:mm", True)
kv(17, "Kész shotok", f'=COUNTIF(Shotok!$N${S1}:$N${S2},"KÉSZ")&" / {NS}"'); kv(18, "Hátralévő shotok", f'={NS}-COUNTIF(Shotok!$N${S1}:$N${S2},"KÉSZ")')
kv(19, "Összes take", f'=COUNTA({R("B")})'); kv(20, "Kopír take-ek", f'=COUNTIF({R("D")},"KOPÍR")')
kv(21, "Felvett anyag összesen (perc:mp)", f'=SUM({R("E")})', "[h]:mm"); kv(22, "Kopír take-ek hossza (perc:mp)", f'=SUMIF({R("D")},"KOPÍR",{R("E")})', "[h]:mm")
nj["C22"] = "Ez nem a film hossza: az átfedő shotok többször számítanak."; nj["C22"].font = Font(italic=True, color="666666")
r0 = 24
for i, t in enumerate(["Jelenet", "Shot összesen", "Kész", "Take", "Kopír hossz", "Mért filmhossz (master alapján)"], 1):
    c = nj.cell(row=r0, column=i, value=t); c.font = WHITE_B; c.fill = HEAD; c.border = BOX; c.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
nj.row_dimensions[r0].height = 32
SC = f"Shotok!$B${S1}:$B${S2}"
for i, (code, name) in enumerate([("1", "1. Konyha"), ("2", "2. Nappali"), ("3", "3. Nappali, este"), ("M", "Montázs")], 1):
    r = r0 + i
    vals = [name, f'=COUNTIF({SC},"{code}")', f'=COUNTIFS({SC},"{code}",Shotok!$N${S1}:$N${S2},"KÉSZ")', f'=SUMIF({SC},"{code}",Shotok!$J${S1}:$J${S2})',
            f'=SUMIF({SC},"{code}",Shotok!$M${S1}:$M${S2})', None]
    for col, v in enumerate(vals, 1):
        c = nj.cell(row=r, column=col, value=v); c.border = BOX; c.fill = INPUT if col == 6 else AUTO; c.alignment = Alignment(horizontal="left" if col == 1 else "center")
    nj.cell(row=r, column=5).number_format = "[h]:mm"; nj.cell(row=r, column=6).number_format = "[h]:mm"
def block(r, heading, cols, rows):
    nj.cell(row=r, column=1, value=heading).font = Font(bold=True, size=12)
    for i, t in enumerate(cols, 1):
        c = nj.cell(row=r + 1, column=i, value=t); c.font = WHITE_B; c.fill = HEAD; c.border = BOX
    for k in range(rows):
        for i in range(1, len(cols) + 1):
            c = nj.cell(row=r + 2 + k, column=i); c.fill = INPUT; c.border = BOX; c.alignment = WRAP
    return r + 2 + rows + 1
nx = block(r0 + 6, "Vadhangok", ["Mi", "Idő", "Hangfájl", "Megjegyzés"], 6)
nx = block(nx, "Pótfelvétel · hiány", ["Shot / mi hiányzik", "Miért", "Mikor pótolható", "Megjegyzés"], 5)
nj.cell(row=nx, column=1, value="Megjegyzés a rendezőnek · vágónak").font = Font(bold=True, size=12)
nj.merge_cells(start_row=nx + 1, start_column=1, end_row=nx + 6, end_column=6)
c = nj.cell(row=nx + 1, column=1); c.fill = INPUT; c.alignment = WRAP

for w in (ws, sh, ko, le, nj):
    w.sheet_view.zoomScale = 100
    w.page_setup.orientation = "landscape"; w.page_setup.fitToWidth = 1; w.page_setup.fitToHeight = 0; w.sheet_properties.pageSetUpPr.fitToPage = True
wb.calculation.fullCalcOnLoad = True
wb.save(OUT)
print("kész:", OUT, "·", NS, "shot ·", NL, "szkript-sor")
