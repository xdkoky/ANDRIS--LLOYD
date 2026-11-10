# Forgatókönyv forgatási sorrendben, a szerkesztett ütemtervhez (V4) igazítva.
# Futtatás: python3 diszpo/build_forgatokonyv.py
#  -> diszpo/Kapaszkodni_utemterv_V5.xlsx  (a V4 + új „Forgatókönyv” lap; a sorok húzhatók,
#                                            az időpontok az Időrend lapról jönnek)
#  -> diszpo/forgatokonyv_sorrend.html     (ebből a build_pdf.js készít PDF-et)
# A shotokhoz rendelt szövegrészek a rendező felsnittje (anyagok/…_FELSNITT.pdf) alapján készültek.
import html, io, json, os, re, shutil, subprocess, tempfile
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.utils.units import pixels_to_EMU
from PIL import Image as PILImage

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "Kapaszkodni_utemterv_V4.xlsx")
OUT = os.path.join(HERE, "Kapaszkodni_utemterv_V5.xlsx")
HTML_OUT = os.path.join(HERE, "forgatokonyv_sorrend.html")

# ---- a szkript elemei az index.html-ből ----
js = r"""const s=require('fs').readFileSync(process.argv[1],'utf8');
console.log(JSON.stringify(eval(s.match(/const SCRIPT=(\[[\s\S]*?\n\]);/)[1])));"""
SCRIPT = json.loads(subprocess.check_output(["node", "-e", js, os.path.join(HERE, "index.html")]))

# A felsnitt kézírásos javításai
FIX = {
    76: ("ac", "Fiú lefekszik a padra. Nagyi előtte áll.", "felsnitt: „leül” → „fekszik”"),
    89: ("ac", "Fiú lock in, elkezdi nyomni a súlyt, hallhatóan küszködik.", "felsnitt: „hátradől” → „lock in”"),
}
def el(i):
    t, txt = SCRIPT[i][0], SCRIPT[i][1]
    if i in FIX:
        t, txt, _ = FIX[i]
    return t, txt

# Szakaszok: (első index, utolsó index, felsnitt-jelölés, shotok)
SEG = [
    (1, 7, "a shotlist alapján (nincs felsnitt-jelölés)", ["1/1"]),
    (1, 48, "a shotlist alapján: teljes jelenet a Fiú oldaláról", ["1/3A"]),
    (1, 48, "a shotlist alapján: teljes jelenet Nagyi oldaláról", ["1/2"]),
    (49, 51, "Dolly átmenet", ["1/3B"]),
    (52, 56, "Medium fiú / Szoba tág / Nagyi CU", ["2/1B", "2/2", "2/3A"]),
    (57, 65, "Medium fiú felül", ["2/1C"]),
    (66, 67, "CU", ["2/3A"]),
    (68, 72, "tág", ["2/2"]),
    (73, 76, "Medium fekszik", ["2/1B"]),
    (77, 79, "CU", ["2/3A"]),
    (80, 81, "Medium fiú fekszik", ["2/1B"]),
    (82, 85, "CU", ["2/3A"]),
    (86, 91, "Medium fekszik", ["2/1B"]),
    (92, 94, "CU", ["2/3A"]),
    (95, 95, "Fiú CU", ["2/1A"]),
    (96, 97, "CU", ["2/3A"]),
    (98, 99, "Medium DOLLY", ["2/4"]),
    (100, 101, "insert Nagyi CU", ["2/5"]),
    (102, 102, "Medium DOLLY – dolly vissza", ["2/4"]),
    (103, 103, "CU", ["2/3A"]),
    (104, 104, "dolly vissza (folyt.)", ["2/4"]),
    (105, 105, "CU", ["2/3A"]),
    (106, 106, "Fiú CU", ["2/1E"]),
    (107, 108, "CU", ["2/3A"]),
    (109, 109, "Handheld bar POV", ["2/6"]),
    (110, 125, "Nagyi CU", ["2/3B"]),
    (128, 129, "a shotlist alapján", ["3/1"]),
    (128, 129, "a shotlist alapján: végső reakció", ["3/2"]),
    (131, 132, "kiegészítő jelenet (nincs a shotlistben)", ["MONTAGE"]),
]
SCENE_OF = lambda i: 1 if i < 50 else 2 if i < 127 else 3 if i < 130 else "M"
GUESS = {"2/1A": "A felsnitt „Fiú CU”-ja; a shotlist szerint ez a 2/1A (MCU) – ellenőrizni.",
         "2/1E": "A felsnitt „Fiú CU”-ja; a shotlist szerint ez a 2/1E (utolsó push).",
         "2/2": "A felsnitt két „tág” szakaszt jelöl; a shotlist szerint ez a master."}

def shots_of_label(label):
    """A V4 Jel./shot cellájából a shotlist-azonosítók (pl. „1/3A + 1/3B” → [1/3A, 1/3B])."""
    if label.strip().upper().startswith("MONTAGE"):
        return ["MONTAGE"]
    return re.findall(r"\d/\d[A-E]?", label)

def segments_for(ids):
    return [s for s in SEG if set(s[3]) & set(ids)]

def plain_lines(a, b):
    out, pend = [], None
    for i in range(a, b + 1):
        t, txt = el(i)
        if t == "sh":
            out.append(txt.upper())
        elif t == "ch":
            pend = txt.upper()
        elif t == "pa":
            pend = (pend or "") + " " + txt
        elif t == "dl":
            out.append(f"{pend}: {txt}" if pend else txt); pend = None
        elif t == "ac":
            out.append(txt)
    return out

# ---- V4 időpontok (LibreOffice-szal újraszámolva, csak az olvasáshoz) ----
tmp = tempfile.mkdtemp()
subprocess.run(["soffice", "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", tmp, SRC],
               check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
vals = load_workbook(os.path.join(tmp, os.path.basename(SRC)), data_only=True)["Időrend"]
ROWS = []
for r in range(9, vals.max_row + 1):
    if vals.cell(r, 6).value == "Forgatás":
        ROWS.append(dict(label=str(vals.cell(r, 7).value).strip(), start=vals.cell(r, 1).value, end=vals.cell(r, 2).value,
                         setup=vals.cell(r, 9).value or "", note=vals.cell(r, 10).value or ""))
shutil.rmtree(tmp)

# ---- V5: a V4 + Forgatókönyv lap ----
wb = load_workbook(SRC)
if "Forgatókönyv" in wb.sheetnames:
    del wb["Forgatókönyv"]
ws = wb.create_sheet("Forgatókönyv", 1)
thin = Side(style="thin", color="999999")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
WRAP = Alignment(wrap_text=True, vertical="top")
ws["A1"] = "KAPASZKODNI – forgatókönyv forgatási sorrendben"; ws["A1"].font = Font(bold=True, size=16)
ws["A2"] = ("Az időpontok és a beállítás az Időrend lapról jönnek. A sorok szabadon húzhatók. "
            "Ha az Időrendben változik a sorrend: jelöld ki az A5-től lefelé a sorokat, majd Adatok → Tartomány rendezése → A oszlop (A→Z).")
ws["A3"] = "A szövegrészek a rendező felsnittje alapján; a dőlt betűs megjegyzés jelzi, ahol értelmezni kellett."
for c in ("A2", "A3"):
    ws[c].font = Font(italic=True, color="666666")
HR = 4
heads = ["Kezdés", "Vég", "Shot", "Storyboard", "Beállítás", "Felsnitt", "Szkript-részlet", "Megjegyzés", "Take / OK", "Szkriptes jegyzet"]
widths = [8, 8, 13, 17, 24, 22, 70, 26, 10, 26]
for i, (h, w) in enumerate(zip(heads, widths), 1):
    c = ws.cell(row=HR, column=i, value=h)
    c.font = Font(bold=True); c.fill = PatternFill("solid", fgColor="C4C4C4"); c.border = BOX; c.alignment = WRAP
    ws.column_dimensions[chr(64 + i)].width = w
ws.row_dimensions[HR].height = 18
for rr in (1, 2, 3):
    ws.row_dimensions[rr].height = 22 if rr == 1 else 15
LOOK = lambda col: (f'IFERROR(INDEX(INDIRECT("\'Időrend\'!{col}9:{col}200"),'
                    f'MATCH($C{{r}},INDIRECT("\'Időrend\'!G9:G200"),0)),"")')
SB_DIR = os.path.join(HERE, "storyboard")
html_blocks = []
for k, row in enumerate(ROWS):
    r = HR + 1 + k
    ids = shots_of_label(row["label"])
    segs = segments_for(ids)
    text_parts, labels = [], []
    for a, b, lab, sh in segs:
        lab_txt = lab + ("" if len(ids) == 1 else f"  ({', '.join(x for x in sh if x in ids)})")
        labels.append(lab_txt)
        text_parts.append(f"[{lab_txt}]\n" + "\n".join(plain_lines(a, b)))
    notes = [GUESS[i] for i in ids if i in GUESS]
    notes += [FIX[i][2] for a, b, _, _ in segs for i in FIX if a <= i <= b]
    ws.cell(row=r, column=1, value="=" + LOOK("A").format(r=r)).number_format = "h:mm"
    ws.cell(row=r, column=2, value="=" + LOOK("B").format(r=r)).number_format = "h:mm"
    ws.cell(row=r, column=3, value=row["label"]).font = Font(bold=True)
    ws.cell(row=r, column=5, value="=" + LOOK("I").format(r=r))
    ws.cell(row=r, column=6, value="\n".join(dict.fromkeys(labels)))
    ws.cell(row=r, column=7, value="\n\n".join(text_parts))
    ws.cell(row=r, column=8, value="\n".join(dict.fromkeys(notes))).font = Font(italic=True, color="555555")
    for col in range(1, 11):
        ws.cell(row=r, column=col).border = BOX
        ws.cell(row=r, column=col).alignment = WRAP
    nlines = sum(1 + len(t) // 85 for t in "\n\n".join(text_parts).split("\n"))
    ws.row_dimensions[r].height = min(409, max(70, 13.5 * nlines + 6))
    files = [os.path.join(SB_DIR, i.replace("/", "-") + ".jpg") for i in ids]
    files = [f for f in files if os.path.exists(f)]
    if files:
        W, H = 360, 270
        canvas = PILImage.new("RGB", (W, H), "white"); w = W // len(files)
        for j, fp in enumerate(files):
            im = PILImage.open(fp).convert("RGB"); im.thumbnail((w - 6, H - 6))
            canvas.paste(im, (j * w + (w - im.width) // 2, (H - im.height) // 2))
        buf = io.BytesIO(); canvas.save(buf, "PNG"); buf.seek(0)
        img = XLImage(buf); img.width, img.height = 114, 84
        img.anchor = OneCellAnchor(_from=AnchorMarker(col=3, colOff=pixels_to_EMU(4), row=r - 1, rowOff=pixels_to_EMU(3)),
                                   ext=XDRPositiveSize2D(pixels_to_EMU(114), pixels_to_EMU(84)))
        ws.add_image(img)
    # HTML-blokk a PDF-hez
    t = lambda x: x.strftime("%H:%M") if x else ""
    body = []
    for a, b, lab, sh in segs:
        lab_txt = lab + ("" if len(ids) == 1 else f" ({', '.join(x for x in sh if x in ids)})")
        body.append(f'<div class="seg"><div class="lab">{html.escape(lab_txt)}</div>')
        for i in range(a, b + 1):
            tt, txt = el(i)
            cls = {"ac": "ac", "ch": "ch", "pa": "pa", "dl": "dl", "sh": "shd"}.get(tt)
            if cls:
                body.append(f'<div class="{cls}">{html.escape(txt)}</div>')
        body.append("</div>")
    sbimg = "".join(f'<img src="storyboard/{os.path.basename(f)}" alt="">' for f in files)
    html_blocks.append(
        f'<section class="shot"><header><div class="time">{t(row["start"])}–{t(row["end"])}</div>'
        f'<div class="id">{html.escape(row["label"])}</div><div class="setup">{html.escape(str(row["setup"]))}'
        f'{"<br><i>" + html.escape(str(row["note"])) + "</i>" if row["note"] else ""}</div>'
        f'<div class="sb">{sbimg}</div></header>'
        + (f'<div class="notes">{"<br>".join(html.escape(n) for n in dict.fromkeys(notes))}</div>' if notes else "")
        + "".join(body) + "</section>")
ws.freeze_panes = f"A{HR + 1}"
ws.page_setup.orientation = "landscape"; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.print_title_rows = f"{HR}:{HR}"
wb.calculation.fullCalcOnLoad = True
wb.save(OUT)
print("kész:", OUT)

with open(HTML_OUT, "w", encoding="utf-8") as fh:
    fh.write("""<!doctype html><html lang="hu"><head><meta charset="utf-8"><title>Kapaszkodni – forgatókönyv forgatási sorrendben</title>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Courier+Prime:ital,wght@0,400;0,700;1,400&family=IBM+Plex+Sans+Condensed:wght@400;600;700&display=swap">
<style>
@page{size:A4;margin:14mm 14mm 16mm}
*{box-sizing:border-box}html,body{margin:0;background:#fff;color:#000}
body{font-family:"Courier Prime","Courier New",monospace;font-size:10.5pt;line-height:1.25}
h1{font-size:15pt;margin:0 0 1mm}.sub{font-family:"IBM Plex Sans Condensed",sans-serif;font-size:8.5pt;color:#555;margin-bottom:5mm}
.shot{margin:0 0 6mm;break-inside:auto}
.shot header{display:grid;grid-template-columns:24mm 26mm 1fr auto;gap:3mm;align-items:center;background:#e2e2e2;border:1.5px solid #000;padding:2mm 3mm;break-after:avoid;break-inside:avoid}
.time{font-weight:700}.id{font-weight:700;font-size:12.5pt}
.setup{font-family:"IBM Plex Sans Condensed",sans-serif;font-size:8.5pt}
.sb img{height:17mm;border:1px solid #999;background:#fff;margin-left:1mm}
.notes{font-family:"IBM Plex Sans Condensed",sans-serif;font-size:8pt;font-style:italic;color:#444;margin:1.5mm 0 0}
.seg{margin:3mm 0 0 0;padding-left:3mm;border-left:2px solid #000}
.lab{break-after:avoid;font-family:"IBM Plex Sans Condensed",sans-serif;font-size:8pt;font-weight:700;text-transform:uppercase;letter-spacing:.04em;margin-bottom:1mm}
.ac{margin:0 0 2.5mm}.ch{margin:2mm 0 0 34%;text-transform:uppercase;break-after:avoid}
.pa{margin:0 0 0 27%;font-style:italic;break-after:avoid}.dl{margin:0 12% 2.5mm 20%}
.shd{font-weight:700;text-transform:uppercase;margin:0 0 2mm}
</style></head><body>
<h1>KAPASZKODNI · forgatókönyv forgatási sorrendben</h1>
<div class="sub">2026.11.12. · a V4 ütemterv sorrendje és időpontjai · a szövegrészek a rendező felsnittje alapján</div>
""" + "\n".join(html_blocks) + "\n</body></html>")
print("kész:", HTML_OUT)
