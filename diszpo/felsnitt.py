# Közös adat: a szkript elemei és a felsnitt szerinti shot-szakaszok.
# A build_forgatokonyv.py és a build_ad_app.py is innen veszi.
import json, os, subprocess
HERE = os.path.dirname(os.path.abspath(__file__))

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

