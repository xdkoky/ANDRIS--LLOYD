# Eredeti minőségű képek kinyerése az anyagok/ PDF-jeiből (veszteség nélkül, a beágyazott JPEG-ek) + kis előnézetek.
# Futtatás: python3 diszpo/build_images.py
#   vilko/full/<id>.jpg       eredeti világítási tervek (3668×2280)
#   vilko/t/<id>.jpg          előnézet a kártyákhoz
#   storyboard/full/<id>.jpg  eredeti storyboard-rajzok
#   storyboard/t/<id>.jpg     előnézet a kártyákhoz és az időrendhez
import glob, os, shutil, subprocess, tempfile, unicodedata
from PIL import Image

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "..", "anyagok")
VILKO = ["1-1", "1-2", "1-3A", "1-3B", "2-1A", "2-2", "2-3A", "2-1B", "2-1C", "2-4", "2-5", "2-1E", "2-6", "2-3B", "3-1", "3-2"]
BOARD = ["1-1", "1-2", "1-3A", "1-3B", "2-1A", "2-2", "2-3A", "2-1B", "2-1C", "2-4", "2-5", "2-1E", "2-6", "2-3B"]

def find(part):
    for f in os.listdir(SRC):
        if part in unicodedata.normalize("NFC", f).lower() and f.lower().endswith(".pdf") and "minta" not in f.lower():
            return os.path.join(SRC, f)
    raise SystemExit("nincs meg: " + part)

def extract(pdf, names, outdir):
    os.makedirs(outdir, exist_ok=True)
    tmp = tempfile.mkdtemp()
    subprocess.run(["pdfimages", "-j", pdf, os.path.join(tmp, "i")], check=True)
    files = sorted(glob.glob(os.path.join(tmp, "i-*")))
    assert len(files) >= len(names), (pdf, len(files))
    for f, n in zip(files, names):
        dst = os.path.join(outdir, n + ".jpg")
        if f.lower().endswith((".jpg", ".jpeg")):
            shutil.copy(f, dst)                       # az eredeti JPEG-adat, újratömörítés nélkül
        else:
            Image.open(f).convert("RGB").save(dst, quality=95, subsampling=0)
    shutil.rmtree(tmp)

extract(find("vilko"), VILKO, os.path.join(HERE, "vilko", "full"))
extract(find("shotlist"), BOARD, os.path.join(HERE, "storyboard", "full"))
os.makedirs(os.path.join(HERE, "vilko", "t"), exist_ok=True)
for n in VILKO:
    im = Image.open(os.path.join(HERE, "vilko", "full", n + ".jpg")).convert("RGB")
    im.thumbnail((720, 720)); im.save(os.path.join(HERE, "vilko", "t", n + ".jpg"), quality=74, optimize=True, progressive=True)
os.makedirs(os.path.join(HERE, "storyboard", "t"), exist_ok=True)
for n in BOARD:
    im = Image.open(os.path.join(HERE, "storyboard", "full", n + ".jpg")).convert("RGB")
    im.thumbnail((480, 480)); im.save(os.path.join(HERE, "storyboard", "t", n + ".jpg"), quality=78, optimize=True, progressive=True)
for d in ("vilko/full", "vilko/t", "storyboard/full", "storyboard/t"):
    fs = glob.glob(os.path.join(HERE, d, "*.jpg"))
    print(d, len(fs), "fájl,", round(sum(os.path.getsize(f) for f in fs) / 1e6, 1), "MB,", Image.open(fs[0]).size)
