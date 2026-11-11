# KAPASZKODNI – diszpó

Interaktív diszpozíció a *Kapaszkodni* rövidfilm forgatásához (2026. 11. 12.).

- `diszpo/index.html` – a diszpó oldal (shotlist vilkóval, kellékek, stáb, karakterek, inspó, szkript)
- Szerkeszthető ütemterv (Google Táblázatokhoz is): `diszpo/Kapaszkodni_utemterv_2026-11-12.xlsx` – újragenerálás: `python3 diszpo/build_xlsx.py`
- Rendezői V4 ütemterv + forgatókönyv forgatási sorrendben: `diszpo/Kapaszkodni_utemterv_V5.xlsx`, `diszpo/Kapaszkodni_forgatokonyv_forgatasi_sorrend.pdf` – újragenerálás: `python3 diszpo/build_forgatokonyv.py && node diszpo/build_pdf.js`
- Ütemterv-PDF a szerkesztett táblázatból: `python3 diszpo/sync_utemterv.py [táblázat.xlsx] && node diszpo/build_pdf.js`
- 1st AD forgatási app: `diszpo/ad_app.html` (sablon: `diszpo/ad_app_template.html`) – újragenerálás a táblázatból: `python3 diszpo/build_ad_app.py [táblázat.xlsx]`
- **Telepíthető, offline app (GitHub Pages):** a `docs/` mappa. Élő cím a Pages bekapcsolása után: https://xdkoky.github.io/ANDRIS--LLOYD/ – újragenerálás: `python3 diszpo/build_ad_app.py && python3 diszpo/build_offline.py`
- Eredeti minőségű képek (vilkó, storyboard) az anyagok/ PDF-jeiből: `python3 diszpo/build_images.py`
- Nyomtatható PDF-ek: `diszpo/Kapaszkodni_diszpo_2026-11-12.pdf`, `diszpo/Kapaszkodni_shotlist.pdf`, `diszpo/Kapaszkodni_forgatokonyv.pdf` – újragenerálás: `node diszpo/build_pdf.js`
- `diszpo/vilko/`, `diszpo/inspo/` – a PDF-ekből kivett képek
- `anyagok/` – az eredeti PDF-ek
