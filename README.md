# KAPASZKODNI – diszpó

Interaktív diszpozíció a *Kapaszkodni* rövidfilm forgatásához (2026. 11. 12.).

- `diszpo/index.html` – a diszpó oldal (shotlist vilkóval, kellékek, stáb, karakterek, inspó, szkript)
- Szerkeszthető ütemterv (Google Táblázatokhoz is): `diszpo/Kapaszkodni_utemterv_2026-11-12.xlsx` – újragenerálás: `python3 diszpo/build_xlsx.py`
- Nyomtatható PDF-ek: `diszpo/Kapaszkodni_diszpo_2026-11-12.pdf`, `diszpo/Kapaszkodni_shotlist.pdf`, `diszpo/Kapaszkodni_forgatokonyv.pdf` – újragenerálás: `node diszpo/build_pdf.js`
- `diszpo/vilko/`, `diszpo/inspo/` – a PDF-ekből kivett képek
- `anyagok/` – az eredeti PDF-ek
