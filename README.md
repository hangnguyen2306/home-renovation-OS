# Renovation OS

A 100% formula-driven home-renovation management workbook for Microsoft Excel 2016+ / Microsoft 365 (Windows & Mac):
no macros, no scripts, no external links.

| Command | What it does |
|---|---|
| `pip install openpyxl` | the only build dependency |
| `python3 build_renovation_os.py` | builds 4 files: `dist/Renovation_OS_{DEMO,BLANK}.xlsx` (Pro) and `dist/Renovation_OS_LITE_{DEMO,BLANK}.xlsx`. One file serves every country: currency, tax, area unit and week start are chosen on Start Here |
| `python3 build_renovation_os.py --edition LITE` | Lite only (Start Here, Dashboard, Budget, Contractors, Payments, Timeline, Issues) |
| `python3 build_renovation_os.py --variant BLANK --currency USD` | one file, with US starting defaults (saved as `…_USD.xlsx`) |
| `python3 test_workbook.py` | builds everything, recalculates in headless LibreOffice (needs `libreoffice-calc` + `python3-uno`), fails on any formula error or circular reference, and checks the demo numbers against `reference_model.py` |

- `PLAN.md` — sheet layouts, cell map, alert engine design.
- `QUICKSTART.md` — 1-page buyer guide.
- `EXCEL_CHECKLIST.md` — manual checks in real Excel before each release.
- Code: `ros/core.py` (styles, helpers, formula whitelist linter), `ros/layout.py` (cell map),
  `ros/sheets_*.py` (one builder per tab group), `demo_data.py` (demo scenario).
