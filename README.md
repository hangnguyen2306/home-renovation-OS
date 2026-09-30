# Renovation OS

A 100% formula-driven home-renovation management workbook (Excel 2016+ and Google Sheets):
no macros, no scripts, no external links.

| Command | What it does |
|---|---|
| `pip install openpyxl` | the only build dependency |
| `python3 build_renovation_os.py` | builds all 8 files: Pro `dist/Renovation_OS_{DEMO,BLANK}_{EUR,USD}.xlsx` and Lite `dist/Renovation_OS_LITE_…` |
| `python3 build_renovation_os.py --edition LITE` | Lite only (Start Here, Dashboard, Budget, Contractors, Payments, Timeline, Issues) |
| `python3 build_renovation_os.py --currency USD --variant BLANK` | one file |
| `python3 test_workbook.py` | builds everything, recalculates in headless LibreOffice (needs `libreoffice-calc` + `python3-uno`), fails on any formula error or circular reference, and checks the demo numbers against `reference_model.py` |

- `PLAN.md` — sheet layouts, cell map, alert engine design.
- `QUICKSTART.md` — 1-page buyer guide.
- `GOOGLE_SHEETS_CHECKLIST.md` — manual checks after uploading to Google Sheets.
- Code: `ros/core.py` (styles, helpers, formula whitelist linter), `ros/layout.py` (cell map),
  `ros/sheets_*.py` (one builder per tab group), `demo_data.py` (demo scenario).
