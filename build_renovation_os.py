#!/usr/bin/env python3
"""Build the Renovation OS workbook (.xlsx) — 100% formulas, no macros, no external links.

Usage:
    python build_renovation_os.py                      # all 8 files (Pro + Lite) into dist/
    python build_renovation_os.py --edition LITE       # Lite only
    python build_renovation_os.py --currency EUR --variant DEMO
    python build_renovation_os.py --currency USD --variant BLANK --out somewhere/
"""
import argparse
import os
import secrets
import sys

from openpyxl import Workbook
from openpyxl.styles import Font

import demo_data
from ros import core as C
from ros import sheets_setup, sheets_money, sheets_exec, sheets_after, sheets_overview

PRESETS = {
    "EUR": dict(currency="EUR €", tax="VAT", rate=0.21, area="m²", week="Monday"),
    "USD": dict(currency="USD $", tax="Sales tax", rate=0.08, area="ft²", week="Sunday"),
}
M2_TO_FT2 = 10.7639


class Context:
    def __init__(self, currency, demo, lite=False):
        self.code = currency
        self.cur = PRESETS[currency]
        self.demo = demo
        self.lite = lite
        self.data = demo_data
        self.wb = Workbook()
        self.wb.remove(self.wb.active)
        for name in (C.LITE_SHEETS if lite else C.SHEET_ORDER):
            self.wb.create_sheet(name)
        normal = self.wb._named_styles["Normal"]
        normal.font = Font(name=C.FONT, size=10)
        self.lists = C.Lists()
        self.dv = C.Validations(self.lists)

    def ws(self, name):
        return self.wb[name]

    def area(self, m2):
        if m2 is None:
            return None
        return round(m2 * M2_TO_FT2) if self.cur["area"] == "ft²" else m2


BUILDERS = [
    (C.S_START, sheets_setup.build_start),
    (C.S_BUDGET, sheets_money.build_budget),
    (C.S_QUOTES, sheets_money.build_quotes),
    (C.S_CON, sheets_money.build_contractors),
    (C.S_PAY, sheets_money.build_payments),
    (C.S_CO, sheets_money.build_change_orders),
    (C.S_TL, sheets_exec.build_timeline),
    (C.S_ROOMS, sheets_exec.build_rooms),
    (C.S_SEL, sheets_exec.build_selections),
    (C.S_ISS, sheets_exec.build_issues),
    (C.S_VAULT, sheets_after.build_vault),
    (C.S_WAR, sheets_after.build_warranty),
    (C.S_ENG, sheets_overview.build_engine),
    (C.S_DASH, sheets_overview.build_dashboard),
    (C.S_WEEK, sheets_overview.build_this_week),
    (C.S_LISTS, sheets_setup.build_lists),
]


PASSWORD_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".protection_password")


def protection_password():
    """Sheet/workbook password: $RENOVATION_OS_PASSWORD, else .protection_password (git-ignored),
    else a new random one saved there. Keep it — you need it to edit your own master files."""
    pw = os.environ.get("RENOVATION_OS_PASSWORD")
    if pw:
        return pw
    if os.path.exists(PASSWORD_FILE):
        return open(PASSWORD_FILE).read().strip()
    pw = secrets.token_urlsafe(12)
    with open(PASSWORD_FILE, "w") as f:
        f.write(pw + "\n")
    return pw


def file_name(edition, variant, currency):
    prefix = "Renovation_OS_LITE" if edition == "LITE" else "Renovation_OS"
    return f"{prefix}_{variant}_{currency}.xlsx"


def build(currency, variant, out_dir, edition="PRO", password=None):
    ctx = Context(currency, variant == "DEMO", lite=edition == "LITE")
    sheets_setup.register_lists(ctx)
    for sheet, fn in BUILDERS:
        if sheet in ctx.wb.sheetnames:
            fn(ctx)
    C.harden(ctx.wb, password or protection_password())

    errors = C.lint_workbook(ctx.wb)
    if errors:
        print("\n".join(errors[:50]))
        raise SystemExit(f"Formula lint failed: {len(errors)} problem(s)")
    ctx.wb.active = 0
    os.makedirs(out_dir, exist_ok=True)
    path = os.path.join(out_dir, file_name(edition, variant, currency))
    ctx.wb.save(path)
    return path


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--currency", choices=["EUR", "USD", "ALL"], default="ALL")
    ap.add_argument("--variant", choices=["DEMO", "BLANK", "ALL"], default="ALL")
    ap.add_argument("--edition", choices=["PRO", "LITE", "ALL"], default="ALL")
    ap.add_argument("--out", default="dist")
    ap.add_argument("--password", help="sheet/workbook protection password "
                    "(default: $RENOVATION_OS_PASSWORD or .protection_password)")
    a = ap.parse_args(argv)
    curs = ["EUR", "USD"] if a.currency == "ALL" else [a.currency]
    vars_ = ["DEMO", "BLANK"] if a.variant == "ALL" else [a.variant]
    eds = ["PRO", "LITE"] if a.edition == "ALL" else [a.edition]
    for ed in eds:
        for cur in curs:
            for var in vars_:
                print("built", build(cur, var, a.out, ed, a.password))


if __name__ == "__main__":
    sys.exit(main())
