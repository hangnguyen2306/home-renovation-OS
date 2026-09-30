#!/usr/bin/env python3
"""Build the Renovation OS workbook (.xlsx) — 100% formulas, no macros, no external links.

Usage:
    python build_renovation_os.py                      # all 4 files into dist/
    python build_renovation_os.py --currency EUR --variant DEMO
    python build_renovation_os.py --currency USD --variant BLANK --out somewhere/
"""
import argparse
import os
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
    def __init__(self, currency, demo):
        self.code = currency
        self.cur = PRESETS[currency]
        self.demo = demo
        self.data = demo_data
        self.wb = Workbook()
        self.wb.remove(self.wb.active)
        for name in C.SHEET_ORDER:
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


def build(currency, variant, out_dir):
    ctx = Context(currency, variant == "DEMO")
    sheets_setup.register_lists(ctx)
    sheets_setup.build_start(ctx)
    sheets_money.build_budget(ctx)
    sheets_money.build_quotes(ctx)
    sheets_money.build_contractors(ctx)
    sheets_money.build_payments(ctx)
    sheets_money.build_change_orders(ctx)
    sheets_exec.build_timeline(ctx)
    sheets_exec.build_rooms(ctx)
    sheets_exec.build_selections(ctx)
    sheets_exec.build_issues(ctx)
    sheets_after.build_vault(ctx)
    sheets_after.build_warranty(ctx)
    sheets_overview.build_engine(ctx)
    sheets_overview.build_dashboard(ctx)
    sheets_overview.build_this_week(ctx)
    sheets_setup.build_lists(ctx)

    errors = C.lint_workbook(ctx.wb)
    if errors:
        print("\n".join(errors[:50]))
        raise SystemExit(f"Formula lint failed: {len(errors)} problem(s)")
    ctx.wb.active = 0
    os.makedirs(out_dir, exist_ok=True)
    path = os.path.join(out_dir, f"Renovation_OS_{variant}_{currency}.xlsx")
    ctx.wb.save(path)
    return path


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--currency", choices=["EUR", "USD", "ALL"], default="ALL")
    ap.add_argument("--variant", choices=["DEMO", "BLANK", "ALL"], default="ALL")
    ap.add_argument("--out", default="dist")
    a = ap.parse_args(argv)
    curs = ["EUR", "USD"] if a.currency == "ALL" else [a.currency]
    vars_ = ["DEMO", "BLANK"] if a.variant == "ALL" else [a.variant]
    for cur in curs:
        for var in vars_:
            print("built", build(cur, var, a.out))


if __name__ == "__main__":
    sys.exit(main())
