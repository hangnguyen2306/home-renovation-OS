#!/usr/bin/env python3
"""Capture real screenshots of the demo workbooks (LibreOffice → PDF → PNG, trimmed).

Used by make_mockups.py. Run from the repo root:  python3 marketing/capture.py
"""
import os
import sys

import pymupdf
import uno
from com.sun.star.beans import PropertyValue
from PIL import Image, ImageChops

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))
import test_workbook as T  # noqa: E402  (reuses the headless LibreOffice helper)

OUT = os.path.join(HERE, "shots")
PRO = os.path.join(os.path.dirname(HERE), "dist", "Renovation_OS_DEMO.xlsx")
LITE = os.path.join(os.path.dirname(HERE), "dist", "Renovation_OS_LITE_DEMO.xlsx")

SHOTS = [  # (file, output name, sheet, range)
    (PRO, "dash_top", "Dashboard", "B4:K23"),
    (PRO, "dash_actions", "Dashboard", "B25:K38"),
    (PRO, "dash_charts", "Dashboard", "B40:K62"),
    (PRO, "dash_full", "Dashboard", "A1:L62"),
    (PRO, "quotes", "Quotes", "B9:J45"),
    (PRO, "contractors", "Contractors", "B5:U16"),
    (PRO, "timeline", "Timeline", "B4:BK30"),
    (PRO, "change_orders", "Change Orders", "B4:Q11"),
    (PRO, "selections", "Selections & Orders", "B4:Y20"),
    (PRO, "rooms", "Rooms", "B5:L27"),
    (PRO, "budget", "Budget", "B7:N64"),
    (PRO, "this_week", "This Week", "B4:G28"),
    (PRO, "start_here", "Start Here", "A1:L36"),
    (PRO, "payments_cash", "Payments", "AH4:AL19"),
    (LITE, "lite_dash", "Dashboard", "A1:L38"),
    (PRO, "con_names", "Contractors", "B5:B15"),
    (PRO, "con_pay", "Contractors", "R5:U15"),
    (PRO, "tl_tasks", "Timeline", "D6:D29"),
    (PRO, "tl_gantt", "Timeline", "X6:BA29"),
    (PRO, "sel_items", "Selections & Orders", "B7:B20"),
    (PRO, "sel_dates", "Selections & Orders", "R7:Y20"),
    (PRO, "co_preview", "Change Orders", "H6:Q11"),
]


def pv(n, v):
    p = PropertyValue()
    p.Name, p.Value = n, v
    return p


def trim(img, pad=6):
    bg = Image.new(img.mode, img.size, (255, 255, 255))
    box = ImageChops.difference(img, bg).getbbox()
    if box:
        l, t, r, b = box
        img = img.crop((max(0, l - pad), max(0, t - pad), min(img.width, r + pad),
                        min(img.height, b + pad)))
    return img


def main():
    os.makedirs(OUT, exist_ok=True)
    office = T.Office()
    try:
        docs = {}
        for path, name, sheet, rng in SHOTS:
            if path not in docs:
                docs[path] = office.open(path)
            doc = docs[path]
            sh = doc.Sheets.getByName(sheet)
            ps = doc.StyleFamilies.getByName("PageStyles").getByName(sh.PageStyle)
            ps.IsLandscape = True
            ps.Width, ps.Height = 84100, 59400          # A1: big page = crisp text
            ps.ScaleToPagesX, ps.ScaleToPagesY = 1, 1
            ps.HeaderIsOn = ps.FooterIsOn = False
            ps.PrintGrid = False
            ps.LeftMargin = ps.RightMargin = ps.TopMargin = ps.BottomMargin = 400
            fd = uno.Any("[]com.sun.star.beans.PropertyValue",
                         (pv("Selection", sh.getCellRangeByName(rng)), pv("ExportNotes", False)))
            pdf = os.path.join(OUT, name + ".pdf")
            doc.storeToURL(uno.systemPathToFileUrl(pdf),
                           (pv("FilterName", "calc_pdf_Export"), pv("FilterData", fd)))
            page = pymupdf.open(pdf)[0]
            pix = page.get_pixmap(dpi=230)
            img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
            img = trim(img)
            img.save(os.path.join(OUT, name + ".png"))
            os.remove(pdf)
            print(f"{name}: {img.size}")
        for d in docs.values():
            d.close(True)
    finally:
        office.close()


if __name__ == "__main__":
    main()
