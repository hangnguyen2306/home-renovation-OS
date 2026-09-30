#!/usr/bin/env python3
"""Render the 10 listing images (2400×1800, 4:3) + a square Gumroad thumbnail.

Layouts are HTML (real workbook screenshots from marketing/shots/, made by capture.py),
captured with headless Chromium. Run from repo root:  python3 marketing/make_mockups.py
"""
import os

from playwright.sync_api import sync_playwright

HERE = os.path.dirname(os.path.abspath(__file__))
SHOTS = os.path.join(HERE, "shots")
OUT = os.path.join(HERE, "images")
CHROME = "/opt/pw-browsers/chromium-1194/chrome-linux/chrome"

W, H = 2400, 1800
TEAL, TEAL_DARK, CREAM, INK, GREY = "#0F5E5E", "#0A4242", "#F6F3EC", "#1F2937", "#5B6472"
CORAL, AMBER, GREEN = "#D9463B", "#F59E0B", "#2E7D32"

CSS = f"""
* {{ box-sizing: border-box; margin: 0; padding: 0; }}
html, body {{ width: {W}px; height: {H}px; }}
body {{ font-family: 'Liberation Sans', 'DejaVu Sans', Arial, sans-serif; background: {CREAM};
       color: {INK}; position: relative; overflow: hidden; }}
.brand {{ position: absolute; top: 70px; left: 110px; font-size: 34px; font-weight: 700;
         letter-spacing: 6px; color: {TEAL}; }}
.brand span {{ color: {AMBER}; }}
.edition {{ position: absolute; top: 62px; right: 110px; font-size: 30px; font-weight: 700;
           color: #fff; background: {TEAL}; padding: 14px 30px; border-radius: 40px; }}
h1 {{ position: absolute; top: 140px; left: 110px; right: 110px; font-size: 104px; line-height: 1.05;
     font-weight: 700; color: {TEAL_DARK}; letter-spacing: -2px; }}
h1 em {{ font-style: normal; color: {CORAL}; }}
.sub {{ position: absolute; left: 112px; right: 160px; font-size: 44px; line-height: 1.3; color: {GREY}; }}
.card {{ position: absolute; background: #fff; border-radius: 26px;
        box-shadow: 0 30px 80px rgba(15, 60, 60, .22), 0 4px 14px rgba(0,0,0,.06); overflow: hidden; }}
.chrome {{ height: 64px; background: #EEF2F2; display: flex; align-items: center; padding: 0 28px;
          border-bottom: 1px solid #DDE4E4; }}
.dot {{ width: 20px; height: 20px; border-radius: 50%; margin-right: 12px; }}
.tab {{ margin-left: 26px; font-size: 26px; font-weight: 700; color: {TEAL};
       background: #fff; padding: 10px 26px; border-radius: 12px 12px 0 0; margin-top: 18px; }}
.shot {{ display: flex; align-items: flex-start; justify-content: center; background: #fff; padding-top: 8px; }}
.shot img {{ display: block; }}
.pill {{ position: absolute; font-size: 40px; font-weight: 700; color: #fff; padding: 26px 40px;
        border-radius: 22px; box-shadow: 0 18px 40px rgba(0,0,0,.25); line-height: 1.2; }}
.pill small {{ display: block; font-size: 28px; font-weight: 400; opacity: .92; margin-top: 6px; }}
.badges {{ position: absolute; left: 110px; right: 110px; bottom: 64px; display: flex; gap: 22px;
          flex-wrap: wrap; }}
.badge {{ font-size: 32px; font-weight: 700; color: {TEAL}; background: #fff; border: 3px solid #CFE3E3;
         padding: 16px 28px; border-radius: 40px; }}
.foot {{ position: absolute; bottom: 60px; right: 110px; font-size: 26px; color: #8A9199; }}
"""


def img(name, width=None, height=None):
    style = []
    if width:
        style.append(f"width:{width}px")
    if height:
        style.append(f"height:{height}px")
    return f'<img src="file://{os.path.join(SHOTS, name)}.png" style="{";".join(style)}">'


def card(tab, inner, left, top, width, height):
    dots = "".join(f'<div class="dot" style="background:{c}"></div>'
                   for c in ("#F87171", "#FBBF24", "#34D399"))
    return (f'<div class="card" style="left:{left}px;top:{top}px;width:{width}px;height:{height}px">'
            f'<div class="chrome">{dots}<div class="tab">{tab}</div></div>'
            f'<div class="shot" style="height:{height - 64}px">{inner}</div></div>')


def pill(text, sub, color, **pos):
    p = ";".join(f"{k}:{v}px" for k, v in pos.items())
    s = f"<small>{sub}</small>" if sub else ""
    return f'<div class="pill" style="background:{color};{p}">{text}{s}</div>'


def page(body, edition="PRO"):
    ed = f'<div class="edition">{edition}</div>' if edition else ""
    return (f"<html><head><meta charset='utf-8'><style>{CSS}</style></head><body>"
            f'<div class="brand">RENOVATION <span>OS</span></div>{ed}{body}</body></html>')


def badges(items):
    return '<div class="badges">' + "".join(f'<div class="badge">{b}</div>' for b in items) + "</div>"


SLIDES = []

# 1 — hero
SLIDES.append(("01_hero", page(
    '<h1>Run your renovation<br>like a pro.</h1>'
    f'<div class="sub" style="top:390px">The all-in-one spreadsheet for homeowners managing a '
    f'renovation themselves — budget, contractors, payments and timeline in one place.</div>'
    + card("Dashboard", img("dash_top", width=1780), 310, 560, 1780, 1070)
    + pill("🔴 30+ actions ranked for you", "overdue payments, late orders, decisions…", CORAL,
           right=70, top=500)
    + badges(["Excel & Google Sheets", "Any currency", "No macros", "Instant download"]))))

# 2 — next actions
SLIDES.append(("02_next_actions", page(
    '<h1>Open it on Monday.<br><em>Know exactly what to do.</em></h1>'
    '<div class="sub" style="top:390px">Every overdue payment, late order, pending decision and '
    'site issue — ranked by urgency, automatically.</div>'
    + card("Dashboard — your next actions", img("dash_actions", width=2000), 200, 590, 2000, 1080)
    + pill("Ranked automatically", "no manual to-do lists", TEAL, right=80, top=520))))

# 3 — quotes
SLIDES.append(("03_hidden_costs", page(
    '<h1>The cheapest quote<br>is <em>not</em> always the cheapest.</h1>'
    '<div class="sub" style="top:390px">Compare up to 4 quotes per trade. The hidden-cost checklist '
    'reveals what is really included.</div>'
    + card("Quotes", img("quotes", height=1100), 330, 560, 1600, 1164)
    + pill("€4,000 of hidden extras", "the “cheap” quote became the expensive one", CORAL,
           right=70, top=1180)
    + pill("🏆 Winner by TRUE cost", None, GREEN, right=70, top=600))))

# 4 — overpayment
SLIDES.append(("04_paying_ahead", page(
    '<h1>Never pay ahead<br>of the work <em>again.</em></h1>'
    '<div class="sub" style="top:390px">Track % of work done vs % paid for every contractor. '
    'Get warned before you lose leverage.</div>'
    + card("Contractors", img("con_names", height=900) + img("con_pay", height=900),
           250, 620, 1900, 1000)
    + pill("⚠️ Paid 60% — only 35% done", "Renovation OS flags it instantly", CORAL,
           right=80, top=540))))

# 5 — timeline
SLIDES.append(("05_timeline", page(
    '<h1>One delay?<br><em>Every date updates itself.</em></h1>'
    '<div class="sub" style="top:390px">Tasks, dependencies and a 52-week Gantt chart. Late tasks '
    'turn red and your move-in date moves with them.</div>'
    + card("Timeline", img("tl_tasks", height=1020) + img("tl_gantt", height=1020),
           330, 590, 1740, 1090)
    + pill("Move-in date recalculated", "11 days after target — you know today", TEAL,
           right=70, top=530))))

# 6 — change orders
SLIDES.append(("06_change_orders", page(
    '<h1>See the impact<br><em>before</em> you say yes.</h1>'
    '<div class="sub" style="top:390px">Every change request shows your new total, contingency '
    'left and move-in date — before you approve it.</div>'
    + card("Change Orders", img("co_preview", width=2100), 150, 560, 2100, 490)
    + pill("This change pushes you over budget", "contingency left: −240", CORAL, right=120,
           top=1110)
    + badges(["New forecast total", "Contingency left", "New move-in date", "% over budget"]))))

# 7 — selections
SLIDES.append(("07_order_on_time", page(
    '<h1>Order on time.<br><em>Every time.</em></h1>'
    '<div class="sub" style="top:390px">Tiles, taps, appliances: decide-by and order-by dates are '
    'calculated from your timeline and supplier lead times.</div>'
    + card("Selections & Orders", img("sel_items", height=820) + img("sel_dates", height=820),
           140, 640, 2120, 900)
    + pill("⚠️ Order late — install at risk", "you'll know weeks earlier", CORAL, right=80,
           top=1600))))

# 8 — rooms
SLIDES.append(("08_rooms", page(
    '<h1>Every room<br>at a glance.</h1>'
    '<div class="sub" style="top:390px">Budget used, progress, decisions and open issues per room '
    '— rename rooms, everything follows.</div>'
    + card("Rooms", img("rooms", width=1700), 350, 580, 1700, 1170))))

# 9 — charts + health
SLIDES.append(("09_charts", page(
    '<h1>Know your final cost<br><em>before the final invoice.</em></h1>'
    '<div class="sub" style="top:390px">Budget vs forecast per category, cash needed per month, '
    'and a live health score for your whole project.</div>'
    + card("Dashboard — charts", img("dash_charts", width=2000), 200, 620, 2000, 900)
    + badges(["Forecast final cost", "Contingency used", "Cash flow per month",
              "🟢🟡🔴 Health score"]))))

# 10 — what's inside / Lite vs Pro
rows = [
    ("Dashboard: ranked next actions + health score", 1, 1),
    ("Budget, forecast & contingency tracking", 1, 1),
    ("Contractors + paying-ahead-of-work warning", 1, 1),
    ("Payments, invoices & 12-month cash flow", 1, 1),
    ("Timeline with dependencies + Gantt chart", 1, 1),
    ("Issues & punch list", 1, 1),
    ("Quote comparison with hidden costs", 0, 1),
    ("Change-order impact preview", 0, 1),
    ("Decide-by / order-by dates for materials", 0, 1),
    ("Room cards · weekly planner · document vault · warranties", 0, 1),
]
tbl = "".join(
    f'<tr><td class="f">{f}</td><td>{"<b class=y>✓</b>" if lite else "<b class=n>—</b>"}</td>'
    f'<td><b class=y>✓</b></td></tr>' for f, lite, _ in rows)
SLIDES.append(("10_lite_vs_pro", page(
    '<h1>What’s inside</h1>'
    '<div class="sub" style="top:270px">Two editions. Same engine. Pick what your project needs.</div>'
    '<style>table{position:absolute;left:180px;width:2040px;top:380px;border-collapse:separate;'
    'border-spacing:0;background:#fff;border-radius:26px;overflow:hidden;'
    'box-shadow:0 30px 80px rgba(15,60,60,.18);font-size:40px}'
    'th{background:#0F5E5E;color:#fff;padding:26px 30px;font-size:46px}'
    'td{padding:19px 30px;border-bottom:2px solid #EEF2F2;text-align:center;width:260px}'
    'td.f{text-align:left;width:auto}.y{color:#2E7D32;font-size:48px}.n{color:#B8BEC4}'
    'th small{display:block;font-size:28px;font-weight:400;opacity:.85;margin-top:6px}</style>'
    '<table><tr><th style="text-align:left">Feature</th><th>LITE<small>7 tabs</small></th>'
    f'<th>PRO<small>16 tabs</small></th></tr>{tbl}</table>'
    + badges(["Excel 2016+ & Google Sheets", "Any currency · m² or ft²", "Demo + blank file",
              "Quick Start guide"]), edition=None)))


def square_thumb():
    css = CSS.replace(f"width: {W}px; height: {H}px;", "width: 1200px; height: 1200px;")
    return (f"<html><head><meta charset='utf-8'><style>{css}"
            "h1{top:120px;font-size:96px}</style></head><body>"
            '<div class="brand">RENOVATION <span>OS</span></div>'
            '<h1>Run your renovation<br>like a pro.</h1>'
            + card("Dashboard", img("dash_top", width=1000), 100, 420, 1000, 600)
            + '<div class="foot" style="left:110px;right:auto;bottom:60px;font-size:34px;'
              'color:#0F5E5E;font-weight:700">Excel &amp; Google Sheets planner</div>'
            + "</body></html>")


def main():
    os.makedirs(OUT, exist_ok=True)
    with sync_playwright() as p:
        browser = p.chromium.launch(executable_path=CHROME)
        pg = browser.new_page(viewport={"width": W, "height": H})
        for name, html in SLIDES:
            path = os.path.join(OUT, f"{name}.html")
            with open(path, "w") as f:
                f.write(html)
            pg.goto(f"file://{path}")
            pg.wait_for_timeout(300)
            pg.screenshot(path=os.path.join(OUT, f"{name}.png"))
            os.remove(path)
            print("image", name)
        pg.set_viewport_size({"width": 1200, "height": 1200})
        path = os.path.join(OUT, "thumb.html")
        with open(path, "w") as f:
            f.write(square_thumb())
        pg.goto(f"file://{path}")
        pg.wait_for_timeout(300)
        pg.screenshot(path=os.path.join(OUT, "gumroad_thumbnail_1200.png"))
        os.remove(path)
        browser.close()


if __name__ == "__main__":
    main()
