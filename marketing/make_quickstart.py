#!/usr/bin/env python3
"""Build the buyer's Quick Start guide PDF (A4) that ships with every order.

Fill in the two settings below, then run from the repo root:
    python3 marketing/make_quickstart.py
Output: marketing/Renovation_OS_Quick_Start.pdf
"""
import os

from playwright.sync_api import sync_playwright

# ---------------------------------------------------------------- settings (edit these)
SUPPORT = ""          # e.g. "hello@yourdomain.com" — empty = "message me on Etsy / Gumroad"
YEAR, AUTHOR = 2026, "Hang Nguyen"
NOTICE = f"© {YEAR} Renovation OS — created by {AUTHOR}"

HERE = os.path.dirname(os.path.abspath(__file__))
SHOTS = os.path.join(HERE, "shots")
CHROME = "/opt/pw-browsers/chromium-1194/chrome-linux/chrome"
OUT = os.path.join(HERE, "Renovation_OS_Quick_Start.pdf")


def shot(name, width="100%"):
    return f'<img class="shot" src="file://{os.path.join(SHOTS, name)}.png" style="width:{width}">'


support = (f"email <b>{SUPPORT}</b>" if SUPPORT else
           "send me a message through the shop where you bought it (Etsy or Gumroad)")
HTML = f"""<!doctype html><html><head><meta charset="utf-8"><style>
@page {{ size: A4; margin: 16mm 16mm 18mm 16mm; }}
* {{ box-sizing: border-box; }}
body {{ font-family: 'Liberation Sans', 'DejaVu Sans', Arial, sans-serif; color: #1F2937;
       font-size: 10.5pt; line-height: 1.5; margin: 0; }}
h1 {{ font-size: 30pt; color: #0A4242; margin: 0 0 4mm; line-height: 1.1; letter-spacing: -.5pt; }}
h2 {{ font-size: 15pt; color: #0F5E5E; margin: 9mm 0 3mm; padding-bottom: 1.5mm;
     border-bottom: 1.2pt solid #0F5E5E; }}
h3 {{ font-size: 11.5pt; color: #0A4242; margin: 5mm 0 1.5mm; }}
p {{ margin: 0 0 2.5mm; }}
.brand {{ font-weight: 700; letter-spacing: 3pt; color: #0F5E5E; font-size: 10pt; }}
.brand span {{ color: #F59E0B; }}
.lead {{ font-size: 13pt; color: #5B6472; }}
.cover {{ page-break-after: always; }}
.shot {{ border: 0.6pt solid #D5DEDE; border-radius: 3mm; box-shadow: 0 2mm 6mm rgba(0,0,0,.08);
        margin: 3mm 0; }}
.steps {{ counter-reset: s; list-style: none; padding: 0; margin: 0; }}
.steps > li {{ counter-increment: s; position: relative; padding: 2.5mm 0 2.5mm 12mm; }}
.steps > li::before {{ content: counter(s); position: absolute; left: 0; top: 2mm; width: 8mm;
                      height: 8mm; border-radius: 50%; background: #0F5E5E; color: #fff;
                      font-weight: 700; text-align: center; line-height: 8mm; }}
.box {{ background: #E3F1F1; border-radius: 3mm; padding: 4mm 5mm; margin: 3mm 0; }}
.warn {{ background: #FFEDD5; }}
.legend td {{ padding: 1.6mm 3mm; vertical-align: middle; }}
.sw {{ width: 12mm; height: 6mm; border: .6pt solid #C9D2D2; border-radius: 1mm; }}
table.tabs {{ width: 100%; border-collapse: collapse; font-size: 9.8pt; }}
table.tabs th {{ background: #0F5E5E; color: #fff; text-align: left; padding: 2mm 3mm; }}
table.tabs td {{ padding: 1.8mm 3mm; border-bottom: .5pt solid #E1E7E7; vertical-align: top; }}
.pro {{ display: inline-block; font-size: 7.5pt; font-weight: 700; color: #fff; background: #0F5E5E;
       border-radius: 2mm; padding: .3mm 1.8mm; margin-left: 1.5mm; vertical-align: 1pt; }}
.two {{ display: flex; gap: 6mm; }} .two > div {{ flex: 1; }}
.small {{ font-size: 8.5pt; color: #6B7280; }}
.nb {{ page-break-inside: avoid; }}
.pb {{ page-break-before: always; }}
</style></head><body>

<section class="cover">
  <div class="brand">RENOVATION <span>OS</span></div>
  <div style="height:14mm"></div>
  <h1>Quick Start Guide</h1>
  <p class="lead">Set up your renovation in 10 minutes — then spend 10 minutes a week
     staying in control of budget, contractors, payments and timeline.</p>
  {shot("dash_top")}
  <div class="box">
    <b>In your download</b>
    <ul style="margin:1.5mm 0 0 5mm;padding:0">
      <li><b>…_DEMO.xlsx</b> — a realistic example renovation, frozen on 15 Jun 2026 so every
          feature shows something. Open this first to explore.</li>
      <li><b>…_BLANK.xlsx</b> — the same workbook, empty and ready for <i>your</i> project.</li>
      <li>This guide.</li>
    </ul>
  </div>
  <p class="small" style="margin-top:8mm">{NOTICE}. Licensed to the
  original purchaser for personal use only — see “Licence” at the end of this guide.</p>
</section>

<h2>1 · Open the file in Microsoft Excel</h2>
<div class="two">
  <div class="nb"><h3>Get started</h3>
    <ol class="steps">
      <li>Save the <b>BLANK</b> file somewhere safe, e.g. a “Renovation” folder on your computer
          or OneDrive.</li>
      <li>Open it in <b>Microsoft Excel</b>. If Excel shows a yellow “Protected View” bar, click
          <b>Enable Editing</b>.</li>
      <li>Start on the <b>Start Here</b> tab.</li>
    </ol></div>
  <div class="nb"><h3>Requirements</h3>
    <ul style="margin:2mm 0 0 5mm;padding:0">
      <li>Microsoft Excel 2016, 2019, 2021, 2024 or Microsoft 365.</li>
      <li>Windows or Mac.</li>
      <li>No macros, add-ons or internet connection needed.</li>
      <li>A computer is recommended — the tabs are designed for a full screen.</li>
    </ul>
    <p class="small" style="margin-top:3mm">Renovation OS is built and supported for Microsoft
    Excel only. Other spreadsheet apps (such as Google Sheets or Apple Numbers) do not support
    all of its features and protection.</p></div>
</div>
<div class="box nb"><b>Tip:</b> keep an untouched copy of the BLANK file as a backup, and save
your project file regularly (Excel ▸ File ▸ Save). If you store it in OneDrive, Excel keeps
version history for you.</div>

<h2>2 · The only rule: yellow = you type</h2>
<table class="legend">
  <tr><td><div class="sw" style="background:#FFF7CC"></div></td>
      <td><b>Light-yellow cells</b> are yours: type, or pick from the ▼ dropdown.</td></tr>
  <tr><td><div class="sw" style="background:#FFFFFF"></div></td>
      <td><b>White / grey cells</b> calculate automatically. They are locked, so nothing can
      break.</td></tr>
  <tr><td><div class="sw" style="background:#D1FAE5"></div></td><td>🟢 On track — nothing to do.</td></tr>
  <tr><td><div class="sw" style="background:#FFEDD5"></div></td><td>🟡 Needs attention soon (within your alert window).</td></tr>
  <tr><td><div class="sw" style="background:#FEE2E2"></div></td><td>🔴 Urgent — overdue, over budget or at risk.</td></tr>
</table>

<h2>3 · Set up in 10 minutes</h2>
<ol class="steps">
  <li><b>Start Here.</b> Project name, start date, target move-in date and total budget
      (including tax). Choose your <b>currency, tax name, m² or ft², and the day your week
      starts</b>. A safety buffer (contingency, 10% by default) is set aside automatically —
      what is left is your <i>working budget</i>.</li>
  <li><b>Budget.</b> Give each category a budget. Rename, add or clear categories freely. Try to
      make the total match your working budget — the check at the top tells you.</li>
  <li><b>Contractors.</b> One row per company: contract value, signed yes/no, start and end
      dates. Update <b>“% work complete”</b> every week — it powers the
      “paying ahead of work” warning.</li>
  <li><b>Payments.</b> Enter each payment milestone (a % of the contract or a fixed amount)
      with its due date. Add invoice and paid dates as they happen.</li>
  <li><b>Timeline.</b> The typical task list is pre-filled. Set durations (days) and the
      <b>Predecessor ID</b> — the task that must finish first. Dates, the Gantt chart and your
      forecast move-in date follow automatically.</li>
</ol>
<div class="box nb"><b>What is a “Predecessor ID”?</b> Every task has a number. If task 11
“Rough-in inspection” can only start after task 8 “Plumbing rough-in”, type <b>8</b> in task 11’s
Predecessor column. When plumbing slips, the inspection — and everything after it — moves too.
The predecessor must be a task listed <i>above</i> (a lower ID).</div>

<h2>4 · Your weekly routine (10 minutes)</h2>
<div class="two">
  <div><h3>Monday — plan</h3>
    <p>Open the <b>Dashboard</b>. Work through <b>“Your next actions”</b> from top to bottom:
    overdue payments, late tasks, decisions and follow-ups are already ranked by urgency.</p>
    <p><span class="pro">PRO</span> <b>This Week</b> shows everything due this week plus a
    Friday review.</p></div>
  <div><h3>During the week — update</h3>
    <p>Record payments and invoices, mark tasks Done, log problems on <b>Issues</b>, and update
    each contractor’s % work complete. Everything else recalculates.</p></div>
</div>
{shot("dash_actions")}

<div class="nb"><h2>5 · Every tab in one line</h2>
<table class="tabs">
  <tr><th style="width:34%">Tab</th><th>What it does</th></tr>
  <tr><td>Start Here</td><td>Your project settings, preferences and this routine.</td></tr>
  <tr><td>Dashboard</td><td>KPIs, 🟢🟡🔴 health score, ranked next actions, charts.</td></tr>
  <tr><td>Budget</td><td>Budget vs committed vs invoiced vs paid, forecast final cost and risk.</td></tr>
  <tr><td>Contractors</td><td>Contracts, insurance/licence checks, paying-ahead warning; communication log on the right →.</td></tr>
  <tr><td>Payments</td><td>Payment schedule and invoices; per-contractor summary and 12-month cash flow on the right →.</td></tr>
  <tr><td>Timeline</td><td>Tasks, dependencies, delays, late warnings and a 52-week Gantt chart.</td></tr>
  <tr><td>Issues &amp; Punch List</td><td>Problems on site and end-of-job snags, with owner, cost and deadline.</td></tr>
  <tr><td>This Week <span class="pro">PRO</span></td><td>Monday plan and Friday review.</td></tr>
  <tr><td>Quotes <span class="pro">PRO</span></td><td>Up to 4 quotes side by side; the hidden-cost checklist gives the TRUE cost and a 🏆 winner.</td></tr>
  <tr><td>Change Orders <span class="pro">PRO</span></td><td>Preview a change’s effect on budget, contingency and move-in date before approving.</td></tr>
  <tr><td>Rooms <span class="pro">PRO</span></td><td>One card per room: money, progress, decisions, issues.</td></tr>
  <tr><td>Selections &amp; Orders <span class="pro">PRO</span></td><td>Choices (tiles, taps…) with automatic decide-by and order-by dates.</td></tr>
  <tr><td>Vault <span class="pro">PRO</span></td><td>Links to contracts, invoices, photos + ✓ checklist of missing documents.</td></tr>
  <tr><td>Warranty &amp; Maintenance <span class="pro">PRO</span></td><td>Warranty expiry dates and recurring home maintenance, with reminders.</td></tr>
</table>
<p class="small">Lite users: there is no Change Orders tab — add approved extras to the
contractor’s contract value. There is no Selections tab — enter big material suppliers
(worktop, appliances…) as contractors.</p></div>

<h2>6 · Good to know</h2>
<div class="nb"><h3>“The cell won’t let me type.”</h3>
<p>It is a calculated (white) cell — it is locked on purpose. Type only in yellow cells.</p></div>
<div class="nb"><h3>How many rows can I use?</h3>
<p>Every list has 1,000 ready rows (Budget: 50 categories). Just use the next empty yellow row —
please don’t insert or delete rows.</p></div>
<div class="nb"><h3>The demo shows dates in 2026 and never changes.</h3>
<p>The demo is frozen on 15 Jun 2026 via Start Here ▸ “Today” override, so every feature shows
data. Clear that cell to see it live. Your BLANK file always uses today’s date.</p></div>
<div class="nb"><h3>Why is a budget line red?</h3>
<p>Its forecast is above its budget. While a category is still <i>Estimating</i> the forecast
uses the highest of budget, best quote, committed and invoiced — deliberately pessimistic until
you sign.</p></div>
<div class="nb"><h3>Amounts show no € or $ sign.</h3>
<p>By design: the symbol is in every column header and follows your Currency choice on Start
Here. Thousand separators follow your computer’s regional settings.</p></div>
<div class="nb"><h3>Can I use it on several projects?</h3>
<p>Yes — for your own home(s), keep a fresh copy of the BLANK file for each project.</p></div>
<div class="box warn nb"><b>Tip that pays for this file:</b> never let the % you have paid a
contractor run far ahead of the % of work done, and hold back the last payment until the punch
list is finished.</div>

<h2>Need help?</h2>
<p>If something is unclear or you think you found a problem, {support}. Please include which
tab and cell you are looking at — a screenshot helps.</p>

<h2>Licence</h2>
<p class="small">{NOTICE}. All rights reserved. The files are licensed,
not sold, to the original purchaser for personal, non-commercial use on their own home
renovation(s). You may not share, copy, resell, redistribute, sublicense, give away or upload the
files — or any modified version or part of them — to any website, marketplace, group or other
person. The formulas, structure, layout and design are protected by copyright and belong to Renovation OS;
removing the copyright notice or the sheet protection is not permitted. Digital product: all
sales are final once the files have been delivered. Provided “as is” as a planning aid without
warranty; it is not financial, legal or construction advice — always check figures before you
act on them.</p>
</body></html>"""


def main():
    tmp = os.path.join(HERE, "_quickstart.html")
    with open(tmp, "w") as f:
        f.write(HTML)
    with sync_playwright() as p:
        b = p.chromium.launch(executable_path=CHROME)
        pg = b.new_page()
        pg.goto(f"file://{tmp}")
        pg.wait_for_timeout(500)
        pg.pdf(path=OUT, format="A4", print_background=True, display_header_footer=True,
               header_template="<span></span>",
               footer_template=(
                   "<div style='font-size:7pt;color:#8A9199;width:100%;padding:0 16mm;"
                   "display:flex;justify-content:space-between;font-family:Arial'>"
                   f"<span>Quick Start · {NOTICE}</span>"
                   "<span><span class='pageNumber'></span> / <span class='totalPages'></span>"
                   "</span></div>"),
               margin={"top": "16mm", "bottom": "18mm", "left": "16mm", "right": "16mm"})
        b.close()
    os.remove(tmp)
    print("wrote", OUT)


if __name__ == "__main__":
    main()
