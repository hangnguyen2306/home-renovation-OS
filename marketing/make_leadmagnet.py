#!/usr/bin/env python3
"""Build the free lead magnet: 'The Renovation Money Checklist' (A4 PDF + cover PNG).

Given away in exchange for an email address (waitlist). No selling inside, just value and a
note that Renovation OS is launching. Run from repo root:  python3 marketing/make_leadmagnet.py
"""
import os

from playwright.sync_api import sync_playwright

YEAR, AUTHOR = 2026, "Hang Nguyen"
NOTICE = f"© {YEAR} Renovation OS — created by {AUTHOR}"
HERE = os.path.dirname(os.path.abspath(__file__))
SHOTS = os.path.join(HERE, "shots")
CHROME = "/opt/pw-browsers/chromium-1194/chrome-linux/chrome"
OUT_PDF = os.path.join(HERE, "Renovation_Money_Checklist.pdf")
OUT_COVER = os.path.join(HERE, "leadmagnet_cover.png")

HIDDEN = [
    ("Demolition &amp; strip-out", "Is removing the old kitchen / tiles / floors included?"),
    ("Waste disposal &amp; skips", "Who pays for skips and tip fees — and how many are included?"),
    ("Delivery, crane or lift", "Are delivery, carrying upstairs and crane hire in the price?"),
    ("Permits &amp; fees", "Who applies for permits and pays the fees?"),
    ("Architect / engineer", "Do you need structural calculations or drawings — and who pays?"),
    ("Protection &amp; daily cleaning", "Are floors, stairs and furniture protected? Final clean included?"),
    ("Electrical upgrades", "Does the new kitchen or heat pump need a bigger fuse box or new circuits?"),
    ("Plumbing re-routing", "What if pipes must be moved? Fixed price or hourly?"),
    ("Unexpected finds", "What happens with asbestos, rot or damp — rates agreed up front?"),
    ("Finishing", "Skirting, sealing, silicone, painting touch-ups — included or “extra”?"),
    ("Fixtures &amp; fittings", "Taps, sockets, switches, handles: which models, and who supplies them?"),
    ("Tax / VAT", "Is the price including tax? Which rate applies to your job?"),
    ("Temporary living", "Kitchen or bathroom out of use: temporary setup, storage, moving costs?"),
    ("Price increases", "Are material prices fixed, and for how long is the quote valid?"),
    ("Contingency", "Have YOU kept 10–15% aside for all of the above?"),
]

DOCS = ["Signed contract per contractor (scope, price, dates, payment schedule)",
        "The quote the contract is based on",
        "Liability insurance certificate of every contractor",
        "Permits, approvals and plans / drawings",
        "Every change order — in writing, with price and extra days",
        "All invoices and proof of payment",
        "Receipts for materials you bought yourself",
        "Inspection reports (electrical, gas, structural…)",
        "Warranties and manuals of installed equipment",
        "Photos at every stage — especially behind walls before they are closed"]

FINAL = ["Every punch-list item is fixed (walk through room by room)",
         "Final inspection passed and certificates received",
         "Warranty and manuals handed over in writing",
         "Invoices match contract + approved change orders — nothing more",
         "Site cleared and cleaned, waste removed",
         "Keys, codes and spare materials handed over"]


def box(items):
    return "".join(f'<li><span class="cb"></span>{t}</li>' for t in items)


rows = "".join(f'<tr><td class="c"><span class="cb"></span></td><td><b>{n}</b></td><td>{q}</td></tr>'
               for n, q in HIDDEN)

HTML = f"""<!doctype html><html><head><meta charset="utf-8"><style>
@page {{ size: A4; margin: 16mm; }}
*{{box-sizing:border-box}}
body{{font-family:'Liberation Sans','DejaVu Sans',Arial,sans-serif;color:#1F2937;font-size:10.5pt;
     line-height:1.45;margin:0}}
.brand{{font-weight:700;letter-spacing:3pt;color:#0F5E5E;font-size:9.5pt}} .brand span{{color:#F59E0B}}
h1{{font-size:34pt;color:#0A4242;line-height:1.05;margin:10mm 0 4mm;letter-spacing:-.5pt}}
h1 em{{font-style:normal;color:#D9463B}}
h2{{font-size:16pt;color:#0F5E5E;margin:8mm 0 3mm;padding-bottom:1.5mm;border-bottom:1.2pt solid #0F5E5E}}
.lead{{font-size:13pt;color:#5B6472}}
.cover{{page-break-after:always;position:relative;min-height:250mm}}
.pillars{{display:flex;gap:4mm;margin:10mm 0}}
.pillars div{{flex:1;background:#E3F1F1;border-radius:3mm;padding:5mm;text-align:center}}
.pillars b{{display:block;font-size:26pt;color:#0F5E5E}}
table{{width:100%;border-collapse:collapse;font-size:9.8pt}}
th{{background:#0F5E5E;color:#fff;text-align:left;padding:2mm 3mm}}
td{{padding:2mm 3mm;border-bottom:.5pt solid #E1E7E7;vertical-align:top}}
td.c{{width:8mm}}
.cb{{display:inline-block;width:4mm;height:4mm;border:1pt solid #0F5E5E;border-radius:.8mm;
    margin-right:2.5mm;vertical-align:-.6mm}}
ul.check{{list-style:none;padding:0;margin:0}} ul.check li{{padding:1.8mm 0;border-bottom:.5pt solid #E1E7E7}}
.box{{background:#E3F1F1;border-radius:3mm;padding:4mm 5mm;margin:3mm 0}}
.warn{{background:#FFEDD5}}
.big{{font-size:28pt;font-weight:700;color:#D9463B}}
.two{{display:flex;gap:6mm}} .two>div{{flex:1}}
.nb{{page-break-inside:avoid}}
.formula{{font-size:12.5pt;font-weight:700;color:#0A4242;background:#fff;border:1pt solid #0F5E5E;
         border-radius:2mm;padding:3mm 4mm;margin:2mm 0;text-align:center}}
.shot{{width:100%;border:.6pt solid #D5DEDE;border-radius:3mm}}
.small{{font-size:8.5pt;color:#6B7280}}
</style></head><body>

<section class="cover">
  <div class="brand">RENOVATION <span>OS</span></div>
  <h1>The Renovation<br><em>Money</em> Checklist</h1>
  <p class="lead">15 hidden costs to check in every quote, the payment rules that protect you,
     and the documents you must keep — on 5 printable pages.</p>
  <div class="pillars">
    <div><b>15</b>hidden costs</div><div><b>10–15%</b>contingency</div><div><b>10</b>documents</div>
  </div>
  <div class="box"><b>How to use it:</b> print it, take it to every contractor meeting, and tick the
  boxes. If an item is not clearly in the quote, ask the question next to it — and get the answer
  in writing.</div>
  <p class="small" style="position:absolute;bottom:0">{NOTICE}. Free to use
  for your own renovation. Please don’t resell or republish it — share the sign-up link instead.
  This checklist is general guidance, not financial, legal or construction advice.</p>
</section>

<h2>1 · 15 hidden costs to check in every quote</h2>
<p>Most budget overruns don’t come from one big disaster — they come from small items nobody put in
the quote. Compare quotes only after you have asked all 15 questions.</p>
<table><tr><th></th><th style="width:33%">Cost</th><th>Ask the contractor</th></tr>{rows}</table>
<div class="box warn nb"><b>Rule:</b> for every “not included” or “unclear” answer, write down your
own estimate and add it to that quote. The quote with the lowest <i>sticker</i> price is often not the
one with the lowest <i>true</i> cost.</div>

<div class="nb"><h2>2 · How much contingency?</h2>
<div class="two">
  <div><div class="big">10–15%</div>
    <p>of your total budget, kept aside for surprises. Older house, structural work or unknown
    wiring? Plan closer to 15–20%.</p></div>
  <div><ul class="check">
    <li><span class="cb"></span>Working budget = total budget − contingency</li>
    <li><span class="cb"></span>Contingency is for surprises, not for upgrades</li>
    <li><span class="cb"></span>Check how much you have used every week</li>
    <li><span class="cb"></span>Above 75% used? Stop all optional changes</li>
  </ul></div>
</div></div>

<div class="nb"><h2>3 · Pay for work done — never ahead of it</h2>
<div>
<p>Your payments are your leverage. Tie every payment to a visible milestone, and compare what you
have paid with how much of the work is really finished.</p>
<div class="formula">% paid to a contractor ≤ % of their work that is finished</div>
<ul class="check">
  <li><span class="cb"></span>Keep the deposit modest and linked to a clear start or material order</li>
  <li><span class="cb"></span>Every payment is linked to a milestone you can check on site</li>
  <li><span class="cb"></span>Update “% work done” per contractor every week</li>
  <li><span class="cb"></span>Hold back the final payment (e.g. 5–10%) until the list below is complete</li>
</ul></div></div>

<h3 style="color:#0A4242;margin:5mm 0 2mm">Before the final payment</h3>
<ul class="check nb">{box(FINAL)}</ul>

<div class="nb"><h2>4 · 10 documents to keep</h2>
<ul class="check">{box(DOCS)}</ul>
<p class="small">Tip: one folder per contractor in your cloud drive, and one “Photos” folder per room
with sub-folders Before · Demolition · Rough-in · Walls · Finishing · Final.</p></div>

<div class="nb"><h2>5 · Order materials on time</h2>
<div>
<p>Late tiles, taps or appliances are one of the most common causes of delay — the installer is
ready, the material isn’t. Work backwards from the day it will be installed:</p>
<div class="formula">Order-by date = install date − supplier lead time − safety buffer (± 5 days)</div>
<div class="formula">Decide-by date = order-by date − 1 week</div>
<p><b>Example:</b> kitchen installed on 2 July, worktop lead time 3 weeks → order by 6 June →
decide by 30 May.</p></div></div>

<div class="nb"><h2>6 · The 10-minute Monday routine</h2>
<ul class="check">
  <li><span class="cb"></span>What is overdue? (payments, decisions, orders, follow-ups)</li>
  <li><span class="cb"></span>What must be decided or ordered this week?</li>
  <li><span class="cb"></span>Which payments are due — and is the work really done?</li>
  <li><span class="cb"></span>Is any task late that other tasks depend on?</li>
  <li><span class="cb"></span>Forecast final cost vs budget — still OK?</li>
</ul></div>

<div class="box nb" style="margin-top:8mm">
  <b>Coming soon: Renovation OS</b> — the Excel planner that does all of this for you. It checks
  hidden costs in quotes, warns you when you pay ahead of the work, calculates order-by dates and
  ranks your to-dos every Monday. You’re on the list — you’ll get the launch discount on launch day.
  {'<br>' + shot if (shot := f'<img class="shot" style="margin-top:3mm" src="file://{os.path.join(SHOTS, "dash_actions.png")}">') else ''}
</div>
</body></html>"""


def main():
    tmp = os.path.join(HERE, "_leadmagnet.html")
    with open(tmp, "w") as f:
        f.write(HTML)
    with sync_playwright() as p:
        b = p.chromium.launch(executable_path=CHROME)
        pg = b.new_page()
        pg.goto(f"file://{tmp}")
        pg.wait_for_timeout(400)
        pg.pdf(path=OUT_PDF, format="A4", print_background=True, display_header_footer=True,
               header_template="<span></span>",
               footer_template=(
                   "<div style='font-size:7pt;color:#8A9199;width:100%;padding:0 16mm;"
                   "display:flex;justify-content:space-between;font-family:Arial'>"
                   f"<span>The Renovation Money Checklist · {NOTICE}</span>"
                   "<span><span class='pageNumber'></span> / <span class='totalPages'></span></span></div>"),
               margin={"top": "16mm", "bottom": "18mm", "left": "16mm", "right": "16mm"})
        b.close()
    os.remove(tmp)
    import pymupdf
    pymupdf.open(OUT_PDF)[0].get_pixmap(dpi=130).save(OUT_COVER)
    print("wrote", OUT_PDF, "and", OUT_COVER)


if __name__ == "__main__":
    main()
