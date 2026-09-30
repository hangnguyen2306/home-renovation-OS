#!/usr/bin/env python3
"""Render vertical Pinterest pins (1000×1500, 2:3) + pins.csv with titles/descriptions/keywords.

Uses the real screenshots in marketing/shots/ (run capture.py first).
Run from repo root:  python3 marketing/make_pins.py
"""
import csv
import os

from playwright.sync_api import sync_playwright

HERE = os.path.dirname(os.path.abspath(__file__))
SHOTS = os.path.join(HERE, "shots")
OUT = os.path.join(HERE, "pins")
CHROME = "/opt/pw-browsers/chromium-1194/chrome-linux/chrome"
LINK = "https://YOUR-GUMROAD-LINK"   # ← replace with your product URL before uploading
W, H = 1000, 1500

TEAL, TEAL_DARK, CREAM, INK, GREY, CORAL, AMBER = ("#0F5E5E", "#0A4242", "#F6F3EC", "#1F2937",
                                                    "#5B6472", "#D9463B", "#F59E0B")
CSS = f"""
*{{box-sizing:border-box;margin:0;padding:0}}
html,body{{width:{W}px;height:{H}px}}
body{{font-family:'Liberation Sans','DejaVu Sans',Arial,sans-serif;background:{CREAM};color:{INK};
     position:relative;overflow:hidden}}
body.dark{{background:{TEAL_DARK};color:#fff}}
.brand{{position:absolute;top:50px;left:64px;font-size:24px;font-weight:700;letter-spacing:5px;
       color:{TEAL}}} .dark .brand{{color:#9FD6D6}} .brand span{{color:{AMBER}}}
.kicker{{position:absolute;top:104px;left:64px;font-size:26px;font-weight:700;color:#fff;
        background:{CORAL};padding:8px 20px;border-radius:30px}}
h1{{position:absolute;left:64px;right:64px;font-size:74px;line-height:1.04;font-weight:700;
   color:{TEAL_DARK};letter-spacing:-1.5px}} .dark h1{{color:#fff}}
h1 em{{font-style:normal;color:{CORAL}}} .dark h1 em{{color:{AMBER}}}
.sub{{position:absolute;left:66px;right:70px;font-size:32px;line-height:1.3;color:{GREY}}}
.dark .sub{{color:#CFE3E3}}
.card{{position:absolute;left:50px;right:50px;background:#fff;border-radius:22px;overflow:hidden;
      box-shadow:0 24px 60px rgba(10,50,50,.25)}}
.chrome{{height:46px;background:#EEF2F2;display:flex;align-items:center;padding:0 20px;
        border-bottom:1px solid #DDE4E4}}
.dot{{width:14px;height:14px;border-radius:50%;margin-right:9px}}
.tab{{margin-left:16px;font-size:20px;font-weight:700;color:{TEAL}}}
.shot{{display:flex;justify-content:center;align-items:flex-start;background:#fff;padding-top:6px}}
.shot img{{display:block}}
.tips{{position:absolute;left:64px;right:64px;list-style:none}}
.tips li{{font-size:36px;line-height:1.25;padding:22px 0 22px 84px;position:relative;
         border-bottom:2px solid rgba(15,94,94,.12)}}
.dark .tips li{{border-color:rgba(255,255,255,.15)}}
.tips li b{{color:{TEAL}}} .dark .tips li b{{color:{AMBER}}}
.tips li::before{{content:attr(data-n);position:absolute;left:0;top:18px;width:58px;height:58px;
                 border-radius:50%;background:{TEAL};color:#fff;font-weight:700;font-size:30px;
                 text-align:center;line-height:58px}}
.dark .tips li::before{{background:{AMBER};color:{TEAL_DARK}}}
.pill{{position:absolute;font-size:30px;font-weight:700;color:#fff;padding:18px 26px;border-radius:18px;
      background:{CORAL};box-shadow:0 14px 30px rgba(0,0,0,.25);line-height:1.2}}
.pill small{{display:block;font-size:22px;font-weight:400;margin-top:4px;opacity:.92}}
.cta{{position:absolute;left:0;right:0;bottom:0;height:130px;background:{TEAL};color:#fff;
     display:flex;align-items:center;justify-content:space-between;padding:0 64px}}
.dark .cta{{background:{AMBER};color:{TEAL_DARK}}}
.cta b{{font-size:34px}} .cta span{{font-size:26px;opacity:.9}}
.big{{position:absolute;left:64px;right:64px;font-size:190px;font-weight:700;line-height:1;
     color:{CORAL}}}
"""


def img(name, width=None, height=None):
    st = ";".join(s for s in (f"width:{width}px" if width else "",
                              f"height:{height}px" if height else "") if s)
    return f'<img src="file://{os.path.join(SHOTS, name)}.png" style="{st}">'


def card(tab, inner, top, height):
    dots = "".join(f'<div class="dot" style="background:{c}"></div>'
                   for c in ("#F87171", "#FBBF24", "#34D399"))
    return (f'<div class="card" style="top:{top}px;height:{height}px"><div class="chrome">{dots}'
            f'<div class="tab">{tab}</div></div><div class="shot">{inner}</div></div>')


def tips(items, top):
    li = "".join(f'<li data-n="{i}">{t}</li>' for i, t in enumerate(items, 1))
    return f'<ul class="tips" style="top:{top}px">{li}</ul>'


def page(body, dark=False, kicker=None, cta="Renovation OS · Excel planner", cta2="Get it →"):
    k = f'<div class="kicker">{kicker}</div>' if kicker else ""
    return (f"<html><head><meta charset='utf-8'><style>{CSS}</style></head>"
            f"<body class='{'dark' if dark else ''}'><div class='brand'>RENOVATION <span>OS</span></div>"
            f"{k}{body}<div class='cta'><b>{cta}</b><span>{cta2}</span></div></body></html>")


def h1(text, top=170):
    return f'<h1 style="top:{top}px">{text}</h1>'


def sub(text, top):
    return f'<div class="sub" style="top:{top}px">{text}</div>'


# (file, html, title, description, board, keywords, alt)
PINS = [
    ("01_budget_template",
     page(h1("Renovation budget template <em>for Excel</em>")
          + sub("Budget, contractors, payments and timeline in one dashboard.", 430)
          + card("Dashboard", img("dash_top", width=880), 560, 520)
          + '<div class="pill" style="right:40px;top:1120px">🔴 Tells you what to do first</div>'),
     "Renovation Budget Template for Excel — Home Renovation Planner",
     "Keep your renovation budget, contractors, payments and timeline in one Excel dashboard. See your forecast final cost, contingency used and a ranked list of what needs your attention this week. Made for homeowners managing their own renovation.",
     "Renovation Planning", "renovation budget template, renovation planner, excel template, home renovation budget, remodel budget spreadsheet",
     "Excel renovation dashboard showing working budget, forecast final cost, contingency used and a red action-required health score"),
    ("02_hidden_costs_list",
     page(h1("5 costs <em>missing</em> from your renovation quote", 160)
          + tips(["<b>Waste disposal</b> — skips add up fast", "<b>Delivery</b> of materials",
                  "<b>Finishing</b> — skirting, sealing, touch-ups", "<b>Electrical upgrades</b> hidden in “extras”",
                  "<b>Cleanup</b> after the job"], 450)
          + card("Quotes — hidden-cost checklist", img("quotes", width=880), 1010, 330),
          kicker="Before you sign"),
     "5 Hidden Costs Missing From Your Renovation Quote",
     "Before you sign a contractor quote, check these 5 costs that are often NOT included: waste disposal, delivery, finishing, electrical upgrades and cleanup. Renovation OS has a hidden-cost checklist that adds them up so you compare the TRUE cost of every quote.",
     "Renovation Tips", "renovation quote, contractor quote checklist, hidden renovation costs, renovation tips, remodel budget",
     "List of five hidden costs often missing from renovation quotes, with an Excel quote comparison checklist"),
    ("03_cheapest_quote",
     page(h1("Your cheapest quote could cost <em>€4,000 more</em>", 160)
          + sub("Compare up to 4 quotes side by side. Missing items are priced in — the TRUE cost wins.", 470)
          + card("Quotes", img("quotes", width=880), 620, 700)
          + '<div class="pill" style="left:40px;top:1210px">⚠️ Hidden costs change the ranking!</div>'),
     "Why the Cheapest Renovation Quote Can Cost More",
     "A quote that looks €2,500 cheaper can end up €1,500 more expensive once you add waste disposal, delivery and electrical work it doesn't include. Compare contractor quotes by TRUE cost with this Excel renovation planner.",
     "Renovation Tips", "compare contractor quotes, renovation quote comparison, kitchen remodel quote, renovation budget, excel template",
     "Excel comparison of three kitchen quotes showing the cheapest quote becomes more expensive after hidden costs"),
    ("04_pay_ahead",
     page(h1("Never pay a contractor <em>ahead of the work</em>", 160)
          + sub("Paid 60% for 35% of the work? You lose your leverage. Get warned automatically.", 470)
          + card("Contractors", img("con_names", height=440) + img("con_pay", height=440), 620, 520)
          + '<div class="pill" style="right:40px;top:1190px">⚠️ Paid 60% — only 35% done<small>flagged automatically</small></div>'),
     "Never Pay a Contractor Ahead of the Work",
     "The golden rule of managing a renovation: the % you have paid should never run far ahead of the % of work done. This Excel planner compares both for every contractor and flags “paying ahead of work” before it's too late.",
     "Renovation Tips", "contractor payment schedule, paying contractors, renovation tips, home renovation, contractor tracker",
     "Excel contractor list comparing percent of work complete with percent paid and a warning for paying ahead of work"),
    ("05_contingency",
     page(h1("How much <em>contingency</em> does a renovation need?", 160)
          + '<div class="big" style="top:500px">10–15%</div>'
          + sub("of your total budget, kept aside for surprises: rotten joists, asbestos, pipe re-routes…", 720)
          + tips(["Budget = what you plan to spend", "Contingency = kept aside, not for upgrades",
                  "Track % used — act when it passes 75%"], 900), dark=True),
     "How Much Contingency Does a Renovation Need?",
     "Plan a contingency of 10–15% of your total renovation budget for surprises like rotten joists, asbestos or rerouted pipes — and track how much of it you've used. Renovation OS shows contingency used on your dashboard and warns you above 75%.",
     "Renovation Budget", "renovation contingency, renovation budget tips, home renovation budget, remodel budget, renovation planning",
     "Tip graphic: renovations need 10 to 15 percent contingency, with three rules for using it"),
    ("06_gantt",
     page(h1("Renovation timeline <em>with Gantt chart</em>")
          + sub("Tasks, dependencies and delays — dates update themselves.", 430)
          + card("Timeline", img("tl_tasks", height=620) + img("tl_gantt", height=620), 560, 690)),
     "Renovation Timeline Template with Gantt Chart (Excel)",
     "Plan your renovation step by step: design, permits, demolition, rough-in, inspections, finishes and move-in. Set durations and dependencies once; when a task slips, every date after it moves — including your move-in date.",
     "Renovation Planning", "renovation timeline, gantt chart excel, renovation schedule, remodel planner, home renovation plan",
     "Excel renovation timeline with task list and weekly Gantt chart, late task in red"),
    ("07_delay",
     page(h1("One delay moved our move-in date <em>11 days</em>", 160)
          + sub("Plumbing ran late. Everything that depended on it moved too — and the dashboard showed it the same day.", 470)
          + card("Timeline", img("tl_tasks", height=620) + img("tl_gantt", height=620), 650, 690)),
     "How One Renovation Delay Moves Your Move-In Date",
     "In a renovation, tasks depend on each other: no drywall before the inspection, no inspection before plumbing. See how one late task pushes the whole schedule — and your move-in date — with an Excel renovation timeline.",
     "Renovation Planning", "renovation delay, renovation timeline, move in date, renovation schedule, gantt chart",
     "Excel Gantt chart where a late plumbing task pushes later tasks and the move-in date"),
    ("08_change_order",
     page(h1("Check this before approving <em>any change</em>", 160)
          + tips(["New <b>total cost</b> if you say yes", "<b>Contingency</b> left afterwards",
                  "New <b>move-in date</b>"], 430)
          + card("Change Orders — impact preview", img("co_preview", width=880), 850, 250)
          + '<div class="pill" style="right:40px;top:1140px">Contingency left: −240</div>'),
     "3 Things to Check Before Approving a Renovation Change Order",
     "Every “small extra” during a renovation changes your budget and timeline. Before you approve a change order, check the new total cost, how much contingency is left, and the new move-in date. Renovation OS previews all three automatically.",
     "Renovation Tips", "change order, renovation extras, renovation budget, contractor change request, home renovation tips",
     "Checklist of three things to check before approving a change order, with an Excel impact preview"),
    ("09_order_by",
     page(h1("When to order tiles, taps &amp; appliances", 160)
          + sub("Order-by = install date − supplier lead time − safety buffer. Decide a week before that.", 400)
          + card("Selections & Orders", img("sel_dates", width=880), 560, 520)
          + '<div class="pill" style="left:40px;top:1120px">⚠️ Order late — install at risk</div>'),
     "When to Order Tiles, Taps and Appliances for Your Renovation",
     "Late materials are one of the most common renovation delays. Work backwards: install date minus supplier lead time minus a safety buffer = your order-by date. Decide a week earlier. This Excel planner calculates both dates for every item.",
     "Renovation Planning", "renovation materials, order tiles, kitchen appliances, renovation checklist, remodel planning",
     "Excel table of materials with needed-on, order-by and decide-by dates and status warnings"),
    ("10_monday",
     page(h1("My 10-minute <em>Monday</em> renovation routine", 160)
          + tips(["Open the dashboard", "Do the ranked actions, top to bottom",
                  "Update payments &amp; % work done"], 450)
          + card("Dashboard — your next actions", img("dash_actions", width=880), 830, 470)),
     "10-Minute Weekly Routine to Stay in Control of Your Renovation",
     "Managing your own renovation doesn't have to take your evenings. Once a week: open the dashboard, work through the ranked next actions (overdue payments, late tasks, decisions), and update what's been paid and done.",
     "Renovation Planning", "renovation routine, project management home, renovation organization, renovation planner, home renovation",
     "Weekly routine with an Excel dashboard listing ranked renovation actions"),
    ("11_cash_flow",
     page(h1("Know what you'll pay <em>each month</em>")
          + sub("Paid and still-to-pay per month — plus cash needed in the next 30 days.", 430)
          + card("Payments — cash flow", img("payments_cash", width=700), 560, 700)),
     "Renovation Cash Flow Planner: What You'll Pay Each Month",
     "Avoid cash surprises during your renovation. See what you've paid and what's still due each month, and how much cash you need in the next 30 days — calculated from your contractor payment schedule.",
     "Renovation Budget", "renovation cash flow, renovation payments, renovation budget, payment schedule, excel budget template",
     "Excel monthly cash flow table for a renovation with paid and still-to-pay amounts"),
    ("12_rooms",
     page(h1("Every room <em>at a glance</em>")
          + sub("Budget used, progress, decisions and issues — per room.", 430)
          + card("Rooms", img("rooms", width=880), 540, 610)),
     "Room-by-Room Renovation Tracker (Excel)",
     "Kitchen over budget? Bathroom waiting on decisions? Room cards show budget used, progress, open issues and contractors for every room of your renovation.",
     "Renovation Planning", "kitchen remodel tracker, bathroom renovation, room by room renovation, renovation planner, home renovation",
     "Excel room cards for kitchen and bathrooms showing budget used and progress"),
    ("13_punch_list",
     page(h1("Don't make the <em>final payment</em> until…", 160)
          + tips(["Every <b>punch-list</b> item is fixed", "Final inspection is <b>passed</b>",
                  "You have the <b>warranty</b> in writing", "Invoices match the <b>contract + approved changes</b>"],
                 460), dark=True),
     "Don't Make the Final Renovation Payment Until…",
     "Your last payment is your leverage. Hold it back until the punch list is done, the final inspection is passed, you have the warranty in writing and the invoices match the contract plus approved changes.",
     "Renovation Tips", "punch list, final payment contractor, renovation tips, renovation checklist, snag list",
     "Checklist of four things to confirm before making the final contractor payment"),
    ("14_documents",
     page(h1("Documents to keep during a renovation", 160)
          + tips(["<b>Contract</b> + quote for every contractor", "<b>Insurance</b> certificates",
                  "<b>Permits</b> and plans", "<b>Invoices</b> and receipts", "<b>Photos</b> before, during, after"],
                 430)),
     "Documents You Should Keep During a Renovation",
     "Keep these for every renovation: contract and quote per contractor, insurance certificates, permits and plans, invoices and receipts, and photos at every stage (especially behind the walls before they are closed).",
     "Renovation Tips", "renovation documents, renovation checklist, home renovation tips, contractor insurance, renovation organization",
     "Checklist of documents to keep during a home renovation"),
    ("15_forecast",
     page(h1("Know your final cost <em>before the final invoice</em>", 160)
          + card("Dashboard — charts", img("dash_charts", width=880), 520, 420)
          + tips(["Budget vs forecast per category", "Contingency used"], 1000)),
     "Know Your Renovation's Final Cost Before the Final Invoice",
     "See forecast vs budget for every category of your renovation — kitchen, bathrooms, electrical, flooring — and know your final cost months before the last invoice.",
     "Renovation Budget", "renovation budget tracker, renovation cost, remodel budget, budget vs actual, excel chart",
     "Excel bar chart of budget versus forecast by renovation category and a monthly cash-flow chart"),
    ("16_question_over_budget",
     page('<h1 style="top:260px;font-size:96px">Is your renovation <em>over budget</em> — and you don’t know it yet?</h1>'
          + sub("Most homeowners find out at the last invoice. A forecast shows it months earlier.", 820)
          + sub("Renovation OS · the Excel planner that tells you what to do next.", 1000), dark=True),
     "Is Your Renovation Over Budget Without You Knowing?",
     "Most homeowners only discover the real cost at the final invoice. A simple forecast — committed contracts, approved extras and open risks — shows it months earlier, while you can still act.",
     "Renovation Budget", "renovation over budget, renovation budget, remodel costs, home renovation, renovation planner",
     "Question graphic asking whether your renovation is over budget"),
    ("17_start_10_minutes",
     page(h1("Set up your renovation planner <em>in 10 minutes</em>", 160)
          + tips(["Dates, budget, currency", "A budget per category", "Your contractors",
                  "Payment milestones", "Timeline durations"], 450)
          + '<div class="sub" style="top:1180px;font-weight:700;color:#0F5E5E">Then 10 minutes a week.</div>'),
     "Set Up Your Renovation Planner in 10 Minutes",
     "Five steps: project dates and budget, a budget per category, your contractors, payment milestones and task durations. After that, 10 minutes a week keeps your whole renovation under control.",
     "Renovation Planning", "renovation planner, renovation spreadsheet, excel planner, home renovation, project planner",
     "Five-step setup list for an Excel renovation planner"),
    ("18_lite",
     page(h1("A <em>simple</em> renovation budget tracker", 160)
          + sub("Budget, contractors, payments, timeline and issues — without the extras.", 430)
          + card("Dashboard — Lite", img("lite_dash", width=880), 560, 610),
          cta="Renovation OS Lite · Excel"),
     "Simple Renovation Budget Tracker for Excel (Lite)",
     "Just the essentials: budget and forecast, contractor payments with a paying-ahead warning, a timeline with Gantt chart and an issues list — on one Excel dashboard.",
     "Renovation Budget", "simple budget tracker, renovation budget spreadsheet, excel budget template, diy renovation, remodel budget",
     "Simple Excel renovation dashboard with budget, payments and health score"),
]


def main():
    os.makedirs(OUT, exist_ok=True)
    with sync_playwright() as p:
        b = p.chromium.launch(executable_path=CHROME)
        pg = b.new_page(viewport={"width": W, "height": H})
        for name, html, *_ in PINS:
            tmp = os.path.join(OUT, name + ".html")
            open(tmp, "w").write(html)
            pg.goto(f"file://{tmp}")
            pg.wait_for_timeout(250)
            pg.screenshot(path=os.path.join(OUT, name + ".png"))
            os.remove(tmp)
        b.close()
    with open(os.path.join(OUT, "pins.csv"), "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["Image", "Title", "Description", "Link", "Board", "Keywords", "Alt text",
                    "Suggested day"])
        for i, (name, _h, title, desc, board, kw, alt) in enumerate(PINS, 1):
            assert len(title) <= 100 and len(desc) <= 500 and len(alt) <= 500, name
            w.writerow([name + ".png", title, desc, LINK, board, kw, alt, f"Day {i}"])
    print(f"{len(PINS)} pins → {OUT}")


if __name__ == "__main__":
    main()
