# Renovation OS — Quick Start (1 page)

**Renovation OS** runs your whole renovation from one spreadsheet: the budget, quotes, contractors, payments, change orders, timeline, orders, issues and documents. Everything is calculated for you, and every Monday it tells you exactly what to do.

Works in **Microsoft Excel 2016+** and **Google Sheets**. No macros, no add-ons and no internet connection needed.

---

## Which file do I open?

| File | Use it to… |
|---|---|
| `Renovation_OS_DEMO_…xlsx` | Explore first. It holds a realistic ~135k renovation, frozen on 15 Jun 2026 so that every alert and chart shows something. |
| `Renovation_OS_BLANK_…xlsx` | Your own project. The structure and formulas are the same, and it is empty and ready to fill in. |

*Google Sheets:* go to Drive ▸ New ▸ File upload, then right-click the file ▸ Open with ▸ Google Sheets ▸ File ▸ Save as Google Sheets.

## The only rule: yellow = you type

- 🟨 **Light-yellow cells**: you type here, or pick from the ▼ dropdown.
- ⬜ **White and grey cells**: automatic. They are protected so you can't break a formula by accident.
- 🟢 on track · 🟡 needs attention soon · 🔴 urgent

## Get started in 10 minutes

1. **Start Here.** Enter the project name, start date, target move-in date and total budget (incl. tax). Pick your currency, tax name, area unit and the day your week starts. The contingency (10% by default) is set aside automatically. What's left is your *working budget*.
2. **Budget.** Give each category a budget. Rename or add categories if you like. Try to make the total match your working budget; the check at the top tells you if it doesn't.
3. **Contractors + Quotes.** Add each contractor. On **Quotes**, put up to 4 offers side by side and fill in the hidden-cost checklist. The **TRUE COST** is often not the cheapest sticker price.
4. **Timeline.** Adjust the pre-filled task list: set durations and the *predecessor ID* (the task that must finish first; it must be listed above). Dates and the Gantt chart update on their own.
5. **Every Monday:** open **Dashboard** and **This Week** and work through the ranked actions. **Every Friday:** fill in the Friday Review boxes. That's the whole routine.

## What each tab does

| Tab | In one line |
|---|---|
| Dashboard | KPIs, the 🟢🟡🔴 health scoreboard, your top 10 next actions, and charts. |
| This Week | Monday plan and Friday review. |
| Budget | Budget, committed, invoiced and paid per category, plus the forecast. |
| Quotes | Side-by-side comparison with hidden costs and weighted scores. 🏆 marks the winner. |
| Contractors | Contracts, insurance and licence checks, a **paying-ahead-of-work** warning, and a communication log. |
| Payments | Payment schedule, overdue invoices, and a 12-month cash flow. |
| Change Orders | Preview the effect on your budget and move-in date **before** you approve a change. |
| Timeline | Dependencies, delays, late tasks, and a 52-week Gantt chart. |
| Rooms | One card per room: money, progress, decisions and issues. |
| Selections & Orders | Decide-by and order-by dates for tiles, taps, appliances and so on. |
| Issues & Punch List | Problems found on site and snags to fix, with owners and deadlines. |
| Vault | Links to your contracts, invoices and photos, plus ✓ checks for missing documents. |
| Warranty & Maintenance | Warranty expiry dates and recurring home maintenance. |

## Good to know

- **"Today" override** (Start Here): leave it empty. The DEMO file has a date there so it looks the same whenever you open it. Clear that cell to see the demo "live".
- **Why is a budget line red?** The forecast is above its budget. While a line is *Estimating*, the forecast takes the highest of budget, best quote, committed and invoiced, so it plans for the worst case until you sign.
- **Paying ahead of work:** update "% work complete" on Contractors every week. If the % you have paid runs ahead by more than your tolerance (10%), you'll get a ⚠️.
- **Changing a formula:** the sheets are protected **without a password**. In Excel use Review ▸ Unprotect Sheet; in Google Sheets use Data ▸ Protect sheets and ranges.
- **Adding rows:** every list has plenty of pre-formatted rows (150–300). Use the next empty yellow row. Don't insert rows in the middle.
- **Amounts** are shown without a currency symbol in the cells. The symbol appears in the column headers and follows your Currency choice. Thousands separators follow your computer's regional settings.
- **USD edition:** the tax name defaults to "Sales tax" at 8%. Set your local rate on Start Here, or 0% if you enter every price including tax.
