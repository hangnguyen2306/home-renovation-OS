# Excel — manual release checklist

The automated tests (`python3 test_workbook.py`) recalculate every file in LibreOffice. They check for formula errors and circular references, and compare the demo numbers with an independent Python model. They can't see how the file looks and behaves in **real Microsoft Excel**, so check that by hand once per release. Do it on **Windows**, and on **Mac** if you can.

Files: `dist/Renovation_OS_DEMO.xlsx`, `…_BLANK.xlsx`, `…_LITE_DEMO.xlsx`, `…_LITE_BLANK.xlsx`.

## 1. Opening
- [ ] Opens with no "We found a problem with some content… repair?" message.
- [ ] After *Enable Editing*, **Formulas ▸ Error Checking** shows no circular references.
- [ ] Microsoft 365: no `@` has been added to formulas. Click a yellow input cell; formulas are hidden anyway, so this only matters if you unprotect with your password to inspect, e.g. Budget ▸ E8.

## 2. Numbers (DEMO)
- [ ] Start Here ▸ "Today (as used by all tabs)" = 15 Jun 2026, and the currency symbol is €.
- [ ] Dashboard headline reads **"🔴 33 ACTIONS NEED YOUR ATTENTION"** (Lite: 18). Forecast 131,840 (Lite: 126,090).
- [ ] No `#VALUE!`, `#REF!`, `#NAME?` or `#N/A` anywhere. Ctrl+F for `#` on each tab, with *Look in: Values*.
- [ ] Budget ▸ *Best quote* filled for Kitchen, Flooring, Painting and Roofing (Pro).

## 3. Protection & licence
- [ ] Clicking a white (formula) cell does nothing. Yellow cells can be selected and typed in.
- [ ] Right-click a sheet tab: *Unhide…* is greyed out (workbook structure locked).
- [ ] Review ▸ Unprotect Sheet asks for the password.
- [ ] Row 3 of every tab shows the © line. The licence block is at the bottom of Start Here. File ▸ Info ▸ Author = Hang Nguyen, and the © line reads "© 2026 Renovation OS — created by Hang Nguyen".
- [ ] Start Here tab links and the Contractors "💬 Communication log →" link jump correctly.

## 4. Dropdowns & behaviour
- [ ] Dropdowns work: Budget ▸ Stage, Contractors ▸ Category, Payments ▸ Paid to, Timeline ▸ Status, Selections ▸ Chosen.
- [ ] Rename a Budget category. The new name appears in the Contractors ▸ Category dropdown.
- [ ] Payments ▸ *Paid to*: typing an unlisted supplier gives a warning you can accept.
- [ ] In the BLANK file, clear nothing but set a project start date. The Timeline gets dates and the Gantt chart fills.
- [ ] Change Start Here ▸ Currency to GBP £. Headers switch to (£).

## 5. Look
- [ ] Only input cells are yellow. Attention highlights are orange, the Gantt current week is blue.
- [ ] Both Dashboard charts render, with category names and month labels.
- [ ] Emoji 🟢🟡🔴⚠️🏆 show. On Excel 2016 for Windows they may appear black-and-white, which is acceptable.
- [ ] Font is Arial and nothing is badly cut off at 100% zoom.

## 6. Mac
- [ ] Repeat 1–3 in Excel for Mac.
