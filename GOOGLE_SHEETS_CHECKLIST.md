# Google Sheets — manual verification checklist

Automated tests recalculate every file in LibreOffice and check for errors and expected values. The following behaviours are specific to Google Sheets and **must be checked by hand** after uploading. Do this once per release, on the EUR DEMO file first and then on the BLANK file.

**Setup:** Drive ▸ Upload `Renovation_OS_DEMO_EUR.xlsx` ▸ Open with Google Sheets ▸ File ▸ **Save as Google Sheets**. Test on the converted copy, not in the xlsx preview.

## 1. Formulas and recalculation
- [ ] No `#ERROR!`, `#NAME?`, `#REF!` or `#VALUE!` anywhere. Use Ctrl+F on each tab and search for `#`.
- [ ] **Start Here ▸ C33** shows `€` and **C34** shows `15 Jun 2026` (DEMO).
- [ ] Dashboard headline reads **"🔴 N ACTIONS NEED YOUR ATTENTION"**, and N matches the LibreOffice test output.
- [ ] Budget ▸ *Best quote* for Kitchen, Flooring, Painting and Roofing is filled in. This column uses `SUMPRODUCT(MAX(...))`, so it confirms Sheets evaluates arrays inside SUMPRODUCT.
- [ ] Payments ▸ per-contractor summary ▸ *Next payment* dates are filled in (same technique).
- [ ] Clear Start Here ▸ C11 ("Today" override). Every tab should recalculate to today's date. Then press Undo.
- [ ] Change Start Here ▸ Currency to `GBP £`. Headers change to `(£)` on Budget, Payments and the Dashboard cards. Change it back.
- [ ] Change **Week starts on** to Sunday. This Week ▸ C4 and the Gantt week headers shift by one day.

## 2. Dropdowns (data validation)
- [ ] Budget ▸ Stage, Contractors ▸ Category/Room/Signed, Payments ▸ Paid to, Timeline ▸ Status, Selections ▸ Chosen (A/B/C), Issues ▸ Severity all show a ▼ list.
- [ ] Rename a category on Budget (e.g. `Other` → `Garage`). The new name appears in the Contractors ▸ Category dropdown.
- [ ] Rename a room card on Rooms. The new name appears in every Room dropdown.
- [ ] Payments ▸ *Paid to*: typing a supplier not on the list shows a **warning**, not a hard rejection. Sheets may show a small red triangle instead of a dialog; that is acceptable.
- [ ] Blank entries at the end of dynamic lists (Categories, Rooms, Contractors) are acceptable. Note if they look ugly.

## 3. Conditional formatting
- [ ] Dashboard health rows: 🟢/🟡/🔴 icons have green, yellow and red backgrounds, and the overall banner is red in DEMO.
- [ ] Budget ▸ Status column is coloured (Kitchen red).
- [ ] Quotes ▸ checklist cells: *Not included* red, *Unclear* yellow, *Included* green. Extra-cost cells turn grey when the item is Included. The winner's ranking cell is green.
- [ ] Quotes ▸ block 1 warning row is **red** ("⚠️ Hidden costs change the ranking!").
- [ ] Timeline Gantt: teal bars for planned work, grey for done, red for the late *Plumbing rough-in*, and a yellow column for the current week.
- [ ] Contractors ▸ Overpayment check is red for *Bright Spark Electric*, and *License ✓ = No* is red.
- [ ] Dashboard KPI "Over / under budget" is red and "Contingency used" is red (above 75%).
- [ ] **Known risk:** some Dashboard KPI rules use named ranges (e.g. `OpenHigh>0`, `OrdersLate>0`). If Sheets drops these rules, note which ones.

## 4. Charts
- [ ] Dashboard ▸ "Budget vs forecast — top 12 categories" (horizontal bars, 2 series) renders with category names.
- [ ] Dashboard ▸ "Monthly cash flow" (stacked columns) shows Mar 2026 – Feb 2027 labels and 2 series.
- [ ] Series colours are acceptable (teal/grey and teal/amber). Sheets may re-theme them; note if they are unreadable.
- [ ] Charts read data from the hidden **Engine** sheet and from Payments. Confirm the charts still render with Engine hidden.

## 5. Protection
- [ ] Sheets converts xlsx sheet protection into *protected ranges/sheets*. Check Data ▸ Protect sheets and ranges: each tab is listed with the yellow cells as exceptions.
- [ ] As a **second Google account with edit access** (the owner is never blocked), try typing in a white formula cell. You should get a warning or be blocked. Typing in a yellow cell should work.
- [ ] Engine and Lists tabs are hidden (View ▸ Hidden sheets shows them).

## 6. Emoji and fonts
- [ ] 🟢 🟡 🔴 ⚠️ 🏆 ✓ ✗ █ ░ 🏠 📅 render in the titles, status cells, progress bars (Dashboard contingency, Rooms cards) and alert list.
- [ ] Rooms cards look clean: the bars are aligned and the contractor names are not cut off badly.
- [ ] Font is Arial everywhere. Row heights look generous, not cramped.

## 7. Links and dates
- [ ] Start Here ▸ "YOUR TABS" links jump to the right tab. Internal links sometimes break when converting to Sheets; if so, note it. The links are only a convenience.
- [ ] Vault ▸ *Open ↗* links open the example URLs in a new tab.
- [ ] Dates display as `15 Jun 2026` (DD MMM YYYY). Month names follow the spreadsheet locale (File ▸ Settings).
- [ ] Switch File ▸ Settings ▸ Locale to Germany. Formulas still work, separators become `12.500`, and the date format stays readable.

## 8. Excel spot-check (Excel 2016 or newer, Windows or Mac)
- [ ] The file opens without a "repair" or "unreadable content" prompt.
- [ ] No `@` inserted into formulas (Microsoft 365). Check Budget ▸ E8 and Payments ▸ AE6.
- [ ] Emoji show in black and white on Excel 2016 for Windows. That is expected and acceptable.
- [ ] Formulas ▸ Error Checking shows no circular references.
