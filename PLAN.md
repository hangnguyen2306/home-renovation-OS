# Renovation OS — Build Plan & Cell Map

Status: **built** — all 4 workbooks generated, `test_workbook.py` passes (0 formula errors, demo values match the Python reference model). Deviations from the draft are noted inline as *Built:*.

---

## 0. Deliverables

| File | Purpose |
|---|---|
| `build_renovation_os.py` | openpyxl generator. `--currency EUR\|USD --variant DEMO\|BLANK\|ALL --out dist/` |
| `test_workbook.py` | LibreOffice headless recalculation (Python-UNO) → error scan + circular-ref scan + asserts vs. Python-computed expected values |
| `demo_data.py` | Demo dataset (single source of truth, used by the builder AND by the test's expected-value calculation) |
| `dist/Renovation_OS_{DEMO,BLANK}_{EUR,USD}.xlsx` | 4 outputs |
| `QUICKSTART.md`, `GOOGLE_SHEETS_CHECKLIST.md` | Buyer guide + your manual QA list |

---

## 1. Global conventions

**Layout on every sheet**
- Column A = 2-char gutter. Row 1 = title (Arial 18 bold, teal). Row 2 = "What this tab does" (Arial 10 italic grey). Row 3 spacer.
- Lists: header row 5 (teal fill, white bold text, frozen at row 5 / freeze pane `B6` or after ID column). Data starts row 6.
- Row height 20 for data rows, 28 for headers, 36 for titles.

**Colours** (tokens in one dict in the builder)
| Token | Hex | Use |
|---|---|---|
| ACCENT | `0F5E5E` (deep teal) | titles, headers, KPI numbers |
| ACCENT_LIGHT | `E3F1F1` | KPI card backgrounds, section bands |
| INPUT | `FFF7CC` (light yellow) | every unlocked cell |
| AUTO | `FFFFFF` / `F3F4F6` | formulas (locked) |
| GREY_TEXT | `6B7280` | descriptions, units |
| GOOD / WARN / BAD | `D1FAE5`/`065F46`, `FEF3C7`/`92400E`, `FEE2E2`/`991B1B` | conditional-format fills/fonts |
| Tab colours | teal `0F5E5E` (Start, Dashboard, This Week), green `2E7D32` (Budget, Quotes, Contractors, Payments, Change Orders), orange `E07B00` (Timeline, Rooms, Selections, Issues), purple `6A1B9A` (Vault, Warranty) |

**Formats**
- Dates: `DD MMM YYYY` (number format, never `TEXT()` — `TEXT` date codes are locale-dependent and break in German/Dutch/French Excel).
- Money: `#,##0;[Red]-#,##0` — **no currency symbol inside number formats**, because the currency is a user choice (see §1a) and a number format cannot change with a dropdown in Google Sheets. The symbol appears in column headers and KPI labels, built by formula, e.g. `="Budget ("&CurSym&")"`. Thousands/decimal separators follow the buyer's own locale automatically (12,500 vs 12.500). Percent `0%`.
- All sheets protected, no password; only INPUT cells unlocked. Protection options allow format columns/rows + sort off + autofilter off.

### 1a. User preferences (dropdowns on Start Here)
Everything "general" is a dropdown the buyer picks once; every sheet reads it through a named range. The EUR/USD build flag only sets the **defaults** of these dropdowns — both files can be switched to any option afterwards.

| Preference | Dropdown options | Default EUR / USD | What it changes |
|---|---|---|---|
| Currency | EUR €, USD $, GBP £, CHF, CAD $, AUD $, NZD $, SEK kr, NOK kr, DKK kr, PLN zł, CZK Kč, Other | EUR € / USD $ | `CurSym` in every money header, KPI label, chart title, alert text |
| Custom currency symbol | free text (only used when "Other") | — | `CurSym` |
| Tax name | VAT, Sales tax, GST, HST, TVA, MwSt, BTW, IVA, None | VAT / Sales tax | `TaxName` in labels: "Price incl. VAT?", "Total budget incl. Sales tax" |
| Default tax rate | % input | 21% / 0% | `VatRate` (quote normalization) |
| Prices usually quoted incl. tax? | Yes / No | Yes / Yes | default value pre-filled into each Quotes "incl. tax?" cell (still overridable per quote) |
| Area unit | m², ft² | m² / ft² | Rooms floor-area label + "Cost per m²/ft²" per room and on Dashboard (new input: floor area per room + total floor area on Start Here) |
| Week starts on | Monday, Sunday | Monday / Sunday | This Week default week-start, Gantt week columns, weekly counts: `AsOf - IF(WeekStart="Sunday", WEEKDAY(AsOf,1)-1, WEEKDAY(AsOf,3))` |
| Health "attention" band on budget lines | 0%, 5%, 10% | 5% | Budget 🟡 threshold per category |
| (existing) Alert window, decide buffer, order buffer, overpayment tolerance, contingency % | number inputs | as before | as before |

Not offered as a dropdown (and why):
- **Date format** — fixed `DD MMM YYYY` (unambiguous in every country; month names auto-translate to the buyer's locale). Switching date formats by dropdown would need conditional-format number formats, which Google Sheets ignores.
- **Symbol placement (€12,500 vs 12,500 €)** — no symbol in cells, so not needed.
- **Language** — English only in v1 (see open questions).

Lists sheet holds a Currency table (label → symbol) so `CurSym = IF(Currency="Other", CustomSym, INDEX(Lists!CurSymbols, MATCH(Currency, Lists!CurLabels, 0)))`.

**Global named ranges** (defined names — supported by Excel 2016 and Google Sheets import)
`CurSym, TaxName, DefInclTax, AreaUnit, WeekStart, BudgetWarnPct, TotalArea, ProjName, ProjStart, TargetEnd, TotalBudget, ContPct, ContAmt, WorkBudget, VatRate, AlertDays, DecideBuf, OrderBuf, OverpayTol, AsOf` (+ a few per-sheet totals used by Dashboard).

**`AsOf` ("today")** = `IF(Start_Here!AsOfOverride="",TODAY(),AsOfOverride)`. Every formula uses `AsOf`, never `TODAY()` directly.
- BLANK: override empty → live today.
- DEMO: override pinned to **15 Jun 2026** with a note *"Demo is frozen on this date so every feature shows data — clear it to use today's date."* This keeps the demo coherent whenever the buyer opens it, and makes the tests deterministic.

**Function whitelist** (Excel 2016 + Sheets): SUM, SUMIFS, COUNTIFS, COUNTIF, COUNTA, COUNT, IF, IFERROR, INDEX, MATCH, SMALL, LARGE, MIN, MAX, TODAY, DATE, YEAR, MONTH, DAY, WEEKDAY, AND, OR, NOT, ROUND, ABS, REPT, LEFT, LEN, ROW, SUMPRODUCT.
- Added vs. your list: `SUMPRODUCT, COUNTIF, COUNTA, COUNT, YEAR/MONTH/DAY, ROW, NOT, LEN` — all Excel 2007+ and native in Sheets.
- **Not** `MINIFS/MAXIFS` (Excel 2019+ only). Conditional minimum = `1/SUMPRODUCT(MAX((crit=X)*InvHelper))` with a helper column holding `1/value` (0 when blank).
- No array (CSE) formulas, no `TEXT()` on dates, no `@`, no tables, no whole-column refs.
- The builder will lint every formula string against the whitelist before saving (fails the build on an unknown function).

**Emoji**: 🟢🟡🔴⚠️🏆✓ produced by formulas as text. Conditional-format colours never depend on matching an emoji — each status has a hidden numeric level column (0/1/2) and CF rules key off that (emoji matching in CF is unreliable across Excel/Sheets).

---

## 2. Sheet order, sizes, cell map

Row capacities (pre-formatted, validated, formula-filled):

| List | Rows | Data range |
|---|---|---|
| Budget categories | 30 | Budget rows 8–37 |
| Quote blocks | 10 × 4 contractors | Quotes |
| Contractors | 30 | rows 6–35 |
| Communication log | 150 | rows 42–191 |
| Payments | 200 | rows 6–205 |
| Change orders | 100 | rows 7–106 |
| Timeline tasks | 150 | rows 7–156 |
| Rooms | 10 blocks | Rooms |
| Selections | 150 | rows 8–157 |
| Issues | 200 | rows 8–207 |
| Vault docs | 300 | rows 6–305 |
| Warranties / Maintenance | 60 / 40 | Warranty rows 7–66, maintenance rows 72–111 (*Built:* reduced from 100 so maintenance is not 100 rows down) |

### 2.1 START HERE (teal, gridlines off)
- B5:C12 **Your project** (yellow C): Project name, Address, Project start, Target move-in, Total budget incl. tax, Total floor area, "Today" override (blank).
- B14:C22 **Preferences** (yellow dropdowns, §1a): Currency, Custom symbol, Tax name, Default tax rate, Prices usually incl. tax?, Area unit, Week starts on, Budget attention band.
- B24:C29 **Alert settings**: Contingency % (10%), Alert window days (7), Decide buffer days (7), Order safety buffer days (5), Overpayment tolerance (10%).
- C31 Contingency amount `=TotalBudget*ContPct`, C32 Working budget `=TotalBudget-ContAmt`.
- E5:J17 **Get started in 10 minutes** — 5 numbered steps (Setup → Budget categories → Contractors & quotes → Timeline → Check Dashboard every Monday / This Week every Friday).
- E19:J22 **Colour legend**: yellow swatch "You type here", white "Automatic — don't type", 🟢🟡🔴 meanings.
- Tab map: one-line description of every tab (hyperlinks via `HYPERLINK("#'Budget'!A1",…)` — internal links only).

### 2.2 DASHBOARD (teal, gridlines off, no inputs)
- Row 1–2 title = `ProjName`, subtitle "as of" `AsOf`.
- Rows 4–12: **KPI cards** 5 per row × 3 rows (each card = 2 merged columns: label small grey, value large teal):
  Working budget · Committed · Invoiced · Paid · Remaining (= WorkBudget − Committed)
  Forecast final cost · Over/under budget (vs WorkBudget, red/green CF) · Contingency used % + `REPT("█",ROUND(pct*10,0))&REPT("░",10-…)` · Day X of Y · % time elapsed
  Forecast completion vs target (date + "N days late/early") · Open issues · Overdue items · Decisions waiting · Orders due
- Rows 14–22: **RENOVATION HEALTH** — 6 rows (Budget, Timeline, Contractors & Payments, Materials, Decisions, Issues) each with level 0/1/2 (hidden col), icon, one-line reason; Overall = MAX(levels) → big merged banner "🟢 ON TRACK / 🟡 NEEDS ATTENTION / 🔴 ACTION REQUIRED".
- Rows 24–36: **🔴 YOUR NEXT ACTIONS** headline `=N&" ACTIONS NEED YOUR ATTENTION"` (N = Engine count). Top-10 table: #, icon, item/alert text, source tab, due date, days left/overdue text ("3 days overdue"/"in 4 days"/"today"), priority.
- Rows 38+: **Charts** — clustered bar "Budget vs Forecast by category" (reads Budget cols, only non-empty categories shown via 30-row range; blanks render as empty bars — acceptable), column chart "Monthly cash flow" (Payments cash-flow table: Paid vs Scheduled, 12 months).

### 2.3 THIS WEEK (teal, gridlines off)
- C4 Week start (yellow) default formula `=AsOf-IF(WeekStart="Sunday",WEEKDAY(AsOf,1)-1,WEEKDAY(AsOf,3))`; C5 week end = C4+6.
- **MONDAY PLAN** counts (COUNTIFS on dates between C4:C5): decisions due, orders due, payments due, deliveries expected, tasks starting, tasks ending, open issues.
- Top-10 Engine alerts with due ≤ week end (Engine has a second ranking column for "due ≤ WeekEnd" — see Engine).
- **FRIDAY REVIEW** auto numbers: paid this week (SUMIFS paid date), COs approved this week + cost, tasks completed (done date), issues opened / closed.
- 5 large merged yellow text boxes (wrap): What was completed? What was delayed? What changed? What is now at risk? Top 3 for next week.

### 2.4 BUDGET (green), rows 8–37, totals row 38
| Col | Field | Formula / input |
|---|---|---|
| B | Category | input (prefilled with 18 defaults in both variants) |
| C | Budget | input |
| D | Stage | dropdown Estimating/Contracted/Complete |
| E | Best quote | `1/SUMPRODUCT(MAX((Quotes_BlockCat=B)*Quotes_BlockInvBest))` wrapped in IFERROR→"" |
| F | Contracted | `SUMIFS(Contractors!ContractValue, Contractors!Category, B, Contractors!Signed,"Y")` |
| G | Materials | `SUMIFS(Selections!SelPrice, Selections!Category, B)` |
| H | Approved COs | `SUMIFS(CO!Cost, CO!Category, B, CO!Status,"Approved")` |
| I | Committed | F+G+H |
| J | Invoiced | `SUMIFS(Payments!AmtInvoiced, Payments!CatHelper, B)` |
| K | Paid | `SUMIFS(Payments!AmtPaid, Payments!CatHelper, B)` |
| L | Forecast | `IF(D="Estimating", MAX(C,E,I), I)` |
| M | Variance | C − L |
| N | Status | 🟢 L≤C · 🟡 L≤C×1.05 · 🔴 above (level in hidden col O) |
- Row 38 totals; row 40 **Risk** = pending COs + open-issue est. cost; row 41 **Forecast incl. risk**.
- Budget alerts (hidden helper rows 44–45): forecast > WorkBudget; contingency used > 75%.

### 2.5 QUOTES (green)
- Row 4 weights (yellow): Price 40, Timeline 20, Warranty 15, Completeness 15, Gut/refs 10 (+ check "weights = 100").
- 10 blocks × 42 rows starting row 8 (block k starts at `8+(k-1)*42`). Columns: B labels; each contractor uses 2 columns (C:D, E:F, G:H, I:J) — left = value/status, right = extra-cost estimate.
- Block header: trade/category (dropdown from Budget categories).
- Per contractor: name, price, price incl. VAT? (Y/N), normalized price `IF(inclVAT="Y",price,price*(1+VatRate))`, timeline weeks, warranty years, refs checked Y/N, gut 1–5, **quote expiry date**.
- 13 hidden-cost rows: status dropdown Included/Not included/Unclear; extra cost yellow (CF greys it out when Included).
- Results: extras sum, **TRUE COST**, completeness %, points (price = min true / true × W, timeline = min weeks / weeks × W, warranty = years / max years × W, completeness = % × W, gut/refs = (gut/5 × 0.7 + refs 0.3) × W), total score, rank (`COUNTIF`-based), 🏆 on max.
- Warning row: if contractor with min sticker (normalized) ≠ contractor with min true cost → "⚠️ Hidden costs change the ranking!".
- Hidden summary table at right (cols N:R, rows 8–17): per block category, best true cost, 1/best (for Budget col E). Hidden per-contractor alert helpers (quote expiry) cols T:W, 40 rows.

### 2.6 CONTRACTORS (green)
- Rows 6–35: Company, Trade, Category (dd), Main room (dd, incl. "Whole house"), Contact, Phone, Email, Contract value, Signed Y/N, Start, Expected completion, Insurance ✓ (Y/N), License ✓, Warranty yrs, % work complete (yellow), Paid to date `SUMIFS(Payments paid, contractor)`, Paid %, Overpayment flag `IF(paid% − work% > OverpayTol,"⚠️ Paying ahead of work","")`.
- Rows 40–191 COMMUNICATION LOG: Date, Contractor (dd), Summary, Follow-up Y/N, Follow-up date, Done Y/N.
- Hidden alert helpers at the far right for both tables.

### 2.7 PAYMENTS (green)
- Rows 6–205: Contractor/supplier (dd), Category if not a contractor (dd), Room (dd, optional), Milestone, % of contract, Amount (auto `IF(amount input,…, % × contract value)` → one input col "Fixed amount" + one auto "Scheduled amount"), Due date, Invoice #, Invoiced date, Amount invoiced, Paid date, Amount paid, Status (Scheduled / Invoiced / Paid / OVERDUE), hidden CatHelper = contractor's category or override.
- Right side (cols T+): **per-contractor summary** (30 rows mirroring Contractors list): contract, invoiced, paid, outstanding, next payment amount + date (`SMALL` over due dates of unpaid rows via helper key col).
- **CASH-FLOW** 12 months from `DATE(YEAR(ProjStart),MONTH(ProjStart)+k,1)`: Paid (SUMIFS on paid date), Scheduled unpaid (SUMIFS on due date, status≠Paid), Cumulative. **Cash needed in next 30 days** = unpaid with due ≤ AsOf+30 (incl. overdue).

### 2.8 CHANGE ORDERS (green)
- Row 4 banner: "⚠️ Never approve a change without checking this preview."
- Rows 7–106: CO#, Date, Requested by, Contractor (dd), Category (dd), Room (dd), Description, Reason (dd: Client choice / Unforeseen / Design error / Code requirement), Extra cost, Extra days, Status (dd), Decision date.
- **Impact preview** (auto, per row): new forecast total `=BudgetForecastTotal + IF(status<>"Approved", cost, 0)`, contingency remaining if approved, new forecast completion `=ForecastEnd + IF(status<>"Approved", days, 0)` (approved COs' days must be entered as Timeline delay days — noted in header comment), % over working budget if approved.

### 2.9 TIMELINE (orange)
- Rows 7–156: ID, Phase (dd), Task, Room (dd), Contractor (dd), Duration days, Predecessor ID, Manual start, **Start**, Delay days, **End**, Status (dd), Done date, Affects N, Late flag.
- Start = `MAX(ProjStart, manual, IFERROR(INDEX($L$7:L{r-1}, MATCH(pred,$B$7:B{r-1},0))+1,0))`
  → **predecessor must be listed above its dependent**. This is deliberate: referencing only rows above makes circular references impossible (a full-column INDEX that includes the row itself is flagged as circular by Excel). A red "⚠️ predecessor must be above" flag shows if the ID isn't found above.
- End = Start + Duration + Delay − 1. Affects N = `COUNTIF(PredCol, ID)`. Late = AsOf > End AND status ≠ Done. Critical = Affects N ≥ 1.
- **Gantt**: cols R:BQ = 52 weeks; row 6 header = week-start dates (`ProjStart − WEEKDAY(ProjStart,3) + 7k`). Cells empty; 4 CF rules in priority order: today's week column (thin teal border/fill), Done (grey), Late (red), In progress / planned (teal).
- Forecast completion = `MAX(End)` → named `ForecastEnd`.

### 2.10 ROOMS (orange)
- 10 room cards in a 2 × 5 grid (each card ~12 rows × 5 cols). Room name yellow (6 prefilled + 4 "Custom room N"). Floor area (yellow, label shows `AreaUnit`), Budget (yellow), Cost per area = Forecast/area, Committed (contractors with Main room = room + selections + approved COs), Paid (payments tagged room), Progress bar `REPT` + % (Timeline done/total for room), decisions waiting, open issues, contractors involved (first 3 distinct from Timeline via hidden helper ranks), status icon.
- Lists!Rooms is formula-linked to these names, so renaming a room renames the dropdown option everywhere.

### 2.11 SELECTIONS & ORDERS (orange)
- Rows 4–5 counters: decisions waiting, orders due in window, late orders.
- Rows 8–157: Item, Room, Category, Option A name/price, B name/price, C name/price, Selected (dd A/B/C), Selected price (auto), Reason, Supplier, Lead time days, Linked task ID, Needed-on override, **Needed-on** (`IF(override<>"",override, INDEX(Timeline start, MATCH(task)))`), **Order-by** = needed − lead − OrderBuf, **Decide-by** = order-by − DecideBuf, Ordered Y/N, Order date, Expected delivery, Delivered Y/N, **Status**:
  Delivered → "Delivered"; Ordered & delivery > needed → "⚠️ Delivery after install date"; Ordered → "Ordered"; not selected → "Decide now" (if decide-by ≤ AsOf+window) else "Waiting for decision"; selected & AsOf > order-by → "⚠️ Order late — install at risk"; selected & order-by ≤ AsOf+window → "Order now"; else "Selected".

### 2.12 ISSUES & PUNCH LIST (orange)
- Summary rows 4–5: open by severity, open punch items, est. cost of open issues (→ Budget risk).
- Rows 8–207: #, Type, Date found, Room, Description, Responsible (dd), Severity, Est. cost, Next action, Action date, Status, Resolved date, Photo link (plain text URL; optional `HYPERLINK` display col).

### 2.13 VAULT (purple)
- Rows 6–305: Document, Type (dd), Contractor, Room, Phase/stage (dd: Before / Demolition / Rough-in / Walls / Finishing / Final / —), Date, Link, Notes.
- Right side **Documents you should have**: per signed contractor (30 rows) ✓/✗ for Contract, Insurance, Quote (`COUNTIFS(type, contractor)`), plus project-level Permit ✓.
- **Photo tracker**: 10 rooms × 6 stages grid, ✓ if `COUNTIFS(Type,"Photo",Room,r,Stage,s)>0`.

### 2.14 WARRANTY & MAINTENANCE (purple)
- Warranties rows 7–106: Item, Supplier/contractor, Start date, Years, Expiry `=DATE(YEAR(s)+y,MONTH(s),DAY(s))`, Days left, Status (Active / Expires soon ≤90 / Expired), Doc link.
- Maintenance rows 112–151: Task, Frequency months, Last done, Next due `=DATE(YEAR(l),MONTH(l)+f,DAY(l))`, Status (OK / Due soon / Overdue).

### 2.15 ENGINE (hidden)
Each source sheet has hidden helper columns per row: **AlertText, Due, Priority** (empty string when no alert). Engine only stacks them.

| Block | Source | Rows | Alert rule → priority |
|---|---|---|---|
| 1 | Selections | 150 | Order late / Delivery after install / decide-by passed → P1; Decide now / Order now within window → P2; due within 2× window → P3 |
| 2 | Payments | 200 | OVERDUE → P1; unpaid due within window → P2 |
| 3 | Quotes (expiry) | 40 | expired & block undecided → P1; expiry within window → P2 |
| 4 | Change orders | 100 | Pending > 3 days → P1; pending ≤ 3 days → P2 |
| 5 | Timeline | 150 | Late → P1; starting within window (not started) → P2 |
| 6 | Contractors overpay | 30 | flag → P1 (due = AsOf) |
| 7 | Comm follow-ups | 150 | follow-up date passed → P1; within window → P2 |
| 8 | Issues | 200 | open High → P1; action date passed → P1; action within window → P2 |
| 9 | Warranties | 100 | expires within window → P2; ≤90 days → P3 |
| 10 | Maintenance | 40 | overdue → P1; within window → P2; within 30 days → P3 |
| 11 | Budget | 2 | forecast > WorkBudget → P1; contingency used > 75% → P1 (due = AsOf) |

Total ≈ 1,162 engine rows. Engine columns: Source tab, AlertText, Due, Priority, **Key** = `IF(text="","",Priority*100000+(Due-DATE(2000,1,1))+ROW()/100000)` (row tiebreaker ⇒ keys unique), **WeekKey** = same but only when Due ≤ WeekEnd (for This Week).
Ranked output (rows 1–25 of an output area): `k`, `SMALL(Key,k)`, `MATCH(thatKey,Key,0)`, then `INDEX` for tab, text, due, priority. `COUNT(Key)` = total alerts. Everything wrapped in `IFERROR(…,"")`.

### 2.16 LISTS (hidden)
Currency table (label, symbol), tax names, area units, week-start options, status lists, phases, severities, Y/N, Included/Not included/Unclear, doc types, photo stages, reasons, A/B/C; **dynamic-by-reference** lists: Categories (→ Budget B8:B37), Rooms (→ Rooms names + "Whole house"), Contractors (→ Contractors B6:B35). Validation uses direct ranges `=Lists!$X$2:$X$31` (blank tail entries are harmless).

---

## 3. Health scoreboard (Dashboard, levels 0/1/2)
- Budget: 0 if Forecast ≤ WorkBudget; 1 if ≤ TotalBudget; else 2.
- Timeline: 2 if ForecastEnd > Target+14 OR any late task with Affects ≥ 1; 1 if ForecastEnd > Target; else 0.
- Contractors & Payments: 2 if any OVERDUE payment or overpay flag; 1 if payment due within window; else 0.
- Materials: 2 if any "Order late"/"Delivery after install"; 1 if any "Order now"; else 0.
- Decisions: 2 if any decide-by passed without selection; 1 if decide-by within window; else 0.
- Issues: 2 if any open High; 1 if any open; else 0.
- Overall = MAX → banner text + CF colour.

---

## 4. Demo dataset (~€135k, 1970s semi-detached house)
Setup: start 02 Mar 2026, target move-in 30 Oct 2026, total budget 135,000, contingency 10% (working 121,500), VAT 21%, AsOf pinned **15 Jun 2026** (day 105 of 242).
Designed to trigger:
- **Budget**: kitchen category over budget (🔴), total forecast ≈ 127k > working budget 121.5k but < 135k → Budget 🟡, contingency used ≈ 45%; a second scenario row: pending CO would push contingency use > 75% (visible in CO preview).
- **Quotes**: Kitchen block where the cheapest sticker (€17,900) misses waste disposal, delivery, installation-finishing → true cost €21,400; runner-up wins → ⚠️ + 🏆. One quote expiring in 4 days.
- **Contractors**: 7 contractors; the electrician paid 60% vs 35% work → ⚠️ overpayment. 2 follow-ups (1 overdue, 1 in 3 days).
- **Payments**: ~20 milestones; one invoice overdue, two due within 7 days, cash-flow across Mar–Nov.
- **Change orders**: 4 (2 approved, 1 rejected, 1 pending 6 days old).
- **Timeline**: ~26 tasks (your sequence) with dependencies; plastering delayed 5 days → one late task with dependents; forecast completion ~9 days after target → Timeline 🟡 (or 🔴 via critical late task).
- **Selections**: ~14 items across statuses incl. one "Decide now" overdue, one "Order now", one "Order late", one "Delivery after install date".
- **Issues**: 6 (1 open High, 2 open Medium, 1 punch item, 2 resolved).
- **Vault**: ~25 docs (missing insurance for one contractor), photos covering Before/Demolition for most rooms.
- **Warranty**: 5 items (1 expiring ≤90 days, from pre-existing boiler); maintenance: 4 tasks, 1 due within window.
Expected result: overall 🔴 ACTION REQUIRED with ~15–20 alerts, every alert type present.

---

## 5. Testing (`test_workbook.py`)
1. Build all 4 files.
2. For each: open via Python-UNO in headless LibreOffice (`soffice --headless --accept=…`), `calculateAll()`, iterate every used cell of every sheet: flag error results (`#REF!`, `#VALUE!`, `#NAME?`, `#DIV/0!`, `#N/A`, `Err:5xx`) — Err:522/523 = circular reference.
3. Save a recalculated copy and also scan it with openpyxl `data_only=True` as a second check.
4. DEMO asserts (expected values computed in Python from `demo_data.py`): Committed, Invoiced, Paid, Forecast, contingency used %, ForecastEnd, alert count, top-3 alert texts/order, health levels, quote winner and hidden-cost warning, overpayment flag, CO preview for the pending CO, cash needed next 30 days.
5. BLANK asserts: zero alerts (except none), all KPIs 0/blank, overall 🟢, no errors.
6. Static lint: every formula uses only whitelisted functions; no external refs; no `[`; no whole-column refs.

---

## 6. Build order
1. Skeleton: styles, Lists, Start Here, named ranges, protection helper → test.
2. Contractors, Payments, Budget (core money) → test.
3. Quotes → test. 4. Change orders, Timeline + Gantt → test.
5. Selections, Issues, Rooms, Vault, Warranty → test.
6. Engine, Dashboard, This Week, charts → test.
7. Demo data tuning until every alert type fires; USD variant; docs.

---

## 7. Open questions (defaults I'll use if you don't answer)
1. **Demo "today" pinned to 15 Jun 2026** (override cell on Start Here) — OK? *(Default: yes.)*
2. **USD variant** = same workbook with preference defaults USD $ / Sales tax 8% (*Built:* 8% instead of 0% so the demo's 'price excl. tax' quote shows the normalisation) / ft² / week starts Sunday; same demo figures (demo room areas converted to ft²). *(Default: yes.)*
3. **Currency symbol only in headers/labels, not in the number cells** (required for a switchable currency that works in Google Sheets). *(Default: yes.)*
4. **Predecessor-above rule** in Timeline (prevents circular refs) — acceptable? *(Default: yes, with a visible warning flag.)*
5. **Room attribution of money**: contractors get one "Main room" (or "Whole house"), payments/selections/COs carry their own room tag. Whole-house items are not split across rooms. OK? *(Default: yes.)*
6. Workbook language: English only for v1? *(Default: yes.)*
