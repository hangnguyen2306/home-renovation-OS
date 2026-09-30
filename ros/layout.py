"""Cell map: every list's position and column keys, shared by all sheet builders."""
from .core import (S_BUDGET, S_CO, S_CON, S_ISS, S_PAY, S_ROOMS, S_SEL, S_TL, S_VAULT, S_WAR,
                   Table)

# ------------------------------------------------------------------ BUDGET
BUDGET = Table(S_BUDGET, 7, 8, 30, [
    "cat", "budget", "stage", "best", "contracted", "materials", "cos", "committed",
    "invoiced", "paid", "forecast", "variance", "status",
    "lvl", "chartkey", "due", "pri", "text",
])
BUDGET_TOTAL_ROW = 38
# rows 40..44: risk block; budget alerts live in hidden due/pri/text on rows 40..41
BUDGET_ALERT_ROWS = (40, 41)

# ------------------------------------------------------------------ QUOTES
QUOTE_BLOCKS = 10
QUOTE_BLOCK_START = 9
QUOTE_BLOCK_H = 39
QUOTE_VCOLS = ["C", "E", "G", "I"]   # value column per contractor
QUOTE_XCOLS = ["D", "F", "H", "J"]   # extra-cost column per contractor
CHECKLIST = ["Demolition", "Waste disposal", "Delivery", "Permits", "Materials specified",
             "Labor specified", "Finishing", "Cleanup", "Tax (VAT) shown", "Electrical upgrades",
             "Plumbing modifications", "Payment schedule defined", "Warranty in writing"]
# row offsets inside a block
QO = dict(head=0, sub=1, name=2, price=3, incl=4, norm=5, weeks=6, warranty=7, refs=8, gut=9,
          valid=10, chk_head=11, chk_first=12, extras=25, true=26, complete=27, score_head=28,
          p_price=29, p_time=30, p_war=31, p_comp=32, p_gut=33, total=34, rank=35, warn=36)
# hidden summary: one row per block (rows 9..18) in M:O; expiry alerts 40 rows in T:V
QSUM_FIRST = 9
QEXP_FIRST = 9


def qblock(k):
    """First row of quote block k (0-based)."""
    return QUOTE_BLOCK_START + k * QUOTE_BLOCK_H


# ------------------------------------------------------------------ CONTRACTORS
CON = Table(S_CON, 5, 6, 30, [
    "company", "trade", "cat", "room", "contact", "phone", "email", "value", "co_add", "total",
    "signed", "start", "end", "insured", "licensed", "warranty", "work", "paid", "paid_pct",
    "flag", "lvl", "due", "pri", "text",
])
COMM = Table(S_CON, 40, 41, 150, [
    "who", "date", "summary", ("fu", "I"), "fu_date", "done", ("due", "W"), "pri", "text",
])

# ------------------------------------------------------------------ PAYMENTS
PAY = Table(S_PAY, 5, 6, 200, [
    "party", "milestone", "room", "cat_in", "pct", "fixed", "sched", "due_date", "inv_no",
    "inv_date", "inv_amt", "paid_date", "paid_amt", "status", "cat", "room_auto", "lvl",
    "unpaid", "inv_due", "due", "pri", "text",
])
PAYSUM = Table(S_PAY, 5, 6, 30, [
    ("name", "Z"), "contract", "invoiced", "paid", "outstanding", "next_date", "next_amt",
])
CASH = Table(S_PAY, 40, 41, 12, [("month", "Z"), "c_paid", "c_sched", "c_total", "c_cum"])
CASH30_CELL = "AC38"

# ------------------------------------------------------------------ CHANGE ORDERS
CO = Table(S_CO, 6, 7, 100, [
    "no", "date", "req", "contractor", "cat", "room", "desc", "reason", "cost", "days", "status",
    "dec_date", "new_total", "cont_left", "new_end", "pct_over", "lvl", "due", "pri", "text",
])

# ------------------------------------------------------------------ TIMELINE
TL = Table(S_TL, 6, 7, 150, [
    "id", "phase", "task", "room", "contractor", "dur", "pred", "manual", "start", "delay", "end",
    "status", "done_date", "affects", "flag", "code", "first", "okey", "due", "pri", "text",
])
GANTT_FIRST = "X"
GANTT_WEEKS = 52
TL_ASOF_CELL = "R4"      # local copy of AsOf for conditional formatting (hidden column)

# ------------------------------------------------------------------ ROOMS
ROOM_COUNT = 10
ROOMDATA = Table(S_ROOMS, 5, 6, ROOM_COUNT, [
    ("name", "Q"), "area", "budget", "committed", "paid", "t_total", "t_done", "progress",
    "decisions", "open_issues", "open_high", "late", "lvl", "c1", "c2", "c3", "per_area",
    "forecast_room",
])
ROOM_CARD_H = 11
ROOM_CARD_COLS = [("B", "F"), ("H", "L")]
ROOM_CARD_TOP = 5


def room_card(k):
    """(first column, last column, top row) of room card k (0-based)."""
    c0, c1 = ROOM_CARD_COLS[k % 2]
    return c0, c1, ROOM_CARD_TOP + (k // 2) * (ROOM_CARD_H + 1)


# ------------------------------------------------------------------ SELECTIONS
SEL = Table(S_SEL, 7, 8, 150, [
    "item", "room", "cat", "a_name", "a_price", "b_name", "b_price", "c_name", "c_price",
    "selected", "price", "reason", "supplier", "lead", "task", "need_in", "need", "order_by",
    "decide_by", "ordered", "order_date", "delivery", "delivered", "status",
    "lvl", "due", "pri", "text", "dec_wait", "dec_over", "dec_soon", "order_now", "late_mat",
    "need_order", "order_late",
])

# ------------------------------------------------------------------ ISSUES
ISS = Table(S_ISS, 7, 8, 200, [
    "no", "type", "found", "room", "desc", "resp", "sev", "cost", "action", "act_date", "status",
    "res_date", "link", "open", "lvl", "due", "pri", "text", "open_cost", "open_high",
    "open_punch",
])

# ------------------------------------------------------------------ VAULT
VAULT = Table(S_VAULT, 5, 6, 300, [
    "name", "type", "contractor", "room", "stage", "date", "link", "open", "notes",
])
DOCCHK = Table(S_VAULT, 5, 6, 30, [
    ("who", "L"), "signed", "contract", "insurance", "quote", "doc_status", "lvl",
])
PHOTO_STAGES = ["Before", "Demolition", "Rough-in", "Walls", "Finishing", "Final"]
PHOTO_HEADER_ROW = 43
PHOTO_FIRST = 44

# ------------------------------------------------------------------ WARRANTY
WAR = Table(S_WAR, 6, 7, 60, [
    "item", "supplier", "room", "start", "years", "expiry", "days", "status", "link",
    "lvl", "due", "pri", "text",
])
MAINT = Table(S_WAR, 71, 72, 40, [
    "task", "room", "freq", "last", "next", "days", "status", ("lvl", "K"), "due", "pri", "text",
])

# ------------------------------------------------------------------ ENGINE blocks
# (tab label, sheet, first row, n rows, text col, due col, pri col)
ENGINE_BLOCKS = [
    ("Selections & Orders", S_SEL, SEL.first, SEL.n, SEL.L("text"), SEL.L("due"), SEL.L("pri")),
    ("Payments", S_PAY, PAY.first, PAY.n, PAY.L("text"), PAY.L("due"), PAY.L("pri")),
    ("Quotes", "Quotes", QEXP_FIRST, QUOTE_BLOCKS * 4, "V", "T", "U"),
    ("Change Orders", S_CO, CO.first, CO.n, CO.L("text"), CO.L("due"), CO.L("pri")),
    ("Timeline", S_TL, TL.first, TL.n, TL.L("text"), TL.L("due"), TL.L("pri")),
    ("Contractors", S_CON, CON.first, CON.n, CON.L("text"), CON.L("due"), CON.L("pri")),
    ("Contractors", S_CON, COMM.first, COMM.n, COMM.L("text"), COMM.L("due"), COMM.L("pri")),
    ("Issues & Punch List", S_ISS, ISS.first, ISS.n, ISS.L("text"), ISS.L("due"), ISS.L("pri")),
    ("Warranty & Maintenance", S_WAR, WAR.first, WAR.n, WAR.L("text"), WAR.L("due"),
     WAR.L("pri")),
    ("Warranty & Maintenance", S_WAR, MAINT.first, MAINT.n, MAINT.L("text"), MAINT.L("due"),
     MAINT.L("pri")),
    ("Budget", S_BUDGET, BUDGET_ALERT_ROWS[0], 2, BUDGET.L("text"), BUDGET.L("due"),
     BUDGET.L("pri")),
]
ENGINE_FIRST = 2
ENGINE_ROWS = sum(b[3] for b in ENGINE_BLOCKS)
ENGINE_LAST = ENGINE_FIRST + ENGINE_ROWS - 1
TOP_N = 25
