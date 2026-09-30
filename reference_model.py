"""Independent Python re-implementation of the Renovation OS rules for the demo data.

test_workbook.py compares the spreadsheet (recalculated by LibreOffice) with these values.
Nothing here reads the workbook: it only uses demo_data.py and the documented rules.
"""
from datetime import date, timedelta

import demo_data as d
from ros import core as C
from ros.layout import (CHECKLIST, CO as T_CO, COMM, CON, ENGINE_BLOCKS, ENGINE_FIRST, ISS, MAINT,
                        PAY, QEXP_FIRST, SEL, TL, WAR, BUDGET_ALERT_ROWS)

EPOCH = date(2000, 1, 1)


def add_months(dt, n):
    """Excel DATE(YEAR(d), MONTH(d)+n, DAY(d)) semantics (day overflow rolls over)."""
    m0 = dt.month - 1 + n
    y, m = dt.year + m0 // 12, m0 % 12 + 1
    return date(y, m, 1) + timedelta(days=dt.day - 1)


def fixed(x):
    return f"{round(x):,}"


class Model:
    def __init__(self, preset):
        self.p = preset
        self.sym = {"EUR €": "€", "USD $": "$"}[preset["currency"]]
        self.rate = preset["rate"]
        s = d.SETUP
        self.asof = d.AS_OF
        self.start, self.target = s["start"], s["target"]
        self.total_budget = s["total_budget"]
        self.cont = self.total_budget * 0.10
        self.work = self.total_budget - self.cont
        self.AD, self.DB, self.OB, self.tol = 7, 7, 5, 0.10
        self.alerts = []   # (engine_row, tab, text, due, pri)
        self._timeline()
        self._contractors()
        self._payments()
        self._selections()
        self._quotes()
        self._budget()
        self._issues()
        self._alerts()

    # ------------------------------------------------------------ helpers
    def pri(self, due):
        if due is None:
            return None
        if due < self.asof:
            return 1
        if due <= self.asof + timedelta(days=self.AD):
            return 2
        if due <= self.asof + timedelta(days=2 * self.AD):
            return 3
        return None

    # ------------------------------------------------------------ timeline
    def _timeline(self):
        self.tasks = []
        ends = {}
        for i, t in enumerate(d.TASKS, 1):
            start = self.start
            if t.get("manual"):
                start = max(start, t["manual"])
            p = t.get("pred")
            if p and p < i and p in ends:
                start = max(start, ends[p] + timedelta(days=1))
            end = start + timedelta(days=max(1, t.get("dur") or 0) + (t.get("delay") or 0) - 1)
            ends[i] = end
            code = 2 if t["status"] == "Done" else (3 if self.asof > end else 1)
            self.tasks.append(dict(t, id=i, start=start, end=end, code=code))
        for t in self.tasks:
            t["affects"] = sum(1 for u in self.tasks if u.get("pred") == t["id"])
        self.forecast_end = max(t["end"] for t in self.tasks)
        self.task_start = {t["id"]: t["start"] for t in self.tasks}

    # ------------------------------------------------------------ contractors / COs
    def _contractors(self):
        self.cos = d.CHANGE_ORDERS
        paid = {}
        for p in d.PAYMENTS:
            if p.get("paid_amt"):
                paid[p["party"]] = paid.get(p["party"], 0) + p["paid_amt"]
        self.con = {}
        for c in d.CONTRACTORS:
            co_add = sum(x["cost"] for x in self.cos
                         if x["contractor"] == c["company"] and x["status"] == "Approved")
            total = (c.get("value") or 0) + co_add
            pd = paid.get(c["company"], 0)
            pct = pd / total if total else None
            lvl = None if pct is None else (2 if pct - (c.get("work") or 0) > self.tol else 0)
            self.con[c["company"]] = dict(c, total=total, paid=pd, paid_pct=pct, lvl=lvl)

    # ------------------------------------------------------------ payments
    def _payments(self):
        self.pay = []
        for p in d.PAYMENTS:
            c = self.con.get(p["party"])
            if p.get("fixed") is not None:
                sched = p["fixed"]
            elif p.get("pct") is not None:
                sched = p["pct"] * ((c or {}).get("value") or 0)
            else:
                sched = None
            if p.get("paid_date"):
                status = "Paid"
            elif p.get("due_date") and p["due_date"] < self.asof:
                status = "OVERDUE"
            elif p.get("inv_date") or p.get("inv_amt"):
                status = "Invoiced"
            else:
                status = "Scheduled"
            cat = p.get("cat_in") or (c or {}).get("cat", "")
            room = p.get("room") or (c or {}).get("room", "")
            unpaid = 0
            if not p.get("paid_date") and p.get("due_date"):
                unpaid = p["inv_amt"] if p.get("inv_amt") is not None else (sched or 0)
            lvl = 2 if status == "OVERDUE" else 0 if status == "Paid" else (
                1 if p.get("due_date") and p["due_date"] <= self.asof + timedelta(days=self.AD)
                else 0)
            self.pay.append(dict(p, sched=sched, status=status, cat=cat, room_auto=room,
                                 unpaid=unpaid, lvl=lvl))
        self.cash30 = sum(p["unpaid"] for p in self.pay
                          if p.get("due_date") and p["due_date"] <= self.asof + timedelta(30))

    # ------------------------------------------------------------ selections
    def _selections(self):
        self.sel = []
        W = timedelta(days=self.AD)
        for s in d.SELECTIONS:
            price = {"A": s.get("a_price"), "B": s.get("b_price"),
                     "C": s.get("c_price")}.get(s.get("selected"))
            need = s.get("need_in") or (self.task_start.get(s["task"]) if s.get("task") else None)
            order_by = need - timedelta(days=(s.get("lead") or 0) + self.OB) if need else None
            decide_by = order_by - timedelta(days=self.DB) if order_by else None
            if s.get("delivered") == "Yes":
                st = C.ST_DELIVERED
            elif s.get("ordered") == "Yes":
                st = (C.ST_DELIV_LATE if s.get("delivery") and need and s["delivery"] > need
                      else C.ST_ORDERED)
            elif not s.get("selected"):
                st = C.ST_DECIDE if decide_by and decide_by <= self.asof + W else C.ST_WAITING
            elif order_by is None:
                st = C.ST_SELECTED
            elif self.asof > order_by:
                st = C.ST_ORDER_LATE
            elif order_by <= self.asof + W:
                st = C.ST_ORDER_NOW
            else:
                st = C.ST_SELECTED
            dec_wait = int(not s.get("selected") and s.get("ordered") != "Yes"
                           and s.get("delivered") != "Yes")
            dec_over = int(dec_wait and decide_by is not None and decide_by < self.asof)
            dec_soon = int(dec_wait and decide_by is not None and self.asof <= decide_by <= self.asof + W)
            late_mat = int(st in (C.ST_ORDER_LATE, C.ST_DELIV_LATE))
            need_order = int(bool(s.get("selected")) and s.get("ordered") != "Yes"
                             and s.get("delivered") != "Yes")
            self.sel.append(dict(s, price=price, need=need, order_by=order_by,
                                 decide_by=decide_by, status=st, dec_wait=dec_wait,
                                 dec_over=dec_over, dec_soon=dec_soon, late_mat=late_mat,
                                 order_now=int(st == C.ST_ORDER_NOW),
                                 order_late=int(st == C.ST_ORDER_LATE), need_order=need_order))

    # ------------------------------------------------------------ quotes
    def _quotes(self):
        self.qblocks = []
        weights = dict(price=40, time=20, war=15, comp=15, gut=10)
        for b in d.QUOTES:
            rows = []
            for c in b["contractors"]:
                norm = c["price"] * (1 + self.rate) if c["incl"] == "No" else c["price"]
                chk = [c["chk"].get(i, ("Included", None)) for i in range(len(CHECKLIST))]
                extras = sum((x or 0) for st, x in chk if st != "Included")
                complete = sum(1 for st, _ in chk if st == "Included") / len(CHECKLIST)
                rows.append(dict(c, norm=norm, true=norm + extras, complete=complete))
            mt = min(r["true"] for r in rows)
            mw = min(r["weeks"] for r in rows)
            mx = max(r["warranty"] for r in rows)
            for r in rows:
                r["total"] = round(mt / r["true"] * weights["price"]
                                   + mw / r["weeks"] * weights["time"]
                                   + r["warranty"] / mx * weights["war"]
                                   + r["complete"] * weights["comp"]
                                   + (r["gut"] / 5 * 0.7 + (0.3 if r["refs"] == "Yes" else 0))
                                   * weights["gut"], 1)
            winner = max(rows, key=lambda r: r["total"])
            cheap_sticker = min(rows, key=lambda r: r["norm"])
            cheap_true = min(rows, key=lambda r: r["true"])
            self.qblocks.append(dict(cat=b["cat"], rows=rows, best=mt, winner=winner["name"],
                                     winner_true=winner["true"],
                                     warn=cheap_sticker["name"] != cheap_true["name"],
                                     cheap_sticker=cheap_sticker["name"],
                                     cheap_true=cheap_true["name"]))

    # ------------------------------------------------------------ budget
    def _budget(self):
        self.cats = []
        for name, bud, stage in d.CATEGORIES:
            best = min([b["best"] for b in self.qblocks if b["cat"] == name], default=None)
            contracted = sum(c.get("value") or 0 for c in d.CONTRACTORS
                             if c["cat"] == name and c.get("signed") == "Yes")
            materials = sum(s["price"] or 0 for s in self.sel if s["cat"] == name)
            cos = sum(x["cost"] for x in self.cos if x["cat"] == name and x["status"] == "Approved")
            committed = contracted + materials + cos
            invoiced = sum(p.get("inv_amt") or 0 for p in self.pay if p["cat"] == name)
            paid = sum(p.get("paid_amt") or 0 for p in self.pay if p["cat"] == name)
            if stage == "Estimating":
                fc = max(bud, best or 0, committed, invoiced)
            else:
                fc = max(committed, invoiced)
            self.cats.append(dict(name=name, budget=bud, stage=stage, best=best,
                                  committed=committed, invoiced=invoiced, paid=paid, forecast=fc))
        self.committed = sum(c["committed"] for c in self.cats)
        self.invoiced = sum(c["invoiced"] for c in self.cats)
        self.paid = sum(c["paid"] for c in self.cats)
        self.forecast = sum(c["forecast"] for c in self.cats)
        self.cont_used = max(0, self.forecast - self.work) / self.cont
        self.pending_co = sum(x["cost"] for x in self.cos if x["status"] == "Pending")

    # ------------------------------------------------------------ issues
    def _issues(self):
        self.iss = []
        for i in d.ISSUES:
            op = int(bool(i.get("desc")) and i.get("status") != "Resolved")
            self.iss.append(dict(i, open=op))
        self.open_issues = sum(i["open"] for i in self.iss)
        self.open_high = sum(1 for i in self.iss if i["open"] and i["sev"] == "High")
        self.open_cost = sum(i.get("cost") or 0 for i in self.iss if i["open"])

    # ------------------------------------------------------------ alerts
    def _alerts(self):
        base = {}
        r = ENGINE_FIRST
        for blk in ENGINE_BLOCKS:
            base[(blk[1], blk[2])] = r
            r += blk[3]
        A = self.alerts
        W = timedelta(days=self.AD)

        def add(sheet, first, idx, tab, text, due, pri):
            if pri is None or due is None:
                return
            A.append((base[(sheet, first)] + idx, tab, text, due, pri))

        for i, s in enumerate(self.sel):
            if s["late_mat"]:
                due = s["need"] if s["status"] == C.ST_DELIV_LATE else s["order_by"]
                pri = 1
                text = ("Delivery after install date: " + s["item"] if s["status"] == C.ST_DELIV_LATE
                        else "Order late — install at risk: " + s["item"])
            elif s["dec_wait"]:
                due = s["decide_by"]
                pri = self.pri(due)
                text = ("Decision overdue: " if pri == 1 else "Decide: ") + s["item"]
            elif s["need_order"]:
                due = s["order_by"]
                pri = self.pri(due)
                text = "Order: " + s["item"] + (" from " + s["supplier"] if s.get("supplier") else "")
            else:
                continue
            add(C.S_SEL, SEL.first, i, "Selections & Orders", text, due, pri)

        for i, p in enumerate(self.pay):
            if p.get("paid_date") or not p.get("due_date"):
                continue
            due = p["due_date"]
            text = (("Overdue payment: " if p["status"] == "OVERDUE" else "Payment due: ")
                    + f'{p["party"]} — {p["milestone"]} ({self.sym} {fixed(p["unpaid"])})')
            add(C.S_PAY, PAY.first, i, "Payments", text, due, self.pri(due))

        stage = {n: st for n, _, st in d.CATEGORIES}
        for k, b in enumerate(d.QUOTES):
            for j, c in enumerate(b["contractors"]):
                if stage.get(b["cat"]) in ("Contracted", "Complete") or c["valid"] < self.asof:
                    continue
                text = f'Quote expires: {c["name"]} ({b["cat"]}) — decide or ask for an extension'
                add("Quotes", QEXP_FIRST, k * 4 + j, "Quotes", text, c["valid"], self.pri(c["valid"]))

        for i, x in enumerate(self.cos):
            if x["status"] not in ("Pending", "") or not x.get("date"):
                continue
            due = x["date"] + timedelta(days=3)
            pri = 1 if self.asof > due else 2
            text = f'Decide on change order {x["no"]}: {x["desc"]} (+{self.sym} {fixed(x["cost"])})'
            add(C.S_CO, T_CO.first, i, "Change Orders", text, due, pri)

        for i, t in enumerate(self.tasks):
            tail = " — " + t["contractor"] if t.get("contractor") else ""
            if t["code"] == 3:
                add(C.S_TL, TL.first, i, "Timeline", "Late task: " + t["task"] + tail, t["end"], 1)
            elif t["status"] in ("Not started", ""):
                due = t["start"]
                text = ("Should have started: " if due < self.asof else "Starts soon: ") + t["task"] + tail
                add(C.S_TL, TL.first, i, "Timeline", text, due, self.pri(due))

        for i, c in enumerate(d.CONTRACTORS):
            m = self.con[c["company"]]
            if m["lvl"] == 2:
                text = (f'Paying ahead of work: {c["company"]} — paid {round(m["paid_pct"] * 100)}%, '
                        f'work done {round((c.get("work") or 0) * 100)}%')
                add(C.S_CON, CON.first, i, "Contractors", text, self.asof, 1)

        for i, m in enumerate(d.COMM_LOG):
            if m.get("fu") == "Yes" and m.get("done") != "Yes" and m.get("fu_date"):
                add(C.S_CON, COMM.first, i, "Contractors",
                    f'Follow up with {m["who"]}: {m["summary"]}', m["fu_date"], self.pri(m["fu_date"]))

        for i, s in enumerate(self.iss):
            if not s["open"]:
                continue
            due = s.get("act_date") or (self.asof if s["sev"] == "High" else None)
            pri = 1 if s["sev"] == "High" else self.pri(due)
            pre = ("High-severity issue: " if s["sev"] == "High" else
                   "Punch item: " if s["type"] == "Punch item" else "Issue action due: ")
            text = pre + s["desc"] + (" → " + s["action"] if s.get("action") else "")
            add(C.S_ISS, ISS.first, i, "Issues & Punch List", text, due, pri)

        for i, w in enumerate(d.WARRANTIES):
            exp = add_months(w["start"], round(w["years"] * 12))
            if exp < self.asof:
                continue
            pri = 2 if exp <= self.asof + W else 3 if exp <= self.asof + timedelta(90) else None
            days = (exp - self.asof).days
            text = (f'Warranty expires in {days} days: {w["item"]}'
                    + (f' ({w["supplier"]})' if w.get("supplier") else "") + " — check for defects")
            add(C.S_WAR, WAR.first, i, "Warranty & Maintenance", text, exp, pri)

        for i, m in enumerate(d.MAINTENANCE):
            if not m.get("last"):
                continue
            nxt = add_months(m["last"], m["freq"])
            pri = 2 if nxt <= self.asof + W else 3 if nxt <= self.asof + timedelta(30) else None
            text = ("Maintenance overdue: " if nxt < self.asof else "Maintenance due: ") + m["task"]
            add(C.S_WAR, MAINT.first, i, "Warranty & Maintenance", text, nxt, pri)

        if self.forecast > self.work:
            add(C.S_BUDGET, BUDGET_ALERT_ROWS[0], 0, "Budget",
                f"Forecast is over your working budget by {self.sym} "
                f"{fixed(self.forecast - self.work)} — review Budget", self.asof, 1)
        if self.cont_used > 0.75:
            add(C.S_BUDGET, BUDGET_ALERT_ROWS[0], 1, "Budget",
                f"Contingency {round(self.cont_used * 100)}% used — pause optional changes",
                self.asof, 1)

        def key(a):
            row, _, _, due, pri = a
            return pri * 100000 + (due - EPOCH).days + row / 100000
        A.sort(key=key)

    # ------------------------------------------------------------ health
    def health(self):
        crit = sum(1 for t in self.tasks if t["code"] == 3 and t["affects"] >= 1)
        h = {}
        h["budget"] = 0 if self.forecast <= self.work else 1 if self.forecast <= self.total_budget else 2
        if crit:
            h["timeline"] = 2
        else:
            late = (self.forecast_end - self.target).days
            h["timeline"] = 2 if late > 14 else 1 if late > 0 else 0
        overdue = sum(1 for p in self.pay if p["status"] == "OVERDUE")
        overpay = sum(1 for c in self.con.values() if c["lvl"] == 2)
        due_w = sum(1 for p in self.pay if p["lvl"] == 1)
        h["contractors"] = 2 if overdue or overpay else 1 if due_w else 0
        late_mat = sum(s["late_mat"] for s in self.sel)
        h["materials"] = 2 if late_mat else 1 if sum(s["order_now"] for s in self.sel) else 0
        h["decisions"] = (2 if sum(s["dec_over"] for s in self.sel) else
                          1 if sum(s["dec_soon"] for s in self.sel) else 0)
        h["issues"] = 2 if self.open_high else 1 if self.open_issues else 0
        h["overall"] = max(h.values())
        return h
