"""After-renovation tabs: Vault (documents & photos) and Warranty & Maintenance."""
from . import core as C
from .core import DATE, INT, put, section, write_table
from .layout import (CON, DOCCHK, MAINT, PHOTO_FIRST, PHOTO_HEADER_ROW, PHOTO_STAGES, ROOM_COUNT,
                     ROOMDATA, VAULT, WAR)


def build_vault(ctx):
    ws = ctx.ws(C.S_VAULT)
    ws.sheet_properties.tabColor = C.TAB_PURPLE
    C.title(ws, "🗄️ Vault",
            "What this tab does: an index of every contract, invoice, permit and photo (paste "
            "the Drive/OneDrive link) — and a checklist of what is still missing (right →).")
    T = VAULT
    spec = [
        dict(key="name", header="Document / photo", width=30, kind="in"),
        dict(key="type", header="Type", width=12, kind="in", list="DocType"),
        dict(key="contractor", header="Contractor", width=20, kind="in", list="Contractors",
             warn=True),
        dict(key="room", header="Room", width=14, kind="in", list="Rooms"),
        dict(key="stage", header="Stage", width=12, kind="in", list="PhotoStage",
             note="For photos: which stage of the work (feeds the photo tracker)."),
        dict(key="date", header="Date", width=12, kind="in", fmt=DATE),
        dict(key="link", header="Link (Drive / OneDrive / Dropbox)", width=30, kind="in",
             note="Paste the share link of the file or folder."),
        dict(key="open", header="Open", width=9, kind="auto", align="center",
             f='=IF({link}="","",HYPERLINK({link},"Open ↗"))'),
        dict(key="notes", header="Notes", width=28, kind="in"),
    ]
    write_table(ws, T, spec, ctx.data.VAULT if ctx.demo else [], ctx.dv)
    ws.column_dimensions["K"].width = 3

    D = DOCCHK
    section(ws, "L4", "DOCUMENTS YOU SHOULD HAVE (signed contractors)", "R")
    tv, tc = T.a("type"), T.a("contractor")

    def has(kind):
        return (f'=IF(OR({{who}}="",{{signed}}<>"Yes"),"",IF(COUNTIFS({tv},"{kind}",{tc},'
                f'{{who}})>0,"✓","✗ missing"))')
    spec2 = [
        dict(key="who", header="Contractor", width=24, kind="auto", bold=True,
             f=f'=IF({C.q(C.S_CON)}!{CON.L("company")}{{r}}="","",{C.q(C.S_CON)}!{CON.L("company")}{{r}})'),
        dict(key="signed", header="Signed?", width=9, kind="auto", align="center",
             f=f'=IF({{who}}="","",{C.q(C.S_CON)}!{CON.L("signed")}{{r}})'),
        dict(key="contract", header="Contract", width=11, kind="auto", align="center",
             f=has("Contract")),
        dict(key="insurance", header="Insurance", width=11, kind="auto", align="center",
             f=has("Insurance")),
        dict(key="quote", header="Quote", width=11, kind="auto", align="center", f=has("Quote")),
        dict(key="doc_status", header="Status", width=19, kind="auto", align="center",
             f='=IF({contract}="","",IF(AND({contract}="✓",{insurance}="✓",{quote}="✓"),'
               '"✓ Complete","Missing documents"))'),
        dict(key="lvl", header="lvl", kind="hid",
             f='=IF({doc_status}="","",IF({doc_status}="✓ Complete",0,1))'),
    ]
    write_table(ws, D, spec2)
    rng = f"N{D.first}:P{D.last}"
    C.add_cf(ws, rng, f'N{D.first}="✗ missing"', C.BAD, bold=True)
    C.add_cf(ws, rng, f'N{D.first}="✓"', C.GOOD)
    C.level_cf(ws, f"Q{D.first}:Q{D.last}", f"$R{D.first}")
    C.define(ctx.wb, "DocsMissing", C.S_VAULT, f"$R${D.first}:$R${D.last}")

    section(ws, "L37", "PROJECT DOCUMENTS", "R")
    for i, (lab, kind) in enumerate([("Building permit", "Permit"),
                                     ("Plans / drawings", "Plan"),
                                     ("Inspection reports", "Inspection")]):
        r = 38 + i
        ws.merge_cells(f"L{r}:M{r}")
        put(ws, f"L{r}", lab, "label", bold=True)
        put(ws, f"N{r}", f'=IF(COUNTIF({tv},"{kind}")>0,"✓","✗ missing")', "auto",
            align="center")
        C.add_cf(ws, f"N{r}", f'N{r}="✗ missing"', C.BAD, bold=True)
        C.add_cf(ws, f"N{r}", f'N{r}="✓"', C.GOOD)

    section(ws, f"L{PHOTO_HEADER_ROW - 1}", "PHOTO TRACKER — ✓ when a Photo row exists for "
            "room + stage", "R")
    put(ws, f"L{PHOTO_HEADER_ROW}", "Room", "header")
    for j, stg in enumerate(PHOTO_STAGES):
        put(ws, f"{C.col_add('M', j)}{PHOTO_HEADER_ROW}", stg, "header")
    for i in range(ROOM_COUNT):
        r = PHOTO_FIRST + i
        src = f"{C.q(C.S_ROOMS)}!{ROOMDATA.c('name', ROOMDATA.first + i)}"
        put(ws, f"L{r}", f'=IF({src}="","",{src})', "auto", bold=True)
        for j in range(len(PHOTO_STAGES)):
            L = C.col_add("M", j)
            put(ws, f"{L}{r}", f'=IF($L{r}="","",IF(COUNTIFS({tv},"Photo",{T.a("room")},$L{r},'
                f'{T.a("stage")},{L}${PHOTO_HEADER_ROW})>0,"✓","·"))', "auto", align="center")
        ws.row_dimensions[r].height = 20
    prng = f"M{PHOTO_FIRST}:R{PHOTO_FIRST + ROOM_COUNT - 1}"
    C.add_cf(ws, prng, f'M{PHOTO_FIRST}="✓"', C.GOOD, bold=True)
    ws.freeze_panes = f"C{T.first}"
    C.protect(ws)


def build_warranty(ctx):
    ws = ctx.ws(C.S_WAR)
    ws.sheet_properties.tabColor = C.TAB_PURPLE
    C.title(ws, "🛡️ Warranty & Maintenance",
            "What this tab does: tracks warranty expiry dates (check for defects before they "
            "run out) and recurring home maintenance. Both create alerts.")
    section(ws, "B4", "WARRANTIES", "J")
    T = WAR
    spec = [
        dict(key="item", header="Item / work", width=28, kind="in"),
        dict(key="supplier", header="Supplier / contractor", width=20, kind="in",
             list="Contractors", warn=True),
        dict(key="room", header="Room", width=14, kind="in", list="Rooms"),
        dict(key="start", header="Purchase / completion date", width=13, kind="in", fmt=DATE),
        dict(key="years", header="Warranty (years)", width=13, kind="in", align="center"),
        dict(key="expiry", header="Expires", width=13, kind="auto", fmt=DATE, bold=True,
             f='=IF(OR({start}="",{years}=""),"",DATE(YEAR({start}),MONTH({start})+'
               'ROUND({years}*12,0),DAY({start})))'),
        dict(key="days", header="Days left", width=10, kind="auto", fmt=INT, align="center",
             f='=IF({expiry}="","",{expiry}-AsOf)'),
        dict(key="status", header="Status", width=18, kind="auto", align="center", bold=True,
             note="Expires soon = within 90 days: inspect and report defects now.",
             f='=IF({expiry}="","",IF({days}<0,"Expired",IF({days}<=90,"Expires soon","Active")))'),
        dict(key="link", header="Warranty document link", width=26, kind="in"),
        dict(key="lvl", header="lvl", kind="hid",
             f='=IF({status}="","",IF({status}="Active",0,IF({status}="Expires soon",1,3)))'),
        dict(key="due", header="due", kind="hid", fmt=DATE,
             f='=IF({expiry}="","",IF({expiry}<AsOf,"",{expiry}))'),
        dict(key="pri", header="pri", kind="hid",
             f='=IF({due}="","",IF({due}<=AsOf+AlertDays,2,IF({due}<=AsOf+90,3,"")))'),
        dict(key="text", header="text", kind="hid",
             f='=IF({pri}="","","Warranty expires in "&{days}&" days: "&{item}&IF({supplier}="",'
               '""," ("&{supplier}&")")&" — check for defects")'),
    ]
    write_table(ws, T, spec, ctx.data.WARRANTIES if ctx.demo else [], ctx.dv)
    S = T.L("status")
    rng = f"{S}{T.first}:{S}{T.last}"
    C.level_cf(ws, rng, f"${T.L('lvl')}{T.first}")
    C.add_cf_fill(ws, rng, f"${T.L('lvl')}{T.first}=3", C.AUTO_BG, C.GREY_TEXT)

    M = MAINT
    section(ws, f"B{M.header_row - 2}", "HOME MAINTENANCE", "J")
    spec2 = [
        dict(key="task", header="Maintenance task", kind="in"),
        dict(key="room", header="Room", kind="in", list="Rooms"),
        dict(key="freq", header="Every (months)", kind="in", fmt=INT, align="center"),
        dict(key="last", header="Last done", kind="in", fmt=DATE),
        dict(key="next", header="Next due", kind="auto", fmt=DATE, bold=True,
             f='=IF(OR({freq}="",{last}=""),"",DATE(YEAR({last}),MONTH({last})+{freq},DAY({last})))'),
        dict(key="days", header="Days left", kind="auto", fmt=INT, align="center",
             f='=IF({next}="","",{next}-AsOf)'),
        dict(key="status", header="Status", kind="auto", align="center", bold=True,
             f='=IF({next}="",IF({task}="","","Enter last-done date"),IF({days}<0,"Overdue",'
               'IF({days}<=30,"Due soon","OK")))'),
        dict(key="lvl", header="lvl", kind="hid",
             f='=IF({next}="","",IF({days}<0,2,IF({days}<=30,1,0)))'),
        dict(key="due", header="due", kind="hid", fmt=DATE, f='=IF({next}="","",{next})'),
        dict(key="pri", header="pri", kind="hid",
             f='=IF({due}="","",IF({due}<=AsOf+AlertDays,2,IF({due}<=AsOf+30,3,"")))'),
        dict(key="text", header="text", kind="hid",
             f='=IF({pri}="","",IF({days}<0,"Maintenance overdue: ","Maintenance due: ")&{task})'),
    ]
    write_table(ws, M, spec2, ctx.data.MAINTENANCE if ctx.demo else [], ctx.dv)
    S = M.L("status")
    C.level_cf(ws, f"{S}{M.first}:{S}{M.last}", f"$K{M.first}")
    ws.freeze_panes = f"C{T.first}"
    C.protect(ws)
