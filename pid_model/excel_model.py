"""
excel_model.py — Standalone, fully recalculating Excel version of the Texas PID model.

Every number in the workbook is a live Excel formula driven by three input tabs
(Inputs, Developer Information, Overlapping Taxes) — the same tabs, layout and
range names as the Python inputs workbook, so the Python model can also read
this file and the two can be reconciled.

    build_excel_model(path, cfg=None, comparables=None)

Calculation flow
  Absorption (monthly)      lots open after delivery + lag; homes close at pace, FIFO by area
  Series 1 … Series 10      one tab per improvement area: per-lot capacity (closed form),
                            sizing (capacity / costs / VTL), annual bond and installment
                            schedule, reserve, CAPI, sources & uses
  Output tabs               Summary, Sources and Uses, Bond Sizing Analysis, Assessment Roll,
                            Annual Installments (selector), Project and PID Rev Analysis,
                            Tax Rate and VTL Tests, Homeowner Cost Comparison

Sizing to costs is solved without iteration: every use other than the project fund
is proportional to par, so  par = costs / (1 − non-project uses as % of par).
"""

from __future__ import annotations

from typing import Optional

from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter as CL
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

from .config import ModelConfig, sample_config
from .inputs import (DI_PHASE_HDR_ROW, DI_PROD_FIRST_COL, DI_PROD_HDR_ROW, MAX_PHASES,
                     MAX_PRODUCTS, MAX_TAXES, build_inputs_template)
from .report import (_BLUE_FILL, _BODY_FONT, _BOLD_FONT, _BORDER, _CENTER, _FMT_DATE, _FMT_DOLLAR,
                     _FMT_DOLLAR2, _FMT_NUM, _FMT_PCT, _FMT_RATE, _FMT_X, _GRAY_FILL, _GREEN_FILL,
                     _HEADER_FONT, _LEFT, _LIGHT_FILL, _NOTE_FONT, _RED_FILL, _RIGHT, _SUBHDR_FILL,
                     _TOTAL_FILL, _TOTAL_FONT, _WHITE_FILL, _band, _cell, _hdr, _section)

DI = "'Developer Information'"
AB = "Absorption"
NP, NS = MAX_PRODUCTS, MAX_PHASES                 # 8 products, 10 series
PROD_ROW = {"name": DI_PROD_HDR_ROW + 1, "asp": DI_PROD_HDR_ROW + 2,
            "lotval": DI_PROD_HDR_ROW + 3, "pace": DI_PROD_HDR_ROW + 4}
PH0 = DI_PHASE_HDR_ROW + 1                        # first phase row on Developer Information (14)
PH1 = PH0 + NS - 1                                # last phase row (23)
LOT_COL0 = 7                                      # column G = lots for product 1


def pcol(p: int) -> str:
    """Developer Information column for product p (1-based): C..J."""
    return CL(DI_PROD_FIRST_COL + p - 1)


def lcol(p: int) -> str:
    """Developer Information lot column for product p: G..N."""
    return CL(LOT_COL0 + p - 1)


def S(k: int) -> str:
    return f"'Series {k}'"


UNITFMT = '0.000000;-0.000000;"-"'
INFL = "(1+INFLATION_ASP)^MAX(0,{y}-INFLATION_START+1)"

# ── Series sheet cell map ─────────────────────────────────────────────────────
R_NUM, R_AREA, R_ACTIVE, R_LABEL, R_LOTDEL, R_BONDDEL, R_COSTS, R_CAPIM, R_CAPIEND = range(6, 15)
R_FP, R_FINAL, R_N, R_RS, R_K, R_LASTCAPI, R_FPX, R_FINX = range(15, 23)
P_HDR = 25
P_NAME, P_LOTS, P_ASP, P_LOTVAL, P_GROSS, P_SHARE, P_ADMIN1, P_CAP, P_CAPTOT, P_ASSESS, \
    P_HOMES, P_VALUE, P_INST1, P_ETR = range(26, 40)
Z_HDR = 42
Z_CAPPAR, Z_VALUE, Z_VTLPAR, Z_LIMIT, Z_USEPCT, Z_COSTPAR, Z_PAR, Z_BIND, Z_SCALE, Z_VTL, \
    Z_URES, Z_UCAPI = range(43, 55)
T_HDR = 58
T0 = 59                                            # first schedule row (t = 1)
NT = 40                                            # schedule rows
T1 = T0 + NT - 1
U_HDR = T1 + 3
U_PAR, U_PF, U_RES, U_DP, U_CAPI, U_ADMIN, U_CITY, U_UWD, U_COI, U_TOT, U_CHK, U_COSTS, U_NET = \
    range(U_HDR + 1, U_HDR + 14)
K = "$C$"                                          # absolute column C prefix


def _label(ws, r, text, bold=False):
    c = ws.cell(row=r, column=2, value=text)
    c.font, c.border, c.alignment = (_BOLD_FONT if bold else _BODY_FONT), _BORDER, _LEFT
    return c


def _val(ws, r, formula, fmt=None, fill=_WHITE_FILL, font=_BODY_FONT, col=3):
    return _cell(ws, r, col, formula, fmt, fill, font)


def _title(ws, line3: str, last_col: int):
    for row, (txt, size, h) in enumerate([("=TITLE1", 14, 22),
                                          ('=CITY&", "&TITLE3&" County, Texas"', 11, 16),
                                          (line3, 11, 16)], 1):
        ws.row_dimensions[row].height = h
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=last_col)
        c = ws.cell(row=row, column=1, value=txt)
        c.fill, c.alignment = _BLUE_FILL, _CENTER
        c.font = Font(name="Calibri", bold=True, color="FFFFFF", size=size)
    ws.row_dimensions[4].height = 6
    ws.sheet_view.showGridLines = False


# ── Absorption ───────────────────────────────────────────────────────────────
AB_PH0, AB_M0, AB_NM = 7, 21, 480                  # phase rows 7-16; months start row 21
AB_M1 = AB_M0 + AB_NM - 1
AB_LOTS0 = 6                                       # F..M lots by product
AB_EARLY0 = 15                                     # O..V lots in earlier-opening areas
AB_HOMES0 = 24                                     # X..AE homes closed before bond delivery
AB_AVAIL0 = 3                                      # monthly: C..J lots available
AB_CUM0 = 11                                       # monthly: K..R cumulative closings


def _absorption(ws):
    _title(ws, "Absorption — Lots Available and Home Closings (Monthly)", AB_HOMES0 + NP - 1)
    ws.column_dimensions["A"].width = 11
    ws.column_dimensions["B"].width = 26
    for c in range(3, AB_HOMES0 + NP):
        ws.column_dimensions[CL(c)].width = 10
    _section(ws, 5, "Improvement areas: open month (lot delivery + MONTHS_LOT_TO_CLOSE), lots by product, "
                    "lots ahead in the FIFO queue, and homes closed before bond delivery", 1, AB_HOMES0 + NP - 1)
    heads = {1: "#", 2: "Improvement Area", 3: "Open Month", 4: "Bond Delivery", 5: "Lot Delivery Yr"}
    for p in range(1, NP + 1):
        heads[AB_LOTS0 + p - 1] = f"Lots P{p}"
        heads[AB_EARLY0 + p - 1] = f"Ahead P{p}"
        heads[AB_HOMES0 + p - 1] = f"Homes P{p}"
    ws.row_dimensions[6].height = 30
    for c, h in heads.items():
        _hdr(ws, 6, c, h)
    opn = f"$C${AB_PH0}:$C${AB_PH0 + NS - 1}"
    months = f"$A${AB_M0}:$A${AB_M1}"
    for i in range(NS):
        r, dr = AB_PH0 + i, PH0 + i
        f = _band(i)
        _cell(ws, r, 1, i + 1, fill=f, align=_CENTER)
        _cell(ws, r, 2, f'=IF({DI}!$B${dr}="","",{DI}!$B${dr})', fill=f, align=_LEFT)
        _cell(ws, r, 3, f'=IF($B{r}="",DATE(2999,1,1),DATE(YEAR(EDATE({DI}!$C${dr},MONTHS_LOT_TO_CLOSE)),'
                        f'MONTH(EDATE({DI}!$C${dr},MONTHS_LOT_TO_CLOSE)),1))', _FMT_DATE, f)
        _cell(ws, r, 4, f'=IF($B{r}="",DATE(2999,1,1),{DI}!$D${dr})', _FMT_DATE, f)
        _cell(ws, r, 5, f'=IF($B{r}="",9999,YEAR({DI}!$C${dr}))', "0", f)
        for p in range(1, NP + 1):
            lc = CL(AB_LOTS0 + p - 1)
            _cell(ws, r, AB_LOTS0 + p - 1, f'=IF($B{r}="",0,N({DI}!{lcol(p)}{dr}))', _FMT_NUM, f)
            lots = f"{lc}${AB_PH0}:{lc}${AB_PH0 + NS - 1}"
            _cell(ws, r, AB_EARLY0 + p - 1,
                  f"=SUMPRODUCT((({opn}<$C{r})+({opn}=$C{r})*(ROW({opn})<ROW($C{r})))*{lots})", _FMT_NUM, f)
            cum = f"{CL(AB_CUM0 + p - 1)}${AB_M0}:{CL(AB_CUM0 + p - 1)}${AB_M1}"
            before = (f"IFERROR(INT(INDEX({cum},MATCH(EDATE(DATE(YEAR($D{r}),MONTH($D{r}),1),"
                      f"IF(DAY($D{r})>1,0,-1)),{months},1))),0)")
            _cell(ws, r, AB_HOMES0 + p - 1,
                  f'=IF($B{r}="",0,MIN({lc}{r},MAX(0,{before}-{CL(AB_EARLY0 + p - 1)}{r})))', _FMT_NUM, f)

    _section(ws, AB_M0 - 2, "Monthly schedule — lots available (opened) and cumulative closings by product "
                            "(pace does not bank while no lots are open)", 1, AB_CUM0 + NP - 1)
    _hdr(ws, AB_M0 - 1, 1, "Month")
    _hdr(ws, AB_M0 - 1, 2, "Year")
    for p in range(1, NP + 1):
        _hdr(ws, AB_M0 - 1, AB_AVAIL0 + p - 1, f"Open Lots P{p}")
        _hdr(ws, AB_M0 - 1, AB_CUM0 + p - 1, f"Cum. Closed P{p}")
    for i in range(AB_NM):
        r = AB_M0 + i
        _cell(ws, r, 1, f"=MIN({opn})" if i == 0 else f"=EDATE(A{r - 1},1)", _FMT_DATE)
        _cell(ws, r, 2, f"=YEAR(A{r})", "0", align=_CENTER)
        for p in range(1, NP + 1):
            lc = CL(AB_LOTS0 + p - 1)
            av = CL(AB_AVAIL0 + p - 1)
            _cell(ws, r, AB_AVAIL0 + p - 1,
                  f"=SUMPRODUCT(({opn}<=$A{r})*{lc}${AB_PH0}:{lc}${AB_PH0 + NS - 1})", _FMT_NUM)
            prev = "0" if i == 0 else f"{CL(AB_CUM0 + p - 1)}{r - 1}"
            _cell(ws, r, AB_CUM0 + p - 1,
                  f"=MIN({av}{r},{prev}+N({DI}!{pcol(p)}${PROD_ROW['pace']}))", "#,##0.0")
    ws.freeze_panes = f"C{AB_M0}"


# ── Series sheets ────────────────────────────────────────────────────────────
def _series(ws, k: int):
    LAST = 20
    _title(ws, f'=C{R_LABEL}&" — "&IF(C{R_AREA}="","(not used)",C{R_AREA})', LAST)
    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 40
    for c in range(3, LAST + 1):
        ws.column_dimensions[CL(c)].width = 15
    dr = f"{PH0 + k - 1}"
    _section(ws, 5, "Series Information", 2, 11)
    rows = [
        (R_NUM, "Series Number", k, "0"),
        (R_AREA, "Improvement Area", f'=IF({DI}!$B${dr}="","",{DI}!$B${dr})', None),
        (R_ACTIVE, "Active (1 = yes)", f'=IF(OR(C{R_AREA}="",$K${P_LOTS}=0),0,1)', "0"),
        (R_LABEL, "Series Label", f'=IF(C{R_ACTIVE}=0,"Series {k} (not used)","Series "&YEAR(C{R_BONDDEL}))',
         None),
        (R_LOTDEL, "Lot Delivery", f'=IF(C{R_AREA}="",DATE(2000,6,1),{DI}!$C${dr})', _FMT_DATE),
        (R_BONDDEL, "Bond Delivery (Dated Date)", f'=IF(C{R_AREA}="",DATE(2000,9,1),{DI}!$D${dr})', _FMT_DATE),
        (R_COSTS, "Improvement Costs", f"=N({DI}!$E${dr})", _FMT_DOLLAR),
        (R_CAPIM, "Capitalized Interest (months)", f'=IF({DI}!$F${dr}="",CAPI_PERIOD,{DI}!$F${dr})', "0"),
        (R_CAPIEND, "Capitalized Interest End", f"=EDATE(C{R_BONDDEL},C{R_CAPIM})", _FMT_DATE),
        (R_FP, "First Principal (Bond Year)",
         f"=YEAR(C{R_CAPIEND})+IF(DATE(YEAR(C{R_CAPIEND}),PRIN_MATURITY,1)<=C{R_CAPIEND},1,0)", "0"),
        (R_FINAL, "Final Maturity (Bond Year)", f"=YEAR(C{R_BONDDEL})+FINAL_MAT_YRS", "0"),
        (R_N, "Number of Annual Installments", f"=C{R_FINAL}-C{R_FP}+1", "0"),
        (R_RS, "Sizing Rate (coupon + add'l interest + haircut)",
         "=BONDRATE+ADDITIONAL_INTEREST+RATE_ADD", _FMT_PCT),
        (R_K, "Revenue Factor (collection % / DSC)", "=TAX_COLLECTION_PRC/DSC_FACTOR", "0.0000"),
        (R_LASTCAPI, "Last Capitalized Interest Date",
         "=MAX(C{bd},{terms})".format(bd=R_BONDDEL, terms=",".join(
             f"IF(DATE(YEAR(C{R_CAPIEND}){d},{m},1)<=C{R_CAPIEND},DATE(YEAR(C{R_CAPIEND}){d},{m},1),0)"
             for d in ("", "-1") for m in ("PRIN_MATURITY", "INT_MATURITY"))), _FMT_DATE),
        (R_FPX, "First Bond Year (active only)", f"=IF(C{R_ACTIVE}=1,C{R_FP},9999)", "0"),
        (R_FINX, "Final Bond Year (active only)", f"=IF(C{R_ACTIVE}=1,C{R_FINAL},0)", "0"),
    ]
    for r, lab, f, fmt in rows:
        _label(ws, r, lab)
        _val(ws, r, f, fmt, font=_BOLD_FONT if r in (R_AREA, R_LABEL) else _BODY_FONT)

    # Product block
    _section(ws, P_HDR - 1, "Assessment Capacity by Lot Type (per lot unless noted)", 2, 11)
    _hdr(ws, P_HDR, 2, "Item")
    for p in range(1, NP + 1):
        _hdr(ws, P_HDR, 2 + p, f"Product {p}")
    _hdr(ws, P_HDR, 11, "Total")
    e = "ADMIN_ESCALATOR"
    prows = [
        (P_NAME, "Lot Type", lambda c, p: f'=IF({DI}!{pcol(p)}${PROD_ROW["name"]}="","",'
                                          f'{DI}!{pcol(p)}${PROD_ROW["name"]})', None, None),
        (P_LOTS, "Lots", lambda c, p: f"=N({DI}!{lcol(p)}${dr})", _FMT_NUM, "sum"),
        (P_ASP, "Average Selling Price", lambda c, p: f"=N({DI}!{pcol(p)}${PROD_ROW['asp']})", _FMT_DOLLAR, None),
        (P_LOTVAL, "As-Developed Lot Value", lambda c, p: f"=N({DI}!{pcol(p)}${PROD_ROW['lotval']})",
         _FMT_DOLLAR, None),
        (P_GROSS, "Gross Annual Installment (ASP/100 x rate)", lambda c, p: f"={c}{P_ASP}/100*PID_RATE",
         _FMT_DOLLAR2, None),
        (P_SHARE, "Share of Per-Series Admin Cost (per lot)",
         lambda c, p: f"=IF(SUMPRODUCT($C${P_LOTS}:$J${P_LOTS},$C${P_ASP}:$J${P_ASP})=0,0,"
                      f"{c}{P_ASP}/SUMPRODUCT($C${P_LOTS}:$J${P_LOTS},$C${P_ASP}:$J${P_ASP}))", "0.000000", None),
        (P_ADMIN1, "Year-1 Administrative Expense", lambda c, p: f"=ADMIN_FEE_FIXED_LOT+ADMIN_FEE_FIXED_SERIES*{c}{P_SHARE}",
         _FMT_DOLLAR2, None),
        (P_CAP, "Maximum Assessment (PV of net installment)",
         lambda c, p: (f"=IF(OR({c}{P_ASP}=0,{K}{R_RS}=0),0,{K}{R_K}*({c}{P_GROSS}*(1-(1+{K}{R_RS})^-{K}{R_N})/{K}{R_RS}"
                       f"-{c}{P_ADMIN1}*IF(ABS({K}{R_RS}-{e})<0.0000001,{K}{R_N}/(1+{K}{R_RS}),"
                       f"(1-((1+{e})/(1+{K}{R_RS}))^{K}{R_N})/({K}{R_RS}-{e}))))"), _FMT_DOLLAR2, None),
        (P_CAPTOT, "Total Assessment Capacity", lambda c, p: f"={c}{P_LOTS}*{c}{P_CAP}", _FMT_DOLLAR, "sum"),
        (P_ASSESS, "Assessment Levied per Lot", lambda c, p: f"={c}{P_CAP}*{K}{Z_SCALE}", _FMT_DOLLAR2, None),
        (P_HOMES, "Homes Closed before Bond Delivery",
         lambda c, p: f"=INDEX({AB}!{CL(AB_HOMES0 + p - 1)}${AB_PH0}:{CL(AB_HOMES0 + p - 1)}${AB_PH0 + NS - 1},{K}{R_NUM})",
         _FMT_NUM, "sum"),
        (P_VALUE, "Estimated Value at Bond Delivery",
         lambda c, p: f"={c}{P_HOMES}*{c}{P_ASP}*{INFL.format(y=f'YEAR({K}{R_BONDDEL})')}"
                      f"+({c}{P_LOTS}-{c}{P_HOMES})*{c}{P_LOTVAL}", _FMT_DOLLAR, "sum"),
        (P_INST1, "First Annual Installment",
         lambda c, p: (f'=IF({c}{P_LOTS}=0,"",{K}{Z_SCALE}*{K}{R_K}*MAX(0,{c}{P_GROSS}-{c}{P_ADMIN1})'
                       f"-{c}{P_ASSESS}*{K}{R_RS}+{c}{P_ASSESS}*(BONDRATE+ADDITIONAL_INTEREST)+{c}{P_ADMIN1})"),
         _FMT_DOLLAR2, None),
        (P_ETR, "Equivalent Tax Rate (first year)",
         lambda c, p: f'=IF(OR({c}{P_LOTS}=0,{c}{P_ASP}=0),"",{c}{P_INST1}/{c}{P_ASP}*100)', _FMT_RATE, None),
    ]
    for i, (r, lab, fn, fmt, tot) in enumerate(prows):
        f = _band(i)
        _cell(ws, r, 2, lab, fill=f, align=_LEFT)
        for p in range(1, NP + 1):
            c = CL(2 + p)
            _cell(ws, r, 2 + p, fn(c, p), fmt, f, align=_CENTER if r == P_NAME else _RIGHT)
        _cell(ws, r, 11, f"=SUM(C{r}:J{r})" if tot else None, fmt if tot else None,
              _TOTAL_FILL if tot else f, _TOTAL_FONT)

    # Sizing
    _section(ws, Z_HDR, "Bond Sizing — par is the smallest of capacity, cost need and value-to-lien limit", 2, 11)
    res_unit = _reserve_formula(f"$T${T0}:$T${T1}", f"{K}{R_N}", "1")
    zrows = [
        (Z_CAPPAR, "Assessment Capacity (rounded down)", f"=INT($K${P_CAPTOT}/ROUND_TO)*ROUND_TO", _FMT_DOLLAR),
        (Z_VALUE, "Estimated Value at Bond Delivery", f"=$K${P_VALUE}", _FMT_DOLLAR),
        (Z_VTLPAR, "Par Allowed by Minimum Value-to-Lien", f"=IF(MIN_VTL>0,C{Z_VALUE}/MIN_VTL,9.99E+307)",
         _FMT_DOLLAR),
        (Z_LIMIT, "Limit before Cost Test",
         f'=IF(AND(LEFT(SIZE_TO_VTL,1)="Y",C{Z_VTLPAR}<C{Z_CAPPAR}),INT(C{Z_VTLPAR}/ROUND_TO)*ROUND_TO,C{Z_CAPPAR})',
         _FMT_DOLLAR),
        (Z_USEPCT, "Non-Project Uses as % of Par",
         f"=C{Z_URES}+C{Z_UCAPI}+DP_RESERVE_DEPOSIT+ADMIN_FEE+CITY_FEE+UWD_FEE+COI_FEE", _FMT_PCT),
        (Z_COSTPAR, "Par Needed to Fund Improvement Costs",
         f'=IF(AND(LEFT(SIZE_TO_COSTS,1)="Y",C{R_COSTS}>0,C{Z_USEPCT}<1),'
         f'ROUNDUP(C{R_COSTS}/(1-C{Z_USEPCT})/ROUND_TO,0)*ROUND_TO,"")', _FMT_DOLLAR),
        (Z_PAR, "PAR AMOUNT", f'=IF(C{R_ACTIVE}=0,0,IF(C{Z_COSTPAR}="",C{Z_LIMIT},MIN(C{Z_LIMIT},C{Z_COSTPAR})))',
         _FMT_DOLLAR),
        (Z_BIND, "Binding Constraint",
         f'=IF(C{R_ACTIVE}=0,"",IF(AND(C{Z_COSTPAR}<>"",N(C{Z_COSTPAR})<C{Z_LIMIT}),"Improvement Costs",'
         f'IF(C{Z_LIMIT}<C{Z_CAPPAR},"Value-to-Lien","Assessment Capacity")))', None),
        (Z_SCALE, "Assessment Levied / Capacity", f"=IF($K${P_CAPTOT}=0,0,C{Z_PAR}/$K${P_CAPTOT})", "0.0000"),
        (Z_VTL, "Value-to-Lien at Delivery", f"=IF(C{Z_PAR}=0,0,C{Z_VALUE}/C{Z_PAR})", _FMT_X),
        (Z_URES, "Reserve Fund as % of Par (unit schedule)", f"=IF($K${P_CAPTOT}=0,0,{res_unit})", _FMT_PCT),
        (Z_UCAPI, "Capitalized Interest as % of Par",
         f"=IF(C{R_LASTCAPI}>C{R_BONDDEL},BONDRATE*DAYS360(C{R_BONDDEL},C{R_LASTCAPI})/360,0)", _FMT_PCT),
    ]
    for r, lab, f, fmt in zrows:
        big = r in (Z_PAR, Z_BIND)
        _label(ws, r, lab, bold=big)
        _val(ws, r, f, fmt, _TOTAL_FILL if big else _WHITE_FILL, _TOTAL_FONT if big else _BODY_FONT)

    # Annual schedule
    _section(ws, T_HDR - 1, "Annual Schedule — bond year ending September 1 (bonds, installments, unit schedule "
                            "for sizing)", 1, LAST)
    heads = ["t", "Bond Year", "Net Revenue for P&I", "Assessment Principal (unrounded)", "Assessment Balance",
             "Bond Principal", "Bonds Outstanding", "Bond Interest", "Capitalized Interest", "Total Debt Service",
             "Net Debt Service", "Installment: Principal", "Installment: Interest", "Installment: Add'l Interest",
             "Installment: Admin", "Total Annual Installments", "P&I Coverage",
             "Unit Principal", "Unit Balance", "Unit Debt Service"]
    ws.row_dimensions[T_HDR].height = 42
    for i, h in enumerate(heads):
        _hdr(ws, T_HDR, 1 + i, h)
    par, n, rs = f"{K}{Z_PAR}", f"{K}{R_N}", f"{K}{R_RS}"
    fp, bd = f"{K}{R_FP}", f"{K}{R_BONDDEL}"
    start_fp = f"MAX({bd},DATE({fp}-1,PRIN_MATURITY,1))"
    for i in range(NT):
        r, t = T0 + i, i + 1
        f = _band(i)
        pe = par if i == 0 else f"E{r - 1}"          # previous assessment balance
        pg = par if i == 0 else f"G{r - 1}"          # previous bonds outstanding
        ps = "1" if i == 0 else f"S{r - 1}"          # previous unit balance
        frac = f"DAYS360(MAX({bd},DATE(B{r}-1,PRIN_MATURITY,1)),DATE(B{r},PRIN_MATURITY,1))/360"
        cells = {
            1: (t, "0"),
            2: (f'=IF(A{r}<={n},{fp}+A{r}-1,"")', "0"),
            3: (f'=IF(B{r}="",0,{K}{R_K}*SUMPRODUCT($C${P_LOTS}:$J${P_LOTS},$C${P_GROSS}:$J${P_GROSS}'
                f'-$C${P_ADMIN1}:$J${P_ADMIN1}*(1+ADMIN_ESCALATOR)^(A{r}-1)))', _FMT_DOLLAR),
            4: (f'=IF(B{r}="",0,IF(A{r}={n},{pe},{K}{Z_SCALE}*C{r}-{pe}*{rs}))', _FMT_DOLLAR2),
            5: (f"={pe}-D{r}", _FMT_DOLLAR2),
            6: (f'=IF(B{r}="",0,IF(A{r}={n},{par}-SUM(F${T0 - 1}:F{r - 1}),MAX(0,ROUND(D{r}/ROUND_TO,0)*ROUND_TO)))'
                if i else
                f'=IF(B{r}="",0,IF(A{r}={n},{par},MAX(0,ROUND(D{r}/ROUND_TO,0)*ROUND_TO)))', _FMT_DOLLAR),
            7: (f"={par}-SUM(F${T0}:F{r})", _FMT_DOLLAR),
            8: (f'=IF(B{r}="",0,{pg}*BONDRATE*{frac})', _FMT_DOLLAR2),
            9: ((f"=IF(C{R_LASTCAPI}>{start_fp},{par}*BONDRATE*DAYS360({start_fp},C{R_LASTCAPI})/360,0)"
                 .replace(f"C{R_LASTCAPI}", f"{K}{R_LASTCAPI}") if i == 0 else 0), _FMT_DOLLAR2),
            10: (f"=F{r}+H{r}", _FMT_DOLLAR2),
            11: (f"=J{r}-I{r}", _FMT_DOLLAR2),
            12: (f"=D{r}", _FMT_DOLLAR2),
            13: (f'=IF(B{r}="",0,{pe}*BONDRATE)', _FMT_DOLLAR2),
            14: (f'=IF(B{r}="",0,{pe}*ADDITIONAL_INTEREST)', _FMT_DOLLAR2),
            15: (f'=IF(B{r}="",0,SUMPRODUCT($C${P_LOTS}:$J${P_LOTS},$C${P_ADMIN1}:$J${P_ADMIN1})'
                 f'*(1+ADMIN_ESCALATOR)^(A{r}-1))', _FMT_DOLLAR2),
            16: (f"=SUM(L{r}:O{r})", _FMT_DOLLAR2),
            17: (f'=IF(K{r}=0,"",(L{r}+M{r})*TAX_COLLECTION_PRC/K{r})', _FMT_X),
            18: (f'=IF(OR(B{r}="",$K${P_CAPTOT}=0),0,IF(A{r}={n},{ps},C{r}/$K${P_CAPTOT}-{ps}*{rs}))', UNITFMT),
            19: (f"={ps}-R{r}", UNITFMT),
            20: (f'=IF(B{r}="",0,R{r}+{ps}*BONDRATE*{frac})', UNITFMT),
        }
        for c, (v, fmt) in cells.items():
            _cell(ws, r, c, v, fmt, f, align=_CENTER if c <= 2 else _RIGHT)
    tr = T1 + 1
    _cell(ws, tr, 1, "Total", fill=_TOTAL_FILL, font=_TOTAL_FONT, align=_CENTER)
    for c in range(2, LAST + 1):
        L = CL(c)
        sumcols = {3, 4, 6, 8, 9, 10, 11, 12, 13, 14, 15, 16}
        _cell(ws, tr, c, f"=SUM({L}{T0}:{L}{T1})" if c in sumcols else None,
              _FMT_DOLLAR if c in sumcols else None, _TOTAL_FILL, _TOTAL_FONT)

    # Sources & uses
    _section(ws, U_HDR, "Sources and Uses of Funds", 2, 4)
    res_act = _reserve_formula(f"$J${T0}:$J${T1}", f"{K}{R_N}", par)
    urows = [
        (U_PAR, "Par Amount", f"={par}"),
        (U_PF, "Deposit to Project Fund (Developer Reimbursement)", f"=C{U_PAR}-SUM(C{U_RES}:C{U_COI})"),
        (U_RES, f"Deposit to Reserve Fund", f"=IF({par}=0,0,{res_act})"),
        (U_DP, "Deposit to Delinquency & Prepayment Reserve", f"={par}*DP_RESERVE_DEPOSIT"),
        (U_CAPI, "Deposit to Capitalized Interest Account",
         f"=IF(C{R_LASTCAPI}>C{R_BONDDEL},{par}*BONDRATE*DAYS360(C{R_BONDDEL},C{R_LASTCAPI})/360,0)"),
        (U_ADMIN, "Deposit to Administrative Fund", f"={par}*ADMIN_FEE"),
        (U_CITY, "City Development / PID Fee", f"={par}*CITY_FEE"),
        (U_UWD, "Underwriter's Discount", f"={par}*UWD_FEE"),
        (U_COI, "Costs of Issuance", f"={par}*COI_FEE"),
        (U_TOT, "Total Uses", f"=SUM(C{U_PF}:C{U_COI})"),
        (U_CHK, "Sources less Uses", f"=C{U_PAR}-C{U_TOT}"),
        (U_COSTS, "Improvement Costs", f"=C{R_COSTS}"),
        (U_NET, "Project Fund less Costs", f"=C{U_PF}-C{U_COSTS}"),
    ]
    for i, (r, lab, f) in enumerate(urows):
        tot = r in (U_PAR, U_TOT, U_CHK, U_NET)
        _label(ws, r, lab, bold=tot)
        _val(ws, r, f, _FMT_DOLLAR, _TOTAL_FILL if tot else _band(i), _TOTAL_FONT if tot else _BODY_FONT)
        _cell(ws, r, 4, f"=IF({par}=0,0,C{r}/{par})" if r <= U_TOT else None, _FMT_PCT,
              _TOTAL_FILL if tot else _band(i))
    ws.freeze_panes = "C5"
    for r in range(T0, T1 + 1):
        ws.row_dimensions[r].outlineLevel = 0


def _reserve_formula(rng: str, n: str, par: str) -> str:
    avg = f"1.25*SUM({rng})/{n}"
    mx = f"MAX({rng})"
    ten = f"0.1*{par}"
    return (f'IF(LEFT(RESERVE_METHOD,3)="Max",{mx},IF(LEFT(RESERVE_METHOD,3)="10%",{ten},'
            f'IF(LEFT(RESERVE_METHOD,3)="Les",MIN({avg},{mx},{ten}),{avg})))')


# ── Output tabs ──────────────────────────────────────────────────────────────
SU_ITEMS = [("Par Amount", U_PAR), ("Deposit to Project Fund (Developer Reimbursement)", U_PF),
            ("Deposit to Reserve Fund", U_RES), ("Deposit to Delinquency & Prepayment Reserve", U_DP),
            ("Deposit to Capitalized Interest Account", U_CAPI), ("Deposit to Administrative Fund", U_ADMIN),
            ("City Development / PID Fee", U_CITY), ("Underwriter's Discount", U_UWD),
            ("Costs of Issuance", U_COI), ("Total Uses", U_TOT), ("Sources less Uses", U_CHK),
            ("Improvement Costs", U_COSTS), ("Project Fund less Costs", U_NET)]


def _sources_uses(ws):
    TOT, PCT = 2 + NS, 3 + NS
    _title(ws, "=\"Sources and Uses of Funds — \"&BOND_TITLE", PCT)
    ws.column_dimensions["A"].width = 46
    for c in range(2, PCT + 1):
        ws.column_dimensions[CL(c)].width = 15
    ws.row_dimensions[5].height = 46
    _hdr(ws, 5, 1, "Item")
    for k in range(1, NS + 1):
        c = ws.cell(row=5, column=1 + k,
                    value=f'=IF({S(k)}!$C${R_ACTIVE}=0,"Not used",{S(k)}!$C${R_LABEL}&CHAR(10)&{S(k)}!$C${R_AREA}'
                          f'&CHAR(10)&TEXT({S(k)}!$C${R_BONDDEL},"mm/dd/yyyy"))')
        c.fill, c.alignment, c.border = _BLUE_FILL, _CENTER, _BORDER
        c.font = Font(name="Calibri", bold=True, color="FFFFFF", size=9)
    for col, txt in ((TOT, "Total"), (PCT, "% of Par")):
        c = ws.cell(row=5, column=col, value=txt)
        c.fill, c.alignment, c.border = _BLUE_FILL, _CENTER, _BORDER
        c.font = Font(name="Calibri", bold=True, color="FFFFFF", size=9)
    for i, (lab, row) in enumerate(SU_ITEMS):
        r = 6 + i
        tot = row in (U_PAR, U_TOT, U_CHK, U_NET)
        fill, font = (_TOTAL_FILL, _TOTAL_FONT) if tot else (_band(i), _BODY_FONT)
        _cell(ws, r, 1, lab, fill=fill, font=font, align=_LEFT)
        for k in range(1, NS + 1):
            _cell(ws, r, 1 + k, f"={S(k)}!$C${row}", _FMT_DOLLAR, fill, font)
        _cell(ws, r, TOT, f"=SUM(B{r}:{CL(TOT - 1)}{r})", _FMT_DOLLAR, fill, _TOTAL_FONT)
        _cell(ws, r, PCT, f"=IF(${CL(TOT)}$6=0,0,{CL(TOT)}{r}/${CL(TOT)}$6)" if row <= U_TOT else None,
              _FMT_PCT, fill, font)
    ws.cell(row=6 + len(SU_ITEMS) + 1, column=1,
            value="Project Fund is the residual after all other uses — the estimated reimbursement to the "
                  "developer. Bonds assumed priced at par.").font = _NOTE_FONT
    ws.freeze_panes = "B6"


def _bond_sizing(ws):
    cols = 1 + 6 + NS + 2
    _title(ws, '="Bond Sizing Analysis — Annual Installments vs. Debt Service (ETR $"&TEXT(PID_RATE,"0.0000")&")"',
           cols)
    heads = (["Bond Year (9/1)", "Assessed Lots", "Installment: Principal", "Installment: Interest",
              "Additional Interest", "Administrative Expenses", "Total Annual Installments"]
             + [f"=IF({S(k)}!$C${R_ACTIVE}=0,\"Not used\",{S(k)}!$C${R_LABEL}&CHAR(10)&\"Net Debt Service\")"
                for k in range(1, NS + 1)]
             + ["Total Net Debt Service", "P&I Coverage"])
    ws.row_dimensions[5].height = 40
    for i, h in enumerate(heads):
        _hdr(ws, 5, 1 + i, h, 12 if i < 2 else 14)
    first = "MIN(" + ",".join(f"{S(k)}!$C${R_FPX}" for k in range(1, NS + 1)) + ")"
    last = "MAX(" + ",".join(f"{S(k)}!$C${R_FINX}" for k in range(1, NS + 1)) + ")"
    NR = 60
    for i in range(NR):
        r = 6 + i
        f = _band(i)
        yr = f"$A{r}"
        _cell(ws, r, 1, f'=IF({first}+{i}>{last},"",{first}+{i})', "0", f, _BOLD_FONT, _CENTER)
        _cell(ws, r, 2, "=" + "+".join(f"IF(AND(N({yr})>={S(k)}!$C${R_FPX},N({yr})<={S(k)}!$C${R_FINX}),"
                                       f"{S(k)}!$K${P_LOTS},0)" for k in range(1, NS + 1)), _FMT_NUM, f)
        for j, col in enumerate("LMNO"):
            _cell(ws, r, 3 + j, "=" + "+".join(f"SUMIF({S(k)}!$B${T0}:$B${T1},{yr},{S(k)}!${col}${T0}:${col}${T1})"
                                               for k in range(1, NS + 1)), _FMT_DOLLAR, f)
        _cell(ws, r, 7, f"=SUM(C{r}:F{r})", _FMT_DOLLAR, f)
        for k in range(1, NS + 1):
            _cell(ws, r, 7 + k, f"=SUMIF({S(k)}!$B${T0}:$B${T1},{yr},{S(k)}!$K${T0}:$K${T1})", _FMT_DOLLAR, f)
        a, b = CL(8), CL(7 + NS)
        _cell(ws, r, 8 + NS, f"=SUM({a}{r}:{b}{r})", _FMT_DOLLAR, f)
        tn = CL(8 + NS)
        _cell(ws, r, 9 + NS, f'=IF({tn}{r}=0,"",(C{r}+D{r})*TAX_COLLECTION_PRC/{tn}{r})', _FMT_X, f)
    tr = 6 + NR
    _cell(ws, tr, 1, "Total", fill=_TOTAL_FILL, font=_TOTAL_FONT, align=_CENTER)
    _cell(ws, tr, 2, None, fill=_TOTAL_FILL)
    for c in range(3, 9 + NS):
        _cell(ws, tr, c, f"=SUM({CL(c)}6:{CL(c)}{tr - 1})", _FMT_DOLLAR, _TOTAL_FILL, _TOTAL_FONT)
    _cell(ws, tr, 9 + NS, None, fill=_TOTAL_FILL)
    ws.cell(row=tr + 2, column=1, value="Installments are billed each October, due January 31, and pay the "
            "March 1 and September 1 debt service of that bond year. Debt service is net of capitalized "
            "interest.").font = _NOTE_FONT
    ws.freeze_panes = "B6"


ROLL0 = 6
ROLL1 = ROLL0 + NS * NP - 1


def _roll(ws):
    cols = ["Series / Improvement Area", "Lot Type", "Lots", "Average Selling Price", "Assessment Per Lot",
            "Total Assessment", "Assessment / ASP", "First Annual Installment Per Lot", "Equivalent Tax Rate",
            "Total First-Year Installments"]
    _title(ws, "Assessment Roll Summary", len(cols))
    ws.row_dimensions[5].height = 40
    for i, (h, w) in enumerate(zip(cols, [36, 16, 8, 14, 14, 16, 12, 14, 12, 16])):
        _hdr(ws, 5, 1 + i, h, w)
    i = 0
    for k in range(1, NS + 1):
        for p in range(1, NP + 1):
            r = ROLL0 + i
            c = CL(2 + p)
            f = _band(k)
            sh = S(k)
            on = f"{sh}!${c}${P_LOTS}=0"
            vals = [
                (f'=IF({on},"",{sh}!$C${R_LABEL}&" — "&{sh}!$C${R_AREA})', None, _LEFT),
                (f'=IF({on},"",{sh}!${c}${P_NAME})', None, _LEFT),
                (f'=IF({on},"",{sh}!${c}${P_LOTS})', _FMT_NUM, _RIGHT),
                (f'=IF({on},"",{sh}!${c}${P_ASP})', _FMT_DOLLAR, _RIGHT),
                (f'=IF({on},"",{sh}!${c}${P_ASSESS})', _FMT_DOLLAR, _RIGHT),
                (f'=IF({on},"",C{r}*E{r})', _FMT_DOLLAR, _RIGHT),
                (f'=IF({on},"",IF(D{r}=0,0,E{r}/D{r}))', _FMT_PCT, _RIGHT),
                (f'=IF({on},"",{sh}!${c}${P_INST1})', _FMT_DOLLAR, _RIGHT),
                (f'=IF({on},"",{sh}!${c}${P_ETR})', _FMT_RATE, _RIGHT),
                (f'=IF({on},"",C{r}*H{r})', _FMT_DOLLAR, _RIGHT),
            ]
            for j, (v, fmt, al) in enumerate(vals):
                _cell(ws, r, 1 + j, v, fmt, f, align=al)
            i += 1
    tr = ROLL1 + 1
    _cell(ws, tr, 1, "Total", fill=_TOTAL_FILL, font=_TOTAL_FONT, align=_LEFT)
    for c in range(2, 11):
        s = c in (3, 6, 10)
        _cell(ws, tr, c, f"=SUM({CL(c)}{ROLL0}:{CL(c)}{ROLL1})" if s else None,
              (_FMT_NUM if c == 3 else _FMT_DOLLAR) if s else None, _TOTAL_FILL, _TOTAL_FONT)
    ws.cell(row=tr + 2, column=1, value="Blank rows are unused lot types / series. Total assessment ties to par "
            "by series; every lot type carries the same equivalent tax rate.").font = _NOTE_FONT
    ws.freeze_panes = "C6"


def _installments(ws):
    _title(ws, "Estimated Annual Installment Schedule (Per Lot) — choose a series and lot type", 8)
    for col, w in {1: 8, 2: 34, 3: 16, 4: 14, 5: 14, 6: 15, 7: 15, 8: 16}.items():
        ws.column_dimensions[CL(col)].width = w
    g = lambda cell: f'INDIRECT("\'Series "&$C$5&"\'!"&{cell})'          # noqa: E731
    pc = 'CHAR(66+$C$6)'
    rows = [
        (5, "Series Number (1-10)", 1, "0", True),
        (6, "Lot Type Number (1-8)", 1, "0", True),
        (7, "Series", f'={g(chr(34) + "C" + str(R_LABEL) + chr(34))}&" — "&{g(chr(34) + "C" + str(R_AREA) + chr(34))}',
         None, False),
        (8, "Lot Type", f"={g(pc + '&' + chr(34) + str(P_NAME) + chr(34))}", None, False),
        (9, "Average Selling Price", f"={g(pc + '&' + chr(34) + str(P_ASP) + chr(34))}", _FMT_DOLLAR, False),
        (10, "Assessment per Lot", f"={g(pc + '&' + chr(34) + str(P_ASSESS) + chr(34))}", _FMT_DOLLAR2, False),
    ]
    for r, lab, v, fmt, inp in rows:
        _cell(ws, r, 2, lab, fill=_WHITE_FILL, font=_BOLD_FONT, align=_LEFT)
        from openpyxl.styles import PatternFill
        _cell(ws, r, 3, v, fmt, PatternFill("solid", fgColor="FFF2CC") if inp else _WHITE_FILL,
              _BOLD_FONT, _RIGHT)
    for cell, n in (("C5", NS), ("C6", NP)):
        dv = DataValidation(type="whole", operator="between", formula1="1", formula2=str(n))
        ws.add_data_validation(dv)
        dv.add(cell)
    # hidden helpers: n, first year, scale, k, rs, gross, admin1
    helpers = [("n", f'{g(chr(34) + "C" + str(R_N) + chr(34))}'), ("fp", f'{g(chr(34) + "C" + str(R_FP) + chr(34))}'),
               ("scale", f'{g(chr(34) + "C" + str(Z_SCALE) + chr(34))}'), ("k", f'{g(chr(34) + "C" + str(R_K) + chr(34))}'),
               ("rs", f'{g(chr(34) + "C" + str(R_RS) + chr(34))}'),
               ("gross", f"{g(pc + '&' + chr(34) + str(P_GROSS) + chr(34))}"),
               ("admin1", f"{g(pc + '&' + chr(34) + str(P_ADMIN1) + chr(34))}")]
    hcol = {}
    for j, (nm, f) in enumerate(helpers):
        col = 10 + j
        ws.cell(row=5, column=col, value=nm).font = _NOTE_FONT
        ws.cell(row=6, column=col, value="=" + f).font = _NOTE_FONT
        hcol[nm] = f"${CL(col)}$6"
    _section(ws, 11, "Per-lot schedule", 1, 8)
    ws.row_dimensions[12].height = 30
    for i, h in enumerate(["Year", "Payment Date", "Principal", "Interest", "Additional Interest",
                           "Administrative Expenses", "Annual Installment", "Principal Balance"]):
        _hdr(ws, 12, 1 + i, h)
    for i in range(NT):
        r = 13 + i
        f = _band(i)
        pb = "$C$10" if i == 0 else f"H{r - 1}"
        on = f"A{r}<={hcol['n']}"
        adm = f"{hcol['admin1']}*(1+ADMIN_ESCALATOR)^(A{r}-1)"
        _cell(ws, r, 1, i + 1, "0", f, align=_CENTER)
        _cell(ws, r, 2, f'=IF({on},DATE({hcol["fp"]}+A{r}-1,PRIN_MATURITY,1),"")', _FMT_DATE, f, align=_CENTER)
        _cell(ws, r, 3, f'=IF({on},IF(A{r}={hcol["n"]},{pb},{hcol["scale"]}*{hcol["k"]}*MAX(0,{hcol["gross"]}-{adm})'
                        f'-{pb}*{hcol["rs"]}),"")', _FMT_DOLLAR2, f)
        _cell(ws, r, 4, f'=IF({on},{pb}*BONDRATE,"")', _FMT_DOLLAR2, f)
        _cell(ws, r, 5, f'=IF({on},{pb}*ADDITIONAL_INTEREST,"")', _FMT_DOLLAR2, f)
        _cell(ws, r, 6, f'=IF({on},{adm},"")', _FMT_DOLLAR2, f)
        _cell(ws, r, 7, f'=IF({on},SUM(C{r}:F{r}),"")', _FMT_DOLLAR2, f)
        _cell(ws, r, 8, f'=IF({on},{pb}-C{r},"")', _FMT_DOLLAR2, f)
    tr = 13 + NT
    _cell(ws, tr, 1, "Total", fill=_TOTAL_FILL, font=_TOTAL_FONT, align=_CENTER)
    for c in range(2, 9):
        _cell(ws, tr, c, f"=SUM({CL(c)}13:{CL(c)}{tr - 1})" if 3 <= c <= 7 else None, _FMT_DOLLAR,
              _TOTAL_FILL, _TOTAL_FONT)
    ws.freeze_panes = "A13"


def _project(ws):
    NY = 40
    cols = (["Year No.", "Year"] + [f"=IF({DI}!{pcol(p)}${PROD_ROW['name']}=\"\",\"Product {p}\","
                                    f"{DI}!{pcol(p)}${PROD_ROW['name']})&\" Homes\"" for p in range(1, NP + 1)]
            + [f"=IF({DI}!{pcol(p)}${PROD_ROW['name']}=\"\",\"Product {p}\","
               f"{DI}!{pcol(p)}${PROD_ROW['name']})&\" Cumulative\"" for p in range(1, NP + 1)]
            + ["Homes Per Year", "Cumulative Homes", "Home Value", "Vacant Lot Value", "Total Taxable Basis",
               "Tax Roll Year"])
    _title(ws, "Project and PID Revenue Analysis — Absorption and Taxable Value", len(cols))
    ws.row_dimensions[5].height = 40
    for i, h in enumerate(cols):
        _hdr(ws, 5, 1 + i, h, 9 if i < 2 else 12)
    first = f"MIN({AB}!$E${AB_PH0}:$E${AB_PH0 + NS - 1})"
    months = f"{AB}!$A${AB_M0}:$A${AB_M1}"
    for i in range(NY):
        r = 6 + i
        f = _band(i)
        show = f"$A{r}<=PROJECTION_YEARS"
        _cell(ws, r, 1, i + 1, "0", f, align=_CENTER)
        _cell(ws, r, 2, f'=IF({show},{first}+{i},"")', "0", f, align=_CENTER)
        for p in range(1, NP + 1):
            cumc = CL(AB_CUM0 + p - 1)
            cum = (f'=IF({show},IFERROR(INT(INDEX({AB}!${cumc}${AB_M0}:${cumc}${AB_M1},'
                   f'MATCH(DATE($B{r},12,1),{months},1))),0),0)')
            hc, cc = 2 + p, 2 + NP + p
            _cell(ws, r, cc, cum, _FMT_NUM, f)
            prev = "0" if i == 0 else f"{CL(cc)}{r - 1}"
            _cell(ws, r, hc, f"={CL(cc)}{r}-{prev}", _FMT_NUM, f)
        hc0, hc1 = CL(3), CL(2 + NP)
        cc0, cc1 = CL(3 + NP), CL(2 + 2 * NP)
        base = 3 + 2 * NP
        _cell(ws, r, base, f"=SUM({hc0}{r}:{hc1}{r})", _FMT_NUM, f)
        _cell(ws, r, base + 1, f"=SUM({cc0}{r}:{cc1}{r})", _FMT_NUM, f)
        _cell(ws, r, base + 2, f'=IF({show},SUMPRODUCT({cc0}{r}:{cc1}{r},{DI}!$C${PROD_ROW["asp"]}:$J${PROD_ROW["asp"]})'
                               f'*{INFL.format(y=f"$B{r}")},0)', _FMT_DOLLAR, f)
        vac = "+".join(
            f"(SUMPRODUCT(({AB}!$E${AB_PH0}:$E${AB_PH0 + NS - 1}<=$B{r})*{AB}!{CL(AB_LOTS0 + p - 1)}${AB_PH0}:"
            f"{CL(AB_LOTS0 + p - 1)}${AB_PH0 + NS - 1})-{CL(2 + NP + p)}{r})*N({DI}!{pcol(p)}${PROD_ROW['lotval']})"
            for p in range(1, NP + 1))
        _cell(ws, r, base + 3, f'=IF({show},{vac},0)', _FMT_DOLLAR, f)
        _cell(ws, r, base + 4, f"={CL(base + 2)}{r}+{CL(base + 3)}{r}", _FMT_DOLLAR, f)
        _cell(ws, r, base + 5, f'=IF({show},$B{r}+AV_LAG_YEARS,"")', "0", f, align=_CENTER)
    ws.cell(row=6 + NY + 1, column=1, value="Homes close MONTHS_LOT_TO_CLOSE after each lot delivery at the "
            "product pace, oldest area first; taxable basis = closed homes at ASP plus delivered, unbuilt lots "
            "at lot value. Monthly detail on the Absorption tab (40-year horizon).").font = _NOTE_FONT
    ws.freeze_panes = "C6"


def _tests(ws):
    _title(ws, "Tax Rate and Value-to-Lien Tests", 6)
    for col, w in {1: 40, 2: 14, 3: 16, 4: 16, 5: 14, 6: 16}.items():
        ws.column_dimensions[CL(col)].width = w
    OT = "'Overlapping Taxes'"
    end = 5 + MAX_TAXES
    r = 5
    _section(ws, r, "Overlapping Tax Rates (per $100 of value)", 1, 6)
    r += 1
    for c, h in ((1, "Taxing Entity"), (2, "Type"), (3, "Tax Rate")):
        _hdr(ws, r, c, h)
    for i in range(MAX_TAXES):
        r += 1
        src = 5 + i
        _cell(ws, r, 1, f'=IF({OT}!$B${src}="","",{OT}!$B${src})', fill=_band(i), align=_LEFT)
        _cell(ws, r, 2, f'=IF({OT}!$C${src}="","",{OT}!$C${src})', fill=_band(i), align=_CENTER)
        _cell(ws, r, 3, f'=IF({OT}!$D${src}="","",{OT}!$D${src})', _FMT_RATE, _band(i))
    for lab, ref in (("Total Overlapping Tax Rate", end), ("PID Equivalent Tax Rate", end + 3),
                     ("Total Tax Rate incl. PID", end + 4), ("Maximum Total Tax Rate (Policy)", end + 1),
                     ("Within Cap?", end + 5)):
        r += 1
        _cell(ws, r, 1, lab, fill=_TOTAL_FILL, font=_TOTAL_FONT, align=_LEFT)
        _cell(ws, r, 2, None, fill=_TOTAL_FILL)
        _cell(ws, r, 3, f"={OT}!$D${ref}", None if lab == "Within Cap?" else _FMT_RATE, _TOTAL_FILL, _TOTAL_FONT)
    r += 3
    _section(ws, r, '="Value-to-Lien at Bond Delivery (minimum "&TEXT(MIN_VTL,"0.00")&"x)"', 1, 6)
    r += 1
    ws.row_dimensions[r].height = 30
    for i, h in enumerate(["Series / Improvement Area", "Lots", "Estimated Value at Delivery", "Par Amount",
                           "Value-to-Lien", "Passes?"]):
        _hdr(ws, r, 1 + i, h)
    for k in range(1, NS + 1):
        r += 1
        f = _band(k)
        sh, on = S(k), f"{S(k)}!$C${R_ACTIVE}=0"
        _cell(ws, r, 1, f'=IF({on},"",{sh}!$C${R_LABEL}&" — "&{sh}!$C${R_AREA})', fill=f, align=_LEFT)
        _cell(ws, r, 2, f'=IF({on},"",{sh}!$K${P_LOTS})', _FMT_NUM, f)
        _cell(ws, r, 3, f'=IF({on},"",{sh}!$C${Z_VALUE})', _FMT_DOLLAR, f)
        _cell(ws, r, 4, f'=IF({on},"",{sh}!$C${Z_PAR})', _FMT_DOLLAR, f)
        _cell(ws, r, 5, f'=IF({on},"",{sh}!$C${Z_VTL})', _FMT_X, f)
        _cell(ws, r, 6, f'=IF({on},"",IF({sh}!$C${Z_VTL}>=MIN_VTL,"Yes","No"))', None, f, align=_CENTER)
    ws.cell(row=r + 2, column=1, value="Value at delivery = closed homes at ASP plus remaining lots at the "
            "as-developed lot value. Replace with the appraisal once available.").font = _NOTE_FONT


def _homeowner(ws):
    LAST = 2 + NP
    _title(ws, "Homeowner Cost Comparison — PID vs. No PID", LAST)
    ws.column_dimensions["A"].width = 52
    ws.column_dimensions["B"].width = 20
    for p in range(1, NP + 1):
        ws.column_dimensions[CL(2 + p)].width = 14
    ws.row_dimensions[5].height = 30
    _hdr(ws, 5, 1, "Item")
    _hdr(ws, 5, 2, "Reference")
    for p in range(1, NP + 1):
        _hdr(ws, 5, 2 + p, f'=IF({DI}!{pcol(p)}${PROD_ROW["name"]}="","",{DI}!{pcol(p)}${PROD_ROW["name"]})')
    SU = "'Sources and Uses'"
    TOTC = CL(2 + NS)
    rb = f"'Assessment Roll'!$B${ROLL0}:$B${ROLL1}"
    rcol = lambda c: f"'Assessment Roll'!${c}${ROLL0}:${c}${ROLL1}"     # noqa: E731
    spec = [
        ("Assumptions", None, None, None),
        ("PID Term (years)", "Years", "=FINAL_MAT_YRS", _FMT_NUM),
        ("PID Equivalent Tax Rate", "A", "=PID_RATE", _FMT_RATE),
        ("Ad Valorem Tax Rate before PID (City, County, ISD, etc.)", "B",
         f"='Overlapping Taxes'!$D${5 + MAX_TAXES}", _FMT_RATE),
        ("Homeowner Mortgage Rate", "Rate", "=MORTGAGE_RATE", _FMT_PCT),
        ("Mortgage Term (years)", "Term", "=MORTGAGE_TERM", _FMT_NUM),
        ("Mortgage Loan as % of Home Price", "C", "=MORTGAGE_LTV", _FMT_PCT),
        ("Net Proceeds as % of Gross PID Bonds", "D-1", f"=IF({SU}!${TOTC}$6=0,0,{SU}!${TOTC}$7/{SU}!${TOTC}$6)",
         _FMT_PCT),
        ("Assessment Per Lot (weighted across series)", "Assess",
         f"=IF(SUMIF({rb},{{h}},{rcol('C')})=0,0,SUMIF({rb},{{h}},{rcol('F')})/SUMIF({rb},{{h}},{rcol('C')}))",
         _FMT_DOLLAR),
        ("PID-Financed Infrastructure per Lot", "D = Assess x D-1", "={c}{Assess}*{c}{D-1}", _FMT_DOLLAR),
        ("Home Price as Multiple of Lot Cost", "E", "=HOME_PRICE_MULTIPLE", "0.00"),
        ("PID Scenario", None, None, None),
        ("Average Home Price", "F", f"=N({DI}!{{pc}}${PROD_ROW['asp']})", _FMT_DOLLAR),
        ("Lot Price assuming PID", "G = F / E", "=IF({c}{E}=0,0,{c}{F}/{c}{E})", _FMT_DOLLAR),
        ("Mortgage with PID", "H = C x F", "={c}{C}*{c}{F}", _FMT_DOLLAR),
        ("Down Payment with PID", "I = F - H", "={c}{F}-{c}{H}", _FMT_DOLLAR),
        ("Monthly Mortgage Payment", "J", "=IF({c}{H}=0,0,PMT({c}{Rate}/12,{c}{Term}*12,-{c}{H}))", _FMT_DOLLAR),
        ("Annual Mortgage Payment", "M = 12 x J", "=12*{c}{J}", _FMT_DOLLAR),
        ("Annual Ad Valorem Taxes", "N = F / 100 x B", "={c}{F}/100*{c}{B}", _FMT_DOLLAR),
        ("Annual PID Installment (first year)", "O",
         f"=IF(SUMIF({rb},{{h}},{rcol('C')})=0,0,SUMIF({rb},{{h}},{rcol('J')})/SUMIF({rb},{{h}},{rcol('C')}))",
         _FMT_DOLLAR),
        ("Total Annual Homeowner Payment with PID", "P = M + N + O", "={c}{M}+{c}{N}+{c}{O}", _FMT_DOLLAR),
        ("No PID Scenario", None, None, None),
        ("Lot Price assuming No PID", "Q = G + D", "={c}{G}+{c}{D}", _FMT_DOLLAR),
        ("Incremental Home Price", "R = D x E", "={c}{D}*{c}{E}", _FMT_DOLLAR),
        ("Home Price without PID", "S = F + R", "={c}{F}+{c}{R}", _FMT_DOLLAR),
        ("Mortgage without PID", "T = C x S", "={c}{C}*{c}{S}", _FMT_DOLLAR),
        ("Down Payment without PID", "U = S - T", "={c}{S}-{c}{T}", _FMT_DOLLAR),
        ("Monthly Mortgage Payment without PID", "V", "=IF({c}{T}=0,0,PMT({c}{Rate}/12,{c}{Term}*12,-{c}{T}))",
         _FMT_DOLLAR),
        ("Annual Mortgage Payment", "W = 12 x V", "=12*{c}{V}", _FMT_DOLLAR),
        ("Annual Ad Valorem Taxes", "X = S / 100 x B", "={c}{S}/100*{c}{B}", _FMT_DOLLAR),
        ("Total Annual Homeowner Payment without PID", "Z = W + X", "={c}{W}+{c}{X}", _FMT_DOLLAR),
        ("Summary", None, None, None),
        ("Annual Savings to Homeowner with PID", "AA = Z - P", "={c}{Z}-{c}{P}", _FMT_DOLLAR),
        ("Down Payment Savings with PID", "BB = U - I", "={c}{U}-{c}{I}", _FMT_DOLLAR),
        ("Lower Home Price with PID", "DD = R", "={c}{R}", _FMT_DOLLAR),
        ("Total Savings over PID Term", "CC = Years x AA + BB", "={c}{Years}*{c}{AA}+{c}{BB}", _FMT_DOLLAR),
    ]
    rows: dict[str, int] = {}
    r, i = 5, 0
    for label, ref, tmpl, fmt in spec:
        r += 1
        if ref is None:
            _section(ws, r, label, 1, LAST)
            continue
        key = ref.split(" =")[0].split(" ")[0]
        rows[key] = r
        summ = key in ("P", "Z", "AA", "BB", "CC", "DD")
        fill, font = (_TOTAL_FILL, _TOTAL_FONT) if summ else (_band(i), _BODY_FONT)
        i += 1
        _cell(ws, r, 1, label, fill=fill, font=font, align=_LEFT)
        _cell(ws, r, 2, ref, fill=fill, font=_NOTE_FONT, align=_CENTER)
        for p in range(1, NP + 1):
            col = CL(2 + p)
            v = tmpl.replace("{h}", f"{col}$5").replace("{pc}", pcol(p))
            for kk in sorted(rows, key=len, reverse=True):
                v = v.replace("{c}{" + kk + "}", f"{col}{rows[kk]}")
            v = f'=IF({col}$5="","",{v[1:]})'
            _cell(ws, r, 2 + p, v, fmt, fill, font)
    ws.cell(row=r + 2, column=1, value="No PID scenario: the developer passes the infrastructure the PID would "
            "have financed into the lot price, and the home price rises by the home-price-to-lot-cost "
            "multiple. Does not capture marketability effects of fewer amenities.").font = _NOTE_FONT
    ws.freeze_panes = "C6"


def _summary(ws):
    LAST = 8
    _title(ws, '=TITLE2&" — "&TEXT(DATE_OF_ANALYSIS,"mmmm d, yyyy")', LAST)
    for col, w in {1: 36, 2: 16, 3: 16, 4: 16, 5: 18, 6: 16, 7: 16, 8: 16}.items():
        ws.column_dimensions[CL(col)].width = w
    SU = "'Sources and Uses'"
    TOTC = CL(2 + NS)
    ROLL = "'Assessment Roll'"
    r = 5
    _section(ws, r, "Estimated Reimbursement Capacity", 1, LAST)
    r += 1
    ws.row_dimensions[r].height = 40
    heads = ["Bond Series", "Effective Tax Rate", "Avg. Gross Annual Assessment",
             "Avg. Est. Total Assessment Per Lot", "Total Par Amount", "Est. Net Proceeds to Developer",
             "Total Improvement Costs", "Net Developer Investment"]
    for i, h in enumerate(heads):
        _hdr(ws, r, 1 + i, h)
    lots = f"{ROLL}!$C${ROLL1 + 1}"
    vals = [
        ('=_xlfn.TEXTJOIN(", ",TRUE,' + ",".join(f'IF({S(k)}!$C${R_ACTIVE}=1,{S(k)}!$C${R_LABEL},"")'
                                                 for k in range(1, NS + 1)) + ")", None),
        ("=PID_RATE", _FMT_RATE),
        (f"=IF({lots}=0,0,SUMPRODUCT({ROLL}!$C${ROLL0}:$C${ROLL1},{ROLL}!$D${ROLL0}:$D${ROLL1})/{lots}"
         f"/100*PID_RATE)", _FMT_DOLLAR),
        (f"=IF({lots}=0,0,{ROLL}!$F${ROLL1 + 1}/{lots})", _FMT_DOLLAR),
        (f"={SU}!${TOTC}$6", _FMT_DOLLAR),
        (f"={SU}!${TOTC}$7", _FMT_DOLLAR),
        (f"=SUM({DI}!$E${PH0}:$E${PH1})", _FMT_DOLLAR),
        (f"=G{r + 1}-F{r + 1}", _FMT_DOLLAR),
    ]
    ws.row_dimensions[r + 1].height = 28
    for i, (v, fmt) in enumerate(vals):
        _cell(ws, r + 1, 1 + i, v, fmt, font=_BOLD_FONT, align=_CENTER)
    r += 4
    _section(ws, r, "Results by Series", 1, LAST)
    r += 1
    ws.row_dimensions[r].height = 30
    for i, h in enumerate(["Series / Improvement Area", "Delivery", "Lots", "Par Amount", "Binding Constraint",
                           "Project Fund", "Improvement Costs", "Value-to-Lien"]):
        _hdr(ws, r, 1 + i, h)
    first = r + 1
    for k in range(1, NS + 1):
        r += 1
        f = _band(k)
        sh, on = S(k), f"{S(k)}!$C${R_ACTIVE}=0"
        cells = [
            (f'=IF({on},"",{sh}!$C${R_LABEL}&" — "&{sh}!$C${R_AREA})', None, _LEFT),
            (f'=IF({on},"",{sh}!$C${R_BONDDEL})', _FMT_DATE, _RIGHT),
            (f'=IF({on},"",{sh}!$K${P_LOTS})', _FMT_NUM, _RIGHT),
            (f'=IF({on},"",{sh}!$C${Z_PAR})', _FMT_DOLLAR, _RIGHT),
            (f'=IF({on},"",{sh}!$C${Z_BIND})', None, _CENTER),
            (f'=IF({on},"",{sh}!$C${U_PF})', _FMT_DOLLAR, _RIGHT),
            (f'=IF({on},"",{sh}!$C${R_COSTS})', _FMT_DOLLAR, _RIGHT),
            (f'=IF({on},"",{sh}!$C${Z_VTL})', _FMT_X, _RIGHT),
        ]
        for j, (v, fmt, al) in enumerate(cells):
            _cell(ws, r, 1 + j, v, fmt, f, align=al)
    r += 1
    _cell(ws, r, 1, "Total", fill=_TOTAL_FILL, font=_TOTAL_FONT, align=_LEFT)
    for c in range(2, LAST + 1):
        s = c in (3, 4, 6, 7)
        _cell(ws, r, c, f"=SUM({CL(c)}{first}:{CL(c)}{r - 1})" if s else None,
              (_FMT_NUM if c == 3 else _FMT_DOLLAR) if s else None, _TOTAL_FILL, _TOTAL_FONT)
    r += 3
    _section(ws, r, "Model Checks", 1, LAST)
    checks = [
        ("Sources equal uses (all series)", f"=IF(ABS({SU}!${TOTC}${6 + 10})<1,\"OK\",\"CHECK\")"),
        ("Total tax rate within policy cap", "=IF(LEFT('Overlapping Taxes'!$D$22,3)=\"Yes\",\"OK\",\"CHECK\")"),
        ("Value-to-lien meets minimum (all series)",
         "=IF(" + "+".join(f"AND({S(k)}!$C${R_ACTIVE}=1,{S(k)}!$C${Z_VTL}<MIN_VTL)" for k in range(1, NS + 1))
         + "=0,\"OK\",\"CHECK\")"),
        ("Improvement costs fully funded (all series)",
         "=IF(" + "+".join(f"AND({S(k)}!$C${R_ACTIVE}=1,{S(k)}!$C${U_NET}<-1)" for k in range(1, NS + 1))
         + "=0,\"OK\",\"SHORTFALL\")"),
        ("Assessment roll ties to par", f"=IF(ABS({ROLL}!$F${ROLL1 + 1}-{SU}!${TOTC}$6)<100,\"OK\",\"CHECK\")"),
    ]
    for i, (lab, f) in enumerate(checks):
        r += 1
        _cell(ws, r, 1, lab, fill=_band(i), align=_LEFT)
        _cell(ws, r, 2, f, fill=_band(i), font=_BOLD_FONT, align=_CENTER)
    from openpyxl.formatting.rule import CellIsRule
    rng = f"B{r - len(checks) + 1}:B{r}"
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"OK"'], fill=_GREEN_FILL))
    ws.conditional_formatting.add(rng, CellIsRule(operator="notEqual", formula=['"OK"'], fill=_RED_FILL))
    vtlr = f"H{first}:H{first + NS - 1}"
    ws.conditional_formatting.add(vtlr, CellIsRule(operator="lessThan", formula=["MIN_VTL"], fill=_RED_FILL))
    r += 3
    _section(ws, r, "Key Assumptions", 1, LAST)
    items = [
        ("Lots / phasing", f'=SUM({DI}!$O${PH0}:$O${PH1})&" lots in "&SUMPRODUCT(--({DI}!$B${PH0}:$B${PH1}<>""))'
                           f'&" improvement area(s); bonds sized and issued by series as each area is delivered"'),
        ("Equivalent tax rate", '="$"&TEXT(PID_RATE,"0.0000")&" per $100 of value ("&PID_RATE_METHOD&")"'),
        ("Bond interest rate", '=TEXT(BONDRATE,"0.00%")&" coupon + "&TEXT(ADDITIONAL_INTEREST,"0.00%")'
                               '&" additional interest = "&TEXT(BONDRATE+ADDITIONAL_INTEREST,"0.00%")&" all-in"'
                               '&IF(RATE_ADD>0,"; sized with "&TEXT(RATE_ADD,"0.00%")&" haircut","")'),
        ("Final maturity", '=FINAL_MAT_YRS&" years; principal "&TEXT(DATE(2000,PRIN_MATURITY,1),"mmmm")&" 1"'),
        ("Capitalized interest", '=CAPI_PERIOD&" months from each delivery (unless overridden by series)"'),
        ("Administrative expenses", '="$"&TEXT(ADMIN_FEE_FIXED_LOT,"#,##0.00")&" per lot + $"&'
                                    'TEXT(ADMIN_FEE_FIXED_SERIES,"#,##0")&" per series, inflating "&'
                                    'TEXT(ADMIN_ESCALATOR,"0%")&" per year"'),
        ("Fees", '="UWD "&TEXT(UWD_FEE,"0.00%")&"; COI "&TEXT(COI_FEE,"0.00%")&"; Admin Fund "&'
                 'TEXT(ADMIN_FEE,"0.00%")&"; D&P Reserve "&TEXT(DP_RESERVE_DEPOSIT,"0.00%")&"; Reserve: "&RESERVE_METHOD'),
        ("Inflation", '=IF(INFLATION_ASP=0,"None assumed",TEXT(INFLATION_ASP,"0.00%")&" per year from "&INFLATION_START)'),
    ]
    for i, (k_, v) in enumerate(items):
        r += 1
        _cell(ws, r, 1, k_, fill=_band(i), font=_BOLD_FONT, align=_LEFT)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=LAST)
        _cell(ws, r, 2, v, fill=_band(i), align=_LEFT)
    r += 2
    ws.cell(row=r, column=1, value="Recalculating model — edit the yellow cells on Inputs, Developer Information "
            "and Overlapping Taxes; every other tab updates automatically.").font = _NOTE_FONT


# ── Public entry point ───────────────────────────────────────────────────────
def build_excel_model(path: str, cfg: Optional[ModelConfig] = None,
                      comparables: Optional[list[dict]] = None) -> str:
    cfg = cfg or sample_config()
    wb = build_inputs_template(None, cfg, comparables, return_workbook=True)
    ws_in = wb["Inputs"]
    ws_in.cell(row=2, column=2, value="Recalculating Excel model — edit the yellow Value cells only; every other "
               "tab updates automatically. Range names in column D are also read by the Python model.")
    wb.defined_names["PID_RATE"] = DefinedName("PID_RATE", attr_text=f"'Overlapping Taxes'!$D${5 + MAX_TAXES + 3}")

    comps = wb["Deal Comparables"]
    order = ["Summary", "Sources and Uses", "Bond Sizing Analysis", "Assessment Roll", "Annual Installments",
             "Project and PID Rev Analysis", "Tax Rate and VTL Tests", "Homeowner Cost Comparison"]
    sheets = {name: wb.create_sheet(name) for name in order}
    series = [wb.create_sheet(f"Series {k}") for k in range(1, NS + 1)]
    ab = wb.create_sheet(AB)
    wb.move_sheet(comps, offset=len(wb.sheetnames) - wb.sheetnames.index("Deal Comparables") - 1)

    _absorption(ab)
    for k, ws in enumerate(series, 1):
        _series(ws, k)
        ws.sheet_properties.tabColor = "2E6FA3"
    _sources_uses(sheets["Sources and Uses"])
    _bond_sizing(sheets["Bond Sizing Analysis"])
    _roll(sheets["Assessment Roll"])
    _installments(sheets["Annual Installments"])
    _project(sheets["Project and PID Rev Analysis"])
    _tests(sheets["Tax Rate and VTL Tests"])
    _homeowner(sheets["Homeowner Cost Comparison"])
    _summary(sheets["Summary"])
    for name in ("Inputs", "Developer Information", "Overlapping Taxes"):
        wb[name].sheet_properties.tabColor = "FFC000"
    ab.sheet_properties.tabColor = "808080"
    for ws in wb.worksheets:
        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToWidth = 1
        ws.sheet_properties.pageSetUpPr.fitToPage = True
    wb.save(path)
    return path
