"""
report.py — Multi-sheet Excel output for the Texas PID analysis.

Sheets
  1. Summary                      memo table, assumptions, series results, credit tests
  2. Sources and Uses             one column per series + total
  3. Bond Sizing Analysis         bond-year installments vs. debt service and coverage
  4. Debt Service                 semiannual schedule by series
  5. Assessment Roll              lots, assessment and installment by series and lot type
  6. Annual Installments          per-lot installment schedule by series / lot type
  7. Project and PID Rev Analysis absorption and taxable value by year
  8. Tax Rate and VTL Tests       overlapping taxes and value-to-lien by series

Line items are written as values from the Python engine; totals, percentages
and checks are live Excel formulas.
"""

from __future__ import annotations

from datetime import date
from typing import Optional

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as CL

from .bonds import BondModel
from .config import ModelConfig
from .development import DevelopmentProjection

# ── Shared style constants (same palette as the reimbursement model report) ──
_BLUE_FILL = PatternFill("solid", fgColor="1F4E79")
_SUBHDR_FILL = PatternFill("solid", fgColor="2E6FA3")
_LIGHT_FILL = PatternFill("solid", fgColor="D6E4F0")
_GRAY_FILL = PatternFill("solid", fgColor="F2F2F2")
_TOTAL_FILL = PatternFill("solid", fgColor="BDD7EE")
_WHITE_FILL = PatternFill("solid", fgColor="FFFFFF")
_RED_FILL = PatternFill("solid", fgColor="FFD0D0")
_GREEN_FILL = PatternFill("solid", fgColor="E2EFDA")

_THIN = Side(style="thin", color="BFBFBF")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)

_HEADER_FONT = Font(name="Calibri", bold=True, color="1F4E79", size=9)
_SUBHDR_FONT = Font(name="Calibri", bold=True, color="FFFFFF", size=10)
_BODY_FONT = Font(name="Calibri", size=9)
_BOLD_FONT = Font(name="Calibri", bold=True, size=9)
_TOTAL_FONT = Font(name="Calibri", bold=True, color="1F4E79", size=9)
_NOTE_FONT = Font(name="Calibri", italic=True, size=9, color="595959")

_FMT_DATE = "MM/DD/YYYY"
_FMT_DOLLAR = '"$"#,##0;("$"#,##0);"-"'
_FMT_DOLLAR2 = '"$"#,##0.00;("$"#,##0.00);"-"'
_FMT_NUM = '#,##0;(#,##0);"-"'
_FMT_PCT = '0.00%;(0.00%);"-"'
_FMT_RATE = '"$"0.0000'
_FMT_X = '0.00"x"'

_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
_RIGHT = Alignment(horizontal="right", vertical="center")
_LEFT = Alignment(horizontal="left", vertical="center")


def _title_block(ws, line1: str, line2: str, line3: str, last_col: int):
    for row, (txt, size, h) in enumerate([(line1, 14, 22), (line2, 11, 16), (line3, 11, 16)], 1):
        ws.row_dimensions[row].height = h
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=last_col)
        c = ws.cell(row=row, column=1, value=txt)
        c.fill, c.alignment = _BLUE_FILL, _CENTER
        c.font = Font(name="Calibri", bold=True, color="FFFFFF", size=size)
    ws.row_dimensions[4].height = 6
    ws.sheet_view.showGridLines = False


def _hdr(ws, row, col, label, width=None):
    c = ws.cell(row=row, column=col, value=label)
    c.font, c.fill, c.border, c.alignment = _HEADER_FONT, _LIGHT_FILL, _BORDER, _CENTER
    if width:
        ws.column_dimensions[CL(col)].width = width
    return c


def _section(ws, row, label, first_col, last_col):
    for col in range(first_col, last_col + 1):
        ws.cell(row=row, column=col).fill = _SUBHDR_FILL
    c = ws.cell(row=row, column=first_col, value=label)
    c.font = _SUBHDR_FONT
    ws.row_dimensions[row].height = 16


def _cell(ws, row, col, value, fmt=None, fill=_WHITE_FILL, font=_BODY_FONT, align=_RIGHT):
    c = ws.cell(row=row, column=col, value=value)
    c.fill, c.font, c.border, c.alignment = fill, font, _BORDER, align
    if fmt:
        c.number_format = fmt
    return c


def _band(i: int):
    return _GRAY_FILL if i % 2 else _WHITE_FILL


def _titles(cfg: ModelConfig):
    return cfg.pid_name, f"{cfg.city}, {cfg.county} County, Texas"


# ── 1. Summary ───────────────────────────────────────────────────────────────
def _summary(ws, cfg: ModelConfig, bm: BondModel, dev: DevelopmentProjection):
    LAST = 8
    t1, t2 = _titles(cfg)
    _title_block(ws, t1, t2, f"{cfg.analysis_type} — {cfg.analysis_date:%B %d, %Y}", LAST)
    for col, w in {1: 34, 2: 16, 3: 16, 4: 16, 5: 16, 6: 16, 7: 16, 8: 16}.items():
        ws.column_dimensions[CL(col)].width = w

    # Memo table (matches the Tierra PID memo)
    all_las = [(la, s.phase.lots[p]) for s in bm.series for p, la in s.lot_assessments.items()]
    lots = sum(n for _, n in all_las)
    avg_gross = sum(la.gross_installment * n for la, n in all_las) / lots if lots else 0
    avg_assess = sum(la.assessment * n for la, n in all_las) / lots if lots else 0
    r = 5
    _section(ws, r, "Estimated Reimbursement Capacity", 1, LAST)
    r += 1
    hdrs = ["Effective Tax Rate", "Avg. Gross Annual Assessment", "Avg. Est. Total Assessment Per Lot",
            "Total Par Amount", "Est. Net Proceeds to Developer", "Total Improvement Costs",
            "Shortfall / (Excess)"]
    ws.row_dimensions[r].height = 40
    for i, h in enumerate(hdrs):
        _hdr(ws, r, 2 + i, h)
    _hdr(ws, r, 1, "Series")
    vals = [
        (", ".join(s.series_label for s in bm.series), None),
        (cfg.effective_pid_rate, _FMT_RATE),
        (avg_gross, _FMT_DOLLAR),
        (avg_assess, _FMT_DOLLAR),
        ("='Sources and Uses'!" + _SU_TOTAL_REF["par"], _FMT_DOLLAR),
        ("='Sources and Uses'!" + _SU_TOTAL_REF["project_fund"], _FMT_DOLLAR),
        (sum(ph.improvement_costs for ph in cfg.phases), _FMT_DOLLAR),
        (f"=G{r + 1}-F{r + 1}", _FMT_DOLLAR),
    ]
    for i, (v, fmt) in enumerate(vals):
        _cell(ws, r + 1, 1 + i, v, fmt, font=_BOLD_FONT, align=_CENTER)
    ws.row_dimensions[r + 1].height = 28
    r += 3

    # Series results
    _section(ws, r, "Results by Series", 1, LAST)
    r += 1
    ws.row_dimensions[r].height = 30
    cols = ["Series / Improvement Area", "Delivery", "Lots", "Par Amount", "Binding Constraint",
            "Project Fund", "Improvement Costs", "Value-to-Lien"]
    for i, h in enumerate(cols):
        _hdr(ws, r, 1 + i, h)
    first = r + 1
    for i, s in enumerate(bm.series):
        r += 1
        f = _band(i)
        _cell(ws, r, 1, f"{s.series_label} — {s.phase.name}", fill=f, align=_LEFT)
        _cell(ws, r, 2, s.delivery, _FMT_DATE, f)
        _cell(ws, r, 3, s.lots, _FMT_NUM, f)
        _cell(ws, r, 4, s.par, _FMT_DOLLAR, f)
        _cell(ws, r, 5, s.binding, fill=f, align=_CENTER)
        _cell(ws, r, 6, s.su.project_fund, _FMT_DOLLAR, f)
        _cell(ws, r, 7, s.phase.improvement_costs, _FMT_DOLLAR, f)
        vf = _GREEN_FILL if s.vtl >= cfg.min_vtl else _RED_FILL
        _cell(ws, r, 8, s.vtl, _FMT_X, vf)
    r += 1
    _cell(ws, r, 1, "Total", fill=_TOTAL_FILL, font=_TOTAL_FONT, align=_LEFT)
    for col in (3, 4, 6, 7):
        _cell(ws, r, col, f"=SUM({CL(col)}{first}:{CL(col)}{r - 1})",
              _FMT_NUM if col == 3 else _FMT_DOLLAR, _TOTAL_FILL, _TOTAL_FONT)
    for col in (2, 5, 8):
        _cell(ws, r, col, None, fill=_TOTAL_FILL)
    r += 3

    # Assumptions (memo bullet points)
    _section(ws, r, "Key Assumptions", 1, LAST)
    first_close = min(ph.lot_delivery for ph in cfg.phases)
    pace = sum(p.absorption_per_month for p in cfg.products)
    asps = ", ".join(f"{p.name} ${p.avg_selling_price:,.0f}" for p in cfg.products)
    phases = "; ".join(f"{sum(ph.lots.values())} lots — bonds {ph.bond_delivery:%b %Y}" for ph in cfg.phases)
    items = [
        ("Lots / phasing", f"{cfg.total_lots} lots in {len(cfg.phases)} improvement area(s): {phases}"),
        ("Average selling prices", asps),
        ("Absorption", f"{pace:g} homes per month in total; first lots delivered {first_close:%B %Y}"),
        ("Equivalent tax rate", f"${cfg.effective_pid_rate:.4f} per $100 of value ({cfg.pid_rate_method})"),
        ("Bond interest rate", f"{cfg.bond_rate:.2%} coupon + {cfg.additional_interest:.2%} additional interest"
                               f" = {cfg.bond_rate + cfg.additional_interest:.2%} all-in"
                               + (f"; sized with {cfg.rate_add:.2%} haircut" if cfg.rate_add else "")),
        ("Final maturity", f"{cfg.final_mat_yrs} years; principal {date(2000, cfg.prin_maturity, 1):%B} 1"),
        ("Capitalized interest", f"{cfg.capi_period} months from each delivery"),
        ("Administrative expenses", f"${cfg.admin_fee_fixed_lot:,.2f} per lot, inflating "
                                    f"{cfg.admin_escalator:.0%} per year"),
        ("Fees", f"UWD {cfg.uwd_fee:.2%}; COI {cfg.coi_fee:.2%}; Admin Fund {cfg.admin_fee:.2%}; "
                 f"D&P Reserve {cfg.dp_reserve_deposit:.2%}; Reserve: {cfg.reserve_method}"),
        ("Inflation", "None assumed" if not cfg.inflation_asp else f"{cfg.inflation_asp:.2%} per year from "
                                                                   f"{cfg.inflation_start}"),
    ]
    for i, (k, v) in enumerate(items):
        r += 1
        _cell(ws, r, 1, k, fill=_band(i), font=_BOLD_FONT, align=_LEFT)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=LAST)
        _cell(ws, r, 2, v, fill=_band(i), align=_LEFT)
    r += 2
    ws.cell(row=r, column=1, value="Many factors could influence the net proceeds to the developer, "
            "including city development fees, final pricing, appraised values and the city's PID policy."
            ).font = _NOTE_FONT


# ── 2. Sources & Uses ────────────────────────────────────────────────────────
_SU_ROWS = [
    ("SOURCES OF FUNDS", None),
    ("Par Amount", "par"),
    ("Total Sources", "=sources"),
    (None, None),
    ("USES OF FUNDS", None),
    ("Deposit to Project Fund (Developer Reimbursement)", "project_fund"),
    ("Deposit to Reserve Fund", "reserve_fund"),
    ("Deposit to Delinquency & Prepayment Reserve", "dp_reserve"),
    ("Deposit to Capitalized Interest Account", "capi"),
    ("Deposit to Administrative Fund", "admin_fund"),
    ("City Development / PID Fee", "city_fee"),
    ("Underwriter's Discount", "uwd"),
    ("Costs of Issuance", "coi"),
    ("Total Uses", "=uses"),
    (None, None),
    ("Sources less Uses", "=check"),
    ("Improvement Costs", "costs"),
    ("Project Fund less Costs", "=net"),
]
_SU_TOTAL_REF: dict[str, str] = {}


def _sources_uses(ws, cfg: ModelConfig, bm: BondModel):
    n = len(bm.series)
    LABEL, FIRST = 1, 2
    TOT = FIRST + n
    PCT = TOT + 1
    LAST = PCT
    t1, t2 = _titles(cfg)
    _title_block(ws, t1, t2, f"Sources and Uses of Funds — {cfg.bond_title}", LAST)
    ws.column_dimensions["A"].width = 46
    for c in range(FIRST, PCT + 1):
        ws.column_dimensions[CL(c)].width = 17
    ws.row_dimensions[5].height = 46
    _hdr(ws, 5, LABEL, "Item")
    for i, s in enumerate(bm.series):
        c = ws.cell(row=5, column=FIRST + i,
                    value=f"{s.series_label}\n{s.phase.name}\n{s.delivery:%m/%d/%Y}")
        c.fill, c.font, c.alignment, c.border = _BLUE_FILL, Font(name="Calibri", bold=True, color="FFFFFF",
                                                                 size=9), _CENTER, _BORDER
    for col, txt in ((TOT, "Total"), (PCT, "% of Par")):
        c = ws.cell(row=5, column=col, value=txt)
        c.fill, c.font, c.alignment, c.border = _BLUE_FILL, Font(name="Calibri", bold=True, color="FFFFFF",
                                                                 size=9), _CENTER, _BORDER
    row_of: dict[str, int] = {}
    r = 6
    for i, (label, key) in enumerate(_SU_ROWS):
        if label is None:
            ws.row_dimensions[r].height = 6
            r += 1
            continue
        if key is None:
            _section(ws, r, label, 1, LAST)
            r += 1
            continue
        is_total = key.startswith("=")
        fill = _TOTAL_FILL if is_total else _band(i)
        font = _TOTAL_FONT if is_total else _BODY_FONT
        _cell(ws, r, LABEL, label, fill=fill, font=font, align=_LEFT)
        for j, s in enumerate(bm.series):
            col = FIRST + j
            L = CL(col)
            if key == "=sources":
                v = f"={L}{row_of['par']}"
            elif key == "=uses":
                v = f"=SUM({L}{row_of['project_fund']}:{L}{row_of['coi']})"
            elif key == "=check":
                v = f"={L}{row_of['=sources']}-{L}{row_of['=uses']}"
            elif key == "=net":
                v = f"={L}{row_of['project_fund']}-{L}{row_of['costs']}"
            elif key == "costs":
                v = s.phase.improvement_costs
            else:
                v = getattr(s.su, key)
            _cell(ws, r, col, v, _FMT_DOLLAR, fill, font)
        _cell(ws, r, TOT, f"=SUM({CL(FIRST)}{r}:{CL(TOT - 1)}{r})", _FMT_DOLLAR, fill, _TOTAL_FONT)
        row_of[key] = r
        if key not in ("costs", "=net", "=check"):
            _cell(ws, r, PCT, f"=IF(${CL(TOT)}${row_of.get('par', r)}=0,0,{CL(TOT)}{r}/${CL(TOT)}"
                              f"${row_of.get('par', r)})", _FMT_PCT, fill, font)
        else:
            _cell(ws, r, PCT, None, fill=fill)
        r += 1
    _SU_TOTAL_REF["par"] = f"${CL(TOT)}${row_of['par']}"
    _SU_TOTAL_REF["project_fund"] = f"${CL(TOT)}${row_of['project_fund']}"
    r += 1
    notes = [
        "Project Fund is the residual after all other uses — the estimated reimbursement to the developer.",
        f"Reserve Fund: {cfg.reserve_method}. D&P Reserve is grown thereafter from the "
        f"{cfg.additional_interest:.2%} additional interest.",
        "Bonds assumed priced at par (no original issue premium/discount).",
    ]
    for t in notes:
        ws.cell(row=r, column=1, value=t).font = _NOTE_FONT
        r += 1
    ws.freeze_panes = "B6"


# ── 3. Bond Sizing Analysis ─────────────────────────────────────────────────
def _bond_sizing(ws, cfg: ModelConfig, bm: BondModel):
    df = bm.annual_table()
    df = df[(df["Total Net Debt Service"] != 0) | (df["Total Annual Installments"] != 0)]
    series_cols = [c for c in df.columns if c.endswith("Net DS")]
    cols = (["Bond Year (9/1)", "Assessed Lots", "Installment - Principal", "Installment - Interest",
             "Additional Interest", "Administrative Expenses", "Total Annual Installments"]
            + series_cols + ["Total Net Debt Service", "P&I Coverage"])
    LAST = len(cols)
    t1, t2 = _titles(cfg)
    _title_block(ws, t1, t2, f"Bond Sizing Analysis — Annual Installments vs. Debt Service "
                             f"(ETR ${cfg.effective_pid_rate:.4f})", LAST)
    ws.row_dimensions[5].height = 40
    for i, c in enumerate(cols):
        _hdr(ws, 5, 1 + i, c.replace(" - ", "\n").replace(" Net DS", "\nNet Debt Service"),
             12 if i < 2 else 15)
    first = 6
    for i, (_, row) in enumerate(df.iterrows()):
        r = first + i
        for j, c in enumerate(cols):
            v = row[c]
            if c == "Bond Year (9/1)":
                _cell(ws, r, 1 + j, str(int(v)), fill=_band(i), font=_BOLD_FONT, align=_CENTER)
            elif c == "P&I Coverage":
                L = CL(1 + j)
                pi = f"({CL(3)}{r}+{CL(4)}{r})*{cfg.tax_collection}"
                ds = f"{CL(j)}{r}"
                _cell(ws, r, 1 + j, f"=IF({ds}=0,\"\",{pi}/{ds})", _FMT_X, _band(i))
            elif c == "Total Net Debt Service":
                a, b = CL(1 + cols.index(series_cols[0])), CL(1 + cols.index(series_cols[-1]))
                _cell(ws, r, 1 + j, f"=SUM({a}{r}:{b}{r})", _FMT_DOLLAR, _band(i))
            elif c == "Total Annual Installments":
                _cell(ws, r, 1 + j, f"=SUM(C{r}:F{r})", _FMT_DOLLAR, _band(i))
            else:
                _cell(ws, r, 1 + j, float(v), _FMT_NUM if c == "Assessed Lots" else _FMT_DOLLAR, _band(i))
    tr = first + len(df)
    _cell(ws, tr, 1, "Total", fill=_TOTAL_FILL, font=_TOTAL_FONT, align=_CENTER)
    _cell(ws, tr, 2, None, fill=_TOTAL_FILL)
    for j in range(2, LAST - 1):
        L = CL(1 + j)
        _cell(ws, tr, 1 + j, f"=SUM({L}{first}:{L}{tr - 1})", _FMT_DOLLAR, _TOTAL_FILL, _TOTAL_FONT)
    _cell(ws, tr, LAST, None, fill=_TOTAL_FILL)
    ws.cell(row=tr + 2, column=1,
            value="Installments are billed each October and due by January 31, then pay the March 1 and "
                  "September 1 debt service of that bond year. Debt service shown net of capitalized "
                  "interest. Coverage = principal and interest installments x collection rate / net debt "
                  "service (target 1.00x when sized to the installment).").font = _NOTE_FONT
    ws.freeze_panes = "B6"


# ── 4. Debt Service ─────────────────────────────────────────────────────────
def _debt_service(ws, cfg: ModelConfig, bm: BondModel):
    LAST = 8
    t1, t2 = _titles(cfg)
    _title_block(ws, t1, t2, "Debt Service Schedules by Series", LAST)
    for col, w in {1: 13, 2: 15, 3: 9, 4: 15, 5: 15, 6: 15, 7: 15, 8: 16}.items():
        ws.column_dimensions[CL(col)].width = w
    r = 5
    for s in bm.series:
        _section(ws, r, f"{cfg.bond_title}, {s.series_label} — {s.phase.name}  |  "
                        f"Par ${s.par:,.0f}  |  Dated/Delivered {s.delivery:%m/%d/%Y}  |  "
                        f"First par call {date(s.delivery.year + cfg.first_par_call_yrs, cfg.prin_maturity, 1):%m/%d/%Y}",
                 1, LAST)
        r += 1
        ws.row_dimensions[r].height = 30
        for i, h in enumerate(["Payment Date", "Principal", "Coupon", "Interest", "Total Debt Service",
                               "Capitalized Interest", "Net Debt Service", "Outstanding Principal"]):
            _hdr(ws, r, 1 + i, h)
        first = r + 1
        for i, d in enumerate(s.ds):
            r += 1
            f = _band(i)
            _cell(ws, r, 1, d.date, _FMT_DATE, f, align=_CENTER)
            _cell(ws, r, 2, d.principal, _FMT_DOLLAR, f)
            _cell(ws, r, 3, d.coupon if d.principal else None, _FMT_PCT if d.principal else None, f)
            _cell(ws, r, 4, round(d.interest, 2), _FMT_DOLLAR2, f)
            _cell(ws, r, 5, f"=B{r}+D{r}", _FMT_DOLLAR2, f)
            _cell(ws, r, 6, -round(d.capi, 2) if d.capi else None, _FMT_DOLLAR2, f)
            _cell(ws, r, 7, f"=E{r}+F{r}", _FMT_DOLLAR2, f)
            prev = f"H{r - 1}" if i else str(s.par)
            _cell(ws, r, 8, f"={prev}-B{r}", _FMT_DOLLAR, f)
        r += 1
        _cell(ws, r, 1, "Total", fill=_TOTAL_FILL, font=_TOTAL_FONT, align=_CENTER)
        _cell(ws, r, 3, None, fill=_TOTAL_FILL)
        _cell(ws, r, 8, None, fill=_TOTAL_FILL)
        for col in (2, 4, 5, 6, 7):
            _cell(ws, r, col, f"=SUM({CL(col)}{first}:{CL(col)}{r - 1})", _FMT_DOLLAR, _TOTAL_FILL, _TOTAL_FONT)
        r += 3


# ── 5. Assessment Roll ──────────────────────────────────────────────────────
def _assessment_roll(ws, cfg: ModelConfig, bm: BondModel):
    cols = ["Series / Improvement Area", "Lot Type", "Lots", "Average Selling Price",
            "Assessment Per Lot", "Total Assessment", "Assessment / ASP",
            "First Annual Installment Per Lot", "Average Annual Installment Per Lot",
            "Equivalent Tax Rate", "Total First-Year Installments"]
    LAST = len(cols)
    t1, t2 = _titles(cfg)
    _title_block(ws, t1, t2, "Assessment Roll Summary", LAST)
    ws.row_dimensions[5].height = 40
    widths = [36, 16, 8, 14, 14, 16, 12, 14, 14, 12, 16]
    for i, (h, w) in enumerate(zip(cols, widths)):
        _hdr(ws, 5, 1 + i, h, w)
    r, i, first = 5, 0, 6
    for s in bm.series:
        for prod, la in s.lot_assessments.items():
            r += 1
            f = _band(i)
            i += 1
            n = s.phase.lots[prod]
            _cell(ws, r, 1, f"{s.series_label} — {s.phase.name}", fill=f, align=_LEFT)
            _cell(ws, r, 2, prod, fill=f, align=_LEFT)
            _cell(ws, r, 3, n, _FMT_NUM, f)
            _cell(ws, r, 4, la.product.avg_selling_price, _FMT_DOLLAR, f)
            _cell(ws, r, 5, round(la.assessment, 2), _FMT_DOLLAR, f)
            _cell(ws, r, 6, f"=C{r}*E{r}", _FMT_DOLLAR, f)
            _cell(ws, r, 7, f"=IF(D{r}=0,0,E{r}/D{r})", _FMT_PCT, f)
            _cell(ws, r, 8, round(la.schedule[0].installment, 2), _FMT_DOLLAR, f)
            _cell(ws, r, 9, round(la.avg_installment, 2), _FMT_DOLLAR, f)
            _cell(ws, r, 10, f"=IF(D{r}=0,0,H{r}/D{r}*100)", _FMT_RATE, f)
            _cell(ws, r, 11, f"=C{r}*H{r}", _FMT_DOLLAR, f)
    r += 1
    _cell(ws, r, 1, "Total", fill=_TOTAL_FILL, font=_TOTAL_FONT, align=_LEFT)
    for col in range(2, LAST + 1):
        if col in (3, 6, 11):
            _cell(ws, r, col, f"=SUM({CL(col)}{first}:{CL(col)}{r - 1})",
                  _FMT_NUM if col == 3 else _FMT_DOLLAR, _TOTAL_FILL, _TOTAL_FONT)
        else:
            _cell(ws, r, col, None, fill=_TOTAL_FILL)
    ws.cell(row=r + 2, column=1, value="Total Assessment ties to par by series (rounding aside). "
            "Assessments are allocated so every lot type carries the same equivalent tax rate.").font = _NOTE_FONT
    ws.freeze_panes = "C6"


# ── 6. Annual Installments per lot ──────────────────────────────────────────
def _installments(ws, cfg: ModelConfig, bm: BondModel):
    LAST = 8
    t1, t2 = _titles(cfg)
    _title_block(ws, t1, t2, "Estimated Annual Installment Schedule (Per Lot)", LAST)
    for col, w in {1: 8, 2: 13, 3: 14, 4: 14, 5: 14, 6: 15, 7: 15, 8: 16}.items():
        ws.column_dimensions[CL(col)].width = w
    r = 5
    for s in bm.series:
        for prod, la in s.lot_assessments.items():
            _section(ws, r, f"{s.series_label} — {s.phase.name}  |  Lot Type: {prod}  |  "
                            f"Assessment ${la.assessment:,.0f}  |  ASP ${la.product.avg_selling_price:,.0f}",
                     1, LAST)
            r += 1
            ws.row_dimensions[r].height = 30
            for i, h in enumerate(["Year", "Payment Date", "Principal", "Interest", "Additional Interest",
                                   "Administrative Expenses", "Annual Installment", "Principal Balance"]):
                _hdr(ws, r, 1 + i, h)
            first = r + 1
            for i, row in enumerate(la.schedule):
                r += 1
                f = _band(i)
                _cell(ws, r, 1, row.year_no, fill=f, align=_CENTER)
                _cell(ws, r, 2, date(row.bond_year, cfg.prin_maturity, 1), _FMT_DATE, f, align=_CENTER)
                _cell(ws, r, 3, round(row.principal, 2), _FMT_DOLLAR2, f)
                _cell(ws, r, 4, round(row.interest, 2), _FMT_DOLLAR2, f)
                _cell(ws, r, 5, round(row.additional_interest, 2), _FMT_DOLLAR2, f)
                _cell(ws, r, 6, round(row.admin, 2), _FMT_DOLLAR2, f)
                _cell(ws, r, 7, f"=SUM(C{r}:F{r})", _FMT_DOLLAR2, f)
                prev = f"H{r - 1}" if i else str(round(la.assessment, 2))
                _cell(ws, r, 8, f"={prev}-C{r}", _FMT_DOLLAR2, f)
            r += 1
            _cell(ws, r, 1, "Total", fill=_TOTAL_FILL, font=_TOTAL_FONT, align=_CENTER)
            _cell(ws, r, 2, None, fill=_TOTAL_FILL)
            _cell(ws, r, 8, None, fill=_TOTAL_FILL)
            for col in range(3, 8):
                _cell(ws, r, col, f"=SUM({CL(col)}{first}:{CL(col)}{r - 1})", _FMT_DOLLAR, _TOTAL_FILL,
                      _TOTAL_FONT)
            r += 3


# ── 7. Project and PID Rev Analysis ─────────────────────────────────────────
def _project(ws, cfg: ModelConfig, dev: DevelopmentProjection):
    df = dev.annual_table()
    last_build = df.loc[df["Homes Per Year"] > 0, "Year"].max()
    df = df[df["Year"] <= (last_build + 3 if last_build == last_build else df["Year"].max())]
    cols = list(df.columns)
    LAST = len(cols)
    t1, t2 = _titles(cfg)
    _title_block(ws, t1, t2, "Project and PID Revenue Analysis — Absorption and Taxable Value", LAST)
    ws.row_dimensions[5].height = 40
    for i, c in enumerate(cols):
        _hdr(ws, 5, 1 + i, c, 9 if c in ("Year No.", "Year", "Tax Roll Year") else 14)
    for i, (_, row) in enumerate(df.iterrows()):
        r = 6 + i
        for j, c in enumerate(cols):
            v = row[c]
            money = c in ("Home Value", "Vacant Lot Value", "Total Taxable Basis")
            if c in ("Year", "Tax Roll Year"):
                _cell(ws, r, 1 + j, str(int(v)), fill=_band(i), align=_CENTER)
            else:
                _cell(ws, r, 1 + j, float(v), _FMT_DOLLAR if money else _FMT_NUM, _band(i),
                      align=_CENTER if c == "Year No." else _RIGHT)
    r = 6 + len(df) + 1
    ws.cell(row=r, column=1, value=f"Home closings begin {cfg.months_lot_to_close} months after each lot "
            f"delivery at the product absorption pace; taxable basis = closed homes at ASP plus delivered, "
            f"unbuilt lots at lot value; value is on the roll {cfg.av_lag_years} year(s) after closing."
            ).font = _NOTE_FONT
    ws.freeze_panes = "C6"


# ── 8. Tests ────────────────────────────────────────────────────────────────
def _tests(ws, cfg: ModelConfig, bm: BondModel):
    LAST = 6
    t1, t2 = _titles(cfg)
    _title_block(ws, t1, t2, "Tax Rate and Value-to-Lien Tests", LAST)
    for col, w in {1: 40, 2: 14, 3: 16, 4: 16, 5: 14, 6: 16}.items():
        ws.column_dimensions[CL(col)].width = w
    r = 5
    _section(ws, r, "Overlapping Tax Rates (per $100 of value)", 1, LAST)
    r += 1
    _hdr(ws, r, 1, "Taxing Entity")
    _hdr(ws, r, 2, "Type")
    _hdr(ws, r, 3, "Tax Rate")
    first = r + 1
    for i, t in enumerate(cfg.overlapping_taxes):
        r += 1
        _cell(ws, r, 1, t.entity, fill=_band(i), align=_LEFT)
        _cell(ws, r, 2, t.entity_type, fill=_band(i), align=_CENTER)
        _cell(ws, r, 3, t.rate, _FMT_RATE, _band(i))
    last_tax = r
    lines = [
        ("Total Overlapping Tax Rate", f"=SUM(C{first}:C{last_tax})" if cfg.overlapping_taxes else 0),
        ("PID Equivalent Tax Rate", cfg.effective_pid_rate),
        ("Total Tax Rate incl. PID", None),
        ("Maximum Total Tax Rate (Policy)", cfg.max_total_tax_rate),
        ("Within Cap?", None),
    ]
    rows = {}
    for lab, v in lines:
        r += 1
        rows[lab] = r
        _cell(ws, r, 1, lab, fill=_TOTAL_FILL, font=_TOTAL_FONT, align=_LEFT)
        _cell(ws, r, 2, None, fill=_TOTAL_FILL)
        if lab == "Total Tax Rate incl. PID":
            v = f"=C{rows['Total Overlapping Tax Rate']}+C{rows['PID Equivalent Tax Rate']}"
        if lab == "Within Cap?":
            v = (f'=IF(C{rows["Total Tax Rate incl. PID"]}<=C{rows["Maximum Total Tax Rate (Policy)"]}'
                 f'+0.00001,"Yes","No")')
            _cell(ws, r, 3, v, None, _TOTAL_FILL, _TOTAL_FONT, _CENTER)
        else:
            _cell(ws, r, 3, v, _FMT_RATE, _TOTAL_FILL, _TOTAL_FONT)
    r += 3
    _section(ws, r, f"Value-to-Lien at Bond Delivery (minimum {cfg.min_vtl:.2f}x)", 1, LAST)
    r += 1
    ws.row_dimensions[r].height = 30
    for i, h in enumerate(["Series / Improvement Area", "Lots", "Estimated Value at Delivery",
                           "Par Amount", "Value-to-Lien", "Passes?"]):
        _hdr(ws, r, 1 + i, h)
    for i, s in enumerate(bm.series):
        r += 1
        f = _band(i)
        _cell(ws, r, 1, f"{s.series_label} — {s.phase.name}", fill=f, align=_LEFT)
        _cell(ws, r, 2, s.lots, _FMT_NUM, f)
        _cell(ws, r, 3, s.value_at_issuance, _FMT_DOLLAR, f)
        _cell(ws, r, 4, s.par, _FMT_DOLLAR, f)
        _cell(ws, r, 5, f"=IF(D{r}=0,0,C{r}/D{r})", _FMT_X, f)
        _cell(ws, r, 6, f'=IF(E{r}>={cfg.min_vtl},"Yes","No")', None, f, align=_CENTER)
    r += 2
    ws.cell(row=r, column=1, value="Value at delivery = closed homes at ASP plus remaining lots at the "
            "as-developed lot value from Developer Information. Replace with the appraisal once available."
            ).font = _NOTE_FONT


# ── Public entry point ───────────────────────────────────────────────────────
def build_excel_report(cfg: ModelConfig, dev: DevelopmentProjection, bm: BondModel,
                       output_path: str) -> str:
    wb = openpyxl.Workbook()
    sheets = ["Summary", "Sources and Uses", "Bond Sizing Analysis", "Debt Service",
              "Assessment Roll", "Annual Installments", "Project and PID Rev Analysis",
              "Tax Rate and VTL Tests"]
    ws_map = {}
    for i, name in enumerate(sheets):
        ws_map[name] = wb.active if i == 0 else wb.create_sheet(name)
        ws_map[name].title = name
    # S&U first: Summary references its total cells
    _sources_uses(ws_map["Sources and Uses"], cfg, bm)
    _summary(ws_map["Summary"], cfg, bm, dev)
    _bond_sizing(ws_map["Bond Sizing Analysis"], cfg, bm)
    _debt_service(ws_map["Debt Service"], cfg, bm)
    _assessment_roll(ws_map["Assessment Roll"], cfg, bm)
    _installments(ws_map["Annual Installments"], cfg, bm)
    _project(ws_map["Project and PID Rev Analysis"], cfg, dev)
    _tests(ws_map["Tax Rate and VTL Tests"], cfg, bm)
    for ws in wb.worksheets:
        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToWidth = 1
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.print_title_rows = "1:5"
    wb.save(output_path)
    return output_path
