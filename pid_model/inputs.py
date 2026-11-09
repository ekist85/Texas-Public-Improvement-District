"""
inputs.py — Build and read the Texas PID Inputs workbook.

Sheets (same convention as the UT PID and Texas Reimbursement models):
  Inputs                — col B Title, col C Value (yellow = edit), col D Range Name, col E Notes
  Developer Information — product setup grid + improvement area / series table
  Overlapping Taxes     — taxing entities and rates; total-rate cap test
  Deal Comparables      — recent Texas PID financings for reference (not read by the model)

    build_inputs_template(path, cfg)   -> writes a styled workbook populated from cfg
    load_inputs(path)                  -> ModelConfig
"""

from __future__ import annotations

import warnings
from datetime import date, datetime
from typing import Optional

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

from .config import ModelConfig, OverlappingTax, Phase, Product, sample_config

# ── House style (matches Viridian PID Inputs / reimbursement report) ─────────
NAVY, MID_BLUE, LIGHT = "1F4E79", "2E6FA3", "D6E4F0"
YELLOW, GRAY_TXT = "FFF2CC", "595959"
_THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)

F_TITLE = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
F_SUB = Font(name="Calibri", size=9, italic=True, color=GRAY_TXT)
F_HDR = Font(name="Calibri", size=9, bold=True, color=NAVY)
F_SECTION = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
F_LABEL = Font(name="Calibri", size=10)
F_INPUT = Font(name="Calibri", size=10, bold=True)
F_CALC = Font(name="Calibri", size=10, color="000000")
F_NOTE = Font(name="Calibri", size=9, italic=True, color=GRAY_TXT)

FILL = lambda c: PatternFill("solid", fgColor=c)  # noqa: E731

FMT_DOLLAR = '"$"#,##0'
FMT_DOLLAR2 = '"$"#,##0.00'
FMT_RATE = '"$"0.0000'
FMT_PCT = "0.000%"
FMT_DATE = "mm/dd/yyyy"
FMT_NUM = "#,##0"
FMT_X = '0.00"x"'

MAX_PRODUCTS = 8
MAX_PHASES = 10
MAX_TAXES = 12

# ── Inputs sheet definition ──────────────────────────────────────────────────
# (label, range_name, attribute, format, note)   — None label = section header
INPUT_ROWS: list[tuple] = [
    ("Basic Inputs", None, None, None, None),
    ("PID Name", "TITLE1", "pid_name", None, "Appears on every output page"),
    ("Analysis", "TITLE2", "analysis_type", None, None),
    ("City (Issuer)", "CITY", "city", None, "Issuer of the special assessment revenue bonds"),
    ("County", "TITLE3", "county", None, None),
    ("Developer", "DEVELOPER_NAME", "developer", None, "Memo addressee"),
    ("Bond Title", "BOND_TITLE", "bond_title", None, None),
    ("Date of the Analysis", "DATE_OF_ANALYSIS", "analysis_date", FMT_DATE, None),

    ("Equivalent Tax Rate", None, None, None, None),
    ("PID Rate Method", "PID_RATE_METHOD", "pid_rate_method", None,
     "Input = use TAX_RATE_PID; Overlapping Cap = MAX_TOTAL_TAX_RATE less overlapping taxes"),
    ("Tax Rate for PID (per $100 AV)", "TAX_RATE_PID", "tax_rate_pid", FMT_RATE,
     "Annual installment = ASP / 100 x this rate"),
    ("Maximum Total Tax Rate (per $100 AV)", "MAX_TOTAL_TAX_RATE", "max_total_tax_rate", FMT_RATE,
     "City PID policy cap on overlapping taxes + PID (commonly ~$3.00)"),

    ("Fees", None, None, None, None),
    ("Costs of Issuance", "COI_FEE", "coi_fee", FMT_PCT, "% of par; legal, MA, trustee, AG, appraisal, SAP"),
    ("Underwriter's Discount", "UWD_FEE", "uwd_fee", FMT_PCT, "% of par; Texas PID deals typically 2.5% - 3.0%"),
    ("Administrative Fund (initial deposit)", "ADMIN_FEE", "admin_fee", FMT_PCT, "% of par"),
    ("City Development / PID Fee", "CITY_FEE", "city_fee", FMT_PCT, "% of par, if the city charges one"),
    ("Fixed Administration Fee Per Lot", "ADMIN_FEE_FIXED_LOT", "admin_fee_fixed_lot", FMT_DOLLAR2,
     "$/lot/yr collected in the annual installment"),
    ("Administrative Expense Escalator", "ADMIN_ESCALATOR", "admin_escalator", FMT_PCT, "Annual growth"),
    ("Delinquency & Prepayment Reserve Deposit", "DP_RESERVE_DEPOSIT", "dp_reserve_deposit", FMT_PCT,
     "% of par at closing; topped up from additional interest"),
    ("Reserve Fund Method", "RESERVE_METHOD", "reserve_method", None,
     "125% Average Annual DS | Maximum Annual DS | 10% of Par | Lesser of 3-Prong"),

    ("Structuring Assumptions", None, None, None, None),
    ("Bond Interest Rate", "BONDRATE", "bond_rate", FMT_PCT,
     "Coupon; also the assessment interest rate (see Deal Comparables)"),
    ("Rate Add (sizing haircut)", "RATE_ADD", "rate_add", FMT_PCT,
     "Added to the discount rate when sizing; 0% = size to full capacity"),
    ("Additional Interest (PID Act §372.018)", "ADDITIONAL_INTEREST", "additional_interest", FMT_PCT,
     "Up to 0.50% on outstanding assessments"),
    ("Final Maturity (years)", "FINAL_MAT_YRS", "final_mat_yrs", FMT_NUM, None),
    ("Principal Maturity (month)", "PRIN_MATURITY", "prin_maturity", FMT_NUM, "Texas PID bonds: September 1"),
    ("Interest Maturity (month)", "INT_MATURITY", "int_maturity", FMT_NUM, "Other semiannual date: March 1"),
    ("Capitalized Interest Period (months)", "CAPI_PERIOD", "capi_period", FMT_NUM,
     "Default for every series; override by series on Developer Information"),
    ("Years to First Par Call", "FIRST_PAR_CALL_YRS", "first_par_call_yrs", FMT_NUM,
     "Recent deals: 7-10 years at par"),
    ("Tax Collection %", "TAX_COLLECTION_PRC", "tax_collection", FMT_PCT, "Applied to sizing revenues"),
    ("Debt Service Coverage Factor", "DSC_FACTOR", "dsc_factor", "0.00", "1.00 = size to full installment"),
    ("Principal Rounding ($)", "ROUND_TO", "round_to", FMT_DOLLAR, "Recent Texas PID deals use $1,000"),
    ("Minimum Value-to-Lien", "MIN_VTL", "min_vtl", FMT_X, "City policy; often 3:1 (some 2:1 with TIRZ)"),
    ("Size Bonds to Improvement Costs?", "SIZE_TO_COSTS", "size_to_costs", None,
     "Yes = issue only what the costs require"),
    ("Cap Bonds at Value-to-Lien?", "SIZE_TO_VTL", "size_to_vtl", None, "Yes = par no greater than value / MIN_VTL"),

    ("Development Assumptions", None, None, None, None),
    ("Inflation Rate for Home Prices", "INFLATION_ASP", "inflation_asp", FMT_PCT, None),
    ("Inflation Start (year)", "INFLATION_START", "inflation_start", "0", None),
    ("Months from Lot Delivery to First Closing", "MONTHS_LOT_TO_CLOSE", "months_lot_to_close", FMT_NUM, None),
    ("Assessed Value Lag (years)", "AV_LAG_YEARS", "av_lag_years", FMT_NUM, "Closing year to tax roll year"),
    ("Projection Horizon (years)", "PROJECTION_YEARS", "projection_years", FMT_NUM, None),
]

INPUT_FIRST_ROW = 5
CHOICES = {
    "PID_RATE_METHOD": '"Input,Overlapping Cap"',
    "RESERVE_METHOD": '"125% Average Annual DS,Maximum Annual DS,10% of Par,Lesser of 3-Prong"',
    "SIZE_TO_COSTS": '"Yes,No"',
    "SIZE_TO_VTL": '"Yes,No"',
}

# Developer Information layout
DI_PROD_HDR_ROW, DI_PROD_FIRST_COL = 5, 3          # product grid header row / first product col (C)
DI_PROD_FIELDS = [("Product Name", "name", None),
                  ("Average Selling Price ($)", "avg_selling_price", FMT_DOLLAR),
                  ("As-Developed Lot Value ($)", "lot_value", FMT_DOLLAR),
                  ("Absorption (homes / month)", "absorption_per_month", "0.0")]
DI_PHASE_HDR_ROW = 13
DI_PHASE_FIXED = [("Improvement Area / Series", 24, None),
                  ("Lot Delivery Date", 14, FMT_DATE),
                  ("Bond Delivery Date", 14, FMT_DATE),
                  ("Improvement Costs ($)", 16, FMT_DOLLAR),
                  ("CAPI Months (blank = default)", 14, FMT_NUM)]


def _title(ws, text: str, sub: str, last_col: int):
    ws.merge_cells(start_row=1, start_column=2, end_row=1, end_column=last_col)
    c = ws.cell(row=1, column=2, value=text)
    c.font, c.fill, c.alignment = F_TITLE, FILL(NAVY), Alignment(vertical="center")
    ws.row_dimensions[1].height = 22
    s = ws.cell(row=2, column=2, value=sub)
    s.font = F_SUB


def _hdr(ws, row: int, col: int, text: str):
    c = ws.cell(row=row, column=col, value=text)
    c.font, c.fill, c.border = F_HDR, FILL(LIGHT), BORDER
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    return c


def _section(ws, row: int, text: str, first_col: int, last_col: int):
    for col in range(first_col, last_col + 1):
        ws.cell(row=row, column=col).fill = FILL(MID_BLUE)
    c = ws.cell(row=row, column=first_col, value=text)
    c.font = F_SECTION


def _input_cell(ws, row: int, col: int, value, fmt: Optional[str]):
    c = ws.cell(row=row, column=col, value=value)
    c.font, c.fill, c.border = F_INPUT, FILL(YELLOW), BORDER
    if fmt:
        c.number_format = fmt
    return c


# ── Template builder ─────────────────────────────────────────────────────────
def build_inputs_template(path: str, cfg: Optional[ModelConfig] = None,
                          comparables: Optional[list[dict]] = None) -> str:
    cfg = cfg or sample_config()
    wb = openpyxl.Workbook()

    # Inputs ------------------------------------------------------------------
    ws = wb.active
    ws.title = "Inputs"
    ws.sheet_view.showGridLines = False
    for col, w in {"A": 2, "B": 44, "C": 44, "D": 26, "E": 62}.items():
        ws.column_dimensions[col].width = w
    _title(ws, f"{cfg.pid_name} — Inputs",
           "Edit the yellow Value cells only. Range names in column D are what the Python model reads.", 5)
    for col, txt in zip(range(2, 6), ["Title", "Value", "Range Name", "Notes"]):
        _hdr(ws, 4, col, txt)

    r = INPUT_FIRST_ROW
    for label, rng, attr, fmt, note in INPUT_ROWS:
        if rng is None:
            _section(ws, r, label, 2, 5)
        else:
            lc = ws.cell(row=r, column=2, value=label)
            lc.font, lc.border = F_LABEL, BORDER
            _input_cell(ws, r, 3, getattr(cfg, attr), fmt)
            ws.cell(row=r, column=3).alignment = Alignment(
                horizontal="left" if isinstance(getattr(cfg, attr), str) else "right")
            dc = ws.cell(row=r, column=4, value=rng)
            dc.font, dc.border = F_NOTE, BORDER
            nc = ws.cell(row=r, column=5, value=note)
            nc.font, nc.border = F_NOTE, BORDER
            if rng in CHOICES:
                dv = DataValidation(type="list", formula1=CHOICES[rng], allow_blank=False)
                ws.add_data_validation(dv)
                dv.add(ws.cell(row=r, column=3))
            wb.defined_names[rng] = openpyxl.workbook.defined_name.DefinedName(
                rng, attr_text=f"Inputs!$C${r}")
        r += 1
    ws.freeze_panes = "B5"

    # Developer Information -----------------------------------------------------
    di = wb.create_sheet("Developer Information")
    di.sheet_view.showGridLines = False
    n_phase_cols = len(DI_PHASE_FIXED) + MAX_PRODUCTS + 1
    last = 1 + n_phase_cols
    di.column_dimensions["A"].width = 2
    di.column_dimensions["B"].width = 30
    _title(di, "Developer Information",
           "Products (one column each) and improvement areas / bond series (one row each). "
           "Yellow cells are inputs; lot columns follow the product order above.", last)

    _section(di, 4, "PRODUCT SETUP  (one column per lot type, up to 8)", 2, 2 + MAX_PRODUCTS)
    _hdr(di, DI_PROD_HDR_ROW, 2, "Parameter")
    for i in range(MAX_PRODUCTS):
        _hdr(di, DI_PROD_HDR_ROW, DI_PROD_FIRST_COL + i, f"Product {i + 1}")
    for j, (label, attr, fmt) in enumerate(DI_PROD_FIELDS):
        row = DI_PROD_HDR_ROW + 1 + j
        c = di.cell(row=row, column=2, value=label)
        c.font, c.border = F_LABEL, BORDER
        for i in range(MAX_PRODUCTS):
            val = getattr(cfg.products[i], attr) if i < len(cfg.products) else None
            _input_cell(di, row, DI_PROD_FIRST_COL + i, val, fmt)
    note_r = DI_PROD_HDR_ROW + 1 + len(DI_PROD_FIELDS)
    di.cell(row=note_r, column=2,
            value="Lot value is the as-developed (finished lot) value used for value-to-lien; "
                  "commonly ~20% of ASP or the appraised value.").font = F_NOTE
    _section(di, DI_PHASE_HDR_ROW - 1, "IMPROVEMENT AREAS / BOND SERIES  (one row per series, up to 10)",
             2, last)
    col = 2
    for label, w, _ in DI_PHASE_FIXED:
        _hdr(di, DI_PHASE_HDR_ROW, col, label)
        di.column_dimensions[openpyxl.utils.get_column_letter(col)].width = max(
            w, di.column_dimensions[openpyxl.utils.get_column_letter(col)].width or 0)
        col += 1
    lot_col0 = col
    for i in range(MAX_PRODUCTS):
        L = openpyxl.utils.get_column_letter(DI_PROD_FIRST_COL + i)
        _hdr(di, DI_PHASE_HDR_ROW, lot_col0 + i,
             f'=IF(\'Developer Information\'!{L}{DI_PROD_HDR_ROW + 1}="","Lots - Product {i + 1}",'
             f'"Lots - "&\'Developer Information\'!{L}{DI_PROD_HDR_ROW + 1})')
        di.column_dimensions[openpyxl.utils.get_column_letter(lot_col0 + i)].width = 13
    tot_col = lot_col0 + MAX_PRODUCTS
    _hdr(di, DI_PHASE_HDR_ROW, tot_col, "Total Lots")
    di.column_dimensions[openpyxl.utils.get_column_letter(tot_col)].width = 11

    names = [p.name for p in cfg.products]
    for k in range(MAX_PHASES):
        row = DI_PHASE_HDR_ROW + 1 + k
        ph = cfg.phases[k] if k < len(cfg.phases) else None
        vals = [ph.name, ph.lot_delivery, ph.bond_delivery, ph.improvement_costs, ph.capi_months] \
            if ph else [None] * 5
        for j, (v, (_, _, fmt)) in enumerate(zip(vals, DI_PHASE_FIXED)):
            _input_cell(di, row, 2 + j, v, fmt)
        for i in range(MAX_PRODUCTS):
            v = ph.lots.get(names[i]) if (ph and i < len(names)) else None
            _input_cell(di, row, lot_col0 + i, v, FMT_NUM)
        a = openpyxl.utils.get_column_letter(lot_col0)
        b = openpyxl.utils.get_column_letter(lot_col0 + MAX_PRODUCTS - 1)
        t = di.cell(row=row, column=tot_col, value=f"=SUM({a}{row}:{b}{row})")
        t.font, t.border, t.number_format = F_CALC, BORDER, FMT_NUM
    trow = DI_PHASE_HDR_ROW + 1 + MAX_PHASES
    c = di.cell(row=trow, column=2, value="Total")
    c.font = F_HDR
    for colx in [5] + list(range(lot_col0, tot_col + 1)):
        L = openpyxl.utils.get_column_letter(colx)
        cc = di.cell(row=trow, column=colx,
                     value=f"=SUM({L}{DI_PHASE_HDR_ROW + 1}:{L}{trow - 1})")
        cc.font, cc.fill, cc.border = F_HDR, FILL(LIGHT), BORDER
        cc.number_format = FMT_DOLLAR if colx == 5 else FMT_NUM
    di.cell(row=trow + 2, column=2,
            value="Each row is one improvement area financed by its own series of bonds. "
                  "Assessments are levied on the lots in that area when its bonds are issued.").font = F_NOTE
    di.freeze_panes = "C6"

    # Overlapping Taxes ---------------------------------------------------------
    ot = wb.create_sheet("Overlapping Taxes")
    ot.sheet_view.showGridLines = False
    for col, w in {"A": 2, "B": 40, "C": 16, "D": 18, "E": 50}.items():
        ot.column_dimensions[col].width = w
    _title(ot, "Overlapping Taxes", "Rates per $100 of assessed value for the taxing units overlapping the PID.", 5)
    for col, txt in zip(range(2, 5), ["Taxing Entity", "Type", "Tax Rate"]):
        _hdr(ot, 4, col, txt)
    for k in range(MAX_TAXES):
        row = 5 + k
        t = cfg.overlapping_taxes[k] if k < len(cfg.overlapping_taxes) else None
        _input_cell(ot, row, 2, t.entity if t else None, None)
        _input_cell(ot, row, 3, t.entity_type if t else None, None)
        _input_cell(ot, row, 4, t.rate if t else None, FMT_RATE)
    end = 5 + MAX_TAXES
    rows = [
        ("Total Overlapping Tax Rate", f"=SUM(D5:D{end - 1})", None),
        ("Maximum Total Tax Rate (Inputs)", "=MAX_TOTAL_TAX_RATE", None),
        ("Room for PID Equivalent Rate", f"=D{end + 1}-D{end}", "Maximum PID rate the cap allows"),
        ("PID Equivalent Rate in Model",
         '=IF(PID_RATE_METHOD="Input",TAX_RATE_PID,MAX(0,D{0}))'.format(end + 2), None),
        ("Total Tax Rate incl. PID", f"=D{end}+D{end + 3}", None),
        ("Within Cap?", f'=IF(D{end + 4}<=D{end + 1}+0.00001,"Yes","No — exceeds cap")', None),
    ]
    for j, (lab, f, note) in enumerate(rows):
        row = end + j
        a = ot.cell(row=row, column=2, value=lab)
        a.font, a.fill, a.border = F_HDR, FILL(LIGHT), BORDER
        c = ot.cell(row=row, column=4, value=f)
        c.font, c.border = F_INPUT, BORDER
        c.number_format = FMT_RATE
        if note:
            ot.cell(row=row, column=5, value=note).font = F_NOTE

    # Deal Comparables ----------------------------------------------------------
    dc = wb.create_sheet("Deal Comparables")
    dc.sheet_view.showGridLines = False
    cols = ["Sale Date", "Issuer", "PID / Improvement Area", "Par ($)", "Final Maturity",
            "Call", "Top Coupon / Yield", "MMD Spread", "Value-to-Lien", "Underwriter", "Source"]
    widths = [12, 22, 40, 14, 12, 16, 16, 11, 11, 16, 60]
    for i, (h, w) in enumerate(zip(cols, widths)):
        dc.column_dimensions[openpyxl.utils.get_column_letter(2 + i)].width = w
    dc.column_dimensions["A"].width = 2
    _title(dc, "Recent Texas PID Financings (reference only — not read by the model)",
           "Compiled from EMMA official statements / limited offering memoranda and public pricing books. "
           "Blank = not disclosed in the source reviewed.", 1 + len(cols))
    for i, h in enumerate(cols):
        _hdr(dc, 4, 2 + i, h)
    for k, deal in enumerate(comparables or []):
        row = 5 + k
        for i, h in enumerate(cols):
            v = deal.get(h)
            c = dc.cell(row=row, column=2 + i, value=v)
            c.font, c.border = F_LABEL, BORDER
            if h == "Par ($)" and v:
                c.number_format = FMT_DOLLAR
            if h == "Sale Date" and isinstance(v, (date, datetime)):
                c.number_format = FMT_DATE
    dc.freeze_panes = "C5"

    wb.save(path)
    return path


# ── Loader ───────────────────────────────────────────────────────────────────
def _to_date(v) -> Optional[date]:
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    return None


def load_inputs(path: str) -> ModelConfig:
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        wb = openpyxl.load_workbook(path, data_only=True)

    # Inputs: range name (col D) -> value (col C)
    ws = wb["Inputs"]
    named = {}
    for row in ws.iter_rows(min_row=INPUT_FIRST_ROW):
        rn = row[3].value if len(row) > 3 else None
        if isinstance(rn, str) and rn.strip():
            named[rn.strip()] = row[2].value

    cfg = ModelConfig()
    for _, rng, attr, _, _ in INPUT_ROWS:
        if rng is None or named.get(rng) is None:
            continue
        cur = getattr(cfg, attr)
        v = named[rng]
        if isinstance(cur, bool):
            v = str(v).lower().startswith("y")
        elif isinstance(cur, int):
            v = int(v)
        elif isinstance(cur, float):
            v = float(v)
        elif isinstance(cur, date):
            v = _to_date(v) or cur
        else:
            v = str(v).strip()
        setattr(cfg, attr, v)

    # Products
    di = wb["Developer Information"]
    products = []
    for i in range(MAX_PRODUCTS):
        col = DI_PROD_FIRST_COL + i
        name = di.cell(row=DI_PROD_HDR_ROW + 1, column=col).value
        if not name:
            continue
        g = lambda j: di.cell(row=DI_PROD_HDR_ROW + 1 + j, column=col).value  # noqa: E731
        products.append((i, Product(str(name).strip(), float(g(1) or 0), float(g(2) or 0),
                                    float(g(3) or 0))))
    cfg.products = [p for _, p in products]

    # Phases
    lot_col0 = 2 + len(DI_PHASE_FIXED)
    phases = []
    for k in range(MAX_PHASES):
        row = DI_PHASE_HDR_ROW + 1 + k
        name = di.cell(row=row, column=2).value
        if not name:
            continue
        lots = {}
        for i, p in products:
            n = di.cell(row=row, column=lot_col0 + i).value
            if n:
                lots[p.name] = int(n)
        if not lots:
            continue
        capi = di.cell(row=row, column=6).value
        phases.append(Phase(
            str(name).strip(),
            _to_date(di.cell(row=row, column=3).value),
            _to_date(di.cell(row=row, column=4).value),
            lots,
            float(di.cell(row=row, column=5).value or 0),
            int(capi) if capi not in (None, "") else None))
    cfg.phases = phases

    # Overlapping taxes
    ot = wb["Overlapping Taxes"]
    taxes = []
    for k in range(MAX_TAXES):
        row = 5 + k
        ent, typ, rate = (ot.cell(row=row, column=c).value for c in (2, 3, 4))
        if ent and rate is not None:
            taxes.append(OverlappingTax(str(ent), str(typ or ""), float(rate)))
    cfg.overlapping_taxes = taxes

    _validate(cfg)
    return cfg


def _validate(cfg: ModelConfig) -> None:
    errs = []
    if not cfg.products:
        errs.append("No products on Developer Information.")
    if not cfg.phases:
        errs.append("No improvement areas / series on Developer Information.")
    for ph in cfg.phases:
        if ph.lot_delivery is None or ph.bond_delivery is None:
            errs.append(f"{ph.name}: lot delivery and bond delivery dates are required.")
    for p in cfg.products:
        if p.avg_selling_price <= 0:
            errs.append(f"{p.name}: average selling price must be > 0.")
    if errs:
        raise ValueError("Input errors:\n  " + "\n  ".join(errs))
