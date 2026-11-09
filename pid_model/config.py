"""
config.py — Model parameters for the Texas PID (Chapter 372) analysis.

Every scalar field maps 1-to-1 to a Range Name in column D of the
'Inputs' sheet of the Texas PID Inputs workbook (see inputs.py).
Product types, improvement-area phases and overlapping taxes come from
the 'Developer Information' and 'Overlapping Taxes' sheets.

Methodology follows Tierra's existing PID analyses (Pretium Farm,
Springhill South): the annual installment on each lot is set equal to an
equivalent tax rate applied to the lot's average selling price, and the
assessment (and bond par) is the present value of that installment net of
administrative expenses and the 0.50% PID Act additional interest.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from typing import Optional


# ── Developer information ─────────────────────────────────────────────────────

@dataclass
class Product:
    """One lot type (e.g. 40' front-load)."""
    name: str
    avg_selling_price: float          # finished home ASP ($)
    lot_value: float                  # as-developed lot value ($) for value-to-lien
    absorption_per_month: float       # home closings per month
    assessment_class_factor: float = 1.0  # optional override weight (1.0 = ETR-equal)


@dataclass
class Phase:
    """One improvement area / bond series.

    lots: {product name: number of lots in this improvement area}
    """
    name: str
    lot_delivery: date                # date lots are finished / delivered
    bond_delivery: date               # date this series of bonds closes
    lots: dict[str, int]
    improvement_costs: float          # authorized improvement costs allocable to the area
    capi_months: Optional[int] = None  # override of CAPI_PERIOD for this series


@dataclass
class OverlappingTax:
    entity: str
    entity_type: str
    rate: float                       # per $100 of AV


# ── Main configuration ────────────────────────────────────────────────────────

@dataclass
class ModelConfig:
    # Basic information
    pid_name: str = "Sample Ranch Public Improvement District"   # TITLE1
    analysis_type: str = "PID Analysis"                          # TITLE2
    city: str = "City of Sample"                                 # CITY
    county: str = "Collin"                                       # TITLE3
    developer: str = "Developer"                                 # DEVELOPER_NAME
    analysis_date: date = field(default_factory=date.today)      # DATE_OF_ANALYSIS
    bond_title: str = "Special Assessment Revenue Bonds"         # BOND_TITLE

    # Equivalent tax rate (per $100 of AV) used to set the annual installment
    pid_rate_method: str = "Input"        # PID_RATE_METHOD  ("Input" | "Overlapping Cap")
    tax_rate_pid: float = 0.90            # TAX_RATE_PID
    max_total_tax_rate: float = 3.00      # MAX_TOTAL_TAX_RATE (overlapping + PID cap)

    # Fees (as % of par unless noted)
    coi_fee: float = 0.035                # COI_FEE
    uwd_fee: float = 0.03                 # UWD_FEE
    admin_fee: float = 0.003              # ADMIN_FEE (initial deposit to Administrative Fund)
    city_fee: float = 0.0                 # CITY_FEE (city development / PID fee)
    admin_fee_fixed_lot: float = 75.0     # ADMIN_FEE_FIXED_LOT ($/lot/yr)
    admin_escalator: float = 0.02         # ADMIN_ESCALATOR
    dp_reserve_deposit: float = 0.0075    # DP_RESERVE_DEPOSIT (Delinquency & Prepayment)
    reserve_method: str = "125% Average Annual DS"  # RESERVE_METHOD
    # options: "125% Average Annual DS" | "Maximum Annual DS" | "10% of Par" | "Lesser of 3-Prong"

    # Structuring
    bond_rate: float = 0.065              # BONDRATE (coupon & assessment interest rate)
    rate_add: float = 0.0                 # RATE_ADD (sizing haircut added to coupon for PV)
    additional_interest: float = 0.005    # ADDITIONAL_INTEREST (PID Act §372.018, max 0.50%)
    final_mat_yrs: int = 30               # FINAL_MAT_YRS
    prin_maturity: int = 9                # PRIN_MATURITY (month)
    int_maturity: int = 3                 # INT_MATURITY (month of the other semiannual date)
    capi_period: int = 12                 # CAPI_PERIOD (months from delivery)
    first_par_call_yrs: int = 8           # FIRST_PAR_CALL_YRS
    tax_collection: float = 1.0           # TAX_COLLECTION_PRC (applied to sizing revenues)
    dsc_factor: float = 1.0               # DSC_FACTOR (sizing coverage; 1.0 = installment-based)
    round_to: int = 1_000                 # ROUND_TO (principal rounding, $)
    min_vtl: float = 2.0                  # MIN_VTL (minimum value-to-lien at issuance)
    size_to_costs: str = "Yes"            # SIZE_TO_COSTS (cap par at what the costs require)
    size_to_vtl: str = "No"               # SIZE_TO_VTL (cap par at value / MIN_VTL)

    # Development assumptions
    inflation_asp: float = 0.0            # INFLATION_ASP (annual home-price growth)
    inflation_start: int = 2027           # INFLATION_START (calendar year)
    months_lot_to_close: int = 6          # MONTHS_LOT_TO_CLOSE (lot delivery → first closing)
    av_lag_years: int = 1                 # AV_LAG_YEARS (closing year → tax roll year)
    projection_years: int = 40            # PROJECTION_YEARS

    # Developer information
    products: list[Product] = field(default_factory=list)
    phases: list[Phase] = field(default_factory=list)
    overlapping_taxes: list[OverlappingTax] = field(default_factory=list)

    # ── Derived ──────────────────────────────────────────────────────────────
    @property
    def overlapping_rate(self) -> float:
        return sum(t.rate for t in self.overlapping_taxes)

    @property
    def effective_pid_rate(self) -> float:
        """PID equivalent tax rate actually used for installments."""
        if self.pid_rate_method.strip().lower().startswith("overlap"):
            return max(0.0, self.max_total_tax_rate - self.overlapping_rate)
        return self.tax_rate_pid

    @property
    def sizing_rate(self) -> float:
        """Discount rate used to size assessments = coupon + additional interest + haircut."""
        return self.bond_rate + self.additional_interest + self.rate_add

    def product(self, name: str) -> Product:
        for p in self.products:
            if p.name == name:
                return p
        raise KeyError(f"Product '{name}' not found on Developer Information sheet")

    @property
    def total_lots(self) -> int:
        return sum(sum(ph.lots.values()) for ph in self.phases)


def sample_config() -> ModelConfig:
    """Illustrative two-phase, two-product project (generic — no client data)."""
    cfg = ModelConfig(analysis_date=date(2026, 11, 9))
    cfg.products = [
        Product("40' Lots", 450_000, 90_000, 2.0),
        Product("50' Lots", 525_000, 105_000, 2.0),
    ]
    cfg.phases = [
        Phase("Improvement Area #1", date(2028, 6, 1), date(2028, 9, 1),
              {"40' Lots": 30, "50' Lots": 22}, 1_900_000),
        Phase("Improvement Area #2", date(2029, 6, 1), date(2029, 9, 1),
              {"40' Lots": 30, "50' Lots": 21}, 2_500_000),
    ]
    cfg.overlapping_taxes = [
        OverlappingTax("City", "City", 0.4500),
        OverlappingTax("Collin County", "County", 0.1493),
        OverlappingTax("Independent School District", "ISD", 1.1500),
        OverlappingTax("Collin College", "College", 0.0812),
    ]
    return cfg
