"""
assessment.py — Per-lot assessment sizing and annual installment schedules.

For each lot type the gross annual installment is set to the equivalent
tax rate times the average selling price:

    Gross Installment = ASP / 100 x PID equivalent tax rate

Each year that installment pays, in order: administrative expenses (fixed
$/lot, escalating), the PID Act additional interest (0.50% of the
outstanding assessment, funds the Delinquency & Prepayment Reserve),
interest at the bond coupon, and principal.  The maximum assessment per
lot is the present value of the installment net of administrative expenses
at (coupon + additional interest + any sizing haircut).
"""

from __future__ import annotations

from dataclasses import dataclass, field

import pandas as pd

from .config import ModelConfig, Product


@dataclass
class InstallmentRow:
    year_no: int
    bond_year: int                    # September 1 principal date of this bond year
    principal: float
    interest: float
    additional_interest: float
    admin: float
    balance: float

    @property
    def installment(self) -> float:
        return self.principal + self.interest + self.additional_interest + self.admin


@dataclass
class LotAssessment:
    product: Product
    gross_installment: float          # level annual installment at full ETR
    assessment: float                 # assessment levied per lot (after any scale-down)
    capacity: float                   # maximum supportable assessment per lot
    schedule: list[InstallmentRow] = field(default_factory=list)

    @property
    def equivalent_rate(self) -> float:
        return self.schedule[0].installment / self.product.avg_selling_price * 100 \
            if self.schedule else 0.0

    @property
    def avg_installment(self) -> float:
        return sum(r.installment for r in self.schedule) / len(self.schedule) \
            if self.schedule else 0.0

    def to_dataframe(self) -> pd.DataFrame:
        return pd.DataFrame([{
            "Year No.": r.year_no, "Payment Date": f"9/1/{r.bond_year}",
            "Principal": r.principal, "Interest": r.interest,
            "Additional Interest": r.additional_interest,
            "Administrative Expenses": r.admin,
            "Annual Installment": r.installment, "Principal Balance": r.balance,
        } for r in self.schedule])


def admin_expense(cfg: ModelConfig, year_no: int) -> float:
    return cfg.admin_fee_fixed_lot * (1 + cfg.admin_escalator) ** (year_no - 1)


def size_lot(cfg: ModelConfig, product: Product, first_bond_year: int,
             n_years: int, scale: float = 1.0) -> LotAssessment:
    """Size one lot's assessment and build its installment schedule.

    scale < 1.0 levies less than the maximum (e.g. when bonds are capped by
    costs or value-to-lien); installments fall proportionally (net of admin).
    """
    gross = product.avg_selling_price / 100 * cfg.effective_pid_rate
    rs = cfg.sizing_rate
    net = [max(0.0, gross - admin_expense(cfg, t)) * cfg.tax_collection / cfg.dsc_factor
           for t in range(1, n_years + 1)]
    capacity = sum(n / (1 + rs) ** t for t, n in enumerate(net, start=1))
    assessment = capacity * scale

    # Amortize at the sizing rate, then charge actual coupon + additional interest
    rows, bal = [], assessment
    for t in range(1, n_years + 1):
        pi_at_sizing = net[t - 1] * scale
        prin = pi_at_sizing - bal * rs
        if t == n_years:
            prin = bal
        interest = bal * cfg.bond_rate
        add_int = bal * cfg.additional_interest
        new_bal = bal - prin
        rows.append(InstallmentRow(t, first_bond_year + t - 1, prin, interest,
                                   add_int, admin_expense(cfg, t), max(0.0, new_bal)))
        bal = new_bal
    return LotAssessment(product, gross, assessment, capacity, rows)
