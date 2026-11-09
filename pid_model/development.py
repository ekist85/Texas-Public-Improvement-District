"""
development.py — Home absorption and assessed-value projection.

Builds a monthly closing schedule for each product within each improvement
area (lots must be delivered before homes can close), then rolls it up to
calendar years and tax-roll years.  Feeds the 'Project and PID Rev Analysis'
and value-to-lien tests.
"""

from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass
from datetime import date

import pandas as pd
from dateutil.relativedelta import relativedelta

from .config import ModelConfig


def month_start(d: date) -> date:
    return date(d.year, d.month, 1)


@dataclass
class DevelopmentProjection:
    cfg: ModelConfig

    def build(self) -> "DevelopmentProjection":
        cfg = self.cfg
        # closings[(phase, product)][month] = homes closed that month
        self.closings: dict[tuple[str, str], dict[date, int]] = defaultdict(dict)

        start = min(month_start(ph.lot_delivery) for ph in cfg.phases)
        horizon = start + relativedelta(years=cfg.projection_years)

        for prod in cfg.products:
            # Lots for this product, released by phase in delivery order
            queue = sorted(
                [(ph.lot_delivery + relativedelta(months=cfg.months_lot_to_close),
                  ph.name, ph.lots.get(prod.name, 0))
                 for ph in cfg.phases if ph.lots.get(prod.name, 0) > 0],
                key=lambda t: t[0])
            if not queue:
                continue
            remaining = {name: n for _, name, n in queue}
            carry = 0.0                               # fractional pace carry-over
            m = month_start(queue[0][0])
            while m < horizon and sum(remaining.values()) > 0:
                carry += prod.absorption_per_month
                can_close = int(carry + 1e-9)
                for open_dt, name, _ in queue:        # oldest open phase first
                    if can_close <= 0:
                        break
                    if month_start(open_dt) <= m and remaining[name] > 0:
                        n = min(can_close, remaining[name])
                        self.closings[(name, prod.name)][m] = n
                        remaining[name] -= n
                        can_close -= n
                        carry -= n
                # pace does not bank: pace left unused because no lots were open is dropped
                if can_close > 0:
                    carry = 0.0
                m += relativedelta(months=1)
        return self

    # ── helpers ──────────────────────────────────────────────────────────────
    def asp(self, product: str, year: int) -> float:
        p = self.cfg.product(product)
        yrs = max(0, year - self.cfg.inflation_start + 1)
        return p.avg_selling_price * (1 + self.cfg.inflation_asp) ** yrs

    def homes_closed_by(self, phase: str, product: str, as_of: date) -> int:
        return sum(n for m, n in self.closings.get((phase, product), {}).items()
                   if m < as_of)

    def phase_value(self, phase_name: str, as_of: date) -> float:
        """Estimated value of an improvement area: closed homes at ASP plus
        remaining lots at as-developed lot value."""
        ph = next(p for p in self.cfg.phases if p.name == phase_name)
        val = 0.0
        for prod_name, lots in ph.lots.items():
            homes = self.homes_closed_by(phase_name, prod_name, as_of)
            val += homes * self.asp(prod_name, as_of.year)
            val += (lots - homes) * self.cfg.product(prod_name).lot_value
        return val

    def annual_table(self) -> pd.DataFrame:
        """Calendar-year closings and tax-roll AV for the whole district."""
        cfg = self.cfg
        first = min(ph.lot_delivery.year for ph in cfg.phases)
        rows = []
        cum = defaultdict(int)
        for i in range(cfg.projection_years):
            yr = first + i
            row = {"Year No.": i + 1, "Year": yr}
            homes_yr_total, cum_total, basis = 0, 0, 0.0
            for prod in cfg.products:
                n = sum(v for (ph, pr), months in self.closings.items() if pr == prod.name
                        for m, v in months.items() if m.year == yr)
                cum[prod.name] += n
                row[f"{prod.name} Homes"] = n
                row[f"{prod.name} Cumulative"] = cum[prod.name]
                homes_yr_total += n
                cum_total += cum[prod.name]
                basis += cum[prod.name] * self.asp(prod.name, yr)
            # lots delivered but not yet built on, at lot value
            vacant = 0.0
            for ph in cfg.phases:
                if ph.lot_delivery.year <= yr:
                    for prod_name, lots in ph.lots.items():
                        built = self.homes_closed_by(ph.name, prod_name, date(yr + 1, 1, 1))
                        vacant += (lots - built) * cfg.product(prod_name).lot_value
            row["Homes Per Year"] = homes_yr_total
            row["Cumulative Homes"] = cum_total
            row["Home Value"] = basis
            row["Vacant Lot Value"] = vacant
            row["Total Taxable Basis"] = basis + vacant
            row["Tax Roll Year"] = yr + cfg.av_lag_years
            rows.append(row)
        return pd.DataFrame(rows)
