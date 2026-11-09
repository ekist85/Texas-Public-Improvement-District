"""
bonds.py — Series sizing, debt service, and sources & uses for each
improvement area.

Par for each series is the smallest of:
  1. Assessment capacity  — lots x maximum assessment per lot (installment-based)
  2. Cost need (optional) — par required to fund the area's improvement costs
  3. Value-to-lien cap (optional) — area value / MIN_VTL
rounded down to ROUND_TO.  Assessments are then levied equal to par.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date

import pandas as pd
from dateutil.relativedelta import relativedelta

from .assessment import LotAssessment, size_lot
from .config import ModelConfig, Phase
from .development import DevelopmentProjection


def days360(a: date, b: date) -> int:
    d1, d2 = min(a.day, 30), b.day
    if d1 == 30 and d2 == 31:
        d2 = 30
    return (b.year - a.year) * 360 + (b.month - a.month) * 30 + (d2 - d1)


@dataclass
class DSRow:
    date: date
    principal: float
    coupon: float
    interest: float
    capi: float                       # interest paid from Capitalized Interest Account
    balance: float

    @property
    def total(self) -> float:
        return self.principal + self.interest

    @property
    def net(self) -> float:
        return self.total - self.capi


@dataclass
class SourcesUses:
    par: float
    project_fund: float
    reserve_fund: float
    dp_reserve: float
    capi: float
    admin_fund: float
    uwd: float
    coi: float
    city_fee: float

    @property
    def total_sources(self) -> float:
        return self.par

    @property
    def total_uses(self) -> float:
        return (self.project_fund + self.reserve_fund + self.dp_reserve + self.capi
                + self.admin_fund + self.uwd + self.coi + self.city_fee)


@dataclass
class Series:
    phase: Phase
    series_label: str
    delivery: date
    capi_end: date
    first_prin_year: int
    final_year: int
    par: float = 0.0
    capacity_par: float = 0.0
    cost_par: float = 0.0
    vtl_par: float = 0.0
    binding: str = ""
    lot_assessments: dict[str, LotAssessment] = field(default_factory=dict)
    ds: list[DSRow] = field(default_factory=list)
    su: SourcesUses | None = None
    value_at_issuance: float = 0.0

    @property
    def n_years(self) -> int:
        return self.final_year - self.first_prin_year + 1

    @property
    def lots(self) -> int:
        return sum(self.phase.lots.values())

    @property
    def vtl(self) -> float:
        return self.value_at_issuance / self.par if self.par else 0.0

    def annual_ds(self) -> dict[int, float]:
        """Bond-year (ending Sept 1) debt service, net of capitalized interest."""
        out: dict[int, float] = {}
        for r in self.ds:
            out[r.date.year] = out.get(r.date.year, 0.0) + r.net
        return out

    def annual_ds_gross(self) -> dict[int, float]:
        out: dict[int, float] = {}
        for r in self.ds:
            out[r.date.year] = out.get(r.date.year, 0.0) + r.total
        return out


class BondModel:
    def __init__(self, cfg: ModelConfig, dev: DevelopmentProjection):
        self.cfg = cfg
        self.dev = dev
        self.series: list[Series] = []

    # ── per-series build ─────────────────────────────────────────────────────
    def _schedule_dates(self, s: Series) -> list[date]:
        cfg = self.cfg
        months = sorted({cfg.prin_maturity, cfg.int_maturity})
        dates, d = [], date(s.delivery.year, 1, 1)
        end = date(s.final_year, cfg.prin_maturity, 1)
        while d <= end:
            if d.month in months and d > s.delivery:
                dates.append(d)
            d += relativedelta(months=1)
        return dates

    def _build_series(self, s: Series, scale: float) -> None:
        cfg = self.cfg
        s.lot_assessments = {
            prod: size_lot(cfg, cfg.product(prod), s.first_prin_year, s.n_years, scale)
            for prod, n in s.phase.lots.items() if n > 0}

        # Principal by bond year = sum of lot principal, rounded; final year plugs
        raw = {}
        for prod, la in s.lot_assessments.items():
            n = s.phase.lots[prod]
            for r in la.schedule:
                raw[r.bond_year] = raw.get(r.bond_year, 0.0) + r.principal * n
        prin, assigned = {}, 0.0
        years = sorted(raw)
        for y in years[:-1]:
            p = round(raw[y] / cfg.round_to) * cfg.round_to
            prin[y] = max(0.0, p)
            assigned += prin[y]
        prin[years[-1]] = s.par - assigned

        # Semiannual debt service
        s.ds, bal, prev = [], s.par, s.delivery
        for d in self._schedule_dates(s):
            frac = days360(prev, d) / 360
            interest = bal * cfg.bond_rate * frac
            p = prin.get(d.year, 0.0) if d.month == cfg.prin_maturity else 0.0
            capi = interest if d <= s.capi_end else 0.0
            bal -= p
            s.ds.append(DSRow(d, p, cfg.bond_rate, interest, capi, max(0.0, bal)))
            prev = d

    def _reserve(self, s: Series) -> float:
        cfg = self.cfg
        ann = [v for y, v in s.annual_ds_gross().items() if y >= s.first_prin_year]
        if not ann:
            return 0.0
        avg125 = 1.25 * sum(ann) / len(ann)
        mads = max(ann)
        ten = 0.10 * s.par
        m = cfg.reserve_method.lower()
        if m.startswith("max"):
            return mads
        if m.startswith("10"):
            return ten
        if m.startswith("lesser"):
            return min(avg125, mads, ten)
        return avg125

    def _sources_uses(self, s: Series) -> SourcesUses:
        cfg = self.cfg
        par = s.par
        capi = sum(r.capi for r in s.ds)
        su = SourcesUses(
            par=par, project_fund=0.0,
            reserve_fund=self._reserve(s),
            dp_reserve=par * cfg.dp_reserve_deposit,
            capi=capi,
            admin_fund=par * cfg.admin_fee,
            uwd=par * cfg.uwd_fee,
            coi=par * cfg.coi_fee,
            city_fee=par * cfg.city_fee)
        su.project_fund = par - (su.total_uses - su.project_fund)
        return su

    def _try_par(self, s: Series, par: float) -> SourcesUses:
        s.par = par
        scale = par / s.capacity_par if s.capacity_par else 0.0
        self._build_series(s, scale)
        return self._sources_uses(s)

    # ── run ─────────────────────────────────────────────────────────────────
    def run(self) -> "BondModel":
        cfg = self.cfg
        for ph in sorted(cfg.phases, key=lambda p: p.bond_delivery):
            capi_m = cfg.capi_period if ph.capi_months is None else ph.capi_months
            capi_end = ph.bond_delivery + relativedelta(months=capi_m)
            first_prin = capi_end.year + (1 if capi_end.month >= cfg.prin_maturity else 0)
            if date(first_prin, cfg.prin_maturity, 1) <= capi_end:
                first_prin += 1
            s = Series(ph, f"Series {ph.bond_delivery.year}", ph.bond_delivery,
                       capi_end, first_prin, ph.bond_delivery.year + cfg.final_mat_yrs)

            # 1. Capacity
            cap = 0.0
            for prod, n in ph.lots.items():
                if n > 0:
                    cap += n * size_lot(cfg, cfg.product(prod), first_prin, s.n_years).capacity
            s.capacity_par = cap
            limit, binding = int(cap // cfg.round_to) * cfg.round_to, "Assessment Capacity"

            # 2. Value-to-lien cap
            s.value_at_issuance = self.dev.phase_value(ph.name, ph.bond_delivery)
            s.vtl_par = s.value_at_issuance / cfg.min_vtl if cfg.min_vtl else float("inf")
            if cfg.size_to_vtl.lower().startswith("y") and s.vtl_par < limit:
                limit, binding = int(s.vtl_par // cfg.round_to) * cfg.round_to, "Value-to-Lien"

            # 3. Cost need — smallest par whose project fund covers the costs
            if cfg.size_to_costs.lower().startswith("y") and ph.improvement_costs > 0:
                lo, hi = 0, limit
                if self._try_par(s, hi).project_fund >= ph.improvement_costs:
                    while hi - lo > cfg.round_to:
                        mid = (lo + hi) // 2 // cfg.round_to * cfg.round_to
                        if mid in (lo, hi):
                            break
                        if self._try_par(s, mid).project_fund >= ph.improvement_costs:
                            hi = mid
                        else:
                            lo = mid
                    s.cost_par = hi
                    if hi < limit:
                        limit, binding = hi, "Improvement Costs"
                else:
                    s.cost_par = float("nan")    # capacity cannot fund full costs

            s.binding = binding
            s.su = self._try_par(s, limit)
            self.series.append(s)
        return self

    # ── roll-ups ─────────────────────────────────────────────────────────────
    @property
    def total_par(self) -> float:
        return sum(s.par for s in self.series)

    @property
    def total_project_fund(self) -> float:
        return sum(s.su.project_fund for s in self.series)

    def annual_table(self) -> pd.DataFrame:
        """Bond-year table: installments billed (by component) vs debt service."""
        cfg = self.cfg
        years = sorted({r.date.year for s in self.series for r in s.ds})
        rows = []
        for y in years:
            row = {"Bond Year (9/1)": y}
            prin_inst = int_inst = ai = admin = lots = 0.0
            tot_ds = 0.0
            for s in self.series:
                for prod, la in s.lot_assessments.items():
                    n = s.phase.lots[prod]
                    for r in la.schedule:
                        if r.bond_year == y:
                            prin_inst += r.principal * n
                            int_inst += r.interest * n
                            ai += r.additional_interest * n
                            admin += r.admin * n
                            lots += n
                ds_y = sum(r.net for r in s.ds if r.date.year == y)
                row[f"{s.series_label} Net DS"] = ds_y
                tot_ds += ds_y
            pi = (prin_inst + int_inst) * cfg.tax_collection
            row.update({
                "Assessed Lots": lots,
                "Installment - Principal": prin_inst,
                "Installment - Interest": int_inst,
                "Additional Interest": ai,
                "Administrative Expenses": admin,
                "Total Annual Installments": prin_inst + int_inst + ai + admin,
                "Total Net Debt Service": tot_ds,
                "P&I Coverage": pi / tot_ds if tot_ds else None,
            })
            rows.append(row)
        return pd.DataFrame(rows)
