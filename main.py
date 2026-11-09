"""
main.py — Entry point for the Texas Public Improvement District (PID) model.

Usage
    python main.py                                  # run on inputs/Texas PID Inputs.xlsx
    python main.py --inputs "path/to/inputs.xlsx"   # run on a specific inputs workbook
    python main.py --template                       # (re)write a blank-ish sample inputs workbook
    python main.py --sample                         # run on the built-in sample project

Outputs (in --output, default 'output/')
    <PID name> PID Analysis.xlsx   multi-sheet Excel report
    *.csv                          the main tables
"""

from __future__ import annotations

import argparse
import csv
import os
import re
import sys
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from pid_model.bonds import BondModel                     # noqa: E402
from pid_model.config import ModelConfig, sample_config   # noqa: E402
from pid_model.development import DevelopmentProjection   # noqa: E402
from pid_model.inputs import build_inputs_template, load_inputs  # noqa: E402
from pid_model.report import build_excel_report           # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_INPUTS = os.path.join(HERE, "inputs", "Texas PID Inputs.xlsx")
COMPS_CSV = os.path.join(HERE, "research", "recent_texas_pid_deals.csv")


def load_comparables() -> list[dict]:
    if not os.path.exists(COMPS_CSV):
        return []
    rows = []
    with open(COMPS_CSV, newline="") as f:
        for r in csv.DictReader(f):
            r = {k: (v if v != "" else None) for k, v in r.items()}
            if r.get("Sale Date"):
                r["Sale Date"] = datetime.strptime(r["Sale Date"], "%Y-%m-%d").date()
            if r.get("Par ($)"):
                r["Par ($)"] = float(r["Par ($)"])
            rows.append(r)
    return rows


def run_model(cfg: ModelConfig, output_dir: str = "output", export_csv: bool = True) -> dict:
    print("=" * 72)
    print(f"  Texas PID Model — {cfg.pid_name}")
    print(f"  {cfg.city}, {cfg.county} County  |  analysis date {cfg.analysis_date:%m/%d/%Y}")
    print("=" * 72)
    print(f"  Equivalent tax rate: ${cfg.effective_pid_rate:.4f}  ({cfg.pid_rate_method})")
    print(f"  Coupon {cfg.bond_rate:.3%} + additional interest {cfg.additional_interest:.2%}"
          f" + haircut {cfg.rate_add:.2%}  |  {cfg.final_mat_yrs}-yr final  |  CAPI {cfg.capi_period} mo")

    dev = DevelopmentProjection(cfg).build()
    bm = BondModel(cfg, dev).run()

    print(f"\n  {'Series':<14}{'Area':<26}{'Lots':>6}{'Par':>14}{'Project Fund':>15}{'Costs':>14}"
          f"{'VTL':>7}  Binding")
    print("  " + "-" * 110)
    for s in bm.series:
        print(f"  {s.series_label:<14}{s.phase.name[:25]:<26}{s.lots:>6}{s.par:>14,.0f}"
              f"{s.su.project_fund:>15,.0f}{s.phase.improvement_costs:>14,.0f}{s.vtl:>7.2f}  {s.binding}")
        check = s.su.total_sources - s.su.total_uses
        if abs(check) > 0.01:
            print(f"     WARNING: sources less uses = {check:,.2f}")
    print("  " + "-" * 110)
    print(f"  {'Total':<40}{cfg.total_lots:>6}{bm.total_par:>14,.0f}{bm.total_project_fund:>15,.0f}"
          f"{sum(p.improvement_costs for p in cfg.phases):>14,.0f}")

    for s in bm.series:
        for prod, la in s.lot_assessments.items():
            print(f"    {s.series_label} {prod:<12} assessment/lot ${la.assessment:>10,.0f}   "
                  f"installment ${la.schedule[0].installment:>8,.0f}   ETR ${la.equivalent_rate:.4f}")

    total_rate = cfg.overlapping_rate + cfg.effective_pid_rate
    print(f"\n  Overlapping ${cfg.overlapping_rate:.4f} + PID ${cfg.effective_pid_rate:.4f} = "
          f"${total_rate:.4f} vs cap ${cfg.max_total_tax_rate:.4f} -> "
          f"{'OK' if total_rate <= cfg.max_total_tax_rate + 1e-6 else 'EXCEEDS CAP'}")

    os.makedirs(output_dir, exist_ok=True)
    safe = re.sub(r"[^\w\s\-&().]", "", cfg.pid_name).strip()
    xlsx = build_excel_report(cfg, dev, bm, os.path.join(output_dir, f"{safe} - PID Analysis.xlsx"))
    if export_csv:
        bm.annual_table().to_csv(os.path.join(output_dir, "bond_sizing_analysis.csv"), index=False)
        dev.annual_table().to_csv(os.path.join(output_dir, "project_and_pid_rev_analysis.csv"), index=False)
        for s in bm.series:
            for prod, la in s.lot_assessments.items():
                fn = re.sub(r"\W+", "_", f"installments_{s.series_label}_{prod}").strip("_").lower()
                la.to_dataframe().to_csv(os.path.join(output_dir, fn + ".csv"), index=False)
    print(f"\n  Excel report: {xlsx}")
    print("=" * 72)
    return {"config": cfg, "development": dev, "bonds": bm, "excel": xlsx}


def main():
    ap = argparse.ArgumentParser(description="Texas PID bond and assessment sizing model")
    ap.add_argument("--inputs", default=DEFAULT_INPUTS, help="inputs workbook (.xlsx)")
    ap.add_argument("--output", default=os.path.join(HERE, "output"), help="output folder")
    ap.add_argument("--template", action="store_true", help="write the sample inputs workbook and exit")
    ap.add_argument("--sample", action="store_true", help="run the built-in sample project")
    args = ap.parse_args()

    if args.template:
        os.makedirs(os.path.dirname(args.inputs), exist_ok=True)
        print("Wrote", build_inputs_template(args.inputs, sample_config(), load_comparables()))
        return
    cfg = sample_config() if args.sample else load_inputs(args.inputs)
    run_model(cfg, args.output)


if __name__ == "__main__":
    main()
