# Texas Public Improvement District (PID) Model

Python model for sizing Chapter 372 special assessment revenue bonds and the
assessments behind them, with an Excel input workbook and a multi-sheet Excel report.
Built to the same methodology as Tierra's existing PID analyses (Pretium Farm,
Springhill South) and the same file conventions as the UT PID and Texas
Reimbursement models.

## Quick start

```bash
pip install -r requirements.txt
python main.py                    # runs inputs/Texas PID Inputs.xlsx -> output/
python main.py --inputs "path/to/My Project Inputs.xlsx"
python main.py --template         # rewrite the sample inputs workbook
python main.py --sample           # run the built-in sample without a workbook
```

## Inputs workbook (`inputs/Texas PID Inputs.xlsx`)

| Sheet | What goes there |
|---|---|
| **Inputs** | Title / Value / Range Name / Notes — edit the yellow Value cells. Equivalent tax rate, fees, coupon, CAPI, maturity, reserve method, VTL minimum, sizing switches. |
| **Developer Information** | Product grid (name, ASP, as-developed lot value, absorption) and one row per improvement area / bond series (lot delivery, bond delivery, improvement costs, CAPI override, lots by product). |
| **Overlapping Taxes** | Taxing entities and rates; shows room under the city's total-rate cap. |
| **Deal Comparables** | Recent Texas PID financings for reference (not read by the model). |

## Methodology

1. **Annual installment per lot** = ASP / 100 × PID equivalent tax rate (or the room left under
   the overlapping-tax cap).
2. **Maximum assessment per lot** = PV of (installment − administrative expenses) at
   coupon + 0.50% additional interest (+ optional sizing haircut). Each installment pays admin,
   additional interest (to the Delinquency & Prepayment Reserve), interest and principal.
3. **Par by series** = the smallest of assessment capacity, the par needed to fund the area's
   improvement costs (optional), and value ÷ minimum VTL (optional). Assessments are levied equal to par.
4. **Debt service**: March 1 / September 1 interest, September 1 principal matched to the
   assessment amortization; capitalized interest through the CAPI period.
5. **Sources & uses**: Project Fund (developer reimbursement) is the residual after the reserve fund,
   D&P reserve, CAPI, administrative fund, city fee, underwriter's discount and costs of issuance.
6. **Tests**: overlapping tax rate vs. cap, value-to-lien at delivery, annual P&I coverage.

Validation: with the July 2026 memo assumptions ($500,000 ASP, $0.9117 rate, $75 admin at 2%,
6.50% + 0.50%) the model's per-lot assessment is $55,424 — identical to the memo.

## Output (`output/<PID name> - PID Analysis.xlsx`)

Summary · Sources and Uses · Bond Sizing Analysis · Debt Service · Assessment Roll ·
Annual Installments · Project and PID Rev Analysis · Tax Rate and VTL Tests (plus CSVs).

## Package layout

```
main.py                 entry point
pid_model/config.py     ModelConfig, Product, Phase, OverlappingTax, sample project
pid_model/inputs.py     builds / reads the inputs workbook
pid_model/development.py absorption and taxable value
pid_model/assessment.py per-lot assessment and installment schedules
pid_model/bonds.py      series sizing, debt service, sources & uses, coverage
pid_model/report.py     Excel report
research/               Texas PID research notes and recent deal comparables
```

## Not yet modeled

Term-bond / yield-curve coupons and original issue premium/discount; TIRZ offsets to the
annual installment; major-improvement-area bonds spread across several areas; prepayment and
delinquency stress cases.
