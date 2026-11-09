# Texas Public Improvement Districts — Research Notes

*Prepared November 2026 for the Texas PID model. Sources are linked inline; deal-level
data is in `recent_texas_pid_deals.csv` and on the **Deal Comparables** tab of the inputs workbook.*

## How a Texas PID financing works (Local Government Code Ch. 372)

- **Creation.** Landowners petition the city (or county); the city creates the PID and adopts
  a **Service and Assessment Plan (SAP)** that lists authorized improvements, allocates costs
  to **Improvement Areas**, and sets the assessment roll.
- **Assessments, not taxes.** Each lot is levied a fixed assessment, repaid in **Annual
  Installments** billed with property taxes (billed ~October 1, delinquent after January 31/February 1).
- **What an Annual Installment contains:** principal + interest on the assessment, **additional
  interest of up to 0.50%** (PID Act §372.018 — funds the Delinquency & Prepayment Reserve, then
  redeems bonds), and **administrative expenses** (PID administrator, trustee, city costs).
- **Bonds.** The city issues *Special Assessment Revenue Bonds*, secured only by assessments in
  that Improvement Area — non-recourse to the city, typically **not rated**, sold through a
  limited offering to sophisticated investors, AG-approved.
- **Typical structure observed:** September 1 principal / March 1 & September 1 interest,
  30-year final maturity, 2–4 term bonds with mandatory sinking funds, par call in ~7–10 years,
  extraordinary call from prepayments, $100,000 minimum denominations (some smaller deals $25,000),
  principal in $1,000 increments.
- **Funds funded at closing:** Project Fund (improvements = developer reimbursement), Reserve Fund
  (MADS, 125% average annual DS, or 10% of par — Austin's Estancia IA#3 is exactly 10% of par),
  Capitalized Interest, Administrative Fund, Delinquency & Prepayment Reserve (some deals),
  costs of issuance, underwriter's discount.

## Recent deals (EMMA / public pricing materials)

| Deal | Takeaway |
|---|---|
| Princeton – Eastridge IA#4, Mar 2025, $15.3M | 5.625% top coupon, MMD + 147 bp, 3.1x oversubscribed; **VTL 2.0x**; 80¢ max gross ETR with a 45% TIRZ offset; DSRF ≈ 7.0% of par; UWD 3.0%; COI ≈ 4.6% |
| Gunter / Hutto / Providence Village comps, Feb–Mar 2025 | Spreads 149–160 bp over MMD; 30–35-year finals; 7–10-year par calls |
| Celina – Wells South NIA #6–7, Apr 2024, $7.7M | Three term bonds (2031 4.625%, 2044 5.50%, 2054 5.75%); 0.50% additional interest; $100k denominations; FMSbonds |
| Denison – Preston Harbor, Aug 2026, $22.4M | 7.17% TIC — 2026 spreads/yields are higher than early 2025 |
| Celina – Ariana IA 1–3, Sep 2026, $19.0M | Serial 2027–2056, callable 9/1/2034 at par |
| FMSbonds 2026 activity | Liberty Hills PID No. 1 (Anna, ~$24.8M), Platinum Ranch (Gunter, ~$50M), Manzano Mile (Marble Falls, ~$23.5M) — heavy North Texas volume |

## Implications built into the model

- Bond rate default **6.50%** (≈ 7.0% all-in with the 0.50% additional interest, matching
  Tierra's July 2026 memo assumption).
- Assessment per lot = PV of (ASP/100 × equivalent rate − admin) at coupon + additional interest.
  Checked against the July 2026 memo: $500k ASP, $0.9117 rate, $75 admin at 2% → **$55,424 per lot,
  identical to the memo**.
- Reserve fund method is selectable (125% average DS — Tierra's prior models — MADS, 10% of par,
  or lesser of the three).
- Value-to-lien and overlapping-tax-rate tests are reported for every series; both can be made
  binding on par.

## Caveats

- EMMA's search pages require a browser session, so deals were pulled from EMMA-hosted official
  statements found by search plus city pricing books. Several 2026 entries come from summaries
  that do not disclose coupons; verify on EMMA before relying on them.
- City PID policies vary (VTL minimums, ETR caps, maximum assessment per lot, TIRZ offsets) —
  confirm the specific city's policy for each project.
