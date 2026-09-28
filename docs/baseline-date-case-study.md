# The baseline-date trap in input-cost analysis

*A two-minute read. Every number below comes from public BLS/EIA series via
FRED and can be reproduced with the commands at the end.*

## The question

You buy copper, steel, resin and corrugated. Your standard costs were last
struck on different dates — some components repriced last quarter, some haven't
been touched since 2022. **How much has your unit cost actually moved?**

The reflex is to pull each input's 12-month change and apply it. That is the
wrong arithmetic, and the error is not small.

## What the reflex misses

A 12-month change answers "how much has this input moved in the last year." It
does not answer "how much has this input moved **since the date my standard cost
was set**." Those are only the same number if every standard was struck exactly
twelve months ago.

Here is a six-component product with realistic, staggered baseline dates:

| Component | Baseline set | Change since **that** date | Flat 12-month | Δ per unit |
|---|---|---:|---:|---:|
| injection-moulded housing | 2024-03 | +6.58% | +6.04% | +$0.01 |
| steel mounting bracket | 2025-09 | +14.44% | +14.35% | +$0.00 |
| **copper wiring harness** | **2022-06** | **+51.91%** | **+32.99%** | **+$0.48** |
| **corrugated shipping carton** | **2026-06** | **+4.96%** | **−3.47%** | **+$0.08** |
| adhesive and solvents | 2025-01 | +10.17% | +13.30% | −$0.02 |
| **plant electricity per unit** | **2023-11** | **+16.67%** | **+3.16%** | **+$0.09** |

Totals:

```
baseline unit cost                $11.10
true increase (own baselines)     +$2.25   (+20.3%)
what a flat 12-month view shows   +$1.59   (+14.4%)
missed                            +$0.66/unit  = 29% of the real increase
```

**At 250,000 units a year that is $164,781 of cost exposure that never appears
in the analysis.**

## Why the errors point in different directions

Three distinct failure modes, all visible above:

1. **Stale baselines understate.** The copper harness was last priced in
   June 2022. Nonferrous metals are up 51.9% since then but only 33.0% over the
   last year — the flat view silently discards two and a half years of drift and
   understates that line by 58%.

2. **Fresh baselines overstate.** The carton was repriced in June 2026. A
   12-month window applies eight months of change that happened *before* the
   standard was set, and it flips the sign: the flat view reports paperboard
   **down 3.5%** when the cost has actually risen 5.0% since the baseline.

3. **Sign reversal is a property of the window, not the market.** Industrial
   natural gas is **−22.9%** measured from June 2022 and **+45.5%** measured
   from September 2025. Same series, same day, opposite conclusion. The window
   is doing the talking.

## The rule

> Price every line from the date its own standard was struck. A cost review
> whose baseline is "twelve months" is measuring the calendar, not your costs.

Two practical corollaries:

- **Show the date you priced from, per line.** If a series has no observation
  near the requested baseline, say so in the cell rather than quietly
  substituting the nearest window — a silently wrong number is worse than a
  flagged gap.
- **Indices are exposure, not invoices.** This arithmetic assumes your purchase
  price tracks its published index 1:1. Contracted or hedged inputs move less.
  Treat the output as where to go *look*, not as what you were billed.

## Reproduce it

```bash
git clone <this repo> && cd costwatch
python3 -m venv .venv && ./.venv/bin/pip install openpyxl
./.venv/bin/python src/build_workbook.py --bom data/mixed_baseline_bom.csv --out review.xlsx
```

Sources: BLS Producer Price Index series (WPU072, WPU3012, WPU102, WPU0911,
WPU061), BLS average price series APU000072610, EIA WTI via FRED. All public,
no API key. Every series ID, publisher and URL is listed on the workbook's
Sources sheet. Figures as of the 2026-08 print, retrieved 2026-09-27.
