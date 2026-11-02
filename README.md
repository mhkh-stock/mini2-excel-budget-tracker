# Mini #2 — Excel Budget & Expense Tracker

**SAMPLE data only.** Fictional household, fictional amounts, Jan 2026.
Never claimed as real client or real business data.

The skill this mini proves: building a tracker a normal person can
actually keep using — type transactions, everything else updates
itself. Every total in the workbook is a **live formula**, not a typed
number: change a transaction and Summary + Budget move with it.

## Deliverable

`outputs/budget-tracker_SAMPLE.xlsx` — 4 sheets:

| Sheet | What it does |
|---|---|
| `Transactions` | 32 rows: date, description, category, type, amount (PKR). Filter on, frozen header, print-fit. |
| `Summary` | Total income, total expenses, balance, transaction count, savings rate — all `SUMIF` / `COUNTA` formulas. |
| `Budget` | Per-category monthly budget vs actual (`SUMIF` from Transactions), remaining, and an `OVER` / `OK` status flag. |
| `How to use` | One paragraph: clear the Transactions rows, keep row 1, type your own month. |

## Verified numbers (Jan 2026 SAMPLE month)

| Measure | Value |
|---|---|
| Transactions | 32 |
| Total income | PKR 145,000 |
| Total expenses | PKR 198,500 |
| Balance | PKR −53,500 (this sample month overspends — on purpose: it proves the flags work) |
| Over-budget categories | Groceries (PKR 64,400 vs 55,000 budget), Utilities (PKR 20,500 vs 20,000) |
| Savings rate | PKR 25,000 / 145,000 ≈ 17.2% |

## Verification (done before delivery)

`python3 src/checks.py` — 11 checks, all passing: sheet names, 32
transactions, income/expense/balance recomputed from the rows,
no negative amounts, no blank categories, every Summary and Budget
total cell holds a formula, every budget category appears in
Transactions (so no `SUMIF` silently returns 0 for a typo), and the
over-budget list matches hand-checked category sums. One note from
verification: the first draft of the checks carried wrong hand-added
expectations (PKR 191,200); the workbook was right and the checks
were corrected to the recomputed PKR 198,500 — the formulas, not the
guess, are the source of truth.

## How to rebuild

1. `python3 src/build_workbook.py` — regenerates the workbook.
2. `python3 src/checks.py` — re-verifies it.

Requires openpyxl (`pip install openpyxl`).
