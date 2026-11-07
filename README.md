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
| Total monthly budget (all categories) | PKR 195,000 (Budget TOTAL row: actual PKR 198,500, remaining PKR −3,500) |
| Savings rate | PKR 25,000 / 145,000 ≈ 17.2% |

## Limits

- One month only (Jan 2026 SAMPLE): no multi-month rollover or yearly view.
- "Savings" is an expense category here (money moved to savings), so it
  counts inside total expenses and the balance — that is why the sample
  balance is negative even though PKR 25,000 was saved.
- The workbook's formulas have no cached values until it is opened in
  Excel / LibreOffice (openpyxl writes formulas, not results); anything
  reading the file without a spreadsheet engine sees blank totals.
- Amounts are whole PKR; no currency conversion.

## Verification (done before delivery)

`python3 src/checks.py` — 14 checks, all passing: sheet names, 32
transactions, dates are real Excel dates (not text), income/expense/
balance/savings-rate recomputed from the rows, no negative amounts,
no blank categories, every Summary and Budget total cell (including
the Budget TOTAL row) holds a formula, every budget category appears
in Transactions (so no `SUMIF` silently returns 0 for a typo), and the
over-budget list matches hand-checked category sums. One note from
verification: the first draft of the checks carried wrong hand-added
expectations (PKR 191,200); the workbook was right and the checks
were corrected to the recomputed PKR 198,500 — the formulas, not the
guess, are the source of truth. A second note from the 2026-11-07
review: the over-budget check used to iterate a Python `set`, so its
order — and the pass/fail — changed with the hash seed (it failed
under `PYTHONHASHSEED=42` with correct numbers); it now walks the
Budget sheet rows in order and passes under every seed tested.

## How to rebuild

1. `python3 src/build_workbook.py` — regenerates the workbook.
2. `python3 src/checks.py` — re-verifies it.

Requires openpyxl (`pip install -r requirements.txt` —
`openpyxl>=3.1.2,<3.2`, verified with 3.1.2; latest in range is 3.1.5).
