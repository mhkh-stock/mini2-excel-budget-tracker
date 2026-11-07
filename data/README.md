# Mini #2 — data note

No external dataset. The 32 transactions in `src/build_workbook.py`
are a fixed, hand-written SAMPLE month (Jan 2026, fictional
household): salary + one freelance payment in, rent / groceries /
transport / utilities / food / entertainment / savings out.

The workbook is the deliverable; there is no separate data file to
keep in sync. Rebuild any time with `python3 src/build_workbook.py`.

## Data dictionary — `Transactions` sheet

| Column | Meaning |
|---|---|
| Date | Real Excel date (`yyyy-mm-dd`), Jan 2026 |
| Description | Free text, e.g. "Grocery — Imtiaz" |
| Category | One of: Income, Rent, Groceries, Transport, Utilities, Food Out, Entertainment, Savings |
| Type | `Income` or `Expense` (note: Savings rows are typed `Expense` — money moved to savings; see Limits in the main README) |
| Amount (PKR) | Whole PKR, always ≥ 0; never negative |

`Budget` sheet: Category | Monthly budget (PKR, typed) | Actual spent
(`SUMIF` from Transactions) | Remaining (budget − actual) |
Status (`OVER` if actual > budget, else `OK`), plus a TOTAL row.
`Summary` sheet: total income / total expenses / balance /
transaction count / savings rate (Savings ÷ income) — all formulas.
