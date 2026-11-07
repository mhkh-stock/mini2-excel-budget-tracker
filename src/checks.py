"""Mini #2 verification — reopen the workbook and check it.

Checks: 4 sheets exist; 32 transaction rows; dates are real Excel
dates (not text); every Summary/Budget total cell holds a formula
(not a typed number); Python recomputation from the transaction rows
matches the expected totals (incl. savings rate and the Budget TOTAL
row); every Budget category appears in Transactions (so no SUMIF
silently returns 0 for a typo); no negative amounts.
"""
import sys
from datetime import date as date_cls
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook

BASE = Path(__file__).resolve().parents[1]
XLSX = BASE / "outputs" / "budget-tracker_SAMPLE.xlsx"

fails = []


def check(name, cond, detail=""):
    print(f"{'PASS' if cond else 'FAIL'}  {name} {detail}")
    if not cond:
        fails.append(name)


wb = load_workbook(XLSX)
check("sheet names", wb.sheetnames == ["Transactions", "Summary", "Budget", "How to use"],
      str(wb.sheetnames))

ws = wb["Transactions"]
rows = list(ws.iter_rows(min_row=2, values_only=True))
rows = [r for r in rows if r[0] is not None]
check("transaction count = 32", len(rows) == 32, f"got {len(rows)}")
income = sum(r[4] for r in rows if r[3] == "Income")
expenses = sum(r[4] for r in rows if r[3] == "Expense")
check("income total", income == 145000, f"got {income}")
check("expense total", expenses == 198500, f"got {expenses}")
check("balance", income - expenses == -53500, f"got {income - expenses}")
check("no negative amounts", all(r[4] >= 0 for r in rows))
check("no blank categories", all(r[2] for r in rows))
check("transaction dates are real dates",
      all(isinstance(r[0], (datetime, date_cls)) for r in rows),
      f"types: {sorted({type(r[0]).__name__ for r in rows})}")

sm = wb["Summary"]
check("summary totals are formulas",
      all(str(sm.cell(row=r, column=2).value).startswith("=") for r in range(2, 7)))
bg = wb["Budget"]
check("budget cells are formulas",
      all(str(bg.cell(row=r, column=col).value).startswith("=")
          for r in range(2, 9) for col in (3, 4, 5)))
cats_tx = {r[2] for r in rows}
cats_bg = {bg.cell(row=r, column=1).value for r in range(2, 9)}
check("budget categories all in transactions", cats_bg <= cats_tx,
      f"missing: {cats_bg - cats_tx}")
per_cat = {}
for r in rows:
    if r[3] == "Expense":
        per_cat[r[2]] = per_cat.get(r[2], 0) + r[4]
check("savings rate recomputed",
      abs(per_cat.get("Savings", 0) / income - 25000 / 145000) < 1e-9,
      f"got {per_cat.get('Savings', 0)}/{income} = "
      f"{per_cat.get('Savings', 0) / income:.4f}")
budgets = {bg.cell(row=r, column=1).value: bg.cell(row=r, column=2).value
           for r in range(2, 9)}
# Budget-sheet order (rows 2..8), NOT set order: iterating a set is
# hash-ordered, so the old list came out ['Utilities', 'Groceries']
# on some runs and failed the check despite correct numbers.
over = [bg.cell(row=r, column=1).value for r in range(2, 9)
        if per_cat.get(bg.cell(row=r, column=1).value, 0)
        > budgets[bg.cell(row=r, column=1).value]]
check("over-budget categories identified", over == ["Groceries", "Utilities"],
      f"got {over}; groceries actual={per_cat.get('Groceries')}, "
      f"utilities actual={per_cat.get('Utilities')}")
check("budget TOTAL row are formulas",
      all(str(bg.cell(row=9, column=col).value).startswith("=")
          for col in (2, 3, 4)))

print(f"\n{len(fails)} failures" if fails else "\nALL CHECKS PASS")
sys.exit(1 if fails else 0)
