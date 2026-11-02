"""Mini #2 verification — reopen the workbook and check it.

Checks: 4 sheets exist; 32 transaction rows; every Summary/Budget
total cell holds a formula (not a typed number); Python recomputation
from the transaction rows matches the expected totals; every Budget
category appears in Transactions (so no SUMIF silently returns 0 for
a typo); no negative amounts.
"""
import sys
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
over = [c for c in cats_bg
        if per_cat.get(c, 0) > dict((bg.cell(row=r, column=1).value,
                                     bg.cell(row=r, column=2).value)
                                    for r in range(2, 9))[c]]
check("over-budget categories identified", over == ["Groceries", "Utilities"],
      f"got {over}; groceries actual={per_cat.get('Groceries')}, "
      f"utilities actual={per_cat.get('Utilities')}")

print(f"\n{len(fails)} failures" if fails else "\nALL CHECKS PASS")
sys.exit(1 if fails else 0)
