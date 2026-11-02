"""Mini #2 — Excel Budget & Expense Tracker builder (SAMPLE data).

Builds outputs/budget-tracker_SAMPLE.xlsx with LIVE formulas (no
hard-coded totals): Transactions + Summary + Budget sheets. Every
total in the workbook is a formula, so a user who edits a transaction
sees the summaries move — that is the skill being demonstrated.

SAMPLE data only: fictional household, fictional amounts, Jan 2026.
"""
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

BASE = Path(__file__).resolve().parents[1]
OUT = BASE / "outputs" / "budget-tracker_SAMPLE.xlsx"

BOLD = Font(bold=True)
HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(bold=True, color="FFFFFF")
MONEY = '#,##0'

# (date, description, category, type, amount PKR)
TRANSACTIONS = [
    ("2026-01-01", "Monthly salary", "Income", "Income", 120000),
    ("2026-01-01", "Freelance design work", "Income", "Income", 25000),
    ("2026-01-02", "House rent", "Rent", "Expense", 35000),
    ("2026-01-03", "Grocery — Imtiaz", "Groceries", "Expense", 18500),
    ("2026-01-04", "Petrol", "Transport", "Expense", 8000),
    ("2026-01-05", "Electricity bill", "Utilities", "Expense", 9500),
    ("2026-01-06", "Grocery — weekly top-up", "Groceries", "Expense", 7200),
    ("2026-01-07", "Dinner with family", "Food Out", "Expense", 4500),
    ("2026-01-08", "Internet bill", "Utilities", "Expense", 3500),
    ("2026-01-09", "Bus / ride fares", "Transport", "Expense", 3200),
    ("2026-01-10", "Grocery — weekly top-up", "Groceries", "Expense", 6800),
    ("2026-01-11", "Cinema tickets", "Entertainment", "Expense", 2800),
    ("2026-01-12", "Mobile load / package", "Utilities", "Expense", 1500),
    ("2026-01-13", "Office lunch", "Food Out", "Expense", 1800),
    ("2026-01-14", "Grocery — weekly top-up", "Groceries", "Expense", 7400),
    ("2026-01-15", "Transferred to savings", "Savings", "Expense", 20000),
    ("2026-01-16", "Petrol", "Transport", "Expense", 7500),
    ("2026-01-17", "Birthday gift", "Entertainment", "Expense", 5000),
    ("2026-01-18", "Grocery — weekly top-up", "Groceries", "Expense", 6900),
    ("2026-01-19", "Gas bill", "Utilities", "Expense", 4200),
    ("2026-01-20", "Takeaway dinner", "Food Out", "Expense", 3200),
    ("2026-01-21", "Grocery — weekly top-up", "Groceries", "Expense", 7100),
    ("2026-01-22", "Ride fares", "Transport", "Expense", 2900),
    ("2026-01-23", "Streaming subscription", "Entertainment", "Expense", 1200),
    ("2026-01-24", "Grocery — month-end top-up", "Groceries", "Expense", 5400),
    ("2026-01-25", "Water bill", "Utilities", "Expense", 1800),
    ("2026-01-26", "Office lunch", "Food Out", "Expense", 1600),
    ("2026-01-27", "Petrol", "Transport", "Expense", 7000),
    ("2026-01-28", "Grocery — month-end top-up", "Groceries", "Expense", 5100),
    ("2026-01-29", "Snacks / chai", "Food Out", "Expense", 1400),
    ("2026-01-30", "Phone accessories", "Entertainment", "Expense", 3500),
    ("2026-01-31", "Extra savings top-up", "Savings", "Expense", 5000),
]

# (category, monthly budget PKR) — expense categories only
BUDGETS = [
    ("Rent", 35000),
    ("Groceries", 55000),
    ("Transport", 30000),
    ("Utilities", 20000),
    ("Food Out", 15000),
    ("Entertainment", 15000),
    ("Savings", 25000),
]


def header_row(ws, row, values):
    for col, v in enumerate(values, start=1):
        c = ws.cell(row=row, column=col, value=v)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center")


def build():
    wb = Workbook()

    # --- Transactions ---
    ws = wb.active
    ws.title = "Transactions"
    header_row(ws, 1, ["Date", "Description", "Category", "Type", "Amount (PKR)"])
    for r, (d, desc, cat, typ, amt) in enumerate(TRANSACTIONS, start=2):
        ws.cell(row=r, column=1, value=d)
        ws.cell(row=r, column=2, value=desc)
        ws.cell(row=r, column=3, value=cat)
        ws.cell(row=r, column=4, value=typ)
        c = ws.cell(row=r, column=5, value=amt)
        c.number_format = MONEY
    last = 1 + len(TRANSACTIONS)
    for col, w in zip("ABCDE", [12, 30, 15, 10, 15]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:E{last}"
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    # --- Summary (all formulas) ---
    sm = wb.create_sheet("Summary")
    header_row(sm, 1, ["Measure", "Amount (PKR)"])
    rows = [
        ("Total income", f"=SUMIF(Transactions!D:D,\"Income\",Transactions!E:E)"),
        ("Total expenses", f"=SUMIF(Transactions!D:D,\"Expense\",Transactions!E:E)"),
        ("Balance (income - expenses)", "=B2-B3"),
        ("Transactions counted", "=COUNTA(Transactions!A:A)-1"),
        ("Savings rate (savings / income)",
         "=SUMIF(Transactions!C:C,\"Savings\",Transactions!E:E)/B2"),
    ]
    for r, (label, formula) in enumerate(rows, start=2):
        sm.cell(row=r, column=1, value=label)
        c = sm.cell(row=r, column=2, value=formula)
        c.number_format = '0.0%' if "rate" in label else MONEY
    sm.column_dimensions["A"].width = 38
    sm.column_dimensions["B"].width = 18

    # --- Budget (all formulas) ---
    bg = wb.create_sheet("Budget")
    header_row(bg, 1, ["Category", "Monthly budget (PKR)",
                        "Actual spent (PKR)", "Remaining (PKR)", "Status"])
    for r, (cat, budget) in enumerate(BUDGETS, start=2):
        bg.cell(row=r, column=1, value=cat)
        c = bg.cell(row=r, column=2, value=budget)
        c.number_format = MONEY
        c = bg.cell(row=r, column=3,
                    value=f"=SUMIF(Transactions!C:C,A{r},Transactions!E:E)")
        c.number_format = MONEY
        c = bg.cell(row=r, column=4, value=f"=B{r}-C{r}")
        c.number_format = MONEY
        bg.cell(row=r, column=5, value=f"=IF(C{r}>B{r},\"OVER\",\"OK\")")
    tr = 2 + len(BUDGETS)
    bg.cell(row=tr, column=1, value="TOTAL").font = BOLD
    for col in (2, 3, 4):
        c = bg.cell(row=tr, column=col,
                    value=f"=SUM({get_column_letter(col)}2:{get_column_letter(col)}{tr-1})")
        c.font = BOLD
        c.number_format = MONEY
    for col, w in zip("ABCDE", [16, 22, 20, 18, 10]):
        bg.column_dimensions[col].width = w

    # --- How to use ---
    how = wb.create_sheet("How to use")
    how["A1"] = ("SAMPLE tracker — fictional household, Jan 2026. "
                 "Add rows to Transactions and Summary + Budget update "
                 "themselves (every total is a formula). "
                 "Start from this file for a real month: clear the "
                 "Transactions rows, keep row 1, and type your own.")
    how["A1"].alignment = Alignment(wrap_text=True, vertical="top")
    how.column_dimensions["A"].width = 110
    how.row_dimensions[1].height = 60

    OUT.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUT)
    income = sum(a for *_, t, a in TRANSACTIONS if t == "Income")
    expenses = sum(a for *_, t, a in TRANSACTIONS if t == "Expense")
    print(f"saved {OUT}")
    print(f"transactions={len(TRANSACTIONS)} income={income} "
          f"expenses={expenses} balance={income - expenses}")


if __name__ == "__main__":
    build()
