"""Unit tests for the three Paycor audit checks:
  1. Review - Casual with coverage
  2. Review - Part-time with medical
  3. Review - Before eligibility date
Run:  python3 audit_checks_test.py
"""
import os, sys
from datetime import date
import pandas as pd

APP = os.path.join(os.path.dirname(os.path.abspath(__file__)), "app.py")
src = open(APP).read().split("# STREAMLIT APP")[0]
ns = {"__file__": APP}
exec(compile(src, APP, "exec"), ns)

run_audit_checks = ns["run_audit_checks"]
compute_eligibility_date = ns["compute_eligibility_date"]
DISPLAY_COLS = ns["DISPLAY_COLS"]

BLANK = {c: None for c in ns["PAYCOR_BENEFIT_COLS"]}

def emp(name, num, status="Full Time", worker="Regular", hire="01/02/2020",
        client="Testville", **amts):
    row = {"Client ID": "1", "Client Name": client, "Full Name": name,
           "Employee Number": num, "Status Type": status, "Worker Type": worker,
           "Hire Date": hire, **BLANK}
    row.update(amts)
    return row

def run(rows, month=date(2026, 7, 1)):
    return run_audit_checks(pd.DataFrame(rows).astype(str).replace("None", ""), month)

FAILS = []
def check(label, cond, detail=""):
    print(("PASS  " if cond else "FAIL  ") + label + (f"  [{detail}]" if detail and not cond else ""))
    if not cond: FAILS.append(label)

# ── constant ──────────────────────────────────────────────────────
check("ELIGIBILITY_WAITING_DAYS constant is 60", ns.get("ELIGIBILITY_WAITING_DAYS") == 60)

# ── eligibility date math ─────────────────────────────────────────
check("hire 4/15 -> +60d = 6/14 -> eligible 7/1",
      compute_eligibility_date(date(2026, 4, 15)) == date(2026, 7, 1))
check("hire 4/2 -> +60d = 6/1 -> eligible 6/1 (lands on a 1st)",
      compute_eligibility_date(date(2026, 4, 2)) == date(2026, 6, 1))
check("hire 12/15/2025 -> +60d = 2/13/2026 -> eligible 3/1/2026 (year wrap)",
      compute_eligibility_date(date(2025, 12, 15)) == date(2026, 3, 1))

# ── check 1: casual with coverage ─────────────────────────────────
res = run([emp("Smith, Casey", "10", worker="Casual", **{"Dent Amount": "12.34", "Vis Amount": "$5.00"})])
c1 = res[res["Action"] == "Review - Casual with coverage"]
check("casual with deductions flagged", len(c1) == 1, res.to_string())
check("casual note lists deductions", len(c1) == 1 and "Dent" in c1.iloc[0]["Note"] and "Vis" in c1.iloc[0]["Note"])

res = run([emp("Smith, Casey", "10", worker="Casual")])
check("casual with no deductions not flagged", len(res[res["Action"] == "Review - Casual with coverage"]) == 0)

res = run([emp("Doe, Jan", "11", worker="Regular", **{"Dent Amount": "12.34"})])
check("regular worker never casual-flagged", len(res[res["Action"] == "Review - Casual with coverage"]) == 0)

res = run([emp("Zero, Zed", "12", worker="Casual", **{"Dent Amount": "0.00"})])
check("casual with $0 deduction not flagged", len(res[res["Action"] == "Review - Casual with coverage"]) == 0)

# ── check 2: part-time with medical ───────────────────────────────
res = run([emp("Park, Tia", "20", status="Part Time", **{"Medical Amount": "50.00"})])
c2 = res[res["Action"] == "Review - Part-time with medical"]
check("part-time with Medical flagged", len(c2) == 1, res.to_string())

res = run([emp("Park, Tia", "20", status="Part Time", **{"Dent Amount": "9.00", "MedicalER2 Amount": "400.00"})])
check("part-time with only other benefits / MedicalER2 not flagged",
      len(res[res["Action"] == "Review - Part-time with medical"]) == 0)

res = run([emp("Full, Tim", "21", status="Full Time", **{"Medical Amount": "50.00"})])
check("full-time with Medical not flagged", len(res[res["Action"] == "Review - Part-time with medical"]) == 0)

# ── check 3: before eligibility ───────────────────────────────────
# hire 5/20/2026 -> +60d = 7/19 -> eligible 8/1; July recon is before that
res = run([emp("New, Nina", "30", hire="05/20/2026", **{"Medical Amount": "50.00"})], month=date(2026, 7, 1))
c3 = res[res["Action"] == "Review - Before eligibility date"]
check("hired 5/20, deduction in July recon flagged", len(c3) == 1, res.to_string())
check("note has hire and eligibility dates",
      len(c3) == 1 and "05/20/2026" in c3.iloc[0]["Note"] and "08/01/2026" in c3.iloc[0]["Note"])

res = run([emp("New, Nina", "30", hire="05/20/2026", **{"Medical Amount": "50.00"})], month=date(2026, 8, 1))
check("same hire, August recon not flagged", len(res[res["Action"] == "Review - Before eligibility date"]) == 0)

res = run([emp("New, Nina", "31", hire="05/20/2026")], month=date(2026, 7, 1))
check("not eligible but no deductions -> not flagged", len(res[res["Action"] == "Review - Before eligibility date"]) == 0)

res = run([emp("Old, Omar", "32", hire="", **{"Medical Amount": "50.00"})], month=date(2026, 7, 1))
check("blank hire date skipped without crash", len(res[res["Action"] == "Review - Before eligibility date"]) == 0)

# ── row shape / multiple checks per employee ──────────────────────
res = run([emp("Multi, Max", "40", status="Part Time", worker="Casual", hire="06/15/2026",
               **{"Medical Amount": "50.00"})], month=date(2026, 7, 1))
check("one row per employee per check (3 rows)", len(res) == 3, res.to_string())
check("columns match DISPLAY_COLS", list(res.columns) == DISPLAY_COLS)
check("all rows Broker /pay 0 (no at-stake impact)", (res["Broker /pay"] == 0).all())
check("all rows Difference /pay 0", (res["Difference /pay"] == 0).all())

# split-check rows for the same employee are combined, one flag not two
res = run([emp("Split, Sam", "50", worker="Casual", **{"Dent Amount": "6.00"}),
           emp("Split, Sam", "50", worker="Casual", **{"Dent Amount": "6.00"})])
check("split checks collapse to one flag per employee",
      len(res[res["Action"] == "Review - Casual with coverage"]) == 1)

print()
sys.exit("FAIL: " + ", ".join(FAILS) if FAILS else print("ALL PASS") or 0)
