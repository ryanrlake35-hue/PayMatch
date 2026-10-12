"""Unit tests for the deduction detail ("long") report layout.
Run:  python3 layout_test.py
Uses synthetic rows only; no client data.
"""
import os
import pandas as pd

APP = os.path.join(os.path.dirname(os.path.abspath(__file__)), "app.py")
src = open(APP).read().split("# STREAMLIT APP")[0]
ns = {"__file__": APP}
exec(compile(src, APP, "exec"), ns)
norm = ns["normalize_deduction_report"]

FAILS = []
def check(label, cond):
    print(("PASS  " if cond else "FAIL  ") + label)
    if not cond: FAILS.append(label)

def row(client, ssn, first, last, code, amt, date, status="Full Time", etype="Regular"):
    return {"Client Name": client, "SSN": ssn, "First Name": first, "Last Name": last,
            "Status": "Active", "Status Type": status, "Hire Date": "2020-01-02 00:00:00",
            "Dist. Deduction Code": code, "Dist. Deduction Description": code,
            "Deduction Current Amount": amt, "Employee Type": etype,
            "Check Date": date + " 00:00:00"}

rows = []
# Client A: two regular payrolls; only the later one should be used
for d in ("2026-09-04", "2026-09-18"):
    rows += [row("A LLC", "111-11-1111", "Jane", "Doe", "Medical", "138.46", d),
             row("A LLC", "111-11-1111", "Jane", "Doe", "Dent", "13.10", d),
             row("A LLC", "111-11-1111", "Jane", "Doe", "Union", "20.00", d),
             row("A LLC", "222-22-2222", "Sam", "Roe", "", "", d, etype="Casual")]
# Split check on the latest run: 2.96 + 10.14 = 13.10
rows += [row("A LLC", "333-33-3333", "Ann", "Lee", "Dent", "2.96", "2026-09-18"),
         row("A LLC", "333-33-3333", "Ann", "Lee", "Dent", "10.14", "2026-09-18")]
# Off-cycle check after the regular run: must be ignored
rows += [row("A LLC", "111-11-1111", "Jane", "Doe", "Medical", "999.00", "2026-09-19")]
# Client B: single payroll, different date
rows += [row("B LLC", "444-44-4444", "Bo", "Ray", "Vis", "3.52", "2026-09-11", status="Part Time")]

out = norm(pd.DataFrame(rows).astype(str))
by = {(r["Client Name"], r["Full Name"]): r for _, r in out.iterrows()}

check("one row per employee with benefit deductions", len(out) == 3)
check("employee with no deductions excluded", ("A LLC", "Roe, Sam") not in by)
check("latest regular check used, not summed across month", by[("A LLC", "Doe, Jane")]["Medical Amount"] == "138.46")
check("off-cycle check ignored", ns["_money"](by[("A LLC", "Doe, Jane")]["Medical Amount"]) != 999.0)
check("split check summed", by[("A LLC", "Lee, Ann")]["Dent Amount"] == "13.10")
check("non-benefit codes dropped", not any("Union" in c for c in out.columns))
check("check date chosen per client", out.attrs["check_dates"] == {"A LLC": "09/18/2026", "B LLC": "09/11/2026"})
check("Employee Type mapped to Worker Type", by[("B LLC", "Ray, Bo")]["Worker Type"] == "Regular")
check("Status Type carried", by[("B LLC", "Ray, Bo")]["Status Type"] == "Part Time")
check("SSN removed from output", "SSN" not in out.columns
      and not out.astype(str).apply(lambda c: c.str.contains("111-11-1111")).any().any())

try:
    norm(pd.DataFrame({"Client Name": ["A"], "Name": ["x"]}))
    check("unknown layout raises a clear error", False)
except ValueError as e:
    check("unknown layout raises a clear error", "layout not recognized" in str(e).lower())

print("\nALL PASS" if not FAILS else f"\n{len(FAILS)} FAILED")
