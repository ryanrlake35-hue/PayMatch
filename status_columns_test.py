"""Tests for the split status columns (Broker Status / Paycor Status /
Paycor Worker Type) and the Review - Status conflict flag.
Run:  python3 status_columns_test.py
"""
import os, sys
import pandas as pd

APP = os.path.join(os.path.dirname(os.path.abspath(__file__)), "app.py")
src = open(APP).read().split("# STREAMLIT APP")[0]
ns = {"__file__": APP}
exec(compile(src, APP, "exec"), ns)

DISPLAY_COLS = ns["DISPLAY_COLS"]
reconcile = ns["reconcile"]
flag_status_conflicts = ns["flag_status_conflicts"]
BENEFIT_CROSSWALK = ns["BENEFIT_CROSSWALK"]

FAILS = []
def check(label, cond, detail=""):
    print(("PASS  " if cond else "FAIL  ") + label + (f"  [{detail}]" if detail and not cond else ""))
    if not cond: FAILS.append(label)

# ── column layout ─────────────────────────────────────────────────
check("DISPLAY_COLS has the three status columns, no old Status",
      "Broker Status" in DISPLAY_COLS and "Paycor Status" in DISPLAY_COLS
      and "Paycor Worker Type" in DISPLAY_COLS and "Status" not in DISPLAY_COLS)

# ── reconcile carries all three ───────────────────────────────────
def paycor_long(rows):
    return pd.DataFrame(rows, columns=["Client Name","Employee Number","Full Name",
                                       "Status Type","Worker Type","Paycor Column","Paycor Amount"])
def broker(rows):
    return pd.DataFrame(rows, columns=["First Name","Last Name","Division",
                                       "Employment Status","Benefit","EE Cost"])

p = paycor_long([("Testville","1","Smith, Casey","Full Time","Regular","Medical Amount",50.0)])
b = broker([("Casey","Smith","Testville","Active","Medical",50.0)])
res = reconcile(b, p, dict(BENEFIT_CROSSWALK), 0.01)
r = res.iloc[0]
check("matched row: Broker Status from broker file", r["Broker Status"] == "Active", res.to_string())
check("matched row: Paycor Status from Status Type", r["Paycor Status"] == "Full Time")
check("matched row: Paycor Worker Type carried", r["Paycor Worker Type"] == "Regular")
check("reconcile output columns == DISPLAY_COLS", list(res.columns) == DISPLAY_COLS)

# unmatched broker row -> Paycor fields blank
b2 = broker([("Nobody","Here","Testville","Active","Medical",50.0)])
res = reconcile(b2, p, dict(BENEFIT_CROSSWALK), 0.01)
nm = res[res["Action"] == "Review - Not in Paycor"].iloc[0]
check("unmatched broker row: Paycor Status/Worker Type blank",
      nm["Paycor Status"] == "" and nm["Paycor Worker Type"] == "" and nm["Broker Status"] == "Active")

# paycor-only row -> Broker Status blank
res = reconcile(broker([]), p, dict(BENEFIT_CROSSWALK), 0.01)
po = res[res["Action"] == "Review - Not in broker file"].iloc[0]
check("paycor-only row: Broker Status blank, Paycor fields filled",
      po["Broker Status"] == "" and po["Paycor Status"] == "Full Time" and po["Paycor Worker Type"] == "Regular")

# ── status conflict flag ──────────────────────────────────────────
def row(bstat, pstat, pval, action="OK", note=""):
    return {"Location":"T","First Name":"A","Last Name":"B","Broker Status":bstat,
            "Paycor Status":pstat,"Paycor Worker Type":"Regular","Benefit":"Medical",
            "Broker /pay":50.0,"Paycor /pay":pval,"Difference /pay":0.0,"Action":action,"Note":note}

res = flag_status_conflicts(pd.DataFrame([row("Terminated","Full Time",50.0)]))
check("Terminated + deductions -> Status conflict", res.iloc[0]["Action"] == "Review - Status conflict")
check("conflict note states both values",
      "Terminated" in res.iloc[0]["Note"] and "Full Time" in res.iloc[0]["Note"])

res = flag_status_conflicts(pd.DataFrame([row("COBRA","Full Time",50.0)]))
check("COBRA + deductions -> flagged", res.iloc[0]["Action"] == "Review - Status conflict")

res = flag_status_conflicts(pd.DataFrame([row("Active","Full Time",50.0)]))
check("Active never flagged", res.iloc[0]["Action"] == "OK")

res = flag_status_conflicts(pd.DataFrame([row("active ","Full Time",50.0)]))
check("case/space-insensitive Active not flagged", res.iloc[0]["Action"] == "OK")

res = flag_status_conflicts(pd.DataFrame([row("","Full Time",50.0)]))
check("blank broker status not flagged", res.iloc[0]["Action"] == "OK")

res = flag_status_conflicts(pd.DataFrame([row("Terminated","Full Time",0.0)]))
check("Terminated but zero Paycor deduction not flagged", res.iloc[0]["Action"] == "OK")

res = flag_status_conflicts(pd.DataFrame([row("Terminated","Full Time",50.0,
      action="Change - Amount differs", note="Broker $50.00 vs Paycor $45.00 per pay")]))
check("existing note preserved when re-flagged", "Broker $50.00 vs Paycor $45.00" in res.iloc[0]["Note"]
      and res.iloc[0]["Action"] == "Review - Status conflict")

check("empty frame handled", len(flag_status_conflicts(pd.DataFrame(columns=DISPLAY_COLS))) == 0)

# ── audit rows use the new columns ────────────────────────────────
from datetime import date
raw = pd.DataFrame([{"Client ID":"1","Client Name":"Testville","Full Name":"Smith, Casey",
    "Employee Number":"1","Status Type":"Part Time","Worker Type":"Casual",
    "Hire Date":"01/02/2020","Medical Amount":"50.00"}]).astype(str)
audit = ns["run_audit_checks"](raw, date(2026, 7, 1))
check("audit rows: Paycor Status + Worker Type filled, Broker Status blank",
      len(audit) > 0 and (audit["Paycor Status"] == "Part Time").all()
      and (audit["Paycor Worker Type"] == "Casual").all() and (audit["Broker Status"] == "").all(),
      audit.to_string() if len(audit) else "no rows")
check("audit columns == DISPLAY_COLS", list(audit.columns) == DISPLAY_COLS)

print()
sys.exit("FAIL: " + ", ".join(FAILS) if FAILS else print("ALL PASS") or 0)
