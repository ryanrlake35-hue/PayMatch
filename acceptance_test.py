"""Acceptance test for the June 2026 CASA reconciliation (spec item 7).
Run:  python3 acceptance_test.py
Expects both June input files in client_data/.
"""
import os, sys
import pandas as pd

APP = os.path.join(os.path.dirname(os.path.abspath(__file__)), "app.py")
BROKER = os.path.join(os.path.dirname(APP), "client_data", "CASA_DEDUCTION_REPORT_June_2026.csv")
PAYCOR = os.path.join(os.path.dirname(APP), "client_data", "Casa_deduction_report_June_Paycor-original_.xlsx")

for f in (BROKER, PAYCOR):
    if not os.path.exists(f):
        sys.exit(f"MISSING INPUT FILE: {f}")

src = open(APP).read().split("# STREAMLIT APP")[0]
ns = {"__file__": APP}
exec(compile(src, APP, "exec"), ns)

class F:
    def __init__(self, path):
        self.name = os.path.basename(path)
        self._f = open(path, "rb")
    def __getattr__(self, a): return getattr(self._f, a)

broker_df, cw = ns["parse_broker_and_crosswalk"](F(BROKER))
paycor_df     = ns["parse_paycor"](F(PAYCOR))
res           = ns["reconcile"](broker_df, paycor_df, cw, 0.01)

ok     = res[res["Action"] == "OK"]
add    = res[res["Action"].str.startswith("Add")]
change = res[res["Action"].str.startswith("Change")]
review = res[res["Action"].str.startswith("Review")]
ok_zero = ok[(ok["Broker /pay"] == 0)]
unmatched = review[review["Action"] == "Review - Not in Paycor"]
paycor_only = review[review["Action"] == "Review - Not in broker file"]
ded_vs_zero = review[review["Action"] == "Review - Deduction vs $0 enrollment"]
at_stake = round(add["Broker /pay"].sum() + change["Difference /pay"].abs().sum()
                 + review["Broker /pay"].sum(), 2)

print(f"OK:      {len(ok)}  (amount matches: {len(ok)-len(ok_zero)}, zero-cost: {len(ok_zero)})   expected 1152 (1078 + 74)")
print(f"Add:     {len(add)}   expected 1")
print(f"Change:  {len(change)}   expected 4")
print(f"Review:  {len(review)}  (unmatched: {len(unmatched)} lines / {unmatched.groupby(['First Name','Last Name']).ngroups} people, "
      f"ded-vs-$0: {len(ded_vs_zero)}, paycor-only: {len(paycor_only)})   expected 144 (122/29, 1, 21)")
print(f"Per-pay $ at stake: ${at_stake:,.2f}   expected ~$4,232")
print("\n--- Change lines (expected Arnold CI 19.80/17.58, Damron Med 138.46/112.38, Novak Med 166.15/142.96, Vaughan Med 129.23/128.42) ---")
print(change[["Last Name","Benefit","Broker /pay","Paycor /pay"]].to_string(index=False))
print("\n--- Unmatched people + earliest Coverage Start Date (all expected 6/1/2026+) ---")
um_names = unmatched[["First Name","Last Name"]].drop_duplicates()
bb = broker_df.merge(um_names, on=["First Name","Last Name"])
bb["_csd"] = pd.to_datetime(bb["Coverage Start Date"], errors="coerce")
print(bb.groupby(["First Name","Last Name"])["_csd"].min().to_string())

passed = (len(ok) == 1152 and len(ok)-len(ok_zero) == 1078 and len(add) == 1 and len(change) == 4
          and len(review) == 144 and len(unmatched) == 122 and len(ded_vs_zero) == 1 and len(paycor_only) == 21)
print("\nACCEPTANCE:", "PASS" if passed else "FAIL — do not adjust logic; report differing lines for review")
