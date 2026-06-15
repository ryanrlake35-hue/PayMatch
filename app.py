"""
PayMatch by ProPayHR — Benefits Reconciliation Platform
"""
import io, warnings, os, re, hashlib, secrets, threading, time
from datetime import datetime, date, timedelta, timezone
from dotenv import load_dotenv
from supabase import create_client, Client as SupabaseClient
from pydantic import BaseModel, field_validator
from typing import Optional
import resend

load_dotenv()

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.chart import BarChart, DoughnutChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import PageSetupProperties

warnings.filterwarnings("ignore")

# ── EXCEL CONSTANTS ───────────────────────────────────────────────
NAVY = "1F3864"; WHITE = "FFFFFF"; SEC = "D9E1F2"; TF = "1F3864"
THIN = Side(style="thin", color="BFBFBF")
BORD = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
AR   = "Arial"
AF   = {"OK":"E2EFDA","Add":"DDEBF7","Change":"FCE4D6","Review":"FFF2CC"}

DISPLAY_COLS = ["Location","First Name","Last Name","Status","Benefit",
                "Broker /pay","Paycor /pay","Difference /pay","Action","Note"]
_APP_DIR    = os.path.dirname(os.path.abspath(__file__))
REPORTS_DIR = os.path.join(_APP_DIR, "reports")
STATUS_OPTIONS = ["Incomplete", "In Progress", "Complete"]

os.makedirs(REPORTS_DIR, exist_ok=True)  # ensure reports folder always exists

MONTHS = ["January","February","March","April","May","June",
          "July","August","September","October","November","December"]

CODE_MAP = {
    "medical":"Medical Amount","dental":"Dent Amount","vis":"Vis Amount",
    "longtermd":"LongTermD Amount","shorttermd":"ShortTermD Amount","acc":"Acc Amount",
    "id protect":None,"legal":None,"hosp ind":"HospIND Amount",
    "critil":"CritIll Amount","critill":"CritIll Amount",
}

BUILTIN_CW = {
    "Medical Plan Employee Per Pay Cost":                    "Medical Amount",
    "Dental Plan Employee Per Pay Cost":                     "Dent Amount",
    "Vision Plan Employee Per Pay Cost":                     "Vis Amount",
    "Voluntary Long-Term Disability Employee Per Pay Cost":  "LongTermD Amount",
    "Voluntary Short-Term Disability Employee Per Pay Cost": "ShortTermD Amount",
    "Accident Plan Employee Per Pay Cost":                   "Acc Amount",
    "Identity Protection Plan Employee Per Pay Cost":        None,
    "Legal Plan Employee Per Pay Cost":                      None,
    "Hospital Indemnity Plan Employee Per Pay Cost":         "HospIND Amount",
    "Critical Illness":                                      "CritIll Amount",
}

# ── SUPABASE ──────────────────────────────────────────────────────
SUPABASE_URL = os.getenv("SUPABASE_URL", "")
SUPABASE_KEY = os.getenv("SUPABASE_KEY", "")

@st.cache_resource
def _sb() -> SupabaseClient:
    return create_client(SUPABASE_URL, SUPABASE_KEY)

def _parse_dt(s):
    if not s: return None
    try:
        dt = datetime.fromisoformat(str(s).replace("Z", "+00:00"))
        return dt if dt.tzinfo else dt.replace(tzinfo=timezone.utc)
    except Exception:
        return None

# ── PYDANTIC MODELS ───────────────────────────────────────────────
class UserCreate(BaseModel):
    name: str
    email: str
    password_hash: str
    salt: str
    verified: bool = False
    verification_code: Optional[str] = None
    reset_code: Optional[str] = None
    failed_attempts: int = 0
    locked_until: Optional[str] = None

    @field_validator("name")
    @classmethod
    def name_not_empty(cls, v: str) -> str:
        v = v.strip()
        if not v:
            raise ValueError("Name cannot be empty")
        return v

    @field_validator("email")
    @classmethod
    def email_valid(cls, v: str) -> str:
        v = v.strip().lower()
        if "@" not in v or "." not in v.split("@")[-1]:
            raise ValueError("Invalid email address")
        return v

    @field_validator("failed_attempts")
    @classmethod
    def attempts_non_negative(cls, v: int) -> int:
        if v < 0:
            raise ValueError("failed_attempts must be >= 0")
        return v


class ReconciliationRecord(BaseModel):
    client: str
    period: str
    run_by: str
    run_date: str
    run_time: str
    total_lines: int
    ok_count: int
    add_count: int
    change_count: int
    review_count: int
    discrepancies: int
    monthly_at_stake: float
    status: str = "Incomplete"
    report_file: str = ""
    notes: str = ""

    @field_validator("client", "period")
    @classmethod
    def not_empty(cls, v: str) -> str:
        if not v.strip():
            raise ValueError("Field cannot be empty")
        return v.strip()

    @field_validator("total_lines", "ok_count", "add_count", "change_count", "review_count", "discrepancies")
    @classmethod
    def non_negative(cls, v: int) -> int:
        if v < 0:
            raise ValueError("Count must be >= 0")
        return v

    @field_validator("monthly_at_stake")
    @classmethod
    def valid_amount(cls, v: float) -> float:
        if v < 0:
            raise ValueError("monthly_at_stake must be >= 0")
        return round(v, 2)

# ── FILE PARSERS ──────────────────────────────────────────────────
def read_file(file):
    n = file.name.lower()
    if n.endswith(".csv"):   return pd.read_csv(file, dtype=str)
    elif n.endswith(".xls"): return pd.read_excel(file, dtype=str, engine="xlrd")
    else:                    return pd.read_excel(file, dtype=str)

def parse_paycor(file):
    df = read_file(file)
    df.columns = [str(c).strip() for c in df.columns]
    for col in df.columns:
        if "Amount" in col or "amount" in col:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    return df

def parse_broker_and_crosswalk(file):
    raw = read_file(file)
    raw.columns = [str(c).strip() for c in raw.columns]
    mapping_row = raw.iloc[0]
    cw = dict(BUILTIN_CW)
    for col in raw.columns:
        val = str(mapping_row.get(col, "")).strip()
        if not val or val.lower() == "nan": continue
        code = val.split(" - ")[0].strip().lower()
        paycor_col = CODE_MAP.get(code)
        for candidate in [
            f"{col} Employee Per Pay Cost",
            f"{col.replace(' Plan','')} Employee Per Pay Cost",
            col,
        ]:
            if candidate in raw.columns and candidate not in cw:
                cw[candidate] = paycor_col; break
    return raw.iloc[1:].reset_index(drop=True), cw

# ── RECONCILIATION ENGINE ─────────────────────────────────────────
def reconcile(broker_df, paycor_df, crosswalk, tolerance):
    rows = []
    for _, emp in broker_df.iterrows():
        fn = str(emp.get("First Name","")).strip(); ln = str(emp.get("Last Name","")).strip()
        div = str(emp.get("Division","")).strip(); status = str(emp.get("Employment Status","")).strip()
        if not fn and not ln: continue
        pmatch = paycor_df[
            (paycor_df.get("First Name", pd.Series(dtype=str)).str.strip()==fn) &
            (paycor_df.get("Last Name",  pd.Series(dtype=str)).str.strip()==ln)]
        for broker_col, paycor_col in crosswalk.items():
            bval = pd.to_numeric(emp.get(broker_col, 0), errors="coerce")
            if pd.isna(bval) or bval == 0: continue
            benefit = broker_col.replace(" Employee Per Pay Cost","").replace(" Plan","").strip()
            pval=0.0; action=""; note=""
            if paycor_col is None:
                action="Review - Not in Paycor export"; note="No Paycor column mapped — verify manually"
            elif len(pmatch)==0:
                action="Review - Not in Paycor"; note="Employee not found in Paycor — possible new hire or name mismatch"
            else:
                pval=float(pmatch.iloc[0].get(paycor_col,0) or 0)
                diff=round(bval-pval,2)
                if abs(diff)<=tolerance: action="OK"; note=""
                elif pval==0: action="Add - Start deduction"; note="Enrolled with broker but zero deduction in Paycor"
                elif pval<0: action="Change - Negative in Paycor"; note=f"Paycor shows negative ${abs(pval):.2f} — possible credit/error"
                else: action="Change - Amount differs"; note=f"Broker ${bval:.2f} vs Paycor ${pval:.2f} per pay"
            rows.append({"Location":div,"First Name":fn,"Last Name":ln,"Status":status,
                "Benefit":benefit,"Broker /pay":round(float(bval),2),"Paycor /pay":round(pval,2),
                "Difference /pay":round(float(bval)-pval,2),"Action":action,"Note":note})
    return pd.DataFrame(rows).reset_index(drop=True)

# ── EXCEL BUILDER ─────────────────────────────────────────────────
def _sw(ws,wmap):
    for ci,w in wmap.items(): ws.column_dimensions[get_column_letter(ci)].width=w

def _nh(ws,txt,n):
    ws.merge_cells(f"A1:{get_column_letter(n)}1")
    c=ws["A1"]; c.value=txt; c.font=Font(name=AR,size=12,bold=True,color=WHITE)
    c.fill=PatternFill("solid",fgColor=NAVY); c.alignment=Alignment(horizontal="left",vertical="center",indent=1)
    ws.row_dimensions[1].height=22

def _kh(ws,cols,r=2):
    for h,cs,ce in cols:
        if cs!=ce: ws.merge_cells(start_row=r,start_column=cs,end_row=r,end_column=ce)
        c=ws.cell(r,cs,h); c.font=Font(name=AR,size=10,bold=True,color=WHITE)
        c.fill=PatternFill("solid",fgColor="404040"); c.alignment=Alignment(horizontal="center",vertical="center"); c.border=BORD
    ws.row_dimensions[r].height=18

def _kr(ws,ri,fc,label,count,meaning,action,me,as_,ae):
    ws.row_dimensions[ri].height=28
    c1=ws.cell(ri,1,label); c1.font=Font(name=AR,size=10,bold=True); c1.fill=PatternFill("solid",fgColor=fc); c1.border=BORD; c1.alignment=Alignment(vertical="center")
    c2=ws.cell(ri,2,count); c2.font=Font(name=AR,size=14,bold=True,color=NAVY); c2.fill=PatternFill("solid",fgColor=fc); c2.border=BORD; c2.alignment=Alignment(horizontal="center",vertical="center")
    ws.merge_cells(start_row=ri,start_column=3,end_row=ri,end_column=me)
    c3=ws.cell(ri,3,meaning); c3.font=Font(name=AR,size=10); c3.fill=PatternFill("solid",fgColor=fc); c3.border=BORD; c3.alignment=Alignment(vertical="center",wrap_text=True)
    ws.merge_cells(start_row=ri,start_column=as_,end_row=ri,end_column=ae)
    c4=ws.cell(ri,as_,f"Action: {action}"); c4.font=Font(name=AR,size=10,bold=True); c4.fill=PatternFill("solid",fgColor=fc); c4.border=BORD; c4.alignment=Alignment(vertical="center",wrap_text=True)

def _kt(ws,ri,lbl,cnt,summary,right,ss,se,as_,ae):
    ws.row_dimensions[ri].height=24
    for ci,v,_ in [(1,lbl,None),(2,cnt,None)]:
        c=ws.cell(ri,ci,v); c.font=Font(name=AR,size=10 if ci==1 else 14,bold=True,color=WHITE)
        c.fill=PatternFill("solid",fgColor=TF); c.border=BORD
        c.alignment=Alignment(indent=1 if ci==1 else 0,horizontal="center" if ci==2 else "left",vertical="center")
    ws.merge_cells(start_row=ri,start_column=ss,end_row=ri,end_column=se)
    c3=ws.cell(ri,ss,summary); c3.font=Font(name=AR,size=10,bold=True,color=WHITE); c3.fill=PatternFill("solid",fgColor=TF); c3.border=BORD; c3.alignment=Alignment(vertical="center",indent=1)
    ws.merge_cells(start_row=ri,start_column=as_,end_row=ri,end_column=ae)
    c4=ws.cell(ri,as_,right); c4.font=Font(name=AR,size=10,bold=True,color=WHITE); c4.fill=PatternFill("solid",fgColor=TF); c4.border=BORD; c4.alignment=Alignment(vertical="center",indent=1)

def _dh(ws,heads,row):
    for ci,h in enumerate(heads,1):
        c=ws.cell(row,ci,h); c.font=Font(name=AR,size=10,bold=True,color=WHITE)
        c.fill=PatternFill("solid",fgColor=NAVY); c.alignment=Alignment(horizontal="center",wrap_text=True); c.border=BORD
    ws.row_dimensions[row].height=28; ws.freeze_panes=f"A{row+1}"

def _dr(ws,src,start):
    for ri,row in src[DISPLAY_COLS].iterrows():
        rn=ri+start; fc=AF.get(str(row["Action"]).split(" - ")[0],"FFFFFF")
        for ci,v in enumerate(row.values,1):
            c=ws.cell(rn,ci,v); c.font=Font(name=AR,size=10); c.fill=PatternFill("solid",fgColor=fc); c.border=BORD
            if ci in (6,7,8): c.number_format="$#,##0.00"

def _nt(ws,r,c1,lbl,c2,val,fmt="0",span=None):
    if span: ws.merge_cells(start_row=r,start_column=c1,end_row=r,end_column=span)
    cx=ws.cell(r,c1,lbl); cx.font=Font(name=AR,size=10,bold=True,color=WHITE); cx.fill=PatternFill("solid",fgColor=TF); cx.border=BORD; cx.alignment=Alignment(indent=1)
    vx=ws.cell(r,c2,val); vx.font=Font(name=AR,size=10,bold=True,color=WHITE); vx.fill=PatternFill("solid",fgColor=TF); vx.border=BORD; vx.alignment=Alignment(horizontal="right"); vx.number_format=fmt

def build_action_items_only(df, client_name="Client"):
    disc=df[df["Action"]!="OK"].reset_index(drop=True)
    ct={"Add":int(df["Action"].str.startswith("Add").sum()),"Change":int(df["Action"].str.startswith("Change").sum()),"Review":int(df["Action"].str.startswith("Review").sum())}
    tdisc=ct["Add"]+ct["Change"]+ct["Review"]; n_emps=disc["First Name"].nunique()
    add_mo=round(df[df["Action"].str.startswith("Add")]["Broker /pay"].sum()*26/12,2)
    wb=Workbook(); ws=wb.active; ws.title="Action Items"; ws.sheet_view.showGridLines=False
    _sw(ws,{1:18,2:10,3:26,4:46,5:12,6:12,7:11,8:28,9:40})
    _nh(ws,f"PayMatch  |  {client_name}  |  Action Items  |  {datetime.now().strftime('%B %d, %Y')}",9)
    ws.merge_cells("A2:I2"); c=ws["A2"]
    c.value=f"Total: {tdisc} lines  |  {n_emps} employees  |  Add: {ct['Add']}  |  Change: {ct['Change']}  |  Review: {ct['Review']}  |  ${add_mo:,.2f}/month not being collected"
    c.font=Font(name=AR,size=10,bold=True,color=WHITE); c.fill=PatternFill("solid",fgColor="2d5299")
    c.alignment=Alignment(horizontal="left",vertical="center",indent=1); ws.row_dimensions[2].height=18
    _kh(ws,[("Color / Action",1,1),("Count",2,2),("What it means",3,5),("What to do",6,9)],r=3)
    aik=[("DDEBF7","Blue  -  Add",ct["Add"],"Enrolled with broker but zero deduction in Paycor.","Set up the deduction in Paycor immediately."),
         ("FCE4D6","Pink  -  Change",ct["Change"],"Both sides have the benefit but dollar amounts differ.","Update the deduction amount in Paycor to match the broker."),
         ("FFF2CC","Yellow  -  Review",ct["Review"],"Benefit not in Paycor export, or employee name did not match.","Verify manually in Paycor.")]
    for ri,(fc,lb,cnt,mn,ac) in enumerate(aik,4): _kr(ws,ri,fc,lb,cnt,mn,ac,5,6,9)
    _kt(ws,7,"TOTAL ACTION ITEMS",tdisc,f"{ct['Add']} Add  +  {ct['Change']} Change  +  {ct['Review']} Review  =  {tdisc} lines",f"{n_emps} employees need attention",3,5,6,9)
    ws.row_dimensions[8].height=8; _dh(ws,DISPLAY_COLS,9); ws.column_dimensions["J"].width=40; _dr(ws,disc,10)
    buf=io.BytesIO(); wb.save(buf); buf.seek(0); return buf.getvalue()

def build_excel(df, client_name="Client"):
    disc=df[df["Action"]!="OK"].reset_index(drop=True)
    ct={"OK":int((df["Action"]=="OK").sum()),"Add":int(df["Action"].str.startswith("Add").sum()),
        "Change":int(df["Action"].str.startswith("Change").sum()),"Review":int(df["Action"].str.startswith("Review").sum())}
    total=len(df); tdisc=ct["Add"]+ct["Change"]+ct["Review"]
    adds=df[df["Action"].str.startswith("Add")]; chgs=df[df["Action"].str.startswith("Change")]
    add_mo=round(adds["Broker /pay"].sum()*26/12,2); cdiff=round(chgs["Difference /pay"].abs().sum(),2)
    locs=df[df["Action"]!="OK"].groupby("Location").size().sort_values(ascending=False)
    n_emps=disc["First Name"].nunique(); wb=Workbook()
    d=wb.active; d.title="Dashboard"; d.sheet_view.showGridLines=False
    for col in "ABCDEFGH": d.column_dimensions[col].width=13
    d.column_dimensions["E"].width=18
    def m(r): d.merge_cells(r)
    def p(co,v=None,font=None,fill=None,align=None,fmt=None,border=False):
        c=d[co]
        if v is not None: c.value=v
        c.font=font or Font(name=AR,size=11)
        if fill: c.fill=PatternFill("solid",fgColor=fill)
        if align: c.alignment=align
        if fmt: c.number_format=fmt
        if border: c.border=BORD
        return c
    m("A1:H1"); p("A1",f"ProPayHR  |  {client_name} Benefits Reconciliation  |  {datetime.now().strftime('%B %d, %Y')}",font=Font(name=AR,size=14,bold=True,color=WHITE),fill=NAVY,align=Alignment(vertical="center",indent=1)); d.row_dimensions[1].height=28
    m("A2:H2"); p("A2","Broker enrollment (per pay) vs Paycor deductions (per pay)  |  Direct per-pay comparison",font=Font(name=AR,size=9,italic=True,color="595959"))
    tiles=[("A","Lines compared",total,"0","DCE6F1"),("C","Matched (OK)",ct["OK"],"0","E2EFDA"),("E","Require action",ct["Add"]+ct["Change"],"0","FCE4D6"),("G","Monthly $ at stake",add_mo,"$#,##0.00","FFF2CC")]
    d.row_dimensions[4].height=16; d.row_dimensions[5].height=30
    for col,lab,val,fmt,fill in tiles:
        c2=chr(ord(col)+1); m(f"{col}4:{c2}4"); p(f"{col}4",lab,font=Font(name=AR,size=10,color="404040"),fill=fill,align=Alignment(horizontal="center"))
        m(f"{col}5:{c2}5"); p(f"{col}5",val,font=Font(name=AR,size=16,bold=True,color=NAVY),fill=fill,align=Alignment(horizontal="center"),fmt=fmt)
    m("A7:D7"); p("A7","Reconciliation outcome",font=Font(name=AR,size=12,bold=True,color=NAVY),fill=SEC,align=Alignment(indent=1))
    outs=[("Matched (OK)",ct["OK"],"E2EFDA"),("Add - start deduction",ct["Add"],"DDEBF7"),("Change - amount differs",ct["Change"],"FCE4D6"),("Review - manual check",ct["Review"],"FFF2CC")]
    for i,(lab,cnt,fc) in enumerate(outs):
        r=8+i; m(f"A{r}:C{r}"); cx=d[f"A{r}"]; cx.value=lab; cx.font=Font(name=AR,size=10); cx.fill=PatternFill("solid",fgColor=fc); cx.border=BORD
        vx=d[f"D{r}"]; vx.value=cnt; vx.font=Font(name=AR,size=10,bold=True); vx.fill=PatternFill("solid",fgColor=fc); vx.alignment=Alignment(horizontal="right"); vx.border=BORD
    _nt(d,12,1,"TOTAL LINES",4,total,span=3)
    for i,(lab,cnt,_) in enumerate(outs): d.cell(70+i,1,lab.split(" - ")[0]); d.cell(70+i,2,cnt)
    dg=DoughnutChart(); dg.title="Outcome"; dg.height=6.5; dg.width=9.5; dg.style=10; dg.varyColors=True
    dg.add_data(Reference(d,min_col=2,min_row=70,max_row=73)); dg.set_categories(Reference(d,min_col=1,min_row=70,max_row=73))
    dl=DataLabelList(); dl.showVal=True; dl.showSerName=False; dl.showCatName=False; dl.showPercent=False; dg.dataLabels=dl; d.add_chart(dg,"E7")
    m("A14:D14"); p("A14","Discrepancies by location",font=Font(name=AR,size=12,bold=True,color=NAVY),fill=SEC,align=Alignment(indent=1))
    m("A15:C15"); p("A15","Location",font=Font(name=AR,size=10,bold=True),fill="F2F2F2",border=True); p("D15","Issues",font=Font(name=AR,size=10,bold=True),fill="F2F2F2",align=Alignment(horizontal="right"),border=True)
    for i,(loc,cnt) in enumerate(locs.items()):
        r=16+i; m(f"A{r}:C{r}"); d[f"A{r}"].value=loc; d[f"A{r}"].font=Font(name=AR,size=10); d[f"A{r}"].border=BORD
        d[f"D{r}"].value=cnt; d[f"D{r}"].font=Font(name=AR,size=10,bold=True); d[f"D{r}"].alignment=Alignment(horizontal="right"); d[f"D{r}"].border=BORD
    tr=16+len(locs); _nt(d,tr,1,"TOTAL DISCREPANCIES",4,tdisc,span=3)
    bg=BarChart(); bg.type="bar"; bg.title="Issues by location"; bg.height=8; bg.width=11; bg.style=11; bg.legend=None
    bg.add_data(Reference(d,min_col=4,min_row=16,max_row=16+len(locs)-1)); bg.set_categories(Reference(d,min_col=1,min_row=16,max_row=16+len(locs)-1))
    bl=DataLabelList(); bl.showVal=True; bl.showSerName=False; bg.dataLabels=bl
    try: bg.series[0].graphicalProperties.solidFill="C0504D"
    except: pass
    d.add_chart(bg,"E14")
    dr2=tr+2; m(f"A{dr2}:D{dr2}"); p(f"A{dr2}","Dollar impact",font=Font(name=AR,size=12,bold=True,color=NAVY),fill=SEC,align=Alignment(indent=1))
    m(f"A{dr2+1}:C{dr2+1}"); p(f"A{dr2+1}","Category",font=Font(name=AR,size=10,bold=True),fill="F2F2F2",border=True); p(f"D{dr2+1}","Per pay",font=Font(name=AR,size=10,bold=True),fill="F2F2F2",align=Alignment(horizontal="right"),border=True); p(f"E{dr2+1}","Per month",font=Font(name=AR,size=10,bold=True),fill="F2F2F2",align=Alignment(horizontal="right"),border=True)
    drows=[("Add - deductions to start",round(adds["Broker /pay"].sum(),2)),("Change - corrections needed",cdiff)]
    for i,(lab,amt) in enumerate(drows):
        r=dr2+2+i; m(f"A{r}:C{r}"); d[f"A{r}"].value=lab; d[f"A{r}"].font=Font(name=AR,size=10); d[f"A{r}"].border=BORD
        d[f"D{r}"].value=amt; d[f"D{r}"].font=Font(name=AR,size=10,bold=True); d[f"D{r}"].number_format="$#,##0.00"; d[f"D{r}"].alignment=Alignment(horizontal="right"); d[f"D{r}"].border=BORD
        d[f"E{r}"].value=round(amt*26/12,2); d[f"E{r}"].font=Font(name=AR,size=10,bold=True); d[f"E{r}"].number_format="$#,##0.00"; d[f"E{r}"].alignment=Alignment(horizontal="right"); d[f"E{r}"].border=BORD
    ta=round(sum(x[1] for x in drows),2); r=dr2+4; m(f"A{r}:C{r}"); _nt(d,r,1,"TOTAL",4,ta,fmt="$#,##0.00",span=3)
    d[f"E{r}"].value=round(ta*26/12,2); d[f"E{r}"].font=Font(name=AR,size=10,bold=True,color=WHITE); d[f"E{r}"].fill=PatternFill("solid",fgColor=TF); d[f"E{r}"].number_format="$#,##0.00"; d[f"E{r}"].alignment=Alignment(horizontal="right"); d[f"E{r}"].border=BORD
    d.page_setup.orientation="landscape"; d.page_setup.fitToWidth=1; d.page_setup.fitToHeight=0; d.sheet_properties.pageSetUpPr=PageSetupProperties(fitToPage=True)
    aw=wb.create_sheet("Action Items"); aw.sheet_view.showGridLines=False
    _sw(aw,{1:18,2:10,3:26,4:46,5:12,6:12,7:11,8:28,9:40})
    _nh(aw,"ACTION ITEMS  |  Color Key & Summary",9)
    _kh(aw,[("Color / Action",1,1),("Count",2,2),("What it means",3,5),("What to do",6,9)])
    aik=[("DDEBF7","Blue  -  Add",ct["Add"],"Enrolled with broker but zero deduction in Paycor. Company may be eating this cost.","Set up the deduction in Paycor immediately."),
         ("FCE4D6","Pink  -  Change",ct["Change"],"Both sides have the benefit but dollar amounts differ.","Update the deduction amount in Paycor to match the broker."),
         ("FFF2CC","Yellow  -  Review",ct["Review"],"Benefit not in Paycor export, or employee name did not match.","Verify manually in Paycor.")]
    for ri,(fc,lb,cnt,mn,ac) in enumerate(aik,3): _kr(aw,ri,fc,lb,cnt,mn,ac,5,6,9)
    _kt(aw,6,"TOTAL ACTION ITEMS",tdisc,f"{ct['Add']} Add  +  {ct['Change']} Change  +  {ct['Review']} Review  =  {tdisc} lines",f"{n_emps} employees need attention",3,5,6,9)
    aw.row_dimensions[7].height=8; _dh(aw,DISPLAY_COLS,8); _dr(aw,disc,9); aw.column_dimensions["J"].width=40
    allw=wb.create_sheet("All Lines"); allw.sheet_view.showGridLines=False
    _sw(allw,{1:18,2:13,3:13,4:11,5:28,6:12,7:12,8:11,9:26,10:40})
    _nh(allw,"ALL RECONCILIATION LINES  |  Color Key & Summary",10)
    _kh(allw,[("Color / Action",1,1),("Count",2,2),("What it means",3,6),("What to do",7,10)])
    allk=[("E2EFDA","Green  -  OK",ct["OK"],"Broker and Paycor match perfectly.","No action needed."),
          ("DDEBF7","Blue  -  Add",ct["Add"],"Enrolled with broker but zero deduction in Paycor.","Set up the deduction in Paycor immediately."),
          ("FCE4D6","Pink  -  Change",ct["Change"],"Both sides have the benefit but dollar amounts differ.","Update the deduction amount in Paycor to match the broker."),
          ("FFF2CC","Yellow  -  Review",ct["Review"],"Benefit not in export or name not matched.","Verify manually in Paycor.")]
    for ri,(fc,lb,cnt,mn,ac) in enumerate(allk,3): _kr(allw,ri,fc,lb,cnt,mn,ac,6,7,10)
    _kt(allw,7,"TOTAL LINES",total,f"{ct['OK']} OK  +  {ct['Add']} Add  +  {ct['Change']} Change  +  {ct['Review']} Review  =  {total}",f"{tdisc} lines need attention  |  {n_emps} employees affected",3,6,7,10)
    allw.row_dimensions[8].height=8; _dh(allw,DISPLAY_COLS,9); _dr(allw,df,10)
    buf=io.BytesIO(); wb.save(buf); buf.seek(0); return buf.getvalue()

# ── HISTORY (Supabase) ────────────────────────────────────────────
def save_to_history(recon_df, client_name, period, run_by="", excel_bytes=None):
    ct = {
        "OK":     int((recon_df["Action"] == "OK").sum()),
        "Add":    int(recon_df["Action"].str.startswith("Add").sum()),
        "Change": int(recon_df["Action"].str.startswith("Change").sum()),
        "Review": int(recon_df["Action"].str.startswith("Review").sum()),
    }
    mo = round(recon_df[recon_df["Action"].str.startswith("Add")]["Broker /pay"].sum() * 26 / 12, 2)
    report_file = ""
    if excel_bytes:
        safe = client_name.strip().replace(" ", "_").replace("/", "_")
        ts   = datetime.now().strftime("%Y%m%d_%H%M%S")
        report_file = os.path.join(REPORTS_DIR, f"PayMatch_{safe}_{period.replace(' ','_')}_{ts}.xlsx")
        with open(report_file, "wb") as f:
            f.write(excel_bytes)
    try:
        existing = _sb().table("reconciliation_history").select("id,notes").eq("client", client_name).eq("period", period).execute()
        record = ReconciliationRecord(
            client=client_name, period=period, run_by=run_by,
            run_date=datetime.now().strftime("%Y-%m-%d"),
            run_time=datetime.now().strftime("%I:%M %p"),
            total_lines=len(recon_df),
            ok_count=ct["OK"], add_count=ct["Add"],
            change_count=ct["Change"], review_count=ct["Review"],
            discrepancies=ct["Add"] + ct["Change"] + ct["Review"],
            monthly_at_stake=mo, report_file=report_file, notes="",
        )
        row = record.model_dump()
        if existing.data:
            row["notes"] = existing.data[0].get("notes", "") or ""
            _sb().table("reconciliation_history").update(row).eq("client", client_name).eq("period", period).execute()
        else:
            _sb().table("reconciliation_history").insert(row).execute()
    except Exception:
        pass

def load_history(client_name=None):
    try:
        q = _sb().table("reconciliation_history").select("*").order("run_date", desc=True)
        if client_name:
            q = q.eq("client", client_name)
        res = q.execute()
        if not res.data:
            return pd.DataFrame()
        return pd.DataFrame(res.data).rename(columns={
            "ok_count": "OK", "add_count": "Add", "change_count": "Change",
            "review_count": "Review", "total_lines": "Total Lines",
            "discrepancies": "Discrepancies",
            "monthly_at_stake": "Monthly $ at stake", "run_by": "Run By",
            "run_date": "Run Date", "run_time": "Run Time",
            "report_file": "Report File", "client": "Client",
            "period": "Period", "notes": "Notes", "status": "Status",
        }).reset_index(drop=True)
    except Exception:
        return pd.DataFrame()

def update_run_status(run_date, client, period, status):
    try:
        _sb().table("reconciliation_history").update({"status": status}).eq("run_date", str(run_date)).eq("client", client).eq("period", period).execute()
    except Exception:
        pass

def delete_run(run_date, client, period, report_file=""):
    try:
        _sb().table("reconciliation_history").delete().eq("run_date", str(run_date)).eq("client", client).eq("period", period).execute()
    except Exception:
        pass
    if report_file:
        rf = report_file if os.path.isabs(report_file) else os.path.join(_APP_DIR, report_file)
        if os.path.exists(rf):
            try: os.remove(rf)
            except Exception: pass

# ── USER MANAGEMENT (Supabase) ────────────────────────────────────
def load_users():
    try:
        res = _sb().table("profiles").select("*").execute()
        return res.data or []
    except Exception:
        return []

def get_user(email):
    try:
        res = _sb().table("profiles").select("*").eq("email", email.strip().lower()).execute()
        return res.data[0] if res.data else None
    except Exception:
        return None

def hash_pw(password, salt=None):
    if salt is None: salt = secrets.token_hex(32)
    return hashlib.sha256((salt + password).encode()).hexdigest(), salt

def create_user(name, email, password):
    pw_hash, salt = hash_pw(password)
    code = f"{secrets.randbelow(900000) + 100000}"
    user = UserCreate(
        name=name,
        email=email.strip().lower(),
        password_hash=pw_hash,
        salt=salt,
        verification_code=code,
    )
    res = _sb().table("profiles").insert(user.model_dump()).execute()
    return res.data[0] if res.data else user.model_dump()

def verify_user(email, code):
    u = get_user(email)
    if u and str(u.get("verification_code", "")) == str(code).strip():
        _sb().table("profiles").update({"verified": True, "verification_code": None}).eq("email", email.lower()).execute()
        return get_user(email)
    return None

def generate_reset_code(email):
    u = get_user(email)
    if not u: return None, None
    code = f"{secrets.randbelow(900000) + 100000}"
    _sb().table("profiles").update({"reset_code": code}).eq("email", email.lower()).execute()
    return get_user(email), code

def apply_reset(email, code, new_pw):
    u = get_user(email)
    if u and str(u.get("reset_code", "")) == str(code).strip():
        pw_hash, salt = hash_pw(new_pw)
        _sb().table("profiles").update({
            "password_hash": pw_hash, "salt": salt,
            "reset_code": None, "verified": True,
        }).eq("email", email.lower()).execute()
        return get_user(email)
    return None

def update_password(email, new_pw):
    pw_hash, salt = hash_pw(new_pw)
    _sb().table("profiles").update({"password_hash": pw_hash, "salt": salt}).eq("email", email.lower()).execute()
    return True

def pw_checks(pw):
    return {
        "length":  len(pw) >= 8,
        "upper":   bool(re.search(r"[A-Z]", pw)),
        "lower":   bool(re.search(r"[a-z]", pw)),
        "digit":   bool(re.search(r"\d", pw)),
        "special": bool(re.search(r'[!@#$%^&*()\-_=+\[\]{}|;:\'",.<>?/`~\\]', pw)),
    }

def pw_ok(pw): return all(pw_checks(pw).values())

def check_account_locked(email):
    u = get_user(email)
    if not u: return False
    locked = _parse_dt(u.get("locked_until"))
    if not locked: return False
    if datetime.now(timezone.utc) < locked:
        return locked
    _reset_lockout(email)
    return False

def _reset_lockout(email):
    try:
        _sb().table("profiles").update({"failed_attempts": 0, "locked_until": None}).eq("email", email.lower()).execute()
    except Exception:
        pass

def record_failed_attempt(email):
    u = get_user(email)
    if not u: return 1
    attempts = u.get("failed_attempts", 0) + 1
    update_data = {"failed_attempts": attempts}
    if attempts >= 5:
        update_data["locked_until"] = (datetime.now(timezone.utc) + timedelta(minutes=15)).isoformat()
    try:
        _sb().table("profiles").update(update_data).eq("email", email.lower()).execute()
    except Exception:
        pass
    return attempts

def get_run_notes(client, period):
    try:
        res = _sb().table("reconciliation_history").select("notes").eq("client", client).eq("period", period).execute()
        return (res.data[0].get("notes", "") or "") if res.data else ""
    except Exception:
        return ""

def update_run_notes(client, period, notes):
    try:
        _sb().table("reconciliation_history").update({"notes": str(notes)}).eq("client", client).eq("period", period).execute()
    except Exception:
        pass

# ── EMAIL (Resend) ────────────────────────────────────────────────
def send_code_email(to_email, name, code, subject="Your PayMatch verification code", heading="Verify your account"):
    api_key = os.getenv("RESEND_API_KEY", "")
    if not api_key:
        return False
    resend.api_key = api_key
    from_addr = os.getenv("RESEND_FROM", "PayMatch <onboarding@resend.dev>")
    html = f"""<!DOCTYPE html><html><body style="margin:0;padding:0;background:#F5EFED;font-family:Lexend,Inter,-apple-system,sans-serif;">
<div style="max-width:480px;margin:40px auto;background:white;border-radius:18px;overflow:hidden;box-shadow:0 4px 32px rgba(61,8,18,0.14);">
<div style="background:linear-gradient(135deg,#3D0812,#5C1020,#8B1A2F);padding:28px;text-align:center;">
  <span style="font-size:1.7rem;font-weight:800;color:white;letter-spacing:-0.04em;">Pay<span style="color:#E8C97A">Match</span></span>
  <div style="color:rgba(255,255,255,0.42);font-size:0.7rem;margin-top:4px;letter-spacing:0.12em;text-transform:uppercase;">by ProPayHR</div>
</div>
<div style="padding:32px;">
  <p style="color:#1A0A0F;font-size:1rem;font-weight:600;margin:0 0 8px;">Hi {name},</p>
  <p style="color:#7D5A5E;font-size:0.875rem;margin:0 0 24px;line-height:1.65;">{heading}</p>
  <div style="background:linear-gradient(135deg,#FDF8EC,#FDF4F0);border:1px solid #E8C97A;border-radius:14px;padding:28px;text-align:center;margin-bottom:24px;">
    <div style="font-size:2.8rem;font-weight:800;color:#8B1A2F;letter-spacing:0.2em;line-height:1;">{code}</div>
    <div style="color:#9D7075;font-size:0.78rem;margin-top:10px;">This code is valid until you close the page.</div>
  </div>
  <p style="color:#B09898;font-size:0.78rem;margin:0;">If you didn't request this, you can safely ignore this email.</p>
</div>
<div style="background:#FAF5F5;padding:14px;text-align:center;border-top:1px solid #EADDD8;">
  <span style="color:#B09898;font-size:0.72rem;">PayMatch by ProPayHR &nbsp;·&nbsp; Benefits Reconciliation Platform</span>
</div></div></body></html>"""
    try:
        resend.Emails.send({
            "from": from_addr,
            "to": [to_email],
            "subject": subject,
            "html": html,
        })
        return True
    except Exception:
        return False

# ─────────────────────────────────────────────────────────────────
# STREAMLIT UI
# ─────────────────────────────────────────────────────────────────
st.set_page_config(page_title="PayMatch by ProPayHR", page_icon="📊", layout="wide")

# ── KEEP-ALIVE ────────────────────────────────────────────────────
@st.cache_resource
def _start_keep_alive():
    """Background thread keeps the Python process warm on Streamlit Cloud."""
    def _loop():
        while True:
            time.sleep(60)
    t = threading.Thread(target=_loop, daemon=True)
    t.start()
    return t

_start_keep_alive()

st.markdown("""
<style>
/* ProPayHR Brand Design System — Trust & Authority / Corporate Minimalism */
@import url('https://fonts.googleapis.com/css2?family=Lexend:wght@300;400;500;600;700;800&family=Source+Sans+3:wght@300;400;500;600;700&display=swap');

/* ── Brand Tokens ────────────────────────────────── */
:root {
    --burg-950: #1A0208;
    --burg-900: #3D0812;
    --burg-800: #5C1020;
    --burg-700: #6B1423;
    --burg-600: #8B1A2F;
    --burg-500: #A02235;
    --burg-400: #C43048;
    --gold-600: #8B6914;
    --gold-500: #A87C28;
    --gold-400: #C4973A;
    --gold-300: #D4AF62;
    --gold-200: #E8C97A;
    --gold-100: #F5E8C0;
    --gold-50:  #FDF8EC;
    --bg:       #FAF7F5;
    --surface:  #FFFFFF;
    --border:   #E8DDD8;
    --text-900: #1A0A0F;
    --text-700: #4A2A30;
    --text-500: #7D5A5E;
    --text-300: #B09898;
    --text-200: #D4C0C0;
}

html, body, [class*="css"] {
    font-family: 'Lexend', 'Source Sans 3', -apple-system, BlinkMacSystemFont, sans-serif !important;
}
#MainMenu, footer, header { visibility: hidden; }
.stApp { background: #FAF7F5; }
.block-container { padding-top: 1.5rem !important; padding-bottom: 2rem !important; }

/* ── Input styling ───────────────────────────────── */
.stTextInput > div > div > input,
.stNumberInput > div > div > input,
.stSelectbox > div > div > div,
.stDateInput > div > div > input {
    background-color: #FFFFFF !important;
    color: #1A0A0F !important;
    border-color: #D8C8C8 !important;
    border-radius: 8px !important;
    font-size: 0.9rem !important;
    font-family: 'Lexend', sans-serif !important;
    padding: 0.55rem 0.75rem !important;
}
.stTextInput > div > div > input::placeholder,
.stNumberInput > div > div > input::placeholder { color: #B09898 !important; }
.stTextInput > div > div > input:focus,
.stNumberInput > div > div > input:focus {
    border-color: #8B1A2F !important;
    box-shadow: 0 0 0 3px rgba(139,26,47,0.12) !important;
    outline: none !important;
}
.stSelectbox > div > div { background: #FFFFFF !important; color: #1A0A0F !important; }
label, .stTextInput label, .stNumberInput label, .stSelectbox label {
    color: #3D0812 !important; font-size: 0.84rem !important; font-weight: 600 !important;
    font-family: 'Lexend', sans-serif !important;
}

/* ── Header ──────────────────────────────────────── */
.pm-header {
    background: linear-gradient(135deg, #1A0208 0%, #3D0812 35%, #5C1020 65%, #8B1A2F 100%);
    border-radius: 16px; margin-bottom: 0;
    box-shadow: 0 8px 48px rgba(26,2,8,0.40), 0 2px 8px rgba(0,0,0,0.18),
                inset 0 1px 0 rgba(255,255,255,0.06);
    overflow: hidden; position: relative;
}
.pm-header::before {
    content:''; position:absolute; top:-60px; right:-60px; width:320px; height:320px;
    background:radial-gradient(circle,rgba(196,151,58,0.10) 0%,transparent 65%); pointer-events:none;
}
.pm-header::after {
    content:''; position:absolute; bottom:-40px; left:20%; width:200px; height:200px;
    background:radial-gradient(circle,rgba(139,26,47,0.15) 0%,transparent 70%); pointer-events:none;
}
.pm-header-body { padding: 1.75rem 2.5rem; position: relative; z-index: 1; }
.pm-header-top { display:flex; align-items:center; justify-content:space-between; margin-bottom:1rem; }
.pm-brand {
    font-size:0.6rem; font-weight:700; letter-spacing:0.24em; text-transform:uppercase;
    color:rgba(255,255,255,0.38);
    padding: 0.2rem 0.6rem;
    background: rgba(255,255,255,0.05);
    border: 1px solid rgba(255,255,255,0.1);
    border-radius: 4px;
}
.pm-user-pill {
    display:inline-flex; align-items:center; gap:0.5rem;
    background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15);
    color:rgba(255,255,255,0.85); padding:0.28rem 0.85rem 0.28rem 0.4rem;
    border-radius:24px; font-size:0.75rem; font-weight:500;
}
.pm-user-avatar {
    width:24px; height:24px; background:rgba(196,151,58,0.22); border:1px solid rgba(196,151,58,0.5);
    border-radius:50%; display:inline-flex; align-items:center; justify-content:center;
    font-size:0.65rem; font-weight:700; color:#E8C97A;
}
.pm-logo-row { display:flex; align-items:baseline; gap:0.5rem; }
.pm-wordmark { font-size:2.2rem; font-weight:800; color:#fff; letter-spacing:-0.04em; line-height:1; }
.pm-wordmark span { color:#E8C97A; }
.pm-byline { font-size:0.72rem; font-weight:500; color:rgba(255,255,255,0.38); letter-spacing:0.08em; text-transform:uppercase; margin-left:0.25rem; }
.pm-sub { color:rgba(255,255,255,0.48); font-size:0.84rem; font-weight:400; margin-top:0.35rem; letter-spacing:0.01em; }
.pm-chips { display:flex; flex-wrap:wrap; gap:0.5rem; margin-top:1.1rem; }
.pm-chip {
    display:inline-flex; align-items:center; gap:0.35rem;
    background:rgba(255,255,255,0.055); border:1px solid rgba(255,255,255,0.1);
    color:rgba(255,255,255,0.6); padding:0.26rem 0.75rem; border-radius:6px;
    font-size:0.68rem; font-weight:500; letter-spacing:0.01em;
    transition: background 0.15s, border-color 0.15s;
}
.pm-chip:hover { background:rgba(255,255,255,0.09); border-color:rgba(255,255,255,0.18); }
.pm-check { color:#E8C97A; font-size:0.7rem; }

/* ── Divider line under header ───────────────────── */
.pm-header-rule { height:3px; background:linear-gradient(90deg,#8B1A2F 0%,#C4973A 50%,#8B1A2F 100%); border-radius:0 0 2px 2px; margin-bottom:1.25rem; opacity:0.7; }

/* ── Nav buttons ─────────────────────────────────── */
div[data-testid="stColumns"] > div [data-testid="stButton"] > button[kind="primary"] {
    background:linear-gradient(180deg,#A02235 0%,#8B1A2F 100%) !important;
    box-shadow:0 1px 2px rgba(0,0,0,0.14),0 3px 12px rgba(139,26,47,0.32) !important;
}
div[data-testid="stColumns"] > div [data-testid="stButton"] > button[kind="secondary"] {
    background: #FFFFFF !important;
    color: #5C1020 !important;
    border: 1px solid #D8C8C8 !important;
    box-shadow: 0 1px 3px rgba(0,0,0,0.06) !important;
}

/* ── Section labels / descriptions ──────────────── */
.slabel {
    font-size:0.6rem; font-weight:700; letter-spacing:0.18em; text-transform:uppercase;
    color:#8B1A2F; margin-bottom:0.55rem; padding-left:0.05rem;
    display:flex; align-items:center; gap:0.4rem;
}
.slabel::before { content:''; display:inline-block; width:16px; height:2px; background:linear-gradient(90deg,#8B1A2F,#C4973A); border-radius:2px; }
.sdesc { font-size:0.82rem; color:#7D5A5E; line-height:1.6; margin-bottom:0.9rem; }

/* ── Cards ───────────────────────────────────────── */
.card {
    background:#fff; border-radius:12px; padding:1.4rem 1.6rem;
    box-shadow:0 1px 3px rgba(26,2,8,0.04),0 4px 18px rgba(26,2,8,0.06);
    border:1px solid #EADDD8; margin-bottom:0.9rem;
    transition:box-shadow 0.2s, border-color 0.2s, transform 0.15s;
}
.card:hover { box-shadow:0 2px 8px rgba(26,2,8,0.07),0 8px 28px rgba(139,26,47,0.10); border-color:#D4B8B8; transform:translateY(-1px); }
.utitle { font-size:0.9rem; font-weight:600; color:#1A0A0F; display:flex; align-items:center; gap:0.55rem; margin-bottom:0.35rem; }
.ustep {
    width:26px; height:26px; background:linear-gradient(135deg,#8B1A2F,#6B1423); color:white;
    border-radius:7px; display:inline-flex; align-items:center; justify-content:center;
    font-size:0.73rem; font-weight:700; flex-shrink:0;
    box-shadow:0 2px 6px rgba(139,26,47,0.3);
}
.utag {
    background:linear-gradient(135deg,#FDF8EC,#F5E8C0); color:#8B6914; border:1px solid #E8C97A;
    border-radius:5px; font-size:0.63rem; font-weight:700; padding:0.13rem 0.5rem;
    letter-spacing:0.07em; text-transform:uppercase;
}
.udesc { font-size:0.82rem; color:#7D5A5E; line-height:1.6; margin-bottom:0.85rem; padding-left:2.1rem; }
.spanel {
    background:#fff; border-radius:12px; padding:1.4rem 1.6rem;
    box-shadow:0 1px 3px rgba(26,2,8,0.04),0 4px 18px rgba(26,2,8,0.06);
    border:1px solid #EADDD8;
}
.info-box {
    background:linear-gradient(135deg,#FDF8EC,#FDF4F0); border-radius:10px;
    padding:0.95rem 1.1rem; margin-top:0.85rem; font-size:0.82rem; color:#5C1020;
    border:1px solid #E8C97A; line-height:1.65;
}
.info-box b { color:#3D0812; display:block; margin-bottom:0.3rem; font-size:0.83rem; font-weight:700; }

/* ── Primary buttons (Run, Sign In, etc.) ────────── */
.stButton > button {
    background:linear-gradient(180deg,#A82338 0%,#8B1A2F 55%,#6B1423 100%) !important;
    color:white !important;
    border:1px solid rgba(255,255,255,0.12) !important;
    border-bottom-color:rgba(0,0,0,0.22) !important;
    border-radius:10px !important;
    padding:0.8rem 2.5rem !important;
    font-size:0.92rem !important;
    font-weight:600 !important;
    font-family:'Lexend',sans-serif !important;
    letter-spacing:0.02em !important;
    box-shadow:0 1px 3px rgba(0,0,0,0.16),0 4px 14px rgba(139,26,47,0.40) !important;
    transition:all 0.18s ease !important;
    width:100% !important;
    min-height: 46px !important;
}
.stButton > button:hover:not(:disabled) {
    background:linear-gradient(180deg,#C43048 0%,#A02235 55%,#8B1A2F 100%) !important;
    box-shadow:0 2px 6px rgba(0,0,0,0.16),0 8px 24px rgba(139,26,47,0.50) !important;
    transform:translateY(-1px) !important;
}
.stButton > button:active:not(:disabled) { transform:translateY(0) !important; }
.stButton > button:disabled { opacity:0.40 !important; cursor:not-allowed !important; }

/* ── Download / secondary action buttons ────────── */
[data-testid="stDownloadButton"] > button {
    background:linear-gradient(180deg,#C4973A 0%,#A87C28 55%,#8B6914 100%) !important;
    color:white !important;
    border:1px solid rgba(255,255,255,0.14) !important;
    border-bottom-color:rgba(0,0,0,0.18) !important;
    border-radius:10px !important;
    font-weight:600 !important;
    font-family:'Lexend',sans-serif !important;
    box-shadow:0 1px 3px rgba(0,0,0,0.12),0 4px 14px rgba(196,151,58,0.38) !important;
    width:100% !important;
    padding:0.8rem !important;
    font-size:0.9rem !important;
    letter-spacing:0.02em !important;
    transition:all 0.18s ease !important;
    min-height: 46px !important;
}
[data-testid="stDownloadButton"] > button:hover {
    background:linear-gradient(180deg,#D4A84A 0%,#C4973A 55%,#A87C28 100%) !important;
    box-shadow:0 2px 6px rgba(0,0,0,0.12),0 8px 24px rgba(196,151,58,0.48) !important;
    transform:translateY(-1px) !important;
}

/* ── Metric tiles ─────────────────────────────────── */
[data-testid="metric-container"] {
    background:white; border-radius:12px; padding:1.1rem 1.3rem;
    box-shadow:0 1px 3px rgba(26,2,8,0.04),0 3px 10px rgba(26,2,8,0.06);
    border:1px solid #EADDD8; border-top:3px solid #8B1A2F;
}
[data-testid="stMetricLabel"] > div {
    font-size:0.68rem !important; font-weight:700 !important; color:#9D7075 !important;
    text-transform:uppercase !important; letter-spacing:0.1em !important;
    font-family:'Lexend',sans-serif !important;
}
[data-testid="stMetricValue"] > div {
    font-size:1.9rem !important; font-weight:800 !important; color:#1A0A0F !important;
    letter-spacing:-0.03em !important; line-height:1.1 !important;
}

/* ── Misc ─────────────────────────────────────────── */
.stAlert { border-radius:10px !important; font-size:0.86rem !important; }
hr { border-color:#E8DDD8 !important; margin:1rem 0 !important; }
[data-testid="stDataFrame"] {
    border-radius:10px !important; overflow:hidden !important;
    box-shadow:0 1px 3px rgba(26,2,8,0.04),0 3px 10px rgba(26,2,8,0.06) !important;
    border:1px solid #EADDD8 !important;
}

/* ── Result / urgency ─────────────────────────────── */
.urgency {
    background:linear-gradient(135deg,#FFFCFC,#FEF2F2);
    border:1px solid #FECACA; border-left:4px solid #DC2626;
    border-radius:12px; padding:1.1rem 1.35rem; margin:0.85rem 0;
    display:flex; align-items:flex-start; gap:1rem;
    box-shadow:0 2px 8px rgba(220,38,38,0.08);
}
.urgency-amount { font-weight:700; color:#991B1B; font-size:1.05rem; }
.urgency-note { color:#7F1D1D; font-size:0.82rem; margin-top:0.25rem; line-height:1.55; }
.res-bar {
    background:linear-gradient(135deg,#FDF8F8,#FDF4F0);
    border:1px solid #E8CCCC; border-left:4px solid #8B1A2F;
    border-radius:12px; padding:1rem 1.35rem; margin:0.85rem 0;
    display:flex; align-items:center; justify-content:space-between; flex-wrap:wrap; gap:0.5rem;
    box-shadow:0 2px 8px rgba(139,26,47,0.08);
}
.res-bar-ttl { font-size:0.9rem; font-weight:700; color:#3D0812; }
.res-bar-sub { font-size:0.78rem; color:#7D5A5E; margin-top:0.12rem; }
.res-bar-count { font-size:1.15rem; font-weight:800; color:#8B1A2F; }
.run-timestamp { font-size:0.75rem; color:#B09898; margin-top:0.45rem; display:flex; align-items:center; gap:0.3rem; }
.ai-hdr { display:flex; align-items:center; justify-content:space-between; margin:1.35rem 0 0.6rem; }
.ai-title { font-size:0.6rem; font-weight:700; letter-spacing:0.16em; text-transform:uppercase; color:#9D7075; }
.ai-badge {
    background:linear-gradient(135deg,#FDF8EC,#FDF4F0); color:#8B1A2F;
    border:1px solid #E8C97A; border-radius:20px;
    padding:0.18rem 0.75rem; font-size:0.72rem; font-weight:700;
}
.empty-prompt {
    text-align:center; color:#B09898; font-size:0.9rem; padding:2rem 1.5rem;
    background:white; border-radius:12px; border:1px dashed #D4B8B8;
    margin-top:0.5rem; line-height:1.65;
    box-shadow:0 1px 3px rgba(26,2,8,0.04);
}
.trend-card {
    background:white; border-radius:12px; padding:1.6rem;
    box-shadow:0 1px 3px rgba(26,2,8,0.04),0 4px 18px rgba(26,2,8,0.06);
    border:1px solid #EADDD8; margin-bottom:1rem;
}
.delta { border-radius:10px; padding:0.9rem 1.35rem; font-weight:600; font-size:0.86rem; margin-top:0.6rem; }
.delta-dn { background:#ECFDF5; border:1px solid #A7F3D0; color:#065F46; }
.delta-up { background:#FEF2F2; border:1px solid #FECACA; color:#991B1B; }

/* ── Download area ───────────────────────────────── */
.dl-section {
    background:linear-gradient(135deg,#FDF8F5,#FDF4EE);
    border:1px solid #E8DDD8; border-top:2px solid #C4973A;
    border-radius:12px; padding:1.1rem 1.35rem; margin-top:1.1rem;
}
.dl-title { font-size:0.68rem; font-weight:700; letter-spacing:0.12em; text-transform:uppercase; color:#8B1A2F; margin-bottom:0.8rem; }
.dl-desc { font-size:0.8rem; color:#7D5A5E; margin-bottom:0.6rem; line-height:1.55; }

/* ── Auth page ───────────────────────────────────── */
[data-testid="stImage"] {
    width: 100% !important;
    display: flex !important;
    justify-content: center !important;
    align-items: center !important;
    background: linear-gradient(160deg, #0D0304 0%, #1A0208 55%, #2A0810 100%) !important;
    border-radius: 16px !important;
    padding: 1.75rem 1.5rem !important;
    margin-bottom: 1.25rem !important;
    box-shadow: 0 6px 32px rgba(13,3,4,0.28) !important;
}
[data-testid="stImage"] img {
    mix-blend-mode: screen !important;
    max-height: 68px !important;
    width: auto !important;
    object-fit: contain !important;
}
.auth-logo { text-align:center; margin-bottom:1.5rem; padding-top:0; }
.auth-logotype { font-size:2.2rem; font-weight:800; color:#1A0A0F; letter-spacing:-0.04em; line-height:1; }
.auth-logotype span { color:#8B1A2F; }
.auth-tagline { color:#B09898; font-size:0.82rem; margin-top:0.35rem; letter-spacing:0.01em; }
.auth-value { font-size:0.88rem; font-weight:500; color:#4A2A30; margin-top:0.65rem; letter-spacing:0.01em; line-height:1.5; font-style:italic; }
.auth-title { font-size:1.15rem; font-weight:700; color:#1A0A0F; margin-bottom:0.28rem; text-align:center; }
.auth-sub { color:#7D5A5E; font-size:0.84rem; margin-bottom:1.5rem; line-height:1.6; text-align:center; }
.auth-switch { text-align:center; font-size:0.82rem; color:#7D5A5E; margin-top:1rem; padding-top:1rem; border-top:1px solid #F5EDEB; }
/* ── Forgot password — text link via key-based CSS ── */
.st-key-goto_forgot { display:flex !important; justify-content:center !important; margin-top:0.1rem !important; }
.st-key-goto_forgot button {
    background: none !important; border: none !important; box-shadow: none !important;
    color: #8B1A2F !important; font-size: 0.82rem !important; font-weight: 500 !important;
    padding: 0.25rem 0.5rem !important; width: auto !important; min-height: unset !important;
    text-decoration: underline !important; text-underline-offset: 2px !important;
    letter-spacing: 0 !important; transform: none !important;
}
.st-key-goto_forgot button:hover:not(:disabled) {
    color: #5C1020 !important; background: none !important;
    box-shadow: none !important; transform: none !important;
}
/* ── Guest button — ghost/secondary style ── */
.st-key-guest_mode_btn button {
    background: white !important; color: #5C1020 !important;
    border: 1.5px solid #D4B8B8 !important; box-shadow: 0 1px 3px rgba(0,0,0,0.05) !important;
}
.st-key-guest_mode_btn button:hover:not(:disabled) {
    background: #FAF5F5 !important; border-color: #C4A8A8 !important;
    box-shadow: 0 1px 5px rgba(0,0,0,0.08) !important; transform: none !important;
}
.pw-reqs { background:#FAF5F5; border-radius:9px; padding:0.75rem 1rem; margin-top:0.4rem; border:1px solid #E8DDD8; }
.pw-req { display:flex; align-items:center; gap:0.4rem; font-size:0.78rem; margin-bottom:0.2rem; line-height:1; }
.pw-req:last-child { margin-bottom:0; }
.pr-ok { color:#059669; font-weight:700; }
.pr-fail { color:#DC2626; font-weight:700; }
.pr-neutral { color:#B09898; }
.verify-box {
    background:linear-gradient(135deg,#FDF8EC,#FDF4F0);
    border:1px solid #E8C97A; border-radius:12px;
    padding:1.25rem; text-align:center; margin-bottom:1.1rem;
}
.verify-code-shown {
    font-size:2.4rem; font-weight:800; color:#8B1A2F;
    letter-spacing:0.22em; line-height:1; font-family:'Courier New',monospace;
}
.verify-note { color:#7D5A5E; font-size:0.8rem; margin-top:0.55rem; line-height:1.55; }

/* ── Profile page ────────────────────────────────── */
.profile-card {
    background:white; border-radius:14px; padding:1.6rem;
    box-shadow:0 1px 3px rgba(26,2,8,0.04),0 4px 18px rgba(26,2,8,0.06);
    border:1px solid #EADDD8; margin-bottom:1rem;
}
.profile-avatar {
    width:58px; height:58px;
    background:linear-gradient(135deg,#8B1A2F,#5C1020);
    border-radius:15px; display:flex; align-items:center; justify-content:center;
    font-size:1.45rem; font-weight:800; color:white; flex-shrink:0;
    box-shadow:0 4px 14px rgba(139,26,47,0.35);
}
.profile-name { font-size:1.15rem; font-weight:700; color:#1A0A0F; }
.profile-email { font-size:0.84rem; color:#7D5A5E; margin-top:0.12rem; }
.profile-since { font-size:0.76rem; color:#B09898; margin-top:0.28rem; }
.profile-stat {
    background:linear-gradient(135deg,#FAF5F5,#FDF8F5);
    border-radius:10px; padding:0.9rem 1rem; border:1px solid #EADDD8; text-align:center;
}
.profile-stat-val { font-size:1.55rem; font-weight:800; color:#1A0A0F; letter-spacing:-0.02em; }
.profile-stat-lbl { font-size:0.68rem; font-weight:700; color:#B09898; text-transform:uppercase; letter-spacing:0.09em; margin-top:0.18rem; }

/* ── Status badges ───────────────────────────────── */
.sbadge { display:inline-flex; align-items:center; gap:0.3rem; border-radius:20px; padding:0.2rem 0.65rem; font-size:0.72rem; font-weight:700; white-space:nowrap; margin-bottom:0.3rem; }
.sbadge-incomplete { background:#FEF2F2; color:#DC2626; border:1px solid #FECACA; }
.sbadge-in-progress { background:#FFFBEB; color:#D97706; border:1px solid #FDE68A; }
.sbadge-complete { background:#ECFDF5; color:#059669; border:1px solid #A7F3D0; }

/* ── Clickable status cycle buttons ─────────────── */
.stMarkdown:has(.scycle-incomplete) + div [data-testid="stButton"] > button,
.stMarkdown:has(.scycle-incomplete) + [data-testid="stButton"] > button {
    background:#FEF2F2 !important; color:#DC2626 !important; border:1px solid #FECACA !important;
    border-radius:20px !important; padding:0.18rem 0.65rem !important; font-size:0.72rem !important;
    font-weight:700 !important; box-shadow:none !important; transform:none !important;
    width:auto !important; min-height:unset !important; line-height:1.4 !important; cursor:pointer !important;
}
.stMarkdown:has(.scycle-in-progress) + div [data-testid="stButton"] > button,
.stMarkdown:has(.scycle-in-progress) + [data-testid="stButton"] > button {
    background:#FFFBEB !important; color:#D97706 !important; border:1px solid #FDE68A !important;
    border-radius:20px !important; padding:0.18rem 0.65rem !important; font-size:0.72rem !important;
    font-weight:700 !important; box-shadow:none !important; transform:none !important;
    width:auto !important; min-height:unset !important; line-height:1.4 !important; cursor:pointer !important;
}
.stMarkdown:has(.scycle-complete) + div [data-testid="stButton"] > button,
.stMarkdown:has(.scycle-complete) + [data-testid="stButton"] > button {
    background:#ECFDF5 !important; color:#059669 !important; border:1px solid #A7F3D0 !important;
    border-radius:20px !important; padding:0.18rem 0.65rem !important; font-size:0.72rem !important;
    font-weight:700 !important; box-shadow:none !important; transform:none !important;
    width:auto !important; min-height:unset !important; line-height:1.4 !important; cursor:pointer !important;
}

/* ── Delete confirm banner ───────────────────────── */
.del-confirm { background:#FEF2F2; border:1px solid #FECACA; border-left:3px solid #DC2626; border-radius:10px; padding:0.8rem 1.1rem; margin:0.2rem 0 0.3rem; font-size:0.84rem; color:#7F1D1D; line-height:1.6; }

/* ── History search ──────────────────────────────── */
.hist-search { background:white; border:1px solid #EADDD8; border-radius:9px; padding:0.55rem 0.9rem; font-size:0.86rem; color:#1A0A0F; width:100%; margin-bottom:1rem; box-shadow:0 1px 3px rgba(26,2,8,0.04); }

/* ── Notes field ─────────────────────────────────── */
.notes-wrap { background:white; border:1px solid #EADDD8; border-radius:14px; padding:1.35rem 1.6rem; margin-top:1rem; box-shadow:0 1px 3px rgba(26,2,8,0.04); }
.notes-label { font-size:0.6rem; font-weight:700; letter-spacing:0.16em; text-transform:uppercase; color:#8B1A2F; margin-bottom:0.55rem; }
.hist-note { font-size:0.78rem; color:#7D5A5E; font-style:italic; padding:0.2rem 0 0.35rem 0.5rem; display:flex; gap:0.4rem; align-items:flex-start; }

/* ── Footer ──────────────────────────────────────── */
.pm-footer {
    text-align:center; color:#B09898; font-size:0.73rem;
    padding:1.5rem 1rem 0.5rem; margin-top:1rem; letter-spacing:0.02em;
    border-top:1px solid #EADDD8;
}
.pm-footer strong { color:#8B1A2F; font-weight:600; }

/* ── Welcome bar ─────────────────────────────────── */
.welcome-bar {
    background:linear-gradient(135deg,#FDF8F5,white);
    border:1px solid #EADDD8; border-left:3px solid #8B1A2F;
    border-radius:10px; padding:0.8rem 1.25rem;
    font-size:0.9rem; color:#4A2A30;
    margin-bottom:1.35rem;
    box-shadow:0 1px 4px rgba(26,2,8,0.04);
}

/* ── Mobile responsive ───────────────────────────── */
@media (max-width: 768px) {
  .block-container { padding: 0.6rem 0.5rem !important; }
  .pm-header-body { padding: 1.1rem 1.25rem !important; }
  .pm-wordmark { font-size: 1.65rem !important; }
  .pm-chips { display: none !important; }
  .pm-sub { font-size: 0.78rem !important; }
  [data-testid="column"] { min-width: 100% !important; flex: none !important; }
  .stButton > button { min-height: 50px !important; font-size: 0.95rem !important; }
  [data-testid="stDownloadButton"] > button { min-height: 50px !important; }
  .hist-table-wrap { overflow-x: auto; -webkit-overflow-scrolling: touch; }
  [data-testid="stMetricValue"] > div { font-size: 1.5rem !important; }
  .pm-header { margin-bottom: 0.5rem !important; }
  div[data-testid="column"]:has(.stButton) .stButton > button { width: 100% !important; }
}

/* ── Streamlit tabs ──────────────────────────────── */
/* ── Auth form card ── */
[data-testid="stTabs"] {
    background: white !important;
    border-radius: 16px !important;
    padding: 1.25rem 1.5rem 0.75rem !important;
    box-shadow: 0 1px 4px rgba(26,2,8,0.04), 0 12px 40px rgba(26,2,8,0.09) !important;
    border: 1px solid #E8DDD8 !important;
    margin-bottom: 0 !important;
}
.stTabs [data-baseweb="tab-list"] { gap:0; background:#F5EDEB; border-radius:9px; padding:3px; }
.stTabs [data-baseweb="tab"] { border-radius:6px; font-size:0.86rem; font-weight:500; padding:0.5rem 1.35rem; color:#9D7075; }
.stTabs [aria-selected="true"] { background:white !important; color:#1A0A0F !important; font-weight:700 !important; box-shadow:0 1px 3px rgba(26,2,8,0.08); }

/* ── Expander ────────────────────────────────────── */
.streamlit-expanderHeader { font-size:0.88rem !important; font-weight:600 !important; color:#3D0812 !important; }

/* ── Spinner ─────────────────────────────────────── */
.stSpinner > div { border-top-color:#8B1A2F !important; }
</style>
""", unsafe_allow_html=True)

# ── SESSION STATE ─────────────────────────────────────────────────
_defaults = {
    "logged_in": False, "current_user": None,
    "auth_tab": "login",
    "pending_email": None, "pending_code": None,
    "page": "app",
    "recon_df": None, "last_run_ts": None,
    "delete_confirm_key": None,
    "last_activity": None,
    "session_expired": False,
    "guest_mode": False,
}
for _k, _v in _defaults.items():
    if _k not in st.session_state:
        st.session_state[_k] = _v

# ─────────────────────────────────────────────────────────────────
# AUTH VIEW
# ─────────────────────────────────────────────────────────────────
if not st.session_state.logged_in:

    if st.session_state.get("session_expired"):
        st.warning("Your session expired after 30 minutes of inactivity. Please sign in again.")
        st.session_state.session_expired = False

    _, auth_col, _ = st.columns([1, 1.15, 1])
    with auth_col:

        st.image("ProPayHR-2024-red-TRANSPARENT.png", width=150)
        st.markdown("""
        <div class="auth-logo">
            <div class="auth-logotype">Pay<span>Match</span></div>
            <div class="auth-tagline">Benefits Reconciliation Platform &nbsp;·&nbsp; by ProPayHR</div>
            <div class="auth-value">Catch payroll errors before they cost you.</div>
        </div>
        """, unsafe_allow_html=True)

        # ── VERIFY ────────────────────────────────────────────────
        if st.session_state.auth_tab == "verify":
            st.markdown('<div class="auth-title">Check your email</div>', unsafe_allow_html=True)
            st.markdown(f'<div class="auth-sub">We sent a 6-digit code to <strong>{st.session_state.pending_email}</strong>. Enter it below to activate your account.</div>', unsafe_allow_html=True)

            if st.session_state.pending_code:
                st.markdown(f"""
                <div class="verify-box">
                    <div style="font-size:0.72rem;font-weight:600;letter-spacing:0.1em;text-transform:uppercase;color:#64748B;margin-bottom:0.5rem;">Dev mode — SMTP not configured</div>
                    <div class="verify-code-shown">{st.session_state.pending_code}</div>
                    <div class="verify-note">Set SMTP_HOST, SMTP_USER, and SMTP_PASS environment variables to send real emails.</div>
                </div>
                """, unsafe_allow_html=True)

            code_input = st.text_input("Verification Code", placeholder="Enter your 6-digit code",
                                       max_chars=6, key="verify_code_input")
            if st.button("Verify & Sign In", key="do_verify"):
                if not code_input.strip():
                    st.error("Please enter the 6-digit code.")
                else:
                    verified = verify_user(st.session_state.pending_email, code_input.strip())
                    if verified:
                        st.session_state.logged_in = True
                        st.session_state.current_user = verified
                        st.session_state.auth_tab = "login"
                        st.session_state.pending_email = None
                        st.session_state.pending_code = None
                        st.rerun()
                    else:
                        st.error("That code doesn't match. Double-check your email and try again.")

            if st.button("← Back to sign up", key="back_signup"):
                st.session_state.auth_tab = "signup"; st.rerun()

        # ── FORGOT PASSWORD ───────────────────────────────────────
        elif st.session_state.auth_tab == "forgot":
            st.markdown('<div class="auth-title">Reset your password</div>', unsafe_allow_html=True)
            st.markdown('<div class="auth-sub">Enter your email address and we\'ll send you a 6-digit reset code.</div>', unsafe_allow_html=True)

            fp_email = st.text_input("Email address", placeholder="you@company.com", key="fp_email")
            if st.button("Send Reset Code", key="do_forgot"):
                if not fp_email.strip() or "@" not in fp_email:
                    st.error("Please enter a valid email address.")
                else:
                    u, code = generate_reset_code(fp_email.strip())
                    if u is None:
                        st.error("No account found with that email address.")
                    else:
                        sent = send_code_email(fp_email.strip(), u["name"].split()[0], code,
                            subject="Reset your PayMatch password",
                            heading="Use the code below to reset your password.")
                        st.session_state.auth_tab = "reset"
                        st.session_state.pending_email = fp_email.strip().lower()
                        st.session_state.pending_code = None if sent else code
                        st.rerun()

            if st.button("← Back to sign in", key="back_login_forgot"):
                st.session_state.auth_tab = "login"; st.rerun()

        # ── RESET PASSWORD ────────────────────────────────────────
        elif st.session_state.auth_tab == "reset":
            st.markdown('<div class="auth-title">Create new password</div>', unsafe_allow_html=True)
            st.markdown(f'<div class="auth-sub">Enter the reset code sent to <strong>{st.session_state.pending_email}</strong> and choose a new password.</div>', unsafe_allow_html=True)

            if st.session_state.pending_code:
                st.markdown(f"""
                <div class="verify-box">
                    <div style="font-size:0.72rem;font-weight:600;letter-spacing:0.1em;text-transform:uppercase;color:#64748B;margin-bottom:0.5rem;">Dev mode — reset code</div>
                    <div class="verify-code-shown">{st.session_state.pending_code}</div>
                </div>
                """, unsafe_allow_html=True)

            rs_code = st.text_input("Reset Code", placeholder="6-digit code", max_chars=6, key="rs_code")
            rs_pw   = st.text_input("New Password", type="password", placeholder="Choose a strong password", key="rs_pw")
            if rs_pw:
                checks = pw_checks(rs_pw)
                rows_html = ""
                for key_c, label in [("length","At least 8 characters"),("upper","Uppercase letter"),("lower","Lowercase letter"),("digit","Number"),("special","Special character")]:
                    ok = checks[key_c]
                    rows_html += f'<div class="pw-req"><span class="{"pr-ok" if ok else "pr-fail"}">{"✓" if ok else "✗"}</span><span style="color:{"#374151" if ok else "#6B7280"};font-size:0.76rem;">{label}</span></div>'
                st.markdown(f'<div class="pw-reqs">{rows_html}</div>', unsafe_allow_html=True)
            rs_pw2  = st.text_input("Confirm Password", type="password", placeholder="Re-enter new password", key="rs_pw2")

            if st.button("Reset Password", key="do_reset"):
                if not rs_code.strip():
                    st.error("Please enter the reset code.")
                elif not pw_ok(rs_pw):
                    st.error("Password doesn't meet all the requirements listed above.")
                elif rs_pw != rs_pw2:
                    st.error("Passwords don't match.")
                else:
                    updated = apply_reset(st.session_state.pending_email, rs_code.strip(), rs_pw)
                    if updated:
                        st.session_state.logged_in = True
                        st.session_state.current_user = updated
                        st.session_state.auth_tab = "login"
                        st.session_state.pending_email = None
                        st.session_state.pending_code = None
                        st.success("Password reset successfully. Welcome back!")
                        st.rerun()
                    else:
                        st.error("That code is incorrect or has expired. Please request a new one.")

            if st.button("← Back", key="back_forgot"):
                st.session_state.auth_tab = "forgot"; st.rerun()

        # ── LOGIN / SIGNUP TABS ───────────────────────────────────
        else:
            login_tab, signup_tab = st.tabs(["Sign In", "Create Account"])

            with login_tab:
                st.markdown('<div style="height:0.75rem"></div>', unsafe_allow_html=True)
                st.markdown('<div class="auth-title">Welcome back</div>', unsafe_allow_html=True)
                st.markdown('<div class="auth-sub">Sign in to access your PayMatch workspace and run reconciliations.</div>', unsafe_allow_html=True)

                with st.form("login_form"):
                    li_email = st.text_input("Email address", placeholder="you@company.com", key="li_email")
                    li_pw    = st.text_input("Password", type="password", placeholder="Your password", key="li_pw")
                    submitted = st.form_submit_button("Sign In", use_container_width=True)

                # Forgot password — styled as text link via .st-key-goto_forgot CSS
                if st.button("Forgot your password?", key="goto_forgot"):
                    st.session_state.auth_tab = "forgot"; st.rerun()

                if submitted:
                    if not li_email or not li_pw:
                        st.error("Please fill in both your email and password.")
                    else:
                        lock_dt = check_account_locked(li_email)
                        if lock_dt:
                            remaining = max(1, int((lock_dt - datetime.now(timezone.utc)).total_seconds() // 60) + 1)
                            st.error(f"Account locked due to too many failed login attempts. Try again in {remaining} minute{'s' if remaining != 1 else ''}.")
                        else:
                            u = get_user(li_email)
                            if u is None:
                                st.error("No account found with that email. Double-check the address or create an account.")
                            else:
                                computed, _ = hash_pw(li_pw, u["salt"])
                                if computed != u["password_hash"]:
                                    attempts = record_failed_attempt(li_email)
                                    remaining_attempts = max(0, 5 - attempts)
                                    if attempts >= 5:
                                        st.error("Too many failed attempts. Your account has been locked for 15 minutes.")
                                    else:
                                        st.error(f"Incorrect password. {remaining_attempts} attempt{'s' if remaining_attempts != 1 else ''} remaining before your account is locked.")
                                elif not u.get("verified", False):
                                    st.warning("Your account hasn't been verified yet. Check your email for the verification code.")
                                    st.session_state.auth_tab = "verify"
                                    st.session_state.pending_email = u["email"]
                                    st.session_state.pending_code = u.get("verification_code")
                                    st.rerun()
                                else:
                                    _reset_lockout(li_email)
                                    st.session_state.logged_in = True
                                    st.session_state.current_user = get_user(li_email)
                                    st.session_state.last_activity = datetime.now()
                                    st.rerun()

            with signup_tab:
                st.markdown('<div style="height:0.75rem"></div>', unsafe_allow_html=True)
                st.markdown('<div class="auth-title">Create your account</div>', unsafe_allow_html=True)
                st.markdown('<div class="auth-sub">Set up your PayMatch account. A verification code will be sent to your email to activate it.</div>', unsafe_allow_html=True)

                su_name  = st.text_input("Full Name", placeholder="Jane Smith", key="su_name")
                su_email = st.text_input("Work Email", placeholder="you@company.com", key="su_email")
                su_pw    = st.text_input("Password", type="password", placeholder="Create a strong password", key="su_pw")

                if su_pw:
                    checks = pw_checks(su_pw)
                    rows_html = ""
                    for key_c, label in [("length","At least 8 characters"),("upper","One uppercase letter (A–Z)"),("lower","One lowercase letter (a–z)"),("digit","One number (0–9)"),("special","One special character (!@#$…)")]:
                        ok = checks[key_c]
                        rows_html += f'<div class="pw-req"><span class="{"pr-ok" if ok else "pr-fail"}">{"✓" if ok else "✗"}</span><span style="color:{"#374151" if ok else "#6B7280"};font-size:0.76rem;">{label}</span></div>'
                    st.markdown(f'<div class="pw-reqs">{rows_html}</div>', unsafe_allow_html=True)
                else:
                    st.markdown('<div class="pw-reqs"><div style="font-size:0.76rem;color:#94A3B8;">Password must include: 8+ characters, uppercase, lowercase, number, special character.</div></div>', unsafe_allow_html=True)

                su_pw2 = st.text_input("Confirm Password", type="password", placeholder="Re-enter your password", key="su_pw2")

                if st.button("Create Account", key="do_signup", use_container_width=True):
                    err = None
                    if not su_name.strip(): err = "Please enter your full name."
                    elif not su_email.strip() or "@" not in su_email: err = "Please enter a valid work email address."
                    elif get_user(su_email): err = "An account with that email already exists. Sign in instead."
                    elif not pw_ok(su_pw): err = "Password doesn't meet all requirements — check the list above."
                    elif su_pw != su_pw2: err = "Passwords don't match. Please re-enter them."
                    if err:
                        st.error(err)
                    else:
                        new_user = create_user(su_name.strip(), su_email.strip(), su_pw)
                        sent = send_code_email(su_email.strip(), su_name.strip().split()[0], new_user["verification_code"])
                        st.session_state.auth_tab = "verify"
                        st.session_state.pending_email = new_user["email"]
                        st.session_state.pending_code = None if sent else new_user["verification_code"]
                        st.rerun()

            st.markdown('<div style="display:flex;align-items:center;gap:0.75rem;margin:1.35rem 0 0.9rem;"><div style="flex:1;height:1px;background:#E8DDD8;"></div><span style="font-size:0.72rem;color:#B09898;letter-spacing:0.08em;text-transform:uppercase;white-space:nowrap;">or continue without an account</span><div style="flex:1;height:1px;background:#E8DDD8;"></div></div>', unsafe_allow_html=True)
            _gc = st.columns(1)
            with _gc[0]:
                if st.button("Try without an account  →", key="guest_mode_btn", type="secondary", use_container_width=True):
                    st.session_state.logged_in = True
                    st.session_state.guest_mode = True
                    st.session_state.current_user = {"name": "Guest", "email": "", "verified": True}
                    st.session_state.last_activity = datetime.now()
                    st.rerun()
            st.markdown("""
            <div style="background:linear-gradient(135deg,#FDF8EC,#FDF4F0);border:1px solid #E8C97A;border-radius:10px;padding:0.7rem 1rem;margin-top:0.65rem;font-size:0.78rem;color:#8B6914;line-height:1.6;">
                <strong style="color:#5C1020;">Guest mode:</strong> Run reconciliations and download reports instantly. Sign in for history, saved reports, and trend tracking.
            </div>
            """, unsafe_allow_html=True)

        st.markdown(f'<div style="text-align:center;color:#C0A8A8;font-size:0.73rem;margin-top:2rem;padding-bottom:1rem;letter-spacing:0.02em;"><strong style="color:#8B1A2F;">PayMatch</strong> by ProPayHR &nbsp;·&nbsp; Benefits Reconciliation Platform &nbsp;·&nbsp; {datetime.now().year}</div>', unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────
# MAIN APP (logged in)
# ─────────────────────────────────────────────────────────────────
else:
    is_guest = st.session_state.get("guest_mode", False)

    # ── SESSION TIMEOUT ───────────────────────────────────────────
    _now = datetime.now()
    if not is_guest and st.session_state.last_activity is not None:
        if (_now - st.session_state.last_activity).total_seconds() > 1800:
            for _k in list(st.session_state.keys()):
                del st.session_state[_k]
            st.session_state.session_expired = True
            st.rerun()
    st.session_state.last_activity = _now

    user    = st.session_state.current_user
    initial = user["name"][0].upper() if user else "?"

    # ── HEADER ────────────────────────────────────────────────────
    hdr_col, btn_col = st.columns([5, 1])
    with btn_col:
        st.markdown("<div style='height:0.5rem'></div>", unsafe_allow_html=True)
        if st.button("Sign Out", key="signout"):
            for k in ["logged_in","current_user","recon_df","last_run_ts","page","guest_mode"]:
                st.session_state[k] = (False if k in ("logged_in","guest_mode") else (None if k!="page" else "app"))
            st.rerun()

    with hdr_col:
        st.markdown(f"""
        <div class="pm-header">
            <div class="pm-header-body">
                <div class="pm-header-top">
                    <div class="pm-brand">ProPayHR &nbsp;·&nbsp; Enterprise Benefits Platform</div>
                    <div class="pm-user-pill" style="{'background:rgba(196,151,58,0.15);border-color:rgba(196,151,58,0.35);' if is_guest else ''}">
                        <span class="pm-user-avatar">{'G' if is_guest else initial}</span>
                        {'Guest Mode' if is_guest else user["name"].split()[0]}
                    </div>
                </div>
                <div class="pm-logo-row">
                    <div class="pm-wordmark">Pay<span>Match</span></div>
                    <span class="pm-byline">by ProPayHR</span>
                </div>
                <div class="pm-sub">Automated Benefits Reconciliation &nbsp;&nbsp;|&nbsp;&nbsp; Broker vs Paycor &nbsp;&nbsp;|&nbsp;&nbsp; Payroll Accuracy Platform</div>
                <div class="pm-chips">
                    <span class="pm-chip"><span class="pm-check">&#10003;</span>&nbsp;Broker vs Paycor Comparison</span>
                    <span class="pm-chip"><span class="pm-check">&#10003;</span>&nbsp;Any Client, Any Month</span>
                    <span class="pm-chip"><span class="pm-check">&#10003;</span>&nbsp;Automatic Benefit Matching</span>
                    <span class="pm-chip"><span class="pm-check">&#10003;</span>&nbsp;Excel Report &amp; Dashboard Export</span>
                    <span class="pm-chip"><span class="pm-check">&#10003;</span>&nbsp;Full Run History &amp; Trend View</span>
                </div>
            </div>
        </div>
        <div class="pm-header-rule"></div>
        """, unsafe_allow_html=True)

    # ── NAV ───────────────────────────────────────────────────────
    _nav_cols = st.columns(1 if is_guest else 2)
    with _nav_cols[0]:
        if st.button("Reconciliation", key="nav_app",
                     type="primary" if st.session_state.page=="app" else "secondary"):
            st.session_state.page = "app"; st.rerun()
    if not is_guest:
        with _nav_cols[1]:
            if st.button("My Profile", key="nav_profile",
                         type="primary" if st.session_state.page=="profile" else "secondary"):
                st.session_state.page = "profile"; st.rerun()

    st.markdown("<div style='height:0.25rem'></div>", unsafe_allow_html=True)

    # ═════════════════════════════════════════════════════════════
    # PAGE: RECONCILIATION
    # ═════════════════════════════════════════════════════════════
    if st.session_state.page == "app":

        if is_guest:
            st.markdown("""
            <div class="welcome-bar" style="border-left-color:#C4973A;background:linear-gradient(135deg,#FDF8EC,#FDFAF5);">
                <strong style="color:#8B6914;">Guest Mode</strong> — Run a reconciliation and download your report below.
                <span style="color:#9D7075;font-size:0.82rem;"> Sign in to save history and access reports anytime.</span>
            </div>
            """, unsafe_allow_html=True)
        else:
            first_name = user["name"].split()[0]
            st.markdown(f'<div class="welcome-bar">Welcome back, <strong style="color:#3D0812;">{first_name}</strong>. Upload your files below to run a reconciliation. It only takes a moment.</div>', unsafe_allow_html=True)

        # ── UPLOAD + CONFIG ───────────────────────────────────────
        left, right = st.columns([3, 2], gap="large")

        with left:
            st.markdown('<div class="slabel">Upload Files</div>', unsafe_allow_html=True)
            st.markdown('<div class="sdesc">Upload both files below, then click Run. PayMatch compares them line by line and flags every mismatch.</div>', unsafe_allow_html=True)

            st.markdown("""
            <div class="card">
                <div class="utitle"><span class="ustep">1</span>Deduction Report<span class="utag">Paycor</span></div>
                <div class="udesc">Export from Paycor. This is what Paycor is currently deducting from each employee's paycheck. Accepted formats: .xlsx, .xls, .csv</div>
            </div>
            """, unsafe_allow_html=True)
            paycor_file = st.file_uploader("Deduction Report", type=["xlsx","xls","csv"],
                                            label_visibility="collapsed", key="paycor")
            if paycor_file and not paycor_file.name.lower().endswith((".xlsx",".xls",".csv")):
                st.error(f"'{paycor_file.name}' isn't a supported file type. Please upload an Excel (.xlsx / .xls) or CSV file.")

            st.markdown("<div style='height:0.4rem'></div>", unsafe_allow_html=True)

            st.markdown("""
            <div class="card">
                <div class="utitle"><span class="ustep">2</span>Master Mapping Report<span class="utag">Broker</span></div>
                <div class="udesc">Enrollment export from your broker. The second row should contain benefit labels — PayMatch uses them to automatically find and compare the right columns. Accepted formats: .xlsx, .xls, .csv</div>
            </div>
            """, unsafe_allow_html=True)
            broker_file = st.file_uploader("Master Mapping Report", type=["xlsx","xls","csv"],
                                            label_visibility="collapsed", key="broker")
            if broker_file and not broker_file.name.lower().endswith((".xlsx",".xls",".csv")):
                st.error(f"'{broker_file.name}' isn't a supported file type. Please upload an Excel (.xlsx / .xls) or CSV file.")

        with right:
            st.markdown('<div class="slabel">Configuration</div>', unsafe_allow_html=True)
            st.markdown('<div class="sdesc">Enter the client name and choose the month you\'re reconciling. This labels your report and saves the run to your history.</div>', unsafe_allow_html=True)

            client_name = st.text_input("Client Name", value="", placeholder="e.g. Acme Corporation",
                help="The name of the client this reconciliation is for. It will appear in the report header.")

            mo_col, yr_col = st.columns(2)
            with mo_col:
                sel_month = st.selectbox("Month", options=MONTHS,
                    index=datetime.now().month - 1, key="period_month")
            with yr_col:
                sel_year = st.number_input("Year", min_value=2020, max_value=2040,
                    value=datetime.now().year, step=1, key="period_year", format="%d")
            period = f"{sel_month} {int(sel_year)}"

            tolerance = st.number_input("Acceptable Difference ($)", min_value=0.0, max_value=1.0,
                value=0.05, step=0.01,
                help="Two amounts this close to each other will be treated as a match.")
            st.markdown('<div style="font-size:0.78rem;color:#7D5A5E;line-height:1.55;margin-top:0.3rem;margin-bottom:0.6rem;">How close two dollar amounts need to be to count as a match. The default ($0.05) works for most cases — it accounts for small rounding differences between systems.</div>', unsafe_allow_html=True)

            st.markdown("""
            <div class="info-box">
                <b>How PayMatch works</b>
                Both files use per-paycheck dollar amounts and are compared directly — no math needed on your end.
                Benefit labels are read automatically from your enrollment file, so nothing needs to be set up manually.
                Results: green = match, blue = missing deduction, orange = amount differs, yellow = needs manual review.
            </div>
            """, unsafe_allow_html=True)

        st.markdown("<div style='height:1rem'></div>", unsafe_allow_html=True)
        st.divider()

        # ── RUN BUTTON ────────────────────────────────────────────
        ready = paycor_file is not None and broker_file is not None and bool(client_name.strip())
        _, btn_col2, _ = st.columns([1,2,1])
        with btn_col2:
            run = st.button("Run Reconciliation", type="primary", disabled=not ready)

        if not ready and not run:
            missing = []
            if paycor_file is None: missing.append("Deduction Report")
            if broker_file is None: missing.append("Master Mapping Report")
            if not client_name.strip(): missing.append("Client Name")
            st.markdown(f'<div class="empty-prompt">Please provide your <strong>{", ".join(missing)}</strong> above to get started.</div>', unsafe_allow_html=True)

        # ── RESULTS ───────────────────────────────────────────────
        if run:
            with st.spinner("Analyzing files and comparing broker enrollment against Paycor deductions…"):
                try:
                    paycor_df            = parse_paycor(paycor_file)
                    broker_df, crosswalk = parse_broker_and_crosswalk(broker_file)
                    if len(paycor_df) == 0:
                        st.error(f"The Deduction Report ({paycor_file.name}) appears to be empty or couldn't be read. Make sure it's a valid Paycor export with at least one row of data.")
                        st.stop()
                    if len(broker_df) == 0:
                        st.error(f"The Master Mapping Report ({broker_file.name}) appears to be empty or couldn't be read. Make sure the file has benefit labels in the second row and employee data below that.")
                        st.stop()
                    recon_df = reconcile(broker_df, paycor_df, crosswalk, tolerance)
                    st.session_state.recon_df = recon_df
                    st.session_state.last_run_ts = datetime.now()
                    # Build Excel once — reuse for download buttons and local storage
                    _excel_full    = build_excel(recon_df, client_name.strip())
                    _excel_actions = build_action_items_only(recon_df, client_name.strip())
                    if not is_guest:
                        save_to_history(recon_df, client_name.strip(), period, run_by=user["name"], excel_bytes=_excel_full)

                    if len(recon_df) == 0:
                        st.warning("No matching employees found. Make sure first and last names are spelled the same way in both files, and that the enrollment file includes per-paycheck dollar amounts.")
                    else:
                        total   = len(recon_df)
                        n_ok    = int((recon_df["Action"]=="OK").sum())
                        n_add   = int(recon_df["Action"].str.startswith("Add").sum())
                        n_chg   = int(recon_df["Action"].str.startswith("Change").sum())
                        n_rev   = int(recon_df["Action"].str.startswith("Review").sum())
                        mo_var  = round(recon_df[recon_df["Action"].str.startswith("Add")]["Broker /pay"].sum()*26/12, 2)
                        n_emps  = recon_df[recon_df["Action"]!="OK"]["First Name"].nunique()
                        disc_ct = n_add + n_chg + n_rev
                        ts      = st.session_state.last_run_ts.strftime("%-I:%M %p on %B %-d, %Y")

                        st.markdown("<div style='height:0.75rem'></div>", unsafe_allow_html=True)

                        if disc_ct == 0:
                            st.success(f"All {total:,} benefit lines match — nothing needs attention.")
                        else:
                            st.markdown(f"""
                            <div class="res-bar">
                                <div>
                                    <div class="res-bar-ttl">Reconciliation Complete — {client_name} · {period}</div>
                                    <div class="res-bar-sub">{total:,} lines compared &nbsp;·&nbsp; {n_emps} employees with discrepancies</div>
                                    <div class="run-timestamp">⏱ Ran at {ts}</div>
                                </div>
                                <div class="res-bar-count">{disc_ct} need attention</div>
                            </div>
                            """, unsafe_allow_html=True)

                        c1,c2,c3,c4,c5 = st.columns(5)
                        c1.metric("Lines Compared", f"{total:,}")
                        c2.metric("Matched OK",     f"{n_ok:,}")
                        c3.metric("Add to Paycor",  f"{n_add:,}")
                        c4.metric("Amount Change",  f"{n_chg:,}")
                        c5.metric("Needs Review",   f"{n_rev:,}")

                        if mo_var > 0:
                            st.markdown(f"""
                            <div class="urgency">
                                <div style="font-size:1.15rem;flex-shrink:0;margin-top:0.05rem;">⚠</div>
                                <div>
                                    <div class="urgency-amount">${mo_var:,.2f} / month not being collected</div>
                                    <div class="urgency-note">{n_add} employees are enrolled with the broker but have no deduction set up in Paycor. These need to be corrected immediately to avoid further financial exposure.</div>
                                </div>
                            </div>
                            """, unsafe_allow_html=True)

                        disc = recon_df[recon_df["Action"]!="OK"]
                        if len(disc) > 0:
                            st.markdown(f"""
                            <div class="ai-hdr">
                                <div class="ai-title">Action Items — what needs fixing</div>
                                <div class="ai-badge">{len(disc)} lines &nbsp;·&nbsp; {n_emps} employees</div>
                            </div>
                            """, unsafe_allow_html=True)
                            st.dataframe(disc[DISPLAY_COLS], use_container_width=True, hide_index=True)

                        # Download section
                        st.markdown("""
                        <div class="dl-section">
                            <div class="dl-title">Download Results</div>
                        </div>
                        """, unsafe_allow_html=True)
                        dl1, dl2 = st.columns(2)
                        safe_name = client_name.strip().replace(' ','_')
                        date_str  = datetime.now().strftime('%Y%m%d')
                        with dl1:
                            st.markdown('<div class="dl-desc"><strong>Full Excel Report</strong><br>Dashboard, all action items, and every reconciliation line with color coding.</div>', unsafe_allow_html=True)
                            st.download_button(
                                "⬇  Download Full Report (.xlsx)",
                                data=_excel_full,
                                file_name=f"PayMatch_{safe_name}_{date_str}_Full.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key="dl_full")
                        with dl2:
                            st.markdown('<div class="dl-desc"><strong>Action Items Only</strong><br>Just the discrepancies that need attention — a clean list to hand off for corrections.</div>', unsafe_allow_html=True)
                            st.download_button(
                                "⬇  Download Action Items (.xlsx)",
                                data=_excel_actions,
                                file_name=f"PayMatch_{safe_name}_{date_str}_ActionItems.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key="dl_actions")

                except Exception as e:
                    err_str = str(e).lower()
                    if "no sheet" in err_str or "worksheet" in err_str:
                        st.error("Couldn't read one of the files — it may be password-protected or have an unexpected layout. Try exporting a fresh copy from Paycor or your broker.")
                    elif "codec" in err_str or "decode" in err_str or "encoding" in err_str:
                        st.error("One of the files couldn't be opened correctly. If it's a CSV file, try opening it in Excel and saving it again before uploading.")
                    elif "column" in err_str or "key" in err_str:
                        st.error("A required column couldn't be found. Make sure the Deduction Report is a standard Paycor export and the enrollment file has employee names in columns labeled 'First Name' and 'Last Name'.")
                    else:
                        st.error(f"Something went wrong while processing: {e}")
                    st.info("If the problem persists, double-check that the Deduction Report is from Paycor and the Master Mapping Report is your broker's enrollment file with benefit labels in the second row.")

        # ── TREND VIEW ────────────────────────────────────────────
        if is_guest and client_name.strip():
            st.markdown("<div style='height:0.75rem'></div>", unsafe_allow_html=True)
            st.markdown("""
            <div style="background:linear-gradient(135deg,#FDF8EC,#FDF4F0);border:1px solid #E8C97A;border-radius:12px;padding:1.25rem 1.5rem;text-align:center;">
                <div style="font-size:0.88rem;font-weight:700;color:#5C1020;margin-bottom:0.35rem;">Want to track this over time?</div>
                <div style="font-size:0.8rem;color:#8B6914;line-height:1.65;">Create a free account to save run history, view month-over-month trends, add notes, and download reports at any time.</div>
            </div>
            """, unsafe_allow_html=True)
        if not is_guest and client_name.strip():
            st.markdown("<div style='height:1rem'></div>", unsafe_allow_html=True)
            history_df = load_history(client_name.strip())

            if len(history_df) > 1:
                st.markdown('<div class="trend-card">', unsafe_allow_html=True)
                st.markdown('<div class="slabel">Month-over-Month Trend</div>', unsafe_allow_html=True)
                st.markdown('<div class="sdesc">How discrepancies and uncollected dollars have changed across reconciliation runs for this client.</div>', unsafe_allow_html=True)

                t1, t2 = st.columns(2)
                with t1:
                    st.markdown("**Discrepancies by month**")
                    chart_data = history_df.sort_values("Period").set_index("Period")[["Add","Change","Review"]]
                    st.bar_chart(chart_data, color=["#3B82F6","#F97316","#FBBF24"])
                with t2:
                    st.markdown("**Monthly $ at stake**")
                    mo_data = history_df.sort_values("Period").set_index("Period")[["Monthly $ at stake"]]
                    st.line_chart(mo_data, color=["#2563EB"])

                st.markdown("**Run history**")
                dh = history_df[["Period","Run Date","Total Lines","OK","Add","Change","Review","Monthly $ at stake"]].copy()
                dh["Monthly $ at stake"] = dh["Monthly $ at stake"].apply(lambda x: f"${x:,.2f}")
                st.dataframe(dh, use_container_width=True, hide_index=True)

                if len(history_df) >= 2:
                    latest = history_df.iloc[0]; prev = history_df.iloc[1]
                    dc = int(latest["Discrepancies"]) - int(prev["Discrepancies"])
                    mc = round(float(latest["Monthly $ at stake"]) - float(prev["Monthly $ at stake"]), 2)
                    cls = "delta-dn" if dc < 0 else "delta-up"
                    st.markdown(f"""
                    <div class="delta {cls}">
                        Discrepancies {"↓" if dc<0 else "↑"} {abs(dc)} vs previous run
                        &nbsp;·&nbsp; Monthly $ {"↓" if mc<0 else "↑"} ${abs(mc):,.2f}
                    </div>
                    """, unsafe_allow_html=True)
                st.markdown('</div>', unsafe_allow_html=True)

            elif len(history_df) == 1:
                st.markdown(f'<div style="background:white;border-radius:12px;padding:1rem 1.5rem;box-shadow:0 1px 3px rgba(26,2,8,0.04),0 4px 18px rgba(26,2,8,0.06);border:1px solid #EADDD8;color:#B09898;font-size:0.84rem;">Trend view will appear after your second reconciliation run for <strong style="color:#7D5A5E">{client_name.strip()}</strong>.</div>', unsafe_allow_html=True)

            # ── NOTES ─────────────────────────────────────────────
            if len(history_df) > 0:
                st.markdown("<div style='height:0.5rem'></div>", unsafe_allow_html=True)
                st.markdown('<div class="notes-wrap">', unsafe_allow_html=True)
                st.markdown('<div class="notes-label">Run Notes</div>', unsafe_allow_html=True)
                st.markdown(f'<div style="font-size:0.82rem;color:#7D5A5E;margin-bottom:0.8rem;">Notes for <strong style="color:#1A0A0F;">{client_name.strip()}</strong> &nbsp;·&nbsp; {period}. These are saved automatically with this run in your history.</div>', unsafe_allow_html=True)
                _nk = f"rn_{''.join(c if c.isalnum() else '_' for c in client_name.strip())}_{period.replace(' ','_')}"
                if _nk not in st.session_state:
                    st.session_state[_nk] = get_run_notes(client_name.strip(), period)
                st.text_area("Notes", key=_nk, placeholder="e.g. Sent corrections to payroll team on 6/10. 3 employees pending broker re-enrollment.", label_visibility="collapsed", height=100)
                if st.button("Save Notes", key="save_run_notes"):
                    update_run_notes(client_name.strip(), period, st.session_state[_nk])
                    st.success("Notes saved.")
                st.markdown('</div>', unsafe_allow_html=True)

    # ═════════════════════════════════════════════════════════════
    # PAGE: PROFILE
    # ═════════════════════════════════════════════════════════════
    elif st.session_state.page == "profile":

        st.markdown("<div style='height:0.5rem'></div>", unsafe_allow_html=True)

        # Account info card
        st.markdown('<div class="slabel">Account</div>', unsafe_allow_html=True)
        st.markdown('<div class="profile-card">', unsafe_allow_html=True)

        a_col, info_col = st.columns([1, 6])
        with a_col:
            st.markdown(f'<div class="profile-avatar">{initial}</div>', unsafe_allow_html=True)
        with info_col:
            created = user.get("created_at","")
            try:
                since = datetime.fromisoformat(created).strftime("%B %-d, %Y")
            except Exception:
                since = "Unknown"
            st.markdown(f"""
            <div class="profile-name">{user["name"]}</div>
            <div class="profile-email">{user["email"]}</div>
            <div class="profile-since">Member since {since}</div>
            """, unsafe_allow_html=True)

        st.markdown("<div style='height:1rem'></div>", unsafe_allow_html=True)

        # Stats row
        all_hist = load_history()
        total_runs = len(all_hist)
        total_lines = int(all_hist["Total Lines"].sum()) if total_runs > 0 else 0
        total_disc  = int(all_hist["Discrepancies"].sum()) if total_runs > 0 else 0
        s1, s2, s3 = st.columns(3)
        with s1:
            st.markdown(f'<div class="profile-stat"><div class="profile-stat-val">{total_runs}</div><div class="profile-stat-lbl">Total Runs</div></div>', unsafe_allow_html=True)
        with s2:
            st.markdown(f'<div class="profile-stat"><div class="profile-stat-val">{total_lines:,}</div><div class="profile-stat-lbl">Lines Processed</div></div>', unsafe_allow_html=True)
        with s3:
            st.markdown(f'<div class="profile-stat"><div class="profile-stat-val">{total_disc:,}</div><div class="profile-stat-lbl">Discrepancies Found</div></div>', unsafe_allow_html=True)

        st.markdown('</div>', unsafe_allow_html=True)

        # Past runs
        st.markdown("<div style='height:0.5rem'></div>", unsafe_allow_html=True)
        st.markdown('<div class="slabel">Past Reconciliation Runs</div>', unsafe_allow_html=True)
        st.markdown('<div class="sdesc">Every reconciliation run saved in this workspace — newest first. Click a status badge to cycle it, and download reports directly from this view.</div>', unsafe_allow_html=True)

        if total_runs > 0:
            # ── Search bar ────────────────────────────────────────
            search_q = st.text_input(
                "Search runs",
                placeholder="Search by client name or period (e.g. CASA or June 2026)…",
                key="hist_search",
                label_visibility="collapsed"
            )
            if search_q.strip():
                q = search_q.strip().lower()
                display_hist = all_hist[
                    all_hist["Client"].astype(str).str.lower().str.contains(q, na=False) |
                    all_hist["Period"].astype(str).str.lower().str.contains(q, na=False)
                ].reset_index(drop=True)
                if len(display_hist) == 0:
                    st.markdown(f'<div class="empty-prompt">No runs match <strong>"{search_q}"</strong>.</div>', unsafe_allow_html=True)
                    display_hist = None
            else:
                display_hist = all_hist.reset_index(drop=True)

            if display_hist is not None and len(display_hist) > 0:
                # Column headers
                _COLS = [1.3, 1.8, 1.3, 0.7, 0.9, 0.6, 0.6, 0.6, 2.1, 1.5, 0.7]
                _HDRS = ["Date","Client","Period","Lines","$/mo","OK","Add","Chg","Status","Report",""]
                hc = st.columns(_COLS)
                for _hdr, _col in zip(_HDRS, hc):
                    _col.markdown(f'<div style="font-size:0.66rem;font-weight:700;color:#9D7075;text-transform:uppercase;letter-spacing:0.1em;padding-bottom:0.35rem;">{_hdr}</div>', unsafe_allow_html=True)
                st.markdown('<hr style="margin:0 0 0.4rem;border-color:#E2E8F0;">', unsafe_allow_html=True)

                _SBADGE = {
                    "Incomplete":  ("sbadge-incomplete",  "●"),
                    "In Progress": ("sbadge-in-progress", "●"),
                    "Complete":    ("sbadge-complete",    "●"),
                }

                for i, row in display_hist.iterrows():
                    run_date   = str(row.get("Run Date",""))
                    client_h   = str(row.get("Client",""))
                    period_h   = str(row.get("Period",""))
                    lines_h    = int(row.get("Total Lines", 0))
                    mo_h       = float(row.get("Monthly $ at stake", 0))
                    ok_h       = int(row.get("OK", 0))
                    add_h      = int(row.get("Add", 0))
                    chg_h      = int(row.get("Change", 0))
                    report_f   = str(row.get("Report File",""))
                    cur_status = str(row.get("Status","Incomplete"))
                    if cur_status not in STATUS_OPTIONS: cur_status = "Incomplete"
                    del_key    = f"{run_date}__{client_h}__{period_h}"

                    # Resolve legacy relative report paths
                    if report_f and not os.path.isabs(report_f):
                        report_f = os.path.join(_APP_DIR, report_f)

                    c1,c2,c3,c4,c5,c6,c7,c8,c9,c10,c11 = st.columns(_COLS)
                    _cell = '<div style="font-size:0.82rem;color:#3D1A1E;padding:0.28rem 0 0.1rem;">'
                    c1.markdown(f'{_cell}{run_date}</div>', unsafe_allow_html=True)
                    c2.markdown(f'{_cell}<strong style="color:#1A0A0F;">{client_h}</strong></div>', unsafe_allow_html=True)
                    c3.markdown(f'{_cell}{period_h}</div>', unsafe_allow_html=True)
                    c4.markdown(f'{_cell}{lines_h:,}</div>', unsafe_allow_html=True)
                    c5.markdown(f'{_cell}${mo_h:,.2f}</div>', unsafe_allow_html=True)
                    c6.markdown(f'<div style="font-size:0.82rem;color:#059669;font-weight:700;padding:0.28rem 0 0.1rem;">{ok_h}</div>', unsafe_allow_html=True)
                    c7.markdown(f'<div style="font-size:0.82rem;color:{"#8B1A2F" if add_h>0 else "#C0A8A8"};font-weight:{"700" if add_h>0 else "400"};padding:0.28rem 0 0.1rem;">{add_h}</div>', unsafe_allow_html=True)
                    c8.markdown(f'<div style="font-size:0.82rem;color:{"#EA580C" if chg_h>0 else "#C0A8A8"};font-weight:{"700" if chg_h>0 else "400"};padding:0.28rem 0 0.1rem;">{chg_h}</div>', unsafe_allow_html=True)

                    # Status: clickable badge that cycles Incomplete → In Progress → Complete
                    with c9:
                        _scycle_cls = {"Incomplete":"scycle-incomplete","In Progress":"scycle-in-progress","Complete":"scycle-complete"}.get(cur_status,"scycle-incomplete")
                        st.markdown(f'<span class="{_scycle_cls}" style="display:none"></span>', unsafe_allow_html=True)
                        if st.button(f"● {cur_status}", key=f"status_btn_{i}", help="Click to change status"):
                            _next = STATUS_OPTIONS[(STATUS_OPTIONS.index(cur_status)+1) % len(STATUS_OPTIONS)]
                            update_run_status(run_date, client_h, period_h, _next)
                            st.rerun()

                    # Report download
                    with c10:
                        if report_f and os.path.exists(report_f):
                            try:
                                with open(report_f, 'rb') as _f:
                                    _bytes = _f.read()
                                safe_c = client_h.replace(' ','_')
                                st.download_button(
                                    "⬇ Report",
                                    data=_bytes,
                                    file_name=f"PayMatch_{safe_c}_{period_h.replace(' ','_')}.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    key=f"hist_dl_{i}"
                                )
                            except Exception:
                                st.markdown('<span style="color:#CBD5E1;font-size:0.72rem;">Read error</span>', unsafe_allow_html=True)
                        else:
                            st.markdown('<span style="color:#CBD5E1;font-size:0.72rem;" title="Re-run this reconciliation to save the report">Re-run to save</span>', unsafe_allow_html=True)

                    # Delete button
                    with c11:
                        if st.session_state.delete_confirm_key == del_key:
                            # Already in confirmation — show nothing here; banner below handles it
                            st.markdown('<span style="font-size:0.75rem;color:#DC2626;font-weight:600;">↓ confirm</span>', unsafe_allow_html=True)
                        else:
                            if st.button("🗑", key=f"del_{i}", help="Delete this run"):
                                st.session_state.delete_confirm_key = del_key
                                st.rerun()

                    # Confirmation banner (full width, shown below the row)
                    if st.session_state.delete_confirm_key == del_key:
                        st.markdown(f'<div class="del-confirm">Delete <strong>{client_h} · {period_h}</strong>? This will remove the run from history and delete the saved report file. <strong>This cannot be undone.</strong></div>', unsafe_allow_html=True)
                        yes_col, no_col, _ = st.columns([1, 1, 8])
                        with yes_col:
                            if st.button("Yes, delete", key=f"confirm_del_{i}", type="primary"):
                                delete_run(run_date, client_h, period_h, report_f)
                                st.session_state.delete_confirm_key = None
                                st.rerun()
                        with no_col:
                            if st.button("Cancel", key=f"cancel_del_{i}"):
                                st.session_state.delete_confirm_key = None
                                st.rerun()

                    # Notes preview under the row
                    notes_h = str(row.get("Notes",""))
                    if notes_h and notes_h.lower() not in ("nan","none",""):
                        st.markdown(f'<div class="hist-note"><span>📝</span><span>{notes_h}</span></div>', unsafe_allow_html=True)

                    st.markdown('<hr style="margin:0.2rem 0;border-color:#F1F5F9;">', unsafe_allow_html=True)
        else:
            st.markdown('<div class="empty-prompt">No reconciliation runs yet. Go to <strong>Reconciliation</strong> to run your first one.</div>', unsafe_allow_html=True)

        # Change password
        st.markdown("<div style='height:0.5rem'></div>", unsafe_allow_html=True)
        st.markdown('<div class="slabel">Security</div>', unsafe_allow_html=True)
        with st.expander("Change Password"):
            st.markdown('<div style="font-size:0.84rem;color:#7D5A5E;margin-bottom:0.85rem;">Enter your current password to confirm your identity, then choose a new one.</div>', unsafe_allow_html=True)
            cp_current = st.text_input("Current Password", type="password", key="cp_current")
            cp_new     = st.text_input("New Password", type="password", key="cp_new")
            if cp_new:
                checks = pw_checks(cp_new)
                rows_html = ""
                for key_c, label in [("length","At least 8 characters"),("upper","Uppercase"),("lower","Lowercase"),("digit","Number"),("special","Special character")]:
                    ok = checks[key_c]
                    rows_html += f'<div class="pw-req"><span class="{"pr-ok" if ok else "pr-fail"}">{"✓" if ok else "✗"}</span><span style="color:{"#374151" if ok else "#6B7280"};font-size:0.76rem;">{label}</span></div>'
                st.markdown(f'<div class="pw-reqs">{rows_html}</div>', unsafe_allow_html=True)
            cp_new2 = st.text_input("Confirm New Password", type="password", key="cp_new2")
            if st.button("Update Password", key="do_cp"):
                computed, _ = hash_pw(cp_current, user["salt"])
                if computed != user["password_hash"]:
                    st.error("Current password is incorrect.")
                elif not pw_ok(cp_new):
                    st.error("New password doesn't meet all requirements.")
                elif cp_new != cp_new2:
                    st.error("New passwords don't match.")
                else:
                    update_password(user["email"], cp_new)
                    updated_user = get_user(user["email"])
                    st.session_state.current_user = updated_user
                    st.success("Password updated successfully.")

        # Logout from profile
        st.markdown("<div style='height:1rem'></div>", unsafe_allow_html=True)
        _, lo_col, _ = st.columns([2, 1, 2])
        with lo_col:
            if st.button("Sign Out", key="signout_profile"):
                for k in ["logged_in","current_user","recon_df","last_run_ts","page","guest_mode"]:
                    st.session_state[k] = (False if k in ("logged_in","guest_mode") else (None if k!="page" else "app"))
                st.rerun()

    # ── FOOTER ────────────────────────────────────────────────────
    st.markdown("<div style='height:1.5rem'></div>", unsafe_allow_html=True)
    st.markdown(f"""
    <div class="pm-footer">
        <strong>PayMatch</strong> by ProPayHR &nbsp;·&nbsp; Benefits Reconciliation Platform &nbsp;·&nbsp; {datetime.now().year}
    </div>
    <script>
    (function(){{
        if(!window.__pmAlive){{
            window.__pmAlive = true;
            setInterval(function(){{
                try{{fetch(window.location.origin + window.location.pathname, {{method:'HEAD', mode:'no-cors', cache:'no-store'}})}}catch(e){{}}
            }}, 45000);
        }}
    }})();
    </script>
    """, unsafe_allow_html=True)
