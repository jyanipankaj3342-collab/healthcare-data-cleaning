"""
Healthcare Data Cleaning & Understanding
Input : healthcare_patients_v2.csv  (synthetic, de-identified)
Output: healthcare_patients_clean.csv  +  healthcare_analysis.xlsx
"""
import sys
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import CellIsRule
from openpyxl.utils import get_column_letter

SRC = sys.argv[1] if len(sys.argv) > 1 else "healthcare_patients_v2.csv"
OUT_CSV, OUT_XLSX = "healthcare_patients_clean.csv", "healthcare_analysis.xlsx"

# ---------- 1. LOAD ----------
raw = pd.read_csv(SRC, dtype=str)
df = raw.copy()
log = []  # (step, check, found, action)

# ---------- 2. IDENTIFY ISSUES ----------
df = df.apply(lambda s: s.str.strip())                      # stray whitespace
log.append(("Identify", "Rows / columns loaded", f"{len(df)} rows x {df.shape[1]} cols", "None"))
miss = int(df.isna().sum().sum()) + int((df == "").sum().sum())
log.append(("Identify", "Missing values", str(miss), "None needed" if miss == 0 else "Rows dropped"))
dup_id = int(df["PatientID"].duplicated().sum())
log.append(("Identify", "Duplicate PatientIDs", str(dup_id), "None needed" if dup_id == 0 else "Kept first"))
bad_fmt = int((~df["PatientID"].str.match(r"^SYN-PAT-\d{5}$")).sum())
log.append(("Identify", "PatientID format (SYN-PAT-#####)", f"{bad_fmt} invalid", "None needed"))

# ---------- 3. TYPE CONVERSIONS / STANDARDISE ----------
num_cols = ["Age", "ComorbidityCount", "PriorAdmissions12M", "LengthOfStayDays"]
for c in num_cols:
    df[c] = pd.to_numeric(df[c], errors="coerce").astype("Int64")
cat_cols = ["Gender", "Diagnosis", "PrimaryMedication", "DischargeDisposition"]
for c in cat_cols:
    df[c] = df[c].str.title().str.replace("Copd", "COPD").astype("category")
for c in ["FollowUpBooked", "Readmitted30Days"]:
    df[c] = df[c].str.capitalize()
log.append(("Convert", "Numeric columns -> integer", ", ".join(num_cols), "Converted"))
log.append(("Convert", "Text columns -> category; Yes/No standardised", "6 columns", "Converted"))

# ---------- 4. RANGE VALIDATION ----------
rules = {"Age": (0, 120), "ComorbidityCount": (0, 20),
         "PriorAdmissions12M": (0, 50), "LengthOfStayDays": (1, 365)}
for c, (lo, hi) in rules.items():
    bad = int(((df[c] < lo) | (df[c] > hi)).sum())
    log.append(("Validate", f"{c} within {lo}-{hi}", f"{bad} out of range",
                "None needed" if bad == 0 else "Removed"))
    df = df[(df[c] >= lo) & (df[c] <= hi)]
for v in ["FollowUpBooked", "Readmitted30Days"]:
    assert set(df[v]) <= {"Yes", "No"}
mism = df.groupby("Diagnosis", observed=True)["PrimaryMedication"].nunique().max()
log.append(("Validate", "Each diagnosis maps to one medication", f"max {mism} per diagnosis", "Consistent"))
df = df.dropna().drop_duplicates("PatientID").reset_index(drop=True)

# ---------- 5. SAVE CLEAN CSV ----------
df.to_csv(OUT_CSV, index=False)
n = len(df)
print(f"Saved {OUT_CSV}: {n} rows")
print(df.describe(include="all").T)

# ---------- 6. COLOURFUL EXCEL REPORT ----------
INDIGO, TEAL, PINK, AMBER = "4F46E5", "0D9488", "DB2777", "F59E0B"
LAV, MINT, ROSE = "EEF2FF", "CCFBF1", "FCE7F3"
F = "Arial"
thin = Side(style="thin", color="D4D4D8")
box = Border(left=thin, right=thin, top=thin, bottom=thin)

def hdr(ws, row, c1, c2, color):
    for c in range(c1, c2 + 1):
        x = ws.cell(row, c)
        x.font = Font(name=F, bold=True, color="FFFFFF")
        x.fill = PatternFill("solid", fgColor=color)
        x.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        x.border = box

def title(ws, text, sub, width):
    ws.sheet_view.showGridLines = False
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=width)
    ws["A1"] = text
    ws["A1"].font = Font(name=F, size=20, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor=INDIGO)
    ws["A1"].alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[1].height = 38
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=width)
    ws["A2"] = sub
    ws["A2"].font = Font(name=F, italic=True, color="52525B")
    ws["A2"].fill = PatternFill("solid", fgColor=LAV)
    ws["A2"].alignment = Alignment(indent=1)

wb = Workbook()

# --- Clean Data sheet (built first so formulas can reference it) ---
wd = wb.active
wd.title = "Clean Data"
cols = list(df.columns)
wd.append(cols)
for r in df.itertuples(index=False):
    wd.append([int(v) if isinstance(v, (int,)) or str(v).isdigit() and False else (v.item() if hasattr(v, "item") else v) for v in r])
hdr(wd, 1, 1, len(cols), INDIGO)
wd.row_dimensions[1].height = 30
for i, c in enumerate(cols, 1):
    wd.column_dimensions[get_column_letter(i)].width = max(14, len(c) + 4)
    for row in range(2, n + 2):
        cell = wd.cell(row, i)
        cell.font = Font(name=F, size=10)
        if row % 2 == 0:
            cell.fill = PatternFill("solid", fgColor="F5F3FF")
wd.freeze_panes = "B2"
wd.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{n+1}"
last = n + 1
wd.conditional_formatting.add(f"K2:K{last}", CellIsRule(operator="equal", formula=['"Yes"'],
    fill=PatternFill("solid", bgColor="FECACA"), font=Font(name=F, color="B91C1C", bold=True)))
wd.conditional_formatting.add(f"K2:K{last}", CellIsRule(operator="equal", formula=['"No"'],
    fill=PatternFill("solid", bgColor="BBF7D0"), font=Font(name=F, color="15803D")))
wd.conditional_formatting.add(f"J2:J{last}", CellIsRule(operator="equal", formula=['"No"'],
    fill=PatternFill("solid", bgColor="FEF3C7"), font=Font(name=F, color="B45309")))

def rng(col): return f"'Clean Data'!${col}$2:${col}${last}"
COL = {c: get_column_letter(i + 1) for i, c in enumerate(cols)}

# --- Summary sheet ---
ws = wb.create_sheet("Summary", 0)
title(ws, "Healthcare Patient Data - Summary", "Synthetic, de-identified learning data - not for clinical decisions", 8)
kpis = [("Patients", f"=COUNTA({rng('A')})", "0", INDIGO),
        ("Avg Age", f"=AVERAGE({rng(COL['Age'])})", "0.0", TEAL),
        ("Avg Stay (days)", f"=AVERAGE({rng(COL['LengthOfStayDays'])})", "0.0", AMBER),
        ("30-Day Readmission", f'=COUNTIF({rng(COL["Readmitted30Days"])},"Yes")/COUNTA({rng("A")})', "0.0%", PINK)]
for i, (lab, f, nf, colr) in enumerate(kpis):
    c = 1 + i * 2
    ws.merge_cells(start_row=4, start_column=c, end_row=4, end_column=c + 1)
    ws.merge_cells(start_row=5, start_column=c, end_row=5, end_column=c + 1)
    ws.cell(4, c, lab).font = Font(name=F, bold=True, color="FFFFFF", size=11)
    ws.cell(5, c, f).font = Font(name=F, bold=True, size=22, color=colr)
    ws.cell(5, c).number_format = nf
    for r in (4, 5):
        for cc in (c, c + 1):
            ws.cell(r, cc).alignment = Alignment(horizontal="center", vertical="center")
            ws.cell(r, cc).fill = PatternFill("solid", fgColor=colr if r == 4 else "FAFAFA")
    ws.row_dimensions[5].height = 42

ws["A7"] = "Readmission by Diagnosis"
ws["A7"].font = Font(name=F, bold=True, size=13, color=INDIGO)
heads = ["Diagnosis", "Patients", "Readmitted", "Rate", "Avg Stay", "Avg Comorbid."]
for i, h in enumerate(heads, 1): ws.cell(8, i, h)
hdr(ws, 8, 1, 6, INDIGO)
diags = sorted(df["Diagnosis"].unique())
for r, d in enumerate(diags, 9):
    ws.cell(r, 1, d)
    ws.cell(r, 2, f'=COUNTIF({rng(COL["Diagnosis"])},A{r})')
    ws.cell(r, 3, f'=COUNTIFS({rng(COL["Diagnosis"])},A{r},{rng(COL["Readmitted30Days"])},"Yes")')
    ws.cell(r, 4, f"=IF(B{r}=0,0,C{r}/B{r})").number_format = "0.0%"
    ws.cell(r, 5, f'=AVERAGEIF({rng(COL["Diagnosis"])},A{r},{rng(COL["LengthOfStayDays"])})').number_format = "0.0"
    ws.cell(r, 6, f'=AVERAGEIF({rng(COL["Diagnosis"])},A{r},{rng(COL["ComorbidityCount"])})').number_format = "0.0"
    for c in range(1, 7):
        x = ws.cell(r, c); x.font = Font(name=F); x.border = box
        x.fill = PatternFill("solid", fgColor=LAV if r % 2 else "FFFFFF")
d_end = 8 + len(diags)

ws["A16"] = "Discharge Disposition"
ws["A16"].font = Font(name=F, bold=True, size=13, color=TEAL)
for i, h in enumerate(["Disposition", "Patients", "Share"], 1): ws.cell(17, i, h)
hdr(ws, 17, 1, 3, TEAL)
disp = sorted(df["DischargeDisposition"].unique())
for r, d in enumerate(disp, 18):
    ws.cell(r, 1, d)
    ws.cell(r, 2, f'=COUNTIF({rng(COL["DischargeDisposition"])},A{r})')
    ws.cell(r, 3, f"=B{r}/SUM($B$18:$B${17+len(disp)})").number_format = "0.0%"
    for c in range(1, 4):
        x = ws.cell(r, c); x.font = Font(name=F); x.border = box
        x.fill = PatternFill("solid", fgColor=MINT if r % 2 else "FFFFFF")
p_end = 17 + len(disp)
for col, w in zip("ABCDEFGH", [22, 12, 12, 12, 12, 14, 14, 14]): ws.column_dimensions[col].width = w

bar = BarChart(); bar.type = "col"; bar.title = "30-Day Readmission Rate by Diagnosis"
bar.add_data(Reference(ws, min_col=4, min_row=8, max_row=d_end), titles_from_data=True)
bar.set_categories(Reference(ws, min_col=1, min_row=9, max_row=d_end))
bar.series[0].graphicalProperties.solidFill = PINK
bar.legend = None; bar.y_axis.numFmt = "0%"; bar.y_axis.majorGridlines = None
bar.dataLabels = DataLabelList(); bar.dataLabels.showVal = True
bar.roundedCorners = False; bar.x_axis.delete = False; bar.y_axis.delete = False
bar.height, bar.width = 7.5, 15
ws.add_chart(bar, "H7")

pie = PieChart(); pie.title = "Discharge Disposition"
pie.add_data(Reference(ws, min_col=2, min_row=17, max_row=p_end), titles_from_data=True)
pie.set_categories(Reference(ws, min_col=1, min_row=18, max_row=p_end))
from openpyxl.chart.series import DataPoint
for i, colr in enumerate([TEAL, AMBER, INDIGO]):
    pt = DataPoint(idx=i); pt.graphicalProperties.solidFill = colr; pie.series[0].dPt.append(pt)
pie.dataLabels = DataLabelList(); pie.dataLabels.showPercent = True; pie.dataLabels.showCatName = False; pie.dataLabels.showSerName = False; pie.dataLabels.showVal = False; pie.dataLabels.showLegendKey = False
pie.legend.overlay = False; pie.roundedCorners = False
pie.height, pie.width = 7.5, 10
ws.add_chart(pie, "D16")

# --- Statistics sheet ---
st = wb.create_sheet("Statistics", 1)
title(st, "Descriptive Statistics", "Numerical attributes (live formulas) and categorical attributes", 9)
for i, h in enumerate(["Metric"] + num_cols, 1): st.cell(4, i, h)
hdr(st, 4, 1, 5, TEAL)
fns = [("Count", "COUNT"), ("Mean", "AVERAGE"), ("Median", "MEDIAN"), ("Std Dev", "STDEV"),
       ("Min", "MIN"), ("Max", "MAX")]
for r, (lab, fn) in enumerate(fns, 5):
    st.cell(r, 1, lab)
    for j, c in enumerate(num_cols, 2):
        x = st.cell(r, j, f"={fn}({rng(COL[c])})")
        x.number_format = "0" if fn in ("COUNT", "MIN", "MAX") else "0.00"
    for c in range(1, 6):
        x = st.cell(r, c); x.font = Font(name=F, bold=(c == 1)); x.border = box
        x.fill = PatternFill("solid", fgColor=MINT if r % 2 else "FFFFFF")
for r, (lab, fn, arg) in enumerate([("25th pct", "QUARTILE", 1), ("75th pct", "QUARTILE", 3)], 11):
    st.cell(r, 1, lab)
    for j, c in enumerate(num_cols, 2):
        st.cell(r, j, f"=QUARTILE({rng(COL[c])},{arg})").number_format = "0.00"
    for c in range(1, 6):
        x = st.cell(r, c); x.font = Font(name=F, bold=(c == 1)); x.border = box
        x.fill = PatternFill("solid", fgColor=MINT if r % 2 else "FFFFFF")

row = 15
for cat, colr, fill in [("Gender", INDIGO, LAV), ("Diagnosis", PINK, ROSE), ("PrimaryMedication", AMBER, "FEF3C7"),
                        ("FollowUpBooked", TEAL, MINT), ("Readmitted30Days", INDIGO, LAV)]:
    for i, h in enumerate([cat, "Count", "Share"], 1): st.cell(row, i, h)
    hdr(st, row, 1, 3, colr)
    vals = sorted(df[cat].unique())
    for k, v in enumerate(vals, row + 1):
        st.cell(k, 1, v)
        st.cell(k, 2, f'=COUNTIF({rng(COL[cat])},A{k})')
        st.cell(k, 3, f"=B{k}/COUNTA({rng('A')})").number_format = "0.0%"
        for c in range(1, 4):
            x = st.cell(k, c); x.font = Font(name=F); x.border = box
            x.fill = PatternFill("solid", fgColor=fill)
    row += len(vals) + 3
for col, w in zip("ABCDE", [22, 14, 20, 22, 22]): st.column_dimensions[col].width = w

# --- Cleaning Log sheet ---
lg = wb.create_sheet("Cleaning Log")
title(lg, "Data Quality & Cleaning Log", "Checks performed by healthcare_cleaning.py (results written as values)", 4)
for i, h in enumerate(["Step", "Check", "Finding", "Action"], 1): lg.cell(4, i, h)
hdr(lg, 4, 1, 4, PINK)
stepcol = {"Identify": "DBEAFE", "Convert": "FEF3C7", "Validate": "D1FAE5"}
for r, rec in enumerate(log, 5):
    for c, v in enumerate(rec, 1):
        x = lg.cell(r, c, v); x.font = Font(name=F); x.border = box
        x.alignment = Alignment(wrap_text=True, vertical="center")
        x.fill = PatternFill("solid", fgColor=stepcol[rec[0]])
for col, w in zip("ABCD", [12, 42, 28, 22]): lg.column_dimensions[col].width = w

wb.save(OUT_XLSX)
print("Saved", OUT_XLSX)
