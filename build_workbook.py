"""Write the Excel scorecard and loan-pricing workbook from outputs/model.json.

    python build_workbook.py

Every pricing and risk number in the workbook is a live formula. Only the fitted
model parameters (WoE values, coefficients) and historical statistics are typed in,
and each of those is marked in blue with its source.
"""
from __future__ import annotations

import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.marker import Marker
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))
from creditrisk.model_io import load_model  # noqa: E402

OUT_PATH = ROOT / "excel" / "Credit_Risk_Pricing_Model.xlsx"

# ---------------------------------------------------------------- styles
FONT = "Arial"
BLUE, BLACK, GREEN, WHITE = "0000FF", "000000", "008000", "FFFFFF"
NAVY, SLATE, PALE, LINE = "1F3A5F", "44546A", "EEF2F7", "C9D1DC"
F_IN = Font(name=FONT, size=10, color=BLUE)
F_CALC = Font(name=FONT, size=10, color=BLACK)
F_LINK = Font(name=FONT, size=10, color=GREEN)
F_BOLD = Font(name=FONT, size=10, bold=True)
F_NOTE = Font(name=FONT, size=9, italic=True, color="595959")
F_TITLE = Font(name=FONT, size=16, bold=True, color=NAVY)
F_SUB = Font(name=FONT, size=10, color="404040")
F_HEAD = Font(name=FONT, size=10, bold=True, color=WHITE)
F_SECTION = Font(name=FONT, size=11, bold=True, color=NAVY)
FILL_HEAD = PatternFill("solid", fgColor=NAVY)
FILL_INPUT = PatternFill("solid", fgColor="FFFF00")
FILL_KEY = PatternFill("solid", fgColor=PALE)
THIN = Side(style="thin", color=LINE)
BOX = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)
TOP_RULE = Border(top=Side(style="thin", color=BLACK))
WRAP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(horizontal="center", vertical="center")

FMT_USD = '$#,##0;($#,##0);"-"'
FMT_USD2 = '$#,##0.00;($#,##0.00);"-"'
FMT_PCT1 = '0.0%;(0.0%);"-"'
FMT_PCT2 = '0.00%;(0.00%);"-"'
FMT_PCT4 = '0.0000%'
FMT_INT = '#,##0;(#,##0);"-"'
FMT_NUM2 = '0.00;(0.00);"-"'
FMT_NUM4 = '0.0000;(0.0000);"-"'
FMT_DEC = '0.000000'


def style(c, font=F_CALC, fmt=None, fill=None, align=None, border=None):
    c.font = font
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = fill
    if align:
        c.alignment = align
    if border:
        c.border = border
    return c


def put(ws, ref, value, font=F_CALC, fmt=None, fill=None, align=None, border=None):
    ws[ref] = value
    return style(ws[ref], font, fmt, fill, align, border)


def header_row(ws, row, col, labels, widths=None):
    for j, lab in enumerate(labels):
        c = ws.cell(row=row, column=col + j, value=lab)
        style(c, F_HEAD, fill=FILL_HEAD, align=Alignment(horizontal="center", vertical="center", wrap_text=True))
    ws.row_dimensions[row].height = 30


def title(ws, text, sub=None):
    put(ws, "B1", text, F_TITLE)
    ws.row_dimensions[1].height = 26
    if sub:
        put(ws, "B2", sub, F_SUB)


def section(ws, ref, text):
    put(ws, ref, text, F_SECTION)


def name(wb, nm, sheet, ref):
    wb.defined_names[nm] = DefinedName(nm, attr_text=f"'{sheet}'!{ref}")


def q(sheet):
    return f"'{sheet}'"


# ---------------------------------------------------------------- display helpers
PCT_POINT_VARS = {"dti", "revol_util"}   # stored in the data as %-points; shown in Excel as fractions


def edge_value(var, e):
    return e / 100.0 if var["var"] in PCT_POINT_VARS else e


def edge_text(var, e):
    v = edge_value(var, e)
    if var["var"] == "term_months":
        return f"{int(v)}"
    if var["unit"] == "$":
        return f"${v:,.0f}"
    if var["unit"] == "pct" or var["var"] in PCT_POINT_VARS:
        return f"{v:.0%}"
    return f"{v:g}"


def numeric_labels(var):
    finite = [b["upper_edge"] for b in var["bins"] if not b["missing"] and b["upper_edge"] is not None]
    labels = []
    if var["var"] == "term_months":
        return ["36 months", "60 months"] + (["Missing"] if any(b["missing"] for b in var["bins"]) else [])
    for i, b in enumerate(var["bins"]):
        if b["missing"]:
            labels.append("Not reported")
        elif i == 0:
            labels.append(f"Up to {edge_text(var, finite[0])}")
        elif b["upper_edge"] is None:
            labels.append(f"Over {edge_text(var, finite[-1])}")
        else:
            labels.append(f"Over {edge_text(var, finite[i - 1])} to {edge_text(var, finite[i])}")
    return labels


INPUT_FORMAT = {"loan_amnt": FMT_USD, "annual_inc": FMT_USD, "term_months": '0', "dti": FMT_PCT1,
                "revol_util": FMT_PCT1, "loan_to_income": FMT_PCT1}
EDGE_FORMAT = {"loan_amnt": FMT_USD, "annual_inc": FMT_USD, "term_months": '0', "dti": '0%',
               "revol_util": '0%', "loan_to_income": '0%'}


# ======================================================================== build
def build(model: dict) -> Path:
    wb = Workbook()
    ws_read = wb.active
    ws_read.title = "Read Me"
    ws_tool = wb.create_sheet("Pricing Tool")
    ws_as = wb.create_sheet("Assumptions")
    ws_sc = wb.create_sheet("Scorecard")
    ws_cf = wb.create_sheet("Cash Flows")
    ws_sv = wb.create_sheet("Solver")
    ws_rb = wb.create_sheet("Risk-Based Pricing")
    ws_mp = wb.create_sheet("Model Performance")
    for ws in wb.worksheets:
        ws.sheet_view.showGridLines = False
        ws.column_dimensions["A"].width = 2

    build_assumptions(wb, ws_as)
    sc_refs = build_scorecard(wb, ws_sc, model)
    tool = build_tool_inputs(wb, ws_tool, model, sc_refs)
    solver = build_solver(wb, ws_sv, model, tool)
    cf = build_cash_flows(wb, ws_cf, tool)
    bands = build_risk_bands(wb, ws_rb, model, solver)
    build_tool_outputs(wb, ws_tool, tool, solver, cf, bands)
    build_performance(wb, ws_mp, model)
    build_readme(wb, ws_read, model)

    wb.active = 1
    OUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUT_PATH)
    return OUT_PATH


# ------------------------------------------------------------------- Assumptions
ASSUMPTIONS = [
    # (row, name, label, value, fmt, note, source, key-input?)
    ("Lender economics", None),
    (6, "Hurdle_Rate", "Required annual return (hurdle rate)", 0.06, FMT_PCT1,
     "Lender's target return on money lent; used to discount all cash flows.", "Assumption", True),
    (7, "Servicing_Fee", "Servicing fee", 0.01, FMT_PCT1,
     "Share of each payment received that goes to the loan servicer.",
     "LendingClub charged investors a 1% service fee on payments received.", True),
    (8, "Upfront_Fee", "Upfront fee kept by lender", 0.0, FMT_PCT1,
     "Share of the loan amount the lender keeps at funding.",
     "Assumption. Investor view: LendingClub's origination fee went to the platform, not the investor.", True),
    (9, "LGD_Rate", "Loss given default (LGD)", 0.90, FMT_PCT1,
     "Share of the outstanding balance lost when a borrower defaults.",
     "Gallo, S. (2021), 'Fintech platforms: lax or careful borrowers' screening?', Financial Innovation 7:17. "
     "Mean LGD 90.2% across 291,664 defaulted LendingClub loans issued 2007-2018.", True),
    (10, "Recovery_Lag", "Recovery lag (months)", 6, '0',
     "Months between default and the recovery cash arriving.", "Assumption", True),
    ("Pricing policy", None),
    (13, "Target_Margin", "Target margin over break-even", 0.01, FMT_PCT2,
     "APR points added on top of the break-even rate.", "Assumption", True),
    (14, "Max_APR", "Maximum APR offered", 0.36, FMT_PCT1,
     "Applications whose risk-based APR exceeds this are declined.",
     "36% is a common consumer-loan rate cap (e.g. the Military Lending Act limit).", True),
    (15, "Max_PD", "Maximum lifetime default probability", 0.35, FMT_PCT1,
     "Applications with a higher predicted PD are declined.", "Assumption", True),
    ("Standard loan for band pricing", None),
    (18, "Band_Amount", "Loan amount", 10_000, FMT_USD,
     "Used on the Risk-Based Pricing sheet. Break-even APR does not depend on loan size.",
     "Median 36-month loan in the data.", False),
    (19, "Band_Term", "Term (months)", 36, '0', "Risk bands are built from 36-month loans.", "Data", False),
    ("Scorecard scaling", None),
    (22, "Score_Base", "Base score", 600, '0', "Score assigned to the base odds.", "Industry convention", False),
    (23, "Score_Odds", "Odds at base score (good : bad)", 50, '0', "Odds of a good loan at the base score.",
     "Industry convention", False),
    (24, "Score_PDO", "Points to double the odds (PDO)", 20, '0',
     "Every 20 points the odds of a good loan double.", "Industry convention", False),
]


def build_assumptions(wb, ws):
    title(ws, "Assumptions",
          "Blue cells on yellow are inputs. Change any of them and every sheet recalculates.")
    header_row(ws, 4, 2, ["Assumption", "Value", "What it means", "Source"])
    for spec in ASSUMPTIONS:
        if spec[1] is None:
            r = {"Lender economics": 5, "Pricing policy": 12, "Standard loan for band pricing": 17,
                 "Scorecard scaling": 21}[spec[0]]
            section(ws, f"B{r}", spec[0])
            continue
        r, nm, label, val, fmt, note, src, key = spec
        put(ws, f"B{r}", label)
        put(ws, f"C{r}", val, F_IN, fmt, FILL_INPUT if key else None, border=BOX)
        put(ws, f"D{r}", note, F_SUB, align=WRAP)
        put(ws, f"E{r}", src, F_NOTE, align=WRAP)
        name(wb, nm, "Assumptions", f"$C${r}")

    section(ws, "B26", "Calculated from the inputs above")
    calc = [
        (27, "Score_Factor", "Scaling factor = PDO / ln(2)", "=Score_PDO/LN(2)", FMT_NUM4),
        (28, "Score_Offset", "Scaling offset = Base score - Factor x ln(Odds)",
         "=Score_Base-Score_Factor*LN(Score_Odds)", FMT_NUM4),
        (29, "Monthly_Disc_Rate", "Monthly discount rate = (1 + hurdle)^(1/12) - 1", "=(1+Hurdle_Rate)^(1/12)-1", FMT_PCT4),
        (30, "Monthly_Disc_Factor", "Monthly discount factor = 1 / (1 + monthly rate)", "=1/(1+Monthly_Disc_Rate)", FMT_DEC),
        (31, "Required_APR", "Hurdle expressed as an APR = 12 x monthly rate", "=12*Monthly_Disc_Rate", FMT_PCT2),
        (32, "Lag_Disc_Factor", "Discount factor over the recovery lag", "=Monthly_Disc_Factor^Recovery_Lag", FMT_DEC),
    ]
    for r, nm, label, f, fmt in calc:
        put(ws, f"B{r}", label)
        put(ws, f"C{r}", f, F_CALC, fmt, border=BOX)
        name(wb, nm, "Assumptions", f"$C${r}")
    for col, w in zip("BCDE", (40, 14, 52, 60)):
        ws.column_dimensions[col].width = w
    for r in range(6, 25):
        ws.row_dimensions[r].height = 27
    ws.freeze_panes = "B5"


# ------------------------------------------------------------------- Scorecard
def build_scorecard(wb, ws, model):
    title(ws, "Scorecard",
          "Points for each variable band. A higher total score means a safer borrower. "
          "WoE and coefficients come from the logistic regression in run_model.py; points are calculated here.")
    put(ws, "B4", "Model intercept (alpha)")
    put(ws, "C4", model["intercept"], F_IN, FMT_NUM4, border=BOX)
    ws["C4"].comment = Comment("Fitted by statsmodels Logit on 114,770 training loans (run_model.py).", "Model")
    name(wb, "Model_Intercept", "Scorecard", "$C$4")
    put(ws, "B5", "Variables in the scorecard")
    put(ws, "B6", "Points formula")
    put(ws, "C6", "Points = ROUND( -(coefficient x WoE + intercept / variables) x Factor + Offset / variables , 0)", F_SUB)
    put(ws, "B7", "WoE (weight of evidence)")
    put(ws, "C7", "ln(share of good loans in the band / share of bad loans in the band). Positive = safer than average.", F_SUB)

    refs = {}
    row = 9
    coef_cells = []
    for var in model["variables"]:
        v = var["var"]
        put(ws, f"B{row}", var["label"], F_SECTION)
        put(ws, f"E{row}", "Coefficient", F_BOLD, align=Alignment(horizontal="right"))
        put(ws, f"F{row}", var["coef"], F_IN, FMT_NUM4, border=BOX)
        coef_cells.append(f"$F${row}")
        coef = f"$F${row}"
        row += 1
        if var["kind"] == "numeric":
            header_row(ws, row, 2, ["Band", "Upper edge (band includes it)", "Loans (training)", "Bad rate",
                                    "WoE", "Points"])
            row += 1
            labels = numeric_labels(var)
            first = row
            missing_row = None
            for b, lab in zip(var["bins"], labels):
                put(ws, f"B{row}", lab, border=BOX)
                if b["missing"]:
                    missing_row = row
                    put(ws, f"C{row}", "blank input", F_NOTE, border=BOX)
                elif b["upper_edge"] is not None:
                    put(ws, f"C{row}", edge_value(var, b["upper_edge"]), F_IN, EDGE_FORMAT[v], border=BOX)
                else:
                    put(ws, f"C{row}", "no limit", F_NOTE, border=BOX)
                put(ws, f"D{row}", b["count"], F_IN, FMT_INT, border=BOX)
                put(ws, f"E{row}", b["bad_rate"], F_IN, FMT_PCT1, border=BOX)
                put(ws, f"F{row}", b["woe"], F_IN, FMT_NUM4, border=BOX)
                put(ws, f"G{row}", f"=ROUND(-({coef}*F{row}+Model_Intercept/Model_Vars)*Score_Factor+Score_Offset/Model_Vars,0)",
                    F_CALC, '0', FILL_KEY, CENTER, BOX)
                row += 1
            n_numeric = len([b for b in var["bins"] if not b["missing"]])
            last_numeric = first + n_numeric - 1
            refs[v] = {"kind": "numeric", "edges": f"Scorecard!$C${first}:$C${last_numeric - 1}",
                       "labels": f"Scorecard!$B${first}:$B${last_numeric}",
                       "points": f"Scorecard!$G${first}:$G${last_numeric}",
                       "missing_label": f"Scorecard!$B${missing_row}" if missing_row else None,
                       "missing_points": f"Scorecard!$G${missing_row}" if missing_row else None}
        else:
            header_row(ws, row, 2, ["As entered", "Scored as", "Loans in group (training)", "Group bad rate",
                                    "WoE", "Points"])
            row += 1
            groups = {b["label"]: b for b in var["bins"]}
            names = model["purpose_names"] if v == "purpose" else model["verification_names"]
            # order entries by group risk, then by name
            order = [b["label"] for b in var["bins"]]
            entries = sorted(var["raw_to_group"].items(), key=lambda kv: (order.index(kv[1]), names.get(kv[0], kv[0])))
            first = row
            for raw, grp in entries:
                b = groups[grp]
                grp_name = names.get(grp, grp) if grp in names else grp
                put(ws, f"B{row}", names.get(raw, raw), border=BOX)
                put(ws, f"C{row}", grp_name, border=BOX)
                put(ws, f"D{row}", b["count"], F_IN, FMT_INT, border=BOX)
                put(ws, f"E{row}", b["bad_rate"], F_IN, FMT_PCT1, border=BOX)
                put(ws, f"F{row}", b["woe"], F_IN, FMT_NUM4, border=BOX)
                put(ws, f"G{row}", f"=ROUND(-({coef}*F{row}+Model_Intercept/Model_Vars)*Score_Factor+Score_Offset/Model_Vars,0)",
                    F_CALC, '0', FILL_KEY, CENTER, BOX)
                row += 1
            refs[v] = {"kind": "categorical", "names": f"Scorecard!$B${first}:$B${row - 1}",
                       "groups": f"Scorecard!$C${first}:$C${row - 1}",
                       "points": f"Scorecard!$G${first}:$G${row - 1}",
                       "first_name": names.get(entries[0][0], entries[0][0])}
        row += 2
    put(ws, "C5", f"=COUNT({','.join(coef_cells)})", F_CALC, '0', border=BOX)
    name(wb, "Model_Vars", "Scorecard", "$C$5")
    put(ws, f"B{row}", "Source: WoE, loan counts and bad rates from the 114,770-loan training sample; coefficients from "
        "the logistic regression of bad_loan on the WoE values (run_model.py).", F_NOTE)
    for col, w in zip("BCDEFG", (30, 26, 16, 12, 12, 10)):
        ws.column_dimensions[col].width = w
    return refs


# ------------------------------------------------------------------- Pricing Tool (inputs + score)
APPLICANT = {"loan_amnt": 12_000, "term_months": 36, "annual_inc": 65_000, "dti": 0.18, "revol_util": 0.45,
             "purpose": "Debt consolidation", "verification_status": "Verified"}


def build_tool_inputs(wb, ws, model, sc_refs):
    title(ws, "Loan Pricing Tool",
          "Enter an application in the yellow cells. The score, default risk, price and decision update automatically.")
    section(ws, "B4", "1. Application")
    rows = {
        "loan_amnt": (5, "Loan amount", "Requested amount, $"),
        "term_months": (6, "Term (months)", "36 or 60"),
        "annual_inc": (7, "Annual income", "Gross income, $"),
        "dti": (8, "Debt-to-income ratio", "Monthly debt payments / monthly income, before this loan"),
        "revol_util": (9, "Revolving credit utilization", "Card balances / card limits. Leave blank if no revolving credit"),
        "purpose": (10, "Loan purpose", "Pick from the list"),
        "verification_status": (11, "Income verified", "Pick from the list"),
    }
    cells = {}
    for v, (r, lab, note) in rows.items():
        put(ws, f"B{r}", lab)
        fmt = INPUT_FORMAT.get(v)
        put(ws, f"C{r}", APPLICANT[v], F_IN, fmt, FILL_INPUT, border=BOX)
        put(ws, f"D{r}", note, F_NOTE)
        cells[v] = f"C{r}"
    put(ws, "B12", "Loan-to-income ratio")
    put(ws, "C12", "=C5/C7", F_CALC, FMT_PCT1, border=BOX)
    put(ws, "D12", "Calculated: loan amount / annual income", F_NOTE)
    cells["loan_to_income"] = "C12"

    # data validation
    dv_term = DataValidation(type="list", formula1='"36,60"', allow_blank=False)
    dv_term.error, dv_term.errorTitle = "Term must be 36 or 60 months.", "Term"
    ws.add_data_validation(dv_term); dv_term.add("C6")
    dv_p = DataValidation(type="list", formula1=sc_refs["purpose"]["names"].replace("Scorecard!", "Scorecard!"),
                          allow_blank=False)
    ws.add_data_validation(dv_p); dv_p.add("C10")
    dv_v = DataValidation(type="list", formula1=sc_refs["verification_status"]["names"], allow_blank=False)
    ws.add_data_validation(dv_v); dv_v.add("C11")
    dv_amt = DataValidation(type="decimal", operator="between", formula1="500", formula2="40000")
    dv_amt.error = "Enter an amount between $500 and $40,000."
    ws.add_data_validation(dv_amt); dv_amt.add("C5")
    dv_inc = DataValidation(type="decimal", operator="greaterThan", formula1="0")
    dv_inc.error = "Income must be above zero."
    ws.add_data_validation(dv_inc); dv_inc.add("C7")
    dv_r = DataValidation(type="decimal", operator="between", formula1="0", formula2="1.5", allow_blank=True)
    dv_r.error = "Enter a percentage between 0% and 150%."
    ws.add_data_validation(dv_r); dv_r.add("C8"); dv_r.add("C9")

    # score build-up
    section(ws, "B14", "2. Credit score")
    header_row(ws, 15, 2, ["Variable", "Value", "Scorecard band", "Points"])
    ws.row_dimensions[15].height = 18
    r = 16
    point_cells = []
    for var in model["variables"]:
        v = var["var"]
        ref = sc_refs[v]
        x = cells[v]
        put(ws, f"B{r}", var["label"], border=BOX)
        put(ws, f"C{r}", f"={x}", F_CALC, INPUT_FORMAT.get(v), border=BOX,
            align=Alignment(horizontal="right"))
        if ref["kind"] == "numeric":
            idx = f"COUNTIF({ref['edges']},\"<\"&{x})+1"
            band = f"INDEX({ref['labels']},{idx})"
            pts = f"INDEX({ref['points']},{idx})"
            if ref["missing_points"]:
                band = f"IF(ISBLANK({x}),{ref['missing_label']},{band})"
                pts = f"IF(ISBLANK({x}),{ref['missing_points']},{pts})"
                ws[f"C{r}"] = f'=IF(ISBLANK({x}),"(blank)",{x})'
        else:
            m = f"MATCH({x},{ref['names']},0)"
            band = f"INDEX({ref['groups']},{m})"
            pts = f"INDEX({ref['points']},{m})"
        put(ws, f"D{r}", f"={band}", F_CALC, border=BOX)
        put(ws, f"E{r}", f"={pts}", F_CALC, '0', align=CENTER, border=BOX)
        point_cells.append(f"E{r}")
        r += 1
    put(ws, f"B{r}", "Total credit score", F_BOLD, border=BOX)
    for col in "CD":
        style(ws[f"{col}{r}"], border=BOX)
    put(ws, f"E{r}", f"=SUM(E16:E{r - 1})", F_BOLD, '0', FILL_KEY, CENTER, BOX)
    score_cell = f"E{r}"
    put(ws, f"B{r + 1}", "Scores in the holdout sample ranged from "
        f"{model['score_stats']['min']:.0f} to {model['score_stats']['max']:.0f} "
        f"(average {model['score_stats']['mean']:.0f}). Each 20 points doubles the odds of repayment.", F_NOTE)
    name(wb, "Applicant_Score", "Pricing Tool", f"${score_cell[0]}${score_cell[1:]}")
    name(wb, "Applicant_Amount", "Pricing Tool", "$C$5")
    name(wb, "Applicant_Term", "Pricing Tool", "$C$6")

    for col, w in zip("BCDEFGHI", (30, 20, 30, 10, 3, 40, 40, 3)):
        ws.column_dimensions[col].width = w
    return {"cells": cells, "score": score_cell, "last_row": r + 1}


# ------------------------------------------------------------------- Solver
SOLVER_PARAMS = ["Loan amount", "Term (months)", "Lifetime PD", "Monthly default hazard h",
                 "Monthly survival s = 1 - h", "Loss given default", "Recovery lag (months)", "Servicing fee",
                 "Upfront fee", "Monthly discount factor d", "q = s x d", "Hq = (1 - q^n) / (1 - q)",
                 "d ^ recovery lag", "Break-even APR", "Check: NPV at break-even ($)"]
N_ITER = 30


def npv_formula(col, apr):
    """Closed-form expected NPV at annual rate `apr` (a cell ref or expression) for the scenario in `col`."""
    p = {k: f"${col}${r}" for k, r in SOLVER_ROWS.items()}
    i = f"(({apr})/12)"
    gn = f"((1+{i})^{p['n']})"
    pay = f"({p['A']}*{i}*{gn}/({gn}-1))"
    x = f"({p['q']}*(1+{i}))"
    hx = f"IF(ABS(1-{x})<1E-12,{p['n']},(1-{x}^{p['n']})/(1-{x}))"
    return (f"{p['A']}*({p['fee']}-1)+{pay}*(1-{p['serv']})*{p['q']}*{p['Hq']}"
            f"+{p['h']}*(1-{p['lgd']})*{p['dk']}*{p['A']}*{p['d']}/({gn}-1)*({gn}*{p['Hq']}-{hx})")


SOLVER_ROWS = {}


def build_solver(wb, ws, model, tool):
    title(ws, "Solver: break-even APR",
          "Finds the APR at which a loan's expected net present value is exactly zero, by bisection.")
    notes = [
        "Expected NPV of one loan (closed form, the sum of geometric series over the term):",
        "   NPV(APR) = A x (upfront fee - 1)  +  P x (1 - servicing) x q x Hq  +  h x (1 - LGD) x d^lag x A x d / ((1+i)^n - 1) x ((1+i)^n x Hq - Hx)",
        "   where i = APR / 12, P = level monthly payment, h = monthly default hazard, d = monthly discount factor, q = (1 - h) x d,",
        "   Hq = sum of q^j and Hx = sum of (q x (1+i))^j for j = 0 .. n-1. Payments come from surviving loans; recoveries arrive "
        "'lag' months after a default.",
        "Bisection: start with APR between 0.01% and 100%. Each step tests the midpoint; if NPV < 0 the rate is too low, so the "
        "midpoint becomes the new low. After 30 steps the answer is accurate to about 0.0000001%.",
        "The Cash Flows sheet rebuilds the applicant's NPV month by month as an independent check.",
    ]
    for k, t in enumerate(notes):
        put(ws, f"B{3 + k}", t, F_SUB if k else F_BOLD)

    top = 10
    labels_rows = {lab: top + 1 + k for k, lab in enumerate(SOLVER_PARAMS)}
    keys = ["A", "n", "pd", "h", "s", "lgd", "lag", "serv", "fee", "d", "q", "Hq", "dk", "be", "check"]
    SOLVER_ROWS.update({k: labels_rows[lab] for k, lab in zip(keys, SOLVER_PARAMS)})
    header_row(ws, top, 2, ["Parameter"])
    for lab, r in labels_rows.items():
        put(ws, f"B{r}", lab, border=BOX)

    scenarios = [("Applicant", "applicant"), ("Applicant if PD were 0", "zero")]
    for b in model["bands"]:
        scenarios.append((f"Band: {b['label']}", "band"))
    band_cols = []
    blocks_start = top + len(SOLVER_PARAMS) + 3
    block_h = N_ITER + 3
    for j, (sname, kind) in enumerate(scenarios):
        col = L(3 + j)
        ws.column_dimensions[col].width = 15
        c = ws.cell(row=top, column=3 + j, value=sname)
        style(c, F_HEAD, fill=FILL_HEAD, align=Alignment(horizontal="center", vertical="center", wrap_text=True))
        R = SOLVER_ROWS
        if kind in ("applicant", "zero"):
            put(ws, f"{col}{R['A']}", "=Applicant_Amount", F_LINK, FMT_USD, border=BOX)
            put(ws, f"{col}{R['n']}", "=Applicant_Term", F_LINK, '0', border=BOX)
            if kind == "applicant":
                put(ws, f"{col}{R['pd']}", "=Applicant_PD", F_LINK, FMT_PCT2, border=BOX)
            else:
                put(ws, f"{col}{R['pd']}", 0, F_IN, FMT_PCT2, border=BOX)
        else:
            band_cols.append(col)
            bi = len(band_cols) - 1
            put(ws, f"{col}{R['A']}", "=Band_Amount", F_LINK, FMT_USD, border=BOX)
            put(ws, f"{col}{R['n']}", "=Band_Term", F_LINK, '0', border=BOX)
            put(ws, f"{col}{R['pd']}", f"='Risk-Based Pricing'!$G${BAND_FIRST_ROW + bi}", F_LINK, FMT_PCT2, border=BOX)
        put(ws, f"{col}{R['h']}", f"=1-(1-{col}{R['pd']})^(1/{col}{R['n']})", F_CALC, FMT_PCT4, border=BOX)
        put(ws, f"{col}{R['s']}", f"=1-{col}{R['h']}", F_CALC, FMT_DEC, border=BOX)
        put(ws, f"{col}{R['lgd']}", "=LGD_Rate", F_LINK, FMT_PCT1, border=BOX)
        put(ws, f"{col}{R['lag']}", "=Recovery_Lag", F_LINK, '0', border=BOX)
        put(ws, f"{col}{R['serv']}", "=Servicing_Fee", F_LINK, FMT_PCT1, border=BOX)
        put(ws, f"{col}{R['fee']}", "=Upfront_Fee", F_LINK, FMT_PCT1, border=BOX)
        put(ws, f"{col}{R['d']}", "=Monthly_Disc_Factor", F_LINK, FMT_DEC, border=BOX)
        put(ws, f"{col}{R['q']}", f"={col}{R['s']}*{col}{R['d']}", F_CALC, FMT_DEC, border=BOX)
        put(ws, f"{col}{R['Hq']}", f"=(1-{col}{R['q']}^{col}{R['n']})/(1-{col}{R['q']})", F_CALC, FMT_NUM4, border=BOX)
        put(ws, f"{col}{R['dk']}", "=Lag_Disc_Factor", F_LINK, FMT_DEC, border=BOX)

        # iteration block
        b0 = blocks_start + j * block_h
        put(ws, f"B{b0}", f"Bisection: {sname}", F_SECTION)
        header_row(ws, b0 + 1, 2, ["Step", "Low APR", "High APR", "Midpoint APR", "Expected NPV at midpoint ($)"])
        for k in range(N_ITER):
            r = b0 + 2 + k
            put(ws, f"B{r}", k + 1, F_CALC, '0', align=CENTER, border=BOX)
            if k == 0:
                put(ws, f"C{r}", 0.0001, F_IN, FMT_PCT4, border=BOX)
                put(ws, f"D{r}", 1.0, F_IN, FMT_PCT4, border=BOX)
            else:
                put(ws, f"C{r}", f"=IF(F{r - 1}<0,E{r - 1},C{r - 1})", F_CALC, FMT_PCT4, border=BOX)
                put(ws, f"D{r}", f"=IF(F{r - 1}<0,D{r - 1},E{r - 1})", F_CALC, FMT_PCT4, border=BOX)
            put(ws, f"E{r}", f"=(C{r}+D{r})/2", F_CALC, FMT_PCT4, border=BOX)
            put(ws, f"F{r}", "=" + npv_formula(col, f"E{r}"), F_CALC, FMT_USD2, border=BOX)
        last = b0 + 1 + N_ITER
        put(ws, f"{col}{R['be']}", f"=E{last}", F_CALC, FMT_PCT4, FILL_KEY, border=BOX)
        style(ws[f"{col}{R['be']}"], Font(name=FONT, size=10, bold=True))
        put(ws, f"{col}{R['check']}", "=" + npv_formula(col, f"{col}{R['be']}"), F_CALC, FMT_USD2, border=BOX)

    name(wb, "Applicant_Breakeven", "Solver", f"$C${SOLVER_ROWS['be']}")
    name(wb, "Applicant_Breakeven_PD0", "Solver", f"$D${SOLVER_ROWS['be']}")
    ws.column_dimensions["B"].width = 30
    for col in "CDEF":
        ws.column_dimensions[col].width = max(ws.column_dimensions[col].width or 0, 16)
    ws.freeze_panes = "C11"
    return {"band_cols": band_cols, "rows": dict(SOLVER_ROWS)}


# ------------------------------------------------------------------- Cash Flows
def build_cash_flows(wb, ws, tool):
    title(ws, "Cash Flows: applicant at the risk-based APR",
          "Expected monthly cash flows, weighting each month by the chance the loan is still paying. "
          "Independent check on the Solver's closed-form NPV.")
    params = [
        (4, "Loan amount", "=Applicant_Amount", FMT_USD, F_LINK),
        (5, "Term (months)", "=Applicant_Term", '0', F_LINK),
        (6, "Risk-based APR", "=Applicant_APR", FMT_PCT2, F_LINK),
        (7, "Monthly interest rate", "=C6/12", FMT_PCT4, F_CALC),
        (8, "Scheduled monthly payment", "=PMT(C7,C5,-C4)", FMT_USD2, F_CALC),
        (9, "Lifetime PD", "=Applicant_PD", FMT_PCT2, F_LINK),
        (10, "Monthly default hazard", "=1-(1-C9)^(1/C5)", FMT_PCT4, F_CALC),
        (11, "Loss given default", "=LGD_Rate", FMT_PCT1, F_LINK),
        (12, "Recovery lag (months)", "=Recovery_Lag", '0', F_LINK),
        (13, "Servicing fee", "=Servicing_Fee", FMT_PCT1, F_LINK),
        (14, "Monthly discount factor", "=Monthly_Disc_Factor", FMT_DEC, F_LINK),
    ]
    for r, lab, f, fmt, font in params:
        put(ws, f"B{r}", lab, border=BOX)
        put(ws, f"C{r}", f, font, fmt, border=BOX)

    hdr = 17
    first = hdr + 1          # month 0
    n_months = 60 + 12       # covers a 60-month loan plus a recovery lag of up to 12 months
    last = first + n_months
    cols = ["Month", "Beginning balance", "Scheduled payment", "Interest", "Principal", "Ending balance",
            "Still paying at start of month", "Chance of default this month", "Expected payment received (net of servicing)",
            "Expected recovery received", "Expected credit loss", "Expected net cash flow", "Discount factor",
            "Present value"]
    header_row(ws, hdr, 2, cols)
    ws.row_dimensions[hdr].height = 44
    r1 = first + 1
    # month 0
    put(ws, f"B{first}", 0, F_CALC, '0', align=CENTER, border=BOX)
    for col in "CDEFGHIJKL":
        style(ws[f"{col}{first}"], border=BOX)
    put(ws, f"M{first}", "=-C4*(1-Upfront_Fee)", F_CALC, FMT_USD2, border=BOX)
    put(ws, f"N{first}", 1, F_CALC, FMT_DEC, border=BOX)
    put(ws, f"O{first}", f"=M{first}*N{first}", F_CALC, FMT_USD2, border=BOX)
    rng = lambda c: f"${c}${r1}:${c}${last}"
    for r in range(r1, last + 1):
        m = f"B{r}"
        put(ws, m, f"=B{r - 1}+1", F_CALC, '0', align=CENTER, border=BOX)
        put(ws, f"C{r}", "=$C$4" if r == r1 else f"=IF({m}<=$C$5,G{r - 1},0)", F_CALC, FMT_USD2, border=BOX)
        put(ws, f"D{r}", f"=IF({m}<=$C$5,$C$8,0)", F_CALC, FMT_USD2, border=BOX)
        put(ws, f"E{r}", f"=C{r}*$C$7", F_CALC, FMT_USD2, border=BOX)
        put(ws, f"F{r}", f"=D{r}-E{r}", F_CALC, FMT_USD2, border=BOX)
        put(ws, f"G{r}", f"=C{r}-F{r}", F_CALC, FMT_USD2, border=BOX)
        put(ws, f"H{r}", f"=IF({m}<=$C$5,(1-$C$10)^({m}-1),0)", F_CALC, FMT_PCT2, border=BOX)
        put(ws, f"I{r}", f"=H{r}*$C$10", F_CALC, FMT_PCT4, border=BOX)
        put(ws, f"J{r}", f"=H{r}*(1-$C$10)*D{r}*(1-$C$13)", F_CALC, FMT_USD2, border=BOX)
        put(ws, f"K{r}", f"=IF({m}-$C$12>=1,INDEX({rng('I')},{m}-$C$12)*INDEX({rng('C')},{m}-$C$12)*(1-$C$11),0)",
            F_CALC, FMT_USD2, border=BOX)
        put(ws, f"L{r}", f"=I{r}*C{r}*$C$11", F_CALC, FMT_USD2, border=BOX)
        put(ws, f"M{r}", f"=J{r}+K{r}", F_CALC, FMT_USD2, border=BOX)
        put(ws, f"N{r}", f"=$C$14^{m}", F_CALC, FMT_DEC, border=BOX)
        put(ws, f"O{r}", f"=M{r}*N{r}", F_CALC, FMT_USD2, border=BOX)
    tot = last + 1
    put(ws, f"B{tot}", "Total", F_BOLD, border=TOP_RULE)
    for col in "IJKLMO":
        put(ws, f"{col}{tot}", f"=SUM({col}{first}:{col}{last})", F_BOLD,
            FMT_PCT2 if col == "I" else FMT_USD2, border=TOP_RULE)

    # summary block
    section(ws, "E3", "Results")
    summary = [
        (4, "Expected NPV ($)", f"=SUM(O{first}:O{last})", FMT_USD2, "CF_NPV"),
        (5, "Expected annual return (loss-adjusted IRR)", f"=(1+IRR(M{first}:M{last},C7))^12-1", FMT_PCT2, "CF_IRR"),
        (6, "Lifetime PD implied by the schedule", f"=I{tot}", FMT_PCT2, None),
        (7, "Exposure at default, expected ($)", f"=SUMPRODUCT(I{r1}:I{last},C{r1}:C{last})/I{tot}", FMT_USD2, "CF_EAD"),
        (8, "Expected credit loss ($)", f"=L{tot}", FMT_USD2, "CF_EL"),
        (9, "Expected credit loss (% of loan)", "=F8/C4", FMT_PCT2, None),
        (10, "Closed-form NPV from Solver ($)", "=" + npv_formula("C", "C6").replace("$C$", "Solver!$C$"), FMT_USD2, None),
        (11, "Difference (should be 0.00)", "=F4-F10", FMT_USD2, None),
    ]
    for r, lab, f, fmt, nm in summary:
        put(ws, f"E{r}", lab, border=BOX)
        put(ws, f"F{r}", f, F_CALC, fmt, FILL_KEY if nm else None, border=BOX)
        if nm:
            name(wb, nm, "Cash Flows", f"$F${r}")
    put(ws, "E12", "Note: payment at the scheduled rate; no prepayment. Rows past the term only carry recoveries.", F_NOTE)
    widths = {"B": 9, "C": 15, "D": 13, "E": 40, "F": 15, "G": 15, "H": 13, "I": 13, "J": 17, "K": 15,
              "L": 13, "M": 15, "N": 11, "O": 13}
    for c, w in widths.items():
        ws.column_dimensions[c].width = w
    ws.freeze_panes = f"C{first}"
    return {"first": first, "last": last}


# ------------------------------------------------------------------- Risk-Based Pricing
BAND_FIRST_ROW = 6


def build_risk_bands(wb, ws, model, solver):
    title(ws, "Risk-Based Pricing by score band",
          "Break-even APR for a standard 36-month loan in each score band, next to the average rate LendingClub "
          "actually charged those borrowers.")
    cols = ["Score band", "Score from", "Score to (exclusive)", "36-month loans", "Observed bad rate",
            "Average model PD", "Break-even APR", "Risk-based APR (with margin)", "LendingClub average rate",
            "LendingClub rate minus break-even", "Expected NPV per loan at LendingClub's rate ($)"]
    header_row(ws, 5, 2, cols)
    ws.row_dimensions[5].height = 44
    bands = model["bands"]
    R = solver["rows"]
    for k, b in enumerate(bands):
        r = BAND_FIRST_ROW + k
        sc = solver["band_cols"][k]
        put(ws, f"B{r}", b["label"], border=BOX)
        put(ws, f"C{r}", 0 if b["score_from"] is None else b["score_from"], F_IN, '0;-0;"-"', border=BOX)
        put(ws, f"D{r}", b["score_to"] if b["score_to"] is not None else "-", F_IN, '0', border=BOX,
            align=Alignment(horizontal="right"))
        put(ws, f"E{r}", b["loans"], F_IN, FMT_INT, border=BOX)
        put(ws, f"F{r}", b["bad_rate"], F_IN, FMT_PCT1, border=BOX)
        put(ws, f"G{r}", b["avg_pd"], F_IN, FMT_PCT1, border=BOX)
        put(ws, f"H{r}", f"=Solver!{sc}${R['be']}", F_LINK, FMT_PCT2, border=BOX)
        put(ws, f"I{r}", f"=H{r}+Target_Margin", F_CALC, FMT_PCT2, border=BOX)
        put(ws, f"J{r}", b["lc_rate"], F_IN, FMT_PCT2, border=BOX)
        put(ws, f"K{r}", f"=J{r}-H{r}", F_CALC, '+0.00%;-0.00%;0.00%', border=BOX)
        npv = npv_formula(sc, f"J{r}").replace(f"${sc}$", f"Solver!${sc}$")
        put(ws, f"L{r}", "=" + npv, F_CALC, '$#,##0;-$#,##0;"-"', border=BOX)
    last = BAND_FIRST_ROW + len(bands) - 1
    tot = last + 1
    put(ws, f"B{tot}", "All 36-month loans", F_BOLD, border=TOP_RULE)
    put(ws, f"E{tot}", f"=SUM(E{BAND_FIRST_ROW}:E{last})", F_BOLD, FMT_INT, border=TOP_RULE)
    for col, fmt in (("F", FMT_PCT1), ("G", FMT_PCT1), ("J", FMT_PCT2)):
        put(ws, f"{col}{tot}", f"=SUMPRODUCT({col}{BAND_FIRST_ROW}:{col}{last},$E${BAND_FIRST_ROW}:$E${last})/$E${tot}",
            F_BOLD, fmt, border=TOP_RULE)
    put(ws, f"L{tot}", f"=SUMPRODUCT(L{BAND_FIRST_ROW}:L{last},$E${BAND_FIRST_ROW}:$E${last})/$E${tot}",
        F_BOLD, '$#,##0;-$#,##0;"-"', border=TOP_RULE)
    put(ws, f"M{tot}", "loan-weighted average", F_NOTE)

    put(ws, f"B{tot + 2}", "Takeaway", F_SECTION)
    put(ws, f"B{tot + 3}",
        f'="LendingClub charged less than the break-even rate in "&COUNTIF(K{BAND_FIRST_ROW}:K{last},"<0")&" of "&'
        f'COUNT(K{BAND_FIRST_ROW}:K{last})&" bands. Its rates rise "&TEXT(100*(J{BAND_FIRST_ROW}-J{last}),"0.0")&'
        f'" percentage points from the safest to the riskiest band, while the break-even rate rises "&TEXT(100*(H{BAND_FIRST_ROW}-H{last}),"0.0")&'
        f'" points. A flatter price curve than the risk curve means safer borrowers subsidize riskier ones."', F_SUB)
    put(ws, f"B{tot + 5}",
        "Sources: loan counts, bad rates, model PD and LendingClub's average interest rate are computed in run_model.py "
        f"from all {sum(b['loans'] for b in bands):,} 36-month loans in the data. Break-even and NPV are live formulas "
        "using the Assumptions sheet. Comparison caveat: this model has no FICO score, which LendingClub's pricing did use.",
        F_NOTE)

    ch = LineChart()
    ch.title = "Break-even APR vs. LendingClub's rate, by score band"
    ch.y_axis.title = "APR"
    ch.y_axis.number_format = "0%"
    ch.x_axis.title = "Credit score band (riskiest to safest)"
    ch.height, ch.width = 8.5, 18
    data_ref = Reference(ws, min_col=8, min_row=5, max_row=last)
    ch.add_data(data_ref, titles_from_data=True)
    ch.add_data(Reference(ws, min_col=10, min_row=5, max_row=last), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=2, min_row=BAND_FIRST_ROW, max_row=last))
    for s, color in zip(ch.series, ("2A78D6", "EB6834")):
        s.graphicalProperties.line.solidFill = color
        s.graphicalProperties.line.width = 28000
        s.marker = Marker(symbol="circle", size=7)
        s.marker.graphicalProperties.solidFill = color
        s.marker.graphicalProperties.line.solidFill = color
        s.smooth = False
    ch.legend.position = "b"
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ws.add_chart(ch, f"B{tot + 7}")

    widths = {"B": 24, "C": 10, "D": 12, "E": 12, "F": 12, "G": 12, "H": 13, "I": 15, "J": 14, "K": 15, "L": 18, "M": 20}
    for c, w in widths.items():
        ws.column_dimensions[c].width = w
    return {"first": BAND_FIRST_ROW, "last": last}


# ------------------------------------------------------------------- Pricing Tool (outputs)
def build_tool_outputs(wb, ws, tool, solver, cf, bands):
    G, H = "G", "H"
    section(ws, "G4", "3. Decision")
    rows = []

    def line(r, label, formula, fmt, font=F_CALC, fill=None, nm=None, bold=False):
        put(ws, f"G{r}", label, F_BOLD if bold else F_CALC, border=BOX)
        f = Font(name=FONT, size=10, bold=bold, color=font.color.rgb[-6:] if font.color else BLACK)
        put(ws, f"H{r}", formula, f, fmt, fill, Alignment(horizontal="right"), BOX)
        if nm:
            name(wb, nm, "Pricing Tool", f"$H${r}")

    put(ws, "G5", "Decision", F_BOLD, border=BOX)
    put(ws, "H5", '=IF(Applicant_PD>Max_PD,"DECLINE: risk",IF(Applicant_APR>Max_APR,"DECLINE: rate cap","APPROVE"))',
        Font(name=FONT, size=11, bold=True), align=CENTER, border=BOX)
    put(ws, "G6", "Reason", border=BOX)
    put(ws, "H6", '=IF(Applicant_PD>Max_PD,"PD of "&TEXT(Applicant_PD,"0.0%")&" is above the "&TEXT(Max_PD,"0%")&" limit",'
        'IF(Applicant_APR>Max_APR,"Risk-based APR of "&TEXT(Applicant_APR,"0.0%")&" is above the "&TEXT(Max_APR,"0%")&" cap",'
        '"Within risk and rate limits"))', F_CALC, align=Alignment(horizontal="right"), border=BOX)
    line(7, "Offered APR", '=IF(H5="APPROVE",Applicant_APR,"Not offered")', FMT_PCT2, bold=True)
    line(8, "Monthly payment", '=IF(H5="APPROVE",PMT(Applicant_APR/12,Applicant_Term,-Applicant_Amount),"-")', FMT_USD2)

    section(ws, "G10", "4. Default risk")
    line(11, "Credit score", f"={tool['score']}", '0', bold=True)
    line(12, "Risk band", f"=INDEX('Risk-Based Pricing'!$B${bands['first']}:$B${bands['last']},"
                          f"MATCH(Applicant_Score,'Risk-Based Pricing'!$C${bands['first']}:$C${bands['last']},1))", None)
    line(13, "Lifetime probability of default (PD)", "=1/(1+EXP((Applicant_Score-Score_Offset)/Score_Factor))",
         FMT_PCT2, fill=FILL_KEY, nm="Applicant_PD", bold=True)
    line(14, "Monthly default hazard", "=1-(1-Applicant_PD)^(1/Applicant_Term)", FMT_PCT4)
    line(15, "Loss given default (LGD)", "=LGD_Rate", FMT_PCT1, font=F_LINK)
    line(16, "Exposure at default, expected (EAD)", "=CF_EAD", FMT_USD, font=F_LINK)
    line(17, "Expected credit loss ($)", "=CF_EL", FMT_USD, font=F_LINK)
    line(18, "Expected credit loss (% of loan)", "=CF_EL/Applicant_Amount", FMT_PCT2)
    put(ws, "G19", "PD x LGD x EAD = " , F_NOTE)
    put(ws, "H19", "=TEXT(Applicant_PD*LGD_Rate*CF_EAD,\"$#,##0\")&\" (same thing)\"", F_NOTE,
        align=Alignment(horizontal="right"))

    section(ws, "G21", "5. Pricing build-up (APR)")
    line(22, "Required return, as an APR", "=Required_APR", FMT_PCT2, font=F_LINK)
    line(23, "+ Servicing and fees", "=Applicant_Breakeven_PD0-Required_APR", FMT_PCT2)
    line(24, "+ Credit risk premium", "=Applicant_Breakeven-Applicant_Breakeven_PD0", FMT_PCT2)
    line(25, "Break-even APR", "=Applicant_Breakeven", FMT_PCT2, bold=True)
    line(26, "+ Target margin", "=Target_Margin", FMT_PCT2, font=F_LINK)
    line(27, "Risk-based APR", "=Applicant_Breakeven+Target_Margin", FMT_PCT2, fill=FILL_KEY, nm="Applicant_APR", bold=True)
    line(28, "LendingClub average rate for this band", f"=INDEX('Risk-Based Pricing'!$J${bands['first']}:$J${bands['last']},"
         f"MATCH(Applicant_Score,'Risk-Based Pricing'!$C${bands['first']}:$C${bands['last']},1))", FMT_PCT2)
    style(ws["H25"], border=Border(top=Side(style="thin", color=BLACK), bottom=THIN, left=THIN, right=THIN))
    style(ws["H27"], border=Border(top=Side(style="thin", color=BLACK), bottom=THIN, left=THIN, right=THIN))

    section(ws, "G30", "6. Expected economics at the risk-based APR")
    line(31, "Expected NPV at the hurdle rate ($)", "=CF_NPV", FMT_USD2, font=F_LINK)
    line(32, "Expected annual return after losses", "=CF_IRR", FMT_PCT2, font=F_LINK)
    line(33, "Required annual return", "=Hurdle_Rate", FMT_PCT2, font=F_LINK)
    put(ws, "G34", "Return exceeds the hurdle by roughly the target margin, after expected defaults.", F_NOTE)

    # legend
    lr = tool["last_row"] + 2
    section(ws, f"B{lr}", "Color key")
    put(ws, f"B{lr + 1}", "Input you can change", F_IN, fill=FILL_INPUT, border=BOX)
    put(ws, f"B{lr + 2}", "Formula", F_CALC, border=BOX)
    put(ws, f"B{lr + 3}", "Link to another sheet", F_LINK, border=BOX)
    put(ws, f"B{lr + 4}", "Key result", F_BOLD, fill=FILL_KEY, border=BOX)
    put(ws, f"B{lr + 6}", "Example application shown. Assumptions (hurdle rate, LGD, margin, limits) are on the Assumptions sheet.",
        F_NOTE)

    ws.conditional_formatting.add("H5", FormulaRule(formula=['LEFT($H$5,7)="APPROVE"'],
                                                    fill=PatternFill("solid", fgColor="D9EAD3"),
                                                    font=Font(name=FONT, bold=True, color="006100")))
    ws.conditional_formatting.add("H5", FormulaRule(formula=['LEFT($H$5,7)="DECLINE"'],
                                                    fill=PatternFill("solid", fgColor="F4CCCC"),
                                                    font=Font(name=FONT, bold=True, color="9C0006")))
    ws.freeze_panes = "A4"


# ------------------------------------------------------------------- Model Performance
def build_performance(wb, ws, model):
    title(ws, "Model Performance",
          "How well the scorecard separates good and bad loans, and whether its predicted default rates match reality. "
          "Values computed in run_model.py.")
    m = model["meta"]
    section(ws, "B4", "Data")
    rows = [("Loans in the data", m["loans"], FMT_INT), ("Training sample (70%)", m["train"], FMT_INT),
            ("Holdout sample (30%, never used to fit)", m["test"], FMT_INT), ("Overall bad-loan rate", m["bad_rate"], FMT_PCT1),
            ("Candidate variables screened", m["n_vars_screened"], '0')]
    for k, (lab, v, fmt) in enumerate(rows):
        put(ws, f"B{5 + k}", lab, border=BOX)
        put(ws, f"C{5 + k}", v, F_IN, fmt, border=BOX)
    put(ws, "B10", "Variables kept (IV of 0.02 or more)", border=BOX)
    put(ws, "C10", "=COUNTIF(E40:E60,\"Yes\")", F_CALC, '0', border=BOX)
    put(ws, "D5", f"Source: {m['source']}", F_NOTE)
    put(ws, "D6", m["source_url"], F_NOTE)

    section(ws, "B12", "Discrimination: how well each model ranks risk")
    header_row(ws, 13, 2, ["Model", "AUC", "Gini = 2 x AUC - 1", "KS statistic"])
    for k, r in enumerate(model["metrics"]):
        row = 14 + k
        put(ws, f"B{row}", r["model"], border=BOX)
        put(ws, f"C{row}", r["auc"], F_IN, '0.000', border=BOX)
        put(ws, f"D{row}", f"=2*C{row}-1", F_CALC, '0.000', border=BOX)
        put(ws, f"E{row}", r["ks"], F_IN, '0.000', border=BOX)
    put(ws, "B19", "AUC 0.5 = random guessing, 1.0 = perfect. The scorecard matches LendingClub's own pricing grade "
        "without using its interest rate or a FICO score.", F_NOTE)

    section(ws, "B21", "Calibration on the holdout set")
    header_row(ws, 22, 2, ["Risk decile (1 = safest)", "Loans", "Predicted PD", "Observed bad rate", "Difference"])
    for k, r in enumerate(model["calibration"]):
        row = 23 + k
        put(ws, f"B{row}", int(r["decile"]), F_IN, '0', align=CENTER, border=BOX)
        put(ws, f"C{row}", int(r["loans"]), F_IN, FMT_INT, border=BOX)
        put(ws, f"D{row}", r["predicted_pd"], F_IN, FMT_PCT1, border=BOX)
        put(ws, f"E{row}", r["observed_bad_rate"], F_IN, FMT_PCT1, border=BOX)
        put(ws, f"F{row}", f"=D{row}-E{row}", F_CALC, '+0.0%;-0.0%;0.0%', border=BOX)
    put(ws, "B34", "Hosmer-Lemeshow statistic", border=BOX)
    put(ws, "C34", model["hosmer_lemeshow"]["stat"], F_IN, '0.0', border=BOX)
    put(ws, "D34", "p-value", border=BOX)
    put(ws, "E34", model["hosmer_lemeshow"]["p_value"], F_IN, '0.000', border=BOX)
    put(ws, "B35", "Population stability index (training vs. holdout scores)", border=BOX)
    put(ws, "C35", model["psi_train_vs_holdout"], F_IN, '0.0000', border=BOX)
    put(ws, "D35", "Below 0.10 = stable", F_NOTE)
    put(ws, "B36", "With 49,000 holdout loans the Hosmer-Lemeshow test flags even small gaps; the decile table shows "
        "predicted and observed rates within about 1.5 points everywhere.", F_NOTE)

    ch = BarChart()
    ch.type = "col"
    ch.title = "Predicted vs. observed default rate by decile"
    ch.y_axis.number_format = "0%"
    ch.y_axis.title = "Default rate"
    ch.x_axis.title = "Risk decile (1 = safest)"
    ch.add_data(Reference(ws, min_col=4, min_row=22, max_row=32), titles_from_data=True)
    ch.add_data(Reference(ws, min_col=5, min_row=22, max_row=32), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=2, min_row=23, max_row=32))
    for s, color in zip(ch.series, ("2A78D6", "EB6834")):
        s.graphicalProperties.solidFill = color
        s.graphicalProperties.line.solidFill = color
    ch.gapWidth = 60
    ch.height, ch.width = 7.5, 14
    ch.legend.position = "b"
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ws.add_chart(ch, "H21")

    section(ws, "B38", "Variable screening by information value (IV)")
    header_row(ws, 39, 2, ["Variable", "Information value", "Strength", "In scorecard"])
    for k, r in enumerate(model["iv"]):
        row = 40 + k
        put(ws, f"B{row}", r["label"], border=BOX)
        put(ws, f"C{row}", r["iv"], F_IN, '0.0000', border=BOX)
        put(ws, f"D{row}", f'=IF(C{row}<0.02,"Not useful",IF(C{row}<0.1,"Weak",IF(C{row}<0.3,"Medium","Strong")))', border=BOX)
        put(ws, f"E{row}", f'=IF(C{row}>=0.02,"Yes","No")', align=CENTER, border=BOX)
    end = 40 + len(model["iv"])
    put(ws, f"B{end}", "Rule of thumb: IV below 0.02 is not useful, 0.02-0.1 weak, 0.1-0.3 medium, above 0.3 strong. "
        "Consumer-loan variables without bureau data are usually weak to medium on their own.", F_NOTE)

    section(ws, f"B{end + 2}", "Logistic regression on WoE values (training set)")
    header_row(ws, end + 3, 2, ["Term", "Coefficient", "Std. error", "p-value"])
    labels = {v["var"]: v["label"] for v in model["variables"]}
    for k, r in enumerate(model["coefficients"]):
        row = end + 4 + k
        put(ws, f"B{row}", "Intercept" if r["term"] == "const" else labels.get(r["term"], r["term"]), border=BOX)
        put(ws, f"C{row}", r["coef"], F_IN, '0.0000', border=BOX)
        put(ws, f"D{row}", r["std_err"], F_IN, '0.0000', border=BOX)
        put(ws, f"E{row}", r["p_value"], F_IN, '0.0000', border=BOX)
    put(ws, f"B{end + 5 + len(model['coefficients'])}",
        "All coefficients are negative, as expected: a higher WoE (safer band) lowers the odds of default. "
        "All are significant at the 1% level.", F_NOTE)
    for c, w in zip("BCDEF", (46, 16, 16, 16, 12)):
        ws.column_dimensions[c].width = w


# ------------------------------------------------------------------- Read Me
def build_readme(wb, ws, model):
    title(ws, "Credit Risk Scorecard & Loan Pricing Model",
          "Predicts a borrower's chance of default from real LendingClub loan data, then prices the loan so the lender "
          "earns its required return after expected losses.")
    r = 4
    blocks = [
        ("How to use it", [
            "1.  Go to the Pricing Tool sheet and enter an application in the yellow cells.",
            "2.  Read the decision, offered APR and the build-up of the price on the right-hand side.",
            "3.  Change lender assumptions (required return, loss given default, margin, limits) on the Assumptions sheet.",
        ]),
        ("What each sheet does", [
            "Pricing Tool  -  application inputs, credit score, default risk, decision and price.",
            "Assumptions  -  every lender and scorecard assumption, with sources.",
            "Scorecard  -  points for each variable band, calculated from the fitted model.",
            "Cash Flows  -  month-by-month expected cash flows for the applicant; checks the Solver's NPV.",
            "Solver  -  finds the break-even APR by bisection on a closed-form expected-NPV formula.",
            "Risk-Based Pricing  -  break-even APR by score band vs. what LendingClub actually charged.",
            "Model Performance  -  AUC, KS, calibration, variable screening and regression output.",
        ]),
        ("Method", [
            "Probability of default (PD): a points scorecard built from logistic regression on weight-of-evidence (WoE) "
            "binned variables, the standard approach in consumer lending because every point can be explained.",
            "Loss given default (LGD): 90%, from published research on LendingClub defaults (see Assumptions).",
            "Exposure at default (EAD): the expected outstanding balance when a default happens, from the amortization schedule.",
            "Expected loss = PD x LGD x EAD. The break-even APR makes the loan's expected NPV zero at the lender's "
            "required return; the offered APR adds a target margin.",
        ]),
        ("Data", [
            f"{model['meta']['loans']:,} completed LendingClub loans, each marked good or bad (charged off or defaulted). "
            f"Overall bad rate {model['meta']['bad_rate']:.1%}.",
            f"Source: {model['meta']['source']}.",
            "LendingClub's interest rate was held out of the model: a pricing model cannot use the price as an input. "
            "It is used only as a benchmark.",
        ]),
        ("Limitations", [
            "No credit bureau score (FICO) in the data, which caps how well any model can rank risk.",
            "Defaults are spread evenly over the loan's life (constant monthly hazard); real defaults cluster in the first two years.",
            "No prepayment modeled. The 70/30 split is random because the file has no issue dates for an out-of-time test.",
        ]),
    ]
    for head, lines in blocks:
        section(ws, f"B{r}", head)
        r += 1
        for t in lines:
            put(ws, f"B{r}", t, F_CALC, align=Alignment(wrap_text=True, vertical="top"))
            ws.row_dimensions[r].height = 27 if len(t) > 110 else 15
            r += 1
        r += 1
    section(ws, f"B{r}", "Color key")
    put(ws, f"B{r + 1}", "Blue text: typed-in input or historical data. Yellow fill: a cell you are meant to change.",
        F_IN, fill=FILL_INPUT)
    put(ws, f"B{r + 2}", "Black text: formula.", F_CALC)
    put(ws, f"B{r + 3}", "Green text: link to a value on another sheet.", F_LINK)
    put(ws, f"B{r + 4}", "Shaded: key result.", F_BOLD, fill=FILL_KEY)
    ws.column_dimensions["B"].width = 120


if __name__ == "__main__":
    path = build(load_model())
    print(f"wrote {path}")
