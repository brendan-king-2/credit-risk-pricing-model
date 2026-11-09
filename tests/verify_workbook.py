"""End-to-end check of the Excel workbook against the Python model.

For several test applications, writes the inputs into a copy of the workbook,
recalculates it in LibreOffice, and compares score, PD, break-even APR and the
decision with the Python implementation.

    python tests/verify_workbook.py path/to/recalc.py
"""
from __future__ import annotations

import json
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
from creditrisk import pricing  # noqa: E402
from creditrisk.model_io import load_model, score_application  # noqa: E402

WORKBOOK = ROOT / "excel" / "Credit_Risk_Pricing_Model.xlsx"

CASES = [
    # Excel inputs (display values)                                   Python inputs (raw codes, %-points)
    dict(amount=12000, term=36, income=65000, dti=0.18, revol=0.45, purpose="Debt consolidation", verified="Verified"),
    dict(amount=30000, term=60, income=48000, dti=0.31, revol=0.92, purpose="Small business", verified="Verified"),
    dict(amount=5000, term=36, income=150000, dti=0.04, revol=None, purpose="Major purchase", verified="Not verified"),
    dict(amount=20000, term=60, income=90000, dti=0.22, revol=0.65, purpose="Wedding", verified="Not verified"),
    dict(amount=35000, term=60, income=30000, dti=0.30, revol=0.95, purpose="Small business", verified="Verified"),
]


def main(recalc: str) -> int:
    model = load_model()
    purpose_code = {v: k for k, v in model["purpose_names"].items()}
    verif_code = {v: k for k, v in model["verification_names"].items()}
    a = {r: c for r, c in [("hurdle", 0.06), ("serv", 0.01), ("lgd", 0.90), ("lag", 6), ("margin", 0.01),
                           ("max_apr", 0.36), ("max_pd", 0.35)]}
    failures = 0
    tmp = Path(tempfile.mkdtemp())
    for k, c in enumerate(CASES):
        path = tmp / f"case{k}.xlsx"
        shutil.copy(WORKBOOK, path)
        wb = load_workbook(path)
        ws = wb["Pricing Tool"]
        ws["C5"], ws["C6"], ws["C7"], ws["C8"] = c["amount"], c["term"], c["income"], c["dti"]
        ws["C9"] = c["revol"]
        ws["C10"], ws["C11"] = c["purpose"], c["verified"]
        wb.save(path)
        out = json.loads(subprocess.run([sys.executable, recalc, str(path), "120"], capture_output=True, text=True).stdout)
        assert out.get("status") == "success", out
        v = load_workbook(path, data_only=True)
        t, s = v["Pricing Tool"], v["Solver"]
        xl = {"score": t["E24"].value, "pd": t["H13"].value, "be": s["C24"].value, "decision": t["H5"].value,
              "cf_check": v["Cash Flows"]["F11"].value}

        py = score_application(model, {
            "loan_amnt": c["amount"], "term_months": c["term"], "annual_inc": c["income"], "dti": c["dti"] * 100,
            "revol_util": None if c["revol"] is None else c["revol"] * 100,
            "purpose": purpose_code[c["purpose"]], "verification_status": verif_code[c["verified"]]})
        terms = pricing.LoanTerms(amount=c["amount"], term=c["term"], lgd=a["lgd"], recovery_lag=a["lag"],
                                  hurdle=a["hurdle"], servicing=a["serv"])
        be = pricing.breakeven_apr(py["pd"], terms)
        decision = ("DECLINE: risk" if py["pd"] > a["max_pd"] else
                    "DECLINE: rate cap" if be + a["margin"] > a["max_apr"] else "APPROVE")
        ok = (xl["score"] == py["score"] and abs(xl["pd"] - py["pd"]) < 1e-9 and abs(xl["be"] - be) < 1e-6
              and xl["decision"] == decision and abs(xl["cf_check"]) < 1e-6)
        failures += not ok
        print(f"case {k}: score xl={xl['score']} py={py['score']} | PD xl={xl['pd']:.4%} py={py['pd']:.4%} | "
              f"break-even xl={xl['be']:.4%} py={be:.4%} | {xl['decision']} vs {decision} | "
              f"CF-vs-closed-form {xl['cf_check']:.2e} -> {'OK' if ok else 'MISMATCH'}")
    shutil.rmtree(tmp)
    return failures


if __name__ == "__main__":
    sys.exit(main(sys.argv[1]))
