"""Load and clean the LendingClub loan file.

Source: LendingClub public loan history (completed loans), as packaged by H2O.ai
for its consumer-loan demo app (github.com/h2oai/app-consumer-loan, data/loan.csv).
One row per loan; `bad_loan` = 1 if the loan was charged off or defaulted over its life.
"""
from __future__ import annotations

from pathlib import Path
from urllib.request import urlretrieve

import numpy as np
import pandas as pd

DATA_URL = "https://raw.githubusercontent.com/h2oai/app-consumer-loan/master/data/loan.csv"
RAW_PATH = Path(__file__).resolve().parents[2] / "data" / "raw" / "lendingclub_loans.csv"

TARGET = "bad_loan"

# Variables offered to the scorecard. LendingClub's own interest rate is held out:
# it is the platform's risk grade, and a pricing model cannot use the price as an input.
NUMERIC = ["term_months", "loan_amnt", "annual_inc", "loan_to_income", "dti", "revol_util",
           "delinq_2yrs", "total_acc", "credit_history_yrs", "emp_length"]
CATEGORICAL = ["home_ownership", "purpose", "verification_status"]

# how each numeric variable's bin edges are rounded and labelled
BIN_SPEC = {
    "term_months": dict(round_to=None, unit="num"),
    "loan_amnt": dict(round_to=1000, unit="$"),
    "annual_inc": dict(round_to=5000, unit="$"),
    "loan_to_income": dict(round_to=0.01, unit="pct"),
    "dti": dict(round_to=1, unit="num"),
    "revol_util": dict(round_to=5, unit="num"),
    "delinq_2yrs": dict(round_to=None, unit="num"),
    "total_acc": dict(round_to=None, unit="num"),
    "credit_history_yrs": dict(round_to=None, unit="num"),
    "emp_length": dict(round_to=None, unit="num"),
}

LABELS = {
    "term_months": "Loan term (months)", "loan_amnt": "Loan amount", "annual_inc": "Annual income",
    "loan_to_income": "Loan-to-income ratio", "dti": "Debt-to-income ratio (%)",
    "revol_util": "Revolving credit utilization (%)", "delinq_2yrs": "Delinquencies, last 2 years",
    "total_acc": "Total credit accounts", "credit_history_yrs": "Credit history (years)",
    "emp_length": "Years employed", "home_ownership": "Home ownership", "purpose": "Loan purpose",
    "verification_status": "Income verified",
}


def download(path: Path = RAW_PATH) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    urlretrieve(DATA_URL, path)
    return path


def load(path: Path = RAW_PATH) -> pd.DataFrame:
    if not path.exists():
        download(path)
    df = pd.read_csv(path)
    df = df.dropna(subset=["annual_inc", "delinq_2yrs", "total_acc", "longest_credit_length"])
    df = df[df["annual_inc"] > 0].copy()
    df["term_months"] = df["term"].str.extract(r"(\d+)").astype(int)
    df["loan_to_income"] = df["loan_amnt"] / df["annual_inc"]
    df["credit_history_yrs"] = df["longest_credit_length"]
    # rare ownership codes (OTHER / NONE / ANY, ~0.1% of loans) folded into one group
    df["home_ownership"] = df["home_ownership"].where(df["home_ownership"].isin(["MORTGAGE", "RENT", "OWN"]), "OTHER")
    df = df.reset_index(drop=True)
    return df


def split(df: pd.DataFrame, test_size: float = 0.3, seed: int = 42):
    """Stratified random train / holdout split (the file has no issue dates for an out-of-time test)."""
    rng = np.random.default_rng(seed)
    test_idx = []
    for _, grp in df.groupby(TARGET):
        n = int(round(len(grp) * test_size))
        test_idx.extend(rng.choice(grp.index.to_numpy(), n, replace=False))
    test_mask = df.index.isin(test_idx)
    return df[~test_mask].reset_index(drop=True), df[test_mask].reset_index(drop=True)
