"""Unit tests. Run with `pytest -q` after `python run_model.py`."""
import sys
from pathlib import Path

import numpy as np
import pandas as pd
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from creditrisk import binning, pricing, scorecard as sc  # noqa: E402
from creditrisk.model_io import MODEL_PATH, load_model, score_application  # noqa: E402

T = pricing.LoanTerms()


@pytest.mark.parametrize("pd_life", [0.0, 0.05, 0.2, 0.45])
@pytest.mark.parametrize("term", [36, 60])
def test_closed_form_npv_matches_monthly_schedule(pd_life, term):
    t = pricing.LoanTerms(term=term)
    for apr in (0.05, 0.12, 0.30):
        assert pricing.npv_closed_form(apr, pd_life, t) == pytest.approx(pricing.npv_from_schedule(apr, pd_life, t), abs=1e-6)


def test_breakeven_has_zero_npv_and_rises_with_risk():
    bes = [pricing.breakeven_apr(p, T) for p in (0.0, 0.1, 0.2, 0.3)]
    assert all(abs(pricing.npv_closed_form(b, p, T)) < 1e-4 for b, p in zip(bes, (0.0, 0.1, 0.2, 0.3)))
    assert bes == sorted(bes)


def test_riskless_loan_without_fees_breaks_even_at_hurdle():
    t = pricing.LoanTerms(servicing=0.0, upfront_fee=0.0)
    monthly = (1 + t.hurdle) ** (1 / 12) - 1
    assert pricing.breakeven_apr(0.0, t) == pytest.approx(12 * monthly, abs=1e-8)


def test_breakeven_does_not_depend_on_loan_size():
    a = pricing.breakeven_apr(0.15, pricing.LoanTerms(amount=2_000))
    b = pricing.breakeven_apr(0.15, pricing.LoanTerms(amount=35_000))
    assert a == pytest.approx(b, abs=1e-10)


def test_hazard_reproduces_lifetime_pd():
    rows, _ = pricing.cash_flow_schedule(0.12, 0.22, T)
    assert sum(r["p_default"] for r in rows) == pytest.approx(0.22)


def test_score_pd_round_trip():
    assert sc.score_to_pd(sc.pd_to_score(0.1)) == pytest.approx(0.1)
    assert sc.score_to_pd(600) == pytest.approx(1 / 51)                     # odds 50:1 at 600
    assert sc.score_to_pd(620) == pytest.approx(1 / 101)                    # PDO: odds double


def test_monotonic_binning():
    rng = np.random.default_rng(0)
    x = pd.Series(rng.uniform(0, 100, 20_000))
    y = pd.Series((rng.uniform(size=20_000) < 0.05 + 0.002 * x).astype(int))
    b = binning.bin_numeric(x, y, "x")
    assert b.table["bad_rate"].is_monotonic_increasing
    assert (b.table["share"] >= 0.05).all()


@pytest.mark.skipif(not MODEL_PATH.exists(), reason="run run_model.py first")
def test_json_scorer_matches_fitted_scorecard_points():
    model = load_model()
    pts = pd.read_csv(MODEL_PATH.parent / "scorecard_points.csv")
    app = {"loan_amnt": 12_000, "term_months": 36, "annual_inc": 65_000, "dti": 18.0, "revol_util": 45.0,
           "purpose": "debt_consolidation", "verification_status": "verified"}
    out = score_application(model, app)
    expected = sum(pts[(pts["var"] == v) & (pts["bin"] == lab)]["points"].iloc[0] for v, (lab, _) in out["detail"].items())
    assert out["score"] == expected
