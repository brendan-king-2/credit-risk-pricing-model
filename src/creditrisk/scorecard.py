"""Points-based credit scorecard (logistic regression on WoE-transformed variables).

Scaling: 600 points = odds of 50 good loans to 1 bad loan, and every 20 points
doubles the odds (PDO = 20). A higher score means a safer borrower.

    Score = Offset + Factor * ln(odds_good)
    Factor = PDO / ln(2),  Offset = 600 - Factor * ln(50)
    PD     = 1 / (1 + exp((Score - Offset) / Factor))
"""
from __future__ import annotations

from dataclasses import dataclass

import numpy as np
import pandas as pd
import statsmodels.api as sm

from .binning import Binning

PDO, BASE_SCORE, BASE_ODDS = 20.0, 600.0, 50.0
FACTOR = PDO / np.log(2)
OFFSET = BASE_SCORE - FACTOR * np.log(BASE_ODDS)


def score_to_pd(score):
    return 1.0 / (1.0 + np.exp((np.asarray(score, dtype=float) - OFFSET) / FACTOR))


def pd_to_score(p):
    p = np.asarray(p, dtype=float)
    return OFFSET + FACTOR * np.log((1 - p) / p)


@dataclass
class Scorecard:
    binnings: dict[str, Binning]
    intercept: float
    coefs: dict[str, float]
    fit_summary: pd.DataFrame
    points: pd.DataFrame        # var, bin, woe, coef, points, count, bad_rate

    def woe_matrix(self, df: pd.DataFrame) -> pd.DataFrame:
        return pd.DataFrame({v: self.binnings[v].woe(df[v]) for v in self.coefs})

    def score(self, df: pd.DataFrame) -> pd.Series:
        """Integer points total, exactly as the Excel scorecard computes it."""
        total = pd.Series(0.0, index=df.index)
        for v in self.coefs:
            pts = self.points[self.points["var"] == v].set_index("bin")["points"]
            total += self.binnings[v].labels(df[v]).map(pts).astype(float)
        return total

    def predict_pd(self, df: pd.DataFrame) -> pd.Series:
        return pd.Series(score_to_pd(self.score(df)), index=df.index)


def fit(train: pd.DataFrame, target: str, binnings: dict[str, Binning]) -> Scorecard:
    X = pd.DataFrame({v: b.woe(train[v]) for v, b in binnings.items()})
    y = train[target]
    model = sm.Logit(y, sm.add_constant(X)).fit(disp=0)
    # logit models P(bad); with WoE = ln(good/bad) the fitted betas are negative
    alpha = float(model.params["const"])
    betas = {v: float(model.params[v]) for v in binnings}
    k = len(betas)
    rows = []
    for v, b in binnings.items():
        for _, r in b.table.iterrows():
            raw = -(betas[v] * r["woe"] + alpha / k) * FACTOR + OFFSET / k
            rows.append({"var": v, "bin": r["bin"], "count": int(r["count"]), "bad_rate": r["bad_rate"],
                         "woe": r["woe"], "coef": betas[v], "points_raw": raw, "points": int(round(raw))})
    pts = pd.DataFrame(rows)
    summ = pd.DataFrame({"coef": model.params, "std_err": model.bse, "z": model.tvalues, "p_value": model.pvalues})
    return Scorecard(binnings=binnings, intercept=alpha, coefs=betas, fit_summary=summ, points=pts)
