"""Discrimination and calibration metrics."""
from __future__ import annotations

import numpy as np
import pandas as pd
from sklearn.metrics import roc_auc_score, roc_curve


def auc(y, p) -> float:
    return float(roc_auc_score(y, p))


def gini(y, p) -> float:
    return 2 * auc(y, p) - 1


def ks(y, p) -> float:
    fpr, tpr, _ = roc_curve(y, p)
    return float(np.max(tpr - fpr))


def calibration_table(y, p, bins: int = 10) -> pd.DataFrame:
    """Predicted vs observed bad rate by predicted-PD decile."""
    d = pd.DataFrame({"y": np.asarray(y), "p": np.asarray(p)})
    d["decile"] = pd.qcut(d["p"].rank(method="first"), bins, labels=range(1, bins + 1)).astype(int)
    t = d.groupby("decile").agg(loans=("y", "size"), predicted_pd=("p", "mean"), observed_bad_rate=("y", "mean"),
                                min_pd=("p", "min"), max_pd=("p", "max"))
    return t.reset_index()


def hosmer_lemeshow(y, p, bins: int = 10) -> tuple[float, float]:
    from scipy import stats
    t = calibration_table(y, p, bins)
    exp_bad = t["predicted_pd"] * t["loans"]
    obs_bad = t["observed_bad_rate"] * t["loans"]
    exp_good = t["loans"] - exp_bad
    obs_good = t["loans"] - obs_bad
    hl = float((((obs_bad - exp_bad) ** 2) / exp_bad + ((obs_good - exp_good) ** 2) / exp_good).sum())
    return hl, float(1 - stats.chi2.cdf(hl, bins - 2))


def psi(expected: np.ndarray, actual: np.ndarray, edges: np.ndarray) -> float:
    """Population stability index between two score distributions."""
    e = np.histogram(expected, edges)[0] / len(expected)
    a = np.histogram(actual, edges)[0] / len(actual)
    e, a = np.clip(e, 1e-6, None), np.clip(a, 1e-6, None)
    return float(np.sum((a - e) * np.log(a / e)))
