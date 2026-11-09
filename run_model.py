"""Fit the credit scorecard, validate it, and export everything the Excel workbook needs.

    python run_model.py            # fit + validate + figures + outputs/model.json
    python build_workbook.py       # write excel/Credit_Risk_Pricing_Model.xlsx

Outputs land in outputs/ (tables, figures, model.json).
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

import numpy as np
import pandas as pd
from sklearn.ensemble import HistGradientBoostingClassifier
from sklearn.metrics import roc_curve

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))

from creditrisk import binning, data, evaluate as ev, pricing, scorecard as sc  # noqa: E402

OUT, FIG = ROOT / "outputs", ROOT / "outputs" / "figures"
IV_CUTOFF = 0.02
BAND_EDGES = [-np.inf, 520, 530, 540, 550, 560, 570, np.inf]

PURPOSE_NAMES = {
    "debt_consolidation": "Debt consolidation", "credit_card": "Credit card refinancing",
    "home_improvement": "Home improvement", "major_purchase": "Major purchase", "small_business": "Small business",
    "other": "Other", "car": "Car", "medical": "Medical", "wedding": "Wedding", "moving": "Moving",
    "house": "House", "vacation": "Vacation", "educational": "Education", "renewable_energy": "Renewable energy",
}
VERIFICATION_NAMES = {"verified": "Verified", "not verified": "Not verified"}


def band_label(lo, hi) -> str:
    if lo is None or pd.isna(lo):
        return f"Below {int(hi)}"
    if hi is None or pd.isna(hi):
        return f"{int(lo)} and above"
    return f"{int(lo)}-{int(hi) - 1}"


def gbm_auc(tr, te, extra=None):
    feats = data.NUMERIC + data.CATEGORICAL + (extra or [])
    Xtr, Xte = tr[feats].copy(), te[feats].copy()
    for c in data.CATEGORICAL:
        Xtr[c] = Xtr[c].astype("category")
        Xte[c] = pd.Categorical(Xte[c], categories=Xtr[c].cat.categories)
    m = HistGradientBoostingClassifier(max_iter=400, learning_rate=0.05, categorical_features="from_dtype",
                                       random_state=0).fit(Xtr, tr[data.TARGET])
    p = m.predict_proba(Xte)[:, 1]
    return p


def main() -> dict:
    OUT.mkdir(exist_ok=True)
    FIG.mkdir(parents=True, exist_ok=True)
    df = data.load()
    tr, te = data.split(df)
    y_tr, y_te = tr[data.TARGET], te[data.TARGET]

    # ---- 1. bin every candidate variable, screen by Information Value -------
    bins = {v: binning.bin_numeric(tr[v], y_tr, v, **data.BIN_SPEC[v]) for v in data.NUMERIC}
    bins.update({v: binning.bin_categorical(tr[v], y_tr, v) for v in data.CATEGORICAL})
    iv = pd.Series({v: b.iv for v, b in bins.items()}).sort_values(ascending=False)
    selected = [v for v in iv.index if iv[v] >= IV_CUTOFF]

    # ---- 2. fit the scorecard ------------------------------------------------
    card = sc.fit(tr, data.TARGET, {v: bins[v] for v in selected})
    assert all(b < 0 for b in card.coefs.values()), "a WoE coefficient flipped sign"

    pd_tr, pd_te = card.predict_pd(tr), card.predict_pd(te)
    score_tr, score_te = card.score(tr), card.score(te)

    # ---- 3. validation ---------------------------------------------------------
    p_gbm = gbm_auc(tr, te)
    p_gbm_rate = gbm_auc(tr.assign(int_rate=tr["int_rate"]), te.assign(int_rate=te["int_rate"]), ["int_rate"])
    metrics = []
    for name, y, p in [("Scorecard (training set)", y_tr, pd_tr), ("Scorecard (holdout set)", y_te, pd_te),
                       ("Gradient boosting, same variables (holdout)", y_te, p_gbm),
                       ("LendingClub's own interest rate (holdout)", y_te, te["int_rate"]),
                       ("Gradient boosting + LendingClub rate (holdout)", y_te, p_gbm_rate)]:
        metrics.append({"model": name, "auc": ev.auc(y, p), "ks": ev.ks(y, p)})
    calib = ev.calibration_table(y_te, pd_te)
    hl, hl_p = ev.hosmer_lemeshow(y_te, pd_te)
    psi = ev.psi(score_tr.to_numpy(), score_te.to_numpy(), np.arange(440, 621, 10))

    # ---- 4. risk bands on all 36-month loans ---------------------------------
    full = df.copy()
    full["score"] = card.score(full)
    full["pd"] = sc.score_to_pd(full["score"])
    f36 = full[full["term_months"] == 36].copy()
    f36["band"] = pd.cut(f36["score"], BAND_EDGES, right=False)
    bands = f36.groupby("band", observed=True).agg(
        loans=("bad_loan", "size"), bad_rate=("bad_loan", "mean"), avg_score=("score", "mean"),
        avg_pd=("pd", "mean"), lc_rate=("int_rate", "mean")).reset_index()
    bands["score_from"] = [None if np.isinf(i.left) else float(i.left) for i in bands["band"]]
    bands["score_to"] = [None if np.isinf(i.right) else float(i.right) for i in bands["band"]]
    bands["lc_rate"] = bands["lc_rate"] / 100.0
    terms = pricing.LoanTerms(amount=10_000, term=36)
    bands["breakeven_apr"] = [pricing.breakeven_apr(p, terms) for p in bands["avg_pd"]]

    # ---- 5. tables + figures --------------------------------------------------
    card.points.to_csv(OUT / "scorecard_points.csv", index=False)
    card.fit_summary.to_csv(OUT / "logistic_coefficients.csv")
    pd.DataFrame(metrics).to_csv(OUT / "model_metrics.csv", index=False)
    calib.to_csv(OUT / "calibration_holdout.csv", index=False)
    iv.rename("information_value").to_csv(OUT / "information_value.csv")
    bands.drop(columns="band").to_csv(OUT / "risk_bands_36m.csv", index=False)
    make_figures(y_te, pd_te, te["int_rate"], p_gbm, calib, bands, score_te, iv, selected)

    # ---- 6. export the model for the workbook -----------------------------------
    def binning_json(v):
        b = card.binnings[v]
        rows = []
        for _, r in b.table.iterrows():
            row = {"label": r["bin"], "count": int(r["count"]), "bad_rate": float(r["bad_rate"]), "woe": float(r["woe"])}
            if b.kind == "numeric" and r["bin"] != binning.MISSING:
                i = [b.bin_label(k) for k in range(len(b.edges))].index(r["bin"])
                row["upper_edge"] = None if np.isinf(b.edges[i]) else b.edges[i]
            row["missing"] = r["bin"] == binning.MISSING
            rows.append(row)
        out = {"var": v, "label": data.LABELS[v], "kind": b.kind, "unit": b.unit, "coef": card.coefs[v], "bins": rows}
        if b.kind == "categorical":
            out["raw_to_group"] = {k: g for k, g in b.groups.items() if k != "__other__"}
        return out

    model = {
        "meta": {
            "source": "LendingClub public loan history (completed loans), via H2O.ai app-consumer-loan data/loan.csv",
            "source_url": data.DATA_URL, "loans": int(len(df)), "train": int(len(tr)), "test": int(len(te)),
            "bad_rate": float(df[data.TARGET].mean()), "n_vars_screened": int(len(iv)),
        },
        "scaling": {"pdo": sc.PDO, "base_score": sc.BASE_SCORE, "base_odds": sc.BASE_ODDS},
        "intercept": card.intercept,
        "variables": [binning_json(v) for v in selected],
        "iv": [{"var": v, "label": data.LABELS[v], "iv": float(iv[v]), "used": v in selected} for v in iv.index],
        "coefficients": [{"term": k, "coef": float(r.coef), "std_err": float(r.std_err), "p_value": float(r.p_value)}
                         for k, r in card.fit_summary.iterrows()],
        "metrics": metrics,
        "calibration": calib.to_dict(orient="records"),
        "hosmer_lemeshow": {"stat": hl, "p_value": hl_p},
        "psi_train_vs_holdout": psi,
        "bands": [{**{k: (None if (isinstance(v, float) and np.isnan(v)) else v) for k, v in r.items()},
                   "label": band_label(r["score_from"], r["score_to"])}
                  for r in bands.drop(columns="band").to_dict(orient="records")],
        "purpose_names": PURPOSE_NAMES, "verification_names": VERIFICATION_NAMES,
        "score_stats": {"mean": float(score_te.mean()), "min": float(score_te.min()), "max": float(score_te.max())},
    }
    (OUT / "model.json").write_text(json.dumps(model, indent=1, default=float))
    return model


def make_figures(y, p, rate, p_gbm, calib, bands, score, iv, selected):
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    import matplotlib.ticker as mt

    INK, INK2, MUTED, GRID, AXIS, SURF = "#0b0b0b", "#52514e", "#898781", "#e1e0d9", "#c3c2b7", "#fcfcfb"
    BLUE, ORANGE, AQUA = "#2a78d6", "#eb6834", "#1baf7a"
    plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 10, "axes.edgecolor": AXIS,
                         "axes.facecolor": SURF, "figure.facecolor": SURF, "axes.grid": True, "grid.color": GRID,
                         "axes.spines.top": False, "axes.spines.right": False, "xtick.color": MUTED,
                         "ytick.color": MUTED, "axes.labelcolor": INK2, "legend.frameon": False,
                         "axes.titleweight": "bold", "axes.titlesize": 12, "axes.titlelocation": "left",
                         "axes.axisbelow": True})

    # ROC
    fig, ax = plt.subplots(figsize=(5.6, 5))
    for name, s, c in [("Scorecard", p, BLUE), ("Gradient boosting", p_gbm, AQUA), ("LendingClub rate", rate, ORANGE)]:
        fpr, tpr, _ = roc_curve(y, s)
        ax.plot(fpr, tpr, color=c, lw=2, label=f"{name} (AUC {ev.auc(y, s):.3f})")
    ax.plot([0, 1], [0, 1], color=AXIS, lw=1)
    ax.set_xlabel("False positive rate (good loans flagged)"); ax.set_ylabel("True positive rate (bad loans caught)")
    ax.set_title("ROC curve, holdout set"); ax.legend(loc="lower right", fontsize=9)
    fig.tight_layout(); fig.savefig(FIG / "roc_curve.png", dpi=160); plt.close(fig)

    # calibration
    fig, ax = plt.subplots(figsize=(6.4, 4.2))
    x = np.arange(1, len(calib) + 1)
    ax.bar(x - 0.18, calib["predicted_pd"], width=0.34, color=BLUE, label="Predicted PD")
    ax.bar(x + 0.18, calib["observed_bad_rate"], width=0.34, color=ORANGE, label="Observed bad rate")
    ax.set_xticks(x); ax.set_xlabel("Predicted-risk decile (1 = safest)")
    ax.yaxis.set_major_formatter(mt.PercentFormatter(1.0, decimals=0))
    ax.set_title("Calibration: predicted vs. observed default, holdout set", pad=24)
    ax.legend(loc="lower left", bbox_to_anchor=(0, 1.0), ncol=2, fontsize=9, borderaxespad=0.2)
    ax.grid(axis="x", visible=False)
    fig.tight_layout(); fig.savefig(FIG / "calibration.png", dpi=160); plt.close(fig)

    # pricing curve
    fig, ax = plt.subplots(figsize=(7.4, 4.2))
    labels = [band_label(f, t) for f, t in zip(bands["score_from"], bands["score_to"])]
    xb = np.arange(len(bands))
    ax.plot(xb, bands["breakeven_apr"], color=BLUE, lw=2, marker="o", ms=6, mec=SURF, mew=2, label="Break-even APR (this model)")
    ax.plot(xb, bands["lc_rate"], color=ORANGE, lw=2, marker="o", ms=6, mec=SURF, mew=2, label="LendingClub average rate")
    last = len(bands) - 1
    for i, col, dy in ((0, "breakeven_apr", 8), (0, "lc_rate", -16), (last, "lc_rate", 8), (last, "breakeven_apr", -16)):
        v = bands[col].iloc[i]
        ax.annotate(f"{v:.1%}", (xb[i], v), textcoords="offset points", xytext=(0, dy), ha="center", fontsize=9, color=INK)
    ax.set_ylim(bands[["breakeven_apr", "lc_rate"]].min().min() - 0.01, bands[["breakeven_apr", "lc_rate"]].max().max() + 0.01)
    ax.set_xticks(xb, labels); ax.set_xlabel("Credit score band (36-month loans)")
    ax.yaxis.set_major_formatter(mt.PercentFormatter(1.0, decimals=0))
    ax.set_title("Risk-based price vs. what LendingClub charged", pad=24)
    ax.legend(loc="lower left", bbox_to_anchor=(0, 1.0), ncol=2, fontsize=9, borderaxespad=0.2)
    fig.tight_layout(); fig.savefig(FIG / "pricing_vs_lendingclub.png", dpi=160); plt.close(fig)

    # IV
    fig, ax = plt.subplots(figsize=(6.4, 4.4))
    ivs = iv.sort_values()
    ax.barh([data.LABELS[v] for v in ivs.index], ivs.values, height=0.55,
            color=[BLUE if v in selected else AXIS for v in ivs.index])
    ax.axvline(IV_CUTOFF, color=INK2, lw=1)
    ax.text(IV_CUTOFF, 1.0, f" cutoff {IV_CUTOFF}", transform=ax.get_xaxis_transform(), fontsize=9, color=INK2, va="bottom")
    fig.suptitle("Information value by variable (blue = in scorecard)", x=0.02, ha="left", fontweight="bold", fontsize=12)
    ax.grid(axis="y", visible=False)
    fig.tight_layout(); fig.savefig(FIG / "information_value.png", dpi=160); plt.close(fig)


if __name__ == "__main__":
    m = main()
    for r in m["metrics"]:
        print(f"{r['model']:<48} AUC {r['auc']:.3f}  KS {r['ks']:.3f}")
    print(f"Hosmer-Lemeshow {m['hosmer_lemeshow']['stat']:.1f} (p={m['hosmer_lemeshow']['p_value']:.3f}); "
          f"PSI {m['psi_train_vs_holdout']:.4f}")
    for b in m["bands"]:
        print(f"score {b['score_from']}-{b['score_to']}: n={b['loans']:>6} bad={b['bad_rate']:.3f} "
              f"PD={b['avg_pd']:.3f} LC={b['lc_rate']:.3%} BE={b['breakeven_apr']:.3%}")
