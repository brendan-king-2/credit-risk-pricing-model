"""Weight-of-Evidence (WoE) binning with monotonic merging.

WoE convention: WoE = ln(%good / %bad). Positive WoE = safer than average.
Information Value (IV) = sum((%good - %bad) * WoE) measures a variable's
standalone power: <0.02 useless, 0.02-0.1 weak, 0.1-0.3 medium, >0.3 strong.
"""
from __future__ import annotations

from dataclasses import dataclass, field

import numpy as np
import pandas as pd

MISSING = "Missing"


@dataclass
class Binning:
    var: str
    kind: str                      # "numeric" or "categorical"
    edges: list[float] = field(default_factory=list)   # numeric upper edges (right-inclusive), last = inf
    groups: dict[str, str] = field(default_factory=dict)  # categorical raw value -> group label
    table: pd.DataFrame | None = None
    has_missing: bool = False
    unit: str = "num"              # label format: "$", "pct" (ratio shown as %), "num"

    @property
    def iv(self) -> float:
        return float(self.table["iv"].sum())

    def labels(self, x: pd.Series) -> pd.Series:
        if self.kind == "numeric":
            idx = np.searchsorted(np.asarray(self.edges), x.to_numpy(dtype=float), side="left")
            idx = np.minimum(idx, len(self.edges) - 1)   # NaN sorts past the end; relabelled below
            out = pd.Series([self.bin_label(i) for i in idx], index=x.index)
            out[x.isna()] = MISSING
            return out
        out = x.map(self.groups).fillna(self.groups.get("__other__", "Other"))
        out[x.isna()] = MISSING
        return out

    def bin_label(self, i: int) -> str:
        lo = -np.inf if i == 0 else self.edges[i - 1]
        hi = self.edges[i]
        fmt = {"$": lambda v: f"${v:,.0f}", "pct": lambda v: f"{v:.0%}",
               "num": lambda v: f"{v:,.4g}"}[self.unit]
        if np.isinf(lo):
            return f"<= {fmt(hi)}"
        if np.isinf(hi):
            return f"> {fmt(lo)}"
        return f"{fmt(lo)} - {fmt(hi)}"

    def woe(self, x: pd.Series) -> pd.Series:
        m = self.table.set_index("bin")["woe"]
        return self.labels(x).map(m).astype(float)


def _woe_table(labels: pd.Series, y: pd.Series, order: list[str]) -> pd.DataFrame:
    g = pd.DataFrame({"bin": labels, "y": y}).groupby("bin")["y"].agg(["size", "sum"])
    g = g.reindex(order).dropna()
    g.columns = ["count", "bads"]
    g["goods"] = g["count"] - g["bads"]
    tot_g, tot_b = g["goods"].sum(), g["bads"].sum()
    pg = (g["goods"] + 0.5) / tot_g          # +0.5 smoothing guards against empty cells
    pb = (g["bads"] + 0.5) / tot_b
    g["share"] = g["count"] / g["count"].sum()
    g["bad_rate"] = g["bads"] / g["count"]
    g["woe"] = np.log(pg / pb)
    g["iv"] = (pg - pb) * g["woe"]
    return g.reset_index()


def bin_numeric(x: pd.Series, y: pd.Series, var: str, max_bins: int = 10, min_share: float = 0.05,
                round_to: float | None = None, unit: str = "num") -> Binning:
    """Quantile bins (edges rounded to readable values), then merge neighbours until the
    bad rate is monotonic and every bin holds at least `min_share` of loans."""
    xv = x.dropna()
    qs = np.quantile(xv, np.linspace(0, 1, max_bins + 1)[1:-1])
    if round_to:
        qs = np.round(qs / round_to) * round_to
    qs = np.unique(qs[qs < xv.max()])
    edges = list(qs) + [np.inf]

    def stats(edges):
        idx = np.searchsorted(np.asarray(edges), xv.to_numpy(dtype=float), side="left")
        d = pd.DataFrame({"b": idx, "y": y.loc[xv.index].to_numpy()}).groupby("b")["y"].agg(["size", "mean"])
        return d.reindex(range(len(edges)), fill_value=0)

    # direction of the overall relationship
    corr = np.corrcoef(xv, y.loc[xv.index])[0, 1]
    sign = 1 if corr >= 0 else -1
    while len(edges) > 2:
        s = stats(edges)
        share = s["size"] / s["size"].sum()
        br = s["mean"].to_numpy()
        viol = [i for i in range(len(br) - 1) if sign * (br[i + 1] - br[i]) < 0]
        small = [i for i in range(len(br)) if share.iloc[i] < min_share]
        if not viol and not small:
            break
        if viol:
            i = min(viol, key=lambda j: abs(br[j + 1] - br[j]))   # merge the closest violating pair
        else:
            j = small[0]
            i = j if j < len(br) - 1 else j - 1
        del edges[i]  # merging bin i and i+1 = drop the edge between them
    b = Binning(var=var, kind="numeric", edges=[float(e) for e in edges], has_missing=bool(x.isna().any()), unit=unit)
    order = [b.bin_label(i) for i in range(len(edges))] + ([MISSING] if b.has_missing else [])
    b.table = _woe_table(b.labels(x), y, order)
    return b


def bin_categorical(x: pd.Series, y: pd.Series, var: str, min_share: float = 0.02,
                    merge: dict[str, str] | None = None) -> Binning:
    """Categories below min_share are pooled; optional explicit merges for interpretability."""
    share = x.value_counts(normalize=True)
    groups = {}
    for v in share.index:
        lab = (merge or {}).get(v, v)
        groups[v] = lab if share[v] >= min_share or v in (merge or {}) else "Other (small)"
    groups["__other__"] = "Other (small)"
    b = Binning(var=var, kind="categorical", groups=groups, has_missing=bool(x.isna().any()))
    labs = b.labels(x)
    order = labs.groupby(labs).size().sort_values(ascending=False).index.tolist()
    b.table = _woe_table(labs, y, order)
    # order rows by risk for readability
    b.table = b.table.sort_values("bad_rate").reset_index(drop=True)
    return b
