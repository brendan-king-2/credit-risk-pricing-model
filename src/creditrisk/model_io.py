"""Score an application straight from outputs/model.json.

This mirrors the Excel scorecard cell for cell (same bins, same rounded points),
so it doubles as the reference when the workbook is verified.
"""
from __future__ import annotations

import json
import math
from pathlib import Path

MODEL_PATH = Path(__file__).resolve().parents[2] / "outputs" / "model.json"


def load_model(path: Path = MODEL_PATH) -> dict:
    return json.loads(Path(path).read_text())


def scaling(model: dict) -> tuple[float, float]:
    s = model["scaling"]
    factor = s["pdo"] / math.log(2)
    offset = s["base_score"] - factor * math.log(s["base_odds"])
    return factor, offset


def bin_points(model: dict, var: dict) -> list[int]:
    factor, offset = scaling(model)
    k = len(model["variables"])
    a = model["intercept"]
    return [_round_half_up(-(var["coef"] * b["woe"] + a / k) * factor + offset / k) for b in var["bins"]]


def _round_half_up(x: float) -> int:
    """Excel's ROUND (halves away from zero)."""
    return int(math.floor(abs(x) + 0.5)) * (1 if x >= 0 else -1)


def score_application(model: dict, app: dict) -> dict:
    """app keys: loan_amnt, term_months, annual_inc, dti (%-points), revol_util (%-points or None),
    purpose (raw code), verification_status (raw code)."""
    app = dict(app)
    app["loan_to_income"] = app["loan_amnt"] / app["annual_inc"]
    total, detail = 0, {}
    for var in model["variables"]:
        pts = bin_points(model, var)
        x = app[var["var"]]
        if var["kind"] == "categorical":
            group = var["raw_to_group"].get(x, "Other (small)")
            i = [b["label"] for b in var["bins"]].index(group)
        elif x is None:
            i = [b["missing"] for b in var["bins"]].index(True)
        else:
            edges = [b["upper_edge"] for b in var["bins"] if not b["missing"] and b["upper_edge"] is not None]
            i = sum(1 for e in edges if e < x)
        total += pts[i]
        detail[var["var"]] = (var["bins"][i]["label"], pts[i])
    factor, offset = scaling(model)
    pd_life = 1.0 / (1.0 + math.exp((total - offset) / factor))
    return {"score": total, "pd": pd_life, "detail": detail}
