"""Risk-based loan pricing for an amortizing installment loan.

Cash-flow model (the same one the Excel workbook implements):
  * Lifetime PD is spread over the term as a constant monthly default hazard h,
    so that 1 - (1 - h)^n = PD.
  * A loan that survives to month t pays the level installment P, less a
    servicing fee taken as a share of each payment.
  * A loan that defaults in month t misses that payment; the lender recovers
    (1 - LGD) x the balance outstanding, `lag` months later.
  * Cash flows are discounted at the lender's required annual return.

Expected NPV has a closed form (sums of geometric series), so the break-even
APR - the rate at which NPV = 0 - is found by bisection.
"""
from __future__ import annotations

from dataclasses import dataclass

import numpy as np


@dataclass
class LoanTerms:
    amount: float = 12_000.0
    term: int = 36
    lgd: float = 0.90               # loss given default (Gallo 2021: mean LGD 90.2% on LendingClub defaults)
    recovery_lag: int = 6           # months from default to recovery cash
    hurdle: float = 0.06            # lender's required annual return
    servicing: float = 0.01         # fee as share of each payment received
    upfront_fee: float = 0.0        # fee retained by lender at funding, share of amount


def monthly_hazard(pd_life: float, term: int) -> float:
    return 1.0 - (1.0 - pd_life) ** (1.0 / term)


def payment(amount: float, apr: float, term: int) -> float:
    i = apr / 12.0
    return amount * i / (1.0 - (1.0 + i) ** -term)


def _geo(x: float, n: int) -> float:
    """sum_{j=0}^{n-1} x^j"""
    return float(n) if abs(1.0 - x) < 1e-12 else (1.0 - x ** n) / (1.0 - x)


def npv_closed_form(apr: float, pd_life: float, t: LoanTerms) -> float:
    n, A = t.term, t.amount
    h = monthly_hazard(pd_life, n)
    s = 1.0 - h
    rm = (1.0 + t.hurdle) ** (1.0 / 12.0) - 1.0
    d = 1.0 / (1.0 + rm)
    q = s * d
    i = apr / 12.0
    g = 1.0 + i
    P = payment(A, apr, n)
    D = g ** n - 1.0
    Hq = _geo(q, n)
    Hx = _geo(q * g, n)
    performing = P * (1.0 - t.servicing) * q * Hq
    recoveries = h * (1.0 - t.lgd) * d ** t.recovery_lag * A * d / D * (g ** n * Hq - Hx)
    return A * (t.upfront_fee - 1.0) + performing + recoveries


def cash_flow_schedule(apr: float, pd_life: float, t: LoanTerms):
    """Month-by-month expected cash flows (used to cross-check the closed form)."""
    n, A = t.term, t.amount
    h = monthly_hazard(pd_life, n)
    P = payment(A, apr, n)
    i = apr / 12.0
    rows = []
    bal, surv = A, 1.0
    horizon = n + t.recovery_lag
    flows = np.zeros(horizon + 1)
    flows[0] = -A + t.upfront_fee * A
    for m in range(1, n + 1):
        p_default = surv * h
        p_pay = surv * (1 - h)
        interest = bal * i
        principal = P - interest
        flows[m] += p_pay * P * (1 - t.servicing)
        flows[m + t.recovery_lag] += p_default * (1 - t.lgd) * bal
        rows.append({"month": m, "balance_start": bal, "survival_start": surv, "p_default": p_default,
                     "expected_payment": p_pay * P * (1 - t.servicing),
                     "expected_loss": p_default * t.lgd * bal})
        bal -= principal
        surv *= 1 - h
    return rows, flows


def npv_from_schedule(apr: float, pd_life: float, t: LoanTerms) -> float:
    _, flows = cash_flow_schedule(apr, pd_life, t)
    rm = (1.0 + t.hurdle) ** (1.0 / 12.0) - 1.0
    return float(np.sum(flows / (1 + rm) ** np.arange(len(flows))))


def breakeven_apr(pd_life: float, t: LoanTerms, lo: float = 0.0005, hi: float = 1.0, iters: int = 40) -> float:
    for _ in range(iters):
        mid = 0.5 * (lo + hi)
        if npv_closed_form(mid, pd_life, t) < 0:
            lo = mid
        else:
            hi = mid
    return 0.5 * (lo + hi)


def expected_loss(pd_life: float, t: LoanTerms, apr: float) -> dict:
    rows, _ = cash_flow_schedule(apr, pd_life, t)
    p_def = sum(r["p_default"] for r in rows)
    ead = sum(r["p_default"] * r["balance_start"] for r in rows) / p_def if p_def > 0 else 0.0
    el = sum(r["expected_loss"] for r in rows)
    return {"pd": p_def, "lgd": t.lgd, "ead": ead, "el": el, "el_pct": el / t.amount}
