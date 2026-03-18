"""
Base financial model: ratio computation engine for Indian listed companies.
Operates on CompanyData (output of ScreenerClient) and produces normalized
metrics used by all sector-specific models and chart_factory.

All monetary values in INR Crores.
All ratios as floats (percentages as 0–100, not 0–1).
All lists indexed 0 = oldest FY.
"""

import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', '..'))

import logging
import math
from dataclasses import dataclass, field
from typing import TYPE_CHECKING, Optional

if TYPE_CHECKING:
    from src.data.screener_client import CompanyData

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Dataclasses
# ---------------------------------------------------------------------------

@dataclass
class FinancialSeries:
    """Aligned time series for one financial metric."""
    fy_labels: list[str]        # ["FY16", "FY17", ..., "FY25"]
    values: list[float | None]  # aligned to fy_labels, None = missing
    actuals_end_fy: str         # last actual FY, e.g. "FY25"
    units: str = "INR Cr"


@dataclass
class NormalizedFinancials:
    """Output of BaseModel.normalize() — clean, aligned financial data."""
    ticker: str
    company_name: str
    is_consolidated: bool
    fy_labels: list[str]
    actuals_end_fy: str

    # Income statement (all INR Cr)
    gross_revenue: list[float | None]
    excise_duty: list[float | None]
    net_revenue: list[float | None]
    raw_materials: list[float | None]
    gross_profit: list[float | None]
    other_opex: list[float | None]
    ebitda: list[float | None]
    depreciation: list[float | None]
    ebit: list[float | None]
    other_income: list[float | None]
    finance_costs: list[float | None]
    pbt: list[float | None]
    tax: list[float | None]
    pat: list[float | None]
    minority_interest: list[float | None]
    pat_after_mi: list[float | None]
    eps: list[float | None]
    dps: list[float | None]

    # Balance sheet
    total_assets: list[float | None]
    fixed_assets: list[float | None]
    cwip: list[float | None]
    investments: list[float | None]
    inventory: list[float | None]
    trade_receivables: list[float | None]
    cash: list[float | None]
    other_current_assets: list[float | None]
    total_equity: list[float | None]
    long_term_debt: list[float | None]
    short_term_debt: list[float | None]
    total_debt: list[float | None]
    trade_payables: list[float | None]
    other_current_liabilities: list[float | None]
    net_debt: list[float | None]   # total_debt - cash

    # Cash flow
    cfo: list[float | None]
    capex: list[float | None]
    fcf: list[float | None]        # cfo - capex
    cfi: list[float | None]
    cff: list[float | None]
    dividends_paid: list[float | None]


@dataclass
class ComputedRatios:
    """All computed ratios, aligned to same fy_labels as NormalizedFinancials."""
    fy_labels: list[str]

    # Profitability (%)
    gross_margin: list[float | None]
    ebitda_margin: list[float | None]
    ebit_margin: list[float | None]
    pat_margin: list[float | None]

    # Returns (%)
    roce: list[float | None]   # NOPAT / Capital Employed
    roe: list[float | None]    # PAT / Average Equity
    roa: list[float | None]    # PAT / Average Total Assets

    # Working capital (days)
    dso: list[float | None]    # Receivables days
    dio: list[float | None]    # Inventory days
    dpo: list[float | None]    # Payable days
    ccc: list[float | None]    # Cash Conversion Cycle

    # Leverage
    debt_to_equity: list[float | None]
    net_debt_to_ebitda: list[float | None]
    interest_coverage: list[float | None]

    # Growth (%)
    revenue_growth: list[float | None]
    ebitda_growth: list[float | None]
    pat_growth: list[float | None]
    eps_growth: list[float | None]

    # Quality
    fcf_conversion: list[float | None]  # FCF / PAT
    cfo_pat_ratio: list[float | None]   # CFO / PAT

    # CAGRs (computed for full period, not per-year)
    revenue_cagr_3y: float | None
    revenue_cagr_5y: float | None
    revenue_cagr_10y: float | None
    pat_cagr_3y: float | None
    pat_cagr_5y: float | None


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def safe_div(a: float | None, b: float | None) -> float | None:
    """Return a / b, or None if b is zero or either operand is None."""
    if a is None or b is None:
        return None
    if b == 0:
        return None
    return a / b


def _to_float(value) -> float | None:
    """Convert a raw value to float, returning None for NaN / missing."""
    if value is None:
        return None
    try:
        f = float(value)
        if math.isnan(f) or math.isinf(f):
            return None
        return f
    except (TypeError, ValueError):
        return None


def _pairwise_op(a_list: list[float | None],
                 b_list: list[float | None],
                 op) -> list[float | None]:
    """Apply a binary operator element-wise to two aligned lists."""
    result = []
    for a, b in zip(a_list, b_list):
        if a is None or b is None:
            result.append(None)
        else:
            try:
                result.append(op(a, b))
            except (ZeroDivisionError, ValueError):
                result.append(None)
    return result


def _average_adjacent(values: list[float | None]) -> list[float | None]:
    """
    Return list of averages of adjacent pairs (used for average balance-sheet
    values).  Output has same length as input; index 0 returns values[0] itself
    (no prior period).
    """
    result = [values[0]]
    for i in range(1, len(values)):
        a, b = values[i - 1], values[i]
        if a is None or b is None:
            result.append(None)
        else:
            result.append((a + b) / 2.0)
    return result


def _yoy_growth(values: list[float | None]) -> list[float | None]:
    """Year-on-year growth in %. Index 0 is None (no prior year)."""
    result = [None]
    for i in range(1, len(values)):
        prev, curr = values[i - 1], values[i]
        if prev is None or curr is None or prev == 0:
            result.append(None)
        else:
            result.append((curr / prev - 1.0) * 100.0)
    return result


# ---------------------------------------------------------------------------
# Screener.in label → field mapping
# ---------------------------------------------------------------------------

# P&L row label variants → NormalizedFinancials field name
PNL_LABEL_MAP: dict[str, str] = {
    "sales +": "gross_revenue",
    "revenue": "gross_revenue",
    "net sales": "net_revenue",
    "excise duty": "excise_duty",
    "raw material": "raw_materials",
    "material cost %": "raw_materials",      # handled separately (percentage)
    "power & fuel": "_power_fuel",
    "other mfr. exp": "_other_mfr",
    "employee cost %": "_employee_cost",
    "selling & admin expenses %": "_selling_admin",
    "operating profit": "ebitda",
    "opm %": "_opm_pct",                     # check only
    "other income +": "other_income",
    "depreciation": "depreciation",
    "interest": "finance_costs",
    "profit before tax": "pbt",
    "tax %": "_tax_pct",                     # percentage, convert to amount
    "net profit +": "pat_after_mi",
    "eps in rs": "eps",
    "dividend payout %": "_div_payout_pct",
    "dividend / share": "dps",
}

BALANCE_SHEET_LABEL_MAP: dict[str, str] = {
    "share capital": "_share_capital",
    "reserves": "_reserves",
    "borrowings": "total_debt",
    "other liabilities": "other_current_liabilities",
    "trade payables": "trade_payables",
    "fixed assets": "fixed_assets",
    "cwip": "cwip",
    "investments": "investments",
    "trade receivables": "trade_receivables",
    "inventory": "inventory",
    "cash equivalents": "cash",
    "other assets": "other_current_assets",
    "total assets": "total_assets",
    "total liabilities": "total_assets",     # alias
}

CASHFLOW_LABEL_MAP: dict[str, str] = {
    "cash from operating activity +": "cfo",
    "cash from operating activity": "cfo",
    "cash from investing activity": "cfi",
    "cash from financing activity": "cff",
    "dividends paid": "dividends_paid",
    "capex": "capex",
    "fixed assets purchased": "capex",
    "purchase of fixed assets": "capex",
}


# ---------------------------------------------------------------------------
# BaseModel
# ---------------------------------------------------------------------------

class BaseModel:
    """
    Core ratio computation engine for Indian listed companies.
    Accepts a CompanyData object (from ScreenerClient) and produces
    NormalizedFinancials and ComputedRatios.
    """

    def __init__(self, company_data: "CompanyData") -> None:
        self.company_data = company_data
        self._nf: NormalizedFinancials | None = None
        self._ratios: ComputedRatios | None = None

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def normalize(self) -> NormalizedFinancials:
        """
        Map screener.in DataFrame row labels to NormalizedFinancials fields.
        Handles missing fields gracefully (sets to None, logs warning).
        Derives composite lines where needed.
        """
        cd = self.company_data

        # Determine FY labels from P&L columns
        fy_labels = self._extract_fy_labels()
        n = len(fy_labels)
        null_series: list[float | None] = [None] * n

        # ---- Income Statement ----
        gross_revenue  = self._extract_series("pnl", ["sales +", "revenue"], fy_labels)
        net_revenue    = self._extract_series("pnl", ["net sales"], fy_labels)
        excise_duty    = self._extract_series("pnl", ["excise duty"], fy_labels)
        raw_materials  = self._extract_series("pnl", ["raw material", "material cost %"], fy_labels, is_pct_of="net_revenue")
        other_income   = self._extract_series("pnl", ["other income +"], fy_labels)
        depreciation   = self._extract_series("pnl", ["depreciation"], fy_labels)
        finance_costs  = self._extract_series("pnl", ["interest"], fy_labels)
        pbt            = self._extract_series("pnl", ["profit before tax"], fy_labels)
        pat_after_mi   = self._extract_series("pnl", ["net profit +"], fy_labels)
        eps            = self._extract_series("pnl", ["eps in rs"], fy_labels)
        dps            = self._extract_series("pnl", ["dividend / share", "dps"], fy_labels)

        # Tax: if stored as %, convert to absolute amount
        tax_pct = self._extract_series("pnl", ["tax %"], fy_labels)
        tax_abs = self._extract_series("pnl", ["tax"], fy_labels)
        tax = self._resolve_tax(tax_abs, tax_pct, pbt)

        # PAT = PAT after MI for most companies; if minority interest is separate:
        mi   = self._extract_series("pnl", ["minority interest", "share of profit from associates"], fy_labels)
        pat  = _pairwise_op(pat_after_mi, mi,
                            lambda a, b: a + b if b else a)  # pat = pat_after_mi + MI

        # ebitda from screener (includes OI per screener convention)
        ebitda_screener = self._extract_series("pnl", ["operating profit"], fy_labels)

        # Derive ebit = pbt + finance_costs - other_income
        ebit = self._derive_ebit(pbt, finance_costs, other_income)

        # Derive ebitda = ebit + depreciation
        ebitda = self._derive_ebitda(ebit, depreciation, ebitda_screener, other_income)

        # Derive gross_profit = net_revenue - raw_materials
        gross_profit = _pairwise_op(net_revenue, raw_materials, lambda a, b: a - b)

        # other_opex components
        power_fuel    = self._extract_series("pnl", ["power & fuel"], fy_labels)
        other_mfr     = self._extract_series("pnl", ["other mfr. exp"], fy_labels)
        emp_cost      = self._extract_series("pnl", ["employee cost %"], fy_labels, is_pct_of="net_revenue")
        selling_admin = self._extract_series("pnl", ["selling & admin expenses %"], fy_labels, is_pct_of="net_revenue")
        other_opex    = self._sum_series([power_fuel, other_mfr, emp_cost, selling_admin], n)

        # net_revenue fallback: gross_revenue - excise_duty
        net_revenue = self._resolve_net_revenue(gross_revenue, net_revenue, excise_duty)

        # ---- Balance Sheet ----
        share_capital = self._extract_series("balance_sheet", ["share capital"], fy_labels)
        reserves      = self._extract_series("balance_sheet", ["reserves"], fy_labels)
        total_equity  = _pairwise_op(share_capital, reserves, lambda a, b: a + b)
        if all(v is None for v in total_equity):
            total_equity = self._extract_series("balance_sheet", ["total equity", "equity"], fy_labels)

        total_debt    = self._extract_series("balance_sheet", ["borrowings"], fy_labels)
        long_term_debt  = self._extract_series("balance_sheet", ["long term borrowings", "long-term borrowings"], fy_labels)
        short_term_debt = self._extract_series("balance_sheet", ["short term borrowings", "short-term borrowings"], fy_labels)

        # Derive long_term_debt / short_term_debt from total if not separately available
        if all(v is None for v in long_term_debt) and all(v is None for v in short_term_debt):
            long_term_debt  = null_series[:]
            short_term_debt = total_debt[:]

        fixed_assets          = self._extract_series("balance_sheet", ["fixed assets"], fy_labels)
        cwip                  = self._extract_series("balance_sheet", ["cwip"], fy_labels)
        investments           = self._extract_series("balance_sheet", ["investments"], fy_labels)
        trade_receivables     = self._extract_series("balance_sheet", ["trade receivables"], fy_labels)
        inventory             = self._extract_series("balance_sheet", ["inventory"], fy_labels)
        cash                  = self._extract_series("balance_sheet", ["cash equivalents"], fy_labels)
        other_current_assets  = self._extract_series("balance_sheet", ["other assets"], fy_labels)
        trade_payables        = self._extract_series("balance_sheet", ["trade payables"], fy_labels)
        other_current_liab    = self._extract_series("balance_sheet", ["other liabilities"], fy_labels)
        total_assets          = self._extract_series("balance_sheet", ["total assets", "total liabilities"], fy_labels)

        net_debt = _pairwise_op(total_debt, cash, lambda d, c: d - c)

        # ---- Cash Flow ----
        cfo           = self._extract_series("cash_flow", ["cash from operating activity +", "cash from operating activity"], fy_labels)
        cfi           = self._extract_series("cash_flow", ["cash from investing activity"], fy_labels)
        cff           = self._extract_series("cash_flow", ["cash from financing activity"], fy_labels)
        dividends_paid = self._extract_series("cash_flow", ["dividends paid"], fy_labels)

        # Capex: look for explicit line, else derive from cfi components
        capex_raw = self._extract_series("cash_flow", ["capex", "fixed assets purchased", "purchase of fixed assets"], fy_labels)
        # Ensure capex is negative (outflow); abs will be used for FCF
        capex = [(-abs(v) if v is not None else None) for v in capex_raw]

        # FCF = CFO - |capex|
        fcf = _pairwise_op(cfo, capex, lambda o, k: o - abs(k))

        # actuals_end_fy: last FY label (assumes data is up to latest actual)
        actuals_end_fy = fy_labels[-1] if fy_labels else "Unknown"

        nf = NormalizedFinancials(
            ticker          = getattr(cd, "ticker", "UNKNOWN"),
            company_name    = getattr(cd, "company_name", ""),
            is_consolidated = getattr(cd, "is_consolidated", True),
            fy_labels       = fy_labels,
            actuals_end_fy  = actuals_end_fy,
            gross_revenue   = gross_revenue,
            excise_duty     = excise_duty,
            net_revenue     = net_revenue,
            raw_materials   = raw_materials,
            gross_profit    = gross_profit,
            other_opex      = other_opex,
            ebitda          = ebitda,
            depreciation    = depreciation,
            ebit            = ebit,
            other_income    = other_income,
            finance_costs   = finance_costs,
            pbt             = pbt,
            tax             = tax,
            pat             = pat,
            minority_interest = mi,
            pat_after_mi    = pat_after_mi,
            eps             = eps,
            dps             = dps,
            total_assets    = total_assets,
            fixed_assets    = fixed_assets,
            cwip            = cwip,
            investments     = investments,
            inventory       = inventory,
            trade_receivables = trade_receivables,
            cash            = cash,
            other_current_assets = other_current_assets,
            total_equity    = total_equity,
            long_term_debt  = long_term_debt,
            short_term_debt = short_term_debt,
            total_debt      = total_debt,
            trade_payables  = trade_payables,
            other_current_liabilities = other_current_liab,
            net_debt        = net_debt,
            cfo             = cfo,
            capex           = capex,
            fcf             = fcf,
            cfi             = cfi,
            cff             = cff,
            dividends_paid  = dividends_paid,
        )
        self._nf = nf
        return nf

    def compute_ratios(self, nf: NormalizedFinancials) -> ComputedRatios:
        """Compute all standard financial ratios from NormalizedFinancials."""
        n = len(nf.fy_labels)

        # ---- Profitability margins ----
        gross_margin   = [safe_div(gp, nr) and safe_div(gp, nr) * 100
                          for gp, nr in zip(nf.gross_profit, nf.net_revenue)]
        gross_margin   = [safe_div(gp, nr) * 100 if (gp is not None and nr) else None
                          for gp, nr in zip(nf.gross_profit, nf.net_revenue)]
        ebitda_margin  = [safe_div(e, nr) * 100 if (e is not None and nr) else None
                          for e, nr in zip(nf.ebitda, nf.net_revenue)]
        ebit_margin    = [safe_div(e, nr) * 100 if (e is not None and nr) else None
                          for e, nr in zip(nf.ebit, nf.net_revenue)]
        pat_margin     = [safe_div(p, nr) * 100 if (p is not None and nr) else None
                          for p, nr in zip(nf.pat_after_mi, nf.net_revenue)]

        # ---- Returns ----
        tax_rates = self.get_effective_tax_rate(nf)
        cap_employed  = self.get_capital_employed(nf)
        avg_cap_emp   = _average_adjacent(cap_employed)
        avg_equity    = _average_adjacent(nf.total_equity)
        avg_assets    = _average_adjacent(nf.total_assets)

        nopat = [safe_div(e * (1 - t / 100.0), 1)
                 if (e is not None and t is not None) else None
                 for e, t in zip(nf.ebit, tax_rates)]

        roce = [safe_div(nop, ace) * 100 if (nop is not None and ace) else None
                for nop, ace in zip(nopat, avg_cap_emp)]
        roe  = [safe_div(p, ae) * 100 if (p is not None and ae) else None
                for p, ae in zip(nf.pat_after_mi, avg_equity)]
        roa  = [safe_div(p, aa) * 100 if (p is not None and aa) else None
                for p, aa in zip(nf.pat_after_mi, avg_assets)]

        # ---- Working Capital (days) ----
        # Use net_revenue / 365 as daily revenue base
        daily_rev = [safe_div(nr, 365) for nr in nf.net_revenue]
        # Use COGS (raw_materials) / 365 as daily cost base for DIO and DPO
        daily_cogs = [safe_div(rm, 365) for rm in nf.raw_materials]

        dso = [safe_div(tr, dr) if (tr is not None and dr) else None
               for tr, dr in zip(nf.trade_receivables, daily_rev)]
        dio = [safe_div(inv, dc) if (inv is not None and dc) else None
               for inv, dc in zip(nf.inventory, daily_cogs)]
        dpo = [safe_div(tp, dc) if (tp is not None and dc) else None
               for tp, dc in zip(nf.trade_payables, daily_cogs)]
        ccc = [None if any(v is None for v in [d, di, dp]) else d + di - dp
               for d, di, dp in zip(dso, dio, dpo)]

        # ---- Leverage ----
        debt_to_equity = [safe_div(td, eq) if (td is not None and eq) else None
                          for td, eq in zip(nf.total_debt, nf.total_equity)]
        net_debt_to_ebitda = [safe_div(nd, eb) if (nd is not None and eb) else None
                               for nd, eb in zip(nf.net_debt, nf.ebitda)]
        interest_coverage  = [safe_div(eb, fc) if (eb is not None and fc) else None
                               for eb, fc in zip(nf.ebitda, nf.finance_costs)]

        # ---- Growth ----
        revenue_growth = _yoy_growth(nf.net_revenue)
        ebitda_growth  = _yoy_growth(nf.ebitda)
        pat_growth     = _yoy_growth(nf.pat_after_mi)
        eps_growth     = _yoy_growth(nf.eps)

        # ---- Quality ----
        fcf_conversion = [safe_div(f, p) * 100 if (f is not None and p) else None
                          for f, p in zip(nf.fcf, nf.pat_after_mi)]
        cfo_pat_ratio  = [safe_div(c, p) * 100 if (c is not None and p) else None
                          for c, p in zip(nf.cfo, nf.pat_after_mi)]

        # ---- CAGRs ----
        rev_cagr_3y  = self.compute_cagr(nf.net_revenue, 3)
        rev_cagr_5y  = self.compute_cagr(nf.net_revenue, 5)
        rev_cagr_10y = self.compute_cagr(nf.net_revenue, 10)
        pat_cagr_3y  = self.compute_cagr(nf.pat_after_mi, 3)
        pat_cagr_5y  = self.compute_cagr(nf.pat_after_mi, 5)

        ratios = ComputedRatios(
            fy_labels          = nf.fy_labels,
            gross_margin       = gross_margin,
            ebitda_margin      = ebitda_margin,
            ebit_margin        = ebit_margin,
            pat_margin         = pat_margin,
            roce               = roce,
            roe                = roe,
            roa                = roa,
            dso                = dso,
            dio                = dio,
            dpo                = dpo,
            ccc                = ccc,
            debt_to_equity     = debt_to_equity,
            net_debt_to_ebitda = net_debt_to_ebitda,
            interest_coverage  = interest_coverage,
            revenue_growth     = revenue_growth,
            ebitda_growth      = ebitda_growth,
            pat_growth         = pat_growth,
            eps_growth         = eps_growth,
            fcf_conversion     = fcf_conversion,
            cfo_pat_ratio      = cfo_pat_ratio,
            revenue_cagr_3y    = rev_cagr_3y,
            revenue_cagr_5y    = rev_cagr_5y,
            revenue_cagr_10y   = rev_cagr_10y,
            pat_cagr_3y        = pat_cagr_3y,
            pat_cagr_5y        = pat_cagr_5y,
        )
        self._ratios = ratios
        return ratios

    def compute_cagr(self, values: list[float | None], n_years: int) -> float | None:
        """
        CAGR over last n_years using last non-None value and value n_years prior.
        Formula: (end/start)^(1/n) - 1, expressed as percentage.
        Returns None if insufficient data or non-positive base.
        """
        if not values or n_years <= 0:
            return None

        # Find end value: last non-None
        end_val: float | None = None
        end_idx: int = -1
        for i in range(len(values) - 1, -1, -1):
            if values[i] is not None:
                end_val = values[i]
                end_idx = i
                break
        if end_val is None:
            return None

        # Find start value: n_years before end_idx
        start_idx = end_idx - n_years
        if start_idx < 0:
            return None
        start_val = values[start_idx]
        if start_val is None or start_val <= 0:
            return None
        if end_val <= 0:
            return None

        cagr = (end_val / start_val) ** (1.0 / n_years) - 1.0
        return cagr * 100.0

    def get_effective_tax_rate(self, nf: NormalizedFinancials) -> list[float | None]:
        """
        PAT / PBT for each year, converted to effective tax rate (%).
        Capped at 40%, floored at 0%.
        """
        result: list[float | None] = []
        for pat, pbt in zip(nf.pat_after_mi, nf.pbt):
            if pbt is None or pbt == 0 or pat is None:
                result.append(None)
                continue
            # tax amount = pbt - pat; rate = tax / pbt
            tax_amount = pbt - pat
            rate = (tax_amount / pbt) * 100.0
            rate = max(0.0, min(40.0, rate))
            result.append(rate)
        return result

    def get_capital_employed(self, nf: NormalizedFinancials) -> list[float | None]:
        """
        Capital Employed = Total Assets - Current Liabilities
        Approximated as: Equity + Long-term Debt (when current liabilities unavailable).
        """
        result: list[float | None] = []
        for eq, ltd, ta, ocl, tp in zip(
            nf.total_equity,
            nf.long_term_debt,
            nf.total_assets,
            nf.other_current_liabilities,
            nf.trade_payables,
        ):
            # Prefer TA - current liabilities if available
            if ta is not None and (ocl is not None or tp is not None):
                cl = (ocl or 0.0) + (tp or 0.0)
                result.append(ta - cl)
            elif eq is not None and ltd is not None:
                result.append(eq + ltd)
            elif eq is not None:
                result.append(eq)
            else:
                result.append(None)
        return result

    # ------------------------------------------------------------------
    # Private helpers
    # ------------------------------------------------------------------

    def _extract_fy_labels(self) -> list[str]:
        """Extract FY labels from company_data DataFrames."""
        cd = self.company_data
        for attr in ("profit_loss", "pnl", "income_statement"):
            df = getattr(cd, attr, None)
            if df is not None and hasattr(df, "columns"):
                cols = [str(c) for c in df.columns if str(c).startswith("FY") or str(c).startswith("Mar")]
                if cols:
                    return self._normalize_fy_labels(cols)
        # Fallback: check balance sheet
        for attr in ("balance_sheet",):
            df = getattr(cd, attr, None)
            if df is not None and hasattr(df, "columns"):
                cols = [str(c) for c in df.columns if str(c).startswith("FY") or str(c).startswith("Mar")]
                if cols:
                    return self._normalize_fy_labels(cols)
        logger.warning("Could not extract FY labels from company_data")
        return []

    def _normalize_fy_labels(self, raw_labels: list[str]) -> list[str]:
        """
        Convert screener.in column headers to standard FY labels.
        e.g. "Mar 2016" → "FY16", "Mar 2025" → "FY25"
        """
        normalized = []
        for label in raw_labels:
            label = label.strip()
            if label.startswith("FY"):
                normalized.append(label)
            elif label.startswith("Mar "):
                year = label.replace("Mar ", "").strip()
                normalized.append(f"FY{year[-2:]}")
            else:
                normalized.append(label)
        return normalized

    def _get_df(self, statement: str):
        """Retrieve the appropriate DataFrame from company_data."""
        cd = self.company_data
        if statement == "pnl":
            for attr in ("profit_loss", "pnl", "income_statement", "pl"):
                df = getattr(cd, attr, None)
                if df is not None:
                    return df
        elif statement == "balance_sheet":
            for attr in ("balance_sheet", "bs"):
                df = getattr(cd, attr, None)
                if df is not None:
                    return df
        elif statement == "cash_flow":
            for attr in ("cash_flow", "cf", "cashflow"):
                df = getattr(cd, attr, None)
                if df is not None:
                    return df
        return None

    def _extract_series(
        self,
        statement: str,
        label_variants: list[str],
        fy_labels: list[str],
        is_pct_of: str | None = None,
    ) -> list[float | None]:
        """
        Extract a time series from a screener.in DataFrame.
        Tries each label variant (case-insensitive).
        If is_pct_of is set, the row contains percentages and will be
        converted to absolute values using the named field (resolved later).
        """
        df = self._get_df(statement)
        if df is None:
            return [None] * len(fy_labels)

        # Build a lowercase-keyed row index
        row_index: dict[str, any] = {}
        for idx in df.index:
            row_index[str(idx).strip().lower()] = idx

        for variant in label_variants:
            key = variant.lower()
            if key in row_index:
                raw_idx = row_index[key]
                row = df.loc[raw_idx]
                series: list[float | None] = []
                for fy in fy_labels:
                    # Try FY label, then "Mar YYYY" format
                    val = None
                    for col_key in self._fy_to_col_keys(fy, df.columns):
                        if col_key in df.columns:
                            val = _to_float(row[col_key])
                            break
                    series.append(val)
                return series

        # Not found — return null series silently (caller handles)
        logger.debug("Label not found in %s: tried %s", statement, label_variants)
        return [None] * len(fy_labels)

    def _fy_to_col_keys(self, fy_label: str, columns) -> list[str]:
        """Generate possible column key formats for a given FY label."""
        keys = [fy_label]
        if fy_label.startswith("FY"):
            yy = fy_label[2:]
            # Handle both 2-digit and 4-digit year
            if len(yy) == 2:
                year4 = "20" + yy if int(yy) <= 30 else "19" + yy
                keys.append(f"Mar {year4}")
                keys.append(f"Mar-{year4}")
            elif len(yy) == 4:
                keys.append(f"Mar {yy}")
        return keys

    def _resolve_tax(
        self,
        tax_abs: list[float | None],
        tax_pct: list[float | None],
        pbt: list[float | None],
    ) -> list[float | None]:
        """Return absolute tax amounts, deriving from percentage if needed."""
        result: list[float | None] = []
        for ta, tp, pb in zip(tax_abs, tax_pct, pbt):
            if ta is not None:
                result.append(ta)
            elif tp is not None and pb is not None:
                result.append(abs(tp / 100.0) * pb)
            else:
                result.append(None)
        return result

    def _derive_ebit(
        self,
        pbt: list[float | None],
        finance_costs: list[float | None],
        other_income: list[float | None],
    ) -> list[float | None]:
        """EBIT = PBT + Finance Costs - Other Income."""
        result: list[float | None] = []
        for pb, fc, oi in zip(pbt, finance_costs, other_income):
            if pb is None:
                result.append(None)
                continue
            fc_ = fc or 0.0
            oi_ = oi or 0.0
            result.append(pb + fc_ - oi_)
        return result

    def _derive_ebitda(
        self,
        ebit: list[float | None],
        depreciation: list[float | None],
        ebitda_screener: list[float | None],
        other_income: list[float | None],
    ) -> list[float | None]:
        """
        Prefer derived EBITDA = EBIT + Depreciation.
        Screener's "Operating Profit" includes other income, so we prefer derivation.
        Fall back to screener value minus other_income if EBIT unavailable.
        """
        result: list[float | None] = []
        for ebit_v, dep, ebitda_s, oi in zip(ebit, depreciation, ebitda_screener, other_income):
            if ebit_v is not None and dep is not None:
                result.append(ebit_v + dep)
            elif ebitda_s is not None:
                # screener includes OI; subtract to get operational EBITDA
                oi_ = oi or 0.0
                result.append(ebitda_s - oi_)
            else:
                result.append(None)
        return result

    def _resolve_net_revenue(
        self,
        gross_revenue: list[float | None],
        net_revenue: list[float | None],
        excise_duty: list[float | None],
    ) -> list[float | None]:
        """
        If net_revenue is all None, derive as gross_revenue - excise_duty.
        Post-GST (FY18+), excise_duty = 0, so net_revenue ≈ gross_revenue.
        """
        if any(v is not None for v in net_revenue):
            return net_revenue
        # Derive from gross - excise
        result: list[float | None] = []
        for gr, ex in zip(gross_revenue, excise_duty):
            if gr is not None:
                result.append(gr - (ex or 0.0))
            else:
                result.append(None)
        return result

    def _sum_series(
        self,
        series_list: list[list[float | None]],
        n: int,
    ) -> list[float | None]:
        """Element-wise sum of multiple series; None if all components are None."""
        result: list[float | None] = []
        for i in range(n):
            values = [s[i] for s in series_list if s and i < len(s) and s[i] is not None]
            result.append(sum(values) if values else None)
        return result
