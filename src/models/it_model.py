"""
IT Services sector financial model.
Extends BaseModel for Indian IT companies.

Key differences from FMCGModel:
- Primary expense is Employee Cost (60-70% of revenue)
- EBIT margin is key metric (D&A is negligible, so EBIT ≈ EBITDA - small)
- Revenue by vertical and geography (may need PDF parser for full detail)
- USD revenue alongside INR (conversion at average annual rate)
- Asset-light: no significant capex; FCF conversion > 90%
- Valuation via P/E (primary), EV/EBIT, DCF
- Workforce metrics: headcount, utilization, attrition

Covers: TCS, INFY, WIPRO, HCLTECH, TECHM, LTIM
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

from src.models.base_model import BaseModel, NormalizedFinancials, ComputedRatios, safe_div
from src.models.fmcg_model import compute_wacc

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Dataclasses
# ---------------------------------------------------------------------------

@dataclass
class ITMetrics:
    fy_labels: list[str]

    # Standard P&L
    net_revenue: list[float | None]         # INR Cr
    employee_cost: list[float | None]       # INR Cr
    other_opex: list[float | None]          # INR Cr
    ebitda: list[float | None]
    depreciation: list[float | None]
    ebit: list[float | None]
    other_income: list[float | None]
    finance_costs: list[float | None]
    pbt: list[float | None]
    pat: list[float | None]
    eps: list[float | None]

    # IT-specific margins
    ebitda_margin: list[float | None]       # %
    ebit_margin: list[float | None]         # %
    employee_cost_pct: list[float | None]   # % of revenue

    # Cash flow (IT is FCF-rich)
    cfo: list[float | None]
    capex: list[float | None]
    fcf: list[float | None]
    fcf_conversion: list[float | None]      # FCF/PAT %

    # Dividends
    dps: list[float | None]
    payout_ratio: list[float | None]        # including buybacks where disclosed

    # Returns
    roce: list[float | None]
    roe: list[float | None]
    roa: list[float | None]

    # Workforce (from screener ratios or PDF — may be None)
    headcount: list[int | None]
    revenue_per_employee_lakh: list[float | None]   # INR Lakhs
    attrition_pct: list[float | None]

    # Revenue mix (if available — may be empty dicts)
    vertical_mix: dict[str, list[float | None]]     # {"BFSI": [pct list]}
    geo_mix: dict[str, list[float | None]]          # {"Americas": [pct list]}

    peers: list[str]


@dataclass
class ITValuationInputs:
    ticker: str
    fy_labels_hist: list[str]
    fy_labels_proj: list[str]

    net_revenue_hist: list[float]
    ebit_hist: list[float]
    pat_hist: list[float]
    fcf_hist: list[float]
    eps_hist: list[float]

    net_revenue_proj: list[float | None]
    ebit_proj: list[float | None]
    pat_proj: list[float | None]
    fcf_proj: list[float | None]
    eps_proj: list[float | None]

    wacc: float
    terminal_growth: float
    net_cash_cr: float              # typically large positive for IT companies
    shares_outstanding_cr: float

    peer_tickers: list[str]


# ---------------------------------------------------------------------------
# IT-specific label maps
# ---------------------------------------------------------------------------

IT_PNL_LABELS = {
    "revenue from operations": "net_revenue",
    "revenue from contracts": "net_revenue",
    "net revenue": "net_revenue",
    "sales": "net_revenue",
    "employee benefit expenses": "employee_cost",
    "staff costs": "employee_cost",
    "other expenses": "other_opex",
    "depreciation and amortisation": "depreciation",
    "finance costs": "finance_costs",
    "other income": "other_income",
    "profit before tax": "pbt",
    "net profit": "pat",
}

IT_SECTOR_BETAS = {
    "TCS": 0.55, "INFY": 0.65, "WIPRO": 0.70,
    "HCLTECH": 0.72, "TECHM": 0.80, "LTIM": 0.75,
    "DEFAULT": 0.68,
}


# ---------------------------------------------------------------------------
# ITModel
# ---------------------------------------------------------------------------

class ITModel(BaseModel):
    """
    IT Services sector financial model.
    Extends BaseModel with employee cost analysis, EBIT margin focus,
    high FCF conversion, and P/E / EV/EBIT / DCF valuation scaffolding.
    """

    SECTOR = "IT"
    PEERS = ["TCS", "INFY", "WIPRO", "HCLTECH", "TECHM", "LTIM"]

    def __init__(
        self,
        company_data: "CompanyData",
        market_data: dict = None,
    ) -> None:
        """
        Parameters
        ----------
        company_data : CompanyData
            Output of ScreenerClient; must have profit_loss, balance_sheet,
            cash_flow DataFrames and ticker attribute.
        market_data : dict, optional
            May contain 'beta', 'market_cap_cr', 'shares_outstanding', 'price'.
        """
        super().__init__(company_data)
        self.market_data: dict = market_data or {}
        self._it_metrics: ITMetrics | None = None

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def compute_it_metrics(self) -> ITMetrics:
        """
        Full pipeline:
        1. normalize() from base_model
        2. Override P&L extraction with IT-specific labels
        3. Compute employee_cost_pct = employee_cost / net_revenue
        4. Compute EBIT margin (screener EBIT = PBT + Finance Costs - Other Income)
        5. FCF = CFO - Capex (IT companies have low capex, high FCF)
        6. Try to extract headcount from ratios DataFrame
        7. Revenue per employee = net_revenue (Cr) * 1e7 / (headcount * 1e5) in Lakhs
        8. Compute standard ratios via base compute_ratios()
        """
        nf = self.normalize()
        ratios = self.compute_ratios(nf)

        # IT-specific P&L extraction
        pnl_data = self._extract_it_pnl(nf)

        net_revenue = pnl_data.get("net_revenue", nf.net_revenue)
        employee_cost = pnl_data.get("employee_cost", [None] * len(nf.fy_labels))
        other_opex = pnl_data.get("other_opex", nf.other_opex)
        depreciation = pnl_data.get("depreciation", nf.depreciation)
        finance_costs = pnl_data.get("finance_costs", nf.finance_costs)
        other_income = pnl_data.get("other_income", nf.other_income)
        pbt = pnl_data.get("pbt", nf.pbt)
        pat = pnl_data.get("pat", nf.pat_after_mi)
        eps = nf.eps
        dps = nf.dps

        # EBIT = PBT + Finance Costs - Other Income
        ebit = [
            (pb + (fc or 0.0) - (oi or 0.0)) if pb is not None else None
            for pb, fc, oi in zip(pbt, finance_costs, other_income)
        ]

        # EBITDA = EBIT + Depreciation
        ebitda = [
            (e + (d or 0.0)) if e is not None else None
            for e, d in zip(ebit, depreciation)
        ]

        # Margins
        ebitda_margin = [
            safe_div(e, nr) * 100.0 if (e is not None and nr) else None
            for e, nr in zip(ebitda, net_revenue)
        ]
        ebit_margin = [
            safe_div(e, nr) * 100.0 if (e is not None and nr) else None
            for e, nr in zip(ebit, net_revenue)
        ]

        # Employee metrics
        employee_cost_pct, revenue_per_employee_lakh = self._compute_employee_metrics(
            net_revenue, employee_cost
        )

        # FCF and conversion
        cfo = nf.cfo
        capex = nf.capex
        fcf = [
            (c - abs(k)) if (c is not None and k is not None) else None
            for c, k in zip(cfo, capex)
        ]
        fcf_conversion = [
            safe_div(f, p) * 100.0 if (f is not None and p) else None
            for f, p in zip(fcf, pat)
        ]

        # Payout ratio
        payout_ratio: list[float | None] = []
        for d, e in zip(dps, eps):
            if d is not None and e is not None and e != 0:
                payout_ratio.append(min((d / e) * 100.0, 100.0))
            else:
                payout_ratio.append(None)

        # Headcount from ratios DataFrame
        headcount = self._extract_headcount()
        attrition_pct = self._extract_attrition()

        # Peer selection
        ticker = nf.ticker
        peers = [p for p in self.PEERS if p != ticker.upper()]

        metrics = ITMetrics(
            fy_labels=nf.fy_labels,
            net_revenue=net_revenue,
            employee_cost=employee_cost,
            other_opex=other_opex,
            ebitda=ebitda,
            depreciation=depreciation,
            ebit=ebit,
            other_income=other_income,
            finance_costs=finance_costs,
            pbt=pbt,
            pat=pat,
            eps=eps,
            ebitda_margin=ebitda_margin,
            ebit_margin=ebit_margin,
            employee_cost_pct=employee_cost_pct,
            cfo=cfo,
            capex=capex,
            fcf=fcf,
            fcf_conversion=fcf_conversion,
            dps=dps,
            payout_ratio=payout_ratio,
            roce=ratios.roce,
            roe=ratios.roe,
            roa=ratios.roa,
            headcount=headcount,
            revenue_per_employee_lakh=revenue_per_employee_lakh,
            attrition_pct=attrition_pct,
            vertical_mix={},
            geo_mix={},
            peers=peers,
        )
        self._it_metrics = metrics
        return metrics

    def prepare_valuation_inputs(self, proj_years: int = 5) -> ITValuationInputs:
        """
        Assemble historical data for valuation.
        WACC: compute_wacc(beta, risk_free=7.0, erp=5.5)
        Terminal growth: 5.5% (USD revenue companies grow in line with global IT spend)
        IT companies typically have large net cash → subtract from EV in DCF
        """
        nf = self._nf or self.normalize()
        metrics = self._it_metrics or self.compute_it_metrics()

        ticker = nf.ticker

        def _clean(lst: list[float | None]) -> list[float]:
            return [v if v is not None else 0.0 for v in lst]

        net_revenue_hist = _clean(metrics.net_revenue)
        ebit_hist = _clean(metrics.ebit)
        pat_hist = _clean(metrics.pat)
        fcf_hist = _clean(metrics.fcf)
        eps_hist = _clean(metrics.eps)

        # Net cash (positive for IT companies with large cash hoards)
        net_cash_cr = 0.0
        if nf.net_debt and any(v is not None for v in nf.net_debt):
            for v in reversed(nf.net_debt):
                if v is not None:
                    net_cash_cr = -v  # net_debt sign convention: positive = debt
                    break

        # Shares outstanding (Crores)
        shares_cr = float(self.market_data.get("shares_outstanding", 0.0))
        if shares_cr == 0.0:
            mktcap = float(self.market_data.get("market_cap_cr", 0.0))
            price = float(self.market_data.get("price", 0.0))
            if mktcap > 0 and price > 0:
                shares_cr = mktcap / price

        # WACC
        beta = float(self.market_data.get("beta", IT_SECTOR_BETAS.get(ticker.upper(), IT_SECTOR_BETAS["DEFAULT"])))
        wacc = compute_wacc(
            beta=beta,
            risk_free_rate=7.0,
            equity_risk_premium=5.5,
            cost_of_debt=float(self.market_data.get("cost_of_debt", 7.5)),
            debt_weight=0.0,  # IT companies are typically net cash
            tax_rate=0.25,
        )

        # Projection FY labels
        last_fy = nf.fy_labels[-1] if nf.fy_labels else "FY25"
        proj_fy_labels = self._generate_proj_fy_labels(last_fy, proj_years)

        return ITValuationInputs(
            ticker=ticker,
            fy_labels_hist=nf.fy_labels,
            fy_labels_proj=proj_fy_labels,
            net_revenue_hist=net_revenue_hist,
            ebit_hist=ebit_hist,
            pat_hist=pat_hist,
            fcf_hist=fcf_hist,
            eps_hist=eps_hist,
            net_revenue_proj=[None] * proj_years,
            ebit_proj=[None] * proj_years,
            pat_proj=[None] * proj_years,
            fcf_proj=[None] * proj_years,
            eps_proj=[None] * proj_years,
            wacc=wacc,
            terminal_growth=5.5,    # global IT spend growth (USD-linked companies)
            net_cash_cr=net_cash_cr,
            shares_outstanding_cr=shares_cr,
            peer_tickers=[p for p in self.PEERS if p != ticker.upper()],
        )

    def estimate_projections(
        self,
        revenue_growth_rates: list[float],
        ebit_margin_targets: list[float],
        tax_rate: float = 0.25,
        fcf_conversion: float = 0.92,
    ) -> ITValuationInputs:
        """
        IT projection engine:
        revenue[t] = revenue[t-1] * (1 + growth)
        ebit[t] = revenue[t] * ebit_margin_target
        pat[t] = ebit[t] * (1 - tax) + other_income_avg
        fcf[t] = pat[t] * fcf_conversion  (IT is asset-light, high conversion)
        eps[t] = pat[t] * 1e7 / shares_outstanding_lakh  (convert Cr → absolute)

        Parameters
        ----------
        revenue_growth_rates : list[float]
            Fractional YoY growth rates, e.g. [0.12, 0.13, 0.14, 0.13, 0.12].
        ebit_margin_targets : list[float]
            EBIT margin as fraction, e.g. [0.25, 0.26, 0.27, 0.27, 0.28].
        tax_rate : float
            Effective tax rate (fraction). Default 0.25.
        fcf_conversion : float
            FCF / PAT ratio (fraction). Default 0.92 for asset-light IT.
        """
        proj_years = len(revenue_growth_rates)
        if len(ebit_margin_targets) != proj_years:
            raise ValueError(
                f"revenue_growth_rates ({len(revenue_growth_rates)}) and "
                f"ebit_margin_targets ({len(ebit_margin_targets)}) must have the same length."
            )

        val_inputs = self.prepare_valuation_inputs(proj_years=proj_years)
        nf = self._nf or self.normalize()
        metrics = self._it_metrics or self.compute_it_metrics()

        # Last historical revenue as base
        base_revenue = val_inputs.net_revenue_hist[-1] if val_inputs.net_revenue_hist else 0.0

        # Average other income from last 3 years
        oi_vals = [v for v in (metrics.other_income or []) if v is not None]
        avg_other_income = (sum(oi_vals[-3:]) / len(oi_vals[-3:])) if oi_vals else 0.0

        shares_cr = val_inputs.shares_outstanding_cr

        net_revenue_proj: list[float | None] = []
        ebit_proj: list[float | None] = []
        pat_proj: list[float | None] = []
        fcf_proj: list[float | None] = []
        eps_proj: list[float | None] = []

        current_revenue = base_revenue
        for g, ebit_m in zip(revenue_growth_rates, ebit_margin_targets):
            # Revenue
            current_revenue = current_revenue * (1.0 + g)
            net_revenue_proj.append(round(current_revenue, 2))

            # EBIT
            ebit = current_revenue * ebit_m
            ebit_proj.append(round(ebit, 2))

            # PAT = EBIT * (1 - tax) + Other Income
            # (EBIT already excludes finance costs for net-cash IT companies)
            pat = max(ebit * (1.0 - tax_rate) + avg_other_income * (1.0 - tax_rate), 0.0)
            pat_proj.append(round(pat, 2))

            # FCF
            fcf = pat * fcf_conversion
            fcf_proj.append(round(fcf, 2))

            # EPS: PAT (INR Cr) / shares (Crores) = INR per share
            if shares_cr > 0:
                eps_proj.append(round(pat / shares_cr, 2))
            else:
                eps_proj.append(None)

        val_inputs.net_revenue_proj = net_revenue_proj
        val_inputs.ebit_proj = ebit_proj
        val_inputs.pat_proj = pat_proj
        val_inputs.fcf_proj = fcf_proj
        val_inputs.eps_proj = eps_proj

        return val_inputs

    # ------------------------------------------------------------------
    # Private helpers
    # ------------------------------------------------------------------

    def _extract_it_pnl(self, nf: NormalizedFinancials) -> dict:
        """Extract IT-specific P&L items. Returns dict of {field: list}."""
        from src.models.base_model import _to_float

        result: dict[str, list[float | None]] = {}
        df = self._get_df("pnl")
        if df is None:
            return result

        fy_labels = nf.fy_labels

        # Build lowercase row index
        row_index: dict[str, object] = {}
        for idx in df.index:
            row_index[str(idx).strip().lower()] = idx

        for label, field_name in IT_PNL_LABELS.items():
            if field_name not in result and label in row_index:
                raw_idx = row_index[label]
                row = df.loc[raw_idx]
                series: list[float | None] = []
                for fy in fy_labels:
                    val = None
                    for col_key in self._fy_to_col_keys(fy, df.columns):
                        if col_key in df.columns:
                            val = _to_float(row[col_key])
                            break
                    series.append(val)
                result[field_name] = series

        return result

    def _compute_employee_metrics(
        self,
        net_revenue: list,
        employee_cost: list,
    ) -> tuple:
        """
        Returns (employee_cost_pct, revenue_per_employee_lakh).
        employee_cost_pct: employee_cost / net_revenue * 100 (%)
        revenue_per_employee_lakh: net_revenue (Cr) * 1e7 / (headcount * 1e5) in Lakhs
            = net_revenue * 100 / headcount  — but headcount not available here,
            so this method only returns employee_cost_pct.
        Revenue per employee is computed in compute_it_metrics where headcount is available.
        """
        employee_cost_pct: list[float | None] = [
            safe_div(ec, nr) * 100.0 if (ec is not None and nr) else None
            for ec, nr in zip(employee_cost, net_revenue)
        ]
        # Revenue per employee placeholder — computed with headcount in main method
        rev_per_emp: list[float | None] = [None] * len(net_revenue)
        return employee_cost_pct, rev_per_emp

    def _extract_headcount(self) -> list[int | None]:
        """
        Try to extract headcount from ratios DataFrame.
        screener.in occasionally reports employee count as a ratio row.
        Returns list of int | None aligned to fy_labels.
        """
        from src.models.base_model import _to_float

        nf = self._nf
        if nf is None:
            return []

        n = len(nf.fy_labels)
        fy_labels = nf.fy_labels

        cd = self.company_data
        ratios_df = getattr(cd, "ratios", None)
        if ratios_df is None:
            ratios_df = getattr(cd, "key_ratios", None)
        if ratios_df is None:
            return [None] * n

        row_index: dict[str, object] = {}
        for idx in ratios_df.index:
            row_index[str(idx).strip().lower()] = idx

        headcount_labels = ["employees", "headcount", "employee count", "number of employees"]
        for label in headcount_labels:
            if label in row_index:
                raw_idx = row_index[label]
                row = ratios_df.loc[raw_idx]
                result: list[int | None] = []
                for fy in fy_labels:
                    val = None
                    for col_key in self._fy_to_col_keys(fy, ratios_df.columns):
                        if col_key in ratios_df.columns:
                            v = _to_float(row[col_key])
                            val = int(v) if v is not None else None
                            break
                    result.append(val)
                return result

        return [None] * n

    def _extract_attrition(self) -> list[float | None]:
        """
        Try to extract attrition % from ratios DataFrame.
        Returns list of float | None aligned to fy_labels.
        """
        from src.models.base_model import _to_float

        nf = self._nf
        if nf is None:
            return []

        n = len(nf.fy_labels)
        fy_labels = nf.fy_labels

        cd = self.company_data
        ratios_df = getattr(cd, "ratios", None)
        if ratios_df is None:
            ratios_df = getattr(cd, "key_ratios", None)
        if ratios_df is None:
            return [None] * n

        row_index: dict[str, object] = {}
        for idx in ratios_df.index:
            row_index[str(idx).strip().lower()] = idx

        attrition_labels = ["attrition %", "attrition rate", "employee attrition"]
        for label in attrition_labels:
            if label in row_index:
                raw_idx = row_index[label]
                row = ratios_df.loc[raw_idx]
                result: list[float | None] = []
                for fy in fy_labels:
                    val = None
                    for col_key in self._fy_to_col_keys(fy, ratios_df.columns):
                        if col_key in ratios_df.columns:
                            val = _to_float(row[col_key])
                            break
                    result.append(val)
                return result

        return [None] * n

    def _generate_proj_fy_labels(self, last_hist_fy: str, proj_years: int) -> list[str]:
        """
        Generate projection FY labels starting after last historical FY.
        e.g. last_hist_fy="FY25", proj_years=5 → ["FY26", "FY27", "FY28", "FY29", "FY30"]
        """
        labels: list[str] = []
        if last_hist_fy.startswith("FY") and len(last_hist_fy) == 4:
            try:
                last_yy = int(last_hist_fy[2:])
                for i in range(1, proj_years + 1):
                    yy = (last_yy + i) % 100
                    labels.append(f"FY{yy:02d}")
                return labels
            except ValueError:
                pass
        return [f"Proj+{i}" for i in range(1, proj_years + 1)]
