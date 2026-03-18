"""
FMCG sector financial model.
Extends BaseModel with FMCG-specific metrics:
- Gross margin reconstruction when direct line unavailable
- Volume/Price/Mix decomposition (where disclosed)
- Segment-level EBIT margins (Cigarettes, FMCG-Others, Hotels, Agribusiness, IT)
- ITC-specific segment mapping

Covers: ITC, HINDUNILVR, NESTLEIND, BRITANNIA, DABUR, MARICO, GODREJCP
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

from src.models.base_model import (
    BaseModel,
    NormalizedFinancials,
    ComputedRatios,
    safe_div,
    _to_float,
    _pairwise_op,
    _yoy_growth,
    _average_adjacent,
)

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Dataclasses
# ---------------------------------------------------------------------------

@dataclass
class FMCGMetrics:
    """FMCG-specific metrics extending base ratios."""
    fy_labels: list[str]

    # Gross margin (FMCG companies focus on this)
    gross_margin: list[float | None]        # %
    gross_profit: list[float | None]        # INR Cr

    # FMCG-specific profitability
    a_and_p_spend: list[float | None]       # Advertising & Promotion, INR Cr
    a_and_p_pct_revenue: list[float | None] # % of net revenue

    # Working capital
    dso: list[float | None]
    dio: list[float | None]
    dpo: list[float | None]
    ccc: list[float | None]

    # FCF quality
    fcf: list[float | None]
    fcf_conversion: list[float | None]      # FCF / PAT %

    # Dividend
    dps: list[float | None]
    payout_ratio: list[float | None]        # DPS / EPS %

    # Returns
    roce: list[float | None]
    roe: list[float | None]

    # Segments (may be None for companies without disclosed segments)
    segments: dict[str, dict]  # {segment_name: {"revenue": [...], "ebit": [...], "ebit_margin": [...]}}


@dataclass
class FMCGValuationInputs:
    """Inputs for DCF and comparables valuation."""
    ticker: str
    fy_labels_hist: list[str]
    fy_labels_proj: list[str]

    # Historical
    net_revenue_hist: list[float]
    ebitda_hist: list[float]
    pat_hist: list[float]
    fcf_hist: list[float]

    # Projections (to be filled by LLM or assumption engine)
    net_revenue_proj: list[float | None]
    ebitda_proj: list[float | None]
    pat_proj: list[float | None]
    fcf_proj: list[float | None]
    eps_proj: list[float | None]

    # DCF inputs
    wacc: float | None
    terminal_growth: float | None
    net_debt: float   # latest, INR Cr (negative = net cash)
    shares_outstanding: float  # Crores

    # Peer benchmarks
    peer_tickers: list[str]


# ---------------------------------------------------------------------------
# WACC helper (module-level for reuse)
# ---------------------------------------------------------------------------

def compute_wacc(
    beta: float,
    risk_free_rate: float = 7.0,
    equity_risk_premium: float = 5.5,
    cost_of_debt: float = 7.5,
    debt_weight: float = 0.0,
    tax_rate: float = 0.25,
) -> float:
    """
    CAPM-based WACC for Indian companies.
    Defaults: India 10yr G-sec ~7%, ERP ~5.5% (Damodaran India estimate).
    For most FMCG companies, debt_weight ≈ 0 (net cash positions).
    Returns WACC as percentage (e.g., 11.5 not 0.115).

    Formula:
        Ke = Rf + beta * ERP
        Kd_after_tax = Kd * (1 - tax_rate)
        WACC = Ke * equity_weight + Kd_after_tax * debt_weight
    """
    if not (0 <= debt_weight <= 1):
        raise ValueError(f"debt_weight must be between 0 and 1, got {debt_weight}")
    if beta < 0:
        raise ValueError(f"beta must be non-negative, got {beta}")

    equity_weight = 1.0 - debt_weight
    cost_of_equity = risk_free_rate + beta * equity_risk_premium
    cost_of_debt_after_tax = cost_of_debt * (1.0 - tax_rate)
    wacc = cost_of_equity * equity_weight + cost_of_debt_after_tax * debt_weight
    return round(wacc, 4)


# ---------------------------------------------------------------------------
# FMCGModel
# ---------------------------------------------------------------------------

class FMCGModel(BaseModel):
    """
    FMCG sector financial model.
    Extends BaseModel with gross margin reconstruction, A&P analysis,
    segment data parsing, and DCF valuation scaffolding.
    """

    SECTOR = "FMCG"

    # Default peers by sub-sector
    PEERS_DIVERSIFIED = ["ITC", "HINDUNILVR", "GODREJCP", "DABUR", "MARICO"]
    PEERS_FOOD = ["NESTLEIND", "BRITANNIA", "TATACONSUM", "HINDUNILVR"]

    # ITC segment names as they appear in screener.in segment data
    ITC_SEGMENTS = [
        "Cigarettes",
        "FMCG - Others",
        "Hotels",
        "Agri Business",
        "Paperboards, Paper & Packaging",
    ]

    # Advertising line label variants across companies
    _A_AND_P_LABELS = [
        "advertising expenses",
        "advertisement expenses",
        "advertising & promotion",
        "advertising and promotion",
        "marketing expenses",
        "brand building expenses",
        "selling & distribution expenses",
        "advertisement and publicity",
    ]

    # Sector-default betas (unleveraged, for net-cash FMCG companies)
    _SECTOR_BETAS: dict[str, float] = {
        "ITC": 0.75,
        "HINDUNILVR": 0.65,
        "NESTLEIND": 0.60,
        "BRITANNIA": 0.70,
        "DABUR": 0.70,
        "MARICO": 0.72,
        "GODREJCP": 0.75,
        "DEFAULT": 0.72,
    }

    def __init__(
        self,
        company_data: "CompanyData",
        market_data: dict | None = None,
    ) -> None:
        """
        Parameters
        ----------
        company_data : CompanyData
            Output of ScreenerClient; must have profit_loss, balance_sheet,
            cash_flow DataFrames and ticker attribute.
        market_data : dict, optional
            May contain 'beta', 'market_cap', 'shares_outstanding', etc.
            sourced from NSE/BSE API or a market data provider.
        """
        super().__init__(company_data)
        self.market_data: dict = market_data or {}
        self._fmcg_metrics: FMCGMetrics | None = None

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def compute_fmcg_metrics(self) -> FMCGMetrics:
        """
        Full pipeline:
        1. normalize() → NormalizedFinancials
        2. compute_ratios() → ComputedRatios
        3. Overlay FMCG-specific metrics
        Returns FMCGMetrics dataclass.
        """
        nf = self.normalize()
        ratios = self.compute_ratios(nf)

        gross_profit, gross_margin = self._compute_gross_margin(nf)
        a_and_p_spend, a_and_p_pct = self._compute_a_and_p(nf)
        segments = self._parse_segments()

        # Dividend payout ratio = DPS / EPS %
        payout_ratio: list[float | None] = []
        for dps, eps in zip(nf.dps, nf.eps):
            if dps is not None and eps is not None and eps != 0:
                payout_ratio.append(min((dps / eps) * 100.0, 100.0))
            else:
                payout_ratio.append(None)

        metrics = FMCGMetrics(
            fy_labels         = nf.fy_labels,
            gross_margin      = gross_margin,
            gross_profit      = gross_profit,
            a_and_p_spend     = a_and_p_spend,
            a_and_p_pct_revenue = a_and_p_pct,
            dso               = ratios.dso,
            dio               = ratios.dio,
            dpo               = ratios.dpo,
            ccc               = ratios.ccc,
            fcf               = nf.fcf,
            fcf_conversion    = ratios.fcf_conversion,
            dps               = nf.dps,
            payout_ratio      = payout_ratio,
            roce              = ratios.roce,
            roe               = ratios.roe,
            segments          = segments,
        )
        self._fmcg_metrics = metrics
        return metrics

    def prepare_valuation_inputs(self, proj_years: int = 5) -> FMCGValuationInputs:
        """
        Assemble historical data for valuation modelling.
        Projection fields are set to None; they must be filled by the LLM
        or a separate assumption engine via estimate_projections().

        WACC is computed using CAPM with:
        - Risk-free rate: 7.0% (India 10yr G-sec)
        - ERP: 5.5% (Damodaran India)
        - Beta: from market_data['beta'] if available, else sector default
        - Debt weight: derived from net_debt / (net_debt + market_cap) if available
        """
        nf = self._nf or self.normalize()

        ticker = nf.ticker

        # Build clean historical lists (replace None with 0 for arrays that
        # need to be numeric; keep as list for downstream flexibility)
        def _clean(lst: list[float | None]) -> list[float]:
            return [v if v is not None else 0.0 for v in lst]

        net_revenue_hist = _clean(nf.net_revenue)
        ebitda_hist      = _clean(nf.ebitda)
        pat_hist         = _clean(nf.pat_after_mi)
        fcf_hist         = _clean(nf.fcf)

        # Latest net debt
        net_debt_latest: float = 0.0
        if nf.net_debt and any(v is not None for v in nf.net_debt):
            for v in reversed(nf.net_debt):
                if v is not None:
                    net_debt_latest = v
                    break

        # Shares outstanding (Crores)
        shares_outstanding: float = float(self.market_data.get("shares_outstanding", 0.0))
        if shares_outstanding == 0.0:
            # Try deriving from market_cap and price
            mktcap = float(self.market_data.get("market_cap_cr", 0.0))
            price  = float(self.market_data.get("price", 0.0))
            if mktcap > 0 and price > 0:
                shares_outstanding = mktcap / price  # result in Crores

        # Beta
        beta = float(self.market_data.get("beta", self._get_sector_beta(ticker)))

        # Debt weight for WACC
        mktcap_cr = float(self.market_data.get("market_cap_cr", 0.0))
        debt_weight = 0.0
        if mktcap_cr > 0 and net_debt_latest > 0:
            total_capital = mktcap_cr + net_debt_latest
            debt_weight = min(max(net_debt_latest / total_capital, 0.0), 0.5)

        wacc = compute_wacc(
            beta=beta,
            risk_free_rate=7.0,
            equity_risk_premium=5.5,
            cost_of_debt=float(self.market_data.get("cost_of_debt", 7.5)),
            debt_weight=debt_weight,
            tax_rate=0.25,
        )

        # Projection FY labels
        last_fy = nf.fy_labels[-1] if nf.fy_labels else "FY25"
        proj_fy_labels = self._generate_proj_fy_labels(last_fy, proj_years)

        # Peer selection
        peers = self._select_peers(ticker)

        val_inputs = FMCGValuationInputs(
            ticker            = ticker,
            fy_labels_hist    = nf.fy_labels,
            fy_labels_proj    = proj_fy_labels,
            net_revenue_hist  = net_revenue_hist,
            ebitda_hist       = ebitda_hist,
            pat_hist          = pat_hist,
            fcf_hist          = fcf_hist,
            net_revenue_proj  = [None] * proj_years,
            ebitda_proj       = [None] * proj_years,
            pat_proj          = [None] * proj_years,
            fcf_proj          = [None] * proj_years,
            eps_proj          = [None] * proj_years,
            wacc              = wacc,
            terminal_growth   = 5.0,   # India long-term nominal GDP growth assumption
            net_debt          = net_debt_latest,
            shares_outstanding = shares_outstanding,
            peer_tickers      = peers,
        )
        return val_inputs

    def estimate_projections(
        self,
        revenue_growth_rates: list[float],
        ebitda_margin_targets: list[float],
        tax_rate: float = 0.25,
    ) -> FMCGValuationInputs:
        """
        Simple projection engine: apply growth rates and margin assumptions.

        Parameters
        ----------
        revenue_growth_rates : list[float]
            Fractional growth rates per projection year, e.g. [0.10, 0.11, 0.12, 0.12, 0.11].
            Length must match proj_years used in prepare_valuation_inputs().
        ebitda_margin_targets : list[float]
            EBITDA margin as fraction per year, e.g. [0.55, 0.57, 0.58, 0.58, 0.59].
        tax_rate : float
            Effective tax rate (fraction). Default 0.25.

        Returns
        -------
        FMCGValuationInputs with proj fields populated.

        Notes
        -----
        PAT approximation: PAT ≈ (EBITDA - D&A) * (1 - tax_rate)
            D&A is held constant at last-historical average.
        FCF = PAT * fcf_conversion_factor (default 0.85 — i.e. 85% PAT-to-FCF).
        EPS = PAT / shares_outstanding.
        """
        proj_years = len(revenue_growth_rates)
        if len(ebitda_margin_targets) != proj_years:
            raise ValueError(
                f"revenue_growth_rates ({len(revenue_growth_rates)}) and "
                f"ebitda_margin_targets ({len(ebitda_margin_targets)}) must have the same length."
            )

        val_inputs = self.prepare_valuation_inputs(proj_years=proj_years)
        nf = self._nf or self.normalize()

        # Last historical revenue as base
        base_revenue = val_inputs.net_revenue_hist[-1] if val_inputs.net_revenue_hist else 0.0

        # Estimate last-3yr average D&A for PAT calculation
        dep_values = [v for v in (nf.depreciation or []) if v is not None]
        avg_da = sum(dep_values[-3:]) / len(dep_values[-3:]) if dep_values else 0.0

        # FCF conversion factor — use historical if available, else 0.85
        fcf_conversions = [
            v for v in (self._ratios.fcf_conversion if self._ratios else []) if v is not None
        ]
        fcf_factor = (sum(fcf_conversions[-3:]) / len(fcf_conversions[-3:]) / 100.0
                      if fcf_conversions else 0.85)
        fcf_factor = max(0.5, min(1.0, fcf_factor))  # clamp to [50%, 100%]

        shares = val_inputs.shares_outstanding

        net_revenue_proj: list[float | None] = []
        ebitda_proj: list[float | None]      = []
        pat_proj: list[float | None]         = []
        fcf_proj: list[float | None]         = []
        eps_proj: list[float | None]         = []

        current_revenue = base_revenue
        for g, m in zip(revenue_growth_rates, ebitda_margin_targets):
            # Revenue
            current_revenue = current_revenue * (1.0 + g)
            net_revenue_proj.append(round(current_revenue, 2))

            # EBITDA
            ebitda = current_revenue * m
            ebitda_proj.append(round(ebitda, 2))

            # EBIT = EBITDA - D&A
            ebit = ebitda - avg_da

            # PAT (simplified; ignores interest since FMCG companies are typically net cash)
            pat = max(ebit * (1.0 - tax_rate), 0.0)
            pat_proj.append(round(pat, 2))

            # FCF
            fcf = pat * fcf_factor
            fcf_proj.append(round(fcf, 2))

            # EPS
            if shares > 0:
                eps_proj.append(round(pat / shares, 2))
            else:
                eps_proj.append(None)

        val_inputs.net_revenue_proj = net_revenue_proj
        val_inputs.ebitda_proj      = ebitda_proj
        val_inputs.pat_proj         = pat_proj
        val_inputs.fcf_proj         = fcf_proj
        val_inputs.eps_proj         = eps_proj

        return val_inputs

    # ------------------------------------------------------------------
    # Private helpers
    # ------------------------------------------------------------------

    def _compute_gross_margin(
        self, nf: NormalizedFinancials
    ) -> tuple[list[float | None], list[float | None]]:
        """
        Returns (gross_profit_list, gross_margin_pct_list).

        For ITC: net_revenue = gross_revenue - excise_duty (pre-GST years).
        Gross profit = net_revenue - raw_materials.
        For companies with direct gross_profit line it is used as-is.
        """
        n = len(nf.fy_labels)
        gross_profit: list[float | None] = []
        gross_margin: list[float | None] = []

        for i in range(n):
            nr = nf.net_revenue[i]
            gp_direct = nf.gross_profit[i] if nf.gross_profit else None
            rm = nf.raw_materials[i] if nf.raw_materials else None

            # For ITC: apply excise-duty adjustment in pre-GST years
            if nf.ticker == "ITC" and nf.excise_duty:
                ex = nf.excise_duty[i]
                gr = nf.gross_revenue[i] if nf.gross_revenue else None
                if gr is not None and ex is not None:
                    nr = gr - ex  # effective net revenue

            if gp_direct is not None:
                gp = gp_direct
            elif nr is not None and rm is not None:
                gp = nr - rm
            else:
                gp = None

            gross_profit.append(gp)
            gm = safe_div(gp, nr)
            gross_margin.append(gm * 100.0 if gm is not None else None)

        return gross_profit, gross_margin

    def _parse_segments(self) -> dict[str, dict]:
        """
        Parse segment data from company_data if available.
        Looks for a 'segments' or 'segment_data' attribute on company_data
        (expected to be a dict of DataFrames keyed by segment name).

        Returns empty dict if not available rather than raising an exception.
        """
        cd = self.company_data
        segments: dict[str, dict] = {}

        # Try multiple attribute names for segment data
        seg_data = None
        for attr in ("segments", "segment_data", "segment_results"):
            seg_data = getattr(cd, attr, None)
            if seg_data is not None:
                break

        if seg_data is None:
            logger.debug("No segment data found for %s", getattr(cd, "ticker", "UNKNOWN"))
            return segments

        nf = self._nf or self.normalize()
        fy_labels = nf.fy_labels

        # Normalize segment names for ITC
        ticker = getattr(cd, "ticker", "")
        expected_segments = self.ITC_SEGMENTS if ticker == "ITC" else []

        # seg_data may be a dict: {segment_name: DataFrame} or a single DataFrame
        if hasattr(seg_data, "items"):
            for seg_name, seg_df in seg_data.items():
                segments[seg_name] = self._extract_segment_metrics(seg_df, fy_labels, seg_name)
        elif hasattr(seg_data, "index"):
            # Single DataFrame with segments as rows
            for seg_name in seg_data.index:
                row_df = seg_data.loc[[seg_name]]
                segments[str(seg_name)] = self._extract_segment_metrics(row_df, fy_labels, str(seg_name))
        else:
            logger.warning("Unrecognized segment data format for %s", ticker)

        return segments

    def _extract_segment_metrics(
        self, seg_df, fy_labels: list[str], seg_name: str
    ) -> dict:
        """
        Extract revenue, EBIT, and EBIT margin for a single segment.
        Returns dict with keys 'revenue', 'ebit', 'ebit_margin', 'fy_labels'.
        """
        result: dict = {"fy_labels": fy_labels, "revenue": [], "ebit": [], "ebit_margin": []}

        if seg_df is None or not hasattr(seg_df, "columns"):
            result["revenue"]     = [None] * len(fy_labels)
            result["ebit"]        = [None] * len(fy_labels)
            result["ebit_margin"] = [None] * len(fy_labels)
            return result

        # Row labels within a segment block
        row_index: dict[str, any] = {}
        for idx in seg_df.index:
            row_index[str(idx).strip().lower()] = idx

        rev_labels  = ["revenue", "net revenue", "segment revenue", "sales"]
        ebit_labels = ["ebit", "segment result", "profit", "operating profit"]

        def _get_row_series(labels: list[str]) -> list[float | None]:
            for lbl in labels:
                if lbl in row_index:
                    row = seg_df.loc[row_index[lbl]]
                    return [_to_float(row.get(fy, row.get(f"Mar {fy[2:]}", None)))
                            for fy in fy_labels]
            return [None] * len(fy_labels)

        revenue = _get_row_series(rev_labels)
        ebit    = _get_row_series(ebit_labels)
        ebit_margin = [
            safe_div(e, r) * 100.0 if (e is not None and r) else None
            for e, r in zip(ebit, revenue)
        ]

        result["revenue"]     = revenue
        result["ebit"]        = ebit
        result["ebit_margin"] = ebit_margin
        return result

    def _compute_a_and_p(
        self, nf: NormalizedFinancials
    ) -> tuple[list[float | None], list[float | None]]:
        """
        Extract Advertising & Promotion spend from P&L.
        Many FMCG companies disclose it as a separate line item.
        If not separately available, falls back to None.

        Returns
        -------
        (a_and_p_cr_list, a_and_p_pct_list)
        """
        fy_labels = nf.fy_labels
        a_and_p_cr: list[float | None] = [None] * len(fy_labels)

        # Try to extract from P&L DataFrame using known label variants
        df = self._get_df("pnl")
        if df is not None:
            row_index: dict[str, any] = {}
            for idx in df.index:
                row_index[str(idx).strip().lower()] = idx

            for label in self._A_AND_P_LABELS:
                if label in row_index:
                    raw_idx = row_index[label]
                    row = df.loc[raw_idx]
                    series: list[float | None] = []
                    for fy in fy_labels:
                        val = None
                        for col_key in self._fy_to_col_keys(fy, df.columns):
                            if col_key in df.columns:
                                val = _to_float(row[col_key])
                                break
                        series.append(abs(val) if val is not None else None)
                    a_and_p_cr = series
                    logger.debug("Found A&P line '%s' for %s", label, nf.ticker)
                    break
            else:
                logger.debug("No explicit A&P line found for %s; returning None series", nf.ticker)

        # Compute as % of net revenue
        a_and_p_pct: list[float | None] = [
            safe_div(ap, nr) * 100.0 if (ap is not None and nr) else None
            for ap, nr in zip(a_and_p_cr, nf.net_revenue)
        ]

        return a_and_p_cr, a_and_p_pct

    def _get_sector_beta(self, ticker: str) -> float:
        """Return the sector-default beta for the given ticker."""
        return self._SECTOR_BETAS.get(ticker.upper(), self._SECTOR_BETAS["DEFAULT"])

    def _select_peers(self, ticker: str) -> list[str]:
        """Select appropriate peer group for comparables valuation."""
        food_companies = {"NESTLEIND", "BRITANNIA", "TATACONSUM"}
        if ticker.upper() in food_companies:
            peers = [p for p in self.PEERS_FOOD if p != ticker.upper()]
        else:
            peers = [p for p in self.PEERS_DIVERSIFIED if p != ticker.upper()]
        return peers

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
        # Fallback: generic labels
        return [f"Proj+{i}" for i in range(1, proj_years + 1)]
