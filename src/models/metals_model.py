"""
Metals & Commodities sector financial model for Indian listed companies.
Extends BaseModel with metals-specific metrics:
- Volume-based metrics: steel/aluminium production in million tonnes
- Per-tonne economics: realization per tonne, EBITDA per tonne
- Segment breakdown: Steel / Aluminium / Copper / Zinc
- Raw material cost mapping: coking coal + iron ore (steel), bauxite + power (aluminium)
- Leverage analysis: net_debt_to_ebitda — key risk flag for highly-levered producers
- EV/EBITDA and EV/tonne valuation (primary for metals sector)
- Conservative DCF crosscheck via WACC

Covers: TATASTEEL, JSWSTEEL, HINDALCO, VEDL, NATIONALUM, SAIL, NMDC,
        COALINDIA, MOIL, APLAPOLLO

All monetary values in INR Crores.
All lists indexed 0 = oldest FY.
None for missing data (not 0).
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
from src.models.fmcg_model import compute_wacc

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Sector-level constants
# ---------------------------------------------------------------------------

# Levered betas for Indian metals companies (reflects commodity cyclicality
# and balance-sheet leverage). Source: Damodaran + NSE regression estimates.
METALS_SECTOR_BETAS: dict[str, float] = {
    "TATASTEEL":  1.40,
    "JSWSTEEL":   1.35,
    "HINDALCO":   1.20,
    "VEDL":       1.45,
    "NATIONALUM": 1.10,
    "SAIL":       1.30,
    "NMDC":       1.00,
    "COALINDIA":  0.85,
    "MOIL":       0.90,
    "APLAPOLLO":  1.15,
}

DEFAULT_METALS_BETA: float = 1.30

# Screener.in label variants for raw-material cost lines
# Steel: coking coal + iron ore are the primary inputs
# Aluminium: bauxite + power costs dominate
_RAW_MATERIAL_LABELS: list[str] = [
    "raw materials",
    "raw material consumed",
    "material cost",
    "material costs",
    "cost of materials consumed",
    "raw material cost",
    "consumption of raw materials",
    "coking coal",
    "iron ore",
    "bauxite",
]

# Segment name variants as they appear in screener.in or annual report data
_STEEL_SEGMENT_LABELS:     list[str] = ["steel", "flat steel", "long steel", "steel products"]
_ALUMINIUM_SEGMENT_LABELS: list[str] = ["aluminium", "aluminum", "aluminium products"]
_COPPER_SEGMENT_LABELS:    list[str] = ["copper", "copper products"]
_ZINC_SEGMENT_LABELS:      list[str] = ["zinc", "zinc products", "zinc-lead"]
_MINING_SEGMENT_LABELS:    list[str] = ["mining", "iron ore mining", "coal mining"]
_POWER_SEGMENT_LABELS:     list[str] = ["power", "energy"]


# ---------------------------------------------------------------------------
# Dataclasses
# ---------------------------------------------------------------------------

@dataclass
class MetalsMetrics:
    """
    Metals & commodities sector-specific metrics.
    All per-FY lists are aligned to fy_labels (index 0 = oldest FY).
    None is used wherever data is not disclosed or not applicable.
    """
    fy_labels: list[str]

    # ---- Standard P&L ----
    net_revenue:    list[float | None]   # INR Cr
    ebitda:         list[float | None]   # INR Cr
    pat:            list[float | None]   # INR Cr (PAT after minority interest)
    eps:            list[float | None]   # INR per share

    # ---- Profitability margins ----
    ebitda_margin:  list[float | None]   # % of net revenue
    pat_margin:     list[float | None]   # % of net revenue

    # ---- Volume metrics ----
    # Set None for tickers that do not produce the relevant metal.
    steel_volume_mt:     list[float | None]  # crude/finished steel, million tonnes
    aluminium_volume_mt: list[float | None]  # primary aluminium, million tonnes

    # Per-tonne economics (INR per tonne; computed from total volume)
    realization_per_tonne: list[float | None]  # net_revenue / total_volume
    ebitda_per_tonne:      list[float | None]  # ebitda / total_volume

    # ---- Cost structure ----
    raw_material_cost_pct: list[float | None]  # raw_materials / net_revenue, %

    # ---- Returns ----
    roce: list[float | None]   # %
    roe:  list[float | None]   # %

    # ---- Leverage — key risk metric for metals ----
    # If net_debt_to_ebitda > 4x, the company is flagged as high-risk.
    net_debt_to_ebitda: list[float | None]

    # ---- Cash flow quality ----
    fcf:            list[float | None]   # INR Cr
    fcf_conversion: list[float | None]   # FCF / PAT, %

    # ---- Capital intensity ----
    capex_to_revenue: list[float | None]  # capex / net_revenue, %

    # ---- Segment breakdown ----
    # {segment_name: {"revenue": [...], "ebit": [...], "ebit_margin": [...], "fy_labels": [...]}}
    # Typical keys: "Steel", "Aluminium", "Copper", "Zinc", "Mining", "Power"
    segments: dict[str, dict]


@dataclass
class MetalsValuationInputs:
    """
    Inputs for DCF and EV/EBITDA valuation of Indian metals companies.
    Projection fields are initially None; populated by estimate_projections().

    Primary valuation method: EV/EBITDA (metals sector convention).
    Secondary: EV/tonne (for commodity producers with disclosed capacity).
    DCF is used as a crosscheck but rarely as the primary method.
    """
    ticker: str
    fy_labels_hist: list[str]
    fy_labels_proj: list[str]

    # ---- Historical actuals ----
    net_revenue_hist: list[float]
    ebitda_hist:      list[float]
    pat_hist:         list[float]
    fcf_hist:         list[float]

    # ---- Projections (populated by estimate_projections) ----
    net_revenue_proj: list[float | None]
    ebitda_proj:      list[float | None]
    pat_proj:         list[float | None]
    fcf_proj:         list[float | None]
    eps_proj:         list[float | None]

    # ---- DCF inputs ----
    wacc:             float | None
    terminal_growth:  float         # default 4.0 — commodity mean-reversion (lower than FMCG)
    net_debt_cr:      float         # latest net debt, INR Cr; typically large for metals cos
    shares_outstanding: float       # Crores

    # ---- EV-based valuation inputs ----
    ev_ebitda_target:  float | None  # peer/historical EV/EBITDA multiple (primary)
    ev_per_tonne:      float | None  # EV / annual production capacity (INR per tonne)

    # ---- Risk flags ----
    high_leverage_flag: bool         # True if net_debt_to_ebitda > 4x

    # ---- Peer comparables ----
    peer_tickers: list[str]


# ---------------------------------------------------------------------------
# MetalsModel
# ---------------------------------------------------------------------------

class MetalsModel(BaseModel):
    """
    Metals & Commodities sector financial model for Indian listed companies.

    Extends BaseModel with:
    - Volume extraction for steel and aluminium producers
    - Per-tonne economics (realization, EBITDA/t)
    - Segment parsing (Steel / Aluminium / Copper / Zinc)
    - Raw material cost mapping (coking coal, iron ore, bauxite, power)
    - EV/EBITDA and EV/tonne valuation scaffolding
    - High-leverage risk flagging (net_debt_to_ebitda > 4x threshold)

    Usage
    -----
    model = MetalsModel(company_data, market_data={"market_cap_cr": 80000})
    metrics = model.compute_metals_metrics()
    val_inputs = model.prepare_valuation_inputs()
    val_inputs = model.estimate_projections(
        volume_growth=0.05,
        realization_growth=0.03,
        margin_pct=0.18
    )
    """

    SECTOR = "Metals"

    # Leverage threshold: net_debt_to_ebitda > 4x → high-risk flag
    HIGH_LEVERAGE_THRESHOLD: float = 4.0

    # Default terminal growth: commodity cycle mean-reversion (lower than FMCG's 5%)
    DEFAULT_TERMINAL_GROWTH: float = 4.0

    # Default projection years
    DEFAULT_PROJ_YEARS: int = 5

    # Peer groups by sub-sector
    PEERS_STEEL:     list[str] = ["TATASTEEL", "JSWSTEEL", "SAIL", "APLAPOLLO"]
    PEERS_ALUMINIUM: list[str] = ["HINDALCO", "NATIONALUM", "VEDL"]
    PEERS_DIVERSIFIED: list[str] = ["VEDL", "HINDALCO", "TATASTEEL", "JSWSTEEL"]
    PEERS_MINING:    list[str] = ["NMDC", "COALINDIA", "MOIL"]

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
            cash_flow DataFrames and a ticker attribute.
        market_data : dict, optional
            May contain: 'beta', 'market_cap_cr', 'shares_outstanding',
            'cost_of_debt', 'price', 'production_capacity_mt' (for EV/tonne).
        """
        super().__init__(company_data)
        self.market_data: dict = market_data or {}
        self._metals_metrics: MetalsMetrics | None = None

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def compute_metals_metrics(self) -> MetalsMetrics:
        """
        Full computation pipeline:
        1. normalize()           → NormalizedFinancials  (stored as self._nf)
        2. compute_ratios()      → ComputedRatios         (stored as self._ratios)
        3. Overlay metals-specific metrics               (stored as self._metrics)

        Returns
        -------
        MetalsMetrics dataclass with all per-FY metrics populated where data
        is available, None where not disclosed or not applicable.
        """
        nf = self.normalize()
        ratios = self.compute_ratios(nf)

        n = len(nf.fy_labels)

        # ---- Volume metrics ----
        steel_vol, aluminium_vol = self._estimate_volume(nf)
        total_volume = self._compute_total_volume(steel_vol, aluminium_vol, n)
        realization_per_tonne, ebitda_per_tonne = self._compute_per_tonne_metrics(
            nf, total_volume
        )

        # ---- Raw material cost % ----
        raw_material_cost_pct: list[float | None] = [
            safe_div(rm, nr) * 100.0 if (rm is not None and nr) else None
            for rm, nr in zip(nf.raw_materials, nf.net_revenue)
        ]

        # ---- Capex intensity ----
        capex_to_revenue: list[float | None] = [
            safe_div(abs(cx), nr) * 100.0 if (cx is not None and nr) else None
            for cx, nr in zip(nf.capex, nf.net_revenue)
        ]

        # ---- Segments ----
        segments = self._extract_segments(nf)

        # ---- Assemble metrics ----
        metrics = MetalsMetrics(
            fy_labels              = nf.fy_labels,
            net_revenue            = nf.net_revenue,
            ebitda                 = nf.ebitda,
            pat                    = nf.pat_after_mi,
            eps                    = nf.eps,
            ebitda_margin          = ratios.ebitda_margin,
            pat_margin             = ratios.pat_margin,
            steel_volume_mt        = steel_vol,
            aluminium_volume_mt    = aluminium_vol,
            realization_per_tonne  = realization_per_tonne,
            ebitda_per_tonne       = ebitda_per_tonne,
            raw_material_cost_pct  = raw_material_cost_pct,
            roce                   = ratios.roce,
            roe                    = ratios.roe,
            net_debt_to_ebitda     = ratios.net_debt_to_ebitda,
            fcf                    = nf.fcf,
            fcf_conversion         = ratios.fcf_conversion,
            capex_to_revenue       = capex_to_revenue,
            segments               = segments,
        )

        self._metals_metrics = metrics
        return metrics

    def prepare_valuation_inputs(self, proj_years: int = DEFAULT_PROJ_YEARS) -> MetalsValuationInputs:
        """
        Assemble historical data for EV/EBITDA and DCF valuation.

        Projection fields are set to None; they must be populated by
        estimate_projections() or an external assumption engine.

        WACC is computed using CAPM with:
        - Risk-free rate: 7.0% (India 10yr G-sec)
        - ERP: 5.5% (Damodaran India estimate)
        - Beta: from market_data['beta'] if provided, else METALS_SECTOR_BETAS
        - Debt weight: derived from net_debt / (net_debt + market_cap)
          Metals companies carry significant debt, so debt_weight is material.

        Returns
        -------
        MetalsValuationInputs with all historical fields populated and
        projection fields set to None.
        """
        nf = self._nf or self.normalize()
        ratios = self._ratios or self.compute_ratios(nf)

        ticker = nf.ticker

        def _clean(lst: list[float | None]) -> list[float]:
            """Replace None with 0.0 for historical numeric arrays."""
            return [v if v is not None else 0.0 for v in lst]

        net_revenue_hist = _clean(nf.net_revenue)
        ebitda_hist      = _clean(nf.ebitda)
        pat_hist         = _clean(nf.pat_after_mi)
        fcf_hist         = _clean(nf.fcf)

        # Latest net debt (expected to be large for metals cos)
        net_debt_latest: float = 0.0
        if nf.net_debt and any(v is not None for v in nf.net_debt):
            for v in reversed(nf.net_debt):
                if v is not None:
                    net_debt_latest = v
                    break

        # High-leverage flag: check latest net_debt_to_ebitda
        high_leverage = self._check_high_leverage(ratios.net_debt_to_ebitda)
        if high_leverage:
            logger.warning(
                "%s: net_debt_to_ebitda > %.1fx — flagged as HIGH LEVERAGE. "
                "DCF projections carry elevated risk.",
                ticker, self.HIGH_LEVERAGE_THRESHOLD,
            )

        # Shares outstanding (Crores)
        shares_outstanding: float = float(self.market_data.get("shares_outstanding", 0.0))
        if shares_outstanding == 0.0:
            mktcap = float(self.market_data.get("market_cap_cr", 0.0))
            price  = float(self.market_data.get("price", 0.0))
            if mktcap > 0 and price > 0:
                shares_outstanding = mktcap / price  # Crores

        # Beta: use market-provided beta, else sector default
        beta = float(self.market_data.get("beta", self._get_sector_beta(ticker)))

        # Debt weight for WACC — metals companies are leverage-heavy
        mktcap_cr   = float(self.market_data.get("market_cap_cr", 0.0))
        debt_weight = 0.0
        if mktcap_cr > 0 and net_debt_latest > 0:
            total_capital = mktcap_cr + net_debt_latest
            # Cap debt_weight at 0.6 to avoid unrealistic WACC
            debt_weight = min(max(net_debt_latest / total_capital, 0.0), 0.60)

        # Cost of debt: metals cos typically borrow at 7.5–9%
        cost_of_debt = float(self.market_data.get("cost_of_debt", 8.0))

        wacc = compute_wacc(
            beta=beta,
            risk_free_rate=7.0,
            equity_risk_premium=5.5,
            cost_of_debt=cost_of_debt,
            debt_weight=debt_weight,
            tax_rate=0.25,
        )

        # EV/EBITDA target: sector-typical ranges
        #   Steel: 5–8x, Aluminium: 5–7x, Diversified: 6–9x, Mining: 7–10x
        ev_ebitda_target = float(self.market_data.get("ev_ebitda_target", 0.0)) or None

        # EV/tonne: requires disclosed production capacity in tonnes
        ev_per_tonne = float(self.market_data.get("ev_per_tonne", 0.0)) or None

        # Projection FY labels
        last_fy = nf.fy_labels[-1] if nf.fy_labels else "FY25"
        proj_fy_labels = self._generate_proj_fy_labels(last_fy, proj_years)

        # Peer group
        peers = self._select_peers(ticker)

        return MetalsValuationInputs(
            ticker             = ticker,
            fy_labels_hist     = nf.fy_labels,
            fy_labels_proj     = proj_fy_labels,
            net_revenue_hist   = net_revenue_hist,
            ebitda_hist        = ebitda_hist,
            pat_hist           = pat_hist,
            fcf_hist           = fcf_hist,
            net_revenue_proj   = [None] * proj_years,
            ebitda_proj        = [None] * proj_years,
            pat_proj           = [None] * proj_years,
            fcf_proj           = [None] * proj_years,
            eps_proj           = [None] * proj_years,
            wacc               = wacc,
            terminal_growth    = self.DEFAULT_TERMINAL_GROWTH,
            net_debt_cr        = net_debt_latest,
            shares_outstanding = shares_outstanding,
            ev_ebitda_target   = ev_ebitda_target,
            ev_per_tonne       = ev_per_tonne,
            high_leverage_flag = high_leverage,
            peer_tickers       = peers,
        )

    def estimate_projections(
        self,
        volume_growth: float = 0.05,
        realization_growth: float = 0.03,
        margin_pct: float = 0.18,
        proj_years: int = DEFAULT_PROJ_YEARS,
        tax_rate: float = 0.25,
    ) -> MetalsValuationInputs:
        """
        Conservative metals projection engine.

        Metals economics are driven by two levers:
        1. Volume growth — capacity expansion / utilization improvement
        2. LME realization — commodity price pass-through

        Revenue = previous_revenue * (1 + volume_growth) * (1 + realization_growth)
        EBITDA  = revenue * margin_pct
        PAT     ≈ (EBITDA - D&A - Finance Costs) * (1 - tax_rate)
        FCF     = PAT * fcf_conversion_factor
        EPS     = PAT / shares_outstanding

        Parameters
        ----------
        volume_growth : float
            Annual production volume growth rate (fraction). Default 5%.
            Reflects capacity ramp-up / greenfield additions.
        realization_growth : float
            Annual net realization per tonne growth rate (fraction). Default 3%.
            Reflects LME price trend + product mix improvement.
        margin_pct : float
            EBITDA margin assumption (fraction). Default 18%.
            Typical range: Steel 12–22%, Aluminium 15–25%, Mining 30–50%.
        proj_years : int
            Number of projection years. Default 5.
        tax_rate : float
            Effective tax rate (fraction). Default 25%.

        Returns
        -------
        MetalsValuationInputs with projection fields populated and a
        high_leverage_flag if the company's balance sheet is stretched.

        Notes
        -----
        If net_debt_to_ebitda > 4x in the latest historical year, a warning
        is logged. The model does not automatically apply a higher discount rate
        for leverage; the caller should adjust WACC or the equity risk premium.
        """
        val_inputs = self.prepare_valuation_inputs(proj_years=proj_years)
        nf = self._nf or self.normalize()

        # Base revenue: last historical data point
        base_revenue = val_inputs.net_revenue_hist[-1] if val_inputs.net_revenue_hist else 0.0

        # Combined top-line growth: volume * realization compounding
        combined_growth = (1.0 + volume_growth) * (1.0 + realization_growth) - 1.0

        # Last-3yr average D&A for EBIT approximation
        dep_values = [v for v in (nf.depreciation or []) if v is not None]
        avg_da = (sum(dep_values[-3:]) / len(dep_values[-3:])) if dep_values else 0.0

        # Last-3yr average finance costs (interest) — material for metals cos
        fin_costs = [v for v in (nf.finance_costs or []) if v is not None]
        avg_finance_cost = (sum(fin_costs[-3:]) / len(fin_costs[-3:])) if fin_costs else 0.0

        # FCF conversion factor: historical average, clamped to [40%, 90%]
        # Metals companies invest heavily in capex, so FCF conversion is lower than FMCG.
        fcf_conversions = [
            v for v in (self._ratios.fcf_conversion if self._ratios else [])
            if v is not None
        ]
        if fcf_conversions:
            fcf_factor = sum(fcf_conversions[-3:]) / len(fcf_conversions[-3:]) / 100.0
        else:
            fcf_factor = 0.60  # conservative default for capital-intensive metals
        fcf_factor = max(0.40, min(0.90, fcf_factor))

        shares = val_inputs.shares_outstanding

        net_revenue_proj: list[float | None] = []
        ebitda_proj:      list[float | None] = []
        pat_proj:         list[float | None] = []
        fcf_proj:         list[float | None] = []
        eps_proj:         list[float | None] = []

        current_revenue = base_revenue
        for _ in range(proj_years):
            # Revenue projection: volume growth + realization growth
            current_revenue = current_revenue * (1.0 + combined_growth)
            net_revenue_proj.append(round(current_revenue, 2))

            # EBITDA
            ebitda = current_revenue * margin_pct
            ebitda_proj.append(round(ebitda, 2))

            # EBIT: subtract D&A (held constant at historical average)
            ebit = ebitda - avg_da

            # PBT: subtract finance costs (debt repayment assumed flat; conservative)
            pbt = ebit - avg_finance_cost

            # PAT
            pat = max(pbt * (1.0 - tax_rate), 0.0)
            pat_proj.append(round(pat, 2))

            # FCF: PAT * conversion factor
            fcf = pat * fcf_factor
            fcf_proj.append(round(fcf, 2))

            # EPS
            if shares > 0:
                eps_proj.append(round(pat / shares, 2))
            else:
                eps_proj.append(None)

        # Log risk flag for high-leverage companies
        if val_inputs.high_leverage_flag:
            logger.warning(
                "%s: HIGH LEVERAGE — projections assume debt remains ~flat. "
                "Equity value highly sensitive to realization assumptions.",
                val_inputs.ticker,
            )

        val_inputs.net_revenue_proj = net_revenue_proj
        val_inputs.ebitda_proj      = ebitda_proj
        val_inputs.pat_proj         = pat_proj
        val_inputs.fcf_proj         = fcf_proj
        val_inputs.eps_proj         = eps_proj

        return val_inputs

    # ------------------------------------------------------------------
    # Volume extraction
    # ------------------------------------------------------------------

    def _estimate_volume(
        self, nf: NormalizedFinancials
    ) -> tuple[list[float | None], list[float | None]]:
        """
        Attempt to extract production volumes from company_data.

        Strategy (in order of preference):
        1. Explicit volume row in P&L notes or a 'volumes' DataFrame attribute
        2. Segment data that labels a row as "Volume (MT)" or similar
        3. Ticker-specific heuristics for TATASTEEL / JSWSTEEL (crude steel)
        4. Return all-None if volume cannot be determined

        For non-steel companies (e.g. NMDC, COALINDIA), steel_volume is None.
        For non-aluminium companies, aluminium_volume is None.

        Returns
        -------
        (steel_volume_mt, aluminium_volume_mt)
            Both are per-FY lists of float | None, in million tonnes.
        """
        n = len(nf.fy_labels)
        null_series: list[float | None] = [None] * n

        ticker = nf.ticker.upper()

        # Determine which metals this company produces
        is_steel_producer     = ticker in {"TATASTEEL", "JSWSTEEL", "SAIL", "APLAPOLLO"}
        is_aluminium_producer = ticker in {"HINDALCO", "NATIONALUM", "VEDL"}

        if not (is_steel_producer or is_aluminium_producer):
            # Mining / diversified companies: volume not separately tracked here
            return null_series[:], null_series[:]

        steel_vol     = null_series[:]
        aluminium_vol = null_series[:]

        # -- Attempt 1: explicit volumes DataFrame on company_data --
        vol_df = getattr(self.company_data, "volumes", None)
        if vol_df is not None and hasattr(vol_df, "index"):
            row_index = {str(idx).strip().lower(): idx for idx in vol_df.index}

            steel_labels = [
                "crude steel production", "steel volume", "steel production",
                "finished steel", "total steel", "volume (mt)", "production volume",
            ]
            alum_labels = [
                "aluminium production", "aluminium volume", "primary aluminium",
                "aluminum production", "total aluminium",
            ]

            if is_steel_producer:
                steel_vol = self._extract_volume_row(vol_df, row_index, steel_labels, nf.fy_labels)
            if is_aluminium_producer:
                aluminium_vol = self._extract_volume_row(vol_df, row_index, alum_labels, nf.fy_labels)

            if any(v is not None for v in steel_vol) or any(v is not None for v in aluminium_vol):
                return steel_vol, aluminium_vol

        # -- Attempt 2: P&L "Other" row sometimes contains volume for screener.in --
        pnl_df = self._get_df("pnl")
        if pnl_df is not None:
            row_index_pnl = {str(idx).strip().lower(): idx for idx in pnl_df.index}
            vol_labels_pnl = ["volume", "sales volume", "production (mt)", "volume (mt)"]
            extracted = self._extract_volume_row(pnl_df, row_index_pnl, vol_labels_pnl, nf.fy_labels)
            if any(v is not None for v in extracted):
                if is_steel_producer:
                    steel_vol = extracted
                elif is_aluminium_producer:
                    aluminium_vol = extracted
                return steel_vol, aluminium_vol

        # -- Attempt 3: ticker-specific crude steel production estimates --
        # These are approximate long-run production estimates derived from
        # public annual report disclosures. They should be replaced with
        # actuals when screener.in or notes provide them.
        if is_steel_producer:
            steel_vol = self._heuristic_steel_volume(ticker, n, nf.fy_labels)
        if is_aluminium_producer:
            aluminium_vol = self._heuristic_aluminium_volume(ticker, n, nf.fy_labels)

        return steel_vol, aluminium_vol

    def _extract_volume_row(
        self,
        df,
        row_index: dict[str, object],
        label_variants: list[str],
        fy_labels: list[str],
    ) -> list[float | None]:
        """
        Extract a volume time-series row from a DataFrame.
        Returns all-None if no matching label is found.
        The values are returned as-is (assumed to be in million tonnes already).
        """
        for label in label_variants:
            key = label.lower()
            if key in row_index:
                raw_idx = row_index[key]
                row = df.loc[raw_idx]
                series: list[float | None] = []
                for fy in fy_labels:
                    val = None
                    for col_key in self._fy_to_col_keys(fy, df.columns):
                        if col_key in df.columns:
                            val = _to_float(row[col_key])
                            break
                    series.append(val)
                logger.debug("Extracted volume row '%s' from DataFrame", label)
                return series
        return [None] * len(fy_labels)

    def _heuristic_steel_volume(
        self, ticker: str, n: int, fy_labels: list[str]
    ) -> list[float | None]:
        """
        Rough production estimates for major steel producers when volume is
        not explicitly disclosed in the screener.in data.

        These are long-run averages from annual reports and should be treated
        as approximate only. The returned series uses a flat estimate scaled
        by the available FY count.

        Values in million tonnes (MT).
        """
        # Approximate latest-year capacity/production (FY24 basis)
        capacity_estimates: dict[str, float] = {
            "TATASTEEL": 21.0,   # India standalone ~21 MT
            "JSWSTEEL":  23.0,   # ~23 MT consolidated
            "SAIL":      18.0,   # ~18 MT
            "APLAPOLLO":  3.5,   # structural steel tubes ~3.5 MT
        }
        base_mt = capacity_estimates.get(ticker.upper())
        if base_mt is None:
            return [None] * n

        # Scale backward at ~5% annual growth to approximate historical volumes
        result: list[float | None] = []
        years_from_end = list(range(n - 1, -1, -1))  # [n-1, n-2, ..., 0]
        for yfe in reversed(years_from_end):
            # volume = base / (1.05 ^ years_from_end)
            vol = base_mt / ((1.05) ** yfe)
            result.append(round(vol, 3))

        logger.debug(
            "%s: using heuristic steel volume (base=%.1f MT, n=%d years)",
            ticker, base_mt, n,
        )
        return result

    def _heuristic_aluminium_volume(
        self, ticker: str, n: int, fy_labels: list[str]
    ) -> list[float | None]:
        """
        Rough primary aluminium production estimates.
        Values in million tonnes (MT). FY24 basis.
        """
        capacity_estimates: dict[str, float] = {
            "HINDALCO":   1.30,   # Hindalco India smelter ~1.3 MT
            "NATIONALUM": 0.46,   # NALCO ~0.46 MT
            "VEDL":       2.30,   # Vedanta aluminium segment ~2.3 MT
        }
        base_mt = capacity_estimates.get(ticker.upper())
        if base_mt is None:
            return [None] * n

        result: list[float | None] = []
        years_from_end = list(range(n - 1, -1, -1))
        for yfe in reversed(years_from_end):
            vol = base_mt / ((1.03) ** yfe)  # 3% growth assumption (aluminium slower)
            result.append(round(vol, 3))

        logger.debug(
            "%s: using heuristic aluminium volume (base=%.2f MT, n=%d years)",
            ticker, base_mt, n,
        )
        return result

    # ------------------------------------------------------------------
    # Per-tonne economics
    # ------------------------------------------------------------------

    def _compute_total_volume(
        self,
        steel_vol: list[float | None],
        aluminium_vol: list[float | None],
        n: int,
    ) -> list[float | None]:
        """
        Total production volume = steel_vol + aluminium_vol (element-wise).
        For companies that produce only one metal, returns the non-None series.
        Volumes in million tonnes.
        """
        total: list[float | None] = []
        for sv, av in zip(steel_vol, aluminium_vol):
            if sv is not None and av is not None:
                total.append(sv + av)
            elif sv is not None:
                total.append(sv)
            elif av is not None:
                total.append(av)
            else:
                total.append(None)
        return total

    def _compute_per_tonne_metrics(
        self,
        nf: NormalizedFinancials,
        total_volume_mt: list[float | None],
    ) -> tuple[list[float | None], list[float | None]]:
        """
        Compute realization per tonne and EBITDA per tonne.

        Conversion note: net_revenue is in INR Crores; total_volume is in
        million tonnes (MT = 10^6 tonnes). To get INR per tonne:
            (net_revenue_cr * 1e7) / (volume_mt * 1e6)
            = net_revenue_cr * 10 / volume_mt

        Returns
        -------
        (realization_per_tonne, ebitda_per_tonne)
            Both in INR per tonne.
        """
        realization_per_tonne: list[float | None] = []
        ebitda_per_tonne:      list[float | None] = []

        for nr, eb, vol_mt in zip(nf.net_revenue, nf.ebitda, total_volume_mt):
            if vol_mt is None or vol_mt == 0:
                realization_per_tonne.append(None)
                ebitda_per_tonne.append(None)
                continue

            # Convert: Crores → units, then divide by MT
            # 1 Crore = 1e7 ; 1 MT = 1e6 tonnes → factor = 1e7 / 1e6 = 10
            conv_factor = 10.0

            r_pt = safe_div(nr * conv_factor, vol_mt) if nr is not None else None
            e_pt = safe_div(eb * conv_factor, vol_mt) if eb is not None else None

            realization_per_tonne.append(round(r_pt, 0) if r_pt is not None else None)
            ebitda_per_tonne.append(round(e_pt, 0) if e_pt is not None else None)

        return realization_per_tonne, ebitda_per_tonne

    # ------------------------------------------------------------------
    # Segment extraction
    # ------------------------------------------------------------------

    def _extract_segments(self, nf: NormalizedFinancials) -> dict[str, dict]:
        """
        Parse segment data from company_data if available.

        Looks for a 'segments', 'segment_data', or 'segment_results' attribute
        on company_data. Recognises the following segment types:
        - Steel / Flat Steel / Long Steel
        - Aluminium / Aluminum
        - Copper / Copper Products
        - Zinc / Zinc-Lead
        - Mining / Iron Ore Mining / Coal Mining
        - Power / Energy

        Returns
        -------
        dict mapping canonical segment name → {
            "fy_labels": [...],
            "revenue": [...],
            "ebit": [...],
            "ebit_margin": [...]
        }
        Returns empty dict if no segment data is available (not an error).
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
            logger.debug(
                "No segment data found for %s — segment dict will be empty",
                nf.ticker,
            )
            return segments

        fy_labels = nf.fy_labels

        if hasattr(seg_data, "items"):
            # dict of {segment_name: DataFrame}
            for seg_name, seg_df in seg_data.items():
                canonical = self._canonicalize_segment_name(str(seg_name))
                segments[canonical] = self._extract_segment_metrics(
                    seg_df, fy_labels, canonical
                )
        elif hasattr(seg_data, "index"):
            # Single DataFrame: segments as rows
            for seg_name in seg_data.index:
                canonical = self._canonicalize_segment_name(str(seg_name))
                row_df = seg_data.loc[[seg_name]]
                segments[canonical] = self._extract_segment_metrics(
                    row_df, fy_labels, canonical
                )
        else:
            logger.warning(
                "Unrecognized segment data format for %s", nf.ticker
            )

        return segments

    def _canonicalize_segment_name(self, raw_name: str) -> str:
        """
        Map raw segment names to canonical metals sector names.

        Examples
        --------
        "flat steel" → "Steel"
        "primary aluminium" → "Aluminium"
        "zinc-lead" → "Zinc"
        """
        lower = raw_name.strip().lower()

        for label in _STEEL_SEGMENT_LABELS:
            if label in lower:
                return "Steel"
        for label in _ALUMINIUM_SEGMENT_LABELS:
            if label in lower:
                return "Aluminium"
        for label in _COPPER_SEGMENT_LABELS:
            if label in lower:
                return "Copper"
        for label in _ZINC_SEGMENT_LABELS:
            if label in lower:
                return "Zinc"
        for label in _MINING_SEGMENT_LABELS:
            if label in lower:
                return "Mining"
        for label in _POWER_SEGMENT_LABELS:
            if label in lower:
                return "Power"

        # Return title-cased original if no canonical match found
        return raw_name.strip().title()

    def _extract_segment_metrics(
        self,
        seg_df,
        fy_labels: list[str],
        seg_name: str,
    ) -> dict:
        """
        Extract revenue, EBIT, and EBIT margin for a single segment DataFrame.

        Handles DataFrames where the rows are labeled 'Revenue', 'EBIT', etc.
        Returns a dict with keys: 'fy_labels', 'revenue', 'ebit', 'ebit_margin'.
        All monetary values in INR Crores; margins in %.
        """
        result: dict = {
            "fy_labels":   fy_labels,
            "revenue":     [None] * len(fy_labels),
            "ebit":        [None] * len(fy_labels),
            "ebit_margin": [None] * len(fy_labels),
        }

        if seg_df is None or not hasattr(seg_df, "columns"):
            return result

        row_index = {str(idx).strip().lower(): idx for idx in seg_df.index}

        rev_labels  = [
            "revenue", "net revenue", "segment revenue", "sales",
            "net sales", "segment sales",
        ]
        ebit_labels = [
            "ebit", "segment result", "segment profit", "operating profit",
            "profit before interest and tax", "pbit",
        ]

        def _get_row(labels: list[str]) -> list[float | None]:
            for lbl in labels:
                if lbl.lower() in row_index:
                    raw_idx = row_index[lbl.lower()]
                    row = seg_df.loc[raw_idx]
                    series: list[float | None] = []
                    for fy in fy_labels:
                        val = None
                        for col_key in self._fy_to_col_keys(fy, seg_df.columns):
                            if col_key in seg_df.columns:
                                val = _to_float(row[col_key])
                                break
                        series.append(val)
                    return series
            return [None] * len(fy_labels)

        revenue = _get_row(rev_labels)
        ebit    = _get_row(ebit_labels)
        ebit_margin = [
            safe_div(e, r) * 100.0 if (e is not None and r) else None
            for e, r in zip(ebit, revenue)
        ]

        result["revenue"]     = revenue
        result["ebit"]        = ebit
        result["ebit_margin"] = ebit_margin
        return result

    # ------------------------------------------------------------------
    # Raw material cost helpers
    # ------------------------------------------------------------------

    def _extract_raw_material_cost(
        self, nf: NormalizedFinancials
    ) -> list[float | None]:
        """
        Extract absolute raw material cost from P&L.

        For steel companies: looks for coking coal + iron ore lines or a
        combined 'raw materials consumed' line.
        For aluminium: bauxite + power costs are the primary inputs.

        Falls back to nf.raw_materials (already extracted by BaseModel) if no
        more granular breakdown is available. Returns INR Crores.
        """
        # BaseModel already attempts to extract 'raw material' row
        if nf.raw_materials and any(v is not None for v in nf.raw_materials):
            return nf.raw_materials

        # Try additional label variants not in the base label map
        fy_labels = nf.fy_labels
        for label in _RAW_MATERIAL_LABELS:
            series = self._extract_series("pnl", [label], fy_labels)
            if any(v is not None for v in series):
                logger.debug(
                    "%s: raw material cost extracted via label '%s'",
                    nf.ticker, label,
                )
                return series

        logger.debug("%s: raw material cost not found; returning None series", nf.ticker)
        return [None] * len(fy_labels)

    # ------------------------------------------------------------------
    # Leverage risk check
    # ------------------------------------------------------------------

    def _check_high_leverage(
        self, net_debt_to_ebitda: list[float | None]
    ) -> bool:
        """
        Return True if the latest available net_debt_to_ebitda exceeds
        HIGH_LEVERAGE_THRESHOLD (4.0x).

        Uses the most recent non-None value in the series. Returns False
        if the series is entirely None (e.g. net-cash company, or data gap).
        """
        for val in reversed(net_debt_to_ebitda):
            if val is not None:
                return val > self.HIGH_LEVERAGE_THRESHOLD
        return False

    # ------------------------------------------------------------------
    # WACC / beta helpers
    # ------------------------------------------------------------------

    def _get_sector_beta(self, ticker: str) -> float:
        """
        Return the metals sector default beta for the given ticker.
        Falls back to DEFAULT_METALS_BETA if ticker is not in the lookup.
        """
        return METALS_SECTOR_BETAS.get(ticker.upper(), DEFAULT_METALS_BETA)

    # ------------------------------------------------------------------
    # Valuation helpers
    # ------------------------------------------------------------------

    def _select_peers(self, ticker: str) -> list[str]:
        """
        Select the appropriate peer group for comparables valuation.

        Steel producers → PEERS_STEEL
        Aluminium producers → PEERS_ALUMINIUM
        Mining companies → PEERS_MINING
        Others → PEERS_DIVERSIFIED
        """
        t = ticker.upper()

        if t in {"TATASTEEL", "JSWSTEEL", "SAIL", "APLAPOLLO"}:
            return [p for p in self.PEERS_STEEL if p != t]

        if t in {"HINDALCO", "NATIONALUM"}:
            return [p for p in self.PEERS_ALUMINIUM if p != t]

        if t in {"NMDC", "COALINDIA", "MOIL"}:
            return [p for p in self.PEERS_MINING if p != t]

        # VEDL is diversified; use diversified peer set
        return [p for p in self.PEERS_DIVERSIFIED if p != t]

    def _generate_proj_fy_labels(self, last_hist_fy: str, proj_years: int) -> list[str]:
        """
        Generate projection FY labels starting the year after last_hist_fy.

        Examples
        --------
        last_hist_fy="FY25", proj_years=5 → ["FY26", "FY27", "FY28", "FY29", "FY30"]
        last_hist_fy="FY30", proj_years=3 → ["FY31", "FY32", "FY33"]
        """
        if last_hist_fy.startswith("FY") and len(last_hist_fy) == 4:
            try:
                last_yy = int(last_hist_fy[2:])
                labels: list[str] = []
                for i in range(1, proj_years + 1):
                    yy = (last_yy + i) % 100
                    labels.append(f"FY{yy:02d}")
                return labels
            except ValueError:
                pass
        # Fallback for unexpected FY label format
        return [f"Proj+{i}" for i in range(1, proj_years + 1)]


# ---------------------------------------------------------------------------
# Module-level convenience function
# ---------------------------------------------------------------------------

def build_metals_model(
    company_data: "CompanyData",
    market_data: dict | None = None,
) -> tuple[MetalsModel, MetalsMetrics]:
    """
    Convenience function: construct a MetalsModel and run compute_metals_metrics()
    in one call.

    Parameters
    ----------
    company_data : CompanyData
        Output of ScreenerClient.
    market_data : dict, optional
        Market inputs: 'market_cap_cr', 'beta', 'shares_outstanding',
        'cost_of_debt', 'price', 'ev_ebitda_target', 'ev_per_tonne'.

    Returns
    -------
    (MetalsModel, MetalsMetrics)
        The model instance (with _nf and _ratios populated) and the metrics.

    Example
    -------
    >>> model, metrics = build_metals_model(company_data, {"market_cap_cr": 80000})
    >>> print(metrics.ebitda_margin[-1])      # latest FY EBITDA margin
    >>> print(metrics.net_debt_to_ebitda[-1]) # latest leverage
    >>> print(metrics.ebitda_per_tonne[-1])   # latest EBITDA/tonne
    """
    model = MetalsModel(company_data, market_data=market_data)
    metrics = model.compute_metals_metrics()
    return model, metrics
