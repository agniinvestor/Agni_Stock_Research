"""
Pharma sector financial model for Indian listed companies.

Extends BaseModel with pharma-specific metrics:
- R&D spend extraction (% of revenue, absolute INR Cr)
- Adjusted EBITDA excluding R&D capitalisation effects
- Domestic vs. exports/US revenue split from segment disclosures
- DSO analysis (pharma receivables tend to be longer than FMCG)
- ANDA pipeline maturation driven margin expansion in projections

Covers: SUNPHARMA, DRREDDY, CIPLA, DIVISLAB, LUPIN, AUROPHARMA,
        TORNTPHARM, ALKEM, IPCALAB, NATCO (and unlisted peers)

All monetary values in INR Crores.
All lists indexed 0 = oldest FY.
None for missing data (not 0).
Percentages stored as 0–100 floats (e.g. 12.5 not 0.125).
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
# Module-level constants
# ---------------------------------------------------------------------------

# Sector betas for Indian pharma companies (leveraged, as-listed beta).
# Source: historical regression against Nifty 50; conservative estimates
# reflecting defensive nature of pharma vs. cyclical sectors.
PHARMA_SECTOR_BETAS: dict[str, float] = {
    "SUNPHARMA":  0.55,   # Largest, most diversified; lower beta
    "DRREDDY":    0.65,   # US generic exposure adds moderate volatility
    "CIPLA":      0.60,   # Domestic + chronic therapy focus; stable
    "DIVISLAB":   0.50,   # API/CRAMS pureplay; lowest beta in sector
    "LUPIN":      0.70,   # US generics + Japan; ANDA-driven volatility
    "AUROPHARMA": 0.75,   # High US exposure; generics pricing risk
    "TORNTPHARM": 0.55,   # Domestic chronic + Germany; low volatility
    "ALKEM":      0.60,   # Domestic acute + chronic mix; mid beta
    "IPCALAB":    0.65,   # Branded generics + API; moderate
    "NATCO":      0.80,   # Specialty/oncology; highest beta (pipeline risk)
}

DEFAULT_PHARMA_BETA: float = 0.65

# R&D spend as % of revenue: valid range for large Indian pharma.
# Used for validation / warning only — data is never discarded on breach.
RND_PCT_LOWER_BOUND: float = 2.0   # Some smaller cos spend less
RND_PCT_UPPER_BOUND: float = 18.0  # Frontier R&D companies may exceed 15%
RND_PCT_TYPICAL_LOWER: float = 5.0
RND_PCT_TYPICAL_UPPER: float = 15.0

# Default terminal growth rate for Indian pharma DCF (% nominal).
# India pharma market grows ~10% nominal; conservative perpetuity assumption.
PHARMA_TERMINAL_GROWTH: float = 5.0

# Peer groups by sub-segment
PEERS_BRANDED_DOMESTIC = ["SUNPHARMA", "CIPLA", "TORNTPHARM", "ALKEM", "IPCALAB"]
PEERS_US_GENERICS      = ["DRREDDY", "LUPIN", "AUROPHARMA", "SUNPHARMA", "CIPLA"]
PEERS_API_CRAMS        = ["DIVISLAB", "DRREDDY", "AUROPHARMA", "NATCO"]

# Screener.in P&L label variants that may contain R&D in pharma financials.
# Order matters: first match wins.
_RND_LABEL_VARIANTS: list[str] = [
    "research & development",
    "research and development",
    "r & d expenses",
    "r&d expenses",
    "r and d expenses",
    "r&d expenditure",
    "research & development expenses",
    "research expenses",
    "product development expenses",
]

# Other manufacturing expense labels — in some screener reports R&D is
# embedded here and not broken out separately.
_OTHER_MFR_LABEL_VARIANTS: list[str] = [
    "other mfr. exp",
    "other manufacturing expenses",
    "other manufacturing exp",
    "manufacturing expenses",
]

# Segment revenue labels for domestic / exports / US parsing
_DOMESTIC_SEGMENT_KEYWORDS: list[str] = [
    "india", "domestic", "formulations india", "india business",
    "branded generics", "domestic formulations",
]
_EXPORT_SEGMENT_KEYWORDS: list[str] = [
    "export", "exports", "international", "rest of world", "row",
    "emerging markets", "global generics",
]
_US_SEGMENT_KEYWORDS: list[str] = [
    "usa", "us", "north america", "united states", "us generics",
    "north america generics", "us formulations",
]


# ---------------------------------------------------------------------------
# Dataclasses
# ---------------------------------------------------------------------------

@dataclass
class PharmaMetrics:
    """
    Pharma-specific metrics extending base ratios.

    All monetary fields in INR Crores.
    All percentage fields as floats in 0–100 range.
    All lists are aligned to fy_labels (index 0 = oldest FY).
    None for any data point that could not be determined.
    """

    fy_labels: list[str]

    # ---- Standard P&L metrics ----
    net_revenue:   list[float | None]   # INR Cr, net of excise/GST
    ebitda:        list[float | None]   # INR Cr, operational EBITDA
    pat:           list[float | None]   # INR Cr, PAT after minority interest
    eps:           list[float | None]   # INR per share

    # ---- Profitability margins ----
    ebitda_margin: list[float | None]   # EBITDA / net_revenue, %
    pat_margin:    list[float | None]   # PAT / net_revenue, %

    # ---- R&D metrics (pharma-specific) ----
    rnd_spend:           list[float | None]  # Absolute R&D spend, INR Cr
    rnd_spend_pct:       list[float | None]  # R&D as % of net revenue
    ebitda_ex_rnd:       list[float | None]  # EBITDA + R&D (adjusted EBITDA)
    ebitda_ex_rnd_margin: list[float | None] # ebitda_ex_rnd / net_revenue, %

    # ---- Revenue mix (latest FY estimates) ----
    domestic_revenue_pct: float | None   # % of revenue from domestic market
    us_revenue_pct:       float | None   # % of revenue from USA / exports

    # ---- Return metrics ----
    roce: list[float | None]   # NOPAT / Capital Employed, %
    roe:  list[float | None]   # PAT / Average Equity, %

    # ---- Cash flow quality ----
    fcf:            list[float | None]   # Free cash flow, INR Cr
    fcf_conversion: list[float | None]   # FCF / PAT, %

    # ---- Working capital ----
    dso: list[float | None]   # Debtor days (receivables / daily revenue)

    # ---- Segment data ----
    segments: dict[str, dict]
    # Structure: {
    #   "India Formulations": {"fy_labels": [...], "revenue": [...], "ebit": [...], "ebit_margin": [...]},
    #   "US Generics":        {"fy_labels": [...], "revenue": [...], "ebit": [...], "ebit_margin": [...]},
    #   ...
    # }


@dataclass
class PharmaValuationInputs:
    """
    Inputs for DCF and comparables-based valuation of Indian pharma companies.

    Extends the FMCG valuation structure with pharma-specific multiples:
    - terminal_growth fixed at 5.0% (conservative for pharma)
    - pe_target: justified P/E for 12-month price target
    - ev_ebitda_target: EV/EBITDA secondary check (typical range 20–35x for large pharma)
    """

    ticker: str

    # Historical FY labels (list of str, e.g. ["FY16", ..., "FY25"])
    fy_labels_hist: list[str]

    # Projection FY labels (e.g. ["FY26", "FY27", "FY28", "FY29", "FY30"])
    fy_labels_proj: list[str]

    # Historical actuals (INR Cr; 0.0 where data unavailable)
    net_revenue_hist: list[float]
    ebitda_hist:      list[float]
    pat_hist:         list[float]
    fcf_hist:         list[float]

    # Projections (None until filled by estimate_projections or LLM)
    net_revenue_proj: list[float | None]
    ebitda_proj:      list[float | None]
    pat_proj:         list[float | None]
    fcf_proj:         list[float | None]
    eps_proj:         list[float | None]

    # DCF inputs
    wacc:             float | None   # Weighted average cost of capital, %
    terminal_growth:  float          # = PHARMA_TERMINAL_GROWTH = 5.0
    net_debt:         float          # Latest net debt, INR Cr (negative = net cash)
    shares_outstanding: float        # Shares in Crores

    # Comparables
    peer_tickers: list[str]

    # Pharma-specific valuation anchors
    pe_target:         float | None  # Justified trailing / forward P/E
    ev_ebitda_target:  float | None  # EV/EBITDA multiple for secondary cross-check


# ---------------------------------------------------------------------------
# PharmaModel
# ---------------------------------------------------------------------------

class PharmaModel(BaseModel):
    """
    Indian pharma sector financial model.

    Extends BaseModel with:
    1. R&D spend extraction from screener.in P&L rows (case-insensitive search)
    2. Adjusted EBITDA (EBITDA before R&D expense deduction)
    3. Domestic / exports / US revenue split from segment data
    4. Pharma-specific projection engine (12% revenue CAGR, ANDA margin expansion)
    5. WACC using CAPM with India risk-free rate 7.0%, ERP 5.5%

    Usage
    -----
    >>> model = PharmaModel(company_data, market_data)
    >>> metrics = model.compute_pharma_metrics()
    >>> val_inputs = model.prepare_valuation_inputs()
    >>> val_inputs = model.estimate_projections()
    """

    SECTOR = "Pharma"

    def __init__(
        self,
        company_data: "CompanyData",
        market_data: dict | None = None,
    ) -> None:
        """
        Parameters
        ----------
        company_data : CompanyData
            Output of ScreenerClient. Must have profit_loss, balance_sheet,
            cash_flow DataFrames and ticker attribute. Consolidated preferred.
        market_data : dict, optional
            May contain: 'beta', 'market_cap_cr', 'shares_outstanding',
            'price', 'cost_of_debt'. Sourced from NSE/BSE market data API.
        """
        super().__init__(company_data)
        self.market_data: dict = market_data or {}
        self._metrics: PharmaMetrics | None = None

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def compute_pharma_metrics(self) -> PharmaMetrics:
        """
        Full computation pipeline:
          1. normalize()           → NormalizedFinancials (stored as self._nf)
          2. compute_ratios()      → ComputedRatios       (stored as self._ratios)
          3. _extract_rnd_spend()  → R&D series
          4. _compute_ebitda_ex_rnd() → Adjusted EBITDA
          5. _extract_segments()  → Domestic / US / Export split
          6. Assemble PharmaMetrics

        Returns
        -------
        PharmaMetrics
            All pharma-specific metrics aligned to nf.fy_labels.

        Side effects
        ------------
        Sets self._nf, self._ratios, self._metrics.
        """
        # Step 1: normalize base financials
        nf = self.normalize()

        # Step 2: compute standard ratios
        ratios = self.compute_ratios(nf)

        # Step 3: R&D spend
        rnd_spend = self._extract_rnd_spend(nf)
        rnd_spend_pct = self._compute_rnd_pct(rnd_spend, nf.net_revenue)

        # Validate R&D % against known pharma range and log warnings
        self._validate_rnd_pct(rnd_spend_pct, nf.ticker)

        # Step 4: Adjusted EBITDA (EBITDA before R&D deduction)
        ebitda_ex_rnd, ebitda_ex_rnd_margin = self._compute_ebitda_ex_rnd(
            nf.ebitda, rnd_spend, nf.net_revenue
        )

        # Step 5: Segment parsing for revenue mix
        segments = self._extract_segments(nf)
        domestic_pct, us_pct = self._estimate_revenue_mix(nf, segments)

        # Step 6: Standard P&L metrics (pulled from nf / ratios)
        ebitda_margin = ratios.ebitda_margin
        pat_margin    = ratios.pat_margin

        # FCF conversion from ratios
        fcf_conversion = ratios.fcf_conversion

        metrics = PharmaMetrics(
            fy_labels             = nf.fy_labels,
            net_revenue           = nf.net_revenue,
            ebitda                = nf.ebitda,
            pat                   = nf.pat_after_mi,
            eps                   = nf.eps,
            ebitda_margin         = ebitda_margin,
            pat_margin            = pat_margin,
            rnd_spend             = rnd_spend,
            rnd_spend_pct         = rnd_spend_pct,
            ebitda_ex_rnd         = ebitda_ex_rnd,
            ebitda_ex_rnd_margin  = ebitda_ex_rnd_margin,
            domestic_revenue_pct  = domestic_pct,
            us_revenue_pct        = us_pct,
            roce                  = ratios.roce,
            roe                   = ratios.roe,
            fcf                   = nf.fcf,
            fcf_conversion        = fcf_conversion,
            dso                   = ratios.dso,
            segments              = segments,
        )

        self._metrics = metrics
        return metrics

    def prepare_valuation_inputs(self, proj_years: int = 5) -> PharmaValuationInputs:
        """
        Assemble historical data for valuation modelling.

        Projection fields are set to None; fill them via estimate_projections()
        or pass LLM-generated assumptions directly.

        WACC computation
        ----------------
        Uses CAPM via compute_wacc() from fmcg_model:
          - Risk-free rate: 7.0% (India 10yr G-sec)
          - ERP: 5.5% (Damodaran India estimate)
          - Beta: from market_data['beta'] if available, else PHARMA_SECTOR_BETAS
          - Debt weight: derived from net_debt / (net_debt + market_cap)

        Returns
        -------
        PharmaValuationInputs
            Historical fields populated; projection fields set to None.
        """
        # Ensure base normalization has been run
        nf = self._nf or self.normalize()

        ticker = nf.ticker

        def _clean(lst: list[float | None]) -> list[float]:
            """Replace None with 0.0 for numeric arrays passed to valuation."""
            return [v if v is not None else 0.0 for v in lst]

        net_revenue_hist = _clean(nf.net_revenue)
        ebitda_hist      = _clean(nf.ebitda)
        pat_hist         = _clean(nf.pat_after_mi)
        fcf_hist         = _clean(nf.fcf)

        # Latest net debt (negative = net cash position)
        net_debt_latest: float = 0.0
        if nf.net_debt and any(v is not None for v in nf.net_debt):
            for v in reversed(nf.net_debt):
                if v is not None:
                    net_debt_latest = v
                    break

        # Shares outstanding (Crores)
        shares_outstanding: float = float(self.market_data.get("shares_outstanding", 0.0))
        if shares_outstanding == 0.0:
            # Derive from market_cap and price if both available
            mktcap = float(self.market_data.get("market_cap_cr", 0.0))
            price  = float(self.market_data.get("price", 0.0))
            if mktcap > 0 and price > 0:
                shares_outstanding = mktcap / price  # result in Crores

        # Beta selection: market_data > PHARMA_SECTOR_BETAS > DEFAULT
        beta = float(self.market_data.get("beta", self._get_pharma_beta(ticker)))

        # Debt weight for WACC (capped at 50% to avoid extreme distortion)
        mktcap_cr = float(self.market_data.get("market_cap_cr", 0.0))
        debt_weight = 0.0
        if mktcap_cr > 0 and net_debt_latest > 0:
            total_capital = mktcap_cr + net_debt_latest
            debt_weight = min(max(net_debt_latest / total_capital, 0.0), 0.50)

        # WACC (% e.g. 11.5 not 0.115)
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

        # Peer selection based on business mix
        peers = self._select_pharma_peers(ticker)

        # Pharma-specific multiples (will be refined after estimate_projections)
        pe_target        = self._estimate_pe_target(ticker)
        ev_ebitda_target = self._estimate_ev_ebitda_target(ticker)

        val_inputs = PharmaValuationInputs(
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
            terminal_growth    = PHARMA_TERMINAL_GROWTH,
            net_debt           = net_debt_latest,
            shares_outstanding = shares_outstanding,
            peer_tickers       = peers,
            pe_target          = pe_target,
            ev_ebitda_target   = ev_ebitda_target,
        )
        return val_inputs

    def estimate_projections(
        self,
        revenue_growth: float = 0.12,
        margin_expansion: float = 0.5,
        rnd_growth: float = 0.08,
        proj_years: int = 5,
        tax_rate: float = 0.25,
    ) -> PharmaValuationInputs:
        """
        Simple pharma-specific projection engine.

        Revenue and margin assumptions
        --------------------------------
        Indian pharma revenue typically grows at ~12% CAGR (blended):
          - Domestic formulations: ~10% (volume ~5–6% + pricing ~4–5%)
          - Exports / US generics: ~15% (new ANDA approvals + market share gain)

        Margin expansion assumption
        ----------------------------
        As ANDA pipelines mature and first-to-file opportunities are realised,
        EBITDA margins expand. margin_expansion is the annual absolute
        improvement in EBITDA margin percentage points (e.g. 0.5pp per year).

        R&D assumption
        ---------------
        R&D grows at rnd_growth p.a. (default 8%) — slower than revenue,
        implying gradual R&D leverage as existing pipeline matures.

        PAT approximation
        ------------------
          EBIT = EBITDA - avg_D&A
          PAT  = EBIT * (1 - tax_rate)
          (Finance costs ignored for net-cash pharma; adjust manually if needed)

        FCF derivation
        ---------------
          FCF = PAT * fcf_conversion_factor
          where fcf_conversion_factor = 3yr avg historical FCF/PAT (clamped 50–100%).

        Parameters
        ----------
        revenue_growth : float
            Annual fractional revenue growth (default 0.12 = 12%).
        margin_expansion : float
            Annual absolute EBITDA margin improvement in percentage points
            (default 0.5pp per year).
        rnd_growth : float
            Annual fractional growth in R&D spend (default 0.08 = 8%).
        proj_years : int
            Number of projection years (default 5).
        tax_rate : float
            Effective tax rate as fraction (default 0.25).

        Returns
        -------
        PharmaValuationInputs
            All projection fields populated.
        """
        val_inputs = self.prepare_valuation_inputs(proj_years=proj_years)
        nf = self._nf or self.normalize()

        # Base revenue: last historical value
        base_revenue = val_inputs.net_revenue_hist[-1] if val_inputs.net_revenue_hist else 0.0

        # Base EBITDA margin (% as float 0–100): use last available historical value
        base_ebitda_margin = self._last_valid(
            [safe_div(e, r) * 100.0 if (e is not None and r) else None
             for e, r in zip(nf.ebitda, nf.net_revenue)]
        )
        if base_ebitda_margin is None:
            base_ebitda_margin = 20.0  # conservative fallback for large pharma

        # Average D&A (last 3 years) for EBIT derivation
        dep_values = [v for v in (nf.depreciation or []) if v is not None]
        avg_da = sum(dep_values[-3:]) / len(dep_values[-3:]) if dep_values else 0.0

        # Historical FCF conversion factor (last 3 years, clamped [50%, 100%])
        ratios = self._ratios  # may be None if compute_pharma_metrics not called
        if ratios is not None:
            fcf_conv_hist = [v for v in ratios.fcf_conversion if v is not None]
        else:
            fcf_conv_hist = []
        fcf_factor = (
            sum(fcf_conv_hist[-3:]) / len(fcf_conv_hist[-3:]) / 100.0
            if fcf_conv_hist else 0.80  # pharma default slightly below FMCG
        )
        fcf_factor = max(0.50, min(1.0, fcf_factor))

        # Shares outstanding for EPS
        shares = val_inputs.shares_outstanding

        # Build projection lists
        net_revenue_proj: list[float | None] = []
        ebitda_proj:      list[float | None] = []
        pat_proj:         list[float | None] = []
        fcf_proj:         list[float | None] = []
        eps_proj:         list[float | None] = []

        current_revenue       = base_revenue
        current_ebitda_margin = base_ebitda_margin   # % (0–100)

        for year_idx in range(proj_years):
            # ---- Revenue ----
            current_revenue = current_revenue * (1.0 + revenue_growth)
            net_revenue_proj.append(round(current_revenue, 2))

            # ---- EBITDA ----
            # Margin expands by margin_expansion pp each year
            current_ebitda_margin = current_ebitda_margin + margin_expansion
            ebitda_val = current_revenue * (current_ebitda_margin / 100.0)
            ebitda_proj.append(round(ebitda_val, 2))

            # ---- EBIT and PAT ----
            ebit = ebitda_val - avg_da
            pat  = max(ebit * (1.0 - tax_rate), 0.0)
            pat_proj.append(round(pat, 2))

            # ---- FCF ----
            fcf = pat * fcf_factor
            fcf_proj.append(round(fcf, 2))

            # ---- EPS ----
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
    # R&D extraction
    # ------------------------------------------------------------------

    def _extract_rnd_spend(self, nf: NormalizedFinancials) -> list[float | None]:
        """
        Extract R&D spend time series from the P&L DataFrame.

        Search strategy (priority order):
        1. Look for explicit "research & development" (or variant) row — case-insensitive.
        2. Look for "r & d expenses" or "r&d expenses" variants.
        3. Fall back to "other mfr. exp" / "other manufacturing expenses" row.
           (For many screener.in reports, R&D is lumped here and not broken out.)
        4. If nothing found, return list of None.

        Note: Step 3 is a last resort. If "other mfr. exp" is used, the full
        value is returned and a DEBUG-level warning is logged, because the row
        may contain more than just R&D.

        Parameters
        ----------
        nf : NormalizedFinancials
            Already-normalized financials (used only for fy_labels and ticker).

        Returns
        -------
        list[float | None]
            Absolute R&D spend in INR Cr, aligned to nf.fy_labels.
        """
        fy_labels = nf.fy_labels
        n = len(fy_labels)

        df = self._get_df("pnl")
        if df is None:
            logger.warning("No P&L DataFrame available for %s; R&D = None", nf.ticker)
            return [None] * n

        # Build a lowercase-stripped row index for case-insensitive lookup
        row_index: dict[str, object] = {}
        for idx in df.index:
            row_index[str(idx).strip().lower()] = idx

        # --- Priority 1 & 2: explicit R&D labels ---
        for variant in _RND_LABEL_VARIANTS:
            key = variant.lower()
            if key in row_index:
                series = self._read_pnl_row(df, row_index[key], fy_labels)
                logger.debug(
                    "R&D extracted from explicit label '%s' for %s", variant, nf.ticker
                )
                return [abs(v) if v is not None else None for v in series]

        # --- Priority 3: partial / fuzzy match inside row_index keys ---
        # Some screener reports write "r&d" or "r & d" inconsistently
        for key in row_index:
            if ("r&d" in key or "r & d" in key or "research" in key):
                raw_idx = row_index[key]
                series = self._read_pnl_row(df, raw_idx, fy_labels)
                logger.debug(
                    "R&D extracted from fuzzy-matched label '%s' for %s", key, nf.ticker
                )
                return [abs(v) if v is not None else None for v in series]

        # --- Priority 4: "other mfr. exp" as proxy (fallback only) ---
        for variant in _OTHER_MFR_LABEL_VARIANTS:
            key = variant.lower()
            if key in row_index:
                series = self._read_pnl_row(df, row_index[key], fy_labels)
                logger.debug(
                    "R&D proxied from '%s' for %s — may overstate R&D. "
                    "Verify annual report disclosure.",
                    variant, nf.ticker,
                )
                return [abs(v) if v is not None else None for v in series]

        # No R&D data found
        logger.debug(
            "No R&D row found in P&L for %s; tried %d label variants",
            nf.ticker, len(_RND_LABEL_VARIANTS) + len(_OTHER_MFR_LABEL_VARIANTS),
        )
        return [None] * n

    def _read_pnl_row(
        self, df, raw_idx: object, fy_labels: list[str]
    ) -> list[float | None]:
        """
        Read a single P&L row from df into a float list aligned to fy_labels.

        Parameters
        ----------
        df : DataFrame
            The P&L DataFrame with FY columns.
        raw_idx : object
            The exact index key as it appears in df.index (not lowercased).
        fy_labels : list[str]
            Target FY labels (e.g. ["FY16", ..., "FY25"]).

        Returns
        -------
        list[float | None]
        """
        row = df.loc[raw_idx]
        result: list[float | None] = []
        for fy in fy_labels:
            val = None
            for col_key in self._fy_to_col_keys(fy, df.columns):
                if col_key in df.columns:
                    val = _to_float(row[col_key])
                    break
            result.append(val)
        return result

    # ------------------------------------------------------------------
    # Adjusted EBITDA (EBITDA ex-R&D)
    # ------------------------------------------------------------------

    def _compute_ebitda_ex_rnd(
        self,
        ebitda: list[float | None],
        rnd_spend: list[float | None],
        net_revenue: list[float | None],
    ) -> tuple[list[float | None], list[float | None]]:
        """
        Compute EBITDA before R&D deduction (adjusted / cash EBITDA).

        Rationale: R&D is expensed in the P&L under Indian GAAP / Ind AS.
        Adding it back reveals the underlying operating cash generation
        before discretionary R&D investment decisions.

        EBITDA_ex_RnD = EBITDA + R&D spend
        Margin = EBITDA_ex_RnD / net_revenue * 100

        Parameters
        ----------
        ebitda : list[float | None]
            Standard EBITDA series, INR Cr.
        rnd_spend : list[float | None]
            R&D spend series, INR Cr (positive values expected).
        net_revenue : list[float | None]
            Net revenue series, INR Cr.

        Returns
        -------
        (ebitda_ex_rnd, ebitda_ex_rnd_margin)
            Both lists aligned to the same fy_labels.
        """
        ebitda_ex_rnd: list[float | None] = []
        ebitda_ex_rnd_margin: list[float | None] = []

        for e, rnd, nr in zip(ebitda, rnd_spend, net_revenue):
            if e is not None and rnd is not None:
                adj = e + abs(rnd)
                ebitda_ex_rnd.append(adj)
                ebitda_ex_rnd_margin.append(
                    safe_div(adj, nr) * 100.0 if nr else None
                )
            elif e is not None:
                # R&D not available — return plain EBITDA as best estimate
                ebitda_ex_rnd.append(e)
                ebitda_ex_rnd_margin.append(
                    safe_div(e, nr) * 100.0 if nr else None
                )
            else:
                ebitda_ex_rnd.append(None)
                ebitda_ex_rnd_margin.append(None)

        return ebitda_ex_rnd, ebitda_ex_rnd_margin

    # ------------------------------------------------------------------
    # Segment extraction
    # ------------------------------------------------------------------

    def _extract_segments(self, nf: NormalizedFinancials) -> dict[str, dict]:
        """
        Parse segment data from company_data if available.

        Looks for a 'segments', 'segment_data', or 'segment_results' attribute
        on company_data. Expected to be a dict of DataFrames keyed by segment
        name, or a single DataFrame with segments as rows.

        Returns empty dict (not None) if segment data is unavailable, so callers
        can safely iterate over result.items() without guarding.

        Parameters
        ----------
        nf : NormalizedFinancials
            Used for fy_labels alignment.

        Returns
        -------
        dict[str, dict]
            Keys are segment names; values are dicts with:
              - "fy_labels": list[str]
              - "revenue":   list[float | None]
              - "ebit":      list[float | None]
              - "ebit_margin": list[float | None]
        """
        cd = self.company_data
        segments: dict[str, dict] = {}

        seg_data = None
        for attr in ("segments", "segment_data", "segment_results"):
            seg_data = getattr(cd, attr, None)
            if seg_data is not None:
                break

        if seg_data is None:
            logger.debug(
                "No segment data attribute found for %s", getattr(cd, "ticker", "UNKNOWN")
            )
            return segments

        fy_labels = nf.fy_labels

        # seg_data may be dict[str, DataFrame] or a single DataFrame
        if hasattr(seg_data, "items"):
            for seg_name, seg_df in seg_data.items():
                segments[str(seg_name)] = self._extract_segment_metrics(
                    seg_df, fy_labels, str(seg_name)
                )
        elif hasattr(seg_data, "index"):
            # Single DataFrame: segment names are row index
            for seg_name in seg_data.index:
                row_df = seg_data.loc[[seg_name]]
                segments[str(seg_name)] = self._extract_segment_metrics(
                    row_df, fy_labels, str(seg_name)
                )
        else:
            logger.warning(
                "Unrecognized segment data format for %s", nf.ticker
            )

        return segments

    def _extract_segment_metrics(
        self, seg_df, fy_labels: list[str], seg_name: str
    ) -> dict:
        """
        Extract revenue, EBIT, and EBIT margin for a single pharma segment.

        Tries multiple row-label variants for revenue and EBIT, since
        screener.in's segment reporting is not standardised across companies.

        Parameters
        ----------
        seg_df : DataFrame
            Segment-level DataFrame (rows = metrics, columns = FY).
        fy_labels : list[str]
            Aligned FY label list.
        seg_name : str
            Human-readable segment name (for logging only).

        Returns
        -------
        dict with keys "fy_labels", "revenue", "ebit", "ebit_margin".
        """
        result: dict = {
            "fy_labels":   fy_labels,
            "revenue":     [None] * len(fy_labels),
            "ebit":        [None] * len(fy_labels),
            "ebit_margin": [None] * len(fy_labels),
        }

        if seg_df is None or not hasattr(seg_df, "columns"):
            return result

        # Build lowercase row index
        row_index: dict[str, object] = {}
        for idx in seg_df.index:
            row_index[str(idx).strip().lower()] = idx

        rev_label_variants  = [
            "revenue", "net revenue", "segment revenue", "sales",
            "net sales", "external revenue", "segment net revenue",
        ]
        ebit_label_variants = [
            "ebit", "segment result", "segment profit", "profit",
            "operating profit", "pbit", "segment ebit",
        ]

        def _get_row(labels: list[str]) -> list[float | None]:
            """Find first matching label and return its FY series."""
            for lbl in labels:
                if lbl in row_index:
                    raw_idx = row_index[lbl]
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

        revenue = _get_row(rev_label_variants)
        ebit    = _get_row(ebit_label_variants)
        ebit_margin = [
            safe_div(e, r) * 100.0 if (e is not None and r) else None
            for e, r in zip(ebit, revenue)
        ]

        result["revenue"]     = revenue
        result["ebit"]        = ebit
        result["ebit_margin"] = ebit_margin
        return result

    # ------------------------------------------------------------------
    # Revenue mix estimation
    # ------------------------------------------------------------------

    def _estimate_revenue_mix(
        self, nf: NormalizedFinancials, segments: dict[str, dict]
    ) -> tuple[float | None, float | None]:
        """
        Estimate domestic and US revenue as % of total from segment data.

        Uses keyword matching on segment names (case-insensitive) to identify
        domestic, exports, and US buckets. Takes the latest year with data.

        If segment data is unavailable, returns (None, None) and logs a debug
        message. Callers should treat None as "not disclosed".

        Parameters
        ----------
        nf : NormalizedFinancials
            Used for latest total net revenue.
        segments : dict[str, dict]
            Output of _extract_segments().

        Returns
        -------
        (domestic_revenue_pct, us_revenue_pct)
            Each as float 0–100 or None if not determinable.
        """
        if not segments:
            logger.debug(
                "No segment data for %s; revenue mix cannot be estimated", nf.ticker
            )
            return None, None

        seg_names_lower = {k.lower(): k for k in segments}

        # Latest total net revenue
        latest_total = self._last_valid(nf.net_revenue)
        if latest_total is None or latest_total == 0:
            return None, None

        def _segment_latest_revenue(keywords: list[str]) -> float | None:
            """Sum revenue across segments matching any keyword."""
            total = 0.0
            found = False
            for kw in keywords:
                for seg_lower, seg_key in seg_names_lower.items():
                    if kw in seg_lower:
                        rev_series = segments[seg_key].get("revenue", [])
                        latest_rev = self._last_valid(rev_series)
                        if latest_rev is not None:
                            total += latest_rev
                            found = True
            return total if found else None

        domestic_rev = _segment_latest_revenue(_DOMESTIC_SEGMENT_KEYWORDS)
        us_rev       = _segment_latest_revenue(_US_SEGMENT_KEYWORDS)

        domestic_pct: float | None = (
            safe_div(domestic_rev, latest_total) * 100.0
            if domestic_rev is not None else None
        )
        us_pct: float | None = (
            safe_div(us_rev, latest_total) * 100.0
            if us_rev is not None else None
        )

        # Sanity-cap at 100%
        if domestic_pct is not None:
            domestic_pct = min(domestic_pct, 100.0)
        if us_pct is not None:
            us_pct = min(us_pct, 100.0)

        return domestic_pct, us_pct

    # ------------------------------------------------------------------
    # R&D validation and helpers
    # ------------------------------------------------------------------

    def _compute_rnd_pct(
        self,
        rnd_spend: list[float | None],
        net_revenue: list[float | None],
    ) -> list[float | None]:
        """
        Compute R&D spend as % of net revenue.

        Parameters
        ----------
        rnd_spend : list[float | None]
            Absolute R&D spend, INR Cr.
        net_revenue : list[float | None]
            Net revenue, INR Cr.

        Returns
        -------
        list[float | None]
            R&D % of revenue (0–100 scale).
        """
        return [
            safe_div(rnd, nr) * 100.0 if (rnd is not None and nr) else None
            for rnd, nr in zip(rnd_spend, net_revenue)
        ]

    def _validate_rnd_pct(
        self, rnd_pct: list[float | None], ticker: str
    ) -> None:
        """
        Log a warning if R&D % falls outside the expected 5–15% range for
        large Indian pharma. Does NOT modify data — purely diagnostic.

        Breaches could indicate:
        - R&D embedded in another expense line (below 5%)
        - Clinical-stage / specialty pipeline spending spike (above 15%)
        - Data extraction error
        """
        valid_values = [v for v in rnd_pct if v is not None]
        if not valid_values:
            logger.debug("No R&D% data to validate for %s", ticker)
            return

        latest = valid_values[-1]
        if latest < RND_PCT_LOWER_BOUND:
            logger.warning(
                "%s R&D%% = %.1f%% (latest) is below floor of %.1f%%. "
                "R&D may be captured under a different expense line.",
                ticker, latest, RND_PCT_LOWER_BOUND,
            )
        elif latest > RND_PCT_UPPER_BOUND:
            logger.warning(
                "%s R&D%% = %.1f%% (latest) exceeds ceiling of %.1f%%. "
                "Verify against annual report footnote.",
                ticker, latest, RND_PCT_UPPER_BOUND,
            )
        elif latest < RND_PCT_TYPICAL_LOWER:
            logger.debug(
                "%s R&D%% = %.1f%% is below typical 5–15%% range for large pharma.",
                ticker, latest,
            )

    # ------------------------------------------------------------------
    # Valuation helpers
    # ------------------------------------------------------------------

    def _get_pharma_beta(self, ticker: str) -> float:
        """
        Return beta for the given ticker.

        Looks up PHARMA_SECTOR_BETAS first; falls back to DEFAULT_PHARMA_BETA.
        """
        return PHARMA_SECTOR_BETAS.get(ticker.upper(), DEFAULT_PHARMA_BETA)

    def _select_pharma_peers(self, ticker: str) -> list[str]:
        """
        Select an appropriate peer group for pharma comparables analysis.

        Heuristic:
        - API / CRAMS focused companies → API_CRAMS peers
        - US generics heavy companies   → US_GENERICS peers
        - Domestic branded leaders      → BRANDED_DOMESTIC peers

        Returns peers excluding the company itself.
        """
        ticker_up = ticker.upper()
        api_cos    = {"DIVISLAB", "AUROPHARMA", "NATCO"}
        us_heavy   = {"DRREDDY", "LUPIN", "SUNPHARMA"}

        if ticker_up in api_cos:
            pool = PEERS_API_CRAMS
        elif ticker_up in us_heavy:
            pool = PEERS_US_GENERICS
        else:
            pool = PEERS_BRANDED_DOMESTIC

        return [p for p in pool if p != ticker_up]

    def _estimate_pe_target(self, ticker: str) -> float | None:
        """
        Return a justified trailing P/E target for the given pharma company.

        Multiples are based on Indian pharma sector median P/E ranges as of
        FY25, calibrated for earnings quality and business mix:
        - Brand-led domestic pharma: 30–40x
        - US generics heavy: 20–30x
        - API / CRAMS: 25–35x

        Returns the mid-point of the relevant range, or None for unlisted/unknown.
        """
        pe_map: dict[str, float] = {
            "SUNPHARMA":  38.0,
            "DIVISLAB":   55.0,   # premium for CRAMS quality
            "TORNTPHARM": 35.0,
            "ALKEM":      30.0,
            "CIPLA":      28.0,
            "DRREDDY":    25.0,
            "LUPIN":      22.0,
            "AUROPHARMA": 18.0,
            "IPCALAB":    28.0,
            "NATCO":      20.0,
        }
        return pe_map.get(ticker.upper())

    def _estimate_ev_ebitda_target(self, ticker: str) -> float | None:
        """
        Return a justified EV/EBITDA target multiple for secondary valuation check.

        Indian large-cap pharma typically trades at 20–35x EV/EBITDA.
        Higher-quality companies (Divis, Sun) command premium multiples.
        """
        ev_map: dict[str, float] = {
            "SUNPHARMA":  28.0,
            "DIVISLAB":   40.0,
            "TORNTPHARM": 25.0,
            "ALKEM":      22.0,
            "CIPLA":      20.0,
            "DRREDDY":    18.0,
            "LUPIN":      16.0,
            "AUROPHARMA": 14.0,
            "IPCALAB":    20.0,
            "NATCO":      16.0,
        }
        return ev_map.get(ticker.upper())

    # ------------------------------------------------------------------
    # Generic private utilities
    # ------------------------------------------------------------------

    def _last_valid(self, values: list[float | None]) -> float | None:
        """
        Return the last non-None value in a list, or None if all are None.

        Parameters
        ----------
        values : list[float | None]
            Time series (index 0 = oldest).

        Returns
        -------
        float | None
        """
        if not values:
            return None
        for v in reversed(values):
            if v is not None:
                return v
        return None

    def _generate_proj_fy_labels(self, last_hist_fy: str, proj_years: int) -> list[str]:
        """
        Generate projection FY labels starting immediately after last historical FY.

        Example: last_hist_fy="FY25", proj_years=5 → ["FY26", "FY27", ..., "FY30"]

        Falls back to generic "Proj+N" labels if the FY string is malformed.

        Parameters
        ----------
        last_hist_fy : str
            Last actual FY label (e.g. "FY25").
        proj_years : int
            Number of projection years.

        Returns
        -------
        list[str]
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
        # Fallback
        return [f"Proj+{i}" for i in range(1, proj_years + 1)]
