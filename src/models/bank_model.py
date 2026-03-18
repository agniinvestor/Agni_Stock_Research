"""
Banking / NBFC sector financial model.
Extends BaseModel for banks and financial companies.

Key differences from FMCGModel:
- No EBITDA concept — uses NII, PPOP, Provisions
- Asset quality metrics: GNPA, NNPA, PCR, Slippage, Credit Cost
- Capital adequacy: CAR, Tier 1 ratio
- NIM (Net Interest Margin) as primary profitability metric
- Valuation via P/B, P/ABV, Gordon Growth (RoE-g/CoE-g), not DCF/EV/EBITDA
- Screener.in uses different row labels for bank P&L

Covers: HDFCBANK, ICICIBANK, KOTAKBANK, AXISBANK, SBIN, INDUSINDBK, BAJFINANCE
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
from src.models.fmcg_model import compute_wacc  # reuse WACC formula

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Dataclasses
# ---------------------------------------------------------------------------

@dataclass
class BankMetrics:
    fy_labels: list[str]

    # P&L restatement
    interest_earned: list[float | None]     # INR Cr
    interest_expended: list[float | None]
    nii: list[float | None]                 # Net Interest Income
    other_income: list[float | None]
    total_income: list[float | None]        # NII + Other Income
    opex: list[float | None]                # Operating Expenses
    ppop: list[float | None]                # Pre-Provision Operating Profit
    provisions: list[float | None]
    pbt: list[float | None]
    pat: list[float | None]

    # Asset quality (from balance sheet / notes — may be None if not in screener)
    gnpa_abs: list[float | None]            # INR Cr
    gnpa_pct: list[float | None]            # % of gross advances
    nnpa_abs: list[float | None]
    nnpa_pct: list[float | None]
    pcr: list[float | None]                 # Provision Coverage Ratio %
    credit_cost: list[float | None]         # Provisions / Avg Advances %

    # Margins and efficiency
    nim: list[float | None]                 # %
    cost_to_income: list[float | None]      # %
    roe: list[float | None]                 # %
    roa: list[float | None]                 # %

    # Capital
    car: list[float | None]                 # Capital Adequacy Ratio %
    tier1_ratio: list[float | None]         # %

    # Business volumes
    advances: list[float | None]            # Loan book, INR Cr
    deposits: list[float | None]
    casa_ratio: list[float | None]          # %
    credit_deposit_ratio: list[float | None]  # %

    # Per share
    book_value_per_share: list[float | None]
    eps: list[float | None]

    # Growth
    advances_growth: list[float | None]     # YoY %
    deposits_growth: list[float | None]
    nii_growth: list[float | None]

    peers: list[str]


@dataclass
class BankValuationInputs:
    ticker: str
    fy_labels_hist: list[str]
    fy_labels_proj: list[str]

    pat_hist: list[float]
    book_value_per_share_hist: list[float]
    roe_hist: list[float | None]
    advances_hist: list[float]

    pat_proj: list[float | None]
    book_value_proj: list[float | None]
    roe_proj: list[float | None]
    advances_proj: list[float | None]

    # Gordon Growth Model inputs
    cost_of_equity: float           # Ke = Rf + beta * ERP  (as %)
    sustainable_roe: float          # 5yr average RoE (as %)
    terminal_growth: float          # Long-term growth ≈ 7% for India

    # P/ABV inputs
    nnpa_abs_latest: float          # INR Cr
    shares_outstanding_cr: float    # Crores

    peer_tickers: list[str]
    is_nbfc: bool = False


# ---------------------------------------------------------------------------
# Bank-specific label maps
# ---------------------------------------------------------------------------

BANK_PNL_LABELS = {
    "interest earned": "interest_earned",
    "interest income": "interest_earned",
    "interest expended": "interest_expended",
    "interest expense": "interest_expended",
    "net interest income": "nii",
    "other income": "other_income",
    "operating expenses": "opex",
    "staff expenses": None,         # part of opex, skip separately
    "other operating expenses": None,
    "provisions and contingencies": "provisions",
    "provisions": "provisions",
    "profit before tax": "pbt",
    "tax": "tax",
    "net profit": "pat",
    "profit after tax": "pat",
}

BANK_BS_LABELS = {
    "advances": "advances",
    "loans and advances": "advances",
    "deposits": "deposits",
    "share capital": "share_capital",
    "reserves": "reserves",
    "borrowings": "borrowings",
    "investments": "investments",
    "cash and balances": "cash",
    "fixed assets": "fixed_assets",
}


# ---------------------------------------------------------------------------
# BankModel
# ---------------------------------------------------------------------------

class BankModel(BaseModel):
    """
    Banking / NBFC sector financial model.
    Extends BaseModel with NII, PPOP, asset quality, NIM, and
    Gordon Growth / P/ABV valuation scaffolding.
    """

    SECTOR = "BANKING"
    PEERS_PRIVATE = ["HDFCBANK", "ICICIBANK", "KOTAKBANK", "AXISBANK", "INDUSINDBK"]
    PEERS_PSU = ["SBIN", "BANKBARODA", "CANBK", "PNB"]
    PEERS_NBFC = ["BAJFINANCE", "HDFC", "LTFH", "CHOLAFIN"]

    DEFAULT_BETAS = {
        "HDFCBANK": 0.85, "ICICIBANK": 0.90, "KOTAKBANK": 0.80,
        "AXISBANK": 0.95, "SBIN": 1.10, "DEFAULT": 0.90,
    }

    def __init__(
        self,
        company_data: "CompanyData",
        market_data: dict = None,
        is_nbfc: bool = False,
    ) -> None:
        """
        Parameters
        ----------
        company_data : CompanyData
            Output of ScreenerClient; must have profit_loss, balance_sheet
            DataFrames and ticker attribute.
        market_data : dict, optional
            May contain 'beta', 'market_cap_cr', 'shares_outstanding', 'price'.
        is_nbfc : bool
            Set True for NBFCs (BAJFINANCE, HDFC, etc.) to adjust cost of equity.
        """
        super().__init__(company_data)
        self.market_data: dict = market_data or {}
        self.is_nbfc: bool = is_nbfc
        self._bank_metrics: BankMetrics | None = None

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def compute_bank_metrics(self) -> BankMetrics:
        """
        Full pipeline:
        1. Call base normalize() to get NormalizedFinancials
        2. Override with bank-specific label extraction
        3. Compute NII = interest_earned - interest_expended
        4. Compute PPOP = total_income - opex
        5. Compute NIM: NII / average advances (approximation)
        6. Compute cost-to-income: opex / total_income
        7. Compute ROE, ROA from base model compute_ratios()
        8. Extract asset quality from quarterly_results if available
        9. Extract GNPA%, NNPA% from balance_sheet ratios section
        """
        nf = self.normalize()
        ratios = self.compute_ratios(nf)

        # Bank-specific P&L extraction
        pnl_data = self._extract_bank_pnl(nf)
        bs_data = self._extract_bank_bs()
        aq_data = self._extract_asset_quality()

        interest_earned = pnl_data.get("interest_earned", nf.fy_labels and [None] * len(nf.fy_labels))
        interest_expended = pnl_data.get("interest_expended", [None] * len(nf.fy_labels))
        other_income = pnl_data.get("other_income", nf.other_income)
        opex = pnl_data.get("opex", [None] * len(nf.fy_labels))
        provisions = pnl_data.get("provisions", [None] * len(nf.fy_labels))
        pbt = pnl_data.get("pbt", nf.pbt)
        pat = pnl_data.get("pat", nf.pat_after_mi)
        eps = nf.eps

        # NII = interest_earned - interest_expended
        nii = [
            safe_div(ie - ix, 1) if (ie is not None and ix is not None) else None
            for ie, ix in zip(interest_earned, interest_expended)
        ]
        # Override with directly extracted NII if available
        nii_direct = pnl_data.get("nii", [None] * len(nf.fy_labels))
        nii = [d if d is not None else c for d, c in zip(nii_direct, nii)]

        # total_income = NII + Other Income
        total_income = [
            (n or 0.0) + (oi or 0.0) if (n is not None or oi is not None) else None
            for n, oi in zip(nii, other_income)
        ]

        # PPOP = total_income - opex
        ppop = [
            (ti - op) if (ti is not None and op is not None) else None
            for ti, op in zip(total_income, opex)
        ]

        # Business volumes
        advances = bs_data.get("advances", [None] * len(nf.fy_labels))
        deposits = bs_data.get("deposits", [None] * len(nf.fy_labels))

        # NIM
        nim = self._compute_nim(nii, advances)

        # Cost-to-income
        cost_to_income = [
            safe_div(op, ti) * 100.0 if (op is not None and ti) else None
            for op, ti in zip(opex, total_income)
        ]

        # CASA ratio
        casa_ratio = [None] * len(nf.fy_labels)

        # Credit-deposit ratio
        credit_deposit_ratio = [
            safe_div(adv, dep) * 100.0 if (adv is not None and dep) else None
            for adv, dep in zip(advances, deposits)
        ]

        # Share capital and reserves for equity
        share_capital = bs_data.get("share_capital", [None] * len(nf.fy_labels))
        reserves = bs_data.get("reserves", [None] * len(nf.fy_labels))
        equity = [
            (sc or 0.0) + (res or 0.0) if (sc is not None or res is not None) else v
            for sc, res, v in zip(share_capital, reserves, nf.total_equity)
        ]

        # Shares outstanding (Crores) for BVPS
        shares_cr = float(self.market_data.get("shares_outstanding", 0.0))
        if shares_cr == 0.0:
            mktcap = float(self.market_data.get("market_cap_cr", 0.0))
            price = float(self.market_data.get("price", 0.0))
            if mktcap > 0 and price > 0:
                shares_cr = mktcap / price

        book_value_per_share = self._compute_book_value_per_share(equity, shares_cr)

        # Asset quality
        gnpa_abs = aq_data.get("gnpa_abs", [None] * len(nf.fy_labels))
        gnpa_pct = aq_data.get("gnpa_pct", [None] * len(nf.fy_labels))
        nnpa_abs = aq_data.get("nnpa_abs", [None] * len(nf.fy_labels))
        nnpa_pct = aq_data.get("nnpa_pct", [None] * len(nf.fy_labels))
        pcr = aq_data.get("pcr", [None] * len(nf.fy_labels))
        credit_cost = self._compute_credit_cost(provisions, advances)

        # CAR / Tier 1 from ratios DataFrame
        car = aq_data.get("car", [None] * len(nf.fy_labels))
        tier1_ratio = aq_data.get("tier1_ratio", [None] * len(nf.fy_labels))

        # Growth
        from src.models.base_model import _yoy_growth
        advances_growth = _yoy_growth(advances)
        deposits_growth = _yoy_growth(deposits)
        nii_growth = _yoy_growth(nii)

        # Peer selection
        ticker = nf.ticker
        peers = self._select_peers(ticker)

        metrics = BankMetrics(
            fy_labels=nf.fy_labels,
            interest_earned=interest_earned,
            interest_expended=interest_expended,
            nii=nii,
            other_income=other_income,
            total_income=total_income,
            opex=opex,
            ppop=ppop,
            provisions=provisions,
            pbt=pbt,
            pat=pat,
            gnpa_abs=gnpa_abs,
            gnpa_pct=gnpa_pct,
            nnpa_abs=nnpa_abs,
            nnpa_pct=nnpa_pct,
            pcr=pcr,
            credit_cost=credit_cost,
            nim=nim,
            cost_to_income=cost_to_income,
            roe=ratios.roe,
            roa=ratios.roa,
            car=car,
            tier1_ratio=tier1_ratio,
            advances=advances,
            deposits=deposits,
            casa_ratio=casa_ratio,
            credit_deposit_ratio=credit_deposit_ratio,
            book_value_per_share=book_value_per_share,
            eps=eps,
            advances_growth=advances_growth,
            deposits_growth=deposits_growth,
            nii_growth=nii_growth,
            peers=peers,
        )
        self._bank_metrics = metrics
        return metrics

    def prepare_valuation_inputs(self, proj_years: int = 5) -> BankValuationInputs:
        """
        Assemble valuation inputs.
        cost_of_equity: Rf=7.0% + beta * ERP=5.5%
        sustainable_roe: 5yr average
        terminal_growth: 7.0% (India nominal GDP growth)
        For NBFC: use higher ERP component (is_nbfc → add 1% to Ke)
        """
        nf = self._nf or self.normalize()
        metrics = self._bank_metrics or self.compute_bank_metrics()

        ticker = nf.ticker

        def _clean(lst: list[float | None]) -> list[float]:
            return [v if v is not None else 0.0 for v in lst]

        pat_hist = _clean(metrics.pat)
        bvps_hist = _clean(metrics.book_value_per_share)
        advances_hist = _clean(metrics.advances)

        # Latest NNPA absolute
        nnpa_abs_latest = 0.0
        for v in reversed(metrics.nnpa_abs):
            if v is not None:
                nnpa_abs_latest = v
                break

        # Shares outstanding (Crores)
        shares_cr = float(self.market_data.get("shares_outstanding", 0.0))
        if shares_cr == 0.0:
            mktcap = float(self.market_data.get("market_cap_cr", 0.0))
            price = float(self.market_data.get("price", 0.0))
            if mktcap > 0 and price > 0:
                shares_cr = mktcap / price

        # Beta and cost of equity
        beta = float(self.market_data.get("beta", self._get_beta(ticker)))
        erp = 5.5 + (1.0 if self.is_nbfc else 0.0)
        cost_of_equity = 7.0 + beta * erp

        # Sustainable ROE: 5yr average of non-None values
        roe_vals = [v for v in (metrics.roe or []) if v is not None]
        sustainable_roe = (sum(roe_vals[-5:]) / len(roe_vals[-5:])) if roe_vals else 15.0

        # Projection FY labels
        last_fy = nf.fy_labels[-1] if nf.fy_labels else "FY25"
        proj_fy_labels = self._generate_proj_fy_labels(last_fy, proj_years)

        # Peers
        peers = self._select_peers(ticker)

        return BankValuationInputs(
            ticker=ticker,
            fy_labels_hist=nf.fy_labels,
            fy_labels_proj=proj_fy_labels,
            pat_hist=pat_hist,
            book_value_per_share_hist=bvps_hist,
            roe_hist=metrics.roe,
            advances_hist=advances_hist,
            pat_proj=[None] * proj_years,
            book_value_proj=[None] * proj_years,
            roe_proj=[None] * proj_years,
            advances_proj=[None] * proj_years,
            cost_of_equity=round(cost_of_equity, 4),
            sustainable_roe=round(sustainable_roe, 2),
            terminal_growth=7.0,
            nnpa_abs_latest=nnpa_abs_latest,
            shares_outstanding_cr=shares_cr,
            peer_tickers=peers,
            is_nbfc=self.is_nbfc,
        )

    def estimate_projections(
        self,
        advances_growth_rates: list[float],
        nim_targets: list[float],
        cost_to_income_targets: list[float],
        credit_cost_targets: list[float],
        tax_rate: float = 0.25,
    ) -> BankValuationInputs:
        """
        Bank projection engine:
        advances[t] = advances[t-1] * (1 + growth)
        nii[t] = advances[t] * nim_target / 100
        total_income[t] = nii[t] * 1.25  (other income = ~25% of NII, historical avg)
        ppop[t] = total_income[t] * (1 - cost_to_income_target)
        provisions[t] = advances[t] * credit_cost_target / 100
        pat[t] = (ppop[t] - provisions[t]) * (1 - tax_rate)
        bvps[t] = bvps[t-1] + eps[t] - dps[t]  (assume 20% payout)

        Parameters
        ----------
        advances_growth_rates : list[float]
            Fractional YoY growth rate for loan book, e.g. [0.18, 0.16, ...].
        nim_targets : list[float]
            NIM as percentage per year, e.g. [3.8, 3.7, ...].
        cost_to_income_targets : list[float]
            Cost-to-income ratio as fraction per year, e.g. [0.42, 0.41, ...].
        credit_cost_targets : list[float]
            Credit cost (provisions/advances) as percentage per year, e.g. [0.5, 0.45, ...].
        tax_rate : float
            Effective tax rate (fraction). Default 0.25.
        """
        proj_years = len(advances_growth_rates)
        if not (len(nim_targets) == len(cost_to_income_targets)
                == len(credit_cost_targets) == proj_years):
            raise ValueError(
                "All projection input lists must have the same length."
            )

        val_inputs = self.prepare_valuation_inputs(proj_years=proj_years)
        metrics = self._bank_metrics or self.compute_bank_metrics()

        # Base values from last historical year
        advances_base = val_inputs.advances_hist[-1] if val_inputs.advances_hist else 0.0
        bvps_base = val_inputs.book_value_per_share_hist[-1] if val_inputs.book_value_per_share_hist else 0.0
        shares_cr = val_inputs.shares_outstanding_cr

        pat_proj: list[float | None] = []
        bvps_proj: list[float | None] = []
        roe_proj: list[float | None] = []
        advances_proj: list[float | None] = []

        current_advances = advances_base
        current_bvps = bvps_base

        for g, nim_t, cti_t, cc_t in zip(
            advances_growth_rates, nim_targets, cost_to_income_targets, credit_cost_targets
        ):
            # Loan book
            current_advances = current_advances * (1.0 + g)
            advances_proj.append(round(current_advances, 2))

            # NII
            nii = current_advances * nim_t / 100.0

            # Total income (other income ~ 25% of NII)
            total_income = nii * 1.25

            # PPOP
            ppop = total_income * (1.0 - cti_t)

            # Provisions
            provisions = current_advances * cc_t / 100.0

            # PAT
            pat = max((ppop - provisions) * (1.0 - tax_rate), 0.0)
            pat_proj.append(round(pat, 2))

            # EPS and BVPS
            if shares_cr > 0:
                eps = pat / shares_cr  # INR per share (Cr / Cr shares = Rs)
                dps = eps * 0.20       # 20% payout assumption
                current_bvps = current_bvps + eps - dps
            else:
                current_bvps = None

            bvps_proj.append(round(current_bvps, 2) if current_bvps is not None else None)

            # ROE = PAT / ((BVPS_start + BVPS_end) / 2 * shares_cr)
            # Simplified: ROE = EPS / avg_bvps * 100
            if current_bvps is not None and shares_cr > 0:
                avg_bvps = (bvps_proj[-2] if len(bvps_proj) >= 2 else bvps_base) or bvps_base
                if avg_bvps and current_bvps:
                    roe_v = safe_div(pat, ((avg_bvps + current_bvps) / 2.0) * shares_cr)
                    roe_proj.append(round(roe_v * 100.0, 2) if roe_v is not None else None)
                else:
                    roe_proj.append(None)
            else:
                roe_proj.append(None)

        val_inputs.pat_proj = pat_proj
        val_inputs.book_value_proj = bvps_proj
        val_inputs.roe_proj = roe_proj
        val_inputs.advances_proj = advances_proj

        return val_inputs

    # ------------------------------------------------------------------
    # Private helpers
    # ------------------------------------------------------------------

    def _extract_bank_pnl(self, nf: NormalizedFinancials) -> dict:
        """Extract bank-specific P&L items from income_statement DataFrame.
        Uses BANK_PNL_LABELS map. Returns dict of {field: list[float|None]}."""
        from src.models.base_model import _to_float

        result: dict[str, list[float | None]] = {}
        df = self._get_df("pnl")
        if df is None:
            return result

        n = len(nf.fy_labels)
        fy_labels = nf.fy_labels

        # Build lowercase row index
        row_index: dict[str, object] = {}
        for idx in df.index:
            row_index[str(idx).strip().lower()] = idx

        for label, field_name in BANK_PNL_LABELS.items():
            if field_name is None:
                continue
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
                    series.append(val)
                # Only store the first match found for each field_name
                if field_name not in result:
                    result[field_name] = series

        return result

    def _extract_bank_bs(self) -> dict:
        """Extract bank balance sheet items. Returns dict of {field: list}."""
        from src.models.base_model import _to_float

        result: dict[str, list[float | None]] = {}
        nf = self._nf
        if nf is None:
            return result

        df = self._get_df("balance_sheet")
        if df is None:
            return result

        fy_labels = nf.fy_labels

        # Build lowercase row index
        row_index: dict[str, object] = {}
        for idx in df.index:
            row_index[str(idx).strip().lower()] = idx

        for label, field_name in BANK_BS_LABELS.items():
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

    def _extract_asset_quality(self) -> dict:
        """
        Try to extract GNPA%, NNPA%, PCR from:
        1. ratios DataFrame (screener.in shows some asset quality ratios)
        2. quarterly_results DataFrame (recent quarters)
        Returns dict with gnpa_pct, nnpa_pct, pcr, credit_cost lists (may be partial).
        """
        from src.models.base_model import _to_float

        nf = self._nf
        if nf is None:
            return {}

        n = len(nf.fy_labels)
        result: dict[str, list[float | None]] = {
            "gnpa_abs": [None] * n,
            "gnpa_pct": [None] * n,
            "nnpa_abs": [None] * n,
            "nnpa_pct": [None] * n,
            "pcr": [None] * n,
            "car": [None] * n,
            "tier1_ratio": [None] * n,
        }

        cd = self.company_data
        # Try ratios DataFrame
        ratios_df = getattr(cd, "ratios", None)
        if ratios_df is None:
            ratios_df = getattr(cd, "key_ratios", None)

        if ratios_df is None:
            return result

        fy_labels = nf.fy_labels

        # Build lowercase row index
        row_index: dict[str, object] = {}
        for idx in ratios_df.index:
            row_index[str(idx).strip().lower()] = idx

        aq_label_map = {
            "gross npa %": "gnpa_pct",
            "gnpa %": "gnpa_pct",
            "net npa %": "nnpa_pct",
            "nnpa %": "nnpa_pct",
            "provision coverage ratio": "pcr",
            "pcr %": "pcr",
            "capital adequacy ratio": "car",
            "car %": "car",
            "tier 1 ratio": "tier1_ratio",
            "tier1 ratio": "tier1_ratio",
        }

        for label, field_name in aq_label_map.items():
            if label in row_index:
                raw_idx = row_index[label]
                row = ratios_df.loc[raw_idx]
                series: list[float | None] = []
                for fy in fy_labels:
                    val = None
                    for col_key in self._fy_to_col_keys(fy, ratios_df.columns):
                        if col_key in ratios_df.columns:
                            val = _to_float(row[col_key])
                            break
                    series.append(val)
                result[field_name] = series

        return result

    def _compute_nim(
        self, nii_list: list, advances_list: list
    ) -> list[float | None]:
        """NIM = NII / Average Advances * 100. Average = (beginning + ending) / 2."""
        from src.models.base_model import _average_adjacent

        avg_advances = _average_adjacent(advances_list)
        nim: list[float | None] = []
        for nii, avg_adv in zip(nii_list, avg_advances):
            if nii is not None and avg_adv:
                nim.append(round(safe_div(nii, avg_adv) * 100.0, 4))
            else:
                nim.append(None)
        return nim

    def _compute_credit_cost(
        self, provisions_list: list, advances_list: list
    ) -> list[float | None]:
        """Credit Cost = Provisions / Average Advances * 100."""
        from src.models.base_model import _average_adjacent

        avg_advances = _average_adjacent(advances_list)
        credit_cost: list[float | None] = []
        for prov, avg_adv in zip(provisions_list, avg_advances):
            if prov is not None and avg_adv:
                credit_cost.append(round(safe_div(prov, avg_adv) * 100.0, 4))
            else:
                credit_cost.append(None)
        return credit_cost

    def _compute_book_value_per_share(
        self, equity_list: list, shares_cr: float
    ) -> list[float | None]:
        """BVPS = Total Equity (Cr) / Shares Outstanding (Crores)."""
        if shares_cr <= 0:
            return [None] * len(equity_list)
        bvps: list[float | None] = []
        for eq in equity_list:
            if eq is not None:
                bvps.append(round(eq / shares_cr, 2))
            else:
                bvps.append(None)
        return bvps

    def _get_beta(self, ticker: str) -> float:
        """Return default beta for the given ticker."""
        return self.DEFAULT_BETAS.get(ticker.upper(), self.DEFAULT_BETAS["DEFAULT"])

    def _select_peers(self, ticker: str) -> list[str]:
        """Select appropriate peer group based on ticker."""
        ticker_upper = ticker.upper()
        if ticker_upper in self.PEERS_NBFC:
            return [p for p in self.PEERS_NBFC if p != ticker_upper]
        elif ticker_upper in self.PEERS_PSU:
            return [p for p in self.PEERS_PSU if p != ticker_upper]
        else:
            return [p for p in self.PEERS_PRIVATE if p != ticker_upper]

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
