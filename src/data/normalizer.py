"""
Data normalization layer — sits between screener_client output and financial models.

Patches CompanyData DataFrames for:
1. Ind AS transition break (FY17) — flags discontinuity, optionally NaN-fills affected items
2. Excise duty gross-up — re-derives net_revenue for pre-GST years (pre-FY18)
3. Segment name continuity — maps historical segment labels to canonical current names
4. USD item flagging — annotates when revenue is in USD

Does NOT modify CompanyData in-place. Returns a patched copy.
"""

import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', '..'))
import logging
import dataclasses
from dataclasses import dataclass, field
from typing import TYPE_CHECKING
import pandas as pd
import numpy as np
if TYPE_CHECKING:
    from src.data.screener_client import CompanyData

logger = logging.getLogger(__name__)


@dataclass
class NormalizerFlags:
    ticker: str
    ind_as_break_detected: bool = False    # True if Ind AS transition discontinuity found
    ind_as_break_fy: str = "FY17"          # FY where break was detected
    excise_duty_adjusted: bool = False     # True if excise was stripped from gross revenue
    excise_affected_fys: list[str] = field(default_factory=list)  # ["FY16", "FY17", ...]
    segment_renames: dict = field(default_factory=dict)  # {old_name: new_name}
    usd_revenue_flagged: bool = False
    warnings: list[str] = field(default_factory=list)


class DataNormalizer:
    """
    Normalizes CompanyData objects for consistent downstream consumption.

    Handles:
    - Ind AS transition discontinuities (FY17)
    - Excise duty gross-up for pre-GST years (pre-FY18)
    - Segment label canonicalization
    - USD revenue annotation
    """

    # Sectors where excise duty gross-up applies (pre-GST)
    EXCISE_SECTORS = {"CIGARETTES", "TOBACCO", "ALCOHOL", "LIQUOR", "PETROLEUM", "OIL"}

    # Segment name aliases per ticker: {ticker: {old_name: canonical_name}}
    SEGMENT_ALIASES = {
        "ITC": {
            "FMCG": "FMCG - Others",
            "FMCG-Others": "FMCG - Others",
            "Paperboards": "Paperboards, Paper & Packaging",
            "Agri": "Agri Business",
        },
        "HINDUNILVR": {
            "Home & Personal Care": "Beauty & Wellbeing",
            "Foods": "Foods & Refreshment",
        },
    }

    def __init__(self, strict: bool = False):
        """
        Parameters
        ----------
        strict : bool
            False (default): silently patch and log warnings.
            True: raise ValueError on ambiguous normalizations.
        """
        self.strict = strict

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def normalize(self, ticker: str, company_data: "CompanyData") -> tuple["CompanyData", NormalizerFlags]:
        """
        Run all normalization steps and return a patched copy of company_data.

        Parameters
        ----------
        ticker : str
            BSE/NSE ticker symbol (e.g. "ITC", "HDFCBANK").
        company_data : CompanyData
            Raw output from screener_client.

        Returns
        -------
        tuple[CompanyData, NormalizerFlags]
            (patched_company_data, flags) — original is not modified.
        """
        flags = NormalizerFlags(ticker=ticker)

        # Work on a cloned copy so we never mutate the caller's object.
        cd = self._clone_company_data(company_data)

        # Step 1: Ind AS transition break detection
        cd, ind_as_detected, ind_as_fy = self._detect_ind_as_break(ticker, cd)
        flags.ind_as_break_detected = ind_as_detected
        flags.ind_as_break_fy = ind_as_fy

        # Step 2: Excise duty gross-up (pre-GST)
        cd, excise_adjusted, excise_fys = self._strip_excise_duty(ticker, cd)
        flags.excise_duty_adjusted = excise_adjusted
        flags.excise_affected_fys = excise_fys

        # Step 3: Segment name canonicalization
        cd, rename_map = self._fix_segment_names(ticker, cd)
        flags.segment_renames = rename_map

        # Step 4: USD revenue flagging
        cd, usd_flagged = self._flag_usd_items(ticker, cd)
        flags.usd_revenue_flagged = usd_flagged

        if flags.warnings:
            for w in flags.warnings:
                logger.warning("[%s] %s", ticker, w)

        return cd, flags

    # ------------------------------------------------------------------
    # Step 1 — Ind AS transition break
    # ------------------------------------------------------------------

    def _detect_ind_as_break(
        self, ticker: str, cd: "CompanyData"
    ) -> tuple["CompanyData", bool, str]:
        """
        Heuristic: if equity changes > 25% YoY in FY17 while total_assets
        change < 15%, flag an Ind AS transition discontinuity and NaN the
        FY17 equity value.

        Detection relies on the balance_sheet DataFrame having FY labels
        as column names (e.g. "FY17", "FY16") and row labels containing
        "Reserves" for equity and "Total Assets" for assets.

        Returns
        -------
        tuple[CompanyData, bool, str]
            (patched_cd, detected, fy_label)
        """
        bs = getattr(cd, "balance_sheet", None)
        if bs is None or bs.empty:
            return cd, False, "FY17"

        # Locate the equity proxy row (reserves) and assets row
        equity_row = self._find_row(bs, ["reserves", "equity"])
        assets_row = self._find_row(bs, ["total assets", "total asset"])

        if equity_row is None or assets_row is None:
            return cd, False, "FY17"

        if "FY17" not in bs.columns or "FY16" not in bs.columns:
            return cd, False, "FY17"

        try:
            eq17 = float(bs.loc[equity_row, "FY17"])
            eq16 = float(bs.loc[equity_row, "FY16"])
            ta17 = float(bs.loc[assets_row, "FY17"])
            ta16 = float(bs.loc[assets_row, "FY16"])
        except (TypeError, ValueError, KeyError):
            return cd, False, "FY17"

        if eq16 == 0 or ta16 == 0:
            return cd, False, "FY17"

        equity_chg = abs(eq17 - eq16) / abs(eq16)
        assets_chg = abs(ta17 - ta16) / abs(ta16)

        detected = equity_chg > 0.25 and assets_chg < 0.15

        if detected:
            msg = (
                f"Ind AS transition break detected in FY17 — "
                f"equity Δ={equity_chg:.1%}, assets Δ={assets_chg:.1%}. "
                f"Setting equity FY17 to NaN."
            )
            logger.warning("[%s] %s", ticker, msg)
            if self.strict:
                raise ValueError(f"[{ticker}] {msg}")

            # NaN out the equity FY17 cell
            bs_patched = bs.copy()
            bs_patched.loc[equity_row, "FY17"] = np.nan
            cd = dataclasses.replace(cd, balance_sheet=bs_patched)

        return cd, detected, "FY17"

    # ------------------------------------------------------------------
    # Step 2 — Excise duty gross-up
    # ------------------------------------------------------------------

    def _strip_excise_duty(
        self, ticker: str, cd: "CompanyData"
    ) -> tuple["CompanyData", bool, list[str]]:
        """
        For companies in EXCISE_SECTORS, subtract excise duty from gross
        revenue to derive/update a "Net Sales" row for pre-GST years.

        The excise duty row is identified by label containing "excise".
        The gross revenue row is identified by "Sales +" or "Total Revenue".
        Post-FY18 columns where excise = 0 are skipped.

        Returns
        -------
        tuple[CompanyData, bool, list[str]]
            (patched_cd, adjusted, affected_fys)
        """
        inc = getattr(cd, "income_statement", None)
        if inc is None or inc.empty:
            return cd, False, []

        # Only apply to sectors that historically charged excise duty
        sector = self._get_sector(cd).upper()
        sector_match = any(s in sector for s in self.EXCISE_SECTORS)
        # Also attempt if an excise row is present regardless of sector
        excise_row = self._find_row(inc, ["excise"])
        if excise_row is None:
            return cd, False, []

        gross_row = self._find_row(inc, ["sales +", "total revenue", "revenue from operations"])
        if gross_row is None:
            logger.debug("[%s] No gross revenue row found; skipping excise strip.", ticker)
            return cd, False, []

        inc_patched = inc.copy()

        # Ensure a "Net Sales" row exists
        net_sales_row = self._find_row(inc_patched, ["net sales"])
        if net_sales_row is None:
            # Append a new row
            net_sales_row = "Net Sales"
            inc_patched.loc[net_sales_row] = np.nan

        affected_fys: list[str] = []

        for col in inc_patched.columns:
            try:
                excise_val = float(inc_patched.loc[excise_row, col])
                gross_val = float(inc_patched.loc[gross_row, col])
            except (TypeError, ValueError):
                continue

            if excise_val == 0 or np.isnan(excise_val):
                # Post-FY18 or already net — no adjustment needed
                continue

            net_val = gross_val - excise_val
            inc_patched.loc[net_sales_row, col] = net_val
            affected_fys.append(str(col))

        adjusted = len(affected_fys) > 0
        if adjusted:
            logger.info(
                "[%s] Excise duty stripped for %d FY(s): %s",
                ticker, len(affected_fys), affected_fys,
            )
            cd = dataclasses.replace(cd, income_statement=inc_patched)

        return cd, adjusted, affected_fys

    # ------------------------------------------------------------------
    # Step 3 — Segment name canonicalization
    # ------------------------------------------------------------------

    def _fix_segment_names(
        self, ticker: str, cd: "CompanyData"
    ) -> tuple["CompanyData", dict]:
        """
        Apply SEGMENT_ALIASES for the given ticker across all DataFrames
        that carry segment labels (index or column values).

        Returns
        -------
        tuple[CompanyData, dict]
            (patched_cd, rename_map) where rename_map = {old: new}
        """
        aliases = self.SEGMENT_ALIASES.get(ticker, {})
        if not aliases:
            return cd, {}

        rename_map: dict = {}
        cd_fields = dataclasses.fields(cd)
        patched_dfs: dict = {}

        for f in cd_fields:
            val = getattr(cd, f.name, None)
            if not isinstance(val, pd.DataFrame):
                continue

            df = val.copy()
            changed = False

            # Rename index entries
            new_index = []
            for idx in df.index:
                canonical = aliases.get(str(idx))
                if canonical is not None and canonical != str(idx):
                    rename_map[str(idx)] = canonical
                    new_index.append(canonical)
                    changed = True
                else:
                    new_index.append(idx)
            if changed:
                df.index = new_index

            # Rename column entries (e.g. segment pivot tables)
            new_cols = []
            col_changed = False
            for col in df.columns:
                canonical = aliases.get(str(col))
                if canonical is not None and canonical != str(col):
                    rename_map[str(col)] = canonical
                    new_cols.append(canonical)
                    col_changed = True
                else:
                    new_cols.append(col)
            if col_changed:
                df.columns = new_cols
                changed = True

            if changed:
                patched_dfs[f.name] = df

        if patched_dfs:
            cd = dataclasses.replace(cd, **patched_dfs)
            logger.info("[%s] Segment renames applied: %s", ticker, rename_map)

        return cd, rename_map

    # ------------------------------------------------------------------
    # Step 4 — USD item flagging
    # ------------------------------------------------------------------

    def _flag_usd_items(
        self, ticker: str, cd: "CompanyData"
    ) -> tuple["CompanyData", bool]:
        """
        Flag if revenue-related metadata suggests USD denomination.
        Checks:
          - raw_metadata dict for "USD" or "$" in revenue fields
          - company sector containing "IT" or "Software" (USD billers)

        The patch adds a top-level attribute `usd_revenue` = True to cd
        via dataclasses.replace when flagged (requires the field to exist).

        Returns
        -------
        tuple[CompanyData, bool]
            (patched_cd, flagged)
        """
        flagged = False

        # Check sector
        sector = self._get_sector(cd).upper()
        if any(kw in sector for kw in ("IT", "SOFTWARE", "INFORMATION TECHNOLOGY")):
            flagged = True

        # Check raw_metadata
        raw_meta = getattr(cd, "raw_metadata", {}) or {}
        revenue_keys = [k for k in raw_meta if "revenue" in k.lower() or "sales" in k.lower()]
        for key in revenue_keys:
            val = str(raw_meta.get(key, ""))
            if "USD" in val.upper() or "$" in val:
                flagged = True
                break

        # Also scan all string values in raw_metadata for currency hints
        if not flagged:
            for key, val in raw_meta.items():
                if isinstance(val, str) and ("USD" in val.upper() or "$ " in val):
                    flagged = True
                    break

        if flagged:
            logger.info("[%s] USD revenue flagged (sector=%s).", ticker, sector)
            # Attach flag if the dataclass supports it
            try:
                cd = dataclasses.replace(cd, usd_revenue=True)
            except TypeError:
                # Field doesn't exist on CompanyData — annotation only
                pass

        return cd, flagged

    # ------------------------------------------------------------------
    # Clone helper
    # ------------------------------------------------------------------

    def _clone_company_data(self, cd: "CompanyData") -> "CompanyData":
        """
        Return a shallow copy of cd with all DataFrame fields deep-copied
        so mutations don't bleed back to the caller's object.
        """
        replacements: dict = {}
        for f in dataclasses.fields(cd):
            val = getattr(cd, f.name, None)
            if isinstance(val, pd.DataFrame):
                replacements[f.name] = val.copy()
            elif isinstance(val, dict):
                replacements[f.name] = dict(val)
            elif isinstance(val, list):
                replacements[f.name] = list(val)
        return dataclasses.replace(cd, **replacements)

    # ------------------------------------------------------------------
    # Internal utilities
    # ------------------------------------------------------------------

    def _find_row(self, df: pd.DataFrame, keywords: list[str]) -> str | None:
        """
        Return the first index label whose lowercased string contains any
        of the given keywords. Returns None if no match found.
        """
        for idx in df.index:
            idx_lower = str(idx).lower()
            for kw in keywords:
                if kw.lower() in idx_lower:
                    return idx
        return None

    def _get_sector(self, cd: "CompanyData") -> str:
        """Extract sector string from CompanyData, falling back to empty string."""
        raw_meta = getattr(cd, "raw_metadata", {}) or {}
        for key in ("Sector", "sector", "Industry", "industry"):
            if key in raw_meta:
                return str(raw_meta[key])
        return getattr(cd, "sector", "") or ""
