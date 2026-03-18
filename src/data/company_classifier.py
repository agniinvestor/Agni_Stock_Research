"""
Company classifier — auto-detects sector, model type, peer group, and valuation method
for any BSE/NSE-listed Indian company.

Three-tier waterfall:
  Tier 1: stock_universe.csv exact match (fastest, O(1))
  Tier 2: screener.in metadata sector/industry field parsing
  Tier 3: company name rules-based fallback

Returns CompanyClassification dataclass.
"""

import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', '..'))
import csv
import logging
from dataclasses import dataclass, field
from typing import TYPE_CHECKING
if TYPE_CHECKING:
    from src.data.screener_client import CompanyData
from config.settings import STOCK_UNIVERSE_CSV

logger = logging.getLogger(__name__)


@dataclass
class CompanyClassification:
    ticker: str
    company_name: str
    sector: str              # "FMCG", "BANKING", "IT", "PHARMA", "GENERIC"
    sub_sector: str          # e.g. "Cigarettes+FMCG", "Private Bank", "IT Services"
    model_type: str          # "fmcg" | "bank" | "it" | "generic"
    peer_group: list[str]    # peer tickers for comps
    valuation_method: str    # "ev_ebitda_pe_dcf" | "pb_abv_roe" | "pe_ev_ebit_dcf" | "pe_ev_ebitda"
    fy_end_month: int        # 3 for most Indian companies, 12 for Nestle
    is_nbfc: bool            # True for Bajaj Finance, HDFC AMC, Muthoot, etc.
    source: str              # "universe_csv" | "screener_metadata" | "rules_based"
    bse_code: str = ""       # BSE scrip code for filing downloads


class CompanyClassifier:
    """
    Classifies any BSE/NSE-listed Indian company into a sector, model type,
    peer group, and valuation method using a three-tier waterfall:

      Tier 1 — stock_universe.csv exact match  (O(1) dict lookup)
      Tier 2 — screener.in metadata parsing
      Tier 3 — company name keyword rules      (always returns a result)
    """

    # ------------------------------------------------------------------
    # Tier 2: screener sector string → model_type
    # ------------------------------------------------------------------
    SCREENER_SECTOR_MAP = {
        "banks": "bank",
        "bank": "bank",
        "banking": "bank",
        "finance": "bank",       # NBFCs
        "nbfc": "bank",
        "software": "it",
        "it": "it",
        "information technology": "it",
        "it services": "it",
        "computers": "it",
        "fmcg": "fmcg",
        "consumer goods": "fmcg",
        "consumer staples": "fmcg",
        "tobacco": "fmcg",
        "food": "fmcg",
        "beverages": "fmcg",
        "personal care": "fmcg",
        "household products": "fmcg",
        "pharmaceuticals": "pharma",
        "pharma": "pharma",
        "drugs": "pharma",
    }

    # Tier 2: screener sector → default peer group
    SECTOR_PEERS = {
        "bank":   ["HDFCBANK", "ICICIBANK", "KOTAKBANK", "AXISBANK", "SBIN"],
        "it":     ["TCS", "INFY", "WIPRO", "HCLTECH", "TECHM"],
        "fmcg":   ["ITC", "HINDUNILVR", "NESTLEIND", "BRITANNIA", "DABUR"],
        "pharma": ["SUNPHARMA", "DRREDDY", "CIPLA", "DIVISLAB", "AUROPHARMA"],
    }

    # Tier 2: sector → preferred valuation method
    SECTOR_VALUATION = {
        "bank":   "pb_abv_roe",
        "it":     "pe_ev_ebit_dcf",
        "fmcg":   "ev_ebitda_pe_dcf",
        "pharma": "pe_ev_ebitda",
    }

    # Tier 2: sector → canonical sector label
    SECTOR_LABEL = {
        "bank":    "BANKING",
        "it":      "IT",
        "fmcg":    "FMCG",
        "pharma":  "PHARMA",
        "generic": "GENERIC",
    }

    # Tier 2: sector → sub_sector default
    SECTOR_SUB = {
        "bank":    "Private Bank",
        "it":      "IT Services",
        "fmcg":    "Consumer Staples",
        "pharma":  "Pharmaceuticals",
        "generic": "General",
    }

    # Tier 3: keyword lists for rules-based fallback
    NBFC_KEYWORDS   = ["finance", "financial", "leasing", "capital", "credit", "lending"]
    BANK_KEYWORDS   = ["bank"]
    IT_KEYWORDS     = [
        "technologies", "infotech", "infosys", "wipro", "software",
        "digital", "solutions", "systems", "tech mahindra",
    ]
    FMCG_KEYWORDS   = ["consumer", "foods", "beverages", "tobacco", "cigarettes", "personal care"]
    PHARMA_KEYWORDS = ["pharma", "pharmaceutical", "drug", "medicine", "biotech", "bio-tech"]

    def __init__(self, universe_csv_path: str = None, db_path: str = None):
        """
        Parameters
        ----------
        universe_csv_path : str, optional
            Path to stock_universe.csv. Defaults to config.settings.STOCK_UNIVERSE_CSV.
        db_path : str, optional
            Reserved for future SQLite look-ups. Unused in current implementation.
        """
        self._universe: dict[str, dict] = {}
        csv_path = universe_csv_path or STOCK_UNIVERSE_CSV
        self._load_universe(csv_path)

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def classify(self, ticker: str, company_data: "CompanyData") -> CompanyClassification:
        """
        Classify a company using the three-tier waterfall.

        Parameters
        ----------
        ticker : str
            BSE/NSE ticker (e.g. "ITC", "HDFCBANK").
        company_data : CompanyData
            Populated CompanyData object from screener_client.

        Returns
        -------
        CompanyClassification
        """
        result = self._tier1_universe_lookup(ticker)
        if result is not None:
            logger.debug("[%s] Classified via Tier 1 (universe CSV).", ticker)
            return result

        result = self._tier2_screener_metadata(ticker, company_data)
        if result is not None:
            logger.debug("[%s] Classified via Tier 2 (screener metadata).", ticker)
            return result

        logger.debug("[%s] Falling back to Tier 3 (name rules).", ticker)
        return self._tier3_name_rules(ticker, company_data)

    # ------------------------------------------------------------------
    # Tier 1 — stock_universe.csv exact match
    # ------------------------------------------------------------------

    def _tier1_universe_lookup(self, ticker: str) -> "CompanyClassification | None":
        """
        Look up ticker in the pre-loaded universe dict.

        Expected CSV columns:
            ticker, bse_code, exchange, sector, sector_model, screener_ticker,
            yf_ticker, company_name, notes

        Returns None if the ticker is not present.
        """
        row = self._universe.get(ticker.upper())
        if row is None:
            return None

        model_type = row.get("sector_model", "").strip().lower() or "generic"
        sector_raw = row.get("sector", "").strip()
        company_name = row.get("company_name", ticker).strip()
        bse_code = row.get("bse_code", "").strip()

        classification = self._build_classification(
            ticker=ticker,
            company_name=company_name,
            model_type=model_type,
            sector=sector_raw,
            bse_code=bse_code,
            source="universe_csv",
        )
        return classification

    # ------------------------------------------------------------------
    # Tier 2 — screener.in metadata
    # ------------------------------------------------------------------

    def _tier2_screener_metadata(
        self, ticker: str, company_data: "CompanyData"
    ) -> "CompanyClassification | None":
        """
        Parse company_data.raw_metadata for "Sector" and "Industry" keys.

        raw_metadata is a plain dict from the screener.in page scrape.
        Sector strings are normalised (lowercased, stripped) before
        lookup in SCREENER_SECTOR_MAP.

        Returns None if the sector string is not in SCREENER_SECTOR_MAP.
        """
        raw_meta = getattr(company_data, "raw_metadata", {}) or {}

        sector_str = (
            raw_meta.get("Sector")
            or raw_meta.get("sector")
            or raw_meta.get("Industry")
            or raw_meta.get("industry")
            or ""
        )
        sector_str = sector_str.strip().lower()

        model_type = None
        # Try exact match first, then substring scan
        if sector_str in self.SCREENER_SECTOR_MAP:
            model_type = self.SCREENER_SECTOR_MAP[sector_str]
        else:
            for key, mt in self.SCREENER_SECTOR_MAP.items():
                if key in sector_str:
                    model_type = mt
                    break

        if model_type is None:
            return None

        company_name = (
            getattr(company_data, "company_name", None)
            or raw_meta.get("Company Name")
            or raw_meta.get("name")
            or ticker
        )
        bse_code = raw_meta.get("BSE Code", "") or raw_meta.get("bse_code", "")

        return self._build_classification(
            ticker=ticker,
            company_name=str(company_name),
            model_type=model_type,
            sector=sector_str,
            bse_code=str(bse_code),
            source="screener_metadata",
        )

    # ------------------------------------------------------------------
    # Tier 3 — name-based rules fallback
    # ------------------------------------------------------------------

    def _tier3_name_rules(
        self, ticker: str, company_data: "CompanyData"
    ) -> CompanyClassification:
        """
        Always returns a result. Parses company_name for keyword signals.

        Priority order: bank → NBFC → IT → FMCG → pharma → generic
        """
        raw_meta = getattr(company_data, "raw_metadata", {}) or {}
        company_name = (
            getattr(company_data, "company_name", None)
            or raw_meta.get("Company Name")
            or raw_meta.get("name")
            or ticker
        )
        name_lower = str(company_name).lower()

        # Determine model_type from name keywords
        model_type = "generic"

        if any(kw in name_lower for kw in self.BANK_KEYWORDS):
            model_type = "bank"
        elif any(kw in name_lower for kw in self.NBFC_KEYWORDS):
            # NBFCs look like banks from a modelling perspective
            model_type = "bank"
        elif any(kw in name_lower for kw in self.IT_KEYWORDS):
            model_type = "it"
        elif any(kw in name_lower for kw in self.FMCG_KEYWORDS):
            model_type = "fmcg"
        elif any(kw in name_lower for kw in self.PHARMA_KEYWORDS):
            model_type = "pharma"

        return self._build_classification(
            ticker=ticker,
            company_name=str(company_name),
            model_type=model_type,
            sector=self.SECTOR_LABEL.get(model_type, "GENERIC"),
            bse_code="",
            source="rules_based",
        )

    # ------------------------------------------------------------------
    # Classification builder helper
    # ------------------------------------------------------------------

    def _build_classification(
        self,
        ticker: str,
        company_name: str,
        model_type: str,
        sector: str,
        bse_code: str = "",
        source: str = "rules_based",
    ) -> CompanyClassification:
        """
        Construct a CompanyClassification from a resolved model_type.

        Fills in:
        - sector label (canonical uppercase)
        - sub_sector
        - peer_group
        - valuation_method
        - fy_end_month (12 for Nestle, 3 for everyone else)
        - is_nbfc (True if model_type == "bank" and name has NBFC keyword)
        """
        # Normalise model_type; unknown types fall back to generic
        mt = model_type.lower().strip()
        if mt not in ("bank", "it", "fmcg", "pharma"):
            mt = "generic"

        sector_label = self.SECTOR_LABEL.get(mt, "GENERIC")
        sub_sector = self.SECTOR_SUB.get(mt, "General")
        peer_group = list(self.SECTOR_PEERS.get(mt, []))
        valuation_method = self.SECTOR_VALUATION.get(mt, "pe_ev_ebitda")

        # Special sub-sector overrides
        if mt == "bank":
            name_lower = company_name.lower()
            # Differentiate PSU banks from private banks
            if "state bank" in name_lower or ticker.upper() in ("SBIN", "BANKBARODA", "PNB", "CANARABANK"):
                sub_sector = "PSU Bank"
            elif any(kw in name_lower for kw in self.NBFC_KEYWORDS):
                sub_sector = "NBFC"
        elif mt == "fmcg":
            name_lower = company_name.lower()
            if any(kw in name_lower for kw in ("cigarette", "tobacco", "itc")):
                sub_sector = "Cigarettes+FMCG"

        # FY end month: Nestle India follows calendar year (Dec end)
        fy_end_month = 12 if "nestle" in company_name.lower() else 3

        # NBFC flag: financial entities that are not full banks
        is_nbfc = (
            mt == "bank"
            and any(kw in company_name.lower() for kw in self.NBFC_KEYWORDS)
        )

        # Remove self from peer group if ticker is in it
        peer_group = [p for p in peer_group if p.upper() != ticker.upper()]

        return CompanyClassification(
            ticker=ticker.upper(),
            company_name=company_name,
            sector=sector_label,
            sub_sector=sub_sector,
            model_type=mt,
            peer_group=peer_group,
            valuation_method=valuation_method,
            fy_end_month=fy_end_month,
            is_nbfc=is_nbfc,
            source=source,
            bse_code=bse_code,
        )

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------

    def _load_universe(self, csv_path: str) -> None:
        """
        Load stock_universe.csv into self._universe as a dict keyed by
        uppercase ticker for O(1) lookup.

        Expected CSV header (flexible — missing columns are tolerated):
            ticker, bse_code, exchange, sector, sector_model,
            screener_ticker, yf_ticker, company_name, notes
        """
        if not csv_path or not os.path.isfile(csv_path):
            logger.warning(
                "stock_universe.csv not found at '%s'. Tier 1 lookup disabled.",
                csv_path,
            )
            return

        try:
            with open(csv_path, newline="", encoding="utf-8") as fh:
                reader = csv.DictReader(fh)
                for row in reader:
                    ticker_key = row.get("ticker", "").strip().upper()
                    if ticker_key:
                        self._universe[ticker_key] = row
            logger.info(
                "Loaded %d tickers from universe CSV: %s",
                len(self._universe), csv_path,
            )
        except Exception as exc:
            logger.error("Failed to load universe CSV '%s': %s", csv_path, exc)
