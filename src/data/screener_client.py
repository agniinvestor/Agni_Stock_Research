"""
Fetches financial data from screener.in for any BSE/NSE-listed Indian company.

screener.in is an unofficial API — HTML scraping with session auth.
All data is cached in SQLite to avoid hammering the site.

Usage:
    from src.data.screener_client import ScreenerClient
    client = ScreenerClient()
    data = client.fetch_company("ITC")
    print(data.income_statement.head())
"""

import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', '..'))

import json
import logging
import re
import sqlite3
import time
from dataclasses import dataclass, field, asdict
from datetime import datetime, timedelta
from typing import Optional

import pandas as pd
import requests
from bs4 import BeautifulSoup

try:
    from config.settings import (
        SCREENER_BASE_URL, SCREENER_LOGIN_URL, SCREENER_COMPANY_URL,
        SCREENER_COMPANY_SA_URL, SCREENER_CACHE_TTL_DAYS, RATE_LIMIT_DELAY_SECS,
        MAX_RETRIES, SCREENER_USERNAME, SCREENER_PASSWORD, SQLITE_DB_PATH
    )
except ImportError:
    SCREENER_BASE_URL = "https://www.screener.in"
    SCREENER_LOGIN_URL = "https://www.screener.in/login/"
    SCREENER_COMPANY_URL = "https://www.screener.in/company/{ticker}/consolidated/"
    SCREENER_COMPANY_SA_URL = "https://www.screener.in/company/{ticker}/"
    SCREENER_CACHE_TTL_DAYS = 1
    RATE_LIMIT_DELAY_SECS = 2.0
    MAX_RETRIES = 3
    SCREENER_USERNAME = os.environ.get("SCREENER_USERNAME", "")
    SCREENER_PASSWORD = os.environ.get("SCREENER_PASSWORD", "")
    SQLITE_DB_PATH = os.path.join(os.path.expanduser("~"), ".india_stock_research", "cache.db")

logger = logging.getLogger(__name__)


@dataclass
class CompanyData:
    ticker: str
    company_name: str
    is_consolidated: bool
    fetched_at: datetime

    # DataFrames indexed by FY label ("FY25", "FY24", ...)
    # Each column is a financial line item
    income_statement: pd.DataFrame   # rows = FY labels
    balance_sheet: pd.DataFrame
    cash_flow: pd.DataFrame
    ratios: pd.DataFrame             # key ratios from screener
    shareholding: pd.DataFrame       # rows = quarter labels ("Mar-25", etc.)
    quarterly_results: pd.DataFrame  # rows = quarter labels

    # Metadata
    raw_metadata: dict               # market cap, face value, etc. from page


def _parse_number(val: str) -> float:
    """Parse an Indian-formatted number string to float. Returns NaN on failure."""
    if not isinstance(val, str):
        try:
            return float(val)
        except (TypeError, ValueError):
            return float("nan")

    val = val.strip()
    if val in ("", "–", "-", "N/A", "NA", "#N/A"):
        return float("nan")

    # Handle negative in parentheses: (1,234) -> -1234
    negative = False
    if val.startswith("(") and val.endswith(")"):
        negative = True
        val = val[1:-1]

    # Strip % and trailing/leading spaces
    val = val.replace("%", "").replace(",", "").strip()

    try:
        result = float(val)
        return -result if negative else result
    except ValueError:
        return float("nan")


def _fy_label(col_header: str) -> str:
    """Convert 'Mar 2025' to 'FY25', 'Mar 2024' to 'FY24', etc."""
    col_header = col_header.strip()
    match = re.search(r"(\d{4})", col_header)
    if match:
        year = match.group(1)
        return f"FY{year[2:]}"
    return col_header


def _quarter_label(col_header: str) -> str:
    """Convert 'Sep 2024' to 'Sep-24', 'Mar 2025' to 'Mar-25', etc."""
    col_header = col_header.strip()
    match = re.match(r"([A-Za-z]+)\s+(\d{4})", col_header)
    if match:
        mon = match.group(1)[:3].capitalize()
        yr = match.group(2)[2:]
        return f"{mon}-{yr}"
    return col_header


def _parse_section_table(section_tag, label_converter=_fy_label) -> pd.DataFrame:
    """
    Parse a screener.in section's <table> into a DataFrame.

    Rows  = financial line items (from first <td> of each <tr> in <tbody>)
    Columns = FY or quarter labels (from <thead>)
    """
    if section_tag is None:
        return pd.DataFrame()

    table = section_tag.find("table")
    if table is None:
        return pd.DataFrame()

    # Extract column headers from thead
    thead = table.find("thead")
    col_labels = []
    if thead:
        header_cells = thead.find_all("th")
        # First th is usually the row label header ("") — skip it
        for th in header_cells[1:]:
            text = th.get_text(strip=True)
            if text:
                col_labels.append(label_converter(text))

    if not col_labels:
        # Try tr in thead
        header_row = thead.find("tr") if thead else None
        if header_row:
            ths = header_row.find_all(["th", "td"])
            for cell in ths[1:]:
                text = cell.get_text(strip=True)
                if text:
                    col_labels.append(label_converter(text))

    tbody = table.find("tbody")
    if tbody is None:
        return pd.DataFrame()

    rows = {}
    for tr in tbody.find_all("tr"):
        cells = tr.find_all("td")
        if not cells:
            continue
        row_label = cells[0].get_text(strip=True)
        if not row_label:
            continue
        values = []
        for td in cells[1:]:
            values.append(_parse_number(td.get_text(strip=True)))

        # Align length with col_labels
        if col_labels:
            if len(values) < len(col_labels):
                values.extend([float("nan")] * (len(col_labels) - len(values)))
            else:
                values = values[:len(col_labels)]

        rows[row_label] = values

    if not rows:
        return pd.DataFrame()

    df = pd.DataFrame(rows, index=col_labels if col_labels else None).T
    df.index.name = "metric"
    return df


class ScreenerClient:
    """Fetches and caches financial data from screener.in."""

    def __init__(
        self,
        db_path: str = None,
        username: str = None,
        password: str = None,
    ):
        self.db_path = db_path or SQLITE_DB_PATH
        self.username = username or SCREENER_USERNAME
        self.password = password or SCREENER_PASSWORD
        self._session: Optional[requests.Session] = None
        self._logged_in = False
        self._init_db()

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def fetch_company(self, ticker: str, force_refresh: bool = False) -> CompanyData:
        """
        Main entry point. Checks cache first.
        Falls back to standalone if consolidated not available.
        """
        ticker = ticker.upper().strip()
        logger.info("Fetching company data for %s (force_refresh=%s)", ticker, force_refresh)

        if not force_refresh:
            cached = self._cache_get(ticker)
            if cached is not None:
                logger.info("Cache hit for %s", ticker)
                return cached

        # Try consolidated first, fall back to standalone
        for consolidated in (True, False):
            try:
                html = self._fetch_html(ticker, consolidated=consolidated)
                if html:
                    data = self._parse_html(html, ticker)
                    data.is_consolidated = consolidated
                    self._cache_set(ticker, data)
                    logger.info(
                        "Successfully fetched %s (%s)",
                        ticker,
                        "consolidated" if consolidated else "standalone",
                    )
                    return data
            except requests.HTTPError as exc:
                if exc.response is not None and exc.response.status_code == 404:
                    if consolidated:
                        logger.info(
                            "Consolidated page not found for %s, trying standalone", ticker
                        )
                        continue
                raise
            except Exception as exc:
                logger.error("Error fetching %s (consolidated=%s): %s", ticker, consolidated, exc)
                if not consolidated:
                    raise

        raise ValueError(f"Could not fetch data for ticker '{ticker}'")

    # ------------------------------------------------------------------
    # Session / HTTP
    # ------------------------------------------------------------------

    def _ensure_session(self) -> requests.Session:
        """Login to screener.in if not already logged in."""
        if self._session is not None and self._logged_in:
            return self._session

        session = requests.Session()
        session.headers.update(
            {
                "User-Agent": (
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/120.0.0.0 Safari/537.36"
                ),
                "Accept-Language": "en-US,en;q=0.9",
                "Referer": SCREENER_BASE_URL,
            }
        )

        if self.username and self.password:
            try:
                # GET login page to extract CSRF token
                login_get = session.get(SCREENER_LOGIN_URL, timeout=15)
                login_get.raise_for_status()
                soup = BeautifulSoup(login_get.text, "html.parser")
                csrf_input = soup.find("input", {"name": "csrfmiddlewaretoken"})
                csrf_token = csrf_input["value"] if csrf_input else ""

                payload = {
                    "username": self.username,
                    "password": self.password,
                    "csrfmiddlewaretoken": csrf_token,
                    "next": "/",
                }
                headers = {"Referer": SCREENER_LOGIN_URL}
                login_post = session.post(
                    SCREENER_LOGIN_URL,
                    data=payload,
                    headers=headers,
                    timeout=15,
                    allow_redirects=True,
                )
                login_post.raise_for_status()

                # Check if login succeeded (screener redirects to home on success)
                if "login" in login_post.url:
                    logger.warning("screener.in login may have failed — still on login page")
                else:
                    logger.info("Logged into screener.in successfully")
                    self._logged_in = True
            except Exception as exc:
                logger.warning("screener.in login failed: %s — proceeding as guest", exc)
        else:
            logger.debug("No screener.in credentials provided — fetching as guest")

        self._session = session
        return session

    def _fetch_html(self, ticker: str, consolidated: bool = True) -> str:
        """GET the company page. Handle 429 with exponential backoff."""
        session = self._ensure_session()
        if consolidated:
            url = SCREENER_COMPANY_URL.format(ticker=ticker)
        else:
            url = SCREENER_COMPANY_SA_URL.format(ticker=ticker)

        backoff = RATE_LIMIT_DELAY_SECS
        for attempt in range(1, MAX_RETRIES + 1):
            try:
                logger.debug("GET %s (attempt %d/%d)", url, attempt, MAX_RETRIES)
                time.sleep(RATE_LIMIT_DELAY_SECS)
                response = session.get(url, timeout=20)

                if response.status_code == 429:
                    logger.warning(
                        "Rate limited (429) on attempt %d — sleeping %.1fs", attempt, backoff
                    )
                    time.sleep(backoff)
                    backoff = min(backoff * 2, 60)
                    continue

                response.raise_for_status()
                return response.text

            except requests.HTTPError as exc:
                if exc.response is not None and exc.response.status_code in (404, 403):
                    raise
                if attempt == MAX_RETRIES:
                    raise
                logger.warning("HTTP error on attempt %d: %s", attempt, exc)
                time.sleep(backoff)
                backoff = min(backoff * 2, 60)

            except requests.RequestException as exc:
                if attempt == MAX_RETRIES:
                    raise
                logger.warning("Request error on attempt %d: %s", attempt, exc)
                time.sleep(backoff)
                backoff = min(backoff * 2, 60)

        raise RuntimeError(f"Failed to fetch {url} after {MAX_RETRIES} attempts")

    # ------------------------------------------------------------------
    # Parsing
    # ------------------------------------------------------------------

    def _parse_html(self, html: str, ticker: str) -> CompanyData:
        """Parse screener.in HTML into a CompanyData object."""
        soup = BeautifulSoup(html, "html.parser")

        # Company name from <h1>
        h1 = soup.find("h1")
        company_name = h1.get_text(strip=True) if h1 else ticker

        # Parse each financial section
        income_statement = _parse_section_table(
            soup.find("section", {"id": "profit-loss"}), label_converter=_fy_label
        )
        balance_sheet = _parse_section_table(
            soup.find("section", {"id": "balance-sheet"}), label_converter=_fy_label
        )
        cash_flow = _parse_section_table(
            soup.find("section", {"id": "cash-flow"}), label_converter=_fy_label
        )
        ratios = _parse_section_table(
            soup.find("section", {"id": "ratios"}), label_converter=_fy_label
        )

        quarterly_results = _parse_section_table(
            soup.find("section", {"id": "quarters"}), label_converter=_quarter_label
        )

        shareholding = self._parse_shareholding(soup)

        raw_metadata = self._parse_metadata(soup)

        return CompanyData(
            ticker=ticker,
            company_name=company_name,
            is_consolidated=True,  # caller overrides
            fetched_at=datetime.utcnow(),
            income_statement=income_statement,
            balance_sheet=balance_sheet,
            cash_flow=cash_flow,
            ratios=ratios,
            shareholding=shareholding,
            quarterly_results=quarterly_results,
            raw_metadata=raw_metadata,
        )

    def _parse_shareholding(self, soup: BeautifulSoup) -> pd.DataFrame:
        """
        Parse shareholding section — quarter columns, holder-type rows.
        Returns DataFrame with rows = holder types, columns = quarter labels.
        """
        section = soup.find("section", {"id": "shareholding"})
        if section is None:
            return pd.DataFrame()

        table = section.find("table")
        if table is None:
            return pd.DataFrame()

        thead = table.find("thead")
        col_labels = []
        if thead:
            ths = thead.find_all("th")
            for th in ths[1:]:
                text = th.get_text(strip=True)
                if text:
                    col_labels.append(_quarter_label(text))

        tbody = table.find("tbody")
        if tbody is None:
            return pd.DataFrame()

        rows = {}
        for tr in tbody.find_all("tr"):
            cells = tr.find_all("td")
            if not cells:
                continue
            row_label = cells[0].get_text(strip=True)
            if not row_label:
                continue
            values = [_parse_number(td.get_text(strip=True)) for td in cells[1:]]
            if col_labels:
                if len(values) < len(col_labels):
                    values.extend([float("nan")] * (len(col_labels) - len(values)))
                else:
                    values = values[:len(col_labels)]
            rows[row_label] = values

        if not rows:
            return pd.DataFrame()

        df = pd.DataFrame(rows, index=col_labels if col_labels else None).T
        df.index.name = "holder"
        return df

    def _parse_metadata(self, soup: BeautifulSoup) -> dict:
        """Extract key metadata (market cap, face value, etc.) from the page."""
        metadata = {}
        # screener.in shows metadata in a <ul class="company-ratios"> or similar
        for li in soup.select("ul.company-ratios li, #top-ratios li"):
            name_tag = li.find("span", class_="name")
            val_tag = li.find("span", class_="value") or li.find("span", class_="number")
            if name_tag and val_tag:
                key = name_tag.get_text(strip=True)
                val = val_tag.get_text(strip=True)
                metadata[key] = val

        # Also try structured data-attributes
        for tag in soup.select("[data-field]"):
            key = tag.get("data-field", "")
            val = tag.get_text(strip=True)
            if key:
                metadata.setdefault(key, val)

        return metadata

    # ------------------------------------------------------------------
    # Cache
    # ------------------------------------------------------------------

    def _init_db(self) -> None:
        """Create SQLite tables if not exist."""
        db_dir = os.path.dirname(self.db_path)
        if db_dir:
            os.makedirs(db_dir, exist_ok=True)

        with sqlite3.connect(self.db_path) as conn:
            conn.execute(
                """
                CREATE TABLE IF NOT EXISTS screener_cache (
                    ticker      TEXT PRIMARY KEY,
                    fetched_at  TEXT,
                    is_consolidated INTEGER,
                    data_json   TEXT
                )
                """
            )
            conn.commit()
        logger.debug("SQLite DB initialised at %s", self.db_path)

    def _cache_get(self, ticker: str) -> Optional["CompanyData"]:
        """Return cached CompanyData if within TTL, else None."""
        try:
            with sqlite3.connect(self.db_path) as conn:
                row = conn.execute(
                    "SELECT fetched_at, is_consolidated, data_json FROM screener_cache WHERE ticker = ?",
                    (ticker,),
                ).fetchone()

            if row is None:
                return None

            fetched_at_str, is_consolidated, data_json_str = row
            fetched_at = datetime.fromisoformat(fetched_at_str)
            age = datetime.utcnow() - fetched_at

            if age > timedelta(days=SCREENER_CACHE_TTL_DAYS):
                logger.debug("Cache expired for %s (age: %s)", ticker, age)
                return None

            return self._deserialize(data_json_str, ticker, bool(is_consolidated), fetched_at)

        except Exception as exc:
            logger.warning("Cache read error for %s: %s", ticker, exc)
            return None

    def _cache_set(self, ticker: str, data: CompanyData) -> None:
        """Store CompanyData in SQLite as JSON."""
        try:
            data_json_str = self._serialize(data)
            with sqlite3.connect(self.db_path) as conn:
                conn.execute(
                    """
                    INSERT OR REPLACE INTO screener_cache
                        (ticker, fetched_at, is_consolidated, data_json)
                    VALUES (?, ?, ?, ?)
                    """,
                    (
                        ticker,
                        data.fetched_at.isoformat(),
                        int(data.is_consolidated),
                        data_json_str,
                    ),
                )
                conn.commit()
            logger.debug("Cached data for %s", ticker)
        except Exception as exc:
            logger.warning("Cache write error for %s: %s", ticker, exc)

    # ------------------------------------------------------------------
    # Serialisation helpers
    # ------------------------------------------------------------------

    @staticmethod
    def _serialize(data: CompanyData) -> str:
        """Serialize CompanyData to a JSON string."""
        df_fields = [
            "income_statement",
            "balance_sheet",
            "cash_flow",
            "ratios",
            "shareholding",
            "quarterly_results",
        ]
        payload = {
            "ticker": data.ticker,
            "company_name": data.company_name,
            "is_consolidated": data.is_consolidated,
            "fetched_at": data.fetched_at.isoformat(),
            "raw_metadata": data.raw_metadata,
        }
        for f in df_fields:
            df: pd.DataFrame = getattr(data, f)
            payload[f] = df.to_json(orient="split") if df is not None and not df.empty else None

        return json.dumps(payload)

    @staticmethod
    def _deserialize(
        data_json_str: str, ticker: str, is_consolidated: bool, fetched_at: datetime
    ) -> CompanyData:
        """Reconstruct CompanyData from a JSON string."""
        payload = json.loads(data_json_str)

        df_fields = [
            "income_statement",
            "balance_sheet",
            "cash_flow",
            "ratios",
            "shareholding",
            "quarterly_results",
        ]
        dfs = {}
        for f in df_fields:
            raw = payload.get(f)
            if raw:
                try:
                    df = pd.read_json(raw, orient="split")
                    dfs[f] = df
                except Exception:
                    dfs[f] = pd.DataFrame()
            else:
                dfs[f] = pd.DataFrame()

        return CompanyData(
            ticker=payload.get("ticker", ticker),
            company_name=payload.get("company_name", ticker),
            is_consolidated=payload.get("is_consolidated", is_consolidated),
            fetched_at=fetched_at,
            income_statement=dfs["income_statement"],
            balance_sheet=dfs["balance_sheet"],
            cash_flow=dfs["cash_flow"],
            ratios=dfs["ratios"],
            shareholding=dfs["shareholding"],
            quarterly_results=dfs["quarterly_results"],
            raw_metadata=payload.get("raw_metadata", {}),
        )
