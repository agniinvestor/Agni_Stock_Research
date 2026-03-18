"""
Fetches price history and market data for NSE/BSE-listed Indian stocks via yfinance.

Usage:
    from src.data.price_client import PriceClient
    client = PriceClient()
    hist = client.get_history("ITC")
    cmp = client.get_current_price("ITC")
"""

import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', '..'))

import json
import logging
import sqlite3
from datetime import datetime, timedelta
from typing import Optional

import pandas as pd
import yfinance as yf

try:
    from config.settings import SQLITE_DB_PATH, PRICE_CACHE_TTL_DAYS, YF_DEFAULT_EXCHANGE
except ImportError:
    SQLITE_DB_PATH = os.path.join(os.path.expanduser("~"), ".india_stock_research", "cache.db")
    PRICE_CACHE_TTL_DAYS = 1
    YF_DEFAULT_EXCHANGE = "NSE"

logger = logging.getLogger(__name__)

_EXCHANGE_SUFFIX = {
    "NSE": ".NS",
    "BSE": ".BO",
}

# INR conversion: yfinance returns marketCap in raw INR for Indian stocks
_CR = 1e7  # 1 Crore = 10,000,000


class PriceClient:
    """Fetches and caches price history and market data for Indian stocks via yfinance."""

    def __init__(self, db_path: str = None):
        self.db_path = db_path or SQLITE_DB_PATH
        self._init_db()

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def get_history(
        self,
        ticker: str,
        exchange: str = None,
        period: str = "5y",
    ) -> pd.DataFrame:
        """
        Return OHLCV DataFrame with DatetimeIndex.
        Columns: Open, High, Low, Close, Volume, Adj_Close.
        Cached with a 1-day TTL.
        """
        exchange = (exchange or YF_DEFAULT_EXCHANGE).upper()
        ticker = ticker.upper().strip()

        cached = self._price_cache_get(ticker, exchange, period)
        if cached is not None:
            logger.info("Price cache hit for %s.%s (%s)", ticker, exchange, period)
            return cached

        yf_str = self.yf_ticker_str(ticker, exchange)
        logger.info("Fetching price history for %s (period=%s)", yf_str, period)

        try:
            raw = yf.Ticker(yf_str).history(period=period, auto_adjust=True)
        except Exception as exc:
            logger.error("yfinance history fetch failed for %s: %s", yf_str, exc)
            raise

        if raw.empty:
            logger.warning("No price data returned for %s", yf_str)
            return pd.DataFrame()

        # Normalise columns
        df = raw.copy()
        df.index = pd.to_datetime(df.index).tz_localize(None)

        rename_map = {}
        for col in df.columns:
            col_lower = col.lower()
            if col_lower == "open":
                rename_map[col] = "Open"
            elif col_lower == "high":
                rename_map[col] = "High"
            elif col_lower == "low":
                rename_map[col] = "Low"
            elif col_lower == "close":
                rename_map[col] = "Close"
            elif col_lower == "volume":
                rename_map[col] = "Volume"
            elif col_lower in ("adj close", "adj_close", "adjclose"):
                rename_map[col] = "Adj_Close"
        df.rename(columns=rename_map, inplace=True)

        # If Adj_Close is missing (auto_adjust=True folds it into Close), create it
        if "Adj_Close" not in df.columns and "Close" in df.columns:
            df["Adj_Close"] = df["Close"]

        # Keep only canonical columns
        keep = [c for c in ("Open", "High", "Low", "Close", "Volume", "Adj_Close") if c in df.columns]
        df = df[keep]

        self._price_cache_set(ticker, exchange, period, df)
        return df

    def get_current_price(self, ticker: str, exchange: str = None) -> float:
        """Return latest closing price in INR."""
        exchange = (exchange or YF_DEFAULT_EXCHANGE).upper()
        ticker = ticker.upper().strip()

        try:
            mdata = self.get_market_data(ticker, exchange)
            cmp = mdata.get("cmp")
            if cmp is not None and cmp > 0:
                return float(cmp)
        except Exception as exc:
            logger.warning("get_market_data failed for CMP, falling back to history: %s", exc)

        # Fallback: last close from history
        hist = self.get_history(ticker, exchange, period="5d")
        if hist.empty or "Close" not in hist.columns:
            raise ValueError(f"Cannot determine current price for {ticker}")
        return float(hist["Close"].dropna().iloc[-1])

    def get_market_data(self, ticker: str, exchange: str = None) -> dict:
        """
        Return dict with market data fields:
            cmp, market_cap_cr, shares_outstanding_cr,
            week52_high, week52_low, pe_ttm, pb,
            dividend_yield_pct, beta
        """
        exchange = (exchange or YF_DEFAULT_EXCHANGE).upper()
        ticker = ticker.upper().strip()

        cached = self._mdata_cache_get(ticker, exchange)
        if cached is not None:
            logger.info("Market data cache hit for %s.%s", ticker, exchange)
            return cached

        yf_str = self.yf_ticker_str(ticker, exchange)
        logger.info("Fetching market data for %s", yf_str)

        try:
            info = yf.Ticker(yf_str).info
        except Exception as exc:
            logger.error("yfinance .info failed for %s: %s", yf_str, exc)
            info = {}

        result = self._extract_market_data(info, ticker, exchange)

        # If core fields are missing, compute from history
        if result.get("cmp") is None or result.get("week52_high") is None:
            result = self._enrich_from_history(result, ticker, exchange, info)

        self._mdata_cache_set(ticker, exchange, result)
        return result

    def get_peer_market_data(
        self,
        tickers: list,
        exchange: str = None,
    ) -> dict:
        """
        Fetch market data for multiple tickers.
        Returns {ticker: market_data_dict}.
        """
        exchange = (exchange or YF_DEFAULT_EXCHANGE).upper()
        results = {}
        for t in tickers:
            try:
                results[t.upper().strip()] = self.get_market_data(t, exchange)
            except Exception as exc:
                logger.warning("Failed to fetch market data for %s: %s", t, exc)
                results[t.upper().strip()] = {}
        return results

    def yf_ticker_str(self, ticker: str, exchange: str) -> str:
        """
        Convert ticker + exchange to yfinance format.
        "ITC", "NSE"    -> "ITC.NS"
        "500875", "BSE" -> "500875.BO"
        """
        suffix = _EXCHANGE_SUFFIX.get(exchange.upper(), ".NS")
        return f"{ticker}{suffix}"

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------

    def _extract_market_data(self, info: dict, ticker: str, exchange: str) -> dict:
        """Extract market data fields from yfinance .info dict."""
        def _get(*keys):
            for k in keys:
                v = info.get(k)
                if v is not None:
                    try:
                        return float(v)
                    except (TypeError, ValueError):
                        pass
            return None

        cmp = _get("currentPrice", "regularMarketPrice", "previousClose")

        raw_market_cap = _get("marketCap")
        market_cap_cr = raw_market_cap / _CR if raw_market_cap is not None else None

        raw_shares = _get("sharesOutstanding")
        shares_outstanding_cr = raw_shares / _CR if raw_shares is not None else None

        week52_high = _get("fiftyTwoWeekHigh")
        week52_low = _get("fiftyTwoWeekLow")

        pe_ttm = _get("trailingPE")
        pb = _get("priceToBook")

        raw_div_yield = _get("dividendYield")
        dividend_yield_pct = (raw_div_yield * 100) if raw_div_yield is not None else None

        beta = _get("beta")

        return {
            "cmp": cmp,
            "market_cap_cr": market_cap_cr,
            "shares_outstanding_cr": shares_outstanding_cr,
            "week52_high": week52_high,
            "week52_low": week52_low,
            "pe_ttm": pe_ttm,
            "pb": pb,
            "dividend_yield_pct": dividend_yield_pct,
            "beta": beta,
        }

    def _enrich_from_history(
        self, result: dict, ticker: str, exchange: str, info: dict
    ) -> dict:
        """Fill in missing market data fields by computing from price history."""
        try:
            hist = self.get_history(ticker, exchange, period="1y")
        except Exception as exc:
            logger.warning("History fetch for enrichment failed: %s", exc)
            return result

        if hist.empty or "Close" not in hist.columns:
            return result

        close = hist["Close"].dropna()
        if close.empty:
            return result

        if result.get("cmp") is None:
            result["cmp"] = float(close.iloc[-1])

        if result.get("week52_high") is None:
            last_252 = close.tail(252)
            result["week52_high"] = float(last_252.max())
            result["week52_low"] = float(last_252.min())

        # Compute market cap from price * shares if available
        if result.get("market_cap_cr") is None and result.get("shares_outstanding_cr") is not None:
            result["market_cap_cr"] = (
                result["cmp"] * result["shares_outstanding_cr"] * _CR / _CR
            )

        return result

    # ------------------------------------------------------------------
    # SQLite cache
    # ------------------------------------------------------------------

    def _init_db(self) -> None:
        """Create SQLite tables if not exist."""
        db_dir = os.path.dirname(self.db_path)
        if db_dir:
            os.makedirs(db_dir, exist_ok=True)

        with sqlite3.connect(self.db_path) as conn:
            conn.execute(
                """
                CREATE TABLE IF NOT EXISTS price_cache (
                    ticker      TEXT,
                    exchange    TEXT,
                    fetched_at  TEXT,
                    period      TEXT,
                    ohlcv_json  TEXT,
                    PRIMARY KEY (ticker, exchange, period)
                )
                """
            )
            conn.execute(
                """
                CREATE TABLE IF NOT EXISTS market_data_cache (
                    ticker      TEXT,
                    exchange    TEXT,
                    fetched_at  TEXT,
                    data_json   TEXT,
                    PRIMARY KEY (ticker, exchange)
                )
                """
            )
            conn.commit()
        logger.debug("PriceClient SQLite DB initialised at %s", self.db_path)

    def _price_cache_get(
        self, ticker: str, exchange: str, period: str
    ) -> Optional[pd.DataFrame]:
        """Return cached OHLCV DataFrame if within TTL, else None."""
        try:
            with sqlite3.connect(self.db_path) as conn:
                row = conn.execute(
                    "SELECT fetched_at, ohlcv_json FROM price_cache WHERE ticker=? AND exchange=? AND period=?",
                    (ticker, exchange, period),
                ).fetchone()

            if row is None:
                return None

            fetched_at = datetime.fromisoformat(row[0])
            age = datetime.utcnow() - fetched_at
            if age > timedelta(days=PRICE_CACHE_TTL_DAYS):
                logger.debug("Price cache expired for %s.%s (%s)", ticker, exchange, period)
                return None

            df = pd.read_json(row[1], orient="split")
            df.index = pd.to_datetime(df.index)
            return df

        except Exception as exc:
            logger.warning("Price cache read error for %s.%s: %s", ticker, exchange, exc)
            return None

    def _price_cache_set(
        self, ticker: str, exchange: str, period: str, df: pd.DataFrame
    ) -> None:
        """Store OHLCV DataFrame in SQLite."""
        try:
            ohlcv_json = df.to_json(orient="split", date_format="iso")
            now = datetime.utcnow().isoformat()
            with sqlite3.connect(self.db_path) as conn:
                conn.execute(
                    """
                    INSERT OR REPLACE INTO price_cache
                        (ticker, exchange, fetched_at, period, ohlcv_json)
                    VALUES (?, ?, ?, ?, ?)
                    """,
                    (ticker, exchange, now, period, ohlcv_json),
                )
                conn.commit()
            logger.debug("Cached price history for %s.%s (%s)", ticker, exchange, period)
        except Exception as exc:
            logger.warning("Price cache write error for %s.%s: %s", ticker, exchange, exc)

    def _mdata_cache_get(self, ticker: str, exchange: str) -> Optional[dict]:
        """Return cached market data dict if within TTL, else None."""
        try:
            with sqlite3.connect(self.db_path) as conn:
                row = conn.execute(
                    "SELECT fetched_at, data_json FROM market_data_cache WHERE ticker=? AND exchange=?",
                    (ticker, exchange),
                ).fetchone()

            if row is None:
                return None

            fetched_at = datetime.fromisoformat(row[0])
            age = datetime.utcnow() - fetched_at
            if age > timedelta(days=PRICE_CACHE_TTL_DAYS):
                logger.debug("Market data cache expired for %s.%s", ticker, exchange)
                return None

            return json.loads(row[1])

        except Exception as exc:
            logger.warning("Market data cache read error for %s.%s: %s", ticker, exchange, exc)
            return None

    def _mdata_cache_set(self, ticker: str, exchange: str, data: dict) -> None:
        """Store market data dict in SQLite."""
        try:
            now = datetime.utcnow().isoformat()
            with sqlite3.connect(self.db_path) as conn:
                conn.execute(
                    """
                    INSERT OR REPLACE INTO market_data_cache
                        (ticker, exchange, fetched_at, data_json)
                    VALUES (?, ?, ?, ?)
                    """,
                    (ticker, exchange, now, json.dumps(data)),
                )
                conn.commit()
            logger.debug("Cached market data for %s.%s", ticker, exchange)
        except Exception as exc:
            logger.warning("Market data cache write error for %s.%s: %s", ticker, exchange, exc)
