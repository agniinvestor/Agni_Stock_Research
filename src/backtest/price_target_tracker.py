"""
Backtesting framework for price target accuracy tracking.

Stores price targets issued with their date, then computes:
- Hit rate: % of BUY recommendations that outperformed Nifty 50 in 12 months
- MAPE: Mean Absolute Percentage Error of earnings estimates vs actuals
- Outcome distribution: 3m / 6m / 12m actual return vs target return

Schema (SQLite):
    price_targets(
        id            INTEGER PRIMARY KEY,
        ticker        TEXT NOT NULL,
        report_date   TEXT NOT NULL,   -- YYYY-MM-DD
        report_type   TEXT,            -- 'initiation' | 'earnings_update'
        rating        TEXT,            -- 'BUY' | 'HOLD' | 'SELL'
        cmp_at_issue  REAL,            -- ₹ at time of report
        price_target  REAL,            -- ₹
        upside_pct    REAL,            -- (pt/cmp - 1) * 100
        implied_return_pct REAL,       -- same as upside_pct
        analyst_pat_fy1e  REAL,        -- PAT estimate for FY+1, INR Cr
        analyst_pat_fy2e  REAL,        -- PAT estimate for FY+2, INR Cr
        notes         TEXT
    )

    price_outcomes(
        id              INTEGER PRIMARY KEY,
        target_id       INTEGER REFERENCES price_targets(id),
        ticker          TEXT,
        outcome_date    TEXT,           -- YYYY-MM-DD
        horizon_months  INTEGER,        -- 3 | 6 | 12
        price_actual    REAL,           -- price on outcome_date
        return_pct      REAL,           -- (actual / cmp_at_issue - 1) * 100
        nifty_return_pct REAL,          -- Nifty 50 return over same period
        alpha_pct       REAL,           -- return_pct - nifty_return_pct
        hit             INTEGER,        -- 1 if price_target reached, 0 if not
        outperformed_nifty INTEGER,     -- 1 if alpha > 0, 0 if not
        pat_actual_fy1  REAL,           -- actual PAT for FY+1 (filled in later)
        pat_mape_fy1    REAL            -- |pat_actual - pat_est| / pat_actual * 100
    )

Usage:
    tracker = PriceTargetTracker(db_path="data/processed/cache.db")

    # Record a new price target when issuing a report
    target_id = tracker.record_target(
        ticker="ITC",
        report_date="2026-03-18",
        rating="BUY",
        cmp=307.0,
        price_target=380.0,
        report_type="initiation",
        analyst_pat_fy1e=20500.0,
        analyst_pat_fy2e=22000.0,
    )

    # Update outcomes (run periodically or on demand)
    tracker.update_outcomes(ticker="ITC", price_client=price_client)

    # Get analytics
    stats = tracker.compute_stats()
    print(stats.hit_rate_12m)     # e.g. 0.65 (65% of BUYs hit target in 12m)
    print(stats.alpha_12m_avg)    # e.g. 8.3 (avg alpha vs Nifty)
"""

import sqlite3
import logging
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from typing import Optional

logger = logging.getLogger(__name__)

NIFTY_TICKER = "^NSEI"   # yfinance ticker for Nifty 50


# ---------------------------------------------------------------------------
# Dataclasses
# ---------------------------------------------------------------------------

@dataclass
class PriceTarget:
    """A recorded price target at time of report issuance."""
    id: int
    ticker: str
    report_date: str
    report_type: str
    rating: str
    cmp_at_issue: float
    price_target: float
    upside_pct: float
    analyst_pat_fy1e: Optional[float]
    analyst_pat_fy2e: Optional[float]
    notes: str


@dataclass
class PriceOutcome:
    """Actual price outcome vs target, measured at a horizon."""
    id: int
    target_id: int
    ticker: str
    outcome_date: str
    horizon_months: int
    price_actual: float
    return_pct: float
    nifty_return_pct: Optional[float]
    alpha_pct: Optional[float]
    hit: bool
    outperformed_nifty: Optional[bool]
    pat_actual_fy1: Optional[float]
    pat_mape_fy1: Optional[float]


@dataclass
class BacktestStats:
    """Aggregated statistics across all tracked recommendations."""
    total_targets: int
    buy_count: int
    hold_count: int
    sell_count: int

    # 12-month metrics (primary)
    hit_rate_12m: Optional[float]           # % of BUYs where actual >= price_target
    outperform_rate_12m: Optional[float]    # % of BUYs with positive alpha
    alpha_12m_avg: Optional[float]          # avg alpha (return - nifty) in %
    alpha_12m_median: Optional[float]

    # 6-month metrics
    hit_rate_6m: Optional[float]
    outperform_rate_6m: Optional[float]
    alpha_6m_avg: Optional[float]

    # 3-month metrics
    hit_rate_3m: Optional[float]
    outperform_rate_3m: Optional[float]

    # Earnings accuracy
    pat_mape_fy1_avg: Optional[float]       # avg MAPE for FY+1 PAT estimates
    pat_mape_fy1_median: Optional[float]

    # By ticker
    per_ticker: dict = field(default_factory=dict)


# ---------------------------------------------------------------------------
# Tracker
# ---------------------------------------------------------------------------

class PriceTargetTracker:

    def __init__(self, db_path: str):
        self.db_path = db_path
        self._init_db()

    # ── Schema init ──────────────────────────────────────────────────────────

    def _init_db(self):
        with sqlite3.connect(self.db_path) as conn:
            conn.execute("""
                CREATE TABLE IF NOT EXISTS price_targets (
                    id               INTEGER PRIMARY KEY AUTOINCREMENT,
                    ticker           TEXT NOT NULL,
                    report_date      TEXT NOT NULL,
                    report_type      TEXT DEFAULT 'initiation',
                    rating           TEXT,
                    cmp_at_issue     REAL,
                    price_target     REAL,
                    upside_pct       REAL,
                    implied_return_pct REAL,
                    analyst_pat_fy1e REAL,
                    analyst_pat_fy2e REAL,
                    notes            TEXT,
                    UNIQUE(ticker, report_date, report_type)
                )
            """)
            conn.execute("""
                CREATE TABLE IF NOT EXISTS price_outcomes (
                    id                   INTEGER PRIMARY KEY AUTOINCREMENT,
                    target_id            INTEGER REFERENCES price_targets(id),
                    ticker               TEXT,
                    outcome_date         TEXT,
                    horizon_months       INTEGER,
                    price_actual         REAL,
                    return_pct           REAL,
                    nifty_return_pct     REAL,
                    alpha_pct            REAL,
                    hit                  INTEGER DEFAULT 0,
                    outperformed_nifty   INTEGER,
                    pat_actual_fy1       REAL,
                    pat_mape_fy1         REAL,
                    UNIQUE(target_id, horizon_months)
                )
            """)
            conn.commit()

    # ── Record ───────────────────────────────────────────────────────────────

    def record_target(
        self,
        ticker: str,
        report_date: str,
        rating: str,
        cmp: float,
        price_target: float,
        report_type: str = "initiation",
        analyst_pat_fy1e: float = None,
        analyst_pat_fy2e: float = None,
        notes: str = "",
    ) -> int:
        """
        Record a price target at time of report issuance.
        Returns the row ID.
        """
        upside = ((price_target / cmp) - 1.0) * 100.0 if cmp else None
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.execute("""
                INSERT OR REPLACE INTO price_targets
                (ticker, report_date, report_type, rating, cmp_at_issue,
                 price_target, upside_pct, implied_return_pct,
                 analyst_pat_fy1e, analyst_pat_fy2e, notes)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (ticker, report_date, report_type, rating.upper(), cmp,
                  price_target, upside, upside,
                  analyst_pat_fy1e, analyst_pat_fy2e, notes))
            conn.commit()
            target_id = cursor.lastrowid
        logger.info(f"[tracker] Recorded {rating} {ticker} PT=₹{price_target:.0f} "
                    f"({upside:+.1f}%) id={target_id}")
        return target_id

    # ── Update outcomes ───────────────────────────────────────────────────────

    def update_outcomes(self, ticker: str = None, price_client=None) -> int:
        """
        Fetch actual prices and compute outcomes for all pending targets.

        For each target where a 3/6/12-month horizon date has passed,
        fetch the actual price (and Nifty 50) and store the outcome.

        Returns count of outcomes updated.
        """
        if price_client is None:
            logger.warning("[tracker] No price_client provided — skipping outcome update")
            return 0

        with sqlite3.connect(self.db_path) as conn:
            conn.row_factory = sqlite3.Row
            query = "SELECT * FROM price_targets"
            params = []
            if ticker:
                query += " WHERE ticker = ?"
                params.append(ticker.upper())
            targets = conn.execute(query, params).fetchall()

        today      = date.today()
        updated    = 0

        for t in targets:
            report_dt = datetime.strptime(t["report_date"], "%Y-%m-%d").date()

            for horizon in (3, 6, 12):
                outcome_date = report_dt + timedelta(days=horizon * 30)
                if outcome_date > today:
                    continue   # horizon not yet reached

                # Check if already recorded
                with sqlite3.connect(self.db_path) as conn:
                    exists = conn.execute(
                        "SELECT 1 FROM price_outcomes WHERE target_id=? AND horizon_months=?",
                        (t["id"], horizon)
                    ).fetchone()
                if exists:
                    continue

                try:
                    actual_price = self._fetch_price_on_date(
                        price_client, t["ticker"], outcome_date
                    )
                    nifty_price_issue   = self._fetch_price_on_date(
                        price_client, NIFTY_TICKER, report_dt
                    )
                    nifty_price_outcome = self._fetch_price_on_date(
                        price_client, NIFTY_TICKER, outcome_date
                    )

                    if actual_price is None or t["cmp_at_issue"] is None:
                        continue

                    ret_pct = ((actual_price / t["cmp_at_issue"]) - 1.0) * 100.0
                    hit     = int(actual_price >= t["price_target"])

                    nifty_ret = None
                    alpha     = None
                    outperformed = None
                    if nifty_price_issue and nifty_price_outcome:
                        nifty_ret    = ((nifty_price_outcome / nifty_price_issue) - 1.0) * 100.0
                        alpha        = ret_pct - nifty_ret
                        outperformed = int(alpha > 0)

                    with sqlite3.connect(self.db_path) as conn:
                        conn.execute("""
                            INSERT OR IGNORE INTO price_outcomes
                            (target_id, ticker, outcome_date, horizon_months,
                             price_actual, return_pct, nifty_return_pct,
                             alpha_pct, hit, outperformed_nifty)
                            VALUES (?,?,?,?,?,?,?,?,?,?)
                        """, (t["id"], t["ticker"], outcome_date.isoformat(), horizon,
                              actual_price, ret_pct, nifty_ret, alpha, hit, outperformed))
                        conn.commit()

                    logger.info(
                        f"[tracker] {t['ticker']} {horizon}m: "
                        f"price={actual_price:.1f} ret={ret_pct:+.1f}% "
                        f"alpha={alpha:+.1f}%" if alpha else
                        f"[tracker] {t['ticker']} {horizon}m: "
                        f"price={actual_price:.1f} ret={ret_pct:+.1f}%"
                    )
                    updated += 1

                except Exception as e:
                    logger.warning(f"[tracker] Failed outcome for {t['ticker']} {horizon}m: {e}")

        return updated

    def record_pat_actual(self, ticker: str, fy_label: str, pat_actual: float):
        """
        Fill in actual PAT for FY+1 once annual results are known.
        Computes MAPE against the analyst estimate recorded at target issuance.
        """
        with sqlite3.connect(self.db_path) as conn:
            conn.row_factory = sqlite3.Row
            # Find all targets for this ticker where FY+1 estimate exists
            targets = conn.execute(
                "SELECT * FROM price_targets WHERE ticker=? AND analyst_pat_fy1e IS NOT NULL",
                (ticker.upper(),)
            ).fetchall()

            for t in targets:
                if t["analyst_pat_fy1e"] and t["analyst_pat_fy1e"] != 0:
                    mape = abs(pat_actual - t["analyst_pat_fy1e"]) / abs(pat_actual) * 100.0
                    conn.execute("""
                        UPDATE price_outcomes
                        SET pat_actual_fy1 = ?, pat_mape_fy1 = ?
                        WHERE target_id = ? AND horizon_months = 12
                    """, (pat_actual, mape, t["id"]))
            conn.commit()

        logger.info(f"[tracker] Updated PAT actuals for {ticker} {fy_label}")

    # ── Analytics ────────────────────────────────────────────────────────────

    def compute_stats(self, ticker: str = None,
                      rating_filter: str = None) -> BacktestStats:
        """
        Compute aggregate accuracy statistics across all tracked targets.

        Args:
            ticker: filter to single ticker (None = all)
            rating_filter: filter to 'BUY'/'HOLD'/'SELL' (None = all)
        """
        with sqlite3.connect(self.db_path) as conn:
            conn.row_factory = sqlite3.Row

            # Targets
            q = "SELECT * FROM price_targets WHERE 1=1"
            p = []
            if ticker:
                q += " AND ticker=?"; p.append(ticker.upper())
            if rating_filter:
                q += " AND rating=?"; p.append(rating_filter.upper())
            targets = conn.execute(q, p).fetchall()

            # All outcomes
            q2 = """
                SELECT po.*, pt.rating, pt.cmp_at_issue, pt.price_target
                FROM price_outcomes po
                JOIN price_targets pt ON po.target_id = pt.id
                WHERE 1=1
            """
            p2 = []
            if ticker:
                q2 += " AND po.ticker=?"; p2.append(ticker.upper())
            if rating_filter:
                q2 += " AND pt.rating=?"; p2.append(rating_filter.upper())
            outcomes = conn.execute(q2, p2).fetchall()

        def _avg(vals):
            vals = [v for v in vals if v is not None]
            return sum(vals) / len(vals) if vals else None

        def _median(vals):
            vals = sorted(v for v in vals if v is not None)
            n = len(vals)
            if not n:
                return None
            return vals[n // 2] if n % 2 else (vals[n//2 - 1] + vals[n//2]) / 2

        def _rate(outcomes_h, condition):
            h = [o for o in outcomes_h if o["rating"] == "BUY"]
            if not h:
                return None
            passing = sum(1 for o in h if condition(o))
            return (passing / len(h)) * 100.0

        by_horizon = {h: [o for o in outcomes if o["horizon_months"] == h]
                      for h in (3, 6, 12)}

        stats = BacktestStats(
            total_targets=len(targets),
            buy_count=sum(1 for t in targets if t["rating"] == "BUY"),
            hold_count=sum(1 for t in targets if t["rating"] == "HOLD"),
            sell_count=sum(1 for t in targets if t["rating"] == "SELL"),

            hit_rate_12m=_rate(by_horizon[12], lambda o: o["hit"] == 1),
            outperform_rate_12m=_rate(by_horizon[12], lambda o: o["outperformed_nifty"] == 1),
            alpha_12m_avg=_avg([o["alpha_pct"] for o in by_horizon[12] if o["rating"] == "BUY"]),
            alpha_12m_median=_median([o["alpha_pct"] for o in by_horizon[12] if o["rating"] == "BUY"]),

            hit_rate_6m=_rate(by_horizon[6], lambda o: o["hit"] == 1),
            outperform_rate_6m=_rate(by_horizon[6], lambda o: o["outperformed_nifty"] == 1),
            alpha_6m_avg=_avg([o["alpha_pct"] for o in by_horizon[6] if o["rating"] == "BUY"]),

            hit_rate_3m=_rate(by_horizon[3], lambda o: o["hit"] == 1),
            outperform_rate_3m=_rate(by_horizon[3], lambda o: o["outperformed_nifty"] == 1),

            pat_mape_fy1_avg=_avg([o["pat_mape_fy1"] for o in outcomes if o["pat_mape_fy1"]]),
            pat_mape_fy1_median=_median([o["pat_mape_fy1"] for o in outcomes if o["pat_mape_fy1"]]),
        )

        # Per-ticker breakdown
        all_tickers = sorted(set(t["ticker"] for t in targets))
        for t in all_tickers:
            t_outcomes_12 = [o for o in by_horizon[12] if o["ticker"] == t and o["rating"] == "BUY"]
            stats.per_ticker[t] = {
                "hit_rate_12m":      _rate(t_outcomes_12, lambda o: o["hit"] == 1),
                "alpha_12m_avg":     _avg([o["alpha_pct"] for o in t_outcomes_12]),
                "pat_mape_fy1_avg":  _avg([o["pat_mape_fy1"] for o in t_outcomes_12 if o["pat_mape_fy1"]]),
            }

        return stats

    def print_scorecard(self, ticker: str = None):
        """Print a formatted scorecard to stdout."""
        stats = self.compute_stats(ticker=ticker, rating_filter="BUY")
        print(f"\n{'='*60}")
        print(f"  PRICE TARGET SCORECARD" + (f" — {ticker}" if ticker else " — ALL"))
        print(f"{'='*60}")
        print(f"  Total BUY recs tracked  : {stats.buy_count}")
        print(f"  12m Hit Rate (pt reached): {stats.hit_rate_12m:.1f}%" if stats.hit_rate_12m else "  12m Hit Rate: N/A")
        print(f"  12m Outperform Rate     : {stats.outperform_rate_12m:.1f}%" if stats.outperform_rate_12m else "  12m Outperform Rate: N/A")
        print(f"  12m Avg Alpha vs Nifty  : {stats.alpha_12m_avg:+.1f}%" if stats.alpha_12m_avg else "  12m Avg Alpha: N/A")
        print(f"  6m Outperform Rate      : {stats.outperform_rate_6m:.1f}%" if stats.outperform_rate_6m else "  6m Outperform Rate: N/A")
        print(f"  PAT Estimate MAPE (FY1) : {stats.pat_mape_fy1_avg:.1f}%" if stats.pat_mape_fy1_avg else "  PAT MAPE: N/A (no actuals yet)")
        if stats.per_ticker:
            print(f"\n  Per-ticker (12m):")
            for t, s in sorted(stats.per_ticker.items()):
                hit  = f"{s['hit_rate_12m']:.0f}%" if s['hit_rate_12m'] is not None else "N/A"
                alph = f"{s['alpha_12m_avg']:+.1f}%" if s['alpha_12m_avg'] is not None else "N/A"
                print(f"    {t:15s} hit={hit:6s}  alpha={alph}")
        print(f"{'='*60}\n")

    def list_targets(self, ticker: str = None) -> list[dict]:
        """Return all recorded targets as list of dicts."""
        with sqlite3.connect(self.db_path) as conn:
            conn.row_factory = sqlite3.Row
            q = "SELECT * FROM price_targets ORDER BY report_date DESC"
            p = []
            if ticker:
                q = "SELECT * FROM price_targets WHERE ticker=? ORDER BY report_date DESC"
                p = [ticker.upper()]
            rows = conn.execute(q, p).fetchall()
        return [dict(r) for r in rows]

    # ── Internal helpers ──────────────────────────────────────────────────────

    @staticmethod
    def _fetch_price_on_date(price_client, ticker: str,
                             target_date: date) -> Optional[float]:
        """
        Fetch closing price for ticker on or just before target_date.
        Uses PriceClient.get_history() with a 5-day window.
        """
        try:
            start = target_date - timedelta(days=7)
            end   = target_date + timedelta(days=1)
            hist  = price_client.get_history(
                ticker, start=start.isoformat(), end=end.isoformat()
            )
            if hist is not None and not hist.empty:
                # Take last row on or before target_date
                hist.index = hist.index.normalize()
                subset = hist[hist.index.date <= target_date]
                if not subset.empty:
                    return float(subset["Close"].iloc[-1])
        except Exception as e:
            logger.debug(f"_fetch_price_on_date({ticker}, {target_date}): {e}")
        return None
