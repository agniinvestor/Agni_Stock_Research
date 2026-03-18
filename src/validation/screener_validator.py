"""
Cross-validation of screener.in scraped data against yfinance market data.

Catches gross scraping errors by comparing key metrics available from both
sources: revenue order of magnitude, market cap, EPS.

Usage:
    from src.validation.screener_validator import ScreenerValidator
    result = ScreenerValidator.validate(report_json, price_client)
"""

import logging
from src.validation.model_validator import CheckResult, ValidationReport

logger = logging.getLogger(__name__)


class ScreenerValidator:
    """
    Compares report_json (derived from screener.in) against yfinance data.
    Returns a ValidationReport (appends checks to an existing report or creates new).
    """

    @staticmethod
    def validate(report_json: dict, price_client=None,
                 existing_report: ValidationReport = None) -> ValidationReport:
        """
        price_client: PriceClient instance (optional — skips yfinance checks if None)
        existing_report: if provided, checks are appended to it
        """
        ticker     = report_json.get("meta", {}).get("ticker", "UNKNOWN")
        model_type = report_json.get("meta", {}).get("model_type", "generic")

        report = existing_report or ValidationReport(ticker=ticker, model_type=model_type)
        checks = report.checks

        mkt_screener = report_json.get("market_data", {})
        fin          = report_json.get("financials", {})
        is_          = fin.get("income_statement", {})
        fy_labels    = fin.get("fy_labels", [])
        actuals_end  = fin.get("actuals_end_idx", len(fy_labels) - 1)

        # ── 1. Market cap sanity (screener vs yfinance) ───────────────────────
        mcap_screener = mkt_screener.get("market_cap_cr")
        if price_client and mcap_screener:
            try:
                live = price_client.get_market_data(ticker)
                mcap_live = live.get("market_cap_cr")
                if mcap_live and mcap_live > 0 and mcap_screener > 0:
                    ratio = abs(mcap_screener - mcap_live) / mcap_live
                    checks.append(CheckResult(
                        name="mcap_cross_check",
                        passed=ratio < 0.25,   # within 25%
                        severity="warning",
                        message=(
                            f"Market cap: screener ₹{mcap_screener:,.0f} Cr vs "
                            f"yfinance ₹{mcap_live:,.0f} Cr ({ratio*100:.0f}% diff)"
                        ),
                    ))
            except Exception as e:
                logger.debug(f"yfinance cross-check skipped: {e}")

        # ── 2. Revenue order-of-magnitude check ───────────────────────────────
        rev_list = [v for v in is_.get("net_revenue", [])[:actuals_end+1]
                    if v is not None]
        if rev_list and mcap_screener and mcap_screener > 0:
            latest_rev = rev_list[-1]
            rev_to_mcap = latest_rev / mcap_screener
            # For most companies: revenue should be 10%-500% of market cap
            # Very asset-heavy cos (metals) may be higher; very capital-light lower
            checks.append(CheckResult(
                name="revenue_vs_mcap_ratio",
                passed=0.01 <= rev_to_mcap <= 10.0,
                severity="warning",
                message=(
                    f"Revenue/MCap = {rev_to_mcap:.2f}x "
                    f"(Rev ₹{latest_rev:,.0f} Cr, MCap ₹{mcap_screener:,.0f} Cr)"
                ),
                detail="Ratio outside expected 0.01x–10x — check data magnitude" if not (0.01 <= rev_to_mcap <= 10.0) else "",
            ))

        # ── 3. EPS sanity (screener vs yfinance EPS TTM) ──────────────────────
        eps_list = [v for v in is_.get("eps", [])[:actuals_end+1] if v is not None]
        eps_yf   = mkt_screener.get("eps_ttm")   # from PriceClient.get_market_data()
        if eps_list and eps_yf and eps_yf != 0:
            eps_screener = eps_list[-1]
            ratio = abs(eps_screener - eps_yf) / abs(eps_yf)
            checks.append(CheckResult(
                name="eps_cross_check",
                passed=ratio < 0.30,   # within 30%
                severity="warning",
                message=(
                    f"EPS: screener ₹{eps_screener:.2f} vs "
                    f"yfinance ₹{eps_yf:.2f} ({ratio*100:.0f}% diff)"
                ),
                detail="Large EPS divergence — check screener EPS row or yfinance TTM" if ratio >= 0.30 else "",
            ))

        # ── 4. FY label alignment check ───────────────────────────────────────
        # Latest FY in screener data should not be more than 2 years stale
        if fy_labels:
            from datetime import date
            current_year = date.today().year
            # FY25 → year 2025; extract last 2 chars
            try:
                latest_fy_year = 2000 + int(fy_labels[actuals_end][-2:])
                staleness = current_year - latest_fy_year
                checks.append(CheckResult(
                    name="data_staleness",
                    passed=staleness <= 2,
                    severity="warning",
                    message=f"Latest actual FY: {fy_labels[actuals_end]} ({staleness} years ago)",
                    detail="Data may be stale — verify screener.in shows latest annual results" if staleness > 2 else "",
                ))
            except (ValueError, IndexError):
                pass

        # ── 5. Consolidated check ─────────────────────────────────────────────
        is_consol = report_json.get("meta", {}).get("is_consolidated", True)
        checks.append(CheckResult(
            name="consolidated_financials",
            passed=is_consol,
            severity="warning",
            message="Using consolidated financials ✓" if is_consol
                    else "Using STANDALONE financials — verify no consolidated available",
        ))

        logger.info(f"[ScreenerValidator] {report.summary()}")
        return report
