"""
Main pipeline entry point for IndiaStockResearch.

Usage:
    python pipeline.py --ticker ITC --task data
    python pipeline.py --ticker ITC --task charts --charts-dir reports/ITC/charts
    python pipeline.py --ticker ITC --task report --output reports/ITC/ITC_Report_2026.docx
    python pipeline.py --ticker ITC --task all

Tasks:
    data    → fetch from screener.in + yfinance, compute ratios, save report_json
    charts  → generate all charts from report_json
    report  → assemble DOCX from report_json + charts
    all     → run data → charts → report in sequence
"""

import argparse
import json
import os
import sys
import logging
from datetime import date

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)

from config.settings import (
    REPORTS_DIR, SQLITE_DB_PATH, SCREENER_USERNAME, SCREENER_PASSWORD
)
from src.data.screener_client import ScreenerClient
from src.data.price_client import PriceClient
from src.models.fmcg_model import FMCGModel
from src.utils.fy_utils import fy_label, fy_range

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s"
)
log = logging.getLogger("pipeline")


def run_data(ticker: str, sector: str = "fmcg", force_refresh: bool = False) -> dict:
    """Fetch data, compute ratios, return report_json dict."""
    log.info(f"Fetching data for {ticker}")

    screener = ScreenerClient(
        db_path=SQLITE_DB_PATH,
        username=SCREENER_USERNAME,
        password=SCREENER_PASSWORD,
    )
    price_client = PriceClient(db_path=SQLITE_DB_PATH)

    company_data = screener.fetch_company(ticker, force_refresh=force_refresh)
    market_data  = price_client.get_market_data(ticker)

    log.info(f"  Company: {company_data.company_name}")
    log.info(f"  CMP: ₹{market_data.get('cmp', 'N/A')}")

    # Select model by sector
    if sector == "fmcg":
        model = FMCGModel(company_data, market_data)
        metrics = model.compute_fmcg_metrics()
        val_inputs = model.prepare_valuation_inputs()
    else:
        raise NotImplementedError(f"Sector model '{sector}' not yet implemented")

    # Assemble report_json
    today = date.today().isoformat()
    report_json = {
        "meta": {
            "ticker": ticker,
            "exchange": "NSE",
            "company_name": company_data.company_name,
            "sector": sector.upper(),
            "report_date": today,
            "analyst": "Institutional Equity Research",
            "currency": "INR",
            "units": "Crores",
            "fy_end_month": 3,
            "is_consolidated": company_data.is_consolidated,
        },
        "market_data": market_data,
        "financials": {
            "fy_labels": metrics.fy_labels,
            "actuals_end_idx": len(metrics.fy_labels) - 1,
            "income_statement": {
                "net_revenue":   metrics.fcf,        # placeholder — use NormalizedFinancials
                "ebitda":        [],
                "pat":           [],
                "eps":           [],
            },
            "segments": metrics.segments,
        },
        "ratios": {
            "fy_labels":      metrics.fy_labels,
            "gross_margin":   metrics.gross_margin,
            "ebitda_margin":  [],
            "roce":           metrics.roce,
            "roe":            metrics.roe,
            "dso":            metrics.dso,
            "dio":            metrics.dio,
            "dpo":            metrics.dpo,
            "ccc":            metrics.ccc,
            "fcf_conversion": metrics.fcf_conversion,
        },
        "projections": {
            "fy_labels": val_inputs.fy_labels_proj,
            "net_revenue": val_inputs.net_revenue_proj or [],
            "ebitda":  [],
            "pat":     val_inputs.pat_proj or [],
            "fcf":     val_inputs.fcf_proj or [],
        },
        "valuation": {
            "dcf": {
                "wacc_pct":            val_inputs.wacc,
                "terminal_growth_pct": val_inputs.terminal_growth,
                "net_debt_cr":         val_inputs.net_debt,
                "shares_cr":           val_inputs.shares_outstanding,
            },
            "peers": [{"ticker": t} for t in val_inputs.peer_tickers],
        },
        "charts": {},
        "narrative": {},
    }

    return report_json


def run_charts(report_json: dict, charts_dir: str) -> dict:
    """Generate all charts. Returns updated report_json with chart paths."""
    from src.charts.chart_factory import (
        ChartData,
        chart_revenue_pat_trend,
        chart_margin_progression,
        chart_roce_roe_vs_wacc,
        chart_fcf_trend,
        chart_shareholding_pie,
    )

    os.makedirs(charts_dir, exist_ok=True)
    ticker = report_json["meta"]["ticker"]
    company_name = report_json["meta"]["company_name"]
    fy_labels = report_json["financials"]["fy_labels"]
    actuals_end_idx = report_json["financials"].get("actuals_end_idx", len(fy_labels) - 1)

    meta = {"company_name": company_name, "ticker": ticker,
            "units": "INR Cr", "source": "Company filings, screener.in"}

    charts = {}

    # Revenue + PAT trend
    rev = report_json["financials"]["income_statement"].get("net_revenue", [])
    pat = report_json["financials"]["income_statement"].get("pat", [])
    if rev and pat:
        cd = ChartData(fy_labels=fy_labels, actuals_end_idx=actuals_end_idx,
                       values={"net_revenue": rev, "pat": pat}, metadata=meta)
        p = os.path.join(charts_dir, "chart_02_revenue_pat_trend.png")
        chart_revenue_pat_trend(cd, p)
        charts["chart_02"] = p
        log.info(f"  Saved chart_02")

    # Margins
    ratios = report_json.get("ratios", {})
    gm = ratios.get("gross_margin", [])
    em = ratios.get("ebitda_margin", [])
    if gm or em:
        cd = ChartData(fy_labels=fy_labels, actuals_end_idx=actuals_end_idx,
                       values={"gross_margin": gm, "ebitda_margin": em}, metadata=meta)
        p = os.path.join(charts_dir, "chart_10_margin_progression.png")
        chart_margin_progression(cd, p)
        charts["chart_10"] = p
        log.info(f"  Saved chart_10")

    report_json["charts"] = charts
    return report_json


def run_report(report_json: dict, output_path: str) -> str:
    """Assemble DOCX from report_json."""
    from src.report.report_builder import ReportBuilder
    builder = ReportBuilder(output_path=output_path)
    return builder.build(report_json)


def save_report_json(report_json: dict, output_dir: str) -> str:
    """Save report_json to disk for inspection/reuse."""
    ticker = report_json["meta"]["ticker"]
    today  = report_json["meta"]["report_date"]
    path   = os.path.join(output_dir, f"{ticker}_report_data_{today}.json")
    with open(path, "w") as f:
        json.dump(report_json, f, indent=2, default=str)
    log.info(f"Saved report_json → {path}")
    return path


def main():
    parser = argparse.ArgumentParser(description="IndiaStockResearch pipeline")
    parser.add_argument("--ticker",  required=True, help="NSE ticker e.g. ITC")
    parser.add_argument("--sector",  default="fmcg", help="Sector model to use")
    parser.add_argument("--task",    default="all",
                        choices=["data", "charts", "report", "all"])
    parser.add_argument("--force-refresh", action="store_true",
                        help="Bypass screener.in cache")
    parser.add_argument("--output", default=None,
                        help="Output DOCX path (default: reports/{TICKER}/)")
    args = parser.parse_args()

    ticker = args.ticker.upper()
    report_dir = os.path.join(REPORTS_DIR, ticker)
    charts_dir = os.path.join(report_dir, "charts")
    today = date.today().isoformat()

    output_path = args.output or os.path.join(
        report_dir, f"{ticker}_Initiation_Report_{today}.docx"
    )
    json_path = os.path.join(report_dir, f"{ticker}_report_data_{today}.json")

    os.makedirs(report_dir, exist_ok=True)

    if args.task in ("data", "all"):
        report_json = run_data(ticker, args.sector, args.force_refresh)
        save_report_json(report_json, report_dir)

    if args.task in ("charts", "all"):
        if args.task == "charts":
            # Load existing report_json from disk
            if not os.path.exists(json_path):
                log.error(f"No report_json found at {json_path}. Run --task data first.")
                sys.exit(1)
            with open(json_path) as f:
                report_json = json.load(f)
        report_json = run_charts(report_json, charts_dir)
        save_report_json(report_json, report_dir)

    if args.task in ("report", "all"):
        if args.task == "report":
            if not os.path.exists(json_path):
                log.error(f"No report_json found at {json_path}. Run --task data first.")
                sys.exit(1)
            with open(json_path) as f:
                report_json = json.load(f)
        result = run_report(report_json, output_path)
        log.info(f"Report saved → {result}")


if __name__ == "__main__":
    main()
