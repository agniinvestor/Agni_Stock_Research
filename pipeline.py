"""
Main pipeline entry point for IndiaStockResearch.

Usage:
    python pipeline.py --ticker ITC --task all
    python pipeline.py --ticker HDFCBANK --task data
    python pipeline.py --ticker TCS --task narrative
    python pipeline.py --ticker ITC --sector fmcg --task charts
    python pipeline.py --ticker ITC --force-refresh --task all

Tasks:
    data      → fetch from screener.in + yfinance, normalize, compute ratios, save report_json
    narrative → generate AI narrative sections via Claude API (requires data task first)
    charts    → generate all charts from report_json
    report    → assemble DOCX from report_json + charts
    all       → data → narrative → charts → report in sequence
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
    REPORTS_DIR, SQLITE_DB_PATH, SCREENER_USERNAME, SCREENER_PASSWORD,
    ANTHROPIC_API_KEY
)
from src.data.screener_client import ScreenerClient
from src.data.price_client import PriceClient
from src.data.company_classifier import CompanyClassifier
from src.data.normalizer import DataNormalizer
from src.utils.fy_utils import fy_label

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s"
)
log = logging.getLogger("pipeline")


# ── Data task ─────────────────────────────────────────────────────────────────

def run_data(ticker: str, sector_override: str = None,
             force_refresh: bool = False) -> dict:
    """
    Fetch company data, auto-classify sector, normalize, compute ratios.
    Returns fully populated report_json (narrative section still empty).
    """
    log.info(f"[data] Fetching {ticker}")

    screener = ScreenerClient(
        db_path=SQLITE_DB_PATH,
        username=SCREENER_USERNAME,
        password=SCREENER_PASSWORD,
    )
    price_client = PriceClient(db_path=SQLITE_DB_PATH)

    company_data = screener.fetch_company(ticker, force_refresh=force_refresh)
    market_data  = price_client.get_market_data(ticker)

    log.info(f"  Company  : {company_data.company_name}")
    log.info(f"  CMP      : ₹{market_data.get('cmp', 'N/A')}")
    log.info(f"  Mkt Cap  : ₹{market_data.get('market_cap_cr', 'N/A'):,.0f} Cr")

    # ── Auto-classify (or use manual override) ────────────────────────────────
    classifier = CompanyClassifier()
    if sector_override:
        # Build a minimal classification from the override
        classification = classifier._build_classification(
            ticker=ticker,
            company_name=company_data.company_name,
            model_type=sector_override.lower(),
            sector=sector_override.upper(),
            source="manual_override",
        )
    else:
        classification = classifier.classify(ticker, company_data)

    log.info(f"  Sector   : {classification.sector} ({classification.model_type})")
    log.info(f"  Source   : {classification.source}")

    # ── Normalize (patch Ind AS break, excise duty, segment renames) ──────────
    normalizer = DataNormalizer()
    company_data, norm_flags = normalizer.normalize(ticker, company_data)
    if norm_flags.excise_duty_adjusted:
        log.info(f"  Excise duty stripped for FYs: {norm_flags.excise_affected_fys}")
    if norm_flags.ind_as_break_detected:
        log.info(f"  Ind AS break flagged at {norm_flags.ind_as_break_fy}")
    for w in norm_flags.warnings:
        log.warning(f"  Normalizer: {w}")

    # ── Route to correct sector model ─────────────────────────────────────────
    model_type = classification.model_type
    nf = ratios = metrics = val_inputs = None

    if model_type == "fmcg":
        from src.models.fmcg_model import FMCGModel
        model   = FMCGModel(company_data, market_data)
        metrics = model.compute_fmcg_metrics()
        nf      = model._nf
        ratios  = model._ratios
        val_inputs = model.prepare_valuation_inputs()

    elif model_type == "bank":
        from src.models.bank_model import BankModel
        model   = BankModel(company_data, market_data,
                            is_nbfc=classification.is_nbfc)
        metrics = model.compute_bank_metrics()
        nf      = model._nf
        ratios  = model._ratios
        val_inputs = model.prepare_valuation_inputs()

    elif model_type == "it":
        from src.models.it_model import ITModel
        model   = ITModel(company_data, market_data)
        metrics = model.compute_it_metrics()
        nf      = model._nf
        ratios  = model._ratios
        val_inputs = model.prepare_valuation_inputs()

    elif model_type == "pharma":
        from src.models.pharma_model import PharmaModel
        model      = PharmaModel(company_data, market_data)
        metrics    = model.compute_pharma_metrics()
        nf         = model._nf
        ratios     = model._ratios
        val_inputs = model.prepare_valuation_inputs()

    elif model_type == "metals":
        from src.models.metals_model import MetalsModel
        model      = MetalsModel(company_data, market_data)
        metrics    = model.compute_metals_metrics()
        nf         = model._nf
        ratios     = model._ratios
        val_inputs = model.prepare_valuation_inputs()

    else:
        # Generic fallback — use BaseModel only
        from src.models.base_model import BaseModel
        model  = BaseModel(company_data)
        nf     = model.normalize()
        ratios = model.compute_ratios(nf)
        metrics = None
        val_inputs = None

    # ── Assemble report_json ──────────────────────────────────────────────────
    today      = date.today().isoformat()
    fy_labels  = nf.fy_labels if nf else []
    actuals_end = nf.actuals_end_fy if nf else ""

    def _list(attr):
        """Safely get list from NormalizedFinancials, return [] if None."""
        v = getattr(nf, attr, None) if nf else None
        return v if v is not None else []

    def _ratio(attr):
        v = getattr(ratios, attr, None) if ratios else None
        return v if v is not None else []

    # Build actuals_end_idx from fy_labels
    actuals_end_idx = len(fy_labels) - 1
    if actuals_end and actuals_end in fy_labels:
        actuals_end_idx = fy_labels.index(actuals_end)

    # Segments
    segments = {}
    if metrics and hasattr(metrics, "segments"):
        segments = metrics.segments or {}

    # Bank-specific or IT-specific extra fields
    bank_metrics_dict = {}
    if model_type == "bank" and metrics:
        bank_metrics_dict = {
            "nii":             getattr(metrics, "nii", []),
            "nim":             getattr(metrics, "nim", []),
            "gnpa_pct":        getattr(metrics, "gnpa_pct", []),
            "nnpa_pct":        getattr(metrics, "nnpa_pct", []),
            "pcr":             getattr(metrics, "pcr", []),
            "casa_ratio":      getattr(metrics, "casa_ratio", []),
            "advances":        getattr(metrics, "advances", []),
            "car":             getattr(metrics, "car", []),
            "cost_to_income":  getattr(metrics, "cost_to_income", []),
        }

    it_metrics_dict = {}
    if model_type == "it" and metrics:
        it_metrics_dict = {
            "employee_cost":     getattr(metrics, "employee_cost", []),
            "employee_cost_pct": getattr(metrics, "employee_cost_pct", []),
            "ebit_margin":       getattr(metrics, "ebit_margin", []),
            "headcount":         getattr(metrics, "headcount", []),
            "attrition_pct":     getattr(metrics, "attrition_pct", []),
            "vertical_mix":      getattr(metrics, "vertical_mix", {}),
            "geo_mix":           getattr(metrics, "geo_mix", {}),
        }

    pharma_metrics_dict = {}
    if model_type == "pharma" and metrics:
        pharma_metrics_dict = {
            "rnd_spend":            getattr(metrics, "rnd_spend", []),
            "rnd_spend_pct":        getattr(metrics, "rnd_spend_pct", []),
            "ebitda_ex_rnd":        getattr(metrics, "ebitda_ex_rnd", []),
            "ebitda_ex_rnd_margin": getattr(metrics, "ebitda_ex_rnd_margin", []),
            "domestic_revenue_pct": getattr(metrics, "domestic_revenue_pct", None),
            "us_revenue_pct":       getattr(metrics, "us_revenue_pct", None),
        }

    metals_metrics_dict = {}
    if model_type == "metals" and metrics:
        metals_metrics_dict = {
            "steel_volume_mt":       getattr(metrics, "steel_volume_mt", []),
            "aluminium_volume_mt":   getattr(metrics, "aluminium_volume_mt", []),
            "realization_per_tonne": getattr(metrics, "realization_per_tonne", []),
            "ebitda_per_tonne":      getattr(metrics, "ebitda_per_tonne", []),
            "raw_material_cost_pct": getattr(metrics, "raw_material_cost_pct", []),
            "net_debt_to_ebitda":    getattr(metrics, "net_debt_to_ebitda", []),
        }

    # Projections
    proj_fy_labels = []
    if val_inputs:
        proj_fy_labels = getattr(val_inputs, "fy_labels_proj", []) or []

    report_json = {
        "meta": {
            "ticker":          ticker,
            "exchange":        classification.source and "NSE",
            "company_name":    company_data.company_name,
            "sector":          classification.sector,
            "sub_sector":      classification.sub_sector,
            "model_type":      model_type,
            "report_date":     today,
            "analyst":         "Institutional Equity Research",
            "currency":        "INR",
            "units":           "Crores",
            "fy_end_month":    classification.fy_end_month,
            "is_consolidated": company_data.is_consolidated,
            "bse_code":        classification.bse_code,
            "screener_ticker": ticker,
        },
        "market_data": market_data,
        "financials": {
            "fy_labels":       fy_labels,
            "actuals_end_fy":  actuals_end,
            "actuals_end_idx": actuals_end_idx,
            "income_statement": {
                "gross_revenue":  _list("gross_revenue"),
                "excise_duty":    _list("excise_duty"),
                "net_revenue":    _list("net_revenue"),      # ← BUG FIXED
                "gross_profit":   _list("gross_profit"),
                "ebitda":         _list("ebitda"),
                "depreciation":   _list("depreciation"),
                "ebit":           _list("ebit"),
                "other_income":   _list("other_income"),
                "finance_costs":  _list("finance_costs"),
                "pbt":            _list("pbt"),
                "tax":            _list("tax"),
                "pat":            _list("pat_after_mi"),
                "eps":            _list("eps"),
                "dps":            _list("dps"),
                **bank_metrics_dict,
                **it_metrics_dict,
                **pharma_metrics_dict,
                **metals_metrics_dict,
            },
            "balance_sheet": {
                "total_assets":   _list("total_assets"),
                "fixed_assets":   _list("fixed_assets"),
                "investments":    _list("investments"),
                "inventory":      _list("inventory"),
                "trade_receivables": _list("trade_receivables"),
                "cash":           _list("cash"),
                "total_equity":   _list("total_equity"),
                "total_debt":     _list("total_debt"),
                "trade_payables": _list("trade_payables"),
                "net_debt":       _list("net_debt"),
            },
            "cash_flow": {
                "cfo":            _list("cfo"),
                "capex":          _list("capex"),
                "fcf":            _list("fcf"),
                "cfi":            _list("cfi"),
                "cff":            _list("cff"),
                "dividends_paid": _list("dividends_paid"),
            },
            "segments": segments,
        },
        "ratios": {
            "fy_labels":        fy_labels,
            "gross_margin":     _ratio("gross_margin"),
            "ebitda_margin":    _ratio("ebitda_margin"),
            "ebit_margin":      _ratio("ebit_margin"),
            "pat_margin":       _ratio("pat_margin"),
            "roce":             _ratio("roce"),
            "roe":              _ratio("roe"),
            "roa":              _ratio("roa"),
            "dso":              _ratio("dso"),
            "dio":              _ratio("dio"),
            "dpo":              _ratio("dpo"),
            "ccc":              _ratio("ccc"),
            "debt_to_equity":   _ratio("debt_to_equity"),
            "interest_coverage":_ratio("interest_coverage"),
            "revenue_growth":   _ratio("revenue_growth"),
            "ebitda_growth":    _ratio("ebitda_growth"),
            "pat_growth":       _ratio("pat_growth"),
            "fcf_conversion":   _ratio("fcf_conversion"),
            "cfo_pat_ratio":    _ratio("cfo_pat_ratio"),
            # CAGRs
            "revenue_cagr_3y":  getattr(ratios, "revenue_cagr_3y", None) if ratios else None,
            "revenue_cagr_5y":  getattr(ratios, "revenue_cagr_5y", None) if ratios else None,
            "revenue_cagr_10y": getattr(ratios, "revenue_cagr_10y", None) if ratios else None,
            "pat_cagr_3y":      getattr(ratios, "pat_cagr_3y", None) if ratios else None,
            "pat_cagr_5y":      getattr(ratios, "pat_cagr_5y", None) if ratios else None,
        },
        "shareholding": _build_shareholding(company_data),
        "projections": {
            "fy_labels":   proj_fy_labels,
            "net_revenue": getattr(val_inputs, "net_revenue_proj", []) or [] if val_inputs else [],
            "ebitda":      [],
            "pat":         getattr(val_inputs, "pat_proj", []) or [] if val_inputs else [],
            "fcf":         getattr(val_inputs, "fcf_proj", []) or [] if val_inputs else [],
            "eps":         getattr(val_inputs, "eps_proj", []) or [] if val_inputs else [],
        },
        "valuation": {
            "dcf": {
                "wacc_pct":            getattr(val_inputs, "wacc", None) if val_inputs else None,
                "terminal_growth_pct": getattr(val_inputs, "terminal_growth", None) if val_inputs else None,
                "net_debt_cr":         getattr(val_inputs, "net_debt", None) if val_inputs else None,
                "shares_cr":           getattr(val_inputs, "shares_outstanding", None) if val_inputs else None,
            },
            "peers": [{"ticker": t} for t in (
                getattr(val_inputs, "peer_tickers", []) if val_inputs else
                classification.peer_group
            )],
            "rating":         None,
            "price_target":   None,
            "cmp":            market_data.get("cmp"),
            "upside_pct":     None,
        },
        "charts": {},
        "narrative": {},
    }

    log.info(f"  FY labels: {fy_labels[:3]}...{fy_labels[-2:] if len(fy_labels) > 3 else ''}")
    if nf and nf.net_revenue:
        non_null = [v for v in nf.net_revenue if v is not None]
        if non_null:
            log.info(f"  Net Rev (latest): ₹{non_null[-1]:,.0f} Cr")

    return report_json


def _build_shareholding(company_data) -> dict:
    """Extract shareholding from CompanyData.shareholding DataFrame."""
    sh = {}
    try:
        df = company_data.shareholding
        if df is not None and not df.empty:
            quarters = list(df.columns)
            sh["quarters"] = quarters
            for idx in df.index:
                key = str(idx).lower().replace(" ", "_").replace("%", "pct")
                sh[key] = [float(v) if v == v else None for v in df.loc[idx]]
    except Exception:
        pass
    return sh


# ── Narrative task ─────────────────────────────────────────────────────────────

def run_narrative(report_json: dict) -> dict:
    """Generate AI narrative sections via Claude API."""
    ticker = report_json["meta"]["ticker"]
    log.info(f"[narrative] Generating narrative for {ticker}")

    if not ANTHROPIC_API_KEY:
        log.warning("ANTHROPIC_API_KEY not set — skipping narrative generation")
        return report_json

    try:
        from src.ai.narrative_generator import NarrativeGenerator
        generator = NarrativeGenerator(db_path=SQLITE_DB_PATH, api_key=ANTHROPIC_API_KEY)
        report_json = generator.generate_all(report_json)
        log.info(f"  Narrative sections generated: {list(report_json.get('narrative', {}).keys())}")
    except ImportError:
        log.warning("narrative_generator not available — skipping")
    except Exception as e:
        log.error(f"Narrative generation failed: {e}")

    return report_json


# ── Charts task ────────────────────────────────────────────────────────────────

def run_charts(report_json: dict, charts_dir: str) -> dict:
    """Generate charts from report_json. Returns updated report_json with chart paths."""
    from src.charts.chart_factory import (
        ChartData,
        chart_revenue_pat_trend,
        chart_margin_progression,
        chart_fcf_trend,
        chart_roce_roe_vs_wacc,
        chart_shareholding_pie,
        chart_dcf_sensitivity_heatmap,
        chart_valuation_football_field,
        chart_working_capital_days,
        chart_capex_vs_net_cash,
        chart_revenue_by_segment,
        chart_scenario_comparison,
    )

    os.makedirs(charts_dir, exist_ok=True)
    ticker       = report_json["meta"]["ticker"]
    company_name = report_json["meta"]["company_name"]
    model_type   = report_json["meta"].get("model_type", "generic")
    fy_labels    = report_json["financials"]["fy_labels"]
    actuals_end_idx = report_json["financials"].get("actuals_end_idx", len(fy_labels) - 1)

    meta = {
        "company_name": company_name, "ticker": ticker,
        "units": "INR Cr", "source": "Company filings, screener.in"
    }
    is_  = report_json["financials"]["income_statement"]
    bs_  = report_json["financials"]["balance_sheet"]
    cf_  = report_json["financials"]["cash_flow"]
    rat_ = report_json.get("ratios", {})
    charts = {}

    def _save(key, fn, cd, fname):
        try:
            p = os.path.join(charts_dir, fname)
            fn(cd, p)
            charts[key] = p
            log.info(f"  ✓ {key}: {fname}")
        except Exception as e:
            log.warning(f"  ✗ {key} failed: {e}")

    rev = is_.get("net_revenue", [])
    pat = is_.get("pat", [])

    # chart_02: Revenue + PAT trend
    if rev or pat:
        _save("chart_02", chart_revenue_pat_trend,
              ChartData(fy_labels, actuals_end_idx,
                        {"net_revenue": rev, "pat": pat}, meta),
              "chart_02_revenue_pat_trend.png")

    # chart_10: Margin progression
    gm = rat_.get("gross_margin", [])
    em = rat_.get("ebitda_margin", [])
    pm = rat_.get("pat_margin", [])
    if any([gm, em, pm]):
        _save("chart_10", chart_margin_progression,
              ChartData(fy_labels, actuals_end_idx,
                        {"gross_margin": gm, "ebitda_margin": em, "pat_margin": pm}, meta),
              "chart_10_margin_progression.png")

    # chart_12: FCF trend
    cfo   = cf_.get("cfo", [])
    capex = cf_.get("capex", [])
    fcf   = cf_.get("fcf", [])
    if any([cfo, capex, fcf]):
        _save("chart_12", chart_fcf_trend,
              ChartData(fy_labels, actuals_end_idx,
                        {"cfo": cfo, "capex": capex, "fcf": fcf}, meta),
              "chart_12_fcf_trend.png")

    # chart_26: ROCE / ROE vs WACC
    roce = rat_.get("roce", [])
    roe  = rat_.get("roe", [])
    wacc_val = report_json["valuation"]["dcf"].get("wacc_pct")
    wacc_list = [wacc_val] * len(fy_labels) if wacc_val else []
    if roce or roe:
        _save("chart_26", chart_roce_roe_vs_wacc,
              ChartData(fy_labels, actuals_end_idx,
                        {"roce": roce, "roe": roe, "wacc": wacc_list}, meta),
              "chart_26_roce_roe_wacc.png")

    # chart_25: Working capital days (FMCG / IT)
    if model_type in ("fmcg", "it"):
        dso = rat_.get("dso", [])
        dio = rat_.get("dio", [])
        dpo = rat_.get("dpo", [])
        ccc = rat_.get("ccc", [])
        if any([dso, dio, dpo]):
            _save("chart_25", chart_working_capital_days,
                  ChartData(fy_labels, actuals_end_idx,
                            {"dso": dso, "dio": dio, "dpo": dpo, "ccc": ccc}, meta),
                  "chart_25_working_capital.png")

    # chart_23: Capex vs net cash
    net_debt = bs_.get("net_debt", [])
    net_cash = [-v if v is not None else None for v in net_debt]
    if capex or net_cash:
        _save("chart_23", chart_capex_vs_net_cash,
              ChartData(fy_labels, actuals_end_idx,
                        {"capex": capex, "net_cash": net_cash}, meta),
              "chart_23_capex_net_cash.png")

    # chart_03: Revenue by segment
    segs = report_json["financials"].get("segments", {})
    if segs:
        seg_rev = {k: v.get("revenue", []) for k, v in segs.items() if "revenue" in v}
        if seg_rev:
            _save("chart_03", chart_revenue_by_segment,
                  ChartData(fy_labels, actuals_end_idx, seg_rev, meta),
                  "chart_03_revenue_by_segment.png")

    # chart_24: Shareholding pie
    sh = report_json.get("shareholding", {})
    if sh:
        holders = {k: v[-1] for k, v in sh.items()
                   if k not in ("quarters",) and isinstance(v, list) and v}
        if holders:
            _save("chart_24", chart_shareholding_pie,
                  ChartData([], -1, holders, meta),
                  "chart_24_shareholding_pie.png")

    report_json["charts"] = charts
    log.info(f"[charts] {len(charts)} charts generated")
    return report_json


# ── Report task ────────────────────────────────────────────────────────────────

def run_report(report_json: dict, output_path: str) -> str:
    """Assemble DOCX from report_json."""
    from src.report.report_builder import ReportBuilder
    builder = ReportBuilder(output_path=output_path)
    return builder.build(report_json)


# ── Validate task ─────────────────────────────────────────────────────────────

def run_validate(report_json: dict):
    """Run data quality checks on report_json. Logs all results."""
    from src.validation.model_validator import ModelValidator
    from src.validation.screener_validator import ScreenerValidator
    from src.data.price_client import PriceClient
    price_client = PriceClient(db_path=SQLITE_DB_PATH)
    report = ModelValidator.validate(report_json)
    report = ScreenerValidator.validate(report_json, price_client, existing_report=report)
    return report


# ── Earnings report task ───────────────────────────────────────────────────────

def run_earnings_report(report_json: dict, output_path: str,
                        quarterly_data: dict = None) -> str:
    """Generate an earnings update DOCX from report_json + quarterly actuals."""
    from src.report.earnings_builder import EarningsBuilder
    if quarterly_data:
        report_json = dict(report_json)
        report_json["quarterly"] = quarterly_data
    builder = EarningsBuilder(output_path=output_path)
    return builder.build(report_json)


# ── Backtest task ──────────────────────────────────────────────────────────────

def run_backtest(ticker: str = None, print_scorecard: bool = True):
    """Update price target outcomes and print accuracy scorecard."""
    from src.backtest.price_target_tracker import PriceTargetTracker
    from src.data.price_client import PriceClient
    price_client = PriceClient(db_path=SQLITE_DB_PATH)
    tracker = PriceTargetTracker(db_path=SQLITE_DB_PATH)
    updated = tracker.update_outcomes(ticker=ticker, price_client=price_client)
    log.info(f"[backtest] Updated {updated} outcomes")
    if print_scorecard:
        tracker.print_scorecard(ticker=ticker)
    return tracker


# ── Utilities ──────────────────────────────────────────────────────────────────

def save_report_json(report_json: dict, output_dir: str) -> str:
    """Save report_json to disk for inspection/reuse."""
    ticker = report_json["meta"]["ticker"]
    today  = report_json["meta"]["report_date"]
    path   = os.path.join(output_dir, f"{ticker}_report_data_{today}.json")
    os.makedirs(output_dir, exist_ok=True)
    with open(path, "w") as f:
        json.dump(report_json, f, indent=2, default=str)
    log.info(f"Saved report_json → {path}")
    return path


def load_report_json(json_path: str) -> dict:
    if not os.path.exists(json_path):
        log.error(f"report_json not found: {json_path}. Run --task data first.")
        sys.exit(1)
    with open(json_path) as f:
        return json.load(f)


# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="IndiaStockResearch pipeline")
    parser.add_argument("--ticker",  required=True, help="NSE ticker e.g. ITC")
    parser.add_argument("--sector",  default=None,
                        help="Override auto-classification: fmcg|bank|it|generic")
    parser.add_argument("--task",    default="all",
                        choices=["data", "narrative", "charts", "report",
                                 "validate", "backtest", "all"])
    parser.add_argument("--report-type", default="initiation",
                        choices=["initiation", "earnings"],
                        help="Report format: initiation (30-50p) or earnings (8-12p)")
    parser.add_argument("--force-refresh", action="store_true",
                        help="Bypass screener.in cache")
    parser.add_argument("--output", default=None,
                        help="Output DOCX path (default: reports/{TICKER}/)")
    parser.add_argument("--skip-narrative", action="store_true",
                        help="Skip narrative generation even in --task all")
    args = parser.parse_args()

    ticker = args.ticker.upper()
    report_dir = os.path.join(REPORTS_DIR, ticker)
    charts_dir = os.path.join(report_dir, "charts")
    today      = date.today().isoformat()

    output_path = args.output or os.path.join(
        report_dir, f"{ticker}_Initiation_Report_{today}.docx"
    )
    json_path = os.path.join(report_dir, f"{ticker}_report_data_{today}.json")

    os.makedirs(report_dir, exist_ok=True)

    report_json = None

    if args.task in ("data", "all"):
        report_json = run_data(ticker, args.sector, args.force_refresh)
        save_report_json(report_json, report_dir)

    if args.task in ("narrative", "all") and not args.skip_narrative:
        if report_json is None:
            report_json = load_report_json(json_path)
        report_json = run_narrative(report_json)
        save_report_json(report_json, report_dir)

    if args.task in ("charts", "all"):
        if report_json is None:
            report_json = load_report_json(json_path)
        report_json = run_charts(report_json, charts_dir)
        save_report_json(report_json, report_dir)

    if args.task in ("validate",):
        if report_json is None:
            report_json = load_report_json(json_path)
        validation = run_validate(report_json)
        if validation.has_errors:
            log.error(f"Validation FAILED: {len(validation.errors)} error(s) — see above")
            sys.exit(1)
        else:
            log.info("Validation passed ✓")

    if args.task in ("backtest",):
        run_backtest(ticker=ticker)

    if args.task in ("report", "all"):
        if report_json is None:
            report_json = load_report_json(json_path)
        report_type = getattr(args, "report_type", "initiation")
        if report_type == "earnings":
            earnings_path = os.path.join(
                report_dir, f"{ticker}_Earnings_Update_{today}.docx"
            )
            result = run_earnings_report(report_json, earnings_path)
            log.info(f"[earnings] Saved → {result}")
        else:
            result = run_report(report_json, output_path)
            log.info(f"[report] Saved → {result}")


if __name__ == "__main__":
    main()
