# IndiaStockResearch — Implementation Plan (Updated 2026-03-18)

## Goal
Build a pipeline that generates full-fledged institutional-quality equity research reports for any Indian NSE/BSE-listed stock. Target: 30–50 page DOCX reports matching JPMorgan/Goldman Sachs quality, generated in under 5 minutes.

---

## ✅ Phase 1 — Data Pipeline & Report Engine (COMPLETE)

**Status**: Committed to `main` at `bdf72cb`. All 19 files, 6,341 lines.

### Data Layer
- [x] `src/data/screener_client.py` — screener.in HTML scraper with session auth, SQLite cache (7-day TTL), exponential backoff on rate limits
- [x] `src/data/price_client.py` — yfinance wrapper for NSE (.NS) / BSE (.BO): price history, market cap, 52w H/L, beta

### Models
- [x] `src/models/base_model.py` — ratio engine: NormalizedFinancials (44 fields), ComputedRatios (25+ ratios including ROCE, ROE, working capital days, CAGRs, DuPont)
- [x] `src/models/fmcg_model.py` — FMCG sector: gross margin reconstruction, A&P spend, segment EBIT, WACC (CAPM), simple projection engine

### Charts
- [x] `src/charts/chart_style.py` — shared palette (NAVY/GREEN/GOLD), rcParams, ChartData dataclass
- [x] `src/charts/chart_factory.py` — 35 typed chart functions: revenue, margins, FCF, valuation (DCF heatmap, football field), shareholding, competitive analysis

### Report
- [x] `src/report/docx_helpers.py` — DRY DOCX primitives: section_bar, make_fin_table, make_rating_box, make_key_metrics_box, add_chart
- [x] `src/report/report_builder.py` — JSON-driven DOCX assembler with 11 section builders (cover → disclosures)

### Foundation
- [x] `src/utils/fy_utils.py` — Indian FY utilities (Apr–Mar FY, quarter labels, FY ranges)
- [x] `config/settings.py` — all constants, paths, color palette, screener + Anthropic config
- [x] `config/stock_universe.csv` — 23 tickers (10 FMCG + 7 Banking + 6 IT) with BSE codes
- [x] `pipeline.py` — CLI entry point
- [x] `requirements.txt`

---

## ✅ Phase 2 — Intelligence Layer (COMPLETE)

**Status**: Committed to `main` at `9b2697a`. 11 files changed, 3,964 lines added.

### Intelligence
- [x] `src/data/company_classifier.py` — 3-tier auto-classification: universe CSV lookup → screener metadata parsing → company name rules
- [x] `src/data/normalizer.py` — Ind AS break detection (FY17), excise duty gross-up strip (pre-FY18/GST), segment name continuity mapping
- [x] `src/data/pdf_parser.py` — BSE annual report PDF download, PyMuPDF section extraction, Claude Haiku summarization into structured JSON

### Sector Models
- [x] `src/models/bank_model.py` — NII, NIM, GNPA%, NNPA%, PCR, CAR, CASA ratio, P/B + Gordon Growth valuation
- [x] `src/models/it_model.py` — EBIT margin, employee cost%, headcount, attrition, FCF conversion, P/E + EV/EBIT valuation

### AI Narrative
- [x] `src/ai/narrative_generator.py` — 7-section Claude API pipeline (investment_summary, business_overview, investment_pillars, risks, financial_analysis, catalysts, industry_overview) with SHA256-keyed SQLite cache (30-day TTL)

### Pipeline Overhaul
- [x] `pipeline.py` — fixed net_revenue bug; added auto-classify + normalizer + narrative steps; full 3-sector routing (fmcg/bank/it/generic); `--skip-narrative` flag; `--task narrative` option
- [x] `config/stock_universe.csv` — extended to 23 tickers with `bse_code` column
- [x] `config/settings.py` — ANTHROPIC_API_KEY, NARRATIVE_MODEL, BSE filing URLs, PDF_CACHE_DIR

---

## 🚧 Phase 3 — Quality & Coverage Expansion (PLANNED)

### Sprint 1: Live Validation (prerequisite: screener.in credentials)
- [ ] Run `python pipeline.py --ticker ITC --task all` end-to-end
- [ ] Validate ITC financials match known values (FY25A Net Rev = ₹43,011 Cr, PAT = ₹22,592 Cr)
- [ ] Validate HDFCBANK classification → bank model → bank metrics
- [ ] Validate TCS classification → IT model → IT metrics
- [ ] Fix any parsing issues in screener_client for real HTML

### Sprint 2: Data Quality Validator
- [ ] `src/validation/model_validator.py` — ratio sanity checks:
  - Balance sheet balance: Total Assets = Total Liabilities + Equity (within 1%)
  - Cash flow tie-out: Opening cash + Net cash change = Closing cash
  - Revenue continuity: No >50% YoY swing without annotation
  - Margin reasonableness: EBITDA margin 0–80%, PAT margin 0–50%
- [ ] `src/validation/screener_validator.py` — cross-check key metrics vs yfinance info

### Sprint 3: Pharma Sector Model
- [ ] `src/models/pharma_model.py` — R&D capitalization, ANDA pipeline, US FDA risk, US generics pricing
  - Key metrics: R&D spend %, Domestic branded vs US generics mix, EBITDA margin ex-R&D
  - Peers: SUNPHARMA, DRREDDY, CIPLA, DIVISLAB, LUPIN, AUROPHARMA
  - Valuation: P/E (primary), EV/EBITDA (secondary), Sum-of-the-parts
- [ ] Add pharma tickers to `stock_universe.csv`

### Sprint 4: Metals & Commodities Model
- [ ] `src/models/metals_model.py` — volume × price model, LME/coking coal sensitivity
  - Key metrics: EBITDA per tonne, Volume growth, Realization per tonne
  - Peers: TATASTEEL, JSWSTEEL, HINDALCO, VEDL, NATIONALUM
  - Valuation: EV/EBITDA, EV/tonne

### Sprint 5: Earnings Update Reports
- [ ] `src/report/earnings_builder.py` — 8–12 page format for post-earnings updates
  - Beat/miss analysis vs estimates
  - Updated model + revised price target
  - Chart delta vs prior quarter
- [ ] `pipeline.py --report-type earnings` option

### Sprint 6: Backtesting Framework
- [ ] `src/backtest/price_target_tracker.py` — store price targets with issue date
  - Track actual price 3/6/12 months later
  - Compute hit rate: % of BUYs that outperformed Nifty in 12 months
  - Track MAPE of earnings estimates

---

## Architecture (Current)

```
screener.in HTML
       ↓
ScreenerClient.fetch_company(ticker)    [SQLite cached 7 days]
       ↓
DataNormalizer.normalize()              [Ind AS, excise, segment fixes]
       ↓
CompanyClassifier.classify()            [auto-detect: FMCG/BANK/IT]
       ↓
FMCGModel / BankModel / ITModel
  .compute_xxx_metrics()  → XxxMetrics
  .prepare_valuation_inputs() → XxxValuationInputs
       ↓
report_json (JSON data contract — source of truth for all downstream)
       ↓
NarrativeGenerator.generate_all()      [7 Claude API calls, cached 30 days]
       ↓
ChartFactory.chart_*()                  [35 charts → PNG @ 300 DPI]
       ↓
ReportBuilder.build()                   [DOCX: 30–50 pages, 10K+ words]
```

---

## JSON Data Contract

The central `report_json` dict is the interface between all modules:

```json
{
  "meta": {
    "ticker": "ITC", "company_name": "ITC Limited",
    "sector": "FMCG", "model_type": "fmcg",
    "report_date": "2026-03-18", "bse_code": "500875"
  },
  "market_data": {
    "cmp": 307.0, "market_cap_cr": 387153,
    "52w_high": 444.0, "52w_low": 272.0
  },
  "financials": {
    "fy_labels": ["FY16", "FY17", ..., "FY25"],
    "actuals_end_fy": "FY25",
    "actuals_end_idx": 9,
    "income_statement": {
      "net_revenue": [...], "ebitda": [...], "pat": [...], "eps": [...]
    },
    "balance_sheet": {
      "total_assets": [...], "total_debt": [...], "net_debt": [...]
    },
    "cash_flow": { "cfo": [...], "capex": [...], "fcf": [...] },
    "segments": { "Cigarettes": {"revenue": [...], "ebit": [...]}, ... }
  },
  "ratios": {
    "fy_labels": [...],
    "gross_margin": [...], "ebitda_margin": [...],
    "roce": [...], "roe": [...],
    "dso": [...], "dio": [...], "dpo": [...], "ccc": [...]
  },
  "projections": {
    "fy_labels": ["FY26E", "FY27E", "FY28E"],
    "net_revenue": [...], "pat": [...], "fcf": [...]
  },
  "valuation": {
    "dcf": {"wacc_pct": 11.59, "terminal_growth_pct": 6.0},
    "rating": "BUY", "price_target": 380, "upside_pct": 23.8
  },
  "charts": { "chart_02": "reports/ITC/charts/chart_02.png", ... },
  "narrative": {
    "investment_summary": [...],
    "investment_pillars": [...],
    "risks": {...}, "catalysts": [...]
  }
}
```

---

## Tech Stack

| Layer | Technology | Purpose |
|---|---|---|
| Data fetching | requests + BeautifulSoup | screener.in HTML scraping |
| Price data | yfinance | NSE/BSE price history + market data |
| Processing | pandas + numpy | DataFrame manipulation |
| Charts | matplotlib | 35 chart types at 300 DPI |
| Report | python-docx | DOCX generation |
| AI | anthropic (Claude) | Narrative generation + PDF summarization |
| PDF parsing | PyMuPDF (fitz) | Annual report text extraction |
| Caching | SQLite (stdlib) | All data + narrative caching |
| CLI | argparse (stdlib) | Pipeline orchestration |

---

## ITC Initiation Report (Completed 2026-03-12 to 2026-03-18)

First output produced before the generic pipeline existed. Built using 6 company-specific scripts.

- **File**: `reports/ITC/ITC_Initiation_Report_2026-03-18.docx`
- **Rating**: BUY | **Price Target**: ₹380 | **CMP**: ₹306.80 | **Upside**: 23.8%
- **Stats**: 30+ pages, 10,745 words, 35 charts embedded, 69 tables
- **Task breakdown**: Research (Task 1) → Financial Model (Task 2) → Valuation (Task 3) → Charts (Task 4) → Report Assembly (Task 5)
- **Scripts**: `build_report_s1.py` through `build_report_s6.py` (all saved in `reports/ITC/`)

---

## Quick Reference

```bash
# Run full pipeline for any ticker
python pipeline.py --ticker ITC --task all
python pipeline.py --ticker HDFCBANK --task all   # auto-detects bank model
python pipeline.py --ticker TCS --task all        # auto-detects IT model

# Individual tasks
python pipeline.py --ticker ITC --task data       # fetch + model only (~30s, no API cost)
python pipeline.py --ticker ITC --task narrative  # AI narrative only
python pipeline.py --ticker ITC --task charts     # charts only
python pipeline.py --ticker ITC --task report     # DOCX only

# Flags
--sector fmcg|bank|it|generic   # override auto-classification
--skip-narrative                # skip Claude API calls
--force-refresh                 # bypass screener.in cache

# Required env vars
export SCREENER_USERNAME=your@email.com
export SCREENER_PASSWORD=yourpassword
export ANTHROPIC_API_KEY=sk-ant-...
```

---

## Key Gotchas

1. **screener.in is unofficial** — cache aggressively; never re-fetch unnecessarily
2. **Ind AS break at FY17** — balance sheet not comparable pre/post; DataNormalizer handles this
3. **Pre-GST revenue (pre-FY18)** — excise duty inflates gross revenue; DataNormalizer strips it
4. **Indian FY is April–March** — use `fy_utils.fy_label()` everywhere; never hardcode year logic
5. **All amounts in INR Crores** — normalize before storing in report_json
6. **NSE vs BSE tickers** — yfinance needs .NS suffix for NSE; some tickers differ between exchanges
7. **Consolidated vs Standalone** — always use consolidated; some small caps only have standalone on screener.in
8. **NestleInd FY ends December** — fy_end_month=12 in stock_universe.csv; affects projection labels
9. **sudo required** — writing to `/mnt/windows-ubuntu/` requires `echo "6875" | sudo -S ...`
10. **Venv** — always use `~/IndiaStockResearch_venv/bin/python3` for running code
