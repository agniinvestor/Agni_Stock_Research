# IndiaStockResearch

An institutional-quality Indian equity research platform that auto-generates sell-side research reports (30–50 pages, 10,000+ words) for any NSE/BSE-listed company. The platform covers FMCG, Banking/NBFC, and IT sectors with sector-specific financial models, deterministic data pipelines, and Claude-powered narrative generation.

---

## Table of Contents

1. [Overview](#overview)
2. [Quick Start](#quick-start)
3. [Architecture](#architecture)
4. [Report JSON Contract](#report-json-contract)
5. [Sector Coverage](#sector-coverage)
6. [Module Reference](#module-reference)
7. [Sample Output — ITC Initiation Report](#sample-output--itc-initiation-report)
8. [Key Design Decisions](#key-design-decisions)
9. [Configuration](#configuration)
10. [Dependencies](#dependencies)
11. [Known Limitations & Roadmap](#known-limitations--roadmap)
12. [File Structure](#file-structure)

---

## Overview

IndiaStockResearch automates the end-to-end workflow of an equity research analyst:

1. **Fetch** — scrapes screener.in for 10 years of financials; pulls live price data via yfinance
2. **Normalize** — handles Ind AS transitions, pre-GST excise duty gross-ups, and segment name drift
3. **Classify** — auto-detects sector (FMCG / Banking / IT) from universe CSV, screener metadata, or company name rules
4. **Model** — runs a sector-specific financial model (ratio engine, WACC/DCF, valuation comps)
5. **Narrate** — calls Claude API for 7 structured narrative sections (investment summary, thesis, risks, etc.)
6. **Chart** — generates 35 publication-quality PNG charts at 300 DPI
7. **Report** — assembles a 30–50 page DOCX with all charts, tables, and narrative

All intermediate state is serialized through a single `report_json` dict, making every stage independently testable and rerunnable.

---

## Quick Start

### 1. Environment Setup

```bash
cd /mnt/windows-ubuntu/IndiaStockResearch
python3 -m venv ~/IndiaStockResearch_venv
source ~/IndiaStockResearch_venv/bin/activate
pip install -r requirements.txt
```

### 2. Credentials

```bash
export SCREENER_USERNAME=your@email.com
export SCREENER_PASSWORD=yourpassword
export ANTHROPIC_API_KEY=sk-ant-...
```

### 3. Run the Pipeline

```bash
# Full pipeline — fetch, model, narrate, chart, assemble report
python pipeline.py --ticker ITC --task all
python pipeline.py --ticker HDFCBANK --task all
python pipeline.py --ticker TCS --task all

# Run individual stages
python pipeline.py --ticker ITC --task data       # fetch + normalize + model (~30s, no LLM)
python pipeline.py --ticker ITC --task narrative  # generate AI narrative (~2 min, ~$0.11)
python pipeline.py --ticker ITC --task charts     # render all charts (~20s)
python pipeline.py --ticker ITC --task report     # assemble DOCX (~10s)

# Override auto-classification
python pipeline.py --ticker NESTLEIND --sector fmcg --task all

# Skip narrative (faster, no API cost)
python pipeline.py --ticker ITC --skip-narrative --task all

# Force-refresh screener.in cache
python pipeline.py --ticker ITC --force-refresh --task data
```

Output reports are written to `reports/<TICKER>/`.

---

## Architecture

```
screener.in HTML
       |
       v
ScreenerClient.fetch_company(ticker)      [SQLite cached — 7 day TTL]
       |
       v
DataNormalizer.normalize()                [Ind AS breaks, excise strip, segment fixes]
       |
       v
CompanyClassifier.classify()              [3-tier: universe CSV -> screener meta -> name rules]
       |
       v
FMCGModel | BankModel | ITModel
  .compute_xxx_metrics()  -> XxxMetrics
  .prepare_valuation_inputs() -> XxxValuationInputs
       |
       v
report_json                               [central JSON data contract — single source of truth]
       |
       v
NarrativeGenerator.generate_all()        [7 Claude API calls, SHA256-keyed, cached 30 days]
       |
       v
ChartFactory.chart_*()                   [35 charts -> PNG @ 300 DPI]
       |
       v
ReportBuilder.build()                    [DOCX: 30-50 pages, 10,000+ words]
```

### Caching Strategy

All network I/O is backed by a single SQLite database at `data/processed/cache.db`:

| Cache | TTL | Contents |
|-------|-----|----------|
| Screener data | 7 days | Company financials, ratios, metadata |
| Price data | 1 day | OHLCV history, market cap |
| AI narrative | 30 days | 7 narrative sections (SHA256-keyed by input) |
| PDF filing | 90 days | Extracted text from BSE annual reports |

---

## Report JSON Contract

The `report_json` dict is the interface between all pipeline modules. No module calls another module directly — all read from and write to this contract.

```json
{
  "meta": {
    "ticker": "ITC",
    "company_name": "ITC Limited",
    "sector": "FMCG",
    "model_type": "fmcg",
    "report_date": "2026-03-18",
    "analyst": "IndiaStockResearch"
  },
  "market_data": {
    "cmp": 306.80,
    "market_cap_cr": 385000,
    "52w_high": 528.50,
    "52w_low": 295.00,
    "price_target": 380.0,
    "rating": "BUY",
    "upside_pct": 23.8
  },
  "financials": {
    "fy_labels": ["FY16", "FY17", "FY18", "FY19", "FY20", "FY21", "FY22", "FY23", "FY24", "FY25"],
    "income_statement": {
      "net_revenue": [...],
      "ebitda": [...],
      "pat": [...]
    },
    "balance_sheet": {
      "total_assets": [...],
      "net_debt": [...]
    },
    "cash_flow": {
      "cfo": [...],
      "capex": [...],
      "fcf": [...]
    },
    "segments": {
      "Cigarettes": { "revenue": [...], "ebit": [...] },
      "FMCG Others": { "revenue": [...], "ebit": [...] }
    }
  },
  "ratios": {
    "gross_margin": [...],
    "ebitda_margin": [...],
    "roce": [...],
    "roe": [...],
    "fcf_yield": [...]
  },
  "projections": {
    "fy_labels": ["FY26E", "FY27E", "FY28E"],
    "net_revenue": [...],
    "ebitda": [...],
    "pat": [...]
  },
  "valuation": {
    "dcf": { "wacc_pct": 11.59, "terminal_growth_pct": 5.0, "intrinsic_value": 365.0 },
    "ev_ebitda": { "target_multiple": 28.0, "implied_value": 380.0 },
    "peers": [...]
  },
  "charts": {
    "chart_02": "reports/ITC/charts/chart_02_revenue_trend.png"
  },
  "narrative": {
    "investment_summary": [...],
    "business_description": [...],
    "financial_analysis": [...],
    "valuation": [...],
    "risks": { "key_risks": [...], "mitigants": [...] },
    "esg": [...],
    "outlook": [...]
  }
}
```

---

## Sector Coverage

| Sector | Model | Key Metrics | Valuation Methods | Universe |
|--------|-------|-------------|-------------------|----------|
| FMCG | `FMCGModel` | Gross margin, A&P spend %, volume growth, working capital days | EV/EBITDA, P/E, DCF (WACC) | ITC, HINDUNILVR, NESTLEIND, BRITANNIA, DABUR, MARICO, GODREJCP, TATACONSUM, COLPAL, EMAMILTD |
| Banking / NBFC | `BankModel` | NII, NIM, GNPA %, NNPA %, PCR, CASA ratio, CAR (Basel III) | P/B, P/ABV, Gordon Growth Model | HDFCBANK, ICICIBANK, KOTAKBANK, AXISBANK, SBIN, INDUSINDBK, BAJFINANCE |
| IT Services | `ITModel` | EBIT margin, employee cost %, headcount, revenue / employee, FCF conversion | P/E, EV/EBIT, DCF | TCS, INFY, WIPRO, HCLTECH, TECHM, LTIM |

---

## Module Reference

### Data Layer

**`src/data/screener_client.py`** (615 lines)
Scrapes screener.in with session-based authentication. Implements exponential backoff on HTTP 429 responses and stores all responses in SQLite with a 7-day TTL. Returns structured dicts covering the income statement, balance sheet, cash flow, and shareholding pattern for up to 10 fiscal years.

**`src/data/price_client.py`** (403 lines)
yfinance wrapper that resolves NSE (`.NS`) and BSE (`.BO`) ticker suffixes automatically. Fetches OHLCV history, calculates 52-week high/low, derives market cap, and returns beta estimates. Results are cached for 1 day.

**`src/data/company_classifier.py`** (407 lines)
3-tier auto-classification pipeline:
1. Looks up the ticker in `config/stock_universe.csv`
2. Inspects screener.in metadata (industry, sector fields)
3. Falls back to name-based rules (e.g., "Bank", "Finance" → BANK; "Tech", "Info" → IT)

**`src/data/normalizer.py`** (433 lines)
Handles three Indian-market data quality issues:
- **Ind AS break detection** — flags the FY year where revenue recognition changed materially
- **Pre-GST excise duty gross-up** — strips excise from revenue in FY17 and earlier to make series comparable
- **Segment name continuity** — maps historical segment label variants to canonical names

**`src/data/pdf_parser.py`** (618 lines)
Downloads annual report PDFs from BSE filing APIs, extracts text using PyMuPDF, and sends selected sections to Claude for structured summarization. Results are cached for 90 days. Gracefully degrades if PyMuPDF is not installed.

---

### Models

**`src/models/base_model.py`** (827 lines)
Core ratio engine used by all sector models. Defines:
- `NormalizedFinancials` — 44-field dataclass covering income statement, balance sheet, and cash flow
- `ComputedRatios` — 25+ ratios including ROCE, ROE, FCFE yield, asset turnover, and working capital days
- CAGR computation utilities for any time series

**`src/models/fmcg_model.py`** (653 lines)
FMCG-specific model extending the base. Computes gross margin decomposition, advertising & promotion spend as percentage of revenue, segment-level EBIT bridges, WACC (using CAPM), and a 3-year projection engine with volume/price/mix assumptions.

**`src/models/bank_model.py`** (750 lines)
Banking/NBFC model. Computes NII, PPOP (pre-provision operating profit), NIM, GNPA %, NNPA %, provision coverage ratio, CASA ratio, and Capital Adequacy Ratio. Valuation via P/B (price-to-book), P/ABV (adjusted book value), and Gordon Growth Model for sustainable RoE.

**`src/models/it_model.py`** (611 lines)
IT services model. Computes EBIT margin, employee cost as percentage of revenue, revenue per employee, headcount trends, and FCF conversion (FCF / PAT). Valuation via P/E, EV/EBIT, and DCF using USD revenue assumptions for hedging-adjusted cash flows.

---

### Charts

**`src/charts/chart_style.py`** (370 lines)
Defines the shared visual identity: NAVY / GREEN / GOLD color palette, matplotlib rcParams defaults, and the `ChartData` dataclass used as a typed input contract for all chart functions.

**`src/charts/chart_factory.py`** (1,758 lines)
35 typed chart functions covering:
- Revenue and profit trends (bar + line combo)
- Margin evolution (EBITDA, PAT, gross margin)
- Free cash flow and FCF yield
- Return ratios (ROCE, ROE, ROIC)
- Valuation bands (EV/EBITDA, P/E historical)
- Shareholding pattern (pie + area over time)
- Competitive benchmarking (radar, grouped bar)
- Segment revenue and EBIT waterfall (FMCG)
- NPA and provisioning trends (Banking)
- Employee metrics and attrition (IT)

All charts output PNG at 300 DPI, sized for A4 DOCX insertion.

---

### Report

**`src/report/docx_helpers.py`** (345 lines)
DRY DOCX primitives to avoid repetition in the report builder:
- `section_bar()` — colored section header bar
- `make_fin_table()` — financial table with alternating row shading and INR formatting
- `make_rating_box()` — BUY/SELL/HOLD rating box with price target and upside
- `make_key_metrics_box()` — 2-column metrics summary box for cover page

**`src/report/report_builder.py`** (919 lines)
JSON-driven DOCX assembler with 11 section builders:
1. Cover page (rating box, key metrics, analyst details)
2. Investment summary (bullet points from narrative)
3. Investment thesis (3–5 pillars with supporting data)
4. Business description (company overview, segment map)
5. Industry overview (market size, competitive dynamics)
6. Financial analysis (10-year historical tables + charts)
7. Projections (3-year forward estimates with assumptions)
8. Valuation (DCF bridge, comps table, sensitivity matrix)
9. Risk factors (key risks + mitigants)
10. ESG overview
11. Appendices (full financial model, ratio table, glossary)

---

### Utilities

**`src/utils/fy_utils.py`** (99 lines)
Indian fiscal year utilities (April–March): FY label generation, quarter identification from calendar date, and FY date range helpers.

**`src/ai/narrative_generator.py`** (674 lines)
7-section Claude API pipeline. Each section sends a structured prompt with relevant financial data and receives a strict JSON response schema. All calls are SHA256-keyed on the input data and cached in SQLite for 30 days, so re-running a report never incurs redundant API costs.

Sections generated:
1. Investment summary (2–3 sentence headline + 5 bullet points)
2. Investment thesis (3–5 thesis pillars with evidence)
3. Business description (company history, segment overview)
4. Financial analysis (10-year narrative with key inflection points)
5. Valuation analysis (method justification, peer comparison)
6. Risk factors (key risks with probability/impact ratings, mitigants)
7. Outlook (12–18 month forward view)

---

## Sample Output — ITC Initiation Report

The first report generated by the platform, produced on 2026-03-18:

| Attribute | Value |
|-----------|-------|
| File | `reports/ITC/ITC_Initiation_Report_2026-03-18.docx` |
| Length | 30+ pages, 10,745 words |
| Charts | 35 |
| Tables | 69 |
| Rating | **BUY** |
| Price Target | ₹380 |
| CMP | ₹306.80 |
| Upside | **23.8%** |

The ITC report was built using 6 sequential scripts (`build_report_s1.py` through `build_report_s6.py`) before the generic pipeline existed. These scripts remain in `reports/ITC/` as reference implementations.

---

## Key Design Decisions

### 1. JSON Data Contract as Single Source of Truth
All modules read from and write to `report_json`. There is no direct module-to-module coupling — the data contract is the API. This makes every stage independently testable, re-runnable from any point, and trivially serializable to disk for debugging.

### 2. SQLite for All Caching
A single `data/processed/cache.db` file backs all caching tiers. This avoids Redis or filesystem scatter, works identically on any OS, and can be inspected with any SQLite browser. TTLs are per-table and enforced at read time.

### 3. LLM for Judgment Only
Data fetching, financial model computation, chart rendering, and DOCX assembly are fully deterministic and require no API calls. Claude is used only to write narrative text — the one task where judgment and language quality matter. This keeps costs low (~$0.11 per full report) and makes the non-narrative pipeline free to run repeatedly.

### 4. Structured JSON Output from Claude
All 7 narrative sections return strict JSON schemas, not free-form text. This makes narrative output machine-readable, allows individual sections to be re-generated independently, and prevents the LLM from hallucinating structural elements like table headers or chart references.

### 5. Graceful Degradation
The pipeline continues with available data if:
- screener.in returns errors (uses cached data or skeleton)
- Anthropic API is unavailable (skips narrative, produces data-only report)
- PyMuPDF is not installed (skips annual report PDF parsing)

### 6. Sector-Specific Models
Banks are valued on P/B and Gordon Growth, not EV/EBITDA. IT companies are benchmarked on EV/EBIT and P/E, not gross margin. FMCG companies use working capital days and A&P intensity as primary operating metrics. The platform enforces this separation — no single generic model handles all sectors.

### 7. Indian Market Specifics
- Fiscal year is April–March (FY25 = April 2024 – March 2025)
- All currency in INR Crores (1 Crore = 10 million)
- Ind AS transition handling for the FY17/FY18 break
- Pre-GST excise duty gross-up applied to revenue for FY17 and earlier
- BSE filing API used for annual report PDF access

---

## Configuration

### `config/settings.py`
Central constants file. Key settings:

```python
CACHE_DB_PATH = "data/processed/cache.db"
REPORTS_DIR = "reports/"
RAW_DATA_DIR = "data/raw/"

SCREENER_BASE_URL = "https://www.screener.in"
SCREENER_CACHE_TTL_DAYS = 7
NARRATIVE_CACHE_TTL_DAYS = 30
PDF_CACHE_TTL_DAYS = 90

# Chart palette
NAVY  = "#1B2A4A"
GREEN = "#2E7D32"
GOLD  = "#F9A825"
```

### `config/stock_universe.csv`
23 tickers with NSE symbol, BSE code, company name, and sector label.

| Column | Description |
|--------|-------------|
| `ticker` | NSE symbol (e.g., ITC) |
| `bse_code` | BSE scrip code (e.g., 500875) |
| `company_name` | Full legal name |
| `sector` | `fmcg` / `banking` / `it` |

---

## Dependencies

```
# Scraping & data
requests
beautifulsoup4
lxml
yfinance
pandas
numpy
openpyxl

# Charts
matplotlib
seaborn

# Report generation
python-docx

# AI narrative (Phase 2)
anthropic

# PDF parsing (Phase 2, optional)
PyMuPDF

# Caching
# sqlite3 is Python stdlib — no install needed
```

Install all dependencies:

```bash
pip install -r requirements.txt
```

---

## Known Limitations & Roadmap

### Current Limitations

| # | Limitation | Impact |
|---|-----------|--------|
| 1 | No live test without screener.in credentials | First run requires valid account |
| 2 | Pharma, Metals, Infra sectors not modeled | Pipeline will reject tickers outside FMCG / Banking / IT |
| 3 | Concall transcript parsing not implemented | Earnings call insights require manual input |
| 4 | Projection assumptions are static defaults | Forward estimates do not adapt to business context automatically |
| 5 | No backtesting framework | Price target accuracy cannot be tracked systematically |

### Phase 3 Roadmap

- [ ] Data quality validation layer — ratio sanity checks, outlier flagging, cross-source reconciliation
- [ ] LLM-driven projection assumptions — derive revenue growth and margin assumptions from business context and management guidance
- [ ] Pharma sector model — revenue breakdown by therapy area, ANDA pipeline, US FDA risk scoring
- [ ] Metals & Mining model — commodity price sensitivity, volume/realization decomposition
- [ ] Backtesting framework — track price targets vs. actual outcomes over 12-month horizon
- [ ] Concall transcript scraper — parse BSE/NSE transcript PDFs, extract forward guidance
- [ ] Web UI — Flask/FastAPI interface for non-CLI users

---

## File Structure

```
IndiaStockResearch/
|
+-- pipeline.py                        # CLI entry point
+-- requirements.txt
|
+-- config/
|   +-- settings.py                    # All constants, paths, color palette
|   +-- stock_universe.csv             # 23 tickers with BSE codes and sectors
|
+-- src/
|   +-- ai/
|   |   +-- narrative_generator.py     # 7-section Claude API pipeline (674 lines)
|   |
|   +-- charts/
|   |   +-- chart_factory.py           # 35 typed chart functions (1,758 lines)
|   |   +-- chart_style.py             # Shared palette + rcParams (370 lines)
|   |
|   +-- data/
|   |   +-- company_classifier.py      # 3-tier sector auto-detection (407 lines)
|   |   +-- normalizer.py              # Ind AS + excise + segment fixes (433 lines)
|   |   +-- pdf_parser.py              # BSE filing download + PyMuPDF extraction (618 lines)
|   |   +-- price_client.py            # yfinance wrapper for NSE/BSE (403 lines)
|   |   +-- screener_client.py         # screener.in scraper + SQLite cache (615 lines)
|   |
|   +-- models/
|   |   +-- bank_model.py              # Banking/NBFC sector model (750 lines)
|   |   +-- base_model.py              # Ratio engine + NormalizedFinancials (827 lines)
|   |   +-- fmcg_model.py              # FMCG sector model (653 lines)
|   |   +-- it_model.py                # IT Services sector model (611 lines)
|   |
|   +-- report/
|   |   +-- docx_helpers.py            # DOCX primitives (345 lines)
|   |   +-- report_builder.py          # JSON-driven DOCX assembler (919 lines)
|   |
|   +-- utils/
|       +-- fy_utils.py                # Indian FY utilities (99 lines)
|
+-- data/
|   +-- raw/                           # Raw screener.in HTML snapshots
|   +-- processed/
|       +-- cache.db                   # SQLite: screener, price, narrative, PDF caches
|
+-- reports/
    +-- ITC/
        +-- ITC_Initiation_Report_2026-03-18.docx   # 30+ pages, 10,745 words
        +-- ITC_Financial_Model_2026-03-12.xlsx
        +-- ITC_Research_Document_2026-03-12.md
        +-- ITC_Valuation_Analysis_2026-03-12.md
        +-- ITC_Charts_2026-03-12.zip
        +-- build_report_s1.py         # Cover + Investment Summary
        +-- build_report_s2.py         # Investment Thesis + Risks
        +-- build_report_s3.py         # Company 101
        +-- build_report_s4.py         # Products + Industry
        +-- build_report_s5.py         # Financial Analysis
        +-- build_report_s6.py         # Valuation + Appendices
```

---

*Built with Python 3.11+ — Powered by screener.in, yfinance, Claude API, and python-docx.*
