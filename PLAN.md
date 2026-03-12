# India Stock Fundamental Research Report Generator — Implementation Plan

## Goal
Build a pipeline that generates full-fledged sell-side style fundamental research reports for Indian NSE/BSE listed stocks.

---

## Report Sections (Target Output)

| Section | Key Data Points |
|---|---|
| Cover Page | Company name, CMP, target price, rating, date |
| Investment Summary | Bull/Bear thesis, key catalysts, risk summary |
| Business Overview | Segments, revenue mix, competitive positioning, moat |
| Industry & Macro | Sector tailwinds/headwinds, regulatory landscape |
| Financial Statements | P&L, Balance Sheet, Cash Flow (5-10yr history) |
| Key Ratios | ROE, ROCE, EBITDA margin, PAT margin, D/E, Current Ratio |
| Valuation | P/E, EV/EBITDA, P/B, DCF model, peer comparison |
| Shareholding Pattern | Promoter %, FII %, DII %, Public %, quarterly trend |
| Management Quality | Promoter pledging, salary/dividend ratio, RPTs |
| Corporate Governance | Audit qualifications, contingent liabilities, ESOP |
| Price Chart | 1yr/3yr chart, volume, moving averages |
| Projections | 2-3yr forward estimates (revenue, EBITDA, PAT, EPS) |
| Risk Factors | Business, regulatory, macro, execution risks |

---

## Data Sources

| Source | Type | Provides | Rating |
|---|---|---|---|
| screener.in JSON API | Free (unofficial) | 10yr P&L, BS, CF, ratios, shareholding | ★★★★★ |
| yfinance (.NS/.BO) | Free | Price, basic financials, market data | ★★★☆☆ |
| BSE API + XBRL | Free (official) | Quarterly filings, shareholding, corporate actions | ★★★★☆ |
| Trendlyne API | Paid | Full financials, ratios, analyst estimates | ★★★★★ |
| Tijori Finance | Paid | Segment data, concall transcripts, subsidiaries | ★★★★★ |
| Refinitiv / Bloomberg | Enterprise | Everything, institutional grade | ★★★★★ |

**Phase 1 (Free)**: screener.in + yfinance + BSE XBRL
**Phase 2 (Paid)**: Add Trendlyne + Tijori Finance

---

## Tech Stack

```
Data Fetching     → screener.in API, yfinance, BSE API
Data Processing   → pandas, numpy
Charting          → Plotly (export PNG via kaleido), matplotlib
AI Narrative      → Claude API (claude-sonnet) via anthropic SDK
Report Templates  → Jinja2 HTML → WeasyPrint PDF
Annual Reports    → pdfplumber / PyMuPDF → Claude summarization
XBRL Parsing      → arelle-release
Excel Export      → openpyxl (optional)
Caching           → SQLite
```

---

## Python Packages

```txt
# Data
yfinance>=0.2.40
requests>=2.31.0
beautifulsoup4>=4.12.0
lxml>=5.0.0

# Processing
pandas>=2.1.0
numpy>=1.26.0

# Charts
plotly>=5.18.0
kaleido>=0.2.1
matplotlib>=3.8.0
mplfinance>=0.12.10b0

# PDF Generation
Jinja2>=3.1.2
WeasyPrint>=62.0

# AI
anthropic>=0.28.0

# Annual Report Parsing
pdfplumber>=0.11.0
PyMuPDF>=1.24.0

# XBRL (BSE filings)
arelle-release>=2.3.0

# Optional
openpyxl>=3.1.2
```

---

## Project Structure

```
IndiaStockResearch/
├── data/
│   ├── raw/                    # Raw fetched data (JSON, XBRL, PDFs)
│   └── processed/              # Cleaned, normalized DataFrames (SQLite cache)
├── reports/
│   ├── output/                 # Generated PDF/HTML reports
│   └── templates/              # Jinja2 HTML templates
│       ├── report_base.html
│       ├── cover_page.html
│       ├── executive_summary.html
│       ├── financials_table.html
│       ├── ratio_section.html
│       ├── charts_section.html
│       ├── shareholding.html
│       ├── valuation.html
│       └── risks.html
├── src/
│   ├── data/
│   │   ├── screener_client.py  # screener.in API: auth, rate-limiting, caching
│   │   ├── bse_client.py       # BSE API: XBRL filings, shareholding, corp actions
│   │   └── price_client.py     # yfinance wrapper for NSE/BSE price data
│   ├── ratios/
│   │   └── calculator.py       # Ratio engine: ROE, ROCE, DuPont, Z-score, etc.
│   ├── charts/
│   │   └── builder.py          # Plotly chart functions (revenue, margins, debt)
│   ├── ai/
│   │   └── narrative_generator.py  # Claude API: section-by-section narrative
│   ├── pdf/
│   │   └── renderer.py         # Jinja2 render → WeasyPrint PDF assembly
│   └── utils/
│       └── annual_report.py    # pdfplumber + Claude for MD&A summarization
├── config/
│   ├── settings.py             # API keys, paths, constants
│   └── stock_universe.csv      # List of stocks to cover
├── notebooks/                  # Exploratory Jupyter notebooks
├── PLAN.md                     # This file
└── requirements.txt
```

---

## Build Phases

### Phase 1 — Data Pipeline
- [ ] `screener_client.py`: screener.in JSON API with session auth, rate limiting, SQLite cache
- [ ] `price_client.py`: yfinance wrapper (.NS suffix), current price, 52-week data, market cap
- [ ] `bse_client.py`: BSE filing downloader, XBRL shareholding parser

### Phase 2 — Processing & Ratios
- [ ] `calculator.py`: Full ratio engine
  - Profitability: ROE, ROCE, EBITDA margin, PAT margin, Asset Turnover
  - Leverage: D/E ratio, Interest Coverage, Debt/EBITDA
  - Liquidity: Current Ratio, Quick Ratio, Cash Conversion Cycle
  - Growth: Revenue CAGR, PAT CAGR, EPS CAGR (3yr, 5yr, 10yr)
  - Quality: DuPont decomposition, Altman Z-score, CFO/PAT ratio
  - Valuation: P/E, EV/EBITDA, P/B, P/S, Dividend Yield

### Phase 3 — Charts
- [ ] `builder.py`: Reusable Plotly chart functions
  - Revenue & PAT trend (bar chart)
  - Margin trends (line chart)
  - Debt & coverage (combo chart)
  - Shareholding pattern (pie + quarterly trend)
  - Price chart with MA (candlestick/line)
  - Peer comparison (horizontal bar)

### Phase 4 — Report Templates
- [ ] Design HTML/CSS base template (professional, print-ready)
- [ ] Section templates for each report block
- [ ] Font embedding (e.g., Inter or Source Sans Pro)
- [ ] Page headers, footers, page numbers

### Phase 5 — AI Narrative (Claude API)
- [ ] `narrative_generator.py`: Claude API integration
  - Business overview generation from structured data
  - Financial analysis narrative (margin trends, growth drivers)
  - Risk factors generation (sector + company-specific)
  - Investment thesis (bull/bear case)
  - Peer comparison commentary
- [ ] Prompt templates per report section

### Phase 6 — PDF Assembly
- [ ] `renderer.py`: Jinja2 → HTML → WeasyPrint → PDF
- [ ] Chart PNG embedding
- [ ] Cover page with CMP and rating

### Phase 7 — Annual Report Processing
- [ ] `annual_report.py`: PDF download from BSE → pdfplumber text extraction
- [ ] Claude summarization of MD&A, Chairman's letter, Notes to Accounts
- [ ] Extraction of segment revenue, RPTs, contingent liabilities

---

## Key Gotchas

1. **screener.in is unofficial** — cache all data in SQLite; never re-fetch what you have
2. **NSE website is hostile to scraping** — use BSE as primary regulatory source
3. **Always use Consolidated financials** — not Standalone, for holding/group companies
4. **Indian FY is April–March** — align all time-series accordingly
5. **All amounts in INR Crores** — normalize carefully if mixing sources
6. **Shareholding data lags ~45 days** — quarterly filings have a delay
7. **Small/mid cap data quality is poor** — validate before rendering

---

## Critical Files to Build First

1. `src/data/screener_client.py` — primary financial data source, must be robust
2. `src/ratios/calculator.py` — computational core of the report
3. `src/ai/narrative_generator.py` — Claude API integration for analyst narrative
4. `reports/templates/report_base.html` — sets the quality bar for output
5. `src/data/bse_client.py` — official regulatory data layer
