# India Stock Research Report — Full Research & Thinking

## 1. What a Full Sell-Side Research Report Contains (Indian Context)

Before picking tools, it is worth anchoring on the deliverable. A typical sell-side or independent fundamental research report for an Indian listed company contains:

| Section | Key Data Points |
|---|---|
| Cover Page | Company name, CMP, target price, rating (Buy/Hold/Sell), analyst name, date |
| Investment Summary | Bull/Bear thesis, key catalysts, risk summary |
| Business Overview | Business segments, revenue mix, competitive positioning, moat analysis |
| Industry & Macro Context | Sector tailwinds/headwinds, regulatory landscape |
| Financial Statements (5-10 yr history) | P&L, Balance Sheet, Cash Flow Statement |
| Key Ratios | ROE, ROCE, EBITDA margin, PAT margin, D/E, Current Ratio, Asset Turnover |
| Valuation | P/E, EV/EBITDA, P/B, DCF model, peer comparison table |
| Shareholding Pattern | Promoter %, FII %, DII %, Public %, quarterly trend |
| Management Quality | Promoter pledging, salary/dividend ratio, related-party transactions |
| Corporate Governance | Audit qualifications, contingent liabilities, ESOP dilution |
| Technical Charts (optional) | 1yr/3yr price chart, volume, moving averages |
| Projections | 2-3 year forward estimates (revenue, EBITDA, PAT, EPS) |
| Risk Factors | Business, regulatory, macro, execution risks |

---

## 2. Free Data Sources for Indian Stocks

### 2.1 yfinance
- **What it provides**: OHLCV price data, basic financials (P&L, balance sheet, cash flow), key statistics, dividends, splits, info dict (sector, industry, market cap, beta, PE, etc.)
- **Indian market support**: Yes, using `.NS` suffix for NSE (e.g., `RELIANCE.NS`) and `.BO` suffix for BSE (e.g., `RELIANCE.BO`). NSE data is generally more reliable.
- **Financial statements**: Provides annual and quarterly P&L, balance sheet, cash flow via `ticker.financials`, `ticker.balance_sheet`, `ticker.cashflow`. These pull from Yahoo Finance's database.
- **Limitations for India**:
  - Data sourced from Yahoo Finance, which often has gaps or inconsistencies in Indian company financials
  - Shareholding pattern NOT available
  - Segment-wise revenue NOT available
  - Subsidiary / consolidation details inconsistent
  - Rate limits are informal and undocumented; heavy use gets IP-blocked
  - Historical financial data depth: typically 4 years only
  - No concall transcripts, annual report text, or management commentary
- **Best use**: Price data, basic ratios, quick screening. Not reliable as sole financial data source.

### 2.2 nsepy
- **What it provides**: NSE-specific price/volume data, derivatives (futures/options) data, index data, corporate actions, VIX
- **Current status**: The library was last actively maintained around 2021-2022. NSE changed its website architecture and added Cloudflare bot protection in 2021-2022, which broke most scraping-based approaches. nsepy is effectively broken/unmaintained as of 2024.
- **What still works**: Some historical OHLCV data fetches still work intermittently, but reliability is poor.
- **Verdict**: Do NOT build a production system on nsepy. Use it only as a reference for understanding NSE data structures.

### 2.3 jugaad-data
- **What it provides**: NSE and BSE historical price data, index data, futures/options data. It is the spiritual successor to nsepy and was actively maintained through 2023.
- **Indian market support**: Strong NSE/BSE focus
- **Limitations**: Primarily a price/derivatives data library; does NOT provide fundamental financial statements. NSE website changes continue to be a challenge.
- **Verdict**: Better than nsepy for price data, but still not a fundamentals source.

### 2.4 Screener.in (Web Scraping / Unofficial)
- **What it provides**: This is arguably the richest free source of Indian fundamental data. Screener.in aggregates 10+ years of financial data from annual reports and quarterly filings for all NSE/BSE-listed companies.
  - Consolidated and standalone P&L, balance sheet, cash flow
  - Quarterly results going back 10+ years
  - Key ratios (computed)
  - Shareholding pattern (quarterly)
  - Peer comparison
  - Promoter pledging data
  - Custom screener queries
- **API access**: Screener.in has an unofficial/undocumented JSON API. For a company like Infosys, `https://www.screener.in/api/company/INFY/` returns structured JSON data including financials, ratios, and shareholding. This is widely used in the Python community.
- **Python library**: No official library, but the community uses `requests` + `BeautifulSoup` or the JSON API directly. There is a library called `screenerpy` (unofficial, use with caution) and several GitHub projects wrapping the API.
- **Limitations**:
  - Terms of service technically prohibit scraping at scale
  - Rate limiting enforced; authentication required for some endpoints
  - Session/cookie handling needed for the JSON API
  - No programmatic bulk export
- **Verdict**: Best free source for fundamental data depth. Use the JSON API with authentication and rate limiting.

### 2.5 BSE India Official Data
- **BSE Corporate Filings**: BSE provides an official API/portal (`api.bseindia.com`) for corporate filings (results, board meetings, shareholding pattern filings in XBRL format).
- **What is available free**:
  - Quarterly financial results (XBRL/PDF filings)
  - Shareholding pattern filings (structured XML/XBRL)
  - Corporate actions (dividends, splits, bonuses)
  - Annual report PDFs
  - Price data via `api.bseindia.com`
- **Python access**: Direct HTTP requests to BSE endpoints. The BSE API is partially documented. Key endpoints: `https://api.bseindia.com/BseIndiaAPI/api/` prefix.
- **XBRL parsing**: XBRL filings can be parsed using the `arelle` Python library or `xbrl` package to extract structured financial data from filings.
- **Limitations**: Unofficial / undocumented endpoints may change without notice. Rate limiting. Some data only in PDF form.

### 2.6 NSE India Official Data
- **NSE has heavily restricted programmatic access** since 2021-2022. Their website uses Cloudflare, session tokens, and rotating headers.
- **What remains accessible**: NSE provides some data via `nseindia.com` but requires session management (cookie extraction via browser simulation).
- **nseindia-api** (unofficial Python package): Handles the session/cookie dance for NSE. Available on GitHub (zerodha/nsetools successor projects).
- **nsetools**: Another library, but like nsepy, affected by NSE website changes.
- **Limitations**: Fragile, dependent on NSE not changing its frontend.

### 2.7 Zerodha Kite Connect (Free tier available)
- **What it provides**: Live and historical OHLCV data (market data), order management
- **For fundamentals**: Does NOT provide financial statements. Only price/volume.
- **Cost**: Free tier limited; paid plans for higher data frequency.

### 2.8 SEBI XBRL / MCA Data
- MCA21 (Ministry of Corporate Affairs) provides structured financial data for all Indian companies (including listed) in XBRL format.
- This is the upstream source for much of what screener.in aggregates.
- **Python access**: Direct download of XBRL ZIP files from MCA website, parse with `arelle`.
- **Limitation**: Complex to set up, data lags by months, not quarterly frequency.

---

## 3. Paid / Freemium APIs

### 3.1 Tickertape (freemium)
- India-focused; provides screener, financials, ratios, peer comparison
- No official public API as of 2024; data accessible via browser but not documented for programmatic use
- Some developers scrape it but TOS prohibits this
- **Verdict**: Good for human use; not suitable for automated pipeline

### 3.2 Trendlyne (freemium/paid)
- India-focused, strong on screener, DVM score, forecasts, shareholding
- Has a paid API offering (`trendlyne.com/developer/`) that provides structured financial data
- Pricing: subscription-based, reasonable for individual developers
- **What it provides via API**: Financial statements, ratios, shareholding, analyst estimates, price targets
- **Verdict**: One of the best India-specific paid API options. Evaluate for production use.

### 3.3 Refinitiv (LSEG) Eikon / Workspace
- Enterprise-grade, used by institutional investors
- Provides deep financial history, analyst estimates, ownership data, news
- **Python**: `refinitiv-data` library (formerly `eikon` Python API)
- **Cost**: Very expensive (enterprise contracts, $10K+/year range)
- **Indian coverage**: Good but not always as deep as domestic providers for smaller companies
- **Verdict**: Overkill for most projects unless institutional

### 3.4 Alpha Vantage
- Provides fundamental data (P&L, balance sheet, cash flow) via REST API
- **Indian market support**: Limited. Primarily US markets. Indian stocks are not well-covered.
- **Verdict**: Not suitable for Indian stock fundamentals

### 3.5 Quandl / Nasdaq Data Link
- Limited Indian fundamentals coverage
- Has some NSE price data via community-contributed datasets
- **Verdict**: Not a primary source for Indian fundamentals

### 3.6 Intrinio
- US-focused; poor Indian coverage
- **Verdict**: Skip for Indian markets

### 3.7 Wisesheets / Stockedge API
- **Stockedge**: India-focused app; has a developer API in beta. Provides financials, ratios, shareholding, news.
- Worth evaluating for a paid tier option.

### 3.8 Tijori Finance API
- India-specific; provides segment-wise revenue, subsidiary analysis, concall data, management commentary
- Niche but excellent for deep fundamental data not available elsewhere
- **Verdict**: Best-in-class for qualitative/segment data if available via API

### 3.9 BSE Data APIs (Official Paid)
- BSE offers official paid data feeds (end-of-day, tick data, corporate actions)
- Contact BSE directly; pricing varies

### 3.10 Upstox API / Angel Broking SmartAPI
- Free broker APIs providing live price data, historical OHLCV
- Do NOT provide financial statement data
- **Verdict**: Use only for price data component

---

## 4. Financial Data Sources Summary Table

| Library/Source | Price Data | P&L | Balance Sheet | Cash Flow | Ratios | Shareholding | Segment Data | Status |
|---|---|---|---|---|---|---|---|---|
| yfinance (.NS/.BO) | Good | Partial | Partial | Partial | Basic | No | No | Active |
| nsepy | Poor | No | No | No | No | No | No | Broken |
| jugaad-data | Good | No | No | No | No | No | No | Mostly active |
| screener.in API | No | Yes (10yr) | Yes (10yr) | Yes (10yr) | Yes | Yes | No | Active (unofficial) |
| BSE API (free) | Basic | Via XBRL | Via XBRL | Via XBRL | No | Yes (filings) | No | Active |
| Trendlyne API (paid) | Yes | Yes | Yes | Yes | Yes | Yes | Partial | Active |
| Tijori (paid) | No | Yes | Yes | Yes | Yes | Yes | Yes | Active |
| Refinitiv (paid) | Yes | Yes | Yes | Yes | Yes | Yes | Partial | Active |

---

## 5. Report Generation Stack

### 5.1 Data Processing
- **pandas**: Core for all financial data manipulation, ratio calculation, time-series alignment
- **numpy**: Financial calculations (CAGR, IRR, DCF computations)
- **openpyxl / xlsxwriter**: If delivering an Excel-based report alongside PDF

### 5.2 Charting and Visualization
- **Plotly**: Best for interactive HTML reports. Supports candlestick charts, bar charts for revenue/profit trends, waterfall charts for bridge analysis. Can export to static PNG via `kaleido`.
- **matplotlib / seaborn**: For static charts in PDF reports. More control over styling.
- **mplfinance**: Specialized for financial charts (OHLCV candlestick, volume bars).
- **Recommendation**: Use Plotly for chart generation, export to PNG via `plotly.io.write_image()` with the `kaleido` backend, then embed PNGs in PDF.

### 5.3 PDF Report Generation — Three Options

**Option A: Jinja2 + WeasyPrint (Recommended)**
- Design report as an HTML template using Jinja2
- Render HTML to PDF using WeasyPrint (pure Python, CSS-based layout)
- Pros: CSS-based design is familiar, easy to create professional layouts, supports charts as embedded images, great for tabular data
- Cons: WeasyPrint can be slow for large reports; complex multi-column layouts require careful CSS
- Fonts: Use Google Fonts or bundle fonts (important for professional output)

**Option B: ReportLab (Programmatic PDF)**
- Full programmatic control over every element on the page
- Supports tables, charts, custom layouts
- Pros: Very precise control, handles complex layouts, good for template reuse
- Cons: Verbose code; designing layouts is time-consuming compared to HTML/CSS approach
- Use `reportlab.platypus` (PLATYPUS framework) for document flow

**Option C: Jinja2 + HTML + Playwright/Puppeteer**
- Render HTML to PDF using headless Chrome (via `playwright` Python library)
- Best CSS/layout fidelity since it uses a real browser engine
- Pros: Pixel-perfect output, supports complex CSS, charts render perfectly
- Cons: Requires Chrome/Chromium installation, heavier dependency

**Recommendation**: Use Jinja2 + WeasyPrint for most cases. Fall back to Playwright if layout quality is paramount.

### 5.4 Template Structure
```
reports/templates/
  report_base.html          # Master layout, CSS, header/footer
  cover_page.html           # Extends base
  executive_summary.html
  financials_table.html     # Reusable for P&L, BS, CF
  ratio_section.html
  charts_section.html
  shareholding.html
  valuation.html
  risks.html
```

---

## 6. AI/LLM Layer (Claude API)

### 6.1 Use Cases for Claude in the Report
- **Narrative generation**: Given structured financial data (revenue CAGR, margin trends, debt levels), generate business analysis paragraphs
- **Ratio interpretation**: "ROCE has improved from 14% to 22% over 3 years; explain the likely drivers and significance"
- **Risk identification**: Given sector and financial profile, identify standard and company-specific risks
- **Peer comparison commentary**: Summarize how the company compares to peers across metrics
- **Management commentary summarization**: Summarize earnings call transcripts or annual report MD&A sections
- **Investment thesis generation**: Given bull/bear data points, draft an investment thesis paragraph

### 6.2 Implementation Pattern
- Use `claude-sonnet` model via the Anthropic Python SDK (`anthropic` package)
- Pass structured data as context in the prompt (JSON financials, ratio tables)
- Use system prompts to define the analyst persona and report style
- Use structured output / tool use to ensure Claude returns data in parseable format when needed
- For long annual reports, use Claude's 200K context window to process full PDF text

### 6.3 Annual Report PDF Processing
- Extract text from annual report PDFs using `pdfplumber` or `PyMuPDF` (fitz)
- Pass MD&A, Chairman's letter, Notes to Accounts sections to Claude for summarization
- Extract specific data points (related party transactions, contingent liabilities, segment revenue) that are only in PDF form

### 6.4 Concall Transcript Analysis
- Sources for transcripts: Motilal Oswal, Edelweiss websites publish free transcripts; Tijori Finance aggregates them
- Use Claude to extract: management guidance, key risks mentioned, capital allocation commentary

---

## 7. Full Pipeline Architecture

```
Data Layer
├── screener.in JSON API  → 10yr financials, ratios, shareholding (primary)
├── yfinance (.NS)        → Current price, market cap, 52-week data
├── BSE API               → Latest quarterly filings, shareholding XBRL
└── Annual Report PDFs    → pdfplumber + Claude for qualitative data

Processing Layer (pandas/numpy)
├── Financial statement normalization
├── Ratio calculation engine (custom, verified against screener)
├── CAGR/growth calculations
├── DCF model
└── Peer comparison aggregation

Visualization Layer
├── Plotly → Revenue/PAT trend charts, margin charts, debt charts
├── matplotlib → Shareholding pie charts, ratio comparison bars
└── Chart export → kaleido PNG for PDF embedding

AI Layer (Claude API)
├── Business description generation
├── Financial analysis narrative
├── Risk factors generation
└── Annual report / concall summarization

Report Generation Layer
├── Jinja2 HTML template rendering
├── WeasyPrint PDF conversion
└── Optional: Excel model export (openpyxl)
```

---

## 8. Python Package Requirements

```txt
# Data Acquisition
yfinance>=0.2.40
requests>=2.31.0
beautifulsoup4>=4.12.0
lxml>=5.0.0

# Financial Calculations
pandas>=2.1.0
numpy>=1.26.0

# PDF Text Extraction (for annual reports)
pdfplumber>=0.11.0
PyMuPDF>=1.24.0  # fitz

# XBRL Parsing (for BSE filings)
arelle-release>=2.3.0

# Charting
plotly>=5.18.0
kaleido>=0.2.1
matplotlib>=3.8.0
mplfinance>=0.12.10b0

# Report Generation
Jinja2>=3.1.2
WeasyPrint>=62.0
# OR: playwright>=1.40.0 (for Chrome-based PDF)

# AI Layer
anthropic>=0.28.0

# Excel output (optional)
openpyxl>=3.1.2
xlsxwriter>=3.1.9
```

---

## 9. Recommended Combination by Use Case

### Solo Developer / Indie Project (Low Cost)
- **Fundamentals**: screener.in JSON API (free, requires account, respect rate limits)
- **Price Data**: yfinance (.NS suffix)
- **Shareholding**: BSE XBRL filings (free)
- **Annual Reports**: Download PDFs from BSE filings portal, process with pdfplumber + Claude
- **Report**: Jinja2 + WeasyPrint
- **AI**: Claude API (Sonnet model for cost efficiency)
- **Estimated cost**: Claude API costs only (a few dollars per report batch)

### Small Team / Startup Product (Medium Cost)
- **Fundamentals**: Trendlyne paid API (structured, reliable, India-specific)
- **Segment data**: Tijori Finance API
- **Price Data**: Zerodha Kite Connect or Upstox API
- **Concall transcripts**: Tijori or direct web sourcing
- **Report**: Jinja2 + Playwright (higher quality output)
- **AI**: Claude API with larger context for full document analysis

### Institutional / High Volume
- **Fundamentals**: Refinitiv LSEG or Bloomberg (via blpapi Python SDK)
- **Custom NLP pipeline**: Fine-tuned models for Indian financial terminology
- **Report**: Fully custom ReportLab templates matching institutional brand standards

---

## 10. Key Limitations & Gotchas

1. **screener.in is not an official API**: It may break without notice. Always implement caching (store fetched data in SQLite or PostgreSQL) so you are not re-fetching constantly and to survive outages.

2. **NSE website is hostile to automation**: Do not rely on direct NSE scraping. Use BSE as the primary regulatory filing source.

3. **Data quality for smaller companies**: For mid/small cap stocks below Rs. 500 crore market cap, data quality drops significantly across all free sources. Some companies file only in PDF (not XBRL).

4. **Quarterly vs Annual data**: yfinance quarterly data is unreliable for Indian companies. Use screener.in or BSE filings for quarterly P&L.

5. **Consolidated vs Standalone**: Always prefer consolidated financials for holding companies. screener.in provides both; make sure to specify.

6. **Currency**: All Indian financial data is in INR (Crores/Lakhs). Some APIs report in thousands or millions. Normalize carefully.

7. **Financial year**: Indian companies follow April-March FY. Align time-series accordingly.

8. **Shareholding lag**: Shareholding patterns are filed quarterly (June, September, December, March). There is always a ~45-day lag.

---

## 11. Build Phases (Detailed)

### Phase 1 — Data Pipeline
- `screener_client.py`: screener.in JSON API with session auth, rate limiting, SQLite cache
- `price_client.py`: yfinance wrapper (.NS suffix), current price, 52-week data, market cap
- `bse_client.py`: BSE filing downloader, XBRL shareholding parser

### Phase 2 — Processing & Ratios
- `calculator.py`: Full ratio engine
  - Profitability: ROE, ROCE, EBITDA margin, PAT margin, Asset Turnover
  - Leverage: D/E ratio, Interest Coverage, Debt/EBITDA
  - Liquidity: Current Ratio, Quick Ratio, Cash Conversion Cycle
  - Growth: Revenue CAGR, PAT CAGR, EPS CAGR (3yr, 5yr, 10yr)
  - Quality: DuPont decomposition, Altman Z-score, CFO/PAT ratio
  - Valuation: P/E, EV/EBITDA, P/B, P/S, Dividend Yield

### Phase 3 — Charts
- `builder.py`: Reusable Plotly chart functions
  - Revenue & PAT trend (bar chart)
  - Margin trends (line chart)
  - Debt & coverage (combo chart)
  - Shareholding pattern (pie + quarterly trend)
  - Price chart with MA (candlestick/line)
  - Peer comparison (horizontal bar)

### Phase 4 — Report Templates
- Design HTML/CSS base template (professional, print-ready)
- Section templates for each report block
- Font embedding (e.g., Inter or Source Sans Pro)
- Page headers, footers, page numbers

### Phase 5 — AI Narrative (Claude API)
- `narrative_generator.py`: Claude API integration
  - Business overview generation from structured data
  - Financial analysis narrative (margin trends, growth drivers)
  - Risk factors generation (sector + company-specific)
  - Investment thesis (bull/bear case)
  - Peer comparison commentary
- Prompt templates per report section

### Phase 6 — PDF Assembly
- `renderer.py`: Jinja2 → HTML → WeasyPrint → PDF
- Chart PNG embedding
- Cover page with CMP and rating

### Phase 7 — Annual Report Processing
- `annual_report.py`: PDF download from BSE → pdfplumber text extraction
- Claude summarization of MD&A, Chairman's letter, Notes to Accounts
- Extraction of segment revenue, RPTs, contingent liabilities

---

## 12. Critical Files to Build First

1. `src/data/screener_client.py` — Core logic for screener.in JSON API authentication, rate-limited fetching, and data normalization; this is the primary financial data source and must be robust with caching
2. `src/data/bse_client.py` — BSE API client for shareholding XBRL, quarterly filings, and corporate actions; handles the official regulatory data layer
3. `src/ratios/calculator.py` — Pure pandas ratio calculation engine covering all standard Indian equity ratios (ROE, ROCE, DuPont, Z-score, etc.); the computational core of the report
4. `reports/templates/report_base.html` — Master Jinja2/HTML/CSS template defining the report layout, typography, and styling; sets the professional quality bar
5. `src/ai/narrative_generator.py` — Claude API integration layer that takes structured financial dicts and produces analyst-quality narrative text for each report section
