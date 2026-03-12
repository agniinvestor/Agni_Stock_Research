# Installed Plugins — India Stock Research Project

Three plugins from `anthropics/financial-services-plugins` are installed and active.

---

## 1. financial-analysis (Core Plugin)

The foundational plugin. Must be installed for the others to work. Provides valuation models, spreadsheet tools, and competitive analysis.

### Commands

| Command | Usage | Description |
|---|---|---|
| `/dcf` | `/dcf RELIANCE.NS` | Build a DCF valuation model with WACC, sensitivity analysis, terminal value — outputs Excel model |
| `/comps` | `/comps "Tata Consultancy Services"` | Build comparable company analysis with trading multiples (P/E, EV/EBITDA, P/B) |
| `/3-statement-model` | `/3-statement-model path/to/template.xlsx` | Fill out a 3-statement model template (P&L, Balance Sheet, Cash Flow) |
| `/lbo` | `/lbo "company name or deal details"` | Build an LBO model for a PE acquisition |
| `/competitive-analysis` | `/competitive-analysis "Indian IT sector"` | Build a competitive landscape deck with market positioning and peer comparison |
| `/check-deck` | `/check-deck path/to/presentation.pptx` | QC a PowerPoint deck for number consistency, errors, and IB-standard language |
| `/debug-model` | `/debug-model path/to/model.xlsx` | Audit a financial model for formula errors and inconsistencies |
| `/ppt-template` | `/ppt-template path/to/template.pptx` | Create a reusable PPT template skill from an existing PowerPoint |

### Skills (Auto-triggered by natural language)

| Skill | Triggers On | What It Does |
|---|---|---|
| **DCF Model** | "build a DCF", "value this company", "intrinsic value" | Retrieves financials from filings, builds cash flow projections, WACC, sensitivity analysis, Excel output |
| **Comps Analysis** | "comparable companies", "trading multiples", "peer valuation" | Institutional-grade comps table with operating metrics and valuation multiples in Excel |
| **3-Statement Model** | "fill out model", "populate financial model", "link IS/BS/CF" | Completes and links integrated financial statement templates |
| **LBO Model** | "LBO model", "leveraged buyout", "PE acquisition model" | Fills LBO model templates with deal assumptions and return analysis |
| **Competitive Analysis** | "competitive landscape", "market map", "who are the competitors to X" | Full competitive landscape deck — positioning, deep-dives, strategic synthesis |
| **Audit Spreadsheet** | "audit this sheet", "check my formulas", "model won't balance" | Audits formulas, finds errors, checks BS balance, cash tie-out |
| **Clean Data** | "clean this data", "fix formatting", "this data is messy" | Cleans spreadsheet data — whitespace, casing, date formats, duplicates |
| **Deck Refresh** | "update deck with Q4 numbers", "refresh the comps", "roll this forward" | Swaps new numbers across an existing presentation without rebuilding it |
| **IB Check Deck** | "check my deck", "QC this pitch", "is this client-ready" | Full IB-standard QC: number consistency, narrative alignment, language polish |
| **PPT Template Creator** | "create a reusable skill from my template" | Converts a PowerPoint template into a reusable skill |

---

## 2. equity-research

Sell-side style research workflows — earnings analysis, coverage initiation, investment theses, and morning notes. Most relevant for this project.

### Commands

| Command | Usage | Description |
|---|---|---|
| `/initiate` | `/initiate HDFCBANK.NS` | Create a full initiating coverage report (5-task workflow: research → model → valuation → charts → report) |
| `/earnings` | `/earnings INFY Q3 2025` | Analyze quarterly earnings and generate an earnings update report (8-12 pages, institutional format) |
| `/earnings-preview` | `/earnings-preview TATAMOTORS.NS` | Build pre-earnings preview with bull/bear scenarios and key metrics to watch |
| `/thesis` | `/thesis WIPRO.NS` | Create or update an investment thesis for a stock |
| `/screen` | `/screen "undervalued midcap Indian FMCG"` | Run a stock screen or generate investment ideas by criteria |
| `/sector` | `/sector "Indian Pharma"` | Create a full sector overview report |
| `/morning-note` | `/morning-note` | Draft a concise morning meeting note with overnight developments and trade ideas |
| `/catalysts` | `/catalysts "next 4 weeks"` | View or update the catalyst calendar for a coverage universe |
| `/model-update` | `/model-update SUNPHARMA.NS` | Update a financial model with new earnings data or guidance |

### Skills (Auto-triggered by natural language)

| Skill | Triggers On | What It Does |
|---|---|---|
| **Initiating Coverage** | "initiate coverage on X", "write an initiation report" | 5-task workflow: company research → financial model → valuation → charts → full DOCX report |
| **Earnings Analysis** | "earnings update", "Q3 results", "analyze quarterly earnings" | Post-earnings report: beat/miss analysis, key metrics, updated estimates, revised thesis |
| **Earnings Preview** | "earnings preview", "pre-earnings", "what to watch for X earnings" | Pre-earnings note with estimate model, bull/bear scenarios, key metrics to watch |
| **Investment Thesis** | "investment thesis for X", "bull case", "bear case" | Creates/updates structured investment thesis with catalysts and risk factors |
| **Idea Generation** | "stock screen", "find ideas", "pitch me something", "new ideas" | Systematic screening — quantitative filters + thematic research + pattern recognition |
| **Sector Overview** | "sector overview", "industry report", "market landscape" | Comprehensive sector report: dynamics, competitive positioning, key players, themes |
| **Morning Note** | "morning note", "morning call prep", "what happened overnight" | Tight, opinionated morning meeting note — overnight events, trade ideas, coverage updates |
| **Catalyst Calendar** | "catalyst calendar", "upcoming events", "earnings calendar" | Builds and maintains calendar of catalysts: earnings, conferences, regulatory events |
| **Model Update** | "update model", "plug earnings", "refresh estimates", "new guidance" | Updates financial model with new actuals, recalculates valuation, flags material changes |
| **Thesis Tracker** | "update thesis for X", "is my thesis still intact", "thesis check" | Maintains and updates investment theses, tracks data points and milestones over time |

---

## 3. wealth-management

Client-facing portfolio management workflows — performance reporting, financial planning, rebalancing, and tax optimization.

### Commands

| Command | Usage | Description |
|---|---|---|
| `/client-report` | `/client-report "Angad Sharma" Q4 2025` | Generate a client-facing performance report with returns, allocation, and market commentary |
| `/client-review` | `/client-review "Angad Sharma"` | Prep for a client review meeting — summary, talking points, action items |
| `/financial-plan` | `/financial-plan "Angad Sharma"` | Build or update a full financial plan (retirement, education, estate, cash flow) |
| `/proposal` | `/proposal "Prospect Name"` | Create an investment proposal for a new client or prospect |
| `/rebalance` | `/rebalance "Angad Sharma"` | Analyze portfolio allocation drift and generate rebalancing trade recommendations |
| `/tlh` | `/tlh "Angad Sharma"` | Identify tax-loss harvesting opportunities across taxable accounts |

### Skills (Auto-triggered by natural language)

| Skill | Triggers On | What It Does |
|---|---|---|
| **Client Report** | "client report", "performance report", "quarterly report for [client]" | Professional client performance report: returns, allocations, market commentary |
| **Client Review Prep** | "client review", "meeting prep", "quarterly review for [client]" | Meeting-ready format: performance summary, allocation analysis, talking points |
| **Financial Plan** | "financial plan", "retirement plan", "can I retire", "education funding" | Comprehensive plan: retirement projections, education funding, estate, cash flow |
| **Investment Proposal** | "investment proposal", "pitch new client", "proposal for [prospect]" | New client pitch: firm approach, proposed allocation, expected outcomes, fee structure |
| **Portfolio Rebalance** | "rebalance", "portfolio drift", "allocation check", "out of balance" | Drift analysis + rebalancing trades, considering tax implications and wash sale rules |
| **Tax-Loss Harvesting** | "tax-loss harvesting", "TLH", "harvest losses", "year-end tax planning" | Identifies unrealized losses, suggests replacement securities, tracks wash sale windows |

---

## Relevance to This Project

| Task | Best Plugin + Command |
|---|---|
| Generate a full research report on a stock | `/initiate TICKER.NS` (equity-research) |
| Analyze latest quarterly earnings | `/earnings TICKER Q4 2025` (equity-research) |
| Build a DCF valuation model | `/dcf TICKER.NS` (financial-analysis) |
| Peer/competitor valuation table | `/comps TICKER.NS` (financial-analysis) |
| Investment thesis write-up | `/thesis TICKER.NS` (equity-research) |
| Screen for stock ideas | `/screen "criteria"` (equity-research) |
| Sector deep-dive report | `/sector "Indian Banking"` (equity-research) |
| Update model after earnings | `/model-update TICKER.NS` (equity-research) |
| QC a generated report or deck | `/check-deck path/to/file.pptx` (financial-analysis) |
| Audit a financial model Excel file | `/debug-model path/to/model.xlsx` (financial-analysis) |

---

## Installation Reference

```bash
# Add marketplace
claude plugin marketplace add anthropics/financial-services-plugins

# Installed plugins
claude plugin install financial-analysis@financial-services-plugins
claude plugin install equity-research@financial-services-plugins
claude plugin install wealth-management@financial-services-plugins

# List installed plugins
claude plugin list
```
