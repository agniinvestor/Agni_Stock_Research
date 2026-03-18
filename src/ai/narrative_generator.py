"""
AI narrative generator using Claude API.

Generates all text sections of the equity research report from structured
financial data. Each section is a separate API call with JSON schema output.
Results cached in SQLite — re-generate only when input data changes.

7 sections generated:
1. investment_summary  — 5 thesis bullets + rating rationale
2. business_overview   — company overview, history, products, governance
3. investment_pillars  — 5 investment thesis pillars with supporting text
4. risks               — 8-12 risks with probability/impact
5. financial_analysis  — P&L narrative, valuation justification, recommendation
6. catalysts           — 5 near-term catalysts with timeframes
7. industry_overview   — industry TAM, competitive landscape, macro context

Cost: ~$0.11 per full report generation (all 7 sections)
Cache TTL: 30 days (re-generates when financial data changes)
"""

import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', '..'))
import json, sqlite3, logging, hashlib, re
from dataclasses import dataclass
from datetime import datetime
from config.settings import (
    ANTHROPIC_API_KEY, SQLITE_DB_PATH, NARRATIVE_MODEL,
    NARRATIVE_MAX_TOKENS, NARRATIVE_CACHE_TTL_DAYS
)

try:
    import anthropic
    HAS_ANTHROPIC = True
except ImportError:
    HAS_ANTHROPIC = False
    logging.getLogger(__name__).warning(
        "anthropic package not installed. Narrative generation disabled. "
        "Install with: pip install anthropic"
    )

logger = logging.getLogger(__name__)


class NarrativeGenerator:
    # Prompt templates as class-level constants for maintainability
    SYSTEM_PROMPT = """You are a senior equity research analyst at a top-tier Indian institutional brokerage (equivalent to Motilal Oswal, Kotak Securities, or IIFL). You are writing an initiating coverage research report.

CRITICAL RULES:
- Respond with ONLY a valid JSON object — no markdown, no preamble, no explanation
- All monetary values in INR Crores unless stated otherwise
- Use Indian number formatting terminology (Crores, Lakhs)
- Be specific and quantitative — lead with numbers
- Investment horizon: 12-18 months
- Rating criteria: BUY if upside > 15%, HOLD if 5-15%, SELL if < 5%"""

    # Ordered list of sections to generate
    SECTION_ORDER = [
        "investment_summary",
        "business_overview",
        "investment_pillars",
        "risks",
        "financial_analysis",
        "catalysts",
        "industry_overview",
    ]

    def __init__(self, db_path: str = None, api_key: str = None):
        """
        Initialize narrative generator.

        Args:
            db_path: SQLite path. Falls back to SQLITE_DB_PATH from settings.
            api_key: Anthropic API key. Falls back to ANTHROPIC_API_KEY from settings.
        """
        self.db_path = db_path or SQLITE_DB_PATH
        self.api_key = api_key or ANTHROPIC_API_KEY or os.environ.get("ANTHROPIC_API_KEY")
        self._init_db()

        if not self.api_key:
            logger.warning("ANTHROPIC_API_KEY not set — narrative generation will be skipped")

    def generate_all(self, report_json: dict) -> dict:
        """
        Generate all 7 sections. Returns updated report_json with narrative filled.
        Sections generated in order: summary → overview → pillars → risks →
        financial → catalysts → industry

        Continues even if individual sections fail (try/except per section).
        """
        ticker = report_json.get("meta", {}).get("ticker", "UNKNOWN")
        logger.info(f"[{ticker}] Starting narrative generation for all sections")

        section_generators = {
            "investment_summary": self._gen_investment_summary,
            "business_overview": self._gen_business_overview,
            "investment_pillars": self._gen_investment_pillars,
            "risks": self._gen_risks,
            "financial_analysis": self._gen_financial_analysis,
            "catalysts": self._gen_catalysts,
            "industry_overview": self._gen_industry_overview,
        }

        if "narrative" not in report_json:
            report_json["narrative"] = {}

        for section in self.SECTION_ORDER:
            try:
                logger.info(f"[{ticker}] Generating section: {section}")
                gen_fn = section_generators[section]
                result = gen_fn(report_json)
                if result:
                    report_json["narrative"][section] = result
                    logger.info(f"[{ticker}] Section '{section}' generated successfully")
                else:
                    logger.warning(f"[{ticker}] Section '{section}' returned empty result")
            except Exception as e:
                logger.error(f"[{ticker}] Failed to generate section '{section}': {e}")
                # Continue with remaining sections

        logger.info(f"[{ticker}] Narrative generation complete. "
                    f"Sections populated: {list(report_json['narrative'].keys())}")
        return report_json

    def _generate_section(self, section: str, ticker: str, context: dict, schema: dict) -> dict:
        """
        Check cache → if miss, call Claude → cache result → return.
        prompt_hash = sha256 of json.dumps(context, sort_keys=True)
        """
        prompt_hash = hashlib.sha256(
            json.dumps(context, sort_keys=True, default=str).encode()
        ).hexdigest()

        # Check cache first
        cached = self._cache_get(ticker, section, prompt_hash)
        if cached is not None:
            logger.info(f"[{ticker}] Cache hit for section '{section}'")
            return cached

        # Build user prompt from context + schema
        user_prompt = (
            f"Generate the '{section}' section for {ticker}.\n\n"
            f"FINANCIAL CONTEXT:\n{json.dumps(context, indent=2, default=str)}\n\n"
            f"OUTPUT SCHEMA (return ONLY this JSON structure):\n"
            f"{json.dumps(schema, indent=2)}"
        )

        result = self._call_claude(user_prompt)

        if result:
            self._cache_set(ticker, section, prompt_hash, result)

        return result

    def _call_claude(self, user_prompt: str, max_tokens: int = 2048) -> dict:
        """
        Call Claude API. Return parsed JSON dict.
        On JSON parse error, try to extract JSON from response with regex.
        On API error, log and return {} (empty dict — pipeline continues).
        Model: NARRATIVE_MODEL from settings (claude-sonnet-4-6).
        """
        if not HAS_ANTHROPIC:
            logger.warning("anthropic not installed — cannot call Claude API")
            return {}
        if not self.api_key:
            logger.warning("ANTHROPIC_API_KEY not set — cannot call Claude API")
            return {}

        try:
            client = anthropic.Anthropic(api_key=self.api_key)
            response = client.messages.create(
                model=NARRATIVE_MODEL,
                max_tokens=max_tokens,
                system=self.SYSTEM_PROMPT,
                messages=[{"role": "user", "content": user_prompt}],
            )
            raw_text = response.content[0].text.strip()

            # Strip markdown fences if present
            raw_text = re.sub(r"^```(?:json)?\s*", "", raw_text)
            raw_text = re.sub(r"\s*```$", "", raw_text)

            try:
                return json.loads(raw_text)
            except json.JSONDecodeError:
                # Try to extract JSON object using regex
                json_match = re.search(r"\{[\s\S]*\}", raw_text)
                if json_match:
                    try:
                        return json.loads(json_match.group(0))
                    except json.JSONDecodeError as e:
                        logger.warning(f"Could not parse extracted JSON: {e}")
                        return {}
                logger.warning("No valid JSON found in Claude response")
                return {}

        except Exception as e:
            logger.error(f"Claude API error: {e}")
            return {}

    # ── Section generators ─────────────────────────────────────────────────────

    def _gen_investment_summary(self, report_json: dict) -> dict:
        """
        Context: meta, market_data, last 3yr P&L (net_revenue, ebitda, pat, eps),
        valuation (price_target, rating, upside_pct), peer multiples.

        Schema output:
        {
          "investment_summary": [string × 5],  // thesis bullets, each <= 40 words
          "rating": "BUY" | "HOLD" | "SELL",
          "price_target_rationale": string      // 1-2 sentences
        }
        """
        ticker = report_json.get("meta", {}).get("ticker", "UNKNOWN")

        compact_fin = self._compact_financials(report_json, n_years=3)
        context = {
            "meta": report_json.get("meta", {}),
            "market_data": report_json.get("market_data", {}),
            "financials_3yr": compact_fin,
            "valuation": report_json.get("valuation", {}),
            "peer_multiples": report_json.get("peer_multiples", {}),
        }
        schema = {
            "investment_summary": ["string (<=40 words each)"],
            "rating": "BUY | HOLD | SELL",
            "price_target_rationale": "string (1-2 sentences)"
        }
        return self._generate_section("investment_summary", ticker, context, schema)

    def _gen_business_overview(self, report_json: dict) -> dict:
        """
        Context: meta (company_name, sector, ticker), segments dict,
        market_data (market_cap_cr), is_consolidated flag.

        Schema output:
        {
          "company_overview": string,      // 2-3 paragraphs, 200-300 words
          "company_history": string,       // 100-150 word timeline paragraph
          "products_services": [string],   // 5-8 product/segment bullet points
          "customer_channels": string,     // 100 word distribution description
          "governance": string             // 100 word governance paragraph
        }
        """
        ticker = report_json.get("meta", {}).get("ticker", "UNKNOWN")

        context = {
            "meta": report_json.get("meta", {}),
            "segments": report_json.get("segments", {}),
            "market_data": {
                "market_cap_cr": report_json.get("market_data", {}).get("market_cap_cr"),
            },
            "is_consolidated": report_json.get("is_consolidated", True),
        }
        schema = {
            "company_overview": "string (200-300 words, 2-3 paragraphs)",
            "company_history": "string (100-150 words, timeline format)",
            "products_services": ["string (5-8 bullet points)"],
            "customer_channels": "string (~100 words)",
            "governance": "string (~100 words)"
        }
        return self._generate_section("business_overview", ticker, context, schema)

    def _gen_investment_pillars(self, report_json: dict) -> dict:
        """
        Context: financials (10yr key metrics), ratios (key trends),
        valuation (price target, upside), market_data.

        Schema output:
        {
          "investment_pillars": [
            {
              "title": string,
              "subtitle": string,       // 1 sentence tagline
              "paragraphs": [string, string, string]  // 3 paragraphs, 80-120 words each
            }
          ]  // exactly 5 items
        }
        """
        ticker = report_json.get("meta", {}).get("ticker", "UNKNOWN")

        compact_fin = self._compact_financials(report_json, n_years=10)
        context = {
            "meta": report_json.get("meta", {}),
            "financials_10yr": compact_fin,
            "ratios": report_json.get("ratios", {}),
            "valuation": report_json.get("valuation", {}),
            "market_data": report_json.get("market_data", {}),
        }
        schema = {
            "investment_pillars": [
                {
                    "title": "string",
                    "subtitle": "string (1 sentence tagline)",
                    "paragraphs": [
                        "string (80-120 words)",
                        "string (80-120 words)",
                        "string (80-120 words)"
                    ]
                }
            ]
        }
        return self._generate_section("investment_pillars", ticker, context, schema)

    def _gen_risks(self, report_json: dict) -> dict:
        """
        Context: meta (sector), financials (leverage, margins), market_data.

        Schema output:
        {
          "risks_intro": string,         // 50 word intro paragraph
          "risks": [
            {
              "title": string,
              "category": "Regulatory" | "Business" | "Financial" | "Macro" | "Execution",
              "probability": "High" | "Medium" | "Low",
              "impact": "High" | "Medium" | "Low",
              "description": string,     // 40-60 words
              "mitigation": string       // 20-30 words
            }
          ],  // 8-12 items
          "extended_risks": string       // 400-500 word detailed risk discussion
        }
        """
        ticker = report_json.get("meta", {}).get("ticker", "UNKNOWN")

        compact_fin = self._compact_financials(report_json, n_years=5)
        # Extract leverage and margin data specifically
        leverage = {}
        margins = {}
        for yr, metrics in compact_fin.items():
            leverage[yr] = {
                "debt_equity": metrics.get("debt_equity"),
                "interest_coverage": metrics.get("interest_coverage"),
            }
            margins[yr] = {
                "ebitda_margin_pct": metrics.get("ebitda_margin_pct"),
                "pat_margin_pct": metrics.get("pat_margin_pct"),
            }

        context = {
            "meta": report_json.get("meta", {}),
            "leverage_5yr": leverage,
            "margins_5yr": margins,
            "market_data": report_json.get("market_data", {}),
        }
        schema = {
            "risks_intro": "string (~50 words)",
            "risks": [
                {
                    "title": "string",
                    "category": "Regulatory | Business | Financial | Macro | Execution",
                    "probability": "High | Medium | Low",
                    "impact": "High | Medium | Low",
                    "description": "string (40-60 words)",
                    "mitigation": "string (20-30 words)"
                }
            ],
            "extended_risks": "string (400-500 words)"
        }
        return self._generate_section("risks", ticker, context, schema)

    def _gen_financial_analysis(self, report_json: dict) -> dict:
        """
        Context: financials (income_statement, balance_sheet, cash_flow — last 5yr),
        ratios (all computed ratios — last 5yr), projections (3yr forward).
        Pass compacted version — only key metrics, not full arrays.

        Schema output:
        {
          "financial_analysis": string,   // 400-600 word P&L + BS + CF narrative
          "projection_assumptions": [     // 8-10 key modelling assumptions
            {"name": string, "value": string, "rationale": string}
          ],
          "valuation_narrative": string,  // 200-300 word valuation methodology justification
          "recommendation": string        // 100-150 word final investment conclusion
        }
        """
        ticker = report_json.get("meta", {}).get("ticker", "UNKNOWN")

        compact_fin = self._compact_financials(report_json, n_years=5)
        context = {
            "meta": report_json.get("meta", {}),
            "financials_5yr": compact_fin,
            "ratios_5yr": report_json.get("ratios", {}),
            "projections_3yr": report_json.get("projections", {}),
            "valuation": report_json.get("valuation", {}),
        }
        schema = {
            "financial_analysis": "string (400-600 words covering P&L, balance sheet, cash flows)",
            "projection_assumptions": [
                {
                    "name": "string",
                    "value": "string",
                    "rationale": "string"
                }
            ],
            "valuation_narrative": "string (200-300 words on methodology and justification)",
            "recommendation": "string (100-150 word final investment conclusion)"
        }
        return self._generate_section("financial_analysis", ticker, context, schema)

    def _gen_catalysts(self, report_json: dict) -> dict:
        """
        Context: meta (sector), valuation (price_target, upside), market_data.

        Schema output:
        {
          "catalysts": [
            {
              "title": string,
              "timeframe": "0-3 months" | "3-6 months" | "6-12 months" | "12-24 months",
              "description": string,   // 40-60 words
              "potential_impact": string  // quantified impact if possible
            }
          ]  // exactly 5 items
        }
        """
        ticker = report_json.get("meta", {}).get("ticker", "UNKNOWN")

        context = {
            "meta": report_json.get("meta", {}),
            "valuation": {
                "price_target": report_json.get("valuation", {}).get("price_target"),
                "upside_pct": report_json.get("valuation", {}).get("upside_pct"),
                "rating": report_json.get("valuation", {}).get("rating"),
            },
            "market_data": report_json.get("market_data", {}),
        }
        schema = {
            "catalysts": [
                {
                    "title": "string",
                    "timeframe": "0-3 months | 3-6 months | 6-12 months | 12-24 months",
                    "description": "string (40-60 words)",
                    "potential_impact": "string (quantified if possible)"
                }
            ]
        }
        return self._generate_section("catalysts", ticker, context, schema)

    def _gen_industry_overview(self, report_json: dict) -> dict:
        """
        Context: meta (company_name, sector), market_data (market_cap_cr),
        financials (net_revenue last 5yr for market share context).

        Schema output:
        {
          "industry_overview": string,         // 300-400 word industry + TAM analysis
          "competitive_landscape": string,     // 200-250 word competitive positioning
          "regulatory_environment": string,    // 100-150 word regulatory context
          "sector_tailwinds": [string],        // 4-5 bullet points
          "sector_headwinds": [string]         // 3-4 bullet points
        }
        """
        ticker = report_json.get("meta", {}).get("ticker", "UNKNOWN")

        compact_fin = self._compact_financials(report_json, n_years=5)
        # Extract only net_revenue for market share context
        revenue_history = {yr: {"net_revenue": m.get("net_revenue")} for yr, m in compact_fin.items()}

        context = {
            "meta": {
                "company_name": report_json.get("meta", {}).get("company_name"),
                "sector": report_json.get("meta", {}).get("sector"),
                "ticker": ticker,
            },
            "market_data": {
                "market_cap_cr": report_json.get("market_data", {}).get("market_cap_cr"),
            },
            "revenue_5yr": revenue_history,
        }
        schema = {
            "industry_overview": "string (300-400 words, include TAM and growth estimates)",
            "competitive_landscape": "string (200-250 words, name key competitors)",
            "regulatory_environment": "string (100-150 words)",
            "sector_tailwinds": ["string (4-5 bullet points)"],
            "sector_headwinds": ["string (3-4 bullet points)"]
        }
        return self._generate_section("industry_overview", ticker, context, schema)

    # ── Cache helpers ──────────────────────────────────────────────────────────

    def _init_db(self) -> None:
        """Initialize SQLite DB and create narrative_cache table."""
        db_dir = os.path.dirname(self.db_path)
        if db_dir:
            os.makedirs(db_dir, exist_ok=True)
        with sqlite3.connect(self.db_path) as conn:
            conn.execute("""
                CREATE TABLE IF NOT EXISTS narrative_cache (
                    ticker        TEXT,
                    section       TEXT,
                    prompt_hash   TEXT,
                    generated_at  TEXT,
                    result_json   TEXT,
                    PRIMARY KEY (ticker, section, prompt_hash)
                )
            """)
            conn.commit()

    def _cache_get(self, ticker: str, section: str, prompt_hash: str) -> dict | None:
        """
        Return cached result if it exists and is within NARRATIVE_CACHE_TTL_DAYS.
        Returns None on miss, stale entry, or error.
        """
        try:
            with sqlite3.connect(self.db_path) as conn:
                row = conn.execute(
                    "SELECT generated_at, result_json FROM narrative_cache "
                    "WHERE ticker = ? AND section = ? AND prompt_hash = ?",
                    (ticker, section, prompt_hash)
                ).fetchone()

            if row is None:
                return None

            generated_at_str, result_json_str = row
            generated_at = datetime.fromisoformat(generated_at_str)
            age_days = (datetime.utcnow() - generated_at).days
            ttl = NARRATIVE_CACHE_TTL_DAYS if NARRATIVE_CACHE_TTL_DAYS else 30

            if age_days > ttl:
                logger.info(
                    f"[{ticker}] Narrative cache stale ({age_days}d > {ttl}d) "
                    f"for section '{section}'"
                )
                return None

            return json.loads(result_json_str)

        except Exception as e:
            logger.warning(f"[{ticker}] Narrative cache read error for '{section}': {e}")
            return None

    def _cache_set(self, ticker: str, section: str, prompt_hash: str, result: dict) -> None:
        """Persist generated section to SQLite narrative_cache."""
        try:
            with sqlite3.connect(self.db_path) as conn:
                conn.execute(
                    """
                    INSERT OR REPLACE INTO narrative_cache
                        (ticker, section, prompt_hash, generated_at, result_json)
                    VALUES (?, ?, ?, ?, ?)
                    """,
                    (
                        ticker,
                        section,
                        prompt_hash,
                        datetime.utcnow().isoformat(),
                        json.dumps(result),
                    )
                )
                conn.commit()
            logger.debug(f"[{ticker}] Cached narrative section '{section}'")
        except Exception as e:
            logger.warning(f"[{ticker}] Narrative cache write error for '{section}': {e}")

    def _compact_financials(self, report_json: dict, n_years: int = 5) -> dict:
        """
        Return last n_years of key financial metrics, compacted for prompt context.

        Pulls from report_json["financials"] which is expected to contain
        year-keyed dicts with income_statement, balance_sheet, and cash_flow data.

        Returns: {year_label: {metric: value, ...}, ...}
        """
        financials = report_json.get("financials", {})
        if not financials:
            return {}

        # Sort years descending and take n_years most recent
        try:
            sorted_years = sorted(financials.keys(), reverse=True)[:n_years]
        except Exception:
            return {}

        compact = {}
        for yr in sorted_years:
            yr_data = financials.get(yr, {})
            income = yr_data.get("income_statement", {})
            balance = yr_data.get("balance_sheet", {})
            cashflow = yr_data.get("cash_flow", {})
            ratios = yr_data.get("ratios", report_json.get("ratios", {}).get(yr, {}))

            compact[yr] = {
                # Income statement
                "net_revenue": income.get("net_revenue") or income.get("revenue"),
                "ebitda": income.get("ebitda"),
                "ebitda_margin_pct": income.get("ebitda_margin_pct"),
                "pat": income.get("pat") or income.get("net_profit"),
                "pat_margin_pct": income.get("pat_margin_pct"),
                "eps": income.get("eps") or income.get("basic_eps"),
                # Balance sheet
                "total_assets": balance.get("total_assets"),
                "net_debt": balance.get("net_debt"),
                "debt_equity": balance.get("debt_equity") or ratios.get("debt_equity"),
                "book_value_per_share": balance.get("book_value_per_share"),
                # Cash flow
                "operating_cash_flow": cashflow.get("operating_cash_flow") or cashflow.get("cfo"),
                "capex": cashflow.get("capex"),
                "free_cash_flow": cashflow.get("free_cash_flow") or cashflow.get("fcf"),
                # Key ratios
                "roe_pct": ratios.get("roe_pct") or ratios.get("roe"),
                "roce_pct": ratios.get("roce_pct") or ratios.get("roce"),
                "interest_coverage": ratios.get("interest_coverage"),
                "pe_ratio": ratios.get("pe_ratio") or ratios.get("pe"),
                "ev_ebitda": ratios.get("ev_ebitda"),
            }
            # Remove None values to keep prompt compact
            compact[yr] = {k: v for k, v in compact[yr].items() if v is not None}

        return compact


# ── CLI helper ─────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import argparse

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(name)s: %(message)s"
    )

    parser = argparse.ArgumentParser(description="Generate AI narrative for equity research report")
    parser.add_argument("report_json_path", help="Path to report JSON file")
    parser.add_argument("--db-path", default="/tmp/india_research.db")
    parser.add_argument(
        "--section",
        choices=NarrativeGenerator.SECTION_ORDER + ["all"],
        default="all",
        help="Which section(s) to generate"
    )
    parser.add_argument("--output", help="Output path for updated report JSON")
    args = parser.parse_args()

    with open(args.report_json_path) as f:
        report = json.load(f)

    gen = NarrativeGenerator(db_path=args.db_path)

    if args.section == "all":
        report = gen.generate_all(report)
    else:
        section_fn_map = {
            "investment_summary": gen._gen_investment_summary,
            "business_overview": gen._gen_business_overview,
            "investment_pillars": gen._gen_investment_pillars,
            "risks": gen._gen_risks,
            "financial_analysis": gen._gen_financial_analysis,
            "catalysts": gen._gen_catalysts,
            "industry_overview": gen._gen_industry_overview,
        }
        fn = section_fn_map[args.section]
        result = fn(report)
        if "narrative" not in report:
            report["narrative"] = {}
        report["narrative"][args.section] = result

    output_path = args.output or args.report_json_path
    with open(output_path, "w") as f:
        json.dump(report, f, indent=2, default=str)

    print(f"Report with narrative saved to: {output_path}")

    # Print a summary of what was generated
    narrative = report.get("narrative", {})
    print(f"\nGenerated sections ({len(narrative)}):")
    for section_name, content in narrative.items():
        if content:
            print(f"  [OK] {section_name}")
        else:
            print(f"  [EMPTY] {section_name}")
