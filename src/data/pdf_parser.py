"""
Annual report PDF parser for Indian listed companies.

Downloads annual report PDFs from BSE and extracts key sections:
- MD&A (Management Discussion and Analysis)
- Chairman's Letter
- Segment Revenue / EBIT from Notes to Accounts
- Key operational data (headcount, capex details, etc.)

Uses PyMuPDF (fitz) for text extraction + Claude API for structured summarization.
All extracted text and structured results cached in SQLite.
"""

import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', '..'))
import json, sqlite3, logging, hashlib, re
from dataclasses import dataclass, field, asdict
from datetime import datetime
from typing import Optional
import requests
from config.settings import (
    ANTHROPIC_API_KEY, SQLITE_DB_PATH, PDF_CACHE_DIR,
    BSE_API_BASE_URL, NARRATIVE_MODEL
)

try:
    import fitz  # PyMuPDF
    HAS_PYMUPDF = True
except ImportError:
    HAS_PYMUPDF = False

try:
    import anthropic
    HAS_ANTHROPIC = True
except ImportError:
    HAS_ANTHROPIC = False

logger = logging.getLogger(__name__)


@dataclass
class AnnualReportData:
    ticker: str
    bse_code: str
    fy_label: str              # e.g. "FY25"
    report_year: int           # e.g. 2025
    pdf_path: str | None       # local path to downloaded PDF

    # Extracted sections (raw text)
    mda_text: str | None = None        # MD&A raw text
    chairman_text: str | None = None
    notes_text: str | None = None      # Notes to Accounts (relevant sections)

    # Structured output (from Claude summarization)
    business_highlights: list = field(default_factory=list)   # 5-8 bullet points from MD&A
    key_risks_mentioned: list = field(default_factory=list)   # risks mentioned by management
    segment_commentary: dict = field(default_factory=dict)    # {segment: management commentary}
    capex_guidance: str | None = None
    revenue_guidance: str | None = None     # if management provided guidance
    margin_commentary: str | None = None

    # Operational data (extracted from text if present)
    headcount: int | None = None
    attrition_pct: float | None = None

    fetched_at: str = field(default_factory=lambda: datetime.utcnow().isoformat())
    parse_errors: list = field(default_factory=list)    # non-fatal errors during parsing


class AnnualReportParser:
    BSE_FILINGS_URL = "https://www.bseindia.com/corporates/ann.html"
    BSE_PDF_BASE = "https://www.bseindia.com/xml-data/corpfiling/AttachLive/"

    # Sections to detect by heading keywords
    SECTION_KEYWORDS = {
        "mda": ["management discussion", "management's discussion", "md&a",
                "management discussion and analysis", "management review"],
        "chairman": ["chairman", "managing director", "chairman's message",
                    "dear shareholder", "letter to shareholders"],
        "notes": ["notes to accounts", "notes to financial statements",
                  "notes forming part"],
        "segment": ["segment information", "segment reporting",
                    "geographical information", "segment revenue"],
    }

    def __init__(self, db_path: str = None, api_key: str = None, pdf_cache_dir: str = None):
        """
        Initialize parser.

        Args:
            db_path: Path to SQLite DB. Falls back to SQLITE_DB_PATH from settings.
            api_key: Anthropic API key. Falls back to ANTHROPIC_API_KEY env var.
            pdf_cache_dir: Directory to store downloaded PDFs.
        """
        self.db_path = db_path or SQLITE_DB_PATH
        self.api_key = api_key or ANTHROPIC_API_KEY or os.environ.get("ANTHROPIC_API_KEY")
        self.pdf_cache_dir = pdf_cache_dir or PDF_CACHE_DIR or "/tmp/pdf_cache"

        os.makedirs(self.pdf_cache_dir, exist_ok=True)
        self._init_db()

        if not HAS_PYMUPDF:
            logger.warning(
                "PyMuPDF (fitz) not installed. PDF text extraction disabled. "
                "Install with: pip install pymupdf"
            )
        if not HAS_ANTHROPIC:
            logger.warning(
                "anthropic package not installed. AI summarization disabled. "
                "Install with: pip install anthropic"
            )
        if not self.api_key:
            logger.warning("ANTHROPIC_API_KEY not set. AI summarization disabled.")

    def parse_latest(self, ticker: str, bse_code: str, fy_year: int = None) -> AnnualReportData:
        """
        Main entry point.
        1. Check SQLite cache — if fresh (<90 days), return cached
        2. Download PDF from BSE
        3. Extract text sections
        4. Summarize with Claude API
        5. Cache and return
        """
        if fy_year is None:
            # Default to the most recently completed financial year
            now = datetime.utcnow()
            # Indian FY ends March 31; if we're past March, FY ended this year
            fy_year = now.year if now.month > 3 else now.year - 1

        fy_label = f"FY{str(fy_year)[2:]}"

        # 1. Check cache
        cached = self._cache_get(bse_code, fy_year)
        if cached is not None:
            logger.info(f"[{ticker}] Returning cached annual report data for {fy_label}")
            return cached

        logger.info(f"[{ticker}] Fetching annual report for {fy_label} (BSE: {bse_code})")
        parse_errors = []

        # 2. Download PDF
        pdf_path = None
        pdf_url = self._find_pdf_url(bse_code, fy_year)
        if pdf_url:
            pdf_path = self._download_pdf(pdf_url, ticker, fy_year)
            if pdf_path is None:
                parse_errors.append(f"PDF download failed for URL: {pdf_url}")
        else:
            parse_errors.append(f"Could not find annual report PDF URL for {bse_code} {fy_label}")
            logger.warning(f"[{ticker}] Could not locate PDF for {fy_label}")

        # 3. Extract text sections
        sections = {}
        if pdf_path and HAS_PYMUPDF:
            try:
                sections = self._extract_sections(pdf_path)
            except Exception as e:
                parse_errors.append(f"Section extraction error: {e}")
                logger.warning(f"[{ticker}] Section extraction failed: {e}")

        mda_text = self._extract_mda(sections)
        chairman_text = sections.get("chairman")
        notes_text = self._extract_segment_notes(sections)

        # 4. Summarize with Claude
        structured = {}
        if mda_text or notes_text:
            structured = self._summarize_with_claude(
                ticker, fy_label,
                mda_text or "",
                notes_text or ""
            )
        else:
            parse_errors.append("No text extracted — skipping AI summarization")

        # 5. Build result
        result = AnnualReportData(
            ticker=ticker,
            bse_code=bse_code,
            fy_label=fy_label,
            report_year=fy_year,
            pdf_path=pdf_path,
            mda_text=mda_text,
            chairman_text=chairman_text,
            notes_text=notes_text,
            business_highlights=structured.get("business_highlights", []),
            key_risks_mentioned=structured.get("key_risks_mentioned", []),
            segment_commentary=structured.get("segment_commentary", {}),
            capex_guidance=structured.get("capex_guidance"),
            revenue_guidance=structured.get("revenue_guidance"),
            margin_commentary=structured.get("margin_commentary"),
            headcount=structured.get("headcount"),
            attrition_pct=structured.get("attrition_pct"),
            fetched_at=datetime.utcnow().isoformat(),
            parse_errors=parse_errors,
        )

        # Cache and return
        self._cache_set(result)
        return result

    def _find_pdf_url(self, bse_code: str, fy_year: int) -> str | None:
        """
        Try to find annual report PDF URL from BSE.
        Strategy: scrape https://www.bseindia.com/corporates/ann.html
        with params: scrip=bse_code, category=Annual Report
        Look for PDF attachment links.
        Return None if not found (will skip PDF parsing gracefully).

        Note: BSE website is often slow/unreliable. Implements 10s timeout.
        Returns None on timeout rather than raising.
        """
        try:
            params = {
                "scrip": bse_code,
                "category": "Annual Report",
            }
            headers = {
                "User-Agent": (
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/120.0.0.0 Safari/537.36"
                ),
                "Referer": "https://www.bseindia.com/",
            }
            resp = requests.get(
                self.BSE_FILINGS_URL,
                params=params,
                headers=headers,
                timeout=10
            )
            resp.raise_for_status()

            # Look for PDF links in the HTML response
            # BSE URLs typically contain .pdf filenames
            pdf_pattern = re.compile(
                r'(?:AttachLive|corpfiling)[^"\']*?\.pdf',
                re.IGNORECASE
            )
            matches = pdf_pattern.findall(resp.text)

            if not matches:
                # Also try looking for direct PDF references
                direct_pattern = re.compile(
                    r'href=["\']([^"\']*?\.pdf)["\']',
                    re.IGNORECASE
                )
                direct_matches = direct_pattern.findall(resp.text)
                matches = [m for m in direct_matches if "annual" in m.lower() or "ar" in m.lower()]

            if matches:
                # Prefer the most recent one (heuristic: last match tends to be newest)
                pdf_path_fragment = matches[-1]
                if pdf_path_fragment.startswith("http"):
                    return pdf_path_fragment
                return f"{self.BSE_PDF_BASE}{pdf_path_fragment.lstrip('/')}"

            logger.info(f"No PDF links found on BSE for code {bse_code}")
            return None

        except requests.Timeout:
            logger.warning(f"BSE request timed out for {bse_code} — skipping PDF")
            return None
        except Exception as e:
            logger.warning(f"Error finding PDF URL for {bse_code}: {e}")
            return None

    def _download_pdf(self, pdf_url: str, ticker: str, fy_year: int) -> str | None:
        """Download PDF to pdf_cache_dir. Return local path. Return None on failure."""
        try:
            url_hash = hashlib.md5(pdf_url.encode()).hexdigest()[:8]
            filename = f"{ticker}_{fy_year}_{url_hash}.pdf"
            local_path = os.path.join(self.pdf_cache_dir, filename)

            # Return cached file if already downloaded
            if os.path.exists(local_path) and os.path.getsize(local_path) > 1024:
                logger.info(f"[{ticker}] PDF already cached at {local_path}")
                return local_path

            headers = {
                "User-Agent": (
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/120.0.0.0 Safari/537.36"
                ),
            }
            logger.info(f"[{ticker}] Downloading PDF from {pdf_url}")
            resp = requests.get(pdf_url, headers=headers, timeout=60, stream=True)
            resp.raise_for_status()

            content_type = resp.headers.get("Content-Type", "")
            if "pdf" not in content_type.lower() and not pdf_url.lower().endswith(".pdf"):
                logger.warning(f"[{ticker}] Unexpected content type: {content_type}")

            with open(local_path, "wb") as f:
                for chunk in resp.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)

            file_size = os.path.getsize(local_path)
            if file_size < 1024:
                os.remove(local_path)
                logger.warning(f"[{ticker}] Downloaded PDF too small ({file_size} bytes)")
                return None

            logger.info(f"[{ticker}] PDF saved to {local_path} ({file_size / 1024:.1f} KB)")
            return local_path

        except Exception as e:
            logger.warning(f"[{ticker}] PDF download failed: {e}")
            return None

    def _extract_sections(self, pdf_path: str) -> dict[str, str]:
        """
        Use PyMuPDF (fitz) to extract text by section.
        Strategy:
        1. Open PDF, iterate all pages
        2. For each text block, check if it's a heading (font size > 12 or ALL CAPS)
        3. Map each heading to a section type via SECTION_KEYWORDS
        4. Collect all text under each heading until the next heading is found
        5. Return {section_name: accumulated_text}

        Import: try fitz import; if not installed, return {} and log warning.
        """
        if not HAS_PYMUPDF:
            logger.warning("PyMuPDF not available — cannot extract PDF sections")
            return {}

        sections: dict[str, list[str]] = {}
        current_section: str | None = None

        try:
            doc = fitz.open(pdf_path)
            total_pages = len(doc)
            logger.info(f"Extracting text from {total_pages}-page PDF: {pdf_path}")

            for page_num in range(total_pages):
                page = doc[page_num]

                # Extract text with block-level metadata
                blocks = page.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE)["blocks"]

                for block in blocks:
                    if block.get("type") != 0:  # 0 = text block
                        continue

                    for line in block.get("lines", []):
                        line_text = ""
                        max_font_size = 0.0
                        is_bold = False

                        for span in line.get("spans", []):
                            span_text = span.get("text", "").strip()
                            if not span_text:
                                continue
                            line_text += span_text + " "
                            font_size = span.get("size", 0)
                            if font_size > max_font_size:
                                max_font_size = font_size
                            flags = span.get("flags", 0)
                            if flags & 16:  # bold flag
                                is_bold = True

                        line_text = line_text.strip()
                        if not line_text:
                            continue

                        # Determine if this line is a heading
                        is_heading = (
                            max_font_size > 12
                            or (len(line_text) < 120 and line_text == line_text.upper() and len(line_text) > 4)
                            or (is_bold and max_font_size >= 11 and len(line_text) < 100)
                        )

                        if is_heading:
                            # Check if this heading matches a known section
                            heading_lower = line_text.lower()
                            matched_section = None
                            for section_name, keywords in self.SECTION_KEYWORDS.items():
                                if any(kw in heading_lower for kw in keywords):
                                    matched_section = section_name
                                    break

                            if matched_section:
                                current_section = matched_section
                                if current_section not in sections:
                                    sections[current_section] = []
                                continue  # Don't add heading text itself

                        # Append text to current section
                        if current_section is not None:
                            sections[current_section].append(line_text)

            doc.close()

        except Exception as e:
            logger.error(f"PDF extraction error for {pdf_path}: {e}")
            return {}

        # Join accumulated text for each section
        return {
            section_name: "\n".join(lines)
            for section_name, lines in sections.items()
            if lines
        }

    def _extract_mda(self, sections: dict) -> str | None:
        """Return MD&A text, truncated to 8000 chars (Claude context management)."""
        # Try "mda" first, then "segment" as fallback partial source
        text = sections.get("mda") or sections.get("segment")
        if not text:
            return None
        return text[:8000]

    def _extract_segment_notes(self, sections: dict) -> str | None:
        """Return segment-related notes text, truncated to 4000 chars."""
        # Prefer dedicated segment section, fall back to notes
        text = sections.get("segment") or sections.get("notes")
        if not text:
            return None
        return text[:4000]

    def _summarize_with_claude(
        self,
        ticker: str,
        fy_label: str,
        mda_text: str,
        segment_text: str
    ) -> dict:
        """
        Single Claude API call to extract structured data from MD&A + segment notes.

        System prompt instructs the model to return ONLY valid JSON.
        Uses claude-haiku-4-5-20251001 for fast, low-cost extraction.
        Returns empty dict on API error.
        """
        if not HAS_ANTHROPIC:
            logger.warning("anthropic not installed — skipping AI summarization")
            return {}
        if not self.api_key:
            logger.warning("ANTHROPIC_API_KEY not set — skipping AI summarization")
            return {}

        system_prompt = (
            "You are an equity research analyst extracting structured data from an "
            "Indian company's annual report. Respond with ONLY a valid JSON object."
        )

        user_prompt = f"""Extract key information from this annual report for {ticker} ({fy_label}).

MD&A TEXT:
{mda_text[:6000] if mda_text else "(not available)"}

SEGMENT / NOTES TEXT:
{segment_text[:2000] if segment_text else "(not available)"}

Return a JSON object with EXACTLY these fields:
{{
  "business_highlights": ["string"],
  "key_risks_mentioned": ["string"],
  "segment_commentary": {{"segment_name": "commentary_string"}},
  "capex_guidance": "string or null",
  "revenue_guidance": "string or null",
  "margin_commentary": "string or null",
  "headcount": integer_or_null,
  "attrition_pct": float_or_null
}}

Rules:
- business_highlights: 5-8 key business highlights from the MD&A
- key_risks_mentioned: specific risks management acknowledged
- segment_commentary: keyed by segment name, value is management's comment on that segment
- capex_guidance: exact quote or paraphrase of capex plans if mentioned, else null
- revenue_guidance: exact quote or paraphrase of revenue guidance if given, else null
- margin_commentary: management's view on margins going forward, else null
- headcount: total employee count if mentioned as an integer, else null
- attrition_pct: attrition rate as float (e.g. 14.5 for 14.5%), else null
"""

        try:
            client = anthropic.Anthropic(api_key=self.api_key)
            response = client.messages.create(
                model="claude-haiku-4-5-20251001",
                max_tokens=1024,
                system=system_prompt,
                messages=[{"role": "user", "content": user_prompt}],
            )
            raw_text = response.content[0].text.strip()

            # Parse JSON — strip any markdown fences if present
            raw_text = re.sub(r"^```(?:json)?\s*", "", raw_text)
            raw_text = re.sub(r"\s*```$", "", raw_text)
            return json.loads(raw_text)

        except json.JSONDecodeError as e:
            logger.warning(f"[{ticker}] Claude returned non-JSON response: {e}")
            return {}
        except Exception as e:
            logger.warning(f"[{ticker}] Claude API error during summarization: {e}")
            return {}

    # ── Cache helpers ──────────────────────────────────────────────────────────

    def _init_db(self) -> None:
        """Initialize SQLite DB and create tables if they don't exist."""
        os.makedirs(os.path.dirname(self.db_path), exist_ok=True) if os.path.dirname(self.db_path) else None
        with sqlite3.connect(self.db_path) as conn:
            conn.execute("""
                CREATE TABLE IF NOT EXISTS pdf_cache (
                    bse_code     TEXT,
                    fy_year      INTEGER,
                    parsed_at    TEXT,
                    pdf_path     TEXT,
                    data_json    TEXT,
                    PRIMARY KEY (bse_code, fy_year)
                )
            """)
            conn.commit()

    def _cache_get(self, bse_code: str, fy_year: int) -> AnnualReportData | None:
        """
        Return cached AnnualReportData if it exists and is < 90 days old.
        Returns None if missing or stale.
        """
        try:
            with sqlite3.connect(self.db_path) as conn:
                row = conn.execute(
                    "SELECT parsed_at, pdf_path, data_json FROM pdf_cache "
                    "WHERE bse_code = ? AND fy_year = ?",
                    (bse_code, fy_year)
                ).fetchone()

            if row is None:
                return None

            parsed_at_str, pdf_path, data_json_str = row
            parsed_at = datetime.fromisoformat(parsed_at_str)
            age_days = (datetime.utcnow() - parsed_at).days

            if age_days > 90:
                logger.info(f"Cache stale ({age_days}d) for {bse_code}/{fy_year}")
                return None

            data_dict = json.loads(data_json_str)
            return AnnualReportData(**data_dict)

        except Exception as e:
            logger.warning(f"Cache read error for {bse_code}/{fy_year}: {e}")
            return None

    def _cache_set(self, data: AnnualReportData) -> None:
        """Persist AnnualReportData to SQLite cache."""
        try:
            data_dict = asdict(data)
            with sqlite3.connect(self.db_path) as conn:
                conn.execute(
                    """
                    INSERT OR REPLACE INTO pdf_cache
                        (bse_code, fy_year, parsed_at, pdf_path, data_json)
                    VALUES (?, ?, ?, ?, ?)
                    """,
                    (
                        data.bse_code,
                        data.report_year,
                        data.fetched_at,
                        data.pdf_path,
                        json.dumps(data_dict),
                    )
                )
                conn.commit()
            logger.info(f"Cached annual report data for {data.bse_code}/{data.report_year}")
        except Exception as e:
            logger.warning(f"Cache write error for {data.bse_code}: {e}")


# ── CLI helper ─────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import argparse

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(name)s: %(message)s"
    )

    parser = argparse.ArgumentParser(description="Parse BSE annual report PDF")
    parser.add_argument("ticker", help="NSE ticker symbol, e.g. TCS")
    parser.add_argument("bse_code", help="BSE scrip code, e.g. 532540")
    parser.add_argument("--fy-year", type=int, help="Financial year end, e.g. 2025")
    parser.add_argument("--db-path", default="/tmp/india_research.db")
    parser.add_argument("--pdf-cache-dir", default="/tmp/pdf_cache")
    args = parser.parse_args()

    ap = AnnualReportParser(
        db_path=args.db_path,
        pdf_cache_dir=args.pdf_cache_dir,
    )
    result = ap.parse_latest(args.ticker, args.bse_code, fy_year=args.fy_year)

    print(f"\n{'='*60}")
    print(f"Annual Report: {result.ticker} {result.fy_label}")
    print(f"{'='*60}")
    print(f"PDF path: {result.pdf_path or '(not downloaded)'}")
    print(f"Parse errors: {result.parse_errors or 'None'}")
    print(f"\nBusiness highlights ({len(result.business_highlights)}):")
    for h in result.business_highlights:
        print(f"  • {h}")
    print(f"\nKey risks ({len(result.key_risks_mentioned)}):")
    for r in result.key_risks_mentioned:
        print(f"  • {r}")
    if result.capex_guidance:
        print(f"\nCapex guidance: {result.capex_guidance}")
    if result.revenue_guidance:
        print(f"Revenue guidance: {result.revenue_guidance}")
    if result.headcount:
        print(f"Headcount: {result.headcount:,}")
    if result.attrition_pct:
        print(f"Attrition: {result.attrition_pct}%")
