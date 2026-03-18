"""
Global configuration for IndiaStockResearch platform.
All paths, constants, and credentials live here.
"""
import os

# ── Root ──────────────────────────────────────────────────────────────────────
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# ── Paths ─────────────────────────────────────────────────────────────────────
DATA_DIR          = os.path.join(ROOT, "data")
RAW_DATA_DIR      = os.path.join(DATA_DIR, "raw")
PROCESSED_DIR     = os.path.join(DATA_DIR, "processed")
SQLITE_DB_PATH    = os.path.join(PROCESSED_DIR, "cache.db")
REPORTS_DIR       = os.path.join(ROOT, "reports")
CHARTS_SUBDIR     = "charts"   # relative to each report's folder

# ── screener.in ───────────────────────────────────────────────────────────────
SCREENER_BASE_URL       = "https://www.screener.in"
SCREENER_LOGIN_URL      = f"{SCREENER_BASE_URL}/login/"
SCREENER_COMPANY_URL    = f"{SCREENER_BASE_URL}/company/{{ticker}}/consolidated/"
SCREENER_COMPANY_SA_URL = f"{SCREENER_BASE_URL}/company/{{ticker}}/"  # standalone fallback
SCREENER_CACHE_TTL_DAYS = 7
RATE_LIMIT_DELAY_SECS   = 2.0   # seconds between requests
MAX_RETRIES             = 3

# Credentials — set via env vars or override here
SCREENER_USERNAME = os.environ.get("SCREENER_USERNAME", "")
SCREENER_PASSWORD = os.environ.get("SCREENER_PASSWORD", "")

# ── yfinance ──────────────────────────────────────────────────────────────────
YF_DEFAULT_EXCHANGE = "NSE"   # "NSE" → .NS suffix, "BSE" → .BO suffix
PRICE_CACHE_TTL_DAYS = 1

# ── Report styling ────────────────────────────────────────────────────────────
FONT_NAME = "Times New Roman"
FONT_SIZE_BODY    = 9
FONT_SIZE_TABLE   = 8
FONT_SIZE_HEADING = 12
FONT_SIZE_TITLE   = 20

# Color palette (hex)
COLOR_NAVY  = "#0D2B55"
COLOR_GREEN = "#1A753F"
COLOR_GOLD  = "#B7950B"
COLOR_LGRAY = "#F2F2F2"
COLOR_DGRAY = "#D9D9D9"
COLOR_WHITE = "#FFFFFF"
COLOR_BLACK = "#000000"
COLOR_BLUE_LIGHT = "#D6E4F0"   # projection column shading

# Color palette (matplotlib RGB tuples, 0–1 scale)
NAVY_RGB  = (0x0D/255, 0x2B/255, 0x55/255)
GREEN_RGB = (0x1A/255, 0x75/255, 0x3F/255)
GOLD_RGB  = (0xB7/255, 0x95/255, 0x0B/255)
LGRAY_RGB = (0xF2/255, 0xF2/255, 0xF2/255)
DGRAY_RGB = (0xD9/255, 0xD9/255, 0xD9/255)

# Chart defaults
CHART_DPI    = 300
CHART_WIDTH  = 12    # inches (full-width)
CHART_HEIGHT = 5     # inches

# ── Financial constants ────────────────────────────────────────────────────────
CR_TO_INR    = 1e7    # 1 Crore = 10,000,000
INR_TO_CR    = 1e-7
LAKH_TO_CR   = 0.01
INDIA_FY_END_MONTH = 3   # March

# ── Stock universe ─────────────────────────────────────────────────────────────
STOCK_UNIVERSE_CSV = os.path.join(ROOT, "config", "stock_universe.csv")
