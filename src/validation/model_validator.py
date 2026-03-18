"""
Data quality validator for report_json.

Runs after run_data() to catch bad scraped data before charts/report generation.
All checks return ValidationResult objects; no exceptions are raised.

Usage:
    from src.validation.model_validator import ModelValidator
    results = ModelValidator.validate(report_json)
    if results.has_errors:
        for r in results.errors:
            print(r)
"""

import logging
from dataclasses import dataclass, field
from typing import Optional

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Result types
# ---------------------------------------------------------------------------

@dataclass
class CheckResult:
    """Result of a single validation check."""
    name: str
    passed: bool
    severity: str           # "error" | "warning" | "info"
    message: str
    detail: str = ""

    def __str__(self):
        icon = "✓" if self.passed else ("✗" if self.severity == "error" else "⚠")
        return f"[{self.severity.upper():7s}] {icon} {self.name}: {self.message}"


@dataclass
class ValidationReport:
    """Aggregated results from all checks."""
    ticker: str
    model_type: str
    checks: list[CheckResult] = field(default_factory=list)

    @property
    def errors(self) -> list[CheckResult]:
        return [c for c in self.checks if not c.passed and c.severity == "error"]

    @property
    def warnings(self) -> list[CheckResult]:
        return [c for c in self.checks if not c.passed and c.severity == "warning"]

    @property
    def has_errors(self) -> bool:
        return len(self.errors) > 0

    @property
    def has_warnings(self) -> bool:
        return len(self.warnings) > 0

    def summary(self) -> str:
        total  = len(self.checks)
        passed = sum(1 for c in self.checks if c.passed)
        errs   = len(self.errors)
        warns  = len(self.warnings)
        return (f"{self.ticker} | {passed}/{total} checks passed | "
                f"{errs} errors | {warns} warnings")

    def log_all(self):
        logger.info(f"=== Validation: {self.summary()} ===")
        for c in self.checks:
            fn = logger.info if c.passed else (
                logger.error if c.severity == "error" else logger.warning)
            fn(str(c))
            if c.detail and not c.passed:
                fn(f"         Detail: {c.detail}")


# ---------------------------------------------------------------------------
# Validator
# ---------------------------------------------------------------------------

class ModelValidator:
    """
    Static methods — no state.  Call ModelValidator.validate(report_json).
    """

    @staticmethod
    def validate(report_json: dict) -> ValidationReport:
        ticker     = report_json.get("meta", {}).get("ticker", "UNKNOWN")
        model_type = report_json.get("meta", {}).get("model_type", "generic")
        report     = ValidationReport(ticker=ticker, model_type=model_type)

        fin    = report_json.get("financials", {})
        is_    = fin.get("income_statement", {})
        bs_    = fin.get("balance_sheet", {})
        cf_    = fin.get("cash_flow", {})
        rat_   = report_json.get("ratios", {})
        fy_labels = fin.get("fy_labels", [])
        actuals_end_idx = fin.get("actuals_end_idx", len(fy_labels) - 1)

        # Only validate actuals (not projections)
        actual_slice = slice(0, actuals_end_idx + 1)

        def _actuals(key_dict, key):
            vals = key_dict.get(key, [])
            if isinstance(vals, list):
                return [v for v in vals[actual_slice] if v is not None]
            return []

        checks = report.checks

        # ── 1. Data completeness ──────────────────────────────────────────────

        # 1a: Revenue series not empty
        rev = _actuals(is_, "net_revenue")
        checks.append(CheckResult(
            name="revenue_data_present",
            passed=len(rev) >= 3,
            severity="error",
            message=f"Net revenue data: {len(rev)} actuals" if rev else "No net revenue data",
        ))

        # 1b: PAT series not empty
        pat = _actuals(is_, "pat")
        checks.append(CheckResult(
            name="pat_data_present",
            passed=len(pat) >= 3,
            severity="error",
            message=f"PAT data: {len(pat)} actuals" if pat else "No PAT data",
        ))

        # 1c: At least 5 FY labels
        checks.append(CheckResult(
            name="fy_labels_sufficient",
            passed=len(fy_labels) >= 5,
            severity="warning",
            message=f"FY labels: {fy_labels}" if fy_labels else "No FY labels",
        ))

        # ── 2. Revenue continuity ─────────────────────────────────────────────
        if len(rev) >= 2:
            swings = []
            for i in range(1, len(rev)):
                if rev[i - 1] > 0:
                    yoy = abs((rev[i] - rev[i - 1]) / rev[i - 1])
                    if yoy > 0.5:
                        fy = fy_labels[i] if i < len(fy_labels) else f"idx {i}"
                        swings.append(f"{fy}: {yoy*100:.0f}% swing")
            checks.append(CheckResult(
                name="revenue_continuity",
                passed=len(swings) == 0,
                severity="warning",
                message=("Revenue has >50% YoY swing(s) — check excise/Ind AS normalizer"
                         if swings else "Revenue continuity OK"),
                detail=", ".join(swings),
            ))

        # ── 3. Margin reasonableness ──────────────────────────────────────────
        if model_type not in ("bank",):   # banks have no EBITDA
            ebitda_margins = [v for v in rat_.get("ebitda_margin", [])[actual_slice]
                              if v is not None]
            if ebitda_margins:
                bad = [f"{v:.1f}%" for v in ebitda_margins if not (0 <= v <= 80)]
                checks.append(CheckResult(
                    name="ebitda_margin_range",
                    passed=len(bad) == 0,
                    severity="warning",
                    message=("EBITDA margins out of range 0-80%"
                             if bad else "EBITDA margins in range"),
                    detail=f"Bad values: {bad}",
                ))

            pat_margins = [v for v in rat_.get("pat_margin", [])[actual_slice]
                           if v is not None]
            if pat_margins:
                bad = [f"{v:.1f}%" for v in pat_margins if not (-20 <= v <= 50)]
                checks.append(CheckResult(
                    name="pat_margin_range",
                    passed=len(bad) == 0,
                    severity="warning",
                    message=("PAT margins out of range -20% to 50%"
                             if bad else "PAT margins in range"),
                    detail=f"Bad values: {bad}",
                ))

        # ── 4. Balance sheet balance check ────────────────────────────────────
        total_assets  = is_["total_assets"] if "total_assets" in is_ else bs_.get("total_assets", [])
        total_equity  = bs_.get("total_equity", [])
        total_debt_   = bs_.get("total_debt", [])

        # Approx check: Total Assets ≈ Total Equity + Total Debt (within 20% tolerance)
        # We don't have explicit 'total_liabilities', so use equity + debt as proxy
        if total_assets and total_equity and total_debt_:
            n = min(len(total_assets), len(total_equity), len(total_debt_), actuals_end_idx + 1)
            bs_breaches = []
            for i in range(n):
                a  = total_assets[i]
                e  = total_equity[i]
                d  = total_debt_[i]
                if a is None or e is None or d is None:
                    continue
                if a <= 0:
                    continue
                proxy_liab = e + d
                ratio = abs(a - proxy_liab) / a
                if ratio > 0.30:   # >30% mismatch = likely bad scrape
                    fy = fy_labels[i] if i < len(fy_labels) else f"idx {i}"
                    bs_breaches.append(
                        f"{fy}: Assets={a:,.0f} vs Equity+Debt={proxy_liab:,.0f} "
                        f"({ratio*100:.0f}% gap)"
                    )
            checks.append(CheckResult(
                name="balance_sheet_proxy_check",
                passed=len(bs_breaches) == 0,
                severity="warning",
                message=("Assets vs (Equity+Debt) mismatch >30% — check BS scrape"
                         if bs_breaches else "Balance sheet proxy check OK"),
                detail="; ".join(bs_breaches[:3]),
            ))

        # ── 5. Cash flow tie-out ──────────────────────────────────────────────
        cfo_  = [v for v in cf_.get("cfo", [])[actual_slice] if v is not None]
        capex_= [v for v in cf_.get("capex", [])[actual_slice] if v is not None]
        fcf_  = [v for v in cf_.get("fcf", [])[actual_slice] if v is not None]

        # FCF = CFO + Capex (capex stored negative in screener.in; verify sign)
        if cfo_ and capex_ and fcf_:
            n = min(len(cfo_), len(capex_), len(fcf_))
            fcf_mismatches = []
            for i in range(n):
                # capex may be stored as negative (outflow) or positive
                implied_fcf_neg = cfo_[i] + capex_[i]   # if capex is negative
                implied_fcf_pos = cfo_[i] - capex_[i]   # if capex is positive
                actual_fcf      = fcf_[i]
                if (abs(implied_fcf_neg - actual_fcf) > max(abs(actual_fcf) * 0.2, 10)
                        and abs(implied_fcf_pos - actual_fcf) > max(abs(actual_fcf) * 0.2, 10)):
                    fy = fy_labels[i] if i < len(fy_labels) else f"idx {i}"
                    fcf_mismatches.append(
                        f"{fy}: FCF={actual_fcf:,.0f} CFO={cfo_[i]:,.0f} Capex={capex_[i]:,.0f}"
                    )
            checks.append(CheckResult(
                name="fcf_tieout",
                passed=len(fcf_mismatches) == 0,
                severity="warning",
                message=("FCF ≠ CFO ± Capex (>20% mismatch)"
                         if fcf_mismatches else "FCF tie-out OK"),
                detail="; ".join(fcf_mismatches[:3]),
            ))

        # ── 6. Sector-specific checks ─────────────────────────────────────────

        if model_type == "bank":
            ModelValidator._check_bank(report_json, checks, fy_labels, actual_slice)

        elif model_type in ("pharma",):
            ModelValidator._check_pharma(report_json, checks, fy_labels, actual_slice)

        elif model_type in ("metals",):
            ModelValidator._check_metals(report_json, checks, fy_labels, actual_slice)

        # ── 7. Market data sanity ─────────────────────────────────────────────
        mkt = report_json.get("market_data", {})
        cmp = mkt.get("cmp")
        mcap = mkt.get("market_cap_cr")

        checks.append(CheckResult(
            name="cmp_positive",
            passed=bool(cmp and cmp > 0),
            severity="error",
            message=f"CMP: ₹{cmp}" if cmp else "CMP missing or zero",
        ))

        if cmp and mcap and mcap > 0:
            # P/E sanity (if revenue available)
            pe = mkt.get("pe_ttm")
            if pe is not None:
                checks.append(CheckResult(
                    name="pe_range",
                    passed=(0 < pe < 200),
                    severity="warning",
                    message=f"P/E TTM = {pe:.1f}" + (
                        "" if 0 < pe < 200 else " — unusually high/negative"),
                ))

        # ── 8. ROCE > 0 in recent years ──────────────────────────────────────
        if model_type not in ("bank",):
            roce = [v for v in rat_.get("roce", [])[max(0, actuals_end_idx-2):actuals_end_idx+1]
                    if v is not None]
            if roce:
                checks.append(CheckResult(
                    name="roce_positive_recent",
                    passed=any(v > 0 for v in roce),
                    severity="warning",
                    message=f"Recent ROCE: {[f'{v:.1f}%' for v in roce]}",
                ))

        report.log_all()
        return report

    # ── Sector-specific check helpers ─────────────────────────────────────────

    @staticmethod
    def _check_bank(report_json, checks, fy_labels, actual_slice):
        is_  = report_json["financials"]["income_statement"]
        nim  = [v for v in is_.get("nim", [])[actual_slice] if v is not None]
        gnpa = [v for v in is_.get("gnpa_pct", [])[actual_slice] if v is not None]
        car  = [v for v in is_.get("car", [])[actual_slice] if v is not None]

        if nim:
            bad = [f"{v:.2f}%" for v in nim if not (0.5 <= v <= 8.0)]
            checks.append(CheckResult(
                name="bank_nim_range",
                passed=len(bad) == 0,
                severity="warning",
                message=f"NIM range check: {[f'{v:.2f}%' for v in nim[-3:]]}",
                detail=f"Out of range [0.5%, 8%]: {bad}" if bad else "",
            ))

        if gnpa:
            bad = [f"{v:.1f}%" for v in gnpa if not (0 <= v <= 30)]
            checks.append(CheckResult(
                name="bank_gnpa_range",
                passed=len(bad) == 0,
                severity="warning",
                message=f"GNPA% range check: {[f'{v:.1f}%' for v in gnpa[-3:]]}",
                detail=f"Out of range [0%, 30%]: {bad}" if bad else "",
            ))

        if car:
            bad = [f"{v:.1f}%" for v in car if not (8 <= v <= 30)]
            checks.append(CheckResult(
                name="bank_car_range",
                passed=len(bad) == 0,
                severity="warning",
                message=f"CAR range check: {[f'{v:.1f}%' for v in car[-3:]]}",
                detail=f"Out of range [8%, 30%]: {bad}" if bad else "",
            ))

    @staticmethod
    def _check_pharma(report_json, checks, fy_labels, actual_slice):
        is_  = report_json["financials"]["income_statement"]
        rnd  = [v for v in is_.get("rnd_spend_pct", [])[actual_slice] if v is not None]

        if rnd:
            # R&D typically 5-15% for Indian pharma
            bad = [f"{v:.1f}%" for v in rnd if not (0 <= v <= 25)]
            checks.append(CheckResult(
                name="pharma_rnd_range",
                passed=len(bad) == 0,
                severity="warning",
                message=f"R&D % range check: {[f'{v:.1f}%' for v in rnd[-3:]]}",
                detail=f"Out of expected range [0%, 25%]: {bad}" if bad else "",
            ))

    @staticmethod
    def _check_metals(report_json, checks, fy_labels, actual_slice):
        is_  = report_json["financials"]["income_statement"]
        ebitda_per_tonne = [v for v in is_.get("ebitda_per_tonne", [])[actual_slice]
                            if v is not None]

        if ebitda_per_tonne:
            bad = [f"₹{v:,.0f}" for v in ebitda_per_tonne if not (0 <= v <= 50000)]
            checks.append(CheckResult(
                name="metals_ebitda_per_tonne_range",
                passed=len(bad) == 0,
                severity="warning",
                message=f"EBITDA/tonne range check (expect ₹2,000-₹25,000)",
                detail=f"Out of expected range: {bad}" if bad else "",
            ))
