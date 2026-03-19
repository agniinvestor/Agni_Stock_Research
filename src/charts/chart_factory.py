"""
Chart factory: all 35 standard charts as typed functions.
Each function takes ChartData + output_path, returns the saved file path.

All charts:
- Use Times New Roman font
- Are saved at 300 DPI as PNG
- Have source note at bottom
- Have projection separator when actuals_end_idx is set
- Are company-agnostic (company name from ChartData.metadata)
"""

from __future__ import annotations

import sys
import os
import math
import warnings

import numpy as np
import matplotlib
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import matplotlib.ticker as mticker
from matplotlib.colors import LinearSegmentedColormap
from matplotlib.lines import Line2D

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", ".."))

# ---------------------------------------------------------------------------
# Style imports — try package path first, fall back to local /tmp copy
# ---------------------------------------------------------------------------
try:
    from src.charts.chart_style import (
        ChartData,
        apply_style,
        add_source_note,
        add_projection_separator,
        format_crore_axis,
        save_chart,
        bar_colors,
        get_fig_ax,
        add_data_labels,
        NAVY_RGB,
        GREEN_RGB,
        GOLD_RGB,
        LGRAY_RGB,
        DGRAY_RGB,
        NAVY_HEX,
        GREEN_HEX,
        GOLD_HEX,
        LGRAY_HEX,
        DGRAY_HEX,
        color_hex_to_rgb,
    )
except ModuleNotFoundError:
    # Running directly from /tmp — import the sibling file
    _here = os.path.dirname(os.path.abspath(__file__))
    sys.path.insert(0, _here)
    from chart_style import (  # type: ignore
        ChartData,
        apply_style,
        add_source_note,
        add_projection_separator,
        format_crore_axis,
        save_chart,
        bar_colors,
        get_fig_ax,
        add_data_labels,
        NAVY_RGB,
        GREEN_RGB,
        GOLD_RGB,
        LGRAY_RGB,
        DGRAY_RGB,
        NAVY_HEX,
        GREEN_HEX,
        GOLD_HEX,
        LGRAY_HEX,
        DGRAY_HEX,
        color_hex_to_rgb,
    )

try:
    from config.settings import FONT_NAME  # type: ignore
except ModuleNotFoundError:
    FONT_NAME = "Times New Roman"

# Ensure style is always active
apply_style()

# ---------------------------------------------------------------------------
# Internal utilities
# ---------------------------------------------------------------------------

_SOURCE_DEFAULT = "Source: Company filings, screener.in"

_SEGMENT_COLORS = [NAVY_RGB, GREEN_RGB, GOLD_RGB, LGRAY_RGB, DGRAY_RGB]


def _source(data: ChartData) -> str:
    return data.metadata.get("source", _SOURCE_DEFAULT)


def _company(data: ChartData) -> str:
    return data.metadata.get("company_name", "Company")


def _safe_vals(lst: list) -> list:
    """Replace None with np.nan so matplotlib skips those points."""
    return [np.nan if v is None else v for v in lst]


def _bar_fill_colors(n_total: int, actuals_end_idx: int) -> list:
    """NAVY for actuals, GOLD for estimates."""
    n_act = actuals_end_idx + 1 if actuals_end_idx >= 0 else n_total
    n_est = n_total - n_act
    return [NAVY_RGB] * n_act + [GOLD_RGB] * n_est


def _x_positions(n: int) -> np.ndarray:
    return np.arange(n, dtype=float)


# ===========================================================================
# CATEGORY 1: Revenue & Earnings  (charts 01–04)
# ===========================================================================


def chart_stock_price_performance(data: ChartData, output_path: str) -> str:
    """Line chart: stock price vs Nifty 50 rebased to 100.

    data.values needs: 'stock_price', 'nifty50' (both as index = 100 at start).
    Marks 52-week high and low with annotations.
    """
    fig, ax = get_fig_ax()
    company = _company(data)

    stock = _safe_vals(data.values.get("stock_price", []))
    nifty = _safe_vals(data.values.get("nifty50", []))
    labels = data.fy_labels
    x = _x_positions(len(labels))

    ax.plot(x, stock, color=NAVY_RGB, linewidth=2.0, label=company, zorder=4)
    ax.plot(x, nifty, color=GOLD_RGB, linewidth=1.5, linestyle="--",
            label="Nifty 50", zorder=3)

    # 52-week high/low annotations on stock line
    stock_clean = [(i, v) for i, v in enumerate(stock) if not math.isnan(v)]
    if stock_clean:
        hi_i, hi_v = max(stock_clean, key=lambda t: t[1])
        lo_i, lo_v = min(stock_clean, key=lambda t: t[1])
        ax.annotate(
            f"52W High\n{hi_v:.0f}",
            xy=(x[hi_i], hi_v),
            xytext=(0, 10),
            textcoords="offset points",
            ha="center",
            fontsize=6,
            color=GREEN_HEX,
            arrowprops=dict(arrowstyle="-", color=GREEN_HEX, lw=0.8),
        )
        ax.annotate(
            f"52W Low\n{lo_v:.0f}",
            xy=(x[lo_i], lo_v),
            xytext=(0, -14),
            textcoords="offset points",
            ha="center",
            fontsize=6,
            color="#CC0000",
            arrowprops=dict(arrowstyle="-", color="#CC0000", lw=0.8),
        )

    ax.set_xticks(x)
    ax.set_xticklabels(labels, rotation=45, ha="right")
    ax.set_ylabel("Indexed Price (Base = 100)")
    ax.set_title(f"{company} — Stock Price vs Nifty 50 (Rebased to 100)")
    ax.axhline(100, color=DGRAY_RGB, linewidth=0.6, linestyle=":")
    ax.legend(loc="upper left")
    add_source_note(ax, _source(data))
    save_chart(fig, output_path)
    return output_path


def chart_revenue_pat_trend(data: ChartData, output_path: str) -> str:
    """Dual-axis bar+line: Net Revenue bars (left axis), PAT line (right axis).

    data.values needs: 'net_revenue', 'pat'.
    Adds YoY growth % labels above each PAT point.
    """
    fig, ax1 = get_fig_ax()
    company = _company(data)

    labels = data.fy_labels
    n = len(labels)
    x = _x_positions(n)
    rev = _safe_vals(data.values.get("net_revenue", [np.nan] * n))
    pat = _safe_vals(data.values.get("pat", [np.nan] * n))

    colors = _bar_fill_colors(n, data.actuals_end_idx)
    bars = ax1.bar(x, rev, color=colors, width=0.55, zorder=2, label="Net Revenue")
    ax1.set_ylabel("Net Revenue (\u20b9 Cr)")
    format_crore_axis(ax1, "y")

    ax2 = ax1.twinx()
    ax2.plot(x, pat, color=GREEN_RGB, linewidth=2.0, marker="o",
             markersize=5, label="PAT", zorder=4)
    ax2.set_ylabel("PAT (\u20b9 Cr)")
    format_crore_axis(ax2, "y")

    # YoY growth labels
    for i in range(1, n):
        if not math.isnan(pat[i]) and not math.isnan(pat[i - 1]) and pat[i - 1] != 0:
            yoy = (pat[i] - pat[i - 1]) / abs(pat[i - 1]) * 100
            ax2.text(
                x[i], pat[i],
                f"{yoy:+.1f}%",
                ha="center", va="bottom",
                fontsize=6, color=GREEN_HEX,
            )

    ax1.set_xticks(x)
    ax1.set_xticklabels(labels, rotation=45, ha="right")
    ax1.set_title(f"{company} — Revenue & PAT Trend")

    lines1, labs1 = ax1.get_legend_handles_labels()
    lines2, labs2 = ax2.get_legend_handles_labels()
    ax1.legend(lines1 + lines2, labs1 + labs2, loc="upper left", fontsize=7)

    add_source_note(ax1, _source(data))
    if data.actuals_end_idx >= 0:
        add_projection_separator(ax1, data.actuals_end_idx, labels)

    save_chart(fig, output_path)
    return output_path


def chart_revenue_by_segment(data: ChartData, output_path: str) -> str:
    """Stacked area chart: revenue breakdown by segment over time.

    data.values needs: dict of {segment_name: values_list}.
    """
    fig, ax = get_fig_ax()
    company = _company(data)
    labels = data.fy_labels
    x = _x_positions(len(labels))

    segments = {k: _safe_vals(v) for k, v in data.values.items()}
    palette = _SEGMENT_COLORS

    keys = list(segments.keys())
    arrays = [np.array(segments[k]) for k in keys]

    ax.stackplot(
        x,
        *arrays,
        labels=keys,
        colors=palette[: len(keys)],
        alpha=0.85,
    )

    ax.set_xticks(x)
    ax.set_xticklabels(labels, rotation=45, ha="right")
    ax.set_ylabel("Revenue (\u20b9 Cr)")
    ax.set_title(f"{company} — Revenue by Segment")
    format_crore_axis(ax, "y")
    ax.legend(loc="upper left", fontsize=7, reverse=True)

    if data.actuals_end_idx >= 0:
        add_projection_separator(ax, data.actuals_end_idx, labels)
    add_source_note(ax, _source(data))
    save_chart(fig, output_path)
    return output_path


def chart_revenue_by_geography(data: ChartData, output_path: str) -> str:
    """Stacked bar: domestic vs export revenue with % labels inside bars.

    data.values needs: 'domestic', 'export' (or more geography keys).
    """
    fig, ax = get_fig_ax()
    company = _company(data)
    labels = data.fy_labels
    n = len(labels)
    x = _x_positions(n)

    geo_keys = list(data.values.keys())
    palette = [NAVY_RGB, GREEN_RGB, GOLD_RGB, LGRAY_RGB, DGRAY_RGB]
    bottoms = np.zeros(n)

    for idx, key in enumerate(geo_keys):
        vals = np.array(_safe_vals(data.values[key]))
        np.nan_to_num(vals, copy=False)
        bars = ax.bar(x, vals, bottom=bottoms, color=palette[idx % len(palette)],
                      width=0.6, label=key.title(), zorder=2)
        # % labels inside each bar segment
        totals = np.nansum(
            [np.array(_safe_vals(data.values[k])) for k in geo_keys], axis=0
        )
        for i, (bar, v, tot) in enumerate(zip(bars, vals, totals)):
            if tot > 0 and v > 0:
                pct = v / tot * 100
                ax.text(
                    bar.get_x() + bar.get_width() / 2,
                    bottoms[i] + v / 2,
                    f"{pct:.0f}%",
                    ha="center", va="center",
                    fontsize=6, color="white", fontweight="bold",
                )
        bottoms += np.nan_to_num(vals)

    ax.set_xticks(x)
    ax.set_xticklabels(labels, rotation=45, ha="right")
    ax.set_ylabel("Revenue (\u20b9 Cr)")
    ax.set_title(f"{company} — Revenue by Geography")
    format_crore_axis(ax, "y")
    ax.legend(loc="upper left", fontsize=7)

    if data.actuals_end_idx >= 0:
        add_projection_separator(ax, data.actuals_end_idx, labels)
    add_source_note(ax, _source(data))
    save_chart(fig, output_path)
    return output_path


# ===========================================================================
# CATEGORY 2: Margins  (charts 10–11)
# ===========================================================================


def chart_margin_progression(data: ChartData, output_path: str) -> str:
    """Multi-line chart: Gross Margin %, EBITDA Margin %, PAT Margin % over time.

    data.values needs: 'gross_margin', 'ebitda_margin', 'pat_margin'.
    Uses GREEN, NAVY, GOLD respectively. Marks latest value at line end.
    """
    fig, ax = get_fig_ax()
    company = _company(data)
    labels = data.fy_labels
    x = _x_positions(len(labels))

    series_cfg = [
        ("gross_margin",  "Gross Margin %",  GREEN_RGB, "o"),
        ("ebitda_margin", "EBITDA Margin %", NAVY_RGB,  "s"),
        ("pat_margin",    "PAT Margin %",    GOLD_RGB,  "^"),
    ]

    for key, label, color, marker in series_cfg:
        vals = _safe_vals(data.values.get(key, []))
        if not vals:
            continue
        ax.plot(x[: len(vals)], vals, color=color, linewidth=2.0,
                marker=marker, markersize=5, label=label, zorder=4)
        # Mark latest non-nan value at line end
        clean = [(i, v) for i, v in enumerate(vals) if not math.isnan(v)]
        if clean:
            li, lv = clean[-1]
            ax.text(
                x[li] + 0.05, lv,
                f"{lv:.1f}%",
                fontsize=6, va="center", color=color,
            )

    ax.set_xticks(x)
    ax.set_xticklabels(labels, rotation=45, ha="right")
    ax.set_ylabel("Margin (%)")
    ax.yaxis.set_major_formatter(mticker.FormatStrFormatter("%.1f%%"))
    ax.set_title(f"{company} — Margin Progression")
    ax.legend(loc="upper left", fontsize=7)

    if data.actuals_end_idx >= 0:
        add_projection_separator(ax, data.actuals_end_idx, labels)
    add_source_note(ax, _source(data))
    save_chart(fig, output_path)
    return output_path


def chart_segment_ebit_margins(data: ChartData, output_path: str) -> str:
    """Grouped bar chart: EBIT margin by segment for last 5 years.

    data.values needs segment margin keys.
    5 groups (one per year), N bars per group (one per segment).
    """
    fig, ax = get_fig_ax()
    company = _company(data)
    labels = data.fy_labels[-5:] if len(data.fy_labels) >= 5 else data.fy_labels
    n_years = len(labels)
    seg_keys = list(data.values.keys())
    n_segs = len(seg_keys)

    width = 0.8 / max(n_segs, 1)
    x = _x_positions(n_years)
    palette = _SEGMENT_COLORS

    for s_idx, seg in enumerate(seg_keys):
        vals_full = _safe_vals(data.values[seg])
        vals = vals_full[-n_years:]
        offset = (s_idx - (n_segs - 1) / 2) * width
        bars = ax.bar(
            x + offset, vals,
            width=width * 0.9,
            color=palette[s_idx % len(palette)],
            label=seg,
            zorder=2,
        )
        add_data_labels(ax, bars, fmt="{:.1f}%", fontsize=6)

    ax.set_xticks(x)
    ax.set_xticklabels(labels, rotation=45, ha="right")
    ax.set_ylabel("EBIT Margin (%)")
    ax.yaxis.set_major_formatter(mticker.FormatStrFormatter("%.1f%%"))
    ax.set_title(f"{company} — Segment EBIT Margins")
    ax.legend(loc="upper left", fontsize=7)
    add_source_note(ax, _source(data))
    save_chart(fig, output_path)
    return output_path


# ===========================================================================
# CATEGORY 3: Cash Flow  (charts 12, 19, 23)
# ===========================================================================


def chart_fcf_trend(data: ChartData, output_path: str) -> str:
    """Grouped bar: CFO vs Capex vs FCF. FCF bars in GREEN when positive.

    data.values needs: 'cfo', 'capex', 'fcf'.
    """
    fig, ax = get_fig_ax()
    company = _company(data)
    labels = data.fy_labels
    n = len(labels)
    x = _x_positions(n)
    width = 0.26

    cfo  = _safe_vals(data.values.get("cfo",   [np.nan] * n))
    capex = _safe_vals(data.values.get("capex", [np.nan] * n))
    fcf  = _safe_vals(data.values.get("fcf",   [np.nan] * n))

    ax.bar(x - width, cfo,  width=width, color=NAVY_RGB,  label="CFO",   zorder=2)
    ax.bar(x,         capex, width=width, color=GOLD_RGB,  label="Capex", zorder=2)

    # FCF: green if positive, red if negative
    fcf_colors = [
        GREEN_RGB if (not math.isnan(v) and v >= 0) else (0.8, 0.1, 0.1)
        for v in fcf
    ]
    for i, (xi, fv, fc) in enumerate(zip(x + width, fcf, fcf_colors)):
        if not math.isnan(fv):
            ax.bar(xi, fv, width=width, color=fc, zorder=2,
                   label="FCF" if i == 0 else "")

    ax.axhline(0, color=DGRAY_HEX, linewidth=0.7)
    ax.set_xticks(x)
    ax.set_xticklabels(labels, rotation=45, ha="right")
    ax.set_ylabel("(\u20b9 Cr)")
    format_crore_axis(ax, "y")
    ax.set_title(f"{company} — Free Cash Flow Trend")
    ax.legend(loc="upper left", fontsize=7)

    if data.actuals_end_idx >= 0:
        add_projection_separator(ax, data.actuals_end_idx, labels)
    add_source_note(ax, _source(data))
    save_chart(fig, output_path)
    return output_path


def chart_dividend_history(data: ChartData, output_path: str) -> str:
    """Bar (DPS) + line (payout ratio %).

    data.values needs: 'dps', 'payout_ratio'.
    """
    fig, ax1 = get_fig_ax()
    company = _company(data)
    labels = data.fy_labels
    n = len(labels)
    x = _x_positions(n)

    dps    = _safe_vals(data.values.get("dps",          [np.nan] * n))
    payout = _safe_vals(data.values.get("payout_ratio", [np.nan] * n))

    colors = _bar_fill_colors(n, data.actuals_end_idx)
    bars = ax1.bar(x, dps, color=colors, width=0.55, zorder=2, label="DPS (\u20b9)")
    add_data_labels(ax1, bars, fmt="{:.1f}", fontsize=6)
    ax1.set_ylabel("Dividend Per Share (\u20b9)")

    ax2 = ax1.twinx()
    ax2.plot(x, payout, color=GOLD_RGB, linewidth=2.0, marker="D",
             markersize=5, label="Payout Ratio %", zorder=4)
    ax2.set_ylabel("Payout Ratio (%)")
    ax2.yaxis.set_major_formatter(mticker.FormatStrFormatter("%.1f%%"))

    ax1.set_xticks(x)
    ax1.set_xticklabels(labels, rotation=45, ha="right")
    ax1.set_title(f"{company} — Dividend History")

    lines1, labs1 = ax1.get_legend_handles_labels()
    lines2, labs2 = ax2.get_legend_handles_labels()
    ax1.legend(lines1 + lines2, labs1 + labs2, loc="upper left", fontsize=7)

    add_source_note(ax1, _source(data))
    save_chart(fig, output_path)
    return output_path


def chart_capex_vs_net_cash(data: ChartData, output_path: str) -> str:
    """Bar (capex) + line (net cash position). Net cash negative = debt.

    data.values needs: 'capex', 'net_cash'.
    """
    fig, ax1 = get_fig_ax()
    company = _company(data)
    labels = data.fy_labels
    n = len(labels)
    x = _x_positions(n)

    capex    = _safe_vals(data.values.get("capex",    [np.nan] * n))
    net_cash = _safe_vals(data.values.get("net_cash", [np.nan] * n))

    colors = _bar_fill_colors(n, data.actuals_end_idx)
    ax1.bar(x, capex, color=colors, width=0.55, zorder=2, label="Capex (\u20b9 Cr)")
    ax1.set_ylabel("Capex (\u20b9 Cr)")
    format_crore_axis(ax1, "y")

    ax2 = ax1.twinx()
    ax2.plot(x, net_cash, color=GREEN_RGB, linewidth=2.0, marker="o",
             markersize=5, label="Net Cash (\u20b9 Cr)", zorder=4)
    ax2.axhline(0, color=DGRAY_HEX, linewidth=0.6, linestyle=":")
    ax2.set_ylabel("Net Cash / (Debt) (\u20b9 Cr)")
    format_crore_axis(ax2, "y")

    ax1.set_xticks(x)
    ax1.set_xticklabels(labels, rotation=45, ha="right")
    ax1.set_title(f"{company} — Capex vs Net Cash Position")

    lines1, labs1 = ax1.get_legend_handles_labels()
    lines2, labs2 = ax2.get_legend_handles_labels()
    ax1.legend(lines1 + lines2, labs1 + labs2, loc="upper left", fontsize=7)

    if data.actuals_end_idx >= 0:
        add_projection_separator(ax1, data.actuals_end_idx, labels)
    add_source_note(ax1, _source(data))
    save_chart(fig, output_path)
    return output_path


# ===========================================================================
# CATEGORY 4: Balance Sheet & Returns  (charts 22, 25, 26)
# ===========================================================================


def chart_balance_sheet_composition(data: ChartData, output_path: str) -> str:
    """Stacked bar: Assets side (Fixed Assets, Investments, Working Capital, Cash).

    data.values needs: 'fixed_assets', 'investments', 'working_capital', 'cash'.
    """
    fig, ax = get_fig_ax()
    company = _company(data)
    labels = data.fy_labels
    n = len(labels)
    x = _x_positions(n)

    asset_keys = ["fixed_assets", "investments", "working_capital", "cash"]
    asset_labels = ["Fixed Assets", "Investments", "Working Capital", "Cash"]
    palette = [NAVY_RGB, GREEN_RGB, GOLD_RGB, LGRAY_RGB]

    bottoms = np.zeros(n)
    for key, lbl, color in zip(asset_keys, asset_labels, palette):
        vals = np.nan_to_num(np.array(_safe_vals(data.values.get(key, [0.0] * n))))
        ax.bar(x, vals, bottom=bottoms, color=color, width=0.6,
               label=lbl, zorder=2)
        bottoms += vals

    ax.set_xticks(x)
    ax.set_xticklabels(labels, rotation=45, ha="right")
    ax.set_ylabel("(\u20b9 Cr)")
    format_crore_axis(ax, "y")
    ax.set_title(f"{company} — Balance Sheet Composition (Assets)")
    ax.legend(loc="upper left", fontsize=7)

    if data.actuals_end_idx >= 0:
        add_projection_separator(ax, data.actuals_end_idx, labels)
    add_source_note(ax, _source(data))
    save_chart(fig, output_path)
    return output_path


def chart_working_capital_days(data: ChartData, output_path: str) -> str:
    """Line chart: DSO, DIO, DPO, CCC days over time.

    data.values needs: 'dso', 'dio', 'dpo', 'ccc'.
    CCC in NAVY (bold), others in lighter colors.
    """
    fig, ax = get_fig_ax()
    company = _company(data)
    labels = data.fy_labels
    x = _x_positions(len(labels))

    series_cfg = [
        ("dso", "DSO",  (0.4, 0.6, 0.8), 1.2, "--", "o"),
        ("dio", "DIO",  (0.4, 0.7, 0.5), 1.2, "--", "s"),
        ("dpo", "DPO",  GOLD_RGB,         1.2, "--", "^"),
        ("ccc", "CCC",  NAVY_RGB,         2.2, "-",  "D"),
    ]

    for key, label, color, lw, ls, marker in series_cfg:
        vals = _safe_vals(data.values.get(key, []))
        if not vals:
            continue
        ax.plot(x[: len(vals)], vals, color=color, linewidth=lw,
                linestyle=ls, marker=marker, markersize=5,
                label=label, zorder=4 if key == "ccc" else 3)

    ax.axhline(0, color=DGRAY_HEX, linewidth=0.6, linestyle=":")
    ax.set_xticks(x)
    ax.set_xticklabels(labels, rotation=45, ha="right")
    ax.set_ylabel("Days")
    ax.set_title(f"{company} — Working Capital Days")
    ax.legend(loc="upper right", fontsize=7)

    add_source_note(ax, _source(data))
    save_chart(fig, output_path)
    return output_path


def chart_roce_roe_vs_wacc(data: ChartData, output_path: str) -> str:
    """Line chart: ROCE %, ROE %, WACC % on same axis.

    data.values needs: 'roce', 'roe', 'wacc'.
    Shades area between ROCE and WACC — green if ROCE > WACC, red if below.
    """
    fig, ax = get_fig_ax()
    company = _company(data)
    labels = data.fy_labels
    x = _x_positions(len(labels))

    roce = np.array(_safe_vals(data.values.get("roce", [np.nan] * len(labels))))
    roe  = np.array(_safe_vals(data.values.get("roe",  [np.nan] * len(labels))))
    wacc = np.array(_safe_vals(data.values.get("wacc", [np.nan] * len(labels))))

    # Shade ROCE vs WACC
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        ax.fill_between(
            x, roce, wacc,
            where=(roce >= wacc),
            interpolate=True,
            color=GREEN_RGB, alpha=0.15, label="_nolegend_",
        )
        ax.fill_between(
            x, roce, wacc,
            where=(roce < wacc),
            interpolate=True,
            color=(0.8, 0.1, 0.1), alpha=0.15, label="_nolegend_",
        )

    ax.plot(x, roce, color=NAVY_RGB,  linewidth=2.0, marker="o", markersize=5, label="ROCE %")
    ax.plot(x, roe,  color=GREEN_RGB, linewidth=1.8, marker="s", markersize=4, label="ROE %")
    ax.plot(x, wacc, color=GOLD_RGB,  linewidth=1.5, linestyle="--",
            marker="^", markersize=4, label="WACC %")

    ax.set_xticks(x)
    ax.set_xticklabels(labels, rotation=45, ha="right")
    ax.set_ylabel("Return / Cost (%)")
    ax.yaxis.set_major_formatter(mticker.FormatStrFormatter("%.1f%%"))
    ax.set_title(f"{company} — ROCE & ROE vs WACC")
    ax.legend(loc="upper left", fontsize=7)

    if data.actuals_end_idx >= 0:
        add_projection_separator(ax, data.actuals_end_idx, labels)
    add_source_note(ax, _source(data))
    save_chart(fig, output_path)
    return output_path


# ===========================================================================
# CATEGORY 5: Shareholding  (charts 24, 27)
# ===========================================================================


def chart_shareholding_pie(data: ChartData, output_path: str) -> str:
    """Donut chart: shareholding breakdown for the latest quarter.

    data.values needs: dict of {holder_name: percentage}.
    """
    fig, ax = plt.subplots(figsize=(8, 6))
    fig.patch.set_facecolor("white")
    ax.set_facecolor("white")
    company = _company(data)

    holders = list(data.values.keys())
    pcts    = [
        (data.values[h][-1] if isinstance(data.values[h], list) else data.values[h])
        for h in holders
    ]
    pcts    = [0.0 if (v is None or math.isnan(v)) else v for v in pcts]

    palette = [NAVY_RGB, GREEN_RGB, GOLD_RGB, LGRAY_RGB, DGRAY_RGB,
               (0.6, 0.3, 0.1), (0.3, 0.5, 0.7)]

    wedges, texts, autotexts = ax.pie(
        pcts,
        labels=holders,
        autopct="%1.1f%%",
        pctdistance=0.75,
        colors=palette[: len(holders)],
        startangle=90,
        wedgeprops=dict(width=0.5, edgecolor="white", linewidth=1.5),
    )
    for at in autotexts:
        at.set_fontsize(7)
    for t in texts:
        t.set_fontsize(8)

    period = data.metadata.get("period", "Latest Quarter")
    ax.set_title(f"{company} — Shareholding Pattern ({period})")
    add_source_note(ax, _source(data))
    save_chart(fig, output_path)
    return output_path


def chart_shareholding_trend(data: ChartData, output_path: str) -> str:
    """Stacked area: Promoter %, FII %, DII %, Public % over quarters.

    data.values needs: 'promoter', 'fii', 'dii', 'public'.
    """
    fig, ax = get_fig_ax()
    company = _company(data)
    labels = data.fy_labels
    x = _x_positions(len(labels))

    keys    = ["promoter", "fii", "dii", "public"]
    lbls    = ["Promoter", "FII", "DII", "Public"]
    palette = [NAVY_RGB, GREEN_RGB, GOLD_RGB, LGRAY_RGB]

    arrays = [np.nan_to_num(np.array(_safe_vals(data.values.get(k, [0]*len(labels))))) for k in keys]
    ax.stackplot(x, *arrays, labels=lbls, colors=palette, alpha=0.85)

    ax.set_xticks(x)
    ax.set_xticklabels(labels, rotation=45, ha="right")
    ax.set_ylabel("Shareholding (%)")
    ax.yaxis.set_major_formatter(mticker.FormatStrFormatter("%.0f%%"))
    ax.set_ylim(0, 100)
    ax.set_title(f"{company} — Shareholding Trend")
    ax.legend(loc="upper right", fontsize=7, reverse=True)
    add_source_note(ax, _source(data))
    save_chart(fig, output_path)
    return output_path


# ===========================================================================
# CATEGORY 6: Valuation  (charts 28–32, 34)
# ===========================================================================


def chart_dcf_sensitivity_heatmap(data: ChartData, output_path: str) -> str:
    """2D heatmap: price sensitivity to WACC (rows) vs terminal growth (cols).

    data.values needs: 'wacc_range' (list), 'tg_range' (list), 'price_matrix' (2D list).
    Color: red (low) → white → green (high). Marks base case cell.
    """
    fig, ax = plt.subplots(figsize=(10, 6))
    fig.patch.set_facecolor("white")
    ax.set_facecolor("white")
    company = _company(data)

    wacc_range = data.values.get("wacc_range", [])
    tg_range   = data.values.get("tg_range",   [])
    matrix     = data.values.get("price_matrix", [[]])

    mat = np.array(matrix, dtype=float)

    cmap = LinearSegmentedColormap.from_list(
        "rdwhgn", ["#CC0000", "#FFFFFF", GREEN_HEX], N=256
    )
    im = ax.imshow(mat, cmap=cmap, aspect="auto")
    plt.colorbar(im, ax=ax, label="Implied Price (\u20b9)")

    ax.set_xticks(range(len(tg_range)))
    ax.set_xticklabels([f"{v:.1f}%" for v in tg_range], fontsize=7)
    ax.set_yticks(range(len(wacc_range)))
    ax.set_yticklabels([f"{v:.1f}%" for v in wacc_range], fontsize=7)
    ax.set_xlabel("Terminal Growth Rate")
    ax.set_ylabel("WACC")

    # Cell text
    for i in range(mat.shape[0]):
        for j in range(mat.shape[1]):
            ax.text(j, i, f"\u20b9{mat[i, j]:.0f}",
                    ha="center", va="center", fontsize=6,
                    color="black")

    # Base case marker
    base_row = data.metadata.get("base_wacc_idx", mat.shape[0] // 2)
    base_col = data.metadata.get("base_tg_idx",   mat.shape[1] // 2)
    ax.add_patch(mpatches.Rectangle(
        (base_col - 0.5, base_row - 0.5), 1, 1,
        fill=False, edgecolor=NAVY_HEX, linewidth=2.0,
    ))

    ax.set_title(f"{company} — DCF Sensitivity: Price vs WACC & Terminal Growth")
    add_source_note(ax, _source(data))
    save_chart(fig, output_path)
    return output_path


def chart_dcf_waterfall(data: ChartData, output_path: str) -> str:
    """Waterfall/bridge: PV of FCF + PV Terminal Value - Net Debt = Equity Value → Per Share.

    data.values needs: 'pv_fcf', 'pv_terminal', 'net_cash', 'equity_value', 'per_share'.
    """
    fig, ax = get_fig_ax(width=10, height=5)
    company = _company(data)

    items = [
        ("PV of FCFs",    data.values.get("pv_fcf",       [0])[0] if isinstance(data.values.get("pv_fcf", 0), list) else data.values.get("pv_fcf", 0)),
        ("PV Terminal",   data.values.get("pv_terminal",  [0])[0] if isinstance(data.values.get("pv_terminal", 0), list) else data.values.get("pv_terminal", 0)),
        ("Net Cash/(Debt)",data.values.get("net_cash",    [0])[0] if isinstance(data.values.get("net_cash", 0), list) else data.values.get("net_cash", 0)),
        ("Equity Value",  data.values.get("equity_value", [0])[0] if isinstance(data.values.get("equity_value", 0), list) else data.values.get("equity_value", 0)),
        ("Per Share (\u20b9)", data.values.get("per_share",   [0])[0] if isinstance(data.values.get("per_share", 0), list) else data.values.get("per_share", 0)),
    ]

    labels_w = [i[0] for i in items]
    values_w = [i[1] if i[1] is not None else 0.0 for i in items]
    n = len(labels_w)
    x = _x_positions(n)

    running = 0.0
    for i, (lbl, val) in enumerate(zip(labels_w, values_w)):
        is_total = lbl in ("Equity Value", "Per Share (\u20b9)")
        if is_total:
            bottom = 0
            color  = NAVY_RGB
        else:
            bottom = running
            color  = GREEN_RGB if val >= 0 else (0.8, 0.1, 0.1)

        ax.bar(x[i], abs(val), bottom=min(bottom, bottom + val) if not is_total else 0,
               color=color, width=0.5, zorder=2)
        ax.text(x[i], (bottom + val / 2) if not is_total else val / 2,
                f"\u20b9{val:,.0f}" if "Share" in lbl else f"{val:,.0f}",
                ha="center", va="center", fontsize=7, color="white", fontweight="bold")

        if not is_total:
            running += val

    ax.set_xticks(x)
    ax.set_xticklabels(labels_w, rotation=20, ha="right", fontsize=8)
    ax.set_ylabel("(\u20b9 Cr / \u20b9 per share)")
    ax.set_title(f"{company} — DCF Waterfall Bridge")
    ax.axhline(0, color=DGRAY_HEX, linewidth=0.6)
    add_source_note(ax, _source(data))
    save_chart(fig, output_path)
    return output_path


def chart_peer_multiples_bar(data: ChartData, output_path: str) -> str:
    """Horizontal grouped bar: P/E and EV/EBITDA for company vs peers.

    data.values needs: 'pe' and 'ev_ebitda' dicts of {company: value}.
    Subject company bar in NAVY, peers in LGRAY.
    """
    fig, axes = plt.subplots(1, 2, figsize=(14, 5))
    fig.patch.set_facecolor("white")
    company = _company(data)
    subject_ticker = data.metadata.get("ticker", company)

    for ax, metric_key, metric_label in zip(
        axes,
        ["pe", "ev_ebitda"],
        ["P/E (x)", "EV/EBITDA (x)"],
    ):
        ax.set_facecolor("white")
        metric_dict = data.values.get(metric_key, {})
        if not isinstance(metric_dict, dict):
            continue
        peers  = list(metric_dict.keys())
        values = [metric_dict[p] if metric_dict[p] is not None else 0.0 for p in peers]
        y      = _x_positions(len(peers))
        colors = [NAVY_RGB if p == subject_ticker else LGRAY_RGB for p in peers]

        bars = ax.barh(y, values, color=colors, height=0.55, zorder=2)
        for bar, val in zip(bars, values):
            ax.text(bar.get_width() + 0.1, bar.get_y() + bar.get_height() / 2,
                    f"{val:.1f}x", va="center", fontsize=7)

        ax.set_yticks(y)
        ax.set_yticklabels(peers, fontsize=8)
        ax.set_xlabel(metric_label)
        ax.set_title(metric_label)
        ax.grid(axis="x", alpha=0.4)
        ax.invert_yaxis()

    fig.suptitle(f"{company} — Peer Multiples Comparison", fontsize=11)
    add_source_note(axes[0], _source(data))
    save_chart(fig, output_path)
    return output_path


def chart_comps_scatter(data: ChartData, output_path: str) -> str:
    """Scatter: EBITDA margin (x) vs EV/EBITDA multiple (y) for peer group.

    data.values needs: 'ebitda_margin', 'ev_ebitda', 'labels' (parallel lists).
    Labels each dot with company ticker. Subject company in larger NAVY dot.
    """
    fig, ax = get_fig_ax(width=10, height=6)
    company = _company(data)
    subject_ticker = data.metadata.get("ticker", company)

    margins    = _safe_vals(data.values.get("ebitda_margin", []))
    multiples  = _safe_vals(data.values.get("ev_ebitda",     []))
    tick_labels = data.values.get("labels", [f"Co{i}" for i in range(len(margins))])

    for i, (mx, my, lbl) in enumerate(zip(margins, multiples, tick_labels)):
        if math.isnan(mx) or math.isnan(my):
            continue
        is_subject = (lbl == subject_ticker)
        ax.scatter(
            mx, my,
            s=120 if is_subject else 60,
            color=NAVY_RGB if is_subject else LGRAY_RGB,
            edgecolors=DGRAY_HEX,
            linewidths=0.8,
            zorder=4 if is_subject else 3,
        )
        ax.annotate(
            lbl,
            (mx, my),
            xytext=(4, 4),
            textcoords="offset points",
            fontsize=7,
            color=NAVY_HEX if is_subject else "#555555",
            fontweight="bold" if is_subject else "normal",
        )

    ax.set_xlabel("EBITDA Margin (%)")
    ax.set_ylabel("EV/EBITDA (x)")
    ax.xaxis.set_major_formatter(mticker.FormatStrFormatter("%.1f%%"))
    ax.set_title(f"{company} — Comps Scatter: Margin vs Multiple")

    subject_patch = mpatches.Patch(color=NAVY_RGB, label=subject_ticker)
    peer_patch    = mpatches.Patch(color=LGRAY_RGB, label="Peers")
    ax.legend(handles=[subject_patch, peer_patch], fontsize=7)

    add_source_note(ax, _source(data))
    save_chart(fig, output_path)
    return output_path


def chart_valuation_football_field(data: ChartData, output_path: str) -> str:
    """Horizontal bar chart showing valuation range by method.

    data.values needs: list of dicts with 'method', 'low', 'high', 'point_estimate'.
    Draws range as gray bar, point estimate as NAVY marker.
    Adds CMP as vertical dashed line and price target as GREEN line.
    """
    fig, ax = get_fig_ax(width=12, height=6)
    company = _company(data)

    scenarios = data.values.get("scenarios", [])
    if not scenarios:
        # Try direct list under any key
        for v in data.values.values():
            if isinstance(v, list) and v and isinstance(v[0], dict):
                scenarios = v
                break

    cmp    = data.metadata.get("cmp", None)
    target = data.metadata.get("price_target", None)

    y_positions = _x_positions(len(scenarios))

    for y_pos, s in zip(y_positions, scenarios):
        lo  = s.get("low",            0.0) or 0.0
        hi  = s.get("high",           0.0) or 0.0
        pt  = s.get("point_estimate", None)
        lbl = s.get("method",         "")

        ax.barh(y_pos, hi - lo, left=lo, height=0.4,
                color=LGRAY_RGB, edgecolor=DGRAY_HEX, linewidth=0.8, zorder=2)

        if pt is not None:
            ax.scatter(pt, y_pos, color=NAVY_RGB, s=60, zorder=5, marker="|",
                       linewidths=3)

    if cmp is not None:
        ax.axvline(cmp, color="#CC0000", linewidth=1.5, linestyle="--",
                   label=f"CMP \u20b9{cmp:.0f}", zorder=6)
    if target is not None:
        ax.axvline(target, color=GREEN_HEX, linewidth=1.5, linestyle="-",
                   label=f"Price Target \u20b9{target:.0f}", zorder=6)

    ax.set_yticks(y_positions)
    ax.set_yticklabels([s.get("method", "") for s in scenarios], fontsize=8)
    ax.set_xlabel("Implied Value (\u20b9 per share)")
    ax.set_title(f"{company} — Valuation Football Field")
    ax.invert_yaxis()
    if cmp is not None or target is not None:
        ax.legend(fontsize=7)

    add_source_note(ax, _source(data))
    save_chart(fig, output_path)
    return output_path


def chart_historical_multiples(data: ChartData, output_path: str) -> str:
    """Line chart: 1yr forward P/E over time with mean ± 1 std band.

    data.values needs: 'pe_1yr_fwd' time series.
    Shades ±1σ band in light blue. Marks current level.
    """
    fig, ax = get_fig_ax()
    company = _company(data)
    labels  = data.fy_labels
    x       = _x_positions(len(labels))

    pe = np.array(_safe_vals(data.values.get("pe_1yr_fwd", [np.nan] * len(labels))))

    clean = pe[~np.isnan(pe)]
    mu    = np.nanmean(clean) if len(clean) else np.nan
    sigma = np.nanstd(clean)  if len(clean) else 0.0

    ax.plot(x, pe, color=NAVY_RGB, linewidth=2.0, marker="o", markersize=4,
            label="1yr Fwd P/E", zorder=4)
    ax.fill_between(x, mu - sigma, mu + sigma,
                    color=(0.5, 0.7, 0.9), alpha=0.25,
                    label=f"\u00b11\u03c3 Band ({mu-sigma:.1f}x\u2013{mu+sigma:.1f}x)")
    ax.axhline(mu, color=GOLD_RGB, linewidth=1.2, linestyle="--",
               label=f"Mean {mu:.1f}x")

    # Mark current (last non-nan) level
    valid_idx = np.where(~np.isnan(pe))[0]
    if len(valid_idx):
        ci = valid_idx[-1]
        ax.annotate(
            f"Current: {pe[ci]:.1f}x",
            xy=(x[ci], pe[ci]),
            xytext=(0, 8), textcoords="offset points",
            fontsize=7, ha="center", color=NAVY_HEX,
        )

    ax.set_xticks(x)
    ax.set_xticklabels(labels, rotation=45, ha="right")
    ax.set_ylabel("1-Year Forward P/E (x)")
    ax.set_title(f"{company} — Historical Forward P/E Band")
    ax.legend(fontsize=7)

    if data.actuals_end_idx >= 0:
        add_projection_separator(ax, data.actuals_end_idx, labels)
    add_source_note(ax, _source(data))
    save_chart(fig, output_path)
    return output_path


# ===========================================================================
# CATEGORY 7: Company Overview  (charts 05–09, 13–18, 20–21)
# ===========================================================================


def chart_company_snapshot(data: ChartData, output_path: str) -> str:
    """Text-based figure: key metrics in a 2×3 grid of metric boxes.

    data.values needs: dict of {metric_name: value_string}.
    e.g.: {'Market Cap': '₹3.9 Lakh Cr', 'CMP': '₹307', 'P/E': '17.4x'}.
    """
    fig, ax = plt.subplots(figsize=(12, 5))
    fig.patch.set_facecolor("white")
    ax.set_facecolor("white")
    ax.axis("off")
    company = _company(data)

    metrics = list(data.values.items())
    n_cols  = 3
    n_rows  = math.ceil(len(metrics) / n_cols)

    cell_w = 1.0 / n_cols
    cell_h = 0.8 / n_rows

    for idx, (name, val) in enumerate(metrics):
        row = idx // n_cols
        col = idx  % n_cols
        x0  = col * cell_w + 0.02
        y0  = 0.9 - (row + 1) * cell_h

        rect = mpatches.FancyBboxPatch(
            (x0, y0), cell_w - 0.04, cell_h - 0.02,
            boxstyle="round,pad=0.01",
            linewidth=0.8,
            edgecolor=DGRAY_HEX,
            facecolor=LGRAY_HEX,
            transform=ax.transAxes,
        )
        ax.add_patch(rect)

        ax.text(
            x0 + (cell_w - 0.04) / 2, y0 + cell_h * 0.62,
            str(name),
            ha="center", va="center",
            fontsize=8, color="#555555",
            transform=ax.transAxes,
        )
        val_str = str(val) if val is not None else "N/A"
        ax.text(
            x0 + (cell_w - 0.04) / 2, y0 + cell_h * 0.28,
            val_str,
            ha="center", va="center",
            fontsize=11, color=NAVY_HEX, fontweight="bold",
            transform=ax.transAxes,
        )

    ax.set_title(f"{company} — Company Snapshot", fontsize=12, pad=10)
    add_source_note(ax, _source(data))
    save_chart(fig, output_path)
    return output_path


def chart_milestones_timeline(data: ChartData, output_path: str) -> str:
    """Horizontal timeline with milestone labels above/below.

    data.values needs: 'years' (list of ints), 'labels' (list of strings),
    'positions' ('above'/'below' alternating).
    """
    fig, ax = get_fig_ax(width=14, height=4)
    company = _company(data)

    years      = data.values.get("years",     [])
    mlabels    = data.values.get("labels",    [])
    positions  = data.values.get("positions", ["above", "below"] * len(years))

    if not years:
        ax.text(0.5, 0.5, "No timeline data", ha="center", va="center",
                transform=ax.transAxes)
    else:
        x_min, x_max = min(years) - 1, max(years) + 1
        ax.axhline(0, color=DGRAY_HEX, linewidth=1.5, zorder=1)
        ax.set_xlim(x_min, x_max)
        ax.set_ylim(-2.5, 2.5)
        ax.axis("off")

        for year, lbl, pos in zip(years, mlabels, positions):
            y_sign  = 1 if pos == "above" else -1
            y_text  = y_sign * 1.4
            y_tick  = y_sign * 0.2

            ax.plot(year, 0, "o", color=NAVY_RGB, markersize=8, zorder=3)
            ax.plot([year, year], [0, y_tick * 5], color=DGRAY_HEX,
                    linewidth=0.8, zorder=2)
            ax.text(year, y_text, f"{year}\n{lbl}",
                    ha="center", va="bottom" if pos == "above" else "top",
                    fontsize=7, color=NAVY_HEX,
                    bbox=dict(boxstyle="round,pad=0.2", facecolor="white",
                              edgecolor=DGRAY_HEX, linewidth=0.6))

    ax.set_title(f"{company} — Corporate Milestones Timeline")
    add_source_note(ax, _source(data))
    save_chart(fig, output_path)
    return output_path


def chart_org_structure(data: ChartData, output_path: str) -> str:
    """Simple org chart using matplotlib text boxes and arrows.

    data.values needs: 'nodes' list of {name, title, level, parent_idx}.
    """
    fig, ax = plt.subplots(figsize=(14, 8))
    fig.patch.set_facecolor("white")
    ax.set_facecolor("white")
    ax.axis("off")
    company = _company(data)

    nodes = data.values.get("nodes", [])
    if not nodes:
        ax.text(0.5, 0.5, "No org structure data", ha="center", va="center",
                transform=ax.transAxes)
        ax.set_title(f"{company} — Organisation Structure")
        add_source_note(ax, _source(data))
        save_chart(fig, output_path)
        return output_path

    # Group nodes by level
    levels: dict[int, list] = {}
    for idx, node in enumerate(nodes):
        lvl = node.get("level", 0)
        levels.setdefault(lvl, []).append((idx, node))

    positions = {}
    for lvl, lvl_nodes in levels.items():
        n_nodes = len(lvl_nodes)
        for rank, (idx, node) in enumerate(lvl_nodes):
            x_pos = (rank + 1) / (n_nodes + 1)
            y_pos = 1.0 - lvl * 0.22
            positions[idx] = (x_pos, y_pos)

    # Draw edges
    for idx, node in enumerate(nodes):
        parent_idx = node.get("parent_idx", None)
        if parent_idx is not None and parent_idx in positions:
            x0, y0 = positions[parent_idx]
            x1, y1 = positions[idx]
            ax.annotate(
                "",
                xy=(x1, y1 + 0.04), xycoords="axes fraction",
                xytext=(x0, y0 - 0.04), textcoords="axes fraction",
                arrowprops=dict(arrowstyle="-|>", color=DGRAY_HEX, lw=1.0),
            )

    # Draw boxes
    for idx, node in enumerate(nodes):
        if idx not in positions:
            continue
        xp, yp = positions[idx]
        name  = node.get("name",  "")
        title = node.get("title", "")
        lvl   = node.get("level", 0)
        bg    = NAVY_HEX if lvl == 0 else (GOLD_HEX if lvl == 1 else "white")
        fc    = "white"  if lvl <= 1 else NAVY_HEX

        ax.text(
            xp, yp,
            f"{name}\n{title}",
            ha="center", va="center",
            fontsize=7, color=fc,
            transform=ax.transAxes,
            bbox=dict(boxstyle="round,pad=0.4", facecolor=bg,
                      edgecolor=DGRAY_HEX, linewidth=0.8),
        )

    ax.set_title(f"{company} — Organisation Structure")
    add_source_note(ax, _source(data))
    save_chart(fig, output_path)
    return output_path


def chart_product_portfolio(data: ChartData, output_path: str) -> str:
    """Bubble chart: segments by Revenue (x), EBIT Margin (y), bubble size = Revenue.

    data.values needs: 'segments' list of {name, revenue, ebit_margin}.
    """
    fig, ax = get_fig_ax(width=10, height=6)
    company = _company(data)

    segments = data.values.get("segments", [])
    palette  = [NAVY_RGB, GREEN_RGB, GOLD_RGB, LGRAY_RGB, DGRAY_RGB]

    if not segments:
        ax.text(0.5, 0.5, "No segment data", ha="center", va="center",
                transform=ax.transAxes)
    else:
        revenues = [s.get("revenue", 0) or 0 for s in segments]
        max_rev  = max(revenues) if revenues else 1

        for i, seg in enumerate(segments):
            rev  = seg.get("revenue",     0) or 0
            mrgn = seg.get("ebit_margin", 0) or 0
            name = seg.get("name",        f"Seg {i}")
            size = 200 + 1500 * (rev / max_rev) if max_rev else 200

            ax.scatter(rev, mrgn, s=size,
                       color=palette[i % len(palette)],
                       edgecolors="white", linewidths=1.0,
                       alpha=0.8, zorder=4)
            ax.annotate(name, (rev, mrgn), ha="center", va="center",
                        fontsize=7, color="white", fontweight="bold")

    ax.set_xlabel("Revenue (\u20b9 Cr)")
    ax.set_ylabel("EBIT Margin (%)")
    ax.yaxis.set_major_formatter(mticker.FormatStrFormatter("%.1f%%"))
    ax.set_title(f"{company} — Product / Segment Portfolio")
    add_source_note(ax, _source(data))
    save_chart(fig, output_path)
    return output_path


def chart_customer_channel(data: ChartData, output_path: str) -> str:
    """Donut chart: revenue by channel.

    data.values needs: dict of {channel: pct}.
    """
    fig, ax = plt.subplots(figsize=(8, 6))
    fig.patch.set_facecolor("white")
    ax.set_facecolor("white")
    company = _company(data)

    channels = list(data.values.keys())
    pcts = [
        (data.values[c][-1] if isinstance(data.values[c], list) else data.values[c])
        for c in channels
    ]
    pcts = [0.0 if (v is None or (isinstance(v, float) and math.isnan(v))) else float(v)
            for v in pcts]

    palette = [NAVY_RGB, GREEN_RGB, GOLD_RGB, LGRAY_RGB, DGRAY_RGB,
               (0.6, 0.3, 0.1), (0.3, 0.5, 0.7)]

    wedges, texts, autotexts = ax.pie(
        pcts,
        labels=channels,
        autopct="%1.1f%%",
        pctdistance=0.75,
        colors=palette[: len(channels)],
        startangle=90,
        wedgeprops=dict(width=0.5, edgecolor="white", linewidth=1.5),
    )
    for at in autotexts:
        at.set_fontsize(7)
    for t in texts:
        t.set_fontsize(8)

    ax.set_title(f"{company} — Revenue by Customer Channel")
    add_source_note(ax, _source(data))
    save_chart(fig, output_path)
    return output_path


def chart_scenario_comparison(data: ChartData, output_path: str) -> str:
    """Grouped bar: Bull/Base/Bear revenue and PAT for projection years.

    data.values needs: 'bull_revenue', 'base_revenue', 'bear_revenue',
    'bull_pat', 'base_pat', 'bear_pat'.
    """
    labels = data.fy_labels
    n      = len(labels)
    x      = _x_positions(n)
    width  = 0.14

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 5))
    fig.patch.set_facecolor("white")
    company = _company(data)

    scenario_colors = {
        "bull": GREEN_RGB,
        "base": NAVY_RGB,
        "bear": (0.8, 0.1, 0.1),
    }

    for ax, metric, ylabel in [(ax1, "revenue", "Revenue (\u20b9 Cr)"),
                                 (ax2, "pat",     "PAT (\u20b9 Cr)")]:
        ax.set_facecolor("white")
        offsets = {"bull": -width, "base": 0.0, "bear": width}
        for scenario, offset in offsets.items():
            key  = f"{scenario}_{metric}"
            vals = _safe_vals(data.values.get(key, [np.nan] * n))
            bars = ax.bar(x + offset, vals[:n], width=width * 0.9,
                          color=scenario_colors[scenario],
                          label=scenario.title(), zorder=2)
        ax.set_xticks(x)
        ax.set_xticklabels(labels, rotation=45, ha="right")
        ax.set_ylabel(ylabel)
        ax.set_title(metric.upper())
        format_crore_axis(ax, "y")
        ax.legend(fontsize=7)

    fig.suptitle(f"{company} — Bull / Base / Bear Scenario Comparison", fontsize=11)
    add_source_note(ax1, _source(data))
    save_chart(fig, output_path)
    return output_path


def chart_price_target_scenarios(data: ChartData, output_path: str) -> str:
    """Horizontal bar: Bull/Base/Bear price targets with probability weights.

    data.values needs: 'scenarios' list of {label, price, probability, color}.
    """
    fig, ax = get_fig_ax(width=10, height=5)
    company = _company(data)

    scenarios = data.values.get("scenarios", [])
    default_colors = [GREEN_RGB, NAVY_RGB, (0.8, 0.1, 0.1)]

    y_pos = _x_positions(len(scenarios))
    for i, s in enumerate(scenarios):
        lbl   = s.get("label",       f"Scenario {i+1}")
        price = s.get("price",       0.0) or 0.0
        prob  = s.get("probability", 0.0) or 0.0
        color = s.get("color",       default_colors[i % len(default_colors)])

        bar = ax.barh(y_pos[i], price, height=0.5,
                      color=color, zorder=2, alpha=0.85)
        ax.text(
            price + max(price * 0.01, 5), y_pos[i],
            f"\u20b9{price:.0f}  ({prob:.0f}% prob.)",
            va="center", fontsize=8, color="#333333",
        )

    ax.set_yticks(y_pos)
    ax.set_yticklabels([s.get("label", "") for s in scenarios], fontsize=9)
    ax.set_xlabel("Price Target (\u20b9 per share)")
    ax.set_title(f"{company} — Price Target Scenarios")
    ax.invert_yaxis()

    cmp = data.metadata.get("cmp", None)
    if cmp is not None:
        ax.axvline(cmp, color=DGRAY_HEX, linewidth=1.2, linestyle="--",
                   label=f"CMP \u20b9{cmp:.0f}")
        ax.legend(fontsize=7)

    add_source_note(ax, _source(data))
    save_chart(fig, output_path)
    return output_path


def chart_tam_evolution(data: ChartData, output_path: str) -> str:
    """Bar + line: TAM size evolution with CAGR annotation.

    data.values needs: 'tam_size' (INR Cr).
    """
    fig, ax = get_fig_ax()
    company = _company(data)
    labels  = data.fy_labels
    n       = len(labels)
    x       = _x_positions(n)

    tam = _safe_vals(data.values.get("tam_size", [np.nan] * n))

    colors = _bar_fill_colors(n, data.actuals_end_idx)
    ax.bar(x, tam, color=colors, width=0.55, zorder=2, label="TAM (\u20b9 Cr)")

    clean = [(i, v) for i, v in enumerate(tam) if not math.isnan(v)]
    if len(clean) >= 2:
        first_i, first_v = clean[0]
        last_i,  last_v  = clean[-1]
        n_periods = last_i - first_i
        if n_periods > 0 and first_v and first_v > 0:
            cagr = ((last_v / first_v) ** (1 / n_periods) - 1) * 100
            mid_x = (x[first_i] + x[last_i]) / 2
            mid_y = (first_v + last_v) / 2
            ax.annotate(
                f"CAGR: {cagr:.1f}%",
                xy=(mid_x, mid_y),
                fontsize=9, color=GREEN_HEX, fontweight="bold",
                ha="center",
                bbox=dict(boxstyle="round,pad=0.3", facecolor="white",
                          edgecolor=GREEN_HEX, linewidth=0.8),
            )

    ax.set_xticks(x)
    ax.set_xticklabels(labels, rotation=45, ha="right")
    ax.set_ylabel("TAM (\u20b9 Cr)")
    format_crore_axis(ax, "y")
    ax.set_title(f"{company} — Total Addressable Market Evolution")
    ax.legend(fontsize=7)

    if data.actuals_end_idx >= 0:
        add_projection_separator(ax, data.actuals_end_idx, labels)
    add_source_note(ax, _source(data))
    save_chart(fig, output_path)
    return output_path


def chart_competitive_positioning(data: ChartData, output_path: str) -> str:
    """2×2 matrix scatter: competitive position map.

    data.values needs: 'companies' list of {name, x, y}
    where x = market_share, y = growth.
    Subject company in NAVY, others in LGRAY.
    """
    fig, ax = get_fig_ax(width=10, height=8)
    company = _company(data)
    subject_ticker = data.metadata.get("ticker", company)

    companies = data.values.get("companies", [])

    all_x = [c.get("x", 0) or 0 for c in companies]
    all_y = [c.get("y", 0) or 0 for c in companies]
    mid_x = np.median(all_x) if all_x else 0
    mid_y = np.median(all_y) if all_y else 0

    # Quadrant dividers
    ax.axvline(mid_x, color=DGRAY_HEX, linewidth=0.8, linestyle="--")
    ax.axhline(mid_y, color=DGRAY_HEX, linewidth=0.8, linestyle="--")

    # Quadrant labels
    for (tx, ty, qlbl) in [
        (0.05, 0.95, "Low Share\nHigh Growth"),
        (0.55, 0.95, "High Share\nHigh Growth"),
        (0.05, 0.05, "Low Share\nLow Growth"),
        (0.55, 0.05, "High Share\nLow Growth"),
    ]:
        ax.text(tx, ty, qlbl, transform=ax.transAxes,
                fontsize=7, color="#AAAAAA", va="top" if "High G" in qlbl else "bottom")

    for c in companies:
        cx   = c.get("x",    0) or 0
        cy   = c.get("y",    0) or 0
        name = c.get("name", "")
        is_subject = (name == subject_ticker or name == company)

        ax.scatter(cx, cy,
                   s=120 if is_subject else 60,
                   color=NAVY_RGB if is_subject else LGRAY_RGB,
                   edgecolors=DGRAY_HEX, linewidths=0.8,
                   zorder=4 if is_subject else 3)
        ax.annotate(name, (cx, cy), xytext=(5, 4),
                    textcoords="offset points", fontsize=7,
                    color=NAVY_HEX if is_subject else "#555555",
                    fontweight="bold" if is_subject else "normal")

    ax.set_xlabel("Market Share (%)")
    ax.set_ylabel("Revenue Growth (%)")
    ax.set_title(f"{company} — Competitive Positioning Matrix")
    add_source_note(ax, _source(data))
    save_chart(fig, output_path)
    return output_path


def chart_market_share(data: ChartData, output_path: str) -> str:
    """Pie or stacked bar: market share by player.

    data.values needs: dict of {company: market_share_pct}.
    Uses pie chart when ≤ 6 players, stacked bar otherwise.
    """
    company = _company(data)
    players = list(data.values.keys())
    pcts    = [
        (data.values[p][-1] if isinstance(data.values[p], list) else data.values[p])
        for p in players
    ]
    pcts    = [0.0 if (v is None or (isinstance(v, float) and math.isnan(v))) else float(v)
               for v in pcts]

    palette = [NAVY_RGB, GREEN_RGB, GOLD_RGB, LGRAY_RGB, DGRAY_RGB,
               (0.6, 0.3, 0.1), (0.3, 0.5, 0.7)]
    subject_ticker = data.metadata.get("ticker", company)

    if len(players) <= 6:
        fig, ax = plt.subplots(figsize=(8, 6))
        fig.patch.set_facecolor("white")
        ax.set_facecolor("white")

        explode = [0.05 if p == subject_ticker else 0.0 for p in players]
        wedges, texts, autotexts = ax.pie(
            pcts,
            labels=players,
            autopct="%1.1f%%",
            explode=explode,
            colors=palette[: len(players)],
            startangle=90,
            wedgeprops=dict(edgecolor="white", linewidth=1.5),
        )
        for at in autotexts:
            at.set_fontsize(7)
        for t in texts:
            t.set_fontsize(8)
    else:
        fig, ax = get_fig_ax()
        x = np.array([0])
        bottom = np.zeros(1)
        for i, (p, pct) in enumerate(zip(players, pcts)):
            ax.bar(x, [pct], bottom=bottom,
                   color=palette[i % len(palette)], width=0.6,
                   label=p, zorder=2)
            ax.text(0, bottom[0] + pct / 2, f"{p}\n{pct:.1f}%",
                    ha="center", va="center", fontsize=7, color="white")
            bottom += pct
        ax.set_xlim(-1, 1)
        ax.set_ylim(0, 110)
        ax.set_ylabel("Market Share (%)")
        ax.set_xticks([])
        ax.legend(fontsize=7, loc="upper right")

    plt.gcf().suptitle(f"{company} — Market Share", fontsize=11)
    add_source_note(plt.gca(), _source(data))
    save_chart(fig, output_path)
    return output_path


def chart_peer_benchmarking(data: ChartData, output_path: str) -> str:
    """Radar/spider chart: 5–6 metrics vs peer average.

    data.values needs: 'metrics' (list of names), 'subject' (values),
    'peer_avg' (values).
    """
    fig, ax = plt.subplots(figsize=(8, 8), subplot_kw=dict(polar=True))
    fig.patch.set_facecolor("white")
    company = _company(data)

    metric_names = data.values.get("metrics",  [])
    subject_vals = data.values.get("subject",  [])
    peer_vals    = data.values.get("peer_avg", [])

    n = len(metric_names)
    if n == 0:
        ax.set_title(f"{company} — Peer Benchmarking")
        add_source_note(ax, _source(data))
        save_chart(fig, output_path)
        return output_path

    angles = [2 * math.pi * i / n for i in range(n)]
    angles += angles[:1]  # close polygon

    subj_v = list(subject_vals) + [subject_vals[0]] if subject_vals else [0] * (n + 1)
    peer_v = list(peer_vals)    + [peer_vals[0]]    if peer_vals    else [0] * (n + 1)

    ax.plot(angles, subj_v, color=NAVY_RGB,  linewidth=2.0, label=company)
    ax.fill(angles, subj_v, color=NAVY_RGB,  alpha=0.15)
    ax.plot(angles, peer_v, color=GOLD_RGB,  linewidth=1.5, linestyle="--", label="Peer Avg")
    ax.fill(angles, peer_v, color=GOLD_RGB,  alpha=0.10)

    ax.set_xticks(angles[:-1])
    ax.set_xticklabels(metric_names, fontsize=8)
    ax.set_title(f"{company} — Peer Benchmarking (Radar)", pad=20)
    ax.legend(loc="upper right", bbox_to_anchor=(1.25, 1.1), fontsize=7)
    add_source_note(ax, _source(data))
    save_chart(fig, output_path)
    return output_path


def chart_fmcg_profitability_journey(data: ChartData, output_path: str) -> str:
    """Line: EBIT margin improvement over time for FMCG division.

    data.values needs: 'ebit_margin' time series.
    Annotates key inflection points from metadata['inflections'] list of
    {idx, label}.
    """
    fig, ax = get_fig_ax()
    company = _company(data)
    labels  = data.fy_labels
    x       = _x_positions(len(labels))

    ebit = _safe_vals(data.values.get("ebit_margin", [np.nan] * len(labels)))

    ax.plot(x, ebit, color=NAVY_RGB, linewidth=2.5, marker="o", markersize=5,
            label="FMCG EBIT Margin %", zorder=4)
    ax.fill_between(x, 0, ebit, alpha=0.08, color=NAVY_RGB)

    # Inflection annotations
    inflections = data.metadata.get("inflections", [])
    for inf in inflections:
        idx = inf.get("idx", 0)
        lbl = inf.get("label", "")
        if 0 <= idx < len(ebit) and not math.isnan(ebit[idx]):
            ax.annotate(
                lbl,
                xy=(x[idx], ebit[idx]),
                xytext=(0, 12), textcoords="offset points",
                ha="center", fontsize=6, color=GOLD_HEX,
                arrowprops=dict(arrowstyle="-", color=GOLD_HEX, lw=0.8),
                bbox=dict(boxstyle="round,pad=0.2", facecolor="white",
                          edgecolor=GOLD_HEX, linewidth=0.6),
            )

    ax.set_xticks(x)
    ax.set_xticklabels(labels, rotation=45, ha="right")
    ax.set_ylabel("EBIT Margin (%)")
    ax.yaxis.set_major_formatter(mticker.FormatStrFormatter("%.1f%%"))
    ax.set_title(f"{company} — FMCG Division Profitability Journey")
    ax.legend(fontsize=7)

    if data.actuals_end_idx >= 0:
        add_projection_separator(ax, data.actuals_end_idx, labels)
    add_source_note(ax, _source(data))
    save_chart(fig, output_path)
    return output_path


def chart_operating_metrics_dashboard(data: ChartData, output_path: str) -> str:
    """2×2 subplot dashboard: Revenue growth %, EBITDA margin %, ROCE %, FCF conversion %.

    data.values needs: 'revenue_growth', 'ebitda_margin', 'roce', 'fcf_conversion'.
    """
    fig, axes = plt.subplots(2, 2, figsize=(14, 8))
    fig.patch.set_facecolor("white")
    company = _company(data)
    labels  = data.fy_labels
    n       = len(labels)
    x       = _x_positions(n)

    cfg = [
        ("revenue_growth",  "Revenue Growth (%)",    NAVY_RGB,  "%"),
        ("ebitda_margin",   "EBITDA Margin (%)",     GREEN_RGB, "%"),
        ("roce",            "ROCE (%)",               GOLD_RGB,  "%"),
        ("fcf_conversion",  "FCF Conversion (%)",    (0.4, 0.3, 0.6), "%"),
    ]

    for ax, (key, title, color, suffix) in zip(axes.flat, cfg):
        ax.set_facecolor("white")
        vals = _safe_vals(data.values.get(key, [np.nan] * n))

        colors = [color if (not math.isnan(v) and v >= 0) else (0.8, 0.1, 0.1)
                  for v in vals]
        bars = ax.bar(x, vals, color=colors, width=0.6, zorder=2)
        add_data_labels(ax, bars, fmt=f"{{:.1f}}{suffix}", fontsize=6)

        ax.axhline(0, color=DGRAY_HEX, linewidth=0.6)
        ax.set_xticks(x)
        ax.set_xticklabels(labels, rotation=45, ha="right", fontsize=7)
        ax.set_ylabel(title, fontsize=8)
        ax.set_title(title, fontsize=9)
        ax.yaxis.set_major_formatter(mticker.FormatStrFormatter(f"%.1f{suffix}"))

        if data.actuals_end_idx >= 0:
            add_projection_separator(ax, data.actuals_end_idx, labels)

    fig.suptitle(f"{company} — Operating Metrics Dashboard", fontsize=12)
    add_source_note(axes[1][0], _source(data))
    save_chart(fig, output_path)
    return output_path


def chart_segment_ebit_waterfall(data: ChartData, output_path: str) -> str:
    """Waterfall: segment EBIT contributions to total EBIT for latest year.

    data.values needs: dict of {segment: ebit_value}.
    """
    fig, ax = get_fig_ax(width=12, height=5)
    company  = _company(data)
    segments = list(data.values.keys())
    year_lbl = data.fy_labels[-1] if data.fy_labels else "Latest"

    raw_vals = []
    for seg in segments:
        v = data.values[seg]
        val = v[-1] if isinstance(v, list) else v
        raw_vals.append(float(val) if val is not None else 0.0)

    # Append total bar
    total = sum(raw_vals)
    seg_labels_w = segments + ["Total EBIT"]
    values_w     = raw_vals + [total]
    n            = len(seg_labels_w)
    x            = _x_positions(n)

    running = 0.0
    for i, (lbl, val) in enumerate(zip(seg_labels_w, values_w)):
        is_total = (i == n - 1)
        if is_total:
            bar_bottom = 0
            color      = NAVY_RGB
        else:
            bar_bottom = running
            color      = GREEN_RGB if val >= 0 else (0.8, 0.1, 0.1)

        ax.bar(x[i], val if is_total else abs(val),
               bottom=0 if is_total else min(bar_bottom, bar_bottom + val),
               color=color, width=0.55, zorder=2)

        label_y = (bar_bottom + val / 2) if not is_total else val / 2
        ax.text(x[i], label_y, f"{val:,.0f}",
                ha="center", va="center",
                fontsize=7, color="white", fontweight="bold")

        if not is_total:
            # Connector line
            next_start = running + val
            if i < n - 2:
                ax.plot([x[i] + 0.28, x[i + 1] - 0.28],
                        [next_start, next_start],
                        color=DGRAY_HEX, linewidth=0.7, linestyle=":")
            running += val

    ax.axhline(0, color=DGRAY_HEX, linewidth=0.6)
    ax.set_xticks(x)
    ax.set_xticklabels(seg_labels_w, rotation=30, ha="right", fontsize=8)
    ax.set_ylabel("EBIT (\u20b9 Cr)")
    format_crore_axis(ax, "y")
    ax.set_title(f"{company} — Segment EBIT Waterfall ({year_lbl})")
    add_source_note(ax, _source(data))
    save_chart(fig, output_path)
    return output_path


# ===========================================================================
# CATEGORY 8: Pharma-specific  (charts 36–38)
# ===========================================================================

def chart_pharma_rnd_trend(data: ChartData, output_path: str) -> str:
    """Dual-axis: R&D spend (INR Cr bar) + R&D as % of revenue (line).

    data.values: {"rnd_spend": [...], "rnd_pct": [...], "net_revenue": [...]}
    """
    fig, ax1 = get_fig_ax(width=12, height=5)
    company  = _company(data)
    n        = len(data.fy_labels)
    x        = _x_positions(n)
    ei       = data.actuals_end_idx

    rnd  = _safe_vals(data.values.get("rnd_spend", [0] * n))
    pct  = _safe_vals(data.values.get("rnd_pct",   [0] * n))

    colors = _bar_fill_colors(n, ei)
    bars   = ax1.bar(x, rnd, color=colors, width=0.55, zorder=2, label="R&D Spend (₹ Cr)")
    add_data_labels(ax1, bars, fmt="{:.0f}", fontsize=7)

    ax2 = ax1.twinx()
    ax2.plot(x, pct, color=GOLD_RGB, linewidth=1.8, marker="o",
             markersize=5, label="R&D as % Revenue", zorder=3)
    ax2.set_ylabel("R&D % Revenue", color=GOLD_HEX, fontsize=9)
    ax2.tick_params(axis="y", labelcolor=GOLD_HEX, labelsize=8)
    ax2.set_ylim(0, max(v for v in pct if not np.isnan(v)) * 1.4 if any(not np.isnan(v) for v in pct) else 20)

    add_projection_separator(ax1, ei, n)
    ax1.set_xticks(x)
    ax1.set_xticklabels(data.fy_labels, rotation=30, ha="right", fontsize=8)
    ax1.set_ylabel("R&D Spend (₹ Cr)")
    format_crore_axis(ax1, "y")
    ax1.set_title(f"{company} — R&D Investment Trend")

    lines1, labels1 = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax1.legend(lines1 + lines2, labels1 + labels2, fontsize=8, loc="upper left")
    add_source_note(ax1, _source(data))
    save_chart(fig, output_path)
    return output_path


def chart_pharma_margin_bridge(data: ChartData, output_path: str) -> str:
    """Bar chart: reported EBITDA margin vs adjusted EBITDA margin (ex-R&D).

    data.values: {"ebitda_margin": [...], "ebitda_ex_rnd_margin": [...]}
    """
    fig, ax = get_fig_ax(width=12, height=5)
    company = _company(data)
    n       = len(data.fy_labels)
    x       = _x_positions(n)
    ei      = data.actuals_end_idx

    rep_m  = _safe_vals(data.values.get("ebitda_margin",        [0] * n))
    adj_m  = _safe_vals(data.values.get("ebitda_ex_rnd_margin", [0] * n))

    width  = 0.38
    x_rep  = x - width / 2
    x_adj  = x + width / 2

    b1 = ax.bar(x_rep, rep_m, width=width, color=NAVY_RGB, label="Reported EBITDA Margin", zorder=2)
    b2 = ax.bar(x_adj, adj_m, width=width, color=GREEN_RGB, label="Adj. EBITDA Margin (ex-R&D)", zorder=2)
    add_data_labels(ax, b1, fmt="{:.1f}%", fontsize=7)
    add_data_labels(ax, b2, fmt="{:.1f}%", fontsize=7)

    add_projection_separator(ax, ei, n)
    ax.set_xticks(x)
    ax.set_xticklabels(data.fy_labels, rotation=30, ha="right", fontsize=8)
    ax.set_ylabel("EBITDA Margin (%)")
    ax.set_title(f"{company} — Reported vs Adjusted EBITDA Margin")
    ax.legend(fontsize=8)
    add_source_note(ax, _source(data))
    save_chart(fig, output_path)
    return output_path


def chart_pharma_revenue_mix(data: ChartData, output_path: str) -> str:
    """Stacked 100% bar: revenue mix by geography/segment (Domestic, US, RoW, API).

    data.values: {segment_name: [pct_values_per_fy]}  — values in % (sum ~100)
    """
    fig, ax = get_fig_ax(width=12, height=5)
    company = _company(data)
    n       = len(data.fy_labels)
    x       = _x_positions(n)
    ei      = data.actuals_end_idx

    segments = [k for k in data.values if k != "total"]
    colors   = [NAVY_RGB, GREEN_RGB, GOLD_RGB, LGRAY_RGB, DGRAY_RGB]

    bottoms = np.zeros(n)
    for i, seg in enumerate(segments):
        vals = _safe_vals(data.values.get(seg, [0] * n))
        vals_arr = np.array([0 if np.isnan(v) else v for v in vals])
        ax.bar(x, vals_arr, bottom=bottoms,
               color=colors[i % len(colors)], width=0.55,
               label=seg, zorder=2)
        for j, (bot, val) in enumerate(zip(bottoms, vals_arr)):
            if val > 5:
                ax.text(x[j], bot + val / 2, f"{val:.0f}%",
                        ha="center", va="center", fontsize=7,
                        color="white", fontweight="bold")
        bottoms += vals_arr

    add_projection_separator(ax, ei, n)
    ax.set_xticks(x)
    ax.set_xticklabels(data.fy_labels, rotation=30, ha="right", fontsize=8)
    ax.set_ylabel("Revenue Mix (%)")
    ax.set_ylim(0, 110)
    ax.set_title(f"{company} — Revenue Mix by Geography / Segment")
    ax.legend(fontsize=8, loc="upper right")
    add_source_note(ax, _source(data))
    save_chart(fig, output_path)
    return output_path


# ===========================================================================
# CATEGORY 9: Metals-specific  (charts 39–41)
# ===========================================================================

def chart_metals_ebitda_per_tonne(data: ChartData, output_path: str) -> str:
    """Dual-axis: EBITDA/tonne (bar, ₹) + steel/aluminium volume (line, MT).

    data.values: {"ebitda_per_tonne": [...], "volume_mt": [...]}
    """
    fig, ax1 = get_fig_ax(width=12, height=5)
    company  = _company(data)
    n        = len(data.fy_labels)
    x        = _x_positions(n)
    ei       = data.actuals_end_idx

    ebt  = _safe_vals(data.values.get("ebitda_per_tonne", [0] * n))
    vol  = _safe_vals(data.values.get("volume_mt",        [None] * n))

    colors = _bar_fill_colors(n, ei)
    bars   = ax1.bar(x, ebt, color=colors, width=0.55, zorder=2, label="EBITDA/tonne (₹)")
    add_data_labels(ax1, bars, fmt="{:,.0f}", fontsize=7)

    ax1.set_ylabel("EBITDA per Tonne (₹)")
    ax1.set_title(f"{company} — EBITDA per Tonne")

    if any(not np.isnan(v) for v in vol):
        ax2 = ax1.twinx()
        ax2.plot(x, vol, color=GOLD_RGB, linewidth=1.8, marker="s",
                 markersize=5, label="Volume (MT)", zorder=3)
        ax2.set_ylabel("Volume (Million Tonnes)", color=GOLD_HEX, fontsize=9)
        ax2.tick_params(axis="y", labelcolor=GOLD_HEX, labelsize=8)
        lines1, labels1 = ax1.get_legend_handles_labels()
        lines2, labels2 = ax2.get_legend_handles_labels()
        ax1.legend(lines1 + lines2, labels1 + labels2, fontsize=8, loc="upper left")
    else:
        ax1.legend(fontsize=8)

    add_projection_separator(ax1, ei, n)
    ax1.set_xticks(x)
    ax1.set_xticklabels(data.fy_labels, rotation=30, ha="right", fontsize=8)
    add_source_note(ax1, _source(data))
    save_chart(fig, output_path)
    return output_path


def chart_metals_leverage(data: ChartData, output_path: str) -> str:
    """Bar: Net Debt/EBITDA ratio with danger zone shading above 3x.

    data.values: {"net_debt_ebitda": [...], "net_debt": [...]}
    """
    fig, ax = get_fig_ax(width=12, height=5)
    company = _company(data)
    n       = len(data.fy_labels)
    x       = _x_positions(n)
    ei      = data.actuals_end_idx

    ratio = _safe_vals(data.values.get("net_debt_ebitda", [0] * n))

    def _color(v):
        if np.isnan(v) or v is None:
            return LGRAY_RGB
        if v < 2.0:
            return GREEN_RGB
        if v < 3.5:
            return GOLD_RGB
        return (0.8, 0.1, 0.1)

    bar_colors = [_color(v) for v in ratio]
    bars = ax.bar(x, [0 if np.isnan(v) else v for v in ratio],
                  color=bar_colors, width=0.55, zorder=2)
    add_data_labels(ax, bars, fmt="{:.1f}x", fontsize=8)

    # Danger zone line at 3.5x
    ax.axhline(3.5, color=(0.8, 0.1, 0.1), linewidth=1.2,
               linestyle="--", label="High Risk (3.5x)", zorder=3)
    ax.axhline(2.0, color=GOLD_HEX, linewidth=1.0,
               linestyle=":", label="Caution (2.0x)", zorder=3)

    add_projection_separator(ax, ei, n)
    ax.set_xticks(x)
    ax.set_xticklabels(data.fy_labels, rotation=30, ha="right", fontsize=8)
    ax.set_ylabel("Net Debt / EBITDA (x)")
    ax.set_title(f"{company} — Leverage: Net Debt / EBITDA")
    ax.legend(fontsize=8)
    add_source_note(ax, _source(data))
    save_chart(fig, output_path)
    return output_path


def chart_metals_realization_trend(data: ChartData, output_path: str) -> str:
    """Line: realization per tonne (₹/t) trend — proxy for commodity price exposure.

    data.values: {"realization_per_tonne": [...], "raw_material_cost_pct": [...]}
    """
    fig, ax1 = get_fig_ax(width=12, height=5)
    company  = _company(data)
    n        = len(data.fy_labels)
    x        = _x_positions(n)
    ei       = data.actuals_end_idx

    real  = _safe_vals(data.values.get("realization_per_tonne", [0] * n))
    rm_pct = _safe_vals(data.values.get("raw_material_cost_pct", [None] * n))

    ax1.plot(x, real, color=NAVY_RGB, linewidth=2, marker="o",
             markersize=5, label="Realization/Tonne (₹)", zorder=3)
    ax1.fill_between(x, real, alpha=0.08, color=NAVY_RGB)

    for i, v in enumerate(real):
        if not np.isnan(v):
            ax1.annotate(f"₹{v:,.0f}", (x[i], v),
                         textcoords="offset points", xytext=(0, 7),
                         ha="center", fontsize=7)

    ax1.set_ylabel("Realization per Tonne (₹)")
    ax1.set_title(f"{company} — Realization & Raw Material Cost Trend")

    if any(not np.isnan(v) for v in rm_pct):
        ax2 = ax1.twinx()
        ax2.bar(x, [0 if np.isnan(v) else v for v in rm_pct],
                color=GOLD_RGB, alpha=0.35, width=0.55,
                label="Raw Material Cost %", zorder=2)
        ax2.set_ylabel("Raw Material Cost (% of Revenue)", color=GOLD_HEX, fontsize=9)
        ax2.tick_params(axis="y", labelcolor=GOLD_HEX, labelsize=8)
        lines1, labels1 = ax1.get_legend_handles_labels()
        lines2, labels2 = ax2.get_legend_handles_labels()
        ax1.legend(lines1 + lines2, labels1 + labels2, fontsize=8, loc="upper left")
    else:
        ax1.legend(fontsize=8)

    add_projection_separator(ax1, ei, n)
    ax1.set_xticks(x)
    ax1.set_xticklabels(data.fy_labels, rotation=30, ha="right", fontsize=8)
    add_source_note(ax1, _source(data))
    save_chart(fig, output_path)
    return output_path
