"""
Shared styling for all charts in IndiaStockResearch.
Import this in chart_factory.py — never set rcParams elsewhere.

Color scheme:
  NAVY  = #0D2B55  (primary, used for bars/lines)
  GREEN = #1A753F  (positive/actual)
  GOLD  = #B7950B  (accent/highlight)
  LGRAY = #F2F2F2  (background, alternating rows)
  DGRAY = #D9D9D9  (borders)
"""

from __future__ import annotations

import matplotlib
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
from dataclasses import dataclass, field
from typing import Optional

# ---------------------------------------------------------------------------
# Color constants — hex strings
# ---------------------------------------------------------------------------
NAVY_HEX  = "#0D2B55"
GREEN_HEX = "#1A753F"
GOLD_HEX  = "#B7950B"
LGRAY_HEX = "#F2F2F2"
DGRAY_HEX = "#D9D9D9"

# ---------------------------------------------------------------------------
# Color constants — matplotlib-compatible RGB tuples (0.0–1.0)
# ---------------------------------------------------------------------------

def color_hex_to_rgb(hex_str: str) -> tuple:
    """Convert a CSS hex color string to an (R, G, B) float tuple.

    Parameters
    ----------
    hex_str:
        A 6-digit hex color such as ``"#0D2B55"`` (leading ``#`` optional).

    Returns
    -------
    tuple
        ``(r, g, b)`` where each component is in *[0.0, 1.0]*.
    """
    hex_str = hex_str.lstrip("#")
    r = int(hex_str[0:2], 16) / 255.0
    g = int(hex_str[2:4], 16) / 255.0
    b = int(hex_str[4:6], 16) / 255.0
    return (r, g, b)


NAVY_RGB  = color_hex_to_rgb(NAVY_HEX)
GREEN_RGB = color_hex_to_rgb(GREEN_HEX)
GOLD_RGB  = color_hex_to_rgb(GOLD_HEX)
LGRAY_RGB = color_hex_to_rgb(LGRAY_HEX)
DGRAY_RGB = color_hex_to_rgb(DGRAY_HEX)

# ---------------------------------------------------------------------------
# ChartData dataclass
# ---------------------------------------------------------------------------

@dataclass
class ChartData:
    """Container for all data needed by a single chart.

    Attributes
    ----------
    fy_labels:
        Ordered list of fiscal-year label strings, e.g.
        ``["FY21A", "FY22A", "FY23A", "FY24E", "FY25E", "FY26E"]``.
    actuals_end_idx:
        Zero-based index of the *last actual* data point.  A vertical
        separator line is drawn between this index and the next.
        Pass ``-1`` to suppress the separator.
    values:
        Mapping of ``series_name → list[float | None]``.  ``None`` entries
        signal missing / not-yet-reported data points.
    metadata:
        Arbitrary key/value pairs used by individual charts, e.g.
        ``company_name``, ``ticker``, ``units``, ``source_text``.
    """

    fy_labels: list[str]
    actuals_end_idx: int
    values: dict[str, list[float | None]]
    metadata: dict = field(default_factory=dict)


# ---------------------------------------------------------------------------
# Global style application
# ---------------------------------------------------------------------------

def apply_style() -> None:
    """Set global matplotlib rcParams for the IndiaStockResearch visual style.

    Call once — this module calls it automatically at import time.
    """
    plt.rcParams.update(
        {
            # --- font ---
            "font.family":          "serif",
            "font.serif":           ["Times New Roman", "DejaVu Serif"],
            "font.size":            9,
            "axes.titlesize":       11,
            "axes.labelsize":       9,
            "xtick.labelsize":      8,
            "ytick.labelsize":      8,
            "legend.fontsize":      8,
            # --- figure background ---
            "figure.facecolor":     "white",
            "figure.dpi":           100,
            # --- axes background ---
            "axes.facecolor":       "white",
            # --- grid ---
            "axes.grid":            True,
            "axes.grid.axis":       "y",
            "grid.color":           LGRAY_HEX,
            "grid.alpha":           0.4,
            "grid.linewidth":       0.6,
            # --- spines ---
            "axes.spines.top":      False,
            "axes.spines.right":    False,
            "axes.spines.left":     True,
            "axes.spines.bottom":   True,
            "axes.edgecolor":       DGRAY_HEX,
            "axes.linewidth":       0.8,
            # --- ticks ---
            "xtick.direction":      "out",
            "ytick.direction":      "out",
            "xtick.major.size":     3,
            "ytick.major.size":     3,
            "xtick.major.width":    0.6,
            "ytick.major.width":    0.6,
            # --- legend ---
            "legend.frameon":       True,
            "legend.framealpha":    0.7,
            "legend.shadow":        False,
            "legend.fontsize":      7,
            "legend.title_fontsize": 7,
            # --- savefig ---
            "savefig.facecolor":    "white",
            "savefig.bbox":         "tight",
            "savefig.dpi":          300,
        }
    )


# ---------------------------------------------------------------------------
# Figure / axes helpers
# ---------------------------------------------------------------------------

def get_fig_ax(width: float = 12, height: float = 5):
    """Create a new figure and primary axes with white backgrounds.

    Parameters
    ----------
    width, height:
        Figure dimensions in inches.

    Returns
    -------
    (fig, ax)
    """
    fig, ax = plt.subplots(figsize=(width, height))
    fig.patch.set_facecolor("white")
    ax.set_facecolor("white")
    return fig, ax


# ---------------------------------------------------------------------------
# Annotation helpers
# ---------------------------------------------------------------------------

def add_source_note(ax, text: str) -> None:
    """Add a small gray source note at the bottom-left of the figure.

    Parameters
    ----------
    ax:
        The primary matplotlib axes whose figure will receive the note.
    text:
        Source text, e.g. ``"Source: Company filings, screener.in"``.
    """
    ax.figure.text(
        0.01, 0.01,
        text,
        ha="left", va="bottom",
        fontsize=6,
        color="#888888",
        style="italic",
    )


def add_projection_separator(ax, idx: int, fy_labels: list) -> None:
    """Draw a dashed vertical separator between actuals and estimates.

    The line is placed at *x = idx + 0.5* (i.e., between bar index *idx*
    and *idx + 1*).  A small label "Estimates →" is placed just to the
    right of the line.

    Parameters
    ----------
    ax:
        Target axes.
    idx:
        Zero-based index of the last actual data point.
    fy_labels:
        Full list of FY label strings (used only for bounds checking).
    """
    if idx < 0 or idx >= len(fy_labels) - 1:
        return

    x_pos = idx + 0.5
    ylim  = ax.get_ylim()

    ax.axvline(
        x=x_pos,
        color=DGRAY_HEX,
        linewidth=1.0,
        linestyle="--",
        zorder=3,
    )
    ax.text(
        x_pos + 0.05,
        ylim[1] * 0.97,
        "Estimates \u2192",
        fontsize=6,
        color="#888888",
        va="top",
    )


# ---------------------------------------------------------------------------
# Axis formatters
# ---------------------------------------------------------------------------

def format_crore_axis(ax, axis: str = "y") -> None:
    """Apply a compact ₹ Crore formatter to an axis.

    * If the maximum value on the axis exceeds 50 000, labels are shown as
      ``₹XXK Cr`` (thousands of crore).
    * Otherwise labels are shown as ``₹XX Cr``.

    Parameters
    ----------
    ax:
        Target axes.
    axis:
        ``'y'`` (default) or ``'x'``.
    """
    target_axis = ax.yaxis if axis == "y" else ax.xaxis
    lim = ax.get_ylim() if axis == "y" else ax.get_xlim()
    max_val = abs(lim[1])

    if max_val > 50_000:
        fmt = mticker.FuncFormatter(lambda v, _: f"\u20b9{v/1000:.0f}K Cr")
    else:
        fmt = mticker.FuncFormatter(lambda v, _: f"\u20b9{v:.0f} Cr")

    target_axis.set_major_formatter(fmt)


# ---------------------------------------------------------------------------
# Color helpers
# ---------------------------------------------------------------------------

def bar_colors(n: int, highlight_last: bool = False) -> list:
    """Return a list of *n* bar colors.

    The first ``actuals_end_idx + 1`` bars should be NAVY; the remaining
    ones GOLD (estimates).  Because this function does not know the split
    point, callers should pass the exact desired list length and use
    ``highlight_last`` to colour only the final bar in GOLD.

    In practice, callers that know the split build the color list themselves
    using NAVY_RGB / GOLD_RGB directly.  This helper is useful when all bars
    are the same color or only the last needs accenting.

    Parameters
    ----------
    n:
        Number of colors to return.
    highlight_last:
        If ``True``, the last color in the list will be GOLD; all others
        will be NAVY.

    Returns
    -------
    list of RGB tuples
    """
    colors = [NAVY_RGB] * n
    if highlight_last and n > 0:
        colors[-1] = GOLD_RGB
    return colors


def _make_bar_colors(n_actuals: int, n_estimates: int) -> list:
    """Return NAVY for actuals, GOLD for estimates."""
    return [NAVY_RGB] * n_actuals + [GOLD_RGB] * n_estimates


# ---------------------------------------------------------------------------
# Data-label helper
# ---------------------------------------------------------------------------

def add_data_labels(ax, bars, fmt: str = "{:.0f}", fontsize: int = 7) -> None:
    """Annotate each bar in *bars* with its height value.

    Parameters
    ----------
    ax:
        Target axes.
    bars:
        A ``BarContainer`` (return value of ``ax.bar(...)``).
    fmt:
        Python format string applied to the bar height, e.g. ``"{:.1f}"``.
    fontsize:
        Font size of the annotation text.
    """
    for bar in bars:
        height = bar.get_height()
        if height is None or height != height:  # catches NaN
            continue
        x = bar.get_x() + bar.get_width() / 2.0
        y = height
        label = fmt.format(height)
        va = "bottom" if height >= 0 else "top"
        offset = 1 if height >= 0 else -1
        ax.annotate(
            label,
            xy=(x, y),
            xytext=(0, offset * 2),
            textcoords="offset points",
            ha="center",
            va=va,
            fontsize=fontsize,
            color="#333333",
        )


# ---------------------------------------------------------------------------
# Save helper
# ---------------------------------------------------------------------------

def save_chart(fig, output_path: str, dpi: int = 300) -> None:
    """Apply tight layout, save figure to *output_path*, then close it.

    Parameters
    ----------
    fig:
        Matplotlib figure to save.
    output_path:
        Destination file path (PNG recommended).
    dpi:
        Output resolution.  Defaults to 300.
    """
    try:
        fig.tight_layout(rect=[0, 0.04, 1, 1])
    except Exception:
        pass
    fig.savefig(output_path, dpi=dpi, bbox_inches="tight", facecolor="white")
    plt.close(fig)


# ---------------------------------------------------------------------------
# Auto-apply style on import
# ---------------------------------------------------------------------------
apply_style()
