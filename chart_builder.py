"""
chart_builder.py
Generates all matplotlib chart / table images for the CBR deck.

Design principles:
  - Every tile (metric box) is its own image.
  - Every chart/table is its own image.
  - Aspect ratios are fixed per function (use the FW/FH constants in
    generate_cbr.py to embed at the correct proportional height).
  - Text sizes in charts/tables are ~2 pt larger than initial version.
  - Tile heights are compact (~0.75" figure height).
"""
import textwrap
from pathlib import Path

import pandas as pd
import matplotlib
if matplotlib.get_backend().lower() != "agg":
    matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import matplotlib.ticker as mticker
import numpy as np

# ── Aviatrix brand palette ────────────────────────────────────────────────── #
ORANGE = "#F4621E"
NAVY   = "#1E2D3D"
WHITE  = "#FFFFFF"
LGRAY  = "#F5F5F5"
MGRAY  = "#E8E8E8"
DGRAY  = "#555555"
RED    = "#C0392B"

# ── Figure-size constants (width, height in inches) ──────────────────────── #
# Import these in generate_cbr.py to compute embed heights:
#   embed_h = embed_w * FIG_H / FIG_W
TILE_FW,    TILE_FH    = 3.5,  0.75   # single metric tile
UC_FW,      UC_FH      = 8.5,  4.6    # use cases table image
STACKED_FW, STACKED_FH = 10,   3.5    # consumption stacked bar
SPEND_FW,   SPEND_FH   = 10,   3.0    # monthly spend chart
CTRL_FW,    CTRL_FH    = 9,    3.5    # controller table
REQ_FW,     REQ_FH     = 9,    3.5    # ticket requestor table
REL_FW,     REL_FH     = 12,   4.5    # release lifecycle table
FR_FW,      FR_FH      = 12,   4.5    # feature requests table
MANUAL_FW,  MANUAL_FH  = 11,   3.5    # manual-slide placeholder (full)
BANNER_FW,  BANNER_FH  = 11,   1.5    # manual-slide placeholder (banner/short)

# ── Increased text sizes (+2 pt vs original) ──────────────────────────────── #
_TICK  = 9.5    # axis tick labels
_LABEL = 10.0   # axis / section labels
_TITLE = 11.0   # chart / table titles
_TH    = 9.5    # table header text
_TC    = 9.5    # table cell text
_BAR   = 8.5    # value labels inside stacked bars
_TILE_LABEL = 10.5   # small label under tile value

PARAM_COLORS = {
    "billing_firenet_firewall_inst": "#1565C0",
    "s2t_intra_att":                 "#5C9BD6",
    "spoke_s2c_att":                 "#F4A460",
    "t2t_inter_att":                 "#E8963E",
    "t2t_intra_att":                 "#2E7D32",
    "transit_s2c_ext_att":           "#43A047",
}
DEFAULT_COLORS = [
    "#1565C0", "#5C9BD6", "#F4A460", "#E8963E",
    "#2E7D32", "#43A047", "#9C27B0", "#00ACC1",
    "#FF7043", "#8D6E63",
]


# ────────────────────────────────────────────────────────────────────────────
# TILE (single metric box)                    figsize = (TILE_FW, TILE_FH)
# ────────────────────────────────────────────────────────────────────────────
def metric_tile(value: str, label: str, out_path: Path) -> Path:
    """
    Single compact metric tile: large value + small label.
    Figure size: (3.5 × 0.75) inches.  Embed preserving aspect ratio.
    """
    fig = plt.figure(figsize=(TILE_FW, TILE_FH), facecolor=WHITE)
    ax  = fig.add_axes([0, 0, 1, 1])
    ax.set_xlim(0, TILE_FW)
    ax.set_ylim(0, TILE_FH)
    ax.axis("off")

    rect = mpatches.FancyBboxPatch(
        (0.05, 0.04), TILE_FW - 0.10, TILE_FH - 0.08,
        boxstyle="round,pad=0.03", linewidth=0.8,
        edgecolor="#CCCCCC", facecolor=LGRAY,
    )
    ax.add_patch(rect)

    # Auto-shrink font for long values
    vlen = len(str(value))
    vfont = 22 if vlen <= 4 else (16 if vlen <= 10 else 12)
    ax.text(TILE_FW / 2, TILE_FH * 0.65, str(value),
            ha="center", va="center", fontsize=vfont,
            fontweight="bold", color=NAVY)
    wrapped = textwrap.fill(label, 32)
    ax.text(TILE_FW / 2, TILE_FH * 0.18, wrapped,
            ha="center", va="center", fontsize=_TILE_LABEL, color=DGRAY)

    try:
        fig.savefig(out_path, dpi=160, bbox_inches="tight", facecolor=WHITE)
    finally:
        plt.close(fig)
    return out_path


# ────────────────────────────────────────────────────────────────────────────
# CONSUMPTION – stacked parameter bar chart    figsize = (STACKED_FW, STACKED_FH)
# ────────────────────────────────────────────────────────────────────────────
def stacked_parameter_chart(pivot_mrr, pivot_usage, out_path: Path) -> Path:
    """
    Stacked bar chart: Consumption Breakdown by Parameter (MRR).
    Figure size: (10 × 3.5) inches.
    """
    params  = list(pivot_mrr.columns)
    months  = pivot_mrr.index
    mlabels = []
    for m in months:
        try:
            mlabels.append(pd.Timestamp(m).strftime("%b '%y"))
        except (ValueError, AttributeError):
            mlabels.append("?")
    colors  = [PARAM_COLORS.get(p, DEFAULT_COLORS[i % len(DEFAULT_COLORS)])
               for i, p in enumerate(params)]

    fig, ax = plt.subplots(figsize=(STACKED_FW, STACKED_FH), facecolor=WHITE)
    ax.set_facecolor("#F7F8FA")
    x, bar_w, bottoms = np.arange(len(months)), 0.65, np.zeros(len(months))

    for i, param in enumerate(params):
        vals = pivot_mrr[param].values
        bars = ax.bar(x, vals, bar_w, bottom=bottoms, color=colors[i], label=param)
        for j, (bar, val) in enumerate(zip(bars, vals)):
            if val < 300:
                continue
            usage_val = pivot_usage[param].iloc[j] if param in pivot_usage.columns else 0
            if pd.isna(usage_val):
                usage_val = 0
            ax.text(bar.get_x() + bar.get_width() / 2,
                    bottoms[j] + val / 2,
                    str(int(round(usage_val))),
                    ha="center", va="center", fontsize=_BAR,
                    color="white", fontweight="bold")
        bottoms += vals

    ax.set_xticks(x)
    ax.set_xticklabels(mlabels, rotation=0, fontsize=_TICK)
    ax.set_ylabel("MRR ($)", fontsize=_LABEL, color=DGRAY)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"${v:,.0f}"))
    ax.tick_params(axis="y", labelsize=_TICK)
    ax.grid(axis="y", linestyle="--", alpha=0.35)
    ax.spines[["top", "right"]].set_visible(False)
    ax.set_title("Consumption Breakdown by Parameter", loc="left",
                 fontsize=_TITLE, fontweight="bold", color=NAVY, pad=5)

    patches = [mpatches.Patch(color=colors[i], label=params[i])
               for i in range(len(params))]
    ax.legend(handles=patches, loc="upper left", bbox_to_anchor=(1.01, 1),
              fontsize=_TICK, frameon=False, title="Parameters",
              title_fontsize=_TICK)
    # Reserve a fixed 25% right margin for the legend so the chart area is
    # always the same width regardless of how long the legend labels are.
    fig.subplots_adjust(left=0.09, right=0.75, bottom=0.13, top=0.93)
    try:
        fig.savefig(out_path, dpi=160, facecolor=WHITE)
    finally:
        plt.close(fig)
    return out_path


# ────────────────────────────────────────────────────────────────────────────
# CONSUMPTION – monthly spend vs usage bar chart   figsize = (SPEND_FW, SPEND_FH)
# ────────────────────────────────────────────────────────────────────────────
def monthly_spend_chart(monthly_spend_df, out_path: Path) -> Path:
    """
    Grouped bar chart: Monthly Spend vs Usage.
    Figure size: (10 × 3.0) inches.
    """
    import pandas as pd
    fig, ax = plt.subplots(figsize=(SPEND_FW, SPEND_FH), facecolor=WHITE)
    ax.set_facecolor("#F7F8FA")

    _spend_ok = False
    if monthly_spend_df is not None and not monthly_spend_df.empty:
        _required = {"Date Month", "Monthly Spend", "Monthly Usage"}
        if not _required.issubset(monthly_spend_df.columns):
            ax.text(0.5, 0.5, "Monthly spend data format not recognised",
                    ha="center", va="center", transform=ax.transAxes,
                    fontsize=_LABEL, color=DGRAY, style="italic")
            ax.axis("off")
        else:
            df = monthly_spend_df.copy()
            if "Account Name" in df.columns:
                df["Account Name"] = df["Account Name"].astype(str)
                df = df[df["Account Name"].str.strip().str.lower() != "nan"]
                df = df[df["Account Name"].str.strip() != ""].copy()
            df["Date Month"] = pd.to_datetime(df["Date Month"], errors="coerce")
            df = df.dropna(subset=["Date Month"])
            df = df.sort_values("Date Month")
            try:
                spend_vals = df["Monthly Spend"].values.astype(float)
                usage_vals = df["Monthly Usage"].values.astype(float)
            except (ValueError, TypeError):
                ax.text(0.5, 0.5, "Non-numeric values in monthly spend data",
                        ha="center", va="center", transform=ax.transAxes,
                        fontsize=_LABEL, color=DGRAY, style="italic")
                ax.axis("off")
            else:
                _spend_ok = True
                sp_labels = []
                for m in df["Date Month"].values:
                    try:
                        sp_labels.append(pd.Timestamp(m).strftime("%b '%y"))
                    except (ValueError, AttributeError):
                        sp_labels.append("?")

                bw = 0.38
                xs = np.arange(len(df))
                ax.bar(xs - bw / 2, spend_vals, bw, color="#1565C0", label="Monthly Spend ($)")
                ax.bar(xs + bw / 2, usage_vals, bw, color=ORANGE,    label="Monthly Usage ($)")
                ax.set_xticks(xs)
                ax.set_xticklabels(sp_labels, rotation=0, fontsize=_TICK)
                ax.set_ylabel("Amount ($)", fontsize=_LABEL, color=DGRAY)
                ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"${v:,.0f}"))
                ax.tick_params(axis="y", labelsize=_TICK)
                ax.grid(axis="y", linestyle="--", alpha=0.35)
                ax.spines[["top", "right"]].set_visible(False)
                ax.set_title("Monthly Spend vs Usage", loc="left",
                             fontsize=_TITLE, fontweight="bold", color=NAVY, pad=5)
                ax.legend(loc="upper left", bbox_to_anchor=(1.01, 1),
                          fontsize=_TICK, frameon=False)

    if not _spend_ok and (monthly_spend_df is None or monthly_spend_df.empty):
        ax.text(0.5, 0.5, "No monthly spend data available",
                ha="center", va="center", transform=ax.transAxes,
                fontsize=_LABEL, color=DGRAY, style="italic")
        ax.axis("off")

    # Same fixed-margin approach as stacked chart — keeps chart area constant.
    fig.subplots_adjust(left=0.09, right=0.75, bottom=0.13, top=0.93)
    try:
        fig.savefig(out_path, dpi=160, facecolor=WHITE)
    finally:
        plt.close(fig)
    return out_path


# ────────────────────────────────────────────────────────────────────────────
# OPERATIONAL HEALTH – controller detail table   figsize = (CTRL_FW, CTRL_FH)
# ────────────────────────────────────────────────────────────────────────────
def controller_table_image(data, out_path: Path) -> Path:
    """
    Controller detail table (no metric boxes – those are separate tiles).
    Figure size: (9 × 3.5) inches.
    """
    fig = plt.figure(figsize=(CTRL_FW, CTRL_FH), facecolor=WHITE)
    ax  = fig.add_axes([0, 0, 1, 1])
    ax.set_xlim(0, CTRL_FW)
    ax.set_ylim(0, CTRL_FH)
    ax.axis("off")

    ax.text(0.15, CTRL_FH - 0.18, "Active Controllers",
            fontsize=_TITLE, fontweight="bold", color=NAVY)

    headers = ["#", "Support Status", "Controller IP",
               "Release No", "Version", "CoPilot"]
    col_x   = [0.15, 0.55, 2.25, 4.35, 5.60, 7.45]
    col_w   = [0.35, 1.65, 2.05, 1.20, 1.80, 1.40]

    hdr_y = CTRL_FH - 0.55
    for cx, cw, hdr in zip(col_x, col_w, headers):
        ax.add_patch(plt.Rectangle(
            (cx, hdr_y - 0.04), cw - 0.05, 0.30, color=MGRAY, zorder=0
        ))
        ax.text(cx + 0.06, hdr_y + 0.09, hdr,
                fontsize=_TH, color=DGRAY, va="center", fontweight="bold")

    rows = []
    if not data.df_detail.empty:
        for _, r in data.df_detail.iterrows():
            rows.append([
                str(r.get("Version Support Status", "")),
                str(r.get("Controller Ip", "")),
                str(r.get("Release No", "")),
                str(r.get("Controller Version", "")),
                str(r.get("Copilot Enabled", "")),
            ])

    row_h = 0.46
    for ri, row in enumerate(rows[:5]):
        ry = hdr_y - row_h * (ri + 1)
        bg = WHITE if ri % 2 == 0 else "#F9F9F9"
        ax.add_patch(plt.Rectangle(
            (0.10, ry - 0.06), CTRL_FW - 0.20, row_h - 0.02, color=bg, zorder=0
        ))
        ax.text(col_x[0] + 0.02, ry + 0.13, str(ri + 1),
                fontsize=_TC, color=DGRAY, va="center")
        for cx, cw, cell in zip(col_x[1:], col_w[1:], row):
            if len(cell) > 25:
                cell = cell[:23] + "…"
            ax.text(cx + 0.06, ry + 0.13, cell,
                    fontsize=_TC, color=NAVY, va="center")
        ax.axhline(ry - 0.06, xmin=0.015, xmax=0.99,
                   color="#DDDDDD", linewidth=0.6)

    if not rows:
        ax.text(CTRL_FW / 2, hdr_y - 0.6, "No controller data available",
                ha="center", fontsize=_LABEL, color=DGRAY, style="italic")

    try:
        fig.savefig(out_path, dpi=160, bbox_inches="tight", facecolor=WHITE)
    finally:
        plt.close(fig)
    return out_path


# ────────────────────────────────────────────────────────────────────────────
# SUPPORT – ticket requestor table               figsize = (REQ_FW, REQ_FH)
# ────────────────────────────────────────────────────────────────────────────
def ticket_requestor_table(df_requestor, out_path: Path) -> Path:
    """
    Table of ticket counts by requestor (replaces the raw ticket list).
    Figure size: (9 × 3.5) inches.
    Auto-detects name and count columns.
    """
    fig = plt.figure(figsize=(REQ_FW, REQ_FH), facecolor=WHITE)
    ax  = fig.add_axes([0, 0, 1, 1])
    ax.set_xlim(0, REQ_FW)
    ax.set_ylim(0, REQ_FH)
    ax.axis("off")

    ax.text(0.15, REQ_FH - 0.18, "Tickets by Requestor",
            fontsize=_TITLE, fontweight="bold", color=NAVY)

    if df_requestor is None or df_requestor.empty or len(df_requestor.columns) == 0:
        ax.text(REQ_FW / 2, REQ_FH / 2, "No requestor data available",
                ha="center", va="center", fontsize=_LABEL,
                color=DGRAY, style="italic")
        try:
            fig.savefig(out_path, dpi=160, bbox_inches="tight", facecolor=WHITE)
        finally:
            plt.close(fig)
        return out_path

    # Detect name and count columns (ensure they are distinct)
    cols = list(df_requestor.columns)
    name_col  = next((c for c in cols if any(k in c.lower()
                      for k in ("name", "requestor", "user", "submitter"))), cols[0])
    count_col = next((c for c in cols if c != name_col and any(k in c.lower()
                      for k in ("count", "ticket", "total", "number", "qty"))),
                     next((c for c in cols if c != name_col), None))
    if count_col is None:
        ax.text(REQ_FW / 2, REQ_FH / 2, "Insufficient columns in requestor data",
                ha="center", va="center", fontsize=_LABEL,
                color=DGRAY, style="italic")
        try:
            fig.savefig(out_path, dpi=160, bbox_inches="tight", facecolor=WHITE)
        finally:
            plt.close(fig)
        return out_path

    col_headers = ["#", "Requestor", "Tickets"]
    col_x = [0.15, 0.60, 7.80]
    col_w = [0.40, 7.15, 1.25]

    hdr_y = REQ_FH - 0.55
    for cx, cw, hdr in zip(col_x, col_w, col_headers):
        ax.add_patch(plt.Rectangle(
            (cx, hdr_y - 0.04), cw - 0.05, 0.30, color=MGRAY, zorder=0
        ))
        ax.text(cx + 0.06, hdr_y + 0.09, hdr,
                fontsize=_TH, color=DGRAY, va="center", fontweight="bold")

    rows_df = df_requestor[[name_col, count_col]].dropna()
    # Drop any row that looks like a repeated header (e.g. Name / Cases)
    rows_df = rows_df[
        ~rows_df[name_col].astype(str).str.strip().str.lower().isin(
            ("name", "requestor", "user", "submitter")
        )
    ]
    rows_df = rows_df.head(6)
    row_h = 0.44
    for ri, (_, row) in enumerate(rows_df.iterrows()):
        ry = hdr_y - row_h * (ri + 1)
        bg = WHITE if ri % 2 == 0 else "#F9F9F9"
        ax.add_patch(plt.Rectangle(
            (0.10, ry - 0.06), REQ_FW - 0.20, row_h - 0.02, color=bg, zorder=0
        ))
        ax.text(col_x[0] + 0.02, ry + 0.13, str(ri + 1),
                fontsize=_TC, color=DGRAY, va="center")
        name = str(row[name_col])
        if len(name) > 45:
            name = name[:43] + "…"
        ax.text(col_x[1] + 0.06, ry + 0.13, name,
                fontsize=_TC, color=NAVY, va="center")
        ax.text(col_x[2] + 0.06, ry + 0.13, str(row[count_col]),
                fontsize=_TC, color=NAVY, va="center", fontweight="bold")
        ax.axhline(ry - 0.06, xmin=0.015, xmax=0.99,
                   color="#DDDDDD", linewidth=0.5)

    try:
        fig.savefig(out_path, dpi=160, bbox_inches="tight", facecolor=WHITE)
    finally:
        plt.close(fig)
    return out_path


# ────────────────────────────────────────────────────────────────────────────
# SOFTWARE UPGRADES – release lifecycle table    figsize = (REL_FW, REL_FH)
# ────────────────────────────────────────────────────────────────────────────
def release_lifecycle_table(data, out_path: Path) -> Path:
    """
    Release Adoption Lifecycle table (no stat boxes – those are separate tiles).
    Figure size: (12 × 4.5) inches.
    """
    fig = plt.figure(figsize=(REL_FW, REL_FH), facecolor=WHITE)
    ax  = fig.add_axes([0, 0, 1, 1])
    ax.set_xlim(0, REL_FW)
    ax.set_ylim(0, REL_FH)
    ax.axis("off")

    ax.text(0.15, REL_FH - 0.18, "Release Adoption Lifecycle",
            fontsize=_TITLE, fontweight="bold", color=NAVY)

    col_headers = ["Controller IP", "Release", "Release Date",
                   "Adoption Date", "Days to Adopt", "Migrated Off", "Days in Release"]
    col_keys    = ["Controller Ip", "Release", "Release Date Date",
                   "Adoption Date", "Days to Adopt", "Migrated Off Date", "Days in Release"]
    col_x  = [0.15, 2.05, 3.40, 4.95, 6.50, 7.85, 9.55]
    col_w  = [1.85, 1.30, 1.50, 1.50, 1.30, 1.65, 1.75]

    hdr_y = REL_FH - 0.55
    for cx, cw, hdr in zip(col_x, col_w, col_headers):
        ax.add_patch(plt.Rectangle(
            (cx, hdr_y - 0.04), cw - 0.05, 0.32, color=MGRAY, zorder=0
        ))
        ax.text(cx + 0.06, hdr_y + 0.10, hdr,
                fontsize=_TH, color=DGRAY, va="center", fontweight="bold")

    rows = []
    df = data.df_release_lifecycle
    if not df.empty:
        if "Adoption Date" in df.columns:
            df = df.copy()
            df["_sort_date"] = pd.to_datetime(df["Adoption Date"], errors="coerce")
            df_s = df.sort_values("_sort_date", ascending=False)
        else:
            df_s = df
        for _, r in df_s.head(8).iterrows():
            rows.append([str(r.get(k, "")) for k in col_keys])

    row_h = 0.46
    for ri, row in enumerate(rows):
        ry = hdr_y - row_h * (ri + 1)
        bg = WHITE if ri % 2 == 0 else "#F9F9F9"
        ax.add_patch(plt.Rectangle(
            (0.10, ry - 0.06), REL_FW - 0.20, row_h - 0.02, color=bg, zorder=0
        ))
        ax.text(0.06, ry + 0.13, str(ri + 1),
                fontsize=_TC - 1, color=DGRAY, va="center")
        for cx, cw, cell in zip(col_x, col_w, row):
            if len(cell) > 22:
                cell = cell[:20] + "…"
            ax.text(cx + 0.06, ry + 0.13, cell,
                    fontsize=_TC, color=NAVY, va="center")
        ax.axhline(ry - 0.06, xmin=0.01, xmax=0.99,
                   color="#DDDDDD", linewidth=0.5)

    if not rows:
        ax.text(REL_FW / 2, hdr_y - 0.7,
                "No release lifecycle data available",
                ha="center", fontsize=_LABEL, color=DGRAY, style="italic")

    try:
        fig.savefig(out_path, dpi=160, bbox_inches="tight", facecolor=WHITE)
    finally:
        plt.close(fig)
    return out_path


# ────────────────────────────────────────────────────────────────────────────
# FEATURE REQUESTS TABLE                       figsize = (FR_FW, FR_FH)
# ────────────────────────────────────────────────────────────────────────────
def feature_requests_table(df_fr, out_path: Path) -> Path:
    """
    Table of feature requests showing Key, Summary, Issue Type, and Status.
    Figure size: (12 × 4.5) inches.
    """
    fig = plt.figure(figsize=(FR_FW, FR_FH), facecolor=WHITE)
    ax  = fig.add_axes([0, 0, 1, 1])
    ax.set_xlim(0, FR_FW)
    ax.set_ylim(0, FR_FH)
    ax.axis("off")

    ax.text(0.15, FR_FH - 0.18, "Feature Requests",
            fontsize=_TITLE, fontweight="bold", color=NAVY)

    if df_fr is None or df_fr.empty:
        ax.text(FR_FW / 2, FR_FH / 2, "No feature requests on record.",
                ha="center", va="center", fontsize=_LABEL,
                color=DGRAY, style="italic")
        try:
            fig.savefig(out_path, dpi=160, bbox_inches="tight", facecolor=WHITE)
        finally:
            plt.close(fig)
        return out_path

    # Detect columns (case-insensitive)
    cols_lower = {c.lower(): c for c in df_fr.columns}
    key_col    = cols_lower.get("key",         next((v for k, v in cols_lower.items() if "key"     in k), df_fr.columns[0]))
    sum_col    = cols_lower.get("summary",     next((v for k, v in cols_lower.items() if "summary" in k), None))
    type_col   = cols_lower.get("issue type",  next((v for k, v in cols_lower.items() if "type"    in k), None))
    status_col = cols_lower.get("status name", next((v for k, v in cols_lower.items() if "status"  in k), None))

    show_cols   = [c for c in [key_col, sum_col, type_col, status_col] if c is not None]
    col_headers = []
    col_x       = []
    col_w       = []

    # Layout: Key(1.4), Summary(fills), Issue Type(2.0), Status(2.0)
    x = 0.15
    layout = []
    if key_col    in show_cols: layout.append((key_col,    "Key",        1.35))
    if sum_col    in show_cols: layout.append((sum_col,    "Summary",    6.10))
    if type_col   in show_cols: layout.append((type_col,   "Issue Type", 2.10))
    if status_col in show_cols: layout.append((status_col, "Status",     2.10))

    for data_col, header, w in layout:
        col_headers.append(header)
        col_x.append(x)
        col_w.append(w)
        x += w

    hdr_y = FR_FH - 0.55
    for cx, cw, hdr in zip(col_x, col_w, col_headers):
        ax.add_patch(plt.Rectangle(
            (cx, hdr_y - 0.04), cw - 0.05, 0.32, color=MGRAY, zorder=0
        ))
        ax.text(cx + 0.06, hdr_y + 0.10, hdr,
                fontsize=_TH, color=DGRAY, va="center", fontweight="bold")

    data_cols = [item[0] for item in layout]
    rows_df = df_fr[data_cols].dropna(how="all").head(8)

    if rows_df.empty:
        ax.text(FR_FW / 2, FR_FH / 2, "No feature requests on record.",
                ha="center", va="center", fontsize=_LABEL,
                color=DGRAY, style="italic")
        try:
            fig.savefig(out_path, dpi=160, bbox_inches="tight", facecolor=WHITE)
        finally:
            plt.close(fig)
        return out_path

    row_h = 0.46
    for ri, (_, row) in enumerate(rows_df.iterrows()):
        ry = hdr_y - row_h * (ri + 1)
        bg = WHITE if ri % 2 == 0 else "#F9F9F9"
        ax.add_patch(plt.Rectangle(
            (0.10, ry - 0.06), FR_FW - 0.20, row_h - 0.02, color=bg, zorder=0
        ))
        for cx, cw, data_col, header in zip(col_x, col_w, data_cols, col_headers):
            cell_val = str(row[data_col]) if pd.notna(row[data_col]) else ""
            max_chars = int(cw * 8)  # rough chars-per-inch estimate
            if len(cell_val) > max_chars:
                cell_val = cell_val[:max_chars - 1] + "…"
            ax.text(cx + 0.06, ry + 0.13, cell_val,
                    fontsize=_TC, color=NAVY, va="center")
        ax.axhline(ry - 0.06, xmin=0.01, xmax=0.99,
                   color="#DDDDDD", linewidth=0.5)

    try:
        fig.savefig(out_path, dpi=160, bbox_inches="tight", facecolor=WHITE)
    finally:
        plt.close(fig)
    return out_path


# ────────────────────────────────────────────────────────────────────────────
# USE CASES TABLE                              figsize = (UC_FW, UC_FH)
# ────────────────────────────────────────────────────────────────────────────
def use_cases_image(data, out_path: Path) -> Path:
    """
    Two-column status table: Features Enabled (left) + Use Cases (right).
    Each row shows a green tick or red cross circle next to the item name.
    Figure size: (8.5 × 4.6) inches.
    """
    PANEL_BG  = "#F4F5F7"
    HEADER_BG = "#E2E4E9"
    GREEN     = "#27AE60"
    _RED      = "#E74C3C"
    TEXT_COL  = "#1E2D3D"

    features = [
        ("Copilot Enabled",            data.copilot_enabled_label == "Y"),
        ("DCF Enabled",                data.dcf_enabled == "Y"),
        ("Firenet Enabled",            data.firenet_enabled == "Y"),
        ("CloudN / Edge Enabled",      data.cloudn_edge == "Y"),
        ("Multi-Region Transit",       data.multi_region == "Y"),
    ]
    use_cases = [
        ("Unified Cloud NW Fabric",    data.uc_unified_nw == "Y"),
        ("Prevent Lateral Movement",   data.uc_lateral == "Y"),
        ("Zero Trust NW Segmentation", data.uc_zt_seg == "Y"),
        ("Secure 3rd Party Access",    data.uc_3rd_party == "Y"),
        ("End-to-End Encryption",      data.uc_e2e_enc == "Y"),
        ("Block Data Exfiltration",    data.uc_block_exfil == "Y"),
        ("Secure Dev Velocity",        data.uc_dev_velocity == "Y"),
    ]

    fw, fh = UC_FW, UC_FH
    fig = plt.figure(figsize=(fw, fh), facecolor=PANEL_BG)
    ax  = fig.add_axes([0, 0, 1, 1])
    ax.set_xlim(0, fw)
    ax.set_ylim(0, fh)
    ax.axis("off")
    ax.set_facecolor(PANEL_BG)

    def _panel(items, panel_label, px, py, pw, ph):
        # Panel background
        ax.add_patch(mpatches.FancyBboxPatch(
            (px, py), pw, ph,
            boxstyle="round,pad=0.04", linewidth=0,
            facecolor=WHITE, zorder=1,
        ))
        # Header bar
        hh = 0.38
        ax.add_patch(plt.Rectangle((px, py + ph - hh), pw, hh,
                                   facecolor=HEADER_BG, zorder=2))
        ax.text(px + 0.22, py + ph - hh / 2, "Status",
                ha="left", va="center", fontsize=8.5, fontweight="bold",
                color=DGRAY, zorder=3)
        ax.text(px + 0.75, py + ph - hh / 2, panel_label,
                ha="left", va="center", fontsize=9.5, fontweight="bold",
                color=TEXT_COL, zorder=3)

        row_h = (ph - hh) / max(len(items), 1)
        for i, (name, active) in enumerate(items):
            ry = py + ph - hh - (i + 1) * row_h
            if i % 2 == 0:
                ax.add_patch(plt.Rectangle((px, ry), pw, row_h,
                                           facecolor="#F9FAFB", zorder=2))
            cx, cy, cr = px + 0.35, ry + row_h / 2, 0.13
            circle = plt.Circle((cx, cy), cr,
                                 color=GREEN if active else _RED, zorder=3)
            ax.add_patch(circle)
            sym = "✓" if active else "✗"
            ax.text(cx, cy, sym, ha="center", va="center",
                    fontsize=9, fontweight="bold", color=WHITE, zorder=4)
            ax.text(px + 0.67, ry + row_h / 2, name,
                    ha="left", va="center", fontsize=9.5,
                    color=TEXT_COL, zorder=3)

    pad = 0.20
    gap = 0.30
    panel_w = (fw - 2 * pad - gap) / 2
    panel_h = fh - 2 * pad

    _panel(features,  "Features Enabled", pad,                 pad, panel_w, panel_h)
    _panel(use_cases, "Use Case",         pad + panel_w + gap, pad, panel_w, panel_h)

    try:
        fig.savefig(out_path, dpi=160, bbox_inches="tight", facecolor=PANEL_BG)
    finally:
        plt.close(fig)
    return out_path


# ────────────────────────────────────────────────────────────────────────────
# MANUAL SLIDE PLACEHOLDER
# ────────────────────────────────────────────────────────────────────────────
def manual_slide_placeholder(out_path: Path, short: bool = False) -> Path:
    """
    Prominent red-text image for slides that require Sales Team completion.

    short=False  →  figsize (MANUAL_FW × MANUAL_FH)  = 11 × 3.5
    short=True   →  figsize (BANNER_FW × BANNER_FH)  = 11 × 1.5  (for ZTMM)
    """
    fw = BANNER_FW if short else MANUAL_FW
    fh = BANNER_FH if short else MANUAL_FH

    fig = plt.figure(figsize=(fw, fh), facecolor=WHITE)
    ax  = fig.add_axes([0, 0, 1, 1])
    ax.set_xlim(0, fw)
    ax.set_ylim(0, fh)
    ax.axis("off")

    pad_x, pad_y = 0.25, 0.10
    ax.add_patch(mpatches.FancyBboxPatch(
        (pad_x, pad_y), fw - 2 * pad_x, fh - 2 * pad_y,
        boxstyle="round,pad=0.08", linewidth=2.0,
        edgecolor=RED, facecolor="#FFF0F0",
    ))
    # Orange accent bar along the top edge of the box
    ax.add_patch(plt.Rectangle(
        (pad_x + 0.12, fh - pad_y - 0.12), fw - 2 * (pad_x + 0.12), 0.07,
        color=ORANGE,
    ))

    if short:
        ax.text(fw / 2, fh / 2,
                "To Be Completed by Sales Team",
                ha="center", va="center",
                fontsize=20, fontweight="bold", color=RED)
    else:
        ax.text(fw / 2, fh * 0.62,
                "To Be Completed",
                ha="center", va="center",
                fontsize=38, fontweight="bold", color=RED)
        ax.text(fw / 2, fh * 0.28,
                "by Sales Team",
                ha="center", va="center",
                fontsize=28, fontweight="bold", color=RED)

    try:
        fig.savefig(out_path, dpi=160, bbox_inches="tight", facecolor=WHITE)
    finally:
        plt.close(fig)
    return out_path
