from __future__ import annotations

from pathlib import Path

import matplotlib
import pandas as pd

# Use a non-interactive backend so chart generation works in CI/headless environments.
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import ticker


PALETTE = {
    "navy": "#0F172A",
    "blue": "#2563EB",
    "teal": "#0F766E",
    "emerald": "#15803D",
    "amber": "#F59E0B",
    "orange": "#EA580C",
    "slate": "#475569",
    "muted": "#94A3B8",
    "panel": "#F8FAFC",
    "grid": "#D7E3F4",
}

SERIES_COLORS = [
    PALETTE["blue"],
    PALETTE["teal"],
    PALETTE["amber"],
    PALETTE["orange"],
    PALETTE["emerald"],
    "#7C3AED",
]

matplotlib.rcParams.update(
    {
        "axes.facecolor": PALETTE["panel"],
        "figure.facecolor": "white",
        "axes.edgecolor": PALETTE["grid"],
        "axes.labelcolor": PALETTE["slate"],
        "axes.titlecolor": PALETTE["navy"],
        "xtick.color": PALETTE["slate"],
        "ytick.color": PALETTE["slate"],
        "font.size": 11,
        "font.family": "DejaVu Sans",
        "grid.color": PALETTE["grid"],
        "grid.alpha": 0.8,
        "axes.grid": True,
        "axes.grid.axis": "y",
    }
)


def _currency_axis() -> ticker.FuncFormatter:
    return ticker.FuncFormatter(lambda value, _: f"{value:,.0f}")


def _polish_axes(ax, title: str, subtitle: str, xlabel: str, ylabel: str) -> None:
    ax.set_title(title, loc="left", fontsize=16, fontweight="bold", pad=18)
    ax.text(
        0.0,
        1.02,
        subtitle,
        transform=ax.transAxes,
        fontsize=10,
        color=PALETTE["slate"],
        va="bottom",
    )
    ax.set_xlabel(xlabel)
    ax.set_ylabel(ylabel)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.yaxis.set_major_formatter(_currency_axis() if ylabel in {"Revenue", "Profit"} else ax.yaxis.get_major_formatter())


def _save_figure(fig: plt.Figure, output_path: Path) -> Path:
    fig.tight_layout()
    fig.savefig(output_path, dpi=170, bbox_inches="tight")
    plt.close(fig)
    return output_path


def _chartable_daily_performance(kpis: dict) -> pd.DataFrame:
    daily = kpis.get("daily_performance")
    if isinstance(daily, pd.DataFrame):
        return daily.copy()
    return pd.DataFrame()


def _chartable_product_performance(kpis: dict) -> pd.DataFrame:
    products = kpis.get("product_performance")
    if isinstance(products, pd.DataFrame):
        return products.copy()
    return pd.DataFrame()


def build_charts(df_clean: pd.DataFrame, kpis: dict, charts_dir: Path) -> list[Path]:
    charts_dir.mkdir(parents=True, exist_ok=True)

    chart_files: list[Path] = []
    daily = _chartable_daily_performance(kpis)
    products = _chartable_product_performance(kpis)

    # 1) Revenue by day
    if not daily.empty:
        p = charts_dir / "revenue_by_day.png"
        fig, ax = plt.subplots(figsize=(10, 5.2))
        dates = pd.to_datetime(daily["date"])
        revenue = daily["revenue"].astype(float)
        ax.plot(
            dates,
            revenue,
            color=PALETTE["blue"],
            linewidth=3,
            marker="o",
            markersize=6,
        )
        ax.fill_between(dates, revenue, color=PALETTE["blue"], alpha=0.16)
        ax.set_xticks(dates)
        ax.tick_params(axis="x", rotation=35)
        _polish_axes(
            ax,
            title="Revenue by Day",
            subtitle="A clean daily trend to show sales momentum over time.",
            xlabel="Date",
            ylabel="Revenue",
        )
        chart_files.append(_save_figure(fig, p))

    # 2) Top products by revenue
    if not products.empty:
        top_revenue = products.head(min(len(products), 8)).sort_values("revenue", ascending=True)
        p = charts_dir / "top_products.png"
        fig, ax = plt.subplots(figsize=(10, 5.2))
        bars = ax.barh(
            top_revenue["product"].astype(str),
            top_revenue["revenue"].astype(float),
            color=[SERIES_COLORS[idx % len(SERIES_COLORS)] for idx in range(len(top_revenue))],
            edgecolor="white",
            linewidth=1.2,
        )
        for bar, value in zip(bars, top_revenue["revenue"].astype(float), strict=False):
            ax.text(
                bar.get_width(),
                bar.get_y() + bar.get_height() / 2,
                f"  {value:,.0f}",
                va="center",
                ha="left",
                color=PALETTE["navy"],
                fontsize=10,
                fontweight="bold",
            )
        _polish_axes(
            ax,
            title="Top Products by Revenue",
            subtitle="Highlights the products driving the largest commercial impact.",
            xlabel="Revenue",
            ylabel="",
        )
        ax.xaxis.set_major_formatter(_currency_axis())
        chart_files.append(_save_figure(fig, p))

        # 3) Profit by product
        p = charts_dir / "profit_by_product.png"
        top_profit = products.head(min(len(products), 8)).copy()
        fig, ax = plt.subplots(figsize=(10, 5.2))
        bar_colors = [
            PALETTE["emerald"] if value >= 0 else "#DC2626"
            for value in top_profit["profit"].astype(float)
        ]
        bars = ax.bar(
            top_profit["product"].astype(str),
            top_profit["profit"].astype(float),
            color=bar_colors,
            edgecolor="white",
            linewidth=1.2,
        )
        for bar, value in zip(bars, top_profit["profit"].astype(float), strict=False):
            offset = 0.02 * max(abs(float(top_profit["profit"].max() or 1.0)), 1.0)
            ax.text(
                bar.get_x() + bar.get_width() / 2,
                value + offset,
                f"{value:,.0f}",
                va="bottom",
                ha="center",
                color=PALETTE["navy"],
                fontsize=10,
                fontweight="bold",
            )
        ax.tick_params(axis="x", rotation=25)
        _polish_axes(
            ax,
            title="Profit by Product",
            subtitle="Makes profitable lines stand out from low-yield catalog entries.",
            xlabel="Product",
            ylabel="Profit",
        )
        chart_files.append(_save_figure(fig, p))

        # 4) Revenue mix
        p = charts_dir / "revenue_mix.png"
        mix = products.head(min(len(products), 5)).copy()
        others_revenue = float(products.iloc[5:]["revenue"].sum()) if len(products) > 5 else 0.0
        if others_revenue > 0:
            mix.loc[len(mix)] = {
                "product": "Other",
                "order_count": 0,
                "total_units": 0,
                "revenue": others_revenue,
                "cost": 0.0,
                "profit": 0.0,
                "margin_pct": 0.0,
            }

        fig, ax = plt.subplots(figsize=(10, 5.2))
        wedges, _, autotexts = ax.pie(
            mix["revenue"].astype(float),
            labels=mix["product"].astype(str),
            colors=[SERIES_COLORS[idx % len(SERIES_COLORS)] for idx in range(len(mix))],
            startangle=90,
            wedgeprops={"width": 0.45, "edgecolor": "white"},
            autopct=lambda pct: f"{pct:.0f}%" if pct >= 4 else "",
            pctdistance=0.78,
        )
        for text in autotexts:
            text.set_color("white")
            text.set_fontweight("bold")
        ax.text(
            0,
            0,
            "Revenue\nMix",
            ha="center",
            va="center",
            fontsize=16,
            fontweight="bold",
            color=PALETTE["navy"],
        )
        ax.set_aspect("equal")
        ax.set_title("Revenue Mix", loc="left", fontsize=16, fontweight="bold", pad=18)
        ax.text(
            0.0,
            1.02,
            "Shows how concentrated the business is across the top products.",
            transform=ax.transAxes,
            fontsize=10,
            color=PALETTE["slate"],
            va="bottom",
        )
        chart_files.append(_save_figure(fig, p))

    return chart_files
