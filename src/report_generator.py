from __future__ import annotations

import logging
from datetime import datetime
from pathlib import Path
from typing import Iterable

from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows
import pandas as pd


PALETTE = {
    "navy": "0F172A",
    "blue": "2563EB",
    "teal": "0F766E",
    "emerald": "15803D",
    "amber": "F59E0B",
    "orange": "EA580C",
    "slate": "475569",
    "sky": "E0F2FE",
    "mint": "D1FAE5",
    "sand": "FEF3C7",
    "rose": "FFEDD5",
    "mist": "F1F5F9",
    "white": "FFFFFF",
    "line": "CBD5E1",
}

THIN_BORDER = Border(
    left=Side(style="thin", color=PALETTE["line"]),
    right=Side(style="thin", color=PALETTE["line"]),
    top=Side(style="thin", color=PALETTE["line"]),
    bottom=Side(style="thin", color=PALETTE["line"]),
)


def _set_col_widths(ws, widths: dict[str, int]) -> None:
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


def _currency_token(currency: str) -> str:
    code = str(currency).strip().upper() or "USD"
    symbols = {
        "USD": "$",
        "EUR": "EUR",
        "GBP": "GBP",
        "JPY": "JPY",
        "ARS": "ARS",
        "MXN": "MXN",
        "GTQ": "GTQ",
    }
    return symbols.get(code, code)


def _currency_number_format(currency: str) -> str:
    token = _currency_token(currency)
    return f'"{token}" #,##0.00;[Red]-"{token}" #,##0.00'


def _currency_text(value: float, currency: str) -> str:
    token = _currency_token(currency)
    if token == "$":
        return f"{token}{value:,.2f}"
    return f"{token} {value:,.2f}"


def _apply_number_format_to_column(ws, col_idx: int, number_format: str, start_row: int = 2) -> None:
    for row_idx in range(start_row, ws.max_row + 1):
        ws.cell(row=row_idx, column=col_idx).number_format = number_format


def _header_positions(ws) -> dict[str, int]:
    headers: dict[str, int] = {}
    for col_idx in range(1, ws.max_column + 1):
        value = ws.cell(row=1, column=col_idx).value
        if value is not None:
            headers[str(value).strip()] = col_idx
    return headers


def _fill_range(
    ws,
    start_row: int,
    end_row: int,
    start_col: int,
    end_col: int,
    *,
    fill: PatternFill | None = None,
    font: Font | None = None,
    alignment: Alignment | None = None,
    border: Border | None = None,
) -> None:
    for row in ws.iter_rows(
        min_row=start_row,
        max_row=end_row,
        min_col=start_col,
        max_col=end_col,
    ):
        for cell in row:
            if fill is not None:
                cell.fill = fill
            if font is not None:
                cell.font = font
            if alignment is not None:
                cell.alignment = alignment
            if border is not None:
                cell.border = border


def _merge_block(
    ws,
    start_row: int,
    end_row: int,
    start_col: int,
    end_col: int,
    value,
    *,
    fill_color: str,
    font: Font,
    alignment: Alignment,
    border: Border = THIN_BORDER,
    number_format: str | None = None,
) -> None:
    ws.merge_cells(
        start_row=start_row,
        start_column=start_col,
        end_row=end_row,
        end_column=end_col,
    )
    _fill_range(
        ws,
        start_row,
        end_row,
        start_col,
        end_col,
        fill=PatternFill("solid", fgColor=fill_color),
        font=font,
        alignment=alignment,
        border=border,
    )
    cell = ws.cell(row=start_row, column=start_col)
    cell.value = value
    if number_format is not None:
        cell.number_format = number_format


def _auto_fit_widths(ws, *, min_width: int = 12, max_width: int = 28, extra: int = 2) -> None:
    widths: dict[int, int] = {}
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is None:
                continue
            text_length = len(str(cell.value))
            widths[cell.column] = max(widths.get(cell.column, min_width), min(text_length + extra, max_width))
    for col_idx, width in widths.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = width


def _style_table(ws) -> None:
    if ws.max_row == 0 or ws.max_column == 0:
        return

    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    ws.row_dimensions[1].height = 24

    for cell in ws[1]:
        cell.fill = PatternFill("solid", fgColor=PALETTE["navy"])
        cell.font = Font(color=PALETTE["white"], bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = THIN_BORDER

    band_a = PatternFill("solid", fgColor=PALETTE["white"])
    band_b = PatternFill("solid", fgColor=PALETTE["mist"])
    for row_idx in range(2, ws.max_row + 1):
        fill = band_a if row_idx % 2 == 0 else band_b
        for col_idx in range(1, ws.max_column + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.fill = fill
            cell.border = THIN_BORDER
            if cell.alignment == Alignment():
                cell.alignment = Alignment(vertical="center")


def _format_date_label(value: str | None) -> str:
    if not value:
        return "N/A"
    try:
        return pd.to_datetime(value).strftime("%b %d, %Y")
    except Exception:
        return str(value)


def _write_dataframe_sheet(
    wb: Workbook,
    title: str,
    df: pd.DataFrame,
    *,
    currency_fmt: str,
    date_cols: Iterable[str] = (),
    currency_cols: Iterable[str] = (),
    percent_cols: Iterable[str] = (),
    decimal_cols: Iterable[str] = (),
    integer_cols: Iterable[str] = (),
) -> None:
    ws = wb.create_sheet(title)
    for row in dataframe_to_rows(df, index=False, header=True):
        ws.append(row)

    headers = _header_positions(ws)
    for col_name in date_cols:
        if col_name in headers:
            _apply_number_format_to_column(ws, headers[col_name], "yyyy-mm-dd")
    for col_name in currency_cols:
        if col_name in headers:
            _apply_number_format_to_column(ws, headers[col_name], currency_fmt)
    for col_name in percent_cols:
        if col_name in headers:
            _apply_number_format_to_column(ws, headers[col_name], '0.0"%"')
    for col_name in decimal_cols:
        if col_name in headers:
            _apply_number_format_to_column(ws, headers[col_name], "#,##0.00")
    for col_name in integer_cols:
        if col_name in headers:
            _apply_number_format_to_column(ws, headers[col_name], "#,##0")

    _style_table(ws)
    _auto_fit_widths(ws)


def _write_kpi_card(
    ws,
    *,
    start_row: int,
    start_col: int,
    end_col: int,
    title: str,
    value,
    note: str,
    fill_color: str,
    number_format: str | None = None,
) -> None:
    _merge_block(
        ws,
        start_row,
        start_row,
        start_col,
        end_col,
        title,
        fill_color=fill_color,
        font=Font(color=PALETTE["white"], bold=True, size=11),
        alignment=Alignment(horizontal="center", vertical="center"),
    )
    _merge_block(
        ws,
        start_row + 1,
        start_row + 2,
        start_col,
        end_col,
        value,
        fill_color=fill_color,
        font=Font(color=PALETTE["white"], bold=True, size=24),
        alignment=Alignment(horizontal="center", vertical="center"),
        number_format=number_format,
    )
    _merge_block(
        ws,
        start_row + 3,
        start_row + 3,
        start_col,
        end_col,
        note,
        fill_color=fill_color,
        font=Font(color=PALETTE["white"], italic=True, size=10),
        alignment=Alignment(horizontal="center", vertical="center"),
    )


def _write_insight_tile(
    ws,
    *,
    start_row: int,
    start_col: int,
    end_col: int,
    title: str,
    value: str,
    note: str,
    fill_color: str,
) -> None:
    _merge_block(
        ws,
        start_row,
        start_row,
        start_col,
        end_col,
        title,
        fill_color=fill_color,
        font=Font(color=PALETTE["navy"], bold=True, size=10),
        alignment=Alignment(horizontal="center", vertical="center"),
    )
    _merge_block(
        ws,
        start_row + 1,
        start_row + 2,
        start_col,
        end_col,
        value,
        fill_color=fill_color,
        font=Font(color=PALETTE["navy"], bold=True, size=16),
        alignment=Alignment(horizontal="center", vertical="center", wrap_text=True),
    )
    _merge_block(
        ws,
        start_row + 3,
        start_row + 3,
        start_col,
        end_col,
        note,
        fill_color=fill_color,
        font=Font(color=PALETTE["slate"], italic=True, size=9),
        alignment=Alignment(horizontal="center", vertical="center", wrap_text=True),
    )


def _embed_chart(ws, image_path: Path, anchor: str, *, width: int, height: int) -> None:
    try:
        image = XLImage(str(image_path))
        image.width = width
        image.height = height
        image.anchor = anchor
        ws.add_image(image)
    except Exception as exc:
        logging.warning("Could not embed chart '%s': %s", image_path, exc)


def _build_summary_sheet(
    ws,
    *,
    kpis: dict,
    currency: str,
    sources: list[str],
    chart_files: list[Path] | None,
) -> None:
    currency_fmt = _currency_number_format(currency)
    total_revenue = float(kpis.get("total_revenue", 0.0))
    total_profit = float(kpis.get("total_profit", 0.0))
    total_cost = float(kpis.get("total_cost", 0.0))
    total_orders = int(kpis.get("total_orders", 0))
    total_units = float(kpis.get("total_units", 0.0))
    avg_order_value = float(kpis.get("avg_order_value", 0.0))
    margin_pct = float(kpis.get("margin_pct", 0.0))
    unique_products = int(kpis.get("unique_products", 0))
    unique_skus = int(kpis.get("unique_skus", 0))
    best_product = kpis.get("best_product") or {}
    best_day = kpis.get("best_day") or {}
    coverage = (
        f"{_format_date_label(kpis.get('date_start'))} to {_format_date_label(kpis.get('date_end'))}"
        if kpis.get("date_start") and kpis.get("date_end")
        else "Coverage unavailable"
    )
    generated_at = datetime.now().strftime("%Y-%m-%d %H:%M")
    top_product_name = str(best_product.get("product", "N/A"))
    top_product_revenue = float(best_product.get("revenue", 0.0))
    top_product_share = (top_product_revenue / total_revenue * 100) if total_revenue else 0.0

    ws.title = "Summary"
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = PALETTE["blue"]
    _set_col_widths(
        ws,
        {
            "A": 14,
            "B": 14,
            "C": 14,
            "D": 14,
            "E": 14,
            "F": 14,
            "G": 14,
            "H": 14,
            "I": 14,
            "J": 14,
            "K": 14,
            "L": 14,
        },
    )

    _fill_range(
        ws,
        1,
        60,
        1,
        12,
        fill=PatternFill("solid", fgColor=PALETTE["mist"]),
    )

    _merge_block(
        ws,
        1,
        2,
        1,
        12,
        f"Sales Performance Dashboard ({currency})",
        fill_color=PALETTE["navy"],
        font=Font(color=PALETTE["white"], bold=True, size=22),
        alignment=Alignment(horizontal="center", vertical="center"),
    )
    _merge_block(
        ws,
        3,
        3,
        1,
        12,
        f"Generated on {generated_at} | {len(sources)} source file(s) | {coverage}",
        fill_color=PALETTE["slate"],
        font=Font(color=PALETTE["white"], size=10),
        alignment=Alignment(horizontal="center", vertical="center"),
    )

    _write_kpi_card(
        ws,
        start_row=5,
        start_col=1,
        end_col=4,
        title="Total Orders",
        value=total_orders,
        note="Commercial transactions processed",
        fill_color=PALETTE["blue"],
        number_format="#,##0",
    )
    _write_kpi_card(
        ws,
        start_row=5,
        start_col=5,
        end_col=8,
        title="Total Revenue",
        value=total_revenue,
        note="Gross sales generated in the period",
        fill_color=PALETTE["navy"],
        number_format=currency_fmt,
    )
    _write_kpi_card(
        ws,
        start_row=5,
        start_col=9,
        end_col=12,
        title="Total Profit",
        value=total_profit,
        note="Bottom-line contribution after cost",
        fill_color=PALETTE["emerald"],
        number_format=currency_fmt,
    )
    _write_kpi_card(
        ws,
        start_row=10,
        start_col=1,
        end_col=4,
        title="Avg Order Value",
        value=avg_order_value,
        note="Average ticket per order",
        fill_color=PALETTE["amber"],
        number_format=currency_fmt,
    )
    _write_kpi_card(
        ws,
        start_row=10,
        start_col=5,
        end_col=8,
        title="Active Products",
        value=f"{unique_products:,}",
        note=f"{unique_skus:,} SKUs represented in the dataset",
        fill_color=PALETTE["teal"],
    )
    _write_kpi_card(
        ws,
        start_row=10,
        start_col=9,
        end_col=12,
        title="Margin Rate",
        value=margin_pct,
        note=f"Total cost base: {_currency_text(total_cost, currency)}",
        fill_color=PALETTE["orange"],
        number_format='0.0"%"',
    )

    _merge_block(
        ws,
        15,
        15,
        1,
        12,
        "Executive Summary",
        fill_color=PALETTE["navy"],
        font=Font(color=PALETTE["white"], bold=True, size=12),
        alignment=Alignment(horizontal="left", vertical="center"),
    )
    executive_summary = (
        f"This reporting window generated {_currency_text(total_revenue, currency)} in revenue from "
        f"{total_orders:,} orders and {total_units:,.0f} units. Profit closed at "
        f"{_currency_text(total_profit, currency)} with an overall margin of {margin_pct:.1f}%.\n\n"
        f"The catalog leader was {top_product_name}, contributing {_currency_text(top_product_revenue, currency)} "
        f"and {top_product_share:.1f}% of total revenue. "
        f"The strongest day landed on {_format_date_label(best_day.get('date'))}, reaching "
        f"{_currency_text(float(best_day.get('revenue', 0.0)), currency)}."
    )
    _merge_block(
        ws,
        16,
        18,
        1,
        12,
        executive_summary,
        fill_color=PALETTE["white"],
        font=Font(color=PALETTE["navy"], size=11),
        alignment=Alignment(horizontal="left", vertical="top", wrap_text=True),
    )

    _write_insight_tile(
        ws,
        start_row=20,
        start_col=1,
        end_col=3,
        title="Revenue Leader",
        value=top_product_name,
        note=f"{_currency_text(top_product_revenue, currency)} in revenue",
        fill_color=PALETTE["sky"],
    )
    _write_insight_tile(
        ws,
        start_row=20,
        start_col=4,
        end_col=6,
        title="Peak Day",
        value=_format_date_label(best_day.get("date")),
        note=f"{_currency_text(float(best_day.get('revenue', 0.0)), currency)} generated",
        fill_color=PALETTE["mint"],
    )
    _write_insight_tile(
        ws,
        start_row=20,
        start_col=7,
        end_col=9,
        title="Revenue Share",
        value=f"{top_product_share:.1f}%",
        note="Contribution from the leading product",
        fill_color=PALETTE["sand"],
    )
    _write_insight_tile(
        ws,
        start_row=20,
        start_col=10,
        end_col=12,
        title="Catalog Spread",
        value=f"{unique_products:,} products",
        note=f"{unique_skus:,} SKUs active in the report",
        fill_color=PALETTE["rose"],
    )

    _merge_block(
        ws,
        25,
        25,
        1,
        12,
        "Visual Story",
        fill_color=PALETTE["navy"],
        font=Font(color=PALETTE["white"], bold=True, size=12),
        alignment=Alignment(horizontal="left", vertical="center"),
    )
    if chart_files:
        for image_path, anchor in zip(chart_files[:2], ["A27", "G27"], strict=False):
            _embed_chart(ws, image_path, anchor, width=520, height=280)

    source_start_row = 47
    _merge_block(
        ws,
        source_start_row,
        source_start_row,
        1,
        12,
        "Source Files",
        fill_color=PALETTE["navy"],
        font=Font(color=PALETTE["white"], bold=True, size=12),
        alignment=Alignment(horizontal="left", vertical="center"),
    )
    for idx, source in enumerate(sources, start=source_start_row + 1):
        _merge_block(
            ws,
            idx,
            idx,
            1,
            12,
            source,
            fill_color=PALETTE["white"],
            font=Font(color=PALETTE["navy"], size=10),
            alignment=Alignment(horizontal="left", vertical="center"),
        )


def _build_charts_sheet(wb: Workbook, chart_files: list[Path]) -> None:
    if not chart_files:
        return

    ws = wb.create_sheet("Charts")
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = PALETTE["teal"]
    _set_col_widths(ws, {get_column_letter(col): 14 for col in range(1, 15)})
    _fill_range(
        ws,
        1,
        50,
        1,
        14,
        fill=PatternFill("solid", fgColor=PALETTE["mist"]),
    )
    _merge_block(
        ws,
        1,
        2,
        1,
        14,
        "Chart Gallery",
        fill_color=PALETTE["navy"],
        font=Font(color=PALETTE["white"], bold=True, size=20),
        alignment=Alignment(horizontal="center", vertical="center"),
    )
    _merge_block(
        ws,
        3,
        3,
        1,
        14,
        "A clean set of visuals for fast commercial review.",
        fill_color=PALETTE["slate"],
        font=Font(color=PALETTE["white"], size=10),
        alignment=Alignment(horizontal="center", vertical="center"),
    )

    anchors = ["A5", "H5", "A24", "H24"]
    for image_path, anchor in zip(chart_files[:4], anchors, strict=False):
        _embed_chart(ws, image_path, anchor, width=600, height=320)


def generate_excel_report(
    report_path: Path,
    df_clean: pd.DataFrame,
    kpis: dict,
    currency: str,
    sources: Iterable[str],
    chart_files: list[Path] | None = None,
) -> None:
    """
    Creates an Excel report with:
    - Summary dashboard
    - Cleaned Data sheet
    - Product analysis sheets
    - Optional chart gallery
    """
    report_path.parent.mkdir(parents=True, exist_ok=True)

    wb = Workbook()
    currency_fmt = _currency_number_format(currency)
    source_list = list(sources)

    summary_ws = wb.active
    _build_summary_sheet(
        summary_ws,
        kpis=kpis,
        currency=currency,
        sources=source_list,
        chart_files=chart_files,
    )

    df_out = kpis.get("df_with_calculations")
    if not isinstance(df_out, pd.DataFrame):
        df_out = df_clean.copy()
        if {"quantity", "unit_price", "unit_cost"}.issubset(df_out.columns):
            if "revenue" not in df_out.columns:
                df_out["revenue"] = df_out["quantity"] * df_out["unit_price"]
            if "cost" not in df_out.columns:
                df_out["cost"] = df_out["quantity"] * df_out["unit_cost"]
            if "profit" not in df_out.columns:
                df_out["profit"] = df_out["revenue"] - df_out["cost"]

    _write_dataframe_sheet(
        wb,
        "Cleaned Data",
        df_out,
        currency_fmt=currency_fmt,
        date_cols=["date"],
        currency_cols=["unit_price", "unit_cost", "revenue", "cost", "profit"],
        percent_cols=["margin_pct"],
        decimal_cols=["quantity"],
    )

    product_performance = kpis.get("product_performance")
    if isinstance(product_performance, pd.DataFrame) and not product_performance.empty:
        top_products_df = product_performance.head(10).copy()
        _write_dataframe_sheet(
            wb,
            "Top Products",
            top_products_df,
            currency_fmt=currency_fmt,
            currency_cols=["revenue", "cost", "profit"],
            percent_cols=["margin_pct"],
            decimal_cols=["total_units"],
            integer_cols=["order_count"],
        )
        _write_dataframe_sheet(
            wb,
            "Product Performance",
            product_performance,
            currency_fmt=currency_fmt,
            currency_cols=["revenue", "cost", "profit"],
            percent_cols=["margin_pct"],
            decimal_cols=["total_units"],
            integer_cols=["order_count"],
        )

    daily_performance = kpis.get("daily_performance")
    if isinstance(daily_performance, pd.DataFrame) and not daily_performance.empty:
        _write_dataframe_sheet(
            wb,
            "Daily Performance",
            daily_performance,
            currency_fmt=currency_fmt,
            date_cols=["date"],
            currency_cols=["revenue", "cost", "profit"],
            percent_cols=["margin_pct"],
            decimal_cols=["units"],
            integer_cols=["orders"],
        )

    if chart_files:
        _build_charts_sheet(wb, chart_files)

    wb.save(report_path)
