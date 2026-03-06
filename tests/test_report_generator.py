from pathlib import Path

from openpyxl import load_workbook

from src.processor import compute_kpis
from src.report_generator import generate_excel_report


def test_generate_excel_report_creates_file(tmp_path: Path, sample_df):
    kpis = compute_kpis(sample_df)
    out = tmp_path / "out.xlsx"

    generate_excel_report(
        report_path=out,
        df_clean=sample_df,
        kpis=kpis,
        currency="USD",
        sources=["sample.csv"],
        chart_files=[],
    )

    assert out.exists()
    assert out.stat().st_size > 0


def test_generate_excel_report_formats_currency_cells(tmp_path: Path, sample_df):
    kpis = compute_kpis(sample_df)
    out = tmp_path / "out.xlsx"

    generate_excel_report(
        report_path=out,
        df_clean=sample_df,
        kpis=kpis,
        currency="USD",
        sources=["sample.csv"],
        chart_files=[],
    )

    wb = load_workbook(out)
    summary = wb["Summary"]
    money_fmt = summary["B5"].number_format  # Total Revenue

    assert "$" in money_fmt
    assert "#,##0.00" in money_fmt

    cleaned = wb["Cleaned Data"]
    cleaned_headers = {
        cleaned.cell(row=1, column=idx).value: idx
        for idx in range(1, cleaned.max_column + 1)
    }
    unit_price_col = cleaned_headers["unit_price"]
    assert cleaned.cell(row=2, column=unit_price_col).number_format == money_fmt

    top = wb["Top Products"]
    top_headers = {top.cell(row=1, column=idx).value: idx for idx in range(1, top.max_column + 1)}
    revenue_col = top_headers["revenue"]
    assert top.cell(row=2, column=revenue_col).number_format == money_fmt
