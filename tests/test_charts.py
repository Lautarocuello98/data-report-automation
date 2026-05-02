from pathlib import Path

from src.charts import build_charts
from src.processor import compute_kpis


def test_build_charts_creates_expected_files(tmp_path: Path, sample_df):
    kpis = compute_kpis(sample_df)
    charts_dir = tmp_path / "charts"

    files = build_charts(sample_df, kpis, charts_dir)
    created = {p.name for p in files}

    assert {
        "revenue_by_day.png",
        "top_products.png",
        "profit_by_product.png",
        "revenue_mix.png",
    }.issubset(created)
    for p in files:
        assert p.exists()
        assert p.stat().st_size > 0
