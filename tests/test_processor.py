import pandas as pd

from src.processor import compute_kpis


def test_compute_kpis(sample_df):
    kpis = compute_kpis(sample_df)

    assert kpis["total_orders"] == 2
    assert kpis["total_units"] == 3
    assert kpis["total_revenue"] == 40.0  # 2*10 + 1*20
    assert kpis["total_cost"] == 24.0     # 2*6 + 1*12
    assert kpis["total_profit"] == 16.0
    assert kpis["margin_pct"] == 40.0
    assert kpis["unique_products"] == 2
    assert kpis["unique_skus"] == 2
    assert kpis["date_start"] == "2026-01-01"
    assert kpis["date_end"] == "2026-01-02"
    assert kpis["best_product"]["revenue"] == 20.0
    assert kpis["best_day"]["revenue"] == 20.0
    assert list(kpis["product_performance"].columns) == [
        "product",
        "order_count",
        "total_units",
        "revenue",
        "cost",
        "profit",
        "margin_pct",
    ]
    assert list(kpis["daily_performance"].columns) == [
        "date",
        "orders",
        "units",
        "revenue",
        "cost",
        "profit",
        "margin_pct",
    ]


def test_compute_kpis_margin_pct_handles_zero_revenue():
    df = pd.DataFrame(
        {
            "date": ["2026-01-01"],
            "sku": ["SKU-0"],
            "product": ["Free Sample"],
            "quantity": [2],
            "unit_price": [0],
            "unit_cost": [0],
        }
    )

    kpis = compute_kpis(df)
    df_out = kpis["df_with_calculations"]

    assert float(df_out.loc[0, "revenue"]) == 0.0
    assert float(df_out.loc[0, "margin_pct"]) == 0.0
