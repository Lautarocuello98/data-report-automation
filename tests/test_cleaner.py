import pandas as pd

from src.cleaner import clean_sales_df


def test_clean_sales_df_basic(sample_df):
    config = {"cleaning": {"drop_duplicates": True, "fill_missing_numeric_with_zero": True}}
    cleaned, summary = clean_sales_df(sample_df, config=config)

    assert len(cleaned) == 2
    assert "bad_dates_coerced_to_na" in summary
    assert cleaned["quantity"].sum() == 3


def test_clean_sales_df_drops_null_and_placeholder_identity_rows():
    df = pd.DataFrame(
        {
            "date": ["2026-01-01", "2026-01-02", "2026-01-03"],
            "sku": [None, "   ", "SKU3"],
            "product": [None, "NaN", ""],
            "quantity": [1, 1, 1],
            "unit_price": [10, 10, 10],
            "unit_cost": [5, 5, 5],
        }
    )

    cleaned, summary = clean_sales_df(df, config={"cleaning": {"strip_strings": True}})

    assert len(cleaned) == 1
    assert cleaned.iloc[0]["sku"] == "SKU3"
    assert summary["dropped_empty_identity_rows"] == 2
