from __future__ import annotations

from typing import Any

import pandas as pd


REQUIRED_COLUMNS = {"date", "sku", "product", "quantity", "unit_price", "unit_cost"}


def _ensure_required_columns(df: pd.DataFrame) -> None:
    missing = REQUIRED_COLUMNS - set(df.columns)
    if missing:
        raise ValueError(
            "Missing required columns. "
            f"Required: {sorted(REQUIRED_COLUMNS)} | Missing: {sorted(missing)}"
        )


def clean_sales_df(df: pd.DataFrame, config: dict[str, Any]) -> tuple[pd.DataFrame, dict[str, Any]]:
    """
    Clean and normalize a sales dataframe.

    Expected columns (after mapping):
    - date, sku, product, quantity, unit_price, unit_cost
    Optional:
    - source_file
    """
    cfg = config.get("cleaning", {})
    out = df.copy()

    _ensure_required_columns(out)

    summary: dict[str, Any] = {}

    # Strip strings
    if cfg.get("strip_strings", True):
        for c in ["sku", "product"]:
            out[c] = out[c].astype("string").str.strip()
        summary["strip_strings"] = True

    # Parse date
    out["date"] = pd.to_datetime(out["date"], errors="coerce")
    bad_dates = int(out["date"].isna().sum())
    summary["bad_dates_coerced_to_na"] = bad_dates

    # Numeric conversions
    for c in ["quantity", "unit_price", "unit_cost"]:
        out[c] = pd.to_numeric(out[c], errors="coerce")

    numeric_na = {c: int(out[c].isna().sum()) for c in ["quantity", "unit_price", "unit_cost"]}
    summary["numeric_na_after_parse"] = numeric_na

    if cfg.get("fill_missing_numeric_with_zero", True):
        out[["quantity", "unit_price", "unit_cost"]] = out[["quantity", "unit_price", "unit_cost"]].fillna(0)
        summary["filled_numeric_na_with_zero"] = True

    # Drop negative quantities if configured
    if cfg.get("drop_negative_quantities", True):
        before = len(out)
        out = out[out["quantity"] >= 0].copy()
        summary["dropped_negative_quantity_rows"] = before - len(out)

    # Drop duplicates
    if cfg.get("drop_duplicates", True):
        before = len(out)
        out = out.drop_duplicates()
        summary["dropped_duplicate_rows"] = before - len(out)

    # Basic sanity: keep only rows with non-empty SKU or product
    before = len(out)
    invalid_text = {"", "nan", "none", "null", "<na>"}
    sku_norm = out["sku"].astype("string").fillna("").str.strip().str.lower()
    product_norm = out["product"].astype("string").fillna("").str.strip().str.lower()
    keep_mask = (~sku_norm.isin(invalid_text)) | (~product_norm.isin(invalid_text))
    out = out[keep_mask].copy()
    summary["dropped_empty_identity_rows"] = before - len(out)

    # Order columns nicely
    preferred = ["date", "sku", "product", "quantity", "unit_price", "unit_cost", "source_file"]
    cols = [c for c in preferred if c in out.columns] + [c for c in out.columns if c not in preferred]
    out = out[cols]

    return out, summary
