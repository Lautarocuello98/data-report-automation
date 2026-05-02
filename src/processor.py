from __future__ import annotations

import pandas as pd


def compute_kpis(df: pd.DataFrame) -> dict:
    """
    Compute core KPIs from cleaned sales data.

    Adds computed columns:
    - revenue, cost, profit, margin_pct
    """
    out = df.copy()

    out["revenue"] = out["quantity"] * out["unit_price"]
    out["cost"] = out["quantity"] * out["unit_cost"]
    out["profit"] = out["revenue"] - out["cost"]
    out["margin_pct"] = (
        (out["profit"] / out["revenue"].where(out["revenue"] != 0))
        .mul(100)
        .fillna(0.0)
    )

    total_orders = int(len(out))
    total_units = float(out["quantity"].sum())
    total_revenue = float(out["revenue"].sum())
    total_cost = float(out["cost"].sum())
    total_profit = float(out["profit"].sum())
    avg_order_value = float(total_revenue / total_orders) if total_orders else 0.0
    margin_pct = float((total_profit / total_revenue) * 100) if total_revenue else 0.0

    product_performance = (
        out.groupby("product", dropna=False)
        .agg(
            order_count=("product", "size"),
            total_units=("quantity", "sum"),
            revenue=("revenue", "sum"),
            cost=("cost", "sum"),
            profit=("profit", "sum"),
        )
        .reset_index()
    )
    if not product_performance.empty:
        product_performance["margin_pct"] = (
            product_performance["profit"]
            .div(product_performance["revenue"].where(product_performance["revenue"] != 0))
            .mul(100)
            .fillna(0.0)
        )
        product_performance = product_performance.sort_values(
            by=["revenue", "profit", "total_units"],
            ascending=[False, False, False],
        ).reset_index(drop=True)
    else:
        product_performance = pd.DataFrame(
            columns=[
                "product",
                "order_count",
                "total_units",
                "revenue",
                "cost",
                "profit",
                "margin_pct",
            ]
        )

    top_products = product_performance.loc[:, ["product", "revenue"]].head(10).copy()

    daily_performance = pd.DataFrame(
        columns=["date", "orders", "units", "revenue", "cost", "profit", "margin_pct"]
    )
    date_start = None
    date_end = None
    best_day = None

    if "date" in out.columns:
        dated = out.copy()
        dated["date"] = pd.to_datetime(dated["date"], errors="coerce")
        dated = dated.dropna(subset=["date"])
        if not dated.empty:
            date_start = dated["date"].min().date().isoformat()
            date_end = dated["date"].max().date().isoformat()
            daily_performance = (
                dated.groupby(dated["date"].dt.date)
                .agg(
                    orders=("date", "size"),
                    units=("quantity", "sum"),
                    revenue=("revenue", "sum"),
                    cost=("cost", "sum"),
                    profit=("profit", "sum"),
                )
                .reset_index()
                .rename(columns={"date": "date"})
            )
            daily_performance["date"] = pd.to_datetime(daily_performance["date"])
            daily_performance["margin_pct"] = (
                daily_performance["profit"]
                .div(daily_performance["revenue"].where(daily_performance["revenue"] != 0))
                .mul(100)
                .fillna(0.0)
            )
            daily_performance = daily_performance.sort_values("date").reset_index(drop=True)
            best_day_row = daily_performance.sort_values(
                by=["revenue", "profit"], ascending=[False, False]
            ).iloc[0]
            best_day = {
                "date": best_day_row["date"].date().isoformat(),
                "revenue": float(best_day_row["revenue"]),
                "profit": float(best_day_row["profit"]),
            }

    best_product = None
    if not product_performance.empty:
        best_product_row = product_performance.iloc[0]
        best_product = {
            "product": str(best_product_row["product"]),
            "revenue": float(best_product_row["revenue"]),
            "profit": float(best_product_row["profit"]),
            "units": float(best_product_row["total_units"]),
        }

    return {
        "total_orders": total_orders,
        "total_units": total_units,
        "total_revenue": total_revenue,
        "total_cost": total_cost,
        "total_profit": total_profit,
        "avg_order_value": avg_order_value,
        "margin_pct": margin_pct,
        "unique_products": int(out["product"].nunique(dropna=True)) if "product" in out.columns else 0,
        "unique_skus": int(out["sku"].nunique(dropna=True)) if "sku" in out.columns else 0,
        "date_start": date_start,
        "date_end": date_end,
        "best_product": best_product,
        "best_day": best_day,
        "top_products": top_products,
        "product_performance": product_performance,
        "daily_performance": daily_performance,
        "df_with_calculations": out,
    }
