from __future__ import annotations

import argparse
import json
import logging
from pathlib import Path

from src.loader import load_input
from src.cleaner import clean_sales_df
from src.processor import compute_kpis
from src.report_generator import generate_excel_report
from src.charts import build_charts


def setup_logging(log_path: Path, verbose: bool) -> None:
    log_path.parent.mkdir(parents=True, exist_ok=True)

    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(
        level=level,
        format="%(asctime)s | %(levelname)s | %(message)s",
        handlers=[
            logging.FileHandler(log_path, encoding="utf-8"),
            logging.StreamHandler(),
        ],
    )


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="Data Report Automation: CSV/Excel -> cleaned data -> KPIs -> Excel report"
    )
    p.add_argument("--input", "-i", required=True, help="Input file or folder (CSV/XLSX)")
    p.add_argument("--output", "-o", default="reports", help="Output folder")
    p.add_argument("--config", "-c", default="config.json", help="Path to config.json")
    p.add_argument("--verbose", "-v", action="store_true", help="Verbose logging")
    return p.parse_args()


def main() -> int:
    args = parse_args()

    input_path = Path(args.input).expanduser().resolve()
    output_dir = Path(args.output).expanduser().resolve()
    config_path = Path(args.config).expanduser().resolve()

    if not config_path.exists():
        raise SystemExit(f"Config not found: {config_path}")

    with config_path.open("r", encoding="utf-8") as f:
        config = json.load(f)

    output_dir.mkdir(parents=True, exist_ok=True)
    log_path = output_dir / "processing.log"
    setup_logging(log_path, args.verbose)

    logging.info("Starting processing")
    logging.info("Input: %s", input_path)
    logging.info("Output: %s", output_dir)
    logging.info("Config: %s", config_path)

    df_raw, sources = load_input(input_path, config=config)
    logging.info("Loaded %d rows from %d source(s)", len(df_raw), len(sources))

    df_clean, cleaning_summary = clean_sales_df(df_raw, config=config)
    logging.info("Cleaning summary: %s", cleaning_summary)

    kpis = compute_kpis(df_clean)
    top_products = kpis.get("top_products")
    top_products_count = int(len(top_products)) if top_products is not None else 0
    kpis_for_log = {
        "total_orders": int(kpis.get("total_orders", 0)),
        "total_units": float(kpis.get("total_units", 0.0)),
        "total_revenue": float(kpis.get("total_revenue", 0.0)),
        "total_cost": float(kpis.get("total_cost", 0.0)),
        "total_profit": float(kpis.get("total_profit", 0.0)),
        "avg_order_value": float(kpis.get("avg_order_value", 0.0)),
        "top_products_count": top_products_count,
    }
    logging.info("KPIs computed: %s", kpis_for_log)

    charts_dir = output_dir / "charts"
    chart_files = []
    if config.get("report", {}).get("include_charts", True):
        chart_files = build_charts(df_clean, kpis, charts_dir)

    excel_name = config.get("report", {}).get("excel_filename", "sales_report.xlsx")
    report_path = output_dir / excel_name
    generate_excel_report(
        report_path=report_path,
        df_clean=df_clean,
        kpis=kpis,
        currency=config.get("report", {}).get("currency", "USD"),
        sources=sources,
        chart_files=chart_files,
    )

    logging.info("Report generated: %s", report_path)
    logging.info("Done.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
