"""
============================================================
 EXCEL / CSV REPORT AUTOMATOR
============================================================
 Turns messy spreadsheet data into a clean, professional
 report - automatically.

 What it does:
   1. Loads any CSV or Excel file
   2. Cleans the data (whitespace, duplicates, dates, blanks)
   3. Computes summary statistics
   4. Generates a formatted Excel report with charts
   5. Prints a console summary + optional HTML report

 Usage:
   python report_automator.py --input data.csv
   python report_automator.py --input sales.xlsx --category Region --value Revenue
   python report_automator.py --input data.csv --html
============================================================
"""

import argparse
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

# ----------------------------------------------------------------------
# 1. LOADING
# ----------------------------------------------------------------------
def load_data(path: Path) -> pd.DataFrame:
    """Load a CSV or Excel file into a DataFrame."""
    suffix = path.suffix.lower()
    try:
        if suffix == ".csv":
            return pd.read_csv(path)
        if suffix in (".xlsx", ".xls", ".xlsm"):
            return pd.read_excel(path)
    except Exception as e:
        print(f"ERROR: could not read '{path}': {e}")
        sys.exit(1)
    print(f"ERROR: unsupported file type '{suffix}'. Use .csv or .xlsx.")
    sys.exit(1)


# ----------------------------------------------------------------------
# 2. CLEANING
# ----------------------------------------------------------------------
def clean_data(df: pd.DataFrame) -> pd.DataFrame:
    """Standard cleanup: strip text, drop exact duplicates, fix dates, fill blanks."""
    report = {}
    before = len(df)

    # Strip whitespace from text columns
    for col in df.columns:
        if pd.api.types.is_string_dtype(df[col]):
            df[col] = df[col].astype(str).str.strip().replace({"nan": pd.NA, "": pd.NA})

    # Drop exact duplicate rows
    df = df.drop_duplicates()
    report["duplicates_removed"] = before - len(df)

    # Try to convert likely-date columns
    date_cols_found = 0
    for col in df.columns:
        if pd.api.types.is_string_dtype(df[col]) and any(
            kw in col.lower() for kw in ("date", "time", "day")
        ):
            converted = pd.to_datetime(df[col], errors="coerce", format="mixed")
            if converted.notna().sum() > 0:
                df[col] = converted
                date_cols_found += 1
    report["date_columns_parsed"] = date_cols_found

    # Count blank cells remaining
    report["blank_cells_remaining"] = int(df.isna().sum().sum())

    return df, report


# ----------------------------------------------------------------------
# 3. SUMMARY STATISTICS
# ----------------------------------------------------------------------
def summarize(df: pd.DataFrame) -> pd.DataFrame:
    """Build a statistics table for every numeric column."""
    numeric = df.select_dtypes(include="number")
    if numeric.empty:
        return pd.DataFrame()

    stats = pd.DataFrame(
        {
            "Column": numeric.columns,
            "Count": [int(numeric[c].count()) for c in numeric.columns],
            "Sum": [numeric[c].sum() for c in numeric.columns],
            "Average": [round(numeric[c].mean(), 2) for c in numeric.columns],
            "Min": [numeric[c].min() for c in numeric.columns],
            "Max": [numeric[c].max() for c in numeric.columns],
        }
    )
    return stats


def category_breakdown(df: pd.DataFrame, category_col: str, value_col: str | None):
    """Aggregate a value column by a category column (e.g. Revenue by Region)."""
    if category_col not in df.columns:
        print(f"NOTE: column '{category_col}' not found - skipping breakdown.")
        return None
    if value_col and value_col in df.columns:
        return (
            df.groupby(category_col, dropna=False)[value_col]
            .agg(["count", "sum", "mean"])
            .round(2)
            .reset_index()
        )
    return df[category_col].value_counts(dropna=False).reset_index()


# ----------------------------------------------------------------------
# 4. OUTPUT
# ----------------------------------------------------------------------
def write_excel_report(
    output_path: Path,
    cleaned: pd.DataFrame,
    stats: pd.DataFrame,
    breakdown,
    cleanup_report: dict,
) -> None:
    """Write a multi-sheet Excel report with a chart on the Summary sheet."""
    try:
        from openpyxl.chart import BarChart, Reference
        from openpyxl.styles import Font, PatternFill
    except ImportError:
        print("ERROR: openpyxl is required. Run: pip install openpyxl")
        sys.exit(1)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        cleaned.to_excel(writer, sheet_name="Cleaned Data", index=False)
        if not stats.empty:
            stats.to_excel(writer, sheet_name="Statistics", index=False)
        if breakdown is not None:
            breakdown.to_excel(writer, sheet_name="Breakdown", index=False)

        # --- Summary sheet -------------------------------------------------
        summary_rows = {
            "Report generated": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "Total rows (after cleaning)": len(cleaned),
            "Total columns": len(cleaned.columns),
            "Duplicate rows removed": cleanup_report["duplicates_removed"],
            "Date columns parsed": cleanup_report["date_columns_parsed"],
            "Blank cells remaining": cleanup_report["blank_cells_remaining"],
        }
        pd.DataFrame(list(summary_rows.items()), columns=["Metric", "Value"]).to_excel(
            writer, sheet_name="Summary", index=False
        )

        # --- Chart on the Summary sheet ------------------------------------
        if breakdown is not None and len(breakdown) > 0:
            numeric_cols = breakdown.select_dtypes(include="number").columns
            if len(numeric_cols) >= 1:
                sheet = writer.sheets["Summary"]
                chart = BarChart()
                chart.title = "Breakdown by Category"
                chart.style = 10
                data = Reference(
                    writer.sheets["Breakdown"],
                    min_col=2, max_col=2,
                    min_row=1, max_row=len(breakdown) + 1,
                )
                cats = Reference(
                    writer.sheets["Breakdown"],
                    min_col=1, min_row=2, max_row=len(breakdown) + 1,
                )
                chart.add_data(data, titles_from_data=True)
                chart.set_categories(cats)
                sheet.add_chart(chart, "E2")

    print(f"Excel report saved: {output_path}")


def write_html_report(path: Path, stats: pd.DataFrame, breakdown) -> None:
    """Write a self-contained HTML report (viewable in any browser)."""
    css = """
    body { font-family: 'Segoe UI', Arial, sans-serif; margin: 40px; color: #222; }
    h1 { color: #1a73e8; } h2 { margin-top: 30px; color: #333; }
    table { border-collapse: collapse; margin-top: 10px; width: 80%; }
    th, td { border: 1px solid #ccc; padding: 8px 14px; text-align: left; }
    th { background: #1a73e8; color: white; }
    tr:nth-child(even) { background: #f5f8fd; }
    """
    parts = [
        "<!DOCTYPE html><html><head><meta charset='utf-8'>",
        f"<title>Report - {path.stem}</title><style>{css}</style></head><body>",
        f"<h1>📊 Data Report: {path.stem}</h1>",
        f"<p>Generated {datetime.now().strftime('%Y-%m-%d %H:%M')}</p>",
    ]
    if not stats.empty:
        parts.append("<h2>Statistics</h2>" + stats.to_html(index=False))
    if breakdown is not None:
        parts.append("<h2>Category Breakdown</h2>" + breakdown.to_html(index=False))
    parts.append("</body></html>")
    path.write_text("\n".join(parts), encoding="utf-8")
    print(f"HTML report saved: {path}")


# ----------------------------------------------------------------------
# 5. CONSOLE REPORT
# ----------------------------------------------------------------------
def print_console_report(df: pd.DataFrame, stats: pd.DataFrame, cleanup: dict) -> None:
    print("\n" + "=" * 60)
    print(" REPORT SUMMARY")
    print("=" * 60)
    print(f" Rows (after cleaning):    {len(df)}")
    print(f" Columns:                  {len(df.columns)}")
    print(f" Duplicates removed:       {cleanup['duplicates_removed']}")
    print(f" Date columns parsed:      {cleanup['date_columns_parsed']}")
    print(f" Blank cells remaining:    {cleanup['blank_cells_remaining']}")
    if not stats.empty:
        print("\n--- Numeric Statistics ---")
        print(stats.to_string(index=False))
    print("=" * 60)


# ----------------------------------------------------------------------
# MAIN
# ----------------------------------------------------------------------
def main() -> None:
    parser = argparse.ArgumentParser(
        description="Turn messy CSV/Excel data into a clean report.",
        epilog="Example: python report_automator.py --input sales.csv --category Region --value Revenue",
    )
    parser.add_argument("--input", "-i", required=True, help="Input CSV or Excel file")
    parser.add_argument("--output", "-o", help="Output Excel report path (default: <input>_report.xlsx)")
    parser.add_argument("--category", "-c", help="Column to group by (e.g. Region)")
    parser.add_argument("--value", "-v", help="Numeric column to aggregate (e.g. Revenue)")
    parser.add_argument("--html", action="store_true", help="Also generate an HTML report")
    args = parser.parse_args()

    input_path = Path(args.input)
    if not input_path.exists():
        print(f"ERROR: file '{input_path}' not found.")
        sys.exit(1)

    df = load_data(input_path)
    print(f"Loaded {len(df)} rows from {input_path.name}")

    df, cleanup = clean_data(df)
    stats = summarize(df)
    breakdown = category_breakdown(df, args.category, args.value)

    print_console_report(df, stats, cleanup)

    output_path = Path(args.output) if args.output else input_path.with_name(input_path.stem + "_report.xlsx")
    write_excel_report(output_path, df, stats, breakdown, cleanup)

    if args.html:
        write_html_report(output_path.with_suffix(".html"), stats, breakdown)


if __name__ == "__main__":
    main()
