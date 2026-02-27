"""
Script 1: Global Stock Lookup Generator
Hospital Pharmacy Inventory Management
"""

import pandas as pd
import os


def create_global_stock_lookup(input_file: str, output_file: str, log_fn=print) -> pd.DataFrame:
    """
    Generate Global Stock Lookup from raw stock report.

    Args:
        input_file:  Path to Material_Global_Stock_Report.CSV
        output_file: Path for Material_Global_Stock_Lookup.xlsx
        log_fn:      Callable for progress messages (default: print)

    Returns:
        pd.DataFrame with columns ['Item Code', 'Total Global Stock']
    """

    # ── Step 1: Load ────────────────────────────────────────────────────────────
    if not os.path.exists(input_file):
        raise FileNotFoundError(f"Input file not found: {input_file}")

    log_fn("Loading input file…")
    df = pd.read_csv(input_file, encoding="utf-8")
    log_fn(f"  Loaded {len(df):,} rows")

    required_cols = ["ItemCode", "Description", "Qty"]
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        raise ValueError(f"Input file missing columns: {missing}")

    if len(df) == 0:
        raise ValueError("Input file is empty")

    # ── Step 2: Clean ────────────────────────────────────────────────────────────
    log_fn("\nCleaning data…")

    # Filter 1 – remove blank Description
    before = len(df)
    df = df[df["Description"].notna() & (df["Description"].astype(str).str.strip() != "")]
    log_fn(f"  Removed {before - len(df):,} rows with blank Description")

    # Filter 2 – remove blank Qty
    before = len(df)
    df = df[df["Qty"].notna()]
    log_fn(f"  Removed {before - len(df):,} rows with blank Qty")

    # Filter 3 – convert Qty to numeric
    df["Qty"] = pd.to_numeric(df["Qty"], errors="coerce")
    before = len(df)
    df = df[df["Qty"].notna()]
    removed = before - len(df)
    if removed:
        log_fn(f"  Removed {removed:,} rows with non-numeric Qty")

    if len(df) == 0:
        raise ValueError("All rows were filtered out – check input data")

    # ── Step 3: Aggregate ────────────────────────────────────────────────────────
    log_fn("\nAggregating…")
    result = df.groupby("ItemCode")["Qty"].sum().reset_index()
    result.columns = ["Item Code", "Total Global Stock"]
    result = result.sort_values("Item Code").reset_index(drop=True)
    log_fn(f"  {len(df):,} rows → {len(result):,} unique items")

    # ── Step 4: Validate ─────────────────────────────────────────────────────────
    assert "Intransit Store" not in result["Item Code"].values, "Intransit Store should be removed"
    assert result["Item Code"].is_unique, "Duplicate item codes found"
    assert (result["Total Global Stock"] >= 0).all(), "Negative quantities found"
    assert result["Total Global Stock"].notna().all(), "Null quantities found"

    # ── Step 5: Export ───────────────────────────────────────────────────────────
    log_fn(f"\nExporting to {output_file}…")
    result.to_excel(output_file, sheet_name="output", index=False)

    log_fn("\n" + "=" * 55)
    log_fn("GLOBAL STOCK LOOKUP – COMPLETE")
    log_fn("=" * 55)
    log_fn(f"  Unique items       : {len(result):,}")
    log_fn(f"  Total stock qty    : {result['Total Global Stock'].sum():,.0f}")
    log_fn(f"  Output file        : {output_file}")
    log_fn("=" * 55)

    return result


if __name__ == "__main__":
    create_global_stock_lookup(
        "Material_Global_Stock_Report.CSV",
        "Material_Global_Stock_Lookup.xlsx",
    )
