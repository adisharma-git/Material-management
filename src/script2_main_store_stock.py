"""
Script 2: Main Store Stock Lookup Generator
Hospital Pharmacy Inventory Management
"""

import pandas as pd
import os


def create_main_store_stock_lookup(input_file: str, output_file: str, log_fn=print) -> pd.DataFrame:
    """
    Generate Main Store Stock Lookup from batch-level stock data.

    Args:
        input_file:  Path to Stock_Report_Material_without_item_category.xlsx
        output_file: Path for Material_Main_Store_Stock_Lookup.xlsx
        log_fn:      Callable for progress messages

    Returns:
        pd.DataFrame with columns ['Item\\nCode', 'Sum of Qty.']
    """

    MAIN_STORE = "Main Medical Store (MMS)-SEPL"
    ITEM_COL   = "Item\nCode"
    QTY_COL    = "Qty."

    # ── Step 1: Load ─────────────────────────────────────────────────────────────
    if not os.path.exists(input_file):
        raise FileNotFoundError(f"Input file not found: {input_file}")

    log_fn("Loading input file…")
    xl = pd.ExcelFile(input_file)
    
    raw_sheet = next((s for s in xl.sheet_names if "raw" in s.lower()), None)
    if raw_sheet is None:
        raise ValueError(f"No 'RAW DATA' sheet found. Available: {xl.sheet_names}")

    df = pd.read_excel(input_file, sheet_name=raw_sheet)

    log_fn(f"  Loaded {len(df):,} rows from 'RAW DATA'")

    # Check required columns
    required = [ITEM_COL, QTY_COL, "Store Name"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        log_fn(f"  Available columns: {df.columns.tolist()}")
        raise ValueError(f"Missing columns: {missing}")

    if len(df) == 0:
        raise ValueError("Input sheet is empty")

    # ── Step 2: Quality checks ───────────────────────────────────────────────────
    log_fn("\nVerifying data quality…")

    # Store name check
    main_rows = (df["Store Name"] == MAIN_STORE).sum()
    log_fn(f"  Main Medical Store rows: {main_rows:,} / {len(df):,}")
    if main_rows != len(df):
        log_fn(f"  WARNING: {len(df) - main_rows:,} rows from other stores – filtering…")
        df = df[df["Store Name"] == MAIN_STORE]

    # Missing data check
    missing_items = df[ITEM_COL].isna().sum()
    missing_qty   = df[QTY_COL].isna().sum()
    log_fn(f"  Missing Item Codes : {missing_items}")
    log_fn(f"  Missing Quantities : {missing_qty}")
    if missing_items > 0 or missing_qty > 0:
        log_fn("  Removing rows with missing data…")
        df = df[df[ITEM_COL].notna() & df[QTY_COL].notna()]

    # Convert qty to numeric
    df[QTY_COL] = pd.to_numeric(df[QTY_COL], errors="coerce")
    df = df[df[QTY_COL].notna()]

    # ── Step 3: Aggregate ────────────────────────────────────────────────────────
    log_fn("\nAggregating by Item Code…")
    input_total = df[QTY_COL].sum()
    result = df.groupby(ITEM_COL)[QTY_COL].sum().reset_index()
    result.columns = [ITEM_COL, "Sum of Qty."]
    result = result.sort_values(ITEM_COL).reset_index(drop=True)
    log_fn(f"  {len(df):,} batch records → {len(result):,} unique items")

    # ── Step 4: Validate ─────────────────────────────────────────────────────────
    assert result[ITEM_COL].is_unique, "Duplicate item codes found"
    assert (result["Sum of Qty."] >= 0).all(), "Negative quantities found"
    assert result["Sum of Qty."].notna().all(), "Null quantities found"
    output_total = result["Sum of Qty."].sum()
    assert abs(input_total - output_total) < 0.01, \
        f"Quantity mismatch: input={input_total}, output={output_total}"
    log_fn(f"  ✓ Quantity conservation: {output_total:,.2f}")

    # ── Step 5: Export ───────────────────────────────────────────────────────────
    log_fn(f"\nExporting to {output_file}…")
    result.to_excel(output_file, sheet_name="OUTPUT", index=False)

    log_fn("\n" + "=" * 55)
    log_fn("MAIN STORE STOCK LOOKUP – COMPLETE")
    log_fn("=" * 55)
    log_fn(f"  Input batches      : {len(df):,}")
    log_fn(f"  Unique items       : {len(result):,}")
    log_fn(f"  Total stock qty    : {result['Sum of Qty.'].sum():,.2f}")
    log_fn(f"  Output file        : {output_file}")
    log_fn("=" * 55)

    return result


if __name__ == "__main__":
    create_main_store_stock_lookup(
        "Stock_Report_Material_without_item_category.xlsx",
        "Material_Main_Store_Stock_Lookup.xlsx",
    )
