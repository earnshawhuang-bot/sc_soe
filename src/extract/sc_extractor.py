"""
SC Extractor - Sales Order Baseline (01-SC)
Extracts and classifies orders from the SC master table.
"""
import pandas as pd
from datetime import datetime


def extract_sc(file_path: str, month: str) -> pd.DataFrame:
    """
    Extract SC baseline orders.

    Args:
        file_path: Path to Order tracking Excel file
        month: Current month string e.g. "2026-05"

    Returns:
        DataFrame with columns: so, plant, cluster, sc_vol_mt, order_type
    """
    # Read the specific sheet
    df = pd.read_excel(file_path, sheet_name="Order Status - Main")

    # Standardize column names (strip whitespace)
    df.columns = df.columns.str.strip()

    # Filter: Item = "main"
    item_col = _find_column(df, "Item")
    df = df[df[item_col].astype(str).str.strip().str.lower() == "main"].copy()

    # Extract key columns
    so_col = _find_column(df, "SC NO")
    plant_col = _find_column(df, "Supply From")
    cluster_col = _find_column(df, "Cluster")
    vol_col = _find_column(df, "SC Vol.-MT", fallback="SCVOL-MT")
    status_col = _find_column(df, "Order Status")
    carryover_col = _find_column(df, "Carryover", col_position="Q")
    carryover_fresh_col = _find_column(df, "carryover/fresh", fallback="Carry Over / Fresh")
    release_col = _find_column(df, "RELEASE DATE")
    loading_col = _find_column(df, "Loading date")

    # Build result DataFrame
    result = pd.DataFrame()
    result["so"] = df[so_col].apply(_to_so_str)
    result["plant"] = df[plant_col].apply(_map_plant)
    result["cluster"] = df[cluster_col].astype(str).str.strip()
    result["sc_vol_mt"] = pd.to_numeric(df[vol_col], errors="coerce").fillna(0)
    result["loading_date_sc"] = pd.to_datetime(df[loading_col], errors="coerce")

    # Classification
    order_status = df[status_col].astype(str).str.strip()
    q_col_values = df[carryover_col]  # Q column
    ao_col_values = df[carryover_fresh_col].astype(str).str.strip()
    release_dates = pd.to_datetime(df[release_col], errors="coerce")

    # Q column empty check (NaN, None, empty string)
    q_is_empty = q_col_values.isna() | (q_col_values.astype(str).str.strip() == "") | (q_col_values.astype(str).str.strip().str.lower() == "nan")

    # Current month check
    target_year = int(month.split("-")[0])
    target_month = int(month.split("-")[1])
    is_current_month = (release_dates.dt.year == target_year) & (release_dates.dt.month == target_month)

    # Four order types (independently verified, MECE)
    result["order_type"] = "Unknown"

    # Type 1: Carry Over Stock
    mask_cos = (order_status.str.lower() == "carry over") & q_is_empty
    result.loc[mask_cos, "order_type"] = "Carry Over Stock"

    # Type 2: Carry Over Unproduced
    mask_coup = (ao_col_values.str.lower() == "last month") & q_is_empty
    result.loc[mask_coup, "order_type"] = "Carry Over Unproduced"

    # Type 3: Fresh Order This Month
    mask_fresh = is_current_month & q_is_empty
    result.loc[mask_fresh, "order_type"] = "Fresh Order This Month"

    # Type 4: Fresh Order Next Month (excluded from baseline)
    mask_next = is_current_month & ~q_is_empty
    result.loc[mask_next, "order_type"] = "Fresh Order Next Month"

    # Baseline = Type 1 + 2 + 3 (exclude Type 4 and Unknown)
    result["in_baseline"] = result["order_type"].isin([
        "Carry Over Stock", "Carry Over Unproduced", "Fresh Order This Month"
    ])

    # Remove rows with invalid SO
    result = result[result["so"].str.match(r"^\d+$", na=False)].copy()

    # Aggregate by SO: sum volumes, keep first for categorical fields
    agg = result.groupby("so").agg(
        plant=("plant", "first"),
        cluster=("cluster", "first"),
        sc_vol_mt=("sc_vol_mt", "sum"),
        loading_date_sc=("loading_date_sc", "first"),
        order_type=("order_type", "first"),
        in_baseline=("in_baseline", "max"),  # True if any row is in baseline
    ).reset_index()

    return agg


def _to_so_str(value) -> str:
    """Convert SO value (often float) to clean string."""
    if pd.isna(value):
        return ""
    try:
        return str(int(float(value)))
    except (ValueError, TypeError):
        return str(value).strip()


def _map_plant(value) -> str:
    """Map Supply From to plant code."""
    val = str(value).strip().lower()
    if "kunshan" in val or "ks" in val:
        return "KS"
    elif "indonesia" in val or "idn" in val:
        return "IDN"
    return "Unknown"


def _find_column(df: pd.DataFrame, name: str, fallback: str = None, col_position: str = None) -> str:
    """Find column by name (case-insensitive, partial match)."""
    name_lower = name.lower()
    for col in df.columns:
        if col.strip().lower() == name_lower:
            return col
    # Partial match
    for col in df.columns:
        if name_lower in col.strip().lower():
            return col
    # Try fallback
    if fallback:
        return _find_column(df, fallback)
    raise ValueError(f"Column '{name}' not found. Available: {list(df.columns)}")
