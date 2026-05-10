"""
Shipped Extractor - Actual Shipments (02-Shipped)
Extracts goods issue (GI) data with SO-level aggregation.
"""
import pandas as pd
import numpy as np


def extract_shipped(file_path: str) -> pd.DataFrame:
    """
    Extract shipped data.

    Args:
        file_path: Path to GI Excel file

    Returns:
        DataFrame with columns: so, plant, shipped_mt, ship_date
    """
    df = pd.read_excel(file_path)
    df.columns = df.columns.str.strip()

    # Key columns
    date_col = _find_col(df, "过账日期")
    plant_col = _find_col(df, "工厂")
    # Q column for sales order (more complete than P column)
    so_col = _find_so_column(df)
    weight_col = _find_col(df, "重量")

    result = pd.DataFrame()
    result["so"] = df[so_col].apply(_to_so_str)
    result["plant"] = df[plant_col].apply(_map_plant_code)
    result["ship_date"] = pd.to_datetime(df[date_col], errors="coerce")

    # Weight: × (-1) ÷ 1000 → MT (system shows negative for outbound)
    weight_raw = pd.to_numeric(df[weight_col], errors="coerce").fillna(0)
    result["shipped_mt"] = (weight_raw * -1) / 1000

    # Remove invalid: no SO, non-positive weight
    result = result[result["so"].str.match(r"^\d+$", na=False)].copy()
    result = result[result["shipped_mt"] > 0].copy()

    # Aggregate by SO (sum shipped_mt, take latest ship_date)
    agg = result.groupby("so").agg(
        shipped_mt=("shipped_mt", "sum"),
        plant=("plant", "first"),
        last_ship_date=("ship_date", "max")
    ).reset_index()

    return agg


def _find_so_column(df: pd.DataFrame) -> str:
    """Find the correct sales order column (Q column = 销售订单.1, contains 10-prefixed SOs)."""
    candidates = [col for col in df.columns if "销售订单" in str(col)]
    if not candidates:
        raise ValueError(f"Sales order column not found. Available: {list(df.columns)}")
    # Prefer the .1 suffix column (Q column) which has actual SC numbers (10-prefix)
    for c in candidates:
        sample = df[c].dropna().head(20)
        if any(str(int(v)).startswith("10") for v in sample if pd.notna(v)):
            return c
    # Fallback to second column if exists (pandas names duplicates as .1)
    if len(candidates) > 1:
        return candidates[1]
    return candidates[0]


def _find_col(df: pd.DataFrame, name: str) -> str:
    """Find column by name."""
    for col in df.columns:
        if name in str(col).strip():
            return col
    raise ValueError(f"Column '{name}' not found. Available: {list(df.columns)}")


def _to_so_str(value) -> str:
    """Convert SO value (often float) to clean string."""
    if pd.isna(value):
        return ""
    try:
        return str(int(float(value)))
    except (ValueError, TypeError):
        return str(value).strip()


def _map_plant_code(value) -> str:
    """Map plant code to KS/IDN."""
    val = str(int(float(value))) if pd.notna(value) else ""
    if val in ("3000", "3001"):
        return "KS"
    elif val == "3301":
        return "IDN"
    return "Unknown"
