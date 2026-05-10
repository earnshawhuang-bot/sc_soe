"""
FG Extractor - Finished Goods Inventory (03-FG)
Extracts current FG stock snapshot. KS only (IDN data absent).
"""
import pandas as pd


def extract_fg(file_path: str) -> pd.DataFrame:
    """
    Extract FG inventory data.

    Args:
        file_path: Path to FG stock Excel file

    Returns:
        DataFrame with columns: so, plant, fg_mt, receipt_date
    """
    df = pd.read_excel(file_path)
    df.columns = df.columns.str.strip()

    # Key columns
    plant_col = _find_col(df, "工厂")
    weight_col = _find_col(df, "重量")
    date_col = _find_col(df, "入库日期")
    so_col = _find_col(df, "合同编码")

    result = pd.DataFrame()
    result["so"] = df[so_col].apply(_to_so_str)
    result["plant"] = df[plant_col].apply(_map_plant_code)
    result["receipt_date"] = pd.to_datetime(df[date_col], errors="coerce")

    # Weight: KG ÷ 1000 → MT
    weight_raw = pd.to_numeric(df[weight_col], errors="coerce").fillna(0)
    result["fg_mt"] = weight_raw / 1000

    # Remove invalid: no SO, zero/negative weight
    result = result[result["so"].str.match(r"^\d+$", na=False)].copy()
    result = result[result["fg_mt"] > 0].copy()

    # Aggregate by SO
    agg = result.groupby("so").agg(
        fg_mt=("fg_mt", "sum"),
        plant=("plant", "first"),
        latest_receipt=("receipt_date", "max")
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


def _find_col(df: pd.DataFrame, name: str) -> str:
    """Find column by name."""
    for col in df.columns:
        if name in str(col).strip():
            return col
    raise ValueError(f"Column '{name}' not found. Available: {list(df.columns)}")


def _map_plant_code(value) -> str:
    """Map plant code to KS/IDN."""
    val = _to_so_str(value)
    if val in ("3000", "3001"):
        return "KS"
    elif val == "3301":
        return "IDN"
    return "Unknown"
