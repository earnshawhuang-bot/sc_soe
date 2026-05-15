"""
PP Extractor - Production Plan (04-PP)
Extracts scheduled and unscheduled production plans.
"""
import pandas as pd
from pathlib import Path


def extract_pp_scheduled(file_paths: list) -> pd.DataFrame:
    """
    Extract scheduled production plans from multiple PP files.

    Args:
        file_paths: List of paths to scheduled PP Excel files

    Returns:
        DataFrame with columns: work_order, so, plant, machine, weight_mt, planned_end_date.
        planned_end_date prefers PP Planned EndTime and falls back to PlannedFinishDate.
    """
    all_dfs = []

    for fp in file_paths:
        df = pd.read_excel(fp)
        df.columns = df.columns.str.strip()

        # Identify columns (unified structure across all 4 files)
        wo_col = _find_col(df, "工单", fallback="WorkOrder")
        so_col = _find_col(df, "SO")
        weight_col = _find_col(df, "TotalWeight", fallback="Weight")
        end_time_col = _find_col(df, "Planned EndTime")
        date_col = _find_col(df, "plannedfinishdate", fallback="PlannedFinishDate")
        machine_col = _find_col(df, "machine", fallback="Machine")

        # Determine plant from filename
        fname = Path(fp).name.lower()
        if "ind" in fname:
            plant = "IDN"
        else:
            plant = "KS"

        temp = pd.DataFrame()
        temp["work_order"] = df[wo_col].apply(_to_str)
        temp["so"] = df[so_col].apply(_to_str)
        temp["plant"] = plant
        temp["machine"] = df[machine_col].astype(str).str.strip()
        temp["weight_mt"] = pd.to_numeric(df[weight_col], errors="coerce").fillna(0)
        planned_finish = (
            pd.to_datetime(df[date_col], errors="coerce")
            if date_col
            else pd.Series(pd.NaT, index=df.index)
        )
        planned_end_time = pd.to_datetime(df[end_time_col], errors="coerce") if end_time_col else pd.Series(pd.NaT, index=df.index)
        temp["planned_end_date"] = planned_end_time.combine_first(planned_finish)

        all_dfs.append(temp)

    result = pd.concat(all_dfs, ignore_index=True)

    # Remove invalid SO
    result = result[result["so"].str.match(r"^\d+$", na=False)].copy()
    result = result[result["weight_mt"] > 0].copy()

    return result


def extract_pp_unscheduled(file_path: str) -> pd.DataFrame:
    """
    Extract unscheduled production plans (has work order, no planned date).

    Args:
        file_path: Path to Global_PP_wo schedule Excel file

    Returns:
        DataFrame with columns: work_order, so, plant, machine, weight_mt
    """
    df = pd.read_excel(file_path)
    df.columns = df.columns.str.strip()

    wo_col = _find_col(df, "工单", fallback="WorkOrder")
    so_col = _find_col(df, "SO")
    weight_col = _find_col(df, "TotalWeight", fallback="Weight")
    machine_col = _find_col(df, "machine", fallback="Machine")
    plant_col = _find_col(df, "Plant", fallback="plant")

    result = pd.DataFrame()
    result["work_order"] = df[wo_col].apply(_to_str)
    result["so"] = df[so_col].apply(_to_str)
    result["machine"] = df[machine_col].astype(str).str.strip()
    result["weight_mt"] = pd.to_numeric(df[weight_col], errors="coerce").fillna(0)

    # Plant mapping
    if plant_col:
        result["plant"] = df[plant_col].apply(_map_plant)
    else:
        result["plant"] = "Unknown"

    # Remove invalid
    result = result[result["so"].str.match(r"^\d+$", na=False)].copy()
    result = result[result["weight_mt"] > 0].copy()

    return result


def aggregate_pp_by_so(pp_sched: pd.DataFrame) -> pd.DataFrame:
    """
    Aggregate scheduled PP to SO level.
    - wip_mt = SUM(weight_mt) per SO
    - planned_end_date = MAX(planned_end_date) per SO, using Planned EndTime when available
    - machines = comma-joined unique machines
    - work_orders = comma-joined unique work orders

    Args:
        pp_sched: Raw scheduled PP DataFrame

    Returns:
        DataFrame with columns: so, wip_mt, planned_end_date, machines, work_orders, plant
    """
    agg = pp_sched.groupby("so").agg(
        wip_mt=("weight_mt", "sum"),
        planned_end_date=("planned_end_date", "max"),
        machines=("machine", lambda x: ", ".join(sorted(set(x)))),
        work_orders=("work_order", lambda x: ", ".join(sorted(set(x)))),
        plant=("plant", "first")
    ).reset_index()

    return agg


def aggregate_pp_unsched_by_so(pp_unsched: pd.DataFrame) -> pd.DataFrame:
    """
    Aggregate unscheduled PP to SO level.

    Returns:
        DataFrame with columns: so, planned_mt, machines, plant
    """
    agg = pp_unsched.groupby("so").agg(
        planned_mt=("weight_mt", "sum"),
        machines=("machine", lambda x: ", ".join(sorted(set(x)))),
        plant=("plant", "first")
    ).reset_index()

    return agg


def _find_col(df: pd.DataFrame, name: str, fallback: str = None) -> str:
    """Find column by name (case-insensitive)."""
    name_lower = name.lower()
    for col in df.columns:
        if col.strip().lower() == name_lower:
            return col
    # Partial match
    for col in df.columns:
        if name_lower in col.strip().lower():
            return col
    if fallback:
        return _find_col(df, fallback)
    # Return None instead of raising - some files may not have all columns
    return None


def _to_str(value) -> str:
    """Convert numeric value (often float) to clean string."""
    if pd.isna(value):
        return ""
    try:
        return str(int(float(value)))
    except (ValueError, TypeError):
        return str(value).strip()


def _map_plant(value) -> str:
    val = str(value).strip().lower()
    if "ks" in val or "kunshan" in val:
        return "KS"
    elif "idn" in val or "ind" in val:
        return "IDN"
    return "Unknown"
