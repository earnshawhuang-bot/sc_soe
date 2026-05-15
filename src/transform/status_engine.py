"""
Status Engine - Assign status labels, calculate gap, determine risk tier.
"""
import pandas as pd
import numpy as np


def assign_status_and_gap(master: pd.DataFrame) -> pd.DataFrame:
    """
    For each SO in master:
    1. Compute quantity waterfall (no_plan_mt)
    2. Assign primary status label
    3. Calculate gap where applicable
    4. Assign risk tier

    Args:
        master: SO master DataFrame from join_engine

    Returns:
        master with added columns: no_plan_mt, status, available_date, gap_days, risk_tier
    """
    df = master.copy()

    # Preserve raw source quantities, then build a mutually exclusive
    # waterfall for management reporting and risk classification.
    if "raw_sc_prior_delivery_mt" not in df.columns:
        df["raw_sc_prior_delivery_mt"] = 0
    df["raw_sc_prior_delivery_mt"] = df["raw_sc_prior_delivery_mt"].fillna(0)
    df["raw_shipped_mt"] = df["shipped_mt"]
    df["raw_fg_mt"] = df["fg_mt"]
    df["raw_wip_mt"] = df["wip_mt"]
    df["raw_unsched_mt"] = df["unsched_mt"]

    remaining = df["sc_vol_mt"].copy()
    df["allocated_sc_prior_shipped_mt"] = np.minimum(df["raw_sc_prior_delivery_mt"], remaining)
    remaining = (remaining - df["allocated_sc_prior_shipped_mt"]).clip(lower=0)

    df["allocated_gi_shipped_mt"] = np.minimum(df["raw_shipped_mt"], remaining)
    remaining = (remaining - df["allocated_gi_shipped_mt"]).clip(lower=0)

    df["allocated_shipped_mt"] = df["allocated_sc_prior_shipped_mt"] + df["allocated_gi_shipped_mt"]

    df["allocated_fg_mt"] = np.minimum(df["raw_fg_mt"], remaining)
    remaining = (remaining - df["allocated_fg_mt"]).clip(lower=0)

    df["allocated_wip_mt"] = np.minimum(df["raw_wip_mt"], remaining)
    remaining = (remaining - df["allocated_wip_mt"]).clip(lower=0)

    df["allocated_unsched_mt"] = np.minimum(df["raw_unsched_mt"], remaining)
    remaining = (remaining - df["allocated_unsched_mt"]).clip(lower=0)

    df["allocated_no_plan_mt"] = remaining

    # Backward-compatible names now refer to mutually exclusive quantities.
    df["shipped_mt"] = df["allocated_shipped_mt"]
    df["fg_mt"] = df["allocated_fg_mt"]
    df["wip_mt"] = df["allocated_wip_mt"]
    df["unsched_mt"] = df["allocated_unsched_mt"]
    df["no_plan_mt"] = df["allocated_no_plan_mt"]

    # --- Status label (based on primary unfulfilled portion) ---
    df["status"] = _determine_status(df)

    # --- Gap calculation ---
    # planned_end_date is the raw PP Lami finish date. available_date keeps
    # the current business buffer explicit before comparing with loading date.
    df["available_date"] = df["planned_end_date"] + pd.Timedelta(days=1)

    # Only for SOs with both available_date AND loading_date. Loading Plan is
    # currently date-grain, so compare by calendar date unless LP carries time.
    has_both = df["planned_end_date"].notna() & df["loading_date"].notna()
    df["gap_days"] = np.nan

    if has_both.any():
        df.loc[has_both, "gap_days"] = _gap_days(
            df.loc[has_both, "loading_date"],
            df.loc[has_both, "available_date"],
        )

    # --- Risk tier ---
    df["risk_tier"] = _determine_risk(df)

    return df


def _determine_status(df: pd.DataFrame) -> pd.Series:
    """Assign primary status based on quantity distribution."""
    status = pd.Series("No Plan", index=df.index)

    # Primary status follows the most severe remaining allocated segment, so
    # partial shipment cannot hide unscheduled/no-plan exposure.
    status[df["allocated_shipped_mt"] >= df["sc_vol_mt"]] = "Shipped"
    status[df["allocated_fg_mt"] > 0] = "In Stock"
    status[df["allocated_wip_mt"] > 0] = "In Production"
    status[df["allocated_unsched_mt"] > 0] = "Planned (Unscheduled)"
    status[df["allocated_no_plan_mt"] > 0] = "No Plan"

    partial_only = (
        (df["allocated_shipped_mt"] > 0)
        & (df["allocated_shipped_mt"] < df["sc_vol_mt"])
        & (df["allocated_fg_mt"] == 0)
        & (df["allocated_wip_mt"] == 0)
        & (df["allocated_unsched_mt"] == 0)
        & (df["allocated_no_plan_mt"] == 0)
    )
    status[partial_only] = "Partially Shipped"

    return status


def _determine_risk(df: pd.DataFrame) -> pd.Series:
    """Assign risk tier based on status and gap."""
    risk = pd.Series("", index=df.index)

    # Shipped / Partially Shipped → Green/Yellow
    risk[df["status"] == "Shipped"] = "Green"
    risk[df["status"] == "Partially Shipped"] = "Yellow"

    # In Stock → Yellow (ready but not shipped yet)
    risk[df["status"] == "In Stock"] = "Yellow"

    # In Production → depends on gap
    in_prod = df["status"] == "In Production"
    has_gap = in_prod & df["gap_days"].notna()
    no_loading = in_prod & df["loading_date"].isna()

    risk[has_gap & (df["gap_days"] > 2)] = "Green"
    risk[has_gap & (df["gap_days"] >= 0) & (df["gap_days"] <= 2)] = "Yellow"
    risk[has_gap & (df["gap_days"] < 0)] = "Red"
    risk[no_loading] = "Orange"  # producing but no shipping arrangement

    # Planned (Unscheduled) → Red
    risk[df["status"] == "Planned (Unscheduled)"] = "Red"

    # No Plan → Critical Red
    risk[df["status"] == "No Plan"] = "Critical"

    return risk


def _gap_days(loading_date: pd.Series, available_date: pd.Series) -> pd.Series:
    """Compare by date when LP is date-only; preserve timestamp comparison when LP has time."""
    loading = pd.to_datetime(loading_date, errors="coerce")
    available = pd.to_datetime(available_date, errors="coerce")
    loading_has_time = (
        loading.dt.hour.fillna(0).ne(0)
        | loading.dt.minute.fillna(0).ne(0)
        | loading.dt.second.fillna(0).ne(0)
        | loading.dt.microsecond.fillna(0).ne(0)
    )
    date_grain_days = (loading.dt.normalize() - available.dt.normalize()).dt.days
    timestamp_days = (loading - available).dt.days
    return timestamp_days.where(loading_has_time, date_grain_days)
