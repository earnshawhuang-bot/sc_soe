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
        master with added columns: no_plan_mt, status, gap_days, risk_tier
    """
    df = master.copy()

    # --- Quantity waterfall ---
    # no_plan_mt = sc_vol - shipped - fg - wip - unsched (floored at 0)
    df["no_plan_mt"] = (
        df["sc_vol_mt"]
        - df["shipped_mt"]
        - df["fg_mt"]
        - df["wip_mt"]
        - df["unsched_mt"]
    ).clip(lower=0)

    # --- Status label (based on primary unfulfilled portion) ---
    df["status"] = _determine_status(df)

    # --- Gap calculation ---
    # Only for SOs with both planned_end_date AND loading_date
    has_both = df["planned_end_date"].notna() & df["loading_date"].notna()
    df["gap_days"] = np.nan

    if has_both.any():
        planned_plus_one = df.loc[has_both, "planned_end_date"] + pd.Timedelta(days=1)
        df.loc[has_both, "gap_days"] = (
            df.loc[has_both, "loading_date"] - planned_plus_one
        ).dt.days

    # --- Risk tier ---
    df["risk_tier"] = _determine_risk(df)

    return df


def _determine_status(df: pd.DataFrame) -> pd.Series:
    """Assign primary status based on quantity distribution."""
    status = pd.Series("No Plan", index=df.index)

    # Priority: Shipped > In Stock > In Production > Planned > No Plan
    # If shipped_mt covers full SC vol → Shipped
    # If partial → Partially Shipped (still check remaining)

    fully_shipped = df["shipped_mt"] >= df["sc_vol_mt"]
    status[fully_shipped] = "Shipped"

    partial_shipped = (df["shipped_mt"] > 0) & ~fully_shipped
    status[partial_shipped] = "Partially Shipped"

    # For non-shipped: check FG
    not_shipped = df["shipped_mt"] == 0
    has_fg = not_shipped & (df["fg_mt"] > 0)
    status[has_fg] = "In Stock"

    # Check WIP (scheduled production)
    no_fg = not_shipped & (df["fg_mt"] == 0)
    has_wip = no_fg & (df["wip_mt"] > 0)
    status[has_wip] = "In Production"

    # Check unscheduled (has work order but no date)
    no_wip = no_fg & (df["wip_mt"] == 0)
    has_unsched = no_wip & (df["unsched_mt"] > 0)
    status[has_unsched] = "Planned (Unscheduled)"

    # Rest = No Plan
    no_plan = no_wip & (df["unsched_mt"] == 0)
    status[no_plan] = "No Plan"

    return status


def _determine_risk(df: pd.DataFrame) -> pd.Series:
    """Assign risk tier based on status and gap."""
    risk = pd.Series("", index=df.index)

    # Shipped / Partially Shipped → Green (done or in progress)
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
