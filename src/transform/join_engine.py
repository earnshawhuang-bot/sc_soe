"""
Join Engine - Merge all data sources on SO as primary key.
"""
import pandas as pd


def build_so_master(
    sc: pd.DataFrame,
    shipped: pd.DataFrame,
    fg: pd.DataFrame,
    pp_sched_agg: pd.DataFrame,
    pp_unsched_agg: pd.DataFrame,
    loading_plan: pd.DataFrame
) -> pd.DataFrame:
    """
    Build the SO master table by left-joining all sources onto SC baseline.

    Args:
        sc: SC baseline (must have 'so' column)
        shipped: Aggregated shipped data
        fg: Aggregated FG data
        pp_sched_agg: Aggregated scheduled PP (SO-level)
        pp_unsched_agg: Aggregated unscheduled PP (SO-level)
        loading_plan: Combined loading plan

    Returns:
        SO master DataFrame with all dimensions merged
    """
    # Start with baseline only
    master = sc[sc["in_baseline"]].copy()

    # LEFT JOIN Shipped
    master = master.merge(
        shipped[["so", "shipped_mt", "last_ship_date"]],
        on="so", how="left"
    )

    # LEFT JOIN FG
    master = master.merge(
        fg[["so", "fg_mt", "latest_receipt"]],
        on="so", how="left"
    )

    # LEFT JOIN PP Scheduled (SO-level aggregated)
    pp_sched_cols = ["so", "wip_mt", "planned_end_date", "machines"]
    if "work_orders" in pp_sched_agg.columns:
        pp_sched_cols.append("work_orders")
    master = master.merge(
        pp_sched_agg[pp_sched_cols],
        on="so", how="left"
    )

    # LEFT JOIN PP Unscheduled
    master = master.merge(
        pp_unsched_agg[["so", "planned_mt"]].rename(columns={"planned_mt": "unsched_mt"}),
        on="so", how="left"
    )

    # LEFT JOIN Loading Plan. Keep optional LP diagnostics when available while
    # preserving the legacy columns used by gap/risk calculations.
    lp_cols = [
        "so", "loading_date", "load_mt", "source", "lp_valid_mt",
        "lp_unconfirmed_mt", "lp_line_count", "lp_date_status", "lp_has_record",
    ]
    lp_cols = [col for col in lp_cols if col in loading_plan.columns]
    master = master.merge(
        loading_plan[lp_cols].rename(columns={"source": "lp_source"}),
        on="so", how="left"
    )

    # Fill NaN with 0 for quantity columns
    qty_cols = ["shipped_mt", "fg_mt", "wip_mt", "unsched_mt", "load_mt", "lp_valid_mt", "lp_unconfirmed_mt"]
    for col in qty_cols:
        if col in master.columns:
            master[col] = master[col].fillna(0)
    if "lp_has_record" in master.columns:
        master["lp_has_record"] = master["lp_has_record"].fillna(False).astype(bool)

    return master
