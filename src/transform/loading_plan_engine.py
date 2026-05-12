"""
Loading Plan analysis engine.

Creates the SC-vs-LP reconciliation and Shipping Readiness Matrix from clean
Loading Plan demand lines and existing supply-status sources.
"""
import pandas as pd


def build_loading_plan_analysis(
    sc: pd.DataFrame,
    master: pd.DataFrame,
    loading_lines: pd.DataFrame,
    shipped: pd.DataFrame,
    fg: pd.DataFrame,
    pp_sched_agg: pd.DataFrame,
    pp_unsched_agg: pd.DataFrame,
    data_date: str,
) -> dict:
    """Return all Loading Plan output frames used by Excel reporting."""
    if loading_lines is None:
        loading_lines = pd.DataFrame()

    clean_detail = loading_lines.copy()
    main_lp = _main_lp_lines(clean_detail)
    supply = _build_supply_by_so(shipped, fg, pp_sched_agg, pp_unsched_agg)
    sc_base = sc[sc["in_baseline"]].copy()

    shipping_readiness = _build_shipping_readiness(main_lp, sc_base, supply, data_date)
    reconciliation = _build_reconciliation(sc_base, master, main_lp, shipping_readiness)

    return {
        "clean_detail": clean_detail,
        "reconciliation": reconciliation,
        "shipping_readiness": shipping_readiness,
        "date_exceptions": main_lp[main_lp["loading_date_status"] != "Valid Date"].copy(),
        "parse_exceptions": clean_detail[clean_detail["so_parse_status"] != "Parsed"].copy(),
        "excluded_prior_invoiced": clean_detail[clean_detail["exclude_from_current_invoice"]].copy(),
    }


def _main_lp_lines(clean_detail: pd.DataFrame) -> pd.DataFrame:
    if clean_detail.empty:
        return clean_detail.copy()
    return clean_detail[
        (~clean_detail["exclude_from_current_invoice"])
        & (clean_detail["so_parse_status"] == "Parsed")
        & clean_detail["so"].notna()
        & (clean_detail["so"].astype(str).str.strip() != "")
    ].copy()


def _build_supply_by_so(
    shipped: pd.DataFrame,
    fg: pd.DataFrame,
    pp_sched_agg: pd.DataFrame,
    pp_unsched_agg: pd.DataFrame,
) -> pd.DataFrame:
    keys = pd.concat(
        [
            shipped[["so"]] if not shipped.empty else pd.DataFrame(columns=["so"]),
            fg[["so"]] if not fg.empty else pd.DataFrame(columns=["so"]),
            pp_sched_agg[["so"]] if not pp_sched_agg.empty else pd.DataFrame(columns=["so"]),
            pp_unsched_agg[["so"]] if not pp_unsched_agg.empty else pd.DataFrame(columns=["so"]),
        ],
        ignore_index=True,
    ).drop_duplicates()

    supply = keys.copy()
    if supply.empty:
        return pd.DataFrame(columns=[
            "so", "raw_shipped_mt", "raw_fg_mt", "raw_wip_mt",
            "raw_unsched_mt", "planned_end_date", "supply_machines",
        ])

    supply = supply.merge(
        shipped[["so", "shipped_mt"]].rename(columns={"shipped_mt": "raw_shipped_mt"}),
        on="so", how="left",
    )
    supply = supply.merge(
        fg[["so", "fg_mt"]].rename(columns={"fg_mt": "raw_fg_mt"}),
        on="so", how="left",
    )
    supply = supply.merge(
        pp_sched_agg[["so", "wip_mt", "planned_end_date", "machines"]].rename(
            columns={"wip_mt": "raw_wip_mt", "machines": "supply_machines"}
        ),
        on="so", how="left",
    )
    supply = supply.merge(
        pp_unsched_agg[["so", "planned_mt"]].rename(columns={"planned_mt": "raw_unsched_mt"}),
        on="so", how="left",
    )

    for col in ["raw_shipped_mt", "raw_fg_mt", "raw_wip_mt", "raw_unsched_mt"]:
        supply[col] = pd.to_numeric(supply[col], errors="coerce").fillna(0)
    return supply


def _build_shipping_readiness(
    main_lp: pd.DataFrame,
    sc_base: pd.DataFrame,
    supply: pd.DataFrame,
    data_date: str,
) -> pd.DataFrame:
    if main_lp.empty:
        return pd.DataFrame()

    df = main_lp.merge(
        sc_base[["so", "plant", "cluster", "sc_vol_mt"]].rename(
            columns={"plant": "sc_plant", "cluster": "sc_cluster", "sc_vol_mt": "sc_baseline_mt"}
        ),
        on="so", how="left",
    )
    df["in_sc_baseline"] = df["sc_baseline_mt"].notna()
    df = df.merge(supply, on="so", how="left")
    for col in ["raw_shipped_mt", "raw_fg_mt", "raw_wip_mt", "raw_unsched_mt"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    run_date = pd.to_datetime(data_date).normalize()
    df["loading_date_bucket"] = "No Valid Date"
    valid_date = df["loading_date_status"] == "Valid Date"
    df.loc[valid_date & (df["loading_date"] < run_date), "loading_date_bucket"] = "Past"
    df.loc[valid_date & (df["loading_date"] >= run_date), "loading_date_bucket"] = "Today/Future"

    df["available_date_from_wip"] = df["planned_end_date"] + pd.Timedelta(days=1)
    df["lp_gap_days"] = pd.NA
    has_gap = valid_date & df["available_date_from_wip"].notna()
    df.loc[has_gap, "lp_gap_days"] = (
        df.loc[has_gap, "loading_date"] - df.loc[has_gap, "available_date_from_wip"]
    ).dt.days

    statuses = df.apply(_readiness_status, axis=1)
    df["shipping_readiness_status"] = [s[0] for s in statuses]
    df["shipping_readiness_tier"] = [s[1] for s in statuses]
    df["lp_fulfillment_note"] = [s[2] for s in statuses]
    df["past_loading_not_shipped"] = (
        (df["loading_date_bucket"] == "Past") & (df["raw_shipped_mt"] <= 0)
    )

    return df


def _readiness_status(row) -> tuple[str, str, str]:
    if not row["in_sc_baseline"]:
        return (
            "LP Not In SC Baseline",
            "Review",
            "Loading Plan has this SO, but the SO is not in current SC baseline.",
        )

    valid_date = row["loading_date_status"] == "Valid Date"
    has_shipped = row["raw_shipped_mt"] > 0
    has_fg = row["raw_fg_mt"] > 0
    has_wip = row["raw_wip_mt"] > 0
    has_unsched = row["raw_unsched_mt"] > 0

    if not valid_date:
        if has_shipped:
            return ("Unconfirmed Loading - Shipped Exists", "Review", "No valid loading date; GI shipped signal exists.")
        if has_fg:
            return ("Unconfirmed Loading - FG Ready", "Orange", "No valid loading date; finished goods are available.")
        if has_wip:
            return ("Unconfirmed Loading - WIP Scheduled", "Orange", "No valid loading date; production schedule exists.")
        if has_unsched:
            return ("Unconfirmed Loading - Unscheduled WO", "Red", "No valid loading date; work order has no finish date.")
        return ("Unconfirmed Loading - No Supply", "Critical", "No valid loading date and no supply signal found.")

    if has_shipped:
        return ("Covered by Shipped", "Green", "GI shipped quantity exists and is treated as fulfillment evidence.")
    if has_fg:
        return ("Covered by FG", "Green", "Finished goods are available for loading.")
    if has_wip and pd.notna(row["available_date_from_wip"]):
        if row["available_date_from_wip"] <= row["loading_date"]:
            gap = row["lp_gap_days"]
            if pd.notna(gap) and gap <= 2:
                return ("WIP Tight", "Yellow", "Production can meet loading date but buffer is tight.")
            return ("WIP On Time", "Green", "Production finish plus one day can meet loading date.")
        return ("WIP Late", "Red", "Production finish plus one day is later than loading date.")
    if has_unsched:
        return ("Unscheduled WO", "Red", "Work order exists but no planned finish date is available.")
    return ("No Supply Signal", "Critical", "No shipped, FG, scheduled WIP, or unscheduled WO signal found.")


def _build_reconciliation(
    sc_base: pd.DataFrame,
    master: pd.DataFrame,
    main_lp: pd.DataFrame,
    shipping_readiness: pd.DataFrame,
) -> pd.DataFrame:
    lp_so = _lp_so_summary(main_lp, shipping_readiness)

    sc_cols = ["so", "plant", "cluster", "sc_vol_mt", "order_type"]
    sc_view = sc_base[sc_cols].copy()
    rec = sc_view.merge(lp_so, on="so", how="outer")

    rec["in_sc_baseline"] = rec["sc_vol_mt"].notna()
    rec["in_loading_plan"] = rec["lp_load_mt"].fillna(0) > 0
    rec["reconciliation_status"] = rec.apply(_reconciliation_status, axis=1)

    master_cols = [
        "so", "status", "risk_tier", "allocated_shipped_mt", "allocated_fg_mt",
        "allocated_wip_mt", "allocated_unsched_mt", "allocated_no_plan_mt",
    ]
    available_master_cols = [c for c in master_cols if c in master.columns]
    rec = rec.merge(master[available_master_cols], on="so", how="left")
    rec["reconciliation_note"] = rec.apply(_reconciliation_note, axis=1)

    for col in ["sc_vol_mt", "lp_load_mt", "lp_valid_mt", "lp_unconfirmed_mt"]:
        if col in rec.columns:
            rec[col] = pd.to_numeric(rec[col], errors="coerce").fillna(0)
    return rec.sort_values(["reconciliation_status", "so"], na_position="last")


def _lp_so_summary(main_lp: pd.DataFrame, shipping_readiness: pd.DataFrame) -> pd.DataFrame:
    if main_lp.empty:
        return pd.DataFrame(columns=["so", "lp_load_mt", "lp_valid_mt", "lp_unconfirmed_mt"])
    lp = main_lp.copy()
    lp["lp_valid_mt"] = lp["load_mt"].where(lp["loading_date_status"] == "Valid Date", 0)
    lp["lp_unconfirmed_mt"] = lp["load_mt"].where(lp["loading_date_status"] != "Valid Date", 0)
    summary = lp.groupby("so").agg(
        lp_load_mt=("load_mt", "sum"),
        lp_valid_mt=("lp_valid_mt", "sum"),
        lp_unconfirmed_mt=("lp_unconfirmed_mt", "sum"),
        lp_sources=("lp_source", lambda x: ", ".join(sorted(set(map(str, x))))),
        lp_line_count=("so", "size"),
    ).reset_index()
    if shipping_readiness is not None and not shipping_readiness.empty:
        risk = shipping_readiness.groupby("so").agg(
            worst_shipping_tier=("shipping_readiness_tier", _worst_tier),
            readiness_statuses=("shipping_readiness_status", lambda x: ", ".join(sorted(set(map(str, x))))),
        ).reset_index()
        summary = summary.merge(risk, on="so", how="left")
    return summary


def _reconciliation_status(row) -> str:
    if row["in_sc_baseline"] and row["in_loading_plan"]:
        return "In SC and In LP"
    if row["in_sc_baseline"] and not row["in_loading_plan"]:
        return "In SC only"
    if not row["in_sc_baseline"] and row["in_loading_plan"]:
        return "In LP only"
    return "No Match"


def _reconciliation_note(row) -> str:
    status = row["reconciliation_status"]
    if status == "In SC and In LP":
        return "Current billing SO has loading arrangement; review Shipping Readiness for fulfillment risk."
    if status == "In LP only":
        return "Loading Plan has this SO, but it is not in current SC baseline."
    if status == "In SC only":
        if row.get("allocated_fg_mt", 0) > 0:
            return "Goods are ready in FG, but no loading arrangement is found."
        if row.get("allocated_wip_mt", 0) > 0:
            return "Production is scheduled, but no loading arrangement is found."
        if row.get("allocated_unsched_mt", 0) > 0:
            return "Work order exists without schedule, and no loading arrangement is found."
        if row.get("allocated_no_plan_mt", 0) > 0:
            return "No loading arrangement and no supply plan for the remaining quantity."
        if row.get("allocated_shipped_mt", 0) > 0:
            return "Shipped signal exists but no loading plan match; verify if LP is missing or already closed."
    return ""


def _worst_tier(values: pd.Series) -> str:
    order = {"Critical": 0, "Red": 1, "Orange": 2, "Review": 3, "Yellow": 4, "Green": 5}
    clean = [v for v in values.dropna().astype(str).tolist() if v]
    if not clean:
        return ""
    return sorted(clean, key=lambda x: order.get(x, 99))[0]
