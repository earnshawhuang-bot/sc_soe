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

    clean_detail = _annotate_lp_scope(loading_lines.copy(), data_date)
    main_lp = _main_lp_lines(clean_detail)
    supply = _build_supply_by_so(shipped, fg, pp_sched_agg, pp_unsched_agg)
    sc_base = sc[sc["in_baseline"]].copy()

    shipping_readiness = _build_shipping_readiness(main_lp, sc_base, supply, data_date)
    lp_evidence = _lp_evidence_summary(clean_detail)
    reconciliation = _build_reconciliation(sc_base, master, main_lp, shipping_readiness, lp_evidence)
    risk_matrix_detail = _build_risk_matrix_detail(master, main_lp, reconciliation, data_date)
    lp_not_in_baseline_detail = _build_lp_not_in_baseline_detail(main_lp, reconciliation, data_date)
    lp_not_in_baseline_summary = _build_lp_not_in_baseline_summary(lp_not_in_baseline_detail)

    return {
        "clean_detail": clean_detail,
        "reconciliation": reconciliation,
        "risk_matrix_detail": risk_matrix_detail,
        "lp_not_in_baseline_summary": lp_not_in_baseline_summary,
        "lp_not_in_baseline_detail": lp_not_in_baseline_detail,
        "shipping_readiness": shipping_readiness,
        "date_exceptions": main_lp[main_lp["loading_date_status"] != "Valid Date"].copy(),
        "parse_exceptions": clean_detail[clean_detail["so_parse_status"] != "Parsed"].copy(),
        "excluded_prior_invoiced": clean_detail[clean_detail["exclude_from_current_invoice"]].copy(),
    }


def _main_lp_lines(clean_detail: pd.DataFrame) -> pd.DataFrame:
    if clean_detail.empty:
        return clean_detail.copy()
    return clean_detail[
        (clean_detail["lp_scope"] == "Current LP")
        & (clean_detail["so_parse_status"] == "Parsed")
        & clean_detail["so"].notna()
        & (clean_detail["so"].astype(str).str.strip() != "")
    ].copy()


def _annotate_lp_scope(clean_detail: pd.DataFrame, data_date: str) -> pd.DataFrame:
    if clean_detail.empty:
        clean_detail["lp_scope"] = pd.Series(dtype=str)
        return clean_detail

    clean_detail["loading_date"] = pd.to_datetime(clean_detail["loading_date"], errors="coerce")
    clean_detail["lp_scope"] = "Out of Scope"

    parsed_so = (
        (clean_detail["so_parse_status"] == "Parsed")
        & clean_detail["so"].notna()
        & (clean_detail["so"].astype(str).str.strip() != "")
    )
    clean_detail.loc[parsed_so, "lp_scope"] = "Current LP"
    clean_detail.loc[clean_detail["exclude_from_current_invoice"].fillna(False), "lp_scope"] = "Excluded LP"

    if data_date:
        month_start = pd.to_datetime(data_date).replace(day=1).normalize()
        historical_valid = (
            parsed_so
            & (~clean_detail["exclude_from_current_invoice"].fillna(False))
            & (clean_detail["loading_date_status"] == "Valid Date")
            & (clean_detail["loading_date"] < month_start)
        )
        clean_detail.loc[historical_valid, "lp_scope"] = "Historical LP"

    return clean_detail


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
            "raw_unsched_mt", "planned_end_date", "supply_machines", "supply_work_orders",
        ])

    supply = supply.merge(
        shipped[["so", "shipped_mt"]].rename(columns={"shipped_mt": "raw_shipped_mt"}),
        on="so", how="left",
    )
    supply = supply.merge(
        fg[["so", "fg_mt"]].rename(columns={"fg_mt": "raw_fg_mt"}),
        on="so", how="left",
    )
    pp_sched_cols = ["so", "wip_mt", "planned_end_date", "machines"]
    if "work_orders" in pp_sched_agg.columns:
        pp_sched_cols.append("work_orders")
    supply = supply.merge(
        pp_sched_agg[pp_sched_cols].rename(
            columns={
                "wip_mt": "raw_wip_mt",
                "machines": "supply_machines",
                "work_orders": "supply_work_orders",
            }
        ),
        on="so", how="left",
    )
    supply = supply.merge(
        pp_unsched_agg[["so", "planned_mt"]].rename(columns={"planned_mt": "raw_unsched_mt"}),
        on="so", how="left",
    )

    for col in ["raw_shipped_mt", "raw_fg_mt", "raw_wip_mt", "raw_unsched_mt"]:
        supply[col] = pd.to_numeric(supply[col], errors="coerce").fillna(0)
    if "supply_work_orders" not in supply.columns:
        supply["supply_work_orders"] = ""
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
    df.loc[has_gap, "lp_gap_days"] = _gap_days(
        df.loc[has_gap, "loading_date"],
        df.loc[has_gap, "available_date_from_wip"],
    )

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
        if _meets_loading_date(row["available_date_from_wip"], row["loading_date"]):
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
    lp_evidence: pd.DataFrame,
) -> pd.DataFrame:
    lp_so = _lp_so_summary(main_lp, shipping_readiness)

    sc_cols = ["so", "plant", "cluster", "sc_vol_mt", "order_type"]
    sc_view = sc_base[sc_cols].copy()
    rec = sc_view.merge(lp_so, on="so", how="outer")
    if lp_evidence is not None and not lp_evidence.empty:
        rec = rec.merge(lp_evidence, on="so", how="left")

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
    rec["lp_match_scope"] = rec.apply(_lp_match_scope, axis=1)
    rec["lp_parse_exception_flag"] = False
    actionable = (rec["sc_vol_mt"].fillna(0) > 0) | (rec["lp_load_mt"].fillna(0) > 0)
    rec = rec[actionable].copy()
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
        lp_loading_date_raw_list=("loading_date_raw", _join_unique),
        lp_earliest_valid_loading_date=("loading_date", "min"),
        lp_loading_date_status_mix=("loading_date_status", _join_unique),
    ).reset_index()
    if shipping_readiness is not None and not shipping_readiness.empty:
        risk = shipping_readiness.groupby("so").agg(
            worst_shipping_tier=("shipping_readiness_tier", _worst_tier),
            readiness_statuses=("shipping_readiness_status", lambda x: ", ".join(sorted(set(map(str, x))))),
        ).reset_index()
        summary = summary.merge(risk, on="so", how="left")
    return summary


def _lp_evidence_summary(clean_detail: pd.DataFrame) -> pd.DataFrame:
    if clean_detail.empty:
        return pd.DataFrame(columns=["so", "lp_all_scopes", "lp_all_line_count"])
    evidence = clean_detail[
        (clean_detail["so_parse_status"] == "Parsed")
        & clean_detail["so"].notna()
        & (clean_detail["so"].astype(str).str.strip() != "")
    ].copy()
    if evidence.empty:
        return pd.DataFrame(columns=["so", "lp_all_scopes", "lp_all_line_count"])
    return evidence.groupby("so").agg(
        lp_all_scopes=("lp_scope", _join_unique),
        lp_all_line_count=("so", "size"),
    ).reset_index()


def _lp_match_scope(row) -> str:
    if row.get("in_loading_plan", False):
        return "Current LP"
    scopes = str(row.get("lp_all_scopes", "") or "")
    if "Historical LP" in scopes:
        return "Historical LP only"
    if "Excluded LP" in scopes:
        return "Excluded LP only"
    return "No LP evidence"


def _build_lp_not_in_baseline_detail(
    main_lp: pd.DataFrame,
    reconciliation: pd.DataFrame,
    data_date: str,
) -> pd.DataFrame:
    columns = [
        "lp_coverage_status",
        "plant",
        "lp_source",
        "source_file",
        "source_sheet",
        "source_row",
        "invoice_no_raw",
        "so",
        "loading_date_raw",
        "loading_date",
        "loading_date_status",
        "load_mt",
        "source_mt",
    ]
    if main_lp is None or main_lp.empty or reconciliation is None or reconciliation.empty:
        return pd.DataFrame(columns=columns)

    lp_only_so = set(
        reconciliation.loc[
            reconciliation["reconciliation_status"] == "In LP only",
            "so",
        ].dropna().astype(str)
    )
    if not lp_only_so:
        return pd.DataFrame(columns=columns)

    detail = main_lp[main_lp["so"].astype(str).isin(lp_only_so)].copy()
    if detail.empty:
        return pd.DataFrame(columns=columns)

    detail["load_mt"] = pd.to_numeric(detail["load_mt"], errors="coerce").fillna(0)
    detail = detail[detail["load_mt"] > 0].copy()
    detail["loading_date"] = pd.to_datetime(detail["loading_date"], errors="coerce")
    run_date = pd.to_datetime(data_date).normalize()
    detail["lp_coverage_status"] = "LP Date Unconfirmed"
    valid = detail["loading_date_status"] == "Valid Date"
    detail.loc[valid & (detail["loading_date"] < run_date), "lp_coverage_status"] = "Past Due LP"
    detail.loc[valid & (detail["loading_date"] >= run_date), "lp_coverage_status"] = "Future Valid LP"

    for col in columns:
        if col not in detail.columns:
            detail[col] = pd.NA

    status_order = {"Past Due LP": 0, "Future Valid LP": 1, "LP Date Unconfirmed": 2}
    detail["_status_order"] = detail["lp_coverage_status"].map(status_order).fillna(99)
    detail = detail.sort_values(
        ["_status_order", "loading_date", "load_mt", "so"],
        ascending=[True, True, False, True],
        na_position="last",
    )
    return detail[columns].reset_index(drop=True)


def _build_lp_not_in_baseline_summary(detail: pd.DataFrame) -> pd.DataFrame:
    statuses = ["Past Due LP", "Future Valid LP", "LP Date Unconfirmed"]
    if detail is None or detail.empty:
        rows = [{"lp_coverage_status": status, "load_mt": 0.0} for status in statuses]
    else:
        grouped = detail.groupby("lp_coverage_status", dropna=False)["load_mt"].sum()
        rows = [
            {"lp_coverage_status": status, "load_mt": float(grouped.get(status, 0.0))}
            for status in statuses
        ]
    rows.append({
        "lp_coverage_status": "Total",
        "load_mt": sum(row["load_mt"] for row in rows),
    })
    return pd.DataFrame(rows)


def _build_risk_matrix_detail(
    master: pd.DataFrame,
    main_lp: pd.DataFrame,
    reconciliation: pd.DataFrame,
    data_date: str,
) -> pd.DataFrame:
    """Build the segment-level fact table behind Summary and Action Required."""
    if master is None or master.empty:
        return pd.DataFrame()

    rec_cols = [
        "so",
        "lp_valid_mt",
        "lp_unconfirmed_mt",
        "lp_loading_date_raw_list",
        "lp_earliest_valid_loading_date",
        "lp_loading_date_status_mix",
        "lp_match_scope",
        "lp_line_count",
        "readiness_statuses",
        "worst_shipping_tier",
    ]
    rec_cols = [c for c in rec_cols if reconciliation is not None and c in reconciliation.columns]
    base = master.copy()
    if rec_cols:
        duplicate_cols = [c for c in rec_cols if c != "so" and c in base.columns]
        if duplicate_cols:
            base = base.drop(columns=duplicate_cols)
        base = base.merge(reconciliation[rec_cols].drop_duplicates("so"), on="so", how="left")

    for col in ["lp_valid_mt", "lp_unconfirmed_mt"]:
        if col not in base.columns:
            base[col] = 0
        base[col] = pd.to_numeric(base[col], errors="coerce").fillna(0)
    if "lp_match_scope" not in base.columns:
        base["lp_match_scope"] = "No LP evidence"
    base["lp_match_scope"] = base["lp_match_scope"].fillna("No LP evidence")
    lp_bucket_map = _lp_bucket_map(main_lp, data_date)

    segment_fields = [
        ("Shipped", "allocated_shipped_mt"),
        ("FG", "allocated_fg_mt"),
        ("WIP Scheduled", "allocated_wip_mt"),
        ("WIP Unscheduled", "allocated_unsched_mt"),
        ("No Supply Signal", "allocated_no_plan_mt"),
    ]

    rows = []
    for _, so_row in base.iterrows():
        covered_mt = 0.0
        so = str(so_row.get("so", "") or "")
        lp_remaining = {
            status: dict(values)
            for status, values in lp_bucket_map.get(so, {}).items()
        }
        for supply_status, qty_col in segment_fields:
            supply_mt = pd.to_numeric(so_row.get(qty_col, 0), errors="coerce")
            if pd.isna(supply_mt) or supply_mt <= 0:
                continue
            pieces = _allocate_lp_coverage(float(supply_mt), lp_remaining)
            for piece in pieces:
                lp_coverage_status = piece["lp_coverage_status"]
                lp_original_status = piece["lp_original_coverage_status"]
                if supply_status == "Shipped" and lp_original_status in ["Past Due LP", "Future Valid LP"]:
                    lp_coverage_status = "Shipped LP Closed"
                has_valid_lp_date = lp_original_status in ["Past Due LP", "Future Valid LP"]
                piece_loading_date = piece.get("loading_date")
                if not has_valid_lp_date or pd.isna(piece_loading_date):
                    piece_loading_date = pd.NaT
                piece_raw_dates = piece.get("lp_loading_date_raw_list", "")
                piece_date_statuses = piece.get("lp_loading_date_status_mix", "")

                planned_end_date = so_row.get("planned_end_date") if supply_status == "WIP Scheduled" else pd.NaT
                available_date = _available_date_from_planned_end(supply_status, planned_end_date)
                machines = so_row.get("machines") if supply_status == "WIP Scheduled" else ""
                work_orders = so_row.get("work_orders") if supply_status == "WIP Scheduled" else ""
                segment_gap_days = _segment_gap_days(supply_status, available_date, piece_loading_date)
                action = _segment_action(
                    supply_status=supply_status,
                    lp_coverage_status=lp_coverage_status,
                    lp_original_coverage_status=lp_original_status,
                    gap_days=segment_gap_days,
                )
                rows.append({
                    "so": so_row.get("so"),
                    "plant": so_row.get("plant"),
                    "cluster": so_row.get("cluster"),
                    "order_type": so_row.get("order_type"),
                    "so_total_mt": so_row.get("sc_vol_mt", 0),
                    "supply_status": supply_status,
                    "supply_segment_total_mt": supply_mt,
                    "lp_coverage_status": lp_coverage_status,
                    "lp_original_coverage_status": lp_original_status,
                    "risk_mt": piece["risk_mt"],
                    "covered_mt": covered_mt,
                    "so_master_status": so_row.get("status"),
                    "so_risk_tier": so_row.get("risk_tier"),
                    "planned_end_date": planned_end_date,
                    "available_date": available_date,
                    "machines": machines,
                    "work_orders": work_orders,
                    "loading_date": piece_loading_date,
                    "lp_gap_days": segment_gap_days,
                    "lp_valid_mt": so_row.get("lp_valid_mt", 0),
                    "lp_unconfirmed_mt": so_row.get("lp_unconfirmed_mt", 0),
                    "lp_match_scope": so_row.get("lp_match_scope", "No LP evidence"),
                    "lp_line_count": so_row.get("lp_line_count"),
                    "lp_loading_date_raw_list": piece_raw_dates,
                    "lp_earliest_valid_loading_date": piece_loading_date,
                    "lp_loading_date_status_mix": piece_date_statuses,
                    "readiness_statuses": so_row.get("readiness_statuses"),
                    "worst_shipping_tier": so_row.get("worst_shipping_tier"),
                    "risk_action": action[0],
                    "action_required": action[1],
                    "suggested_owner": action[2],
                    "action_note": action[3],
                })
            covered_mt += float(supply_mt)

    detail = pd.DataFrame(rows)
    if detail.empty:
        return detail
    detail["risk_mt"] = pd.to_numeric(detail["risk_mt"], errors="coerce").fillna(0)
    detail["covered_mt"] = pd.to_numeric(detail["covered_mt"], errors="coerce").fillna(0)
    detail["action_required"] = detail["action_required"].fillna(False).astype(bool)
    return detail.sort_values(["action_required", "risk_action", "risk_mt"], ascending=[False, True, False])


def _available_date_from_planned_end(supply_status: str, planned_end_date):
    if supply_status != "WIP Scheduled":
        return pd.NaT
    if pd.isna(planned_end_date):
        return pd.NaT
    planned = pd.to_datetime(planned_end_date, errors="coerce")
    if pd.isna(planned):
        return pd.NaT
    return planned + pd.Timedelta(days=1)


def _segment_gap_days(supply_status: str, available_date, loading_date):
    if supply_status != "WIP Scheduled":
        return pd.NA
    if pd.isna(available_date) or pd.isna(loading_date):
        return pd.NA
    available = pd.to_datetime(available_date, errors="coerce")
    loading = pd.to_datetime(loading_date, errors="coerce")
    if pd.isna(available) or pd.isna(loading):
        return pd.NA
    if _is_date_grain(loading):
        return (loading.normalize() - available.normalize()).days
    return (loading - available).days


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


def _meets_loading_date(available_date, loading_date) -> bool:
    available = pd.to_datetime(available_date, errors="coerce")
    loading = pd.to_datetime(loading_date, errors="coerce")
    if pd.isna(available) or pd.isna(loading):
        return False
    if _is_date_grain(loading):
        return available.normalize() <= loading.normalize()
    return available <= loading


def _is_date_grain(value) -> bool:
    ts = pd.to_datetime(value, errors="coerce")
    if pd.isna(ts):
        return True
    return not (ts.hour or ts.minute or ts.second or ts.microsecond)


def _lp_bucket_map(main_lp: pd.DataFrame, data_date: str) -> dict:
    if main_lp is None or main_lp.empty:
        return {}
    lp = main_lp.copy()
    lp["load_mt"] = pd.to_numeric(lp["load_mt"], errors="coerce").fillna(0)
    lp = lp[lp["load_mt"] > 0].copy()
    if lp.empty:
        return {}

    run_date = pd.to_datetime(data_date).normalize()
    lp["loading_date"] = pd.to_datetime(lp["loading_date"], errors="coerce")
    lp["lp_coverage_status"] = "LP Date Unconfirmed"
    valid = lp["loading_date_status"] == "Valid Date"
    lp.loc[valid & (lp["loading_date"] < run_date), "lp_coverage_status"] = "Past Due LP"
    lp.loc[valid & (lp["loading_date"] >= run_date), "lp_coverage_status"] = "Future Valid LP"

    grouped = lp.groupby(["so", "lp_coverage_status"]).agg(
        remaining_mt=("load_mt", "sum"),
        lp_loading_date_raw_list=("loading_date_raw", _join_unique),
        lp_loading_date_status_mix=("loading_date_status", _join_unique),
        loading_date=("loading_date", "min"),
    ).reset_index()

    result = {}
    for _, row in grouped.iterrows():
        result.setdefault(str(row["so"]), {})[row["lp_coverage_status"]] = {
            "remaining_mt": float(row["remaining_mt"]),
            "lp_loading_date_raw_list": row.get("lp_loading_date_raw_list"),
            "lp_loading_date_status_mix": row.get("lp_loading_date_status_mix"),
            "loading_date": row.get("loading_date"),
        }
    return result


def _allocate_lp_coverage(quantity: float, lp_remaining: dict) -> list[dict]:
    pieces = []
    remaining_qty = float(quantity)
    for status in ["Past Due LP", "Future Valid LP", "LP Date Unconfirmed"]:
        if remaining_qty <= 0:
            break
        bucket = lp_remaining.get(status)
        available = float(bucket.get("remaining_mt", 0)) if bucket else 0.0
        if available <= 0:
            continue
        take = min(remaining_qty, available)
        bucket["remaining_mt"] = available - take
        pieces.append({
            "lp_coverage_status": status,
            "lp_original_coverage_status": status,
            "risk_mt": take,
            "lp_loading_date_raw_list": bucket.get("lp_loading_date_raw_list"),
            "lp_loading_date_status_mix": bucket.get("lp_loading_date_status_mix"),
            "loading_date": bucket.get("loading_date"),
        })
        remaining_qty -= take
    if remaining_qty > 0.000001:
        pieces.append({
            "lp_coverage_status": "No Current LP",
            "lp_original_coverage_status": "No Current LP",
            "risk_mt": remaining_qty,
            "lp_loading_date_raw_list": "",
            "lp_loading_date_status_mix": "",
            "loading_date": pd.NaT,
        })
    return pieces


def _segment_action(
    supply_status: str,
    lp_coverage_status: str,
    lp_original_coverage_status: str,
    gap_days,
) -> tuple[str, bool, str, str]:
    if supply_status == "Shipped":
        if lp_coverage_status == "Shipped LP Closed":
            return ("Fulfilled - LP Closed", False, "Sales Ops", "Shipment exists and LP evidence closes the loop.")
        if lp_coverage_status == "LP Date Unconfirmed":
            return ("Shipped but LP Date Unconfirmed", False, "Sales Ops", "Shipment exists; LP closure can be reviewed as data hygiene.")
        return ("Shipped but LP Missing", False, "Sales Ops", "Shipment exists; LP gap is a data review item, not fulfillment exposure.")

    if supply_status == "FG":
        if lp_coverage_status == "Past Due LP":
            return ("Past Due LP - FG Ready", True, "Logistics", "Goods are ready, but the loading date is already past and not closed by shipment.")
        if lp_coverage_status == "Future Valid LP":
            return ("FG Ready with Loading Plan", False, "Logistics", "Goods are ready and a confirmed loading date exists.")
        if lp_coverage_status == "LP Date Unconfirmed":
            return ("FG Ready - LP Date Unconfirmed", True, "Logistics", "Goods are ready, but loading date is not confirmed.")
        return ("FG without Loading Plan", True, "Logistics", "Goods are ready, but no current loading arrangement is found.")

    if supply_status == "WIP Scheduled":
        if lp_coverage_status == "Past Due LP":
            return ("Past Due LP - WIP Scheduled", True, "Planning", "Loading date is already past while production is still scheduled/open.")
        if lp_coverage_status == "Future Valid LP":
            if pd.notna(gap_days) and gap_days < 0:
                return ("WIP Late vs Loading Date", True, "Planning", "Production finish is later than confirmed loading date.")
            return ("WIP Scheduled with Loading Plan", False, "Planning", "Production has schedule and loading date can be checked by gap.")
        if lp_coverage_status == "LP Date Unconfirmed":
            return ("WIP Scheduled - LP Date Unconfirmed", True, "Logistics", "Production is scheduled, but loading date is not confirmed.")
        return ("WIP Scheduled without Loading Plan", True, "Logistics", "Production is scheduled, but no current loading arrangement is found.")

    if supply_status == "WIP Unscheduled":
        if lp_coverage_status == "Past Due LP":
            return ("Past Due LP - Production Unscheduled", True, "Planning", "Loading date is already past and work order has no planned finish date.")
        if lp_coverage_status == "Future Valid LP":
            return ("Production Unscheduled vs Loading Plan", True, "Planning", "Loading demand exists, but work order has no planned finish date.")
        if lp_coverage_status == "LP Date Unconfirmed":
            return ("Production and LP Date Both Unconfirmed", True, "Planning", "Production finish and loading date are both uncertain.")
        return ("Production Unscheduled without Loading Plan", True, "Planning", "Work order exists without finish date and no loading arrangement is found.")

    if supply_status == "No Supply Signal":
        if lp_coverage_status == "Past Due LP":
            return ("Past Due LP without Supply Signal", True, "Planning", "Loading date is already past, but no shipped / FG / PP signal supports it.")
        if lp_coverage_status == "Future Valid LP":
            return ("Loading Plan without Supply Signal", True, "Planning", "Shipping is arranged, but no shipped / FG / PP signal supports it.")
        if lp_coverage_status == "LP Date Unconfirmed":
            return ("LP Unconfirmed and No Supply Signal", True, "Planning", "LP exists but date is unclear, and no supply signal exists.")
        return ("No Supply and No Loading Signal", True, "Planning", "Baseline quantity has neither supply evidence nor loading arrangement.")

    return ("Review", True, "Sales Ops", "Unclassified risk segment; review source data.")


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


def _join_unique(values: pd.Series) -> str:
    clean = []
    for value in values.dropna().tolist():
        if pd.isna(value):
            continue
        if isinstance(value, pd.Timestamp):
            if value.hour or value.minute or value.second:
                text = value.strftime("%Y-%m-%d %H:%M")
            else:
                text = value.strftime("%Y-%m-%d")
        else:
            text = str(value).strip()
        if text:
            clean.append(text)
    return " | ".join(sorted(set(clean)))
