"""
S&OE Order Fulfillment Tracking - Main Entry Point
Usage: python run_soe.py
"""
import yaml
from pathlib import Path

from src.extract.sc_extractor import extract_sc
from src.extract.shipped_extractor import extract_shipped
from src.extract.fg_extractor import extract_fg
from src.extract.pp_extractor import (
    extract_pp_scheduled,
    extract_pp_unscheduled,
    aggregate_pp_by_so,
    aggregate_pp_unsched_by_so,
)
from src.extract.lp_extractor import (
    extract_lp_ks,
    extract_lp_idn_export,
    extract_lp_idn_domestic,
    combine_loading_plans,
)
from src.transform.join_engine import build_so_master
from src.transform.status_engine import assign_status_and_gap
from src.output.excel_writer import write_excel
from src.output.html_report import write_html_report


def main():
    # Load config
    config_path = Path(__file__).parent / "config.yaml"
    with open(config_path, "r", encoding="utf-8") as f:
        cfg = yaml.safe_load(f)

    root = Path(__file__).parent / cfg["data_root"]
    month = cfg["month"]
    target_mt = cfg["target_mt"]
    data_date = cfg["data_date"]
    output_dir = Path(__file__).parent / cfg["output_dir"]

    print(f"=== S&OE Tracking: {month} ===")
    print(f"Data date: {data_date} | Target: {target_mt:,} MT")
    print()

    # --- Phase 1: Extract ---
    print("[1/6] Extracting SC baseline...")
    customer_mapping = cfg["files"].get("customer_mapping")
    customer_mapping_path = str(root / customer_mapping) if customer_mapping else None
    sc = extract_sc(str(root / cfg["files"]["sc"]), month, customer_mapping_path, data_date)
    baseline = sc[sc["in_baseline"]]
    print(f"  SC total: {len(sc)} SOs, Baseline: {len(baseline)} SOs, {baseline['sc_vol_mt'].sum():,.0f} MT")
    unmatched_customer = sc.attrs.get("unmatched_customer")
    if unmatched_customer is not None and not unmatched_customer.empty:
        print(f"  Customer mapping unmatched: {len(unmatched_customer)} rows, {unmatched_customer['raw_sc_vol_mt'].sum():,.0f} MT")

    print("[2/6] Extracting Shipped...")
    shipped = extract_shipped(str(root / cfg["files"]["shipped"]))
    print(f"  Shipped: {len(shipped)} SOs, {shipped['shipped_mt'].sum():,.0f} MT")

    print("[3/6] Extracting FG inventory...")
    fg = extract_fg(str(root / cfg["files"]["fg"]))
    print(f"  FG: {len(fg)} SOs, {fg['fg_mt'].sum():,.0f} MT")

    print("[4/6] Extracting Production Plan...")
    pp_files = [str(root / f) for f in cfg["files"]["pp_scheduled"]]
    pp_sched = extract_pp_scheduled(pp_files)
    pp_unsched = extract_pp_unscheduled(str(root / cfg["files"]["pp_unscheduled"]))
    pp_sched_agg = aggregate_pp_by_so(pp_sched)
    pp_unsched_agg = aggregate_pp_unsched_by_so(pp_unsched)
    print(f"  PP Scheduled: {len(pp_sched_agg)} SOs, {pp_sched_agg['wip_mt'].sum():,.0f} MT")
    print(f"  PP Unscheduled: {len(pp_unsched_agg)} SOs, {pp_unsched_agg['planned_mt'].sum():,.0f} MT")

    print("[5/6] Extracting Loading Plans...")
    lp_ks = extract_lp_ks(str(root / cfg["files"]["lp_ks"]), data_date)
    lp_idn_ex = extract_lp_idn_export(str(root / cfg["files"]["lp_idn_export"]), data_date)
    lp_idn_dom = extract_lp_idn_domestic(str(root / cfg["files"]["lp_idn_domestic"]), data_date)
    loading_plan = combine_loading_plans(lp_ks, lp_idn_ex, lp_idn_dom)
    print(f"  Loading Plan: {len(loading_plan)} SOs (KS:{len(lp_ks)}, IDN_Ex:{len(lp_idn_ex)}, IDN_Dom:{len(lp_idn_dom)})")

    # --- Phase 2: Join ---
    print("[6/6] Building SO Master & calculating gap...")
    master = build_so_master(sc, shipped, fg, pp_sched_agg, pp_unsched_agg, loading_plan)
    master = assign_status_and_gap(master)
    print(f"  Master: {len(master)} SOs")
    print()

    # --- Phase 3: Output ---
    print("Writing outputs...")
    excel_path = write_excel(
        master,
        str(output_dir),
        month,
        target_mt,
        data_date,
        sc.attrs,
        cfg.get("excel_version_suffix", ""),
    )
    print(f"  Excel: {excel_path}")

    html_path = write_html_report(master, str(output_dir), month, target_mt, data_date)
    print(f"  HTML:  {html_path}")

    # Quick summary
    print()
    print("=== Summary ===")
    print(f"  Baseline:  {master['sc_vol_mt'].sum():>10,.0f} MT ({len(master)} SOs)")
    print(f"  Shipped:   {master['allocated_shipped_mt'].sum():>10,.0f} MT")
    print(f"  In Stock:  {master['allocated_fg_mt'].sum():>10,.0f} MT")
    print(f"  WIP:       {master['allocated_wip_mt'].sum():>10,.0f} MT")
    print(f"  At Risk:   {(master['allocated_unsched_mt'].sum() + master['allocated_no_plan_mt'].sum()):>10,.0f} MT")
    print()
    risk_counts = master["risk_tier"].value_counts()
    for tier in ["Green", "Yellow", "Orange", "Red", "Critical"]:
        count = risk_counts.get(tier, 0)
        if count > 0:
            print(f"  {tier:10s}: {count} SOs")
    print()
    print("Done.")


if __name__ == "__main__":
    main()
