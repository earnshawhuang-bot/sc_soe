"""
S&OE Order Fulfillment Tracking - Main Entry Point
Usage: python run_soe.py
"""
import re
import yaml
from datetime import datetime
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
    combine_loading_plan_lines,
    aggregate_loading_plan_for_master,
)
from src.transform.join_engine import build_so_master
from src.transform.status_engine import assign_status_and_gap
from src.transform.loading_plan_engine import build_loading_plan_analysis
from src.output.excel_writer import write_excel
from src.output.risk_dashboard import write_risk_dashboard


def main():
    # Load config
    config_path = Path(__file__).parent / "config.yaml"
    with open(config_path, "r", encoding="utf-8") as f:
        cfg = yaml.safe_load(f)

    root = Path(__file__).parent / cfg["data_root"]
    month = cfg["month"]
    target_mt = cfg["target_mt"]
    runtime_files, data_date = _resolve_runtime_inputs(root, cfg, month)
    output_dir = Path(__file__).parent / "output" / data_date
    run_stamp = datetime.now().strftime("%Y%m%d-%H%M")

    print(f"=== S&OE Tracking: {month} ===")
    print(f"Data date: {data_date} | Target: {target_mt:,} MT")
    print("Runtime source files:")
    for label in ["sc", "shipped", "fg", "lp_ks", "lp_idn_export", "lp_idn_domestic", "pp_unscheduled"]:
        if label in runtime_files:
            print(f"  {label}: {runtime_files[label]}")
    print()

    # --- Phase 1: Extract ---
    print("[1/6] Extracting SC baseline...")
    customer_mapping = cfg["files"].get("customer_mapping")
    customer_mapping_path = str(root / customer_mapping) if customer_mapping else None
    sc = extract_sc(str(root / runtime_files["sc"]), month, customer_mapping_path, data_date)
    baseline = sc[sc["in_baseline"]]
    print(f"  SC total: {len(sc)} SOs, Baseline: {len(baseline)} SOs, {baseline['sc_vol_mt'].sum():,.0f} MT")
    unmatched_customer = sc.attrs.get("unmatched_customer")
    if unmatched_customer is not None and not unmatched_customer.empty:
        print(f"  Customer mapping unmatched: {len(unmatched_customer)} rows, {unmatched_customer['raw_sc_vol_mt'].sum():,.0f} MT")

    print("[2/6] Extracting Shipped...")
    shipped = extract_shipped(str(root / runtime_files["shipped"]))
    print(f"  Shipped: {len(shipped)} SOs, {shipped['shipped_mt'].sum():,.0f} MT")

    print("[3/6] Extracting FG inventory...")
    fg = extract_fg(str(root / runtime_files["fg"]))
    print(f"  FG: {len(fg)} SOs, {fg['fg_mt'].sum():,.0f} MT")

    print("[4/6] Extracting Production Plan...")
    pp_files = [str(root / f) for f in runtime_files["pp_scheduled"]]
    pp_sched = extract_pp_scheduled(pp_files)
    pp_unsched = extract_pp_unscheduled(str(root / runtime_files["pp_unscheduled"]))
    pp_sched_agg = aggregate_pp_by_so(pp_sched)
    pp_unsched_agg = aggregate_pp_unsched_by_so(pp_unsched)
    print(f"  PP Scheduled: {len(pp_sched_agg)} SOs, {pp_sched_agg['wip_mt'].sum():,.0f} MT")
    print(f"  PP Unscheduled: {len(pp_unsched_agg)} SOs, {pp_unsched_agg['planned_mt'].sum():,.0f} MT")

    print("[5/6] Extracting Loading Plans...")
    lp_ks = extract_lp_ks(str(root / runtime_files["lp_ks"]), data_date)
    lp_idn_ex = extract_lp_idn_export(str(root / runtime_files["lp_idn_export"]), data_date)
    lp_idn_dom = extract_lp_idn_domestic(str(root / runtime_files["lp_idn_domestic"]), data_date)
    loading_lines = combine_loading_plan_lines(lp_ks, lp_idn_ex, lp_idn_dom)
    loading_plan = aggregate_loading_plan_for_master(loading_lines, data_date)
    print(
        f"  Loading Plan Lines: {len(loading_lines)} rows "
        f"(KS:{len(lp_ks)}, IDN_Ex:{len(lp_idn_ex)}, IDN_Dom:{len(lp_idn_dom)})"
    )
    print(f"  Loading Plan SO-level: {len(loading_plan)} SOs")

    # --- Phase 2: Join ---
    print("[6/6] Building SO Master & calculating gap...")
    master = build_so_master(sc, shipped, fg, pp_sched_agg, pp_unsched_agg, loading_plan)
    master = assign_status_and_gap(master)
    print(f"  Master: {len(master)} SOs")
    lp_outputs = build_loading_plan_analysis(
        sc, master, loading_lines, shipped, fg, pp_sched_agg, pp_unsched_agg, data_date
    )
    print(
        "  LP analysis: "
        f"{len(lp_outputs['reconciliation'])} reconciliation rows, "
        f"{len(lp_outputs['shipping_readiness'])} readiness rows"
    )
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
        lp_outputs,
        run_stamp,
    )
    print(f"  Excel: {excel_path}")

    dashboard_path = write_risk_dashboard(master, lp_outputs, str(output_dir), month, target_mt, data_date, run_stamp)
    print(f"  Dashboard: {dashboard_path}")

    # Quick summary
    print()
    print("=== Summary ===")
    print(f"  Baseline:  {master['sc_vol_mt'].sum():>10,.0f} MT ({len(master)} SOs)")
    print(f"  Shipped:   {master['allocated_shipped_mt'].sum():>10,.0f} MT")
    print(f"  In Stock:  {master['allocated_fg_mt'].sum():>10,.0f} MT")
    print(f"  WIP:       {master['allocated_wip_mt'].sum():>10,.0f} MT")
    print(f"  Scheduling:{(master['allocated_unsched_mt'].sum() + master['allocated_no_plan_mt'].sum()):>10,.0f} MT")
    print()
    risk_counts = master["risk_tier"].value_counts()
    for tier in ["Green", "Yellow", "Orange", "Red", "Critical"]:
        count = risk_counts.get(tier, 0)
        if count > 0:
            print(f"  {tier:10s}: {count} SOs")
    print()
    print("Done.")


def _resolve_runtime_inputs(root: Path, cfg: dict, month: str) -> tuple[dict, str]:
    """Resolve latest source files and derive data_date from latest 02-Shipped suffix."""
    files = dict(cfg["files"])

    shipped_rel, data_date = _latest_shipped_file(root, files["shipped"], month)
    files["shipped"] = shipped_rel

    files["sc"] = _latest_file_in_same_folder(root, files["sc"], "Order tracking*.xlsx", month)
    files["fg"] = _latest_file_in_same_folder(root, files["fg"], "FG stock*.xlsx", month)

    return files, data_date


def _latest_shipped_file(root: Path, configured_rel: str, month: str) -> tuple[str, str]:
    folder = (root / configured_rel).parent
    candidates = []
    for path in folder.glob("*.xlsx"):
        if path.name.startswith("~$"):
            continue
        parsed = _parse_shipped_end_date(path.name, month)
        if parsed is None:
            continue
        candidates.append((parsed, path))

    if not candidates:
        fallback = _fallback_data_date(month)
        return configured_rel, fallback

    end_date, path = max(candidates, key=lambda item: item[0])
    return _relative_to_root(path, root), end_date.strftime("%Y-%m-%d")


def _latest_file_in_same_folder(root: Path, configured_rel: str, pattern: str, month: str) -> str:
    folder = (root / configured_rel).parent
    candidates = []
    for path in folder.glob(pattern):
        if path.name.startswith("~$"):
            continue
        parsed = _parse_mmdd_date(path.name, month)
        sort_key = parsed or datetime.fromtimestamp(path.stat().st_mtime)
        candidates.append((sort_key, path))
    if not candidates:
        return configured_rel
    _, path = max(candidates, key=lambda item: item[0])
    return _relative_to_root(path, root)


def _parse_shipped_end_date(filename: str, month: str):
    match = re.search(r"(\d{6})-(\d{4})", filename)
    if not match:
        return _parse_mmdd_date(filename, month)
    start_yymmdd = match.group(1)
    end_mmdd = match.group(2)
    year = 2000 + int(start_yymmdd[:2])
    return datetime(year, int(end_mmdd[:2]), int(end_mmdd[2:]))


def _parse_mmdd_date(filename: str, month: str):
    matches = re.findall(r"(?<!\d)(\d{4})(?!\d)", filename)
    if not matches:
        return None
    year = int(month.split("-")[0])
    mmdd = matches[-1]
    try:
        return datetime(year, int(mmdd[:2]), int(mmdd[2:]))
    except ValueError:
        return None


def _relative_to_root(path: Path, root: Path) -> str:
    return str(path.relative_to(root)).replace("\\", "/")


def _fallback_data_date(month: str) -> str:
    year, mon = month.split("-")
    return f"{year}-{mon}-01"


if __name__ == "__main__":
    main()
