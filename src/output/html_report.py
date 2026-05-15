"""
HTML Narrative Report - Single-page story telling the monthly S&OE status.
"""
import pandas as pd
from pathlib import Path
from datetime import datetime


def write_html_report(
    master: pd.DataFrame,
    output_dir: str,
    month: str,
    target_mt: float,
    data_date: str,
    version_suffix: str = "",
):
    """
    Generate an HTML narrative report.

    Args:
        master: Final SO master DataFrame
        output_dir: Output directory
        month: Month string
        target_mt: Monthly target MT
        data_date: Data extraction date
    """
    Path(output_dir).mkdir(parents=True, exist_ok=True)
    suffix = f"-{version_suffix}" if version_suffix else ""
    output_path = Path(output_dir) / f"SOE_Report_{month}{suffix}.html"

    # Compute metrics
    metrics = _compute_metrics(master, target_mt)
    risk_table = _build_risk_table(master)
    top_risks = _build_top_risks(master)
    plant_breakdown = _build_plant_breakdown(master)

    html = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>S&OE Monthly Report - {month}</title>
    <style>
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{ font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; background: #f5f7fa; color: #1a1a2e; line-height: 1.6; padding: 40px 20px; }}
        .container {{ max-width: 1000px; margin: 0 auto; }}
        .header {{ text-align: center; margin-bottom: 40px; }}
        .header h1 {{ font-size: 28px; color: #1a1a2e; margin-bottom: 8px; }}
        .header .meta {{ color: #666; font-size: 14px; }}
        .section {{ background: white; border-radius: 12px; padding: 32px; margin-bottom: 24px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); }}
        .section h2 {{ font-size: 20px; color: #1a1a2e; margin-bottom: 16px; border-bottom: 2px solid #e8ecf0; padding-bottom: 8px; }}
        .kpi-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(160px, 1fr)); gap: 16px; margin-bottom: 20px; }}
        .kpi {{ text-align: center; padding: 16px; border-radius: 8px; background: #f8f9fc; }}
        .kpi .value {{ font-size: 28px; font-weight: 700; }}
        .kpi .label {{ font-size: 12px; color: #666; margin-top: 4px; }}
        .kpi.green .value {{ color: #22c55e; }}
        .kpi.yellow .value {{ color: #f59e0b; }}
        .kpi.red .value {{ color: #ef4444; }}
        .kpi.blue .value {{ color: #3b82f6; }}
        .waterfall {{ display: flex; height: 40px; border-radius: 8px; overflow: hidden; margin: 16px 0; }}
        .waterfall .seg {{ display: flex; align-items: center; justify-content: center; color: white; font-size: 12px; font-weight: 600; }}
        .seg-shipped {{ background: #22c55e; }}
        .seg-fg {{ background: #84cc16; }}
        .seg-wip {{ background: #f59e0b; }}
        .seg-unsched {{ background: #f97316; }}
        .seg-noplan {{ background: #ef4444; }}
        .seg-gap {{ background: #e5e7eb; color: #666 !important; }}
        table {{ width: 100%; border-collapse: collapse; font-size: 14px; }}
        th, td {{ padding: 10px 12px; text-align: left; border-bottom: 1px solid #e8ecf0; }}
        th {{ background: #f8f9fc; font-weight: 600; color: #555; }}
        .risk-green {{ color: #22c55e; font-weight: 600; }}
        .risk-yellow {{ color: #f59e0b; font-weight: 600; }}
        .risk-orange {{ color: #f97316; font-weight: 600; }}
        .risk-red {{ color: #ef4444; font-weight: 600; }}
        .risk-critical {{ color: #dc2626; font-weight: 700; }}
        .footer {{ text-align: center; color: #999; font-size: 12px; margin-top: 40px; }}
    </style>
</head>
<body>
<div class="container">
    <div class="header">
        <h1>S&OE Order Fulfillment Report</h1>
        <div class="meta">{month} | Data as of {data_date} | Generated {datetime.now().strftime('%Y-%m-%d %H:%M')}</div>
    </div>

    <!-- Section 1: Overview -->
    <div class="section">
        <h2>1. Monthly Overview</h2>
        <div class="kpi-grid">
            <div class="kpi blue"><div class="value">{metrics['target']:,.0f}</div><div class="label">Target (MT)</div></div>
            <div class="kpi blue"><div class="value">{metrics['baseline']:,.0f}</div><div class="label">Baseline Orders (MT)</div></div>
            <div class="kpi green"><div class="value">{metrics['shipped']:,.0f}</div><div class="label">Shipped (MT)</div></div>
            <div class="kpi green"><div class="value">{metrics['fg']:,.0f}</div><div class="label">In Stock (MT)</div></div>
            <div class="kpi yellow"><div class="value">{metrics['wip']:,.0f}</div><div class="label">In Production (MT)</div></div>
            <div class="kpi red"><div class="value">{metrics['at_risk']:,.0f}</div><div class="label">Scheduling (MT)</div></div>
        </div>
        <!-- Waterfall bar -->
        <div class="waterfall">
            {_waterfall_segments(metrics)}
        </div>
        <div style="display:flex; gap:16px; font-size:12px; color:#666; justify-content:center; margin-top:8px;">
            <span><span style="color:#22c55e;">&#9632;</span> Shipped</span>
            <span><span style="color:#84cc16;">&#9632;</span> In Stock</span>
            <span><span style="color:#f59e0b;">&#9632;</span> In Production</span>
            <span><span style="color:#f97316;">&#9632;</span> Unscheduled</span>
            <span><span style="color:#ef4444;">&#9632;</span> No Plan</span>
        </div>
    </div>

    <!-- Section 2: Risk Distribution -->
    <div class="section">
        <h2>2. Risk Distribution</h2>
        {risk_table}
    </div>

    <!-- Section 3: Plant Breakdown -->
    <div class="section">
        <h2>3. By Plant</h2>
        {plant_breakdown}
    </div>

    <!-- Section 4: Top Risks -->
    <div class="section">
        <h2>4. Top Risk Items (Action Required)</h2>
        {top_risks}
    </div>

    <div class="footer">
        S&OE Tracking System | Auto-generated monthly report
    </div>
</div>
</body>
</html>"""

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)

    return str(output_path)


def _compute_metrics(master: pd.DataFrame, target_mt: float) -> dict:
    return {
        "target": target_mt,
        "baseline": master["sc_vol_mt"].sum(),
        "shipped": master["shipped_mt"].sum(),
        "fg": master["fg_mt"].sum(),
        "wip": master["wip_mt"].sum(),
        "unsched": master["unsched_mt"].sum(),
        "noplan": master["no_plan_mt"].sum(),
        "at_risk": master["unsched_mt"].sum() + master["no_plan_mt"].sum(),
        "so_count": len(master),
    }


def _waterfall_segments(metrics: dict) -> str:
    total = metrics["baseline"] if metrics["baseline"] > 0 else 1
    segments = [
        ("seg-shipped", metrics["shipped"], "Shipped"),
        ("seg-fg", metrics["fg"], "FG"),
        ("seg-wip", metrics["wip"], "WIP"),
        ("seg-unsched", metrics["unsched"], "Unsched"),
        ("seg-noplan", metrics["noplan"], "No Plan"),
    ]
    html = ""
    for cls, val, label in segments:
        pct = (val / total) * 100
        if pct > 0:
            display = f"{val:,.0f}" if pct > 8 else ""
            html += f'<div class="seg {cls}" style="width:{pct:.1f}%">{display}</div>'
    return html


def _build_risk_table(master: pd.DataFrame) -> str:
    risk_order = ["Green", "Yellow", "Orange", "Red", "Critical"]
    rows = ""
    for tier in risk_order:
        subset = master[master["risk_tier"] == tier]
        count = len(subset)
        mt = subset["sc_vol_mt"].sum()
        css_class = f"risk-{tier.lower()}"
        rows += f'<tr><td class="{css_class}">{tier}</td><td>{count}</td><td>{mt:,.1f} MT</td></tr>'

    return f"""<table>
        <thead><tr><th>Risk Tier</th><th>SO Count</th><th>Volume</th></tr></thead>
        <tbody>{rows}</tbody>
    </table>"""


def _build_plant_breakdown(master: pd.DataFrame) -> str:
    plants = master.groupby("plant").agg(
        so_count=("so", "count"),
        baseline_mt=("sc_vol_mt", "sum"),
        shipped_mt=("shipped_mt", "sum"),
        fg_mt=("fg_mt", "sum"),
        wip_mt=("wip_mt", "sum"),
        risk_red=("risk_tier", lambda x: ((x == "Red") | (x == "Critical")).sum())
    ).reset_index()

    rows = ""
    for _, r in plants.iterrows():
        rows += f"""<tr>
            <td><strong>{r['plant']}</strong></td>
            <td>{r['so_count']}</td>
            <td>{r['baseline_mt']:,.0f}</td>
            <td>{r['shipped_mt']:,.0f}</td>
            <td>{r['fg_mt']:,.0f}</td>
            <td>{r['wip_mt']:,.0f}</td>
            <td class="risk-red">{r['risk_red']}</td>
        </tr>"""

    return f"""<table>
        <thead><tr><th>Plant</th><th>SOs</th><th>Baseline MT</th><th>Shipped</th><th>FG</th><th>WIP</th><th>Red/Critical</th></tr></thead>
        <tbody>{rows}</tbody>
    </table>"""


def _build_top_risks(master: pd.DataFrame) -> str:
    # Show top 20 highest-risk SOs
    risk_priority = {"Critical": 0, "Red": 1, "Orange": 2, "Yellow": 3, "Green": 4}
    top = master.copy()
    top["_priority"] = top["risk_tier"].map(risk_priority).fillna(5)
    top = top.sort_values(["_priority", "gap_days"], ascending=[True, True]).head(20)

    if top.empty:
        return "<p>No risk items found.</p>"

    rows = ""
    for _, r in top.iterrows():
        tier = r["risk_tier"]
        css_class = f"risk-{tier.lower()}"
        gap_str = f"{r['gap_days']:.0f}d" if pd.notna(r.get("gap_days")) else "-"
        rows += f"""<tr>
            <td>{r['so']}</td>
            <td>{r['plant']}</td>
            <td>{r.get('cluster', '-')}</td>
            <td>{r['sc_vol_mt']:,.1f}</td>
            <td>{r['status']}</td>
            <td class="{css_class}">{tier}</td>
            <td>{gap_str}</td>
        </tr>"""

    return f"""<table>
        <thead><tr><th>SO</th><th>Plant</th><th>Cluster</th><th>Volume MT</th><th>Status</th><th>Risk</th><th>Gap</th></tr></thead>
        <tbody>{rows}</tbody>
    </table>"""
