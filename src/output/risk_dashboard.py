"""
Interactive HTML dashboard for the S&OE executive summary and risk matrix.
"""
import json
from datetime import datetime
from pathlib import Path

import pandas as pd


def write_risk_dashboard(
    master: pd.DataFrame,
    lp_outputs: dict,
    output_dir: str,
    month: str,
    target_mt: float,
    data_date: str,
    version_suffix: str = "",
) -> str:
    """Write a standalone interactive dashboard HTML file."""
    Path(output_dir).mkdir(parents=True, exist_ok=True)
    suffix = f"-{version_suffix}" if version_suffix else ""
    output_path = Path(output_dir) / f"SOE_Dashboard_{month}{suffix}.html"

    payload = _build_payload(master, lp_outputs or {}, month, target_mt, data_date)
    html = _render_html(payload)
    output_path.write_text(html, encoding="utf-8")
    return str(output_path)


def _build_payload(master: pd.DataFrame, lp_outputs: dict, month: str, target_mt: float, data_date: str) -> dict:
    risk = lp_outputs.get("risk_matrix_detail", pd.DataFrame())
    if risk is None:
        risk = pd.DataFrame()
    lp_only = lp_outputs.get("lp_not_in_baseline_detail", pd.DataFrame())
    if lp_only is None:
        lp_only = pd.DataFrame()

    master_records = _records(master)
    risk_records = _records(risk)
    lp_only_records = _records(lp_only)
    plants = sorted([p for p in master["plant"].dropna().astype(str).unique().tolist() if p])
    return {
        "month": month,
        "targetMt": float(target_mt),
        "dataDate": data_date,
        "generatedAt": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "plants": plants,
        "master": master_records,
        "riskDetail": risk_records,
        "lpOnlyDetail": lp_only_records,
    }


def _records(df: pd.DataFrame) -> list[dict]:
    if df is None or df.empty:
        return []
    clean = df.copy()
    for col in clean.columns:
        if pd.api.types.is_datetime64_any_dtype(clean[col]):
            clean[col] = _format_datetime_series(clean[col])
    clean = clean.replace({pd.NA: None})
    clean = clean.where(pd.notna(clean), None)
    records = clean.to_dict(orient="records")
    return [{key: _json_value(value) for key, value in record.items()} for record in records]


def _json_value(value):
    if isinstance(value, pd.Timestamp):
        if pd.isna(value):
            return None
        if value.hour or value.minute or value.second:
            return value.strftime("%Y-%m-%d %H:%M")
        return value.strftime("%Y-%m-%d")
    if hasattr(value, "item"):
        try:
            return value.item()
        except (ValueError, TypeError):
            pass
    return value


def _format_datetime_series(series: pd.Series) -> pd.Series:
    non_blank = series.dropna()
    has_time = False
    if not non_blank.empty:
        has_time = bool(
            (
                (non_blank.dt.hour != 0)
                | (non_blank.dt.minute != 0)
                | (non_blank.dt.second != 0)
            ).any()
        )
    return series.dt.strftime("%Y-%m-%d %H:%M" if has_time else "%Y-%m-%d")


def _render_html(payload: dict) -> str:
    data = json.dumps(payload, ensure_ascii=False)
    return f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>S&OE Fulfillment Dashboard</title>
  <style>
    :root {{
      --navy: #16233d;
      --blue: #2f6f85;
      --blue-dark: #234f63;
      --gold: #c6a05c;
      --line: #d7d0c0;
      --text: #17233d;
      --muted: #6f6b61;
      --panel: #f5f1e8;
      --row: #f8f5ed;
      --row-alt: #efede6;
      --white: #ffffff;
      --green: #146b45;
      --amber: #9a6a00;
      --red: #b23a32;
      --shadow: 0 10px 28px rgba(22, 35, 61, 0.10);
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      font-family: "Segoe UI", Arial, sans-serif;
      background: #eee9df;
      color: var(--text);
    }}
    .page {{
      max-width: 1320px;
      margin: 0 auto;
      background: #fbf8f1;
      min-height: 100vh;
      box-shadow: var(--shadow);
    }}
    .topbar {{
      background: var(--navy);
      color: var(--white);
      text-align: center;
      padding: 14px 20px 10px;
    }}
    .topbar h1 {{
      margin: 0;
      font-size: 20px;
      letter-spacing: 0;
      font-weight: 700;
    }}
    .topbar .meta {{
      margin-top: 12px;
      color: #d8d1c2;
      font-size: 12px;
    }}
    .controls {{
      display: flex;
      justify-content: space-between;
      align-items: center;
      gap: 16px;
      padding: 14px 18px;
      background: #f2eee5;
      border-bottom: 1px solid var(--line);
    }}
    .control-group {{
      display: flex;
      align-items: center;
      gap: 10px;
      font-size: 13px;
      color: var(--muted);
    }}
    select, button {{
      height: 32px;
      border: 1px solid var(--line);
      border-radius: 4px;
      background: var(--white);
      color: var(--text);
      padding: 0 10px;
      font: inherit;
    }}
    button {{
      cursor: pointer;
      font-weight: 600;
    }}
    button:hover, select:hover {{ border-color: var(--blue); }}
    .content {{ padding: 16px 18px 28px; }}
    .kpis {{
      display: grid;
      grid-template-columns: repeat(5, minmax(150px, 1fr));
      gap: 0;
      background: var(--panel);
      border-bottom: 1px solid var(--line);
      margin-bottom: 20px;
    }}
    .kpi {{
      text-align: center;
      padding: 10px 8px 13px;
      border-top: 3px solid var(--navy);
      min-height: 88px;
    }}
    .kpi:nth-child(2), .kpi:nth-child(3) {{ border-top-color: var(--green); }}
    .kpi:nth-child(4) {{ border-top-color: var(--gold); }}
    .kpi:nth-child(5) {{ border-top-color: var(--red); }}
    .kpi .label {{
      font-size: 10px;
      color: var(--muted);
      text-transform: uppercase;
      margin: 0 0 12px;
    }}
    .kpi .value {{
      font-size: 22px;
      font-weight: 750;
      color: var(--navy);
    }}
    .kpi:nth-child(2) .value, .kpi:nth-child(3) .value {{ color: var(--green); }}
    .kpi:nth-child(4) .value {{ color: var(--amber); }}
    .kpi:nth-child(5) .value {{ color: var(--red); }}
    .kpi .sub {{
      margin-top: 8px;
      font-size: 11px;
      color: var(--muted);
    }}
    .grid-two {{
      display: grid;
      grid-template-columns: minmax(420px, 0.9fr) minmax(460px, 1fr);
      gap: 18px;
      align-items: start;
      margin-bottom: 24px;
    }}
    .section-title {{
      background: var(--navy);
      color: var(--white);
      font-weight: 700;
      font-size: 12px;
      padding: 7px 8px;
      text-transform: uppercase;
    }}
    .table-wrap {{
      border: 1px solid var(--line);
      background: var(--white);
      overflow: auto;
    }}
    table {{
      width: 100%;
      border-collapse: collapse;
      table-layout: fixed;
      font-size: 12px;
    }}
    th {{
      background: var(--blue-dark);
      color: var(--white);
      font-weight: 700;
      padding: 8px 6px;
      border: 1px solid var(--line);
      text-align: center;
    }}
    td {{
      padding: 9px 7px;
      border: 1px solid var(--line);
      text-align: center;
      background: var(--row);
      height: 36px;
    }}
    tr:nth-child(even) td {{ background: var(--row-alt); }}
    tr.total-row td {{
      background: #eadfc7 !important;
      font-weight: 750;
    }}
    td.row-head {{
      text-align: left;
      font-weight: 700;
      color: var(--text);
    }}
    .clickable {{
      cursor: pointer;
      color: var(--navy);
      font-weight: 700;
      position: relative;
    }}
    .clickable:hover {{
      background: #eadfc7 !important;
      outline: 2px solid var(--gold);
      outline-offset: -2px;
    }}
    .clickable.active {{
      background: #e5d2a6 !important;
      outline: 2px solid var(--gold);
      outline-offset: -2px;
    }}
    .matrix-area {{ margin-top: 2px; }}
    .detail-panel {{
      margin-top: 20px;
      border: 1px solid var(--line);
      background: var(--white);
    }}
    .detail-panel .table-wrap {{
      max-height: 520px;
      overflow: auto;
    }}
    .detail-head {{
      display: flex;
      justify-content: space-between;
      align-items: center;
      gap: 12px;
      background: #f2eee5;
      border-bottom: 1px solid var(--line);
      padding: 10px 12px;
    }}
    .detail-head h2 {{
      margin: 0;
      font-size: 15px;
      color: var(--navy);
    }}
    .detail-actions {{
      display: flex;
      align-items: center;
      justify-content: flex-end;
      gap: 10px;
      min-width: 360px;
    }}
    .detail-head .hint {{
      font-size: 12px;
      color: var(--muted);
      text-align: right;
    }}
    .download-button {{
      min-width: 122px;
      background: var(--navy);
      color: var(--white);
      border-color: var(--navy);
    }}
    .download-button:hover {{
      background: var(--blue);
      border-color: var(--blue);
    }}
    .download-button:disabled {{
      cursor: not-allowed;
      opacity: 0.48;
    }}
    .detail-table th, .detail-table td {{
      font-size: 11px;
      text-align: left;
      white-space: nowrap;
    }}
    .detail-table {{
      table-layout: auto;
      min-width: 1120px;
    }}
    .detail-table th {{
      position: sticky;
      top: 0;
      z-index: 1;
    }}
    .detail-table td.num, .detail-table th.num {{ text-align: right; }}
    .detail-table td.note {{
      white-space: normal;
      min-width: 280px;
    }}
    .detail-table td.gap-neg {{
      background: #f4cccc !important;
      color: #9f1d1d;
      font-weight: 750;
    }}
    .detail-table td.gap-tight {{
      background: #fff2cc !important;
      color: #7a4f00;
      font-weight: 750;
    }}
    .detail-table td.gap-ok {{
      background: #d9ead3 !important;
      color: #1f6b32;
      font-weight: 750;
    }}
    .detail-table td.gap-blank {{
      color: var(--muted);
    }}
    .empty {{
      padding: 20px;
      color: var(--muted);
      font-size: 13px;
    }}
    .footer-note {{
      margin-top: 12px;
      color: var(--muted);
      font-size: 11px;
      line-height: 1.5;
    }}
    @media (max-width: 900px) {{
      .kpis {{ grid-template-columns: 1fr 1fr; }}
      .grid-two {{ grid-template-columns: 1fr; }}
      .controls {{ align-items: flex-start; flex-direction: column; }}
      .detail-head {{ align-items: flex-start; flex-direction: column; }}
      .detail-actions {{ align-items: flex-start; flex-direction: column; min-width: 0; }}
      .detail-head .hint {{ text-align: left; }}
    }}
  </style>
</head>
<body>
  <div class="page">
    <header class="topbar">
      <h1>S&amp;OE ORDER FULFILLMENT TRACKING&nbsp;&nbsp; | &nbsp;&nbsp;<span id="month"></span></h1>
      <div class="meta" id="meta"></div>
    </header>
    <div class="controls">
      <div class="control-group">
        <label for="plantFilter">Plant</label>
        <select id="plantFilter"></select>
        <button id="resetSelection" type="button">Clear selection</button>
      </div>
      <div class="control-group" id="selectionText">Click a matrix cell to inspect detail.</div>
    </div>
    <main class="content">
      <section class="kpis" id="kpis"></section>
      <section class="grid-two">
        <div>
          <div class="section-title">By Plant</div>
          <div class="table-wrap"><table id="plantTable"></table></div>
        </div>
        <div>
          <div class="section-title">Shipped Data Closure (MT)</div>
          <div class="table-wrap"><table id="closureTable"></table></div>
        </div>
      </section>
      <section class="matrix-area">
        <div class="section-title">Open Fulfillment Risk Matrix (MT)</div>
        <div class="table-wrap"><table id="riskMatrix"></table></div>
      </section>
      <section class="matrix-area">
        <div class="section-title">LP Not In Current SC Baseline (MT)</div>
        <div class="table-wrap"><table id="lpOnlyTable"></table></div>
      </section>
      <section class="detail-panel">
        <div class="detail-head">
          <h2 id="detailTitle">Risk Matrix Detail</h2>
          <div class="detail-actions">
            <div class="hint" id="detailHint">默认展示全部 action-required 明细；点击矩阵格子后展示对应明细。</div>
            <button class="download-button" id="downloadDetail" type="button">Download CSV</button>
          </div>
        </div>
        <div class="table-wrap"><table class="detail-table" id="detailTable"></table></div>
      </section>
      <div class="footer-note">
        口径说明：Shipped 单独作为数据闭环审计；Open Risk Matrix 只看本月 baseline 未发货量。LP Not In Current SC Baseline 是 Loading Plan-led 对账异常，不混入履约风险矩阵。
      </div>
    </main>
  </div>
  <script>
    const DATA = {data};
    const SUPPLY_ORDER = ["FG", "WIP Scheduled", "WIP Unscheduled", "No Supply Signal"];
    const LP_ORDER = ["Past Due LP", "Future Valid LP", "LP Date Unconfirmed", "No Current LP"];
    const LP_ONLY_ORDER = ["Past Due LP", "Future Valid LP", "LP Date Unconfirmed"];
    const BASE_DETAIL_COLUMNS = [
      ["so", "SO"],
      ["plant", "Plant"],
      ["cluster", "Cluster"],
      ["order_type", "Order Type"],
      ["supply_status", "Supply"],
      ["lp_coverage_status", "LP Coverage"],
      ["risk_mt", "Risk MT"]
    ];
    const WIP_PRODUCTION_CONTEXT_COLUMNS = [
      ["machines", "Machine"],
      ["planned_end_date", "Planned End"],
      ["available_date", "Available Date"]
    ];
    const WORK_ORDER_DETAIL_COLUMNS = [
      ["work_orders", "Work Order"]
    ];
    const VALID_LP_DETAIL_COLUMNS = [
      ["loading_date", "Loading Date"]
    ];
    const UNCONFIRMED_LP_DETAIL_COLUMNS = [
      ["lp_loading_date_raw_list", "LP Date Raw"],
      ["lp_loading_date_status_mix", "LP Date Status"]
    ];
    const GAP_DETAIL_COLUMNS = [
      ["lp_gap_days", "Gap"]
    ];
    const ACTION_LEADING_COLUMNS = [
      ["risk_action", "Risk Action"],
      ["suggested_owner", "Owner"]
    ];
    const ACTION_NOTE_COLUMNS = [
      ["action_note", "Action Note"]
    ];
    const LP_ONLY_DETAIL_COLUMNS = [
      ["so", "SO"],
      ["plant", "Plant"],
      ["lp_source", "LP Source"],
      ["source_sheet", "Source Sheet"],
      ["invoice_no_raw", "Invoice No Raw"],
      ["loading_date_raw", "Loading Date Raw"],
      ["loading_date", "Loading Date"],
      ["loading_date_status", "LP Date Status"],
      ["load_mt", "Load MT"]
    ];
    let selected = null;

    const fmt = value => Number(value || 0).toLocaleString("en-US", {{ maximumFractionDigits: 0 }});
    const fmt1 = value => Number(value || 0).toLocaleString("en-US", {{ minimumFractionDigits: 1, maximumFractionDigits: 1 }});
    const samePlant = row => {{
      const plant = document.getElementById("plantFilter").value;
      return plant === "ALL" || String(row.plant || "") === plant;
    }};
    const master = () => DATA.master.filter(samePlant);
    const risk = () => DATA.riskDetail.filter(samePlant);
    const lpOnly = () => DATA.lpOnlyDetail.filter(samePlant);
    const sum = (rows, field) => rows.reduce((acc, row) => acc + Number(row[field] || 0), 0);

    function init() {{
      document.getElementById("month").textContent = DATA.month;
      document.getElementById("meta").textContent = `Data as of ${{DATA.dataDate}}   |   Target: ${{fmt(DATA.targetMt)}} MT/month   |   Generated: ${{DATA.generatedAt}}`;
      const select = document.getElementById("plantFilter");
      select.innerHTML = `<option value="ALL">All Plants</option>` + DATA.plants.map(p => `<option value="${{p}}">${{p}}</option>`).join("");
      select.addEventListener("change", () => {{ selected = null; renderAll(); }});
      document.getElementById("resetSelection").addEventListener("click", () => {{ selected = null; renderAll(); }});
      document.getElementById("downloadDetail").addEventListener("click", downloadSelectedDetail);
      renderAll();
    }}

    function renderAll() {{
      renderKpis();
      renderPlantTable();
      renderClosureTable();
      renderRiskMatrix();
      renderLpOnlyTable();
      renderDetailTable();
      renderSelectionText();
    }}

    function renderKpis() {{
      const rows = master();
      const values = [
        ["BASELINE ORDERS", sum(rows, "sc_vol_mt"), "Adjusted baseline"],
        ["SHIPPED", sum(rows, "allocated_shipped_mt"), "Allocated shipped"],
        ["IN STOCK (FG)", sum(rows, "allocated_fg_mt"), "Allocated FG"],
        ["IN PRODUCTION", sum(rows, "allocated_wip_mt"), "Allocated WIP"],
        ["SCHEDULING", sum(rows, "allocated_unsched_mt") + sum(rows, "allocated_no_plan_mt"), "Unscheduled + no plan"]
      ];
      document.getElementById("kpis").innerHTML = values.map(([label, value, sub]) => `
        <div class="kpi">
          <div class="label">${{label}}</div>
          <div class="value">${{fmt(value)}} MT</div>
          <div class="sub">${{sub}}</div>
        </div>
      `).join("");
    }}

    function renderPlantTable() {{
      const selectedPlant = document.getElementById("plantFilter").value;
      const plants = selectedPlant === "ALL" ? DATA.plants : DATA.plants.filter(p => p === selectedPlant);
      const rows = plants.map(plant => {{
        const subset = DATA.master.filter(r => String(r.plant || "") === plant);
        return {{
          plant,
          baseline: sum(subset, "sc_vol_mt"),
          shipped: sum(subset, "allocated_shipped_mt"),
          fg: sum(subset, "allocated_fg_mt"),
          wip: sum(subset, "allocated_wip_mt"),
          unscheduled: sum(subset, "allocated_unsched_mt"),
          noPlan: sum(subset, "allocated_no_plan_mt")
        }};
      }});
      const total = {{
        plant: "Total",
        baseline: sum(rows, "baseline"),
        shipped: sum(rows, "shipped"),
        fg: sum(rows, "fg"),
        wip: sum(rows, "wip"),
        unscheduled: sum(rows, "unscheduled"),
        noPlan: sum(rows, "noPlan")
      }};
      const displayRows = [...rows, total];
      document.getElementById("plantTable").innerHTML = `
        <thead><tr><th>Plant</th><th>Baseline MT</th><th>Shipped</th><th>FG</th><th>WIP</th><th>Unscheduled</th><th>No Plan</th></tr></thead>
        <tbody>${{displayRows.map(r => `
          <tr class="${{r.plant === "Total" ? "total-row" : ""}}">
            <td class="row-head">${{r.plant}}</td>
            <td>${{fmt(r.baseline)}}</td>
            <td>${{fmt(r.shipped)}}</td>
            <td>${{fmt(r.fg)}}</td>
            <td>${{fmt(r.wip)}}</td>
            <td>${{fmt(r.unscheduled)}}</td>
            <td>${{fmt(r.noPlan)}}</td>
          </tr>`).join("")}}</tbody>`;
    }}

    function renderClosureTable() {{
      const rows = risk().filter(r => r.supply_status === "Shipped");
      const total = sum(rows, "risk_mt");
      const closed = sum(rows.filter(r => r.lp_coverage_status === "Shipped LP Closed"), "risk_mt");
      const missing = total - closed;
      document.getElementById("closureTable").innerHTML = `
        <thead><tr><th>Supply Status</th><th>Total MT</th><th>LP Closed / Matched</th><th>LP Missing or Unclear</th></tr></thead>
        <tbody><tr>
          <td class="row-head">Shipped</td>
          <td class="clickable" data-type="closure" data-status="ALL">${{fmt(total)}}</td>
          <td class="clickable" data-type="closure" data-status="Shipped LP Closed">${{fmt(closed)}}</td>
          <td class="clickable" data-type="closure" data-status="missing">${{fmt(missing)}}</td>
        </tr></tbody>`;
      document.querySelectorAll("#closureTable .clickable").forEach(cell => {{
        cell.addEventListener("click", () => {{
          selected = {{ kind: "closure", status: cell.dataset.status }};
          renderAll();
        }});
      }});
    }}

    function renderRiskMatrix() {{
      const rows = risk().filter(r => r.supply_status !== "Shipped");
      const body = SUPPLY_ORDER.map(supply => {{
        const supplyRows = rows.filter(r => r.supply_status === supply);
        const cells = LP_ORDER.map(lp => {{
          const value = sum(supplyRows.filter(r => r.lp_coverage_status === lp), "risk_mt");
          const active = selected && selected.kind === "matrix" && selected.supply === supply && selected.lp === lp ? " active" : "";
          return `<td class="clickable${{active}}" data-supply="${{supply}}" data-lp="${{lp}}">${{fmt(value)}}</td>`;
        }}).join("");
        return `<tr><td class="row-head">${{supply}}</td><td>${{fmt(sum(supplyRows, "risk_mt"))}}</td>${{cells}}</tr>`;
      }});
      const totalRow = `<tr>
        <td class="row-head">Open Total</td>
        <td>${{fmt(sum(rows, "risk_mt"))}}</td>
        ${{LP_ORDER.map(lp => `<td>${{fmt(sum(rows.filter(r => r.lp_coverage_status === lp), "risk_mt"))}}</td>`).join("")}}
      </tr>`;
      document.getElementById("riskMatrix").innerHTML = `
        <thead><tr><th>Supply Status</th><th>Supply Status Total</th>${{LP_ORDER.map(lp => `<th>${{lp}}</th>`).join("")}}</tr></thead>
        <tbody>${{body.join("")}}${{totalRow}}</tbody>`;
      document.querySelectorAll("#riskMatrix .clickable").forEach(cell => {{
        cell.addEventListener("click", () => {{
          selected = {{ kind: "matrix", supply: cell.dataset.supply, lp: cell.dataset.lp }};
          renderAll();
        }});
      }});
    }}

    function renderLpOnlyTable() {{
      const rows = lpOnly();
      const body = LP_ONLY_ORDER.map(status => {{
        const value = sum(rows.filter(r => r.lp_coverage_status === status), "load_mt");
        const active = selected && selected.kind === "lpOnly" && selected.status === status ? " active" : "";
        return `<tr>
          <td class="row-head">${{status}}</td>
          <td class="clickable${{active}}" data-status="${{status}}">${{fmt(value)}}</td>
        </tr>`;
      }}).join("");
      const totalActive = selected && selected.kind === "lpOnly" && selected.status === "ALL" ? " active" : "";
      const totalRow = `<tr>
        <td class="row-head">Total</td>
        <td class="clickable${{totalActive}}" data-status="ALL">${{fmt(sum(rows, "load_mt"))}}</td>
      </tr>`;
      document.getElementById("lpOnlyTable").innerHTML = `
        <thead><tr><th>LP Status</th><th>MT</th></tr></thead>
        <tbody>${{body}}${{totalRow}}</tbody>`;
      document.querySelectorAll("#lpOnlyTable .clickable").forEach(cell => {{
        cell.addEventListener("click", () => {{
          selected = {{ kind: "lpOnly", status: cell.dataset.status }};
          renderAll();
        }});
      }});
    }}

    function selectedRows() {{
      let rows = risk();
      if (selected && selected.kind === "lpOnly") {{
        rows = lpOnly();
        if (selected.status !== "ALL") {{
          rows = rows.filter(r => r.lp_coverage_status === selected.status);
        }}
        return rows.sort((a, b) => Number(b.load_mt || 0) - Number(a.load_mt || 0));
      }}
      if (selected && selected.kind === "matrix") {{
        rows = rows.filter(r => r.supply_status === selected.supply && r.lp_coverage_status === selected.lp);
      }} else if (selected && selected.kind === "closure") {{
        rows = rows.filter(r => r.supply_status === "Shipped");
        if (selected.status === "Shipped LP Closed") {{
          rows = rows.filter(r => r.lp_coverage_status === "Shipped LP Closed");
        }} else if (selected.status === "missing") {{
          rows = rows.filter(r => r.lp_coverage_status !== "Shipped LP Closed");
        }}
      }} else {{
        rows = rows.filter(r => r.action_required === true || String(r.action_required).toLowerCase() === "true");
      }}
      return rows.sort((a, b) => Number(b.risk_mt || 0) - Number(a.risk_mt || 0));
    }}

    function detailColumns() {{
      if (selected && selected.kind === "lpOnly") {{
        return LP_ONLY_DETAIL_COLUMNS;
      }}
      const columns = [...BASE_DETAIL_COLUMNS];
      if (selected && selected.kind === "matrix") {{
        if (selected.supply === "WIP Scheduled") {{
          columns.push(...WIP_PRODUCTION_CONTEXT_COLUMNS);
        }}
        if (selected.lp === "Past Due LP" || selected.lp === "Future Valid LP") {{
          columns.push(...VALID_LP_DETAIL_COLUMNS);
          if (selected.supply === "WIP Scheduled") columns.push(...GAP_DETAIL_COLUMNS);
        }} else if (selected.lp === "LP Date Unconfirmed") {{
          columns.push(...UNCONFIRMED_LP_DETAIL_COLUMNS);
        }}
      }} else if (selected && selected.kind === "closure") {{
        columns.push(...VALID_LP_DETAIL_COLUMNS, ...UNCONFIRMED_LP_DETAIL_COLUMNS);
      }}
      columns.push(...ACTION_LEADING_COLUMNS);
      if (selected && selected.kind === "matrix" && selected.supply === "WIP Scheduled") {{
        columns.push(...WORK_ORDER_DETAIL_COLUMNS);
      }}
      columns.push(...ACTION_NOTE_COLUMNS);
      return columns;
    }}

    function cellClass(field, value) {{
      const classes = [];
      if (field === "risk_mt" || field === "lp_gap_days" || field === "load_mt") classes.push("num");
      if (field === "action_note") classes.push("note");
      if (field === "lp_gap_days") {{
        if (value === null || value === undefined || value === "") {{
          classes.push("gap-blank");
        }} else {{
          const num = Number(value);
          if (Number.isFinite(num) && num < 0) classes.push("gap-neg");
          else if (Number.isFinite(num) && num <= 2) classes.push("gap-tight");
          else if (Number.isFinite(num)) classes.push("gap-ok");
        }}
      }}
      return classes.join(" ");
    }}

    function displayValue(field, value) {{
      if (field === "risk_mt" || field === "load_mt") return fmt1(value);
      if (field === "lp_gap_days") {{
        if (value === null || value === undefined || value === "") return "";
        const num = Number(value);
        return Number.isFinite(num) ? num.toLocaleString("en-US", {{ maximumFractionDigits: 0 }}) : "";
      }}
      return value ?? "";
    }}

    function renderDetailTable() {{
      const rows = selectedRows();
      const shown = rows.slice(0, 300);
      const columns = detailColumns();
      const table = document.getElementById("detailTable");
      const downloadButton = document.getElementById("downloadDetail");
      downloadButton.disabled = rows.length === 0;
      table.style.minWidth = `${{Math.max(1040, columns.length * 115)}}px`;
      if (!shown.length) {{
        table.innerHTML = `<tbody><tr><td class="empty">No detail rows for the current selection.</td></tr></tbody>`;
        document.getElementById("detailHint").textContent = "0 rows, 0.0 MT.";
        return;
      }}
      const headers = columns.map(([field, label]) => `<th class="${{field === "risk_mt" || field === "lp_gap_days" || field === "load_mt" ? "num" : ""}}">${{label}}</th>`).join("");
      const body = shown.map(row => `<tr>${{columns.map(([field]) => {{
        let value = row[field];
        return `<td class="${{cellClass(field, value)}}">${{displayValue(field, value)}}</td>`;
      }}).join("")}}</tr>`).join("");
      table.innerHTML = `<thead><tr>${{headers}}</tr></thead><tbody>${{body}}</tbody>`;
      const mtField = selected && selected.kind === "lpOnly" ? "load_mt" : "risk_mt";
      const mtLabel = selected && selected.kind === "lpOnly" ? "Load MT" : "Risk MT";
      document.getElementById("detailHint").textContent = `${{fmt(rows.length)}} rows, ${{fmt1(sum(rows, mtField))}} MT. Showing top ${{fmt(shown.length)}} rows by ${{mtLabel}}.`;
    }}

    function csvEscape(value) {{
      const text = value === null || value === undefined ? "" : String(value);
      if (/[",\\r\\n]/.test(text)) return `"${{text.replace(/"/g, '""')}}"`;
      return text;
    }}

    function safeFilePart(value) {{
      return String(value || "ALL")
        .trim()
        .replace(/[^a-zA-Z0-9\u4e00-\u9fa5]+/g, "-")
        .replace(/^-+|-+$/g, "") || "ALL";
    }}

    function downloadSelectedDetail() {{
      const rows = selectedRows();
      if (!rows.length) return;
      const columns = detailColumns();
      const header = columns.map(([, label]) => csvEscape(label)).join(",");
      const lines = rows.map(row => columns.map(([field]) => {{
        return csvEscape(displayValue(field, row[field]));
      }}).join(","));
      const csv = "\uFEFF" + [header, ...lines].join("\\r\\n");
      const plant = document.getElementById("plantFilter").value;
      let scope = "Action-Required";
      if (selected && selected.kind === "matrix") scope = `${{selected.supply}}-${{selected.lp}}`;
      if (selected && selected.kind === "closure") scope = `Shipped-${{selected.status}}`;
      if (selected && selected.kind === "lpOnly") scope = `LP-Only-${{selected.status}}`;
      const filename = `risk_detail_${{safeFilePart(DATA.month)}}_${{safeFilePart(plant)}}_${{safeFilePart(scope)}}.csv`;
      const blob = new Blob([csv], {{ type: "text/csv;charset=utf-8;" }});
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = filename;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url);
    }}

    function renderSelectionText() {{
      const plant = document.getElementById("plantFilter").value;
      let text = `Plant: ${{plant === "ALL" ? "All" : plant}}`;
      if (selected && selected.kind === "matrix") text += ` | ${{selected.supply}} × ${{selected.lp}}`;
      if (selected && selected.kind === "closure") text += ` | Shipped closure: ${{selected.status}}`;
      if (selected && selected.kind === "lpOnly") text += ` | LP not in baseline: ${{selected.status}}`;
      document.getElementById("selectionText").textContent = text;
      let title = "Action Required Detail";
      if (selected && selected.kind === "lpOnly") title = "LP Not In Baseline Detail";
      else if (selected) title = "Selected Risk Detail";
      document.getElementById("detailTitle").textContent = title;
    }}

    init();
  </script>
</body>
</html>"""
