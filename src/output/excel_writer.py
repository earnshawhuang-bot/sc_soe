"""
Excel Writer - Professional consulting-style output with full formatting.
"""
import pandas as pd
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side,
    GradientFill
)
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.formatting.rule import CellIsRule, Rule
from openpyxl.styles.differential import DifferentialStyle
from datetime import datetime


# ── Color palette (McKinsey-inspired) ──────────────────────────────────────
C_NAVY       = "1F3864"   # Header background
C_DARK_BLUE  = "2E4D8A"   # Sub-header
C_MID_BLUE   = "4472C4"   # Accent
C_LIGHT_BLUE = "D6E4F7"   # Alternate row
C_WHITE      = "FFFFFF"
C_OFF_WHITE  = "F8F9FC"
C_LIGHT_GREY = "F2F2F2"
C_BORDER     = "BFC9D6"

C_GREEN      = "D6F0E0"   # Risk tier fills
C_GREEN_FG   = "1A7A3F"
C_YELLOW     = "FFF8D6"
C_YELLOW_FG  = "8A6800"
C_ORANGE     = "FDEBD0"
C_ORANGE_FG  = "B85C00"
C_RED        = "FADADD"
C_RED_FG     = "9B1C1C"
C_CRITICAL   = "E8B4B8"
C_CRITICAL_FG= "6B0000"


def _fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def _font(bold=False, size=10, color=C_NAVY, name="Calibri"):
    return Font(bold=bold, size=size, color=color, name=name)

def _border_thin():
    s = Side(style="thin", color=C_BORDER)
    return Border(left=s, right=s, top=s, bottom=s)

def _border_bottom():
    return Border(bottom=Side(style="medium", color=C_BORDER))

def _align(h="left", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

RISK_STYLE = {
    "Green":    (_fill(C_GREEN),    _font(bold=True, color=C_GREEN_FG)),
    "Yellow":   (_fill(C_YELLOW),   _font(bold=True, color=C_YELLOW_FG)),
    "Orange":   (_fill(C_ORANGE),   _font(bold=True, color=C_ORANGE_FG)),
    "Red":      (_fill(C_RED),      _font(bold=True, color=C_RED_FG)),
    "Critical": (_fill(C_CRITICAL), _font(bold=True, color=C_CRITICAL_FG)),
}


def write_excel(master: pd.DataFrame, output_dir: str, month: str, target_mt: float, data_date: str = ""):
    Path(output_dir).mkdir(parents=True, exist_ok=True)
    output_path = Path(output_dir) / f"SOE_Tracking_{month}.xlsx"

    wb = Workbook()
    wb.remove(wb.active)
    master.attrs["data_date"] = data_date

    _sheet_summary(wb, master, month, target_mt)
    _sheet_master(wb, master)
    _sheet_gap_detail(wb, master)
    _sheet_action(wb, master)

    wb.save(str(output_path))
    return str(output_path)


# ── Sheet 1: Executive Summary ─────────────────────────────────────────────
def _sheet_summary(wb, master, month, target_mt):
    ws = wb.create_sheet("Summary")
    ws.sheet_view.showGridLines = False

    # ── Title banner ──
    ws.merge_cells("A1:J1")
    ws.row_dimensions[1].height = 36
    c = ws["A1"]
    c.value = f"S&OE ORDER FULFILLMENT TRACKING  │  {month}"
    c.font = Font(bold=True, size=16, color=C_WHITE, name="Calibri")
    c.fill = _fill(C_NAVY)
    c.alignment = _align("center")

    ws.merge_cells("A2:J2")
    c = ws["A2"]
    c.value = f"Data as of {master.attrs.get('data_date', '')}   │   Target: {target_mt:,.0f} MT/month   │   Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}"
    c.font = Font(size=9, color="AABBCC", name="Calibri")
    c.fill = _fill(C_NAVY)
    c.alignment = _align("center")
    ws.row_dimensions[2].height = 18

    ws.row_dimensions[3].height = 8  # spacer

    # ── KPI cards (row 4-8) ──
    _kpi_header(ws, row=4)
    kpis = [
        ("BASELINE ORDERS", f"{master['sc_vol_mt'].sum():,.0f} MT", f"{len(master)} SOs", C_DARK_BLUE),
        ("SHIPPED",         f"{master['shipped_mt'].sum():,.0f} MT",
                            f"{(master['status']=='Shipped').sum() + (master['status']=='Partially Shipped').sum()} SOs", C_GREEN_FG),
        ("IN STOCK (FG)",   f"{master['fg_mt'].sum():,.0f} MT",
                            f"{(master['status']=='In Stock').sum()} SOs", "2E7D32"),
        ("IN PRODUCTION",   f"{master['wip_mt'].sum():,.0f} MT",
                            f"{(master['status']=='In Production').sum()} SOs", C_YELLOW_FG),
        ("AT RISK",         f"{(master['unsched_mt']+master['no_plan_mt']).sum():,.0f} MT",
                            f"{master['status'].isin(['Planned (Unscheduled)','No Plan']).sum()} SOs", C_RED_FG),
    ]
    col_starts = [1, 3, 5, 7, 9]
    for (label, val, sub, color), col in zip(kpis, col_starts):
        _kpi_card(ws, row=5, col=col, label=label, value=val, sub=sub, accent=color)

    ws.row_dimensions[10].height = 8  # spacer

    # ── Risk Distribution table (row 11+) ──
    _section_header(ws, row=11, col=1, title="RISK DISTRIBUTION", span=4)
    risk_headers = ["Risk Tier", "SOs", "Volume (MT)", "% of Baseline"]
    _table_header(ws, row=12, col=1, headers=risk_headers, widths=[18, 8, 16, 16])
    baseline_mt = master["sc_vol_mt"].sum() or 1
    risk_order = ["Green", "Yellow", "Orange", "Red", "Critical"]
    for i, tier in enumerate(risk_order):
        sub = master[master["risk_tier"] == tier]
        vol = sub["sc_vol_mt"].sum()
        row_data = [tier, len(sub), round(vol, 1), f"{vol/baseline_mt*100:.1f}%"]
        r = 13 + i
        fill_c, font_c = RISK_STYLE.get(tier, (_fill(C_OFF_WHITE), _font()))
        for j, val in enumerate(row_data):
            c = ws.cell(row=r, column=1 + j, value=val)
            c.fill = fill_c if j == 0 else _fill(C_OFF_WHITE if i % 2 == 0 else C_LIGHT_GREY)
            c.font = font_c if j == 0 else _font()
            c.alignment = _align("center" if j > 0 else "left")
            c.border = _border_thin()
        ws.row_dimensions[r].height = 20

    # ── Plant Breakdown table ──
    _section_header(ws, row=11, col=6, title="BY PLANT", span=5)
    plant_headers = ["Plant", "SOs", "Baseline MT", "Shipped", "FG", "WIP", "Red+Critical"]
    _table_header(ws, row=12, col=6, headers=plant_headers, widths=[10, 8, 14, 12, 10, 10, 14])
    plants = master.groupby("plant").agg(
        so_count=("so", "count"),
        baseline_mt=("sc_vol_mt", "sum"),
        shipped_mt=("shipped_mt", "sum"),
        fg_mt=("fg_mt", "sum"),
        wip_mt=("wip_mt", "sum"),
        risk_red=("risk_tier", lambda x: x.isin(["Red","Critical"]).sum())
    ).reset_index()
    for i, (_, row) in enumerate(plants.iterrows()):
        r = 13 + i
        vals = [row["plant"], row["so_count"], round(row["baseline_mt"],1),
                round(row["shipped_mt"],1), round(row["fg_mt"],1),
                round(row["wip_mt"],1), row["risk_red"]]
        bg = C_OFF_WHITE if i % 2 == 0 else C_LIGHT_GREY
        for j, val in enumerate(vals):
            c = ws.cell(row=r, column=6 + j, value=val)
            c.fill = _fill(bg)
            c.font = _font(bold=(j == 0))
            c.alignment = _align("center" if j > 0 else "left")
            c.border = _border_thin()
        ws.row_dimensions[r].height = 20

    # ── Order Type Breakdown ──
    start_row = 11 + 5 + 4
    _section_header(ws, row=start_row, col=1, title="ORDER TYPE BREAKDOWN", span=9)
    ot_headers = ["Order Type", "SOs", "Volume (MT)", "Shipped", "FG", "WIP", "Unscheduled", "No Plan"]
    _table_header(ws, row=start_row+1, col=1, headers=ot_headers, widths=[28,8,14,12,10,10,14,12])
    for i, ot in enumerate(["Carry Over Stock","Carry Over Unproduced","Fresh Order This Month"]):
        sub = master[master["order_type"] == ot]
        vals = [ot, len(sub), round(sub["sc_vol_mt"].sum(),1),
                round(sub["shipped_mt"].sum(),1), round(sub["fg_mt"].sum(),1),
                round(sub["wip_mt"].sum(),1), round(sub["unsched_mt"].sum(),1),
                round(sub["no_plan_mt"].sum(),1)]
        r = start_row + 2 + i
        bg = C_OFF_WHITE if i % 2 == 0 else C_LIGHT_GREY
        for j, val in enumerate(vals):
            c = ws.cell(row=r, column=1+j, value=val)
            c.fill = _fill(bg)
            c.font = _font(bold=(j==0))
            c.alignment = _align("center" if j > 0 else "left")
            c.border = _border_thin()
        ws.row_dimensions[r].height = 20

    # Column widths for summary sheet
    for col in range(1, 13):
        ws.column_dimensions[get_column_letter(col)].width = 14


# ── Sheet 2: SO Master ─────────────────────────────────────────────────────
def _sheet_master(wb, master):
    ws = wb.create_sheet("SO Master")
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A3"

    _banner(ws, "SO MASTER — ALL BASELINE ORDERS", span=17)

    col_defs = [
        ("SO Number",        "so",               12),
        ("Plant",            "plant",              8),
        ("Cluster",          "cluster",           10),
        ("Order Type",       "order_type",        22),
        ("SC Vol (MT)",      "sc_vol_mt",         12),
        ("Shipped (MT)",     "shipped_mt",        12),
        ("FG Stock (MT)",    "fg_mt",             12),
        ("WIP (MT)",         "wip_mt",            12),
        ("Unscheduled (MT)", "unsched_mt",        16),
        ("No Plan (MT)",     "no_plan_mt",        14),
        ("Status",           "status",            22),
        ("Risk Tier",        "risk_tier",         12),
        ("Planned End Date", "planned_end_date",  16),
        ("Loading Date",     "loading_date",      14),
        ("Gap (days)",       "gap_days",          12),
        ("Machines",         "machines",          20),
        ("LP Source",        "lp_source",         14),
    ]
    headers = [h for h, _, _ in col_defs]
    widths  = [w for _, _, w in col_defs]
    fields  = [f for _, f, _ in col_defs]

    _table_header(ws, row=2, col=1, headers=headers, widths=widths)

    # Sort: Critical → Red → Orange → Yellow → Green
    risk_priority = {"Critical": 0, "Red": 1, "Orange": 2, "Yellow": 3, "Green": 4}
    sorted_df = master.copy()
    sorted_df["_rp"] = sorted_df["risk_tier"].map(risk_priority).fillna(5)
    sorted_df = sorted_df.sort_values(["_rp", "gap_days"], ascending=[True, True])

    for i, (_, row) in enumerate(sorted_df.iterrows()):
        r = 3 + i
        bg = C_OFF_WHITE if i % 2 == 0 else C_LIGHT_GREY
        for j, field in enumerate(fields):
            val = row.get(field)
            if pd.isna(val) if not isinstance(val, str) else False:
                val = ""
            if isinstance(val, float) and val == int(val) and field.endswith("_mt"):
                val = round(val, 1)
            c = ws.cell(row=r, column=1+j, value=val)
            c.fill = _fill(bg)
            c.font = _font()
            c.alignment = _align("center" if j > 1 else "left")
            c.border = _border_thin()
            # Risk tier column coloring
            if field == "risk_tier" and val in RISK_STYLE:
                c.fill, c.font = RISK_STYLE[val]
            # Gap coloring
            if field == "gap_days" and isinstance(val, (int, float)):
                if val < 0:
                    c.fill = _fill(C_RED); c.font = _font(bold=True, color=C_RED_FG)
                elif val <= 2:
                    c.fill = _fill(C_YELLOW); c.font = _font(bold=True, color=C_YELLOW_FG)
        ws.row_dimensions[r].height = 18

    ws.auto_filter.ref = f"A2:{get_column_letter(len(col_defs))}2"


# ── Sheet 3: Gap Detail ────────────────────────────────────────────────────
def _sheet_gap_detail(wb, master):
    ws = wb.create_sheet("Gap Analysis")
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A3"

    gap_df = master[
        (master["status"] == "In Production") & master["gap_days"].notna()
    ].sort_values("gap_days").copy()

    _banner(ws, f"GAP ANALYSIS — IN PRODUCTION SOs ({len(gap_df)} orders)", span=10)

    col_defs = [
        ("SO Number",        "so",               12),
        ("Plant",            "plant",              8),
        ("Cluster",          "cluster",           10),
        ("Order Type",       "order_type",        22),
        ("SC Vol (MT)",      "sc_vol_mt",         12),
        ("WIP (MT)",         "wip_mt",            12),
        ("Planned End Date", "planned_end_date",  16),
        ("Loading Date",     "loading_date",      14),
        ("Gap (days)",       "gap_days",          12),
        ("Risk Tier",        "risk_tier",         12),
    ]
    headers = [h for h, _, _ in col_defs]
    widths  = [w for _, _, w in col_defs]
    fields  = [f for _, f, _ in col_defs]

    _table_header(ws, row=2, col=1, headers=headers, widths=widths)

    for i, (_, row) in enumerate(gap_df.iterrows()):
        r = 3 + i
        bg = C_OFF_WHITE if i % 2 == 0 else C_LIGHT_GREY
        for j, field in enumerate(fields):
            val = row.get(field)
            if pd.isna(val) if not isinstance(val, str) else False:
                val = ""
            c = ws.cell(row=r, column=1+j, value=val)
            c.fill = _fill(bg)
            c.font = _font()
            c.alignment = _align("center" if j > 1 else "left")
            c.border = _border_thin()
            if field == "gap_days" and isinstance(val, (int, float)):
                if val < 0:
                    c.fill = _fill(C_RED); c.font = _font(bold=True, color=C_RED_FG)
                elif val <= 2:
                    c.fill = _fill(C_YELLOW); c.font = _font(bold=True, color=C_YELLOW_FG)
                else:
                    c.fill = _fill(C_GREEN); c.font = _font(bold=True, color=C_GREEN_FG)
            if field == "risk_tier" and val in RISK_STYLE:
                c.fill, c.font = RISK_STYLE[val]
        ws.row_dimensions[r].height = 18

    ws.auto_filter.ref = f"A2:{get_column_letter(len(col_defs))}2"


# ── Sheet 4: Action Required ───────────────────────────────────────────────
def _sheet_action(wb, master):
    ws = wb.create_sheet("Action Required")
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A3"

    action_df = master[
        master["status"].isin(["No Plan", "Planned (Unscheduled)"])
    ].sort_values(["status", "sc_vol_mt"], ascending=[True, False]).copy()

    _banner(ws, f"ACTION REQUIRED — {len(action_df)} SOs NEED IMMEDIATE ATTENTION", span=8, urgent=True)

    col_defs = [
        ("SO Number",        "so",               12),
        ("Plant",            "plant",              8),
        ("Cluster",          "cluster",           10),
        ("Order Type",       "order_type",        22),
        ("SC Vol (MT)",      "sc_vol_mt",         12),
        ("Status",           "status",            22),
        ("Risk Tier",        "risk_tier",         12),
        ("Loading Date",     "loading_date",      14),
    ]
    headers = [h for h, _, _ in col_defs]
    widths  = [w for _, _, w in col_defs]
    fields  = [f for _, f, _ in col_defs]

    _table_header(ws, row=2, col=1, headers=headers, widths=widths)

    for i, (_, row) in enumerate(action_df.iterrows()):
        r = 3 + i
        bg = C_OFF_WHITE if i % 2 == 0 else C_LIGHT_GREY
        for j, field in enumerate(fields):
            val = row.get(field)
            if pd.isna(val) if not isinstance(val, str) else False:
                val = ""
            c = ws.cell(row=r, column=1+j, value=val)
            c.fill = _fill(bg)
            c.font = _font()
            c.alignment = _align("center" if j > 1 else "left")
            c.border = _border_thin()
            if field == "risk_tier" and val in RISK_STYLE:
                c.fill, c.font = RISK_STYLE[val]
        ws.row_dimensions[r].height = 18

    ws.auto_filter.ref = f"A2:{get_column_letter(len(col_defs))}2"


# ── Helper functions ────────────────────────────────────────────────────────
def _banner(ws, title, span=10, urgent=False):
    ws.merge_cells(f"A1:{get_column_letter(span)}1")
    c = ws["A1"]
    c.value = title
    c.font = Font(bold=True, size=13, color=C_WHITE, name="Calibri")
    c.fill = _fill("7B0000" if urgent else C_NAVY)
    c.alignment = _align("left")
    ws.row_dimensions[1].height = 30


def _section_header(ws, row, col, title, span=4):
    end_col = get_column_letter(col + span - 1)
    ws.merge_cells(f"{get_column_letter(col)}{row}:{end_col}{row}")
    c = ws.cell(row=row, column=col, value=title)
    c.font = Font(bold=True, size=10, color=C_WHITE, name="Calibri")
    c.fill = _fill(C_DARK_BLUE)
    c.alignment = _align("left")
    ws.row_dimensions[row].height = 22


def _table_header(ws, row, col, headers, widths):
    for j, (h, w) in enumerate(zip(headers, widths)):
        c = ws.cell(row=row, column=col+j, value=h)
        c.font = Font(bold=True, size=9, color=C_WHITE, name="Calibri")
        c.fill = _fill(C_MID_BLUE)
        c.alignment = _align("center")
        c.border = _border_thin()
        ws.column_dimensions[get_column_letter(col+j)].width = w
    ws.row_dimensions[row].height = 22


def _kpi_header(ws, row):
    ws.row_dimensions[row].height = 6


def _kpi_card(ws, row, col, label, value, sub, accent):
    # Label row
    ws.merge_cells(f"{get_column_letter(col)}{row}:{get_column_letter(col+1)}{row}")
    c = ws.cell(row=row, column=col, value=label)
    c.font = Font(bold=True, size=8, color="888888", name="Calibri")
    c.fill = _fill(C_OFF_WHITE)
    c.alignment = _align("center")
    c.border = Border(top=Side(style="medium", color=accent))
    ws.row_dimensions[row].height = 18

    # Value row
    ws.merge_cells(f"{get_column_letter(col)}{row+1}:{get_column_letter(col+1)}{row+1}")
    c = ws.cell(row=row+1, column=col, value=value)
    c.font = Font(bold=True, size=18, color=accent, name="Calibri")
    c.fill = _fill(C_OFF_WHITE)
    c.alignment = _align("center")
    ws.row_dimensions[row+1].height = 30

    # Sub row
    ws.merge_cells(f"{get_column_letter(col)}{row+2}:{get_column_letter(col+1)}{row+2}")
    c = ws.cell(row=row+2, column=col, value=sub)
    c.font = Font(size=9, color="888888", name="Calibri")
    c.fill = _fill(C_OFF_WHITE)
    c.alignment = _align("center")
    ws.row_dimensions[row+2].height = 16

    # Bottom border spacer
    ws.merge_cells(f"{get_column_letter(col)}{row+3}:{get_column_letter(col+1)}{row+3}")
    c = ws.cell(row=row+3, column=col)
    c.fill = _fill(C_OFF_WHITE)
    ws.row_dimensions[row+3].height = 6
