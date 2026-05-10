import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch

plt.rcParams['font.family'] = ['PingFang HK', 'Heiti TC', 'STHeiti', 'SimHei', 'sans-serif']
plt.rcParams['axes.unicode_minus'] = False

fig, ax = plt.subplots(1, 1, figsize=(28, 20))
ax.set_xlim(0, 28)
ax.set_ylim(0, 20)
ax.axis('off')
fig.patch.set_facecolor('#FAFAFA')

# ── Color palette ──
C_SOURCE = '#E3F2FD'
C_SOURCE_BORDER = '#1565C0'
C_CLEAN = '#FFF3E0'
C_CLEAN_BORDER = '#E65100'
C_JOIN = '#E8F5E9'
C_JOIN_BORDER = '#2E7D32'
C_ENGINE = '#FCE4EC'
C_ENGINE_BORDER = '#C62828'
C_OUTPUT = '#F3E5F5'
C_OUTPUT_BORDER = '#6A1B9A'
C_ARROW = '#546E7A'
C_KEY_ARROW = '#D32F2F'

def draw_box(x, y, w, h, text, bg, border, fontsize=8, bold=False, alpha=0.9):
    box = FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.15",
                         facecolor=bg, edgecolor=border, linewidth=1.5, alpha=alpha)
    ax.add_patch(box)
    weight = 'bold' if bold else 'normal'
    ax.text(x + w/2, y + h/2, text, ha='center', va='center',
            fontsize=fontsize, fontweight=weight, color='#212121',
            linespacing=1.4, wrap=True)

def draw_arrow(x1, y1, x2, y2, color=C_ARROW, style='->', lw=1.2, ls='-'):
    ax.annotate('', xy=(x2, y2), xytext=(x1, y1),
                arrowprops=dict(arrowstyle=style, color=color, lw=lw, ls=ls,
                               connectionstyle='arc3,rad=0'))

def draw_label(x, y, text, fontsize=11, color='#212121', bold=True):
    ax.text(x, y, text, ha='center', va='center', fontsize=fontsize,
            fontweight='bold' if bold else 'normal', color=color,
            bbox=dict(boxstyle='round,pad=0.3', facecolor='white', edgecolor='#BDBDBD', alpha=0.95))

# ═══════════════════════════════════════════
# PHASE 0: Title
# ═══════════════════════════════════════════
ax.text(14, 19.5, 'S&OE Order Fulfillment Tracking — Data Flow', ha='center', va='center',
        fontsize=18, fontweight='bold', color='#1A237E')
ax.text(14, 19.0, '数据流全景图：从原始 Excel 到交付风险预警', ha='center', va='center',
        fontsize=12, color='#546E7A')

# ═══════════════════════════════════════════
# PHASE 1: Raw Data Sources (Layer 1)
# ═══════════════════════════════════════════
draw_label(14, 18.2, 'Phase 1: Raw Data Sources  原始数据源', fontsize=12, color=C_SOURCE_BORDER)

src_y = 16.2
src_h = 1.6
src_w = 3.8

# 01-SC
draw_box(0.3, src_y, src_w, src_h,
    '01-SC  销售订单\n─────────────\nOrder tracking 0507.xlsx\n1,266 行\n\n关键: SC NO, Supply from,\nCluster, SC Vol.-MT,\nOrder status, Carryover,\nLoading date',
    C_SOURCE, C_SOURCE_BORDER, fontsize=7.5)

# 02-Shipped
draw_box(4.5, src_y, src_w, src_h,
    '02-Shipped  已发货\n─────────────\nGI (KS&IDN) 260501-0507\n4,986 行\n\n关键: 过账日期, 工厂,\n销售订单(Q列), 重量KG\n⚡ 负数→正数, KG→MT',
    C_SOURCE, C_SOURCE_BORDER, fontsize=7.5)

# 03-FG
draw_box(8.7, src_y, src_w, src_h,
    '03-FG  成品库存\n─────────────\nFG stock 050826.xlsx\n3,940 行 (仅KS)\n\n关键: 工厂, 重量KG,\n入库日期, 合同编码(→SO)\n⚠️ IDN 数据缺失',
    C_SOURCE, C_SOURCE_BORDER, fontsize=7.5)

# 04-PP
draw_box(12.9, src_y, src_w, src_h,
    '04-PP  生产计划\n─────────────\n有排产: SAM/DAVIS/KSII/IND\n共 1,768 行\n无排产: Global_PP_wo\n221 行\n\n关键: 工单号, SO,\nPlanned End Date, Plant',
    C_SOURCE, C_SOURCE_BORDER, fontsize=7.5)

# 05-GR
draw_box(17.1, src_y, src_w, src_h,
    '05-GR  实际入库\n─────────────\nMB51 GR data\n6,117 行\n\n关键: 过账日期, 工厂,\n订单(=工单号), 重量KG\n⚠️ 无SO号! 需经PP中转',
    C_SOURCE, C_SOURCE_BORDER, fontsize=7.5)

# 06-Loading Plan
draw_box(21.3, src_y, 6.3, src_h,
    '06-Loading Plan  发货计划\n──────────────────────────\nKS: LoadingPlan_20260508.xlsx (357行)\n    发票号拆分 → SO, 柜型→MT换算\n\nIDN出口: Schedule Planning Dispatch\n    ORDER OUTSTANDING sheet (255柜)\n    SC No. + ELD + Cont Qty/Size\n\nIDN内销: NEW DOMESTIC TRACKING (327行)\n    INV NO.(需拆分) + Weight MT + ELD',
    C_SOURCE, C_SOURCE_BORDER, fontsize=7.5)

# ═══════════════════════════════════════════
# PHASE 2: Standardization (Layer 2)
# ═══════════════════════════════════════════
draw_label(14, 14.7, 'Phase 2: Extract & Standardize  数据清洗标准化', fontsize=12, color=C_CLEAN_BORDER)

clean_y = 12.6
clean_h = 1.5
clean_w = 3.2

# Arrows from sources to clean
for sx in [2.2, 6.4, 10.6, 14.8, 19.0, 24.4]:
    draw_arrow(sx, src_y, sx if sx < 22 else 21, clean_y + clean_h, color='#90A4AE', lw=0.8)

draw_box(0.5, clean_y, clean_w, clean_h,
    'sc_df\n─────────\nSO号, Plant(KS/IDN)\nCluster, 订单量MT\n订单类型(CO/Fresh)',
    C_CLEAN, C_CLEAN_BORDER, fontsize=7.5)

draw_box(4.0, clean_y, clean_w, clean_h,
    'shipped_df\n─────────\nSO号, Plant\n发货日期\n发货量MT',
    C_CLEAN, C_CLEAN_BORDER, fontsize=7.5)

draw_box(7.5, clean_y, clean_w, clean_h,
    'fg_df\n─────────\nSO号, Plant(仅KS)\n库存量MT\n入库日期',
    C_CLEAN, C_CLEAN_BORDER, fontsize=7.5)

draw_box(11.0, clean_y, clean_w+0.5, clean_h,
    'pp_sched_df + pp_unsched_df\n─────────\n工单号, SO号, Plant\n计划完工日期(有/无)\n重量MT, Machine',
    C_CLEAN, C_CLEAN_BORDER, fontsize=7.5)

draw_box(15.2, clean_y, clean_w, clean_h,
    'gr_df\n─────────\n工单号, Plant\n入库日期, 入库量MT\n移动类型(101/102)',
    C_CLEAN, C_CLEAN_BORDER, fontsize=7.5)

draw_box(19.0, clean_y, clean_w+1.0, clean_h,
    'lp_ks_df + lp_idn_df\n+ lp_idn_dom_df\n─────────\nSO号, 装货日期(Loading Date)\n柜量MT, Cluster',
    C_CLEAN, C_CLEAN_BORDER, fontsize=7.5)

# ═══════════════════════════════════════════
# PHASE 3: Join (Layer 3)
# ═══════════════════════════════════════════
draw_label(14, 11.5, 'Phase 3: Join & Enrich  数据关联', fontsize=12, color=C_JOIN_BORDER)

join_y = 9.4
join_h = 1.5

# GR enrichment sub-step
draw_box(13.0, 10.8, 5.5, 0.55,
    'GR + PP 中转关联:  gr.订单号 → pp.工单号 → pp.SO  →  gr_enriched_df (补上SO号)',
    '#FFFDE7', '#F57F17', fontsize=7.5)
draw_arrow(12.25, clean_y, 14.5, 11.35, color=C_KEY_ARROW, lw=1.5, style='->')
draw_arrow(16.8, clean_y, 16.0, 11.35, color=C_KEY_ARROW, lw=1.5, style='->')

# SO Master join
draw_box(3.5, join_y, 21, join_h,
    'so_master_df — 以 SO 为主键的全量关联表\n'
    '═══════════════════════════════════════════════════════════════════════════════\n'
    'sc_df  ←LEFT JOIN→  shipped_df     ←LEFT JOIN→  fg_df     ←LEFT JOIN→  pp_sched_df / pp_unsched_df\n'
    '                     (on SO号)                    (on SO号)               (on SO号)\n'
    '       ←LEFT JOIN→  gr_enriched_df  ←LEFT JOIN→  lp_ks_df / lp_idn_df / lp_idn_dom_df\n'
    '                     (on SO号)                    (on SO号)',
    C_JOIN, C_JOIN_BORDER, fontsize=8)

# Arrows from clean to join
for cx in [2.1, 5.6, 9.1, 12.5, 20.5]:
    draw_arrow(cx, clean_y, cx if cx < 22 else 20, join_y + join_h, color='#66BB6A', lw=1.0)
draw_arrow(15.75, 10.8, 15.75, join_y + join_h, color=C_KEY_ARROW, lw=1.5)

# ═══════════════════════════════════════════
# PHASE 4: Status Engine (Layer 4)
# ═══════════════════════════════════════════
draw_label(14, 8.5, 'Phase 4: Status Engine  状态判定 + Gap 计算', fontsize=12, color=C_ENGINE_BORDER)

eng_y = 5.8
eng_h = 2.2

draw_arrow(14, join_y, 14, eng_y + eng_h, color=C_ENGINE_BORDER, lw=2.0)

draw_box(1.0, eng_y, 8.5, eng_h,
    '状态判定逻辑 (按优先级)\n'
    '════════════════════════════\n'
    '① shipped_qty ≥ order_qty  →  Fully Shipped ✅\n'
    '② shipped_qty > 0          →  Partially Shipped 🟡\n'
    '③ fg_qty > 0               →  In Stock (待发货) 🟢\n'
    '④ gr_qty > 0               →  Produced (已入库) 🟢\n'
    '⑤ has planned_end_date     →  In Production (排产中) ⏳\n'
    '⑥ has work_order, no date  →  Planned Only (仅工单) 🟠\n'
    '⑦ none of above            →  No Plan (无计划) 🔴',
    C_ENGINE, C_ENGINE_BORDER, fontsize=8)

draw_box(10.5, eng_y, 8.0, eng_h,
    'Gap 计算 (仅对 ⑤ In Production)\n'
    '════════════════════════════\n'
    'Gap = Required_Loading_Date\n'
    '      - (Planned_End_Date + 1天)\n'
    '─────────────────────────────\n'
    'Gap < 0   →  🔴 生产滞后 |Gap|天\n'
    '              需催产 或 推船期\n'
    'Gap 0~2   →  🟡 紧张, 无缓冲\n'
    'Gap > 2   →  🟢 有缓冲',
    C_ENGINE, C_ENGINE_BORDER, fontsize=8)

draw_box(19.5, eng_y, 8.0, eng_h,
    '风险等级汇总\n'
    '════════════════════════════\n'
    '🔴🔴 Critical: No Plan\n'
    '     有订单但完全无生产安排\n\n'
    '🔴  High: Gap<0 或 Planned Only\n'
    '     生产滞后 或 未排具体日期\n\n'
    '🟡  Medium: Gap 0~2 或 Partial Ship\n'
    '     紧张但尚可调度\n\n'
    '🟢  Low: Gap>2 / In Stock / Shipped',
    C_ENGINE, C_ENGINE_BORDER, fontsize=8)

# ═══════════════════════════════════════════
# PHASE 5: Output (Layer 5)
# ═══════════════════════════════════════════
draw_label(14, 4.9, 'Phase 5: Output  输出', fontsize=12, color=C_OUTPUT_BORDER)

out_y = 2.8
out_h = 1.6

draw_arrow(14, eng_y, 14, out_y + out_h, color=C_OUTPUT_BORDER, lw=2.0)

draw_box(1.5, out_y, 10.5, out_h,
    'Excel 明细表\n'
    '═══════════════════════════════\n'
    'Sheet 1: SO Master 全量明细 (状态/gap/风险/量)\n'
    'Sheet 2: Risk Summary 按 Cluster × Plant 汇总\n'
    'Sheet 3: Gap Detail 排产中订单按gap排序\n'
    'Sheet 4: Action List 无计划/仅工单 待跟进清单',
    C_OUTPUT, C_OUTPUT_BORDER, fontsize=8.5)

draw_box(13.5, out_y, 13.5, out_h,
    'HTML Interactive Dashboard\n'
    '═══════════════════════════════════════\n'
    '目标达成: 25,000吨 vs 当前进度 (瀑布图)\n'
    '状态分布: Shipped/Stock/Production/NoPlan by Plant (堆叠柱状图)\n'
    'Gap 热力图: 日历视图 Loading Date vs Planned End Date\n'
    '风险清单: 可排序/筛选交互表格 | Cluster维度交付率',
    C_OUTPUT, C_OUTPUT_BORDER, fontsize=8.5)

# ═══════════════════════════════════════════
# Key linkage annotations
# ═══════════════════════════════════════════
ax.text(14, 1.8, '🔑 核心关联键: SO号 (销售订单号)  |  ⚠️ GR→SO 需经 PP 工单号中转  |  📅 月度复用: config.yaml 驱动',
        ha='center', va='center', fontsize=10, color='#37474F',
        bbox=dict(boxstyle='round,pad=0.4', facecolor='#ECEFF1', edgecolor='#78909C'))

# Legend
legend_y = 1.0
for i, (label, color, border) in enumerate([
    ('原始数据源', C_SOURCE, C_SOURCE_BORDER),
    ('清洗标准化', C_CLEAN, C_CLEAN_BORDER),
    ('关联合并', C_JOIN, C_JOIN_BORDER),
    ('状态/风险引擎', C_ENGINE, C_ENGINE_BORDER),
    ('输出', C_OUTPUT, C_OUTPUT_BORDER),
]):
    x = 4 + i * 4.5
    box = FancyBboxPatch((x, legend_y - 0.2), 0.5, 0.4, boxstyle="round,pad=0.05",
                         facecolor=color, edgecolor=border, linewidth=1.2)
    ax.add_patch(box)
    ax.text(x + 0.7, legend_y, label, va='center', fontsize=9, color='#424242')

plt.tight_layout()
plt.savefig('/Users/0xiaobo0/Desktop/Projects/Analytics/siop_target_simulation/SOE_DataFlow.png',
            dpi=180, bbox_inches='tight', facecolor='#FAFAFA')
print('Saved: SOE_DataFlow.png')
