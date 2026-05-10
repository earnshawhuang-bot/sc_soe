# S&OE Order Fulfillment Tracking System - Blueprint

## 1. Purpose

For each sales order (SO) in hand, answer three questions:
- Has it shipped? Has it been produced? Is it scheduled?
- If scheduled, can production finish before the required loading date?
- If not scheduled, flag it as a risk.

This is a **monthly recurring mechanism**: feed new data each month, output the same structured risk view. Target: ~25,000 MT/month across Kunshan (KS) and Indonesia (IDN).

**Data nature**: Shipped = MTD cumulative (e.g. May 1–7, next time May 1–10); FG = point-in-time snapshot on data extraction day. Both refresh each run cycle.

---

## 2. Data Flow (one paragraph)

SC table defines the **baseline** (what SOs we need to fulfill this month). We then check each SO's progress: Shipped tells us what's already gone; FG tells us what's sitting in warehouse ready to go; PP tells us what's in production and when it finishes. Loading Plan tells us the **deadline** (when each SO must be loaded). The gap between PP's planned finish date and Loading Plan's loading date is our core risk signal. SOs with no PP record and no Loading Plan entry are the highest risk.

---

## 3. Per-Table Extraction

### 3.1 SC - Sales Order Baseline (01-SC)

**File**: `Order tracking 0507.xlsx` → Sheet "Order Status - Main"

**Filter**: Item column = "main"

**Key columns** (use fixed column index positions — header names can shift):
| Index | Letter | Field | Usage |
|-------|--------|-------|-------|
| 11 | L | SC NO | SO number (primary key) |
| 5  | F | Supply From | Plant: Kunshan→KS, Indonesia→IDN |
| 3  | D | Cluster | Region dimension |
| 28 | AC | SC Vol.-MT | Order quantity (MT) |
| 10 | K | Order Status | Classification input |
| 16 | Q | Carryover | **Null** = this month's billing |
| 40 | AO | carryover/fresh production | "last month" = Carry Over Unproduced |
| 8  | I | RELEASE DATE | Month determination |
| 15 | P | Loading Date | **NOT RELIABLE** — do not use as baseline |

**Classification (4 types, independently verified MECE)**:
1. **Carry Over Stock**: K = `"Carryover"` AND Q is null
2. **Carry Over Unproduced**: AO = `"last month"` AND Q is null
3. **Fresh Order This Month**: Release Date in current month AND Q is null
4. **Fresh Order Next Month**: Release Date in current month AND Q has value → **excluded from baseline**

**Baseline** = Type 1 + Type 2 + Type 3

**Aggregation**: One SO can appear on multiple rows (different SKUs). Aggregate by SO: `sc_vol_mt = SUM`, categorical fields = first.

> ⚠️ **Pitfall**: K column value is `"Carryover"` (single word, capital C) — NOT `"carry over"` or `"Carry Over"`. Matching against the wrong string returns 0 rows for Carry Over Stock.

> ⚠️ **Pitfall**: Q column (Carryover) holds a numeric pcs value when filled, not text. "Empty" means `pd.isna()` — do NOT check for empty string.

> ⚠️ **Pitfall**: AO column header contains a newline: `"carryover/fresh\nproduction"`. Use index position (40) instead of name matching.

---

### 3.2 Shipped - Actual Shipments (02-Shipped)

**File**: `GI (KS&IDN) 260501-0507.xlsx`

**Key columns**:
| Column | Field | Usage |
|--------|-------|-------|
| B | Post Date | Actual shipment date |
| C | Plant | 3000/3001=KS, 3301=IDN |
| Q (销售订单.1) | Sales Order | SO number (join key) |
| S | Weight (KG) | × (-1) ÷ 1000 → MT |

**Join**: Shipped.SO → SC.SO (direct match)

> ⚠️ **Pitfall**: Two columns share the name "销售订单" (pandas deduplicates to `.1`). The correct one is `销售订单.1` (Q column), which holds 10-prefix SC numbers. `销售订单` (P column) holds internal order numbers starting with 8.

> ⚠️ **Pitfall**: All SO numbers in Excel are stored as `float64` (e.g. `1000011733.0`). Must convert via `int(float(value))` before string matching — direct `.astype(str)` produces `"1000011733.0"` which fails regex match.

---

### 3.3 FG - Finished Goods Inventory (03-FG)

**File**: `FG stock 050826.xlsx`

**Key columns**:
| Column | Field | Usage |
|--------|-------|-------|
| A | Plant | Only KS (3000/3001). IDN data absent. |
| J | Weight (KG) | ÷ 1000 → MT |
| L | Receipt Date | When it entered warehouse |
| AG | Contract Code | SO number (join key) |

**Join**: FG.Contract Code → SC.SO (direct match)

**Note**: IDN has no FG data currently. IDN orders skip "In Stock" status.

---

### 3.4 PP - Production Plan (04-PP)

**Scheduled (4 files, unified structure)**:
- `SAM日生产计划表.xlsx` — KS
- `KS Ⅱ 日生产计划表.xlsx` — KS
- `DAVIS日生产计划表.xlsx` — KS
- `IND Production Plan.xlsx` — IDN

**Key columns**: Work Order No., SO, TotalWeight/T, PlannedFinishDate, Machine, Plant

**Unscheduled (1 file)**:
- `Global_PP_wo schedule.xlsx`

**Key columns**: Work Order No., SO, TotalWeight/T, Machine, Plant (**no date**)

**Join**: PP.SO → SC.SO (direct match)

**Multi-work-order aggregation** (one SO can have 20+ work orders across machines/files):
- `wip_mt` = SUM of all work orders' TotalWeight under that SO
- `planned_end_date` = MAX(PlannedFinishDate) across all work orders under that SO
  - Rationale: the SO cannot be loaded until ALL sub-batches are complete; latest date is the bottleneck

> ⚠️ **Pitfall**: PP SO numbers are also stored as `float64`. Same `int(float())` conversion required.

---

### 3.5 Loading Plan - Required Loading Dates (06-Loading Plan)

#### 3.5.1 KS Loading Plan

**File**: `LoadingPlan_20260508.xlsx` → Sheet "Loading Plan"

**Key columns**:
| Column | Field | Usage |
|--------|-------|-------|
| A | Invoice No. | → split to get SO number |
| D | 20GP qty | × 15 MT |
| E | 40GP qty | × 24.5 MT |
| F | 40HQ qty | × 24.5 MT |
| H | Loading Date | **Required loading date (deadline)** |

**Invoice No. splitting rules** (only 4.2% need split):
- Underscore `_` = multi-SO separator
  - `1000012350_11631_11479` → 3 SOs: 1000012350, 1000011631, 1000011479
  - Short numbers inherit prefix from the longest number in the group
- Dash `-` = container sequence within same SO (strip suffix)
  - `1000012299-1` → SO: 1000012299
- Combined: strip dash first, then split by underscore, then prefix-complete short numbers
- Tonnage: after split, total MT evenly distributed across SOs

**Time filter**: Loading Date >= (data_date − 3 days)

**Join**: Split SO → SC.SO

---

#### 3.5.2 IDN Export - Schedule Planning Dispatch

**File**: `Schedule Planning Dispatch 260508.xlsx` → Sheet `"ORDER OUTSTANDING "` *(trailing space in sheet name)*

**Key columns**:
| Column | Field | Usage |
|--------|-------|-------|
| C | SC No. | SO number (direct, no split needed) |
| E | Region | Cluster (SEA/ISU/MEA) |
| J | Cont Qty | Container quantity |
| K | Cont Size | 40'HC=24.5MT, 20'=15MT |
| S | ELD | **Required loading date (deadline)** |

**Tonnage**: Cont Qty × corresponding MT per size

**Time filter**: ELD >= (data_date − 3 days)

**Join**: SC No. → SC.SO (direct match, 100% pure numbers)

> ⚠️ **Pitfall**: Sheet name has a trailing space: `"ORDER OUTSTANDING "`. Must match exactly.

---

#### 3.5.3 IDN Domestic - New Domestic Tracking

**File**: `NEW DOMESTIC TRACKING 260508.xlsx` → Sheet `"Order List"`, headers in **row 2**

**Key columns**:
| Column | Field | Usage |
|--------|-------|-------|
| D | INV NO. | → split to get SO number |
| G | ELD | **Required loading date (deadline)** |
| P | Weight | Actual weight in MT |

**INV NO. splitting rules** (for 10-prefixed orders only; `LMIDSAM*` and other non-10 prefixes → discard):
- Comma `,` or Ampersand `&` = multi-SO separator
- Dash `-` = batch/trip number → strip suffix, keep SO number before dash
- Processing order: split by `,` or `&` first → strip `-N` from each part
- Tonnage: P column Weight evenly distributed across SOs after split

**Time filter**: ELD >= (data_date − 3 days)

**Join**: Split SO → SC.SO

---

## 4. Join & Status Determination

All joins converge on **SC.SO** as primary key.

```
SC (baseline)
├── LEFT JOIN Shipped     → shipped_mt (per SO)
├── LEFT JOIN FG          → fg_mt (per SO, KS only)
├── LEFT JOIN PP_sched    → planned_end_date, wip_mt, machines
├── LEFT JOIN PP_unsched  → unsched_mt
└── LEFT JOIN Loading Plan (KS + IDN Export + IDN Domestic)
                          → loading_date, load_mt, lp_source
```

**Status assignment (quantity waterfall — not a single label)**:

Each SO gets a quantity breakdown:
- `shipped_mt`: already shipped
- `fg_mt`: in warehouse, ready to ship
- `wip_mt`: in production with planned_end_date
- `unsched_mt`: has work order but no date
- `no_plan_mt`: remainder = SC qty − (all above), floored at 0

**Primary status label** (priority order):
1. Shipped / Partially Shipped
2. In Stock (FG > 0)
3. In Production (wip_mt > 0, has planned_end_date)
4. Planned (Unscheduled) (work order exists, no date)
5. No Plan

---

## 5. Gap Calculation

**Applies only to SOs with BOTH**:
- `planned_end_date` from PP (= MAX across all work orders for that SO)
- `loading_date` from Loading Plan

**Formula**:
```
Gap (days) = Loading_Date − (MAX(Planned_End_Date) + 1)
```
- `+1 day`: production finish → next day = earliest warehouse receipt = earliest loadable

**Interpretation**:
- Gap > 2 → Green: buffer exists
- Gap 0–2 → Yellow: tight but feasible
- Gap < 0 → Red: production behind, |gap| = days overdue

**Risk tiers for SOs without a computable gap**:
- Has PP schedule but no Loading Plan entry → Orange
- Has work order, no schedule date → Red
- No work order at all → Critical

---

## 6. Output

### Excel (4 sheets, consulting-style formatted)
- **Summary**: Banner + KPI cards + Risk Distribution + **Plant × Region (Cluster) breakdown** + Order Type Breakdown
- **SO Master**: Full detail per SO, sorted by risk, with conditional gap coloring and auto-filter
- **Gap Analysis**: In-Production SOs only, sorted by gap ascending, gap cells color-coded
- **Action Required**: No Plan + Unscheduled SOs, urgent red banner

### HTML Report (single-page narrative)
- Section 1: Monthly overview — target vs progress waterfall bar
- Section 2: Risk distribution table
- Section 3: By Plant table
- Section 4: Top 20 risk items

---

## 7. Management View Dimensions

Management reviews across **three levels of granularity**:

| Level | Dimension | Used for |
|-------|-----------|---------|
| L1 | Total | Monthly target vs actual headline |
| L2 | Plant (KS / IDN) | Factory-level capacity and risk ownership |
| L3 | Plant × Region (Cluster) | Customer segment delivery visibility |

**Clusters observed in data**: CHINA, CIS, MEA, SEA, ISU, and others.

The Summary sheet presents all three levels: total KPIs → by Plant → by Plant × Cluster.

---

## 8. Logical Closed Loop

```
        SC Baseline (what we OWE)
              │
    ┌─────────┼─────────┐
    ▼         ▼         ▼
 Shipped    FG Stock   PP Schedule
 (done)    (ready)    (in progress)
                         │
                         ▼
              Loading Plan (DEADLINE)
                         │
                         ▼
                  Gap = Deadline − Production
                         │
                         ▼
              Risk Signal → Action Required
```

Every SO in the baseline lands in exactly ONE path. Waterfall quantities sum back to SC Vol.-MT. This is the MECE closure.

---

## 9. Known Pitfalls & Lessons Learned

| # | Where | Issue | Fix |
|---|-------|-------|-----|
| 1 | SC K column | Value is `"Carryover"` not `"carry over"` — wrong string = 0 matches, entire Carry Over Stock category lost | Match exact string `"Carryover"` |
| 2 | SC Q column | Holds numeric pcs value when filled; "empty" = `pd.isna()`, not empty string check | Use `.isna()` |
| 3 | SC AO column | Header has embedded newline `\nproduction`; name-based search hits Q column first | Use fixed column index (40) |
| 4 | SC rows | One SO can have multiple rows (different SKUs) — naively treating rows as SOs inflates count | Aggregate by SO: sum vol, first for categoricals |
| 5 | All numeric IDs | SO/plant codes stored as `float64` in Excel (e.g. `1000011733.0`) — `.astype(str)` gives `"1000011733.0"`, breaks regex | Convert via `str(int(float(value)))` |
| 6 | Shipped SO column | Two columns named "销售订单"; pandas deduplicates to `.1`; P column has internal order numbers, Q column (`.1`) has SC numbers | Prefer the column where values start with `"10"` |
| 7 | IDN Dispatch sheet | Sheet name has trailing space: `"ORDER OUTSTANDING "` | Match with trailing space |
| 8 | SC Loading Date (P) | Unreliable — values like "Pending" are common | Never use as fallback for gap calculation |
