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

**Key columns**:
| Column | Field | Usage |
|--------|-------|-------|
| L | SC NO | SO number (primary key) |
| F | Supply From | Plant: Kunshan→KS, Indonesia→IDN |
| D | Cluster | Region dimension |
| AC | SC Vol.-MT | Order quantity (MT) |
| K | Order Status | Classification input |
| Q | Carryover | Empty = this month's billing |
| AO | Carryover/Fresh | "Last Month" = Carry Over Unproduced |
| I | Release Date | Month determination |
| P | Loading Date | **NOT RELIABLE** - only as last resort |

**Classification (4 types, independently verified MECE)**:
1. **Carry Over Stock**: Order Status = "Carry Over" AND Q = empty
2. **Carry Over Unproduced**: AO = "Last Month" AND Q = empty
3. **Fresh Order This Month**: Release Date in current month AND Q = empty
4. **Fresh Order Next Month**: Release Date in current month AND Q has value → **excluded from baseline**

**Baseline** = Type 1 + Type 2 + Type 3

---

### 3.2 Shipped - Actual Shipments (02-Shipped)

**File**: `GI (KS&IDN) 260501-0507.xlsx`

**Key columns**:
| Column | Field | Usage |
|--------|-------|-------|
| B | Post Date | Actual shipment date |
| C | Plant | 3000/3001=KS, 3301=IDN |
| Q | Sales Order | SO number (join key) |
| S | Weight (KG) | × (-1) ÷ 1000 → MT |

**Join**: Shipped.SO → SC.SO (direct match)

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
- `SAM日生产计划表.xlsx` (445 rows) — KS
- `KS Ⅱ 日生产计划表.xlsx` (40 rows) — KS
- `DAVIS日生产计划表.xlsx` (935 rows) — KS
- `IND Production Plan.xlsx` (348 rows) — IDN

**Key columns**: Work Order No., SO, TotalWeight/T, PlannedFinishDate, Machine, Plant

**Unscheduled (1 file)**:
- `Global_PP_wo schedule.xlsx` (221 rows)

**Key columns**: Work Order No., SO, TotalWeight/T, Machine, Plant (**no date**)

**Join**: PP.SO → SC.SO (direct match)

**Multi-work-order aggregation** (common: one SO can have 20+ work orders across machines/files):
- `wip_mt` = SUM of all work orders' TotalWeight under that SO
- `planned_end_date` = MAX(PlannedFinishDate) across all work orders under that SO
  - Rationale: the SO cannot be loaded until ALL sub-batches are complete; latest date is the bottleneck

**Significance**: MAX(PlannedFinishDate) is one side of the gap calculation.

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
- Dash `-` = container sequence within same SO (strip it)
  - `1000012299-1` → SO: 1000012299
- Combined: underscore first, then strip dash from each part
- Tonnage: after split, total MT evenly distributed across SOs

**Time filter**: Loading Date >= (data_date - 3 days). E.g. data from May 8 → filter >= May 5

**Join**: Split SO → SC.SO

---

#### 3.5.2 IDN Export - Schedule Planning Dispatch

**File**: `Schedule Planning Dispatch 260508.xlsx` → Sheet "ORDER OUTSTANDING"

**Key columns**:
| Column | Field | Usage |
|--------|-------|-------|
| C | SC No. | SO number (direct, no split needed) |
| E | Region | Cluster (SEA/ISU/MEA) |
| J | Cont Qty | Container quantity |
| K | Cont Size | 40'HC=24.5MT, 20'=15MT |
| S | ELD | **Required loading date (deadline)** |

**Tonnage**: Cont Qty × corresponding MT per size

**Time filter**: ELD >= (data_date - 3 days)

**Join**: SC No. → SC.SO (direct match, 100% pure numbers)

---

#### 3.5.3 IDN Domestic - New Domestic Tracking

**File**: `NEW DOMESTIC TRACKING 260508.xlsx` → Sheet "Order List"

**Key columns**:
| Column | Field | Usage |
|--------|-------|-------|
| D | INV NO. | → split to get SO number |
| G | ELD | **Required loading date (deadline)** |
| P | Weight | Actual weight in MT |

**INV NO. splitting rules** (for 10-prefixed orders only; non-10 prefix like `LMIDSAM*` → remove):
- Comma `,` = multi-SO separator (43 entries)
- Ampersand `&` = multi-SO separator (9 entries)
- Dash `-` = batch/trip number, strip suffix, keep SO number before dash
- Processing order: split by `,` or `&` first → then strip `-N` from each part
- Tonnage after split: P column Weight evenly distributed across SOs (same as KS rule)

**Example**: `1000007826-7, 1000008015-4` → split by comma → `1000007826-7` and `1000008015-4` → strip dash → SO `1000007826` and SO `1000008015` → Weight ÷ 2 each

**Time filter**: ELD >= (data_date - 3 days)

**Join**: Split SO → SC.SO

---

## 4. Join & Status Determination

All joins converge on **SC.SO** as primary key.

```
SC (baseline)
├── LEFT JOIN Shipped     → shipped_mt (per SO)
├── LEFT JOIN FG          → fg_mt (per SO, KS only)
├── LEFT JOIN PP_sched    → planned_end_date, wip_mt
├── LEFT JOIN PP_unsched  → has_work_order (no date)
└── LEFT JOIN Loading Plan (KS + IDN Export + IDN Domestic)
                          → required_loading_date, planned_load_mt
```

**Status assignment (quantity waterfall, not single label)**:

Each SO gets a quantity breakdown:
- `shipped_mt`: already shipped
- `fg_mt`: in warehouse, ready to ship
- `wip_mt`: in production (has planned_end_date)
- `planned_mt`: has work order but no schedule
- `no_plan_mt`: remainder (SC qty - all above)

**Primary status label** (based on largest unfulfilled portion):
1. Shipped (fully or partially)
2. In Stock (FG available)
3. In Production - Scheduled (has planned_end_date)
4. Planned - Unscheduled (has work order, no date)
5. No Plan (nothing in PP)

---

## 5. Gap Calculation

**Applies only to SOs with BOTH**:
- A `planned_end_date` from PP (= MAX across all work orders for that SO)
- A `required_loading_date` from Loading Plan (shipping side)

**Formula**:
```
Gap (days) = Required_Loading_Date - (MAX(Planned_End_Date) + 1)
```

- `+1 day`: production finish → next day = earliest warehouse receipt = earliest loadable

**Interpretation**:
- Gap > 2 → Green: buffer exists
- Gap 0~2 → Yellow: tight but feasible
- Gap < 0 → Red: production behind schedule, |gap| = days of delay

**Risk tiers (for SOs WITHOUT a computable gap)**:
- Has PP schedule but no Loading Plan entry → Orange: production moving, but no shipping arrangement
- Has work order but no schedule (PP_unsched) → Red
- No work order at all → Critical Red

---

## 6. Output

### Excel
- Sheet 1: SO Master (full detail with status, quantities, gap, risk tier)
- Sheet 2: Risk Summary (aggregated by Plant × Cluster)
- Sheet 3: Gap Detail (only "In Production" SOs, sorted by gap ascending)
- Sheet 4: No Plan / Unscheduled (action list)

### HTML Dashboard
- Progress: 25,000 MT target vs current status waterfall
- Distribution: stacked bar by Plant (Shipped / FG / WIP / Planned / No Plan)
- Gap view: timeline or table sorted by urgency
- Filterable by Plant, Cluster, Machine

---

## 7. Logical Closed Loop

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
                  Gap = Deadline - Production
                         │
                         ▼
              Risk Signal → Action Required
```

Every SO in the baseline must land in exactly ONE of these paths. No order is left unaccounted for. The waterfall quantities must sum back to SC Vol.-MT. This is the MECE closure.
