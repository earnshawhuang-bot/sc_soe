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

**File**: monthly `Order tracking *.xlsx` → Sheet "Order Status - Main"

**Filter**: Item column = "main"

**Key columns** (use normalized header lookup; column positions can shift month to month):
| Field | Usage |
|-------|-------|
| SC NO | SO number (primary key) |
| Supply From | Plant: Kunshan→KS, Indonesia→IDN |
| Cluster | Region dimension |
| End User Cust ID | Customer mapping validation |
| SC Vol.-MT | Raw order quantity (MT) |
| Order Status | Classification input |
| Carryover | **Null** = this month's billing |
| carryover/fresh production | "last month" = Carry Over Unproduced |
| RELEASE DATE | Month determination |
| Delivery PCS | SC-side delivery signal; negative rows are excluded from prior-month delivery pre-allocation |
| Loading Date | SC-side loading date; not used as the required loading deadline |

**Customer mapping validation**:
- File: `07-Mapping/dim_cc_region.xlsx`
- Sheet: `dim->new cc`
- Join: `SC.End User Cust ID` → `dim->new cc.new customer code`
- End customer unmatched rows are excluded from the main baseline and stored in the `Unmatched End Customer` audit sheet.
- `Material Code` mapping is not used as a baseline filter in v2.

**Classification (4 types, independently verified MECE)**:
1. **Carry Over Stock**: Order Status = `"Carryover"` AND Carryover is null
2. **Carry Over Unproduced**: carryover/fresh production = `"last month"` AND Carryover is null
3. **Fresh Order This Month**: Release Date in current month AND Q is null
4. **Fresh Order Next Month**: Release Date in current month AND Q has value → **excluded from baseline**

**Business adjustment factor**:
| Order Type | Adjusted baseline volume |
|------------|--------------------------|
| Carry Over Stock | `SC Vol.-MT` |
| Carry Over Unproduced | `SC Vol.-MT × 0.975` |
| Fresh Order This Month | `SC Vol.-MT × 0.975` |
| Fresh Order Next Month | `SC Vol.-MT × 0.975` for next-month visibility only; excluded from current baseline |

**Baseline** = Carry Over Stock + adjusted Carry Over Unproduced + adjusted Fresh Order This Month.

**Aggregation**: Classify and adjust at SC row level first, then aggregate to SO. Preserve type-specific adjusted columns (`carry_over_stock_mt`, `carry_over_unproduced_mt`, `fresh_this_month_mt`) so mixed-type SOs do not get misallocated to the first row's type.

**SC prior-month delivery pre-allocation**:
- Build a separate SC-derived delivery DataFrame from the same `Order Status - Main` rows.
- Date window: previous calendar month relative to the run `data_date`.
  - Example: if `data_date = 2026-05-11`, use `2026-04-01` through `2026-04-30`.
- Filter:
  - `Item = main`
  - `Loading Date` falls in the previous-month window
  - `Delivery PCS` is not negative; negative delivery rows are removed
  - End customer mapping passes the same `End User Cust ID` validation
- Quantity:
  - Use `SC Vol.-MT` as the delivery volume field.
- Purpose:
  - This SC-derived quantity is treated as already fulfilled and is allocated into `Allocated Shipped` before normal 02-Shipped GI data.
  - It is an internal SC pre-allocation signal; it does not replace 02-Shipped.

> ⚠️ **Pitfall**: K column value is `"Carryover"` (single word, capital C) — NOT `"carry over"` or `"Carry Over"`. Matching against the wrong string returns 0 rows for Carry Over Stock.

> ⚠️ **Pitfall**: Q column (Carryover) holds a numeric pcs value when filled, not text. "Empty" means `pd.isna()` — do NOT check for empty string.

> ⚠️ **Pitfall**: AO column header contains a newline: `"carryover/fresh\nproduction"`. Normalize header names instead of relying on either fixed position or exact raw text.

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

**Status assignment uses a mutually exclusive quantity waterfall**.

Raw source quantities are preserved as `raw_sc_prior_delivery_mt`, `raw_shipped_mt`, `raw_fg_mt`, `raw_wip_mt`, and `raw_unsched_mt`. Management KPIs and status use allocated quantities:

```
allocated_sc_prior_shipped = min(raw_sc_prior_delivery, adjusted_sc_vol)
remaining_0                = adjusted_sc_vol - allocated_sc_prior_shipped

allocated_shipped = min(raw_shipped, remaining_0)
remaining_1       = remaining_0 - allocated_shipped

allocated_fg      = min(raw_fg, remaining_1)
remaining_2       = remaining_1 - allocated_fg

allocated_wip     = min(raw_wip, remaining_2)
remaining_3       = remaining_2 - allocated_wip

allocated_unsched = min(raw_unsched, remaining_3)
allocated_no_plan = remaining_3 - allocated_unsched
```

This guarantees:

```
allocated_sc_prior_shipped
+ allocated_shipped
+ allocated_fg
+ allocated_wip
+ allocated_unsched
+ allocated_no_plan
= adjusted SC baseline
```

`Allocated Shipped` in management reporting equals `allocated_sc_prior_shipped + allocated_shipped`. The split is retained in SO Master so users can distinguish SC prior-month delivery pre-allocation from normal GI shipped data.

**Primary status label** follows the most severe remaining allocated segment:
1. `allocated_no_plan > 0` → No Plan
2. `allocated_unsched > 0` → Planned (Unscheduled)
3. `allocated_wip > 0` → In Production
4. `allocated_fg > 0` → In Stock
5. fully shipped → Shipped

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

**Risk tier logic used in the Summary sheet**:
| Risk Tier | Trigger |
|-----------|---------|
| Green | Fully shipped, or in production with gap > 2 days |
| Yellow | In stock / partial shipped, or in production with gap 0-2 days |
| Orange | Allocated WIP exists but no loading plan date |
| Red | Allocated unscheduled work order exists, or production gap < 0 |
| Critical | Allocated no-plan quantity remains after shipped, FG, WIP, and unscheduled |

`Red+Critical MT` in Summary tables is the adjusted baseline volume of SOs whose final `risk_tier` is Red or Critical. It is not an SO count, and it is not simply `Unscheduled + No Plan`; Red can also include scheduled WIP with a negative production gap.

---

## 6. Output

### Excel (v7 sheets, consulting-style formatted)
- **Summary**: Banner + KPI cards + Risk Distribution + By Plant + Risk Tier Logic + Order Type Breakdown + **Plant × Region (Cluster) breakdown**
- **SO Master**: Full detail per SO, including adjusted baseline, SC prior-month delivery pre-allocation, raw source quantities, allocated waterfall quantities, gap, and risk
- **Gap Analysis**: In-Production SOs only, sorted by gap ascending, gap cells color-coded
- **Action Required**: No Plan + Unscheduled SOs, urgent red banner
- **Overlap Audit**: SOs where raw source quantities exceed adjusted baseline; used for explanation, not management KPI
- **SC Row Detail / SC Fresh Next Month / SC Unknown Type / Unmatched End Customer**: SC baseline audit sheets

**Summary sheet presentation rules**:
- All volume metrics are MT, shown with thousands separators and no decimals.
- SO counts are intentionally not shown in the Summary tables; the management view is volume-driven.
- KPI cards:
  - `BASELINE ORDERS` = adjusted SC baseline volume
  - `SHIPPED` = allocated shipped volume, including SC prior-month delivery pre-allocation and normal GI shipped
  - `IN STOCK (FG)` = allocated FG volume
  - `IN PRODUCTION` = allocated WIP volume
  - `SCHEDULING` = allocated unscheduled + allocated no-plan volume
- Risk Distribution columns: `Risk Tier`, `Volume (MT)`, `% of Baseline`.
- By Plant columns: `Plant`, `Baseline MT`, `Shipped`, `FG`, `WIP`, `Unscheduled`, `No Plan`, `Red+Critical MT`.
- Order Type Breakdown columns: `Order Type`, `Volume (MT)`, `Shipped`, `FG`, `WIP`, `Unscheduled`, `No Plan`.
- By Plant × Region (Cluster) columns: `Plant`, `Cluster`, `Baseline MT`, `Shipped`, `FG`, `WIP`, `Unscheduled`, `No Plan`, `Red+Critical MT`, `% Fulfilled`.
- By Plant × Region includes plant-level `TOTAL` rows before cluster-level detail rows.
- `% Fulfilled` = `(Shipped + FG) / Baseline MT`; WIP, Unscheduled, and No Plan remain future fulfillment exposure.

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

Every SO in the baseline can have multiple raw source signals, but the allocated waterfall is mutually exclusive and sums back to adjusted SC baseline. This is the MECE closure used for management KPIs.

---

## 9. Known Pitfalls & Lessons Learned

| # | Where | Issue | Fix |
|---|-------|-------|-----|
| 1 | SC K column | Value is `"Carryover"` not `"carry over"` — wrong string = 0 matches, entire Carry Over Stock category lost | Match exact string `"Carryover"` |
| 2 | SC Q column | Holds numeric pcs value when filled; "empty" = `pd.isna()`, not empty string check | Use `.isna()` |
| 3 | SC AO column | Header has embedded newline `\nproduction`; exact-name lookup and fixed positions can both drift | Normalize headers and match semantically |
| 4 | SC rows | One SO can have multiple rows and even multiple order types | Classify/adjust at row level, then aggregate to SO with type-specific columns |
| 5 | All numeric IDs | SO/plant codes stored as `float64` in Excel (e.g. `1000011733.0`) — `.astype(str)` gives `"1000011733.0"`, breaks regex | Convert via `str(int(float(value)))` |
| 6 | Shipped SO column | Two columns named "销售订单"; pandas deduplicates to `.1`; P column has internal order numbers, Q column (`.1`) has SC numbers | Prefer the column where values start with `"10"` |
| 7 | IDN Dispatch sheet | Sheet name has trailing space: `"ORDER OUTSTANDING "` | Match with trailing space |
| 8 | SC Loading Date (P) | Not reliable as a required loading deadline, but useful for the SC prior-month delivery pre-allocation window | Never use as fallback for gap calculation; only use for the explicit previous-month SC delivery filter |
| 9 | Raw source overlap | Same SO can appear in Shipped, FG, and PP simultaneously, causing raw sums to exceed baseline | Use allocated waterfall for KPIs; keep raw quantities in Overlap Audit |
| 10 | SC Delivery PCS | Negative delivery PCS rows reverse/correct delivery signals | Exclude negative Delivery PCS rows from SC prior-month delivery pre-allocation |

---

## 10. Loading Plan Shipping Readiness Blueprint

### 10.1 High-Level Purpose

The model is upgraded from a pure order execution tracker into a **monthly billing-order shipping fulfillment model**.

It answers four management questions:
- Which SC baseline orders are this month's billing responsibility?
- Which of those orders already have a Loading Plan arrangement?
- For all Loading Plan demand, is the shipping arrangement confirmed or still unconfirmed?
- Does Shipped / FG / PP evidence support the loading demand, and where is the risk?

The key design principle is:

```
SC Baseline = current-month billing responsibility
Loading Plan = shipping arrangement / open delivery ledger
Shipped / FG / PP = supply execution status
```

Loading Plan is **not** a peer status to Shipped / FG / PP. It is a separate shipping-arrangement dimension that must be cross-checked against supply readiness.

---

### 10.2 Loading Plan Source Tables

#### KS Loading Plan

Source:
- Folder: `06-Loading Plan/ks_loading plan/`
- File: `Loading plan-May.xlsx`
- Sheet: `Loading plan`

Key fields:
| Field | Usage |
|-------|-------|
| Invoice No | Split / normalize into SO |
| 20GP / 40GP / 40HQ | Convert container count into MT |
| Loading | Required loading date or unconfirmed loading text |
| MT | Source-table MT for audit only |
| Unnamed: 20 | Current observed value: `4月已开票`; excluded from current invoice scope but kept in audit |

Container MT conversion:
| Container | MT |
|-----------|----|
| 20GP | 14.5 |
| 40GP | 24.5 |
| 40HQ | 24.5 |

KS SO parsing:
- `_` separates multiple SOs.
- `-N` is a batch/container suffix and is stripped.
- `-N~M` is a multi-container range and is stripped to the base SO.
- Short SO fragments inherit the prefix from the longest 10-digit SO in the invoice group.
- Non-standard `LM*` invoices are kept in parse-exception audit and excluded from SO matching.

#### IDN Export Loading Plan

Source:
- Folder: `06-Loading Plan/idn_loading plan/`
- File: `Schedule Planning Dispatch 260511.xlsx`
- Sheet: `ORDER OUTSTANDING `, with trailing space.

Key fields:
| Field | Usage |
|-------|-------|
| SC No. | Direct SO key, normalized from Excel numeric format |
| Region | Cluster reference |
| Cont Qty / Cont Size | Convert to MT using the same 14.5 / 24.5 rule |
| Rough Ton | Source-table MT for audit only |
| ELD | Loading date or unconfirmed loading value such as TBA |

#### IDN Domestic Loading Plan

Source:
- Folder: `06-Loading Plan/idn_loading plan/`
- File: `NEW DOMESTIC TRACKING 260511.xlsx`
- Sheet: `Order List`, header row 2.

Key fields:
| Field | Usage |
|-------|-------|
| SC NO. | Preferred SO key when valid |
| INV NO. | Fallback split source when SC NO. is not valid |
| ELD | Loading date or unconfirmed loading value |
| Weight | Loading MT |
| STATUS | Operational reference |

IDN Domestic SO parsing:
- Prefer valid `SC NO.`.
- If `SC NO.` is missing or invalid, split `INV NO.`.
- Comma and ampersand split multiple SOs.
- `-N` and `-N~M` suffixes are stripped.
- `LMIDSAM*` and other non-SO records remain in parse-exception audit.

---

### 10.3 Clean Loading Plan Demand-Line Output

The cleaned Loading Plan layer preserves demand-line granularity. It does **not** aggregate to SO during extraction.

Each row represents:

```
one plant + one source + one source row + one parsed SO + one loading demand
```

Standard fields:
| Field | Meaning |
|-------|---------|
| plant | KS / IDN |
| lp_source | KS_LP / IDN_Export / IDN_Domestic |
| source_file / source_sheet / source_row | Traceability back to source workbook |
| invoice_no_raw | Raw invoice / SC key text |
| so | Parsed SO |
| so_parse_status | Parsed / Non-SC / Parse Failed |
| loading_date_raw | Original Loading / ELD value |
| loading_date | Parsed date when valid |
| loading_date_status | Valid Date / TBA / Blank / Text Month / Invalid Text |
| load_mt | Business-rule MT used by the model |
| source_mt | Source workbook MT for audit comparison |
| exclude_from_current_invoice | True for records such as `4月已开票` |
| exclude_reason | Raw exclusion reason |

Important rule:

```
All non-excluded Loading Plan records participate in cross-validation.
Only Valid Date records participate in gap calculation.
```

TBA, blank, text-month, and invalid-text loading values are not dropped. They are treated as unconfirmed shipping risks.

---

### 10.4 SC vs Loading Plan Reconciliation

This layer compares two different scopes:

```
SC Baseline = current-month billing orders
Loading Plan = sales-side open delivery / loading arrangement ledger
```

Reconciliation categories:
| Category | Meaning |
|----------|---------|
| In SC and In LP | Current billing SO has loading arrangement |
| In SC only | Current billing SO has no loading arrangement |
| In LP only | Loading Plan has SO not in current SC baseline |
| LP excluded - prior invoiced | Excluded from current billing scope but retained in audit |

The key management interpretation:
- `In SC only + FG`: goods are ready but no loading arrangement exists.
- `In SC only + WIP`: production is planned but shipping is not arranged yet.
- `In SC only + No Plan`: both production and shipping arrangement are missing.
- `In LP only`: the loading ledger contains demand outside the current billing baseline and must be reviewed separately.

---

### 10.5 Shipping Readiness Matrix

Shipping readiness cross-checks Loading Plan demand against supply evidence:

| Loading Plan State | Supply Evidence | Interpretation |
|--------------------|-----------------|----------------|
| Valid Date | Shipped | Covered by Shipped |
| Valid Date | FG | Covered by FG |
| Valid Date | WIP on time | Production can support loading |
| Valid Date | WIP late | Production misses loading date |
| Valid Date | Unscheduled WO | Work order exists but completion date is missing |
| Valid Date | No supply | Loading arranged but no supply signal |
| TBA / Blank / Text / Invalid | Any supply state | Loading date is unconfirmed; still participates in risk review |

Gap logic remains:

```
LP Gap Days = Loading Date - (Planned End Date + 1 day)
```

But this gap is computed only when Loading Date is valid.

---

### 10.6 Excel Output Additions

The Excel workbook includes these Loading Plan sheets:
- `Loading Plan Clean Detail`
- `SC vs LP Reconciliation`
- `Shipping Readiness`
- `LP Date Exceptions`
- `LP Parse Exceptions`
- `LP Excluded Prior Invoiced`

Summary also includes a Loading Plan readiness panel:
- SC Baseline
- In Loading Plan
- SC without Loading Plan
- Valid Loading Date
- Unconfirmed Loading Date
- LP with Supply Risk
- Past Loading Not Shipped

---

### 10.7 Phase 2 Roadmap: Data Consumption Layer

Excel, HTML, and Power BI are presentation layers. The system should later output a stable machine-readable data mart so all presentation layers consume the same business logic.

Phase 2 target structure:

```
output/<run>/data_mart/
  clean_loading_plan_lines.csv
  sc_lp_reconciliation.csv
  shipping_readiness.csv
  risk_summary.csv
  lp_date_exceptions.csv
  lp_parse_exceptions.csv
  lp_excluded_prior_invoiced.csv
  manifest.yaml
  data_dictionary.md
```

Recommended storage:
- CSV as the first machine-readable table output because it is easy for Excel, Power BI, and HTML tooling to consume.
- YAML or JSON only for run metadata such as source files, row counts, and run parameters.
- Parquet or SQLite can be considered later if data volume or querying needs grow.
