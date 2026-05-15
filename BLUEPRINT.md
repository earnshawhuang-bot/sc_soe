# S&OE Order Fulfillment Tracking System - Blueprint

## 1. Purpose

For each current-month baseline sales order (SO), answer four management questions:
- What is this month's billing responsibility pool?
- Where is each baseline SO in the supply execution waterfall: shipped, FG, scheduled WIP, unscheduled WO, or no supply signal?
- Does each baseline SO have a current Loading Plan arrangement, and is the loading date confirmed?
- If the order cannot be fulfilled smoothly, which action queue does it belong to?

This is a **monthly recurring mechanism**: feed new data each month, output the same structured risk view. Target: ~25,000 MT/month across Kunshan (KS) and Indonesia (IDN).

**Data nature**: Shipped = MTD cumulative (e.g. May 1–7, next time May 1–10); FG = point-in-time snapshot on data extraction day. Both refresh each run cycle.

**Run anchor rule**:
- `data_date` is derived at runtime from the latest file suffix in `02-Shipped`, for example `GI (KS&IDN) 260501-0513.xlsx` -> `2026-05-13`.
- The report should not use the computer's current date as cutoff, because source data may lag behind the day when the report is opened.
- Output files are written under `output/<data_date>/` and use a minute-level run suffix such as `SOE_Tracking_2026-05-20260514-1531.xlsx` to avoid overwriting same-day reruns.

---

## 2. Data Flow (one paragraph)

SC table defines the **baseline**: what SOs belong to the current-month billing responsibility pool. Shipped, FG, and PP form the **supply execution waterfall**: shipped, in stock, scheduled production, unscheduled work order, or no supply signal. Loading Plan is a separate **shipping arrangement ledger**, not a peer status to Shipped / FG / PP. It is cross-checked against the baseline and supply waterfall to determine whether the order has a confirmed loading date, an unconfirmed loading signal, or no current loading arrangement. Valid loading dates are used for gap calculation; TBA / blank / invalid loading values remain as action risks rather than being dropped.

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

**Key columns**: Work Order No., SO, TotalWeight/T, Planned EndTime, PlannedFinishDate, Machine, Plant

`Planned EndTime` is the preferred planned-end field because it carries hour/minute/second precision and represents the planned Lami finish timestamp. `PlannedFinishDate` is retained only as a date-level reference or fallback when `Planned EndTime` is missing.

**Unscheduled (1 file)**:
- `Global_PP_wo schedule.xlsx`

**Key columns**: Work Order No., SO, TotalWeight/T, Machine, Plant (**no date**)

**Join**: PP.SO → SC.SO (direct match)

**Multi-work-order aggregation** (one SO can have 20+ work orders across machines/files):
- `wip_mt` = SUM of all work orders' TotalWeight under that SO
- `planned_end_date` = MAX(Planned EndTime) across all work orders under that SO
  - Fallback: use `PlannedFinishDate` only when `Planned EndTime` is missing
  - Rationale: the SO cannot be loaded until ALL sub-batches are complete; latest finish timestamp is the bottleneck

> ⚠️ **Pitfall**: PP SO numbers are also stored as `float64`. Same `int(float())` conversion required.

---

### 3.5 Loading Plan - Shipping Arrangement Ledger (06-Loading Plan)

Loading Plan is the shipping-arrangement ledger. It is not an execution status beside Shipped / FG / PP. It is used to identify whether a baseline SO has a confirmed loading date, an unconfirmed loading signal, or no current loading arrangement.

#### 3.5.1 KS Loading Plan

**File**: `06-Loading Plan/ks_loading plan/Loading plan-May.xlsx` → Sheet `"Loading plan"`

**Key columns**:
| Field | Usage |
|-------|-------|
| Invoice No | Split / normalize into SO |
| 20GP / 40GP / 40HQ | Convert container count into model MT |
| Loading | Raw loading date or unconfirmed text |
| MT | Source-table MT, kept for audit |
| Unnamed: 20 | Prior-invoiced marker such as `4月已开票`; excluded from current invoice scope but kept in audit |

**Container MT conversion**:
| Container | MT |
|-----------|----|
| 20GP | 14.5 |
| 40GP | 24.5 |
| 40HQ | 24.5 |

**Invoice No. parsing rules**:
- `_` separates multiple SOs.
- `-N` is a batch/container suffix and is stripped.
- `-N~M` is a multi-container range and is stripped to the same base SO.
- Short SO fragments inherit the prefix from the longest 10-digit SO in the invoice group.
- Non-standard `LM*` invoices are kept in parse-exception audit and excluded from SO matching.
- Tonnage is evenly allocated across parsed SOs when one source row maps to multiple SOs.

#### 3.5.2 IDN Export - Schedule Planning Dispatch

**File**: `06-Loading Plan/idn_loading plan/Schedule Planning Dispatch 260511.xlsx`

This workbook is now treated as two appendable Loading Plan sources:

| Sheet | Business meaning | Current-scope rule |
|-------|------------------|--------------------|
| `ORDER OUTSTANDING ` | Open / outstanding IDN export loading plan | Kept as current LP evidence, including TBA / blank / invalid ELD values as unconfirmed loading risks |
| `DISPATCH` | Already dispatched IDN export loading plan | Appended to IDN export LP so IDN has a full loading ledger; valid ELD rows are kept only from the run-month first day onward |

The `DISPATCH` sheet uses the same header family as `ORDER OUTSTANDING `. Some trailing columns may be missing; extraction should align by header name rather than fixed column count.

**Key columns**:
| Field | Usage |
|-------|-------|
| SC No. | Direct SO key, normalized from Excel numeric format |
| Region | Cluster reference |
| Cont Qty / Cont Size | Convert to MT using 14.5 / 24.5 rule |
| Rough Ton | Source-table MT, kept for audit |
| ELD | Raw loading date or unconfirmed value such as TBA |

`DISPATCH` inclusion rule:
- For a run month such as `2026-05`, only valid `ELD >= 2026-05-01` enters the current main LP ledger from `DISPATCH`.
- Valid `ELD` earlier than the run-month first day is treated as historical dispatched LP evidence and excluded from current management analysis.
- Blank / invalid `ELD` in `DISPATCH` should be retained in audit first; whether it enters the main current ledger requires business confirmation because the sheet is supposed to represent already dispatched orders.

> ⚠️ **Pitfall**: Sheet name has a trailing space: `"ORDER OUTSTANDING "`. Must match exactly. `DISPATCH` does not have this trailing-space requirement.

#### 3.5.3 IDN Domestic - New Domestic Tracking

**File**: `06-Loading Plan/idn_loading plan/NEW DOMESTIC TRACKING 260511.xlsx` → Sheet `"Order List"`, headers in **row 2**

**Key columns**:
| Field | Usage |
|-------|-------|
| SC NO. | Preferred SO key when valid |
| INV NO. | Fallback split source when SC NO. is not valid |
| ELD | Raw loading date or unconfirmed value |
| Weight | Loading MT |
| STATUS | Operational reference |

**IDN Domestic parsing rules**:
- Prefer valid `SC NO.`.
- If `SC NO.` is missing or invalid, split `INV NO.`.
- Comma and ampersand split multiple SOs.
- `-N` and `-N~M` suffixes are stripped.
- `LMIDSAM*` and other non-SO records remain in parse-exception audit.

#### 3.5.4 Main-Scope Filtering

Extraction keeps clean detail records at demand-line level and does not aggregate to SO. Main analysis then applies scope logic:
- Exclude explicit prior-invoiced records from current invoice scope, but retain them in audit.
- Exclude valid loading dates before the run-month start, but retain them as historical LP evidence.
- Keep TBA / blank / text-month / invalid-text loading values in main analysis because they are unconfirmed loading risks.
- Only in-scope valid loading dates participate in production-vs-loading gap calculation.

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
- **Risk Matrix Detail**: Segment-level fact table behind the Summary Shipped Data Closure and Open Fulfillment Risk Matrix
- **Action Required**: Action-required subset of Risk Matrix Detail, no longer generated from SO-level `Status`
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
  - Legacy `SCHEDULING` = allocated unscheduled + allocated no-plan volume; after Loading Plan redesign, this should be split into `Production Unscheduled` and `No Supply Signal` because they require different actions
- Risk Distribution columns: `Risk Tier`, `Volume (MT)`, `% of Baseline`.
- By Plant columns: `Plant`, `Baseline MT`, `Shipped`, `FG`, `WIP`, `Unscheduled`, `No Plan`, `Red+Critical MT`.
- Dashboard should keep a compact `BY PLANT` row-level summary table near the top: `Plant`, `Baseline MT`, `Shipped`, `FG`, `WIP`, `Unscheduled`, `No Plan`. It is a factory split of the first-layer baseline execution status, not a separate risk layer. Cells should remain numeric MT values, with units in the title or headers.
- Order Type Breakdown columns: `Order Type`, `Volume (MT)`, `Shipped`, `FG`, `WIP`, `Unscheduled`, `No Plan`.
- By Plant × Region (Cluster) columns: `Plant`, `Cluster`, `Baseline MT`, `Shipped`, `FG`, `WIP`, `Unscheduled`, `No Plan`, `Red+Critical MT`, `% Fulfilled`.
- By Plant × Region includes plant-level `TOTAL` rows before cluster-level detail rows.
- `% Fulfilled` = `(Shipped + FG) / Baseline MT`; WIP, Unscheduled, and No Plan remain future fulfillment exposure.

Loading Plan redesign note:
- Do not present `No Supply Signal` and `No Loading Arrangement` as the same thing.
- `Production Unscheduled` belongs to the supply execution dimension: a work order exists but has no planned finish date.
- `No Supply Signal` also belongs to the supply execution dimension: no shipped, FG, scheduled WIP, or unscheduled WO evidence exists.
- `No Loading Arrangement` belongs to the shipping arrangement dimension: no current-scope Loading Plan evidence exists.
- The Summary page should not add a standalone Loading Plan coverage block. It should show baseline execution first, then use the risk matrix to show how loading arrangement changes the action interpretation.

### Dashboard HTML
- `SOE_Dashboard_<month>-<run_suffix>.html` is the primary management-facing front end.
- The legacy narrative `SOE_Report_<month>-<run_suffix>.html` is no longer required once the dashboard carries the executive summary, matrix, and drill-down details.
- Visual style should use an executive-report palette: warm ivory background, deep navy titles, muted teal/blue and gold accents, and light warm-gray table rules rather than heavy saturated blue headers.

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
SC Baseline (current-month responsibility)
        |
        v
Supply Execution Waterfall
  - Shipped
  - FG
  - WIP Scheduled
  - WIP Unscheduled
  - No Supply Signal
        |
        v
Shipped Data Closure
  - Shipped quantity is audit / closure, not open fulfillment risk
  - Cross-check shipped quantity against current-month LP evidence
        |
        v
Open Fulfillment Risk Matrix
  rows    = open supply execution status: FG / WIP Scheduled / WIP Unscheduled / No Supply Signal
  columns = Past Due LP / Future Valid LP / LP Date Unconfirmed / No Current LP
        |
        v
Action Queue
  - FG without Loading Plan
  - Loading Plan without Supply Signal
  - WIP late vs Loading Date
  - LP Date Unconfirmed
  - Production Unscheduled
  - No Supply and No Loading Signal
```

Every SO in the baseline can have multiple raw source signals, but the allocated supply waterfall is mutually exclusive and sums back to adjusted SC baseline. This is the MECE closure used for management KPIs.

Loading Plan is not a separate execution bucket in the Summary page. It is the shipping-arrangement dimension used inside the risk matrix. Shipped quantity is handled as data closure: if it has matching LP evidence, the loop is closed; if not, it is a data consistency exception. Open quantity then enters the risk matrix.

`Past Due LP` must be defined by quantity, not by date alone:

```text
Past Due LP = valid LP quantity with loading_date < data_date
              minus the quantity already covered by shipped evidence
```

This avoids calling an old LP date "past due" when the same demand has already been shipped. `Past Due LP` is shown before `Future Valid LP` because it is already late and needs earlier management attention.

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
| 7 | IDN Schedule Planning Dispatch | `ORDER OUTSTANDING ` has a trailing space; `DISPATCH` is a second sheet in the same workbook and should be appended for IDN full LP coverage | Match `ORDER OUTSTANDING ` exactly; also read `DISPATCH`, align by headers, and keep only valid `ELD >= run-month first day` in current main scope |
| 8 | SC Loading Date (P) | Not reliable as a required loading deadline, but useful for the SC prior-month delivery pre-allocation window | Never use as fallback for gap calculation; only use for the explicit previous-month SC delivery filter |
| 9 | Raw source overlap | Same SO can appear in Shipped, FG, and PP simultaneously, causing raw sums to exceed baseline | Use allocated waterfall for KPIs; keep raw quantities in Overlap Audit |
| 10 | SC Delivery PCS | Negative delivery PCS rows reverse/correct delivery signals | Exclude negative Delivery PCS rows from SC prior-month delivery pre-allocation |

---

## 10. Loading Plan Shipping Readiness Blueprint

### 10.1 High-Level Purpose

The model is upgraded from a pure order execution tracker into a **monthly billing-order shipping fulfillment model**.

It answers one page-level decision mission:

```
Can the current-month SC baseline orders be fulfilled, and if not, where is the action risk?
```

The management story is intentionally baseline-led:
- First, use SC Baseline as the denominator and classify where the goods are in the supply execution waterfall.
- Second, use the risk action matrix to cross-check the supply execution status with the Loading Plan shipping-arrangement signal.
- Third, route exceptions into action queues such as FG without Loading Plan, Loading Plan without supply signal, production missing loading date, and unconfirmed loading date.

The key design principle is:

```
SC Baseline = current-month billing responsibility
Loading Plan = shipping arrangement / open delivery ledger
Shipped / FG / PP = supply execution status
```

Loading Plan is **not** a peer status to Shipped / FG / PP. It is a separate shipping-arrangement dimension that must be cross-checked against supply readiness.

Summary and management views must therefore avoid mixing these dimensions into one flat status list. They also should not introduce Loading Plan as a standalone Summary block. The correct narrative is:

```
Baseline responsibility -> supply execution status -> risk action matrix -> action queue
```

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
- Sheets:
  - `ORDER OUTSTANDING `, with trailing space: open / outstanding export loading plan.
  - `DISPATCH`: already dispatched export loading plan; append to the same IDN export LP stream.

Key fields:
| Field | Usage |
|-------|-------|
| SC No. | Direct SO key, normalized from Excel numeric format |
| Region | Cluster reference |
| Cont Qty / Cont Size | Convert to MT using the same 14.5 / 24.5 rule |
| Rough Ton | Source-table MT for audit only |
| ELD | Loading date or unconfirmed loading value such as TBA |

IDN export append rules:
- `ORDER OUTSTANDING ` and `DISPATCH` share the same header family. `DISPATCH` may miss trailing columns; align by semantic header names.
- Add source lineage fields so the output can distinguish `lp_source = IDN_Export` and `source_sheet = ORDER OUTSTANDING / DISPATCH`.
- `DISPATCH` contains historical rows. For current management analysis, keep valid `ELD` rows only from the run-month first day onward, for example `ELD >= 2026-05-01` for a May 2026 run.
- Earlier valid `DISPATCH` ELD rows are historical dispatched LP evidence and stay available for audit rather than current open-risk analysis.
- Blank / invalid `DISPATCH` ELD rows should be retained in audit first; whether they enter main current scope needs business confirmation.

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
Clean detail keeps all Loading Plan records for audit.
Main analysis excludes historical valid loading dates before the run-month start and explicit prior-invoiced records.
TBA / blank / text-month / invalid-text records remain in main analysis as unconfirmed loading risks.
Only in-scope Valid Date records participate in gap calculation.
```

For a run month such as `2026-05`, a parsed valid loading date before `2026-05-01` is treated as historical Loading Plan evidence. It is retained in audit detail but does not participate in the current-month SC vs LP reconciliation or Shipping Readiness main view. A valid loading date in the run month or after the run month remains in scope because it can still explain whether the current baseline order is arranged, late, or pushed out.

TBA, blank, text-month, and invalid-text loading values are not dropped. They are treated as unconfirmed shipping risks because the business confirmed these values usually represent orders that still require attention.

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
| In SC and In LP | Current billing SO has positive-MT, main-scope Loading Plan arrangement |
| In SC only | Current billing SO has no positive-MT, main-scope Loading Plan match |
| In LP only | Main-scope Loading Plan has SO not in current SC baseline |
| LP excluded - prior invoiced | Excluded from current billing scope but retained in audit |

Important interpretation:
- `In SC only` does **not** always mean the SO never appeared anywhere in the raw Loading Plan. It means the SO has no usable current-scope LP quantity after parsing, prior-month valid-date exclusion, prior-invoiced exclusion, and zero-MT handling.
- If an SO has TBA / blank / invalid loading date but has positive MT and can be parsed, it remains `In SC and In LP` and is flagged as unconfirmed loading.
- Rows where both SC MT and LP MT are zero should be removed from the management reconciliation output because they do not represent an actionable business population.

Recommended reconciliation support fields:
| Field | Purpose |
|-------|---------|
| lp_loading_date_raw_list | Show original Loading / ELD values from source files |
| lp_earliest_valid_loading_date | Earliest parsed valid loading date, if any |
| lp_loading_date_status_mix | Date quality mix such as `Valid Date`, `TBA`, `Blank`, `Invalid Text` |
| lp_match_scope | Current LP / Historical LP only / Excluded LP only / No LP evidence |
| lp_line_count | Count of source LP lines supporting the SO |
| lp_parse_exception_flag | Indicates possible missed match due to invoice/SO parsing issue |

The key management interpretation:
- `In SC only + FG`: goods are ready but no loading arrangement exists.
- `In SC only + WIP`: production is planned but shipping is not arranged yet.
- `In SC only + No Plan`: both production and shipping arrangement are missing.
- `In LP only`: the loading ledger contains demand outside the current billing baseline and must be reviewed separately.

---

### 10.5 Shipping Readiness Matrix

Shipping readiness cross-checks baseline supply status against loading arrangement coverage. The Summary view is split into two pieces so shipped quantity does not distort open fulfillment risk.

#### 10.5.1 Shipped Data Closure

`Shipped` quantity is already fulfilled. It should not be mixed with open risk rows. It is reviewed as data closure:

| Supply Status | Total MT | LP Closed / Matched | LP Missing or Unclear |
|---------------|---------:|--------------------:|----------------------:|
| Shipped | shipped baseline quantity | shipped quantity with current-month LP evidence | shipped quantity without usable LP evidence |

This block answers whether shipped baseline quantity has a matching LP trail. The new IDN `DISPATCH` sheet is important here because it contains IDN export records that may already have shipped and therefore should close the LP loop for shipped IDN orders.

#### 10.5.2 Open Fulfillment Risk Matrix

Open quantity excludes shipped quantity and focuses on what still needs action:

| Supply Status | Past Due LP | Future Valid LP | LP Date Unconfirmed | No Current LP |
|---------------|-------------|-----------------|---------------------|---------------|
| FG | Goods are ready but the planned loading date has passed and is not closed by shipped evidence | Ready to ship; check whether loading date is reasonable | Goods ready but shipping date unconfirmed | **FG without Loading Plan** |
| WIP Scheduled | Production exists, but the loading date has passed and is not closed by shipped evidence | Check whether planned finish can meet loading date | Production exists but shipping date unconfirmed | Production scheduled but no shipping arrangement |
| WIP Unscheduled | Loading date has passed, but production has no finish date | Shipping demand exists but production has no finish date | Production and shipping date are both uncertain | Production and shipping are both unconfirmed |
| No Supply Signal | **Past-due loading arrangement but no supply signal** | **Loading arranged but no supply signal** | LP unconfirmed and no supply signal | Baseline order has no supply or shipping signal |

Definitions:
- `Supply Status` comes from the allocated baseline waterfall: Shipped -> FG -> WIP Scheduled -> WIP Unscheduled -> No Supply Signal.
- `Past Due LP` means valid LP quantity with `loading_date < data_date` that is not already covered by shipped quantity.
- `Future Valid LP` means valid LP quantity with `loading_date >= data_date` for remaining open demand.
- `LP Date Unconfirmed` means the SO has LP evidence but loading date is TBA / blank / text-month / invalid.
- `No Current LP` means no positive-MT LP coverage remains after current-scope filtering.
- LP status should be allocated at quantity level where possible. Do not simply paste one SO-level LP status onto every supply segment if that would duplicate LP coverage.

Gap logic remains:

```
Available Date = Planned EndTime + 1 day
LP Gap Days = Loading Date - Available Date
```

Because current LP loading dates are mostly date-level, gap should be compared at date grain for now. If LP later carries explicit loading timestamps, the model can upgrade to timestamp-level comparison. This gap is computed only for open WIP Scheduled quantity with valid loading date.

---

### 10.6 Excel Output Additions

The Excel workbook includes these Loading Plan sheets:
- `Loading Plan Clean Detail`
- `SC vs LP Reconciliation`
- `Shipping Readiness`
- `LP Date Exceptions`
- `LP Parse Exceptions`
- `LP Excluded Prior Invoiced`

Summary page decision mission:

```
Can current-month baseline orders be fulfilled, and what needs management action first?
```

The Summary page is baseline-led. It should not present Loading Plan as another execution status beside Shipped / FG / WIP, and it should not add a separate standalone Loading Plan coverage section. Loading Plan appears inside the risk matrix as the shipping-arrangement axis.

Block 1 - Baseline Execution Status:
| Metric | Purpose |
|--------|---------|
| Baseline MT | Current-month responsibility denominator |
| Shipped MT | Already fulfilled |
| FG MT | Goods ready, shipping execution required |
| WIP Scheduled MT | Production has planned finish date |
| WIP Unscheduled MT | Work order exists, finish date missing |
| No Supply Signal MT | No shipped, FG, scheduled WIP, or unscheduled WO signal |

Block 2 - Shipped Data Closure:
| Supply Status | Total MT | LP Closed / Matched | LP Missing or Unclear |
|---------------|---------:|--------------------:|----------------------:|
| Shipped | numeric shipped MT | numeric matched MT | numeric exception MT |

This is an audit / closure block, not the open action-risk matrix. It answers whether already shipped baseline volume can be tied back to current-month LP evidence. The IDN `DISPATCH` append is part of this closure design.

Block 3 - Open Fulfillment Risk Matrix:
| Supply Status | Supply Status Total | Past Due LP | Future Valid LP | LP Date Unconfirmed | No Current LP |
|---------------|--------------------:|------------:|----------------:|--------------------:|--------------:|
| FG | numeric MT | numeric MT | numeric MT | numeric MT | numeric MT |
| WIP Scheduled | numeric MT | numeric MT | numeric MT | numeric MT | numeric MT |
| WIP Unscheduled | numeric MT | numeric MT | numeric MT | numeric MT | numeric MT |
| No Supply Signal | numeric MT | numeric MT | numeric MT | numeric MT | numeric MT |
| Open Total | numeric MT | numeric MT | numeric MT | numeric MT | numeric MT |

Supporting drill-down fields may still include historical LP evidence, LP-only MT, raw loading date lists, and parse exception flags. They should support investigation, but they should not become a separate Summary narrative layer.

Primary sorting anchor for management action is MT volume, not SO count. SO count can be used as supporting context in drill-down tables.

Summary matrix formatting rule:
- Matrix cell values must be numeric MT values, not text strings.
- Do not append `MT` or SO counts inside matrix cells.
- Add a numeric `Supply Status Total` column immediately after `Supply Status`; it equals the row sum across Past Due LP, Future Valid LP, LP Date Unconfirmed, and No Current LP.
- Place `Past Due LP` before `Future Valid LP` because already-late loading commitments should be reviewed first.
- Express the unit in the section title or header, for example `Open Fulfillment Risk Matrix (MT)`.
- This keeps the matrix directly usable for Excel formulas, row totals, column totals, and audit checks back to baseline.

#### Risk Matrix Detail as the fact table

`Risk Matrix Detail` is the single fact table behind the Summary risk matrix and the Action Required sheet.

Grain:

```
one SO + one allocated supply segment + one Loading Plan coverage segment
```

This means one SO may appear multiple times if its baseline quantity is split across Shipped, FG, WIP Scheduled, WIP Unscheduled, and No Supply Signal. It may also split by LP coverage status when only part of the quantity has past-due LP, future valid LP, unconfirmed LP, or no LP. This avoids overstating the whole SO as a risk when only part of its quantity is exposed.

Recommended fields:
| Field | Meaning |
|-------|---------|
| so | Sales order |
| plant / cluster / order_type | Management dimensions |
| so_total_mt | Full adjusted baseline quantity for context |
| supply_status | Shipped / FG / WIP Scheduled / WIP Unscheduled / No Supply Signal |
| lp_coverage_status | Past Due LP / Future Valid LP / LP Date Unconfirmed / No Current LP / Shipped LP Closed |
| risk_mt | Allocated MT for this segment; the primary action anchor |
| covered_mt | SO quantity already covered by better supply segments, for context |
| shipped_closed_mt | LP quantity already closed by shipped evidence, for shipped audit and past-due calculation |
| lp_loading_date_raw_list | Original Loading / ELD values |
| lp_earliest_valid_loading_date | Earliest valid current-scope loading date |
| planned_end_date | Raw PP Planned EndTime, aggregated to latest timestamp per SO; represents Lami finish time |
| available_date | `planned_end_date + 1 day`; current available-to-ship reference timestamp used for gap checks |
| lp_gap_days | Loading date minus available date, when computable; current comparison is date-grain if LP has no time |
| risk_action | Business action label such as `FG without Loading Plan` |
| action_required | Boolean flag used to derive Action Required |
| suggested_owner | Suggested owner such as Logistics / Planning / Plant / Sales Ops |
| action_note | Short business explanation |

Dashboard display rules:

- Excel / `Risk Matrix Detail` remains a stable fact table with complete fields for filtering, pivoting, and audit.
- The dashboard detail area is a business workbench; after a matrix cell is clicked, detail columns may change by scenario.
- `planned_end_date / available_date / machines / work_orders` are shown only for `WIP Scheduled`.
- `loading_date` is shown only for `Past Due LP / Future Valid LP`; it is hidden for `No Current LP`.
- `LP Date Unconfirmed` does not show a formal `loading_date`; it shows `lp_loading_date_raw_list / lp_loading_date_status_mix` instead.
- `lp_gap_days` is calculated and shown only for `WIP Scheduled + valid loading date`, and highlighted in the dashboard.
- `risk_mt / risk_action / suggested_owner / action_note` are shown in every dashboard detail scenario.
- For `WIP Scheduled`, the business reading order is: `Risk MT -> Machine -> Planned End -> Available Date -> Loading Date -> Gap -> Risk Action -> Owner -> Work Order -> Action Note`.
- PP inputs should use `Planned EndTime`; Excel and dashboard outputs should preserve at least minute-level display. Seconds can remain in the data layer while the UI shows minutes.
- `Available Date = Planned EndTime + 1 day`; current LP date-grain means gap is compared by date, and can later move to timestamp-grain when LP carries time.
- The dashboard `Selected Risk Detail` area should provide a CSV download button. It exports the current Plant filter and selected matrix cell detail; if no cell is selected, it exports the default action-required detail.
- The dashboard HTML remains a standalone static snapshot that can be opened directly in Edge / Chrome, but it embeds detail data and must be distributed with access control in mind.

#### LP Not In Current SC Baseline

`Open Fulfillment Risk Matrix` is baseline-led; its denominator is current-month SC baseline open quantity.

`LP Not In Current SC Baseline` is LP-led reconciliation exception reporting. Its denominator is current-scope Loading Plan quantity that cannot be matched back to the current-month SC baseline. It should not be mixed into the fulfillment risk matrix; it should be shown as a separate third block.

Summary structure:

| LP Status | MT |
|-----------|---:|
| Past Due LP | numeric MT |
| Future Valid LP | numeric MT |
| LP Date Unconfirmed | numeric MT |
| Total | numeric MT |

Do not show SO Count in this block. Management should first read MT volume; SO count can be derived from the detail table or CSV if needed.

Detail grain:

```
one current-scope Loading Plan detail line
```

Recommended dashboard / Excel detail fields:

```
SO / Plant / LP Source / Source Sheet / Invoice No Raw
Loading Date Raw / Loading Date / LP Date Status / Load MT
```

This block answers whether Loading Plan contains future-month orders, missing baseline orders, cross-month arrangements, or loading demand that still needs business ownership clarification.

Derived outputs:
| Output | Source | Purpose |
|--------|--------|---------|
| Summary Shipped Data Closure | Pivot shipped rows / shipped closure fields from `Risk Matrix Detail` | Show whether shipped quantity has LP evidence |
| Summary Open Fulfillment Risk Matrix | Pivot non-shipped rows from `Risk Matrix Detail` | Show MT by open supply status and LP coverage status |
| Dashboard LP Not In Current SC Baseline | Summarize `lp_not_in_baseline_detail` | Show LP-led reconciliation exceptions outside the baseline fulfillment matrix |
| Action Required | Filter `Risk Matrix Detail` where `action_required = True` | Show only rows requiring follow-up |

`Action Required` should therefore not be generated from SO-level `Status`. SO-level `Status` can remain in `SO Master` as a worst-exposure signal, but management actions should be driven by `risk_mt` at segment level.

---

### 10.7 Phase 2 Roadmap: Data Consumption Layer

Excel, HTML, and Power BI are presentation layers. The system should later output a stable machine-readable data mart so all presentation layers consume the same business logic.

Phase 2 target structure:

```
output/<run>/data_mart/
  clean_loading_plan_lines.csv
  sc_lp_reconciliation.csv
  risk_matrix_detail.csv
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
