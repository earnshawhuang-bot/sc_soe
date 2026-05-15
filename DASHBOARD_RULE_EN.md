# S&OE Dashboard Reading Guide

## 1. What This Dashboard Is For

This HTML dashboard is designed to answer one management question:

> Can the current-month SC baseline orders be fulfilled and shipped? If not, where is the risk stuck, and who should follow up?

It is not just a data display page. It is an action-management page.
When reading it, do not start with "how many SOs do we have?" Start with:

```text
How many MT have already been shipped?
How many MT are already ready to ship?
How many MT are still in production or have no supply signal?
Do the open MT have a Loading Plan?
If there is a Loading Plan, is the date confirmed, overdue, and feasible against production?
```

The dashboard is baseline-led. The starting point is the current-month SC baseline responsibility pool. Loading Plan is not another execution status. It is the shipping-arrangement dimension used to cross-check whether the fulfillment loop is closed.

## 2. Three Core Definitions

### 2.1 Data Date

`Data as of` is the cut-off date of the current data package.

The current rule derives it from the latest GI file suffix in the `02-Shipped` folder. Example:

```text
GI (KS&IDN) 260501-0513.xlsx -> data_date = 2026-05-13
```

It is not the computer's current date.
Business data usually has a refresh rhythm, so the dashboard should use the actual cut-off date of the source data.

### 2.2 Baseline

`Baseline MT` is the current-month billing responsibility pool. It is the order volume that management expects to fulfill in the current month.

It comes from SC baseline, not from Loading Plan.
Therefore, every fulfillment judgment starts with:

```text
Does this SO belong to the current-month SC baseline?
```

### 2.3 Loading Plan

Loading Plan is the shipping-arrangement ledger. It answers:

```text
Does this order have a shipping arrangement?
Is the loading date confirmed?
Is the date before the data_date, or on / after the data_date?
```

Loading Plan does not mean shipped, in stock, or in production.
It must be cross-checked against Shipped / FG / PP.

## 3. Recommended Reading Sequence

Read the page from top to bottom in four layers. Do not drill into detail first.

```text
Layer 1: Check whether the total baseline loop is closed
Layer 2: Check plant-level differences
Layer 3: Check shipped data closure
Layer 4: Check open fulfillment risk matrix and drill down to actions
```

This sequence follows the natural management questions:

```text
How big is the total responsibility pool?
Where are we now?
Which plant has more pressure, KS or IDN?
Can shipped volume be explained by Loading Plan evidence?
For open volume, is the risk in logistics, planning, production, or order ownership?
```

## 4. How To Read Top KPI

The top KPI block is the first closure layer. It shows where the current-month baseline is currently sitting.

| KPI | Business Meaning | How To Read |
|---|---|---|
| Baseline Orders | Current-month responsibility pool | Denominator. All statuses should reconcile back to this number |
| Shipped | Fulfilled volume | Already shipped; this is an outcome |
| In Stock (FG) | Finished goods in stock | Goods are ready; check whether shipping is arranged |
| In Production | Scheduled WIP | Production is running; check whether it can meet the loading date |
| Scheduling | Unclear supply volume | Includes WIP Unscheduled + No Plan; this is a risk pool requiring breakdown |

These statuses are MECE:

```text
Shipped + FG + WIP + Unscheduled + No Plan = Baseline
```

`Scheduling` is a management-level combined display. For action tracking, it should be split in the matrix into:

```text
WIP Unscheduled
No Supply Signal
```

These two categories require different follow-up actions.

## 5. How To Read BY PLANT

`BY PLANT` is a plant-level breakdown of the top KPI. It is not a new risk matrix.

It answers:

> What is the execution structure of KS and IDN baseline respectively?

| Field | Meaning |
|---|---|
| Plant | Plant |
| Baseline MT | Current-month baseline of this plant |
| Shipped | Shipped volume |
| FG | Finished goods in stock |
| WIP | Scheduled production in progress |
| Unscheduled | Work order exists, but no confirmed finish time |
| No Plan | No supply signal |
| Total | Total row under the current Plant filter |

Recommended reading:

```text
First, compare baseline size;
then compare Shipped + FG ratio;
then check which plant has heavier WIP / Unscheduled / No Plan.
```

This table helps identify which plant carries the main pressure. It does not answer whether Loading Plan exists.
Loading Plan coverage should be read in the risk matrix.

The `Total` row is the plant-level summary.
When `All Plants` is selected, it equals KS + IDN. When only `KS` or `IDN` is selected, it equals the selected plant total. This allows business users to reconcile back to the top KPI without manually summing plant rows.

## 6. How To Read Shipped Data Closure

`Shipped Data Closure` only looks at shipped volume.

It is not open risk, because the shipment has already happened.

It answers:

> Can the shipped volume be explained by Loading Plan evidence?

| Field | Meaning |
|---|---|
| Total MT | Baseline volume already shipped |
| LP Closed / Matched | Shipped volume that can be closed by current LP evidence |
| LP Missing or Unclear | Shipped volume with missing or unclear LP evidence |

Business interpretation:

```text
High LP Closed / Matched means shipped and LP evidence reconcile well.
High LP Missing or Unclear does not mean goods were not shipped.
It means the data chain needs review.
```

The usual owner is Sales Ops / Logistics / data process owner, not production.

## 7. How To Read Open Fulfillment Risk Matrix

This is the core area of the dashboard.

It only looks at unshipped baseline volume. It answers:

> For open volume, what is the supply status, what is the shipping-arrangement status, and what action is required?

Matrix rows are supply execution statuses:

| Row | Business Meaning |
|---|---|
| FG | Goods are in stock and theoretically ready to ship |
| WIP Scheduled | Production is running and has a planned end time |
| WIP Unscheduled | Work order exists but has no confirmed end time |
| No Supply Signal | No shipped, stock, scheduled PP, or unscheduled WO signal |

Matrix columns are Loading Plan coverage statuses:

| Column | Business Meaning |
|---|---|
| Past Due LP | Confirmed LP date exists, but it is earlier than data_date and not closed by shipped evidence |
| Future Valid LP | Confirmed LP date exists and is on / after data_date |
| LP Date Unconfirmed | LP record exists, but the date is TBA / blank / non-standard text |
| No Current LP | No usable current-scope LP coverage |

Each cell is MT, not SO count.
Each row should reconcile as:

```text
Supply Status Total
= Past Due LP
+ Future Valid LP
+ LP Date Unconfirmed
+ No Current LP
```

The source table behind this matrix is Excel `Risk Matrix Detail`.
The dashboard matrix is calculated by summing `risk_mt` by `supply_status` and `lp_coverage_status`. When a cell is clicked, the detail list below is the filtered detail from the same fact table.

### 7.1 Priority Logic

Start with the largest MT cells, then judge business severity.

Suggested priority:

```text
1. Past Due LP
2. No Supply Signal + Future Valid / Past Due LP
3. FG + No Current LP
4. WIP Scheduled + Future Valid LP with Gap < 0
5. LP Date Unconfirmed
6. WIP Unscheduled / No Current LP
```

Why:

- `Past Due LP` is already overdue. Check whether it was actually shipped, whether data was not updated, or whether the LP needs rescheduling.
- `No Supply Signal + LP` means there is a shipping arrangement but no supply-side support signal. This is a strong risk.
- `FG + No Current LP` means goods are ready but shipping is not arranged. This is usually a logistics / sales arrangement issue.
- `WIP Scheduled + Future Valid LP` must be checked against production feasibility.
- `LP Date Unconfirmed` does not mean no arrangement. It means the shipping date is unclear and needs confirmation.

## 8. How To Read Selected Risk Detail

When you click a matrix cell, `Selected Risk Detail` shows the corresponding list.

The detail table does not always show fixed columns. It displays different columns based on the selected scenario.

Principle:

> Only show fields that are truly needed for the current risk judgment, so background fields are not misread as facts.

Common fields:

| Field | Purpose |
|---|---|
| SO | Identify the order |
| Plant | Plant |
| Cluster | Region / customer cluster |
| Order Type | Order type |
| Supply | Current supply status |
| LP Coverage | Current LP coverage status |
| Risk MT | Volume requiring attention in the selected cell |
| Risk Action | Suggested action category |
| Owner | Suggested owner |
| Action Note | Business explanation |

### 8.1 WIP Scheduled Scenario

When supply status is `WIP Scheduled`, the detail table additionally shows:

| Field | Meaning |
|---|---|
| Machine | Production machine |
| Planned End | `Planned EndTime` from PP source, using the latest time by SO |
| Available Date | `Planned End + 1 day` |
| Loading Date | Required loading date from LP |
| Gap | `Loading Date - Available Date` |
| Work Order | Work order |

Recommended reading:

```text
First, check Risk MT;
then use Machine / Work Order to locate production;
then check Planned End and Available Date;
finally compare Loading Date and Gap.
```

`Planned End` is the planned Lami finish time.
`Available Date` is the reference date after allowing one additional day for stock-in / shipping readiness.

Most current LP loading dates are date-level only, so the gap is compared at date level:

```text
Loading Date = 2026-05-20
Available Date = 2026-05-20 23:20
Gap = 0
```

It will not be wrongly marked as -1 just because `23:20` is later than `00:00`.
Only when future LP provides clear hour / minute information should the comparison move to time level.

Gap colors:

| Gap | Meaning |
|---:|---|
| `< 0` | Production cannot meet the loading date |
| `0-2` | Tight but potentially feasible |
| `> 2` | Buffer exists |

### 8.2 FG Scenario

FG means goods are already in stock, so machine, work order, and planned end are usually not needed.

Focus on:

```text
Is there an LP?
Is the LP date confirmed?
If there is no LP, why are goods ready but not arranged for shipment?
```

Typical actions:

```text
FG + No Current LP -> Logistics / Sales to confirm shipping arrangement
FG + LP Date Unconfirmed -> Confirm loading date
FG + Past Due LP -> Check whether shipment happened but was not updated, or LP needs rescheduling
```

### 8.3 WIP Unscheduled Scenario

WIP Unscheduled means a work order exists, but no confirmed planned end time is available.

The focus is not gap, because there is no planned end date to compare.
Focus on:

```text
Why is there no planned finish time?
If LP already exists, can production provide a schedule?
If LP does not exist, both production and shipping plans need clarification.
```

### 8.4 No Supply Signal Scenario

No Supply Signal is the lowest supply-signal level.

It means the current process cannot find:

```text
Shipped
FG
Scheduled PP
Unscheduled WO
```

If it also has `Future Valid LP` or `Past Due LP`, sales / logistics has a shipping arrangement, but supply has no support signal. Planning should check it first.

If it is `No Current LP`, it means:

```text
The order is in baseline,
but neither supply nor shipping has a signal.
```

These items usually need SC / PP / business confirmation on whether the order truly belongs to the current-month responsibility pool.

## 9. How To Read LP Not In Current SC Baseline

This block is not part of the baseline fulfillment risk matrix.
It is an LP-led reconciliation exception.

It answers:

> Loading Plan contains shipping demand, but these SOs are not in the current-month SC baseline.

Common reasons:

| Reason | Explanation |
|---|---|
| Future-month order | LP includes orders not included in current-month billing responsibility |
| Missing baseline order | SC baseline may be missing these SOs |
| Cross-month arrangement | LP may cover cross-month shipping, but ownership is not confirmed |
| SO parsing issue | Invoice / SC No. parsing causes baseline matching failure |

Read this block by MT first, not SO count.
If details are needed, click the cell or download CSV.

## 10. How To Use Plant Filter And Download

### 10.1 Plant Filter

Plant filter affects:

- Top KPI
- BY PLANT
- Shipped Data Closure
- Open Fulfillment Risk Matrix
- LP Not In Current SC Baseline
- Selected Risk Detail

If `KS` is selected, the page shows KS only.
If `IDN` is selected, the page shows IDN only.
If `All Plants` is selected, the page shows the total view.

### 10.2 Download CSV

`Download CSV` exports the current detail view.

Rules:

| Current State | Download Content |
|---|---|
| No cell clicked | Default action-required detail |
| Risk matrix cell clicked | Current Plant + selected risk cell detail |
| Shipped closure clicked | Current Plant + shipped closure detail |
| LP Not In Baseline clicked | Current Plant + LP-led exception detail |

CSV columns follow the current dashboard view. They are not the same as the full Excel fact table.
For complete pivoting and audit, use Excel `Risk Matrix Detail`.

## 11. Common Misreadings To Avoid

### 11.1 Loading Plan Is Not An Execution Status

Do not treat Loading Plan as a peer status to Shipped / FG / WIP.
It is the shipping-arrangement dimension.

Correct interpretation:

```text
Supply status answers: where are the goods?
LP status answers: is shipping arranged?
```

### 11.2 No Current LP Does Not Always Mean No Historical LP

`No Current LP` means no usable LP coverage under the current main scope.
It does not necessarily mean the SO never appeared in any raw LP table.

Possible reasons include:

- Only historical valid LP exists
- Prior invoiced exclusion
- LP quantity is 0
- Raw number cannot be parsed

### 11.3 Missing LP For Shipped Is Not Fulfillment Risk

Shipped is already an outcome.
If shipped volume lacks LP evidence, it is a data-closure issue, not a risk that goods cannot ship.

### 11.4 FG Does Not Need Planned End

FG is already in stock.
For FG, planned end / machine / work order is not the key risk context.

The core FG question is:

```text
Goods are ready. Why is there still no clear loading arrangement?
```

### 11.5 Only WIP Scheduled Needs Gap

Gap is meaningful only for `WIP Scheduled + confirmed LP loading date`.
Only in this case do we know both:

```text
When production is expected to finish
When loading is required
```

## 12. Recommended Management Action Loop

Each dashboard review meeting can follow this rhythm:

```text
1. Confirm baseline total and data_date
2. Read top KPI to understand shipped / stock / production / scheduling structure
3. Read BY PLANT to identify whether KS or IDN carries more pressure
4. Read Shipped Data Closure to confirm whether shipped data can be explained
5. Read Open Fulfillment Risk Matrix and identify largest MT risk cells
6. Click cells and drill down to SO detail
7. Assign follow-up by Risk Action / Owner
8. Download CSV for business owners to review line by line
9. Refresh next cycle and check whether risk MT decreases
```

The final goal is not to explain every number. The goal is to reduce risk MT.

## 13. One-Sentence Summary

The dashboard should be read as:

```text
Use SC baseline to define the current-month responsibility pool,
use the supply waterfall to locate where the goods are,
use Loading Plan to check whether shipping arrangement is closed,
and use the risk matrix to assign unfulfilled volume into action queues.
```

Management reads the summary, business teams click cells for detail, and execution teams download CSV for line-by-line follow-up.
