# Loading Plan 引入后的履约蓝图

## 1. 本轮目标

本模型的核心问题不是“Loading Plan 本身完成了多少”，而是：

> 5月 SC Baseline 里的订单能不能完成发运？如果不能，风险卡在哪个环节？

因此，Summary 和明细表必须以 `SC Baseline` 作为主线和分母。`Loading Plan` 是发运安排底账，用来交叉验证 baseline 订单是否已有发运安排、日期是否明确、安排是否被货源支持。

本轮运行锚点补充：

- `data_date` 不手工取电脑当天，而是从 `02-Shipped` 文件夹最新 GI 文件后缀推断。
- 例如 `GI (KS&IDN) 260501-0513.xlsx` 代表本轮 shipped 数据截至 `2026-05-13`，因此 `Past Due LP / Future Valid LP` 也以 `2026-05-13` 为 cutoff。
- 输出文件使用分钟级后缀，例如 `20260514-1531`，同一天重复跑数不会覆盖前一版。

## 2. 三类数据的地位

| 数据层 | 业务地位 | 在模型中的作用 |
|---|---|---|
| SC Baseline | 本月开票责任池 | 定义本月必须兑现的订单范围 |
| Shipped / FG / PP | 货源执行状态 | 判断货在哪里、是否已经具备发运条件 |
| Loading Plan | 发运安排底账 | 判断是否有发运安排、日期是否确认、安排是否和货源匹配 |

关键结论：

- `Loading Plan` 不是 `Shipped / FG / PP` 的同级执行状态。
- Summary 必须先讲 baseline 订单执行到哪里，再用风险矩阵把 Loading Plan 放进发运风险判断里。
- 最终输出应落到行动队列：哪里需要物流、计划、工厂或业务跟进。

## 3. Loading Plan 主口径过滤

`Loading Plan Clean Detail` 保留所有原始清洗记录，用于追溯和审计。

### 3.1 IDN Schedule Planning Dispatch 新增口径

印尼出口的 `Schedule Planning Dispatch 260511.xlsx` 不能只读取 `ORDER OUTSTANDING `。同一文件里的 `DISPATCH` sheet 也要 append 进来，原因是：

> `ORDER OUTSTANDING ` 代表未完成 / outstanding 的发运安排，`DISPATCH` 代表已经 dispatch / 已发运的记录。两者合在一起才更接近印尼出口全量 Loading Plan。

本轮对齐口径：

| Sheet | 业务含义 | 主口径处理 |
|---|---|---|
| `ORDER OUTSTANDING ` | 未完成的印尼出口发运安排 | 正常进入 LP 清洗，TBA / 空 / 非标准日期继续作为日期未确认风险保留 |
| `DISPATCH` | 已经 dispatch / 已发运的印尼出口记录 | append 到 IDN Export LP；有效 `ELD` 只取本月第一天及之后 |

补充说明：

- `ORDER OUTSTANDING ` 的 sheet 名末尾有空格，必须精确匹配。
- `DISPATCH` 与 `ORDER OUTSTANDING ` 表头基本一致，最后几列缺失不影响主字段抽取。
- `DISPATCH` 中有历史数据，因此例如 2026-05 运行时，只把有效 `ELD >= 2026-05-01` 的记录放入本月主分析。
- `DISPATCH` 中有效 `ELD < 本月第一天` 的记录进入历史审计，不进入本月主分析。
- `DISPATCH` 中空 / 非标准 `ELD` 先进入审计，是否进入主口径需后续业务确认。

进入主分析口径的规则：

| 情况 | 是否进入主分析 | 原因 |
|---|---:|---|
| 有效 loading date 且早于本月第一天 | 否 | 属于历史发运安排，不应影响本月 baseline 判断 |
| 明确标记为 prior invoiced / 已开票 | 否 | 不属于本月开票责任口径 |
| TBA | 是 | 代表发运日期未确认，是需要跟进的风险 |
| 空 loading date | 是 | 代表发运安排不完整，是需要跟进的风险 |
| 非标准日期文本 | 是 | 代表发运日期不清晰，是需要跟进的风险 |
| 有效 loading date 且在本月或未来 | 是 | 用于判断是否有发运安排，以及生产是否赶得上 |

对于 `2026-05` 这类运行月份，本月第一天为 `2026-05-01`。有效日期早于该日期的 LP 记录进入审计，不进入 `SC vs LP Reconciliation` 和 `Shipping Readiness` 主口径。

## 4. SC vs LP Reconciliation 口径

这张表是第一层闭环，用来回答：

> 本月 baseline 订单是否进入当前 Loading Plan 主口径？

分类建议：

| 分类 | 业务解释 |
|---|---|
| In SC and In LP | 本月 baseline 订单有当前主口径 LP 记录 |
| In SC only | 本月 baseline 订单没有当前主口径 LP 记录 |
| In LP only | LP 里有发运安排，但不属于本月 SC baseline |

需要特别注意：

- `In SC only / LP MT = 0` 不是“有 LP 只是没有日期”。
- 如果 LP 里有 TBA / 空 / 非标准日期，但有吨数且 SO 可解析，应仍然算 `In SC and In LP`，只是归为日期未确认。
- `In SC only` 更准确的含义是：经过当前主口径过滤、SO 解析、排除项处理后，没有匹配到可用 LP 数量。
- SC 和 LP 两边吨数都为 0 的空行不进入管理输出。

建议新增字段：

| 字段 | 用途 |
|---|---|
| lp_loading_date_raw_list | 展示原始 Loading / ELD 值 |
| lp_earliest_valid_loading_date | 展示最早有效 loading date |
| lp_loading_date_status_mix | 展示日期状态组合，如 Valid Date / TBA / Blank |
| lp_match_scope | 标记是当前 LP、历史 LP、排除 LP，还是完全无 LP |
| lp_line_count | 对应 LP 原始行数 |
| lp_parse_exception_flag | 提醒是否可能因为单号解析失败导致漏匹配 |

## 5. Summary 叙事框架

Summary 的页面使命：

> 当前 baseline 订单是否可以完成？如果不能，优先处理哪些风险？

### 第一层：Baseline Execution Status

先用 baseline 做分母，回答订单执行到哪里。

| 状态 | 业务含义 |
|---|---|
| Shipped | 已兑现，已经发货 |
| FG | 货已经在库，理论上具备发运条件 |
| WIP Scheduled | 货还在生产，但已有计划完工日期 |
| WIP Unscheduled | 有工单，但没有明确完工日期 |
| No Supply Signal | 没有发货、库存、排产、工单信号 |

注意：旧 Summary 里的 `SCHEDULING` 不是一个足够准确的业务状态。它混合了 `WIP Unscheduled` 和 `No Supply Signal`，后续应拆开。

### 第二层：Shipped Data Closure

这一层单独看已经发货的闭环。

`Shipped` 不是 open risk，它已经是兑现结果。继续把它和 FG / WIP 放在同一个风险矩阵里，会把“数据一致性检查”和“未发货行动风险”混在一起。

| 货源状态 | Total MT | LP Closed / Matched | LP Missing or Unclear |
|---|---:|---:|---:|
| Shipped | 已发货 baseline 量 | 能被本月 LP 证据解释的量 | 找不到可用 LP 证据或 LP 日期不清晰的量 |

这块回答：

> 已经发货的订单，在 Loading Plan 侧是否也能闭环？

IDN 新增 `DISPATCH` sheet 的价值就在这里：它补齐了印尼出口已经发货的 LP 记录，减少“Shipped 但 No Current LP”的假异常。

### 第三层：Open Fulfillment Risk Matrix

这一层用于管理层判断“未发货的量卡在哪里、谁要行动”。

| 货源状态 | Supply Status Total | Past Due LP | Future Valid LP | LP Date Unconfirmed | No Current LP |
|---|---:|---:|---:|---:|---:|
| FG | 数字 MT | 已过期 LP 但未被 shipped 闭环 | 有未来明确发运日期 | 货已好但发运日期未确认 | **在库但无 Loading Plan** |
| WIP Scheduled | 数字 MT | 已过期 LP 但生产/发运未闭环 | 看完工日是否 meet loading date | 生产中但发运日期未确认 | 生产中但未安排发运 |
| WIP Unscheduled | 数字 MT | 已过期 LP 但生产无完工日 | 有发运要求但生产无日期 | 双重不确定 | 生产和发运都未明确 |
| No Supply Signal | 数字 MT | **已过期发运安排但无货源信号** | **有发运安排但无货源信号** | LP 不确认且无货源 | Baseline 订单无供应/发运信号 |

注意：

- Summary 页面不再单独拆一个 `Loading Arrangement Coverage` 模块。
- Loading Plan 不是独立展示层，而是 open risk matrix 里的发运安排维度。
- `Past Due LP` 放在 `Future Valid LP` 前面，因为前者已经晚于承诺/计划日期。
- `Past Due LP` 不是只看日期小于今天，而是要扣掉已经被 shipped 闭环的量：

```text
Past Due LP = loading_date < data_date 的有效 LP 量 - 已被 shipped 闭环的量
```

- Summary 矩阵的单元格必须是纯数字 MT，不在格子里写 `MT` 或 SO 数。
- Summary 矩阵需要在 `Supply Status` 右侧增加 `Supply Status Total` 列，先显示每个供应状态的总量，再向右展示 LP 覆盖状态拆解。

## 6. 管理层行动队列

Summary 最终应把风险归到几类行动队列：

| 行动队列 | 业务含义 |
|---|---|
| FG without Loading Plan | 货已在库，但没有当前发运安排 |
| Loading Plan without Supply Signal | 有发运安排，但没有货源信号支撑 |
| WIP Late vs Loading Date | 已排产，但完工日赶不上 loading date |
| LP Date Unconfirmed | 有 LP 记录，但 loading date 不明确 |
| Production Unscheduled | 有工单，但生产日期未排清楚 |
| No Supply and No Loading Signal | 本月 baseline 订单既没有货源信号，也没有发运安排 |

排序锚点建议使用 MT 量，而不是 SO 数。SO 数可以作为辅助信息出现在明细或 drill-down 表里。

## 7. Risk Matrix Detail 与 Action Required

这次对齐后，`Risk Matrix Detail` 是唯一事实底表。

它同时支撑：

- Summary 上的 Shipped Data Closure 和 Open Fulfillment Risk Matrix
- `Action Required` 行动清单

逻辑关系：

```text
Risk Matrix Detail
  ├── 透视汇总 -> Summary Shipped Data Closure / Open Fulfillment Risk Matrix
  └── 筛选 action_required = True -> Action Required
```

`Risk Matrix Detail` 的粒度不是“一个 SO 一行”，而是：

```text
一个 SO + 一个 allocated 供应分段 + 一个 LP 覆盖分段
```

这样可以避免把一个 SO 的全部 baseline 都误解为同一种风险。例如一个 SO 总量 100 MT，其中 90 MT 已经被 Shipped / FG / WIP 覆盖，只有 10 MT 是 No Supply Signal，那么 `Risk Matrix Detail` 只会把 10 MT 作为风险量。后续进一步要求是：LP 也尽量按数量覆盖分配，避免把一个 SO 级 LP 状态重复贴到多个供应分段上。

关键字段建议：

| 字段 | 含义 |
|---|---|
| so | SO 编号 |
| supply_status | Shipped / FG / WIP Scheduled / WIP Unscheduled / No Supply Signal |
| lp_coverage_status | Past Due LP / Future Valid LP / LP Date Unconfirmed / No Current LP / Shipped LP Closed |
| risk_mt | 当前风险分段吨数 |
| so_total_mt | SO 总 baseline，仅作为背景 |
| covered_mt | 已被更高优先级状态覆盖的量 |
| shipped_closed_mt | 已被 shipped 闭环的 LP 量，用于避免 Past Due LP 被误判 |
| risk_action | 风险行动标签 |
| action_required | 是否进入 Action Required |
| suggested_owner | 建议责任方 |
| action_note | 业务解释 |

`Action Required` 不再使用 SO 级 `Status` 直接生成，而是 `Risk Matrix Detail` 的一个筛选视图。

## 8. 下一步执行范围

后续代码落地建议按以下顺序：

1. IDN Export LP 增加 `DISPATCH` sheet append，并按 `ELD >= 本月第一天` 过滤有效历史数据。
2. 更新 Loading Plan 主口径过滤：历史有效日期不进主分析，异常日期继续保留。
3. 更新 `SC vs LP Reconciliation`：增加原始 loading date、日期状态、LP 匹配范围等字段。
4. 移除 SC 和 LP 两边吨数都为 0 的空行。
5. 重构 Summary 叙事：Baseline Execution Status + Shipped Data Closure + Open Fulfillment Risk Matrix。
6. 重构 `Risk Matrix Detail`：从 SO 级 LP 状态升级为数量级 LP 覆盖状态，支持 Past Due LP / Future Valid LP。
7. 让 Summary Matrix 与 Action Required 都从 `Risk Matrix Detail` 派生。
8. 重新跑一版 Excel，重点检查“在库但无 Loading Plan”和“Past Due LP 未被 shipped 闭环”是否可以直接回答。
