# S&OE 订单履约跟踪系统 - 中文蓝图

## 1. 目标

本系统围绕“本月 SC Baseline 订单能否完成发运”来建立一套月度循环机制。

对每一个进入本月 baseline 的 SO，模型要回答四个管理问题：

- 本月开票责任池里有哪些订单，合计多少 MT？
- 每个 baseline SO 现在执行到哪里：已发货、在库、已排产、工单未排期，还是没有供应信号？
- 这些订单是否有当前 Loading Plan 发运安排，loading date 是否明确？
- 如果不能顺利完成，风险卡在哪个环节，应该进入哪个行动队列？

这是一个月度复用机制：每月放入最新源文件，输出同一套结构化风险视图。当前管理目标约为 `25,000 MT/月`，覆盖昆山（KS）与印尼（IDN）。

数据性质说明：

- `Shipped` 是月初至数据日的累计发货量。
- `FG` 是数据抽取日当天的成品库存快照。
- `PP` 是生产计划与工单状态。
- `Loading Plan` 是发运安排底账，不是供应执行状态。

运行锚点规则：

- `data_date` 运行时从 `02-Shipped` 文件夹最新 GI 文件后缀自动推断，例如 `GI (KS&IDN) 260501-0513.xlsx` 推断为 `2026-05-13`。
- 报表 cutoff 不使用电脑当天日期，因为源数据可能晚于或早于打开报表的日期。
- 输出文件写入 `output/<data_date>/`，文件名追加分钟级运行后缀，例如 `SOE_Tracking_2026-05-20260514-1531.xlsx`，避免同一天多次运行互相覆盖。

## 2. 总体数据流

`SC Baseline` 定义本月开票责任池，也就是“这个月我们欠客户哪些 SO”。

之后模型按两条线看同一批 baseline SO：

- 供应执行线：`Shipped / FG / PP`，回答“货在哪里”。
- 发运安排线：`Loading Plan`，回答“有没有发运安排、日期是否明确”。

最终不是把 Loading Plan 单独做成一层 Summary，而是把它放进风险矩阵里，和供应执行状态交叉判断。

核心结构：

```text
SC Baseline = 本月开票责任池
Shipped / FG / PP = 供应执行状态
Loading Plan = 发运安排底账
Risk Matrix = 供应状态 × 发运安排
```

## 3. 源表抽取规则

### 3.1 SC - 销售订单 Baseline（01-SC）

来源：

- 文件：每月 `Order tracking *.xlsx`
- Sheet：`Order Status - Main`
- 基础过滤：`Item = main`

关键字段：

| 字段 | 用途 |
|---|---|
| SC NO | SO 编号，主关联键 |
| Supply From | 工厂来源，映射为 KS / IDN |
| Cluster | 区域 / 客户群维度 |
| End User Cust ID | 客户映射校验 |
| SC Vol.-MT | 原始订单吨数 |
| Order Status | 订单类型判断 |
| Carryover | 为空代表本月开票口径 |
| carryover/fresh production | 判断 Carry Over Unproduced |
| RELEASE DATE | 判断 Fresh Order 所属月份 |
| Delivery PCS | SC 侧交付信号，负数行需要排除 |
| Loading Date | 只作 SC 侧参考，不作为可靠 LP loading date |

客户映射：

- 文件：`07-Mapping/dim_cc_region.xlsx`
- Sheet：`dim->new cc`
- Join：`SC.End User Cust ID -> dim->new cc.new customer code`
- 映射失败的 end customer 不进入主 baseline，进入 `Unmatched End Customer` 审计表。

订单分类：

| 类型 | 判断逻辑 | 是否进入本月 baseline |
|---|---|---|
| Carry Over Stock | `Order Status = Carryover` 且 `Carryover` 为空 | 是 |
| Carry Over Unproduced | `carryover/fresh production = last month` 且 `Carryover` 为空 | 是 |
| Fresh Order This Month | `RELEASE DATE` 在当前月且 `Carryover` 为空 | 是 |
| Fresh Order Next Month | `RELEASE DATE` 在当前月但 `Carryover` 有值 | 否，仅保留审计 |

调整系数：

| 类型 | 调整后 baseline |
|---|---|
| Carry Over Stock | `SC Vol.-MT` |
| Carry Over Unproduced | `SC Vol.-MT × 0.975` |
| Fresh Order This Month | `SC Vol.-MT × 0.975` |
| Fresh Order Next Month | `SC Vol.-MT × 0.975`，但不进本月 baseline |

Baseline 口径：

```text
Baseline = Carry Over Stock + Carry Over Unproduced + Fresh Order This Month
```

聚合规则：

- 先在 SC 行级判断类型和调整吨数，再按 SO 聚合。
- 保留各类型调整后吨数字段，避免一个 SO 多行、多类型时被错误归到第一行类型。

SC 上月交付预分配：

- 使用同一张 SC 表构造 SC 派生交付记录。
- 时间窗口：相对 `data_date` 的上一个自然月。
- 过滤：
  - `Item = main`
  - `Loading Date` 落在上月窗口
  - `Delivery PCS` 不是负数
  - End customer 映射通过
- 数量：使用 `SC Vol.-MT`
- 用途：作为已兑现信号，优先分配到 `Allocated Shipped`，但不替代 02-Shipped GI 数据。

### 3.2 Shipped - 实际发货（02-Shipped）

来源：`GI (KS&IDN) *.xlsx`

关键字段：

| 字段 | 用途 |
|---|---|
| Post Date | 实际发货日期 |
| Plant | 工厂代码，映射 KS / IDN |
| Sales Order | SO 编号 |
| Weight (KG) | 发货重量，转 MT |

处理规则：

```text
shipped_mt = Weight(KG) × -1 / 1000
```

说明：系统里的发货重量通常以负数显示，所以先乘以 `-1` 再从 KG 转成 MT。

### 3.3 FG - 成品库存（03-FG）

来源：`FG stock*.xlsx`

关键字段：

| 字段 | 用途 |
|---|---|
| Plant | 工厂 |
| Weight (KG) | 库存重量，转 MT |
| Receipt Date | 入库日期 |
| Contract Code | SO 编号 |

处理规则：

```text
fg_mt = Weight(KG) / 1000
```

FG 表回答的是：订单对应的货是否已经在库，理论上是否具备发运条件。

### 3.4 PP - 生产计划（04-PP）

PP 分两类：

| 类型 | 业务含义 |
|---|---|
| Scheduled PP | 有工单且有计划完工日期 |
| Unscheduled PP | 有工单，但没有计划完工日期 |

Scheduled PP 关键字段：

- Work Order No.
- SO
- TotalWeight/T
- Planned EndTime（优先使用，带小时分钟/秒，代表 Lami 计划结束时间）
- PlannedFinishDate（仅作为日期级参考或缺失兜底）
- Machine
- Plant

Unscheduled PP 关键字段：

- Work Order No.
- SO
- TotalWeight/T
- Machine
- Plant

聚合规则：

- 同一个 SO 可能对应多个工单。
- `wip_mt` = 同一 SO 下所有工单吨数之和。
- `planned_end_date` = 同一 SO 下所有工单最晚的 `Planned EndTime`。
- 若 `Planned EndTime` 缺失，再用 `PlannedFinishDate` 作为兜底。
- 原因：一个 SO 只有等所有子批次完成后才具备完整发运条件，所以最晚结束时间是瓶颈。

### 3.5 Loading Plan - 发运安排底账（06-Loading Plan）

Loading Plan 用来判断是否有发运安排，不是 `Shipped / FG / PP` 的同级执行状态。

#### KS Loading Plan

来源：

- 文件夹：`06-Loading Plan/ks_loading plan/`
- 文件：`Loading plan-May.xlsx`
- Sheet：`Loading plan`

关键字段：

| 字段 | 用途 |
|---|---|
| Invoice No | 拆解 / 标准化为 SO |
| 20GP / 40GP / 40HQ | 柜型数量，转换为模型吨数 |
| Loading | 原始 loading date 或异常文本 |
| MT | 源表吨数，仅用于审计 |
| Unnamed: 20 | 如 `4月已开票`，排除本月主口径但保留审计 |

柜型吨数：

| 柜型 | MT |
|---|---:|
| 20GP | 14.5 |
| 40GP | 24.5 |
| 40HQ | 24.5 |

KS SO 拆解规则：

- `_` 表示多个 SO。
- `-N` 是批次 / 柜序号后缀，去掉后缀。
- `-N~M` 是多柜范围，识别为同一 SO。
- 短号继承同组中最长 10 位 SO 的前缀。
- `LM*` 等非标准编号进入 parse exception 审计。
- 一行拆成多个 SO 时，吨数按 SO 数量均分。

#### IDN Export Loading Plan

来源：

- 文件夹：`06-Loading Plan/idn_loading plan/`
- 文件：`Schedule Planning Dispatch 260511.xlsx`
- Sheet：
  - `ORDER OUTSTANDING `：印尼出口未完成 / outstanding 的发运安排，注意末尾有空格。
  - `DISPATCH`：印尼出口已经 dispatch / 已发运的发运记录，需要 append 到同一个 IDN Export Loading Plan 口径里。

业务口径：

- `ORDER OUTSTANDING ` 和 `DISPATCH` 的表头结构基本一致，`DISPATCH` 最后几列可能缺失，不影响主字段抽取。
- 抽取时不能按固定列数硬套，应该按字段名对齐。
- `DISPATCH` 用来补齐印尼已经发运的 Loading Plan 记录，否则 IDN 的 LP 只看 outstanding 会不完整。
- `DISPATCH` 中 `ELD` 有历史数据；本月主口径只保留有效 `ELD >= 本月第一天` 的记录。例如 2026-05 的运行口径为 `ELD >= 2026-05-01`。
- `DISPATCH` 中有效 `ELD` 早于本月第一天的记录，作为历史已发运 LP 证据保留在审计，不进入本月主分析。
- `DISPATCH` 中空 / 非标准 `ELD` 先进入审计；是否进入主口径需要后续和业务确认，因为该 sheet 按定义应代表已发运记录。

关键字段：

| 字段 | 用途 |
|---|---|
| SC No. | 直接作为 SO key |
| Region | 区域参考 |
| Cont Qty / Cont Size | 按柜型转换吨数 |
| Rough Ton | 源表吨数，仅用于审计 |
| ELD | 原始 loading date 或 TBA 等异常值 |

#### IDN Domestic Loading Plan

来源：

- 文件夹：`06-Loading Plan/idn_loading plan/`
- 文件：`NEW DOMESTIC TRACKING 260511.xlsx`
- Sheet：`Order List`，表头在第 2 行。

关键字段：

| 字段 | 用途 |
|---|---|
| SC NO. | 优先使用的 SO key |
| INV NO. | 当 SC NO. 不可用时作为拆解来源 |
| ELD | 原始 loading date 或异常值 |
| Weight | 发运吨数 |
| STATUS | 业务参考字段 |

IDN Domestic 规则：

- 优先使用有效 `SC NO.`。
- 如果 `SC NO.` 缺失或无效，再拆 `INV NO.`。
- 逗号和 `&` 表示多个 SO。
- `-N` 和 `-N~M` 后缀需要去掉。
- `LMIDSAM*` 和无法解析为 SO 的记录进入异常审计。

Loading Plan 主口径过滤：

| 情况 | 是否进入主分析 | 说明 |
|---|---:|---|
| 有效 loading date 且早于本月第一天 | 否 | 视为历史 LP 证据，保留审计 |
| 明确 prior invoiced / 已开票 | 否 | 不属于本月开票责任池 |
| TBA | 是 | 代表发运日期未确认，是风险 |
| 空 loading date | 是 | 代表安排不完整，是风险 |
| 非标准日期文本 | 是 | 代表发运日期不清晰，是风险 |
| 有效 loading date 且在本月或未来 | 是 | 用于发运安排和 gap 判断 |

## 4. Join 与状态判断

所有主关联都收敛到 `SO`。

```text
SC Baseline
  LEFT JOIN Shipped
  LEFT JOIN FG
  LEFT JOIN PP Scheduled
  LEFT JOIN PP Unscheduled
  LEFT JOIN Loading Plan
```

管理 KPI 使用“分配后瀑布”，而不是直接把所有源表 raw 数量相加。

分配逻辑：

```text
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

闭环公式：

```text
allocated_sc_prior_shipped
+ allocated_shipped
+ allocated_fg
+ allocated_wip
+ allocated_unsched
+ allocated_no_plan
= adjusted SC baseline
```

这保证了 Summary 的主执行状态是 MECE 的：每一吨 baseline 量只落在一个供应执行状态里。

供应执行状态建议命名：

| 状态 | 业务含义 |
|---|---|
| Shipped | 已兑现，已经发货 |
| FG | 货已在库，理论上具备发运条件 |
| WIP Scheduled | 生产中，且有计划完工日期 |
| WIP Unscheduled | 有工单，但没有计划完工日期 |
| No Supply Signal | 没有发货、库存、排产、工单信号 |

## 5. Gap 计算

Gap 只在同时满足以下条件时计算：

- SO 有 PP Scheduled 的 `planned_end_date`
- SO 有当前主口径内的有效 `loading_date`

公式：

```text
Available Date = Planned EndTime + 1 day
LP Gap Days = Loading Date - Available Date
```

解释：

- `+1 day` 表示生产完成后，通常次日才可视为具备入库 / 可装柜条件。
- `Planned End Date` 应来自 PP 底表 `Planned EndTime`，保留小时分钟/秒；`Available Date` 也应继承这个时间精度。
- 当前 Loading Plan 的 `Loading Date` 多数是日期级。如果 LP 没有具体小时分钟，gap 判断先按日期比较，避免把 `2026-05-17 04:35` 和 `2026-05-17` 误判为晚一天。
- 只有当未来 LP 也提供明确装柜小时分钟时，才进入小时级 gap 判断。
- Gap > 2：有缓冲。
- Gap 0-2：紧张但可能可行。
- Gap < 0：生产赶不上 loading date。

注意：Loading Plan 不只服务于 PP gap。即使 SO 是 FG、Shipped 或 No Supply Signal，也仍然要通过风险矩阵判断发运安排是否闭环。

## 6. Summary 页面设计

Summary 页面使命：

> 当前 baseline 订单是否可以完成？如果不能，优先处理哪些风险？

Summary 不再单独拆一个 `Loading Plan Coverage` 模块。Loading Plan 直接进入风险矩阵，作为发运安排维度。

### 6.1 Block 1 - Baseline Execution Status

第一块回答：本月 baseline 订单现在执行到哪里了？

| 指标 | 用途 |
|---|---|
| Baseline MT | 本月责任池总量 |
| Shipped MT | 已兑现 |
| FG MT | 货已准备，等待发运执行 |
| WIP Scheduled MT | 有排产日期 |
| WIP Unscheduled MT | 有工单但无完工日期 |
| No Supply Signal MT | 没有供应信号 |

Dashboard 顶部应保留 `BY PLANT` 行级汇总表，作为第一层执行状态的工厂拆解：

| Plant | Baseline MT | Shipped | FG | WIP | Unscheduled | No Plan |
|---|---:|---:|---:|---:|---:|---:|
| IDN | 数字 MT | 数字 MT | 数字 MT | 数字 MT | 数字 MT | 数字 MT |
| KS | 数字 MT | 数字 MT | 数字 MT | 数字 MT | 数字 MT | 数字 MT |

这里的行级汇总不是新的业务维度，而是把 `Baseline Execution Status` 按工厂拆开，方便管理层快速看到 KS / IDN 各自卡在哪个状态。单元格保持纯数字，单位放在标题或表头。

### 6.2 Block 2 - Shipped Data Closure

第二块先回答：已经发货的 baseline 量，数据闭环了吗？

`Shipped` 已经是兑现结果，不再和 FG / WIP 一起放进 open risk 里。它要单独作为数据闭环审计：

| 货源状态 | Total MT | LP Closed / Matched | LP Missing or Unclear |
|---|---:|---:|---:|
| Shipped | 已发货 baseline 量 | 能在本月 LP 证据中闭环的量 | 找不到可用 LP 证据或日期不清晰的量 |

这块的意义不是说“已发货还有执行风险”，而是检查：系统 shipped 和 Loading Plan 是否能互相解释。IDN 新增 `DISPATCH` sheet 后，印尼已发货部分的 LP 证据会更完整。

### 6.3 Block 3 - Open Fulfillment Risk Matrix

第三块回答：未发货的 open baseline 量，风险卡在哪里？

| 货源状态 | Supply Status Total | Past Due LP | Future Valid LP | LP Date Unconfirmed | No Current LP |
|---|---:|---:|---:|---:|---:|
| FG | 数字 MT | 已过期 LP 但未被 shipped 闭环 | 有未来明确发运日期 | 有 LP 但日期不明确 | **在库但无 Loading Plan** |
| WIP Scheduled | 数字 MT | 已过期 LP 但生产/发运未闭环 | 看完工日是否 meet loading date | 生产中但发运日期未确认 | 生产中但未安排发运 |
| WIP Unscheduled | 数字 MT | 已过期 LP 但生产无完工日 | 有发运要求但生产无日期 | 双重不确定 | 生产和发运都未明确 |
| No Supply Signal | 数字 MT | **已过期发运安排但无货源信号** | **有发运安排但无货源信号** | LP 不确认且无货源 | Baseline 订单无供应/发运信号 |

这里的列不是简单的 SO 标签，而是要尽量按数量分配：

- `Past Due LP`：有效 LP 日期早于 `data_date`，且没有被 shipped 数量闭环的 LP 量。
- `Future Valid LP`：有效 LP 日期在 `data_date` 当天或之后，且仍属于 open fulfillment 的 LP 量。
- `LP Date Unconfirmed`：有 LP 记录，但日期是 TBA / 空 / 非标准文本。
- `No Current LP`：当前主口径下没有可用 LP 覆盖。
- `Past Due LP` 放在 `Future Valid LP` 前面，因为它已经晚于计划，更需要优先管理层关注。

管理层行动队列：

| 行动队列 | 业务含义 |
|---|---|
| FG without Loading Plan | 货已在库，但没有当前发运安排 |
| Loading Plan without Supply Signal | 有发运安排，但没有货源信号支撑 |
| WIP Late vs Loading Date | 已排产，但完工日赶不上 loading date |
| LP Date Unconfirmed | 有 LP 记录，但 loading date 不明确 |
| Production Unscheduled | 有工单，但生产日期未排清楚 |
| No Supply and No Loading Signal | baseline 订单既无货源信号，也无发运安排 |

排序锚点建议使用 MT 量。SO 数可以作为辅助字段放在明细表中。

Summary 矩阵格式规则：

- 矩阵单元格必须是可计算的数字 MT，不要写成 `3,589 MT` 这种文本。
- 不要在矩阵单元格里放 SO 数。
- 单位放在标题或表头，例如 `Open Fulfillment Risk Matrix (MT)`。
- 在 `Supply Status` 右侧新增 `Supply Status Total` 列，先展示该供应状态总量，再向右拆解到 `Past Due LP / Future Valid LP / LP Date Unconfirmed / No Current LP`。
- 这样每一行可以直接对回对应的供应状态总量，所有行合计可以对回 baseline。

### 6.4 Risk Matrix Detail 与 Action Required 的关系

`Risk Matrix Detail` 是 Summary 风险矩阵和 `Action Required` 的唯一事实底表。

粒度：

```text
一个 SO + 一个 allocated 供应分段 + 一个 LP 覆盖分段
```

这意味着一个 SO 可能出现多行。例如同一个 SO 可能一部分已经 Shipped，一部分在 FG，还有一部分是 No Supply Signal；同一个供应分段也可能再按 LP 覆盖状态拆成 Past Due LP、Future Valid LP、LP Date Unconfirmed、No Current LP。这样做的原因是避免把整个 SO 都严重化成一个风险，而是只把真正暴露的风险吨数拿出来管理。

推荐字段：

| 字段 | 含义 |
|---|---|
| so | SO 编号 |
| plant / cluster / order_type | 管理维度 |
| so_total_mt | SO 总 baseline，作为背景 |
| supply_status | Shipped / FG / WIP Scheduled / WIP Unscheduled / No Supply Signal |
| lp_coverage_status | Past Due LP / Future Valid LP / LP Date Unconfirmed / No Current LP / Shipped LP Closed |
| risk_mt | 当前分段的 allocated MT，是行动管理主锚点 |
| covered_mt | 该 SO 已被更高优先级供应状态覆盖的量 |
| shipped_closed_mt | 已被 shipped 闭环的 LP 量，用于判断 Past Due LP 是否真实未闭环 |
| lp_loading_date_raw_list | 原始 Loading / ELD 汇总 |
| lp_earliest_valid_loading_date | 当前主口径最早有效 loading date |
| planned_end_date | PP 底表 Planned EndTime，按 SO 取最晚时间，代表 Lami 结束时间 |
| available_date | `planned_end_date + 1 day`，当前用于 gap 判断的可发运参考日期 |
| lp_gap_days | `Loading Date - Available Date`，可计算时的 LP gap；若 LP 只有日期级，则先按日期比较 |
| risk_action | 风险行动标签 |
| action_required | 是否进入 Action Required |
| suggested_owner | 建议责任方，例如物流 / 计划 / 工厂 / 业务 |
| action_note | 简短业务解释 |

Dashboard 展示规则：

- Excel / `Risk Matrix Detail` 是稳定底表，字段保持完整，便于筛选、透视和审计。
- Dashboard 是业务工作台，点击矩阵格子后，明细列可以按场景动态展示。
- `planned_end_date / available_date / machines / work_orders` 只在 `WIP Scheduled` 场景展示。
- `loading_date` 只在 `Past Due LP / Future Valid LP` 场景展示；`No Current LP` 不展示 loading date。
- `LP Date Unconfirmed` 场景不展示正式 `loading_date`，改展示 `lp_loading_date_raw_list / lp_loading_date_status_mix`。
- `lp_gap_days` 只在 `WIP Scheduled + 有明确 loading date` 时计算和展示，并在 dashboard 中高亮。
- `risk_mt / risk_action / suggested_owner / action_note` 是所有 dashboard 明细场景都要展示的行动字段。
- `WIP Scheduled` 场景的业务阅读顺序为：`Risk MT -> Machine -> Planned End -> Available Date -> Loading Date -> Gap -> Risk Action -> Owner -> Work Order -> Action Note`。
- PP 底表应优先使用 `Planned EndTime`，并在 Excel 和 dashboard 保留到分钟级；如需展示秒，可在底层保留、前端按分钟展示。
- `Available Date = Planned EndTime + 1 day`，但 gap 在当前 LP 日期级口径下按日期比较；未来 LP 若提供小时分钟，再升级为小时级比较。
- Dashboard 视觉风格采用管理层汇报风格：暖白/象牙色背景、深海军蓝标题、青蓝与金色作为主要对比色、浅暖灰表格线，避免过重的纯蓝表头。
- Dashboard 的 `Selected Risk Detail` 应提供 CSV 下载按钮，导出当前 Plant filter 和当前点击格子对应的明细；若未点击格子，则导出默认 action-required 明细。
- Dashboard HTML 应保持单文件静态快照，便于直接分发给业务方用 Edge / Chrome 打开；同时需要提示该文件内嵌明细数据，分发时要注意权限。

#### LP Not In Current SC Baseline

`Open Fulfillment Risk Matrix` 是 baseline-led 的履约风险矩阵，分母是 current-month SC baseline open quantity。

`LP Not In Current SC Baseline` 是 LP-led 的对账异常，分母是 current-scope Loading Plan 中无法对回 current-month SC baseline 的 LP quantity。它不应混入履约风险矩阵，而应作为第三块独立呈现。

Summary 结构：

| LP Status | MT |
|---|---:|
| Past Due LP | numeric MT |
| Future Valid LP | numeric MT |
| LP Date Unconfirmed | numeric MT |
| Total | numeric MT |

不展示 SO Count，避免管理层注意力从量转移到单数。需要 SO 数时，可从底表或 CSV 下载后透视。

明细粒度：

```text
一条 current-scope Loading Plan 明细行
```

推荐 dashboard / Excel 明细字段：

```text
SO / Plant / LP Source / Source Sheet / Invoice No Raw
Loading Date Raw / Loading Date / LP Date Status / Load MT
```

这块用于回答：LP 是否包含未来月订单、本月 baseline 漏单、跨月安排、或业务侧尚未归属清楚的 loading demand。

派生关系：

| 输出 | 来源 | 用途 |
|---|---|---|
| Summary Shipped Data Closure | 从 `Risk Matrix Detail` 的 shipped 闭环字段透视汇总 | 检查已发货量是否有 LP 证据闭环 |
| Summary Open Fulfillment Risk Matrix | 从 `Risk Matrix Detail` 的非 shipped 行透视汇总 | 展示未发货供应状态 × LP 覆盖状态的 MT |
| Dashboard LP Not In Current SC Baseline | 从 `lp_not_in_baseline_detail` 汇总 | 展示 LP-led 对账异常，不混入 baseline 履约矩阵 |
| Action Required | 筛选 `Risk Matrix Detail` 中 `action_required = True` 的行 | 只展示需要人跟进的风险清单 |

因此，`Action Required` 不再从 SO 级 `Status` 直接生成。`SO Master` 里的 `Status` 可以继续作为最严重风险提示，但管理动作应以 `Risk Matrix Detail.risk_mt` 为准。

## 7. 输出

Excel 输出建议包含：

- `Summary`
- `SO Master`
- `Gap Analysis`
- `Risk Matrix Detail`
- `Action Required`
- `Overlap Audit`
- `SC Row Detail`
- `SC Fresh Next Month`
- `SC Unknown Type`
- `Unmatched End Customer`
- `Loading Plan Clean Detail`
- `SC vs LP Reconciliation`
- `Shipping Readiness`
- `LP Date Exceptions`
- `LP Parse Exceptions`
- `LP Excluded Prior Invoiced`

Dashboard 是本轮主要管理层呈现入口，文件名形如 `SOE_Dashboard_<month>-<run_suffix>.html`。旧的叙事型 `SOE_Report_<month>-<run_suffix>.html` 不再作为必要输出，后续可停止生成，避免同一套信息出现两个前端版本。

Loading Plan 明细层标准字段：

| 字段 | 含义 |
|---|---|
| plant | KS / IDN |
| lp_source | KS_LP / IDN_Export / IDN_Domestic |
| source_file / source_sheet / source_row | 源文件追溯 |
| invoice_no_raw | 原始发票号 / SC key |
| so | 解析后的 SO |
| so_parse_status | Parsed / Non-SC / Parse Failed |
| loading_date_raw | 原始 Loading / ELD |
| loading_date | 可解析的日期 |
| loading_date_status | Valid Date / TBA / Blank / Text Month / Invalid Text |
| load_mt | 模型使用吨数 |
| source_mt | 源表吨数 |
| exclude_from_current_invoice | 是否排除本月主口径 |
| exclude_reason | 排除原因 |

`SC vs LP Reconciliation` 建议增加字段：

| 字段 | 用途 |
|---|---|
| lp_loading_date_raw_list | 原始 Loading / ELD 汇总 |
| lp_earliest_valid_loading_date | 最早有效 loading date |
| lp_loading_date_status_mix | 日期状态组合 |
| lp_match_scope | Current LP / Historical LP only / Excluded LP only / No LP evidence |
| lp_line_count | 支撑该 SO 的 LP 原始行数 |
| lp_parse_exception_flag | 是否可能存在 SO 解析异常导致漏匹配 |

## 8. 逻辑闭环

新版闭环是管理决策闭环，不是旧版“PP -> Loading Plan -> Gap”的单一路径。

```text
SC Baseline（本月责任池）
        |
        v
供应执行瀑布
  - Shipped
  - FG
  - WIP Scheduled
  - WIP Unscheduled
  - No Supply Signal
        |
        v
风险行动矩阵
  行 = 供应执行状态
  列 = Past Due LP / Future Valid LP / LP Date Unconfirmed / No Current LP
        |
        v
行动队列
  - 在库但无 Loading Plan
  - 有发运安排但无货源信号
  - 生产赶不上 Loading Date
  - Loading Date 未确认
  - 生产未排清楚
  - 无供应且无发运安排
```

说明：

- 供应执行瀑布必须 MECE，且加总回 adjusted SC baseline。
- Loading Plan 不作为 Summary 的单独执行状态。
- Loading Plan 在风险矩阵中作为发运安排维度。
- Shipped 已经是兑现结果，单独做数据闭环审计，不混入未发货 open risk。
- Past Due LP 必须按数量判断：已过期有效 LP 量减去已经被 shipped 闭环的量，不能只按日期粗判。
- 有效 LP 日期只有在 open WIP Scheduled 量上，才参与 gap 计算。

## 9. 已知易错点

| 位置 | 易错点 | 修正 |
|---|---|---|
| SC `Order Status` | 值是 `Carryover`，不是 `carry over` | 精确匹配 |
| SC `Carryover` | 有值时可能是数字，不是文本 | 用真正空值判断 |
| SC 表头 | 可能有换行或插列 | 使用表头标准化查找 |
| SO 编号 | Excel 可能存成浮点数 | 转成整数再转字符串 |
| Shipped SO 列 | 可能存在两个销售订单列 | 优先取值以 `10` 开头的列 |
| IDN Schedule Planning Dispatch | `ORDER OUTSTANDING ` 有尾随空格；`DISPATCH` 是同一工作簿里的已发运记录，需要 append | 精确匹配 `ORDER OUTSTANDING `；同时读取 `DISPATCH`，按表头对齐，并且本月主口径只保留有效 `ELD >= 本月第一天` |
| SC Loading Date | 不可靠，不作为 LP deadline | 只用于明确的 SC 上月交付预分配 |
| Raw source overlap | 同一 SO 可能同时出现在 Shipped / FG / PP | KPI 使用 allocated waterfall，raw 保留审计 |
| SC Delivery PCS | 负数可能是反冲 / 修正 | 上月交付预分配时排除 |

## 10. Phase 2 路线：数据消费层

Excel、HTML、PBI 都是呈现层。后续应沉淀一套机器可消费的数据层，让所有前端都消费同一套业务逻辑。

Phase 2 设想结构：

```text
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

建议：

- 主明细数据优先使用 CSV，便于 Excel、PBI、HTML 读取。
- YAML 或 JSON 只记录运行元数据，例如源文件、行数、参数、生成时间。
- Parquet / SQLite 可以作为未来增强，不进入当前 Phase 1。
