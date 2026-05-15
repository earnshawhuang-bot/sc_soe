# S&OE Dashboard 解读规则

## 1. 这份 Dashboard 解决什么问题

这份 HTML dashboard 的核心问题只有一个：

> 本月 SC baseline 订单能不能完成发运兑现？如果不能，风险卡在哪个环节，应该由谁跟进？

它不是一张“数据展示页”，而是一张“管理行动页”。
看它时不要先问“有多少 SO”，而要先问：

```text
有多少 MT 已经兑现？
有多少 MT 已经具备发运条件？
有多少 MT 仍在生产或没有供应信号？
这些未发货的 MT 是否有 Loading Plan？
如果有 LP，日期是否明确、是否已经过期、生产是否赶得上？
```

本页的主线是 baseline-led，也就是以本月 SC baseline 责任池为起点。Loading Plan 不是另一个执行状态，而是用来交叉验证发运安排是否闭环。

## 2. 三个基础口径

### 2.1 Data Date

页面顶部的 `Data as of` 是本轮数据包的截止日。

当前规则是从 `02-Shipped` 文件夹中最新 GI 文件名后缀推导，例如：

```text
GI (KS&IDN) 260501-0513.xlsx -> data_date = 2026-05-13
```

它不是电脑当天日期。
原因是业务数据通常有更新节奏，dashboard 应该以源数据的实际截止日为准。

### 2.2 Baseline

`Baseline MT` 是本月开票责任池，也就是本月管理上需要兑现的订单量。

它来自 SC baseline，不来自 Loading Plan。
所以所有履约判断都先问：

```text
这个 SO 是否属于本月 SC baseline？
```

### 2.3 Loading Plan

Loading Plan 是发运安排底账，用来回答：

```text
这个订单有没有发运安排？
发运日期是否明确？
日期是在 data_date 之前，还是 data_date 当天及以后？
```

Loading Plan 不等于已发货，也不等于在库或在产。
它和 Shipped / FG / PP 是交叉检查关系。

## 3. 推荐阅读顺序

建议从上到下按四层阅读，不要一上来就点明细。

```text
第一层：看总量是否闭环
第二层：看工厂差异
第三层：看已发货数据闭环
第四层：看未发货风险矩阵并下钻行动
```

这个顺序对应管理层的自然问题：

```text
总盘子多大？
现在兑现到哪里？
KS / IDN 谁压力更大？
已发货是否能被 LP 解释？
未发货的风险到底在物流、计划、生产，还是销售订单归属？
```

## 4. 顶部 KPI 怎么读

顶部 KPI 是第一层闭环，回答本月 baseline 执行到哪里。

| KPI | 业务含义 | 怎么看 |
|---|---|---|
| Baseline Orders | 本月责任池总量 | 分母，所有状态最终都要对回它 |
| Shipped | 已兑现量 | 已经发货，属于结果 |
| In Stock (FG) | 在库量 | 货已准备，主要看是否安排发运 |
| In Production | 已排产 WIP | 生产中，主要看能否赶上 loading date |
| Scheduling | 未明确供应量 | 包含 WIP Unscheduled + No Plan，是需要拆解的风险池 |

这里的状态是 MECE 的。
也就是说：

```text
Shipped + FG + WIP + Unscheduled + No Plan = Baseline
```

注意：这里的 `Scheduling` 是管理层概览用的合并显示。真正行动时，需要在下面矩阵里拆成：

```text
WIP Unscheduled
No Supply Signal
```

因为这两类的行动逻辑不同。

## 5. By Plant 行级汇总怎么读

`BY PLANT` 是顶部 KPI 的工厂拆解，不是新的风险矩阵。

它回答：

> KS 和 IDN 各自的 baseline 执行结构是什么？

| 字段 | 含义 |
|---|---|
| Plant | 工厂 |
| Baseline MT | 该工厂本月 baseline |
| Shipped | 已发货 |
| FG | 已在库 |
| WIP | 已排产生产中 |
| Unscheduled | 有工单但没有明确完工时间 |
| No Plan | 没有供应信号 |
| Total | 当前 Plant filter 下的合计行 |

读法建议：

```text
先看 Baseline 谁大；
再看 Shipped + FG 占比；
再看 WIP / Unscheduled / No Plan 哪个工厂更重。
```

这张表适合用来判断“问题主要在哪个工厂”，但它不回答“是否有 Loading Plan”。
Loading Plan 的交叉判断要看下面的风险矩阵。

`Total` 行是 plant 级别汇总。
如果选择 `All Plants`，它等于 KS + IDN；如果只选择 `KS` 或 `IDN`，它等于当前被筛选工厂的汇总。这个设计是为了让业务方不用手工加总，也能快速对回顶部 KPI。

## 6. Shipped Data Closure 怎么读

`Shipped Data Closure` 只看已发货量。

它不是 open risk。
因为货已经发了，执行结果已经发生。

这块回答的是：

> 已发货的量，能不能在 Loading Plan 侧找到解释？

| 字段 | 含义 |
|---|---|
| Total MT | 已发货 baseline 量 |
| LP Closed / Matched | 能被当前 LP 证据闭环的已发货量 |
| LP Missing or Unclear | 已发货但 LP 缺失或 LP 日期不清晰的量 |

业务解读：

```text
LP Closed / Matched 高，说明 shipped 和 LP 之间闭环较好。
LP Missing or Unclear 高，不代表货没发，而代表数据链路需要核对。
```

这块的 owner 通常不是生产，而是 Sales Ops / Logistics / 数据口径维护方。

## 7. Open Fulfillment Risk Matrix 怎么读

这是 dashboard 的核心区。

它只看未发货的 baseline 量，回答：

> 未发货的量，货源状态是什么？发运安排状态是什么？风险应该怎么处理？

矩阵行是供应执行状态：

| 行 | 业务含义 |
|---|---|
| FG | 货已在库，理论上可发 |
| WIP Scheduled | 生产中，且有计划结束时间 |
| WIP Unscheduled | 有工单，但无明确结束时间 |
| No Supply Signal | 没有发货、库存、排产、工单信号 |

矩阵列是 Loading Plan 覆盖状态：

| 列 | 业务含义 |
|---|---|
| Past Due LP | 有明确 LP 日期，但日期早于 data_date，且未被 shipped 闭环 |
| Future Valid LP | 有明确 LP 日期，且日期在 data_date 当天或之后 |
| LP Date Unconfirmed | 有 LP 记录，但日期是 TBA / 空 / 非标准文本 |
| No Current LP | 当前主口径下没有可用 LP 覆盖 |

每个格子都是 MT，不是 SO 数。
每行应满足：

```text
Supply Status Total
= Past Due LP
+ Future Valid LP
+ LP Date Unconfirmed
+ No Current LP
```

这张矩阵的底表是 Excel 里的 `Risk Matrix Detail`。
页面上的矩阵数字等于按 `supply_status` 和 `lp_coverage_status` 对 `risk_mt` 求和；点击格子后，下方明细也是同一张底表的筛选结果。

### 7.1 怎么判断优先级

优先看 MT 大的格子，再看业务风险严重性。

一般优先顺序建议：

```text
1. Past Due LP
2. No Supply Signal + Future Valid / Past Due LP
3. FG + No Current LP
4. WIP Scheduled + Future Valid LP 且 Gap < 0
5. LP Date Unconfirmed
6. WIP Unscheduled / No Current LP
```

原因：

- `Past Due LP` 已经过期，优先核对是否实际发货、是否数据未更新、是否需要重排。
- `No Supply Signal + 有 LP` 说明业务有发运安排，但供应侧完全没有信号，是强风险。
- `FG + No Current LP` 说明货已经在库，但没有发运安排，通常是物流/业务安排问题。
- `WIP Scheduled + Future Valid LP` 要看生产是否赶得上 loading date。
- `LP Date Unconfirmed` 不是没安排，而是日期不清晰，需要业务确认。

## 8. 点击矩阵后怎么看明细

点击任意矩阵格子，下方 `Selected Risk Detail` 会展示对应清单。

这张明细不是固定列，而是按场景动态展示。
原则是：

> 只展示当前风险判断真正需要看的字段，避免把背景字段误读成事实。

通用字段：

| 字段 | 用途 |
|---|---|
| SO | 定位订单 |
| Plant | 工厂 |
| Cluster | 区域 / 客户群 |
| Order Type | 订单类型 |
| Supply | 当前供应状态 |
| LP Coverage | 当前 LP 覆盖状态 |
| Risk MT | 当前格子中需要关注的吨数 |
| Risk Action | 建议行动分类 |
| Owner | 建议责任方 |
| Action Note | 业务解释 |

### 8.1 WIP Scheduled 场景

当供应状态是 `WIP Scheduled` 时，明细会额外展示：

| 字段 | 含义 |
|---|---|
| Machine | 生产机器 |
| Planned End | PP 底表 `Planned EndTime`，按 SO 取最晚时间 |
| Available Date | `Planned End + 1 day` |
| Loading Date | LP 要求的装载日期 |
| Gap | `Loading Date - Available Date` |
| Work Order | 工单号 |

读法：

```text
先看 Risk MT 多大；
再看 Machine / Work Order 定位生产对象；
再看 Planned End 和 Available Date；
最后看 Loading Date 和 Gap。
```

`Planned End` 代表 Lami 计划结束时间。
`Available Date` 是考虑次日可入库/可发运后的参考时间。

当前 LP loading date 多数只有日期，所以 gap 按日期比较：

```text
Loading Date = 2026-05-20
Available Date = 2026-05-20 23:20
Gap = 0
```

不会因为 `23:20` 晚于 `00:00` 就误判为 -1。
只有未来 LP 也提供具体装柜小时分钟时，才升级到小时级比较。

Gap 颜色：

| Gap | 含义 |
|---:|---|
| `< 0` | 生产赶不上 loading date |
| `0-2` | 紧张但可能可行 |
| `> 2` | 有缓冲 |

### 8.2 FG 场景

FG 已经在库，所以通常不需要看机器、工单、planned end。

重点看：

```text
有没有 LP？
LP 日期是否明确？
如果没有 LP，为什么货在库但没有发运安排？
```

典型行动：

```text
FG + No Current LP -> Logistics / Sales 确认发运安排
FG + LP Date Unconfirmed -> 确认 loading date
FG + Past Due LP -> 确认是否已发未更新，或 LP 是否需要重排
```

### 8.3 WIP Unscheduled 场景

WIP Unscheduled 表示有工单，但没有明确计划结束时间。

重点不在 gap，因为没有 planned end 无法算 gap。
重点是：

```text
为什么工单没有排出结束时间？
如果已有 LP，生产是否能给出计划？
如果没有 LP，生产和发运都需要补计划。
```

### 8.4 No Supply Signal 场景

No Supply Signal 是最底层的供应信号缺失。

它表示当前没有找到：

```text
Shipped
FG
Scheduled PP
Unscheduled WO
```

如果它同时有 `Future Valid LP` 或 `Past Due LP`，说明业务侧有发运安排，但供应侧没有任何支撑信号，需要计划端优先核对。

如果它是 `No Current LP`，则说明：

```text
baseline 有订单，
但供应和发运两边都没有信号。
```

这类通常要回到 SC / PP / 业务确认订单是否真实需要本月兑现。

## 9. LP Not In Current SC Baseline 怎么读

这块不属于 baseline 履约风险矩阵。
它是 LP-led 的对账异常。

它回答：

> Loading Plan 里有发运需求，但这些 SO 不在本月 SC baseline 里。

常见原因：

| 原因 | 说明 |
|---|---|
| 未来月订单 | LP 包含未纳入本月开票责任池的订单 |
| baseline 漏单 | SC baseline 可能没有包含这部分 SO |
| 跨月安排 | LP 安排可能跨月，但业务归属尚未确认 |
| SO 解析问题 | Invoice / SC No. 拆解导致匹配不到 baseline |

这块建议看 MT，不先看 SO Count。
如果需要核对明细，点击对应格子或下载 CSV。

## 10. Plant Filter 和下载怎么用

### 10.1 Plant Filter

Plant filter 会同时影响：

- 顶部 KPI
- BY PLANT
- Shipped Data Closure
- Open Fulfillment Risk Matrix
- LP Not In Current SC Baseline
- Selected Risk Detail

如果选择 `KS`，页面只看 KS。
如果选择 `IDN`，页面只看 IDN。
如果选择 `All Plants`，页面看整体。

### 10.2 Download CSV

`Download CSV` 导出当前明细视图。

规则：

| 当前状态 | 下载内容 |
|---|---|
| 没有点击格子 | 默认 action-required 明细 |
| 点击风险矩阵格子 | 当前 Plant + 当前格子的风险明细 |
| 点击 Shipped closure | 当前 Plant + 已发货闭环明细 |
| 点击 LP Not In Baseline | 当前 Plant + LP-led 异常明细 |

CSV 的列与 dashboard 当前展示列一致，不等于 Excel 底表的全部字段。
如果要做完整透视和审计，仍然建议使用 Excel 的 `Risk Matrix Detail`。

## 11. 不要误读的地方

### 11.1 Loading Plan 不是执行状态

不要把 Loading Plan 当成 Shipped / FG / WIP 的同级状态。
它是发运安排维度。

正确理解是：

```text
供应状态回答：货在哪里？
LP 状态回答：发运有没有安排？
```

### 11.2 No Current LP 不等于完全没有历史 LP

`No Current LP` 表示当前主口径下没有可用 LP 覆盖。
它不一定说明这个 SO 从来没有出现在任何 LP 原始表中。

可能原因包括：

- 只有历史有效 LP
- 被 prior invoiced 排除
- LP 量为 0
- 原始编号无法解析

### 11.3 Shipped 的 LP 缺失不是履约风险

Shipped 已经是结果。
如果已发货但 LP 缺失，这是数据闭环问题，不是“货发不出去”的执行风险。

### 11.4 FG 不需要看 Planned End

FG 已经在库。
对 FG 来说，planned end / machine / work order 不是当前风险判断重点。

FG 的核心问题是：

```text
货已经好了，为什么还没有明确 loading？
```

### 11.5 WIP Scheduled 才需要看 Gap

Gap 只对 `WIP Scheduled + 有明确 LP loading date` 有业务意义。
因为只有这种情况下，系统同时知道：

```text
生产预计什么时候结束
发运要求什么时候装载
```

## 12. 推荐管理动作闭环

每次 review dashboard，可以按这个节奏开会：

```text
1. 确认 baseline 总量和 data_date
2. 看顶部 KPI，确认已发货 / 在库 / 在产 / 待安排结构
3. 看 BY PLANT，判断 KS / IDN 哪边压力更大
4. 看 Shipped Data Closure，确认已发货数据是否能解释
5. 看 Open Fulfillment Risk Matrix，按 MT 找最大风险格子
6. 点击格子，下钻 SO 明细
7. 按 Risk Action / Owner 分派任务
8. 下载 CSV 给对应业务方核对
9. 下轮刷新数据，看风险格子的 MT 是否下降
```

最终目标不是解释所有数字，而是推动风险 MT 下降。

## 13. 一句话总结

这份 dashboard 的阅读逻辑是：

```text
先用 SC baseline 定义本月责任池，
再用供应瀑布看货在哪里，
再用 Loading Plan 判断发运安排是否闭环，
最后用风险矩阵把未兑现量分派到具体行动队列。
```

管理层看 summary，业务方点格子看明细，执行团队下载 CSV 逐条核对。
