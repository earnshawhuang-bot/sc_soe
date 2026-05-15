# Risk Matrix 使用说明

## 1. 这张表解决什么问题

`Risk Matrix Detail` 用来回答：

> 本月 baseline 订单里，未完成发货的量卡在哪个环节？Loading Plan 是否支持这些量？需要谁跟进？

它不是普通的 SO 清单，而是 Summary 风险矩阵的事实底表。可以把它理解成 PBI 里的明细事实表：Summary 只是透视后的管理视图，真正能追溯到 SO 和吨数的是 `Risk Matrix Detail`。

## 2. Summary 应该先看哪几块

### 2.1 Baseline Execution Status

第一块先回答：本月 baseline 订单执行到哪里了？

| 状态 | 含义 |
|---|---|
| Shipped | 已经发货，属于已兑现 |
| FG | 货已在库，理论上具备发运条件 |
| WIP Scheduled | 生产中，且已有计划完工日期 |
| WIP Unscheduled | 有工单，但没有明确完工日期 |
| No Supply Signal | 没有发货、库存、排产、工单信号 |

这一块必须 MECE。也就是说，每一吨 baseline 量只能落在一个供应状态里，所有状态相加要对回 baseline 总量。

### 2.2 Shipped Data Closure

`Shipped` 已经是结果，不再放进 open risk 矩阵里。

它单独作为闭环审计：

| Supply Status | Total MT | LP Closed / Matched | LP Missing or Unclear |
|---|---:|---:|---:|
| Shipped | 已发货量 | 能被 LP 证据解释的量 | 找不到可用 LP 证据或 LP 日期不清晰的量 |

这块回答的是：

> 已发货的数据，Loading Plan 侧能不能解释得上？

IDN 新增 `DISPATCH` sheet 后，会补齐印尼已经发货的 LP 记录，减少“已发货但 LP 缺失”的假异常。

### 2.3 Open Fulfillment Risk Matrix

未发货的量进入 open risk matrix：

| Supply Status | Supply Status Total | Past Due LP | Future Valid LP | LP Date Unconfirmed | No Current LP |
|---|---:|---:|---:|---:|---:|
| FG | 数字 MT | 数字 MT | 数字 MT | 数字 MT | 数字 MT |
| WIP Scheduled | 数字 MT | 数字 MT | 数字 MT | 数字 MT | 数字 MT |
| WIP Unscheduled | 数字 MT | 数字 MT | 数字 MT | 数字 MT | 数字 MT |
| No Supply Signal | 数字 MT | 数字 MT | 数字 MT | 数字 MT | 数字 MT |
| Open Total | 数字 MT | 数字 MT | 数字 MT | 数字 MT | 数字 MT |

读法：

```text
Supply Status Total
= Past Due LP
+ Future Valid LP
+ LP Date Unconfirmed
+ No Current LP
```

每个格子都必须是纯数字 MT，不要写 `MT` 文本，也不要把 SO 数塞进格子里。

这张矩阵不是单独手工汇总出来的表。它的底表仍然是 Excel 里的 `Risk Matrix Detail`：

```text
Risk Matrix Detail
  -> 按 supply_status 作为行
  -> 按 lp_coverage_status 作为列
  -> 对 risk_mt 求和
  -> 得到 Open Fulfillment Risk Matrix
```

所以截图里的 `FG × No Current LP = 787` 这类数字，本质上就是在 `Risk Matrix Detail` 里筛选：

```text
supply_status = FG
lp_coverage_status = No Current LP
sum(risk_mt)
```

Dashboard 点击某个格子后，下方明细也来自同一张事实底表，只是按当前格子的筛选条件把相关行列出来。

### 2.4 LP Not In Current SC Baseline

这一块是 Loading Plan-led 的对账异常，不属于 `Open Fulfillment Risk Matrix`。

它回答的是：

> Loading Plan 里有发运安排，但这些 SO 不在本月 SC baseline 里，这些量是什么？

推荐 summary：

| LP Status | MT |
|---|---:|
| Past Due LP | 数字 MT |
| Future Valid LP | 数字 MT |
| LP Date Unconfirmed | 数字 MT |
| Total | 数字 MT |

这里不放 SO Count，管理层先看量。SO 数如果需要，可以后续在底表或下载明细里自行透视。

推荐明细粒度是 Loading Plan 行级，而不是 SO 聚合级：

```text
一条 current-scope Loading Plan 明细行
```

推荐明细字段：

```text
SO / Plant / LP Source / Source Sheet / Invoice No Raw
Loading Date Raw / Loading Date / LP Date Status / Load MT
```

原因是 loading date 和 load MT 本来就是 LP 行级信息；如果先聚合到 SO，多个 loading date 和多个来源会被揉在一起，反而不方便业务核对。

## 3. 四个 LP 覆盖状态怎么理解

| 字段 | 业务含义 |
|---|---|
| Past Due LP | 有有效 LP 日期，且 loading date 已早于 data_date，但这部分量没有被 shipped 闭环 |
| Future Valid LP | 有有效 LP 日期，且 loading date 在 data_date 当天或之后 |
| LP Date Unconfirmed | 有 LP 记录，但日期是 TBA / 空 / 非标准文本 |
| No Current LP | 当前主口径下没有可用 LP 覆盖 |

关键点：

```text
Past Due LP 不是只看 loading date < data_date。
Past Due LP = 已过期有效 LP 量 - 已被 shipped 闭环的量。
```

这里的 `data_date` 来自 `02-Shipped` 文件夹最新 GI 文件后缀。例如最新 shipped 文件是 `GI (KS&IDN) 260501-0513.xlsx`，则本轮 `data_date = 2026-05-13`。

这样可以避免把“已经发掉、只是 LP 日期在过去”的量误判成风险。

## 4. Risk Matrix Detail 怎么看

`Risk Matrix Detail` 的粒度是：

```text
一个 SO + 一个 allocated 供应分段 + 一个 LP 覆盖分段
```

所以同一个 SO 可能出现多行。原因是一个 SO 的数量可能一部分已发货、一部分在库、一部分排产、一部分还没有供应信号；同时 LP 也可能一部分已经过期、一部分未来有效、一部分日期未确认。

核心字段：

| 字段 | 怎么看 |
|---|---|
| `supply_status` | 当前这一行的供应状态，比如 FG / WIP Scheduled / No Supply Signal |
| `lp_coverage_status` | 当前这一行对应的 LP 覆盖状态 |
| `risk_mt` | 当前这一行对应的吨数，是最重要的管理量 |
| `covered_mt` | 同一个 SO 中已经被更高优先级供应状态覆盖的量 |
| `so_total_mt` | 这个 SO 的总 baseline 量，只作背景 |
| `risk_action` | 这一行对应的行动分类 |
| `suggested_owner` | 建议责任方，例如物流、计划、工厂或业务 |
| `action_note` | 简短业务解释 |

最重要的组合是：

```text
supply_status + lp_coverage_status + risk_mt
```

### 4.1 Excel 底表和 Dashboard 明细的区别

`Risk Matrix Detail` 在 Excel 里是底表，字段保持稳定，方便筛选、透视和审计。

Dashboard 里的点击明细是业务工作台，不需要把所有字段都固定展示。它会根据点击的矩阵格子动态显示字段，只展示当前场景真正有意义的信息。

通用规则：

| 字段 | Dashboard 展示规则 |
|---|---|
| `risk_mt` | 所有场景都显示，是主行动量 |
| `planned_end_date` | 只在 `WIP Scheduled` 场景显示 |
| `available_date` | 只在 `WIP Scheduled` 场景显示，等于 `planned_end_date + 1 day` |
| `machines` | 只在 `WIP Scheduled` 场景显示 |
| `work_orders` | 只在 `WIP Scheduled` 场景显示 |
| `loading_date` | 只在 `Past Due LP` 或 `Future Valid LP` 场景显示 |
| `lp_loading_date_raw_list` | 只在 `LP Date Unconfirmed` 场景显示 |
| `lp_loading_date_status_mix` | 只在 `LP Date Unconfirmed` 场景显示 |
| `lp_gap_days` | 只在 `WIP Scheduled + 有明确 loading date` 时显示和高亮 |
| `risk_action / suggested_owner / action_note` | 所有场景都显示 |

举例：

| 点击的矩阵格子 | Dashboard 明细额外显示 |
|---|---|
| `FG × Future Valid LP` | `Loading Date` |
| `FG × No Current LP` | 不显示生产字段，也不显示 Loading Date |
| `WIP Scheduled × Future Valid LP` | `Machine / Planned End / Available Date / Loading Date / Gap / Work Order` |
| `WIP Scheduled × No Current LP` | `Machine / Planned End / Available Date / Work Order` |
| `LP Date Unconfirmed` | `LP Date Raw / LP Date Status` |

这条规则的目的，是避免把 SO 级背景字段误读成当前风险分段的事实。例如 FG 已经在库，就不应该再让业务看机器、工单和计划完工日；`No Current LP` 没有当前可用 LP 覆盖，就不应该展示 loading date。

`WIP Scheduled` 的 dashboard 明细列顺序需要服务业务判断，推荐顺序是：

```text
SO / Plant / Cluster / Order Type / Supply / LP Coverage / Risk MT
Machine / Planned End / Available Date / Loading Date / Gap
Risk Action / Owner / Work Order / Action Note
```

这个顺序的含义是：先定位 SO 和风险量，再看机器、计划完工日、考虑缓冲后的可发运日期、loading 要求和 gap，随后看行动标签、责任方，最后用工单号作为追溯和执行对象。

口径：

```text
Planned End = PP 底表 Planned EndTime，按 SO 取最晚时间，代表 Lami 结束时间
Available Date = Planned End + 1 day
Gap = Loading Date - Available Date
```

PP 底表里的 `Planned EndTime` 带小时分钟/秒，Excel 和 dashboard 应至少保留到分钟级展示。由于当前 Loading Plan 的 loading date 多数是日期级，gap 判断先按日期比较；只有未来 LP 也提供明确小时分钟时，才升级为小时级比较。

Dashboard 视觉风格建议采用管理层汇报用色：暖白/象牙色背景、深海军蓝标题、青蓝与金色作为主要对比色、浅暖灰表格线，避免过重的纯蓝表头。

Dashboard 顶部的 `BY PLANT` 是行级汇总，用来把第一层 `Baseline Execution Status` 按工厂拆开：

| Plant | Baseline MT | Shipped | FG | WIP | Unscheduled | No Plan |
|---|---:|---:|---:|---:|---:|---:|
| IDN | 数字 MT | 数字 MT | 数字 MT | 数字 MT | 数字 MT | 数字 MT |
| KS | 数字 MT | 数字 MT | 数字 MT | 数字 MT | 数字 MT | 数字 MT |
| Total | 数字 MT | 数字 MT | 数字 MT | 数字 MT | 数字 MT | 数字 MT |

它不是新的风险矩阵，只是帮助业务方快速看 KS / IDN 各自的执行状态。单元格保持纯数字，单位放在标题或表头。

`Total` 行是 plant 级别汇总行，用来快速确认所有工厂相加是否能对回顶部 KPI。Plant filter 会影响这张表：

```text
All Plants -> 显示 KS / IDN / Total
KS -> 显示 KS / Total
IDN -> 显示 IDN / Total
```

因此 `BY PLANT` 的作用是回答“哪个工厂压力更大、各工厂执行结构如何”，而 `Open Fulfillment Risk Matrix` 才回答“这些未发货量的风险动作是什么”。

Dashboard 的 `Selected Risk Detail` 支持下载当前明细：

- 下载按钮只导出当前 Plant filter 和当前点击格子对应的明细。
- 如果没有点击格子，则导出默认的 action-required 明细。
- 下载格式为 CSV，方便直接用 Excel 打开或发给业务方核对。
- CSV 使用 dashboard 当前动态展示的列，而不是 Excel 底表的全部稳定字段。
- HTML dashboard 是单文件静态快照，可以直接发给同事用 Edge / Chrome 打开；但由于明细数据已嵌入 HTML，分发时要注意权限。

## 5. 用 WIP Scheduled 举例

假设 Summary 里 `WIP Scheduled` 这一行如下：

| Supply Status | Supply Status Total | Past Due LP | Future Valid LP | LP Date Unconfirmed | No Current LP |
|---|---:|---:|---:|---:|---:|
| WIP Scheduled | 8,693 | 1,200 | 3,600 | 2,948 | 945 |

这行可以读成：

```text
本月 baseline 中，有 8,693 MT 处于已排产状态。
其中：
1,200 MT 的 LP 已过期，且没有被 shipped 闭环。
3,600 MT 有未来有效 loading date，需要看生产完工日是否赶得上。
2,948 MT 有 LP 记录，但日期未确认。
945 MT 当前没有可用 LP 覆盖。
```

### 5.1 WIP Scheduled + Past Due LP

业务含义：

> 生产有计划，但发运日期已经过了，且 shipped 侧没有闭环。

这类要优先看：

```text
loading_date
data_date
planned_end_date
available_date
shipped_closed_mt
risk_mt
```

常见行动是确认：是否实际已发但 shipped 未更新，还是发运计划已经过期需要重排。

### 5.2 WIP Scheduled + Future Valid LP

业务含义：

> 生产已经排了，也有未来明确 loading date。

这类需要看：

```text
planned_end_date
available_date
loading_date
lp_gap_days
```

如果 `lp_gap_days < 0`，说明考虑 `Available Date` 后赶不上 loading date，会进入 `WIP Late vs Loading Date`。

### 5.3 WIP Scheduled + LP Date Unconfirmed

业务含义：

> 生产已经排了，但发运日期没有确认。

这类通常不是生产没有动作，而是物流 / 业务侧需要确认 loading date。

### 5.4 WIP Scheduled + No Current LP

业务含义：

> 生产已经排了，但当前 Loading Plan 里没有可用发运安排。

这类通常要确认：货生产出来后准备什么时候发？是否漏进 Loading Plan？是否在其他版本或历史 LP 中？

## 6. 最后怎么用

建议使用顺序：

1. 先看 `Baseline Execution Status`，确认总量和供应状态是否闭环。
2. 再看 `Shipped Data Closure`，确认已发货量是否能被 LP 解释。
3. 再看 `Open Fulfillment Risk Matrix`，找最大风险格子。
4. 到 `Risk Matrix Detail` 按 `supply_status` 和 `lp_coverage_status` 筛选。
5. 按 `risk_mt` 从大到小处理，再看 `risk_action / suggested_owner / action_note`。
6. 如果使用 dashboard，直接点击风险矩阵中的格子，下面明细会自动收起不适用字段。

一句话：

> Summary 告诉你风险在哪个格子，Risk Matrix Detail 告诉你具体是哪几个 SO、多少吨、该找谁处理。
