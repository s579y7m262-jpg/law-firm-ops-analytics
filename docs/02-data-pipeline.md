# 02 数据链路：从 CRM 导出到看板

## 1. 数据源

| 文件 | 粒度 | 规模 | 用途 |
|---|---|---|---|
| `订单表_2026.1-2026.9.xlsx` | 一行 = 一次约见/一个案件 | 166 列 / 1,357 行 | 约见、差旅费、合同额、咨询费 |
| `订单表_2026.10.xlsx` | 同上 | 64 行 | 补 10 月（与 1–9 月无订单编号重叠） |
| `收入表_2026.1-2026.10.xlsx` | 一行 = 一笔认款 | 每月一个 sheet，2,702 行 | 业绩、成案、客单价（**主口径**） |
| `花名册.xlsx` | 一行 = 一名员工 | 955 人 | 人员 → 小组 → 区域 映射 |
| `结案退费返点汇总表.xlsx` | 每月一个 sheet，混装三类 | 14 个 sheet | 退费（月度完整）、结案与返点（按批次记账，不完整） |

关键差异：**订单表记"签了什么"，收入表记"认了多少钱"**。二者不能混用，否则业绩会被重复计一遍。

## 2. 三层结构

```
Bronze  原始导出（不做任何修改，只做只读加载）
  │
Silver  清洗：金额文本解析 / 空值唯一键 / 列名归一化 / 组织回挂
  │
Gold    fact_all  业绩事实·行级   收入表全部行求和
        fact_new  成案事实·案件级 金额按行、件数按业绩比例拆
        fact_fee  咨询费事实·线索级 三规则并集，可重复计入
  │
  └── build/payload.json  →  看板（单文件 HTML）+ 底表（13 张 sheet）
```

## 3. 连接键与命中率

| 连接 | 命中 | 说明 |
|---|---|---|
| `案件编号` 订单表 ↔ 收入表 | 965 / 973（99.2%） | 主键，唯一可靠的跨表锚点 |
| `订单归属人` → 花名册 `坐席` | 90.9% | 未命中的多为离职/兼岗，进「其他（未归属）」 |
| 花名册 `小组` ↔ 收入表 `所属小组` | 59 / 68 完全对齐 | 同一套命名体系，故小组级可直接并排比较 |
| 订单表 `归属部门` → 区域 | 仅 73% | 与收入表「小组」不是同一套命名（如「A区十组」vs「A区行政八组」），只作兜底 |

> 组织回挂的教训：**能用同一套命名体系连，就不要用两套做模糊匹配**。V5 之前咨询费停在"人→区域"，就是因为绕道走了「归属部门」；改走「人→花名册小组→区域」后映射率从 94.7% 升到 99.4%，小镇级下钻一次打通。

## 4. 脚本与产物

| 脚本 | 输入 | 输出 |
|---|---|---|
| `src/etl/build_gold.py` | 4 张源表 | `build/payload.json`、`new_cases.csv`、`consult.csv` |
| `src/render.py` | payload + `src/templates/*.html` | `dashboards/01~05_*.html` |
| `src/excel/build_workbook.py` | payload + 3 个 csv | `build/指标底表.xlsx`（13 sheet） |
| `src/analysis/attribution.py` | payload + 订单表 + 台账 | `build/payload_attribution.json` |
| `src/analysis/censoring_forecast.py` | 订单表 | `build/payload_censoring.json` |
| `src/analysis/revenue_scenarios.py` | 收入表 + 订单表 | `build/payload_scenarios.json` |
| `src/analysis/payment_capacity.py` | 订单表 | `build/payload_payment.json` |
| `src/analysis/region_deepdive.py` | 2 个 csv | `build/区域经营分析_*.xlsx` |
| `src/monitor/monitor_etl.py` | 订单表 + 收入表（自动发现） | `dashboards/monitor.html` + `monitor_data.js` |

## 5. 为什么看板是单文件 HTML

受众是律所合伙人/销售负责人，电脑上没有 Python、没有 BI 账号、IT 也不给装。
所以每张看板把样式、数据、图形**全部内联进一个 HTML**：双击即开、可微信发送、可离线、断网也能看。
代价是文件 40–700 KB，但对这个场景是正确的取舍——**能把结论送到决策者眼前，比技术优雅重要**。

图形全部手写 SVG，不引第三方图表库：一是断网可用，二是能精确控制标注位置（见 [05](05-data-quality.md) 的呈现陷阱）。
