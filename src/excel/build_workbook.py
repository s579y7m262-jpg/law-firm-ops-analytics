# -*- coding: utf-8 -*-
"""指标底表（Excel）：把 build/payload.json 摊成 13 张可审计的工作表。
每张表都能回溯到看板的某一块，供业务复核口径用。
"""
import os, json
import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
BUILD = os.path.join(ROOT, "build")
P = json.load(open(os.path.join(BUILD, "payload.json"), encoding="utf-8"))
OUTF = os.path.join(BUILD, "指标底表.xlsx")

NAVY = "1F3864"
BLUE = "2F5BEA"
thin = Side(style="thin", color="E6E8EC")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)


def title(ws, text, sub=""):
    ws["A1"] = text
    ws["A1"].font = Font(size=14, bold=True, color=NAVY)
    ws.row_dimensions[1].height = 26
    if sub:
        ws["A2"] = sub
        ws["A2"].font = Font(size=9.5, color="5B6572")


def write_table(ws, df, start_row, widths=None, num_fmt=None):
    cols = list(df.columns)
    for j, c in enumerate(cols, 1):
        cell = ws.cell(row=start_row, column=j, value=c)
        cell.font = Font(bold=True, color="FFFFFF", size=10)
        cell.fill = PatternFill("solid", fgColor=BLUE)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
    ws.row_dimensions[start_row].height = 28
    for i, row in enumerate(df.itertuples(index=False), start_row + 1):
        for j, v in enumerate(row, 1):
            if isinstance(v, (np.integer,)):
                v = int(v)
            elif isinstance(v, (np.floating,)):
                v = float(v)
            if isinstance(v, float) and (np.isnan(v) or np.isinf(v)):
                v = None
            cell = ws.cell(row=i, column=j, value=v)
            cell.font = Font(size=10)
            cell.border = BORDER
            if j > 1:
                cell.alignment = Alignment(horizontal="right")
                if num_fmt and isinstance(v, (int, float)):
                    cell.number_format = num_fmt
            else:
                cell.alignment = Alignment(horizontal="left")
        if (i - start_row) % 2 == 0:
            for j in range(1, len(cols) + 1):
                ws.cell(row=i, column=j).fill = PatternFill("solid", fgColor="FAFBFD")
    if widths:
        for j, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(j)].width = w
    ws.auto_filter.ref = f"A{start_row}:{get_column_letter(len(cols))}{start_row+len(df)}"
    ws.freeze_panes = ws.cell(row=start_row + 1, column=1)


wb = Workbook()

# ── 00 口径说明 ──
ws = wb.active
ws.title = "00_口径说明"
ws.sheet_view.showGridLines = False
title(ws, "2026 经营指标底表 · 口径说明", "窗口 2026.1–10 ｜ 生成 " + P["meta"]["generated_at"])
rows = [
    ("指标 / 规则", "定义", "说明"),
    ("总业绩", "收入表「业绩金额」全部行求和", "含新接、中期尾款、新增户、后期效果费、续约等全部款项"),
    ("成案量", "新接案件按业绩比例拆分的件数（V 列）",
     "参考_成案量口径：W=同案件含新接金额合计，X=同案件新接案件记录数；单记录 V=1，多记录 V=本行金额÷W"),
    ("成案金额", "新接记录金额按行求和", "拆分后各组认款额相加即为案件金额"),
    ("客单价", "成案金额 ÷ 成案量", "元/件"),
    ("咨询费数量", "订单表满足任一条件、按差旅码去重的线索数",
     "①差旅费金额>0 ②主场约见且合同额≥5000 ③直播签约且合同额≥5000；覆盖 1–10 月"),
    ("", "", ""),
    ("件数边界修正①", "参考表对「新接案件」标签直接给 V=1",
     "会使 43 个「两行都标新接案件」的案件各算 2 件。已改为：X≥2 时一律按金额比例拆"),
    ("件数边界修正②", "2 个案件仅有 1 条「新接案件-撞单」记录（202608031、GS2605034）",
     "参考公式判 V=0 会丢件。已改为：X<2 时算 1 件。最终每案恰 1 件，合计 1,115"),
    ("撞单处理", "多行按业绩拆分求和，不去重",
     "求和 vs 订单表合同总金额：完全一致率 56.8%，中位差 0 元"),
    ("区域范围", "仅A区、B区、C区、D区",
     "E区/F区/G区/H区/律师条线等已解散或非区域条线归「其他」"),
    ("咨询费区域归属", "订单归属人 → 花名册区域",
     "「是否合作约见=是」的线索改用约见归属人1/2，按各自个数（通常 0.5+0.5）拆分；非合作案即使填了约见归属人也以订单归属人为准"),
    ("咨询费未命中", "10 条份额人名不在花名册且非合作约见", "已并入「其他」，不进四大区域"),
    ("咨询费下钻", "仅支持到区域级", "订单表「归属部门」与收入表「小组」是两套命名，无法对齐"),
    ("10 月", "业绩：收入表截至 2026-10-27，仅 81 行，未完整",
     "咨询费：已并入订单表_2026.10（10/2–10/29），10 月完整可比"),
]
r0 = 4
for i, (a, b, c) in enumerate(rows):
    rr = r0 + i
    for j, v in enumerate([a, b, c], 1):
        cell = ws.cell(row=rr, column=j, value=v)
        cell.font = Font(size=10, bold=(i == 0), color="FFFFFF" if i == 0 else "141A22")
        cell.alignment = Alignment(vertical="top", wrap_text=True)
        cell.border = BORDER
        if i == 0:
            cell.fill = PatternFill("solid", fgColor=BLUE)
        elif a.startswith("件数边界") or a in ("撞单处理", "区域范围", "咨询费区域归属"):
            cell.fill = PatternFill("solid", fgColor="FFF8EC")
    if i:
        n = max(1, -(-len(str(b)) // 40), -(-len(str(c)) // 62))
        ws.row_dimensions[rr].height = max(17, n * 15)
ws.column_dimensions["A"].width = 18
ws.column_dimensions["B"].width = 52
ws.column_dimensions["C"].width = 80

# ── 01 指标汇总 ──
ws = wb.create_sheet("01_指标汇总")
ws.sheet_view.showGridLines = False
k = P["kpi"]
title(ws, "核心指标汇总", "2026.1–10 ｜ 收入表为准")
summ = pd.DataFrame([
    ["总业绩", round(k["total_revenue"] / 10000, 1), "万元", "收入表全部行求和"],
    ["成案量", k["case_count"], "件", "按业绩比例拆分（V 列），每案恰 1 件"],
    ["成案金额", round(k["case_amount"] / 10000, 1), "万元", "新接记录金额求和"],
    ["客单价", k["arpu"], "元/件", "成案金额 ÷ 成案量"],
    ["咨询费数量", k["fee_count"], "条", "订单表线索去重，覆盖 1–10 月"],
    ["  其中：差旅费驱动", k["fee_travel"], "条", "可与其他条件重叠"],
    ["  其中：主场成案", k["fee_main"], "条", "合同额≥5000"],
    ["  其中：直播签约", k["fee_live"], "条", "合同额≥5000"],
    ["  人员可映射到小组", k.get("fee_map_rate"), "%", "订单归属人/约见归属人经花名册映射"],
    ["新接占比", k["new_share"], "%", "成案金额 ÷ 总业绩"],
], columns=["指标", "数值", "单位", "说明"])
write_table(ws, summ, 4, widths=[20, 16, 10, 50], num_fmt="#,##0.0")

# ── 02 月度 ──
ws = wb.create_sheet("02_月度")
ws.sheet_view.showGridLines = False
title(ws, "月度指标", "件数为拆分后份额，可出现小数；10 月业绩未完整")
mon = pd.DataFrame([{
    "月份": f'{r["mon"]}月', "总业绩(万元)": round(r["revenue"] / 10000, 1),
    "成案量(件)": r["case_count"], "成案金额(万元)": round(r["case_amount"] / 10000, 1),
    "客单价(元)": round(r["arpu"]) if r["arpu"] else None,
    "咨询费(条)": r["fee_count"],
    "业绩状态": "完整" if r["complete"] else "未完整",
} for r in P["month"]])
write_table(ws, mon, 4, widths=[10, 15, 12, 15, 13, 12, 10])

# ── 03 区域 ──
ws = wb.create_sheet("03_区域")
ws.sheet_view.showGridLines = False
title(ws, "区域指标", "A区/B区/C区/D区 + 其他（已解散，含未归属）")
reg = pd.DataFrame([{
    "区域": ("其他（已解散）" if r["k"] == "其他" else r["k"]),
    "总业绩(万元)": round(r["revenue"] / 10000, 1),
    "成案量(件)": r["case_count"],
    "成案金额(万元)": round(r["case_amount"] / 10000, 1),
    "客单价(元)": round(r["arpu"]) if r["arpu"] else None,
    "新接占比(%)": r["new_share"],
    "咨询费(条)": r["fee_count"],
} for r in P["region"]])
write_table(ws, reg, 4, widths=[16, 14, 12, 15, 13, 13, 12], num_fmt="#,##0.0")

# ── 04 小组 ──
ws = wb.create_sheet("04_小组")
ws.sheet_view.showGridLines = False
title(ws, "小组指标", "咨询费已按人员映射到小组（与收入表同一命名）；4 个小组有咨询费无业绩")
tm = pd.DataFrame([{
    "小组": r["k"],
    "总业绩(万元)": round(r["revenue"] / 10000, 1),
    "成案量(件)": r["case_count"],
    "成案金额(万元)": round(r["case_amount"] / 10000, 1),
    "客单价(元)": round(r["arpu"]) if r["arpu"] else None,
    "咨询费(条)": r.get("fee_count"),
} for r in P["team"]])
tm["区域"] = tm["小组"].map(lambda x: next((r["k"] for r in P["region"] if r["k"] != "其他" and x.startswith(r["k"])), "其他（已解散）"))
tm = tm[["小组", "区域", "总业绩(万元)", "成案量(件)", "成案金额(万元)", "客单价(元)", "咨询费(条)"]]
tm = tm.sort_values("总业绩(万元)", ascending=False)
write_table(ws, tm, 4, widths=[24, 16, 14, 12, 15, 13, 12], num_fmt="#,##0.0")

# ── 05 区域月份矩阵 ──
ws = wb.create_sheet("05_区域月份矩阵")
ws.sheet_view.showGridLines = False
title(ws, "区域 × 月份 总业绩矩阵（万元）", "可直接透视")
months = [r["mon"] for r in P["month"]]
piv = {}
for r in P["reg_mon"]:
    piv.setdefault(r["region"], {})[r["mon"]] = r["revenue"]
mat = pd.DataFrame([{
    "区域": ("其他（已解散）" if rg == "其他" else rg),
    **{f"{m}月": round(piv.get(rg, {}).get(m, 0) / 10000, 1) for m in months},
    "合计": round(sum(piv.get(rg, {}).values()) / 10000, 1),
} for rg in [x["k"] for x in P["region"]]])
write_table(ws, mat, 4, widths=[16] + [10] * len(months) + [12], num_fmt="#,##0.0")

# ── 06 区域月份件数 ──
ws = wb.create_sheet("06_区域月份件数")
ws.sheet_view.showGridLines = False
title(ws, "区域 × 月份 成案量矩阵（件）", "按业绩比例拆分，合计 = 1,115")
pc = {}
for r in P["reg_mon"]:
    pc.setdefault(r["region"], {})[r["mon"]] = r["case_count"]
matc = pd.DataFrame([{
    "区域": ("其他（已解散）" if rg == "其他" else rg),
    **{f"{m}月": pc.get(rg, {}).get(m, 0) for m in months},
    "合计": round(sum(pc.get(rg, {}).values()), 2),
} for rg in [x["k"] for x in P["region"]]])
write_table(ws, matc, 4, widths=[16] + [10] * len(months) + [12], num_fmt="#,##0.0#")

# ── 07 咨询费月度 ──
ws = wb.create_sheet("07_咨询费月度")
ws.sheet_view.showGridLines = False
title(ws, "咨询费月度与构成（条）", "同一线索可命中多个规则，三列之和 ≥ 总量；含 10 月")
ft = {x["tag"]: x for x in P["fee_tag"]}
fm = pd.DataFrame([{
    "月份": f'{m}月',
    "咨询费(条)": ft["总线索"]["mon"].get(str(m), 0),
    "差旅费驱动": ft["差旅费"]["mon"].get(str(m), 0),
    "主场成案": ft["主场成案"]["mon"].get(str(m), 0),
    "直播签约": ft["直播签约"]["mon"].get(str(m), 0),
} for m in P["meta"]["fee_months"]])
write_table(ws, fm, 4, widths=[10, 13, 14, 12, 12], num_fmt="#,##0.0#")

# ── 08 咨询费区域月份 ──
ws = wb.create_sheet("08_咨询费区域月份")
ws.sheet_view.showGridLines = False
title(ws, "咨询费 区域 × 月份（条）", "合作约见按约见归属人 0.5/0.5 拆分，故有小数")
pv2 = {}
for o in P["fee_rm"]:
    pv2.setdefault(o["region"], {})[o["mon"]] = o["fee_count"]
fmt2 = pd.DataFrame([{
    "区域": ("其他（已解散/含未归属）" if rg == "其他" else rg),
    **{f"{m}月": pv2.get(rg, {}).get(m, None) for m in P["meta"]["fee_months"]},
    "合计": round(sum(v for v in pv2.get(rg, {}).values()), 2),
} for rg in [x["k"] for x in P["region"]]])
write_table(ws, fmt2, 4, widths=[24] + [9] * len(P["meta"]["fee_months"]) + [10], num_fmt="#,##0.0#")

# ── 09 新接案件明细 ──
ws = wb.create_sheet("09_新接案件明细")
ws.sheet_view.showGridLines = False
title(ws, "新接案件明细（含 V/W/X 计算列）",
      f'案件 {k["case_count"]:,} 件 / 记录 {k["case_rows"]:,} 行；W=案件新接合计，X=记录数，件数份额=本行金额÷W')
nc = pd.read_csv(os.path.join(BUILD, "new_cases.csv"))
nc = nc[["case_id", "mon", "region", "team", "company", "tag", "amt", "W", "X", "v"]].copy()
nc.columns = ["案件编号", "月份", "区域", "小组", "所属公司", "款项标签",
              "业绩金额(元)", "W 案件新接合计", "X 记录数", "件数份额"]
nc["区域"] = nc["区域"].replace({"其他": "其他（已解散）"})
nc["业绩金额(元)"] = nc["业绩金额(元)"].round(2)
nc["W 案件新接合计"] = nc["W 案件新接合计"].round(2)
nc = nc.sort_values(["月份", "案件编号"])
write_table(ws, nc, 4, widths=[16, 8, 16, 24, 20, 18, 15, 16, 9, 10], num_fmt="#,##0.00")

# ── 10 咨询费线索明细 ──
ws = wb.create_sheet("10_咨询费线索明细")
ws.sheet_view.showGridLines = False
title(ws, "咨询费线索明细（订单表 1–10 月）",
      f'{k["fee_count"]:,} 条线索；合作约见拆成两行（各 0.5）；小组由人员经花名册映射，区域由小组推导')
cf = pd.read_csv(os.path.join(BUILD, "consult.csv"))
cf = cf[["mon", "lead", "region", "team", "owner", "w", "f_travel", "f_main", "f_live", "lv", "ht"]].copy()
cf.columns = ["月份", "线索ID", "区域", "小组", "订单归属人", "条数份额",
              "差旅费驱动", "主场成案", "直播签约", "差旅费(元)", "合同额(元)"]
cf["区域"] = cf["区域"].replace({"其他": "其他（已解散/含未归属）"})
for c in ["差旅费驱动", "主场成案", "直播签约"]:
    cf[c] = cf[c].map(lambda x: "是" if x else "")
cf["差旅费(元)"] = cf["差旅费(元)"].round(2)
cf["合同额(元)"] = cf["合同额(元)"].round(2)
cf = cf.sort_values(["月份", "区域"])
write_table(ws, cf, 4, widths=[8, 20, 24, 24, 14, 10, 11, 11, 11, 13, 14], num_fmt="#,##0.00")

# ── 11 业绩明细 ──
ws = wb.create_sheet("11_业绩明细")
ws.sheet_view.showGridLines = False
title(ws, "业绩明细（全部款项类型，未去重）",
      "账面合计 %.1f 万元" % (k["total_revenue"] / 10000))
ra = pd.read_csv(os.path.join(BUILD, "revenue_raw.csv"))
ra = ra[["mon", "case_id", "kind", "tag", "region", "team", "company", "amt"]].copy()
ra.columns = ["月份", "案件编号", "款项类型", "原始标签", "区域", "小组", "所属公司", "业绩金额(元)"]
ra["区域"] = ra["区域"].replace({"其他": "其他（已解散）"})
ra["业绩金额(元)"] = ra["业绩金额(元)"].round(2)
ra = ra.sort_values(["月份", "案件编号"])
write_table(ws, ra, 4, widths=[8, 16, 12, 18, 16, 24, 20, 15], num_fmt="#,##0.00")

# ── 12 咨询费小组月份 ──
ws = wb.create_sheet("12_咨询费小组月份")
ws.sheet_view.showGridLines = False
title(ws, "咨询费 · 小组 × 月份矩阵", "单元格为咨询费条数（合作案件按份额拆到两个小组）；仅列有咨询费的小组")
gm = cf.pivot_table(index=["区域", "小组"], columns="月份", values="条数份额",
                    aggfunc="sum", fill_value=0)
gm = gm.reindex(columns=sorted(gm.columns))
gm["合计"] = gm.sum(axis=1)
gm = gm.reset_index()
gm["合计"] = gm["合计"].round(1)
gm.columns = ["区域", "小组"] + [f"{c}月" if isinstance(c, (int,)) else str(c) for c in gm.columns[2:]]
gm = gm.sort_values("合计", ascending=False)
write_table(ws, gm, 4, num_fmt="#,##0.0")

# ── 13 区域月份客单价（矩阵） ──
from openpyxl.formatting.rule import ColorScaleRule
REGION_ORDER = [r["k"] for r in P["region"]]
MONTHS = sorted({o["mon"] for o in P["fact_new"]})
cell = {}
for o in P["fact_new"]:   # 用案件级事实重算，保证与 03_区域 完全一致
    k = (o["region"], o["mon"])
    a, c = cell.get(k, (0.0, 0.0))
    cell[k] = (a + float(o["amt"]), c + float(o["v"]))

ws = wb.create_sheet("13_区域月份客单价")
ws.sheet_view.showGridLines = False
title(ws, "区域 × 月份 · 客单价矩阵（元/件）",
      "客单价 = 当月该区域成案金额 ÷ 当月该区域成案量（件数）；「—」表示当月无成案")
mtx = []
tot_amt = {m: 0.0 for m in MONTHS}
tot_cnt = {m: 0.0 for m in MONTHS}
for reg in REGION_ORDER:
    row = {"区域": reg if reg != "其他" else "其他（已解散）"}
    for m in MONTHS:
        amt, cnt = cell.get((reg, m), (0.0, 0.0))
        tot_amt[m] += amt
        tot_cnt[m] += cnt
        row[f"{m}月"] = round(amt / cnt) if cnt > 0 else None
    a = sum(cell.get((reg, m), (0.0, 0.0))[0] for m in MONTHS)
    c = sum(cell.get((reg, m), (0.0, 0.0))[1] for m in MONTHS)
    row["合计"] = round(a / c) if c > 0 else None
    mtx.append(row)
tot_row = {"区域": "合计"}
for m in MONTHS:
    tot_row[f"{m}月"] = round(tot_amt[m] / tot_cnt[m]) if tot_cnt[m] > 0 else None
tot_row["合计"] = round(sum(tot_amt.values()) / sum(tot_cnt.values()))
n_reg = len(mtx)
mtx = pd.DataFrame(mtx + [tot_row])
end = write_table(ws, mtx, 4, widths=[18] + [11] * len(MONTHS) + [12], num_fmt="#,##0")
# 合计行加粗
for j in range(1, len(mtx.columns) + 1):
    c = ws.cell(row=4 + len(mtx), column=j)
    c.font = Font(size=10, bold=True, color=NAVY)
    c.fill = PatternFill("solid", fgColor="EFF3FB")

# 数据色阶（仅区域行，不含合计）
first_col = 2
last_col = 1 + len(MONTHS) + 1
rng = f"{get_column_letter(first_col)}5:{get_column_letter(last_col)}{4 + n_reg}"
ws.conditional_formatting.add(rng, ColorScaleRule(
    start_type="min", start_color="FFFFFF",
    mid_type="percentile", mid_value=50, mid_color="EDF2FC",
    end_type="max", end_color="C2D3F0"))
ws.freeze_panes = "B5"

nr = 4 + len(mtx) + 2
for t in ["读法：横向看同一区域逐月客单价变化，纵向看同月各区差异；",
          "「其他（已解散）」含律师条线、E区/F区/G区/H区，客单价参考意义有限。",
          "客单价按月独立计算，与「02_月度」的B区口径一致（成案金额 ÷ 拆分件数）。"]:
    ws.cell(row=nr, column=1, value=t).font = Font(size=9, color="8C96A3")
    nr += 1

# ── 14 区域月份客单价明细（长表，可透视/看环比） ──
ws = wb.create_sheet("14_区域月份客单价明细")
ws.sheet_view.showGridLines = False
title(ws, "区域 × 月份 · 客单价明细（长表）",
      "共 5 区域 × 10 个月；环比 = 本月客单价 ÷ 上月客单价 − 1，上月无成案则留空")
rows = []
for reg in REGION_ORDER:
    prev = None
    for m in MONTHS:
        amt, cnt = cell.get((reg, m), (0.0, 0.0))
        kd = amt / cnt if cnt > 0 else None
        rows.append({
            "区域": reg if reg != "其他" else "其他（已解散）",
            "月份": f"{m}月",
            "成案量(件)": round(cnt, 2),
            "成案金额(元)": round(amt, 2),
            "客单价(元)": round(kd) if kd is not None else None,
            "环比(%)": (round(kd / prev - 1, 4) if (kd is not None and prev) else None),
        })
        prev = kd if kd is not None else prev
ld = pd.DataFrame(rows)
write_table(ws, ld, 4, widths=[18, 8, 12, 16, 14, 11], num_fmt="#,##0.00")
# 逐列设置格式
_fmt = {"成案量(件)": "#,##0.00", "成案金额(元)": "#,##0.00",
        "客单价(元)": "#,##0", "环比(%)": "0.0%"}
for j, col in enumerate(ld.columns, 1):
    if col in _fmt:
        for i in range(5, 5 + len(ld)):
            ws.cell(row=i, column=j).number_format = _fmt[col]


# 打印设置：宽度一页、高度自动，明细页重复表头
from openpyxl.worksheet.properties import PageSetupProperties
for sh in wb.worksheets:
    sh.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    sh.page_setup.fitToWidth = 1
    sh.page_setup.fitToHeight = 0
    sh.page_setup.orientation = "portrait"
    sh.page_margins.left = sh.page_margins.right = 0.5
    sh.page_margins.top = sh.page_margins.bottom = 0.6
    if sh.max_row > 60:
        sh.print_title_rows = "4:4"

wb.save(OUTF)
print("saved", OUTF, f"{os.path.getsize(OUTF)/1024:.0f} KB")
print("sheets:", [s.title for s in wb.worksheets])
