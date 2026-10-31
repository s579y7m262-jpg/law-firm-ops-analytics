# -*- coding: utf-8 -*-
"""单区域客单价分档分析：成案金额按区间分箱，看「量的结构」而不是只看总量。"""
import os
import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
DATA = os.path.join(ROOT, "data", "raw")
BUILD = os.path.join(ROOT, "build")
os.makedirs(BUILD, exist_ok=True)
REGION = os.environ.get("REGION", "B区")   # 目标区域，需与 build_gold 的区域代号一致
OUTF = os.path.join(BUILD, f"区域经营分析_{REGION}.xlsx")

NAVY = "1F3864"
BLUE = "2F5BEA"
GREEN = "0E9F6E"
thin = Side(style="thin", color="E6E8EC")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

BINS = [0, 50000, 100000, 150000, 200000, 300000, 10 ** 12]
LABELS = ["5万以下", "5~10万", "10~15万", "15~20万", "20~30万", "30万以上"]


def title(ws, text, sub=""):
    ws["A1"] = text
    ws["A1"].font = Font(size=14, bold=True, color=NAVY)
    ws.row_dimensions[1].height = 26
    if sub:
        ws["A2"] = sub
        ws["A2"].font = Font(size=9.5, color="5B6572")


def write_table(ws, df, start_row, widths=None, num_fmt=None, fmts=None, total_row=None):
    cols = list(df.columns)
    for j, c in enumerate(cols, 1):
        cell = ws.cell(row=start_row, column=j, value=c)
        cell.font = Font(bold=True, color="FFFFFF", size=10)
        cell.fill = PatternFill("solid", fgColor=BLUE)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
    ws.row_dimensions[start_row].height = 30
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
                f = (fmts or {}).get(cols[j - 1], num_fmt)
                if f and isinstance(v, (int, float)):
                    cell.number_format = f
            else:
                cell.alignment = Alignment(horizontal="left")
        if (i - start_row) % 2 == 0:
            for j in range(1, len(cols) + 1):
                ws.cell(row=i, column=j).fill = PatternFill("solid", fgColor="FAFBFD")
    if total_row is not None:
        r = start_row + len(df) + 1
        for j, v in enumerate(total_row, 1):
            cell = ws.cell(row=r, column=j, value=v)
            cell.font = Font(size=10, bold=True, color=NAVY)
            cell.fill = PatternFill("solid", fgColor="EFF3FB")
            cell.border = Border(left=thin, right=thin,
                                 top=Side(style="medium", color="C9D3E5"), bottom=thin)
            if j > 1:
                cell.alignment = Alignment(horizontal="right")
                f = (fmts or {}).get(cols[j - 1], num_fmt)
                if f and isinstance(v, (int, float)):
                    cell.number_format = f
    if widths:
        for j, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = ws.cell(row=start_row + 1, column=1)
    return start_row + len(df) + (1 if total_row is not None else 0)


# ══════════════════════════════════════════════
# 数据
# ══════════════════════════════════════════════
n = pd.read_csv(os.path.join(BUILD, "new_cases.csv"))
allc = n.groupby("case_id").agg(tot=("amt", "sum"), regs=("region", "nunique"))
cd = n[n["region"] == REGION].copy()

g = cd.groupby("case_id").agg(amt=("amt", "sum"), v=("v", "sum"), mon=("mon", "min"), rows=("amt", "size"))
# 客单价 = 该案在B区的金额 ÷ 该案在B区的件数份额（跨区域案件由此还原为整案单价）
g["kd"] = g["amt"] / g["v"]
main_team = cd.sort_values("amt", ascending=False).drop_duplicates("case_id").set_index("case_id")
g["team"] = main_team["team"]
g["company"] = main_team["company"]
g["tag"] = main_team["tag"]
g["mixed"] = g.index.map(allc["regs"]) > 1
g["total_amt"] = g.index.map(allc["tot"])
g["seg"] = pd.cut(g["kd"], bins=BINS, labels=LABELS, right=False)

segs = g.groupby("seg", observed=False).agg(cnt=("v", "sum"), cases=("amt", "size"), amt=("amt", "sum"))
segs["share"] = segs["cnt"] / segs["cnt"].sum() * 100
segs["avg"] = segs["amt"] / segs["cnt"]
segs["amt_share"] = segs["amt"] / segs["amt"].sum() * 100
segs["cum"] = segs["share"].cumsum()

fee = pd.read_csv(os.path.join(BUILD, "consult.csv"))
fee_cd = fee[fee["region"] == REGION]

wb = Workbook()

# ── 00 说明 ──
ws = wb.active
ws.title = "00_口径说明"
ws.sheet_view.showGridLines = False
title(ws, "B区域经营分析 · 口径说明",
      "窗口 2026.1–10 ｜ 数据源 收入表_2026.1-2026.10 ｜ 生成 " + pd.Timestamp.now().strftime("%Y-%m-%d %H:%M"))
rows = [
    ("项目", "口径", "说明"),
    ("区域范围", "收入表「所属小组」以「B区」开头", "B区行政一至八组、B区谈案部；不含其他区域"),
    ("成案量（件）", "新接记录按业绩比例拆分的件数（V 列），口径与看板完全一致",
     "每案恰 1 件；跨区域案件按业绩比例在各区域分摊，B区合计 196.9 件"),
    ("成案金额", "该案件在B区的新接业绩金额合计", "跨区域拆分的案件只计B区认款部分"),
    ("客单价", "成案金额 ÷ 成案量（件）",
     "单案客单价 = 该案在B区的新接案件金额 ÷ 该案的件数份额（份额为 1 时即该案金额）；"
     "区间平均客单价 = 区间金额 ÷ 区间件数"),
    ("客单价区间", "左闭右开：<5万 / 5–10万 / 10–15万 / 15–20万 / 20–30万 / ≥30万",
     "以案件在B区的成案金额分档，非合同额"),
    ("数量占比", "区间笔数 ÷ B区总笔数", "B区合计 = 100%"),
    ("总业绩金额", "区间内各案件成案金额之和", "仅新接（成案）金额，不含中期尾款/续约等后续款项"),
    ("", "", ""),
    ("与看板的一致性", "本表成案量 196.9 件 = 看板B区成案量",
     "两端同源：均为「新接记录按业绩比例拆分的件数」；案件个数另列 203 个备查"),
    ("跨区域案件", "11 个案件件数份额 <1，B区金额 75.8 万元（占 2.7%）",
     "已按 金额 ÷ 件数份额 还原整案客单价再分档，区间口径不受跨区域折损影响"),
    ("10 月数据", "收入表截至 2026-10-27，10 月仅 5 笔", "月度对比时 10 月不完整，勿直接与整月环比"),
    ("咨询费转化率", "当月成案笔数 ÷ 当月咨询费条数", "分子分母非同批线索，可能 >100%，仅作趋势参考"),
]
for i, (a, b, c) in enumerate(rows, start=4):
    ws.cell(row=i, column=1, value=a).font = Font(size=10, bold=(i == 4), color="FFFFFF" if i == 4 else "141A22")
    ws.cell(row=i, column=2, value=b).font = Font(size=10, bold=(i == 4), color="FFFFFF" if i == 4 else "141A22")
    ws.cell(row=i, column=3, value=c).font = Font(size=9.5, color="FFFFFF" if i == 4 else "5B6572")
    for j in range(1, 4):
        cell = ws.cell(row=i, column=j)
        cell.border = BORDER
        cell.alignment = Alignment(vertical="top", wrap_text=True)
        if i == 4:
            cell.fill = PatternFill("solid", fgColor=BLUE)
        elif i % 2 == 0:
            cell.fill = PatternFill("solid", fgColor="FAFBFD")
    ws.row_dimensions[i].height = 8 if (a == "" and b == "") else 30
ws.column_dimensions["A"].width = 20
ws.column_dimensions["B"].width = 52
ws.column_dimensions["C"].width = 62

# ── 01 总览 ──
ws = wb.create_sheet("01_总览")
ws.sheet_view.showGridLines = False
title(ws, "B区域 · 经营总览（2026.1–10）", "成案口径（新接）｜ 成案量按业绩比例拆分为件数，与看板B区口径一致")
cards = [
    ("成案量", round(float(segs["cnt"].sum()), 1), "件"),
    ("成案金额", round(float(segs["amt"].sum()) / 10000, 1), "万元"),
    ("平均客单价", round(float(segs["amt"].sum()) / float(segs["cnt"].sum())), "元/件"),
    ("咨询费数量", round(float(fee_cd["w"].sum()), 1), "条"),
    ("占全域成案金额", round(float(segs["amt"].sum()) / float(n["amt"].sum()) * 100, 1), "%"),
]
col = 1
for name, val, unit in cards:
    c = ws.cell(row=4, column=col, value=name)
    c.font = Font(size=9.5, color="5B6572")
    c.fill = PatternFill("solid", fgColor="F4F6FA")
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    v = ws.cell(row=5, column=col, value=val)
    v.font = Font(size=20, bold=True, color=NAVY)
    v.fill = PatternFill("solid", fgColor="F4F6FA")
    v.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    v.number_format = "#,##0.0" if isinstance(val, float) else "#,##0"
    u = ws.cell(row=5, column=col + 0, value=val)
    ws.cell(row=6, column=col, value=unit).font = Font(size=9, color="8C96A3")
    ws.cell(row=6, column=col).fill = PatternFill("solid", fgColor="F4F6FA")
    ws.cell(row=6, column=col).alignment = Alignment(horizontal="left", indent=1)
    for r in (4, 5, 6):
        for cc in range(col, col + 1):
            ws.cell(row=r, column=cc).border = Border(
                left=Side(style="thin", color="E6E8EC"), right=Side(style="thin", color="E6E8EC"),
                top=Side(style="thin", color="E6E8EC") if r == 4 else Side(style="hair", color="E6E8EC"),
                bottom=Side(style="thin", color="E6E8EC") if r == 6 else Side(style="hair", color="E6E8EC"))
    ws.column_dimensions[get_column_letter(col)].width = 18
    col += 1

ws.cell(row=9, column=1, value="关键结论").font = Font(size=11, bold=True, color=NAVY)
top = segs.sort_values("amt", ascending=False)
small = segs.loc["5万以下"]
big = segs.loc["30万以上"]
mid = segs.loc[["5~10万", "10~15万"]]
lines = [
    f"1. 客单价呈「两头尖、中间厚」：5–15 万区间 {mid['cnt'].sum():.0f} 件，占 {mid['cnt'].sum()/segs['cnt'].sum()*100:.1f}%，贡献 {mid['amt'].sum()/segs['amt'].sum()*100:.1f}% 金额，是B区的主力成交带。",
    f"2. 小额单量大但贡献低：5 万以下 {small['cnt']:.0f} 件占 {small['share']:.1f}%，金额只占 {small['amt_share']:.1f}%，平均客单价 {small['avg']:,.0f} 元 —— 数量看着不少，收入拉动有限。",
    f"3. 大单决定业绩：30 万以上仅 {big['cnt']:.0f} 件（{big['share']:.1f}%），却贡献 {big['amt_share']:.1f}% 金额，平均 {big['avg']:,.0f} 元；单件价值是 5 万以下的 {big['avg']/small['avg']:.1f} 倍。",
    f"4. 中间层最值钱：20–30 万区间 {segs.loc['20~30万','cnt']:.0f} 件即贡献 {segs.loc['20~30万','amt_share']:.1f}% 金额，单价高、成交难度低于 30 万以上，是可以优先放量的区间。",
    f"5. 结构提示：前 3 档（≤15 万）合计 {segs.loc[LABELS[:3],'cnt'].sum():.0f} 件占 {segs.loc[LABELS[:3],'share'].sum():.1f}%，金额占 {segs.loc[LABELS[:3],'amt_share'].sum():.1f}% —— 若考核只看件数，容易掩盖客单价偏低的问题。",
]
for i, t in enumerate(lines, start=10):
    c = ws.cell(row=i, column=1, value=t)
    c.font = Font(size=10, color="25303F")
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=5)
    ws.row_dimensions[i].height = 40

# ── 02 客单价区间（主表） ──
ws = wb.create_sheet("02_客单价区间")
ws.sheet_view.showGridLines = False
title(ws, "B区 · 客单价区间分布（2026.1–10）",
      "区间按单案客单价划分（B区金额 ÷ 件数份额，跨区域案件已还原整案单价），左闭右开；占比合计 100%")
tbl = pd.DataFrame({
    "客单价区间": LABELS,
    "成案量(件)": segs["cnt"].round(1).values,
    "数量占比": (segs["share"] / 100).round(4).values,
    "平均客单价(元)": segs["avg"].round(0).values,
    "总业绩金额(元)": segs["amt"].round(2).values,
    "金额占比": (segs["amt_share"] / 100).round(4).values,
    "案件数(个)": segs["cases"].astype(int).values,
    "累计占比": (segs["cum"] / 100).round(4).values,
})
end = write_table(ws, tbl,
                  start_row=4,
                  widths=[14, 12, 11, 15, 17, 11, 11, 11],
                  fmts={"成案量(件)": "#,##0.0", "数量占比": "0.0%", "平均客单价(元)": "#,##0",
                        "总业绩金额(元)": "#,##0.00", "金额占比": "0.0%", "案件数(个)": "#,##0", "累计占比": "0.0%"},
                  total_row=["合计", round(float(tbl["成案量(件)"].sum()), 1), 1.0,
                             round(float(segs["amt"].sum() / segs["cnt"].sum()), 0),
                             round(float(segs["amt"].sum()), 2), 1.0,
                             int(tbl["案件数(个)"].sum()), None])

ch = BarChart()
ch.type = "col"
ch.style = 10
ch.title = "数量占比 vs 金额占比"
ch.y_axis.numFmt = "0%"
ch.height = 9.5
ch.width = 19
data = Reference(ws, min_col=3, max_col=3, min_row=4, max_row=4 + len(tbl))
data2 = Reference(ws, min_col=6, max_col=6, min_row=4, max_row=4 + len(tbl))
cats = Reference(ws, min_col=1, min_row=5, max_row=4 + len(tbl))
ch.add_data(data, titles_from_data=True)
ch.add_data(data2, titles_from_data=True)
ch.set_categories(cats)
ch.series[0].graphicalProperties.solidFill = "2F5BEA"
ch.series[1].graphicalProperties.solidFill = "E8952A"
ch.legend.position = "r"
ch.x_axis.delete = False
ch.y_axis.delete = False
ch.x_axis.majorTickMark = "out"
ch.gapWidth = 55
ws.add_chart(ch, "A15")

note_r = end + 2
ws.cell(row=note_r, column=1, value="注：平均客单价 = 该区间总业绩金额 ÷ 该区间成案量（件）；案件数仅供核对案件个数，跨区域案件件数份额 <1 故与件数不等。").font = Font(size=9, color="8C96A3")

# ── 03 区间×月份（笔数） ──
ws = wb.create_sheet("03_区间x月份_件数")
ws.sheet_view.showGridLines = False
title(ws, "B区 · 客单价区间 × 月份（件数）", "单元格为拆分后的件数（可为小数）；右侧为区间合计与占比")
mv = g.pivot_table(index="seg", columns="mon", values="v", aggfunc="sum", fill_value=0, observed=False)
mv = mv.reindex(LABELS)
mamt = g.pivot_table(index="seg", columns="mon", values="amt", aggfunc="sum", fill_value=0, observed=False).reindex(LABELS)
mv.columns = [f"{c}月" for c in mv.columns]
mv = mv.round(2)
mv["合计"] = mv.sum(axis=1).round(1)
rows = mv.reset_index().rename(columns={"seg": "客单价区间"})
rows["合计占比"] = (rows["合计"] / rows["合计"].sum()).round(4)
write_table(ws, rows, 4, widths=[14] + [7] * (len(mv.columns) - 1) + [9, 11],
            fmts={c: "#,##0.0" for c in rows.columns if c.endswith("月")} | {"合计": "#,##0.0", "合计占比": "0.0%"},
            total_row=["合计"] + [round(float(v), 1) for v in mv.sum(axis=0).tolist()] + [1.0])
ws.row_dimensions[5 + len(rows)].height = 20

# ── 04 区间×月份（占比） ──
ws = wb.create_sheet("04_区间x月份_占比")
ws.sheet_view.showGridLines = False
title(ws, "B区 · 客单价区间 × 月份（数量占比）", "按月归一：每一列合计 100%，可直接看当月成交结构变化")
pct = mv.drop(columns=["合计"]).div(mv.drop(columns=["合计"]).sum(axis=0), axis=1)
rows = pct.reset_index().rename(columns={"seg": "客单价区间"})
for c in rows.columns[1:]:
    rows[c] = rows[c].round(4)
write_table(ws, rows, 4, widths=[14] + [8] * (len(rows.columns) - 1),
            fmts={c: "0.0%" for c in rows.columns[1:]},
            total_row=["合计"] + [1.0] * (len(rows.columns) - 1))

# ── 05 区间×小组 ──
ws = wb.create_sheet("05_区间x小组")
ws.sheet_view.showGridLines = False
title(ws, "B区 · 客单价区间 × 小组（件数）", "小组取该案件金额最大的主责组；件数按业绩比例分摊，可为小数")
tv = g.pivot_table(index="team", columns="seg", values="v", aggfunc="sum", fill_value=0, observed=False)
tv = tv.reindex(columns=LABELS, fill_value=0).round(2)
tv["合计"] = tv.sum(axis=1).round(1)
tv = tv.sort_values("合计", ascending=False)
tamt = g.groupby("team")["amt"].sum()
tv["总业绩(万元)"] = (tamt.reindex(tv.index) / 10000).round(1)
tv["平均客单价(元)"] = (tamt.reindex(tv.index) / tv["合计"]).round(0)
rows = tv.reset_index().rename(columns={"team": "小组"})
rows.columns = ["小组"] + LABELS + ["合计", "总业绩(万元)", "平均客单价(元)"]
write_table(ws, rows, 4, widths=[16] + [9] * 6 + [8, 13, 13],
            fmts={**{c: "#,##0.0" for c in LABELS}, "合计": "#,##0.0",
                  "总业绩(万元)": "#,##0.0", "平均客单价(元)": "#,##0"},
            total_row=["合计"] + [round(float(v), 1) for v in tv[LABELS].sum().tolist()]
            + [round(float(g["v"].sum()), 1), round(float(tamt.sum() / 10000), 1),
               round(float(tamt.sum() / g["v"].sum()), 0)])

# ── 06 案件明细 ──
ws = wb.create_sheet("06_案件明细")
ws.sheet_view.showGridLines = False
title(ws, f"B区 · 成案案件明细（{len(g)} 个案件 / {g['v'].sum():.1f} 件）",
      "「跨区域拆分」为是时金额仅为B区认款部分、件数份额 <1；客单价 = B区金额 ÷ 件数份额")
det = pd.DataFrame({
    "案件编号": g.index,
    "月份": g["mon"].astype(int).values,
    "小组": g["team"].values,
    "所属公司": g["company"].values,
    "客单价区间": g["seg"].astype(str).values,
    "成案金额(元)": g["amt"].round(2).values,
    "件数份额": g["v"].round(2).values,
    "客单价(元)": g["kd"].round(0).values,
    "记录数": g["rows"].astype(int).values,
    "跨区域拆分": np.where(g["mixed"], "是", ""),
    "案件全线金额(元)": g["total_amt"].round(2).values,
}).sort_values(["月份", "成案金额(元)"], ascending=[True, False])
write_table(ws, det, 4, widths=[16, 8, 16, 20, 12, 15, 10, 13, 8, 11, 17],
            fmts={"成案金额(元)": "#,##0.00", "件数份额": "0.00", "客单价(元)": "#,##0",
                  "案件全线金额(元)": "#,##0.00"})

# ── 07 月度趋势 ──
ws = wb.create_sheet("07_月度趋势")
ws.sheet_view.showGridLines = False
title(ws, "B区 · 月度走势", "10 月收入表截至 10/27，不完整，勿直接环比")
mo = g.groupby("mon").agg(件数=("v", "sum"), 金额=("amt", "sum")).reset_index()
mo["平均客单价(元)"] = (mo["金额"] / mo["件数"]).round(0)
mo["金额(万元)"] = (mo["金额"] / 10000).round(1)
feem = fee_cd.groupby("mon")["w"].sum().round(1)
mo["咨询费(条)"] = mo["mon"].map(feem).fillna(0)
mo["咨询费转化率"] = (mo["件数"] / mo["咨询费(条)"]).round(3)
mo = mo.rename(columns={"mon": "月份"})
mo["月份"] = mo["月份"].map(lambda m: f"{int(m)}月")
mo = mo[["月份", "件数", "金额(万元)", "平均客单价(元)", "咨询费(条)", "咨询费转化率"]]
write_table(ws, mo, 4, widths=[8, 10, 14, 15, 13, 15],
            fmts={"件数": "#,##0.0", "金额(万元)": "#,##0.0", "平均客单价(元)": "#,##0",
                  "咨询费(条)": "#,##0.0", "咨询费转化率": "0.0%"},
            total_row=["合计", round(float(mo["件数"].sum()), 1), round(float(segs["amt"].sum()) / 10000, 1),
                       round(float(segs["amt"].sum() / segs["cnt"].sum()), 0),
                       round(float(mo["咨询费(条)"].sum()), 1), None])
nr = 4 + len(mo) + 3
for t in ["注 1：咨询费转化率 = 当月成案量（件） ÷ 当月咨询费条数。分子分母不是同一批线索（当月成交可能来自前期咨询），",
          "      故该值可能大于 100%，只看趋势与量级，不可当作「线索转化率」对外引用。",
          "注 2：10 月收入表截至 10/27，仅 5 笔，转化率与客单价均不具可比性，勿直接环比。"]:
    ws.cell(row=nr, column=1, value=t).font = Font(size=9, color="8C96A3")
    nr += 1

# ── 08 四区域对比 ──
ws = wb.create_sheet("08_四区域对比")
ws.sheet_view.showGridLines = False
title(ws, "四区域 · 客单价结构对比（数量占比）", "同一分档口径，看B区在全域中的位置")
comp = []
for reg, sub in n.groupby("region"):
    gg = sub.groupby("case_id").agg(amt=("amt", "sum"), v=("v", "sum"))
    gg["kd"] = gg["amt"] / gg["v"].replace(0, pd.NA)
    s = pd.cut(gg["kd"], bins=BINS, labels=LABELS, right=False)
    sh = gg.groupby(s, observed=False)["v"].sum().reindex(LABELS).fillna(0)
    comp.append(pd.concat([pd.Series({"区域": reg, "件数": round(float(gg["v"].sum()), 1),
                                      "平均客单价(元)": round(float(gg["amt"].sum() / gg["v"].sum()), 0)}),
                           (sh / sh.sum()).round(4)]))
cmp_df = pd.DataFrame(comp)
cmp_df["合计"] = cmp_df[LABELS].sum(axis=1).round(4)
cmp_df = cmp_df[["区域", "件数", "平均客单价(元)"] + LABELS + ["合计"]]
cmp_df = cmp_df.sort_values("平均客单价(元)", ascending=False)
write_table(ws, cmp_df, 4, widths=[14, 9, 15] + [10] * 7,
            fmts={"件数": "#,##0.0", "平均客单价(元)": "#,##0", **{c: "0.0%" for c in LABELS + ["合计"]}})

from openpyxl.worksheet.properties import PageSetupProperties
for sh in wb.worksheets:
    sh.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    sh.page_setup.fitToWidth = 1
    sh.page_setup.fitToHeight = 0 if sh.title.startswith("06") else 1
    sh.page_setup.orientation = "portrait"
    sh.page_margins.left = sh.page_margins.right = 0.5
    sh.page_margins.top = sh.page_margins.bottom = 0.6
    if sh.title.startswith("06"):
        sh.print_title_rows = "4:4"

wb.save(OUTF)
print("saved", OUTF, f"{os.path.getsize(OUTF)/1024:.0f} KB")
print("sheets:", [s.title for s in wb.worksheets])
print("\n区间分布:\n", tbl.to_string(index=False))
