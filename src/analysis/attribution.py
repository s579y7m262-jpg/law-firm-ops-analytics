# -*- coding: utf-8 -*-
"""
业绩下滑归因 —— 数据构建
口径：
  业绩：收入表「业绩金额」按月汇总（1 月 1,439 万经业务确认无误）
  约见：订单表「约见时间」，按主/客场拆分；1 月主场 312 条经业务确认异常，趋势图标注但不作基期
  退费：结案退费返点台账中「列12＝是/否」的退费板块
输出：build/payload_attribution.json
"""
import pandas as pd, numpy as np, re, json, os, warnings
warnings.filterwarnings("ignore")

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
DATA = os.path.join(ROOT, "data", "raw")
BUILD = os.path.join(ROOT, "build")
os.makedirs(BUILD, exist_ok=True)


def num(s):
    if isinstance(s, (int, float, np.integer, np.floating)):
        return float(s) if pd.notna(s) else np.nan
    t = re.sub(r"[^\d.\-]", "", str(s))
    try:
        return float(t) if t not in ("", "-", ".") else np.nan
    except Exception:
        return np.nan


# ── 1. 业绩（收入表口径，沿用 payload_v5）──
P5 = json.load(open(os.path.join(BUILD, "payload.json"), encoding="utf-8"))
rev = {m["mon"]: m["revenue"] / 1e4 for m in P5["month"]}
fee = {m["mon"]: m["fee_count"] for m in P5["month"]}
case_n = {m["mon"]: m["case_count"] for m in P5["month"]}

# ── 2. 约见（订单表）──
d1 = pd.read_excel(os.path.join(DATA, "订单表_2026.1-2026.9.xlsx"), sheet_name="订单导出明细")
d2 = pd.read_excel(os.path.join(DATA, "订单表_2026.10.xlsx"), sheet_name="订单导出明细")
a = pd.concat([d1, d2], ignore_index=True)
a["约见"] = pd.to_datetime(a["约见时间"], errors="coerce")
a = a[a["约见"].notna()].copy()
a["m"] = a["约见"].dt.month
a["合同"] = a["合同总金额"].map(num).fillna(0)
a["成交"] = pd.to_datetime(a["成交时间"], errors="coerce").notna()
a["类"] = np.where(a["约见类型"].astype(str).str.contains("主场"), "主场", "客场")

g = a.groupby(["m", "类"]).agg(
    约见=("订单编号", "size"), 成交=("成交", "sum"),
    合同=("合同", lambda x: x.sum() / 1e4)).reset_index()

# ── 3. 退费（新台账退费板块）──
from openpyxl import load_workbook
wb = load_workbook(os.path.join(DATA, "给数据组结案，退费，返点汇总表10-30.xlsx"), data_only=True)
fr = []
for sn in wb.sheetnames:
    ws = wb[sn]
    for r in range(2, ws.max_row + 1):
        v11 = ws.cell(row=r, column=11).value
        if not isinstance(v11, (int, float)):
            continue
        if ws.cell(row=r, column=12).value in ("是", "否"):
            fr.append({"ym": "20" + sn[:2] + "-" + sn[2:], "amt": float(v11)})
fdf = pd.DataFrame(fr)
fdf = fdf[fdf["ym"].str[:4] == "2026"]
fdf["m"] = fdf["ym"].str[-2:].astype(int)
fv = fdf.groupby("m")["amt"].sum().abs() / 1e4

# ── 4. 组装月度 ──
month = []
for m in range(1, 11):
    sub = g[g["m"] == m]
    h = sub[sub["类"] == "主场"]
    k = sub[sub["类"] == "客场"]
    hv = float(h["约见"].sum()); kv = float(k["约见"].sum())
    hc = float(h["合同"].sum()); kc = float(k["合同"].sum())
    hd = float(h["成交"].sum()); kd = float(k["成交"].sum())
    month.append({
        "mon": m,
        "rev": round(rev.get(m, 0), 1),
        "fee": fee.get(m),
        "case_n": round(case_n.get(m, 0), 1),
        "meet": int(hv + kv), "meet_h": int(hv), "meet_k": int(kv),
        "deal_h": int(hd), "deal_k": int(kd),
        "deal": int(hd + kd),
        "cont_h": round(hc, 1), "cont_k": round(kc, 1), "cont": round(hc + kc, 1),
        "arpu_h": round(hc / hd, 0) if hd else None,
        "arpu_k": round(kc / kd, 0) if kd else None,
        "meet_arpu_h": round(hc / hv, 2) if hv else None,
        "meet_arpu_k": round(kc / kv, 2) if kv else None,
        "refund": round(float(fv.get(m, np.nan)), 1) if m in fv.index else None,
        "refund_rate": round(float(fv.get(m, np.nan)) / rev.get(m, np.nan) * 100, 1) if m in fv.index else None,
        "complete": m < 10,
    })

# ── 5. 归因拆解（3 月 → 9 月，两因子 × 两类型）──
def pick(m):
    return next(x for x in month if x["mon"] == m)
b, e = pick(3), pick(9)
attr = {"base_mon": 3, "end_mon": 9, "total_from": b["cont"], "total_to": e["cont"],
        "total_delta": round(e["cont"] - b["cont"], 1), "items": []}
for key, label in [("h", "主场约见"), ("k", "客场约见")]:
    mv0, mv1 = b[f"meet_{key}"], e[f"meet_{key}"]
    ar0, ar1 = b[f"meet_arpu_{key}"], e[f"meet_arpu_{key}"]
    c0, c1 = b[f"cont_{key}"], e[f"cont_{key}"]
    attr["items"].append({
        "label": label,
        "meet_from": mv0, "meet_to": mv1,
        "per_from": ar0, "per_to": ar1,
        "cont_from": c0, "cont_to": c1, "delta": round(c1 - c0, 1),
        "qty_effect": round((mv1 - mv0) * ar0, 1),
        "price_effect": round((ar1 - ar0) * mv1, 1),
    })
attr["qty_total"] = round(sum(i["qty_effect"] for i in attr["items"]), 1)
attr["price_total"] = round(sum(i["price_effect"] for i in attr["items"]), 1)

# ── 6. 相关性 / 弹性（2–9 月，剔除 1 月异常）──
s = pd.DataFrame([x for x in month if 2 <= x["mon"] <= 9])
def corr(x, y): return float(np.corrcoef(x, y)[0, 1])
stat = {
    "corr_meet_deal": round(corr(s["meet"], s["deal"]), 3),
    "corr_meet_rev": round(corr(s["meet"], s["rev"]), 3),
    "corr_k_rev": round(corr(s["meet_k"], s["rev"]), 3),
    "corr_h_rev": round(corr(s["meet_h"], s["rev"]), 3),
    "corr_rate_rev": round(corr(s["deal"] / s["meet"] * 100, s["rev"]), 3),
    "corr_arpu_rev": round(corr(s["cont"] / s["deal"].replace(0, np.nan), s["rev"]), 3),
}
lg = np.polyfit(np.log(s["meet"]), np.log(s["cont"]), 1)
stat["elasticity"] = round(float(lg[0]), 2)

# ── 7. 数据质量 ──
dq = [
    {"level": "warn", "item": "1 月主场约见 312 条异常",
     "detail": "1 月主场约见 312 条，2–10 月仅 10–58 条；业务确认为统计异常，业绩 1,439 万不受影响",
     "impact": "趋势与归因一律以 2 月起算，归因基期取 3 月"},
    {"level": "warn", "item": "结案与返点记账不完整",
     "detail": "结案：2 月 0 笔、4 月 22 笔、7 月 2 笔；返点：2026 年 2–6 月、9 月整月为 0",
     "impact": "按批次记账，不可用于月度趋势，本次未纳入归因"},
    {"level": "info", "item": "退费口径",
     "detail": "取台账中列12＝「是/否」的退费板块，金额取绝对值；1–9 月合计 740.5 万",
     "impact": "退费率 7.9%，无恶化趋势，非本次下滑主因"},
    {"level": "info", "item": "业绩与合同额口径不同源",
     "detail": "业绩用收入表（含中期尾款等全部款项），合同额用订单表合同总金额",
     "impact": "两者绝对值不可直接相减，趋势分别看"},
    {"level": "warn", "item": "9–10 月主场样本过小",
     "detail": "9 月主场 11 场（8 场有合同）、10 月 10 场（2 场有合同）",
     "impact": "主场单均合同额的跳升主要来自小样本，不宜解读为价格提升"},
    {"level": "info", "item": "10 月未完整",
     "detail": "约见截至 2026-10-29，10 月 63 场；业绩仅 332 万",
     "impact": "10 月不参与趋势判断"},
]

payload = {
    "meta": {"generated_at": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M"),
             "window": "2026.1–10", "last_date": str(a["约见"].max().date())},
    "month": month, "attr": attr, "stat": stat, "dq": dq,
}
json.dump(payload, open(os.path.join(BUILD, "payload_attribution.json"), "w", encoding="utf-8"),
          ensure_ascii=False, default=float)

print("归因：", attr["total_from"], "→", attr["total_to"], "差异", attr["total_delta"])
for i in attr["items"]:
    print(f"  {i['label']}: {i['cont_from']}→{i['cont_to']} ({i['delta']:+.1f}) 量效应 {i['qty_effect']:+.1f} 价效应 {i['price_effect']:+.1f}")
print(f"  量效应合计 {attr['qty_total']:+.1f} / 价效应合计 {attr['price_total']:+.1f}")
print("统计：", stat)
print("\n月度：")
for x in month:
    print(f"  {x['mon']}月 业绩{x['rev']:7.1f} 约见{x['meet']:4d}(主{x['meet_h']:3d}/客{x['meet_k']:3d}) 成交{x['deal']:3d} 合同{x['cont']:7.1f} 退费率{x['refund_rate']}")
